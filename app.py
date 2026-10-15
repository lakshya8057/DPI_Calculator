"""
DPI Calculator & Image Enhancement Service
==========================================
Measures effective print DPI for placed figures in books, enhances them
to a target DPI (Lanczos + light sharpen or optional Real-ESRGAN), exports
them with embedded DPI metadata, and extracts original figures from DOCX books.
"""

from __future__ import annotations

import base64
import csv
import io
import json
import math
import os
import shutil
import subprocess
import tempfile
import zipfile
from typing import Optional, Tuple

import numpy as np
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image, ImageFilter, ImageOps

from docx_figures import build_extracted_figures_zip, parse_docx_figures

MAX_FILE_SIZE = 25 * 1024 * 1024  # 25 MB for single images
MAX_DOCX_SIZE = 150 * 1024 * 1024  # 150 MB for docx books
MAX_DIMENSION = 12000  # Hard limit on pixel width or height
ALLOWED_FORMATS = {"PNG", "JPEG", "JPG", "WEBP", "TIFF"}

app = FastAPI(title="DPI Calculator & Image Enhancer", version="1.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "Content-Disposition",
        "X-Enhanced-Width",
        "X-Enhanced-Height",
        "X-Effective-DPI",
        "X-Upscale-Factor",
        "X-Warning",
    ],
)


def is_realesrgan_available() -> bool:
    """Check if the realesrgan-ncnn-vulkan executable is on PATH."""
    return shutil.which("realesrgan-ncnn-vulkan") is not None


def calculate_effective_dpi(width: int, placed_width_in: float) -> int:
    """Calculate effective print DPI rounded to the nearest integer.

    Effective DPI = image pixel width / placed width in inches
    """
    if placed_width_in <= 0:
        raise ValueError("Placed width must be greater than 0")
    return int(round(width / placed_width_in))


def get_status_badge(dpi: int) -> dict:
    """Return status category and message based on effective DPI."""
    if dpi >= 300:
        return {"level": "ok", "label": "OK", "color": "green"}
    elif dpi >= 200:
        return {"level": "warning", "label": "Re-export to 300 DPI", "color": "amber"}
    else:
        return {"level": "danger", "label": "Will print soft", "color": "red"}


def calculate_target_dimensions(
    width: int, height: int, placed_width_in: float, target_dpi: int
) -> Tuple[int, int, bool]:
    """Calculate target dimensions.

    Returns:
        (target_width, target_height, needs_resize)
    """
    target_width = math.ceil(placed_width_in * target_dpi)
    if width >= target_width:
        return width, height, False

    aspect_ratio = height / width
    target_height = int(round(target_width * aspect_ratio))
    return target_width, target_height, True


def detect_diagram_heuristic(img: Image.Image) -> bool:
    """Heuristic to detect diagrams, charts, or text-heavy screenshots.

    Diagrams typically have sharp lines and a limited palette of unique colors.
    """
    try:
        sample = img.copy()
        sample.thumbnail((256, 256), Image.Resampling.NEAREST)
        rgb_sample = sample.convert("RGB")

        # Fast check: getcolors returns None if unique colors exceed maxcolors
        colors = rgb_sample.getcolors(maxcolors=1024)
        if colors is not None and len(colors) < 512:
            return True

        arr = np.array(rgb_sample).reshape(-1, 3)
        unique_count = len(np.unique(arr, axis=0))
        return unique_count < 1000
    except Exception:
        return False


def sanitize_input_image(img: Image.Image) -> Tuple[Image.Image, bool]:
    """Transpose EXIF orientation, convert palette images, and detect alpha.

    Returns:
        (processed_image, has_alpha)
    """
    img = ImageOps.exif_transpose(img)

    has_alpha = False
    if img.mode in ("RGBA", "LA"):
        has_alpha = True
    elif img.mode == "P":
        if "transparency" in img.info:
            img = img.convert("RGBA")
            has_alpha = True
        else:
            img = img.convert("RGB")
    elif img.mode == "CMYK":
        img = img.convert("RGB")

    return img, has_alpha


def enhance_with_lanczos(
    img: Image.Image, target_width: int, target_height: int, has_alpha: bool
) -> Image.Image:
    """Resize image using Lanczos with light unsharp mask, preserving transparency."""
    if has_alpha:
        rgba = img.convert("RGBA")
        r, g, b, a = rgba.split()
        rgb = Image.merge("RGB", (r, g, b)).resize(
            (target_width, target_height), Image.Resampling.LANCZOS
        )
        rgb_sharp = rgb.filter(ImageFilter.UnsharpMask(radius=1.2, percent=60, threshold=2))
        a_resized = a.resize((target_width, target_height), Image.Resampling.LANCZOS)
        r_new, g_new, b_new = rgb_sharp.split()
        return Image.merge("RGBA", (r_new, g_new, b_new, a_resized))
    else:
        rgb = img.convert("RGB")
        resized = rgb.resize((target_width, target_height), Image.Resampling.LANCZOS)
        return resized.filter(ImageFilter.UnsharpMask(radius=1.2, percent=60, threshold=2))


def enhance_with_realesrgan(
    img: Image.Image, target_width: int, target_height: int
) -> Image.Image:
    """Run Real-ESRGAN binary with temp files and Lanczos downscaling if needed."""
    binary = shutil.which("realesrgan-ncnn-vulkan")
    if not binary:
        raise RuntimeError("realesrgan-ncnn-vulkan binary not found on PATH")

    # Pick smallest scale in [2, 3, 4] that reaches target_width
    scale = 4
    for s in [2, 3, 4]:
        if img.width * s >= target_width:
            scale = s
            break

    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "in.png")
        out_path = os.path.join(tmpdir, "out.png")

        img.save(in_path, format="PNG")

        cmd = [binary, "-i", in_path, "-o", out_path, "-s", str(scale)]
        proc = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90
        )
        if proc.returncode != 0 or not os.path.exists(out_path):
            err_msg = proc.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"Real-ESRGAN process error ({proc.returncode}): {err_msg}")

        with Image.open(out_path) as out_img:
            out_img.load()
            if out_img.width > target_width or out_img.height > target_height:
                return out_img.resize(
                    (target_width, target_height), Image.Resampling.LANCZOS
                )
            return out_img.copy()


def process_image(
    image_bytes: bytes,
    placed_width_in: float,
    target_dpi: int = 300,
    method: str = "lanczos",
    output_format: str = "png",
) -> Tuple[bytes, dict]:
    """Enhance an image and return encoded output bytes with metadata."""
    if placed_width_in <= 0:
        raise HTTPException(status_code=400, detail="Placed width must be positive.")
    if target_dpi < 50 or target_dpi > 2400:
        raise HTTPException(status_code=400, detail="Target DPI must be between 50 and 2400.")

    out_fmt = output_format.strip().lower()
    if out_fmt in ("jpg", "jpeg"):
        out_fmt = "jpeg"
    elif out_fmt == "png":
        out_fmt = "png"
    else:
        raise HTTPException(status_code=400, detail=f"Unsupported output format: {output_format}")

    try:
        raw_img = Image.open(io.BytesIO(image_bytes))
        raw_img.load()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid image file: {str(exc)}")

    img, has_alpha = sanitize_input_image(raw_img)

    if has_alpha and out_fmt == "jpeg":
        raise HTTPException(
            status_code=400,
            detail="Cannot export image with transparency as JPEG. Use PNG to preserve transparency.",
        )

    orig_width, orig_height = img.width, img.height
    target_width, target_height, needs_resize = calculate_target_dimensions(
        orig_width, orig_height, placed_width_in, target_dpi
    )

    if target_width > MAX_DIMENSION or target_height > MAX_DIMENSION:
        raise HTTPException(
            status_code=400,
            detail=f"Target dimensions ({target_width} x {target_height}) exceed maximum allowed limit of {MAX_DIMENSION} px.",
        )

    warning_msg = None
    if not needs_resize:
        enhanced = img.copy()
    else:
        if method.lower() == "realesrgan":
            try:
                enhanced = enhance_with_realesrgan(img, target_width, target_height)
            except Exception as e:
                warning_msg = f"Real-ESRGAN failed or not found ({e}). Fell back to Lanczos."
                enhanced = enhance_with_lanczos(img, target_width, target_height, has_alpha)
        else:
            enhanced = enhance_with_lanczos(img, target_width, target_height, has_alpha)

    # Encode with embedded DPI metadata
    out_buf = io.BytesIO()
    if out_fmt == "png":
        enhanced.save(out_buf, format="PNG", dpi=(target_dpi, target_dpi), optimize=True)
        media_type = "image/png"
    else:
        rgb_out = enhanced.convert("RGB") if enhanced.mode != "RGB" else enhanced
        rgb_out.save(
            out_buf,
            format="JPEG",
            dpi=(target_dpi, target_dpi),
            quality=95,
            subsampling=0,
        )
        media_type = "image/jpeg"

    out_bytes = out_buf.getvalue()
    effective_dpi_after = calculate_effective_dpi(enhanced.width, placed_width_in)
    upscale_factor = round(enhanced.width / orig_width, 2)

    meta = {
        "width": enhanced.width,
        "height": enhanced.height,
        "effective_dpi": effective_dpi_after,
        "upscale_factor": upscale_factor,
        "media_type": media_type,
        "format": out_fmt,
        "warning": warning_msg,
    }
    return out_bytes, meta


def make_thumbnail_data_url(raw_bytes: bytes, max_size: int = 100) -> Optional[str]:
    """Generate a lightweight base64 thumbnail for browser preview."""
    try:
        with Image.open(io.BytesIO(raw_bytes)) as img:
            img.thumbnail((max_size, max_size))
            thumb_buf = io.BytesIO()
            if img.mode in ("RGBA", "LA", "P"):
                img.convert("RGBA").save(thumb_buf, format="PNG")
                b64 = base64.b64encode(thumb_buf.getvalue()).decode("ascii")
                return f"data:image/png;base64,{b64}"
            else:
                img.convert("RGB").save(thumb_buf, format="JPEG", quality=80)
                b64 = base64.b64encode(thumb_buf.getvalue()).decode("ascii")
                return f"data:image/jpeg;base64,{b64}"
    except Exception:
        return None


@app.get("/api/capabilities")
async def get_capabilities():
    """Return backend capabilities, such as whether Real-ESRGAN binary is available."""
    return {
        "realesrgan_available": is_realesrgan_available(),
        "max_dimension": MAX_DIMENSION,
        "max_file_size_mb": MAX_FILE_SIZE // (1024 * 1024),
        "max_docx_size_mb": MAX_DOCX_SIZE // (1024 * 1024),
    }


@app.post("/api/analyze")
async def analyze_image(file: UploadFile = File(...)):
    """Analyze an uploaded image and return its dimensions, embedded DPI, and heuristics."""
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File exceeds maximum allowed size of {MAX_FILE_SIZE // (1024 * 1024)} MB.",
        )

    try:
        img = Image.open(io.BytesIO(contents))
        img.load()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Cannot parse image: {str(e)}")

    orig_format = (img.format or "UNKNOWN").upper()
    if orig_format not in ALLOWED_FORMATS and orig_format not in ("JPEG", "JPG"):
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported format '{orig_format}'. Supported formats: PNG, JPG, WEBP, TIFF.",
        )

    sanitized_img, has_alpha = sanitize_input_image(img)

    raw_dpi = img.info.get("dpi")
    embedded_dpi = None
    if raw_dpi and isinstance(raw_dpi, (tuple, list)) and len(raw_dpi) >= 2:
        try:
            dpi_x = float(raw_dpi[0])
            dpi_y = float(raw_dpi[1])
            if dpi_x > 1 and dpi_y > 1:
                embedded_dpi = [round(dpi_x), round(dpi_y)]
        except (ValueError, TypeError):
            embedded_dpi = None

    is_diagram = detect_diagram_heuristic(sanitized_img)

    return {
        "filename": file.filename,
        "width": sanitized_img.width,
        "height": sanitized_img.height,
        "format": orig_format,
        "mode": sanitized_img.mode,
        "has_alpha": has_alpha,
        "embedded_dpi": embedded_dpi,
        "is_diagram_heuristic": is_diagram,
        "realesrgan_available": is_realesrgan_available(),
    }


@app.post("/api/enhance")
async def enhance_endpoint(
    file: UploadFile = File(...),
    placed_width_in: float = Form(4.88),
    target_dpi: int = Form(300),
    method: str = Form("lanczos"),
    output_format: str = Form("png"),
):
    """Enhance the uploaded image to target DPI and return as downloadable file."""
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File exceeds maximum size of {MAX_FILE_SIZE // (1024 * 1024)} MB.",
        )

    out_bytes, meta = process_image(
        image_bytes=contents,
        placed_width_in=placed_width_in,
        target_dpi=target_dpi,
        method=method,
        output_format=output_format,
    )

    stem = os.path.splitext(file.filename or "image")[0]
    ext = "png" if meta["format"] == "png" else "jpg"
    download_filename = f"{stem}_{target_dpi}dpi.{ext}"

    headers = {
        "Content-Disposition": f'attachment; filename="{download_filename}"',
        "X-Enhanced-Width": str(meta["width"]),
        "X-Enhanced-Height": str(meta["height"]),
        "X-Effective-DPI": str(meta["effective_dpi"]),
        "X-Upscale-Factor": str(meta["upscale_factor"]),
    }
    if meta.get("warning"):
        headers["X-Warning"] = meta["warning"]

    return Response(content=out_bytes, media_type=meta["media_type"], headers=headers)


@app.post("/api/batch-enhance")
async def batch_enhance_endpoint(
    files: list[UploadFile] = File(...),
    placed_width_in: float = Form(4.88),
    target_dpi: int = Form(300),
    method: str = Form("lanczos"),
    output_format: str = Form("png"),
    csv_widths: Optional[UploadFile] = File(None),
):
    """Batch enhance multiple images, returning a ZIP archive with report.csv."""
    custom_widths = {}
    if csv_widths is not None and csv_widths.filename:
        csv_bytes = await csv_widths.read()
        try:
            reader = csv.reader(io.StringIO(csv_bytes.decode("utf-8-sig")))
            for row in reader:
                if len(row) >= 2:
                    try:
                        fname = row[0].strip()
                        w_val = float(row[1].strip())
                        if fname and w_val > 0:
                            custom_widths[fname] = w_val
                    except ValueError:
                        continue
        except Exception:
            pass

    zip_buffer = io.BytesIO()
    report_rows = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_out:
        for file in files:
            contents = await file.read()
            if len(contents) > MAX_FILE_SIZE or len(contents) == 0:
                continue

            fname = file.filename or "image"
            item_placed_width = custom_widths.get(fname, placed_width_in)

            try:
                # Analyze initial dimensions & DPI
                with Image.open(io.BytesIO(contents)) as orig_raw:
                    orig_img, _ = sanitize_input_image(orig_raw)
                    before_dpi = calculate_effective_dpi(orig_img.width, item_placed_width)
                    status_before = get_status_badge(before_dpi)["label"]
                    before_pixels = f"{orig_img.width}x{orig_img.height}"

                out_bytes, meta = process_image(
                    image_bytes=contents,
                    placed_width_in=item_placed_width,
                    target_dpi=target_dpi,
                    method=method,
                    output_format=output_format,
                )

                stem = os.path.splitext(fname)[0]
                ext = "png" if meta["format"] == "png" else "jpg"
                out_name = f"{stem}_{target_dpi}dpi.{ext}"

                zip_out.writestr(out_name, out_bytes)

                report_rows.append(
                    {
                        "filename": fname,
                        "pixels": before_pixels,
                        "placed_width_in": item_placed_width,
                        "effective_dpi_before": before_dpi,
                        "status_before": status_before,
                        "output_pixels": f"{meta['width']}x{meta['height']}",
                        "effective_dpi_after": meta["effective_dpi"],
                    }
                )
            except Exception as e:
                report_rows.append(
                    {
                        "filename": fname,
                        "pixels": "Error",
                        "placed_width_in": item_placed_width,
                        "effective_dpi_before": 0,
                        "status_before": f"Failed: {str(e)}",
                        "output_pixels": "-",
                        "effective_dpi_after": 0,
                    }
                )

        # Build report.csv
        csv_buf = io.StringIO()
        fieldnames = [
            "filename",
            "pixels",
            "placed_width_in",
            "effective_dpi_before",
            "status_before",
            "output_pixels",
            "effective_dpi_after",
        ]
        writer = csv.DictWriter(csv_buf, fieldnames=fieldnames)
        writer.writeheader()
        for r in report_rows:
            writer.writerow(r)

        zip_out.writestr("report.csv", csv_buf.getvalue())

    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="enhanced_batch.zip"'},
    )


# =========================================================================
# DOCX Figure Extraction Endpoints
# =========================================================================


@app.post("/api/docx/analyze")
async def docx_analyze_endpoint(
    file: UploadFile = File(...),
    caption_position: str = Form("auto"),
):
    """Analyze a DOCX document to find all figures and captions, returning a preview table JSON."""
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a .docx archive.")

    # Write to a temporary file for streaming extraction
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp_path = tmp.name
    try:
        total_read = 0
        while chunk := await file.read(1024 * 1024):
            total_read += len(chunk)
            if total_read > MAX_DOCX_SIZE:
                raise HTTPException(
                    status_code=413,
                    detail=f"DOCX file exceeds {MAX_DOCX_SIZE // (1024 * 1024)} MB limit.",
                )
            tmp.write(chunk)
        tmp.flush()
        tmp.close()  # Must close file handle before reading or unlinking on Windows

        parse_res = parse_docx_figures(tmp_path, caption_position=caption_position)

        # Generate preview data URLs for table display
        with zipfile.ZipFile(tmp_path, "r") as docx_zip:
            parts_json: dict[str, list[dict]] = {}
            for part_num, figs in parse_res.parts.items():
                figs_list = []
                for f in figs:
                    thumb = None
                    try:
                        raw = docx_zip.read(f.image.media_path)
                        thumb = make_thumbnail_data_url(raw, max_size=90)
                    except Exception:
                        pass

                    figs_list.append(
                        {
                            "sequence": f.sequence,
                            "part": f.part,
                            "figure_number": f.figure_number,
                            "page_number": f.page_number,
                            "caption_text": f.caption_text,
                            "media_filename": f.image.media_filename,
                            "original_media_file": f.image.media_path,
                            "output_path": f.output_path,
                            "extension": f.image.extension,
                            "pixel_width": f.image.pixel_width,
                            "pixel_height": f.image.pixel_height,
                            "pixels_str": (
                                f"{f.image.pixel_width}x{f.image.pixel_height}"
                                if f.image.pixel_width
                                else ("Vector" if f.image.is_vector else "-")
                            ),
                            "placed_width_in": f.image.placed_width_in,
                            "effective_dpi": f.effective_dpi,
                            "print_status": f.print_status,
                            "sha256": f.image.sha256,
                            "warnings": f.warnings,
                            "thumbnail": thumb,
                        }
                    )
                parts_json[str(part_num)] = figs_list

            unmatched_list = []
            for u in parse_res.unmatched_images:
                thumb = None
                try:
                    raw = docx_zip.read(u.media_path)
                    thumb = make_thumbnail_data_url(raw, max_size=90)
                except Exception:
                    pass

                status = "OK (Vector)" if u.is_vector else (
                    "OK" if u.effective_dpi and u.effective_dpi >= 300 else (
                        "Re-export to 300 DPI" if u.effective_dpi and u.effective_dpi >= 200 else "Will print soft"
                    )
                )
                unmatched_list.append(
                    {
                        "doc_index": u.doc_index,
                        "page_number": u.page_number,
                        "media_filename": u.media_filename,
                        "original_media_file": u.media_path,
                        "output_path": f"_unmatched/Unmatched_{u.doc_index:03d}__{u.media_filename}",
                        "extension": u.extension,
                        "pixel_width": u.pixel_width,
                        "pixel_height": u.pixel_height,
                        "pixels_str": (
                            f"{u.pixel_width}x{u.pixel_height}"
                            if u.pixel_width
                            else ("Vector" if u.is_vector else "-")
                        ),
                        "placed_width_in": u.placed_width_in,
                        "effective_dpi": u.effective_dpi,
                        "print_status": status,
                        "sha256": u.sha256,
                        "warnings": ["No caption matched within search window"],
                        "thumbnail": thumb,
                    }
                )

            unmatched_captions_list = [
                {"paragraph_idx": c.paragraph_idx, "full_text": c.full_text}
                for c in parse_res.unmatched_captions
            ]

            total_matched = len(parse_res.matched_figures)
            total_unmatched = len(parse_res.unmatched_images)
            total_images = total_matched + total_unmatched

            return {
                "filename": file.filename,
                "summary": {
                    "total_images": total_images,
                    "matched": total_matched,
                    "unmatched": total_unmatched,
                    "duplicates": len(parse_res.duplicates),
                    "gaps": len(parse_res.gaps),
                },
                "parts": parts_json,
                "unmatched": unmatched_list,
                "unmatched_captions": unmatched_captions_list,
                "duplicates": parse_res.duplicates,
                "gaps": parse_res.gaps,
                "report_text": parse_res.report_text,
                "warnings": parse_res.warnings,
            }
    finally:
        try:
            tmp.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.post("/api/docx/extract")
async def docx_extract_endpoint(
    file: UploadFile = File(...),
    caption_position: str = Form("auto"),
    enhanced_files: list[UploadFile] = File(default=[]),
    enhanced_meta: Optional[str] = Form(None),
):
    """Extract figures into an uncompressed ZIP, using enhanced bytes for any enhanced images."""
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a .docx archive.")

    stem = os.path.splitext(file.filename or "book")[0]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp_path = tmp.name
    try:
        total_read = 0
        while chunk := await file.read(1024 * 1024):
            total_read += len(chunk)
            if total_read > MAX_DOCX_SIZE:
                raise HTTPException(
                    status_code=413,
                    detail=f"DOCX file exceeds {MAX_DOCX_SIZE // (1024 * 1024)} MB limit.",
                )
            tmp.write(chunk)
        tmp.flush()
        tmp.close()

        # Read any enhanced files sent from the frontend
        enhanced_images: dict[str, bytes] = {}
        for ef in enhanced_files:
            if ef.filename:
                content = await ef.read()
                if content:
                    enhanced_images[ef.filename] = content

        meta_dict: dict[str, dict] = {}
        if enhanced_meta:
            try:
                parsed = json.loads(enhanced_meta)
                if isinstance(parsed, list):
                    for item in parsed:
                        if "media_path" in item:
                            meta_dict[item["media_path"]] = item
                elif isinstance(parsed, dict):
                    meta_dict = parsed
            except Exception:
                pass

        zip_bytes, _ = build_extracted_figures_zip(
            tmp_path,
            caption_position=caption_position,
            root_folder_name=f"{stem}_figures",
            enhanced_images=enhanced_images,
            enhanced_meta=meta_dict,
        )

        headers = {
            "Content-Disposition": f'attachment; filename="{stem}_figures.zip"',
        }
        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers=headers,
        )
    finally:
        try:
            tmp.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.post("/api/docx/enhance-figure")
async def docx_enhance_figure_endpoint(
    file: UploadFile = File(...),
    media_path: str = Form(...),
    placed_width_in: float = Form(4.88),
    target_dpi: int = Form(300),
    method: str = Form("lanczos"),
    output_format: str = Form("png"),
):
    """Extract and enhance a single figure image directly from DOCX in place."""
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a .docx archive.")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp_path = tmp.name
    try:
        while chunk := await file.read(1024 * 1024):
            tmp.write(chunk)
        tmp.flush()
        tmp.close()

        with zipfile.ZipFile(tmp_path, "r") as docx_zip:
            if media_path not in docx_zip.namelist():
                raise HTTPException(
                    status_code=404,
                    detail=f"Media entry '{media_path}' not found in docx.",
                )
            raw_bytes = docx_zip.read(media_path)

        ext = os.path.splitext(media_path)[1].lower()
        if ext in (".emf", ".wmf", ".svg"):
            raise HTTPException(
                status_code=400,
                detail=f"Vector format '{ext}' is resolution-independent and cannot be raster-enhanced.",
            )

        out_bytes, meta = process_image(
            image_bytes=raw_bytes,
            placed_width_in=placed_width_in,
            target_dpi=target_dpi,
            method=method,
            output_format=output_format,
        )

        stem = os.path.splitext(os.path.basename(media_path))[0]
        download_filename = f"{stem}_{target_dpi}dpi.png"

        headers = {
            "Content-Disposition": f'attachment; filename="{download_filename}"',
            "X-Enhanced-Width": str(meta["width"]),
            "X-Enhanced-Height": str(meta["height"]),
            "X-Effective-DPI": str(meta["effective_dpi"]),
            "X-Upscale-Factor": str(meta["upscale_factor"]),
            "X-Original-Media-Path": media_path,
        }
        if meta.get("warning"):
            headers["X-Warning"] = meta["warning"]

        return Response(content=out_bytes, media_type="image/png", headers=headers)
    finally:
        try:
            tmp.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.post("/api/docx/extract-single")
async def docx_extract_single_endpoint(
    file: UploadFile = File(...),
    media_path: str = Form(...),
):
    """Extract a single original media file from the DOCX to send directly to Enhancer."""
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a .docx archive.")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp_path = tmp.name
    try:
        while chunk := await file.read(1024 * 1024):
            tmp.write(chunk)
        tmp.flush()
        tmp.close()

        with zipfile.ZipFile(tmp_path, "r") as docx_zip:
            if media_path not in docx_zip.namelist():
                raise HTTPException(
                    status_code=404,
                    detail=f"Media entry '{media_path}' not found in docx.",
                )
            data = docx_zip.read(media_path)

        ext = os.path.splitext(media_path)[1].lower()
        media_type = "image/png" if ext == ".png" else (
            "image/jpeg" if ext in (".jpg", ".jpeg") else "application/octet-stream"
        )

        headers = {
            "Content-Disposition": f'attachment; filename="{os.path.basename(media_path)}"',
        }
        return Response(content=data, media_type=media_type, headers=headers)
    finally:
        try:
            tmp.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


# Mount static directory for frontend
os.makedirs("static", exist_ok=True)
app.mount("/", StaticFiles(directory="static", html=True), name="static")

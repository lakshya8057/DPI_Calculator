"""
DOCX Figure Extraction & Analysis Module
========================================
Parses Word .docx archives directly using zipfile and lxml to locate
all figure images, extract their original uncompressed bytes without modification,
match them with figure captions, and package them into structured ZIP archives.
"""

from __future__ import annotations

import csv
import hashlib
import io
import os
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

from lxml import etree
from PIL import Image

# XML Namespaces used in OpenXML WordprocessingML
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Regex to match caption patterns: "Figure 1.1: ...", "Fig. 1.1 - ...", "Figure 1.1. ..."
CAPTION_REGEX = re.compile(
    r"^\s*(?:fig(?:ure)?\.?)\s*(\d+)\.(\d+[a-z]?)(?:\s*[:.\-–—]\s*|\s+|$)(.*)$",
    re.IGNORECASE,
)

EMU_PER_INCH = 914400.0


@dataclass
class DocxImage:
    """Represents an image reference found in the DOCX."""
    doc_index: int
    r_id: str
    media_path: str  # Path inside zip, e.g. "word/media/image1.png"
    media_filename: str  # Basename, e.g. "image1.png"
    extension: str  # e.g. ".png"
    placed_width_in: float
    placed_height_in: float
    width_emu: int
    height_emu: int
    is_linked: bool
    is_cropped: bool
    location: str  # "body", "table", "textbox", "header", "footnote"
    paragraph_idx: int
    sha256: str = ""
    pixel_width: Optional[int] = None
    pixel_height: Optional[int] = None
    effective_dpi: Optional[int] = None
    is_vector: bool = False
    crop_info: Optional[str] = None


@dataclass
class DocxCaption:
    """Represents a matched figure caption."""
    paragraph_idx: int
    full_text: str
    part: int
    fig_num: str
    caption_body: str
    has_caption_style: bool
    is_seq: bool = False


@dataclass
class MatchedFigure:
    """An image successfully paired with a figure caption."""
    sequence: int
    part: int
    figure_number: str
    base_figure_number: str
    caption_text: str
    image: DocxImage
    output_path: str
    warnings: List[str] = field(default_factory=list)
    print_status: str = "OK"
    effective_dpi: Optional[int] = None


@dataclass
class ParseResult:
    """Complete analysis result of a DOCX document."""
    matched_figures: List[MatchedFigure]
    unmatched_images: List[DocxImage]
    unmatched_captions: List[DocxCaption]
    parts: Dict[int, List[MatchedFigure]]
    duplicates: List[str]
    gaps: List[str]
    warnings: List[str]
    report_text: str
    manifest_rows: List[Dict[str, Any]]


def sanitize_filename(name: str) -> str:
    """Remove or replace characters unsafe for Windows or POSIX filenames."""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def read_relationships(zip_file: zipfile.ZipFile, rels_path: str) -> Dict[str, Dict[str, str]]:
    """Parse a .rels XML file and map rId to relationship info."""
    rels: Dict[str, Dict[str, str]] = {}
    if rels_path not in zip_file.namelist():
        return rels

    try:
        xml_content = zip_file.read(rels_path)
        root = etree.fromstring(xml_content)
        for rel in root.xpath("//rel:Relationship", namespaces=NAMESPACES):
            r_id = rel.get("Id")
            target = rel.get("Target", "")
            target_mode = rel.get("TargetMode", "Internal")
            rel_type = rel.get("Type", "")
            if r_id:
                rels[r_id] = {
                    "target": target,
                    "target_mode": target_mode,
                    "type": rel_type,
                }
    except Exception:
        pass
    return rels


def resolve_media_zip_path(zip_file: zipfile.ZipFile, target: str, base_dir: str = "word") -> Optional[str]:
    """Resolve a target path from a rels entry to an actual zip member path."""
    clean_target = target.replace("\\", "/")
    if clean_target.startswith("/"):
        candidate = clean_target.lstrip("/")
        if candidate in zip_file.namelist():
            return candidate

    # Relative to base_dir
    norm = posixpath.normpath(posixpath.join(base_dir, clean_target))
    if norm in zip_file.namelist():
        return norm

    # Fallback search in word/media or anywhere in zip
    base_name = posixpath.basename(clean_target)
    for name in zip_file.namelist():
        if name.endswith("/" + base_name) or name == base_name:
            return name

    return None


def extract_paragraph_text(p_elem: etree._Element) -> Tuple[str, bool]:
    """Extract joined visible text from a paragraph, handling fields and SEQ runs.

    Returns:
        (text, is_seq_field)
    """
    is_seq = False
    # Check for SEQ in simple fields or instruction text
    for fld in p_elem.xpath(".//w:fldSimple", namespaces=NAMESPACES):
        instr = fld.get(f"{{{NAMESPACES['w']}}}instr", "") or fld.get("instr", "")
        if "SEQ" in instr.upper():
            is_seq = True

    for instr_elem in p_elem.xpath(".//w:instrText", namespaces=NAMESPACES):
        if instr_elem.text and "SEQ" in instr_elem.text.upper():
            is_seq = True

    # Gather visible text from w:t elements
    text_parts: List[str] = []
    for node in p_elem.iter():
        tag = etree.QName(node).localname
        if tag == "t" and node.text:
            text_parts.append(node.text)
        elif tag == "tab":
            text_parts.append("\t")
        elif tag == "br":
            text_parts.append(" ")

    full_text = "".join(text_parts).replace("\u00a0", " ").strip()
    return full_text, is_seq


def inspect_image_metadata(
    zip_file: zipfile.ZipFile, media_path: str, placed_width_in: float
) -> Tuple[Optional[int], Optional[int], Optional[int], str, bool]:
    """Read image bytes to compute SHA256 and (read-only) dimensions/DPI.

    Returns:
        (pixel_width, pixel_height, effective_dpi, sha256, is_vector)
    """
    try:
        raw_bytes = zip_file.read(media_path)
    except Exception:
        return None, None, None, "", False

    sha256 = hashlib.sha256(raw_bytes).hexdigest()
    ext = os.path.splitext(media_path)[1].lower()
    is_vector = ext in (".emf", ".wmf", ".svg", ".eps")

    pixel_width = None
    pixel_height = None
    effective_dpi = None

    if not is_vector:
        try:
            with Image.open(io.BytesIO(raw_bytes)) as pil_img:
                pixel_width, pixel_height = pil_img.size
                if placed_width_in > 0 and pixel_width > 0:
                    effective_dpi = int(round(pixel_width / placed_width_in))
        except Exception:
            pass

    return pixel_width, pixel_height, effective_dpi, sha256, is_vector


def parse_extent(elem: etree._Element) -> Tuple[int, int, float, float]:
    """Parse cx/cy from wp:extent or v:shape style.

    Returns:
        (cx, cy, width_in, height_in)
    """
    extent = elem.xpath(".//wp:extent", namespaces=NAMESPACES)
    if extent:
        try:
            cx = int(extent[0].get("cx", 0))
            cy = int(extent[0].get("cy", 0))
            return cx, cy, cx / EMU_PER_INCH, cy / EMU_PER_INCH
        except (ValueError, TypeError):
            pass

    # Try v:shape style="width:300pt;height:200pt"
    style = elem.get("style", "") or ""
    w_match = re.search(r"width:\s*([\d.]+)\s*(pt|in|px)?", style, re.IGNORECASE)
    h_match = re.search(r"height:\s*([\d.]+)\s*(pt|in|px)?", style, re.IGNORECASE)
    if w_match:
        val = float(w_match.group(1))
        unit = (w_match.group(2) or "pt").lower()
        if unit == "pt":
            w_in = val / 72.0
        elif unit == "in":
            w_in = val
        else:
            w_in = val / 96.0
        cx = int(w_in * EMU_PER_INCH)
        cy = cx
        if h_match:
            h_val = float(h_match.group(1))
            h_unit = (h_match.group(2) or "pt").lower()
            h_in = h_val / 72.0 if h_unit == "pt" else (h_val if h_unit == "in" else h_val / 96.0)
            cy = int(h_in * EMU_PER_INCH)
        return cx, cy, w_in, cy / EMU_PER_INCH

    # Default fallback: 4.88 inches
    default_emu = int(4.88 * EMU_PER_INCH)
    return default_emu, default_emu, 4.88, 4.88


def find_images_in_element(
    elem: etree._Element,
    rels: Dict[str, Dict[str, str]],
    zip_file: zipfile.ZipFile,
    location: str,
    p_idx: int,
    start_doc_index: int,
) -> List[DocxImage]:
    """Locate all images within a container element, ignoring mc:Fallback."""
    images: List[DocxImage] = []
    curr_idx = start_doc_index

    # 1. DrawingML images (w:drawing)
    for drawing in elem.xpath(".//w:drawing", namespaces=NAMESPACES):
        # Ignore drawings inside mc:Fallback
        if drawing.xpath("ancestor::mc:Fallback", namespaces=NAMESPACES):
            continue

        blips = drawing.xpath(".//a:blip", namespaces=NAMESPACES)
        if not blips:
            continue

        cx, cy, w_in, h_in = parse_extent(drawing)
        for blip in blips:
            embed_id = blip.get(f"{{{NAMESPACES['r']}}}embed") or blip.get("r:embed")
            link_id = blip.get(f"{{{NAMESPACES['r']}}}link") or blip.get("r:link")

            is_linked = False
            r_id = embed_id
            if not r_id and link_id:
                r_id = link_id
                is_linked = True

            if not r_id or r_id not in rels:
                continue

            target = rels[r_id]["target"]
            media_path = resolve_media_zip_path(zip_file, target)
            if not media_path:
                continue

            # Check cropping in a:srcRect
            is_cropped = bool(drawing.xpath(".//a:srcRect", namespaces=NAMESPACES))
            crop_info = "Cropped in Word (original uncropped image extracted)" if is_cropped else None

            pix_w, pix_h, eff_dpi, sha256, is_vec = inspect_image_metadata(
                zip_file, media_path, w_in
            )

            images.append(
                DocxImage(
                    doc_index=curr_idx,
                    r_id=r_id,
                    media_path=media_path,
                    media_filename=posixpath.basename(media_path),
                    extension=posixpath.splitext(media_path)[1].lower(),
                    placed_width_in=round(w_in, 2),
                    placed_height_in=round(h_in, 2),
                    width_emu=cx,
                    height_emu=cy,
                    is_linked=is_linked,
                    is_cropped=is_cropped,
                    crop_info=crop_info,
                    location=location,
                    paragraph_idx=p_idx,
                    sha256=sha256,
                    pixel_width=pix_w,
                    pixel_height=pix_h,
                    effective_dpi=eff_dpi,
                    is_vector=is_vec,
                )
            )
            curr_idx += 1

    # 2. Legacy VML images (w:pict)
    for pict in elem.xpath(".//w:pict", namespaces=NAMESPACES):
        if pict.xpath("ancestor::mc:Fallback", namespaces=NAMESPACES):
            continue

        imagedatas = pict.xpath(".//v:imagedata", namespaces=NAMESPACES)
        if not imagedatas:
            continue

        cx, cy, w_in, h_in = parse_extent(pict)
        for imgdata in imagedatas:
            r_id = (
                imgdata.get(f"{{{NAMESPACES['r']}}}id")
                or imgdata.get("r:id")
                or imgdata.get(f"{{{NAMESPACES['r']}}}href")
                or imgdata.get("r:href")
            )
            if not r_id or r_id not in rels:
                continue

            target = rels[r_id]["target"]
            media_path = resolve_media_zip_path(zip_file, target)
            if not media_path:
                continue

            pix_w, pix_h, eff_dpi, sha256, is_vec = inspect_image_metadata(
                zip_file, media_path, w_in
            )

            images.append(
                DocxImage(
                    doc_index=curr_idx,
                    r_id=r_id,
                    media_path=media_path,
                    media_filename=posixpath.basename(media_path),
                    extension=posixpath.splitext(media_path)[1].lower(),
                    placed_width_in=round(w_in, 2),
                    placed_height_in=round(h_in, 2),
                    width_emu=cx,
                    height_emu=cy,
                    is_linked=False,
                    is_cropped=False,
                    crop_info=None,
                    location=location,
                    paragraph_idx=p_idx,
                    sha256=sha256,
                    pixel_width=pix_w,
                    pixel_height=pix_h,
                    effective_dpi=eff_dpi,
                    is_vector=is_vec,
                )
            )
            curr_idx += 1

    return images


def parse_docx_figures(
    docx_file_or_bytes: Any, caption_position: str = "auto"
) -> ParseResult:
    """Parse a DOCX file to find figures, match captions, and prepare extraction report.

    Args:
        docx_file_or_bytes: Path string, bytes, or file-like object of the .docx archive.
        caption_position: "auto", "below", or "above".
    """
    if isinstance(docx_file_or_bytes, (bytes, bytearray)):
        zip_file = zipfile.ZipFile(io.BytesIO(docx_file_or_bytes))
    elif isinstance(docx_file_or_bytes, str):
        zip_file = zipfile.ZipFile(docx_file_or_bytes, "r")
    else:
        zip_file = zipfile.ZipFile(docx_file_or_bytes)

    # 1. Read document relationships
    doc_rels = read_relationships(zip_file, "word/_rels/document.xml.rels")

    if "word/document.xml" not in zip_file.namelist():
        raise ValueError("Invalid DOCX archive: missing word/document.xml")

    doc_xml = zip_file.read("word/document.xml")
    root = etree.fromstring(doc_xml)

    # Clean out mc:Fallback nodes so fallback copies are never processed
    for fallback in root.xpath("//mc:AlternateContent/mc:Fallback", namespaces=NAMESPACES):
        parent = fallback.getparent()
        if parent is not None:
            parent.remove(fallback)

    # 2. Walk all paragraphs in depth-first order
    paragraph_records: List[Dict[str, Any]] = []
    all_images: List[DocxImage] = []
    all_captions: List[DocxCaption] = []

    doc_image_counter = 1

    for p_idx, p in enumerate(root.xpath("//w:p", namespaces=NAMESPACES)):
        # Determine location
        in_tbl = bool(p.xpath("ancestor::w:tbl", namespaces=NAMESPACES))
        in_txbx = bool(p.xpath("ancestor::w:txbxContent", namespaces=NAMESPACES))
        location = "table" if in_tbl else ("textbox" if in_txbx else "body")

        # Extract text and caption info
        text, is_seq = extract_paragraph_text(p)
        style_vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NAMESPACES)
        has_caption_style = any("caption" in str(s).lower() for s in style_vals)

        caption_obj: Optional[DocxCaption] = None
        if text:
            m = CAPTION_REGEX.match(text)
            if m:
                part_num = int(m.group(1))
                fig_num_str = f"{m.group(1)}.{m.group(2)}"
                caption_obj = DocxCaption(
                    paragraph_idx=p_idx,
                    full_text=text,
                    part=part_num,
                    fig_num=fig_num_str,
                    caption_body=m.group(3).strip(),
                    has_caption_style=has_caption_style,
                    is_seq=is_seq,
                )
                all_captions.append(caption_obj)

        # Extract images from this paragraph
        p_images = find_images_in_element(
            p, doc_rels, zip_file, location, p_idx, doc_image_counter
        )
        doc_image_counter += len(p_images)
        all_images.extend(p_images)

        paragraph_records.append(
            {
                "p_idx": p_idx,
                "text": text,
                "is_empty": (len(text) == 0 and len(p_images) == 0),
                "images": p_images,
                "caption": caption_obj,
            }
        )

    # Check headers and footnotes for images and flag them
    for member in zip_file.namelist():
        if re.match(r"^word/(header\d*|footnotes)\.xml$", member):
            try:
                xml_data = zip_file.read(member)
                sub_root = etree.fromstring(xml_data)
                rels_name = f"word/_rels/{posixpath.basename(member)}.rels"
                sub_rels = read_relationships(zip_file, rels_name)
                sub_loc = "header" if "header" in member else "footnote"
                sub_images = find_images_in_element(
                    sub_root, sub_rels, zip_file, sub_loc, -1, doc_image_counter
                )
                doc_image_counter += len(sub_images)
                all_images.extend(sub_images)
            except Exception:
                pass

    # 3. Match images with captions
    matched_figures: List[MatchedFigure] = []
    used_captions: Set[int] = set()  # set of paragraph_idx of used captions
    matched_image_indices: Set[int] = set()

    for p_rec in paragraph_records:
        p_idx = p_rec["p_idx"]
        images_here = p_rec["images"]
        if not images_here:
            continue

        multi_image_flag = len(images_here) > 1

        for img in images_here:
            best_caption: Optional[DocxCaption] = None

            # First priority: check same paragraph if unused
            if p_rec["caption"] and p_rec["caption"].paragraph_idx not in used_captions:
                best_caption = p_rec["caption"]
            else:
                # Search candidate paragraphs according to caption_position preference
                candidates: List[Tuple[int, DocxCaption, str]] = []  # (distance, caption, direction)

                # Search below (after)
                if caption_position in ("auto", "below"):
                    non_empty_count = 0
                    for j in range(p_idx + 1, len(paragraph_records)):
                        if paragraph_records[j]["is_empty"]:
                            continue
                        non_empty_count += 1
                        cap = paragraph_records[j]["caption"]
                        if cap and cap.paragraph_idx not in used_captions:
                            candidates.append((non_empty_count, cap, "below"))
                            break  # Nearest unused caption below
                        if non_empty_count >= 3:
                            break

                # Search above (before)
                if caption_position in ("auto", "above"):
                    non_empty_count = 0
                    for j in range(p_idx - 1, -1, -1):
                        if paragraph_records[j]["is_empty"]:
                            continue
                        non_empty_count += 1
                        cap = paragraph_records[j]["caption"]
                        if cap and cap.paragraph_idx not in used_captions:
                            candidates.append((non_empty_count, cap, "above"))
                            break  # Nearest unused caption above
                        if non_empty_count >= 3:
                            break

                if candidates:
                    # Preference logic:
                    # 1. Caption style
                    # 2. Smaller distance
                    # 3. 'below' over 'above' in auto
                    def sort_key(c: Tuple[int, DocxCaption, str]):
                        dist, cap_obj, direction = c
                        style_pref = 0 if cap_obj.has_caption_style else 1
                        dir_pref = 0 if direction == "below" else 1
                        return (dist, style_pref, dir_pref)

                    candidates.sort(key=sort_key)
                    best_caption = candidates[0][1]

            if best_caption:
                used_captions.add(best_caption.paragraph_idx)
                matched_image_indices.add(img.doc_index)

                warnings: List[str] = []
                if multi_image_flag:
                    warnings.append("Multiple images in one paragraph; matched nearest unused caption")
                if img.crop_info:
                    warnings.append(img.crop_info)
                if img.is_linked:
                    warnings.append("Linked image, not embedded")
                if img.is_vector:
                    warnings.append("Vector format (resolution-independent)")
                if img.location in ("header", "footnote"):
                    warnings.append(f"Located in document {img.location}")

                # Determine print status
                if img.is_vector:
                    status = "OK (Vector)"
                elif img.effective_dpi is not None:
                    if img.effective_dpi >= 300:
                        status = "OK"
                    elif img.effective_dpi >= 200:
                        status = "Re-export to 300 DPI"
                    else:
                        status = "Will print soft"
                else:
                    status = "Unknown"

                matched_figures.append(
                    MatchedFigure(
                        sequence=0,  # Will be assigned after ordering
                        part=best_caption.part,
                        figure_number=best_caption.fig_num,
                        base_figure_number=best_caption.fig_num,
                        caption_text=best_caption.full_text,
                        image=img,
                        output_path="",
                        warnings=warnings,
                        print_status=status,
                        effective_dpi=img.effective_dpi,
                    )
                )

    # 4. Handle unmatched images and captions
    unmatched_images = [img for img in all_images if img.doc_index not in matched_image_indices]
    unmatched_captions = [cap for cap in all_captions if cap.paragraph_idx not in used_captions]

    # 5. Order figures within each part by natural figure number and document order
    parts: Dict[int, List[MatchedFigure]] = {}
    for fig in matched_figures:
        parts.setdefault(fig.part, []).append(fig)

    duplicates: List[str] = []
    gaps: List[str] = []

    # Sort and resolve duplicates
    ordered_matched: List[MatchedFigure] = []
    seq_counter = 1

    for part_num in sorted(parts.keys()):
        part_figs = parts[part_num]

        # Natural sort by figure number, tie-breaking by document index
        def natural_fig_key(f: MatchedFigure):
            nums = re.findall(r"\d+", f.base_figure_number)
            sub_num = int(nums[1]) if len(nums) > 1 else 0
            return (sub_num, f.image.doc_index)

        part_figs.sort(key=natural_fig_key)

        # Detect duplicates and assign _a, _b suffixes
        num_counts: Dict[str, List[MatchedFigure]] = {}
        for f in part_figs:
            num_counts.setdefault(f.base_figure_number, []).append(f)

        for base_num, group in num_counts.items():
            if len(group) > 1:
                duplicates.append(
                    f"Part {part_num}: Duplicate figure number {base_num} found ({len(group)} occurrences)"
                )
                letters = "abcdefghijklmnopqrstuvwxyz"
                for idx, item in enumerate(group):
                    suffix = f"_{letters[idx]}" if idx < len(letters) else f"_{idx+1}"
                    item.figure_number = f"{base_num}{suffix}"
                    item.warnings.append(
                        f"Duplicate figure number {base_num} (assigned suffix {suffix})"
                    )

        # Detect gaps in expected integer sequence
        seen_numbers: List[int] = []
        for f in part_figs:
            nums = re.findall(r"\d+", f.base_figure_number)
            if len(nums) > 1:
                seen_numbers.append(int(nums[1]))

        if seen_numbers:
            min_val = min(seen_numbers)
            max_val = max(seen_numbers)
            missing = set(range(min_val, max_val + 1)) - set(seen_numbers)
            for m_val in sorted(missing):
                gaps.append(f"Part {part_num}: Missing Fig {part_num}.{m_val} in sequence")

        # Assign output path and sequence
        for f in part_figs:
            f.sequence = seq_counter
            seq_counter += 1
            f.output_path = (
                f"Part_{part_num}/Fig_{f.figure_number}__{f.image.media_filename}"
            )
            ordered_matched.append(f)

    # 6. Build Manifest Rows
    manifest_rows: List[Dict[str, Any]] = []
    for f in ordered_matched:
        manifest_rows.append(
            {
                "sequence": f.sequence,
                "part": f.part,
                "figure_number": f.figure_number,
                "caption_text": f.caption_text,
                "original_media_file": f.image.media_path,
                "output_path": f.output_path,
                "extension": f.image.extension,
                "pixel_width": f.image.pixel_width or "",
                "pixel_height": f.image.pixel_height or "",
                "placed_width_in": f.image.placed_width_in,
                "effective_dpi": f.effective_dpi or "",
                "print_status": f.print_status,
                "sha256": f.image.sha256,
                "warnings": "; ".join(f.warnings),
            }
        )

    for idx, un_img in enumerate(unmatched_images, 1):
        un_path = f"_unmatched/Unmatched_{un_img.doc_index:03d}__{un_img.media_filename}"
        status = "OK (Vector)" if un_img.is_vector else (
            "OK" if un_img.effective_dpi and un_img.effective_dpi >= 300 else (
                "Re-export to 300 DPI" if un_img.effective_dpi and un_img.effective_dpi >= 200 else "Will print soft"
            )
        )
        manifest_rows.append(
            {
                "sequence": f"U{idx:03d}",
                "part": "-",
                "figure_number": "-",
                "caption_text": "",
                "original_media_file": un_img.media_path,
                "output_path": un_path,
                "extension": un_img.extension,
                "pixel_width": un_img.pixel_width or "",
                "pixel_height": un_img.pixel_height or "",
                "placed_width_in": un_img.placed_width_in,
                "effective_dpi": un_img.effective_dpi or "",
                "print_status": status,
                "sha256": un_img.sha256,
                "warnings": "No caption matched within search window",
            }
        )

    # 7. Generate report.txt content
    report_lines = [
        "DOCX FIGURE EXTRACTION REPORT",
        "=======================================================",
        f"Total Images in Document: {len(all_images)}",
        f"Matched Figures: {len(ordered_matched)}",
        f"Unmatched Images: {len(unmatched_images)}",
        f"Unmatched Captions: {len(unmatched_captions)}",
        "",
        "PART BREAKDOWN",
        "-------------------------------------------------------",
    ]
    if parts:
        for p_k in sorted(parts.keys()):
            report_lines.append(f"Part {p_k}: {len(parts[p_k])} figures")
    else:
        report_lines.append("No parts detected.")

    report_lines.extend(["", "DUPLICATES & SUFFIXES", "-------------------------------------------------------"])
    if duplicates:
        report_lines.extend(f"- {d}" for d in duplicates)
    else:
        report_lines.append("None detected.")

    report_lines.extend(["", "NUMBERING GAPS", "-------------------------------------------------------"])
    if gaps:
        report_lines.extend(f"- {g}" for g in gaps)
    else:
        report_lines.append("No gaps in numeric figure sequences.")

    report_lines.extend(["", "UNMATCHED IMAGES", "-------------------------------------------------------"])
    if unmatched_images:
        for u in unmatched_images:
            report_lines.append(
                f"- {u.media_filename} (Doc Index {u.doc_index}, Placed width: {u.placed_width_in} in) -> saved to _unmatched/"
            )
    else:
        report_lines.append("All images were successfully matched to captions.")

    report_lines.extend(["", "UNMATCHED CAPTIONS (Captions without images)", "-------------------------------------------------------"])
    if unmatched_captions:
        for c in unmatched_captions:
            report_lines.append(f"- \"{c.full_text}\" (Paragraph {c.paragraph_idx})")
    else:
        report_lines.append("All captions were successfully matched to images.")

    report_lines.extend(["", "IMAGE WARNINGS & NOTICES", "-------------------------------------------------------"])
    warning_count = 0
    for f in ordered_matched:
        if f.warnings:
            warning_count += 1
            report_lines.append(f"- Fig {f.figure_number} ({f.image.media_filename}): {', '.join(f.warnings)}")
    if warning_count == 0:
        report_lines.append("No warnings.")

    report_text = "\n".join(report_lines)

    all_general_warnings: List[str] = []
    if duplicates:
        all_general_warnings.append(f"{len(duplicates)} duplicate figure numbers resolved with suffixes")
    if gaps:
        all_general_warnings.append(f"{len(gaps)} gaps detected in figure numbering")
    if unmatched_images:
        all_general_warnings.append(f"{len(unmatched_images)} images had no captions (moved to _unmatched/)")
    if unmatched_captions:
        all_general_warnings.append(f"{len(unmatched_captions)} captions had no accompanying image")

    return ParseResult(
        matched_figures=ordered_matched,
        unmatched_images=unmatched_images,
        unmatched_captions=unmatched_captions,
        parts=parts,
        duplicates=duplicates,
        gaps=gaps,
        warnings=all_general_warnings,
        report_text=report_text,
        manifest_rows=manifest_rows,
    )


def build_extracted_figures_zip(
    docx_file_or_bytes: Any,
    caption_position: str = "auto",
    root_folder_name: str = "extracted_figures",
) -> Tuple[bytes, ParseResult]:
    """Extract figures with original byte preservation and build an uncompressed ZIP.

    Returns:
        (zip_bytes, parse_result)
    """
    if isinstance(docx_file_or_bytes, (bytes, bytearray)):
        docx_zip = zipfile.ZipFile(io.BytesIO(docx_file_or_bytes))
    elif isinstance(docx_file_or_bytes, str):
        docx_zip = zipfile.ZipFile(docx_file_or_bytes, "r")
    else:
        docx_zip = zipfile.ZipFile(docx_file_or_bytes)

    result = parse_docx_figures(docx_file_or_bytes, caption_position=caption_position)

    zip_buf = io.BytesIO()

    # Write output ZIP strictly with zipfile.ZIP_STORED (no recompression or pass over bytes)
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_STORED) as out_zip:
        # Write matched figures
        for fig in result.matched_figures:
            entry_path = f"{root_folder_name}/{fig.output_path}"
            # Direct byte stream copy: zero Pillow modification or recompression
            raw_bytes = docx_zip.read(fig.image.media_path)
            out_zip.writestr(entry_path, raw_bytes, compress_type=zipfile.ZIP_STORED)

        # Write unmatched images
        for un_img in result.unmatched_images:
            un_name = f"Unmatched_{un_img.doc_index:03d}__{un_img.media_filename}"
            entry_path = f"{root_folder_name}/_unmatched/{un_name}"
            raw_bytes = docx_zip.read(un_img.media_path)
            out_zip.writestr(entry_path, raw_bytes, compress_type=zipfile.ZIP_STORED)

        # Write manifest.csv
        csv_buf = io.StringIO()
        fieldnames = [
            "sequence",
            "part",
            "figure_number",
            "caption_text",
            "original_media_file",
            "output_path",
            "extension",
            "pixel_width",
            "pixel_height",
            "placed_width_in",
            "effective_dpi",
            "print_status",
            "sha256",
            "warnings",
        ]
        writer = csv.DictWriter(csv_buf, fieldnames=fieldnames)
        writer.writeheader()
        for row in result.manifest_rows:
            writer.writerow(row)

        out_zip.writestr(
            f"{root_folder_name}/manifest.csv",
            csv_buf.getvalue().encode("utf-8-sig"),
            compress_type=zipfile.ZIP_STORED,
        )

        # Write report.txt
        out_zip.writestr(
            f"{root_folder_name}/report.txt",
            result.report_text.encode("utf-8"),
            compress_type=zipfile.ZIP_STORED,
        )

    return zip_buf.getvalue(), result

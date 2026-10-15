"""
Tests for DOCX Figure Extraction and Caption Matching.
"""

import hashlib
import io
import sys
import zipfile
from pathlib import Path

import docx
import pytest
from docx.shared import Inches
from PIL import Image

# Ensure project root is on sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx_figures import build_extracted_figures_zip, parse_docx_figures


def create_solid_image(width: int, height: int, color: tuple) -> bytes:
    """Generate in-memory PNG image bytes with distinct colors."""
    img = Image.new("RGB", (width, height), color)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def test_parts_1_to_3_ordered_by_figure_number():
    """Parts 1-3 with 2-3 figures each; the extracted order matches figure numbers

    even when the document order is shuffled by captions.
    """
    doc = docx.Document()

    # Create dummy images with distinct colors
    img1 = create_solid_image(100, 100, (255, 0, 0))
    img2 = create_solid_image(100, 100, (0, 255, 0))
    img3 = create_solid_image(100, 100, (0, 0, 255))
    img4 = create_solid_image(100, 100, (255, 255, 0))
    img5 = create_solid_image(100, 100, (255, 0, 255))
    img6 = create_solid_image(100, 100, (0, 255, 255))
    img7 = create_solid_image(100, 100, (128, 128, 128))

    # Intentionally shuffle: Part 1 Fig 1.2 placed BEFORE Fig 1.1
    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(img1), width=Inches(4.0))
    doc.add_paragraph("Figure 1.2: Second figure of part 1")

    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(img2), width=Inches(4.0))
    doc.add_paragraph("Figure 1.1: First figure of part 1")

    # Part 2: Fig 2.1, 2.2
    p3 = doc.add_paragraph()
    p3.add_run().add_picture(io.BytesIO(img3), width=Inches(3.5))
    doc.add_paragraph("Fig. 2.1 - Component layout")

    p4 = doc.add_paragraph()
    p4.add_run().add_picture(io.BytesIO(img4), width=Inches(3.5))
    doc.add_paragraph("Fig 2.2: Data pipeline")

    # Part 3: Fig 3.3, 3.1, 3.2 (shuffled)
    p5 = doc.add_paragraph()
    p5.add_run().add_picture(io.BytesIO(img5), width=Inches(4.5))
    doc.add_paragraph("Figure 3.3: Analysis chart")

    p6 = doc.add_paragraph()
    p6.add_run().add_picture(io.BytesIO(img6), width=Inches(4.5))
    doc.add_paragraph("Figure 3.1: Overview map")

    p7 = doc.add_paragraph()
    p7.add_run().add_picture(io.BytesIO(img7), width=Inches(4.5))
    doc.add_paragraph("Figure 3.2: Network topology")

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    result = parse_docx_figures(docx_bytes)

    # Check Part 1 order: 1.1 then 1.2
    part1_figs = [f.figure_number for f in result.parts[1]]
    assert part1_figs == ["1.1", "1.2"]

    # Check Part 2 order: 2.1 then 2.2
    part2_figs = [f.figure_number for f in result.parts[2]]
    assert part2_figs == ["2.1", "2.2"]

    # Check Part 3 order: 3.1, 3.2, 3.3 (sorted naturally despite document shuffle)
    part3_figs = [f.figure_number for f in result.parts[3]]
    assert part3_figs == ["3.1", "3.2", "3.3"]


def test_caption_below_vs_above():
    """Verify that captions placed above or below the image are correctly matched."""
    # Test 1: Caption ABOVE image
    doc_above = docx.Document()
    doc_above.add_paragraph("Figure 1.1: Caption above image")
    p_img = doc_above.add_paragraph()
    p_img.add_run().add_picture(io.BytesIO(create_solid_image(80, 80, (10, 20, 30))))

    buf_above = io.BytesIO()
    doc_above.save(buf_above)
    res_above = parse_docx_figures(buf_above.getvalue(), caption_position="auto")
    assert len(res_above.matched_figures) == 1
    assert res_above.matched_figures[0].figure_number == "1.1"

    # Test 2: Caption BELOW image
    doc_below = docx.Document()
    p_img2 = doc_below.add_paragraph()
    p_img2.add_run().add_picture(io.BytesIO(create_solid_image(80, 80, (40, 50, 60))))
    doc_below.add_paragraph("Figure 1.2: Caption below image")

    buf_below = io.BytesIO()
    doc_below.save(buf_below)
    res_below = parse_docx_figures(buf_below.getvalue(), caption_position="auto")
    assert len(res_below.matched_figures) == 1
    assert res_below.matched_figures[0].figure_number == "1.2"


def test_duplicate_figure_numbers_produce_suffixes_and_report_flag():
    """A duplicate figure number produces _a and _b files plus a report flag."""
    doc = docx.Document()

    # Image 1 with Figure 5.11
    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(create_solid_image(120, 80, (100, 10, 10))))
    doc.add_paragraph("Figure 5.11: Primary diagram")

    # Image 2 also captioned Figure 5.11 (duplicate!)
    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(create_solid_image(120, 80, (10, 100, 10))))
    doc.add_paragraph("Figure 5.11: Detailed inset")

    buf = io.BytesIO()
    doc.save(buf)

    result = parse_docx_figures(buf.getvalue())
    assert len(result.parts[5]) == 2
    f1, f2 = result.parts[5]

    assert f1.figure_number == "5.11_a"
    assert f2.figure_number == "5.11_b"
    assert "Fig_5.11_a__" in f1.output_path
    assert "Fig_5.11_b__" in f2.output_path

    # Verify duplicate is logged in report
    assert any("5.11" in d for d in result.duplicates)
    assert "DUPLICATE" in result.report_text
    assert "5.11" in result.report_text


def test_seq_field_caption_split_across_runs():
    """A SEQ-field caption split across multiple runs matches correctly."""
    doc = docx.Document()

    p_img = doc.add_paragraph()
    p_img.add_run().add_picture(io.BytesIO(create_solid_image(90, 90, (30, 40, 50))))

    # Create a caption paragraph where runs are split
    p_cap = doc.add_paragraph()
    r1 = p_cap.add_run("Figure ")
    r2 = p_cap.add_run("4.3")  # Represents field-generated number
    r3 = p_cap.add_run(": Architecture flow")

    buf = io.BytesIO()
    doc.save(buf)

    result = parse_docx_figures(buf.getvalue())
    assert len(result.matched_figures) == 1
    assert result.matched_figures[0].figure_number == "4.3"
    assert "Architecture flow" in result.matched_figures[0].caption_text


def test_unmatched_image_lands_in_unmatched():
    """An image with no caption lands in _unmatched."""
    doc = docx.Document()

    # Image with no caption anywhere near it
    p_img = doc.add_paragraph()
    p_img.add_run().add_picture(io.BytesIO(create_solid_image(50, 50, (200, 200, 200))))

    doc.add_paragraph("Just standard paragraph body text with no figure numbering at all.")
    doc.add_paragraph("Another standard paragraph describing topic details.")
    doc.add_paragraph("Yet another generic paragraph.")
    doc.add_paragraph("Fourth generic paragraph.")

    buf = io.BytesIO()
    doc.save(buf)

    result = parse_docx_figures(buf.getvalue())
    assert len(result.matched_figures) == 0
    assert len(result.unmatched_images) == 1
    assert result.unmatched_images[0].doc_index == 1

    # In manifest rows, it should be marked as unmatched
    assert any(row["output_path"].startswith("_unmatched/") for row in result.manifest_rows)


def test_sha256_unmodified_byte_identity():
    """SHA-256 of every extracted file equals the original media entry in the docx."""
    doc = docx.Document()

    img1_bytes = create_solid_image(150, 100, (12, 34, 56))
    img2_bytes = create_solid_image(200, 150, (78, 90, 12))

    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(img1_bytes))
    doc.add_paragraph("Figure 1.1: First test image")

    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(img2_bytes))
    doc.add_paragraph("Figure 1.2: Second test image")

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    # Read original SHA256 directly from the source docx zip
    original_hashes = {}
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as src_zip:
        for name in src_zip.namelist():
            if name.startswith("word/media/"):
                original_hashes[name] = hashlib.sha256(src_zip.read(name)).hexdigest()

    assert len(original_hashes) == 2

    # Build the output ZIP
    out_zip_bytes, result = build_extracted_figures_zip(
        docx_bytes, root_folder_name="test_figures"
    )

    # Read extracted files and check SHA-256
    with zipfile.ZipFile(io.BytesIO(out_zip_bytes), "r") as out_zip:
        for fig in result.matched_figures:
            zip_entry = f"test_figures/{fig.output_path}"
            extracted_bytes = out_zip.read(zip_entry)
            extracted_sha = hashlib.sha256(extracted_bytes).hexdigest()

            # The extracted file's SHA-256 must match the original media file exactly
            orig_sha = original_hashes[fig.image.media_path]
            assert extracted_sha == orig_sha, f"SHA mismatch for {fig.output_path}"


def test_figure_in_table_cell_found_in_right_order():
    """A figure in a table cell is found in the right order."""
    doc = docx.Document()

    doc.add_paragraph("Document introduction")

    # Table with 2 rows, 1 column: image in top row, caption in bottom row
    tbl = doc.add_table(rows=2, cols=1)
    cell_top = tbl.cell(0, 0)
    p_img = cell_top.paragraphs[0]
    p_img.add_run().add_picture(io.BytesIO(create_solid_image(110, 90, (80, 70, 60))))

    cell_bottom = tbl.cell(1, 0)
    cell_bottom.paragraphs[0].text = "Figure 2.1: Table cell figure caption"

    # Subsequent image outside table
    p_after = doc.add_paragraph()
    p_after.add_run().add_picture(io.BytesIO(create_solid_image(90, 90, (10, 20, 30))))
    doc.add_paragraph("Figure 2.2: Figure outside table")

    buf = io.BytesIO()
    doc.save(buf)

    result = parse_docx_figures(buf.getvalue())
    assert len(result.parts[2]) == 2
    assert result.parts[2][0].figure_number == "2.1"
    assert result.parts[2][0].image.location == "table"
    assert result.parts[2][1].figure_number == "2.2"
    assert result.parts[2][1].image.location == "body"


def test_docx_api_endpoints_analyze_and_extract():
    """Verify HTTP endpoints /api/docx/analyze and /api/docx/extract on Windows."""
    from app import app
    from fastapi.testclient import TestClient

    client = TestClient(app)

    doc = docx.Document()
    p = doc.add_paragraph()
    p.add_run().add_picture(io.BytesIO(create_solid_image(120, 80, (20, 40, 60))))
    doc.add_paragraph("Figure 1.1: Test API figure")

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    # 1. Test /api/docx/analyze
    resp_analyze = client.post(
        "/api/docx/analyze",
        files={"file": ("book_test.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"caption_position": "auto"},
    )
    assert resp_analyze.status_code == 200, resp_analyze.text
    data = resp_analyze.json()
    assert data["summary"]["matched"] == 1
    assert "1" in data["parts"]

    # 2. Test /api/docx/extract
    resp_extract = client.post(
        "/api/docx/extract",
        files={"file": ("book_test.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"caption_position": "auto"},
    )
    assert resp_extract.status_code == 200, resp_extract.text
    assert "attachment" in resp_extract.headers["Content-Disposition"]
    assert "book_test_figures.zip" in resp_extract.headers["Content-Disposition"]

    # Verify extracted zip content
    with zipfile.ZipFile(io.BytesIO(resp_extract.content)) as z:
        namelist = z.namelist()
        assert any(n.endswith("manifest.csv") for n in namelist)
        assert any(n.endswith("report.txt") for n in namelist)
        assert any("Fig_1.1__" in n for n in namelist)


def test_docx_page_number_tracking():
    """Verify that images across hard page breaks, soft breaks, and section breaks have accurate page numbers."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from app import app
    from fastapi.testclient import TestClient

    doc = docx.Document()

    # Page 1: Figure 1.1
    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(create_solid_image(100, 100, (255, 0, 0))))
    doc.add_paragraph("Figure 1.1: Image on page 1")

    # Hard page break
    doc.add_page_break()

    # Page 2: Figure 1.2
    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(create_solid_image(100, 100, (0, 255, 0))))
    doc.add_paragraph("Figure 1.2: Image on page 2")

    # Soft page break (w:lastRenderedPageBreak)
    p3 = doc.add_paragraph()
    run = p3.add_run()
    soft_break = OxmlElement("w:lastRenderedPageBreak")
    run._r.append(soft_break)

    # Page 3: Figure 1.3
    p4 = doc.add_paragraph()
    p4.add_run().add_picture(io.BytesIO(create_solid_image(100, 100, (0, 0, 255))))
    doc.add_paragraph("Figure 1.3: Image on page 3")

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    # 1. Check parse_docx_figures result
    res = parse_docx_figures(docx_bytes)
    figs = res.parts[1]
    assert len(figs) == 3
    assert figs[0].figure_number == "1.1"
    assert figs[0].page_number == 1
    assert figs[1].figure_number == "1.2"
    assert figs[1].page_number == 2
    assert figs[2].figure_number == "1.3"
    assert figs[2].page_number == 3

    # 2. Check build_extracted_figures_zip manifest.csv
    zip_bytes, _ = build_extracted_figures_zip(docx_bytes)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        manifest_text = z.read("extracted_figures/manifest.csv").decode("utf-8-sig")
        assert "page_number" in manifest_text
        lines = manifest_text.strip().split("\r\n")
        header = lines[0].split(",")
        page_idx = header.index("page_number")
        assert lines[1].split(",")[page_idx] == "1"
        assert lines[2].split(",")[page_idx] == "2"
        assert lines[3].split(",")[page_idx] == "3"

    # 3. Check /api/docx/analyze API endpoint
    client = TestClient(app)
    resp = client.post(
        "/api/docx/analyze",
        files={"file": ("pages_test.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"caption_position": "auto"},
    )
    assert resp.status_code == 200
    api_figs = resp.json()["parts"]["1"]
    assert api_figs[0]["page_number"] == 1
    assert api_figs[1]["page_number"] == 2
    assert api_figs[2]["page_number"] == 3


def test_docx_enhance_figure_direct_and_zip_inclusion():
    """Verify that figures enhanced directly in DOCX tab are included in the downloaded ZIP

    with updated DPI, pixels, and OK status, while untouched images preserve original bytes.
    """
    import json
    from app import app
    from fastapi.testclient import TestClient

    doc = docx.Document()

    # Image 1: 100x100 at 4.88" -> effective DPI = 20 (low DPI)
    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(create_solid_image(100, 100, (200, 50, 50))), width=Inches(4.88))
    doc.add_paragraph("Figure 1.1: Low resolution diagram")

    # Image 2: 1500x1500 at 4.88" -> effective DPI = 307 (high DPI, OK)
    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(create_solid_image(1500, 1500, (50, 200, 50))), width=Inches(4.88))
    doc.add_paragraph("Figure 1.2: High resolution chart")

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    client = TestClient(app)

    # 1. Analyze DOCX
    resp_an = client.post(
        "/api/docx/analyze",
        files={"file": ("test_doc.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"caption_position": "auto"},
    )
    assert resp_an.status_code == 200
    figs = resp_an.json()["parts"]["1"]
    fig1 = figs[0]
    assert fig1["figure_number"] == "1.1"
    assert fig1["effective_dpi"] < 300
    assert fig1["print_status"] == "Will print soft"

    fig2 = figs[1]
    assert fig2["figure_number"] == "1.2"
    assert fig2["effective_dpi"] >= 300
    assert fig2["print_status"] == "OK"

    # 2. Call /api/docx/enhance-figure directly for Figure 1.1
    resp_enh = client.post(
        "/api/docx/enhance-figure",
        files={"file": ("test_doc.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={
            "media_path": fig1["original_media_file"],
            "placed_width_in": fig1["placed_width_in"],
            "target_dpi": 300,
            "method": "lanczos",
        },
    )
    assert resp_enh.status_code == 200
    assert resp_enh.headers["X-Effective-DPI"] == "300"
    enhanced_bytes = resp_enh.content
    enhanced_w = int(resp_enh.headers["X-Enhanced-Width"])
    enhanced_h = int(resp_enh.headers["X-Enhanced-Height"])
    assert enhanced_w >= 1464

    # 3. Request ZIP from /api/docx/extract with enhanced_files attached
    meta_json = json.dumps([{
        "media_path": fig1["original_media_file"],
        "pixel_width": enhanced_w,
        "pixel_height": enhanced_h,
        "effective_dpi": 300,
    }])

    resp_zip = client.post(
        "/api/docx/extract",
        files=[
            ("file", ("test_doc.docx", docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
            ("enhanced_files", (fig1["original_media_file"], enhanced_bytes, "image/png")),
        ],
        data={
            "caption_position": "auto",
            "enhanced_meta": meta_json,
        },
    )
    assert resp_zip.status_code == 200

    # 4. Inspect the generated ZIP:
    with zipfile.ZipFile(io.BytesIO(resp_zip.content)) as out_zip:
        # Figure 1.1 must contain the enhanced image bytes!
        fig1_entry = [n for n in out_zip.namelist() if "Fig_1.1__" in n][0]
        fig1_zip_bytes = out_zip.read(fig1_entry)
        assert fig1_zip_bytes == enhanced_bytes
        assert hashlib.sha256(fig1_zip_bytes).hexdigest() == hashlib.sha256(enhanced_bytes).hexdigest()

        # Figure 1.2 (unenhanced) must contain the exact original bytes from the docx!
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as in_zip:
            orig_fig2_bytes = in_zip.read(fig2["original_media_file"])

        fig2_entry = [n for n in out_zip.namelist() if "Fig_1.2__" in n][0]
        fig2_zip_bytes = out_zip.read(fig2_entry)
        assert fig2_zip_bytes == orig_fig2_bytes
        assert hashlib.sha256(fig2_zip_bytes).hexdigest() == hashlib.sha256(orig_fig2_bytes).hexdigest()

        # Check manifest.csv:
        # Fig 1.1 row must show 300 DPI, OK status, and enhanced dimensions!
        manifest_csv = out_zip.read("test_doc_figures/manifest.csv").decode("utf-8-sig")
        lines = manifest_csv.strip().split("\r\n")
        header = lines[0].split(",")
        dpi_idx = header.index("effective_dpi")
        status_idx = header.index("print_status")
        width_idx = header.index("pixel_width")

        row1 = lines[1].split(",")
        assert row1[dpi_idx] == "300"
        assert row1[status_idx] == "OK"
        assert int(row1[width_idx]) == enhanced_w

        # Check report.txt mentions enhanced figures:
        report_txt = out_zip.read("test_doc_figures/report.txt").decode("utf-8")
        assert "ENHANCED FIGURES IN THIS ARCHIVE" in report_txt
        assert "Fig 1.1" in report_txt



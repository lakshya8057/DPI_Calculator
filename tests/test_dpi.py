"""
Unit and integration tests for DPI Calculator & Image Enhancer.
"""

import io
import math
import sys
from pathlib import Path

import pytest
from PIL import Image

# Ensure project root is in sys.path when running `pytest` directly
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import (
    app,
    calculate_effective_dpi,
    calculate_target_dimensions,
    process_image,
    sanitize_input_image,
)
from fastapi.testclient import TestClient

client = TestClient(app)


def create_test_image(
    width: int, height: int, mode: str = "RGB", color: tuple = (100, 150, 200)
) -> bytes:
    """Helper to generate in-memory PNG image bytes."""
    img = Image.new(mode, (width, height), color)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def test_effective_dpi_1024_at_4_88():
    """1024 px at 4.88" -> 210 DPI (round to nearest integer)."""
    dpi = calculate_effective_dpi(1024, 4.88)
    assert dpi == 210
    # Exactly round(1024 / 4.88) = round(209.8360655737705) = 210


def test_effective_dpi_468_at_6_5():
    """468 px at 6.5" -> 72 DPI."""
    dpi = calculate_effective_dpi(468, 6.5)
    assert dpi == 72
    # Exactly round(468 / 6.5) = round(72.0) = 72


def test_effective_dpi_2048_at_4_88_no_resize():
    """2048 px at 4.88" -> 420 DPI, with no resize needed."""
    dpi = calculate_effective_dpi(2048, 4.88)
    assert dpi == 420

    target_w, target_h, needs_resize = calculate_target_dimensions(
        width=2048, height=1000, placed_width_in=4.88, target_dpi=300
    )
    # Required width = ceil(4.88 * 300) = 1464 px
    assert target_w == 2048
    assert target_h == 1000
    assert needs_resize is False


def test_output_width_exact_ceil_from_1024():
    """An output width of exactly ceil(4.88 x 300) = 1464 px when starting from 1024 px."""
    img_bytes = create_test_image(1024, 768, mode="RGB")
    out_bytes, meta = process_image(
        image_bytes=img_bytes,
        placed_width_in=4.88,
        target_dpi=300,
        method="lanczos",
        output_format="png",
    )

    expected_width = math.ceil(4.88 * 300)  # 1464
    assert meta["width"] == 1464
    assert expected_width == 1464

    # Verify Pillow reads the enhanced image width
    with Image.open(io.BytesIO(out_bytes)) as out_img:
        assert out_img.width == 1464
        # Aspect ratio check: 1464 * (768/1024) = 1098
        assert out_img.height == 1098


def test_rgba_png_keeps_alpha():
    """An RGBA PNG keeps its alpha after enhancement."""
    # Create an RGBA image with varying alpha
    img = Image.new("RGBA", (100, 100), (255, 0, 0, 128))
    # Add a fully transparent corner
    for x in range(20):
        for y in range(20):
            img.putpixel((x, y), (0, 0, 0, 0))

    buf = io.BytesIO()
    img.save(buf, format="PNG")

    out_bytes, meta = process_image(
        image_bytes=buf.getvalue(),
        placed_width_in=2.0,  # target_width = 600 px (needs resize from 100 px)
        target_dpi=300,
        method="lanczos",
        output_format="png",
    )

    with Image.open(io.BytesIO(out_bytes)) as out_img:
        assert out_img.mode == "RGBA"
        # Check that alpha channel exists and is preserved
        assert "A" in out_img.getbands()
        alpha = out_img.split()[3]
        # Transparent corner should remain transparent (alpha == 0)
        assert alpha.getpixel((0, 0)) == 0


def test_dpi_metadata_reads_back_300_300():
    """The output file's DPI metadata reads back as (300, 300)."""
    img_bytes = create_test_image(500, 500, mode="RGB")
    out_bytes, _ = process_image(
        image_bytes=img_bytes,
        placed_width_in=3.0,
        target_dpi=300,
        method="lanczos",
        output_format="png",
    )

    with Image.open(io.BytesIO(out_bytes)) as out_img:
        raw_dpi = out_img.info.get("dpi")
        assert raw_dpi is not None
        # In PNG pHYs chunks, 300 DPI is encoded in dots/meter (~11811 dpm)
        # Reading back gives (299.9994, 299.9994), which rounds to (300, 300).
        rounded_dpi = (round(raw_dpi[0]), round(raw_dpi[1]))
        assert rounded_dpi == (300, 300)


def test_hard_limit_exceeding_12000px():
    """Hard limit: reject outputs above 12000 px on either side."""
    img_bytes = create_test_image(200, 200, mode="RGB")
    # 50 inches at 300 DPI = 15000 px > 12000 px
    with pytest.raises(Exception) as exc_info:
        process_image(
            image_bytes=img_bytes,
            placed_width_in=50.0,
            target_dpi=300,
            method="lanczos",
            output_format="png",
        )
    assert "exceed" in str(exc_info.value).lower()


def test_jpeg_transparency_rejection():
    """Never output JPEG if the source has transparency."""
    rgba_bytes = create_test_image(100, 100, mode="RGBA", color=(255, 0, 0, 128))
    with pytest.raises(Exception) as exc_info:
        process_image(
            image_bytes=rgba_bytes,
            placed_width_in=2.0,
            target_dpi=300,
            method="lanczos",
            output_format="jpeg",
        )
    assert "transparency" in str(exc_info.value).lower()


def test_api_analyze_and_enhance_endpoints():
    """Test full HTTP endpoints for /api/analyze and /api/enhance."""
    img_bytes = create_test_image(1024, 768, mode="RGB")

    # Test /api/analyze
    analyze_resp = client.post(
        "/api/analyze",
        files={"file": ("chart.png", img_bytes, "image/png")},
    )
    assert analyze_resp.status_code == 200
    data = analyze_resp.json()
    assert data["width"] == 1024
    assert data["height"] == 768
    assert data["format"] == "PNG"

    # Test /api/enhance
    enhance_resp = client.post(
        "/api/enhance",
        files={"file": ("chart.png", img_bytes, "image/png")},
        data={
            "placed_width_in": "4.88",
            "target_dpi": "300",
            "method": "lanczos",
            "output_format": "png",
        },
    )
    assert enhance_resp.status_code == 200
    assert enhance_resp.headers["X-Enhanced-Width"] == "1464"
    assert enhance_resp.headers["X-Effective-DPI"] == "300"
    assert "attachment" in enhance_resp.headers["Content-Disposition"]

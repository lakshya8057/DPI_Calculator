# Print DPI Calculator & Image Enhancer

A specialized web application for print and book production. It measures an image's effective print DPI based on its placed dimensions in a book, enhances diagrams and figures to a target DPI (300 DPI standard), embeds print DPI metadata into exported files, and extracts original, uncompressed figure images from Word `.docx` book manuscripts.

---

## Features

### 1. Single Image DPI Calculator & Enhancer
- **Effective Print DPI Calculation**: Computes `round(pixel_width / placed_width_in)` in real time.
- **Book Preset Buttons**: Quick presets for standard placed widths (`4.88"`, `4.37"`, `3.57"`, `2.56"`, `1.32"`).
- **Status Badges**:
  - `≥ 300 DPI`: Green **OK**
  - `200–299 DPI`: Amber **Re-export to 300 DPI**
  - `< 200 DPI`: Red **Will print soft**
- **Enhancement Methods**:
  - **Lanczos + light unsharp mask** (default): Sharpens diagram text without haloing or ringing.
  - **Transparency preservation**: Separates and resizes the alpha channel independently to prevent dark edge fringing.
  - **AI Upscaling (Real-ESRGAN)**: Optional ncnn-vulkan integration enabled automatically when detected on `PATH`.
- **Side-by-Side Visual Inspection**:
  - Fit, 100%, and 200% zoom controls for fine diagram and text inspection.
  - Before vs. after pixel dimensions and effective DPI comparison.
- **Honest Warnings**:
  - Alerts for upscale factors > 2× and > 4× reminding users to re-export vector sources if available.
  - Heuristic detection for diagrams/charts with warnings regarding AI text distortion.
  - Clarification that embedded DPI metadata sets physical print scale and does not invent pixel detail.

### 2. Batch Mode Processing
- Upload multiple figures at once.
- Optional CSV import to assign custom placed widths per file (`filename,placed_width_in`).
- Downloads a single ZIP archive containing all enhanced images and a summary `report.csv`.

### 3. DOCX Figure Extractor (Direct Manuscript Mining)
- **Zero Loss / Original Quality Guarantee**:
  - Direct byte extraction from `word/media/` inside the `.docx` archive.
  - **No re-encoding, no conversion, no resizing, and no recompression**. Pillow is only used in read-only mode to report dimensions.
  - Vector files (EMF, WMF, SVG, EPS) are extracted in their original format.
  - Output ZIP is written with `zipfile.ZIP_STORED` (no second pass or compression over bytes).
  - SHA-256 hashes of extracted files match the original DOCX media entries bit-for-bit.
- **Intelligent Caption Matching**:
  - Tolerant regex supporting "Figure 1.1", "Fig. 1.1 - ...", "Figure 1.1. ...", and SEQ fields split across runs.
  - Search order: checks adjacent paragraphs above/below (auto, below, or above preference).
  - Handles figures inside tables, text boxes, and content controls.
- **Organized Output ZIP Structure**:
  ```
  <docx name>_figures/
    Part_1/
      Fig_1.1__image37.png
      Fig_1.2__image68.png
      ...
    Part_2/
      Fig_2.1__image86.png
    ...
    Part_6/
    _unmatched/          (images with no caption found)
    manifest.csv
    report.txt
  ```
- **Duplicate & Sequence Handling**:
  - Automatically resolves duplicate figure numbers by appending `_a` and `_b` suffixes.
  - Detects numbering gaps and lists them in `report.txt`.
- **One-Click Send to Enhancer**:
  - Clicking "Enhance" on any table row preloads that figure and its placed width into the Single Image tab.

---

## Project Structure

```
DPI Calculator/
├── app.py                  # FastAPI backend with DPI calculation, enhancement, and DOCX endpoints
├── docx_figures.py         # Standalone DOCX XML walker, caption matcher & ZIP packager
├── static/
│   └── index.html          # Single-page vanilla JavaScript frontend (light mode, calm colors)
├── tests/
│   ├── test_dpi.py         # Pytest suite for DPI calculations and enhancement
│   └── test_docx_figures.py# Pytest suite for DOCX figure extraction and SHA256 integrity
├── requirements.txt        # Python dependencies
└── README.md               # Documentation and instructions
```

---

## Installation & Setup

### 1. Prerequisites
- Python 3.11 or higher (Python 3.12 recommended)

### 2. Install Dependencies

Open PowerShell or terminal in the project directory:

```powershell
pip install -r requirements.txt
```

### 3. (Optional) Install Real-ESRGAN Vulkan Binary

If you wish to use the AI upscaler option:

1. Download the latest release of `realesrgan-ncnn-vulkan` for your OS:
   - **Windows / Linux / macOS**: [Real-ESRGAN Releases on GitHub](https://github.com/xinntao/Real-ESRGAN/releases)
2. Extract the archive (e.g. `realesrgan-ncnn-vulkan-20220424-windows.zip`).
3. Place `realesrgan-ncnn-vulkan.exe` and its accompanying `.bin` / `.param` model files into a directory included in your system's `PATH`.
4. Restart the web server. The app will automatically detect it and show "AI upscale (Real-ESRGAN)" in the dropdown.

---

## Running the Web App

Start the application with Uvicorn:

```powershell
uvicorn app:app --reload --port 8000
```

Then open your browser and navigate to:

```
http://127.0.0.1:8000
```

---

## Running the Tests

Execute the automated test suite with pytest:

```powershell
pytest -v
```

All 16 unit and integration tests will run:
- **`tests/test_dpi.py`**:
  - Effective DPI formula tests (1024 px @ 4.88" = 210 DPI, 468 px @ 6.5" = 72 DPI, 2048 px @ 4.88" = 420 DPI with no resize).
  - Exact ceil dimensioning (`ceil(4.88 * 300) = 1464 px`).
  - RGBA alpha preservation with no dark fringing.
  - DPI metadata round-trip verification (`(300, 300)`).
  - 12,000 px hard limit rejection and JPEG transparency guard.
  - Full HTTP endpoints (`/api/analyze` and `/api/enhance`).
- **`tests/test_docx_figures.py`**:
  - Parts 1-3 natural figure numbering order even when document order is shuffled.
  - Captions placed above vs below images.
  - Duplicate figure numbers producing `_a` and `_b` files plus report flags.
  - Word SEQ-field captions split across runs.
  - Unmatched images landing in `_unmatched/`.
  - Bitwise SHA-256 identity between extracted files and original `word/media/*` entries.
  - Figures inside table cells extracted in reading order.

---

## API Documentation

### Image Enhancer Endpoints
- `POST /api/analyze`: Multipart image upload (max 25 MB). Returns image dimensions, embedded DPI, and heuristics.
- `POST /api/enhance`: Multipart image upload + form fields (`placed_width_in`, `target_dpi`, `method`, `output_format`). Returns enhanced image bytes with print DPI metadata.
- `POST /api/batch-enhance`: Multi-file upload + optional CSV placed widths. Returns ZIP with enhanced images and `report.csv`.

### DOCX Extractor Endpoints
- `POST /api/docx/analyze`: Multipart `.docx` upload (max 150 MB) + `caption_position` (`auto`, `below`, `above`). Returns structured JSON analysis of all figures, captions, parts, thumbnails, and warnings.
- `POST /api/docx/extract`: Multipart `.docx` upload (max 150 MB) + `caption_position`. Returns `<docx name>_figures.zip` containing untouched original images, `manifest.csv`, and `report.txt`.
- `POST /api/docx/extract-single`: Multipart `.docx` upload + `media_path`. Returns the raw bytes of a single figure for seamless transfer to the Enhancer tab.

# PDF to Excel Extractor

Converts one-page scanned PDF files into a single Excel workbook using OCR and table extraction, served through a Streamlit UI.

Designed for scanned technical drawings, parts lists, and table-heavy documents where each PDF is a single-page image. Runs entirely on CPU with open-source Python tooling.

> **Deployment note:** This app requires PaddlePaddle and PaddleOCR (~1.5 GB of dependencies) and cannot be hosted on Streamlit Community Cloud. Use Docker on any Linux server, VPS, or local machine instead.

---

## What It Does

- Two input modes:
  - **Upload PDFs** — upload files directly in the browser (session-based)
  - **Read Folder** — reads all `.pdf` files from a mounted input folder (resumable)
- Extracts structured table data from each page using `PPStructure`
- Tries multiple image preprocessing variants automatically to maximize OCR yield
- Produces one Excel workbook with one sheet — each file gets:
  - row 1: headers
  - row 2: values
  - row 3: blank separator
- Text found outside tables is captured in `Unstructured_1`, `Unstructured_2`, … columns
- Shows live previews and progress during extraction
- Folder mode saves progress so a stopped run can be resumed

---

## Repository Layout

```
.
├── app.py
├── Dockerfile
├── docker-compose.yml
├── packages.txt
├── requirements.txt
└── .streamlit/
    └── config.toml
```

---

## Deployment — Docker (recommended)

### Requirements

- Docker and Docker Compose installed
- Any Linux server, VPS, or local machine with ~2 GB RAM

### Steps

```bash
git clone https://github.com/wael-fahmy/pdf-excel
cd pdf-excel
mkdir -p input_pdfs output_excel
docker compose build --no-cache
docker compose up -d
```

Open `http://localhost:8501` (or `http://<server-ip>:8501` for a remote server).

### Folder mode workflow

1. Copy your PDF files into `input_pdfs/`
2. Open the app in the browser
3. Select **Read Folder** in the sidebar
4. Keep input path as `/data/input` and output path as `/data/output`
5. Click **Start Extraction**
6. Collect `output_excel/Extracted_Data.xlsx` when done

### Upload mode workflow

1. Open the app in the browser
2. Keep **Upload PDFs** selected
3. Upload one or more PDF files
4. Click **Start Extraction**
5. Download `Extracted_Data.xlsx` from the browser

---

## Hosting Options

Any platform that supports Docker works:

| Platform | Notes |
|---|---|
| **VPS (Hetzner, DigitalOcean, etc.)** | `docker compose up -d`, ~$5/month |
| **Railway** | Connect GitHub repo, select Dockerfile |
| **Render** | Free tier supports Docker containers |
| **Fly.io** | `fly launch` from the repo root |

---

## Output Files

```
output_excel/
├── Extracted_Data.xlsx
└── .pdf_cache/
    ├── progress.json
    ├── records/*.json
    └── samples/*.png
```

The Excel file and resume state are saved automatically to disk after each successful file.

---

## Resume Behavior (Folder Mode)

If the app is stopped mid-run:

- Completed files remain in cache
- The Excel file stays on disk with all successful results so far
- The next run continues from the last saved point

Sidebar options:

| Option | Effect |
|---|---|
| **Resume previous run** | Skips files already marked done or empty |
| **Retry failed files** | Re-processes files that previously errored |
| **Clear Cache** | Deletes cached progress and starts fresh |

Turning off Resume clears the old cache and Excel file before starting.

---

## Expected Input

Best results:

- Scanned technical tables
- Engineering drawings with parts lists
- Machine-printed text inside table cells
- Single-page PDF images at 150 DPI or higher

Harder cases:

- Very faint or low-contrast scans
- Handwritten notes
- Heavily skewed or rotated pages
- Multi-page PDFs (only the first page is processed)

---

## OCR Details

Uses `paddleocr` with `PPStructure` on CPU. For each page the app tries four image variants and keeps the one that yields the most structured data:

1. Original image
2. CLAHE-enhanced denoised grayscale
3. Adaptive threshold (high contrast binary)
4. Sharpened version of the thresholded image

---

## Troubleshooting

**`ImportError: cannot import name 'PPStructure' from 'paddleocr'`**

Rebuild the image to reinstall pinned dependencies:

```bash
docker compose build --no-cache
docker compose up -d
```

**No PDFs found in folder mode**

Make sure files are inside `input_pdfs/` and have a `.pdf` extension.

**Extraction quality is poor**

- Use higher-resolution scans (200–300 DPI)
- Crop unnecessary whitespace or borders before input
- Test with a small batch (`File limit = 5`) first

**Output file not updating**

Confirm the `output_excel/` folder exists and Docker has write access to it.

---

## Notes for Non-Technical Operators

Normal Docker workflow after initial setup:

```bash
docker compose up -d   # start
docker compose down    # stop
```

1. Copy PDFs into `input_pdfs/`
2. Open the browser UI at `http://localhost:8501`
3. Click **Start Extraction**
4. Collect the Excel file from `output_excel/`

No command-line interaction is needed beyond starting and stopping the container.

# PDF to Excel Extractor

Converts one-page scanned PDF files into a single Excel workbook using OCR and table extraction, served through a Streamlit UI.

Designed for scanned technical drawings, parts lists, and table-heavy documents where each PDF is a single-page image. Runs entirely on CPU with open-source Python tooling.

> **Deployment note:** This app requires PaddlePaddle and PaddleOCR (~1.5 GB of dependencies) and cannot be hosted on Streamlit Community Cloud. You can run it directly on Windows (see the Windows guide below) or use Docker on any Linux server or VPS.

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

## Deployment — Windows (No Docker)

This guide is for running the app directly on your Windows computer, step by step. No technical background is needed — just follow the steps in order.

> **Before you start:** The installation will download about **1.5–2 GB** of files (Python libraries for AI/OCR). Make sure you have a stable internet connection and enough disk space.

---

### Step 1 — Install Python

1. Open your browser and go to: **https://www.python.org/downloads/**
2. Click the big yellow **"Download Python 3.11.x"** button
3. Run the downloaded installer
4. **Important:** On the first screen of the installer, check the box that says **"Add Python to PATH"** (it is unchecked by default)
5. Click **Install Now** and wait for it to finish
6. Click **Close** when done

To confirm it worked:
- Press `Win + R`, type `cmd`, press Enter
- In the black window, type `python --version` and press Enter
- You should see something like `Python 3.11.9`

---

### Step 2 — Install Poppler (needed to read PDF files)

Poppler is a small free tool that lets Python open PDF files.

1. Go to: **https://github.com/oschwartz10612/poppler-windows/releases**
2. Download the latest file named something like `Release-XX.XX.X-0.zip`
3. Unzip the file — you will get a folder (e.g. `poppler-24.08.0`)
4. Move that folder to `C:\poppler` so the path looks like `C:\poppler\Library\bin`
5. Now add it to Windows PATH:
   - Press `Win + S`, search for **"Edit the system environment variables"** and open it
   - Click **"Environment Variables..."**
   - Under **"System variables"**, find the row called **Path** and double-click it
   - Click **New** and paste: `C:\poppler\Library\bin`
   - Click **OK** on all windows to save

---

### Step 3 — Download the App

**Option A — if you have Git installed:**
- Open `cmd` and run:
  ```
  git clone https://github.com/wael-fahmy/pdf-excel
  cd pdf-excel
  ```

**Option B — download as ZIP (easier):**
1. Go to the GitHub page of the project
2. Click the green **"Code"** button → **"Download ZIP"**
3. Unzip the downloaded file to a folder you can find easily, e.g. `C:\pdf-excel`

---

### Step 4 — Create a Virtual Environment

A virtual environment keeps the app's libraries separate from the rest of your computer.

1. Open `cmd` (press `Win + R`, type `cmd`, press Enter)
2. Navigate to the project folder. For example:
   ```
   cd C:\pdf-excel
   ```
3. Create the virtual environment:
   ```
   python -m venv venv
   ```
4. Activate it:
   ```
   venv\Scripts\activate
   ```
   You will see `(venv)` appear at the start of the line — this means it is active.

---

### Step 5 — Install the App's Libraries

Still in the same `cmd` window (with `(venv)` showing), run:

```
pip install -r requirements.txt
```

This will download and install everything the app needs. **It will take several minutes** (the PaddlePaddle AI library alone is about 1.5 GB). Let it run — do not close the window.

When it finishes you will see something like `Successfully installed ...`

---

### Step 6 — Create the Input and Output Folders

In the same `cmd` window, run:

```
mkdir input_pdfs
mkdir output_excel
```

---

### Step 7 — Run the App

In the same `cmd` window, run:

```
streamlit run app.py
```

After a few seconds you will see a message like:

```
You can now view your Streamlit app in your browser.
Local URL: http://localhost:8501
```

Open your browser and go to: **http://localhost:8501**

The app is now running on your computer.

---

### Step 8 — Use the App

**To process PDFs from a folder:**
1. Copy your PDF files into the `input_pdfs` folder
2. In the app sidebar, select **Read Folder**
3. Set the input path to the full path of your `input_pdfs` folder (e.g. `C:\pdf-excel\input_pdfs`)
4. Set the output path to the full path of your `output_excel` folder (e.g. `C:\pdf-excel\output_excel`)
5. Click **Start Extraction**
6. When done, your Excel file will be at `C:\pdf-excel\output_excel\Extracted_Data.xlsx`

**To upload PDFs directly in the browser:**
1. In the app sidebar, keep **Upload PDFs** selected
2. Click the upload area and choose your PDF files
3. Click **Start Extraction**
4. Download the Excel file directly from the browser

---

### How to Start the App Again Later

Every time you want to use the app after the initial setup:

1. Open `cmd`
2. Go to the project folder:
   ```
   cd C:\pdf-excel
   ```
3. Activate the environment:
   ```
   venv\Scripts\activate
   ```
4. Start the app:
   ```
   streamlit run app.py
   ```
5. Open **http://localhost:8501** in your browser

To stop the app: go back to the `cmd` window and press `Ctrl + C`.

---

### Troubleshooting

**"python is not recognized"**
You forgot to check "Add Python to PATH" during installation. Uninstall Python and reinstall it, making sure to check that box.

**"No module named paddleocr" or similar**
Make sure you activated the virtual environment (`venv\Scripts\activate`) before running the app.

**PDF pages are not being read**
Confirm Poppler is installed and its `bin` folder is in the PATH (Step 2). Restart `cmd` after changing environment variables.

**The app is slow on the first file**
This is normal — the AI model loads into memory on the first run. Subsequent files in the same session are faster.

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

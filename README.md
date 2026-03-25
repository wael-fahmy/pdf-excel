# PDF to Excel Extractor

This project converts one-page PDF image files into a single Excel sheet using OCR and table extraction through a Streamlit app.

It is designed for scanned technical drawings, parts lists, and similar table-heavy documents where each PDF is a single page image. The app runs on CPU, uses open-source Python tooling, supports a friendly Streamlit UI, and can run either as a hosted Streamlit app or in Docker on a private machine.

## What It Does

- Supports two input modes:
  - upload PDFs directly in the Streamlit UI
  - read all `.pdf` files from an input folder
- Extracts structured table data from each file
- Tries multiple image preprocessing variants automatically to improve OCR
- Saves one Excel workbook with one sheet for all files
- Writes each file in this layout:
  - row 1: headers
  - row 2: values
  - row 3: blank separator
- Adds extra columns like `Unstructured_1` when text is found outside the main table
- Shows sample previews in the Streamlit UI
- Saves progress and cached records in folder mode so a stopped run can continue later

## Repository Layout

```text
.
|-- app.py
|-- Dockerfile
|-- docker-compose.yml
|-- packages.txt
|-- requirements.txt
|-- .streamlit/config.toml
|-- input_pdfs/
`-- output_excel/
```

## Output Files

During a run, the app writes:

- `output_excel/Extracted_Data.xlsx`
- `output_excel/.pdf_cache/progress.json`
- `output_excel/.pdf_cache/records/*.json`
- `output_excel/.pdf_cache/samples/*.png`

This means the Excel file and resume state are both saved automatically to disk while the app is running.

## Run Options

### Option 1: Streamlit Deployment

This is the best option when the app will be deployed as a Streamlit app and users should upload files in the browser.

1. Deploy the repo.
2. Make sure [packages.txt](d:/pdf-excel/packages.txt) is included so Poppler and OCR system libraries are installed.
3. Open the app in the browser.
4. Choose `Upload PDFs`.
5. Upload one or more PDF files.
6. Click `Start Extraction`.
7. Download `Extracted_Data.xlsx`.

Notes for hosted Streamlit:

- upload mode is the recommended mode
- the output file is available for download in the session
- folder-based resume is mainly intended for Docker/private-server usage

### Option 2: Docker / Private Server

1. Put your PDF files in [input_pdfs](d:/pdf-excel/input_pdfs).
2. Start the app:

```powershell
docker compose build --no-cache
docker compose up
```

3. Open `http://localhost:8501`
4. In the app choose `Read Folder`
5. Keep input path as `/data/input`
6. Keep output path as `/data/output`
7. Click `Start Extraction`
8. Get the result from [output_excel](d:/pdf-excel/output_excel) as `Extracted_Data.xlsx`

## Resume Behavior

In folder mode, if the app is stopped:

- finished files stay saved in cache
- the Excel file remains on disk
- the next run can continue from the last saved point

Useful options in the UI for folder mode:

- `Resume previous run`: keeps existing progress
- `Retry failed files`: reruns only files that previously failed
- `Clear Cache`: removes cached progress and starts fresh

If you turn off resume, the app clears the old cache and old Excel file before starting a new run.

In upload mode, the app is intentionally session-based and optimized for hosted Streamlit use.

## Expected Input Type

Best results are expected for:

- scanned technical tables
- engineering drawings with parts lists
- machine-generated text inside table boxes
- single-page PDF image files

Harder cases:

- very faint scans
- rotated pages
- handwritten notes
- badly skewed or low-resolution images

## OCR Notes

The app uses `paddleocr` with `PPStructure` on CPU only. It tries several versions of the page image:

- original image
- denoised grayscale
- thresholded high-contrast image
- sharpened image

It keeps the best extraction result automatically.

## Streamlit UI Notes

The UI is built to be friendly for non-technical users:

- upload mode for hosted usage
- folder mode for server and Docker usage
- progress metrics during extraction
- sample previews for extracted records
- direct Excel download from the browser
- a clear summary of files that failed

## Deployment Files

- [requirements.txt](d:/pdf-excel/requirements.txt): Python dependencies
- [packages.txt](d:/pdf-excel/packages.txt): system packages needed by `pdf2image` and OCR on Streamlit deployments
- [.streamlit/config.toml](d:/pdf-excel/.streamlit/config.toml): Streamlit config including larger upload size and theme colors
- [docker-compose.yml](d:/pdf-excel/docker-compose.yml): Docker deployment for local/server use

## Docker Notes

The service in [docker-compose.yml](d:/pdf-excel/docker-compose.yml) includes:

- volume mapping for input and output folders
- automatic restart with `unless-stopped`
- a healthcheck for the Streamlit app

## Troubleshooting

### `ImportError: cannot import name 'PPStructure' from 'paddleocr'`

Rebuild the Docker image so the pinned dependencies from [requirements.txt](d:/pdf-excel/requirements.txt) are installed:

```powershell
docker compose build --no-cache
docker compose up
```

### No PDFs found

In folder mode, make sure your files are inside [input_pdfs](d:/pdf-excel/input_pdfs) and end with `.pdf`.

### Upload mode on hosted Streamlit is preferred

If the app is deployed on Streamlit, use `Upload PDFs` instead of `Read Folder` unless the deployment has access to mounted folders.

### Output file not updating

Check that [output_excel](d:/pdf-excel/output_excel) exists and Docker can write to it.

### Extraction quality is weak

Try:

- cleaner scans
- higher resolution PDFs
- cropping unnecessary borders before input
- smaller test batches first

## Notes For Hand-Off

For a non-technical operator, the normal workflow is:

1. copy PDFs into the input folder
2. run Docker
3. open the browser UI
4. click one button
5. collect the Excel file from the output folder

No command-line interaction is needed after startup.

# PDF2Excel — Functional Website

This is a real working PDF → Excel web application, not just a UI mockup.

## Features

- Upload PDF from browser
- Drag & drop
- 25 MB limit
- Server-side PDF parsing
- Table detection with `pdfplumber`
- Multi-page table extraction
- Fallback text extraction when no table is detected
- Creates real `.xlsx` files using `pandas` + `openpyxl`
- Automatic column sizing
- Freeze panes and filters
- Download endpoint
- JSON API
- Docker deployment support

## Run on Windows

Install Python 3.11+.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open:

http://127.0.0.1:5000

## API

### POST /api/convert

Form field:

`file` = PDF

Example using curl:

```bash
curl -X POST -F "file=@sample.pdf" http://127.0.0.1:5000/api/convert
```

Response contains:

- `job_id`
- `pages`
- `rows`
- `sheets`
- `filename`

### GET /api/download/<job_id>

Downloads the generated Excel workbook.

## Production upgrade

For scanned PDFs, add OCR with Tesseract/PaddleOCR. For difficult financial/invoice tables, add Camelot/Tabula and a better table-detection strategy.

For production, also add:
- automatic deletion of old job folders
- authentication/rate limiting
- antivirus scanning
- persistent job storage if needed
- HTTPS
- background workers for large PDFs

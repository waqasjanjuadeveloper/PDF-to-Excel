from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from pathlib import Path
import tempfile
import uuid
import re
import shutil
import os
import base64

import pdfplumber
import pandas as pd

# Fix template folder path for Vercel
template_dir = str(Path(__file__).parent.parent / "templates")
app = Flask(__name__, template_folder=template_dir)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# For Vercel, we must use /tmp for writable files
BASE = Path(tempfile.gettempdir()) / "pdf2excel_jobs"

def ensure_base():
    """Ensure the base jobs directory exists."""
    if not BASE.exists():
        try:
            BASE.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error creating BASE directory: {e}")

ALLOWED = {".pdf"}

LEDGER_COLUMNS = [
    "Date", "Type", "Number", "Customer Reference", "Ref. Date",
    "Debit", "Credit", "PDC Value", "Balance", "Due Date", "Paid Date"
]
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
MONEY_RE = re.compile(r"-?(?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2}|\.\d{2}")


def clean_cell(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def words_by_y(page):
    words = page.extract_words(
        x_tolerance=1,
        y_tolerance=3,
        keep_blank_chars=False,
    )
    groups = {}
    for word in words:
        y = round(float(word["top"]) / 3) * 3
        groups.setdefault(y, []).append(word)
    for y in groups:
        groups[y] = sorted(groups[y], key=lambda w: w["x0"])
    return groups


def extract_ledger_page(page):
    groups = words_by_y(page)
    lines = [(y, groups[y]) for y in sorted(groups)]

    date_indices = [
        i for i, (_, line) in enumerate(lines)
        if line and DATE_RE.fullmatch(line[0]["text"].strip())
    ]

    records = []
    start_indices = []
    for date_idx in date_indices:
        start_idx = date_idx
        if date_idx > 0:
            prev_line = lines[date_idx - 1][1]
            pre_number_words = [
                w["text"].strip() for w in prev_line
                if 85 <= float(w["x0"]) < 155
            ]
            has_number = any(re.search(r"\d", value) for value in pre_number_words)
            has_reference = any(155 <= float(w["x0"]) < 323 for w in prev_line)
            prev_is_date = bool(prev_line and DATE_RE.fullmatch(prev_line[0]["text"].strip()))
            if has_number and has_reference and not prev_is_date:
                start_idx = date_idx - 1
        start_indices.append(start_idx)

    for pos, date_idx in enumerate(date_indices):
        start_idx = start_indices[pos]
        next_start_idx = start_indices[pos + 1] if pos + 1 < len(start_indices) else len(lines)
        transaction_lines = lines[start_idx:next_start_idx]
        if not transaction_lines:
            continue

        date = lines[date_idx][1][0]["text"].strip()
        type_value = ""
        number = ""
        customer_parts = []
        ref_date_words = []
        amount_words = {"debit": [], "credit": [], "pdc": [], "balance": []}

        for _, line in transaction_lines:
            for word in line:
                text = clean_cell(word["text"])
                x0 = float(word["x0"])

                if 55 <= x0 < 85 and not type_value:
                    type_value = text
                    continue
                if 85 <= x0 < 155 and not number:
                    number = text
                    continue
                if 155 <= x0 < 323:
                    customer_parts.append(text)
                    continue
                if 323 <= x0 < 410:
                    if DATE_RE.search(text):
                        ref_date_words.append(text)
                    continue
                if 410 <= x0 < 490:
                    amount_words["debit"].append(text)
                elif 490 <= x0 < 570:
                    amount_words["credit"].append(text)
                elif 570 <= x0 < 630:
                    amount_words["pdc"].append(text)
                elif x0 >= 630:
                    amount_words["balance"].append(text)

        def first_money(parts):
            m = MONEY_RE.search(" ".join(parts))
            return m.group(0) if m else ""

        balance_text = " ".join(amount_words["balance"])
        balance_match = MONEY_RE.search(balance_text)
        balance = balance_match.group(0) if balance_match else ""
        balance_dates = DATE_RE.findall(balance_text)
        due_date = balance_dates[0] if balance_dates else ""
        paid_date = balance_dates[1] if len(balance_dates) > 1 else ""
        ref_date = ref_date_words[0] if ref_date_words else ""

        records.append({
            "Date": date,
            "Type": type_value,
            "Number": number,
            "Customer Reference": clean_cell(" ".join(customer_parts)),
            "Ref. Date": ref_date,
            "Debit": first_money(amount_words["debit"]),
            "Credit": first_money(amount_words["credit"]),
            "PDC Value": first_money(amount_words["pdc"]),
            "Balance": balance,
            "Due Date": due_date,
            "Paid Date": paid_date,
        })

    return records


def extract_ledger(pdf_path):
    records = []
    page_records = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            page_rows = extract_ledger_page(page)
            if page_rows:
                page_records[page_no] = page_rows
                records.extend(page_rows)
    return records, page_records, len(page_records)


def extract_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            try:
                page_tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "intersection_tolerance": 3,
                })
            except Exception:
                page_tables = []
            for table_no, table in enumerate(page_tables or [], start=1):
                cleaned = []
                for row in table:
                    row = [clean_cell(v) for v in (row or [])]
                    if any(row):
                        cleaned.append(row)
                if cleaned:
                    width = max(len(r) for r in cleaned)
                    tables.append({"page": page_no, "table": table_no,
                                   "data": [r + [""] * (width-len(r)) for r in cleaned]})
    return tables


def format_workbook(writer):
    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 70)


def make_excel(job_dir, records, page_records, total_pages):
    output = job_dir / "converted.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if records:
            df = pd.DataFrame(records, columns=LEDGER_COLUMNS)
            df.to_excel(writer, sheet_name="All Data", index=False)

            for page_no, rows in page_records.items():
                pd.DataFrame(rows, columns=LEDGER_COLUMNS).to_excel(
                    writer, sheet_name=f"Page {page_no}"[:31], index=False
                )
        else:
            tables = extract_tables(job_dir / "source.pdf") if (job_dir / "source.pdf").exists() else []
            if tables:
                all_rows = []
                for item in tables:
                    for row in item["data"]:
                        all_rows.append([item["page"], item["table"]] + row)
                width = max(len(r) for r in all_rows)
                all_rows = [r + [""] * (width-len(r)) for r in all_rows]
                cols = ["PDF Page", "Table"] + [f"Column {i}" for i in range(1, width-1)]
                pd.DataFrame(all_rows, columns=cols).to_excel(writer, sheet_name="All Data", index=False)
            else:
                pd.DataFrame([["No extractable data found"]]).to_excel(
                    writer, sheet_name="Result", index=False, header=False
                )
        format_workbook(writer)
    return output


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/convert")
def convert():
    ensure_base()
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="No PDF file uploaded."), 400
    if Path(file.filename).suffix.lower() not in ALLOWED:
        return jsonify(error="Only PDF files are allowed."), 400

    job_id = uuid.uuid4().hex
    job_dir = BASE / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = job_dir / "source.pdf"

    try:
        file.save(pdf_path)
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)

        records, page_records, ledger_pages = extract_ledger(pdf_path)
        output = make_excel(job_dir, records, page_records, total_pages)

        # Base64 encoded Excel for fallback download in stateless serverless
        with open(output, "rb") as f:
            excel_data = base64.b64encode(f.read()).decode("utf-8")

        return jsonify(
            job_id=job_id,
            filename="converted.xlsx",
            pages=total_pages,
            rows=len(records),
            sheets=(1 + ledger_pages) if records else 1,
            excel_data=excel_data,
            message=f"Successfully extracted {len(records)} ledger rows.",
        )
    except Exception as error:
        print(f"Conversion error: {error}")
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=f"Conversion failed: {str(error)}"), 500


@app.get("/api/download/<job_id>")
def download(job_id):
    ensure_base()
    job_dir = BASE / secure_filename(job_id)
    output = job_dir / "converted.xlsx"
    if not output.exists():
        return jsonify(error="File not found or expired. Please use the backup download."), 404
    return send_file(output, as_attachment=True, download_name="converted.xlsx")


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="File is too large. Maximum size is 25 MB."), 413


# For local testing
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

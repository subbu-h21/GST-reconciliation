import asyncio
import contextlib
import json
import os
import threading
import tempfile
import uuid
from io import BytesIO
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)

# {job_id: {"status": "running"|"done"|"error", "log": [...], "error": str|None, "output": bytes|None}}
_jobs: dict = {}
_jobs_lock = threading.Lock()

# Known synonyms for each canonical Inward Register field.
# Used to pre-populate the mapping form in the UI.
_FIELD_ALIASES: dict[str, list[str]] = {
    "Supplier": [
        "Party", "Party Name", "Vendor", "Vendor Name", "Supplier Name",
        "Creditor", "Name", "Trade Name", "Ledger Name", "Ledger",
    ],
    "Gross Total": [
        "Total Amount", "Invoice Value", "Bill Amount", "Total",
        "Grand Total", "Amount", "Invoice Amount", "Net Amount",
        "Total Value", "Gross Amount",
    ],
    "GSTIN/UIN": [
        "GSTIN", "GST No", "GST Number", "Supplier GSTIN", "Vendor GSTIN",
        "UIN", "GSTIN No", "GSTIN of Supplier", "Party GSTIN",
        "GSTIN/UIN of Supplier",
    ],
    "Voucher Number": [
        "Invoice No", "Invoice Number", "Bill No", "Bill Number",
        "Voucher No", "Doc No", "Document No", "Ref No", "Reference No",
        "Invoice Ref", "Entry No", "Voucher No.",
    ],
    "Type": [
        "Transaction Type", "State Type", "Supply Type", "Supply Category",
        "Intra/Inter", "Place of Supply Type",
    ],
    "Accounting Date": [
        "Invoice Date", "Bill Date", "Date", "Voucher Date",
        "Transaction Date", "Entry Date", "Doc Date", "Document Date",
        "Inv Date",
    ],
    "Taxable @ 0%":  ["0% Taxable", "Taxable 0%", "Nil Rated", "Zero Rated", "Exempted"],
    "Taxable @ 5%":  ["5% Taxable", "Taxable 5%",  "GST 5%",  "5%"],
    "Taxable @ 12%": ["12% Taxable", "Taxable 12%", "GST 12%", "12%"],
    "Taxable @ 18%": ["18% Taxable", "Taxable 18%", "GST 18%", "18%"],
    "Taxable @ 28%": ["28% Taxable", "Taxable 28%", "GST 28%", "28%"],
}


def _auto_detect(actual_cols: list[str]) -> dict[str, str]:
    """Return {canonical_field: best_matching_actual_col} for auto-detected fields."""
    upper_map = {c.strip().upper(): c for c in actual_cols}
    result: dict[str, str] = {}
    for canonical, aliases in _FIELD_ALIASES.items():
        if canonical.strip().upper() in upper_map:
            result[canonical] = upper_map[canonical.strip().upper()]
            continue
        for alias in aliases:
            if alias.strip().upper() in upper_map:
                result[canonical] = upper_map[alias.strip().upper()]
                break
    return result


class _LiveWriter:
    """Captures stdout line-by-line and appends to the job log in real time."""

    def __init__(self, job_id: str):
        self._job_id = job_id
        self._buf = ""

    def write(self, text: str) -> int:
        self._buf += text
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            with _jobs_lock:
                _jobs[self._job_id]["log"].append(line)
        return len(text)

    def flush(self):
        pass


def _run_job(
    job_id: str,
    books_bytes: bytes,
    b2b_bytes: bytes,
    col_map: dict,
    header_row: int,
) -> None:
    import gst_reconcillation_2 as gst
    gst._AI_LOG.clear()

    with tempfile.TemporaryDirectory() as tmpdir:
        books_path = Path(tmpdir) / "books.xlsx"
        b2b_path   = Path(tmpdir) / "b2b.xlsx"
        out_path   = Path(tmpdir) / "output.xlsx"

        books_path.write_bytes(books_bytes)
        b2b_path.write_bytes(b2b_bytes)

        writer = _LiveWriter(job_id)
        try:
            with contextlib.redirect_stdout(writer):
                asyncio.run(gst.main(
                    str(books_path), str(b2b_path), out_path,
                    col_map=col_map or None,
                    header_row=header_row,
                ))

            excel_bytes = out_path.read_bytes()
            with _jobs_lock:
                _jobs[job_id]["status"] = "done"
                _jobs[job_id]["output"] = excel_bytes
        except Exception as exc:
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["error"]  = str(exc)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/columns", methods=["POST"])
def columns():
    """Read the books file headers and return them with auto-detected guesses."""
    books_file = request.files.get("books")
    if not books_file:
        return jsonify({"error": "No file provided"}), 400

    header_row = int(request.form.get("header_row", 1))

    try:
        df = pd.read_excel(
            BytesIO(books_file.read()), header=header_row,
            engine="openpyxl", nrows=0,
        )
        cols = [str(c) for c in df.columns]
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400

    return jsonify({"columns": cols, "guesses": _auto_detect(cols)})


@app.route("/run", methods=["POST"])
def run():
    books_file = request.files.get("books")
    b2b_file   = request.files.get("b2b")

    if not books_file or not b2b_file:
        return jsonify({"error": "Both files are required."}), 400

    header_row = int(request.form.get("header_row", 1))
    try:
        col_map = json.loads(request.form.get("col_map", "{}"))
    except json.JSONDecodeError:
        col_map = {}
    col_map = {k: v for k, v in col_map.items() if v}

    job_id = str(uuid.uuid4())
    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "log": [], "error": None, "output": None}

    thread = threading.Thread(
        target=_run_job,
        args=(job_id, books_file.read(), b2b_file.read(), col_map, header_row),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)

    if not job:
        return jsonify({"error": "Job not found"}), 404

    return jsonify({
        "status": job["status"],
        "log":    list(job["log"]),
        "error":  job["error"],
    })


@app.route("/download/<job_id>")
def download(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)

    if not job or job["status"] != "done" or not job["output"]:
        return "Not ready", 404

    return send_file(
        BytesIO(job["output"]),
        download_name="gst_reconciliation_output.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(port=port)

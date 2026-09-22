import os
import re
import uuid
import shutil
import subprocess
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from pypdf import PdfReader, PdfWriter

try:
    from openpyxl import load_workbook
    from openpyxl.worksheet.page import PageMargins
except Exception:
    load_workbook = None
    PageMargins = None

try:
    from pdf_to_xlsx import convert_pdf_to_xlsx
except Exception:
    convert_pdf_to_xlsx = None

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.environ.get("TOOLNEST_TEMP_DIR", tempfile.gettempdir())) / "toolnest"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024

OFFICE_EXTENSIONS = {".docx", ".doc", ".odt", ".rtf", ".txt", ".xlsx", ".xls", ".ods", ".csv", ".pptx", ".ppt", ".odp"}

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition, X-ToolNest-Files"
    return response


def safe_name(name: str) -> str:
    base = Path(name or "document").name
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(base).stem).strip("._") or "document"
    ext = Path(base).suffix.lower()
    return stem + ext


def find_office_binary():
    for candidate in ("libreoffice", "soffice"):
        path = shutil.which(candidate)
        if path:
            return path
    raise RuntimeError("LibreOffice is not installed on the conversion server.")


def prepare_spreadsheet_for_pdf(input_path: Path, output_dir: Path) -> Path:
    """Prepare Excel/Calc workbooks for clean PDF pagination without changing the user's file."""
    ext = input_path.suffix.lower()
    if ext not in {".xlsx", ".xlsm", ".xltx", ".xltm"} or load_workbook is None:
        return input_path

    prepared = output_dir / f"prepared-{input_path.name}"
    keep_vba = ext in {".xlsm", ".xltm"}
    wb = load_workbook(input_path, keep_vba=keep_vba)
    for ws in wb.worksheets:
        # Office spreadsheets commonly have content wider than a portrait page.
        # Fit the complete used range to one page wide, while allowing rows to
        # continue onto additional pages vertically. This prevents isolated
        # right-side fragments and blank-looking pages in the PDF.
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = None
        ws.page_margins = PageMargins(
            left=0.25, right=0.25, top=0.35, bottom=0.35,
            header=0.1, footer=0.1
        )
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = False
        if not ws.print_area:
            ws.print_area = ws.calculate_dimension()
    wb.save(prepared)
    return prepared


def convert_one_office(input_path: Path, output_dir: Path) -> Path:
    binary = find_office_binary()
    profile = output_dir / f"profile-{uuid.uuid4().hex}"
    profile.mkdir(parents=True, exist_ok=True)
    try:
        cmd = [
            binary,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(input_path),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        output_pdf = output_dir / (input_path.stem + ".pdf")
        if completed.returncode != 0 or not output_pdf.exists() or output_pdf.stat().st_size == 0:
            detail = (completed.stderr or completed.stdout or "LibreOffice did not produce a PDF.").strip()
            raise RuntimeError(detail[-1500:])
        return output_pdf
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def merge_pdfs(pdf_paths, output_path: Path):
    writer = PdfWriter()
    for path in pdf_paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    with output_path.open("wb") as fh:
        writer.write(fh)


@app.get("/")
def home():
    return jsonify({"ok": True, "service": "ToolNest conversion API", "status": "running"})


@app.get("/health")
def health():
    office = shutil.which("libreoffice") or shutil.which("soffice")
    return jsonify({"ok": True, "service": "ToolNest conversion API", "libreoffice": bool(office)})


@app.post("/api/office-to-pdf")
def office_to_pdf():
    uploads = request.files.getlist("files") or request.files.getlist("file")
    uploads = [f for f in uploads if f and f.filename]
    if not uploads:
        return jsonify({"error": "Please upload at least one supported file."}), 400

    work = UPLOAD_DIR / f"office-{uuid.uuid4().hex}"
    work.mkdir(parents=True, exist_ok=True)
    try:
        input_paths = []
        for idx, uploaded in enumerate(uploads):
            name = safe_name(uploaded.filename)
            if Path(name).suffix.lower() not in OFFICE_EXTENSIONS:
                return jsonify({"error": f"Unsupported Office file: {uploaded.filename}"}), 400
            path = work / f"{idx}-{name}"
            uploaded.save(path)
            input_paths.append(path)

        prepared_paths = []
        for path in input_paths:
            if path.suffix.lower() in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
                prepared_paths.append(prepare_spreadsheet_for_pdf(path, work))
            else:
                prepared_paths.append(path)

        pdfs = [convert_one_office(path, work) for path in prepared_paths]
        output = work / "converted-to-pdf.pdf"
        merge_pdfs(pdfs, output)
        if not output.exists() or output.stat().st_size == 0:
            raise RuntimeError("The PDF conversion produced an empty file.")

        return send_file(
            output,
            as_attachment=True,
            download_name="converted-to-pdf.pdf",
            mimetype="application/pdf",
        )
    except subprocess.TimeoutExpired:
        return jsonify({"error": "The document took too long to convert."}), 504
    except Exception as exc:
        return jsonify({"error": "Unable to convert the selected Office file(s) to PDF.", "details": str(exc)}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@app.post("/api/pdf-to-xlsx")
def pdf_to_xlsx():
    if convert_pdf_to_xlsx is None:
        return jsonify({"error": "PDF-to-XLSX converter is not installed on this server."}), 503
    if "file" not in request.files:
        return jsonify({"error": "Please upload a PDF file."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "Please choose a PDF file."}), 400
    original = safe_name(uploaded.filename)
    if not original.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400
    job_id = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{job_id}-{original}"
    output_path = UPLOAD_DIR / f"{job_id}-{Path(original).stem}.xlsx"
    try:
        uploaded.save(input_path)
        result = convert_pdf_to_xlsx(input_path, output_path)
        if not output_path.exists() or output_path.stat().st_size == 0:
            raise RuntimeError("The converter did not produce an Excel file.")
        response = send_file(output_path, as_attachment=True,
                             download_name=f"{Path(original).stem}.xlsx",
                             mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response.headers["X-ToolNest-Mode"] = result.get("mode", "unknown")
        response.headers["X-ToolNest-Tables"] = str(result.get("tables", 0))
        return response
    except Exception as exc:
        return jsonify({"error": "Unable to convert this PDF to Excel.", "details": str(exc)}), 500
    finally:
        input_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "The upload is too large. Maximum upload size is 50 MB."}), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

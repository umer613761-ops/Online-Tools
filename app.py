import os
import re
import uuid
import tempfile
import subprocess
import shutil
from pathlib import Path

from flask import Flask, jsonify, request, send_file
import fitz
from flask_cors import CORS

from pdf_convert import convert_txt, convert_docx, convert_html, convert_xlsx, parse_pages, safe_stem

UPLOAD_DIR = Path(os.environ.get("TOOLNEST_TEMP_DIR", tempfile.gettempdir())) / "toolnest"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}}, expose_headers=["X-ToolNest-Mode", "X-ToolNest-Tables"])
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024


def safe_filename(name: str) -> str:
    name = Path(name or "document.pdf").name
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._") or "document"
    return stem + ".pdf"


def convert_office_to_pdf(source_path: Path, output_path: Path):
    """Render Office-compatible documents to PDF using LibreOffice."""
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice is not installed on the conversion server.")

    work_dir = source_path.parent / f"lo-{uuid.uuid4().hex}"
    work_dir.mkdir(parents=True, exist_ok=True)
    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(work_dir),
                str(source_path),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        generated = work_dir / f"{source_path.stem}.pdf"
        if result.returncode != 0 or not generated.exists():
            details = (result.stderr or result.stdout or "LibreOffice conversion failed.").strip()
            raise RuntimeError(details)
        shutil.move(str(generated), str(output_path))
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.post("/api/office-to-pdf")
def office_to_pdf():
    if "file" not in request.files:
        return jsonify({"error": "Please upload a document file."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "Please choose a document file."}), 400

    original_name = Path(uploaded.filename).name
    allowed = {
        ".doc", ".docx", ".odt", ".rtf",
        ".xls", ".xlsx", ".ods", ".csv",
        ".ppt", ".pptx", ".odp",
    }
    suffix = Path(original_name).suffix.lower()
    if suffix not in allowed:
        return jsonify({"error": "Unsupported document format."}), 400

    job_id = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{job_id}-{re.sub(r'[^A-Za-z0-9._-]+', '_', original_name)}"
    output_path = UPLOAD_DIR / f"{job_id}-{Path(original_name).stem}.pdf"
    uploaded.save(input_path)

    try:
        convert_office_to_pdf(input_path, output_path)
        if not output_path.exists() or output_path.stat().st_size == 0:
            raise RuntimeError("The converter did not produce a PDF.")
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"{Path(original_name).stem}.pdf",
            mimetype="application/pdf",
        )
    except Exception as exc:
        return jsonify({
            "ok": False,
            "error": "Unable to convert this document to PDF.",
            "details": str(exc) or exc.__class__.__name__,
            "exception": exc.__class__.__name__,
        }), 500
    finally:
        try:
            input_path.unlink(missing_ok=True)
        except Exception:
            pass
        try:
            output_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.get("/")
def root():
    return jsonify({"ok": True, "service": "ToolNest conversion API", "status": "running"})


@app.get("/health")
def health():
    return jsonify({
        "ok": True,
        "service": "ToolNest conversion API",
        "status": "running",
        "conversions": ["txt", "docx", "html", "xlsx"],
    })




def pdf_request_setup(extension):
    if "file" not in request.files:
        return None, jsonify({"error": "Please upload a PDF file."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename:
        return None, jsonify({"error": "Please choose a PDF file."}), 400
    original = safe_filename(uploaded.filename)
    if not original.lower().endswith(".pdf"):
        return None, jsonify({"error": "Only PDF files are supported."}), 400
    job_id = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{job_id}-{original}"
    output_path = UPLOAD_DIR / f"{job_id}-{Path(original).stem}.{extension}"
    uploaded.save(input_path)
    return (uploaded, original, input_path, output_path, job_id), None, None


def serve_pdf_conversion(extension, converter, mimetype):
    setup, error, code = pdf_request_setup(extension)
    if error:
        return error, code
    uploaded, original, input_path, output_path, job_id = setup
    try:
        pdf=fitz.open(input_path)
        pages=parse_pages(request.form.get("pages"), pdf.page_count)
        pdf.close()
        converter(input_path, output_path, pages)
        if not output_path.exists() or output_path.stat().st_size == 0:
            raise RuntimeError("The converter did not produce a file.")
        return send_file(output_path, as_attachment=True, download_name=f"{Path(original).stem}.{extension}", mimetype=mimetype)
    except Exception as exc:
        # Keep the user-facing message useful while exposing the real
        # converter exception for debugging instead of hiding it.
        details=str(exc) or exc.__class__.__name__
        if extension == "xlsx" and details.startswith("No tables or tabular data were found"):
            return jsonify({"ok": False, "error": details}), 422
        return jsonify({
            "ok": False,
            "error": f"Unable to convert this PDF to {extension.upper()}.",
            "details": details,
            "exception": exc.__class__.__name__,
        }), 500
    finally:
        try:
            input_path.unlink(missing_ok=True)
        except Exception:
            pass
        try:
            output_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.post("/api/pdf-to-txt")
def pdf_to_txt():
    return serve_pdf_conversion("txt", convert_txt, "text/plain; charset=utf-8")


@app.post("/api/pdf-to-docx")
def pdf_to_docx():
    return serve_pdf_conversion("docx", convert_docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")




@app.post("/api/pdf-to-xlsx")
def pdf_to_xlsx():
    return serve_pdf_conversion("xlsx", convert_xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/pdf-to-html")
def pdf_to_html():
    return serve_pdf_conversion("html", convert_html, "text/html; charset=utf-8")


def _cleanup_old_outputs(max_age_seconds=3600):
    import time
    now = time.time()
    for path in UPLOAD_DIR.glob("*"):
        try:
            if now - path.stat().st_mtime > max_age_seconds:
                path.unlink(missing_ok=True)
        except OSError:
            pass


@app.before_request
def cleanup_outputs():
    _cleanup_old_outputs()


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "The PDF is too large. Maximum upload size is 50 MB."}), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

import os
import re
import uuid
import tempfile
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

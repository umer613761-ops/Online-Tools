import os
import re
import uuid
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from weasyprint import HTML

from pdf_to_xlsx import convert_pdf_to_xlsx

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.environ.get("TOOLNEST_TEMP_DIR", tempfile.gettempdir())) / "toolnest"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition, X-ToolNest-Mode, X-ToolNest-Tables"
    return response

def safe_filename(name: str) -> str:
    name = Path(name or "document.pdf").name
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._") or "document"
    return stem + ".pdf"

def safe_stem(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name or "document.html").stem).strip("._") or "document"

@app.get("/")
def home():
    return jsonify({"ok": True, "service": "ToolNest conversion API", "status": "running"})

@app.get("/health")
def health():
    return jsonify({"ok": True, "service": "ToolNest conversion API"})

@app.post("/api/pdf-to-xlsx")
def pdf_to_xlsx():
    if "file" not in request.files:
        return jsonify({"error": "Please upload a PDF file."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "Please choose a PDF file."}), 400
    original = safe_filename(uploaded.filename)
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
        response = send_file(output_path, as_attachment=True, download_name=f"{Path(original).stem}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response.headers["X-ToolNest-Mode"] = result.get("mode", "unknown")
        response.headers["X-ToolNest-Tables"] = str(result.get("tables", 0))
        return response
    except Exception as exc:
        return jsonify({"error": "Unable to convert this PDF to Excel.", "details": str(exc)}), 500
    finally:
        input_path.unlink(missing_ok=True)

@app.post("/api/html-to-pdf")
def html_to_pdf():
    if "file" not in request.files:
        return jsonify({"error": "Please upload an HTML file."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename or not re.search(r"\.html?$", uploaded.filename, re.I):
        return jsonify({"error": "Only HTML files are supported."}), 400
    job_id = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{job_id}.html"
    output_path = UPLOAD_DIR / f"{job_id}.pdf"
    try:
        uploaded.save(input_path)
        HTML(filename=str(input_path), base_url=input_path.parent.as_uri() + "/").write_pdf(str(output_path))
        if not output_path.exists() or output_path.stat().st_size == 0:
            raise RuntimeError("The HTML renderer did not produce a PDF.")
        return send_file(output_path, as_attachment=True, download_name=f"{safe_stem(uploaded.filename)}-converted.pdf", mimetype="application/pdf")
    except Exception as exc:
        return jsonify({"error": "Unable to convert this HTML file to PDF.", "details": str(exc)}), 500
    finally:
        input_path.unlink(missing_ok=True)

@app.before_request
def cleanup_outputs():
    import time
    now = time.time()
    for pattern in ("*.xlsx", "*.pdf", "*.html"):
        for path in UPLOAD_DIR.glob(pattern):
            try:
                if now - path.stat().st_mtime > 3600:
                    path.unlink(missing_ok=True)
            except OSError:
                pass

@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "The upload is too large. Maximum upload size is 50 MB."}), 413

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

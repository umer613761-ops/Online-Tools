import os
import re
import uuid
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file

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


@app.get("/")
def home():
    return jsonify({
        "ok": True,
        "service": "ToolNest conversion API",
        "status": "running"
    })


@app.get("/health")
def health():
    return jsonify({
        "ok": True,
        "service": "ToolNest conversion API"
    })


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

        response = send_file(
            output_path,
            as_attachment=True,
            download_name=f"{Path(original).stem}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers["X-ToolNest-Mode"] = result.get("mode", "unknown")
        response.headers["X-ToolNest-Tables"] = str(result.get("tables", 0))
        return response
    except Exception as exc:
        return jsonify({
            "error": "Unable to convert this PDF to Excel.",
            "details": str(exc),
        }), 500
    finally:
        try:
            input_path.unlink(missing_ok=True)
        except Exception:
            pass


def _cleanup_old_outputs(max_age_seconds=3600):
    import time
    now = time.time()
    for path in UPLOAD_DIR.glob("*.xlsx"):
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


def _office_to_pdf(kind):
    if "file" not in request.files:
        return jsonify({"error": "Please upload a file."}), 400

    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "Please choose a file."}), 400

    expected = ".docx" if kind == "docx" else ".xlsx"
    if Path(uploaded.filename).suffix.lower() != expected:
        return jsonify({"error": f"Only {expected[1:].upper()} files are supported."}), 400

    job_id = uuid.uuid4().hex
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(uploaded.filename).stem).strip("._") or "document"
    input_path = UPLOAD_DIR / f"{job_id}-{stem}{expected}"
    output_dir = UPLOAD_DIR / job_id
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{stem}.pdf"

    try:
        uploaded.save(input_path)
        import subprocess
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_path)],
            capture_output=True, text=True, timeout=180
        )

        if result.returncode != 0 or not output_path.exists() or output_path.stat().st_size == 0:
            details = (result.stderr or result.stdout or "LibreOffice did not produce a PDF.").strip()
            raise RuntimeError(details)

        with output_path.open("rb") as fh:
            if fh.read(5) != b"%PDF-":
                raise RuntimeError("The conversion engine returned a non-PDF file.")

        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"{stem}.pdf",
            mimetype="application/pdf"
        )
    except Exception as exc:
        return jsonify({
            "error": f"Unable to convert {expected[1:].upper()} to PDF.",
            "details": str(exc)
        }), 500
    finally:
        try:
            input_path.unlink(missing_ok=True)
        except OSError:
            pass
        import shutil
        shutil.rmtree(output_dir, ignore_errors=True)


@app.post("/api/docx-to-pdf")
def docx_to_pdf():
    return _office_to_pdf("docx")


@app.post("/api/xlsx-to-pdf")
def xlsx_to_pdf():
    return _office_to_pdf("xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

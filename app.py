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
except Exception:
    load_workbook = None

try:
    from pdf_to_xlsx import convert_pdf_to_xlsx
except Exception:
    convert_pdf_to_xlsx = None

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.environ.get("TOOLNEST_TEMP_DIR", tempfile.gettempdir())) / "toolnest"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024

SPREADSHEET_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

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


def _prepare_spreadsheet_for_pdf(input_path: Path, output_dir: Path) -> Path:
    """Prepare modern Excel workbooks for PDF printing without changing the upload.

    Repeats the detected table/header row on continuation pages while preserving
    the workbook's existing formatting and the previously fixed fit-to-width settings.
    """
    if load_workbook is None:
        raise RuntimeError("Excel workbook support is not installed on the conversion server.")

    prepared = output_dir / f"prepared-{input_path.name}"
    wb = load_workbook(input_path, keep_vba=input_path.suffix.lower() in {".xlsm", ".xltm"})
    for ws in wb.worksheets:
        max_scan = min(ws.max_row, 12)
        header_row = None
        for row_idx in range(1, max_scan + 1):
            values = [ws.cell(row_idx, c).value for c in range(1, min(ws.max_column, 30) + 1)]
            nonempty = sum(v not in (None, "") for v in values)
            if nonempty >= 3:
                header_row = row_idx
                break
        if header_row is not None and ws.max_row > header_row:
            ws.print_title_rows = f"{header_row}:{header_row}"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.save(prepared)
    return prepared


def convert_one_office(input_path: Path, output_dir: Path) -> Path:
    binary = find_office_binary()
    profile = output_dir / f"profile-{uuid.uuid4().hex}"
    profile.mkdir(parents=True, exist_ok=True)
    try:
        conversion_input = input_path
        if input_path.suffix.lower() in SPREADSHEET_EXTENSIONS:
            conversion_input = _prepare_spreadsheet_for_pdf(input_path, output_dir)

        cmd = [
            binary,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(conversion_input),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        output_pdf = output_dir / (conversion_input.stem + ".pdf")
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

        pdfs = [convert_one_office(path, work) for path in input_paths]
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


def _parse_pages(value, total):
    if not value:
        return list(range(1, total + 1))
    pages=[]
    for part in str(value).split(','):
        part=part.strip()
        if not part:
            continue
        try:
            n=int(part)
        except ValueError:
            continue
        if 1 <= n <= total and n not in pages:
            pages.append(n)
    return sorted(pages)


def _pdf_input(uploaded):
    if not uploaded or not uploaded.filename:
        raise ValueError('Please choose a PDF file.')
    original=safe_name(uploaded.filename)
    if Path(original).suffix.lower() != '.pdf':
        raise ValueError('Only PDF files are supported.')
    job=UPLOAD_DIR / f"pdf-from-{uuid.uuid4().hex}"
    job.mkdir(parents=True, exist_ok=True)
    path=job / original
    uploaded.save(path)
    return job, path, Path(original).stem


def _page_numbers(pdf_path, pages_value):
    import fitz
    doc=fitz.open(pdf_path)
    try:
        return _parse_pages(pages_value, doc.page_count)
    finally:
        doc.close()


@app.post('/api/pdf-to-txt')
def pdf_to_txt():
    try:
        job, pdf_path, stem = _pdf_input(request.files.get('file'))
        import fitz
        doc=fitz.open(pdf_path)
        try:
            pages=_parse_pages(request.form.get('pages'), doc.page_count)
            chunks=[]
            for n in pages:
                text=doc.load_page(n-1).get_text('text').strip()
                chunks.append(f'Page {n}\n\n{text}')
            output=job / f'{stem}.txt'
            output.write_text('\n\n--------------------------------\n\n'.join(chunks), encoding='utf-8')
        finally:
            doc.close()
        return send_file(output, as_attachment=True, download_name=f'{stem}.txt', mimetype='text/plain; charset=utf-8')
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': 'Unable to convert this PDF to text.', 'details': str(exc)}), 500
    finally:
        try: shutil.rmtree(locals().get('job'), ignore_errors=True)
        except Exception: pass


@app.post('/api/pdf-to-html')
def pdf_to_html():
    try:
        job, pdf_path, stem = _pdf_input(request.files.get('file'))
        import fitz, html
        doc=fitz.open(pdf_path)
        try:
            pages=_parse_pages(request.form.get('pages'), doc.page_count)
            out=['<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF to HTML</title>',
                 '<style>html,body{margin:0;padding:0;background:#e9edf1;font-family:Arial,Helvetica,sans-serif}.document{padding:24px}.page{position:relative;margin:0 auto 24px;background:#fff;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.14)}.text{position:absolute;white-space:pre;transform-origin:0 0;line-height:1}</style></head><body><div class="document">']
            for n in pages:
                page=doc.load_page(n-1); rect=page.rect; out.append(f'<div class="page" style="width:{rect.width:.2f}px;height:{rect.height:.2f}px">')
                data=page.get_text('dict')
                for block in data.get('blocks',[]):
                    if block.get('type') != 0: continue
                    for line in block.get('lines',[]):
                        for span in line.get('spans',[]):
                            text=span.get('text','')
                            if not text: continue
                            x0,y0,x1,y1=span.get('bbox',[0,0,0,0]); size=float(span.get('size') or 10)
                            out.append(f'<span class="text" style="left:{x0:.2f}px;top:{y0:.2f}px;font-size:{size:.2f}px">{html.escape(text)}</span>')
                out.append('</div>')
            out.append('</div></body></html>')
            output=job/f'{stem}.html'; output.write_text(''.join(out),encoding='utf-8')
        finally: doc.close()
        return send_file(output, as_attachment=True, download_name=f'{stem}.html', mimetype='text/html')
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': 'Unable to convert this PDF to HTML.', 'details': str(exc)}), 500
    finally:
        try: shutil.rmtree(locals().get('job'), ignore_errors=True)
        except Exception: pass


@app.post('/api/pdf-to-docx')
def pdf_to_docx():
    try:
        job, pdf_path, stem = _pdf_input(request.files.get('file'))
        import fitz
        from docx import Document
        from docx.shared import Pt
        doc_pdf=fitz.open(pdf_path)
        try:
            pages=_parse_pages(request.form.get('pages'), doc_pdf.page_count)
            doc=Document()
            style=doc.styles['Normal']; style.font.name='Arial'; style.font.size=Pt(10)
            for idx,n in enumerate(pages):
                if idx: doc.add_page_break()
                text=doc_pdf.load_page(n-1).get_text('text').strip()
                for line in text.splitlines():
                    doc.add_paragraph(line)
            output=job/f'{stem}.docx'; doc.save(output)
        finally: doc_pdf.close()
        return send_file(output, as_attachment=True, download_name=f'{stem}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': 'Unable to convert this PDF to Word.', 'details': str(exc)}), 500
    finally:
        try: shutil.rmtree(locals().get('job'), ignore_errors=True)
        except Exception: pass


@app.post('/api/pdf-to-pptx')
def pdf_to_pptx():
    try:
        job, pdf_path, stem = _pdf_input(request.files.get('file'))
        import fitz
        from pptx import Presentation
        from pptx.util import Inches, Pt
        doc_pdf=fitz.open(pdf_path)
        try:
            pages=_parse_pages(request.form.get('pages'), doc_pdf.page_count)
            prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
            blank=prs.slide_layouts[6]
            for n in pages:
                slide=prs.slides.add_slide(blank)
                page=doc_pdf.load_page(n-1); rect=page.rect
                scale=min(prs.slide_width/rect.width, prs.slide_height/rect.height)
                tx=(prs.slide_width-rect.width*scale)/2; ty=(prs.slide_height-rect.height*scale)/2
                for block in page.get_text('blocks'):
                    if len(block)<5: continue
                    x0,y0,x1,y1,text=block[:5]
                    if not text.strip(): continue
                    box=slide.shapes.add_textbox(int(tx+x0*scale),int(ty+y0*scale),max(1,int((x1-x0)*scale)),max(1,int((y1-y0)*scale)))
                    tf=box.text_frame; tf.clear(); para=tf.paragraphs[0]; run=para.add_run(); run.text=text.strip(); run.font.size=Pt(max(8,min(24,10*scale)))
            output=job/f'{stem}.pptx'; prs.save(output)
        finally: doc_pdf.close()
        return send_file(output, as_attachment=True, download_name=f'{stem}.pptx', mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': 'Unable to convert this PDF to PowerPoint.', 'details': str(exc)}), 500
    finally:
        try: shutil.rmtree(locals().get('job'), ignore_errors=True)
        except Exception: pass


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "The upload is too large. Maximum upload size is 50 MB."}), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))

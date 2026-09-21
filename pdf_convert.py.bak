import html
import io
import os
import re
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.shared import Pt, Inches


def parse_pages(value, page_count):
    if not value:
        return list(range(1, page_count + 1))
    pages=[]
    for part in str(value).split(','):
        try:
            n=int(part.strip())
            if 1 <= n <= page_count and n not in pages:
                pages.append(n)
        except ValueError:
            pass
    return pages or list(range(1, page_count + 1))


def page_text(page):
    data=page.get_text('dict')
    blocks=[]
    for block in data.get('blocks',[]):
        if block.get('type') != 0:
            continue
        lines=[]
        for line in block.get('lines',[]):
            parts=[]
            for span in line.get('spans',[]):
                text=span.get('text','')
                if text:
                    parts.append(text)
            text=''.join(parts).rstrip()
            if text:
                lines.append(text)
        if lines:
            blocks.append('\n'.join(lines))
    return '\n'.join(blocks).strip()


def _text_quality(text):
    """Return a rough readability score for extracted PDF text (0..1).

    Some scanned PDFs contain an invisible OCR/text layer. PyMuPDF can extract
    that layer successfully even when the layer itself is badly corrupted.
    In that case, using the extracted text verbatim produces results like
    ``Cti\u2018\u201cC*rLS`` instead of the words visible on the page. This score
    helps us decide when to re-OCR the rendered page.
    """
    if not text or not text.strip():
        return 0.0

    compact = re.sub(r'\s+', ' ', text).strip()
    chars = len(compact)
    if chars < 20:
        return 0.25

    alpha_num = sum(ch.isalnum() for ch in compact)
    printable = sum(ch.isprintable() for ch in compact)
    replacement = compact.count('\ufffd')
    symbol_runs = len(re.findall(r'[^\w\s]{2,}', compact, flags=re.UNICODE))
    words = re.findall(r"[A-Za-z]{2,}", compact)

    alpha_ratio = alpha_num / max(chars, 1)
    printable_ratio = printable / max(chars, 1)
    word_ratio = min(len(words) / max(len(compact.split()), 1), 1.0)
    penalty = min(0.45, symbol_runs * 0.025 + replacement * 0.05)

    score = (
        alpha_ratio * 0.40
        + printable_ratio * 0.15
        + word_ratio * 0.45
        - penalty
    )
    return max(0.0, min(1.0, score))


def _has_large_page_image(page):
    """Detect a scanned page with a near-full-page image behind an OCR layer."""
    page_area = max(float(page.rect.width * page.rect.height), 1.0)
    image_area = 0.0
    for image in page.get_images(full=True):
        try:
            rects = page.get_image_rects(image[0])
            image_area = max(image_area, max((r.width * r.height for r in rects), default=0.0))
        except Exception:
            continue
    return image_area / page_area >= 0.70


def ocr_page(page, psm=3, scale=2.5):
    try:
        import pytesseract
        from PIL import Image, ImageOps

        pix = page.get_pixmap(
            matrix=fitz.Matrix(scale, scale),
            alpha=False,
            colorspace=fitz.csRGB,
        )
        img = Image.open(io.BytesIO(pix.tobytes('png'))).convert('L')
        img = ImageOps.autocontrast(img)
        config = f'--oem 3 --psm {psm}'
        text = pytesseract.image_to_string(img, config=config, lang='eng').strip()
        data = pytesseract.image_to_data(
            img, config=config, lang='eng', output_type=pytesseract.Output.DICT
        )
        confidences = []
        for value, conf in zip(data.get('text', []), data.get('conf', [])):
            if value.strip():
                try:
                    confidence = float(conf)
                    if confidence >= 0:
                        confidences.append(confidence)
                except (TypeError, ValueError):
                    pass
        confidence = sum(confidences) / len(confidences) if confidences else 0.0
        return text, confidence
    except Exception:
        return '', 0.0


def get_page_text(page):
    extracted = page_text(page)

    # Scanned PDFs commonly contain a bad hidden OCR layer. When a large page
    # image is present, trust fresh OCR of the visible page instead of that
    # hidden layer.
    if _has_large_page_image(page):
        candidates = []
        for psm in (3, 6):
            text, confidence = ocr_page(page, psm=psm)
            if text:
                candidates.append((confidence, text))
        if candidates:
            # Confidence alone can favour a cleaner-looking OCR pass that has
            # silently dropped whole paragraphs. Prefer the fuller pass when
            # its confidence is close to the best pass. This matters for
            # scanned letters/certificates where PSM 3 can miss text near
            # signatures, stamps, or the lower part of the page.
            candidates.sort(key=lambda item: item[0], reverse=True)
            best_conf, best_text = candidates[0]
            fuller = max(candidates, key=lambda item: len(item[1]))
            if (fuller is not candidates[0]
                    and fuller[0] >= best_conf - 5.0
                    and len(fuller[1]) >= len(best_text) * 1.12):
                return fuller[1]
            return best_text

    if extracted:
        return extracted

    candidates = []
    for psm in (3, 6):
        text, confidence = ocr_page(page, psm=psm)
        if text:
            candidates.append((confidence, text))
    if candidates:
        return max(candidates, key=lambda item: item[0])[1]
    return ''

def safe_stem(name):
    stem=Path(name).stem
    return re.sub(r'[^A-Za-z0-9._-]+','_',stem).strip('._') or 'converted-from-pdf'


def convert_txt(pdf_path, output_path, pages):
    doc=fitz.open(pdf_path)
    chunks=[]
    for n in pages:
        text=get_page_text(doc[n-1])
        chunks.append(f'Page {n}\n\n{text}')
    Path(output_path).write_text('\n\n'.join(chunks)+'\n',encoding='utf-8')
    doc.close()


def _docx_paragraphs(text):
    """Group OCR lines into readable Word paragraphs without destroying rows."""
    raw=[re.sub(r'[ \t]+', ' ', line).strip() for line in text.splitlines()]
    lines=[line for line in raw if line]
    paragraphs=[]
    current=[]

    def flush():
        nonlocal current
        if current:
            paragraphs.append(current)
            current=[]

    def is_heading(line):
        letters=re.sub(r'[^A-Za-z]', '', line)
        return bool(letters) and len(line) <= 80 and letters.upper() == letters and not line.endswith('.')

    def is_date(line):
        return bool(re.fullmatch(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', line))

    for line in lines:
        if is_heading(line) or is_date(line):
            flush()
            paragraphs.append([line])
            continue

        current.append(line)
        # A sentence-ending line is normally the end of a printed paragraph.
        # Keep short administrative/table lines separate unless they clearly
        # continue a sentence.
        if re.search(r'[.!?]["\'\)]?$', line):
            flush()

    flush()
    return paragraphs

def _render_page_image(page, scale=2.5):
    """Render a PDF page for OCR/table detection."""
    from PIL import Image
    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        alpha=False,
        colorspace=fitz.csRGB,
    )
    return Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')


def _docx_paragraphs(text):
    """Group extracted lines into readable Word paragraphs."""
    raw = [re.sub(r'[ \t]+', ' ', line).strip() for line in text.splitlines()]
    lines = [line for line in raw if line]
    paragraphs = []
    current = []

    def flush():
        nonlocal current
        if current:
            paragraphs.append(current)
            current = []

    def is_heading(line):
        letters = re.sub(r'[^A-Za-z]', '', line)
        return bool(letters) and len(line) <= 80 and letters.upper() == letters and not line.endswith('.')

    def is_date(line):
        return bool(re.fullmatch(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', line))

    for line in lines:
        if is_heading(line) or is_date(line):
            flush()
            paragraphs.append([line])
            continue
        current.append(line)
        if re.search(r'[.!?]["\'\)]?$', line):
            flush()

    flush()
    return paragraphs


def _ocr_page_lines(page, scale=2.5):
    """Return OCR lines with editable text for a scanned page."""
    try:
        import pytesseract
        from PIL import Image, ImageOps
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False, colorspace=fitz.csRGB)
        img = Image.open(io.BytesIO(pix.tobytes('png'))).convert('L')
        img = ImageOps.autocontrast(img)
        # Try two layouts and keep the fuller readable result.
        candidates = []
        for psm in (3, 6):
            cfg = f'--oem 3 --psm {psm}'
            text = pytesseract.image_to_string(img, config=cfg, lang='eng').strip()
            data = pytesseract.image_to_data(img, config=cfg, lang='eng', output_type=pytesseract.Output.DICT)
            confs=[]
            for t,c in zip(data.get('text',[]), data.get('conf',[])):
                if t.strip():
                    try:
                        c=float(c)
                        if c>=0: confs.append(c)
                    except: pass
            conf=sum(confs)/len(confs) if confs else 0
            candidates.append((conf,text))
        candidates.sort(key=lambda x:x[0], reverse=True)
        best=candidates[0]
        fuller=max(candidates,key=lambda x:len(x[1]))
        text = fuller[1] if fuller[0] >= best[0]-5 and len(fuller[1]) >= len(best[1])*1.08 else best[1]
        return [re.sub(r'[ \t]+',' ',ln).strip() for ln in text.splitlines() if ln.strip()]
    except Exception:
        return []


def _native_page_lines(page):
    text = page_text(page)
    return [re.sub(r'[ \t]+',' ',ln).strip() for ln in text.splitlines() if ln.strip()]


def _add_editable_lines(docx, lines):
    """Write extracted/OCR lines as real editable Word text."""
    for line in lines:
        p = docx.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.05
        p.add_run(line)


def _add_signature_hint(docx, page):
    """Keep the conversion editable; signatures/stamps remain non-editable visual marks only."""
    # Deliberately do not turn the whole page into an image.
    return


def convert_docx(pdf_path, output_path, pages):
    """Create an editable DOCX from native or scanned PDFs.

    Native PDF text is extracted directly. Scanned pages are OCR'd and the
    OCR result is written as real Word paragraphs, so the document is editable
    rather than a collection of full-page screenshots.
    """
    docx = Document()
    pdf = fitz.open(pdf_path)

    for idx, n in enumerate(pages):
        page = pdf[n - 1]
        if idx:
            docx.add_page_break()

        lines = _native_page_lines(page)
        # A scanned page may have a useless hidden OCR layer. Re-OCR visible
        # content whenever the page is image-heavy.
        if _has_large_page_image(page) or not lines:
            lines = _ocr_page_lines(page)

        # Keep a small page marker only when there is no extracted content.
        if not lines:
            lines = [f'[Page {n}: no readable text detected]']

        _add_editable_lines(docx, lines)

    pdf.close()
    docx.save(output_path)

def convert_html(pdf_path, output_path, pages):
    pdf=fitz.open(pdf_path)
    out=['<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
         '<title>PDF to HTML</title><style>html,body{margin:0;padding:0;background:#e9edf1;color:#111;font-family:Arial,Helvetica,sans-serif}.pdf-document{padding:24px}.pdf-page{position:relative;margin:0 auto 24px;background:#fff;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.14)}.pdf-text{position:absolute;white-space:pre;transform-origin:0 0;line-height:1}.pdf-ocr{white-space:pre-wrap;padding:32px;font:14px/1.55 Arial,sans-serif;box-sizing:border-box}@media print{body{background:#fff}.pdf-document{padding:0}.pdf-page{margin:0;box-shadow:none;break-after:page}.pdf-page:last-child{break-after:auto}}</style></head><body><div class="pdf-document">']
    for n in pages:
        page=pdf[n-1]
        rect=page.rect
        data=page.get_text('dict')
        has_spans=any(b.get('type')==0 and any(s.get('text','').strip() for l in b.get('lines',[]) for s in l.get('spans',[])) for b in data.get('blocks',[]))
        out.append(f'<section class="pdf-page" style="width:{rect.width}px;height:{rect.height}px">')
        if has_spans:
            for block in data.get('blocks',[]):
                if block.get('type')!=0: continue
                for line in block.get('lines',[]):
                    for span in line.get('spans',[]):
                        text=span.get('text','')
                        if not text: continue
                        x0,y0,x1,y1=span.get('bbox',[0,0,0,0])
                        size=max(1,float(span.get('size',10)))
                        font=html.escape(span.get('font','Arial'))
                        flags=int(span.get('flags',0))
                        weight='700' if flags & 16 else '400'
                        angle=0
                        out.append(f'<span class="pdf-text" style="left:{x0}px;top:{y0}px;font-size:{size}px;font-family:{font},Arial,sans-serif;font-weight:{weight}">{html.escape(text).replace(" ","&nbsp;")}</span>')
        else:
            text=html.escape(ocr_page(page))
            out.append(f'<div class="pdf-ocr">{text}</div>')
        out.append('</section>')
    out.append('</div></body></html>')
    Path(output_path).write_text(''.join(out),encoding='utf-8')
    pdf.close()

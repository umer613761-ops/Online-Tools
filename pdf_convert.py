import html
import io
import os
import re
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.shared import Pt


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
            candidates.sort(key=lambda item: item[0], reverse=True)
            return candidates[0][1]

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


def convert_docx(pdf_path, output_path, pages):
    docx=Document()
    normal=docx.styles['Normal']
    normal.font.name='Arial'
    normal.font.size=Pt(10.5)
    pdf=fitz.open(pdf_path)
    for idx,n in enumerate(pages):
        if idx:
            docx.add_page_break()
        text=get_page_text(pdf[n-1])
        if not text:
            p=docx.add_paragraph()
            p.add_run(f'Page {n}').bold=True
            continue
        page_lines=text.splitlines()
        first=True
        for line in page_lines:
            p=docx.add_paragraph()
            p.paragraph_format.space_after=Pt(3)
            p.add_run(line)
            first=False
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

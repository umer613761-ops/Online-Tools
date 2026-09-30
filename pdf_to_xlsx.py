import os
import re
import math
import tempfile
from pathlib import Path

import fitz
import pdfplumber
try:
    import cv2
    import numpy as np
    import pytesseract
except ImportError:
    cv2 = None
    np = None
    pytesseract = None
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# PDF points -> Excel canvas. 5 PDF points per Excel column/row gives a
# reasonably fine editable canvas while keeping worksheets manageable.
GRID_PT = 5.0
PX_PER_PT = 96.0 / 72.0


def _safe_sheet_name(name, used):
    name = re.sub(r'[\[\]:*?/\\]', ' ', name).strip() or "Page"
    name = name[:31]
    base = name
    n = 2
    while name in used:
        suffix = f" ({n})"
        name = base[:31-len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


def _numeric_value(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s
    m = re.fullmatch(r'(-?[\d,]+(?:\.\d+)?)\s*%', s)
    if m:
        return float(m.group(1).replace(',', '')) / 100.0
    cleaned = s.replace(',', '')
    if re.fullmatch(r'\$[-+]?\d+(?:\.\d+)?', cleaned):
        try:
            return float(cleaned[1:])
        except ValueError:
            return s
    if re.fullmatch(r'[-+]?\d+(?:\.\d+)?', cleaned):
        try:
            n = float(cleaned)
            return int(n) if n.is_integer() else n
        except ValueError:
            pass
    return s


def _extract_images(doc, page, page_index, temp_dir):
    records = []
    seen = set()
    for img in page.get_images(full=True):
        xref = img[0]
        if xref in seen:
            continue
        seen.add(xref)
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        try:
            data = doc.extract_image(xref)
            ext = data.get('ext', 'png')
            path = os.path.join(temp_dir, f"p{page_index+1}_img{xref}.{ext}")
            with open(path, 'wb') as f:
                f.write(data['image'])
        except Exception:
            continue
        for rect in rects:
            records.append({
                'path': path,
                'x0': max(0.0, float(rect.x0)),
                'y0': max(0.0, float(rect.y0)),
                'x1': min(float(page.rect.width), float(rect.x1)),
                'y1': min(float(page.rect.height), float(rect.y1)),
            })
    return records


def _rect_center(r):
    return ((r['x0'] + r['x1']) / 2.0, (r['y0'] + r['y1']) / 2.0)


def _inside(rect, bbox, margin=0):
    cx, cy = _rect_center(rect)
    return (bbox[0] - margin <= cx <= bbox[2] + margin and
            bbox[1] - margin <= cy <= bbox[3] + margin)


def _grid_col(x):
    return max(1, int(math.floor(x / GRID_PT)) + 1)


def _grid_row(y):
    return max(1, int(math.floor(y / GRID_PT)) + 1)


def _grid_range(x0, y0, x1, y1):
    # Snap both edges to the nearest 5pt canvas boundary. Using the same
    # boundary for the end of one PDF cell and the start of the next avoids
    # overlapping Excel merged ranges while retaining page geometry closely.
    c1 = max(1, int(round(x0 / GRID_PT)) + 1)
    r1 = max(1, int(round(y0 / GRID_PT)) + 1)
    c2 = max(c1, int(round(x1 / GRID_PT)))
    r2 = max(r1, int(round(y1 / GRID_PT)))
    return r1, c1, r2, c2


def _set_border_rect(ws, r1, c1, r2, c2, side):
    # Apply borders to the actual canvas cells without drawing an artificial
    # 5-point grid over the entire worksheet.
    for c in range(c1, c2 + 1):
        top = ws.cell(r1, c)
        bottom = ws.cell(r2, c)
        top.border = Border(top=side, left=top.border.left, right=top.border.right, bottom=top.border.bottom)
        bottom.border = Border(bottom=side, left=bottom.border.left, right=bottom.border.right, top=bottom.border.top)
    for r in range(r1, r2 + 1):
        left = ws.cell(r, c1)
        right = ws.cell(r, c2)
        left.border = Border(left=side, top=left.border.top, bottom=left.border.bottom, right=left.border.right)
        right.border = Border(right=side, top=right.border.top, bottom=right.border.bottom, left=right.border.left)


def _anchor_for_rect(x0, y0, x1, y1):
    r1, c1, r2, c2 = _grid_range(x0, y0, x1, y1)
    # openpyxl's offsets are EMU; using zero offsets is deliberate. The 5pt
    # canvas resolution is fine enough for normal PDF business documents.
    return TwoCellAnchor(
        _from=AnchorMarker(col=c1 - 1, colOff=0, row=r1 - 1, rowOff=0),
        to=AnchorMarker(col=c2, colOff=0, row=r2, rowOff=0),
        editAs='twoCell',
    )


def _add_image_exact(ws, img_record, max_box=None):
    x0, y0, x1, y1 = img_record['x0'], img_record['y0'], img_record['x1'], img_record['y1']
    if max_box:
        bx0, by0, bx1, by1 = max_box
        x0 = max(x0, bx0); y0 = max(y0, by0)
        x1 = min(x1, bx1); y1 = min(y1, by1)
    if x1 <= x0 or y1 <= y0:
        return
    ximg = XLImage(img_record['path'])
    ximg.width = max(8, (x1 - x0) * PX_PER_PT)
    ximg.height = max(8, (y1 - y0) * PX_PER_PT)
    ximg.anchor = _anchor_for_rect(x0, y0, x1, y1)
    ws.add_image(ximg)


def _table_cell_bboxes(table):
    return [[tuple(cell) if cell else None for cell in row.cells] for row in table.rows]


def _words_to_lines(page, table_bboxes):
    try:
        words = page.extract_words(use_text_flow=True, keep_blank_chars=False, extra_attrs=['size', 'fontname'])
    except Exception:
        return []
    usable = []
    for w in words:
        x0, top, x1, bottom = map(float, (w['x0'], w['top'], w['x1'], w['bottom']))
        cx, cy = (x0+x1)/2, (top+bottom)/2
        if any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in table_bboxes):
            continue
        usable.append(w)
    lines = []
    for w in sorted(usable, key=lambda z: (float(z['top']), float(z['x0']))):
        placed = False
        wt = float(w['top'])
        for line in reversed(lines[-5:]):
            if abs(line['top'] - wt) <= 2.5:
                line['words'].append(w)
                line['x0'] = min(line['x0'], float(w['x0']))
                line['x1'] = max(line['x1'], float(w['x1']))
                line['bottom'] = max(line['bottom'], float(w['bottom']))
                line['size'] = max(line['size'], float(w.get('size') or 10))
                placed = True
                break
        if not placed:
            lines.append({'top': wt, 'bottom': float(w['bottom']), 'x0': float(w['x0']), 'x1': float(w['x1']), 'size': float(w.get('size') or 10), 'words': [w]})
    out = []
    for line in sorted(lines, key=lambda z: z['top']):
        text = ' '.join(w['text'] for w in sorted(line['words'], key=lambda z: float(z['x0']))).strip()
        if text:
            out.append({**line, 'text': text})
    return out


def _render_vector_regions(doc_page, table_bboxes, image_records, temp_dir, page_no):
    # Preserve substantial vector-only graphics (charts, diagrams, shapes) as
    # images. Table borders are explicitly ignored. This is intentionally
    # conservative so ordinary text does not become a page screenshot.
    rects = []
    for d in doc_page.get_drawings():
        r = d.get('rect')
        if not r or r.width < 8 or r.height < 8:
            continue
        bbox = (float(r.x0), float(r.y0), float(r.x1), float(r.y1))
        if any(_bbox_overlap(bbox, tb) > 0.80 for tb in table_bboxes):
            continue
        if any(_bbox_overlap(bbox, (im['x0'], im['y0'], im['x1'], im['y1'])) > 0.80 for im in image_records):
            continue
        # Ignore tiny rules; keep larger vector objects.
        if bbox[2]-bbox[0] < 30 and bbox[3]-bbox[1] < 30:
            continue
        rects.append(bbox)
    if not rects:
        return []

    # Merge nearby vector rectangles into graphic regions.
    merged = []
    for r in sorted(rects, key=lambda z: (z[1], z[0])):
        hit = None
        for i, m in enumerate(merged):
            if _bbox_gap(m, r) <= 12:
                hit = i; break
        if hit is None:
            merged.append(list(r))
        else:
            m = merged[hit]
            m[0] = min(m[0], r[0]); m[1] = min(m[1], r[1]); m[2] = max(m[2], r[2]); m[3] = max(m[3], r[3])
    out = []
    for idx, r in enumerate(merged):
        area = max(0, r[2]-r[0]) * max(0, r[3]-r[1])
        if area < 1800:
            continue
        clip = fitz.Rect(*r)
        pix = doc_page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), clip=clip, alpha=False)
        path = os.path.join(temp_dir, f'p{page_no}_vector{idx}.png')
        pix.save(path)
        out.append({'path': path, 'x0': r[0], 'y0': r[1], 'x1': r[2], 'y1': r[3]})
    return out


def _bbox_overlap(a, b):
    x0=max(a[0],b[0]); y0=max(a[1],b[1]); x1=min(a[2],b[2]); y1=min(a[3],b[3])
    inter=max(0,x1-x0)*max(0,y1-y0)
    area=max(1,(a[2]-a[0])*(a[3]-a[1]))
    return inter/area


def _bbox_gap(a, b):
    dx=max(0, max(a[0],b[0])-min(a[2],b[2]))
    dy=max(0, max(a[1],b[1])-min(a[3],b[3]))
    return math.hypot(dx,dy)


def _style_table_cell(cell, value, header=False):
    cell.value = _numeric_value(value)
    cell.alignment = Alignment(vertical='center', horizontal='center' if header else 'left', wrap_text=True)
    if header:
        cell.fill = PatternFill('solid', fgColor='263B5A')
        cell.font = Font(color='FFFFFF', bold=True, size=10)
    else:
        cell.font = Font(size=10)



def _scanned_fallback(wb, doc_page, page_no, temp_dir, used):
    """Fallback for scanned/image-only pages: preserve the page image and
    expose OCR text below it when Tesseract is available."""
    title = f"Page {page_no} - Scanned"
    ws = wb.create_sheet(_safe_sheet_name(title, used))
    ws.sheet_view.showGridLines = False
    pix = doc_page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    path = os.path.join(temp_dir, f"p{page_no}_scanned.png")
    pix.save(path)
    img = XLImage(path)
    target_w = 800
    scale = target_w / max(1, img.width)
    img.width = target_w
    img.height = max(1, img.height * scale)
    img.anchor = 'A1'
    ws.add_image(img)
    ws.column_dimensions['A'].width = 18

    if pytesseract is not None and cv2 is not None and np is not None:
        try:
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if pix.n == 4:
                arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2GRAY)
            else:
                arr = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
            threshold = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
            data = pytesseract.image_to_data(threshold, config='--oem 3 --psm 6', output_type=pytesseract.Output.DICT)
            row = max(60, int(img.height / 18) + 5)
            ws.cell(row, 1, 'EDITABLE OCR TEXT').font = Font(size=12, bold=True)
            row += 1
            for i, raw in enumerate(data.get('text', [])):
                raw = _numeric_value(raw)
                if raw in ('', None):
                    continue
                try:
                    conf = float(data['conf'][i])
                except Exception:
                    conf = 0
                if conf < 20:
                    continue
                ws.cell(row, 1, str(raw)).alignment = Alignment(wrap_text=True, vertical='top')
                row += 1
        except Exception:
            pass
    return ws

def convert_pdf_to_xlsx(input_path, output_path):
    input_path = str(input_path)
    output_path = str(output_path)
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    table_count = 0
    thin = Side(style='thin', color='9AA4AE')

    with tempfile.TemporaryDirectory() as temp_dir, pdfplumber.open(input_path) as pdf, fitz.open(input_path) as doc:
        for page_no, (pl_page, doc_page) in enumerate(zip(pdf.pages, doc), 1):
            width, height = float(pl_page.width), float(pl_page.height)
            text = pl_page.extract_text() or ''
            lines = [s.strip() for s in text.splitlines() if s.strip()]
            tables = pl_page.find_tables()
            if not text.strip() and not tables:
                _scanned_fallback(wb, doc_page, page_no, temp_dir, used)
                continue
            title = lines[0] if lines else f'Page {page_no}'
            ws = wb.create_sheet(_safe_sheet_name(f'Page {page_no} - {title}', used))
            ws.sheet_view.showGridLines = False
            ws.freeze_panes = 'A1'

            # A fixed page canvas: one Excel sheet corresponds to one PDF page.
            total_cols = int(math.ceil(width / GRID_PT)) + 1
            total_rows = int(math.ceil(height / GRID_PT)) + 1
            for c in range(1, total_cols + 1):
                ws.column_dimensions[get_column_letter(c)].width = 0.72
            for r in range(1, total_rows + 1):
                ws.row_dimensions[r].height = GRID_PT

            table_bboxes = [tuple(t.bbox) for t in tables]
            images = _extract_images(doc, doc_page, page_no - 1, temp_dir)
            vector_images = _render_vector_regions(doc_page, table_bboxes, images, temp_dir, page_no)

            # Editable text outside tables, placed at its PDF coordinates.
            for line in _words_to_lines(pl_page, table_bboxes):
                r1, c1, r2, c2 = _grid_range(line['x0'], line['top'], line['x1'], max(line['bottom'], line['top'] + line['size'] + 1))
                if r1 == r2 and c1 == c2:
                    c2 += 2
                # Do not let one extracted text line collide with another
                # line's merged canvas range. PDF word boxes occasionally
                # overlap by a fraction of a point.
                target = ws.cell(r1, c1)
                if isinstance(target, MergedCell):
                    continue
                try:
                    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
                except ValueError:
                    continue
                cell = ws.cell(r1, c1)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = line['text']
                size = max(7, min(18, line['size']))
                cell.font = Font(size=size, bold=(r1 <= 10 and size >= 13))
                cell.alignment = Alignment(vertical='center', wrap_text=True)

            # Every detected table is placed at its original page coordinates.
            table_count += len(tables)
            for table in tables:
                rows = table.extract()
                if not rows:
                    continue
                for r_idx, row_obj in enumerate(table.rows):
                    values = rows[r_idx] if r_idx < len(rows) else []
                    for c_idx, cell_bbox in enumerate(row_obj.cells):
                        if not cell_bbox:
                            continue
                        x0, y0, x1, y1 = map(float, cell_bbox)
                        rr1, cc1, rr2, cc2 = _grid_range(x0, y0, x1, y1)
                        # Keep table cell as one editable Excel cell.
                        # Apply borders before merging so edge cells remain styled.
                        _set_border_rect(ws, rr1, cc1, rr2, cc2, thin)
                        if rr2 > rr1 or cc2 > cc1:
                            try:
                                ws.merge_cells(start_row=rr1, start_column=cc1, end_row=rr2, end_column=cc2)
                            except ValueError:
                                continue
                        value = values[c_idx] if c_idx < len(values) else ''
                        cell = ws.cell(rr1, cc1)
                        if isinstance(cell, MergedCell):
                            continue
                        _style_table_cell(cell, value, header=(r_idx == 0))
                        target_h = max(GRID_PT, (y1-y0) * 0.95)
                        per_row = target_h / max(1, rr2-rr1+1)
                        for rr in range(rr1, rr2+1):
                            ws.row_dimensions[rr].height = max(ws.row_dimensions[rr].height or GRID_PT, per_row)

                # Images whose centers fall inside table cells stay inside those cells.
                for im in images:
                    cx, cy = _rect_center(im)
                    hit = None
                    for row_obj in table.rows:
                        for cell_bbox in row_obj.cells:
                            if cell_bbox and cell_bbox[0] <= cx <= cell_bbox[2] and cell_bbox[1] <= cy <= cell_bbox[3]:
                                hit = tuple(cell_bbox); break
                        if hit: break
                    if hit:
                        pad = 2
                        _add_image_exact(ws, im, (hit[0]+pad, hit[1]+pad, hit[2]-pad, hit[3]-pad))

            # Standalone raster images remain exactly where they appeared.
            for im in images:
                if not any(_inside(im, tb) for tb in table_bboxes):
                    _add_image_exact(ws, im)

            # Vector charts/diagrams are preserved as positioned images.
            for vim in vector_images:
                _add_image_exact(ws, vim)

            # Page setup: the worksheet prints as one physical page, matching the source page.
            ws.print_area = f'A1:{get_column_letter(total_cols)}{total_rows}'
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.page_setup.orientation = 'portrait' if height >= width else 'landscape'
            ws.page_setup.paperSize = ws.PAPERSIZE_A4 if abs(width-height) > 0 else ws.PAPERSIZE_A4
            ws.page_margins.left = 0
            ws.page_margins.right = 0
            ws.page_margins.top = 0
            ws.page_margins.bottom = 0
            ws.sheet_properties.outlinePr.summaryBelow = True

        wb.save(output_path)
    return {
        'mode': 'layout-aware-page-canvas',
        'sheets': len(wb.worksheets),
        'tables': table_count,
    }


if __name__ == '__main__':
    import sys
    if len(sys.argv) != 3:
        raise SystemExit('Usage: pdf_to_xlsx.py input.pdf output.xlsx')
    convert_pdf_to_xlsx(sys.argv[1], sys.argv[2])

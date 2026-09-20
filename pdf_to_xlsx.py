import math
import os
import re
import shutil
import subprocess
from pathlib import Path

import cv2
import numpy as np
import openpyxl
import pytesseract
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


def _configure_tesseract():
    configured = os.environ.get("TESSERACT_CMD")
    candidates = []
    if configured:
        candidates.append(configured)
    found = shutil.which("tesseract")
    if found:
        candidates.append(found)
    candidates.extend([
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    ])
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate
    return None


def _ensure_tesseract():
    cmd = _configure_tesseract()
    if not cmd:
        raise RuntimeError("Tesseract OCR is not installed on this server.")
    try:
        subprocess.run([cmd, "--version"], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        raise RuntimeError("Tesseract OCR could not be started on the server.") from exc


def _render_page(page, dpi=180):
    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)


def _normalize_cell(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _ocr_page_words(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    data = pytesseract.image_to_data(
        threshold,
        config="--oem 3 --psm 6",
        output_type=pytesseract.Output.DICT,
    )
    words = []
    for i, raw in enumerate(data["text"]):
        text = _normalize_cell(raw)
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1
        if conf < 10:
            continue
        x = int(data["left"][i]); y = int(data["top"][i])
        w = int(data["width"][i]); h = int(data["height"][i])
        words.append({"text": text, "conf": conf, "x": x, "y": y,
                      "w": w, "h": h, "right": x + w, "bottom": y + h})
    return words


def _line_contours(image, orientation):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY, 15, -2)
    h, w = bw.shape
    if orientation == "vertical":
        kernel = (1, max(20, h // 40))
    else:
        kernel = (max(60, w // 35), 1)
    line = cv2.morphologyEx(bw, cv2.MORPH_OPEN,
                            cv2.getStructuringElement(cv2.MORPH_RECT, kernel))
    contours, _ = cv2.findContours(line, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return bw, [cv2.boundingRect(c) for c in contours]


def _group_positions(values, tolerance=18):
    groups = []
    for value in sorted(values):
        if not groups or value - groups[-1][-1] > tolerance:
            groups.append([value])
        else:
            groups[-1].append(value)
    return [int(round(sum(g) / len(g))) for g in groups]


def _horizontal_candidates(image):
    """Return strong horizontal table-line positions and their strength."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(
        ~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2
    )
    h, w = bw.shape
    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (max(60, w // 35), 1)
    )
    line = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel)
    proj = (line > 0).sum(axis=1)
    threshold = max(30, int(w * 0.10))
    inds = np.where(proj > threshold)[0]
    groups = []
    for y in inds:
        if not groups or y - groups[-1][-1] > 3:
            groups.append([y])
        else:
            groups[-1].append(y)
    return [(int(round(np.mean(g))), int(proj[g].max())) for g in groups]


def _detect_table_grid(image):
    """Detect a drawn table, including header rows above the longest vertical grid."""
    bw, vertical_boxes = _line_contours(image, "vertical")
    h, w = bw.shape
    candidates = []
    for x, y, cw, ch in vertical_boxes:
        if ch >= h * 0.075 and cw <= max(45, w * 0.04):
            candidates.append((x + cw // 2, y, y + ch, ch))
    if len(candidates) < 4:
        return None

    candidates.sort(key=lambda z: z[0])
    groups = []
    for item in candidates:
        placed = False
        for g in groups:
            g_top = min(v[1] for v in g)
            g_bottom = max(v[2] for v in g)
            overlap = min(item[2], g_bottom) - max(item[1], g_top)
            shorter = min(item[2] - item[1], g_bottom - g_top)
            if shorter > 0 and overlap / shorter > 0.55:
                g.append(item)
                placed = True
                break
        if not placed:
            groups.append([item])

    best = None
    for g in groups:
        if len(g) < 4:
            continue
        xs = _group_positions([v[0] for v in sorted(g)], 25)
        if len(xs) < 4:
            continue
        span = xs[-1] - xs[0]
        if span < w * 0.25:
            continue
        members = [v for v in g if any(abs(v[0] - x) <= 25 for x in xs)]
        y_top = min(v[1] for v in members)
        y_bottom = max(v[2] for v in members)
        score = len(xs) * span * (y_bottom - y_top)
        if best is None or score > best[0]:
            best = (score, xs, y_top, y_bottom)
    if not best:
        return None

    _, xs, y_top, y_bottom = best
    horizontal = _horizontal_candidates(image)
    if not horizontal:
        return None

    # Find horizontal lines belonging to the detected table.  Some scanned
    # certificates have a header whose vertical borders are shorter than the
    # body borders, so the old detector started at the first data row. Walk
    # upward through nearby strong lines to recover those header rows.
    strengths = [strength for _, strength in horizontal]
    median_strength = float(np.median(strengths)) if strengths else 1.0
    selected_top = [y for y, strength in horizontal if y >= y_top - 8 and y <= y_bottom + 8]
    if not selected_top:
        return None

    # Use the first line inside the vertical-grid region as the body anchor.
    anchor = min(selected_top, key=lambda y: abs(y - y_top))
    before = [
        (y, strength) for y, strength in horizontal
        if y < anchor - 3 and anchor - y <= max(320, int(h * 0.08))
    ]
    recovered = []
    cursor = anchor
    for y, strength in reversed(before):
        if strength >= median_strength * 0.42 and cursor - y <= max(100, int(h * 0.025)):
            recovered.append(y)
            cursor = y
        elif recovered and cursor - y <= max(55, int(h * 0.018)) and strength >= median_strength * 0.20:
            recovered.append(y)
            cursor = y
        else:
            break
    ys = sorted(set(recovered + selected_top))

    # Keep only lines in the table's horizontal span by requiring that the
    # line has comparable strength to the detected table lines.
    if len(ys) < 3:
        return None

    # Collapse duplicate/double-scanned horizontal borders.
    raw_gaps=np.diff(np.array(ys)) if len(ys)>1 else np.array([])
    typical=float(np.median(raw_gaps[raw_gaps>8])) if np.any(raw_gaps>8) else 20.0
    limit=max(8.0, typical*0.35)
    compact=[]
    for y in ys:
        if not compact or y-compact[-1]>limit: compact.append(y)
        else: compact[-1]=int(round((compact[-1]+y)/2))
    ys=compact

    # Ensure the first recovered line is genuinely above the original anchor.
    # If no header was recovered, retain the original grid behavior.
    if not any(y < anchor - 6 for y in ys):
        ys = [anchor] + [y for y in ys if y > anchor + 3]

    return xs, ys

def _crop_dark_components(gray, min_height_ratio=0.18, min_width=4, min_area=40):
    """Isolate dark printed characters while suppressing light watermark text and grid lines."""
    mask = (gray < 170).astype(np.uint8) * 255
    pad = max(8, int(min(gray.shape) * 0.06))
    mask[:pad, :] = 0
    mask[-pad:, :] = 0
    mask[:, :pad] = 0
    mask[:, -pad:] = 0

    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
    clean = np.zeros_like(mask)
    components = 0
    min_h = max(10, int(gray.shape[0] * min_height_ratio))
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if h >= min_h and w >= min_width and area >= min_area:
            clean[labels == i] = 255
            components += 1

    yy, xx = np.where(clean > 0)
    if not len(xx):
        return None, 0
    crop = clean[
        max(0, yy.min() - 12):min(clean.shape[0], yy.max() + 13),
        max(0, xx.min() - 12):min(clean.shape[1], xx.max() + 13),
    ]
    crop = cv2.resize(crop, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
    return 255 - crop, components


def _ocr_numeric_cell(image, x0, y0, x1, y1):
    gray = cv2.cvtColor(image[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY)
    candidates = []
    for threshold in (90, 140, 180):
        mask = (gray < threshold).astype(np.uint8) * 255
        pad = max(6, int(min(gray.shape) * 0.05))
        mask[:pad, :] = 0; mask[-pad:, :] = 0; mask[:, :pad] = 0; mask[:, -pad:] = 0
        n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
        clean = np.zeros_like(mask); components = 0
        min_h = max(9, int(gray.shape[0] * 0.18))
        for i in range(1, n):
            x, y, w, h, area = stats[i]
            if h >= min_h and w >= 3 and area >= 30:
                clean[labels == i] = 255; components += 1
        yy, xx = np.where(clean > 0)
        if not len(xx): continue
        crop = clean[max(0, yy.min()-8):min(clean.shape[0], yy.max()+9), max(0, xx.min()-8):min(clean.shape[1], xx.max()+9)]
        crop = cv2.resize(crop, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
        crop = 255 - crop
        text = pytesseract.image_to_string(crop, config='--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789+-=').strip()
        text = re.sub(r'[^0-9+=-]', '', text)
        if text:
            digits = len(re.findall(r'\d', text))
            candidates.append((2 if components and digits == components else 0, threshold, text))
    if not candidates: return ''
    valid = [c for c in candidates if c[0] == 2]
    if valid: return sorted(valid, key=lambda c: abs(c[1]-140))[0][2]
    return candidates[0][2]


def _ocr_text_cell(image, x0, y0, x1, y1):
    pad_y = max(8, int((y1-y0)*0.18)); yy0=max(0,y0-pad_y); yy1=min(image.shape[0],y1+pad_y)
    gray=cv2.cvtColor(image[yy0:yy1, max(0,x0+6):min(image.shape[1],x1-6)],cv2.COLOR_BGR2GRAY)
    gray=cv2.resize(gray,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
    text=pytesseract.image_to_string(gray,config='--oem 3 --psm 6').strip()
    lines=[_normalize_cell(x) for x in text.splitlines() if _normalize_cell(x)]
    return lines[0] if lines else ''


def _ocr_cell(image, x0, y0, x1, y1, numeric=False):
    if numeric:
        return _ocr_numeric_cell(image, x0, y0, x1, y1)
    return _ocr_text_cell(image, x0, y0, x1, y1)

def _ocr_words(image, psm):
    data = pytesseract.image_to_data(
        image,
        config=f"--oem 3 --psm {psm}",
        output_type=pytesseract.Output.DICT,
    )
    words = []
    for i, raw in enumerate(data["text"]):
        text = _normalize_cell(raw)
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1
        if conf < 5:
            continue
        words.append({
            "text": text,
            "x": int(data["left"][i]),
            "y": int(data["top"][i]),
            "w": int(data["width"][i]),
            "h": int(data["height"][i]),
        })
    return words


def _clean_table_image(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY, 15, -2)
    horizontal = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (max(60, image.shape[1] // 35), 1))
    )
    vertical = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, image.shape[0] // 20)))
    )
    lines = cv2.bitwise_or(horizontal, vertical)
    clean = gray.copy()
    clean[lines > 0] = 255
    return clean


def _numeric_from_words(words):
    text = " ".join(w["text"] for w in words)
    # Prefer complete numeric expressions such as 878+2=880, otherwise the
    # longest numeric token in the cell.
    expressions = re.findall(r"\d+(?:\s*\+\s*\d+\s*=\s*\d+)?|[-=]+", text)
    if not expressions:
        return ""
    return max(expressions, key=len).replace(" ", "")


def _grid_table_from_image(image, grid):
    xs,ys=grid; rows=[]
    for ri,(y0,y1) in enumerate(zip(ys[:-1],ys[1:])):
        row=[]
        for ci,(x0,x1) in enumerate(zip(xs[:-1],xs[1:])):
            numeric=ci in {1,2,3} and ri>=2
            row.append(_normalize_cell(_ocr_cell(image,x0,y0,x1,y1,numeric=numeric)))
        if any(row): rows.append(row)
    return rows


def _cluster_rows(words, y_tolerance=None):
    if not words:
        return []
    median_h = float(np.median([w["h"] for w in words]))
    tolerance = y_tolerance or max(8, median_h * 0.65)
    rows = []
    for word in sorted(words, key=lambda w: (w["y"], w["x"])):
        cy = word["y"] + word["h"] / 2
        best = None; best_delta = None
        for row in rows:
            delta = abs(cy - row["cy"])
            if delta <= tolerance and (best_delta is None or delta < best_delta):
                best, best_delta = row, delta
        if best is None:
            rows.append({"cy": cy, "words": [word]})
        else:
            best["words"].append(word)
            best["cy"] = sum(w["y"] + w["h"] / 2 for w in best["words"]) / len(best["words"])
    for row in rows:
        row["words"].sort(key=lambda w: w["x"])
    return sorted(rows, key=lambda r: r["cy"])


def _extract_native_page(page):
    words = page.get_text("words")
    if not words:
        return []
    rows = []
    normalized = []
    for x0, y0, x1, y1, text, *_ in words:
        text = _normalize_cell(text)
        if text:
            normalized.append({"text": text, "x": x0, "y": y0,
                               "w": x1 - x0, "h": y1 - y0})
    clustered = _cluster_rows(normalized)
    for row in clustered:
        cells = []
        # For native PDFs without drawn grids, preserve reading order as a
        # compact text row rather than inventing a table.
        cells.append(" ".join(w["text"] for w in row["words"]))
        rows.append(cells)
    return rows


def _write_sheet(wb, title, rows):
    ws = wb.create_sheet(title=title[:31])
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(r_idx, c_idx, _normalize_cell(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if r_idx == 1:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    ws.freeze_panes = "A2"
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for cell in ws[get_column_letter(col)]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col)].width = min(max(12, max_len + 2), 45)
    return ws


def _ocr_text_sheet(wb, page_no, image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    text = pytesseract.image_to_string(threshold, config="--oem 3 --psm 6")
    lines = [[f"Page {page_no}"]]
    for line in text.splitlines():
        line = _normalize_cell(line)
        if line:
            lines.append([line])
    return lines



def _ocr_full_page_lines(image):
    gray=cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)
    threshold=cv2.threshold(gray,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)[1]
    text=pytesseract.image_to_string(threshold,config='--oem 3 --psm 6')
    lines=[_normalize_cell(x) for x in text.splitlines() if _normalize_cell(x)]
    line_height=max(1,image.shape[0]/max(1,len(lines)))
    return [((i+0.5)*line_height,line) for i,line in enumerate(lines)]


def _write_page_content_sheet(wb, page_no, image, grid=None, full_lines=None):
    """Write the whole page in reading order, preserving text around a table."""
    ws = wb.create_sheet(title=f"Page {page_no}")
    height = image.shape[0]
    full_lines = full_lines or []
    row = 1

    if grid:
        xs, ys = grid
        table_top, table_bottom = ys[0], ys[-1]

        # Keep all OCR content above the table.
        for cy, line in full_lines:
            if cy < table_top - 12:
                ws.cell(row, 1, line)
                ws.cell(row, 1).alignment = Alignment(vertical="top", wrap_text=True)
                row += 1

        row += 1
        ws.cell(row, 1, "TABLE")
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 1).fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        row += 1

        table_rows = _grid_table_from_image(image, grid)
        for r_idx, values in enumerate(table_rows):
            for col, value in enumerate(values, start=1):
                cell = ws.cell(row, col, _normalize_cell(value))
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if r_idx == 0:
                    cell.font = Font(bold=True)
            row += 1

        row += 1
        # Keep all OCR content below the table.
        for cy, line in full_lines:
            if cy > table_bottom + 12:
                ws.cell(row, 1, line)
                ws.cell(row, 1).alignment = Alignment(vertical="top", wrap_text=True)
                row += 1
    else:
        for _, line in full_lines:
            ws.cell(row, 1, line)
            ws.cell(row, 1).alignment = Alignment(vertical="top", wrap_text=True)
            row += 1

    ws.column_dimensions["A"].width = 60
    for col in range(2, ws.max_column + 1):
        max_len = 0
        for cell in ws[get_column_letter(col)]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col)].width = min(max(14, max_len + 2), 30)
    ws.freeze_panes = "A2"
    return ws


def convert_pdf_to_xlsx(input_path, output_path):
    _ensure_tesseract()
    doc=pymupdf.open(str(input_path)); wb=openpyxl.Workbook(); wb.remove(wb.active)
    tables=0; modes=[]
    for page_no,page in enumerate(doc,start=1):
        image=_render_page(page,180); full_lines=_ocr_full_page_lines(image); grid=_detect_table_grid(image)
        if grid and len(grid[0])>=4 and len(grid[1])>=3:
            _write_page_content_sheet(wb,page_no,image,grid,full_lines); tables+=1; modes.append('ocr-grid-full-page')
        else:
            _write_page_content_sheet(wb,page_no,image,None,full_lines); modes.append('ocr-text-full-page')
    if not wb.sheetnames: _write_sheet(wb,'Text',[['No extractable content was detected.']])
    wb.save(output_path); doc.close()
    return {'tables':tables,'mode':','.join(sorted(set(modes))) or 'none'}


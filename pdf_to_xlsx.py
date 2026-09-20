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


def _render_page(page, dpi=216):
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


def _detect_table_grid(image):
    """Detect a drawn table and return its column/row boundaries."""
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
            g_top = min(v[1] for v in g); g_bottom = max(v[2] for v in g)
            overlap = min(item[2], g_bottom) - max(item[1], g_top)
            shorter = min(item[2] - item[1], g_bottom - g_top)
            if shorter > 0 and overlap / shorter > 0.55:
                g.append(item); placed = True; break
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
    crop = bw[max(0, y_top - 20):min(h, y_bottom + 80), xs[0]:xs[-1] + 1]
    kernel_width = max(60, (xs[-1] - xs[0]) // 35)
    horizontal = cv2.morphologyEx(
        crop, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_width, 1))
    )
    proj = (horizontal > 0).sum(axis=1)
    threshold = max(30, (xs[-1] - xs[0]) * 0.10)
    inds = np.where(proj > threshold)[0]
    groups_y = []
    for yy in inds:
        if not groups_y or yy - groups_y[-1][-1] > 3:
            groups_y.append([yy])
        else:
            groups_y[-1].append(yy)

    candidates_y = []
    for g in groups_y:
        row = horizontal[g[len(g) // 2]] > 0
        # Longest continuous run is a useful discriminator against text strokes.
        longest = 0; run = 0
        for bit in row:
            if bit:
                run += 1; longest = max(longest, run)
            else:
                run = 0
        candidates_y.append((
            int(round(np.mean(g))) + max(0, y_top - 20),
            int(proj[g].max()),
            longest,
        ))
    if len(candidates_y) < 3:
        return None

    # Estimate the dominant row spacing. Large certificate tables have
    # ~100-150 px rows; compact marks tables have ~35-55 px rows at this
    # render resolution.
    ys_all = np.array([v[0] for v in candidates_y])
    gaps = np.diff(ys_all)
    large = gaps[(gaps >= 110) & (gaps <= 160)]
    compact = gaps[(gaps >= 25) & (gaps <= 60)]
    if len(large) >= 4:
        spacing = float(np.median(large))
    elif len(compact) >= 4:
        spacing = float(np.median(compact))
    else:
        valid = gaps[gaps >= 15]
        if len(valid) == 0:
            return None
        spacing = float(np.median(valid))

    # The first real table boundary is normally the strongest line in the
    # first small vertical window. This avoids a faint line just above a table.
    first_window = [v for v in candidates_y if v[0] <= y_top + 120]
    if not first_window:
        first_window = candidates_y[:3]
    start = max(first_window, key=lambda v: v[2])

    selected = [start]
    current = start[0]
    used = {id(start)}
    max_run = max(v[2] for v in candidates_y) or 1

    while True:
        target = current + spacing
        pool = [v for v in candidates_y
                if v[0] > current and v[0] - current >= spacing * 0.35
                and v[0] - current <= spacing * 2.0]
        if not pool:
            break
        # Distance matters, but a strong continuous horizontal line is given
        # substantial weight so text/watermark strokes do not win over the
        # actual table border.
        chosen = min(
            pool,
            key=lambda v: abs(v[0] - target) / spacing - 1.5 * (v[2] / max_run),
        )
        selected.append(chosen)
        current = chosen[0]
        if len(selected) > 200:
            break

    ys = [v[0] for v in selected]
    if len(ys) < 3:
        return None
    return xs, ys

def _ocr_cell(image, x0, y0, x1, y1, numeric=False):
    pad = 6
    x0 += pad; y0 += pad; x1 -= pad; y1 -= pad
    if x1 <= x0 or y1 <= y0:
        return ""
    gray = cv2.cvtColor(image[y0:y1, x0:x1], cv2.COLOR_BGR2GRAY)
    gray = cv2.copyMakeBorder(gray, 12, 12, 12, 12, cv2.BORDER_CONSTANT, value=255)
    if numeric:
        # At the tested resolution Otsu often recovers faint leading digits,
        # while the original grayscale handles ordinary one/two digit cells.
        variants = [gray, cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]]
        results = []
        for variant in variants:
            text = pytesseract.image_to_string(
                variant,
                config="--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789-=",
            ).strip()
            text = re.sub(r"[^0-9=-]", "", text)
            if text:
                results.append(text)
        if not results:
            return ""
        # Prefer a multi-digit result over a one-digit result when they differ;
        # this fixes faint leading digits such as 17/80 without hard-coding
        # document values.
        if len(set(results)) > 1:
            multi = [r for r in results if len(re.sub(r"[^0-9]", "", r)) >= 2]
            if multi:
                return multi[-1]
        return results[0]

    text = pytesseract.image_to_string(
        gray, config="--oem 3 --psm 7"
    ).strip()
    return _normalize_cell(text)


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
    xs, ys = grid
    crop = image[ys[0]:ys[-1] + 1, xs[0]:xs[-1] + 1]
    clean = _clean_table_image(crop)
    text_words = _ocr_words(clean, 11)
    numeric_words = _ocr_words(clean, 11)
    rows = []

    for ri, (y0, y1) in enumerate(zip(ys[:-1], ys[1:])):
        cy0 = y0 - ys[0]
        cy1 = y1 - ys[0]
        row = []
        for ci, (x0, x1) in enumerate(zip(xs[:-1], xs[1:])):
            cx0 = x0 - xs[0]
            cx1 = x1 - xs[0]
            source = numeric_words if ci in {1, 2, 3} else text_words
            inside = [w for w in source
                      if cy0 <= w["y"] + w["h"] / 2 <= cy1
                      and cx0 <= w["x"] + w["w"] / 2 <= cx1]
            if ci in {1, 2, 3}:
                value = _numeric_from_words(inside)
                # A second OCR layout is useful when the sparse layout misses
                # a numeric cell. Prefer a multi-digit fallback over a lone
                # digit when both are available.
                if not value:
                    fallback = [w for w in text_words
                                if cy0 <= w["y"] + w["h"] / 2 <= cy1
                                and cx0 <= w["x"] + w["w"] / 2 <= cx1]
                    value = _numeric_from_words(fallback)
            else:
                value = _normalize_cell(" ".join(w["text"] for w in inside))
            row.append(value)
        if any(row):
            rows.append(row)
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


def convert_pdf_to_xlsx(input_path, output_path):
    _ensure_tesseract()
    doc = pymupdf.open(str(input_path))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    tables = 0
    modes = []
    text_rows = []

    for page_no, page in enumerate(doc, start=1):
        native_text = page.get_text("text").strip()
        image = _render_page(page)
        grid = _detect_table_grid(image)
        if grid:
            rows = _grid_table_from_image(image, grid)
            if len(rows) >= 2 and len(rows[0]) >= 2:
                _write_sheet(wb, f"Table {tables + 1}", rows)
                tables += 1
                modes.append("ocr-grid")
                continue

        if native_text:
            # Preserve ordinary native-PDF text without pretending it is a table.
            text_rows.append([f"Page {page_no}"])
            for line in native_text.splitlines():
                line = _normalize_cell(line)
                if line:
                    text_rows.append([line])
            modes.append("native-text")
        else:
            text_rows.extend(_ocr_text_sheet(wb, page_no, image))
            modes.append("ocr-text")

    if text_rows:
        _write_sheet(wb, "Text", text_rows)
    if not wb.sheetnames:
        _write_sheet(wb, "Text", [["No extractable text or table was detected."]])
    wb.save(output_path)
    doc.close()
    return {"tables": tables, "mode": ",".join(sorted(set(modes))) or "none"}

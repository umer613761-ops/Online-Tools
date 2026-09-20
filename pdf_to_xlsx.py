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
        raise RuntimeError(
            "Tesseract OCR is not installed on this server. "
            "Install the tesseract-ocr system package or set TESSERACT_CMD."
        )
    try:
        subprocess.run([cmd, "--version"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        raise RuntimeError("Tesseract OCR could not be started on the server.") from exc


def _render_page(page, dpi=220):
    scale = dpi / 72.0
    matrix = pymupdf.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img


def _normalize_cell(text):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text


def _ocr_data(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # Keep a clean grayscale copy for OCR. Otsu is helpful on scans while
    # preserving enough character detail for ordinary certificates.
    _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    for candidate in (threshold, gray):
        data = pytesseract.image_to_data(
            candidate,
            config="--oem 3 --psm 6",
            output_type=pytesseract.Output.DICT,
        )
        words = []
        for i, raw in enumerate(data["text"]):
            text = _normalize_cell(raw)
            try:
                conf = float(data["conf"][i])
            except (ValueError, TypeError):
                conf = -1
            if text and conf >= 15:
                words.append({
                    "text": text,
                    "conf": conf,
                    "x": int(data["left"][i]),
                    "y": int(data["top"][i]),
                    "w": int(data["width"][i]),
                    "h": int(data["height"][i]),
                    "right": int(data["left"][i]) + int(data["width"][i]),
                    "bottom": int(data["top"][i]) + int(data["height"][i]),
                })
        if len(words) > 0:
            return words
    return []


def _cluster_rows(words, y_tolerance=None):
    if not words:
        return []
    median_h = float(np.median([w["h"] for w in words]))
    tolerance = y_tolerance or max(8, median_h * 0.65)
    rows = []
    for word in sorted(words, key=lambda w: (w["y"], w["x"])):
        cy = word["y"] + word["h"] / 2
        best = None
        best_delta = None
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


def _detect_grid(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY, 15, -2)
    h, w = bw.shape
    horizontal_len = max(20, w // 30)
    vertical_len = max(20, h // 30)
    horizontal = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (horizontal_len, 1))
    )
    vertical = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (1, vertical_len))
    )
    grid = cv2.add(horizontal, vertical)
    contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    for contour in contours:
        x, y, cw, ch = cv2.boundingRect(contour)
        if cw >= 40 and ch >= 18:
            boxes.append((x, y, cw, ch))
    return boxes


def _table_from_grid(image, words):
    boxes = _detect_grid(image)
    if not boxes:
        return None
    xs = sorted(set([x for x, y, w, h in boxes] + [x + w for x, y, w, h in boxes]))
    ys = sorted(set([y for x, y, w, h in boxes] + [y + h for x, y, w, h in boxes]))
    # Merge nearly identical boundaries created by line thickness/contours.
    def merge(vals, tol=8):
        out = []
        for v in vals:
            if not out or abs(v - out[-1]) > tol:
                out.append(v)
            else:
                out[-1] = int(round((out[-1] + v) / 2))
        return out
    xs, ys = merge(xs), merge(ys)
    if len(xs) < 3 or len(ys) < 3:
        return None

    # Only keep a plausible table region rather than isolated lines.
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    if (x1 - x0) < image.shape[1] * 0.15 or (y1 - y0) < image.shape[0] * 0.08:
        return None

    rows = []
    for top, bottom in zip(ys[:-1], ys[1:]):
        if bottom - top < 10:
            continue
        row = []
        for left, right in zip(xs[:-1], xs[1:]):
            if right - left < 15:
                continue
            inside = [w for w in words if w["x"] >= left - 3 and w["right"] <= right + 3
                      and w["y"] >= top - 4 and w["bottom"] <= bottom + 4]
            row.append(" ".join(w["text"] for w in sorted(inside, key=lambda z: z["x"])))
        if any(c for c in row):
            rows.append(row)
    if len(rows) < 2:
        return None
    # Remove empty trailing columns and rows.
    while rows and not any(rows[-1]):
        rows.pop()
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    while width > 1 and all(not r[-1] for r in rows):
        width -= 1
        rows = [r[:width] for r in rows]
    return rows if len(rows) >= 2 and width >= 2 else None


def _aligned_text_table(words):
    rows = _cluster_rows(words)
    if len(rows) < 2:
        return []
    # Use x-centers to infer columns. Large recurring gaps are stronger than
    # raw word positions and work for PDFs whose tables have no drawn borders.
    centers = []
    for row in rows:
        for w in row["words"]:
            centers.append(w["x"] + w["w"] / 2)
    if not centers:
        return []
    centers = sorted(centers)
    gaps = [(centers[i + 1] - centers[i], i) for i in range(len(centers) - 1)]
    median_gap = np.median([g for g, _ in gaps]) if gaps else 0
    split_gap = max(45, median_gap * 3.0)
    # Candidate column anchors are based on stable left edges across rows.
    x_positions = sorted([w["x"] for r in rows for w in r["words"]])
    anchors = []
    for x in x_positions:
        if not anchors or x - anchors[-1] > 28:
            anchors.append(x)
    if len(anchors) < 2:
        return []

    # Collapse anchors that are not repeated across enough rows. Then assign
    # each word to the nearest stable anchor.
    counts = []
    for a in anchors:
        count = sum(1 for r in rows if any(abs(w["x"] - a) <= 30 for w in r["words"]))
        counts.append(count)
    stable = [a for a, c in zip(anchors, counts) if c >= max(2, math.ceil(len(rows) * 0.18))]
    if len(stable) < 2:
        return []
    stable = sorted(stable)

    table = []
    for r in rows:
        cells = [""] * len(stable)
        for w in r["words"]:
            idx = min(range(len(stable)), key=lambda i: abs(w["x"] - stable[i]))
            # Reject an implausibly distant assignment unless this is the only
            # sensible column. This prevents paragraph text becoming a table.
            if abs(w["x"] - stable[idx]) > 90 and len(stable) > 2:
                continue
            cells[idx] = (cells[idx] + " " + w["text"]).strip()
        if any(cells):
            table.append(cells)
    return table


def _extract_native_page(page):
    words = page.get_text("words")
    if not words:
        return []
    normalized = []
    for x0, y0, x1, y1, text, block, line, word_no in words:
        text = _normalize_cell(text)
        if text:
            normalized.append({
                "text": text, "x": x0, "y": y0,
                "w": x1 - x0, "h": y1 - y0,
                "right": x1, "bottom": y1,
            })
    return _aligned_text_table(normalized)


def _ocr_page(page):
    image = _render_page(page)
    words = _ocr_data(image)
    grid = _table_from_grid(image, words)
    if grid:
        return grid, "ocr-grid"
    aligned = _aligned_text_table(words)
    return aligned, "ocr-positioned" if aligned else "ocr-text"


def _looks_like_table(rows):
    if not rows or len(rows) < 2:
        return False
    widths = [len(r) for r in rows]
    width = max(widths)
    if width < 2:
        return False
    nonempty_ratio = sum(bool(c) for r in rows for c in r) / max(1, len(rows) * width)
    return nonempty_ratio >= 0.25


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


def convert_pdf_to_xlsx(input_path, output_path):
    _ensure_tesseract()
    doc = pymupdf.open(str(input_path))
    wb = openpyxl.Workbook()
    default = wb.active
    wb.remove(default)

    text_sheet_rows = []
    tables = 0
    modes = []

    for page_no, page in enumerate(doc, start=1):
        native = _extract_native_page(page)
        if _looks_like_table(native):
            _write_sheet(wb, f"Table {tables + 1}", native)
            tables += 1
            modes.append("native")
            continue

        # For image/scanned pages, OCR is used. This is intentionally also
        # attempted on native PDFs when no convincing native table exists.
        rows, mode = _ocr_page(page)
        modes.append(mode)
        if _looks_like_table(rows):
            _write_sheet(wb, f"Table {tables + 1}", rows)
            tables += 1
        elif rows:
            text_sheet_rows.append([f"Page {page_no}"])
            text_sheet_rows.extend(rows)

    if text_sheet_rows:
        _write_sheet(wb, "Text", text_sheet_rows)
    if not wb.sheetnames:
        _write_sheet(wb, "Text", [["No extractable text or table was detected."]])

    wb.save(output_path)
    doc.close()
    return {"tables": tables, "mode": ",".join(sorted(set(modes))) or "none"}

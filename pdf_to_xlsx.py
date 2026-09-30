
import os
import re
import math
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter


def _safe_sheet_name(name, used):
    name = re.sub(r'[\[\]:*?/\\]', ' ', name).strip() or "Page"
    name = name[:31]
    base = name
    n = 2
    while name in used:
        suffix = f" ({n})"
        name = (base[:31-len(suffix)] + suffix)
        n += 1
    used.add(name)
    return name


def _numeric_value(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return ""
    # Keep date-like strings as text. Some PDFs can contain non-calendar values such as 2026-13.
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s
    if re.fullmatch(r"\d{4}-\d{2}", s):
        return s
    # Percentages
    m = re.fullmatch(r'(-?[\d,]+(?:\.\d+)?)\s*%', s)
    if m:
        return float(m.group(1).replace(",", "")) / 100.0
    # Currency / numeric
    cleaned = s.replace(",", "")
    if re.fullmatch(r'[-+]?\$?\d+(?:\.\d+)?', cleaned):
        if cleaned.startswith("$"):
            try:
                return float(cleaned[1:])
            except ValueError:
                return s
        try:
            n = float(cleaned)
            return int(n) if n.is_integer() else n
        except ValueError:
            pass
    return s


def _extract_images(pdf_path, page_index, temp_dir):
    """Return image records with x/y/w/h in PDF points and extracted file path."""
    records = []
    doc = fitz.open(pdf_path)
    page = doc[page_index]
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
            ext = data.get("ext", "png")
            path = os.path.join(temp_dir, f"p{page_index+1}_img{xref}.{ext}")
            with open(path, "wb") as f:
                f.write(data["image"])
        except Exception:
            continue
        for rect in rects:
            records.append({
                "path": path,
                "x0": float(rect.x0), "top": float(rect.y0),
                "x1": float(rect.x1), "bottom": float(rect.y1),
                "width": float(rect.width), "height": float(rect.height),
            })
    doc.close()
    return records


def _table_rows_and_cells(page):
    tables = page.find_tables()
    if not tables:
        return [], None
    table = tables[0]
    rows = table.extract()
    return rows, table


def _line_text_above_table(page, table_top):
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    words = [w for w in words if float(w["bottom"]) <= float(table_top) - 1]
    lines = {}
    for w in words:
        key = round(float(w["top"]) / 2) * 2
        lines.setdefault(key, []).append(w)
    out = []
    for y in sorted(lines):
        line = " ".join(w["text"] for w in sorted(lines[y], key=lambda x: x["x0"])).strip()
        if line:
            out.append(line)
    return out


def _cell_matrix(table):
    # find_tables().cells is a flat list in row-major order.
    # Build row groups by identical top coordinate.
    cells = table.cells
    if not cells:
        return []
    groups = {}
    for c in cells:
        groups.setdefault(round(c[1], 3), []).append(c)
    return [sorted(groups[y], key=lambda c: c[0]) for y in sorted(groups)]


def _image_in_cell(img, cell):
    ix = (img["x0"] + img["x1"]) / 2
    iy = (img["top"] + img["bottom"]) / 2
    return cell[0] <= ix <= cell[2] and cell[1] <= iy <= cell[3]


def convert_pdf_to_xlsx(input_path, output_path):
    input_path = str(input_path)
    output_path = str(output_path)
    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)
    used_names = set()
    thin = Side(style="thin", color="A6A6A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="263B5A")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=16, bold=True, color="263B5A")
    desc_font = Font(size=10, italic=True, color="666666")

    with tempfile.TemporaryDirectory() as temp_dir, pdfplumber.open(input_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            rows, table = _table_rows_and_cells(page)
            page_images = _extract_images(input_path, page_no - 1, temp_dir)

            # Use first meaningful line as sheet title.
            page_text = page.extract_text() or ""
            lines = [x.strip() for x in page_text.splitlines() if x.strip()]
            title = lines[0] if lines else f"Page {page_no}"
            ws = wb.create_sheet(_safe_sheet_name(f"Page {page_no} - {title}", used_names))
            ws.sheet_view.showGridLines = False

            if table and rows:
                pre = _line_text_above_table(page, table.bbox[1])
                # Put title/description above the table.
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(6, len(rows[0])))
                ws.cell(1, 1, title).font = title_font
                next_row = 2
                for line in pre[1:]:
                    if line == title:
                        continue
                    ws.merge_cells(start_row=next_row, start_column=1, end_row=next_row, end_column=max(6, len(rows[0])))
                    ws.cell(next_row, 1, line).font = desc_font
                    next_row += 1
                    if next_row >= 4:
                        break
                table_start = max(4, next_row + 1)
                # Write extracted table.
                for r, row in enumerate(rows, table_start):
                    for c, value in enumerate(row, 1):
                        cell = ws.cell(r, c, _numeric_value(value))
                        cell.border = border
                        cell.alignment = Alignment(vertical="center", wrap_text=True)
                        if r == table_start:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                # Reasonable widths.
                for c in range(1, len(rows[0]) + 1):
                    max_len = max([len(str(row[c-1] or "")) for row in rows[:20] if len(row) >= c] + [10])
                    ws.column_dimensions[get_column_letter(c)].width = min(max(max_len + 2, 11), 42)

                # Place images. If an image falls inside a table cell, anchor it there.
                cell_rows = _cell_matrix(table)
                for img in page_images:
                    target = None
                    tr = tc = None
                    for rr, row_cells in enumerate(cell_rows):
                        for cc, cellbox in enumerate(row_cells):
                            if _image_in_cell(img, cellbox):
                                tr, tc = rr, cc
                                target = cellbox
                                break
                        if target:
                            break
                    ximg = XLImage(img["path"])
                    # cap size while respecting aspect ratio
                    max_w = 125
                    max_h = 85
                    ratio = img["width"] / max(img["height"], 1)
                    w = min(max_w, max_h * ratio)
                    h = w / ratio
                    ximg.width = w
                    ximg.height = h
                    if target and tr is not None:
                        excel_row = table_start + tr
                        excel_col = tc + 1
                        ws.add_image(ximg, f"{get_column_letter(excel_col)}{excel_row}")
                        ws.row_dimensions[excel_row].height = max(ws.row_dimensions[excel_row].height or 15, h * 0.75 + 8)
                        ws.column_dimensions[get_column_letter(excel_col)].width = max(ws.column_dimensions[get_column_letter(excel_col)].width or 10, min(24, w / 6))
                    else:
                        # Standalone page image: place it to the right of the table.
                        anchor_col = max(len(rows[0]) + 2, 8)
                        ximg.width = min(360, max(180, img["width"] * 1.0))
                        ximg.height = ximg.width / ratio
                        ws.add_image(ximg, f"{get_column_letter(anchor_col)}2")
                        ws.column_dimensions[get_column_letter(anchor_col)].width = 18
                        ws.row_dimensions[2].height = max(ws.row_dimensions[2].height or 15, min(220, ximg.height * 0.75))
            else:
                # Text-only fallback: one line per row.
                ws.merge_cells("A1:F1")
                ws["A1"] = title
                ws["A1"].font = title_font
                for r, line in enumerate(lines[1:], 3):
                    ws.cell(r, 1, line)
                    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
                ws.column_dimensions["A"].width = 100

            ws.freeze_panes = "A5"
            ws.auto_filter.ref = None

        # Save while the temporary extracted image files still exist.
        wb.save(output_path)

    return {"mode": "table-aware-with-images", "tables": len(wb.worksheets)}

if __name__ == "__main__":
    import sys
    convert_pdf_to_xlsx(sys.argv[1], sys.argv[2])

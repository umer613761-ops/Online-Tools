ToolNest Convert to PDF restoration

Files:
- convert-to-pdf.html
- app.py

DOCX -> PDF and XLSX -> PDF use the Railway/LibreOffice backend routes:
- POST /api/docx-to-pdf
- POST /api/xlsx-to-pdf

The existing PDF -> XLSX route remains in app.py.
HTML -> PDF was not changed.

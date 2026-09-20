# ToolNest PDF to XLSX backend

This backend is designed for the Railway deployment used by ToolNest.

Key fixes:
- CORS enabled for the GitHub Pages frontend, preventing browser "Failed to fetch" errors.
- Root and /health endpoints return a simple running status.
- Tesseract OCR is installed inside the Docker image; visitors do not install anything.
- PDF pages are preserved as images inside the Excel workbook in addition to editable OCR/table data.
- Existing `/api/pdf-to-xlsx` endpoint is preserved.

ToolNest - Convert to PDF - Office conversion fix

This patch fixes DOCX/XLSX/PPTX conversion in convert-to-pdf.html.

Problem fixed:
The previous browser code treated Office files as plain text. DOCX/XLSX/PPTX are ZIP-based Office packages, so the result was the internal XML/ZIP data rendered into the PDF.

New behavior:
- DOCX, DOC, ODT, RTF, XLSX, XLS, ODS, CSV, PPTX, PPT and ODP are sent to /api/office-to-pdf.
- The backend uses LibreOffice headless to render the actual Office document to PDF.
- Multiple Office files can be uploaded together and are merged in upload order.
- Existing image/text/HTML browser conversion remains in place.

Backend requirements:
- Python 3
- Flask
- pypdf
- LibreOffice/soffice installed on the server

Frontend endpoint:
https://toolnest-api-production-dda7.up.railway.app/api/office-to-pdf

Important:
The new endpoint must be deployed to the ToolNest API server before the live website can convert DOCX/XLSX/PPTX. The browser page is already wired to that endpoint.

Local verification performed:
- Country Education Profiles-Pakistan.docx -> real PDF, 43 pages, A4.
- Panel University's Important Guidelines_VICPAK.xlsx -> real PDF, 24 pages, A4.
- Both outputs were visually rendered as documents/spreadsheets rather than ZIP/XML contents.

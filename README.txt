TOOLNEST CONVERT-TO-PDF RAILWAY FIX

Replace the deployment files in the existing ToolNest repository with:

- convert-to-pdf.html (updated frontend)
- backend/app.py (adds /api/office-to-pdf while retaining /api/pdf-to-xlsx)
- backend/requirements.txt
- Dockerfile

IMPORTANT:
The Dockerfile installs LibreOffice. This is required for DOCX/XLSX/PPTX/ODS/etc. rendering.

Expected repository structure:

Dockerfile
backend/
  app.py
  requirements.txt

The website's convert-to-pdf.html calls:
https://toolnest-api-production-dda7.up.railway.app/api/office-to-pdf

The backend accepts one or more Office files using multipart field name `files`, renders each
with LibreOffice, then merges the generated PDFs in upload order.

This patch was tested locally with the supplied DOCX and XLSX samples using LibreOffice.

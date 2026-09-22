ToolNest Convert to PDF - Railway root-layout fix

The Railway repository has app.py and requirements.txt at the repository root, not in backend/.
This package therefore uses root-level COPY instructions.

Files:
- Dockerfile
- app.py
- requirements.txt
- convert-to-pdf.html

Important: Do not delete existing repository modules such as pdf_to_xlsx.py. The Dockerfile copies the whole repo so the existing PDF-to-XLSX converter remains available.

After committing these files to the connected GitHub repository, trigger a new Railway deployment. Check /health and confirm libreoffice:true.

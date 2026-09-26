ToolNest Change PDF Page Size - Rotation Fix

Replace these two files in the project:
- app.py
- change-pdf-page-size.html

Dockerfile: no change required.

Fix: handles PDFs with portrait MediaBox + 90/270 degree page rotation correctly,
so pages such as Original.pdf page 2 remain landscape when converted to A4.

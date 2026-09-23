ToolNest Convert to PDF - XLSX Header Repeat Fix

Replace the ROOT app.py in the Railway/GitHub repository with this version.
Do not create a backend folder.

This change only affects modern Excel workbook conversion (.xlsx/.xlsm/.xltx/.xltm):
- detects the first likely table/header row on each worksheet
- repeats that row on continuation PDF pages
- keeps fit-to-width pagination enabled
- leaves the original uploaded workbook untouched

DOCX/PPTX/other Office conversion remains on the existing LibreOffice path.

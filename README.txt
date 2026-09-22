ToolNest Convert to PDF backend update

HTML -> PDF now uses the backend /api/html-to-pdf endpoint with WeasyPrint.
This preserves real PDF text, CSS page breaks, A4 @page rules, and repeating table headers.

Files:
- app-updated-cors.py  Updated Flask backend with /api/html-to-pdf
- convert-to-pdf.html  Updated frontend; HTML files are sent to the backend renderer
- requirements.txt      Runtime dependencies

Deploy the backend with WeasyPrint installed, then serve the updated convert-to-pdf.html.
If the API is on a different origin, set window.TOOLNEST_API_BASE to the API base URL before the page script runs. Otherwise the page uses /api automatically.

Important: HTML conversion is intentionally limited to one HTML file per conversion. Image/text conversion remains client-side and unchanged.

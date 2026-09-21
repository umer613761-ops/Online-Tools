# ToolNest PDF conversion backend

This backend provides the PDF conversion API used by the ToolNest frontend.

## Supported PDF conversions

- `POST /api/pdf-to-txt`
- `POST /api/pdf-to-docx`
- `POST /api/pdf-to-html`
- `GET /health`

The frontend can also convert PDF pages to JPG/PNG locally in the browser using PDF.js.

## Deployment

The Docker image installs Tesseract OCR and runs the Flask application with Gunicorn. The service listens on Railway's `PORT` environment variable (default `10000`).

The API returns JSON error details for conversion failures so the frontend can display the actual server-side error instead of hiding it behind a generic message.

## Upload limits

The default maximum PDF upload size is 50 MB. This can be changed with `MAX_UPLOAD_MB`.

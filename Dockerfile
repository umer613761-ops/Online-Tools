FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Force a fresh Railway image build for the Change PDF Page Size backend fix.
ARG TOOLNEST_BUILD=2026-09-26-page-size-fix
LABEL toolnest.build="${TOOLNEST_BUILD}"

RUN apt-get update \
    && apt-get install -y --no-install-recommends ghostscript tesseract-ocr tesseract-ocr-eng libgl1 libglib2.0-0 libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 2 --timeout 180 app:app"]

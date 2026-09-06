FROM python:3.12-slim

# OCI Metadata Labels
LABEL org.opencontainers.image.title="office-worker-mcp" \
      org.opencontainers.image.description="The complete document MCP for AI agents: Word, Excel, PowerPoint & PDF with professional styling" \
      org.opencontainers.image.url="https://github.com/CaFra-House/office-worker-mcp" \
      org.opencontainers.image.source="https://github.com/CaFra-House/office-worker-mcp" \
      org.opencontainers.image.licenses="MIT" \
      org.opencontainers.image.vendor="CaFra-House"

# Avoid interactive prompts during package install
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    HOME=/tmp

# System dependencies:
# - LibreOffice headless minimal (writer + calc + impress for docx/xlsx/pptx -> pdf conversions)
# - Tesseract OCR + Spanish language data
# - WeasyPrint C dependencies (pango, cairo, gdk-pixbuf, ffi, mime)
# - Poppler utilities (pdftoppm preview)
# - Liberation fonts (ensures metric-compatible Arial/Times/Courier rendering in LibreOffice)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer-nogui \
    libreoffice-calc-nogui \
    libreoffice-impress-nogui \
    tesseract-ocr \
    tesseract-ocr-spa \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    poppler-utils \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Non-root user setup
RUN useradd -m -u 10001 -s /bin/bash appuser \
    && mkdir -p /data \
    && chown -R appuser:appuser /data

# Install office-worker-mcp with full feature set extras:
# [pdf] for WeasyPrint premium HTML->PDF rendering
# [sign] for cryptographic PDF signature & verification
# [book] for EPUB multi-chapter export
# [ocr] for pytesseract + Pillow scanned PDF & image text extraction
# [pptx] for html-to-pptx (native PPTX charts fall back gracefully without Chromium)
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir ".[sign,book,ocr,pptx,pdf]"

# Set working directory to default persistent output volume
WORKDIR /data
VOLUME ["/data"]

USER appuser

# Healthcheck validating python environment and capabilities status
HEALTHCHECK --interval=30s --timeout=10s --retries=3 CMD owi doctor || exit 1

# Default entrypoint runs the MCP stdio server
ENTRYPOINT ["office-worker-mcp"]

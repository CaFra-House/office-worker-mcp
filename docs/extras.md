# Optional Extras & Platform Notes

## Optional pip extras

The core package works with zero optional dependencies. Install extras only when you need the specific capability:

```bash
pip install office-worker-mcp            # core: MCP server + CLI 'owi' (no extras)
pip install "office-worker-mcp[pdf]"     # render_document premium HTML→PDF via WeasyPrint
pip install "office-worker-mcp[ocr]"     # OCR over scans/images via Tesseract + Pillow
pip install "office-worker-mcp[sign]"    # real PAdES cryptographic signing via pyhanko
pip install "office-worker-mcp[book]"   # EPUB export for create_book via ebooklib
pip install "office-worker-mcp[pptx]"   # native PowerPoint charts via Playwright/Chromium
```

- `[pdf]` enables `render_document` premium HTML→PDF via WeasyPrint (requires native Pango/Cairo libraries on bare metal; install them with your system package manager on Linux/macOS, or use Docker / WSL2 on Windows).
- `[ocr]` enables `pdf_ocr` (searchable text layer over scanned PDFs and images).
- `[sign]` enables real PAdES signatures in `sign_pdf` and multi-engine verification in `verify_pdf_signature`.
- `[book]` enables parallel `.epub` output from `create_book`.
- `[pptx]` is required by `create_pptx` (Playwright + Chromium; run `playwright install chromium`).

## Platform limitations (honest)

For full transparency about what the tools do and do not guarantee:

- **Windows is excluded from the CI matrix.** WeasyPrint needs natively compiled C libraries (`libpango`, `libpangocairo`, `libgdk-pixbuf`, `libffi`) with no official unattended Windows installer, and LibreOffice headless requires registry-path binaries. On Windows, use the **official multi-arch Docker image** (`ghcr.io/cafra-house/office-worker-mcp`, amd64 + arm64 — see [docker.md](docker.md)) or run under **WSL2 (Ubuntu)**.
- **LibreOffice headless font substitution.** Converting to PDF may substitute fonts if exact corporate typefaces are not installed on the host. Wide spreadsheets without a fixed print area can get automatic page breaks - always reported in `warnings`.
- **VBA macros are preserved but NOT executed.** `openpyxl` keeps `.xlsm` macros safely via `keep_vba=True`; no VBA code is ever run.
- **`sign_pdf` without a certificate stamps only the visible PNG seal** - it does not embed a cryptographic signature. In that case `verify_pdf_signature` honestly reports `has_signature: false`, `valid: false`. Pass `cert_pem` or set `auto_generate_test_cert=true` for a real verifiable PAdES signature.
- **`document_diff` is a textual difflib comparison**, not a legal-grade semantic redline and not Word OOXML track-changes. It emits an explicit warning to that effect.
- **`add_pivot` produces a static consolidated pandas view** styled as an Excel sheet - it does not insert Microsoft's binary OLAP pivot cache.
- **`create_book` TOC** uses WeasyPrint CSS Paged Media GCPM (`target-counter(attr(href), page)`) to resolve real page numbers; EPUB export requires the `[book]` extra.
- **`design_mode="premium"`** applies local packaged editorial CSS (`PREMIUM_CSS`) - zero external design/AI API calls, zero network, zero added token cost.
- **`owi doctor` / `environment_status`** inspect binaries and Python libraries read-only and return exact per-OS install commands (`apt`/`brew`/`dnf`) without executing privileged commands or modifying the system.


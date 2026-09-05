---
name: office-worker
description: "Use when the agent must create, edit, convert, manipulate, sign, compress, redact, flatten, preview or read office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX). The Office Worker MCP: 21 tools for batch pipelines (office_batch), in-place edits (edit_excel, edit_word), packaged templates (docxtpl), native Excel charts, structured tables, visual PNG previews (pdf_preview), irreversible redaction (pdf_redact), structured Markdown/JSON extraction for RAG (pdf_extract_structured), smart document splitting (split_smart), PDF flattening (flatten), PDF all-in-one reader (read_pdf flags), PDF-to-Excel (pdf_to_excel), Office reader (read_office), AcroForm filling, OCR, compression, signatures, and conversion. Deterministic one-path-per-task with hard anti-loop limits."
---

# The Office Worker — Agent Operating Manual (v0.6.0)

Generates, edits, secures, and transforms professional corporate office documents (**PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX**) using the `office-worker` MCP server. 100% local, zero external API keys, fully private, deterministic, and safe.

The Golden Rule: **One tool per format, one deterministic path per task.**

---

## Hard Rules (Anti-Loop & Operational Discipline)

1. **One tool per task:** Pick the EXACT tool corresponding to your target format and objective. Never guess or improvise incompatible workflows.
2. **Hard Anti-Loop Limit (2-Strike Stop):** If any tool call returns `{"status": "error"}` twice consecutively on the same task, **YOU MUST IMMEDIATELY STOP**. Do NOT attempt a 3rd blind variation or alternate syntax. Report the exact error string (`error`) honestly to the user and ask for clarification. Blind retries cause context degradation and hallucinations.
3. **Always verify on disk:** When a tool finishes, verify that the returned `bytes` is > 0. If a generated file is 0 bytes, the operation failed regardless of status.
4. **Visual inspection before delivery (`pdf_preview`):** Before handing a generated or signed PDF over to the user, call `pdf_preview(input_path=..., max_pages=1, dpi=110)`. The tool returns a Base64 data URL (`data:image/png;base64,...`) for your multimodal vision inspection.
5. **Packaged templates first:** For standard business documents (meeting minutes, executive reports, simple invoices, formal letters, audit checklists), call `list_templates()` first, then invoke `create_word(out_path=..., template_docx="<name>", context={...})`.
6. **In-place modifications:** Never regenerate entire files from scratch just to alter a few values. Use `edit_excel` or `edit_word` to preserve existing fonts, layout, and VBA macros (`keep_vba`).
7. **Multi-step pipelines with `office_batch`:** When chaining document operations (e.g. generate Word -> convert to PDF -> sign -> preview), run them in a single call to `office_batch`.
8. **Irreversible redaction (`pdf_redact`):** When sanitizing confidential information, use `pdf_redact`. PyMuPDF permanently destroys underlying text and vector paths.
9. **Form flattening (`pdf_manipulate op="flatten"`):** Before archiving or distributing completed AcroForms, flatten them to freeze inputs and prevent subsequent tampering.
10. **Structured extraction for RAG (`pdf_extract_structured`):** When feeding documents into knowledge bases or LLM context, use `pdf_extract_structured` with `format="markdown"` (pipe tables) or `format="json"`.

---

## Tool Directory (21 Specialized Tools)

| Tool | Returns | Primary Use Case | When NOT to use |
|---|---|---|---|
| `pdf_redact` | JSON `{status, path, bytes, redactions_count}` | Permanently removes sensitive PII/passwords via text search or coordinates | Do NOT use for standard text editing (use `edit_word`) |
| `pdf_extract_structured` | JSON `{status, format, content?, pages, n_tables}` | Extracts text, tables (Markdown pipes / arrays), and metadata for RAG workflows | Do NOT use for scanned images without text (use `pdf_ocr`) |
| `pdf_preview` | JSON `{status, data_url, pages, path?, bytes?}` | Renders PDF pages to PNG (Base64 data URL) for multimodal vision verification | Do NOT use for extracting selectable text (use `read_pdf`) |
| `read_pdf` | JSON `{pages, metadata, tables?, fields?, images?}` | All-in-one PDF inspection: text, tables, form fields, and Base64 images | Do NOT use on non-digital scanned PDFs (use `pdf_ocr`) |
| `pdf_manipulate` | JSON `{status, path/files, bytes, warnings?}` | Merges, rotates, extracts ranges, flattens forms (`flatten`), or smart-splits (`split_smart`) | Do NOT use for editing internal paragraph text |
| `edit_excel` | JSON `{status, path, fidelity, warnings}` | In-place updates: cells, rows, columns, tables, auto-filters, formulas, native charts | Do NOT use to build spreadsheets from nothing (use `create_excel`) |
| `edit_word` | JSON `{status, path, fidelity, warnings}` | In-place updates: paragraphs, text replacement, heading inserts, tables | Do NOT use to create brand-new files from scratch (use `create_word`) |
| `office_batch` | JSON `{status, total, succeeded, failed, results}` | Chains multi-document pipelines in 1 roundtrip with fault isolation | Do NOT use for simple one-off commands |
| `pdf_to_excel` | JSON `{status, path, n_tables, fidelity, warnings}` | Extracts structured PDF balance/pricing tables directly to styled `.xlsx` | Do NOT use for text-heavy narrative PDFs without tables |
| `read_office` | JSON `{status, format, paragraphs/slides/sheets, text}` | Extracts text and structure directly from `.docx`, `.pptx`, and `.xlsx` | Do NOT use on PDF files (use `read_pdf`) |
| `render_document` | JSON `{status, path, bytes}` | Renders high-end PDFs from Jinja HTML templates, corporate themes, logos, watermarks | Do NOT use for editable Office formats (`.docx`, `.xlsx`, `.pptx`) |
| `create_word` | JSON `{status, path, bytes}` | Generates new `.docx` files from blocks or official packaged templates | Do NOT use to edit existing documents (use `edit_word`) |
| `create_excel` | JSON `{status, path, bytes}` | Generates styled multi-sheet `.xlsx` workbooks with auto-filters and native charts | Do NOT use to update existing workbooks (use `edit_excel`) |
| `create_pptx` | JSON `{status, path, bytes}` | Generates editable PowerPoint decks (native text, cards, tables, badges) | Do NOT use for long narrative documents or balance sheets |
| `convert_to_pdf` | JSON `{status, path, fidelity, warnings}` | Converts Office files (`.docx`, `.xlsx`, `.pptx`) to PDF via headless LibreOffice | Do NOT use when generating documents from HTML templates |
| `sign_pdf` | JSON `{status, path, bytes}` | Stamps visual PNG signature/seal and applies cryptographic PAdES digital signature | Do NOT use to alter document text |
| `pdf_compress` | JSON `{status, path, bytes, savings_percent}` | Shrinks heavy PDFs by downsampling images and cleaning unused objects | Do NOT use on text-only PDFs with no images |
| `pdf_fill_form` | JSON `{status, path, bytes}` | Fills interactive AcroForm fields deterministically with key-value data | Do NOT use on flat PDFs without AcroForm fields |
| `pdf_ocr` | JSON `{status, text, path?}` | Performs Tesseract OCR on scanned documents or images, generating searchable PDF | Do NOT use on digital PDFs with existing text |
| `list_templates` | JSON `{status, templates}` | Inspects official packaged `.docx` templates and required context variables | Do NOT use to generate documents directly |
| `list_themes` | JSON `{status, theme}` | Inspects corporate palettes and typography configurations | Do NOT use to generate files directly |

---

## Conversational Workflows & Examples (EN)

### 1. Invoice Workflow (`factura_simple` & `render_document`)
*User:* "Generate an official invoice for client Acme Corp for USD 4,500 consulting services with corporate styling."
*Agent:*
```json
{
  "out_path": "invoice_acme.docx",
  "template_docx": "factura_simple",
  "context": {
    "numero_factura": "INV-2026-0089",
    "fecha": "2026-09-05",
    "emisor_nombre": "Enterprise Soluciones S.A.",
    "emisor_cuit": "30-71234567-9",
    "cliente_nombre": "Acme Corp S.A.",
    "items": [
      {"descripcion": "Senior AI Architecture Consulting (40 hrs)", "cantidad": 40, "precio_unitario": 100, "total": 4000},
      {"descripcion": "Cloud Infrastructure Setup", "cantidad": 1, "precio_unitario": 500, "total": 500}
    ],
    "subtotal": 4500,
    "iva": 945,
    "total": 5445
  }
}
```

### 2. Executive Report Workflow (`informe_ejecutivo` & `create_word`)
*User:* "Prepare an executive summary report for Q3 performance."
*Agent:*
```json
{
  "out_path": "q3_executive_report.docx",
  "template_docx": "informe_ejecutivo",
  "context": {
    "titulo": "Q3 Performance Executive Summary",
    "subtitulo": "Regional Sales & Operational Milestones",
    "fecha": "September 2026",
    "autor": "Chief Operations Officer",
    "resumen_ejecutivo": "Revenue exceeded quarterly forecasts by 14% driven by enterprise AI adoption.",
    "secciones": [
      {"titulo": "Financial Highlights", "contenido": "Gross ARR reached $12.4M with 88% gross margin."},
      {"titulo": "Operational Goals", "contenido": "All ISO 27001 audit controls verified compliant."}
    ],
    "conclusiones": "Maintain aggressive expansion into European markets."
  }
}
```

### 3. Board Meeting Minutes (`acta_meeting`)
*User:* "Draft board meeting minutes approving the annual budget."
*Agent:*
```json
{
  "out_path": "board_minutes_2026_09.docx",
  "template_docx": "acta_meeting",
  "context": {
    "titulo": "Minutes of the Executive Board N° 128",
    "fecha": "2026-09-05",
    "hora": "14:30",
    "lugar": "Main Boardroom / Virtual",
    "asistentes": [
      {"nombre": "Julio Cardozo", "rol": "Chairman"},
      {"nombre": "Elena Rostova", "rol": "Chief Legal Counsel"}
    ],
    "puntos": [
      {"orden": 1, "tema": "FY2027 Budget Allocation", "discusion": "Budget reviewed and unanimously endorsed."}
    ],
    "acuerdos": [
      {"acuerdo": "Approve CapEx of $1.5M", "responsable": "JC", "fecha_limite": "2026-10-01"}
    ],
    "firmas": [
      {"nombre": "Julio Cardozo", "cargo": "Chairman"}
    ]
  }
}
```

### 4. Contract Revision (`edit_word`)
*User:* "Update our vendor agreement contract.docx: replace the counterparty name and append an audit clause."
*Agent:* Call `edit_word`:
```json
{
  "input_path": "contract.docx",
  "operations": [
    {"op": "replace_text", "find": "COUNTERPARTY_LEGAL_NAME", "replace": "Global Logistics Corp"},
    {"op": "append_paragraph", "text": "Clause 14 (Audit Rights): Either party may audit security compliance upon 5 days notice.", "bold": true}
  ]
}
```

### 5. Audit Checklist (`checklist_auditoria` & `create_excel`)
*User:* "Create an ISO 27001 audit checklist in Excel with autofilters."
*Agent:* Call `create_excel`:
```json
{
  "out_path": "iso27001_audit.xlsx",
  "title": "ISO 27001 Compliance Audit",
  "sheets_json": "[{\"name\": \"Controls\", \"headers\": [\"Control ID\", \"Requirement\", \"Status\", \"Auditor Note\"], \"rows\": [[\"A.9.1\", \"Access Control Policy\", \"Compliant\", \"Verified MFA\"], [\"A.12.1\", \"Operational Procedures\", \"Compliant\", \"Runbooks updated\"]]}]",
  "auto_filter": true,
  "table_style": "TableStyleMedium2"
}
```

### 6. Batch Multi-Step Pipeline (`office_batch`)
*User:* "Create our quarterly report in Word, convert it to PDF, sign it with our digital stamp, and generate a preview image."
*Agent:*
```json
{
  "operations": [
    {"tool": "create_word", "args": {"out_path": "report.docx", "template_docx": "informe_ejecutivo", "context": {"titulo": "Quarterly Report"}}},
    {"tool": "convert_to_pdf", "args": {"input_file": "report.docx", "output": "report.pdf"}},
    {"tool": "sign_pdf", "args": {"input_pdf": "report.pdf", "output": "report_signed.pdf", "sello_img_path": "seal.png"}},
    {"tool": "pdf_preview", "args": {"input_path": "report_signed.pdf", "max_pages": 1, "dpi": 110}}
  ]
}
```

### 7. In-Place Spreadsheet Editing (`edit_excel`)
*User:* "In budget.xlsx, update cell C4 to 12500, append a new Q4 projection row, and add a bar chart."
*Agent:*
```json
{
  "input_path": "budget.xlsx",
  "operations": [
    {"op": "set_cell", "sheet": "Sheet1", "coordinate": "C4", "value": 12500},
    {"op": "append_row", "sheet": "Sheet1", "row": ["Q4 Projection", 18200]},
    {"op": "add_chart", "sheet": "Sheet1", "chart_type": "bar", "title": "Quarterly Trends", "target_cell": "E2"}
  ]
}
```

### 8. Structured Extraction for RAG (`pdf_extract_structured`)
*User:* "Extract the balance sheet tables and sections from annual_report.pdf to Markdown so I can load it into our knowledge base."
*Agent:*
```json
{
  "input_path": "annual_report.pdf",
  "format": "markdown",
  "output": "annual_report.md"
}
```
*Result:* Clean GitHub Markdown with `# Title`, `## Page N`, and formatted `| Col | Col |` pipe tables ready for semantic chunking.

### 9. Permanent Confidential Redaction (`pdf_redact`)
*User:* "Censor all occurrences of customer SSN 001-23-4567 and the account box at coordinates [100, 150, 280, 180] in tax_filing.pdf."
*Agent:*
```json
{
  "input_path": "tax_filing.pdf",
  "output": "tax_filing_redacted.pdf",
  "search_text": "001-23-4567",
  "regions": [
    {"page": 1, "x0": 100, "y0": 150, "x1": 280, "y1": 180}
  ],
  "fill_color": "black"
}
```
*Fidelity Guarantee:* The redacted content is irreversibly purged from the PDF stream via PyMuPDF.

### 10. Form Flattening & Archival Compliance (`pdf_manipulate op="flatten"`)
*User:* "Flatten the filled registration form so the fields cannot be edited anymore."
*Agent:*
```json
{
  "operation": "flatten",
  "input_path": "filled_registration.pdf",
  "output": "archived_registration.pdf"
}
```

### 11. Smart Document Splitting (`pdf_manipulate op="split_smart"`)
*User:* "This 50-page PDF scanned_batch.pdf contains several contracts separated by blank separator pages. Split them automatically."
*Agent:*
```json
{
  "operation": "split_smart",
  "input_path": "scanned_batch.pdf",
  "output": "contract_part.pdf"
}
```
*Result:* Returns `{"files": ["contract_part_1.pdf", "contract_part_2.pdf", ...], "n_splits": 2}`. If boundaries are uncertain, returns honest warnings without fabricating arbitrary cuts.

### 12. Digital Signing & Optimization (`sign_pdf` & `pdf_compress`)
*User:* "Compress this heavy scanned report and stamp our corporate approval seal on the last page."
*Agent:*
```json
{
  "input_path": "raw_scan.pdf",
  "output": "compressed.pdf",
  "quality": "med"
}
```
Followed by:
```json
{
  "input_pdf": "compressed.pdf",
  "output": "approved_document.pdf",
  "sello_img_path": "approval_stamp.png",
  "reason": "Audited and Approved"
}
```

---

## Honest Fidelity Framework

Tools producing derivative documents report `fidelity` transparently:
- `"rich"`: Bit-accurate native file generation or in-place edit preserving formatting and metadata.
- `"clean"`: Standard high-fidelity transformation (e.g. headless LibreOffice or PDF tabular conversion). Includes explicit `warnings` for minor pagination or font shifts.
- `"lossy"`: Fallback or partial extraction when original vector layout is missing.

Always inspect `warnings` in the JSON response before claiming perfect fidelity to the user.

---
name: office-worker-orchestration
description: "Intent-to-tools workflow orchestration for The Office Worker. Maps business intentions (factura, informe, acta, contrato, auditoria, batch, redact-flow, sign-flow, rag-ingest, compliance-flow) directly to deterministic tool chains with strict anti-loop limits."
---

# The Office Worker — Orchestration & Intent Workflows (v0.9.0)

Directs high-level business goals into deterministic, single-turn or pipelined MCP tool chains. 100% local, zero cloud dependencies, private, and deterministic.

---

## Core Intent-to-Tools Routing Map

| Business Intent | Primary Tools Sequence | Output & Verification |
|---|---|---|
| **Factura (Invoice)** | `render_document(design_mode="premium")` or `create_word(template_docx="factura_simple")` &rarr; `convert_to_pdf` | Final PDF invoice; verify via `pdf_preview` |
| **Informe (Executive Report)** | `create_word(template_docx="informe_ejecutivo")` or `create_pptx` &rarr; `convert_to_pdf` | High-fidelity executive deck or docx/pdf |
| **Libro / Manual (Multi-chapter Book)** | `create_book(title, chapters=[...], theme, epub=True)` &rarr; `pdf_preview` | Multi-chapter PDF with automated TOC (real pages) + EPUB |
| **Análisis Dinámico (Pivot Tables)** | `create_excel` or `edit_excel(operations=[{op: "add_pivot", ...}])` | Excel workbook with styled pivot sheet, sum/count/avg, and auto-filters |
| **Acta (Meeting Minutes)** | `create_word(template_docx="acta_meeting")` &rarr; `convert_to_pdf` &rarr; `sign_pdf` | Approved, timestamped board minutes |
| **Contrato (Contract & Agreement)** | `create_word(template_docx="carta_formal")` or `mail_merge` &rarr; `edit_word` &rarr; `document_diff` &rarr; `sign_pdf` &rarr; `verify_pdf_signature` | Revision-tracked, cryptographically signed legal contract |
| **Auditoría (Audit Checklist)** | `create_excel(table_style="TableStyleMedium9", auto_filter=True)` or `create_word(template_docx="checklist_auditoria")` &rarr; `convert_to_pdf` | Structured checklist with active filters and corporate palette |
| **Batch (Pipelines)** | `office_batch(operations=[...])` | Executes up to 20 operations in 1 roundtrip without intermediate chat turns |
| **Redact Flow (Sanitization)** | `pdf_redact` (permanent redaction) &rarr; `scrub_metadata` (author, editor, history wipe) &rarr; `pdf_preview` | Sanitized document with zero PII leaks or metadata traces |
| **Sign Flow (Digital Signature)** | `sign_pdf` (sello PNG + cert PEM o `auto_generate_test_cert=True`) &rarr; `verify_pdf_signature` (audit cryptographic validity) | Verifiable PAdES digital signature |
| **RAG Ingest (Knowledge Base)** | `read_office(format="markdown")` for Office or `pdf_extract_structured(format="markdown")` for PDF | Markdown with headings `#` and pipe tables `\|` for LLM embeddings |
| **Compliance Flow (Enterprise Redline & Protection)** | `document_diff` (detect differences) &rarr; `scrub_metadata` (clean personal data) &rarr; `protect_office` (AES password lock) &rarr; `verify_pdf_signature` | Comprehensive enterprise compliance and audit trail |
| **Doctor (Environment Diagnostics)** | `environment_status()` or CLI `owi doctor` | Audits system binaries/python libs and outputs exact OS install commands |

---

## Intent Workflows Explained

### 1. Factura (Invoice Flow)
* **Goal:** Generate a professional, audit-ready invoice with calculation breakdown.
* **Flow:**
  1. Call `list_templates()` to verify required keys for `factura_simple`.
  2. Invoke `create_word(out_path="invoice.docx", template_docx="factura_simple", context={...})` or `render_document(template_html=..., out_path="invoice.pdf", data_json=...)`.
  3. If Word was generated, call `convert_to_pdf(input_file="invoice.docx", output="invoice.pdf")`.
  4. Call `pdf_preview(input_path="invoice.pdf", max_pages=1)` to visually inspect before delivery.

### 2. Informe (Executive Report Flow)
* **Goal:** Present quarterly performance or technical findings to stakeholders.
* **Flow:**
  1. For text/narrative reports: `create_word(out_path="report.docx", template_docx="informe_ejecutivo", context={...})`.
  2. For visual presentations: `create_pptx(out_path="deck.pptx", slides_json=...)` with native bar/line/pie charts.
  3. Export to PDF via `convert_to_pdf` and inspect layout via `pdf_preview`.

### 3. Acta (Board / Team Minutes Flow)
* **Goal:** Record attendees, decisions, action items, and required sign-offs.
* **Flow:**
  1. Call `create_word(out_path="acta.docx", template_docx="acta_meeting", context={...})`.
  2. Convert to PDF via `convert_to_pdf`.
  3. Apply visual endorsement via `sign_pdf(input_pdf="acta.pdf", output="acta_signed.pdf", sello_img_path="seal.png")`.

### 4. Contrato (Contract & Revision Tracking Flow)
* **Goal:** Draft, revise, track changes, and execute binding contracts.
* **Flow:**
  1. Draft base contract with `create_word` or `mail_merge`.
  2. Apply amendments in place using `edit_word(operations=[{"op": "replace_text", ...}])`.
  3. Compare versions using `document_diff(path_a="contract_v1.docx", path_b="contract_v2.docx")` to inspect added, deleted, and modified clauses.
  4. Convert to PDF with `convert_to_pdf`.
  5. Sign cryptographically with `sign_pdf(cert_pem="cert.pem")`.
  6. Audit signature integrity with `verify_pdf_signature(input="contract_signed.pdf")`.

### 5. Auditoría (Audit Checklist Flow)
* **Goal:** Track ISO, IT security, or financial compliance items.
* **Flow:**
  1. Call `create_excel(out_path="audit.xlsx", sheets_json=[...], table_style="TableStyleMedium9", auto_filter=True)`.
  2. If updates are needed, patch in place via `edit_excel(operations=[{"op": "set_cell", ...}])`.

### 6. Batch (Multi-Step Pipeline Flow)
* **Goal:** Chain document generation, conversion, redaction, and signing without intermediate turns.
* **Flow:**
  1. Construct operations array for `office_batch`.
  2. Call `office_batch(operations=[{"tool": "create_word", ...}, {"tool": "convert_to_pdf", ...}, {"tool": "sign_pdf", ...}])`.
  3. Check results array for granular step status.

### 7. Redact Flow (Sanitization & Privacy Flow)
* **Goal:** Scrub PII, confidential figures, or trade secrets before external transmission.
* **Flow:**
  1. Black out sensitive text/coordinates in PDF via `pdf_redact(search_text="SECRET", fill_color="#000000")`.
  2. Scrub underlying author, editor, and revision metadata via `scrub_metadata(input="doc.pdf", output="doc_clean.pdf")`.
  3. Verify clean appearance with `pdf_preview`.

### 8. Sign Flow (Digital Signature & Endorsement)
* **Goal:** Cryptographically seal documents and verify existing signatures.
* **Flow:**
  1. Apply visual stamp and PAdES X.509 signature via `sign_pdf`.
  2. Audit signature validity, intactness, and certificate details via `verify_pdf_signature`.

### 9. RAG Ingest (Knowledge Base Extraction Flow)
* **Goal:** Ingest multi-format office documents into vector stores or LLM contexts.
* **Flow:**
  1. For `.docx`, `.pptx`, `.xlsx`: call `read_office(path=..., format="markdown")` to extract structural markdown headings and pipe tables.
  2. For `.pdf`: call `pdf_extract_structured(input_path=..., format="markdown")`.
  3. Feed extracted markdown blocks into context without lossy external OCR unless working with scanned images.

### 10. Compliance Flow (Enterprise Audit & Protection)
* **Goal:** End-to-end governance: compare revisions, scrub personal data, password-protect files, and verify digital signatures.
* **Flow:**
  1. Compare revisions: `document_diff(path_a="draft_v1.docx", path_b="draft_v2.docx")`.
  2. Clean identifying metadata: `scrub_metadata(input="draft_v2.docx", output="draft_clean.docx")`.
  3. Lock file with AES encryption: `protect_office(input="draft_clean.docx", output="draft_locked.docx", password="SecurePassword!")`.
  4. For PDFs, audit digital signatures: `verify_pdf_signature(input="signed_audit.pdf")`.

---

## Anti-Loop & Operational Discipline

1. **Follow `next_steps`:** When a creation tool returns `"next_steps"`, prioritize the suggested follow-up tool to complete the business lifecycle.
2. **2-Strike Stop:** If a tool call fails 2 consecutive times with `{"status": "error"}`, STOP immediately. Report the exact error message and ask the user for clarification.
3. **Verify File Bytes:** Always verify that returned `bytes` is > 0.
4. **Honest Warnings:** Always communicate warnings returned by `document_diff`, `convert_to_pdf`, or `verify_pdf_signature` (e.g. self-signed certificates or textual diff limitations).

# Conversational Workflows

Ten end-to-end examples of how an agent drives The Office Worker in natural conversation. Each shows the user request, the exact tool call (JSON), and the result.

### Flujo 1: Bulk Mail Merge (Generating Employee Welcome Letters)
*User:* "I have a CSV with 50 new hires and a Word template. Generate personalized welcome letters for all of them."
*Agent:*
```json
{
  "template_path": "templates/welcome_letter.docx",
  "dataset_csv": "data/new_hires.csv",
  "output_prefix": "output/letters/welcome"
}
```
*Result:* Returns `{"status": "ok", "n_docs": 50, "paths": [".../welcome_1.docx", ...]}`.

### Flujo 2: Ingesting Office Spreadsheets into RAG Context
*User:* "Read our quarterly financial balance sheet and give me the markdown table so I can analyze it."
*Agent:*
```json
{
  "path": "reports/q3_balance.xlsx",
  "format": "markdown"
}
```
*Result:* Returns clean Markdown pipe tables per sheet with exact numeric values.

### Flujo 3: Executive Deck with Native PowerPoint Charts
*User:* "Create an executive presentation with 2 slides: title slide and a quarterly sales bar chart."
*Agent:*
```json
{
  "out_path": "presentations/sales_q3.pptx",
  "slides_json": [
    {
      "title": "Quarterly Business Review",
      "kicker": "Executive Summary",
      "bullets": ["Q3 Key Performance Indicators", "Next Steps"]
    },
    {
      "title": "Revenue Growth 2026",
      "kicker": "Performance",
      "chart": {
        "type": "bar",
        "title": "Quarterly Revenue ($M)",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "values": [12.4, 14.2, 16.8, 19.5]
      }
    }
  ]
}
```

### Flujo 4: Comparing Contract Revisions (`document_diff`)
*User:* "Compare agreement_draft_v1.docx and agreement_draft_v2.docx and tell me what changed."
*Agent:*
```json
{
  "path_a": "agreement_draft_v1.docx",
  "path_b": "agreement_draft_v2.docx",
  "format": "markdown"
}
```
*Result:* Returns markdown summary with modified/added clauses and transparent warnings about textual difflib boundaries.

### Flujo 5: Scrubbing Sensitive Metadata before External Dispatch (`scrub_metadata`)
*User:* "Scrub all author, editor, and revision metadata from executive_report.docx before we email it to partners."
*Agent:*
```json
{
  "input": "executive_report.docx",
  "output": "executive_report_clean.docx"
}
```
*Result:* Returns `{"status": "ok", "scrubbed_fields": ["author", "comments", "last_modified_by", "revision", "title"]}`.

### Flujo 6: Password-Protecting Spreadsheets (`protect_office`)
*User:* "Lock compensation_matrix.xlsx with password 'HR-Enterprise-2026!'."
*Agent:*
```json
{
  "input": "compensation_matrix.xlsx",
  "output": "compensation_matrix_encrypted.xlsx",
  "password": "HR-Enterprise-2026!"
}
```
*Result:* Returns `{"status": "ok", "encrypted": true}` producing standard agile AES encrypted spreadsheet.

### Flujo 7: Cryptographic PDF Digital Signature & Verification (`sign_pdf` & `verify_pdf_signature`)
*User:* "Sign this vendor contract with an automated test certificate and verify that the signature is valid."
*Agent:* Call `sign_pdf` with `auto_generate_test_cert=True`:
```json
{
  "input_pdf": "contracts/vendor_contract.pdf",
  "output": "contracts/vendor_contract_signed.pdf",
  "auto_generate_test_cert": true,
  "reason": "Approved and Signed",
  "location": "Buenos Aires"
}
```
*Result:* Returns `{"status": "ok", "path": ".../vendor_contract_signed.pdf", "bytes": 84520, "next_steps": ["Verify digital signature with verify_pdf_signature", "Preview signed document with pdf_preview"]}`.

Then verify with `verify_pdf_signature`:
```json
{
  "input": "contracts/vendor_contract_signed.pdf"
}
```
*Result:* Returns `{"status": "ok", "has_signature": true, "valid": true, "intact": true, "signer": "Office Worker Test Certificate (Non-Production)", "reason": "Approved and Signed", "warnings": ["Certificate is self-signed or not anchored in a trusted system certificate authority."]}`.

### Flujo 8: Multi-Chapter Technical Manual with Automatic TOC (`create_book`)
*User:* "Generate a 3-chapter security manual in PDF and EPUB format with a cover page, table of contents, and corporate theme."
*Agent:* Call `create_book`:
```json
{
  "output": "manuals/security_architecture.pdf",
  "title": "Enterprise Security Architecture",
  "author": "Security Engineering Team",
  "theme": "corporate-blue",
  "epub": true,
  "chapters": [
    {
      "title": "Threat Modeling",
      "content_html": "<p>Identification of assets, threat actors, and attack vectors.</p>"
    },
    {
      "title": "Zero Trust Network Architecture",
      "content_html": "<p>Microsegmentation, continuous authentication, and least privilege access.</p>"
    },
    {
      "title": "Cryptographic Controls",
      "content_html": "<p>Key rotation, AES-256 storage, and TLS 1.3 enforcement.</p>"
    }
  ]
}
```
*Result:* Returns `{"status": "ok", "path": "manuals/security_architecture.pdf", "chapters_count": 3, "epub_path": "manuals/security_architecture.epub", "next_steps": ["Inspect book layout with pdf_preview", "Verify digital signature with verify_pdf_signature"]}`.

### Flujo 9: Excel Pivot Table Aggregation (`edit_excel` with `add_pivot`)
*User:* "Take our raw sales sheet in sales_q3.xlsx and generate a pivot table sheet summing revenue by region and product line."
*Agent:* Call `edit_excel`:
```json
{
  "input_path": "sales_q3.xlsx",
  "operations": [
    {
      "op": "add_pivot",
      "sheet": "RawData",
      "rows": "Region",
      "cols": "ProductLine",
      "values": "Revenue",
      "agg": "sum",
      "pivot_sheet": "Pivot_Sales_Summary"
    }
  ]
}
```
*Result:* Returns `{"status": "ok", "path": "sales_q3.xlsx", "fidelity": "rich", "sheets_modified": ["Pivot_Sales_Summary"]}` with styled headers, alternating rows, borders, and auto-filters.

### Flujo 10: Environment Health Check & Missing Dependencies (`environment_status`)
*User:* "Check if this machine has all binaries and Python libraries needed for document generation and OCR."
*Agent:* Call `environment_status`:
```json
{}
```
*Result:* Returns `{"status": "ok", "os": "linux", "package_manager": "apt", "core_ready": true, "all_ready": true, "capabilities": {...}}`. If any dependency is missing, returns exact copy-paste command (e.g., `sudo apt-get install -y libreoffice`).

---


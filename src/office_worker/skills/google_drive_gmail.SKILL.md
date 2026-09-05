---
name: office-worker-google
description: "Optional workflow extension: Upload office-worker generated documents (PDF, Word, Excel, PowerPoint) to Google Drive or compose Gmail drafts with attached files using the workspace-mcp backend. If Google Workspace is not configured or desired, this skill can be safely ignored."
---

# Google Drive & Gmail Workflow for Office Worker (Optional)

This optional workflow guides agents on how to connect **The Office Worker** (100% local, zero-API-key document generator) with **workspace-mcp** (Google Workspace backend) to upload generated files to Google Drive or email them via Gmail.

> **Privacy & Architecture Note:**
> The core `office-worker-mcp` package remains 100% local, zero-API-key, and self-contained. OAuth2 and cloud tokens are never bundled into the core library. This skill acts purely as an agent-level coordination pattern. If you or your organization do not use Google Workspace, you can safely ignore this skill.

---

## Prerequisites

1. The core `office-worker-mcp` server is running and configured in your agent.
2. The `workspace-mcp` server (or equivalent Google integration) is configured with user authentication.

---

## Common Workflows

### Workflow 1: Generate Document Locally & Upload to Google Drive

When the user asks to save a newly generated report or spreadsheet into a specific Google Drive folder:

1. **Generate the document locally with Office Worker:**
   ```json
   {
     "tool": "create_word",
     "args": {
       "out_path": "/tmp/informe_mensual.docx",
       "template_docx": "informe_ejecutivo",
       "context": {
         "titulo": "Informe Mensual de Gestión",
         "fecha": "2026-09-05"
       }
     }
   }
   ```
2. **Convert to PDF or compress locally if needed:**
   ```json
   {
     "tool": "convert_to_pdf",
     "args": {
       "input_file": "/tmp/informe_mensual.docx",
       "output": "/tmp/informe_mensual.pdf"
     }
   }
   ```
3. **Upload to Google Drive using `workspace-mcp`:**
   Invoke the `workspace-mcp` tool (e.g. `google_drive_upload_file` or `drive_upload`):
   ```json
   {
     "file_path": "/tmp/informe_mensual.pdf",
     "folder_name": "Reportes 2026",
     "mime_type": "application/pdf"
   }
   ```
4. **Report confirmation to user:**
   Share the generated local path, the Google Drive link, and the file size.

---

### Workflow 2: Compose a Gmail Draft with Generated Document

When the user asks to email an invoice, signed contract, or audit report:

1. **Inspect document visually with `pdf_preview`:**
   ```json
   {
     "tool": "pdf_preview",
     "args": {
       "input_path": "/tmp/factura_acme.pdf",
       "max_pages": 1
     }
   }
   ```
2. **Draft email in Gmail via `workspace-mcp`:**
   Call the Gmail tool (e.g. `gmail_create_draft` or `gmail_send_message`):
   ```json
   {
     "to": "billing@acmecorp.com",
     "subject": "Factura N° 0089 - Servicios Profesionales CaFra",
     "body": "Estimados,\n\nAdjunto la factura correspondiente al mes en curso.\n\nSaludos cordiales,\nEquipo de Finanzas",
     "attachments": ["/tmp/factura_acme.pdf"]
   }
   ```
3. **Notify user:**
   Confirm draft creation without sending automatically, allowing the user to review before final dispatch.

---

### Workflow 3: Chat Delivery (Host Agent Behavior)

In chat-based agents (e.g. Hermes over Telegram, Discord, or Web Chat):
- The host agent automatically forwards local generated files (PDF, DOCX, XLSX) into the chat window.
- No special tools are needed; simply inform the user of the final file path and summary metrics.

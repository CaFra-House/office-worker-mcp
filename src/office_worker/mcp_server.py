"""The Office Worker — MCP server (cara delgada sobre office_worker.core).

Meta: herramientas poderosas y especializadas (~19 aquí), bajo overhead de contexto (<1700 tok).
Cada tool crea/lee un archivo real y devuelve {status, path, ...} o {status:"error", error}.
Regla anti-loop en instructions: si una tool falla 2 veces seguidas, NO reintentar con variantes.
"""
from __future__ import annotations
import os, json
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from office_worker.core import (
    render_pdf as _render_pdf,
    create_word as _create_word,
    edit_word as _edit_word,
    mail_merge as _mail_merge,
    create_excel as _create_excel,
    edit_excel as _edit_excel,
    csv_excel_convert as _csv_excel_convert,
    create_pptx as _create_pptx,
    read_pdf as _read_pdf,
    pdf_preview as _pdf_preview,
    pdf_extract_structured as _pdf_extract_structured,
    fill_pdf_form as _fill_pdf_form,
    ocr_pdf as _ocr_pdf,
    convert_office_to_pdf as _convert_office_to_pdf,
    manipulate_pdf as _manipulate_pdf,
    sign_pdf as _sign_pdf,
    compress_pdf as _compress_pdf,
    pdf_redact as _pdf_redact,
    pdf_to_excel as _pdf_to_excel,
    read_office as _read_office,
    list_packaged_templates as _list_packaged_templates,
    document_diff as _document_diff,
    scrub_metadata as _scrub_metadata,
    protect_office as _protect_office,
    verify_pdf_signature as _verify_pdf_signature,
    create_book as _create_book,
    check_environment as _check_environment,
    load_theme, DEFAULT_THEME,
)

mcp = FastMCP(
    "office-worker",
    instructions=(
        "The Office Worker: generates, edits, and secures professional office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX) "
        "locally with corporate styling, packaged Word templates (docxtpl), bulk mail merge, multi-chapter books and EPUB (create_book), "
        "in-place edits, Excel pivot tables (add_pivot), native charts, formulas, bidirectional CSV-Excel conversion, "
        "Office-to-Markdown/JSON structured extraction, document comparison/diff, metadata scrubbing, password protection, digital signature verification, "
        "environment diagnostics (environment_status), batch pipelines, PDF processing (preview PNG, permanent redaction, structured RAG extraction, flattening, smart split, text, tables, forms, OCR, compression, signature), and Office-to-PDF conversion. "
        "Choose the exact tool matching your target document format. "
        "Hard Anti-Loop Rule: if any tool call fails 2 consecutive times, STOP immediately and report the exact error to the user."
    ),
)


def _ok(path: str, next_steps: list[str] | None = None) -> dict:
    d = {"status": "ok", "path": os.path.abspath(path), "bytes": os.path.getsize(path)}
    if next_steps:
        d["next_steps"] = next_steps
    return d


@mcp.tool()
def render_document(
    template_html: str,
    out_path: str,
    data_json: str = "{}",
    theme: str = "",
    logo: str = "",
    password: str = "",
    watermark_text: str = "",
    footer_left: str = "",
    footer_right: str = "",
    page_numbers: bool = True,
    design_mode: str = "standard",
) -> dict:
    """Generates professional PDF document (invoice, report, executive summary, letter, balance statement) from Jinja HTML template, corporate theme, data, and design options (watermark, header logo, footers, page numbers, encryption, design_mode='standard'|'premium'). Returns JSON with status, absolute file path, and byte size. When to use: Use when creating polished, high-fidelity PDFs from structured data and HTML/CSS templates. When NOT to use: Do NOT use for editing existing PDF/Office files (use edit_word/edit_excel/pdf_manipulate) or when pure Office formats (.docx, .xlsx, .pptx) are required. Keywords: invoice, report, letter, contract, executive summary, balance statement, template, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        data = json.loads(data_json or "{}") or {}
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"data_json inválido: {e}"}
    if isinstance(data.get("rows"), list):
        headers = data.get("headers") or []
        thead = "".join(f"<th>{h}</th>" for h in headers)
        body = "\n".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in data["rows"])
        data["tabla"] = f"<table><thead><tr>{thead}</tr></thead><tbody>{body}</tbody></table>"
    try:
        return _ok(
            _render_pdf(
                template_html,
                out_path,
                data=data,
                theme=theme or None,
                logo=logo or None,
                password=password or None,
                watermark_text=watermark_text,
                footer_left=footer_left,
                footer_right=footer_right,
                page_numbers=page_numbers,
                design_mode=design_mode or "standard",
            ),
            next_steps=["Preview visual layout with pdf_preview", "Sign document with sign_pdf"],
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def create_word(
    out_path: str,
    title: str = "",
    subtitle: str = "",
    blocks_json: str = "[]",
    theme: str = "",
    template_docx: str = "",
    context: dict = {},
) -> dict:
    """Creates professional Word .docx document (minutes, letter, contract, report, audit checklist) from declarative blocks or packaged templates (acta_meeting, informe_ejecutivo, factura_simple, carta_formal, checklist_auditoria). Returns JSON with status, absolute file path, and byte size. When to use: Use to generate new .docx documents with consistent corporate theme or standard templates with {{ variables }}. When NOT to use: Do NOT use to modify existing .docx files in place (use edit_word instead) or to produce PDFs directly (use render_document or convert_to_pdf). Keywords: minutes, letter, contract, report, audit, checklist, invoice, executive summary, template, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        blocks = json.loads(blocks_json or "[]") if isinstance(blocks_json, str) else (blocks_json or [])
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"blocks_json inválido: {e}"}
    if isinstance(context, str):
        try:
            ctx = json.loads(context or "{}") if context else {}
        except json.JSONDecodeError as e:
            return {"status": "error", "error": f"context inválido: {e}"}
    else:
        ctx = dict(context or {})
    try:
        return _ok(
            _create_word(out_path, title=title, subtitle=subtitle or None, blocks=blocks, theme=theme or None, template_docx=template_docx or None, context=ctx),
            next_steps=["Convert to PDF with convert_to_pdf", "Preview rendered document with pdf_preview"],
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def mail_merge(
    template_path: str,
    dataset_csv: str = "",
    dataset_json: str = "",
    output_prefix: str = "",
    fields: list = [],
) -> dict:
    """Generates N .docx documents from docxtpl template with {{placeholders}} fed by CSV or JSON dataset. Returns JSON with status, n_docs, and paths. When to use: Use for personalized contracts, letters, or reports in bulk. When NOT to use: Do NOT use for single documents (use create_word). Keywords: mail merge, template, dataset, bulk, csv, json, local, safe."""
    try:
        f_list = fields
        if isinstance(fields, str):
            try:
                f_list = json.loads(fields or "[]")
            except Exception:
                f_list = [f.strip() for f in fields.split(",") if f.strip()]
        res = _mail_merge(
            template_path=template_path,
            dataset_csv=dataset_csv,
            dataset_json=dataset_json,
            output_prefix=output_prefix,
            fields=f_list or None,
        )
        if isinstance(res, dict) and res.get("status") == "ok":
            res["next_steps"] = ["Convert generated documents to PDF with convert_to_pdf", "Preview first generated document with pdf_preview"]
        return res
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def create_excel(
    out_path: str,
    title: str = "",
    sheets_json: str = "[]",
    theme: str = "",
    table_style: str = "",
    auto_filter: bool = False,
) -> dict:
    """Creates professional multi-sheet Excel .xlsx workbook (balance statement, financial report, audit checklist) with corporate styling, optional structured tables, auto-filters, formulas, and native charts (bar, line, pie). Returns JSON with status, absolute file path, and byte size. When to use: Use to generate new spreadsheets with tables, headers, zebra-striping, and charts. When NOT to use: Do NOT use to modify existing workbooks in place (use edit_excel instead) or for VBA macros (openpyxl does not run macros). Keywords: balance statement, report, audit, checklist, chart, formulas, local, private, no api key, offline, cross-platform, no office install needed, deterministic, safe."""
    try:
        sheets = json.loads(sheets_json or "[]") if isinstance(sheets_json, str) else (sheets_json or [])
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"sheets_json inválido: {e}"}
    try:
        return _ok(
            _create_excel(out_path, title=title, sheets=sheets or None, theme=theme or None, table_style=table_style or None, auto_filter=auto_filter),
            next_steps=["Convert to PDF with convert_to_pdf", "Export or edit with csv_excel_convert or edit_excel"],
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def create_pptx(out_path: str, slides_json: str = "[]", theme: str = "") -> dict:
    """Creates editable PowerPoint .pptx presentation (deck, executive summary, status report) with corporate theme and optional native charts (bar, line, pie). Returns JSON with status, absolute file path, and byte size. When to use: Use to generate editable PowerPoint slides with typography, kickers, bullets, tables, and native charts. When NOT to use: Do NOT use for PDFs (use render_document). Keywords: executive summary, report, deck, presentation, chart, local, private, safe."""
    try:
        slides = json.loads(slides_json or "[]") if isinstance(slides_json, str) else (slides_json or [])
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"slides_json inválido: {e}"}
    try:
        return _ok(
            _create_pptx(out_path, slides=slides or None, theme=theme or None),
            next_steps=["Convert to PDF with convert_to_pdf", "Preview slides with pdf_preview"],
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def read_pdf(
    path: str,
    max_pages: int = 0,
    extract_tables: bool = False,
    list_forms: bool = False,
    extract_images: bool = False,
    max_images: int = 10,
) -> dict:
    """Reads PDF text, metadata, tables (extract_tables=True), interactive AcroForm fields (list_forms=True), and embedded images as base64 data URLs (extract_images=True, max_images=10) for agent vision. All-in-one PDF inspection tool. Returns JSON with text, metadata, tables, form fields, and base64 images. When to use: Use as primary tool to inspect and analyze any PDF document. When NOT to use: Do NOT use for scanned image PDFs without digital text (use pdf_ocr). Keywords: read, text, table, form, images, vision, local, safe."""
    try:
        return _read_pdf(
            path,
            max_pages=max_pages or None,
            extract_tables=extract_tables,
            list_forms=list_forms,
            extract_images=extract_images,
            max_images=max_images,
        ) | {"status": "ok"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_preview(
    input_path: str,
    output: str = "",
    max_pages: int = 0,
    dpi: int = 110,
) -> dict:
    """Renders PDF pages to PNG images via PyMuPDF (fitz) returning base64 data URLs for agent vision inspection before final delivery, with optional file saving. Returns JSON with status, data_url, rendered pages, dimensions, and optional output path. When to use: Use when the agent needs to visually inspect or verify a generated PDF layout, signature, or stamp before delivering it to the user. When NOT to use: Do NOT use for text extraction (use read_pdf) or scanned OCR (use pdf_ocr). Keywords: preview, png, image, vision, inspect, review, render, base64, local, safe."""
    try:
        return _pdf_preview(
            input_path,
            output=output or None,
            max_pages=max_pages or None,
            dpi=dpi or 110,
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def list_themes(theme_name_or_path: str = "") -> dict:
    """Returns corporate color palette and typography configuration for consistent document styling (aden, claro, oscuro, minimal, corporate-blue). Returns JSON with theme details. When to use: Use to inspect active color codes and font settings before creating documents. When NOT to use: Do NOT use to create documents directly. Keywords: theme, style, colors, report, local, private, offline."""
    t = load_theme(theme_name_or_path or None) if theme_name_or_path else dict(DEFAULT_THEME)
    return {"status": "ok", "theme": t}


@mcp.tool()
def pdf_fill_form(input_pdf: str, output: str, fields: dict = {}) -> dict:
    """Fills interactive AcroForm form fields in a PDF (contract, tax form, registration) with key-value data. Returns JSON with status, filled file path, and byte size. When to use: Use to fill standardized PDF forms deterministically. When NOT to use: Do NOT use for flat or scanned PDFs without AcroForm fields (use render_document or pdf_ocr). Keywords: fill form, contract, report, audit, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        f = json.loads(fields) if isinstance(fields, str) else (fields or {})
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"fields inválido: {e}"}
    try:
        return _ok(_fill_pdf_form(input_pdf, f, output))
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_ocr(input_path: str, lang: str = "spa", output: str = "", max_pages: int = 0) -> dict:
    """Performs optical character recognition (OCR) on scanned documents (invoice, receipt, minutes, audit checklist) or images using Tesseract, returning extracted text and optionally generating a searchable PDF. Returns JSON with text and output path. When to use: Use on scanned documents lacking digital text layers. When NOT to use: Do NOT use on digital PDFs with existing selectable text (use read_pdf instead). Keywords: ocr, scan, invoice, minutes, receipt, audit, convert, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        text = _ocr_pdf(input_path, lang=lang, out=output or None, max_pages=max_pages or None)
        if output:
            res = _ok(output)
            res["text"] = text
            return res
        return {"status": "ok", "text": text, "lang": lang}
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def convert_to_pdf(input_file: str, output: str) -> dict:
    """Converts Office documents (.docx, .xlsx, .pptx: contract, report, letter, balance statement) to PDF locally using headless LibreOffice, reporting honest fidelity (clean/lossy) and conversion warnings. Returns JSON with status, output path, byte size, fidelity rating, and warnings list. When to use: Use to export Office files to PDF without external cloud APIs. When NOT to use: Do NOT use when creating documents from scratch (use render_document). Keywords: convert, contract, letter, report, balance statement, local, private, no api key, offline, cross-platform, linux, mac, windows, deterministic, safe."""
    try:
        res_path = _convert_office_to_pdf(input_file, output)
        res = _ok(res_path)
        ext = Path(input_file).suffix.lower()
        warnings = []
        if ext in (".docx", ".doc"):
            warnings.append("LibreOffice headless conversion may have minor font substitution or pagination shifts.")
            fidelity = "clean"
        elif ext in (".xlsx", ".xls"):
            warnings.append("LibreOffice headless converts worksheets without configured print areas into automatic page breaks.")
            fidelity = "clean"
        elif ext in (".pptx", ".ppt"):
            warnings.append("Slide transitions, animations, and embedded media are flattened.")
            fidelity = "clean"
        else:
            warnings.append(f"Format {ext} converted with best-effort standard filter.")
            fidelity = "lossy"
        res["fidelity"] = fidelity
        res["warnings"] = warnings
        res["next_steps"] = ["Preview visual layout with pdf_preview", "Sign document with sign_pdf"]
        return res
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_manipulate(
    operation: str,
    output: str,
    input_path: str = "",
    files: list = [],
    pages: str = "",
    angle: int = 90,
    password: str = "",
) -> dict:
    """Manipulates PDF documents: merge files, extract pages (split_by), split_smart by doc boundaries, flatten forms/annots, or rotate pages. Returns JSON with status, output path, and details. When to use: Use for PDF page reorganization, flattening, and splitting. When NOT to use: Do NOT use for text editing. Keywords: merge, split, flatten, rotate, local, safe."""
    try:
        file_list = files
        if isinstance(files, str):
            try:
                file_list = json.loads(files)
            except Exception:
                file_list = [f.strip() for f in files.split(",") if f.strip()]
        res = _manipulate_pdf(
            operation=operation,
            out=output,
            input_path=input_path or None,
            files=file_list or None,
            pages=pages or None,
            angle=angle,
            password=password or None,
        )
        if isinstance(res, dict):
            return res
        return _ok(res)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def list_templates() -> dict:
    """Lists packaged Word templates (acta_meeting, informe_ejecutivo, factura_simple, carta_formal, checklist_auditoria) with descriptions, keywords, and expected variable schemas. Returns JSON with template catalog. When to use: Call before create_word to inspect available official templates and required context keys. When NOT to use: Do NOT use to generate files (use create_word with template_docx). Keywords: template, minutes, report, invoice, letter, audit, checklist, local, private, offline."""
    return {"status": "ok", "templates": _list_packaged_templates()}


@mcp.tool()
def pdf_compress(input_path: str, output: str, quality: str = "med") -> dict:
    """Compresses and optimizes PDF documents (report, scan, balance statement) by downsampling embedded images and cleaning unused objects via PyMuPDF. Returns JSON with status, output path, before/after byte size, and savings percentage. When to use: Use to shrink large PDFs before emailing or archiving. When NOT to use: Do NOT use on text-only PDFs where image optimization yields no benefit. Keywords: compress, optimize, report, scan, balance statement, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        return _compress_pdf(input_path, output, quality=quality)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def sign_pdf(
    input_pdf: str,
    output: str,
    sello_img_path: str = "",
    cert_pem: str = "",
    key_pem: str = "",
    passphrase: str = "",
    reason: str = "",
    location: str = "",
    auto_generate_test_cert: bool = False,
) -> dict:
    """Signs PDF (contract, invoice, letter): stamps visual seal and applies cryptographic PAdES signature if cert_pem or auto_generate_test_cert=True. Returns JSON with status, path, bytes, next_steps. When to use: Use to sign or approve documents. When NOT to use: Do NOT use to edit content. Keywords: sign, pades, seal, contract, local, safe."""
    try:
        return _ok(
            _sign_pdf(
                input_pdf=input_pdf,
                output=output,
                sello_img_path=sello_img_path or None,
                cert_pem=cert_pem or None,
                key_pem=key_pem or None,
                passphrase=passphrase or None,
                reason=reason or None,
                location=location or None,
                auto_generate_test_cert=bool(auto_generate_test_cert),
            ),
            next_steps=["Verify digital signature with verify_pdf_signature", "Preview signed document with pdf_preview"],
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def edit_excel(
    input_path: str,
    operations: list = [],
    output_path: str = "",
) -> dict:
    """Modifies an existing Excel workbook (.xlsx / .xlsm: balance statement, financial report, audit) in place or to a new file via openpyxl, preserving styles, formatting, and VBA macros (keep_vba). Supports set_cell, append_row, add_column, add_chart (bar, line, pie), add_table, auto_filter, add_pivot (pandas pivot tables on new sheet), and formulas (SUM, SUMIF, AVERAGEIF, VLOOKUP, XLOOKUP, COUNTIFS). Returns JSON with status, path, fidelity, warnings, and operations count. When to use: Use to update existing spreadsheets without recreating them from scratch. When NOT to use: Do NOT use to build new spreadsheets from nothing (use create_excel) or to run VBA macros. Keywords: edit in place, balance statement, report, audit, formula, chart, local, private, no api key, offline, cross-platform, no office install needed, deterministic, safe."""
    try:
        return _edit_excel(input_path, operations=operations, output_path=output_path or None)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def edit_word(
    input_path: str,
    operations: list = [],
    output_path: str = "",
) -> dict:
    """Modifies an existing Word document (.docx: contract, letter, minutes, report, audit checklist) in place or to a new file via python-docx, preserving styles. Supports append_paragraph, replace_text, insert_after_heading, and append_table. Returns JSON with status, path, fidelity, warnings, and operations count. When to use: Use to revise, update, or patch existing .docx files. When NOT to use: Do NOT use to create new documents from scratch (use create_word) or to edit PDFs directly. Keywords: edit in place, contract, letter, minutes, report, audit, checklist, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        return _edit_word(input_path, operations=operations, output_path=output_path or None)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_to_excel(
    input_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    max_pages: int = 0,
) -> dict:
    """Extracts structured tabular data from PDF files (balance statement, invoice, audit report) into a clean Excel .xlsx workbook via pdfplumber and openpyxl with honest fidelity reporting. Returns JSON with status, path, tables count, pages processed, fidelity, and warnings. When to use: Use when tabular data locked in PDF reports needs to be exported to Excel for analysis. When NOT to use: Do NOT use for scanned image PDFs without selectable tables (use pdf_ocr) or unstructured text PDFs. Keywords: convert, balance statement, invoice, report, audit, table, local, private, no api key, offline, cross-platform, no office install needed, deterministic, safe."""
    try:
        return _pdf_to_excel(input_path, output_path, sheet_name=sheet_name or "Sheet1", max_pages=max_pages or None)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def read_office(
    path: str,
    max_rows: int = 500,
    format: str = "",
) -> dict:
    """Extracts text, slides, and tables from Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) files in Markdown (format='markdown') or JSON (format='json', default). Returns JSON with format, structured elements, and content. When to use: Use to read or ingest Office files into RAG/LLM context. When NOT to use: Do NOT use for PDF (use read_pdf). Keywords: read, markdown, json, rag, table, local, safe."""
    try:
        return _read_office(path, max_rows=max_rows, format=format or "json")
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def csv_excel_convert(
    input: str,
    output: str,
    direction: str = "",
    sheet: str = "",
) -> dict:
    """Converts bidirectionally between CSV and structured Excel .xlsx workbooks. Supports single sheet or all sheets. Returns JSON with status, path, rows, and warnings. When to use: Use to convert CSV to styled Excel or export spreadsheets to CSV. When NOT to use: Do NOT use for editing existing spreadsheets (use edit_excel). Keywords: csv, excel, convert, xlsx, table, sheet, local, safe."""
    try:
        return _csv_excel_convert(
            input_path=input,
            output_path=output,
            direction=direction or "csv_to_xlsx",
            sheet=sheet,
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_redact(
    input_path: str,
    output: str,
    search_text: str = "",
    regions: list = [],
    fill_color: str = "",
) -> dict:
    """Permanently removes sensitive info from PDF by text search or coordinate regions via PyMuPDF apply_redactions. Returns JSON with status, path, bytes, and redactions_count. When to use: Use to erase PII or confidential data. When NOT to use: Do NOT use to edit text. Keywords: redact, censor, privacy, security, pii, sanitize, local, safe."""
    try:
        reg_list = regions
        if isinstance(regions, str):
            try:
                reg_list = json.loads(regions or "[]")
            except Exception:
                reg_list = []
        res = _pdf_redact(
            input_path=input_path,
            output=output,
            search_text=search_text,
            regions=reg_list,
            fill_color=fill_color or None,
        )
        if isinstance(res, dict) and res.get("status") == "ok":
            res["next_steps"] = ["Preview redacted document with pdf_preview", "Scrub document metadata with scrub_metadata"]
        return res
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def pdf_extract_structured(
    input_path: str,
    format: str = "markdown",
    output: str = "",
) -> dict:
    """Extracts structured text, tables (pipe format or arrays), and metadata from PDF pages to Markdown or JSON for RAG workflows. Returns JSON with format, content, and page structures. When to use: Use to ingest PDF documents and tables into RAG or LLM knowledge bases. When NOT to use: Do NOT use for scanned image PDFs (use pdf_ocr). Keywords: rag, extract, markdown, json, table, structure, knowledge, local, safe."""
    try:
        return _pdf_extract_structured(
            input_path=input_path,
            format=format or "markdown",
            output=output or None,
        )
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def document_diff(
    path_a: str,
    path_b: str,
    format: str = "json",
) -> dict:
    """Compares two Word (.docx) or PDF documents returning honest textual differences (added, deleted, modified paragraphs) via difflib. Returns JSON with status, changes summary, diffs list, and textual warnings. When to use: Use to compare revisions or contract versions. When NOT to use: Do NOT use for legal-grade redlines or pixel diffs. Keywords: diff, compare, changes, revision, contract, local, safe."""
    try:
        return _document_diff(path_a=path_a, path_b=path_b, format=format or "json")
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def scrub_metadata(
    input: str,
    output: str,
    fields: list = [],
) -> dict:
    """Removes sensitive metadata (author, title, revision history, editors, custom properties) from PDF, Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) documents. Returns JSON with status, path, scrubbed fields, and byte size. When to use: Call before sharing documents externally. When NOT to use: Do NOT use to redact visual document content (use pdf_redact). Keywords: scrub, sanitize, metadata, privacy, author, clean, local, safe."""
    try:
        f_list = fields
        if isinstance(fields, str):
            try:
                f_list = json.loads(fields or "[]")
            except Exception:
                f_list = [f.strip() for f in fields.split(",") if f.strip()]
        return _scrub_metadata(input_path=input, output=output, fields=f_list or None)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def protect_office(
    input: str,
    output: str,
    password: str,
) -> dict:
    """Protects Office documents (.docx, .xlsx, .pptx) with standard agile AES encryption password via msoffcrypto. Returns JSON with status, output path, and byte size. When to use: Use to password-protect sensitive workbooks, presentations, or documents. When NOT to use: Do NOT use for PDFs (use render_document or pdf_manipulate). Keywords: protect, password, encrypt, confidential, office, lock, local, safe."""
    try:
        return _protect_office(input_path=input, output=output, password=password)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def verify_pdf_signature(
    input: str,
) -> dict:
    """Verifies cryptographic digital signatures in a PDF file via pyhanko and pypdf, returning honest validation status. Returns JSON with has_signature, valid, signer, date, reason, and warnings. When to use: Use to audit or verify signed contracts and invoices. When NOT to use: Do NOT use to sign documents (use sign_pdf). Keywords: verify signature, pades, audit, contract, integrity, local, safe."""
    try:
        return _verify_pdf_signature(input_path=input)
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def create_book(
    output: str,
    title: str,
    author: str = "",
    chapters: list = [],
    theme: str = "",
    epub: bool = False,
) -> dict:
    """Generates multi-chapter PDF books with cover, automatic numbered TOC with page numbers, and professional margins via WeasyPrint, with optional EPUB export via ebooklib. Returns JSON with status, path, and chapter count. When to use: Use for long multi-chapter manuals, books, or reports. When NOT to use: Do NOT use for single-page documents (use render_document). Keywords: book, epub, manual, toc, chapters, pdf, local, safe."""
    if isinstance(chapters, str):
        try:
            ch_list = json.loads(chapters or "[]")
        except json.JSONDecodeError as e:
            return {"status": "error", "error": f"chapters JSON inválido: {e}"}
    else:
        ch_list = list(chapters or [])
    try:
        res = _create_book(
            out_path=output,
            title=title,
            author=author or "",
            chapters=ch_list,
            theme=theme or None,
            epub=bool(epub),
        )
        if isinstance(res, dict) and res.get("status") == "ok":
            res["next_steps"] = ["Preview book layout with pdf_preview", "Inspect table of contents with read_pdf"]
        return res
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def environment_status() -> dict:
    """Audits system environment and reports active capabilities (LibreOffice, Tesseract OCR, WeasyPrint, PyMuPDF, Pandas) and exact installation commands for missing tools. Returns JSON with status, OS, capabilities, and install hints. When to use: Call to check installed document processing engines on host. When NOT to use: Do NOT use to process documents. Keywords: doctor, status, environment, system, check, install, diagnose, local, safe."""
    try:
        return _check_environment()
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def office_batch(
    operations: list = [],
) -> dict:
    """Executes multiple document operations sequentially in a single turn without round-trips; continues on error and reports individual results. Returns JSON with total, succeeded, failed count, and step results list. When to use: Use to execute multi-step document pipelines (e.g. generate report + convert to PDF + sign). When NOT to use: Do NOT use for single operations or when later steps require conversational branching. Keywords: batch, pipeline, workflow, invoice, report, minutes, letter, contract, audit, checklist, executive summary, balance statement, convert, sign, local, private, no api key, offline, cross-platform, deterministic, safe."""
    if isinstance(operations, str):
        try:
            ops = json.loads(operations or "[]") or []
        except json.JSONDecodeError as e:
            return {"status": "error", "error": f"operations JSON inválido: {e}"}
    else:
        ops = list(operations or [])

    dispatcher = {
        "render_document": render_document,
        "create_word": create_word,
        "edit_word": edit_word,
        "mail_merge": mail_merge,
        "create_excel": create_excel,
        "edit_excel": edit_excel,
        "csv_excel_convert": csv_excel_convert,
        "create_pptx": create_pptx,
        "read_pdf": read_pdf,
        "pdf_preview": pdf_preview,
        "list_themes": list_themes,
        "pdf_fill_form": pdf_fill_form,
        "pdf_ocr": pdf_ocr,
        "convert_to_pdf": convert_to_pdf,
        "pdf_manipulate": pdf_manipulate,
        "list_templates": list_templates,
        "pdf_compress": pdf_compress,
        "sign_pdf": sign_pdf,
        "pdf_to_excel": pdf_to_excel,
        "read_office": read_office,
        "pdf_redact": pdf_redact,
        "pdf_extract_structured": pdf_extract_structured,
        "document_diff": document_diff,
        "scrub_metadata": scrub_metadata,
        "protect_office": protect_office,
        "verify_pdf_signature": verify_pdf_signature,
        "create_book": create_book,
        "environment_status": environment_status,
    }

    results = []
    succeeded = 0
    failed = 0

    for i, op in enumerate(ops):
        tool_name = op.get("tool") or op.get("name")
        args = op.get("args") or op.get("arguments") or {}

        if not tool_name or tool_name not in dispatcher:
            results.append({
                "index": i,
                "tool": tool_name,
                "status": "error",
                "error": f"Tool desconocida o no soportada: '{tool_name}'. Disponibles: {list(dispatcher.keys())}",
            })
            failed += 1
            continue

        fn = dispatcher[tool_name]
        try:
            res = fn(**args)
            if isinstance(res, dict) and res.get("status") == "error":
                results.append({"index": i, "tool": tool_name, "status": "error", "error": res.get("error")})
                failed += 1
            else:
                results.append({"index": i, "tool": tool_name, "status": "ok", "result": res})
                succeeded += 1
        except Exception as exc:
            results.append({"index": i, "tool": tool_name, "status": "error", "error": str(exc)})
            failed += 1

    status = "ok" if failed == 0 else ("partial_error" if succeeded > 0 else "error")
    return {
        "status": status,
        "total": len(ops),
        "succeeded": succeeded,
        "failed": failed,
        "results": results,
    }


def main():
    mcp.run()


if __name__ == "__main__":
    main()

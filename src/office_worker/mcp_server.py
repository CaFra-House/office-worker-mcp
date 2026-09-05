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
    create_excel as _create_excel,
    edit_excel as _edit_excel,
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
    load_theme, DEFAULT_THEME,
)

mcp = FastMCP(
    "office-worker",
    instructions=(
        "The Office Worker: generates, edits, and secures professional office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX) "
        "locally with corporate styling, packaged Word templates (docxtpl), in-place edits, native Excel charts, formulas, "
        "batch pipelines, PDF processing (preview PNG, permanent redaction, structured RAG extraction, flattening, smart split, text, tables, forms, OCR, compression, signature), and Office-to-PDF conversion. "
        "Choose the exact tool matching your target document format. "
        "Hard Anti-Loop Rule: if any tool call fails 2 consecutive times, STOP immediately and report the exact error to the user."
    ),
)


def _ok(path: str) -> dict:
    return {"status": "ok", "path": os.path.abspath(path), "bytes": os.path.getsize(path)}


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
) -> dict:
    """Generates professional PDF document (invoice, report, executive summary, letter, balance statement) from Jinja HTML template, corporate theme, data, and design options (watermark, header logo, footers, page numbers, encryption). Returns JSON with status, absolute file path, and byte size. When to use: Use when creating polished, high-fidelity PDFs from structured data and HTML/CSS templates. When NOT to use: Do NOT use for editing existing PDF/Office files (use edit_word/edit_excel/pdf_manipulate) or when pure Office formats (.docx, .xlsx, .pptx) are required. Keywords: invoice, report, letter, contract, executive summary, balance statement, template, local, private, no api key, offline, cross-platform, deterministic, safe."""
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
        return _ok(_render_pdf(
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
        ))
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
        return _ok(_create_word(out_path, title=title, subtitle=subtitle or None, blocks=blocks, theme=theme or None, template_docx=template_docx or None, context=ctx))
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
        return _ok(_create_excel(out_path, title=title, sheets=sheets or None, theme=theme or None, table_style=table_style or None, auto_filter=auto_filter))
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def create_pptx(out_path: str, slides_json: str = "[]", theme: str = "") -> dict:
    """Creates editable native PowerPoint .pptx presentation (deck, executive summary, status report) from structured slide definitions with corporate theme. Returns JSON with status, absolute file path, and byte size. When to use: Use to generate editable PowerPoint slides with real typography, kicker badges, bullet lists, and tables. When NOT to use: Do NOT use for PDFs (use render_document) or Word documents (use create_word). Keywords: executive summary, report, deck, presentation, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        slides = json.loads(slides_json or "[]") if isinstance(slides_json, str) else (slides_json or [])
    except json.JSONDecodeError as e:
        return {"status": "error", "error": f"slides_json inválido: {e}"}
    try:
        return _ok(_create_pptx(out_path, slides=slides or None, theme=theme or None))
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
    reason: str = "",
    location: str = "",
) -> dict:
    """Signs PDF documents (contract, letter, minutes, invoice): stamps visual PNG signature/seal on target page and applies cryptographic PAdES digital signature if X.509 PEM certificate is provided. Returns JSON with status, signed file path, and byte size. When to use: Use to approve, endorse, or sign official business documents. When NOT to use: Do NOT use for modifying document content (use edit_word or render_document). Keywords: sign, stamp, contract, minutes, letter, invoice, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        return _ok(_sign_pdf(
            input_pdf=input_pdf,
            output=output,
            sello_img_path=sello_img_path or None,
            cert_pem=cert_pem or None,
            reason=reason or None,
            location=location or None,
        ))
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def edit_excel(
    input_path: str,
    operations: list = [],
    output_path: str = "",
) -> dict:
    """Modifies an existing Excel workbook (.xlsx / .xlsm: balance statement, financial report, audit) in place or to a new file via openpyxl, preserving styles, formatting, and VBA macros (keep_vba). Supports set_cell, append_row, add_column, add_chart (bar, line, pie), add_table, auto_filter, and formulas (SUM, SUMIF, AVERAGEIF, VLOOKUP, XLOOKUP, COUNTIFS). Returns JSON with status, path, fidelity, warnings, and operations count. When to use: Use to update existing spreadsheets without recreating them from scratch. When NOT to use: Do NOT use to build new spreadsheets from nothing (use create_excel) or to run VBA macros. Keywords: edit in place, balance statement, report, audit, formula, chart, local, private, no api key, offline, cross-platform, no office install needed, deterministic, safe."""
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
) -> dict:
    """Extracts structured text, paragraphs, slides, and tables from Word (.docx), PowerPoint (.pptx), and Excel (.xlsx / .xlsm) files. Returns JSON with format, structured elements (paragraphs, slides, sheets), table contents, and full text. When to use: Use to inspect, search, and extract content from Office documents directly without converting to PDF. When NOT to use: Do NOT use for PDF files (use read_pdf instead). Keywords: read, report, minutes, letter, contract, deck, balance statement, local, private, no api key, offline, cross-platform, deterministic, safe."""
    try:
        return _read_office(path, max_rows=max_rows)
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
        return _pdf_redact(
            input_path=input_path,
            output=output,
            search_text=search_text,
            regions=reg_list,
            fill_color=fill_color or None,
        )
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
        "create_excel": create_excel,
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
        "edit_excel": edit_excel,
        "edit_word": edit_word,
        "pdf_to_excel": pdf_to_excel,
        "read_office": read_office,
        "pdf_redact": pdf_redact,
        "pdf_extract_structured": pdf_extract_structured,
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

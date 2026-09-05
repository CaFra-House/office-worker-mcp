"""Lectura de PDF (input) — el gap que nadie cubre bien en los MCPs existentes.

Usa PyMuPDF (fitz) para texto/páginas y pdfplumber para tablas. pypdf para metadatos.
"""
from __future__ import annotations
import os

def read_pdf(path, max_pages=None, extract_tables=False, list_forms=False, extract_images=False, max_images=10):
    """Extrae texto por página + metadatos. Devuelve dict {pages:[{n,text}], metadata, n_pages, ...}.

    Si extract_tables=True, incluye tables_by_page y n_tables.
    Si list_forms=True, incluye is_form y fields.
    Si extract_images=True, incluye images (base64 data URLs) y n_images.
    """
    import base64
    import fitz  # PyMuPDF
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path): raise FileNotFoundError(path)
    doc = fitz.open(path)
    pages = []
    limit = len(doc) if not max_pages else min(len(doc), int(max_pages))
    for i in range(limit):
        pg = doc[i]
        pages.append({"n": i+1, "text": pg.get_text("text").strip()})
    meta = doc.metadata or {}
    out = {"path": path, "n_pages": len(doc), "metadata": {k:meta.get(k) for k in ("title","author","creator","producer") if meta.get(k)}, "pages": pages}

    if extract_images:
        images = []
        processed_xrefs = set()
        max_img_count = int(max_images) if max_images else 10
        for i in range(limit):
            pg = doc[i]
            for img_info in pg.get_images():
                xref = img_info[0]
                if xref in processed_xrefs:
                    continue
                processed_xrefs.add(xref)
                try:
                    base_img = doc.extract_image(xref)
                    if not base_img or not base_img.get("image"):
                        continue
                    ext = base_img.get("ext", "png")
                    b64 = base64.b64encode(base_img["image"]).decode("ascii")
                    images.append({
                        "page": i + 1,
                        "xref": xref,
                        "width": base_img.get("width", 0),
                        "height": base_img.get("height", 0),
                        "format": ext,
                        "data_url": f"data:image/{ext};base64,{b64}",
                    })
                    if len(images) >= max_img_count:
                        break
                except Exception:
                    continue
            if len(images) >= max_img_count:
                break
        out["images"] = images
        out["n_images"] = len(images)

    doc.close()

    if extract_tables:
        tb = globals()["extract_tables"](path, max_pages=max_pages)
        out["tables_by_page"] = tb.get("tables_by_page", {})
        out["n_tables"] = tb.get("n_tables", 0)

    if list_forms:
        fm = globals()["list_form_fields"](path)
        out["is_form"] = fm.get("is_form", False)
        out["fields"] = fm.get("fields", {})

    return out

def extract_tables(path, max_pages=None):
    """Extrae tablas por página con pdfplumber. Devuelve {page: [tabla(...)]}."""
    import pdfplumber
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path): raise FileNotFoundError(path)
    result = {}
    with pdfplumber.open(path) as pdf:
        limit = len(pdf.pages) if not max_pages else min(len(pdf.pages), int(max_pages))
        for i in range(limit):
            tbls = pdf.pages[i].extract_tables() or []
            if tbls: result[i+1] = tbls
    return {"path": path, "tables_by_page": result, "n_tables": sum(len(v) for v in result.values())}

def list_form_fields(path):
    """Lista campos de formulario AcroForm (pypdf). Vacío si no es un form."""
    from pypdf import PdfReader
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path): raise FileNotFoundError(path)
    r = PdfReader(path)
    fields = r.get_fields() or {}
    return {"path": path, "is_form": bool(fields), "fields": {k:(v.get("/FT"), ) for k,v in fields.items()}}


def pdf_preview(
    input_path: str,
    output: str | None = None,
    max_pages: int | None = None,
    dpi: int = 110,
) -> dict:
    """Renderiza páginas de un archivo PDF a imagen PNG vía PyMuPDF page.get_pixmap().

    Devuelve dict con:
    - status: 'ok'
    - input_path: ruta absoluta del PDF fuente
    - n_pages: total de páginas del PDF
    - rendered_pages: cantidad de páginas renderizadas
    - pages: lista de dicts [{'page': n, 'width': w, 'height': h, 'dpi': dpi, 'data_url': 'data:image/png;base64,...', 'path'?: str}]
    - data_url: data URL base64 de la primera página renderizada
    - path: ruta al archivo PNG si se guardó en disco
    - bytes: tamaño en bytes del archivo guardado
    """
    import base64
    import fitz  # PyMuPDF
    from .security import safe_out

    input_path = os.path.abspath(os.path.expanduser(str(input_path)))
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"PDF no encontrado: {input_path}")

    doc = fitz.open(input_path)
    total_pages = len(doc)
    limit = min(total_pages, int(max_pages)) if (max_pages and int(max_pages) > 0) else total_pages

    dpi_val = int(dpi) if (dpi and int(dpi) > 0) else 110

    pages_data = []
    saved_main_path = None

    for i in range(limit):
        page = doc[i]
        pix = page.get_pixmap(dpi=dpi_val)
        png_bytes = pix.tobytes("png")
        b64 = base64.b64encode(png_bytes).decode("ascii")
        data_url = f"data:image/png;base64,{b64}"

        page_info = {
            "page": i + 1,
            "width": pix.width,
            "height": pix.height,
            "dpi": dpi_val,
            "data_url": data_url,
        }

        if output:
            out_raw = str(output).strip()
            if limit == 1:
                target_file = safe_out(out_raw)
            else:
                stem, ext = os.path.splitext(out_raw)
                ext = ext or ".png"
                target_file = safe_out(f"{stem}_p{i+1}{ext}")

            with open(target_file, "wb") as f:
                f.write(png_bytes)

            page_info["path"] = os.path.abspath(target_file)
            if saved_main_path is None:
                saved_main_path = page_info["path"]

        pages_data.append(page_info)

    doc.close()

    result = {
        "status": "ok",
        "input_path": input_path,
        "n_pages": total_pages,
        "rendered_pages": len(pages_data),
        "pages": pages_data,
        "data_url": pages_data[0]["data_url"] if pages_data else "",
    }

    if saved_main_path:
        result["path"] = saved_main_path
        result["bytes"] = os.path.getsize(saved_main_path)

    return result


def _table_to_markdown(table: list[list]) -> str:
    """Convierte una matriz de celdas a tabla Markdown con formato pipe."""
    if not table:
        return ""
    cleaned = []
    for row in table:
        cleaned.append(["" if c is None else str(c).replace("\n", " ").replace("|", "\\|").strip() for c in row])
    if not cleaned:
        return ""
    header = cleaned[0]
    num_cols = len(header)
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join(["---"] * num_cols) + " |"]
    for row in cleaned[1:]:
        if len(row) < num_cols:
            row = row + [""] * (num_cols - len(row))
        lines.append("| " + " | ".join(row[:num_cols]) + " |")
    return "\n".join(lines)


def pdf_extract_structured(
    input_path: str,
    format: str = "markdown",
    output: str | None = None,
    max_pages: int | None = None,
) -> dict:
    """Extrae texto, tablas estructuradas y metadatos de un PDF a Markdown o JSON.

    - input_path: ruta al archivo PDF.
    - format: 'markdown' (o 'md') para Markdown legible con tablas pipe, o 'json' para estructura de datos.
    - output: ruta opcional de archivo de salida en disco.
    - max_pages: límite opcional de páginas a procesar.

    Devuelve dict con status, format, pages, content (si markdown), path y bytes (si output).
    """
    import json
    import fitz  # PyMuPDF
    import pdfplumber
    from pathlib import Path
    from .security import safe_out

    input_path = os.path.abspath(os.path.expanduser(str(input_path)))
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"PDF no encontrado: {input_path}")

    fmt = str(format or "markdown").lower().strip()
    if fmt not in ("markdown", "md", "json"):
        raise ValueError(f"Formato no soportado: '{format}'. Use 'markdown' o 'json'.")

    doc_fitz = fitz.open(input_path)
    total_pages = len(doc_fitz)
    limit = min(total_pages, int(max_pages)) if (max_pages and int(max_pages) > 0) else total_pages

    pages_data = []
    with pdfplumber.open(input_path) as doc_plumber:
        for i in range(limit):
            pg_fitz = doc_fitz[i]
            pg_plumber = doc_plumber.pages[i] if i < len(doc_plumber.pages) else None

            text = pg_fitz.get_text("text").strip()
            images_n = len(pg_fitz.get_images())
            tables = (pg_plumber.extract_tables() or []) if pg_plumber else []

            clean_tables = []
            for tbl in tables:
                clean_tbl = []
                for row in tbl:
                    clean_tbl.append(["" if cell is None else str(cell).strip() for cell in row])
                clean_tables.append(clean_tbl)

            pages_data.append({
                "page": i + 1,
                "text": text,
                "tables": clean_tables,
                "images_n": images_n,
            })

    meta = doc_fitz.metadata or {}
    title = meta.get("title") or Path(input_path).stem
    doc_fitz.close()

    total_tables = sum(len(p["tables"]) for p in pages_data)

    if fmt == "json":
        result = {
            "status": "ok",
            "format": "json",
            "title": title,
            "n_pages": limit,
            "total_pages": total_pages,
            "n_tables": total_tables,
            "pages": pages_data,
        }
        if output:
            out_file = safe_out(output)
            with open(out_file, "w", encoding="utf-8") as f:
                json.dump(result, f, indent=2, ensure_ascii=False)
            result["path"] = os.path.abspath(out_file)
            result["bytes"] = os.path.getsize(out_file)
        return result

    # Generar Markdown estructurado
    md_sections = []
    if title:
        md_sections.append(f"# {title}\n")

    for p in pages_data:
        p_num = p["page"]
        section_lines = [f"## Página {p_num}\n"]
        if p["text"]:
            section_lines.append(p["text"] + "\n")
        if p["tables"]:
            for t_idx, tbl in enumerate(p["tables"], 1):
                section_lines.append(f"### Tabla {t_idx} (Pág. {p_num})")
                section_lines.append(_table_to_markdown(tbl) + "\n")
        if p["images_n"] > 0:
            section_lines.append(f"*[Página {p_num}: {p['images_n']} imágenes detectadas]*\n")
        md_sections.append("\n".join(section_lines))

    md_content = "\n---\n\n".join(md_sections).strip()

    result = {
        "status": "ok",
        "format": "markdown",
        "content": md_content,
        "n_pages": limit,
        "total_pages": total_pages,
        "n_tables": total_tables,
        "pages": pages_data,
    }
    if output:
        out_file = safe_out(output)
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(md_content)
        result["path"] = os.path.abspath(out_file)
        result["bytes"] = os.path.getsize(out_file)
    return result


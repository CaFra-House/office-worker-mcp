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

"""Lectura de PDF (input) — el gap que nadie cubre bien en los MCPs existentes.

Usa PyMuPDF (fitz) para texto/páginas y pdfplumber para tablas. pypdf para metadatos.
"""
from __future__ import annotations
import os

def read_pdf(path, max_pages=None):
    """Extrae texto por página + metadatos. Devuelve dict {pages:[{n,text}], metadata, n_pages}."""
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
    doc.close()
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

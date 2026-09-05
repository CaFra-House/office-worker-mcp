"""Generación de Word (.docx) con python-docx, aplicando tema corporativo."""
from __future__ import annotations
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def create_word(out_path, title="", subtitle=None, blocks=None, theme=None):
    """Crea un .docx profesional.

    blocks: lista de dicts. Tipos soportados:
      {"type":"h1"|"h2"|"p", "text": "..."}
      {"type":"table", "headers":[...], "rows":[[...], ...]}
    Si blocks es None → documento mínimo con título. Devuelve ruta absoluta.
    """
    from .themes import load_theme
    th = load_theme(theme)
    primary = _hex_to_rgb(th["primary"])

    out_path = os.path.abspath(os.path.expanduser(out_path))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    doc = Document()
    # estilo base de fuente/cuerpo según tema
    st = doc.styles["Normal"]; st.font.name = th.get("font_body","Arial"); st.font.size = Pt(10.5)

    if title:
        h = doc.add_heading(title, level=0); 
        for r in h.runs: r.font.color.rgb = primary
    if subtitle:
        p = doc.add_paragraph(subtitle)
        if p.runs: p.runs[0].font.color.rgb = _hex_to_rgb(th["muted"])

    for b in (blocks or []):
        t = b.get("type","p")
        if t == "h1":
            hh = doc.add_heading(b["text"], level=1)
            for r in hh.runs: r.font.color.rgb = primary
        elif t == "h2":
            hh = doc.add_heading(b["text"], level=2)
            for r in hh.runs: r.font.color.rgb = primary
        elif t == "table":
            headers = b.get("headers",[]); rows = b.get("rows",[])
            tbl = doc.add_table(rows=1, cols=len(headers)); tbl.style="Light Grid Accent 1"
            hdr = tbl.rows[0].cells
            for i,hv in enumerate(headers): hdr[i].text=str(hv)
            for row in rows:
                cells = tbl.add_row().cells
                for i,cv in enumerate(row):
                    if i < len(cells): cells[i].text=str(cv)
        else:  # p / default
            doc.add_paragraph(b.get("text",""))

    doc.save(out_path)
    if os.path.getsize(out_path) < 300: raise RuntimeError(f"DOCX sospechosamente pequeño ({os.path.getsize(out_path)}B)")
    return out_path

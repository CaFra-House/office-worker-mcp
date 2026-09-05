"""The Office Worker — MCP server (cara delgada sobre office_worker.core).

Meta: pocas tools poderosas (~8 aquí), no 47 CRUD genéricas → bajo overhead de contexto.
Cada tool crea/lee un archivo real y devuelve {status, path, ...} o {status:"error", error}.
Regla anti-loop en instructions: si una tool falla 2 veces, NO reintentar con variantes.
"""
from __future__ import annotations
import os, json

from mcp.server.fastmcp import FastMCP

from office_worker.core import (
    render_pdf as _render_pdf,
    create_word as _create_word,
    create_excel as _create_excel,
    create_pptx as _create_pptx,
    read_pdf as _read_pdf,
    extract_tables as _extract_tables,
    list_form_fields as _list_form_fields,
    load_theme, DEFAULT_THEME,
)

mcp = FastMCP(
    "office-worker",
    instructions=(
        "The Office Worker: genera documentos de oficina profesionales (PDF/Word/Excel/PPTX) "
        "desde datos + tema corporativo, y lee PDFs (texto/tablas/formularios). "
        "Usa la tool que corresponda al formato pedido. Regla anti-loop: si una llamada falla 2 "
        "veces, NO reintentes con variantes; reporta el error exacto y detente."
    ),
)


def _ok(path): return {"status": "ok", "path": os.path.abspath(path), "bytes": os.path.getsize(path)}

@mcp.tool()
def render_document(template_html: str, out_path: str, data_json: str = "{}", theme: str | None = None) -> dict:
    """Genera un PDF profesional desde plantilla HTML/Jinja + datos + tema."""
    try: data = json.loads(data_json or "{}") or {}
    except json.JSONDecodeError as e: return {"status":"error","error":f"data_json inválido: {e}"}
    if isinstance(data.get("rows"), list):
        headers=data.get("headers") or []
        thead="".join(f"<th>{h}</th>" for h in headers)
        body="\n".join("<tr>"+"".join(f"<td>{c}</td>" for c in r)+"</tr>" for r in data["rows"])
        data["tabla"]=f"<table><thead><tr>{thead}</tr></thead><tbody>{body}</tbody></table>"
    try: return _ok(_render_pdf(template_html, out_path, data=data, theme=theme))
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def create_word(out_path: str, title: str = "", subtitle: str | None = None, blocks_json: str = "[]", theme: str | None = None) -> dict:
    """Crea un .docx. blocks_json: JSON lista de {"type":"h1|h2|p|table","text"|"headers"/"rows"}."""
    try: blocks=json.loads(blocks_json or "[]") or []
    except json.JSONDecodeError as e: return {"status":"error","error":f"blocks_json inválido: {e}"}
    try: return _ok(_create_word(out_path,title=title,subtitle=subtitle,blocks=blocks,theme=theme))
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def create_excel(out_path: str, title: str = "", sheets_json: str = "[]", theme: str | None = None) -> dict:
    """Crea un .xlsx. sheets_json: JSON lista de {"name","headers":[...],"rows":[[...]]}."""
    try: sheets=json.loads(sheets_json or "[]") or []
    except json.JSONDecodeError as e: return {"status":"error","error":f"sheets_json inválido: {e}"}
    try: return _ok(_create_excel(out_path,title=title,sheets=sheets or None,theme=theme))
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def create_pptx(out_path: str, slides_json: str = "[]", theme: str | None = None) -> dict:
    """Crea un .pptx EDITABLE. slides_json: JSON lista de {"title","kicker","bullets":[...]}."""
    try: slides=json.loads(slides_json or "[]") or []
    except json.JSONDecodeError as e: return {"status":"error","error":f"slides_json inválido: {e}"}
    try: return _ok(_create_pptx(out_path,slides=slides or None,theme=theme))
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def read_pdf(path: str, max_pages: int | None = None) -> dict:
    """Lee texto + metadatos de un PDF (input). Devuelve páginas con su texto."""
    try: return _read_pdf(path,max_pages=max_pages) | {"status":"ok"}
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def pdf_extract_tables(path: str, max_pages: int | None = None) -> dict:
    """Extrae tablas estructuradas de un PDF (pdfplumber)."""
    try: return _extract_tables(path,max_pages=max_pages) | {"status":"ok"}
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def pdf_list_form_fields(path: str) -> dict:
    """Lista campos de formulario AcroForm de un PDF (vacío si no es form)."""
    try: return _list_form_fields(path) | {"status":"ok"}
    except Exception as e: return {"status":"error","error":str(e)}

@mcp.tool()
def list_themes(theme_name_or_path: str | None = None) -> dict:
    """Devuelve el tema resuelto (paleta+fuente) para usarlo como referencia de diseño."""
    t = load_theme(theme_name_or_path) if theme_name_or_path else dict(DEFAULT_THEME)
    return {"status":"ok","theme":t}


def main(): mcp.run()

if __name__ == "__main__": main()

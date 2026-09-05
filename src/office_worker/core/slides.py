"""Generación de PPTX editable vía html-to-pptx (Playwright Chromium).

Cada slide es un <section class="slide">. Inyecta el tema como CSS vars para
consistencia visual. Produce PPTX NATIVO editable (textos reales, no imágenes).
Requiere Playwright + Chromium ya instalados (verificado en Fase B).
"""
from __future__ import annotations
import os, tempfile

_SLIDE_CSS = """
* { box-sizing: border-box; margin: 0; }
body { font-family: var(--ow-font-body); color: var(--ow-text); }
.slide { width: 1280px; height: 720px; padding: 56px 64px; background: var(--ow-bg); display:flex; flex-direction:column; }
.slide h1 { font-family: var(--ow-font-title); font-size: 44px; color: var(--ow-primary); border-bottom: 3px solid var(--ow-primary); padding-bottom: 12px; margin-bottom: 24px; }
.slide h2 { font-family: var(--ow-font-title); font-size: 30px; color: var(--ow-primary); margin-bottom: 18px; }
.slide p, .slide li { font-size: 22px; line-height: 1.5; }
.slide ul { padding-left: 32px; margin-top: 10px; }
.slide table { width:100%; border-collapse:collapse; font-size:20px; margin-top:16px;}
.slide th{ background:var(--ow-primary); color:#fff; text-align:left; padding:10px 12px;}
.slide td{ padding:9px 12px; border-bottom:.5pt solid #c8d4e0;}
.slide tr:nth-child(even) td{ background:var(--ow-row-alt);}
.kicker{ color:var(--ow-accent); font-weight:bold; letter-spacing:.5px; text-transform:uppercase; font-size:18px; margin-bottom:8px;}
"""

def create_pptx(out_path, slides=None, theme=None):
    """slides: lista de dicts {"title":..., "kicker":..., "bullets":[...]} o {"html":"..."} crudo.

    Si slides es None → una portada genérica. Devuelve ruta absoluta del .pptx editable.
    """
    from .themes import load_theme, css_vars
    from .security import safe_out
    th = load_theme(theme)

    parts = []
    for s in (slides or [{"title": "The Office Worker", "bullets": ["Presentación generada"]}]):
        if "html" in s and s["html"]:
            parts.append(f'<section class="slide">{s["html"]}</section>')
            continue
        k = f'<div class="kicker">{s.get("kicker","")}</div>' if s.get("kicker") else ""
        items = "".join(f"<li>{b}</li>" for b in s.get("bullets", []))
        body = f"<ul>{items}</ul>" if items else f"<p>{s.get('subtitle','')}</p>" if s.get('subtitle') else ""
        parts.append(
            f'<section class="slide"><h1>{s.get("title","")}</h1>{k}{body}</section>'
        )

    deck_html = (f'<!DOCTYPE html><html><head><meta charset="utf-8">'
                 f'<style>{css_vars(th)}{_SLIDE_CSS}</style></head><body>{"".join(parts)}</body></html>')

    out_path = safe_out(out_path)

    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
        tf.write(deck_html); tmp_html = tf.name

    try:
        import asyncio, concurrent.futures
        try:
            from html_to_pptx import convert
        except ImportError as e:
            raise ImportError(
                "html-to-pptx no está disponible. Instale el extra opcional: pip install 'office-worker-mcp[pptx]' "
                "y ejecute 'playwright install chromium'."
            ) from e
        # El MCP server corre con su propio event loop; asyncio.run() fallaría ahí.
        # Ejecutamos convert() en un hilo con su propio loop aislado.
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
            result = ex.submit(asyncio.run, convert(tmp_html, out_path)).result(timeout=300)
    finally:
        try: os.unlink(tmp_html)
        except OSError: pass

    if not os.path.exists(result) or os.path.getsize(result) < 3000:
        raise RuntimeError(f"PPTX generado inválido o vacío ({result})")
    return result

"""Núcleo de render: template Jinja + tema → documento.

H1 implementa el camino PDF (HTML/CSS → WeasyPrint), que es el motor de diseño ya
probado en producción (caso ISO 27001). docx/xlsx/pptx se agregan en hitos siguientes
reusando la misma entrada (template + data + theme).
"""
from __future__ import annotations
import os
from jinja2 import Environment, FileSystemLoader, BaseLoader, StrictUndefined

from .themes import load_theme, css_vars
from .security import safe_out, safe_url_fetcher

# CSS base compartido por todas las plantillas (usa las variables del tema).
_BASE_CSS = """
@page { size: A4; margin: 1.6cm 1.5cm 1.8cm 1.5cm;
  @bottom-center { content: "Página " counter(page) " de " counter(pages); font-size: 8pt; color: var(--ow-muted); } }
* { box-sizing: border-box; }
body { font-family: var(--ow-font-body); font-size: 10pt; line-height: 1.45; color: var(--ow-text); }
h1 { font-family: var(--ow-font-title); font-size: 17pt; color: var(--ow-primary); border-bottom: 2px solid var(--ow-primary); padding-bottom: 5pt; margin-top: 0; }
h2 { font-family: var(--ow-font-title); font-size: 12.5pt; color: var(--ow-primary); margin-top: 14pt; }
p, li { orphans: 3; widows: 3; }
table { width: 100%; border-collapse: collapse; font-size: 9pt; table-layout: fixed; margin-top: 8pt; }
th { background: var(--ow-primary); color: #fff; text-align: left; padding: 4pt 6pt; }
td { padding: 4pt 6pt; border-bottom: .5pt solid #c8d4e0; vertical-align: top; }
tr:nth-child(even) td { background: var(--ow-row-alt); }
.muted { color: var(--ow-muted); font-size: 9pt; }
"""

def _env(template_dir=None):
    if template_dir and os.path.isdir(template_dir):
        return Environment(loader=FileSystemLoader(template_dir), undefined=StrictUndefined, autoescape=False)
    return Environment(loader=BaseLoader(), undefined=StrictUndefined, autoescape=False)

import html

def render_pdf(
    template_html_or_path,
    out_path,
    data=None,
    theme=None,
    template_dir=None,
    logo=None,
    password=None,
    watermark_text=None,
    footer_left=None,
    footer_right=None,
    page_numbers=True,
):
    """Rellena una plantilla HTML (Jinja) con `data` + `theme` y exporta a PDF vía WeasyPrint.

    - template_html_or_path: string HTML (con placeholders {{ }}) o ruta a un .html.
    - data: dict de variables para Jinja. Si hay lista 'rows' con 'headers', genera tabla.
    - theme: dict o nombre/ruta (ver themes.load_theme). None → tema ADEN por defecto.
    - logo: ruta a archivo de imagen (PNG/JPG) a insertar en la cabecera vía CSS WeasyPrint.
    - password: clave opcional para cifrar el PDF resultante (pypdf encrypt).
    - watermark_text: texto diagonal semitransparente (ej: 'CONFIDENCIAL').
    - footer_left / footer_right: texto en las esquinas inferiores izquierda y derecha.
    - page_numbers: booleano (default True) para mostrar 'Página X de Y' en el centro inferior.
    Devuelve la ruta absoluta del PDF generado.

    Seguridad:
    - Valida out_path con safe_out contra path traversal y sobreescritura de sistema.
    - Utiliza safe_url_fetcher contra SSRF remoto y acceso a archivos sensibles (/etc).
    Limitación residual: WeasyPrint puede leer archivos locales permitidos legibles por el proceso si no se define OFFICE_WORKER_ALLOWED_DIR.
    """
    import pathlib
    from weasyprint import HTML

    th = load_theme(theme)

    # ¿es ruta a archivo o HTML inline?
    if template_html_or_path and os.path.exists(str(template_html_or_path)):
        with open(template_html_or_path, "r", encoding="utf-8") as f:
            tpl_text = f.read()
    else:
        tpl_text = str(template_html_or_path or "")

    env = _env(template_dir)
    tmpl = env.from_string(tpl_text)

    # helper opcional para tablas declarativas en la plantilla via contexto simple
    ctx = dict(data or {})

    logo_css = ""
    if logo:
        logo_path = os.path.abspath(os.path.expanduser(str(logo)))
        if not os.path.exists(logo_path):
            raise FileNotFoundError(f"Logo no encontrado: {logo_path}")
        logo_uri = pathlib.Path(logo_path).as_uri()
        logo_css = f"""
@page {{
  margin-top: 2.2cm;
  @top-right {{
    content: url('{logo_uri}');
    max-height: 1.2cm;
    vertical-align: middle;
  }}
}}
"""
        ctx.setdefault("logo", logo_uri)

    footer_css_parts = []
    if not page_numbers:
        footer_css_parts.append("@bottom-center { content: none; }")
    if footer_left:
        esc_fl = str(footer_left).replace("\\", "\\\\").replace('"', '\\"')
        footer_css_parts.append(f'@bottom-left {{ content: "{esc_fl}"; font-size: 8pt; color: var(--ow-muted); }}')
    if footer_right:
        esc_fr = str(footer_right).replace("\\", "\\\\").replace('"', '\\"')
        footer_css_parts.append(f'@bottom-right {{ content: "{esc_fr}"; font-size: 8pt; color: var(--ow-muted); }}')

    footer_css = f"\n@page {{\n  {' '.join(footer_css_parts)}\n}}\n" if footer_css_parts else ""

    watermark_css = ""
    watermark_html = ""
    if watermark_text:
        watermark_css = """
.ow-watermark {
  position: fixed;
  top: 35%;
  left: 5%;
  width: 90%;
  text-align: center;
  transform: rotate(-35deg);
  font-size: 52pt;
  font-weight: bold;
  color: rgba(180, 180, 180, 0.25);
  z-index: -1000;
  pointer-events: none;
  text-transform: uppercase;
}
"""
        watermark_html = f'<div class="ow-watermark">{html.escape(str(watermark_text))}</div>'

    body = tmpl.render(**ctx)

    html_doc = f"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<style>{css_vars(th)}{_BASE_CSS}{logo_css}{footer_css}{watermark_css}</style></head><body>{watermark_html}{body}</body></html>"""

    out_path = safe_out(out_path)
    HTML(string=html_doc, url_fetcher=safe_url_fetcher).write_pdf(out_path)

    if password:
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(out_path)
        writer = PdfWriter()
        writer.append(reader)
        writer.encrypt(user_password=password)
        with open(out_path, "wb") as f_enc:
            writer.write(f_enc)

    if os.path.getsize(out_path) < 500:
        raise RuntimeError(f"PDF generado sospechosamente pequeño ({os.path.getsize(out_path)}B): revisar plantilla/datos")
    return out_path

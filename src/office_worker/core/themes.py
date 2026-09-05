"""Capa de temas corporativos.

Un tema es un dict con paleta + tipografía que se inyecta como CSS variables en las
plantillas HTML (para PDF vía WeasyPrint) y como constantes para docx/xlsx/pptx.
Esto resuelve el gap #3 de office-mcp: diseño consistente aplicable de una vez,
en vez de parámetros sueltos por llamada.
"""
from __future__ import annotations
import os

# Tema por defecto (paleta ADEN usada ya en producción).
DEFAULT_THEME = {
    "name": "aden",
    "primary": "#003366",      # azul oscuro — encabezados, headers de tabla
    "accent": "#3B82F6",       # acento — cabeceras destacadas, enlaces
    "text": "#1A202C",         # cuerpo
    "muted": "#718096",        # secundario / pies
    "row_alt": "#F5F7FA",      # filas alternas de tabla
    "bg": "#FFFFFF",           # fondo
    "font_title": "Helvetica Neue, Arial, sans-serif",
    "font_body": "Helvetica Neue, Arial, sans-serif",
}

def load_theme(name_or_path: str | None) -> dict:
    """Devuelve un tema por nombre conocido o ruta a YAML/JSON. Sin arg → DEFAULT_THEME."""
    if not name_or_path:
        return dict(DEFAULT_THEME)
    p = name_or_path
    if os.path.exists(p):
        text = open(p).read()
        if p.endswith((".yaml", ".yml")):
            import yaml  # opcional; solo si cargan tema desde archivo YAML
            data = yaml.safe_load(text) or {}
        else:
            import json as _json
            data = _json.loads(text) or {}
        base = dict(DEFAULT_THEME); base.update(data); base.setdefault("name", os.path.basename(p))
        return base
    # nombre conocido (extender después con más temas empaquetados)
    return dict(DEFAULT_THEME)

def css_vars(theme: dict) -> str:
    """Bloque :root con las variables CSS del tema, listo para inyectar en <style>."""
    return (
        f":root{{ --ow-primary:{theme['primary']}; --ow-accent:{theme['accent']};"
        f"--ow-text:{theme['text']}; --ow-muted:{theme['muted']};"
        f"--ow-row-alt:{theme['row_alt']}; --ow-bg:{theme['bg']};"
        f"--ow-font-title:{theme['font_title']}; --ow-font-body:{theme['font_body']}; }}"
    )

"""Capa de temas corporativos.

Un tema es un dict con paleta + tipografía que se inyecta como CSS variables en las
plantillas HTML (para PDF vía WeasyPrint) y como constantes para docx/xlsx/pptx.
Esto resuelve el gap #3 de office-mcp: diseño consistente aplicable de una vez,
en vez de parámetros sueltos por llamada.
"""
from __future__ import annotations
import os

from dataclasses import dataclass, asdict

@dataclass
class Theme:
    """Definición tipada de tema corporativo."""
    name: str
    primary: str
    accent: str
    text: str
    muted: str
    row_alt: str
    bg: str
    font_title: str
    font_body: str

    def to_dict(self) -> dict:
        return asdict(self)

# Tema por defecto (paleta ADEN usada ya en producción).
ADEN_THEME = {
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

DEFAULT_THEME = ADEN_THEME

CLARO_THEME = {
    "name": "claro",
    "primary": "#2563EB",      # azul moderno
    "accent": "#38BDF8",       # celeste vivo
    "text": "#1E293B",         # pizarra oscuro
    "muted": "#64748B",        # gris medio
    "row_alt": "#F8FAFC",      # fondo alterno muy suave
    "bg": "#FFFFFF",           # blanco
    "font_title": "Helvetica Neue, Arial, sans-serif",
    "font_body": "Helvetica Neue, Arial, sans-serif",
}

OSCURO_THEME = {
    "name": "oscuro",
    "primary": "#60A5FA",      # azul suave contrastado
    "accent": "#93C5FD",       # celeste claro
    "text": "#F3F4F6",         # texto claro
    "muted": "#9CA3AF",        # gris medio
    "row_alt": "#1F2937",      # fila alterna oscura
    "bg": "#111827",           # fondo grafito oscuro
    "font_title": "Helvetica Neue, Arial, sans-serif",
    "font_body": "Helvetica Neue, Arial, sans-serif",
}

MINIMAL_THEME = {
    "name": "minimal",
    "primary": "#111827",      # negro neutro
    "accent": "#4B5563",       # gris grafito
    "text": "#111827",         # texto negro
    "muted": "#9CA3AF",        # gris tenue
    "row_alt": "#F9FAFB",      # fila alterna sutil
    "bg": "#FFFFFF",           # fondo blanco
    "font_title": "Georgia, Times New Roman, serif",
    "font_body": "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
}

CORPORATE_BLUE_THEME = {
    "name": "corporate-blue",
    "primary": "#0F4C81",      # Classic Navy Pantone
    "accent": "#1E40AF",       # azul corporativo intenso
    "text": "#1E293B",         # texto sobrio
    "muted": "#64748B",        # secundario
    "row_alt": "#EFF6FF",      # tinte azulado sutil
    "bg": "#FFFFFF",           # blanco
    "font_title": "Helvetica Neue, Arial, sans-serif",
    "font_body": "Helvetica Neue, Arial, sans-serif",
}

THEMES: dict[str, dict] = {
    "aden": ADEN_THEME,
    "claro": CLARO_THEME,
    "oscuro": OSCURO_THEME,
    "minimal": MINIMAL_THEME,
    "corporate-blue": CORPORATE_BLUE_THEME,
    "corporate_blue": CORPORATE_BLUE_THEME,
}

def load_theme(name_or_path: str | None) -> dict:
    """Devuelve un tema por nombre conocido o ruta a YAML/JSON. Sin arg → DEFAULT_THEME."""
    if not name_or_path:
        return dict(DEFAULT_THEME)
    p = str(name_or_path).strip()
    key = p.lower()
    if key in THEMES:
        return dict(THEMES[key])
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            text = f.read()
        if p.endswith((".yaml", ".yml")):
            import yaml  # opcional; solo si cargan tema desde archivo YAML
            data = yaml.safe_load(text) or {}
        else:
            import json as _json
            data = _json.loads(text) or {}
        base = dict(DEFAULT_THEME); base.update(data); base.setdefault("name", os.path.basename(p))
        return base
    # Si no coincide, retorna el tema por defecto
    return dict(DEFAULT_THEME)

def css_vars(theme: dict) -> str:
    """Bloque :root con las variables CSS del tema, listo para inyectar en <style>."""
    return (
        f":root{{ --ow-primary:{theme['primary']}; --ow-accent:{theme['accent']};"
        f"--ow-text:{theme['text']}; --ow-muted:{theme['muted']};"
        f"--ow-row-alt:{theme['row_alt']}; --ow-bg:{theme['bg']};"
        f"--ow-font-title:{theme['font_title']}; --ow-font-body:{theme['font_body']}; }}"
    )

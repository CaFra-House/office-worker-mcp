"""Módulo de plantillas empaquetadas profesionales (.docx / docxtpl).

Permite listar y resolver plantillas oficiales incluidas en The Office Worker:
- acta_meeting: acta de reunión (asistentes, temas, acuerdos, firmas)
- informe_ejecutivo: informe de gestión (resumen, secciones, tabla KPIs, conclusiones)
- factura_simple: factura comercial (emisor, receptor, ítems, IVA, totales)
- carta_formal: carta formal corporativa (remitente, destinatario, asunto, cuerpo, firma)
- checklist_auditoria: matriz de auditoría/control (criterios, estados, evidencia)
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Any

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates_data"

TEMPLATES_CATALOG: dict[str, dict[str, Any]] = {
    "acta_meeting": {
        "name": "acta_meeting",
        "description": "Acta de reunión / minutes con fecha, asistentes, puntos tratados, acuerdos y firmas.",
        "keywords": ["acta", "minutes", "reunión", "directorio", "acuerdos"],
        "variables": {
            "titulo": "str (ej: 'Acta de Reunión N° 12')",
            "fecha": "str (ej: '2026-09-05')",
            "hora": "str (ej: '10:00 AM')",
            "lugar": "str (ej: 'Sala de Conferencias')",
            "asistentes": "list[dict] [{'nombre': str, 'rol': str}]",
            "puntos": "list[dict] [{'orden': int|str, 'tema': str, 'discusion': str}]",
            "acuerdos": "list[dict] [{'acuerdo': str, 'responsable': str, 'fecha_limite': str}]",
            "firmas": "list[dict] [{'nombre': str, 'cargo': str}]",
        },
    },
    "informe_ejecutivo": {
        "name": "informe_ejecutivo",
        "description": "Informe ejecutivo / executive summary con resumen, secciones temáticas, tabla de indicadores (KPIs) y conclusiones.",
        "keywords": ["informe", "report", "resumen ejecutivo", "executive summary", "balance", "kpi"],
        "variables": {
            "titulo": "str (ej: 'Informe de Gestión Trimestral')",
            "subtitulo": "str (ej: 'Tercer Trimestre 2026')",
            "fecha": "str (ej: '2026-09-05')",
            "autor": "str (ej: 'Gerencia General')",
            "resumen": "str",
            "secciones": "list[dict] [{'titulo': str, 'contenido': str}]",
            "indicadores": "list[dict] [{'kpi': str, 'meta': str, 'actual': str, 'estado': str}]",
            "conclusiones": "str",
        },
    },
    "factura_simple": {
        "name": "factura_simple",
        "description": "Factura comercial / invoice con datos de emisor, receptor, tabla de ítems/servicios, subtotal, IVA y total.",
        "keywords": ["factura", "invoice", "cobro", "presupuesto", "iva", "servicios"],
        "variables": {
            "numero_factura": "str (ej: 'A-0001-00004523')",
            "fecha": "str (ej: '2026-09-05')",
            "fecha_vencimiento": "str (ej: '2026-09-20')",
            "emisor": "dict {'nombre': str, 'cuit_nif': str, 'direccion': str, 'contacto': str}",
            "receptor": "dict {'nombre': str, 'cuit_nif': str, 'direccion': str}",
            "items": "list[dict] [{'descripcion': str, 'cantidad': str|int, 'precio_unitario': str, 'subtotal': str}]",
            "subtotal": "str (ej: '$100.000')",
            "iva": "str (ej: '$21.000')",
            "total": "str (ej: '$121.000')",
            "condiciones_pago": "str (ej: 'Transferencia a 15 días')",
        },
    },
    "carta_formal": {
        "name": "carta_formal",
        "description": "Carta formal / letter institucional con remitente, destinatario, asunto, saludo, cuerpo de carta y bloque de firma.",
        "keywords": ["carta", "letter", "comunicado", "notificación", "formal", "remitente"],
        "variables": {
            "lugar_fecha": "str (ej: 'Buenos Aires, 5 de Septiembre de 2026')",
            "destinatario": "dict {'nombre': str, 'cargo': str, 'organizacion': str, 'direccion': str}",
            "asunto": "str (ej: 'Presentación de propuesta')",
            "saludo": "str (ej: 'De mi mayor consideración:')",
            "cuerpo": "str (texto del mensaje)",
            "despedida": "str (ej: 'Sin otro particular, saluda atentamente,')",
            "firma": "dict {'nombre': str, 'cargo': str, 'organizacion': str}",
        },
    },
    "checklist_auditoria": {
        "name": "checklist_auditoria",
        "description": "Checklist de auditoría / audit control con matriz de criterios, estado de cumplimiento, evidencia y dictamen.",
        "keywords": ["checklist", "auditoria", "audit", "control", "cumplimiento", "inspección"],
        "variables": {
            "titulo": "str (ej: 'Checklist de Auditoría Interna ISO 27001')",
            "auditor": "str (ej: 'Equipo de Calidad')",
            "fecha": "str (ej: '2026-09-05')",
            "alcance": "str (ej: 'Sistemas de Información')",
            "items": "list[dict] [{'item': str, 'criterio': str, 'estado': str, 'evidencia': str, 'observaciones': str}]",
            "conclusion": "str (ej: 'Cumplimiento satisfactorio')",
        },
    },
}


def list_packaged_templates() -> list[dict[str, Any]]:
    """Devuelve la lista de plantillas empaquetadas con su descripción y variables esperadas."""
    results = []
    for key, info in TEMPLATES_CATALOG.items():
        results.append({
            "name": info["name"],
            "description": info["description"],
            "keywords": info.get("keywords", []),
            "variables": info["variables"],
        })
    return results


def resolve_template_path(name_or_path: str) -> str:
    """Resuelve la ruta absoluta de una plantilla.

    1. Si es una ruta existente en disco, la devuelve tal cual (normalizada).
    2. Si coincide con una plantilla empaquetada (con o sin extensión .docx),
       resuelve la ruta dentro de templates_data.
    3. Si no existe, lanza FileNotFoundError.
    """
    if not name_or_path or not str(name_or_path).strip():
        raise ValueError("El nombre o ruta de la plantilla no puede estar vacío.")

    raw = str(name_or_path).strip()
    direct_path = os.path.abspath(os.path.expanduser(raw))
    if os.path.exists(direct_path):
        return direct_path

    # Buscar en templates empaquetados
    clean_name = raw
    if clean_name.lower().endswith(".docx"):
        clean_name = clean_name[:-5]

    packaged = TEMPLATES_DIR / f"{clean_name}.docx"
    if packaged.exists():
        return str(packaged.resolve())

    raise FileNotFoundError(
        f"Plantilla no encontrada: '{name_or_path}'. "
        f"Plantillas empaquetadas disponibles: {list(TEMPLATES_CATALOG.keys())}"
    )


def get_template_schema(name: str) -> dict[str, Any]:
    """Devuelve el esquema y descripción de una plantilla empaquetada."""
    clean_name = name.strip()
    if clean_name.lower().endswith(".docx"):
        clean_name = clean_name[:-5]
    if clean_name in TEMPLATES_CATALOG:
        return TEMPLATES_CATALOG[clean_name]
    raise KeyError(
        f"Plantilla '{name}' no existe en el catálogo. Disponibles: {list(TEMPLATES_CATALOG.keys())}"
    )

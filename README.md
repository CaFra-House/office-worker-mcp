# The Office Worker — `office-worker-mcp`

![CI](https://github.com/CaFra-House/office-worker-mcp/actions/workflows/ci.yml/badge.svg)
![PyPI version](https://img.shields.io/pypi/v/office-worker-mcp)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/office-worker-mcp)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

> **Your agent's document clerk.** Genera documentos de oficina profesionales —
> **PDF, Word, Excel y PowerPoint** — desde datos + un tema corporativo, y procesa PDFs
> (texto, tablas, formularios, OCR, manipulación y conversión Office). Diseñado para agentes de IA (Hermes, Claude, Cursor)
> vía [MCP](https://modelcontextprotocol.io), 100% local-first.

```bash
pip install office-worker-mcp          # instala librería + MCP server + CLI 'owi'
pip install "office-worker-mcp[ocr]"   # extra para OCR con Tesseract + Pillow
```

Conectalo a tu agente (ejemplo config MCP / stdio):

```json
{ "mcpServers": { "office-worker": { "command": "office-worker-mcp" } } }
```

Y listo: le pedís *"haceme una factura/informe/deck"*, *"llená este formulario"*, o *"convertí este docx a PDF"* en lenguaje natural y tu agente
usa las tools del oficinista para producir el archivo bien diseñado.

## Por qué existe (diferenciadores)

Los MCPs de documentos existentes son CRUD genéricos (47+ tools que inflan el contexto)
y no cubren diseño ni operaciones integrales sobre PDF. El Office Worker hace lo contrario:

| | office-mcp / takos | **The Office Worker** |
|---|---|---|
| Tools | ~47–80 CRUD genéricas | **12 poderosas** (~2K tok/turno vs ~6.4K) |
| Plantillas con variables | ❌ | ✅ Jinja (`{{ var }}`) |
| Temas corporativos (paleta/fuente) | ❌ parámetros sueltos | ✅ 5 temas incluidos + logo opcional en header |
| PDF como **input** & procesamiento | ❌ | ✅ PyMuPDF + pdfplumber + pypdf + Tesseract OCR |
| Formularios AcroForm | ❌ | ✅ Relleno determinista (`pdf_fill_form`) |
| Conversión Office → PDF | ❌ | ✅ LibreOffice headless 100% local |
| Manipulación de PDF | ❌ | ✅ Merge, extract y rotate en una sola tool |
| PPTX **editable** (textos nativos, no imágenes) | parcial | ✅ html-to-pptx verificado |
| Guía anti-loop para agentes | ❌ | ✅ SKILL.md con reglas probadas en producción |

## Las 12 tools

| Tool | Qué hace |
|---|---|
| `render_document` | PDF profesional desde plantilla HTML/Jinja + datos + tema (+ logo opcional en header) (WeasyPrint) |
| `create_word` | `.docx` con h1/h2/párrafos/tablas, tema aplicado (python-docx) |
| `create_excel` | `.xlsx` multi-hoja con paleta y filas alternas (openpyxl) |
| `create_pptx` | `.pptx` **editable** desde slides declarativas (html-to-pptx + Playwright) |
| `read_pdf` | Extrae texto por página + metadatos de un PDF (PyMuPDF) |
| `pdf_extract_tables` | Tablas estructuradas de un PDF (pdfplumber) |
| `pdf_list_form_fields` | Campos de formulario AcroForm de un PDF (pypdf) |
| `list_themes` | Devuelve la paleta/fuente del tema activo como referencia de diseño |
| `pdf_fill_form` | Rellena campos de un formulario PDF interactivo (pypdf) |
| `pdf_ocr` | OCR sobre imágenes o PDFs escaneados; opcionalmente genera PDF con capa de texto (Tesseract + fitz) |
| `convert_to_pdf` | Convierte `.docx`, `.xlsx`, `.pptx` a PDF con LibreOffice headless (timeout 120s) |
| `pdf_manipulate` | Manipula PDFs: unir (`merge`), extraer páginas (`extract`), o rotar (`rotate`) (pypdf) |

## Temas corporativos & Logo

El sistema incluye 5 temas listos para usar en cualquier tool mediante el parámetro `theme`:

1. **`aden`** (por defecto): Azul marino oscuro (`#003366`), acento azul (`#3B82F6`), texto oscuro.
2. **`claro`**: Estilo moderno limpio en azul `#2563EB` sobre fondo blanco puro.
3. **`oscuro`**: Modo oscuro elegante con azul `#60A5FA` sobre fondo grafito `#111827`.
4. **`minimal`**: Blanco y negro minimalista tipográfico (fuente serif `Georgia` en títulos).
5. **`corporate-blue`**: Estilo institucional clásico con Pantone Classic Blue `#0F4C81`.

Además, `render_document` soporta el parámetro opcional `logo` (ruta a PNG/JPG). El logo se inyecta automáticamente en la esquina superior derecha (`@top-right`) del encabezado de página mediante CSS nativo de WeasyPrint.

## Uso rápido (sin agente, vía CLI o Python)

```python
from office_worker.core import (
    create_word, create_excel, render_pdf,
    fill_pdf_form, convert_office_to_pdf, manipulate_pdf, ocr_pdf
)

# Generar Word y Excel con tema corporate-blue
create_word("acta.docx", title="Acta", subtitle="Reunión", blocks=[
    {"type":"h2","text":"Puntos"}, {"type":"p","text":"Decisión tomada"}
], theme="corporate-blue")

create_excel("balance.xlsx", title="Balance", sheets=[{"name":"Resumen",
    "headers":["Rubro","Monto"],"rows":[["Ingresos",100],["Gastos",40]]}])

# Convertir Word a PDF
convert_office_to_pdf("acta.docx", "acta.pdf")

# Rellenar formulario PDF
fill_pdf_form("solicitud.pdf", {"nombre": "Ana Gomez", "cargo": "Directora"}, "solicitud_firmada.pdf")

# Manipular PDF: unir dos documentos
manipulate_pdf(operation="merge", out="completo.pdf", files=["acta.pdf", "solicitud_firmada.pdf"])

# OCR de imagen a PDF con texto buscable
ocr_pdf("recibo_escaneado.png", lang="spa", out="recibo_buscable.pdf")
```

## Requisitos locales

- Python ≥ 3.9 · WeasyPrint (necesita fontconfig/pango en Linux) · LibreOffice (`soffice`) para conversión Office → PDF.
- Tesseract OCR (`tesseract`) para la herramienta `pdf_ocr`.
- **PPTX editable** (opcional): `pip install office-worker-mcp[pptx]` + `playwright install chromium`.
  Sin este extra, las tools de PDF/Word/Excel funcionan igual; solo `create_pptx` queda deshabilitada.
- Todo corre en local; sin dependencias cloud ni telemetría.

## Desarrollo & tests

```bash
pip install -e ".[dev]"
pytest -q            # test_core.py (núcleo) + test_e2e_mcp.py (MCP end-to-end, 12/12 tools)
```

El test E2E levanta el servidor MCP real, hace handshake y ejecuta cada tool creando/leyendo archivos válidos en disco.

## Licencia

MIT — ver [LICENSE](LICENSE). Agradecimientos en [NOTICE](NOTICE).

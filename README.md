# The Office Worker — `office-worker-mcp`

> **Your agent's document clerk.** Genera documentos de oficina profesionales —
> **PDF, Word, Excel y PowerPoint** — desde datos + un tema corporativo, y lee PDFs
> (texto, tablas y formularios). Diseñado para agentes de IA (Hermes, Claude, Cursor)
> vía [MCP](https://modelcontextprotocol.io), 100% local-first.

```
pip install office-worker-mcp          # instala librería + MCP server + CLI 'owi'
```

Conectalo a tu agente (ejemplo config MCP / stdio):

```json
{ "mcpServers": { "office-worker": { "command": "office-worker-mcp" } } }
```

Y listo: le pedís *"haceme una factura/informe/deck"* en lenguaje natural y tu agente
usa las tools del oficinista para producir el archivo bien diseñado.

## Por qué existe (diferenciadores)

Los MCPs de documentos existentes son CRUD genéricos (47+ tools que inflan el contexto)
y no cubren diseño ni PDF como entrada. El Office Worker hace lo contrario:

| | office-mcp / takos | **The Office Worker** |
|---|---|---|
| Tools | ~47–80 CRUD genéricas | **8 poderosas** (~2K tok/turno vs ~6.4K) |
| Plantillas con variables | ❌ | ✅ Jinja (`{{ var }}`) |
| Temas corporativos (paleta/fuente) | ❌ parámetros sueltos | ✅ aplicable de una vez a todos los formatos |
| PDF como **input** (leer/tablas/formularios) | ❌ | ✅ PyMuPDF + pdfplumber + pypdf |
| PPTX **editable** (textos nativos, no imágenes) | parcial | ✅ html-to-pptx verificado |
| Guía anti-loop para agentes | ❌ | ✅ SKILL.md con reglas probadas en producción |

## Las 8 tools

| Tool | Qué hace |
|---|---|
| `render_document` | PDF profesional desde plantilla HTML/Jinja + datos + tema (WeasyPrint) |
| `create_word` | `.docx` con h1/h2/párrafos/tablas, tema aplicado (python-docx) |
| `create_excel` | `.xlsx` multi-hoja con paleta y filas alternas (openpyxl) |
| `create_pptx` | `.pptx` **editable** desde slides declarativas (html-to-pptx + Playwright) |
| `read_pdf` | Extrae texto por página + metadatos de un PDF (input) |
| `pdf_extract_tables` | Tablas estructuradas de un PDF (pdfplumber) |
| `pdf_list_form_fields` | Campos de formulario AcroForm de un PDF (pypdf) |
| `list_themes` | Devuelve la paleta/fuente del tema activo como referencia de diseño |

## Uso rápido (sin agente, vía CLI o Python)

```python
from office_worker.core import create_word, create_excel, render_pdf

create_word("acta.docx", title="Acta", subtitle="Reunión", blocks=[
    {"type":"h2","text":"Puntos"}, {"type":"p","text":"Decisión tomada"}])

create_excel("balance.xlsx", title="Balance", sheets=[{"name":"Resumen",
    "headers":["Rubro","Monto"],"rows":[["Ingresos",100],["Gastos",40]]}])
```

## Requisitos locales

- Python ≥ 3.9 · WeasyPrint (necesita fontconfig/pango en Linux) · LibreOffice opcional.
- Para PPTX: Playwright + Chromium (`playwright install chromium`). Todo corre en local; sin cloud.

## Desarrollo & tests

```bash
pip install -e ".[dev]"
pytest            # test_core.py (núcleo) + test_e2e_mcp.py (MCP end-to-end, 8/8 tools)
```

El test E2E levanta el servidor MCP real, hace handshake y ejecuta cada tool creando/leyendo archivos válidos en disco.

## Licencia

MIT — ver [LICENSE](LICENSE). Agradecimientos en [NOTICE](NOTICE).

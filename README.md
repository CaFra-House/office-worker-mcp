# The Office Worker — `office-worker-mcp`

![CI](https://github.com/CaFra-House/office-worker-mcp/actions/workflows/ci.yml/badge.svg)
![PyPI version](https://img.shields.io/pypi/v/office-worker-mcp)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/office-worker-mcp)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

> **Your agent's document clerk.** Genera y edita documentos de oficina profesionales —
> **PDF, Word, Excel y PowerPoint** — desde datos + un tema corporativo, plantillas empaquetadas (docxtpl),
> edición in-place, gráficos nativos Excel, extracción de imágenes para visión, conversión PDF a Excel,
> pipelines batch y reporte honesto de fidelidad. Diseñado para agentes de IA (Hermes, Claude, Cursor)
> vía [MCP](https://modelcontextprotocol.io), 100% local-first, sin API keys, determinista y seguro.

```bash
pip install office-worker-mcp          # instala librería + MCP server + CLI 'owi'
pip install "office-worker-mcp[ocr]"   # extra para OCR con Tesseract + Pillow
pip install "office-worker-mcp[sign]"  # extra para firma digital PAdES con pyhanko
```

Conectalo a tu agente (ejemplo config MCP / stdio):

```json
{ "mcpServers": { "office-worker": { "command": "office-worker-mcp" } } }
```

## Por qué existe (diferenciadores)

Los MCPs de documentos existentes son CRUD genéricos (47–80 tools que inflan el contexto)
y no cubren diseño, fidelidad honesta ni operaciones integrales sobre documentos. The Office Worker hace lo contrario:

| Capacidad | office-mcp / takos | **The Office Worker (v0.5.0)** |
|---|---|---|
| Tools & Contexto | ~47–80 CRUD genéricas (>6.4K tok) | **20 herramientas especializadas** (~1960 tok/turno total, ~98 tok/tool) |
| Batch pipelines | ❌ Múltiples turnos lentos | ✅ `office_batch`: ejecuta secuencias en un solo turno con tolerancia a fallas parciales |
| Edición in-place | ❌ Rehacer archivos enteros | ✅ `edit_excel` y `edit_word` preservando estilos originales y macros VBA |
| Reporte de fidelidad honesta | ❌ Silencio o afirmaciones falsas | ✅ `fidelity` (`rich` \| `clean` \| `lossy`) + `warnings` transparentes |
| Visión en PDFs | ❌ Solo texto plano | ✅ `read_pdf(extract_images=True)` devuelve imágenes base64 para agentes multimodales |
| PDF a Excel | ❌ No disponible | ✅ `pdf_to_excel` extrae tablas de PDFs directamente a hojas `.xlsx` estilizadas |
| Extracción Office unificada | ❌ Requiere parsers dispersos | ✅ `read_office` lee `.docx`, `.pptx` y `.xlsx` en formato estructurado |
| Gráficos nativos Excel | ❌ Solo celdas planas | ✅ Gráficos nativos de barras, líneas y torta embebidos en `.xlsx` |
| Tablas estructuradas & Fórmulas | ❌ Solo texto crudo | ✅ Tablas Excel con estilos oficiales, `auto_filter` y fórmulas (`SUM`, `VLOOKUP`, etc.) |
| Macros VBA | ❌ Corrupción de archivos `.xlsm` | ✅ Preservación explícita de código macro vía `keep_vba=True` (sin ejecución VBA) |
| Plantillas Word empaquetadas | ❌ | ✅ 5 plantillas oficiales listas con Jinja/docxtpl (`list_templates`) |
| Temas corporativos (paleta/fuente) | ❌ parámetros sueltos | ✅ 5 temas corporativos listos + logo en header |
| Watermark & Footers | ❌ | ✅ Marca de agua diagonal + pie de página configurable (@page) |
| Firma de PDF | ❌ | ✅ Sello PNG visible siempre + firma digital PAdES si hay certificado PEM |
| Compresión de PDF | ❌ | ✅ Optimización PyMuPDF con recolección de basura y downsampling |
| Conversión Office → PDF | ❌ | ✅ LibreOffice headless 100% local con aislamiento de subproceso |
| Reglas duras anti-loop | ❌ | ✅ `SKILL.md` probado: detención obligatoria tras 2 fallos consecutivos |

---

## Las 20 Tools (v0.5.0)

| Tool | Qué retorna | Cuándo usarla | Cuándo NO usarla |
|---|---|---|---|
| `office_batch` | Resumen JSON `{status, total, succeeded, failed, results}` | Para ejecutar secuencias de operaciones documentales en un solo turno (ej: crear docx + convertir a pdf + firmar) | No usar para operaciones aisladas de un solo paso o con bifurcación conversacional |
| `edit_excel` | JSON `{status, path, bytes, fidelity, warnings}` | Para modificar celdas, agregar filas/columnas, tablas o gráficos en archivos `.xlsx`/`.xlsm` existentes | No usar para crear planillas desde cero (usar `create_excel`) ni para crear macros VBA |
| `edit_word` | JSON `{status, path, bytes, fidelity, warnings}` | Para editar `.docx` existentes (agregar párrafos, reemplazar texto, insertar tras títulos, agregar tablas) | No usar para crear documentos nuevos desde cero (usar `create_word`) ni para editar PDFs |
| `pdf_to_excel` | JSON `{status, path, bytes, n_tables, fidelity, warnings}` | Para extraer tablas desde PDFs (balances, facturas, informes) a planillas `.xlsx` limpias | No usar para PDFs escaneados sin tablas seleccionables (usar `pdf_ocr`) |
| `read_office` | JSON `{status, format, paragraphs/slides/sheets, text}` | Para extraer texto y datos estructurados directamente de `.docx`, `.pptx` o `.xlsx` | No usar para archivos PDF (usar `read_pdf`) |
| `read_pdf` | JSON `{pages, metadata, tables, fields, images}` | Para inspeccionar texto, tablas, formularios e imágenes base64 de cualquier PDF | No usar para PDFs escaneados que requieran OCR completo (usar `pdf_ocr`) |
| `render_document` | JSON `{status, path, bytes}` | Para generar PDFs impecables desde plantillas HTML/Jinja + tema corporativo + logo + watermark | No usar para editar documentos existentes (usar `edit_word`/`pdf_manipulate`) |
| `create_word` | JSON `{status, path, bytes}` | Para crear documentos Word `.docx` nuevos desde bloques o plantillas empaquetadas | No usar para editar documentos en disco (usar `edit_word`) ni para generar PDFs directos |
| `create_excel` | JSON `{status, path, bytes}` | Para crear libros Excel `.xlsx` multi-hoja con tablas estructuradas, autofiltro y gráficos | No usar para actualizar libros existentes (usar `edit_excel`) |
| `create_pptx` | JSON `{status, path, bytes}` | Para crear presentaciones PowerPoint editables (texto nativo, kickers, bullets, tablas) | No usar para documentos de texto corrido o reportes en PDF |
| `convert_to_pdf` | JSON `{status, path, bytes, fidelity, warnings}` | Para exportar archivos Office (`.docx`, `.xlsx`, `.pptx`) a PDF localmente con LibreOffice | No usar cuando se genera un documento nuevo desde plantilla HTML (usar `render_document`) |
| `sign_pdf` | JSON `{status, path, bytes}` | Para estampar sello visual PNG y aplicar firma digital criptográfica PAdES con certificado PEM | No usar para editar contenido textual de un documento |
| `pdf_compress` | JSON `{status, path, bytes, savings_percent}` | Para optimizar y comprimir PDFs pesados reduciendo imágenes y limpiando objetos | No usar en PDFs de texto puro donde no hay imágenes que optimizar |
| `pdf_manipulate` | JSON `{status, path, bytes}` | Para unir (`merge`), dividir/extraer páginas (`extract`) o rotar páginas con clave opcional | No usar para modificar texto o imágenes dentro de las páginas |
| `pdf_fill_form` | JSON `{status, path, bytes}` | Para rellenar formularios interactivos AcroForm con campos clave-valor | No usar en PDFs planos o escaneados sin campos de formulario |
| `pdf_ocr` | JSON `{status, text, path?}` | Para extraer texto y generar PDF buscable sobre imágenes o escaneos con Tesseract | No usar sobre PDFs digitales que ya tengan capa de texto seleccionable |
| `list_templates` | JSON `{status, templates}` | Para consultar el catálogo de plantillas oficiales `.docx` y las variables que requieren | No usar para generar archivos (usar `create_word`) |
| `list_themes` | JSON `{status, theme}` | Para consultar paleta de colores y fuentes de temas corporativos | No usar para generar documentos directamente |
| `pdf_extract_tables` | *(Deprecated)* JSON con tablas | Preferir `read_pdf(path, extract_tables=True)` | No usar para lectura integral |
| `pdf_list_form_fields`| *(Deprecated)* JSON con campos | Preferir `read_pdf(path, list_forms=True)` | No usar para lectura integral |

---

## Nuevas Capacidades v0.5.0

### 1. Batch Operations (`office_batch`)
Permite despachar una lista secuencial de operaciones en una sola llamada de herramienta. Si un paso falla, el despachador continúa ejecutando los pasos restantes y reporta exactamente qué falló:

```json
{
  "operations": [
    {
      "tool": "create_word",
      "args": { "out_path": "/workspace/informe.docx", "template_docx": "informe_ejecutivo", "context": { "titulo": "Reporte Q3" } }
    },
    {
      "tool": "convert_to_pdf",
      "args": { "input_file": "/workspace/informe.docx", "output": "/workspace/informe.pdf" }
    },
    {
      "tool": "sign_pdf",
      "args": { "input_pdf": "/workspace/informe.pdf", "output": "/workspace/informe_firmado.pdf", "sello_img_path": "/workspace/sello.png" }
    }
  ]
}
```

### 2. Edición In-Place (`edit_excel` y `edit_word`)
Modifica documentos existentes en el mismo archivo o en uno nuevo sin perder estilos previos:
- **`edit_excel`**: Modifica celdas (`set_cell`), agrega filas (`append_row`), agrega columnas (`add_column`), agrega tablas estilizadas (`add_table`), activa filtros (`auto_filter`) e inserta gráficos (`add_chart`).
- **`edit_word`**: Agrega párrafos (`append_paragraph`), reemplaza texto preservando formato de run (`replace_text`), inserta contenido inmediatamente después de encabezados específicos (`insert_after_heading`) o añade tablas (`append_table`).

### 3. Reporte de Fidelidad Honesta
Tanto `convert_to_pdf` como las herramientas de edición devuelven:
- `fidelity`: `"rich"` (edición sin pérdida o nativa), `"clean"` (conversión estándar de alta calidad), o `"lossy"` (degradación esperable por limitaciones de formato).
- `warnings`: Lista explícita de advertencias (ej: sustitución de fuentes tipográficas, saltos de página automáticos en hojas anchas, macros no ejecutadas). No se inventa fidelidad alta falsa.

### 4. Visión de Documentos: Extracción de Imágenes en PDF
`read_pdf(path, extract_images=True, max_images=10)` extrae las imágenes embebidas en el PDF y las retorna como data URLs en Base64 (`data:image/png;base64,...`). Esto permite a los modelos de visión de los agentes inspeccionar logotipos, firmas, gráficos o fotos directamente.

### 5. PDF a Excel (`pdf_to_excel`)
Convierte tablas encerradas en informes PDF a un archivo Excel `.xlsx` estructurado, formateando encabezados con paleta corporativa y ajustando automáticamente el ancho de columnas.

### 6. Lector Estructurado de Office (`read_office`)
Extrae de manera uniforme párrafos de Word (`.docx`), diapositivas y textos de PowerPoint (`.pptx`), y hojas con datos de Excel (`.xlsx`), sin requerir LibreOffice ni conversión previa a PDF.

### 7. Gráficos Nativos y Tablas en Excel
Tanto `create_excel` como `edit_excel` soportan gráficos nativos de Excel (`bar`, `line`, `pie`), tablas estructuradas con estilos oficiales (`TableStyleMedium9`), autofiltros y fórmulas (`SUM`, `SUMIF`, `AVERAGEIF`, `VLOOKUP`, `XLOOKUP`, `COUNTIFS`).
> [!NOTE]
> **Macros VBA:** openpyxl no crea ni ejecuta código VBA. Para archivos `.xlsm`, las macros existentes se preservan de forma segura utilizando `keep_vba=True`.

---

## Ejemplos Conversacionales Globales (EN)

### Example 1: Creating and signing an executive report in batch
*User:* "Generate an executive report for Q3 with our financial indicators, export it to PDF, and apply our corporate seal."
*Agent Workflow:*
The agent invokes `office_batch`:
1. `create_word`: Uses template `informe_ejecutivo` with revenue and KPI context.
2. `convert_to_pdf`: Converts the `.docx` into `.pdf` with headless LibreOffice.
3. `sign_pdf`: Stamps the corporate signature seal PNG on the final page.

### Example 2: Extracting financial tables from PDF to Excel
*User:* "I have an audit report in PDF containing balance sheets. Can you extract the tables into an Excel spreadsheet?"
*Agent Workflow:*
The agent calls `pdf_to_excel(input_path="audit_2026.pdf", output_path="audit_tables.xlsx")`, which extracts the data cleanly and reports fidelity and table count.

### Example 3: Updating an existing spreadsheet and adding a chart
*User:* "In our budget spreadsheet `budget.xlsx`, set cell B3 to 12500, append an expenses row, and add a bar chart titled Revenue Overview."
*Agent Workflow:*
The agent calls `edit_excel(input_path="budget.xlsx", operations=[{"op":"set_cell","coordinate":"B3","value":12500}, {"op":"append_row","row":["Cloud Ops", 3200]}, {"op":"add_chart","chart_type":"bar","title":"Revenue Overview"}])`.

### Example 4: Agent multimodal vision on PDF images
*User:* "What does the stamp or logo inside this contract PDF look like?"
*Agent Workflow:*
The agent calls `read_pdf(path="contract.pdf", extract_images=True, max_images=3)`, receives base64 data URLs, and inspects the visual elements.

---

## Seguridad (Local-First)

- **Validación estricta de rutas (`safe_out`):** Bloquea path traversal y sobreescritura en rutas críticas (`/etc`, `/boot`, `/root`, `/sys`).
- **Aislamiento WeasyPrint (`safe_url_fetcher`):** Bloquea llamadas de red remotas (SSRF) y lecturas de archivos protegidos (LFI).
- **Aislamiento de procesos:** Conversiones de LibreOffice con timeout estricto de 120s y terminación de proceso aislada (`killpg`).
- **Cifrado AES:** Soporte de protección por contraseña tanto en generación como en manipulación.

## Desarrollo & Tests

```bash
pip install -e ".[dev]"
pytest -q            # 26 tests verificados en disco: test_core.py + test_e2e_mcp.py
python count_tokens.py # Auditoría en tiempo real de tokens por schema
```

## Licencia

MIT — ver [LICENSE](LICENSE). Agradecimientos en [NOTICE](NOTICE).



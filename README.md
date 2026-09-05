# The Office Worker — `office-worker-mcp`

![CI](https://github.com/CaFra-House/office-worker-mcp/actions/workflows/ci.yml/badge.svg)
![PyPI version](https://img.shields.io/pypi/v/office-worker-mcp)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/office-worker-mcp)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

> **Your agent's document clerk.** Genera, edita, asegura y extrae documentos de oficina profesionales —
> **PDF, Word, Excel y PowerPoint** — desde datos + tema corporativo, plantillas empaquetadas (docxtpl),
> edición in-place, gráficos nativos Excel, extracción de imágenes para visión, preview PNG para inspección previa a entrega,
> redacción irreversible permanente (`pdf_redact`), aplanado para cumplimiento de archivo (`flatten`),
> extracción estructurada Markdown/JSON para RAG (`pdf_extract_structured`), división inteligente por límites (`split_smart`),
> conversión PDF a Excel, skills empaquetadas y reporte honesto de fidelidad. Diseñado para agentes de IA (Hermes, Claude, Cursor)
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

| Capacidad | office-mcp / takos | **The Office Worker (v0.6.0)** |
|---|---|---|
| Tools & Contexto | ~47–80 CRUD genéricas (>6.4K tok) | **21 herramientas especializadas** (~1880 tok/turno total, <1900 tok, ~89 tok/tool) |
| Redacción Irreversible | ❌ No disponible | ✅ `pdf_redact`: borrado permanente de PII/claves por texto o coordenadas con PyMuPDF |
| Extracción RAG (Markdown / JSON) | ❌ Solo texto plano | ✅ `pdf_extract_structured`: tablas pipe Markdown y JSON estructurado para bases de conocimiento |
| Aplanado de Formularios | ❌ No disponible | ✅ `pdf_manipulate(op="flatten")`: sella AcroForms a contenido estático para archivo seguro |
| División Inteligente (Split) | ❌ Solo rangos fijos | ✅ `pdf_manipulate(op="split_smart")`: detecta límites heurísticos (páginas en blanco / carátulas) |
| Skills Empaquetadas & CLI | ❌ No disponible | ✅ `owi skill install [name]`: despliega manuales y flujos oficiales en `~/.hermes/skills/` |
| Vista Previa PNG / Visión | ❌ No disponible | ✅ `pdf_preview`: renderiza páginas a PNG (base64 data URL) para inspección visual previa |
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

## Las 21 Tools (v0.6.0)

| Tool | Qué retorna | Cuándo usarla | Cuándo NO usarla |
|---|---|---|---|
| `pdf_redact` | JSON `{status, path, bytes, redactions_count}` | Para borrar permanentemente información confidencial (texto o coordenadas) | No usar para edición normal de texto (usar `edit_word`) |
| `pdf_extract_structured` | JSON `{status, format, content?, pages, n_tables}` | Para extraer texto y tablas en Markdown (pipes) o JSON estructurado para RAG | No usar en PDFs escaneados sin capa de texto (usar `pdf_ocr`) |
| `pdf_preview` | JSON `{status, data_url, pages, path?, bytes?}` | Para renderizar páginas PDF a imagen PNG y revisarlo visualmente antes de entregarlo | No usar para extraer texto seleccionable (usar `read_pdf`) |
| `read_pdf` | JSON `{pages, metadata, tables?, fields?, images?}` | Herramienta todo-en-uno: texto, metadatos, tablas (`extract_tables`), forms (`list_forms`), imágenes base64 (`extract_images`) | No usar para PDFs escaneados que requieran OCR completo (usar `pdf_ocr`) |
| `office_batch` | Resumen JSON `{status, total, succeeded, failed, results}` | Para ejecutar secuencias de operaciones documentales en un solo turno sin roundtrips | No usar para operaciones aisladas de un solo paso |
| `edit_excel` | JSON `{status, path, bytes, fidelity, warnings}` | Para modificar celdas, agregar filas/columnas, tablas o gráficos en `.xlsx`/`.xlsm` existentes | No usar para crear planillas desde cero (usar `create_excel`) |
| `edit_word` | JSON `{status, path, bytes, fidelity, warnings}` | Para editar `.docx` existentes (agregar párrafos, reemplazar texto, insertar tras títulos, tablas) | No usar para crear documentos nuevos desde cero (usar `create_word`) |
| `pdf_to_excel` | JSON `{status, path, bytes, n_tables, fidelity, warnings}` | Para extraer tablas desde PDFs (balances, facturas, informes) a planillas `.xlsx` limpias | No usar para PDFs escaneados sin tablas seleccionables (usar `pdf_ocr`) |
| `read_office` | JSON `{status, format, paragraphs/slides/sheets, text}` | Para extraer texto y datos estructurados directamente de `.docx`, `.pptx` o `.xlsx` | No usar para archivos PDF (usar `read_pdf`) |
| `render_document` | JSON `{status, path, bytes}` | Para generar PDFs impecables desde plantillas HTML/Jinja + tema corporativo + logo + watermark | No usar para editar documentos existentes (usar `edit_word`/`pdf_manipulate`) |
| `create_word` | JSON `{status, path, bytes}` | Para crear documentos Word `.docx` nuevos desde bloques o plantillas empaquetadas | No usar para editar documentos en disco (usar `edit_word`) |
| `create_excel` | JSON `{status, path, bytes}` | Para crear libros Excel `.xlsx` multi-hoja con tablas estructuradas, autofiltro y gráficos | No usar para actualizar libros existentes (usar `edit_excel`) |
| `create_pptx` | JSON `{status, path, bytes}` | Para crear presentaciones PowerPoint editables (texto nativo, kickers, bullets, tablas) | No usar para documentos de texto corrido o reportes en PDF |
| `convert_to_pdf` | JSON `{status, path, bytes, fidelity, warnings}` | Para exportar archivos Office (`.docx`, `.xlsx`, `.pptx`) a PDF localmente con LibreOffice | No usar cuando se genera un documento nuevo desde plantilla HTML |
| `sign_pdf` | JSON `{status, path, bytes}` | Para estampar sello visual PNG y aplicar firma digital criptográfica PAdES con certificado PEM | No usar para editar contenido textual de un documento |
| `pdf_compress` | JSON `{status, path, bytes, savings_percent}` | Para optimizar y comprimir PDFs pesados reduciendo imágenes y limpiando objetos | No usar en PDFs de texto puro donde no hay imágenes |
| `pdf_manipulate` | JSON `{status, path/files, bytes, warnings?}` | Para unir (`merge`), extraer (`extract`/`split_by`), aplanar (`flatten`), rotar (`rotate`) o dividir inteligentemente (`split_smart`) | No usar para modificar texto dentro de las páginas |
| `pdf_fill_form` | JSON `{status, path, bytes}` | Para rellenar formularios interactivos AcroForm con campos clave-valor | No usar en PDFs planos o escaneados sin campos de formulario |
| `pdf_ocr` | JSON `{status, text, path?}` | Para extraer texto y generar PDF buscable sobre imágenes o escaneos con Tesseract | No usar sobre PDFs digitales que ya tengan texto seleccionable |
| `list_templates` | JSON `{status, templates}` | Para consultar el catálogo de plantillas oficiales `.docx` y las variables que requieren | No usar para generar archivos (usar `create_word`) |
| `list_themes` | JSON `{status, theme}` | Para consultar paleta de colores y fuentes de temas corporativos | No usar para generar documentos directamente |

---

## Capacidades Destacadas (v0.6.0)

### 1. Redacción Irreversible y Permanente (`pdf_redact`)
Permite al agente censurar permanentemente información confidencial (PII, números de tarjetas, contraseñas, saldos o nombres) antes de distribuir un documento PDF.
Soporta búsqueda por texto o regiones rectangulares explícitas `[x0, y0, x1, y1]`, aplicando anotaciones de redacción vía PyMuPDF (`page.apply_redactions()`). A diferencia de simples rectángulos negros superpuestos, la redacción **destruye físicamente el contenido textual y vectorial del stream del PDF**, haciendo imposible su recuperación o copia:

```json
{
  "input_path": "/workspace/declaracion_jurada.pdf",
  "output": "/workspace/declaracion_censurada.pdf",
  "search_text": "30-71234567-9",
  "regions": [{"page": 1, "x0": 120, "y0": 340, "x1": 280, "y1": 365}],
  "fill_color": "black"
}
```

### 2. Aplanado de Formularios para Cumplimiento de Archivo (`flatten`)
Convierte campos de formularios interactivos AcroForm y anotaciones en contenido gráfico y textual estático mediante PyMuPDF (`doc.bake(annots=True, widgets=True)`). Garantiza que contratos y formularios completados no puedan ser manipulados ni alterados posteriormente:

```json
{
  "operation": "flatten",
  "input_path": "/workspace/formulario_relleno.pdf",
  "output": "/workspace/formulario_archivado.pdf"
}
```

### 3. Extracción Estructurada Markdown / JSON para RAG (`pdf_extract_structured`)
Optimizado para flujos de trabajo de ingesta de conocimiento y RAG (Retrieval-Augmented Generation). Extrae texto limpio, metadatos y matrices de tablas mediante `pdfplumber` + `PyMuPDF`:
- **`format: "markdown"`**: Genera encabezados jerárquicos (`# Titulo`, `## Página N`) y tablas convertidas a formato pipe Markdown estándar (`| Col | Col |`).
- **`format: "json"`**: Retorna objetos `{pages: [{page, text, tables, images_n}]}` estructurados para segmentación semántica de datos.

```json
{
  "input_path": "/workspace/balance_anual.pdf",
  "format": "markdown",
  "output": "/workspace/balance_rag.md"
}
```

### 4. División Inteligente de Documentos (`split_smart`)
Detecta heurísticamente los límites entre documentos concatenados (por páginas separadoras en blanco sin texto ni gráficos, o por carátulas/encabezados repetidos) y los divide automáticamente en $N$ archivos independientes válidos:

```json
{
  "operation": "split_smart",
  "input_path": "/workspace/lote_escaneado.pdf",
  "output": "/workspace/documento_part.pdf"
}
```
*Honestidad técnica:* Si no se detecta un límite confiable, la herramienta emite advertencias honestas (`warnings`) y mantiene el documento íntegro en lugar de inventar particiones arbitrarias.

### 5. Skills Empaquetadas & Comando CLI (`owi skill install`)
The Office Worker empaqueta manuales de operación (`SKILL.md`) con ejemplos conversacionales en inglés, reglas duras anti-loop y flujos opcionales de Google Drive / Gmail (`workspace-mcp`).
Instala la skill en el entorno de tu agente con un solo comando idempotente:

```bash
owi skill list                       # Lista las skills oficiales disponibles
owi skill install office-worker      # Copia la skill a ~/.hermes/skills/office-worker/SKILL.md
owi skill install all                # Instala todas las skills empaquetadas
```

### 6. Preview PNG para Visión de Agentes (`pdf_preview`)
Renderiza páginas PDF a imágenes PNG de alta resolución (`dpi=110`) mediante PyMuPDF (`page.get_pixmap()`), retornando Data URLs en base64 (`data:image/png;base64,...`) para inspección visual previa a la entrega final.

### 7. Batch Operations (`office_batch`)
Despacha secuencias completas de operaciones documentales en un solo turno con manejo de errores parciales:
```json
{
  "operations": [
    { "tool": "create_word", "args": { "out_path": "rep.docx", "template_docx": "informe_ejecutivo", "context": { "titulo": "Reporte Q3" } } },
    { "tool": "convert_to_pdf", "args": { "input_file": "rep.docx", "output": "rep.pdf" } },
    { "tool": "pdf_redact", "args": { "input_path": "rep.pdf", "output": "rep_redacted.pdf", "search_text": "CONFIDENCIAL" } },
    { "tool": "pdf_preview", "args": { "input_path": "rep_redacted.pdf", "max_pages": 1 } }
  ]
}
```

### 8. Edición In-Place (`edit_excel` y `edit_word`)
Modifica documentos existentes en el mismo archivo o en uno nuevo sin perder estilos previos:
- **`edit_excel`**: Modifica celdas (`set_cell`), agrega filas (`append_row`), agrega columnas (`add_column`), agrega tablas estilizadas (`add_table`), activa filtros (`auto_filter`) e inserta gráficos (`add_chart`).
- **`edit_word`**: Agrega párrafos (`append_paragraph`), reemplaza texto preservando formato de run (`replace_text`), inserta contenido tras encabezados (`insert_after_heading`) o añade tablas (`append_table`).

### 9. Reporte de Fidelidad Honesta
`convert_to_pdf`, `edit_excel`, `edit_word` y `pdf_to_excel` devuelven:
- `fidelity`: `"rich"` (edición sin pérdida o nativa), `"clean"` (conversión estándar de alta fidelidad), o `"lossy"` (degradación esperable por limitaciones de formato).
- `warnings`: Lista explícita de advertencias (sustitución tipográfica, saltos de página automáticos en planillas anchas, macros no ejecutadas).

---

## Limitaciones Conocidas de Plataforma (Honestas)

Para garantizar total transparencia técnica frente a los usuarios y administradores del sistema:

1. **Soporte Windows:**
   - WeasyPrint requiere librerías nativas C compiladas (`libpango-1.0-0`, `libpangocairo`, `libgdk-pixbuf`, `libffi`) que no disponen de un instalador desatendido oficial en Windows sin recurrir a pilas complejas de MSYS2 o gvsbuild.
   - La conversión headless de LibreOffice en Windows requiere invocar binarios en rutas del Registro como `C:\Program Files\LibreOffice\program\soffice.exe`.
   - Por estas razones, `windows-latest` queda excluido de la matriz de CI oficial. En entornos Windows, **se recomienda ejecutar bajo WSL2 (Ubuntu) o dentro de un contenedor Linux Docker**.
2. **LibreOffice Headless:**
   - La conversión a PDF con LibreOffice headless puede generar sustitución de fuentes si las tipografías corporativas exactas no están instaladas en el sistema anfitrión.
   - En hojas de cálculo anchas sin área de impresión fijada, LibreOffice puede insertar saltos de página automáticos. Esto siempre se reporta en `warnings`.
3. **Macros VBA:**
   - `openpyxl` preserva las macros de archivos `.xlsm` de forma segura mediante `keep_vba=True`, pero **no ejecuta código VBA**.
4. **Presentaciones PowerPoint (`create_pptx`):**
   - Requiere el extra opcional `[pptx]` junto con Playwright y Chromium instalado (`playwright install chromium`).
5. **Firmas Digitales Criptográficas:**
   - El estampado visual PNG funciona siempre. La firma criptográfica digital PAdES requiere un certificado X.509 válido en formato PEM.

---

## CI Multiplataforma

El repositorio cuenta con integración continua automatizada en GitHub Actions (`.github/workflows/ci.yml`) con matriz multiplataforma:
- **Ubuntu:** `ubuntu-latest` (x86_64) en Python 3.11 y 3.12, más `ubuntu-24.04-arm` (ARM64 experimental).
- **macOS:** `macos-latest` (Apple Silicon ARM64) en Python 3.11 y 3.12 con Homebrew (`pango`, `cairo`, `libffi`, `tesseract`).
- **Verificación Smoke:** Ejecución obligatoria de suite completa de pruebas (`pytest -q`), verificación de herramientas CLI (`owi --help`) y auditoría de presupuesto de tokens (`count_tokens.py`).

---

## Desarrollo & Tests

```bash
pip install -e ".[dev]"
pytest -q            # Suite completa verificada en disco: test_core.py + test_e2e_mcp.py
python count_tokens.py # Auditoría en tiempo real de tokens por schema (<1900 tok v0.6.0)
```

## Licencia

MIT — ver [LICENSE](LICENSE). Agradecimientos en [NOTICE](NOTICE).

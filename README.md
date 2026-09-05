# The Office Worker — `office-worker-mcp`

![CI](https://github.com/CaFra-House/office-worker-mcp/actions/workflows/ci.yml/badge.svg)
![PyPI version](https://img.shields.io/pypi/v/office-worker-mcp)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/office-worker-mcp)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

> **Your agent's document clerk.** Genera, edita, asegura, audita y extrae documentos de oficina profesionales —
> **PDF, Word, Excel y PowerPoint** — desde datos + tema corporativo, plantillas empaquetadas (docxtpl),
> combinación de correspondencia por lotes (`mail_merge`), conversión bidireccional CSV-Excel (`csv_excel_convert`),
> extracción Office a Markdown/JSON para RAG (`read_office`), comparación honesta de revisiones (`document_diff`),
> limpieza de metadatos sensibles (`scrub_metadata`), protección con clave estándar Office (`protect_office`),
> verificación criptográfica de firmas digitales (`verify_pdf_signature`), guiado proactivo con recomendaciones (`next_steps`),
> gráficos nativos en PowerPoint y Excel, edición in-place, preview PNG para inspección previa a entrega,
> redacción irreversible permanente (`pdf_redact`), aplanado para archivo (`flatten`),
> división inteligente por límites (`split_smart`), conversión PDF a Excel, skills empaquetadas y reporte honesto de fidelidad.
> Diseñado para agentes de IA (Hermes, Claude, Cursor) vía [MCP](https://modelcontextprotocol.io), 100% local-first, sin API keys, determinista y seguro.

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

| Capacidad | office-mcp / takos | **The Office Worker (v0.8.0)** |
|---|---|---|
| Tools & Contexto | ~47–80 CRUD genéricas (>6.4K tok) | **27 herramientas especializadas** (~2335 tok/turno total, <2350 tok, ~86 tok/tool) |
| Comparación de Documentos | ❌ No disponible | ✅ `document_diff`: diff textual honesto (párrafos agregados/eliminados/modificados) Word/PDF con warnings |
| Limpieza de Metadatos | ❌ No disponible | ✅ `scrub_metadata`: borrado de autor, título, revisiones y propiedades en PDF, Word, Excel y PowerPoint |
| Cifrado / Contraseña Office | ❌ No disponible | ✅ `protect_office`: cifrado robusto ECMA-376 agile AES con contraseña para `.docx`, `.xlsx`, `.pptx` |
| Verificación de Firmas | ❌ No disponible | ✅ `verify_pdf_signature`: validación criptográfica de integridad y certificados X.509 en PDF (pyhanko) |
| Guiado Proactivo | ❌ Ejecución ciega | ✅ `next_steps` en tools de creación + skill de orquestación (`orchestration.SKILL.md`) |
| Mail Merge Word | ❌ No disponible | ✅ `mail_merge`: genera N documentos `.docx` desde plantillas docxtpl + dataset CSV/JSON |
| Office a Markdown / RAG | ❌ Solo texto plano disperso | ✅ `read_office(format="markdown")`: encabezados (`#`) y tablas pipe (`\|...\|`) de Word, Excel y PPTX |
| Gráficos Nativos PowerPoint | ❌ Solo imágenes o HTML plano | ✅ `create_pptx`: diapositivas con gráficos nativos DrawingML (barras, líneas, torta) |
| CSV <-> Excel Bidireccional | ❌ No disponible | ✅ `csv_excel_convert`: CSV a `.xlsx` estructurado (tablas/autofiltro) y `.xlsx` a CSV con warnings de tipo |
| Redacción Irreversible | ❌ No disponible | ✅ `pdf_redact`: borrado permanente de PII/claves por texto o coordenadas con PyMuPDF |
| Extracción RAG PDF (Markdown / JSON) | ❌ Solo texto plano | ✅ `pdf_extract_structured`: tablas pipe Markdown y JSON estructurado para bases de conocimiento |
| Aplanado de Formularios | ❌ No disponible | ✅ `pdf_manipulate(op="flatten")`: sella AcroForms a contenido estático para archivo seguro |
| División Inteligente (Split) | ❌ Solo rangos fijos | ✅ `pdf_manipulate(op="split_smart")`: detecta límites heurísticos (páginas en blanco / carátulas) |
| Skills Empaquetadas & CLI | ❌ No disponible | ✅ `owi skill install [name]`: despliega manuales y flujos oficiales en `~/.hermes/skills/` |
| Vista Previa PNG / Visión | ❌ No disponible | ✅ `pdf_preview`: renderiza páginas a PNG (base64 data URL) para inspección visual previa |
| Batch pipelines | ❌ Múltiples turnos lentos | ✅ `office_batch`: ejecuta secuencias en un solo turno con tolerancia a fallas parciales |
| Edición in-place | ❌ Rehacer archivos enteros | ✅ `edit_excel` y `edit_word` preservando estilos originales y macros VBA |
| Reporte de fidelidad honesta | ❌ Silencio o afirmaciones falsas | ✅ `fidelity` (`rich` \| `clean` \| `lossy`) + `warnings` transparentes |
| Visión en PDFs | ❌ Solo texto plano | ✅ `read_pdf(extract_images=True)` devuelve imágenes base64 para agentes multimodales |
| PDF a Excel | ❌ No disponible | ✅ `pdf_to_excel` extrae tablas de PDFs directamente a hojas `.xlsx` estilizadas |
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

## Las 27 Tools (v0.8.0)

| Tool | Qué retorna | Cuándo usarla | Cuándo NO usarla |
|---|---|---|---|
| `document_diff` | JSON `{status, summary, diffs, warnings}` | Para comparar dos documentos Word (.docx) o PDF y obtener diferencias de párrafos con warnings honestos | No usar para redlines legales semánticos o diffs de píxeles |
| `scrub_metadata` | JSON `{status, path, bytes, scrubbed_fields}` | Para borrar autor, historial de revisiones, editores y propiedades personales en PDF, Word, Excel y PPTX | No usar para censurar texto visible en páginas (usar `pdf_redact`) |
| `protect_office` | JSON `{status, path, bytes, encrypted}` | Para cifrar y proteger con contraseña documentos Office (.docx, .xlsx, .pptx) con AES estándar | No usar para PDFs (usar `render_document` o `pdf_manipulate`) |
| `verify_pdf_signature` | JSON `{status, has_signature, valid, signer, date, warnings}` | Para verificar la integridad criptográfica y validez de firmas digitales en documentos PDF | No usar para firmar documentos (usar `sign_pdf`) |
| `mail_merge` | JSON `{status, template, n_docs, paths, fields}` | Para generar N documentos Word `.docx` personalizados desde plantilla `docxtpl` + CSV/JSON | No usar para documentos aislados de un solo paso (usar `create_word`) |
| `csv_excel_convert` | JSON `{status, path, bytes, n_rows, fidelity, warnings}` | Para convertir bidireccionalmente CSV a Excel estructurado y `.xlsx` a CSV | No usar para editar hojas de cálculo existentes (usar `edit_excel`) |
| `read_office` | JSON `{status, format, content?, paragraphs/slides/sheets}` | Para extraer texto y tablas de Word, PowerPoint o Excel a Markdown (pipes) o JSON | No usar para archivos PDF (usar `read_pdf` o `pdf_extract_structured`) |
| `create_pptx` | JSON `{status, path, bytes}` | Para crear presentaciones PowerPoint editables con tipografía, kickers y gráficos nativos | No usar para documentos de texto corrido o reportes en PDF |
| `pdf_redact` | JSON `{status, path, bytes, redactions_count}` | Para borrar permanentemente información confidencial (texto o coordenadas) | No usar para edición normal de texto (usar `edit_word`) |
| `pdf_extract_structured` | JSON `{status, format, content?, pages, n_tables}` | Para extraer texto y tablas en Markdown (pipes) o JSON estructurado para RAG | No usar en PDFs escaneados sin capa de texto (usar `pdf_ocr`) |
| `pdf_preview` | JSON `{status, data_url, pages, path?, bytes?}` | Para renderizar páginas PDF a imagen PNG y revisarlo visualmente antes de entregarlo | No usar para extraer texto seleccionable (usar `read_pdf`) |
| `read_pdf` | JSON `{pages, metadata, tables?, fields?, images?}` | Herramienta todo-en-uno: texto, metadatos, tablas (`extract_tables`), forms (`list_forms`), imágenes base64 (`extract_images`) | No usar para PDFs escaneados que requieran OCR completo (usar `pdf_ocr`) |
| `office_batch` | Resumen JSON `{status, total, succeeded, failed, results}` | Para ejecutar secuencias de operaciones documentales en un solo turno sin roundtrips | No usar para operaciones aisladas de un solo paso |
| `edit_excel` | JSON `{status, path, bytes, fidelity, warnings}` | Para modificar celdas, agregar filas/columnas, tablas o gráficos en `.xlsx`/`.xlsm` existentes | No usar para crear planillas desde cero (usar `create_excel`) |
| `edit_word` | JSON `{status, path, bytes, fidelity, warnings}` | Para editar `.docx` existentes (agregar párrafos, reemplazar texto, insertar tras títulos, tablas) | No usar para crear documentos nuevos desde cero (usar `create_word`) |
| `pdf_to_excel` | JSON `{status, path, bytes, n_tables, fidelity, warnings}` | Para extraer tablas desde PDFs (balances, facturas, informes) a planillas `.xlsx` limpias | No usar para PDFs escaneados sin tablas seleccionables (usar `pdf_ocr`) |
| `render_document` | JSON `{status, path, bytes}` | Para generar PDFs impecables desde plantillas HTML/Jinja + tema corporativo + logo + watermark | No usar para editar documentos existentes (usar `edit_word`/`pdf_manipulate`) |
| `create_word` | JSON `{status, path, bytes}` | Para crear documentos Word `.docx` nuevos desde bloques o plantillas empaquetadas | No usar para editar documentos en disco (usar `edit_word`) |
| `create_excel` | JSON `{status, path, bytes}` | Para crear libros Excel `.xlsx` multi-hoja con tablas estructuradas, autofiltro y gráficos | No usar para actualizar libros existentes (usar `edit_excel`) |
| `convert_to_pdf` | JSON `{status, path, bytes, fidelity, warnings}` | Para exportar archivos Office (`.docx`, `.xlsx`, `.pptx`) a PDF localmente con LibreOffice | No usar cuando se genera un documento nuevo desde plantilla HTML |
| `sign_pdf` | JSON `{status, path, bytes}` | Para estampar sello visual PNG y aplicar firma digital criptográfica PAdES con certificado PEM | No usar para editar contenido textual de un documento |
| `pdf_compress` | JSON `{status, path, bytes, savings_percent}` | Para optimizar y comprimir PDFs pesados reduciendo imágenes y limpiando objetos | No usar en PDFs de texto puro donde no hay imágenes |
| `pdf_manipulate` | JSON `{status, path/files, bytes, warnings?}` | Para unir (`merge`), extraer (`extract`/`split_by`), aplanar (`flatten`), rotar (`rotate`) o dividir inteligentemente (`split_smart`) | No usar para modificar texto dentro de las páginas |
| `pdf_fill_form` | JSON `{status, path, bytes}` | Para rellenar formularios interactivos AcroForm con campos clave-valor | No usar en PDFs planos o escaneados sin campos de formulario |
| `pdf_ocr` | JSON `{status, text, path?}` | Para extraer texto y generar PDF buscable sobre imágenes o escaneos con Tesseract | No usar sobre PDFs digitales que ya tengan texto seleccionable |
| `list_templates` | JSON `{status, templates}` | Para consultar el catálogo de plantillas oficiales `.docx` y las variables que requieren | No usar para generar archivos (usar `create_word`) |
| `list_themes` | JSON `{status, theme}` | Para consultar paleta de colores y fuentes de temas corporativos | No usar para generar documentos directamente |

---

## Capacidades Destacadas (v0.7.0)

### 1. Mail Merge Word en Lotes (`mail_merge`)
Genera $N$ documentos `.docx` independientes y personalizados combinando una plantilla `docxtpl` con placeholders `{{ variable }}` y un dataset tabular (CSV o JSON):
- Alimenta variables complejas desde archivos `.csv` o estructuras JSON estructuradas (`pandas` + `docxtpl`).
- Nomenclatura automática basada en prefijo (`<prefix>_1.docx`, `<prefix>_2.docx`).
- Filtrado opcional de columnas (`fields`) para control granular de contexto.

```json
{
  "template_path": "/workspace/plantilla_contrato.docx",
  "dataset_csv": "/workspace/nomina_empleados.csv",
  "output_prefix": "/workspace/contratos/contrato_firmado"
}
```

### 2. Extracción Office a Markdown / JSON para RAG (`read_office`)
Extrae directamente el contenido de archivos Word (`.docx`), PowerPoint (`.pptx`) y Excel (`.xlsx` / `.xlsm`) sin necesidad de convertirlos a PDF:
- **`format: "markdown"`**: Genera texto legible con jerarquía de encabezados (`# Titulo`, `## Subtitulo`) y tablas estructuradas en formato pipe Markdown (`| ... |`) listas para ingestión en bases de conocimiento RAG.
- **`format: "json"`** *(default, backward-compatible)*: Retorna objetos limpios con listas de párrafos, diapositivas y hojas de cálculo para inspección programática.

```json
{
  "path": "/workspace/balance_gestion.xlsx",
  "format": "markdown"
}
```

### 3. Gráficos Nativos en PowerPoint (`create_pptx`)
Soporta la inserción de gráficos nativos DrawingML (`bar`, `line`, `pie`) en diapositivas individuales mediante `python-pptx`:
- Gráficos 100% nativos: no son imágenes rasterizadas planas. Al abrir el `.pptx` en PowerPoint o Google Slides, los datos y series son completamente editables.
- Datos inline estructurados con categorías y valores numéricos o series multi-variable.

```json
{
  "out_path": "/workspace/reporte_ejecutivo.pptx",
  "slides_json": [
    {
      "title": "Evolución de Rendimiento 2026",
      "kicker": "Finanzas",
      "chart": {
        "type": "bar",
        "title": "Ingresos por Trimestre (kUSD)",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "values": [120, 160, 210, 290]
      }
    }
  ]
}
```

### 4. Conversión Bidireccional CSV <-> Excel (`csv_excel_convert`)
Convierte archivos CSV a libros Excel estructurados (`.xlsx`) y viceversa:
- **CSV -> Excel**: Genera tablas oficiales (`TableStyleMedium9`), autofiltro, ajuste de columnas y preservación de códigos numéricos con ceros a la izquierda (ej: códigos postales, identificadores).
- **Excel -> CSV**: Exporta la hoja activa o todas las hojas del libro (`sheet="all"`) a archivos CSV limpios, emitiendo advertencias honestas (`warnings`) ante cualquier ambigüedad de tipo.

```json
{
  "input": "/workspace/datos_clientes.csv",
  "output": "/workspace/clientes_estructurado.xlsx",
  "direction": "csv_to_xlsx",
  "sheet": "Clientes"
}
```

### 5. Redacción Irreversible y Permanente (`pdf_redact`)
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

### 6. Aplanado de Formularios para Cumplimiento de Archivo (`flatten`)
Convierte campos de formularios interactivos AcroForm y anotaciones en contenido gráfico y textual estático mediante PyMuPDF (`doc.bake(annots=True, widgets=True)`). Garantiza que contratos y formularios completados no puedan ser manipulados ni alterados posteriormente:

```json
{
  "operation": "flatten",
  "input_path": "/workspace/formulario_relleno.pdf",
  "output": "/workspace/formulario_archivado.pdf"
}
```

### 7. Extracción Estructurada Markdown / JSON para RAG en PDF (`pdf_extract_structured`)
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

### 8. División Inteligente de Documentos (`split_smart`)
Detecta heurísticamente los límites entre documentos concatenados (por páginas separadoras en blanco sin texto ni gráficos, o por carátulas/encabezados repetidos) y los divide automáticamente en $N$ archivos independientes válidos:

```json
{
  "operation": "split_smart",
  "input_path": "/workspace/lote_escaneado.pdf",
  "output": "/workspace/documento_part.pdf"
}
```
*Honestidad técnica:* Si no se detecta un límite confiable, la herramienta emite advertencias honestas (`warnings`) y mantiene el documento íntegro en lugar de inventar particiones arbitrarias.

### 9. Skills Empaquetadas & Comando CLI (`owi skill install`)
The Office Worker empaqueta manuales de operación (`SKILL.md`) con ejemplos conversacionales en inglés, reglas duras anti-loop y flujos opcionales de Google Drive / Gmail (`workspace-mcp`).
Instala la skill en el entorno de tu agente con un solo comando idempotente:

```bash
owi skill list                       # Lista las skills oficiales disponibles
owi skill install office-worker      # Copia la skill a ~/.hermes/skills/office-worker/SKILL.md
owi skill install all                # Instala todas las skills empaquetadas
```

### 10. Preview PNG para Visión de Agentes (`pdf_preview`)
Renderiza páginas PDF a imágenes PNG de alta resolución (`dpi=110`) mediante PyMuPDF (`page.get_pixmap()`), retornando Data URLs en base64 (`data:image/png;base64,...`) para inspección visual previa a la entrega final.

### 11. Batch Operations (`office_batch`)
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

### 12. Edición In-Place (`edit_excel` y `edit_word`)
Modifica documentos existentes en el mismo archivo o en uno nuevo sin perder estilos previos:
- **`edit_excel`**: Modifica celdas (`set_cell`), agrega filas (`append_row`), agrega columnas (`add_column`), agrega tablas estilizadas (`add_table`), activa filtros (`auto_filter`) e inserta gráficos (`add_chart`).
- **`edit_word`**: Agrega párrafos (`append_paragraph`), reemplaza texto preservando formato de run (`replace_text`), inserta contenido tras encabezados (`insert_after_heading`) o añade tablas (`append_table`).

### 13. Reporte de Fidelidad Honesta
`convert_to_pdf`, `edit_excel`, `edit_word` y `pdf_to_excel` devuelven:
- `fidelity`: `"rich"` (edición sin pérdida o nativa), `"clean"` (conversión estándar de alta fidelidad), o `"lossy"` (degradación esperable por limitaciones de formato).
- `warnings`: Lista explícita de advertencias (sustitución tipográfica, saltos de página automáticos en planillas anchas, macros no ejecutadas).

### 14. Comparación y Diff Honesto de Documentos (`document_diff`)
Compara dos documentos Word (`.docx`) o PDF y calcula las diferencias textuales reales entre versiones mediante `difflib`:
- Reporta párrafos agregados (`added`), eliminados (`deleted`), modificados (`modified`) e intactos (`unchanged`).
- Formatos disponibles: JSON estructurado o Markdown formateado (`format="markdown"`).
- **Advertencia honesta:** La herramienta emite explícitamente un warning indicando que se trata de una comparación textual aproximada y no de un redline legal-grade semántico.

### 15. Limpieza de Metadatos Sensibles (`scrub_metadata`)
Elimina metadatos de auditoría y rastreo antes de compartir documentos externamente:
- **PDF**: Borra diccionario de metadatos vía PyMuPDF (`set_metadata({})`) y ejecuta compactación profunda (`garbage=4, clean=True`).
- **Word (.docx)**: Limpia `core_properties` (autor, último editor, título, asunto, comentarios, categoría, resetea número de revisión).
- **Excel (.xlsx)**: Limpia propiedades de libro (creador, último modificador, título, descripción).
- **PowerPoint (.pptx)**: Limpia propiedades de presentación (autor, editores, asunto).
- Soporta filtrado por campos específicos (`fields=['author', 'title']`) o limpieza total por defecto.

### 16. Cifrado y Protección con Contraseña en Office (`protect_office`)
Aplica cifrado ágil estándar por contraseña (ECMA-376 Agile Encryption con AES-128/256) sobre archivos Word (`.docx`), Excel (`.xlsx`) y PowerPoint (`.pptx`) mediante `msoffcrypto`:
- Al abrir el archivo sin la contraseña correcta, aplicaciones como Microsoft Office, LibreOffice u openpyxl rechazan la lectura (`BadZipFile`).
- Abre y descifra de forma transparente al proveer la clave correcta.

### 17. Verificación de Firmas Digitales PDF (`verify_pdf_signature`)
Audita y valida documentos PDF firmados digitalmente mediante `pyhanko` y `pypdf`:
- Inspecciona si el documento contiene firmas digitales embebidas (`has_signature: true/false`).
- Valida la integridad criptográfica del digest (`intact: true`) comprobando que el archivo no haya sido alterado tras la firma.
- Extrae metadatos del firmante (Common Name del certificado X.509), fecha de firma (`D:YYYYMMDD...`), motivo y ubicación.
- Emite advertencias transparentes cuando el certificado es auto-firmado o no está anclado en la cadena de confianza del sistema.

### 18. Guiado Proactivo (`next_steps` y Skill de Orquestación)
- Las herramientas principales de generación y transformación (`create_word`, `create_excel`, `create_pptx`, `render_document`, `convert_to_pdf`, `sign_pdf`, `pdf_redact`, `mail_merge`) retornan un campo opcional `"next_steps"` con la siguiente acción lógica sugerida en inglés (ej. tras `create_word` &rarr; `["Convert to PDF with convert_to_pdf", "Preview rendered document with pdf_preview"]`).
- Incluye la skill empaquetada `orchestration.SKILL.md` instalable con `owi skill install orchestration`, mapeando intenciones de negocio (factura, informe, acta, contrato, auditoria, batch, redact, sign, rag, compliance) a cadenas de herramientas deterministas.

---

## Flujos Conversacionales de Ejemplo (EN)

### Flujo 1: Bulk Mail Merge (Generating Employee Welcome Letters)
*User:* "I have a CSV with 50 new hires and a Word template. Generate personalized welcome letters for all of them."
*Agent:*
```json
{
  "template_path": "templates/welcome_letter.docx",
  "dataset_csv": "data/new_hires.csv",
  "output_prefix": "output/letters/welcome"
}
```
*Result:* Returns `{"status": "ok", "n_docs": 50, "paths": [".../welcome_1.docx", ...]}`.

### Flujo 2: Ingesting Office Spreadsheets into RAG Context
*User:* "Read our quarterly financial balance sheet and give me the markdown table so I can analyze it."
*Agent:*
```json
{
  "path": "reports/q3_balance.xlsx",
  "format": "markdown"
}
```
*Result:* Returns clean Markdown pipe tables per sheet with exact numeric values.

### Flujo 3: Executive Deck with Native PowerPoint Charts
*User:* "Create an executive presentation with 2 slides: title slide and a quarterly sales bar chart."
*Agent:*
```json
{
  "out_path": "presentations/sales_q3.pptx",
  "slides_json": [
    {
      "title": "Quarterly Business Review",
      "kicker": "Executive Summary",
      "bullets": ["Q3 Key Performance Indicators", "Next Steps"]
    },
    {
      "title": "Revenue Growth 2026",
      "kicker": "Performance",
      "chart": {
        "type": "bar",
        "title": "Quarterly Revenue ($M)",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "values": [12.4, 14.2, 16.8, 19.5]
      }
    }
  ]
}
```

### Flujo 4: Comparing Contract Revisions (`document_diff`)
*User:* "Compare agreement_draft_v1.docx and agreement_draft_v2.docx and tell me what changed."
*Agent:*
```json
{
  "path_a": "agreement_draft_v1.docx",
  "path_b": "agreement_draft_v2.docx",
  "format": "markdown"
}
```
*Result:* Returns markdown summary with modified/added clauses and transparent warnings about textual difflib boundaries.

### Flujo 5: Scrubbing Sensitive Metadata before External Dispatch (`scrub_metadata`)
*User:* "Scrub all author, editor, and revision metadata from executive_report.docx before we email it to partners."
*Agent:*
```json
{
  "input": "executive_report.docx",
  "output": "executive_report_clean.docx"
}
```
*Result:* Returns `{"status": "ok", "scrubbed_fields": ["author", "comments", "last_modified_by", "revision", "title"]}`.

### Flujo 6: Password-Protecting Spreadsheets (`protect_office`)
*User:* "Lock compensation_matrix.xlsx with password 'HR-Enterprise-2026!'."
*Agent:*
```json
{
  "input": "compensation_matrix.xlsx",
  "output": "compensation_matrix_encrypted.xlsx",
  "password": "HR-Enterprise-2026!"
}
```
*Result:* Returns `{"status": "ok", "encrypted": true}` producing standard agile AES encrypted spreadsheet.

### Flujo 7: Verifying Cryptographic PDF Digital Signatures (`verify_pdf_signature`)
*User:* "Verify if this signed vendor contract vendor_contract_signed.pdf has an intact digital signature."
*Agent:*
```json
{
  "input": "vendor_contract_signed.pdf"
}
```
*Result:* Returns `{"has_signature": true, "valid": true, "intact": true, "signer": "Common Name: Julio Cardozo", "reason": "Approved"}`.

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
6. **Cifrado Office (`protect_office`):**
   - Aplica cifrado agile estándar ECMA-376 con AES mediante `msoffcrypto`. Protege la apertura del archivo en Word, Excel y PowerPoint. La protección con clave de archivos PDF se realiza mediante `render_document(password=...)` o `pdf_manipulate(operation="encrypt", password=...)`.
7. **Comparación de Documentos (`document_diff`):**
   - La comparación se realiza a nivel textual mediante `difflib` sobre el texto y párrafos extraídos de Word (.docx) y PDF. No constituye un redline semántico ni legal-grade con seguimiento de marcas Word OOXML track changes nativo.

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

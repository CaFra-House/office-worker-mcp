---
name: office-worker
description: "Use when the agent must create, edit, convert, manipulate, sign, compress or read office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX). The Office Worker MCP: 20 tools for batch pipelines (office_batch), in-place edits (edit_excel, edit_word), packaged templates (docxtpl), native Excel charts, structured tables, vision image extraction (read_pdf extract_images), PDF-to-Excel (pdf_to_excel), Office reader (read_office), AcroForm filling, OCR, compression, signatures, and conversion. Deterministic one-path-per-task with hard anti-loop limits."
---

# The Office Worker — Guía de uso para agentes (v0.5.0)

Genera y modifica documentos de oficina **profesionales** (diseño consistente, no genéricos) y procesa PDFs usando el
MCP `office-worker`. 100% local, sin API keys, determinista y seguro. La clave: **una tool por formato, un camino determinista**.

## Reglas duras (anti-colapso / anti-loop — probadas en producción)

1. **Un camino por tarea.** Elegí UNA tool según el objetivo pedido y usala. No improvises alternativas ni pruebes métodos incompatibles.
2. **Falla 2 veces → PARÁ OBLIGATORIAMENTE (Hard Anti-Loop).** Si una tool devuelve `{status:"error"}` dos veces seguidas en la misma tarea, **QUEDA ESTRICTAMENTE PROHIBIDO** hacer un 3er intento o probar variantes a ciegas. Reportá el error exacto recibido (`error`) al usuario y detenete pidiendo clarificación. El 3er intento ciego provoca alucinación, degradación de contexto y gasto inútil de tokens.
3. **Flujos multi-paso → usá `office_batch`.** Si tenés que generar, convertir y firmar un documento, o procesar varios archivos juntos, despachalos en una sola llamada a `office_batch`. Si una operación intermedia falla, el batch continúa ejecutando las demás y te reporta el detalle de éxitos y fallos.
4. **Modificar archivo existente → `edit_excel` o `edit_word`.** Si el usuario pide cambiar una celda, agregar una fila, insertar un gráfico en un `.xlsx`, o cambiar texto en un `.docx`, **NUNCA** regeneres el archivo entero desde cero: usá `edit_excel` o `edit_word`. Preservan formato, estilos, fuentes y macros VBA (`keep_vba`).
5. **Plantillas empaquetadas primero:** Si el usuario pide un acta, informe ejecutivo, factura, carta formal o checklist de auditoría, llamá a `list_templates()` para conocer las variables y usá `create_word(out_path=..., template_docx="<nombre>", context={...})`. No inventes estructuras desde cero cuando ya existe una plantilla oficial probada.
6. **Tablas en PDF → `pdf_to_excel`:** Si un PDF contiene balances, listas de precios o tablas financieras y querés exportarlas a planilla, usá `pdf_to_excel`. No copies texto a mano.
7. **Lectura directa de Office → `read_office`:** Si necesitás extraer texto, párrafos, diapositivas o celdas de un `.docx`, `.pptx` o `.xlsx`, llamá a `read_office(path)`. No lo conviertas a PDF para luego leerlo.
8. **Visión / logos / sellos en PDF → `read_pdf(extract_images=True)`:** Cuando tu modelo de visión necesite inspeccionar imágenes embebidas, sellos o diagramas de un PDF, activá `extract_images=True` (con `max_images=5` o `10`).
9. **Archivos grandes → nunca heredoc.** Para plantillas o JSON >8KB, escribilos a disco con `write_file` y pasá la ruta / contenido; jamás armarlos en un heredoc de terminal.
10. **Verificá en disco.** Después de generar o transformar un archivo, confirmá que el archivo existe y tiene tamaño razonable (el tool devuelve `bytes`; si es ~0 algo falló aunque diga ok).

---

## Cuál tool usar

| Objetivo | Tool | Args clave |
|---|---|---|
| Ejecutar pipeline de múltiples documentos en 1 turno | `office_batch` | `operations: [{tool, args}]` |
| Editar planilla Excel existente in-place (.xlsx/.xlsm) | `edit_excel` | `input_path`, `operations` ([set_cell, append_row, add_column, add_chart, add_table, auto_filter]), opcional `output_path` |
| Editar documento Word existente in-place (.docx) | `edit_word` | `input_path`, `operations` ([append_paragraph, replace_text, insert_after_heading, append_table]), opcional `output_path` |
| Extraer tablas de un PDF a planilla Excel (.xlsx) | `pdf_to_excel` | `input_path`, `output_path`, opcional `sheet_name`, `max_pages` |
| Extraer texto y estructura de Office (.docx/.pptx/.xlsx) | `read_office` | `path`, opcional `max_rows=500` |
| Leer texto, tablas, forms e imágenes de un PDF | `read_pdf` | `path`, opcional `max_pages`, `extract_tables=True`, `list_forms=True`, `extract_images=True`, `max_images=10` |
| Consultar catálogo de plantillas Word oficiales | `list_templates` | *(Sin argumentos)* Devuelve catálogo con variables esperadas |
| Documento Word con plantilla oficial empaquetada | `create_word` | `out_path`, `template_docx="acta_meeting"\|"informe_ejecutivo"\|"factura_simple"\|"carta_formal"\|"checklist_auditoria"`, `context` |
| Documento Word nuevo desde bloques | `create_word` | `out_path`, `title`, `blocks_json`, `theme` |
| Planilla Excel nueva con tablas, autofiltro y gráficos | `create_excel` | `out_path`, `title`, `sheets_json`, `theme`, opcional `table_style`, `auto_filter` |
| Presentación PowerPoint editable nueva | `create_pptx` | `out_path`, `slides_json`, `theme` |
| Informe / factura / cartilla en PDF bonito | `render_document` | `template_html` (Jinja), `out_path`, `data_json`, `theme`, opcional `watermark_text`, `footer_left`, `footer_right`, `logo`, `password` |
| Comprimir / optimizar tamaño de PDF | `pdf_compress` | `input_path`, `output`, `quality` ("low"\|"med"\|"high") |
| Firmar PDF (sello PNG visual o PAdES digital) | `sign_pdf` | `input_pdf`, `output`, opcional `sello_img_path`, `cert_pem`, `reason`, `location` |
| Convertir Office (.docx/.xlsx/.pptx) a PDF | `convert_to_pdf` | `input_file`, `output` |
| Unir, extraer páginas o rotar PDF | `pdf_manipulate` | `operation` ("merge"\|"extract"\|"rotate"), `output`, `files`/`input_path`, `pages`, `angle`, opcional `password` |
| Rellenar formulario PDF (AcroForm) | `pdf_fill_form` | `input_pdf`, `output`, `fields` (dict o JSON) |
| OCR sobre imagen o PDF escaneado | `pdf_ocr` | `input_path`, opcional `lang`, `output` (PDF buscable), `max_pages` |
| Consultar paleta del tema activo | `list_themes` | opcional `theme_name_or_path` |
| *(Deprecated)* Extraer tablas | `pdf_extract_tables` | Preferir `read_pdf(path, extract_tables=True)` |
| *(Deprecated)* Listar formularios | `pdf_list_form_fields` | Preferir `read_pdf(path, list_forms=True)` |

---

## Guía Operativa de Nuevas Tools

### `office_batch` (Pipelines multi-paso)
Usa `office_batch` para evitar idas y vueltas con el usuario cuando una tarea involucra varias etapas:
```json
{
  "operations": [
    { "tool": "create_word", "args": { "out_path": "report.docx", "template_docx": "informe_ejecutivo", "context": { "titulo": "Reporte Semanal" } } },
    { "tool": "convert_to_pdf", "args": { "input_file": "report.docx", "output": "report.pdf" } },
    { "tool": "sign_pdf", "args": { "input_pdf": "report.pdf", "output": "report_firmado.pdf", "sello_img_path": "sello.png" } }
  ]
}
```
Si un paso falla (ej: falta el sello), el despachador reporta `"status": "partial_error"` indicando exactamente qué paso falló y conservando los archivos exitosos.

### `edit_excel` (Edición in-place de planillas)
Modifica hojas existentes sin rehacer el libro:
```json
{
  "input_path": "finanzas.xlsx",
  "operations": [
    { "op": "set_cell", "coordinate": "C5", "value": 15000 },
    { "op": "append_row", "sheet": "Gastos", "row": ["Servicios Cloud", 450, "=B5*1.21"] },
    { "op": "add_chart", "chart_type": "bar", "title": "Gastos por Rubro", "target_cell": "F2" },
    { "op": "add_table", "table_style": "TableStyleMedium9" },
    { "op": "auto_filter" }
  ]
}
```

### `edit_word` (Edición in-place de documentos)
```json
{
  "input_path": "contrato.docx",
  "operations": [
    { "op": "replace_text", "find": "PROVEEDOR_PENDIENTE", "replace": "Acme Corp S.A." },
    { "op": "insert_after_heading", "heading_text": "Cláusula Tercera", "text": "El plazo de pago acordado será de 30 días hábiles." },
    { "op": "append_paragraph", "text": "Firmado en conformidad.", "bold": true }
  ]
}
```

### `read_pdf` con `extract_images=True`
Extrae las imágenes del PDF como data URLs en Base64 para que las analices con tu visión multimodal:
```json
{ "path": "documento.pdf", "extract_images": true, "max_images": 5 }
```

### `pdf_to_excel`
Extrae tablas de un PDF directamente a un `.xlsx` estilizado:
```json
{ "input_path": "balance_anual.pdf", "output_path": "balance_tablas.xlsx", "sheet_name": "Balance" }
```

### `read_office`
Extrae texto y estructura de cualquier archivo Office sin convertirlo a PDF:
```json
{ "path": "presentacion.pptx" }
```

---

## Fidelidad Honesta

Tanto `convert_to_pdf` como `edit_excel`, `edit_word` y `pdf_to_excel` retornan el campo `fidelity`:
- `"rich"`: Modificación nativa completa sin pérdida.
- `"clean"`: Conversión estándar de alta fidelidad (ej: LibreOffice headless). Incluye advertencias de posibles variantes tipográficas o saltos de página.
- `"lossy"`: Archivo o extracción degradada (ej: PDF sin tablas estructuradas).
Si hay degradación, examiná `warnings` antes de asumir que el documento es perfecto.

## Errores típicos y cómo resolverlos

- `"data_json inválido"`, `"blocks_json inválido"` o `"context inválido"` → JSON malformado. Revisá la sintaxis UNA vez y reintentá. Máx 2 intentos.
- `"Acceso denegado: ruta de salida en directorio de sistema protegido"` → Intentaste escribir en `/etc` u otra ruta de sistema protegida. Escribí en una ruta de trabajo válida o bajo tu directorio de trabajo.
- `"El archivo PDF no contiene campos de formulario interactivos"` → el PDF es plano o escaneado. Usá `pdf_ocr` o `read_pdf`, no `pdf_fill_form`.
- `"LibreOffice no está instalado"` → asegurate de que `soffice` esté en el PATH del sistema.
- PPTX lento (~15–30s) → normal: renderiza con Chromium headless. Esperá el timeout.



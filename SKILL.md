---
name: office-worker
description: "Use when the agent must create, convert, manipulate or read office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX). The Office Worker MCP: 12 tools to render professional docs from data + corporate theme, fill AcroForms, run OCR, convert Office files, manipulate PDFs, and read text/tables. Deterministic one-path-per-task with hard anti-loop limits."
---

# The Office Worker — guía de uso para agentes

Genera documentos de oficina **profesionales** (diseño consistente, no genéricos) y procesa PDFs usando el
MCP `office-worker`. 100% local. La clave: **una tool por formato, un camino determinista**.

## Reglas duras (anti-colapso — probadas en producción)

1. **Un camino por tarea.** Elegí UNA tool según el objetivo pedido y usala. No improvises alternativas ni pruebes métodos incompatibles.
2. **Falla 2 veces → PARÁ.** Si una tool devuelve `{status:"error"}` dos veces seguidas en la misma tarea, NO reintentes con variantes. Reportá el error exacto (`error`) y detenete. El 3er intento ciego provoca degradación y bucles de contexto.
3. **Archivos grandes → nunca heredoc.** Para plantillas o JSON >8KB, escribilos a disco con `write_file` y pasá la ruta / contenido; jamás armarlos en un heredoc de terminal.
4. **Plan-first para decks.** Antes de `create_pptx`, definí las slides (título + bullets) como lista clara. Máx 2 reintentos por deck (`MAX_SLIDE_RETRIES=2`).
5. **Verificá en disco.** Después de generar o transformar un archivo, confirmá que el archivo existe y tiene tamaño razonable (el tool devuelve `bytes`; si es ~0 algo falló aunque diga ok).

## Cuál tool usar

| Pide el usuario | Tool | Args clave |
|---|---|---|
| Informe/factura/cartilla en PDF bonito | `render_document` | `template_html` (Jinja), `out_path`, `data_json`, `theme`, opcional `logo` |
| Documento Word editable | `create_word` | `out_path`, `title`, `blocks_json` (h1/h2/p/table), `theme` |
| Planilla / tabla de datos | `create_excel` | `out_path`, `title`, `sheets_json` [{name,headers,rows}], `theme` |
| Presentación / deck editable | `create_pptx` | `out_path`, `slides_json` [{title,kicker,bullets}], `theme` |
| Leer texto de un PDF | `read_pdf` | `path`, opcional `max_pages` |
| Extraer tablas estructuradas de un PDF | `pdf_extract_tables` | `path`, opcional `max_pages` |
| Consultar campos de formulario de un PDF | `pdf_list_form_fields` | `path` |
| Rellenar formulario PDF (AcroForm) | `pdf_fill_form` | `input_pdf`, `output`, `fields` (dict o JSON) |
| OCR sobre imagen o PDF escaneado | `pdf_ocr` | `input_path`, opcional `lang` ("spa", "eng"), opcional `output` (PDF buscable) |
| Convertir Office (.docx/.xlsx/.pptx) a PDF | `convert_to_pdf` | `input_file`, `output` |
| Unir, extraer páginas o rotar PDF | `pdf_manipulate` | `operation` ("merge"\|"extract"\|"rotate"), `output`, `files`/`input_path`, `pages`, `angle` |
| Consultar paleta del tema activo | `list_themes` | opcional `theme_name_or_path` |

## Temas corporativos & Logo

Todas las tools de generación aceptan `theme` (`"aden"` por defecto, `"claro"`, `"oscuro"`, `"minimal"`, `"corporate-blue"`, o ruta a YAML/JSON).
Para consistencia visual en una misma organización, usá siempre el mismo tema.
Además, `render_document` acepta el parámetro `logo` (ruta absoluta o relativa a archivo PNG/JPG), que WeasyPrint sitúa en la esquina superior derecha del encabezado.

## Plantilla HTML para PDF (ejemplo mínimo)

```html
<h1>{{ titulo }}</h1>
<p class="muted">{{ subtitulo }} · {{ fecha }}</p>
{% if tabla is defined and tabla %}{{ tabla }}{% endif %}
```

Para tablas declarativas, pasá en `data_json`: `{"headers":[...], "rows":[[...]]}` — el server arma `<table>` automáticamente y lo inyecta en `{{ tabla }}`.

## Errores típicos y cómo resolverlos

- `"data_json inválido"` o `"fields inválido"` → JSON malformado. Revisá la sintaxis UNA vez y reintentá. Máx 2 intentos.
- `"El archivo PDF no contiene campos de formulario interactivos"` → el PDF es plano o escaneado. Usá `pdf_ocr` o `read_pdf`, no `pdf_fill_form`.
- `"LibreOffice no está instalado"` → asegurate de que `soffice` esté en el PATH del sistema.
- PPTX lento (~15–30s) → normal: renderiza con Chromium headless. Esperá el timeout.

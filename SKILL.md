---
name: office-worker
description: "Use when the agent must create or read office documents (PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX). The Office Worker MCP: 8 tools to render professional docs from data + corporate theme and to read PDFs (text/tables/forms). Deterministic one-path-per-task with hard anti-loop limits."
---

# The Office Worker — guía de uso para agentes

Genera documentos de oficina **profesionales** (diseño consistente, no genéricos) usando el
MCP `office-worker`. 100% local. La clave: **una tool por formato, un camino determinista**.

## Reglas duras (anti-colapso — probadas en producción)

1. **Un camino por tarea.** Elegí UNA tool según el formato pedido y usala. No improvises alternativas ni pruebes 3 métodos distintos.
2. **Falla 2 veces → PARÁ.** Si una tool devuelve `{status:"error"}` dos veces seguidas, NO reintentes con variantes. Reportá el error exacto (`error`) y detente. El 3er intento ciego es lo que colapsó perfiles antes.
3. **Archivos grandes → nunca heredoc.** Para plantillas o JSON >8KB, escribilos a disco con `write_file` y pasá la ruta / contenido; jamás armarlos en un heredoc de terminal (causa degradación del modelo en contexto largo).
4. **Plan-first para decks.** Antes de `create_pptx`, definí las slides (título + bullets) como lista clara. Máx 2 reintentos por deck (`MAX_SLIDE_RETRIES=2`).
5. **Verificá en disco.** Después de generar, confirmá que el archivo existe y tiene tamaño razonable (el tool ya devuelve `bytes`; si es ~0 o <500B algo falló aunque diga ok).

## Cuál tool usar

| Pide el usuario | Tool | Args clave |
|---|---|---|
| Informe/factura/cartilla en PDF bonito | `render_document` | `template_html` (Jinja), `data_json`, `theme` |
| Documento Word editable | `create_word` | `title`, `blocks_json` (h1/h2/p/table) |
| Planilla / tabla de datos | `create_excel` | `sheets_json` [{name,headers,rows}] |
| Presentación / deck | `create_pptx` | `slides_json` [{title,kicker,bullets}] |
| Leer texto de un PDF | `read_pdf` | `path`, opcional `max_pages` |
| Extraer tablas de un PDF | `pdf_extract_tables` | `path` |
| ¿El PDF tiene formulario? | `pdf_list_form_fields` | `path` |
| ¿Qué paleta usa mi tema? | `list_themes` | opcional nombre/ruta tema |

## Tema corporativo

Todas las tools aceptan `theme` (nombre `"aden"` por defecto, o ruta a YAML/JSON con la paleta).
Siempre pasá el mismo tema en un documento multi-página para consistencia visual. Paleta ADEN default:
primario `#003366`, acento `#3B82F6`, texto `#1A202C`. Usalo para headers/tablas; no inventes colores sueltos.

## Plantilla HTML para PDF (ejemplo mínimo)

```html
<h1>{{ titulo }}</h1>
<p class="muted">{{ subtitulo }} · {{ fecha }}</p>
{% if tabla is defined and tabla %}{{ tabla }}{% endif %}
```

Para tablas declarativas, pasá en `data_json`: `{"headers":[...], "rows":[[...]]}` — el server arma `<table>` solo e inyecta en `{{ tabla }}`.

## Errores típicos y cómo leerlos

- `"asyncio.run() cannot be called..."` → bug interno ya corregido; si aparece, reportalo (no reintentar).
- `"data_json inválido"` → tu JSON malformado; corregí la sintaxis UNA vez y volvé a intentar. Máx 2 intentos totales.
- PPTX lento (~15–30s) → normal: renderiza con Playwright Chromium headless. No es un fallo; esperá el timeout.

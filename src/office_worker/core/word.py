"""Generación de Word (.docx) con python-docx, aplicando tema corporativo."""
from __future__ import annotations
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .security import safe_out

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def create_word(out_path, title="", subtitle=None, blocks=None, theme=None, template_docx=None, context=None):
    """Crea un .docx profesional desde bloques declarativos o rellenando una plantilla con docxtpl.

    - blocks: lista de dicts. Tipos soportados:
      {"type":"h1"|"h2"|"p", "text": "..."}
      {"type":"table", "headers":[...], "rows":[[...], ...]}
    - template_docx: ruta opcional a plantilla .docx existente con placeholders Jinja {{ variable }}.
    - context: diccionario opcional de variables para rellenar template_docx.
    Si blocks es None y no hay template_docx → documento mínimo con título. Devuelve ruta absoluta.
    """
    out_path = safe_out(out_path)

    if template_docx:
        from .templates_pack import resolve_template_path
        tpl_path = resolve_template_path(template_docx)
        from docxtpl import DocxTemplate
        tpl = DocxTemplate(tpl_path)
        ctx = dict(context or {})
        if title and "title" not in ctx: ctx["title"] = title
        if subtitle and "subtitle" not in ctx: ctx["subtitle"] = subtitle
        tpl.render(ctx)
        tpl.save(out_path)
        if os.path.getsize(out_path) < 300:
            raise RuntimeError(f"DOCX generado sospechosamente pequeño ({os.path.getsize(out_path)}B)")
        return out_path

    from .themes import load_theme
    th = load_theme(theme)
    primary = _hex_to_rgb(th["primary"])

    doc = Document()
    # estilo base de fuente/cuerpo según tema
    st = doc.styles["Normal"]; st.font.name = th.get("font_body","Arial"); st.font.size = Pt(10.5)

    if title:
        h = doc.add_heading(title, level=0); 
        for r in h.runs: r.font.color.rgb = primary
    if subtitle:
        p = doc.add_paragraph(subtitle)
        if p.runs: p.runs[0].font.color.rgb = _hex_to_rgb(th["muted"])

    for b in (blocks or []):
        t = b.get("type","p")
        if t == "h1":
            hh = doc.add_heading(b["text"], level=1)
            for r in hh.runs: r.font.color.rgb = primary
        elif t == "h2":
            hh = doc.add_heading(b["text"], level=2)
            for r in hh.runs: r.font.color.rgb = primary
        elif t == "table":
            headers = b.get("headers",[]); rows = b.get("rows",[])
            tbl = doc.add_table(rows=1, cols=len(headers)); tbl.style="Light Grid Accent 1"
            hdr = tbl.rows[0].cells
            for i,hv in enumerate(headers): hdr[i].text=str(hv)
            for row in rows:
                cells = tbl.add_row().cells
                for i,cv in enumerate(row):
                    if i < len(cells): cells[i].text=str(cv)
        else:  # p / default
            doc.add_paragraph(b.get("text",""))

    doc.save(out_path)
    if os.path.getsize(out_path) < 300: raise RuntimeError(f"DOCX sospechosamente pequeño ({os.path.getsize(out_path)}B)")
    return out_path


def edit_word(
    input_path: str,
    operations: list[dict[str, Any]] | str,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Modifica un documento Word existente (.docx) preservando estilos y formato.

    Operaciones soportadas (lista de dicts):
    - {"op": "append_paragraph", "text": str, "style"?: str, "bold"?: bool, "italic"?: bool}
    - {"op": "replace_text", "find"|"old": str, "replace"|"new": str, "count"?: int}
    - {"op": "insert_after_heading", "heading_text": str, "text": str, "style"?: str, "bold"?: bool, "italic"?: bool}
    - {"op": "append_table", "headers": list, "rows": list[list], "style"?: str}

    Devuelve dict con fidelity honesta ("clean"), warnings y ruta absoluta.
    """
    import json

    input_path = os.path.abspath(os.path.expanduser(str(input_path)))
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Archivo Word no encontrado: {input_path}")

    target_out = safe_out(output_path if output_path else input_path)
    doc = Document(input_path)

    warnings = ["python-docx edits document body; embedded OLE objects, drawing canvas and SmartArt are not updated."]
    fidelity = "clean"

    if isinstance(operations, str):
        try:
            ops = json.loads(operations)
        except json.JSONDecodeError as exc:
            raise ValueError(f"operations inválido (JSON malformado): {exc}")
    else:
        ops = list(operations or [])

    for op in ops:
        op_name = str(op.get("op") or op.get("operation") or "").lower().strip()

        if op_name == "append_paragraph":
            text = op.get("text", "")
            style = op.get("style")
            p = doc.add_paragraph(text, style=style)
            bold = op.get("bold")
            italic = op.get("italic")
            if bold is not None or italic is not None:
                for r in p.runs:
                    if bold is not None:
                        r.bold = bool(bold)
                    if italic is not None:
                        r.italic = bool(italic)

        elif op_name == "replace_text":
            old = op.get("find") or op.get("old") or ""
            new = op.get("replace") or op.get("new") or ""
            count = op.get("count")
            max_c = int(count) if count is not None else -1
            c_left = max_c

            if not old:
                continue

            # Reemplazar en párrafos
            for p in doc.paragraphs:
                if old in p.text:
                    replaced_in_run = False
                    for r in p.runs:
                        if old in r.text:
                            if c_left > 0:
                                r.text = r.text.replace(old, new, c_left)
                                c_left -= 1
                            elif c_left == -1:
                                r.text = r.text.replace(old, new)
                            replaced_in_run = True
                            if c_left == 0:
                                break
                    if not replaced_in_run:
                        if c_left > 0:
                            p.text = p.text.replace(old, new, c_left)
                            c_left -= 1
                        elif c_left == -1:
                            p.text = p.text.replace(old, new)
                    if c_left == 0:
                        break

            # Reemplazar en tablas
            if c_left != 0:
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if old in p.text:
                                    if c_left > 0:
                                        p.text = p.text.replace(old, new, c_left)
                                        c_left -= 1
                                    elif c_left == -1:
                                        p.text = p.text.replace(old, new)
                                    if c_left == 0:
                                        break

        elif op_name == "insert_after_heading":
            target_heading = str(op.get("heading_text", "")).strip().lower()
            text = op.get("text", "")
            style = op.get("style", "Normal")
            matched_p = None
            for p in doc.paragraphs:
                if target_heading in p.text.strip().lower():
                    matched_p = p
                    break

            if matched_p is not None:
                new_p = doc.add_paragraph(text, style=style)
                bold = op.get("bold")
                italic = op.get("italic")
                if bold is not None or italic is not None:
                    for r in new_p.runs:
                        if bold is not None:
                            r.bold = bool(bold)
                        if italic is not None:
                            r.italic = bool(italic)
                # Mover XML para insertar inmediatamente después del encabezado
                matched_p._p.addnext(new_p._p)
            else:
                # Si no encontró el heading, hacer append como fallback con warning
                p = doc.add_paragraph(text, style=style)
                warnings.append(f"Heading '{op.get('heading_text')}' not found; appended paragraph at end.")

        elif op_name == "append_table":
            headers = op.get("headers", [])
            rows = op.get("rows", [])
            tbl_style = op.get("style", "Light Grid Accent 1")
            tbl = doc.add_table(rows=1, cols=len(headers))
            tbl.style = tbl_style
            hdr = tbl.rows[0].cells
            for i, hv in enumerate(headers):
                hdr[i].text = str(hv)
            for row in rows:
                cells = tbl.add_row().cells
                for i, cv in enumerate(row):
                    if i < len(cells):
                        cells[i].text = str(cv)

        else:
            raise ValueError(f"Operación de edición no soportada: '{op_name}'")

    doc.save(target_out)
    return {
        "status": "ok",
        "path": os.path.abspath(target_out),
        "bytes": os.path.getsize(target_out),
        "fidelity": fidelity,
        "warnings": warnings,
        "operations_applied": len(ops),
    }


def mail_merge(
    template_path: str,
    dataset_csv: str = "",
    dataset_json: str = "",
    output_prefix: str = "",
    fields: list[str] | str | None = None,
) -> dict[str, Any]:
    """Genera N documentos .docx combinando una plantilla docxtpl con un dataset CSV o JSON.

    - template_path: ruta a la plantilla .docx con placeholders {{ variable }}.
    - dataset_csv: ruta a archivo CSV o texto CSV plano.
    - dataset_json: ruta a archivo JSON o texto/estructura JSON.
    - output_prefix: prefijo para los archivos generados (ej: 'docs/carta' -> 'docs/carta_1.docx').
    - fields: lista o nombres de campos opcionales para filtrar columnas.

    Devuelve dict con status, n_docs, paths y fields.
    """
    import io
    import json
    from pathlib import Path
    import pandas as pd
    from docxtpl import DocxTemplate
    from .templates_pack import resolve_template_path

    try:
        tpl_path = resolve_template_path(template_path)
    except Exception:
        tpl_path = os.path.abspath(os.path.expanduser(str(template_path)))

    if not os.path.exists(tpl_path):
        raise FileNotFoundError(f"Plantilla no encontrada: {template_path}")

    df = None
    if dataset_csv:
        csv_str = str(dataset_csv).strip()
        if os.path.isfile(csv_str):
            df = pd.read_csv(csv_str)
        else:
            df = pd.read_csv(io.StringIO(csv_str))
    elif dataset_json:
        if isinstance(dataset_json, str):
            json_str = dataset_json.strip()
            if os.path.isfile(json_str):
                with open(json_str, "r", encoding="utf-8") as f:
                    raw_data = json.load(f)
            else:
                raw_data = json.loads(json_str)
        else:
            raw_data = dataset_json

        if isinstance(raw_data, list):
            df = pd.DataFrame(raw_data)
        elif isinstance(raw_data, dict):
            rows = raw_data.get("rows") or raw_data.get("data") or [raw_data]
            df = pd.DataFrame(rows)
        else:
            raise ValueError(f"Formato JSON no reconocido para dataset: {type(raw_data)}")
    else:
        raise ValueError("Debe especificarse dataset_csv o dataset_json.")

    if df is None or df.empty:
        raise ValueError("El dataset está vacío o no contiene filas.")

    if fields:
        if isinstance(fields, str):
            try:
                f_list = json.loads(fields)
            except Exception:
                f_list = [f.strip() for f in fields.split(",") if f.strip()]
        else:
            f_list = list(fields)
        valid_cols = [c for c in f_list if c in df.columns]
        if valid_cols:
            df = df[valid_cols]

    df.columns = [str(c) for c in df.columns]
    records = df.fillna("").to_dict(orient="records")

    if not output_prefix:
        base_dir = os.path.dirname(os.path.abspath(tpl_path))
        stem = Path(tpl_path).stem
        output_prefix = os.path.join(base_dir, f"{stem}_merged")
    else:
        output_prefix = str(output_prefix)
        if output_prefix.endswith(".docx"):
            output_prefix = output_prefix[:-5]

    paths = []
    for idx, rec in enumerate(records, start=1):
        target_file = safe_out(f"{output_prefix}_{idx}.docx")
        tpl = DocxTemplate(tpl_path)
        tpl.render(rec)
        tpl.save(target_file)
        if os.path.getsize(target_file) < 300:
            raise RuntimeError(f"DOCX generado sospechosamente pequeño ({os.path.getsize(target_file)}B): {target_file}")
        paths.append(os.path.abspath(target_file))

    return {
        "status": "ok",
        "template": os.path.abspath(tpl_path),
        "n_docs": len(paths),
        "paths": paths,
        "fields": list(df.columns),
    }


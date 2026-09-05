"""Extracción estructurada de texto y tablas desde documentos Office (.docx, .pptx, .xlsx, .xlsm)."""
from __future__ import annotations
import os
from pathlib import Path
from typing import Any


def read_office(path: str, max_rows: int = 500) -> dict[str, Any]:
    """Extrae texto estructurado, tablas y diapositivas de archivos Office (.docx, .pptx, .xlsx, .xlsm).

    - path: ruta al archivo de Office.
    - max_rows: límite de filas por hoja al leer planillas Excel (default 500 para control de tokens).

    Devuelve dict con estructura específica según el formato detectado.
    """
    path = os.path.abspath(os.path.expanduser(str(path)))
    if not os.path.exists(path):
        raise FileNotFoundError(f"Archivo no encontrado: {path}")

    ext = Path(path).suffix.lower()

    if ext == ".docx":
        from docx import Document
        doc = Document(path)
        paragraphs = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                st = p.style.name if p.style else "Normal"
                paragraphs.append({"text": txt, "style": st})

        tables = []
        for tbl in doc.tables:
            tbl_rows = []
            for row in tbl.rows:
                tbl_rows.append([cell.text.strip() for cell in row.cells])
            tables.append(tbl_rows)

        full_text_parts = [p["text"] for p in paragraphs]
        for tbl in tables:
            for r in tbl:
                full_text_parts.append(" | ".join(r))

        return {
            "status": "ok",
            "path": path,
            "format": "docx",
            "n_paragraphs": len(paragraphs),
            "n_tables": len(tables),
            "paragraphs": paragraphs,
            "tables": tables,
            "text": "\n\n".join(full_text_parts).strip(),
        }

    elif ext == ".pptx":
        from pptx import Presentation
        prs = Presentation(path)
        slides_data = []
        all_text_parts = []

        for idx, slide in enumerate(prs.slides, start=1):
            slide_texts = []
            slide_tables = []
            title = ""

            for shape in slide.shapes:
                if shape.has_text_frame:
                    t = shape.text_frame.text.strip()
                    if t:
                        if not title and (shape == slide.shapes.title or "title" in shape.name.lower()):
                            title = t
                        slide_texts.append(t)
                        all_text_parts.append(t)
                if shape.has_table:
                    tbl_content = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
                    slide_tables.append(tbl_content)
                    for r in tbl_content:
                        all_text_parts.append(" | ".join(r))

            slides_data.append({
                "slide": idx,
                "title": title or (slide_texts[0] if slide_texts else f"Slide {idx}"),
                "text": slide_texts,
                "tables": slide_tables,
            })

        return {
            "status": "ok",
            "path": path,
            "format": "pptx",
            "n_slides": len(prs.slides),
            "slides": slides_data,
            "text": "\n\n".join(all_text_parts).strip(),
        }

    elif ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheets_data = []
        limit_r = int(max_rows) if max_rows else 500

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if r_idx > limit_r:
                    break
                if any(c is not None and str(c).strip() != "" for c in row):
                    rows.append([str(c) if c is not None else "" for c in row])

            headers = rows[0] if rows else []
            data_rows = rows[1:] if len(rows) > 1 else []
            sheets_data.append({
                "name": sheet_name,
                "max_row": ws.max_row or len(rows),
                "max_column": ws.max_column or (len(headers) if headers else 0),
                "headers": headers,
                "rows": data_rows,
            })
        wb.close()

        return {
            "status": "ok",
            "path": path,
            "format": "xlsx",
            "n_sheets": len(sheets_data),
            "sheets": sheets_data,
        }

    else:
        raise ValueError(f"Formato no soportado para lectura estructurada: '{ext}'. Formatos válidos: .docx, .pptx, .xlsx, .xlsm")

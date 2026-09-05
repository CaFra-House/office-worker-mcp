"""Generación y edición de Excel (.xlsx / .xlsm) con openpyxl, aplicando tema corporativo."""
from __future__ import annotations
import json
import os
from typing import Any, Sequence

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

from .security import safe_out

# Fórmulas de asistencia para agentes
FORMULA_HELPERS = {
    "SUM": "=SUM({range})",
    "SUMIF": "=SUMIF({range}, {criteria}, {sum_range})",
    "AVERAGEIF": "=AVERAGEIF({range}, {criteria}, {avg_range})",
    "VLOOKUP": "=VLOOKUP({lookup_value}, {table_array}, {col_index}, {range_lookup})",
    "XLOOKUP": "=XLOOKUP({lookup_value}, {lookup_array}, {return_array})",
    "COUNTIFS": "=COUNTIFS({criteria_range1}, {criteria1})",
}


def _hex(h: str) -> str:
    """Convierte hex #003366 a formato ARGB FF003366 de openpyxl."""
    return "FF" + h.lstrip("#").upper()


def _apply_chart(ws, chart_spec: dict) -> None:
    """Inserta un gráfico nativo de Excel en la hoja de cálculo."""
    c_type = str(chart_spec.get("type") or chart_spec.get("chart_type") or "bar").lower().strip()
    title = chart_spec.get("title", "")
    target_cell = chart_spec.get("cell") or chart_spec.get("target_cell") or "E2"
    data_range = chart_spec.get("data_range") or chart_spec.get("data")
    cats_range = chart_spec.get("categories_range") or chart_spec.get("categories")

    if c_type in ("pie", "piechart", "pie_chart"):
        chart = PieChart()
    elif c_type in ("line", "linechart", "line_chart"):
        chart = LineChart()
    else:
        chart = BarChart()

    if title:
        chart.title = title

    if data_range and isinstance(data_range, str) and ":" in data_range:
        min_c, min_r, max_c, max_r = range_boundaries(data_range)
        data = Reference(ws, min_col=min_c, min_row=min_r, max_col=max_c, max_row=max_r)
        chart.add_data(data, titles_from_data=True)
    elif ws.max_column >= 2 and ws.max_row >= 2:
        # Por defecto toma desde columna 2 hasta fin de datos
        data = Reference(ws, min_col=2, min_row=1, max_col=ws.max_column, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)

    if cats_range and isinstance(cats_range, str) and ":" in cats_range:
        c_min_c, c_min_r, c_max_c, c_max_r = range_boundaries(cats_range)
        cats = Reference(ws, min_col=c_min_c, min_row=c_min_r, max_col=c_max_c, max_row=c_max_r)
        chart.set_categories(cats)
    elif ws.max_row >= 2:
        # Por defecto categorías en columna 1 excluyendo encabezado
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        chart.set_categories(cats)

    ws.add_chart(chart, target_cell)


def _apply_table_style(ws, table_style: str | None = None, table_name: str | None = None, table_range: str | None = None) -> None:
    """Añade una tabla estructurada con estilo y autofiltro integrado."""
    if not table_range:
        if ws.max_column < 1 or ws.max_row < 1:
            return
        start_row = 1
        if ws.max_row >= 2 and ws.max_column >= 2:
            if ws.cell(1, 2).value is None and ws.cell(2, 1).value is not None:
                start_row = 2
        last_col = get_column_letter(ws.max_column)
        table_range = f"A{start_row}:{last_col}{ws.max_row}"

    min_c, min_r, max_c, max_r = range_boundaries(table_range)
    # openpyxl requiere que todos los encabezados de tabla sean strings no vacíos
    for c in range(min_c, max_c + 1):
        cell = ws.cell(min_r, c)
        if cell.value is None or str(cell.value).strip() == "":
            cell.value = f"Column{c}"
        else:
            cell.value = str(cell.value)

    t_name = (table_name or f"Table_{len(ws.tables) + 1}")[:31].replace(" ", "_")
    tab = Table(displayName=t_name, ref=table_range)
    style_name = table_style if (table_style and isinstance(table_style, str)) else "TableStyleMedium9"
    tab.tableStyleInfo = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )
    ws.add_table(tab)


def _apply_autofilter(ws, filter_range: str | None = None) -> None:
    """Activa el autofiltro en el rango indicado o en todo el rango con datos."""
    if not filter_range:
        if ws.max_column < 1 or ws.max_row < 1:
            return
        last_col = get_column_letter(ws.max_column)
        filter_range = f"A1:{last_col}{ws.max_row}"
    ws.auto_filter.ref = filter_range


def create_excel(
    out_path: str,
    title: str = "",
    sheets: list[dict[str, Any]] | None = None,
    theme: str | None = None,
    table_style: str | None = None,
    auto_filter: bool = False,
) -> str:
    """Crea un archivo .xlsx profesional con soporte multi-hoja, tablas estructuradas, autofiltro y gráficos.

    - sheets: lista de dicts {"name":"Hoja", "headers":[...], "rows":[[...], ...], "charts":[...], "table_style":..., "auto_filter":...}
    - table_style: estilo opcional de tabla estructurada (ej: "TableStyleMedium9", "TableStyleLight1").
    - auto_filter: si es True, activa autofiltro en los encabezados.
    - theme: paleta corporativa aplicada a encabezados y filas alternas.

    Devuelve ruta absoluta del archivo generado.
    """
    from .themes import load_theme

    th = load_theme(theme)
    primary = _hex(th["primary"])
    alt = _hex(th.get("row_alt", "#F5F7FA"))

    out_path = safe_out(out_path)

    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="FFC8D4E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if not sheets:
        ws = wb.create_sheet("Datos")
        ws["A1"] = title or "The Office Worker"
        ws["A1"].font = Font(bold=True, color=primary, size=12)
        wb.save(out_path)
        return out_path

    for sh in sheets:
        ws = wb.create_sheet(sh.get("name", "Hoja")[:31])
        headers = sh.get("headers", [])
        rows = sh.get("rows", [])
        charts = sh.get("charts", [])
        sh_table_style = sh.get("table_style", table_style)
        sh_auto_filter = sh.get("auto_filter", auto_filter)

        if title:
            ws.cell(1, 1, title).font = Font(bold=True, color=primary, size=12)
        hrow = 2 if title else 1

        for c, hv in enumerate(headers, start=1):
            cell = ws.cell(hrow, c, str(hv))
            cell.font = Font(bold=True, color="FFFFFFFF")
            cell.fill = PatternFill("solid", fgColor=primary)
            cell.alignment = Alignment(horizontal="left")
            cell.border = border

        ridx = hrow + 1
        for i, row in enumerate(rows):
            for c, cv in enumerate(row, start=1):
                cell = ws.cell(ridx, c, cv)
                cell.border = border
                cell.alignment = Alignment(vertical="top")
                if i % 2 == 1:
                    cell.fill = PatternFill("solid", fgColor=alt)
            ridx += 1

        # Auto ajuste de ancho aproximado de columnas (máx 40)
        for c in range(1, len(headers) + 1):
            maxlen = max([len(str(headers[c - 1]))] + [len(str(r[c - 1])) for r in rows if c - 1 < len(r)] or [8])
            ws.column_dimensions[get_column_letter(c)].width = min(max(maxlen + 2, 9), 40)

        # Tablas estructuradas o autofiltro si se solicitan
        if sh_table_style and headers:
            last_col = get_column_letter(len(headers))
            last_row = max(hrow + len(rows), hrow + 1)
            _apply_table_style(ws, table_style=sh_table_style, table_range=f"A{hrow}:{last_col}{last_row}")
        elif sh_auto_filter and headers:
            last_col = get_column_letter(len(headers))
            last_row = max(hrow + len(rows), hrow + 1)
            _apply_autofilter(ws, filter_range=f"A{hrow}:{last_col}{last_row}")

        # Gráficos si se declararon en la hoja
        for ch in charts:
            try:
                _apply_chart(ws, ch)
            except Exception:
                pass

    wb.save(out_path)
    if os.path.getsize(out_path) < 300:
        raise RuntimeError(f"XLSX sospechosamente pequeño ({os.path.getsize(out_path)}B)")
    return out_path


def edit_excel(
    input_path: str,
    operations: list[dict[str, Any]] | str,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Modifica un archivo Excel existente (.xlsx / .xlsm) preservando estilos y formato original.

    Operaciones soportadas (lista de dicts):
    - {"op": "set_cell", "sheet"?: str, "coordinate"?: "B5", "row"?: int, "column"?: int|str, "value": Any}
    - {"op": "append_row", "sheet"?: str, "row": list}
    - {"op": "add_column", "sheet"?: str, "header"?: str, "values": list, "column"?: int|str}
    - {"op": "add_chart", "sheet"?: str, "chart_type": "bar"|"line"|"pie", "title"?: str, "data_range"?: str, "categories_range"?: str, "target_cell"?: str}
    - {"op": "add_table", "sheet"?: str, "name"?: str, "table_style"?: str, "range"?: str}
    - {"op": "auto_filter", "sheet"?: str, "range"?: str}

    Preserva macros VBA si el archivo es .xlsm (mediante keep_vba=True).
    Devuelve dict con fidelity honesta ("rich"|"clean"), warnings y ruta absoluta.
    """
    input_path = os.path.abspath(os.path.expanduser(str(input_path)))
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Archivo Excel no encontrado: {input_path}")

    target_out = safe_out(output_path if output_path else input_path)

    is_xlsm = input_path.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(input_path, keep_vba=is_xlsm)

    warnings: list[str] = []
    fidelity = "rich"

    if is_xlsm:
        warnings.append("VBA macros preserved via keep_vba=True; openpyxl does not execute or edit VBA macro code.")
        fidelity = "clean"

    if isinstance(operations, str):
        try:
            ops = json.loads(operations)
        except json.JSONDecodeError as exc:
            raise ValueError(f"operations inválido (JSON malformado): {exc}")
    else:
        ops = list(operations or [])

    has_formulas = False
    sheets_touched = set()

    for op in ops:
        op_name = str(op.get("op") or op.get("operation") or "").lower().strip()
        sh_name = op.get("sheet")
        ws = wb[sh_name] if (sh_name and sh_name in wb.sheetnames) else wb.active
        sheets_touched.add(ws.title)

        if op_name == "set_cell":
            val = op.get("value")
            if val is None and "formula" in op:
                val = op.get("formula")
            coord = op.get("coordinate") or op.get("cell")
            if coord:
                ws[coord] = val
            else:
                r = int(op.get("row", 1))
                c_raw = op.get("column", 1)
                c = column_index_from_string(c_raw) if isinstance(c_raw, str) and not c_raw.isdigit() else int(c_raw)
                ws.cell(row=r, column=c, value=val)

            if isinstance(val, str) and val.startswith("="):
                has_formulas = True

        elif op_name == "append_row":
            row_data = op.get("row") or op.get("values") or []
            ws.append(row_data)
            if any(isinstance(v, str) and v.startswith("=") for v in row_data):
                has_formulas = True

        elif op_name == "add_column":
            values = op.get("values") or []
            header = op.get("header")
            col_target = op.get("column")
            if col_target is not None:
                c_idx = column_index_from_string(col_target) if isinstance(col_target, str) and not col_target.isdigit() else int(col_target)
            else:
                c_idx = ws.max_column + 1

            start_row = 1
            if header is not None:
                ws.cell(row=1, column=c_idx, value=header)
                start_row = 2

            for offset, val in enumerate(values):
                ws.cell(row=start_row + offset, column=c_idx, value=val)
                if isinstance(val, str) and val.startswith("="):
                    has_formulas = True

        elif op_name in ("add_chart", "chart"):
            _apply_chart(ws, op)

        elif op_name in ("add_table", "table"):
            t_style = op.get("table_style") or op.get("style") or "TableStyleMedium9"
            t_name = op.get("name") or f"Table_{len(ws.tables) + 1}"
            t_range = op.get("range")
            _apply_table_style(ws, table_style=t_style, table_name=t_name, table_range=t_range)

        elif op_name in ("auto_filter", "autofilter"):
            f_range = op.get("range")
            _apply_autofilter(ws, filter_range=f_range)

        else:
            raise ValueError(f"Operación de edición no soportada: '{op_name}'")

    if has_formulas:
        warnings.append("Formulas written to cells; values will be calculated when opened in spreadsheet software.")

    wb.save(target_out)
    return {
        "status": "ok",
        "path": os.path.abspath(target_out),
        "bytes": os.path.getsize(target_out),
        "fidelity": fidelity,
        "warnings": warnings,
        "operations_applied": len(ops),
        "sheets_modified": list(sheets_touched),
    }

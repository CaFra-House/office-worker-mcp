"""Generación de Excel (.xlsx) con openpyxl, aplicando tema corporativo."""
from __future__ import annotations
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def _hex(h):  # "#003366" -> "FF003366" para openpyxl (ARGB)
    return "FF" + h.lstrip("#").upper()

def create_excel(out_path, title="", sheets=None, theme=None):
    """Crea un .xlsx profesional.

    sheets: lista de dicts {"name":"Hoja", "headers":[...], "rows":[[...], ...]}
            Si None → una hoja única "Datos" vacía con título en A1.
    Devuelve ruta absoluta. Aplica paleta del tema al header y filas alternas.
    """
    from .themes import load_theme
    th = load_theme(theme)
    primary = _hex(th["primary"]); alt = _hex(th.get("row_alt","#F5F7FA"))

    out_path = os.path.abspath(os.path.expanduser(out_path))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style="thin", color="FFC8D4E0")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)

    if not sheets:
        ws = wb.create_sheet("Datos"); ws["A1"]=title or "The Office Worker"; ws["A1"].font=Font(bold=True,color=primary,size=12); wb.save(out_path); return out_path

    for sh in sheets:
        ws = wb.create_sheet(sh.get("name","Hoja")[:31])
        headers = sh.get("headers",[]); rows = sh.get("rows",[])
        start_row = 1
        if title:
            ws.cell(1,1,title).font = Font(bold=True,color=primary,size=12)
        hrow = 2 if title else 1
        for c,hv in enumerate(headers, start=1):
            cell = ws.cell(hrow,c,str(hv)); cell.font=Font(bold=True,color="FFFFFFFF")
            cell.fill=PatternFill("solid",fgColor=primary); cell.alignment=Alignment(horizontal="left"); cell.border=border
        ridx = hrow+1; i=0
        for row in rows:
            for c,cv in enumerate(row,start=1):
                cell=ws.cell(ridx,c,cv); cell.border=border; cell.alignment=Alignment(vertical="top")
                if i%2==1: cell.fill=PatternFill("solid",fgColor=alt)
            ridx+=1; i+=1
        # auto ancho aprox por contenido (máx 40)
        for c in range(1,len(headers)+1):
            maxlen=max([len(str(headers[c-1]))] + [len(str(r[c-1])) for r in rows if c-1<len(r)] or [8])
            ws.column_dimensions[get_column_letter(c)].width=min(max(maxlen+2,9),40)

    wb.save(out_path)
    if os.path.getsize(out_path) < 300: raise RuntimeError(f"XLSX sospechosamente pequeño ({os.path.getsize(out_path)}B)")
    return out_path

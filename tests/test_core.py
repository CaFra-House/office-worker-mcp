"""Test del núcleo: render_document → PDF real. (H1)"""
import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from office_worker.core.templates import render_pdf

TPL = """<h1>{{ titulo }}</h1><p class="muted">{{ subtitulo }} · {{ fecha }}</p>
{% if tabla is defined and tabla %}{{ tabla }}{% endif %}"""

def test_render_pdf(tmp_path):
    out = str(tmp_path / "out.pdf")
    data = {
        "titulo": "Informe de Prueba",
        "subtitulo": "The Office Worker H1",
        "fecha": "2026-09-05",
        "headers": ["Rubro", "Valor"],
        "rows": [["Ingresos", "$1.200.000"], ["Gastos", "$800.000"]],
    }
    # emular el helper de tabla que hace el MCP server antes de llamar al núcleo
    headers = data["headers"]; rows = data["rows"]
    thead = "".join(f"<th>{h}</th>" for h in headers)
    body = "\n".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    data["tabla"] = f"<table><thead><tr>{thead}</tr></thead><tbody>{body}</tbody></table>"

    path = render_pdf(TPL, out, data=data)  # tema por defecto ADEN
    assert os.path.exists(path), "PDF no creado"
    size = os.path.getsize(path)
    assert size > 500, f"PDF demasiado chico: {size}"
    head = open(path, "rb").read(5)
    assert head == b"%PDF-", f"No es PDF válido: {head!r}"
    print(f"OK — PDF generado: {path} ({size} bytes)")

if __name__ == "__main__":
    import pathlib, tempfile
    with tempfile.TemporaryDirectory() as d:
        test_render_pdf(pathlib.Path(d))

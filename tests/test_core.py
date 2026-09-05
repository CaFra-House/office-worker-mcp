"""Test del núcleo: render_document, pdf_tools (form fill, ocr, convert, manipulate), temas. (v0.2.0)"""
import os, shutil, sys, pathlib
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from office_worker.core.templates import render_pdf
from office_worker.core.word import create_word
from office_worker.core.themes import load_theme, THEMES
from office_worker.core.pdf_tools import (
    fill_pdf_form,
    ocr_pdf,
    convert_office_to_pdf,
    manipulate_pdf,
)

TPL = """<h1>{{ titulo }}</h1><p class="muted">{{ subtitulo }} · {{ fecha }}</p>
{% if tabla is defined and tabla %}{{ tabla }}{% endif %}"""


def test_render_pdf_and_themes(tmp_path):
    out = str(tmp_path / "out.pdf")
    data = {
        "titulo": "Informe de Prueba",
        "subtitulo": "The Office Worker v0.2.0",
        "fecha": "2026-09-05",
        "headers": ["Rubro", "Valor"],
        "rows": [["Ingresos", "$1.200.000"], ["Gastos", "$800.000"]],
    }
    headers = data["headers"]; rows = data["rows"]
    thead = "".join(f"<th>{h}</th>" for h in headers)
    body = "\n".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    data["tabla"] = f"<table><thead><tr>{thead}</tr></thead><tbody>{body}</tbody></table>"

    # Test con tema nuevo corporate-blue
    path = render_pdf(TPL, out, data=data, theme="corporate-blue")
    assert os.path.exists(path), "PDF no creado"
    size = os.path.getsize(path)
    assert size > 500, f"PDF demasiado chico: {size}"
    assert open(path, "rb").read(5) == b"%PDF-", "No es PDF válido"


def test_render_pdf_with_logo(tmp_path):
    try:
        from PIL import Image
    except ImportError:
        pytest.skip("Pillow no instalado")

    logo_path = str(tmp_path / "logo.png")
    img = Image.new("RGB", (150, 40), color=(15, 76, 129))
    img.save(logo_path)

    out = str(tmp_path / "out_logo.pdf")
    path = render_pdf("<h1>Con Logo</h1>", out, theme="claro", logo=logo_path)
    assert os.path.exists(path)
    assert os.path.getsize(path) > 500
    assert open(path, "rb").read(5) == b"%PDF-"


def test_themes_catalog():
    for name in ["aden", "claro", "oscuro", "minimal", "corporate-blue"]:
        t = load_theme(name)
        assert t["name"] == name or name in t["name"]
        assert "primary" in t and "bg" in t and "text" in t


def test_fill_pdf_form(tmp_path):
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.colors import black, white
        from pypdf import PdfReader
    except ImportError:
        pytest.skip("reportlab o pypdf no disponibles")

    form_pdf = str(tmp_path / "form.pdf")
    filled_pdf = str(tmp_path / "filled.pdf")

    # Crear PDF con AcroForm interactivo usando reportlab
    c = canvas.Canvas(form_pdf)
    c.drawString(100, 700, "Nombre:")
    form = c.acroForm
    form.textfield(name="nombre", tooltip="Nombre completo", x=150, y=695, width=200, height=20, borderColor=black, fillColor=white)
    c.showPage()
    c.save()

    assert os.path.exists(form_pdf)

    # Rellenar formulario
    res_path = fill_pdf_form(form_pdf, {"nombre": "Ana Gomez"}, filled_pdf)
    assert os.path.exists(res_path)
    assert open(res_path, "rb").read(5) == b"%PDF-"

    # Verificar campo relleno
    r = PdfReader(res_path)
    fields = r.get_fields()
    assert fields is not None and "nombre" in fields
    assert fields["nombre"].get("/V") == "Ana Gomez"

    # Verificar excepción si fields_dict está vacío
    with pytest.raises(ValueError):
        fill_pdf_form(form_pdf, {}, str(tmp_path / "err.pdf"))


def test_convert_office_to_pdf(tmp_path):
    if not (shutil.which("soffice") or shutil.which("libreoffice")):
        pytest.skip("soffice/LibreOffice no disponible en el sistema")

    docx_path = str(tmp_path / "doc.docx")
    pdf_out = str(tmp_path / "doc.pdf")

    create_word(docx_path, title="Prueba Conversión", subtitle="LibreOffice", blocks=[{"type": "p", "text": "Contenido para convertir"}])
    assert os.path.exists(docx_path)
    # Magic bytes PK para docx (zip)
    assert open(docx_path, "rb").read(2) == b"PK"

    res_path = convert_office_to_pdf(docx_path, pdf_out)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 0
    assert open(res_path, "rb").read(5) == b"%PDF-"


def test_manipulate_pdf(tmp_path):
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        pytest.skip("pypdf no disponible")

    pdf1 = str(tmp_path / "p1.pdf")
    pdf2 = str(tmp_path / "p2.pdf")
    merged = str(tmp_path / "merged.pdf")
    extracted = str(tmp_path / "extracted.pdf")
    rotated = str(tmp_path / "rotated.pdf")

    w1 = PdfWriter(); w1.add_blank_page(width=100, height=100); w1.write(pdf1)
    w2 = PdfWriter(); w2.add_blank_page(width=100, height=100); w2.write(pdf2)

    # 1. Merge
    res_m = manipulate_pdf(operation="merge", out=merged, files=[pdf1, pdf2])
    assert os.path.exists(res_m)
    assert len(PdfReader(res_m).pages) == 2
    assert open(res_m, "rb").read(5) == b"%PDF-"

    # 2. Extract
    res_e = manipulate_pdf(operation="extract", out=extracted, input_path=merged, pages="1")
    assert os.path.exists(res_e)
    assert len(PdfReader(res_e).pages) == 1
    assert open(res_e, "rb").read(5) == b"%PDF-"

    # 3. Rotate
    res_r = manipulate_pdf(operation="rotate", out=rotated, input_path=extracted, angle=90)
    assert os.path.exists(res_r)
    r_doc = PdfReader(res_r)
    assert r_doc.pages[0].rotation == 90
    assert open(res_r, "rb").read(5) == b"%PDF-"


def test_ocr_pdf(tmp_path):
    if not shutil.which("tesseract"):
        pytest.skip("Tesseract OCR no instalado en el sistema")
    try:
        from PIL import Image, ImageDraw
        import pytesseract
    except ImportError:
        pytest.skip("PIL o pytesseract no disponibles")

    img_path = str(tmp_path / "test_ocr.png")
    pdf_out = str(tmp_path / "searchable.pdf")

    img = Image.new("RGB", (300, 80), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((20, 30), "HELLO WORLD", fill=(0, 0, 0))
    img.save(img_path)

    text = ocr_pdf(img_path, lang="eng", out=pdf_out)
    clean_text = text.replace(" ", "").upper()
    assert "HELLO" in clean_text or "HELLOWORLD" in clean_text
    assert os.path.exists(pdf_out)
    assert os.path.getsize(pdf_out) > 0
    assert open(pdf_out, "rb").read(5) == b"%PDF-"

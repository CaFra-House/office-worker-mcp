"""Herramientas avanzadas de PDF: formularios, OCR, conversión y manipulación.

Complementa office_worker.core.pdf con operaciones de escritura y transformación:
- fill_pdf_form: rellenar formularios AcroForm con pypdf
- ocr_pdf: extracción de texto y generación de PDF con capa de texto (Tesseract + PyMuPDF)
- convert_office_to_pdf: conversión de documentos Office (.docx, .xlsx, .pptx) vía LibreOffice
- manipulate_pdf: merge, extract (rango de páginas) y rotate (ángulo)
"""
from __future__ import annotations
import io
import os
import shutil
import subprocess
from pathlib import Path
from typing import Mapping, Sequence


def fill_pdf_form(pdf_in: str, fields_dict: dict, out: str) -> str:
    """Rellena campos de un formulario AcroForm en un PDF.

    - pdf_in: ruta al archivo PDF con formulario interactivo.
    - fields_dict: diccionario con pares campo-valor (ej: {'nombre': 'Juan'}).
    - out: ruta donde se guardará el PDF completado.

    Lanza ValueError si fields_dict está vacío o si el PDF no contiene formulario.
    Devuelve la ruta absoluta del PDF generado.
    """
    from pypdf import PdfReader, PdfWriter

    if not fields_dict or not isinstance(fields_dict, (dict, Mapping)):
        raise ValueError("fields_dict no puede estar vacío y debe ser un diccionario.")

    pdf_in = os.path.abspath(os.path.expanduser(str(pdf_in)))
    out = os.path.abspath(os.path.expanduser(str(out)))
    if not os.path.exists(pdf_in):
        raise FileNotFoundError(f"PDF de entrada no encontrado: {pdf_in}")

    reader = PdfReader(pdf_in)
    existing_fields = reader.get_fields()
    if not existing_fields:
        raise ValueError(f"El archivo PDF no contiene campos de formulario interactivos: {pdf_in}")

    writer = PdfWriter(clone_from=pdf_in)
    writer.update_page_form_field_values(None, fields_dict)

    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "wb") as f:
        writer.write(f)

    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise RuntimeError(f"Error al escribir el PDF rellenado en: {out}")
    return out


def ocr_pdf(path: str, lang: str = "spa", out: str | None = None) -> str:
    """Extrae texto mediante OCR de imágenes o PDFs escaneados y opcionalmente crea un PDF con capa de texto.

    - path: ruta a una imagen (.png, .jpg, etc.) o archivo PDF.
    - lang: código de idioma Tesseract (ej: 'spa', 'eng', 'spa+eng').
    - out: ruta opcional para guardar el PDF con capa de texto buscable.

    Detecta si el PDF ya tiene capa de texto vía PyMuPDF (fitz get_text).
    Devuelve el texto extraído.
    """
    import fitz
    from PIL import Image
    import pytesseract

    path = os.path.abspath(os.path.expanduser(str(path)))
    if not os.path.exists(path):
        raise FileNotFoundError(f"Archivo no encontrado: {path}")

    out_path = os.path.abspath(os.path.expanduser(str(out))) if out else None
    if out_path:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

    lower_path = path.lower()
    is_img = lower_path.endswith((".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp"))

    if is_img:
        img = Image.open(path)
        text = pytesseract.image_to_string(img, lang=lang).strip()
        if out_path:
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="pdf")
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
        return text

    # Tratar como PDF
    doc = fitz.open(path)
    pages_text = [page.get_text().strip() for page in doc]
    existing_text = "\n\n".join(t for t in pages_text if t).strip()

    # Si ya tiene capa de texto existente
    if existing_text:
        if out_path:
            doc.save(out_path)
        doc.close()
        return existing_text

    # PDF escaneado sin capa de texto: procesar página por página
    extracted = []
    out_doc = fitz.open() if out_path else None

    for page in doc:
        pix = page.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        page_text = pytesseract.image_to_string(img, lang=lang).strip()
        extracted.append(page_text)
        if out_doc is not None:
            page_pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="pdf")
            page_pdf = fitz.open("pdf", page_pdf_bytes)
            out_doc.insert_pdf(page_pdf)
            page_pdf.close()

    doc.close()
    if out_doc is not None:
        out_doc.save(out_path)
        out_doc.close()

    return "\n\n".join(extracted).strip()


def convert_office_to_pdf(file_in: str, out: str) -> str:
    """Convierte archivos Office (.docx, .xlsx, .pptx) a PDF usando LibreOffice headless.

    - file_in: ruta al archivo office fuente.
    - out: ruta del PDF destino.

    Usa subprocess con timeout de 120s y verifica que el archivo resultante sea > 0 bytes.
    Devuelve la ruta absoluta del PDF generado.
    """
    file_in = os.path.abspath(os.path.expanduser(str(file_in)))
    out = os.path.abspath(os.path.expanduser(str(out)))
    if not os.path.exists(file_in):
        raise FileNotFoundError(f"Archivo de entrada no encontrado: {file_in}")

    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        raise FileNotFoundError("LibreOffice (soffice) no está instalado o disponible en PATH.")

    os.makedirs(os.path.dirname(out), exist_ok=True)
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        cmd = [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", td, file_in]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        except subprocess.TimeoutExpired as exc:
            raise TimeoutError(f"La conversión de LibreOffice excedió el timeout de 120s para {file_in}") from exc

        if res.returncode != 0:
            raise RuntimeError(f"Error en LibreOffice (código {res.returncode}): {res.stderr or res.stdout}")

        stem = Path(file_in).stem
        converted_file = os.path.join(td, f"{stem}.pdf")
        if not os.path.exists(converted_file) or os.path.getsize(converted_file) == 0:
            raise RuntimeError(f"LibreOffice no produjo un PDF válido (>0 bytes) para {file_in}")

        shutil.copy2(converted_file, out)

    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise RuntimeError(f"Error al escribir el archivo PDF convertido en: {out}")
    return out


def _parse_page_range(pages: str | Sequence[int], total_pages: int) -> list[int]:
    """Parsea rangos de páginas estilo '1-3,5' o listas a índices 0-based válidos."""
    if isinstance(pages, (list, tuple)):
        indices = [int(p) - 1 for p in pages]
    elif isinstance(pages, str):
        indices = []
        parts = [p.strip() for p in pages.split(",") if p.strip()]
        for part in parts:
            if "-" in part:
                start_s, end_s = part.split("-", 1)
                start, end = int(start_s), int(end_s)
                for p in range(start, end + 1):
                    indices.append(p - 1)
            else:
                indices.append(int(part) - 1)
    else:
        indices = [int(pages) - 1]

    for idx in indices:
        if idx < 0 or idx >= total_pages:
            raise IndexError(f"Página {idx + 1} fuera de rango (el documento tiene {total_pages} páginas)")
    return indices


def manipulate_pdf(
    operation: str | dict,
    out: str | None = None,
    input_path: str | None = None,
    files: list[str] | None = None,
    pages: str | Sequence[int] | None = None,
    angle: int = 90,
    **kwargs,
) -> str:
    """Manipula PDFs: merge, extract (rango de páginas) o rotate (ángulo).

    - operation: 'merge' | 'extract' | 'rotate' (o dict con los parámetros).
    - out: ruta del PDF resultante.
    - files: lista de rutas a PDFs (requerido para merge).
    - input_path: ruta al PDF fuente (requerido para extract y rotate).
    - pages: páginas a extraer o rotar (ej: '1-3', '2,4', [1, 2]). 1-indexed.
    - angle: ángulo de rotación horario en grados (ej: 90, 180, 270).

    Devuelve la ruta absoluta del PDF resultante.
    """
    from pypdf import PdfReader, PdfWriter

    if isinstance(operation, dict):
        ops = operation
        op = str(ops.get("operation") or ops.get("op") or "").lower().strip()
        out_path = ops.get("out") or ops.get("output") or out
        inp = ops.get("input_path") or ops.get("input_file") or ops.get("pdf_in") or input_path
        fls = ops.get("files") or ops.get("input_files") or files
        pgs = ops.get("pages") or pages
        ang = ops.get("angle", angle)
    else:
        op = str(operation).lower().strip()
        out_path = out or kwargs.get("output")
        inp = input_path or kwargs.get("input_file") or kwargs.get("pdf_in")
        fls = files or kwargs.get("input_files")
        pgs = pages
        ang = angle

    if not out_path:
        raise ValueError("Se requiere especificar una ruta de salida 'out'")
    out_path = os.path.abspath(os.path.expanduser(str(out_path)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    if op == "merge":
        if not fls or not isinstance(fls, (list, tuple)):
            raise ValueError("La operación 'merge' requiere una lista de archivos en 'files'")
        writer = PdfWriter()
        for f in fls:
            fp = os.path.abspath(os.path.expanduser(str(f)))
            if not os.path.exists(fp):
                raise FileNotFoundError(f"Archivo para merge no encontrado: {fp}")
            writer.append(fp)
        with open(out_path, "wb") as f_out:
            writer.write(f_out)

    elif op == "extract":
        if not inp:
            raise ValueError("La operación 'extract' requiere 'input_path'")
        inp = os.path.abspath(os.path.expanduser(str(inp)))
        if not os.path.exists(inp):
            raise FileNotFoundError(f"Archivo de entrada no encontrado: {inp}")
        if not pgs:
            raise ValueError("La operación 'extract' requiere especificar 'pages' (ej: '1-3', '2')")

        reader = PdfReader(inp)
        indices = _parse_page_range(pgs, len(reader.pages))
        writer = PdfWriter()
        for idx in indices:
            writer.add_page(reader.pages[idx])
        with open(out_path, "wb") as f_out:
            writer.write(f_out)

    elif op == "rotate":
        if not inp:
            raise ValueError("La operación 'rotate' requiere 'input_path'")
        inp = os.path.abspath(os.path.expanduser(str(inp)))
        if not os.path.exists(inp):
            raise FileNotFoundError(f"Archivo de entrada no encontrado: {inp}")

        reader = PdfReader(inp)
        total_pages = len(reader.pages)
        target_indices = _parse_page_range(pgs, total_pages) if pgs else None
        writer = PdfWriter()
        deg = int(ang)
        for i, page in enumerate(reader.pages):
            p = writer.add_page(page)
            if target_indices is None or i in target_indices:
                p.rotate(deg)
        with open(out_path, "wb") as f_out:
            writer.write(f_out)

    else:
        raise ValueError(f"Operación no soportada: '{op}'. Operaciones válidas: 'merge', 'extract', 'rotate'")

    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        raise RuntimeError(f"Error al escribir el PDF manipulado en: {out_path}")
    return out_path

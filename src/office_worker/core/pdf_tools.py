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
import signal
import subprocess
from pathlib import Path
from typing import Mapping, Sequence

from .security import safe_out


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
    out = safe_out(out)
    if not os.path.exists(pdf_in):
        raise FileNotFoundError(f"PDF de entrada no encontrado: {pdf_in}")

    reader = PdfReader(pdf_in)
    existing_fields = reader.get_fields()
    if not existing_fields:
        raise ValueError(f"El archivo PDF no contiene campos de formulario interactivos: {pdf_in}")

    writer = PdfWriter(clone_from=pdf_in)
    writer.update_page_form_field_values(None, fields_dict)

    with open(out, "wb") as f:
        writer.write(f)

    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise RuntimeError(f"Error al escribir el PDF rellenado en: {out}")
    return out


def ocr_pdf(path: str, lang: str = "spa", out: str | None = None, max_pages: int | None = None) -> str:
    """Extrae texto mediante OCR de imágenes o PDFs escaneados y opcionalmente crea un PDF con capa de texto.

    - path: ruta a una imagen (.png, .jpg, etc.) o archivo PDF.
    - lang: código de idioma Tesseract (ej: 'spa', 'eng', 'spa+eng').
    - out: ruta opcional para guardar el PDF con capa de texto buscable.
    - max_pages: límite opcional de páginas a procesar para PDFs largos.

    Detecta si el PDF ya tiene capa de texto vía PyMuPDF (fitz get_text).
    Devuelve el texto extraído.
    """
    import fitz
    from PIL import Image
    import pytesseract

    path = os.path.abspath(os.path.expanduser(str(path)))
    if not os.path.exists(path):
        raise FileNotFoundError(f"Archivo no encontrado: {path}")

    out_path = safe_out(out) if out else None

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
    limit = len(doc) if not max_pages else min(len(doc), int(max_pages))
    pages_text = [doc[i].get_text().strip() for i in range(limit)]
    existing_text = "\n\n".join(t for t in pages_text if t).strip()

    # Si ya tiene capa de texto existente
    if existing_text:
        if out_path:
            if max_pages and int(max_pages) < len(doc):
                sub_doc = fitz.open()
                sub_doc.insert_pdf(doc, from_page=0, to_page=limit - 1)
                sub_doc.save(out_path)
                sub_doc.close()
            else:
                doc.save(out_path)
        doc.close()
        return existing_text

    # PDF escaneado sin capa de texto: procesar página por página
    extracted = []
    out_doc = fitz.open() if out_path else None

    for i in range(limit):
        page = doc[i]
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
    out = safe_out(out)
    if not os.path.exists(file_in):
        raise FileNotFoundError(f"Archivo de entrada no encontrado: {file_in}")

    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        raise FileNotFoundError("LibreOffice (soffice) no está instalado o disponible en PATH.")

    import tempfile
    with tempfile.TemporaryDirectory() as td:
        cmd = [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", td, file_in]
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                start_new_session=True,
            )
            stdout, stderr = proc.communicate(timeout=120)
        except subprocess.TimeoutExpired as exc:
            try:
                os.killpg(proc.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            proc.communicate()
            raise TimeoutError(f"La conversión de LibreOffice excedió el timeout de 120s para {file_in}") from exc

        if proc.returncode != 0:
            raise RuntimeError(f"Error en LibreOffice (código {proc.returncode}): {stderr or stdout}")

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
    password: str | None = None,
    **kwargs,
) -> str:
    """Manipula PDFs: merge, extract (rango de páginas) o rotate (ángulo), con password opcional.

    - operation: 'merge' | 'extract' | 'rotate' (o dict con los parámetros).
    - out: ruta del PDF resultante.
    - files: lista de rutas a PDFs (requerido para merge).
    - input_path: ruta al PDF fuente (requerido para extract y rotate).
    - pages: páginas a extraer o rotar (ej: '1-3', '2,4', [1, 2]). 1-indexed.
    - angle: ángulo de rotación horario en grados (ej: 90, 180, 270).
    - password: clave opcional para cifrar el PDF generado.

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
        pwd = ops.get("password") or password
    else:
        op = str(operation).lower().strip()
        out_path = out or kwargs.get("output")
        inp = input_path or kwargs.get("input_file") or kwargs.get("pdf_in")
        fls = files or kwargs.get("input_files")
        pgs = pages
        ang = angle
        pwd = password or kwargs.get("password")

    if not out_path:
        raise ValueError("Se requiere especificar una ruta de salida 'out'")
    out_path = safe_out(out_path)

    if op == "merge":
        if not fls or not isinstance(fls, (list, tuple)):
            raise ValueError("La operación 'merge' requiere una lista de archivos en 'files'")
        writer = PdfWriter()
        for f in fls:
            fp = os.path.abspath(os.path.expanduser(str(f)))
            if not os.path.exists(fp):
                raise FileNotFoundError(f"Archivo para merge no encontrado: {fp}")
            writer.append(fp)
        if pwd:
            writer.encrypt(user_password=pwd)
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
        if pwd:
            writer.encrypt(user_password=pwd)
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
        if pwd:
            writer.encrypt(user_password=pwd)
        with open(out_path, "wb") as f_out:
            writer.write(f_out)

    else:
        raise ValueError(f"Operación no soportada: '{op}'. Operaciones válidas: 'merge', 'extract', 'rotate'")

    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        raise RuntimeError(f"Error al escribir el PDF manipulado en: {out_path}")
    return out_path


def sign_pdf(
    input_pdf: str,
    output: str,
    sello_img_path: str | None = None,
    cert_pem: str | None = None,
    key_pem: str | None = None,
    passphrase: str | None = None,
    reason: str | None = None,
    location: str | None = None,
    page: int = -1,
) -> str:
    """Firma un documento PDF: estampa sello PNG visible y aplica firma digital criptográfica si hay certificado.

    - input_pdf: ruta al archivo PDF a firmar.
    - output: ruta donde se guardará el PDF firmado.
    - sello_img_path: ruta opcional a un archivo PNG/JPG de sello o rúbrica visible.
    - cert_pem: ruta opcional a certificado X.509 (.pem) para firma digital criptográfica PAdES.
    - key_pem: ruta opcional a clave privada (.pem o .key). Si no se indica, se busca dentro de cert_pem.
    - passphrase: clave de desbloqueo opcional si la clave privada está protegida.
    - reason: motivo de la firma (ej: 'Aprobado', 'Conforme').
    - location: ubicación geográfica de la firma.
    - page: índice de página donde estampar el sello visual (0-indexed, default -1 para última página).

    Limitación honesta:
    El sello PNG visible se estampa siempre sobre la página indicada. La firma digital criptográfica
    PKCS#7/PAdES requiere obligatoriamente un certificado X.509 válido provisto en cert_pem.
    Si no se dispone de certificado, el documento se genera con el sello visual estampado.
    Devuelve la ruta absoluta del PDF firmado.
    """
    import fitz
    import tempfile

    input_pdf = os.path.abspath(os.path.expanduser(str(input_pdf)))
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"PDF de entrada no encontrado: {input_pdf}")

    out_path = safe_out(output)

    doc = fitz.open(input_pdf)
    if sello_img_path:
        sello_file = os.path.abspath(os.path.expanduser(str(sello_img_path)))
        if not os.path.exists(sello_file):
            doc.close()
            raise FileNotFoundError(f"Imagen de sello no encontrada: {sello_file}")

        total_pages = len(doc)
        target_idx = page if page >= 0 else (total_pages + page)
        target_idx = max(0, min(total_pages - 1, target_idx))
        target_page = doc[target_idx]

        p_rect = target_page.rect
        stamp_w, stamp_h = 160, 60
        r = fitz.Rect(
            max(20, p_rect.width - stamp_w - 40),
            max(20, p_rect.height - stamp_h - 40),
            p_rect.width - 40,
            p_rect.height - 40,
        )
        target_page.insert_image(r, filename=sello_file)

    if cert_pem:
        cert_file = os.path.abspath(os.path.expanduser(str(cert_pem)))
        if not os.path.exists(cert_file):
            doc.close()
            raise FileNotFoundError(f"Certificado PEM no encontrado: {cert_file}")
        key_file = os.path.abspath(os.path.expanduser(str(key_pem))) if key_pem else cert_file
        if not os.path.exists(key_file):
            doc.close()
            raise FileNotFoundError(f"Clave privada PEM no encontrada: {key_file}")

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_stamped = tmp.name
        try:
            doc.save(tmp_stamped)
            doc.close()

            from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
            from pyhanko.sign import fields, signers
            from pyhanko.sign.signers import SimpleSigner

            pass_bytes = passphrase.encode("utf-8") if passphrase else None
            signer = SimpleSigner.load(key_file=key_file, cert_file=cert_file, key_passphrase=pass_bytes)

            with open(tmp_stamped, "rb") as inf:
                w = IncrementalPdfFileWriter(inf)
                fields.append_signature_field(
                    w, sig_field_spec=fields.SigFieldSpec(sig_field_name="Signature1")
                )
                with open(out_path, "wb") as outf:
                    signers.sign_pdf(
                        w,
                        signers.PdfSignatureMetadata(
                            field_name="Signature1",
                            reason=reason or "Documento aprobado y firmado",
                            location=location or "",
                        ),
                        signer=signer,
                        output=outf,
                    )
        finally:
            if os.path.exists(tmp_stamped):
                try:
                    os.unlink(tmp_stamped)
                except OSError:
                    pass
    else:
        doc.save(out_path)
        doc.close()

    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        raise RuntimeError(f"Error al escribir el PDF firmado en: {out_path}")
    return out_path


def compress_pdf(input_path: str, output: str, quality: str = "med") -> dict:
    """Comprime y optimiza un archivo PDF mediante PyMuPDF (garbage collection + downsampling de imágenes).

    - input_path: ruta al archivo PDF original.
    - output: ruta de destino para el PDF optimizado.
    - quality: nivel de calidad ('low', 'med', 'high').
      * 'low': máxima compresión (resolución máx 1000px, JPEG q=45).
      * 'med': balance estándar (resolución máx 1600px, JPEG q=65).
      * 'high': preservación de alta calidad (resolución máx 2400px, JPEG q=85).

    Devuelve dict con detalles de tamaño antes, después y porcentaje de ahorro.
    """
    import fitz
    from PIL import Image

    input_path = os.path.abspath(os.path.expanduser(str(input_path)))
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Archivo PDF no encontrado: {input_path}")

    out_path = safe_out(output)
    size_before = os.path.getsize(input_path)

    q = (quality or "med").lower().strip()
    if q == "low":
        max_dim, jpeg_q = 1000, 45
    elif q == "high":
        max_dim, jpeg_q = 2400, 85
    else:
        max_dim, jpeg_q = 1600, 65

    doc = fitz.open(input_path)
    processed_xrefs = set()

    for page in doc:
        for img_info in page.get_images():
            xref = img_info[0]
            if xref in processed_xrefs:
                continue
            processed_xrefs.add(xref)
            try:
                base_img = doc.extract_image(xref)
                if not base_img or not base_img.get("image"):
                    continue
                raw_bytes = base_img["image"]
                im = Image.open(io.BytesIO(raw_bytes))
                w, h = im.size

                if w > max_dim or h > max_dim:
                    scale = min(max_dim / w, max_dim / h)
                    im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)

                out_buf = io.BytesIO()
                if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                    im.save(out_buf, format="PNG", optimize=True)
                else:
                    im_rgb = im.convert("RGB")
                    im_rgb.save(out_buf, format="JPEG", quality=jpeg_q, optimize=True)

                new_bytes = out_buf.getvalue()
                if len(new_bytes) < len(raw_bytes):
                    doc.update_stream(xref, new_bytes)
            except Exception:
                continue

    doc.save(out_path, garbage=4, deflate=True, clean=True)
    doc.close()

    size_after = os.path.getsize(out_path)
    saved_bytes = max(0, size_before - size_after)
    savings_pct = round((saved_bytes / size_before) * 100, 2) if size_before > 0 else 0.0

    return {
        "status": "ok",
        "path": out_path,
        "size_before": size_before,
        "size_after": size_after,
        "saved_bytes": saved_bytes,
        "savings_percent": savings_pct,
    }


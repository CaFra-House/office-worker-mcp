"""Generación de libros y manuales multi-capítulo profesionales (PDF y EPUB)."""
from __future__ import annotations
import html
import os
import uuid
from pathlib import Path
from typing import Any

from .security import safe_out, safe_url_fetcher
from .themes import load_theme, css_vars


_BOOK_CSS = """
@page {
  size: A4;
  margin: 2.4cm 2cm 2.4cm 2cm;
  @top-left {
    content: "The Office Worker";
    font-size: 8pt;
    color: var(--ow-muted);
    font-family: var(--ow-font-body);
  }
  @top-right {
    content: string(book-title);
    font-size: 8pt;
    color: var(--ow-muted);
    font-family: var(--ow-font-body);
  }
  @bottom-center {
    content: "Página " counter(page) " de " counter(pages);
    font-size: 8.5pt;
    color: var(--ow-muted);
    font-family: var(--ow-font-body);
  }
}

@page :first {
  margin: 0;
  @top-left { content: none; }
  @top-right { content: none; }
  @bottom-center { content: none; }
}

* { box-sizing: border-box; }
body {
  font-family: var(--ow-font-body);
  font-size: 10.5pt;
  line-height: 1.6;
  color: var(--ow-text);
  margin: 0;
  padding: 0;
}

.book-title-meta {
  string-set: book-title content();
  display: none;
}

/* Portada editorial profesional */
.cover-page {
  page-break-before: avoid;
  page-break-after: always;
  break-after: page;
  height: 100vh;
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding: 4cm 2.5cm;
  background: linear-gradient(135deg, var(--ow-bg) 0%, var(--ow-row-alt) 100%);
  border-left: 14pt solid var(--ow-primary);
}

.cover-kicker {
  text-transform: uppercase;
  letter-spacing: 0.15em;
  font-size: 11pt;
  font-weight: 700;
  color: var(--ow-accent);
  margin-bottom: 12pt;
}

.cover-title {
  font-family: var(--ow-font-title);
  font-size: 32pt;
  font-weight: 800;
  line-height: 1.15;
  color: var(--ow-primary);
  margin: 0 0 16pt 0;
}

.cover-author {
  font-size: 14pt;
  color: var(--ow-text);
  font-weight: 500;
  margin-top: 8pt;
}

.cover-meta {
  margin-top: auto;
  font-size: 9pt;
  color: var(--ow-muted);
  border-top: 1pt solid var(--ow-muted);
  padding-top: 12pt;
}

/* Tabla de Contenidos (TOC) */
.toc-page {
  page-break-before: always;
  break-before: page;
  padding-top: 1cm;
}

.toc-heading {
  font-family: var(--ow-font-title);
  font-size: 20pt;
  color: var(--ow-primary);
  border-bottom: 2pt solid var(--ow-primary);
  padding-bottom: 6pt;
  margin-bottom: 20pt;
}

.toc-list {
  list-style: none;
  padding: 0;
  margin: 0;
}

.toc-item {
  margin-bottom: 10pt;
  font-size: 11pt;
}

.toc-item a {
  text-decoration: none;
  color: var(--ow-text);
  display: block;
}

.toc-item a::after {
  content: leader('.') " " target-counter(attr(href), page);
  font-weight: 700;
  color: var(--ow-primary);
  float: right;
}

/* Capítulos */
.chapter-container {
  page-break-before: always;
  break-before: page;
  padding-top: 1cm;
}

.chapter-header {
  margin-bottom: 24pt;
  border-bottom: 1.5pt solid var(--ow-accent);
  padding-bottom: 8pt;
}

.chapter-number {
  font-size: 11pt;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  font-weight: 700;
  color: var(--ow-accent);
  margin-bottom: 4pt;
}

.chapter-title {
  font-family: var(--ow-font-title);
  font-size: 22pt;
  font-weight: 700;
  color: var(--ow-primary);
  margin: 0;
}

.chapter-body {
  font-size: 10.5pt;
  line-height: 1.65;
  color: var(--ow-text);
}

.chapter-body h3 {
  font-family: var(--ow-font-title);
  font-size: 13pt;
  color: var(--ow-primary);
  margin-top: 18pt;
  margin-bottom: 6pt;
}

.chapter-body p {
  margin-top: 0;
  margin-bottom: 10pt;
  orphans: 3;
  widows: 3;
}

.chapter-body table {
  width: 100%;
  border-collapse: collapse;
  margin: 14pt 0;
  font-size: 9.5pt;
}

.chapter-body th {
  background: var(--ow-primary);
  color: #fff;
  padding: 6pt 8pt;
  text-align: left;
}

.chapter-body td {
  padding: 5pt 8pt;
  border-bottom: 0.5pt solid #c8d4e0;
}

.chapter-body tr:nth-child(even) td {
  background: var(--ow-row-alt);
}
"""


def create_book(
    out_path: str,
    title: str,
    author: str = "",
    chapters: list[dict[str, Any]] | None = None,
    theme: str | None = None,
    epub: bool = False,
) -> dict[str, Any]:
    """Generates a professional multi-chapter PDF book or report with cover, automatic TOC with page numbers,

    and numbered chapters via WeasyPrint, and optional EPUB export via ebooklib.

    - out_path: output path for the PDF (or EPUB if out_path ends with .epub).
    - title: main title of the publication.
    - author: author or organization name.
    - chapters: list of dicts [{"title": "Chapter Name", "content_html": "<p>...</p>"}, ...]
    - theme: corporate theme name or dict (e.g. 'aden', 'corporate-blue', 'minimal').
    - epub: if True, also exports a validated .epub alongside the PDF.

    Returns dict with status, path, bytes, chapters_count, and epub_path (if generated).
    """
    from weasyprint import HTML

    if not title:
        raise ValueError("El título del libro es obligatorio (title no puede estar vacío).")

    chapters_list = list(chapters or [])
    if not chapters_list:
        raise ValueError("Debe incluir al menos un capítulo en 'chapters' ([{'title': ..., 'content_html': ...}]).")

    th = load_theme(theme)
    out_path_clean = os.path.abspath(os.path.expanduser(str(out_path)))
    target_pdf = safe_out(out_path_clean)
    is_epub_only = target_pdf.lower().endswith(".epub")

    epub_path = None
    if is_epub_only:
        epub_path = target_pdf
        target_pdf = str(Path(target_pdf).with_suffix(".pdf"))
    elif epub:
        epub_path = safe_out(str(Path(target_pdf).with_suffix(".epub")))

    # 1. Construir HTML para WeasyPrint
    escaped_title = html.escape(title)
    escaped_author = html.escape(author) if author else ""

    toc_items_html = []
    chapters_html = []

    for idx, chap in enumerate(chapters_list, start=1):
        chap_id = f"chap-{idx}"
        chap_title = html.escape(str(chap.get("title") or f"Capítulo {idx}"))
        chap_content = str(chap.get("content_html") or chap.get("content") or "")

        toc_items_html.append(
            f'<li class="toc-item"><a href="#{chap_id}">Capítulo {idx}: {chap_title}</a></li>'
        )

        chapters_html.append(f"""
<div id="{chap_id}" class="chapter-container">
  <div class="chapter-header">
    <div class="chapter-number">Capítulo {idx}</div>
    <h2 class="chapter-title">{chap_title}</h2>
  </div>
  <div class="chapter-body">
    {chap_content}
  </div>
</div>
""")

    toc_full_html = f"""
<div class="toc-page">
  <h2 class="toc-heading">Índice General</h2>
  <ul class="toc-list">
    {''.join(toc_items_html)}
  </ul>
</div>
"""

    author_markup = f'<div class="cover-author">{escaped_author}</div>' if escaped_author else ""

    full_html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>{escaped_title}</title>
<style>
{css_vars(th)}
{_BOOK_CSS}
</style>
</head>
<body>
<div class="book-title-meta">{escaped_title}</div>

<div class="cover-page">
  <div class="cover-kicker">Publicación Oficial</div>
  <h1 class="cover-title">{escaped_title}</h1>
  {author_markup}
  <div class="cover-meta">Generado con The Office Worker · Edición Profesional</div>
</div>

{toc_full_html}

{''.join(chapters_html)}
</body>
</html>"""

    # 2. Renderizar PDF con WeasyPrint
    HTML(string=full_html, url_fetcher=safe_url_fetcher).write_pdf(target_pdf)
    if not os.path.exists(target_pdf) or os.path.getsize(target_pdf) < 500:
        raise RuntimeError(f"Error al generar PDF de libro: tamaño inválido ({os.path.getsize(target_pdf) if os.path.exists(target_pdf) else 0}B)")

    warnings: list[str] = []

    # 3. Exportar EPUB opcional si se solicitó
    if epub_path:
        try:
            from ebooklib import epub as _epub
            book = _epub.EpubBook()
            book.set_identifier(f"urn:uuid:{uuid.uuid4()}")
            book.set_title(title)
            book.set_language("es")
            if author:
                book.add_author(author)

            epub_chapters = []
            for idx, chap in enumerate(chapters_list, start=1):
                c_title = str(chap.get("title") or f"Capítulo {idx}")
                c_content = str(chap.get("content_html") or chap.get("content") or "")
                c_item = _epub.EpubHtml(title=c_title, file_name=f"chap_{idx}.xhtml", lang="es")
                c_item.content = f"<h1>{html.escape(c_title)}</h1>{c_content}"
                book.add_item(c_item)
                epub_chapters.append(c_item)

            book.toc = tuple(epub_chapters)
            book.add_item(_epub.EpubNcx())
            book.add_item(_epub.EpubNav())
            book.spine = ["nav"] + epub_chapters

            _epub.write_epub(epub_path, book, {})

            if not os.path.exists(epub_path) or open(epub_path, "rb").read(2) != b"PK":
                warnings.append("EPUB generado no tiene los magic bytes estándar PK de archivo ZIP.")
        except ImportError:
            warnings.append("Exportación EPUB omitida: paquete 'ebooklib' no disponible en el entorno. Instalar con: pip install ebooklib")
            epub_path = None
        except Exception as exc:
            warnings.append(f"Fallo al exportar EPUB: {exc}")
            epub_path = None

    final_path = epub_path if is_epub_only and epub_path else target_pdf
    res: dict[str, Any] = {
        "status": "ok",
        "path": os.path.abspath(final_path),
        "bytes": os.path.getsize(final_path),
        "chapters_count": len(chapters_list),
        "title": title,
        "fidelity": "rich",
    }
    if epub_path and not is_epub_only:
        res["epub_path"] = os.path.abspath(epub_path)
        res["epub_bytes"] = os.path.getsize(epub_path)
    if warnings:
        res["warnings"] = warnings
    return res

from .security import safe_out, safe_url_fetcher, PROTECTED_SYSTEM_DIRS
from .themes import load_theme, DEFAULT_THEME, css_vars, Theme, THEMES, PREMIUM_CSS
from .templates import render_pdf
from .word import create_word, edit_word, mail_merge
from .excel import create_excel, edit_excel, csv_excel_convert
from .slides import create_pptx
from .pdf import read_pdf, extract_tables, list_form_fields, pdf_preview, pdf_extract_structured
from .pdf_tools import fill_pdf_form, ocr_pdf, convert_office_to_pdf, manipulate_pdf, sign_pdf, compress_pdf, pdf_redact
from .pdf_to_excel import pdf_to_excel
from .office_reader import read_office
from .templates_pack import list_packaged_templates, resolve_template_path, get_template_schema, TEMPLATES_CATALOG
from .diff import document_diff
from .compliance import scrub_metadata, protect_office, verify_pdf_signature
from .book import create_book
from .doctor import check_environment, detect_os, get_install_hint

__all__ = [
    "safe_out", "safe_url_fetcher", "PROTECTED_SYSTEM_DIRS",
    "load_theme", "DEFAULT_THEME", "css_vars", "Theme", "THEMES", "PREMIUM_CSS",
    "render_pdf", "create_word", "edit_word", "mail_merge", "create_excel", "edit_excel", "csv_excel_convert", "create_pptx",
    "read_pdf", "extract_tables", "list_form_fields", "pdf_preview", "pdf_extract_structured",
    "fill_pdf_form", "ocr_pdf", "convert_office_to_pdf", "manipulate_pdf",
    "sign_pdf", "compress_pdf", "pdf_redact", "pdf_to_excel", "read_office",
    "list_packaged_templates", "resolve_template_path", "get_template_schema", "TEMPLATES_CATALOG",
    "document_diff", "scrub_metadata", "protect_office", "verify_pdf_signature",
    "create_book", "check_environment", "detect_os", "get_install_hint",
]


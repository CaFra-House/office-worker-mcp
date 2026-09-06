"""Environment doctor: audits host system binaries and Python libraries for office-worker capabilities."""
from __future__ import annotations
import glob
import importlib.util
import os
import platform
import re
import shutil
from typing import Any


def detect_os() -> tuple[str, str]:
    """Detects host operating system and preferred package manager.

    Returns:
        tuple[os_name, package_manager] e.g. ('linux', 'apt'), ('macos', 'brew'), ('windows', 'winget')
    """
    sys_name = platform.system().lower()
    if sys_name == "darwin":
        return "macos", "brew"
    elif sys_name == "windows":
        return "windows", "winget"
    elif sys_name == "linux":
        if os.path.exists("/etc/os-release"):
            try:
                with open("/etc/os-release", "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read().lower()
                if any(distro in content for distro in ("fedora", "rhel", "centos", "rocky", "alma")):
                    return "linux", "dnf"
                elif "arch" in content or "manjaro" in content:
                    return "linux", "pacman"
                elif "alpine" in content:
                    return "linux", "apk"
            except Exception:
                pass
        return "linux", "apt"
    return "unknown", "unknown"


def _check_binary(name: str) -> bool:
    """Checks if a binary executable exists in PATH."""
    return shutil.which(name) is not None


def find_libreoffice_binary() -> str | None:
    """Busca el ejecutable de LibreOffice (soffice) en PATH y rutas estándar por plataforma.

    Prueba en orden:
    1) shutil.which("soffice")
    2) shutil.which("libreoffice")
    3) Rutas estándar por plataforma (Windows, macOS, Linux).

    Returns:
        str | None: Ruta absoluta al binario o None si no se encuentra.
    """
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return os.path.abspath(found)

    sys_name = platform.system().lower()
    candidates: list[str] = []

    if sys_name == "windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        pf = os.environ.get("ProgramFiles")
        if pf:
            cand = os.path.join(pf, "LibreOffice", "program", "soffice.exe")
            if cand not in candidates:
                candidates.insert(0, cand)
        pfx86 = os.environ.get("ProgramFiles(x86)")
        if pfx86:
            cand = os.path.join(pfx86, "LibreOffice", "program", "soffice.exe")
            if cand not in candidates:
                candidates.append(cand)
    elif sys_name == "darwin":
        candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    else:
        posix = ["/usr/bin/soffice", "/usr/local/bin/soffice"]
        if sys_name == "linux":
            posix.append("/snap/bin/libreoffice")
            opt_cands = glob.glob("/opt/libreoffice*/program/soffice")
            posix.extend(
                sorted(
                    opt_cands,
                    key=lambda p: [int(n) for n in re.findall(r"\d+", p)] or [0],
                    reverse=True,
                )
            )
        candidates = posix

    for cand in candidates:
        if os.path.isfile(cand) and os.access(cand, os.X_OK):
            return os.path.abspath(cand)

    return None


def _check_python_module(module_name: str) -> bool:
    """Checks if a Python module is importable without heavy execution."""
    try:
        spec = importlib.util.find_spec(module_name)
        return spec is not None
    except Exception:
        return False


_WEASYPROBE_CACHE: dict[str, Any] = {}


def _weasyprint_render_probe() -> tuple[bool, str]:
    """Real render probe: importing weasyprint can succeed on Windows while the native
    Pango/Cairo/GdkPixbuf shared libraries are missing — in that case any actual render
    fails. We do one minimal HTML→PDF render to memory (cached per process) so doctor
    never reports a false 'ACTIVO'.

    Returns:
        tuple[works: bool, error_detail: str]
    """
    if "result" in _WEASYPROBE_CACHE:
        return _WEASYPROBE_CACHE["result"]
    works, detail = False, ""
    try:
        import contextlib
        import io
        import sys

        err_stream = sys.stderr if sys.stderr is not None else io.StringIO()
        with contextlib.redirect_stdout(err_stream):
            from weasyprint import HTML

            HTML(string="<p>probe</p>").write_pdf(io.BytesIO())
        works, detail = True, ""
    except ImportError as e:
        works, detail = False, f"import error: {e}"
    except Exception as e:  # OSError/dlopen/cairo errors etc.
        works, detail = False, f"{type(e).__name__}: {e}"
    _WEASYPROBE_CACHE["result"] = (works, detail)
    return works, detail


def _render_document_install_hint(pkg_mgr: str, probe_error: str = "") -> str:
    """Install hint for render_document when the real render probe fails.

    If weasyprint imports but rendering fails (missing native libs), the pip package is
    already present — the hint must point at the C libraries or the Docker image, not at
    'pip install weasyprint' again.
    """
    if probe_error and "import error" not in probe_error.lower():
        hints = {
            "apt": "sudo apt update && sudo apt install -y libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0",
            "dnf": "sudo dnf install -y pango cairo gdk-pixbuf2",
            "pacman": "sudo pacman -S pango cairo gdk-pixbuf2",
            "apk": "apk add pango cairo gdk-pixbuf2",
            "brew": "brew install pango cairo gdk-pixbuf",
            "winget": ("WeasyPrint needs native Pango/Cairo libs with no unattended Windows installer — "
                       "use the official Docker image (docker pull ghcr.io/cafra-house/office-worker-mcp) or WSL2"),
        }
        base = hints.get(pkg_mgr, hints["apt"])
        return (base + f"  [probe failed: {probe_error}]")[:400]
    return get_install_hint("render_document", pkg_mgr)


def get_install_hint(cap_key: str, pkg_mgr: str) -> str:
    """Returns the exact installation command for a missing capability on the detected OS."""
    hints = {
        "convert_to_pdf": {
            "apt": "sudo apt update && sudo apt install -y libreoffice",
            "dnf": "sudo dnf install -y libreoffice",
            "pacman": "sudo pacman -S libreoffice-fresh",
            "apk": "apk add libreoffice",
            "brew": "brew install --cask libreoffice",
            "winget": "winget install TheDocumentFoundation.LibreOffice",
        },
        "pdf_ocr": {
            "apt": "sudo apt update && sudo apt install -y tesseract-ocr tesseract-ocr-spa",
            "dnf": "sudo dnf install -y tesseract tesseract-langpack-spa",
            "pacman": "sudo pacman -S tesseract tesseract-data-spa",
            "apk": "apk add tesseract-ocr tesseract-ocr-data-spa",
            "brew": "brew install tesseract tesseract-lang",
            "winget": "winget install UB-Mannheim.TesseractOCR",
        },
        "poppler": {
            "apt": "sudo apt update && sudo apt install -y poppler-utils",
            "dnf": "sudo dnf install -y poppler-utils",
            "pacman": "sudo pacman -S poppler",
            "apk": "apk add poppler-utils",
            "brew": "brew install poppler",
            "winget": "winget install osdn.poppler",
        },
        "render_document": {
            "apt": 'pip install "office-worker-mcp[pdf]"',
            "dnf": 'pip install "office-worker-mcp[pdf]"',
            "pacman": 'pip install "office-worker-mcp[pdf]"',
            "apk": 'pip install "office-worker-mcp[pdf]"',
            "brew": 'brew install pango cairo gdk-pixbuf && pip install "office-worker-mcp[pdf]"',
            "winget": 'pip install "office-worker-mcp[pdf]"',
        },
        "pandas": {
            "apt": "pip install pandas",
            "dnf": "pip install pandas",
            "pacman": "pip install pandas",
            "apk": "pip install pandas",
            "brew": "pip install pandas",
            "winget": "pip install pandas",
        },
        "ebooklib": {
            "apt": "pip install ebooklib",
            "dnf": "pip install ebooklib",
            "pacman": "pip install ebooklib",
            "apk": "pip install ebooklib",
            "brew": "pip install ebooklib",
            "winget": "pip install ebooklib",
        },
        "python_core": {
            "apt": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
            "dnf": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
            "pacman": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
            "apk": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
            "brew": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
            "winget": "pip install PyMuPDF pdfplumber openpyxl python-docx python-pptx docxtpl reportlab pypdf",
        },
    }
    cap_hints = hints.get(cap_key, {})
    return cap_hints.get(pkg_mgr, f"pip install {cap_key}")


def check_environment() -> dict[str, Any]:
    """Audits local system environment and reports active capabilities and missing binaries/packages.

    Returns:
        dict with status, os, package_manager, all_ready, and capabilities details.
    """
    os_name, pkg_mgr = detect_os()

    # System binaries
    has_soffice = find_libreoffice_binary() is not None
    has_tesseract = _check_binary("tesseract")
    has_poppler = _check_binary("pdftoppm")

    # Python packages
    has_weasyprint, weasy_error = _weasyprint_render_probe()  # real render probe (no false ACTIVO on Windows)
    has_fitz = _check_python_module("fitz")
    has_pdfplumber = _check_python_module("pdfplumber")
    has_openpyxl = _check_python_module("openpyxl")
    has_docx = _check_python_module("docx")
    has_pptx = _check_python_module("pptx")
    has_docxtpl = _check_python_module("docxtpl")
    has_pandas = _check_python_module("pandas")
    has_ebooklib = _check_python_module("ebooklib")
    has_pyhanko = _check_python_module("pyhanko")
    has_msoffcrypto = _check_python_module("msoffcrypto")

    capabilities = {
        "convert_to_pdf": {
            "active": has_soffice,
            "type": "binary",
            "name": "libreoffice (soffice)",
            "required_for": "Local Office to PDF conversion via headless LibreOffice",
            "install_hint": "" if has_soffice else get_install_hint("convert_to_pdf", pkg_mgr),
        },
        "pdf_ocr": {
            "active": has_tesseract,
            "type": "binary",
            "name": "tesseract",
            "required_for": "Optical character recognition on scanned PDFs/images",
            "install_hint": "" if has_tesseract else get_install_hint("pdf_ocr", pkg_mgr),
        },
        "poppler": {
            "active": has_poppler,
            "type": "binary (optional)",
            "name": "pdftoppm (poppler-utils)",
            "required_for": "Alternative PDF rasterization preview utility",
            "install_hint": "" if has_poppler else get_install_hint("poppler", pkg_mgr),
        },
        "render_document": {
            "active": has_weasyprint,
            "type": "python_library",
            "name": "weasyprint (+ native Pango/Cairo libs)",
            "required_for": "HTML/CSS to PDF rendering engine",
            "install_hint": "" if has_weasyprint else _render_document_install_hint(pkg_mgr, weasy_error),
        },
        "pdf_core": {
            "active": has_fitz,
            "type": "python_library",
            "name": "PyMuPDF (fitz)",
            "required_for": "Fast PDF text extraction, preview PNG rendering, redaction, compression",
            "install_hint": "" if has_fitz else "pip install PyMuPDF",
        },
        "pdf_to_excel": {
            "active": has_pdfplumber,
            "type": "python_library",
            "name": "pdfplumber",
            "required_for": "PDF table extraction into Excel workbooks",
            "install_hint": "" if has_pdfplumber else "pip install pdfplumber",
        },
        "excel_core": {
            "active": has_openpyxl,
            "type": "python_library",
            "name": "openpyxl",
            "required_for": "Reading and writing Excel .xlsx files",
            "install_hint": "" if has_openpyxl else "pip install openpyxl",
        },
        "word_core": {
            "active": has_docx,
            "type": "python_library",
            "name": "python-docx",
            "required_for": "Word .docx document generation and editing",
            "install_hint": "" if has_docx else "pip install python-docx",
        },
        "pptx_core": {
            "active": has_pptx,
            "type": "python_library",
            "name": "python-pptx",
            "required_for": "PowerPoint .pptx presentation generation and charts",
            "install_hint": "" if has_pptx else "pip install python-pptx",
        },
        "templates_word": {
            "active": has_docxtpl,
            "type": "python_library",
            "name": "docxtpl",
            "required_for": "Template rendering and mail merge in Word",
            "install_hint": "" if has_docxtpl else "pip install docxtpl",
        },
        "pandas_pivot": {
            "active": has_pandas,
            "type": "python_library",
            "name": "pandas",
            "required_for": "Excel pivot tables and tabular analysis",
            "install_hint": "" if has_pandas else get_install_hint("pandas", pkg_mgr),
        },
        "book_epub": {
            "active": has_ebooklib,
            "type": "python_library (optional)",
            "name": "ebooklib",
            "required_for": "EPUB export in create_book",
            "install_hint": "" if has_ebooklib else get_install_hint("ebooklib", pkg_mgr),
        },
        "pdf_signature": {
            "active": has_pyhanko,
            "type": "python_library",
            "name": "pyhanko",
            "required_for": "Cryptographic PAdES digital signature and verification",
            "install_hint": "" if has_pyhanko else "pip install pyhanko cryptography",
        },
        "office_protect": {
            "active": has_msoffcrypto,
            "type": "python_library",
            "name": "msoffcrypto-tool",
            "required_for": "Office document AES password encryption",
            "install_hint": "" if has_msoffcrypto else "pip install msoffcrypto-tool",
        },
    }

    critical_keys = ["render_document", "pdf_core", "excel_core", "word_core", "pandas_pivot"]
    core_ready = all(capabilities[k]["active"] for k in critical_keys)
    all_ready = all(v["active"] for v in capabilities.values() if "optional" not in v["type"])

    return {
        "status": "ok",
        "os": os_name,
        "package_manager": pkg_mgr,
        "core_ready": core_ready,
        "all_ready": all_ready,
        "capabilities": capabilities,
    }

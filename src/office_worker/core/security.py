"""Módulo de seguridad para The Office Worker.

Implementa:
- safe_out: validación de rutas de salida contra Path Traversal y sobreescritura del sistema.
- safe_url_fetcher: URL fetcher para WeasyPrint que bloquea SSRF remoto y lectura de archivos de sistema sensibles (LFI).
"""
from __future__ import annotations
import os
import urllib.parse
from pathlib import Path
from urllib.request import url2pathname

# Prefijos de directorios del sistema protegidos contra escritura arbitraria
PROTECTED_SYSTEM_DIRS = (
    "/etc",
    "/boot",
    "/sys",
    "/proc",
    "/dev",
    "/sbin",
    "/bin",
    "/usr",
    "/lib",
    "/lib64",
    "/root",
    "/var/run",
)


def safe_out(path: str | Path, base_dir: str | Path | None = None, ensure_parent: bool = True) -> str:
    """Valida y resuelve una ruta de salida segura contra Path Traversal y sobreescritura del sistema.

    - path: ruta de salida propuesta (relativa o absoluta).
    - base_dir: directorio base opcional al que restringir la salida. Si no se indica,
      se consulta la variable de entorno OFFICE_WORKER_ALLOWED_DIR o OFFICE_WORKER_BASE_DIR.
    - ensure_parent: si es True, crea automáticamente las carpetas intermedias.

    Lanza:
    - ValueError: si la ruta está vacía o es un directorio existente.
    - PermissionError: si la ruta apunta a directorios del sistema protegidos o fuera del base_dir permitido.
    Devuelve la ruta absoluta normalizada.
    """
    if not path or not str(path).strip():
        raise ValueError("La ruta de salida no puede estar vacía.")

    resolved = os.path.abspath(os.path.expanduser(str(path)))

    # Verificar si es un directorio existente
    if os.path.isdir(resolved):
        raise ValueError(f"La ruta de salida es un directorio existente, debe ser un archivo: {resolved}")

    # Verificar si se especificó o configuró un base_dir permitido
    allowed_base = base_dir or os.environ.get("OFFICE_WORKER_ALLOWED_DIR") or os.environ.get("OFFICE_WORKER_BASE_DIR")
    if allowed_base:
        resolved_base = os.path.abspath(os.path.expanduser(str(allowed_base)))
        try:
            Path(resolved).relative_to(resolved_base)
        except ValueError:
            raise PermissionError(
                f"Acceso denegado: la ruta de salida '{resolved}' está fuera del directorio permitido '{resolved_base}'"
            )

    # Bloqueo de rutas de sistema protegidas
    for p_dir in PROTECTED_SYSTEM_DIRS:
        if resolved == p_dir or resolved.startswith(p_dir + os.sep):
            raise PermissionError(f"Acceso denegado: ruta de salida en directorio de sistema protegido: {resolved}")

    if ensure_parent:
        parent = os.path.dirname(resolved)
        if parent:
            os.makedirs(parent, exist_ok=True)

    return resolved


def safe_url_fetcher(url: str, timeout: int = 10, ssl_context=None, http_headers=None):
    """Fetcher seguro para WeasyPrint contra SSRF remoto y lectura de archivos sensibles del sistema (LFI).

    - Bloquea esquemas de red remotos (http, https, ftp) para mitigar SSRF.
    - Permite esquema 'data:' para imágenes/fuentes inline en base64.
    - Para esquema 'file:', bloquea acceso a archivos del sistema protegidos (/etc, /proc, /sys, /root, /dev).
    - Si OFFICE_WORKER_ALLOWED_DIR / OFFICE_WORKER_BASE_DIR está activo, restringe file:// a dicho directorio.

    Limitación residual:
    En esquemas 'file://', se permite lectura de archivos locales no protegidos que el usuario del proceso
    pueda leer (p. ej. logos, plantillas, imágenes locales). Para un aislamiento estricto total,
    definir la variable OFFICE_WORKER_ALLOWED_DIR o ejecutar en un contenedor con permisos mínimos.
    """
    from weasyprint import default_url_fetcher

    parsed = urllib.parse.urlparse(url)
    scheme = parsed.scheme.lower()

    if scheme in ("http", "https", "ftp"):
        raise PermissionError(f"SSRF bloqueado: acceso a red remota denegado para URL '{url}'")

    if scheme == "file":
        filepath = os.path.abspath(os.path.expanduser(url2pathname(parsed.path)))
        # Verificar sistema protegido
        for p_dir in PROTECTED_SYSTEM_DIRS:
            if filepath == p_dir or filepath.startswith(p_dir + os.sep):
                raise PermissionError(f"LFI bloqueado: acceso denegado a recurso local protegido: '{filepath}'")

        allowed_base = os.environ.get("OFFICE_WORKER_ALLOWED_DIR") or os.environ.get("OFFICE_WORKER_BASE_DIR")
        if allowed_base:
            resolved_base = os.path.abspath(os.path.expanduser(str(allowed_base)))
            try:
                Path(filepath).relative_to(resolved_base)
            except ValueError:
                raise PermissionError(
                    f"Acceso denegado: recurso '{filepath}' fuera de directorio permitido '{resolved_base}'"
                )

    try:
        from weasyprint import URLFetcher
        return URLFetcher()(url)
    except (ImportError, AttributeError):
        from weasyprint import default_url_fetcher
        return default_url_fetcher(url, timeout=timeout, ssl_context=ssl_context, http_headers=http_headers)


def run_in_worker_thread_if_async(func, *args, **kwargs):
    """Ejecuta una función síncrona en un worker thread si se invoca dentro de un event loop activo.

    Bibliotecas como PyHanko invocan internamente `asyncio.run()`, lo que genera RuntimeError
    si el llamador ya se encuentra en un contexto asyncio (por ejemplo el servidor FastMCP o
    runners de tests asíncronos). Ejecutar en un ThreadPoolExecutor garantiza un entorno sin
    event loop activo donde `asyncio.run()` opera de forma segura.
    """
    import asyncio
    import concurrent.futures

    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = None

    if loop is not None and loop.is_running():
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
            return pool.submit(func, *args, **kwargs).result()
    return func(*args, **kwargs)


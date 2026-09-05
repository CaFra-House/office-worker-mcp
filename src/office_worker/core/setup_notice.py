"""Startup warning banner helper for CLI and MCP server.

Audits environment capabilities using doctor.py and emits a concise
banner to stderr when non-optional capabilities are missing.
"""
from __future__ import annotations
import sys
from typing import Any

from office_worker.core.doctor import check_environment

TOOL_BY_CAPABILITY: dict[str, str] = {
    "convert_to_pdf": "convert_to_pdf",
    "pdf_ocr": "pdf_ocr",
    "render_document": "render_document",
    "pdf_core": "read_pdf, pdf_preview, pdf_redact",
    "pdf_to_excel": "pdf_to_excel",
    "excel_core": "create_excel, edit_excel",
    "word_core": "create_word, edit_word",
    "pptx_core": "create_pptx, edit_pptx",
    "templates_word": "mail_merge",
    "pandas_pivot": "create_excel / edit_excel (pivot)",
    "pdf_signature": "sign_pdf, verify_pdf_signature",
    "office_protect": "protect_office",
}


def generate_setup_notice(env: dict[str, Any] | None = None) -> str | None:
    """Generates a short warning banner (max ~8 lines) detailing disabled tools,
    install commands, and the official Docker alternative.
    Returns None if all required capabilities are active.
    """
    if env is None:
        env = check_environment()

    if env.get("all_ready", False):
        return None

    capabilities = env.get("capabilities", {})
    inactive = [
        (k, v) for k, v in capabilities.items()
        if not v.get("active") and "optional" not in v.get("type", "")
    ]

    if not inactive:
        return None

    lines = [
        "[office-worker] Warning: Missing host capabilities for full document functionality:",
    ]
    max_items = 5
    for k, v in inactive[:max_items]:
        tool = TOOL_BY_CAPABILITY.get(k, k)
        hint = v.get("install_hint", "")
        hint_str = f" -> {hint}" if hint else ""
        name = v.get("name", k)
        lines.append(f"  • Tool '{tool}' disabled ({name}){hint_str}")

    if len(inactive) > max_items:
        lines.append(f"  • ... and {len(inactive) - max_items} more (run 'owi doctor' for full report)")

    lines.append("Frictionless full setup: docker pull ghcr.io/cafra-house/office-worker-mcp")
    return "\n".join(lines)


def print_setup_notice_if_needed(file: Any = None, env: dict[str, Any] | None = None) -> bool:
    """Prints warning banner to stderr (or specified file) if capabilities are missing.
    Returns True if banner was printed, False otherwise.
    """
    banner = generate_setup_notice(env)
    if not banner:
        return False
    target = sys.stderr if file is None else file
    print(banner, file=target)
    if hasattr(target, "flush"):
        try:
            target.flush()
        except Exception:
            pass
    return True

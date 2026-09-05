"""Tests for the startup setup notice (missing-capability warning banner)."""
from __future__ import annotations

import sys

import pytest

from office_worker.core import setup_notice


def _fake_env(inactive: dict[str, str]) -> dict:
    """Build a check_environment()-shaped dict with given capabilities inactive."""
    caps = {
        "convert_to_pdf": {
            "active": "convert_to_pdf" not in inactive,
            "type": "binary",
            "name": "libreoffice (soffice)",
            "install_hint": inactive.get("convert_to_pdf", ""),
        },
        "pdf_ocr": {
            "active": "pdf_ocr" not in inactive,
            "type": "binary",
            "name": "tesseract",
            "install_hint": inactive.get("pdf_ocr", ""),
        },
        "poppler": {
            "active": True,
            "type": "binary (optional)",
            "name": "pdftoppm (poppler-utils)",
            "install_hint": "",
        },
    }
    return {"status": "ok", "all_ready": not any(not v["active"] and "(optional)" not in v["type"] for v in caps.values()), "capabilities": caps}


def test_all_ready_returns_none():
    assert setup_notice.generate_setup_notice(_fake_env({})) is None


def test_missing_soffice_banner_mentions_tool_and_hint():
    env = _fake_env({"convert_to_pdf": "sudo apt update && sudo apt install -y libreoffice"})
    banner = setup_notice.generate_setup_notice(env)
    assert banner is not None
    assert "convert_to_pdf" in banner          # tool name affected
    assert "libreoffice" in banner             # missing binary named
    assert "apt install" in banner              # exact install command surfaced
    assert "docker pull ghcr.io/cafra-house/office-worker-mcp" in banner  # zero-friction path


def test_optional_capability_does_not_trigger_banner():
    # poppler is typed 'binary (optional)' — even if inactive it must not warn.
    env = _fake_env({})
    env["capabilities"]["poppler"]["active"] = False
    env["all_ready"] = True  # doctor marks optional gaps as non-blocking when core ready
    assert setup_notice.generate_setup_notice(env) is None


def test_print_goes_to_stderr_not_stdout(capfd):
    env = _fake_env({"pdf_ocr": "brew install tesseract"})
    printed = setup_notice.print_setup_notice_if_needed(env=env)
    out, err = capfd.readouterr()
    assert printed is True
    assert out == ""                            # stdout untouched (MCP JSON-RPC safe)
    assert "tesseract" in err                   # banner on stderr


def test_no_output_when_ready(capfd):
    printed = setup_notice.print_setup_notice_if_needed(env=_fake_env({}))
    out, err = capfd.readouterr()
    assert printed is False
    assert out == "" and err == ""

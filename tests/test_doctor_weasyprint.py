"""Tests for the real WeasyPrint render probe in doctor.py (no false 'ACTIVO' on Windows)."""
from __future__ import annotations

import pytest

from office_worker.core import doctor


@pytest.fixture(autouse=True)
def _clear_probe_cache():
    doctor._WEASYPROBE_CACHE.clear()
    yield
    doctor._WEASYPROBE_CACHE.clear()


def test_probe_ok_marks_capability_active(monkeypatch):
    monkeypatch.setattr(doctor, "_weasyprint_render_probe", lambda: (True, ""))
    env = doctor.check_environment()
    cap = env["capabilities"]["render_document"]
    assert cap["active"] is True
    assert cap["install_hint"] == ""


def test_probe_native_lib_failure_marks_inactive_with_pango_hint(monkeypatch):
    """Windows case: weasyprint imports but Pango/Cairo native libs are missing."""
    err = "OSError: cannot load library 'libpango-1.0-0': dlopen failed"
    monkeypatch.setattr(doctor, "_weasyprint_render_probe", lambda: (False, err))
    env = doctor.check_environment()
    cap = env["capabilities"]["render_document"]
    assert cap["active"] is False
    hint = cap["install_hint"]
    # must point at the C libraries (or Docker on winget), NOT 'pip install weasyprint' again
    if env["package_manager"] == "winget":
        assert "Docker" in hint or "WSL2" in hint
    else:
        assert "pango" in hint.lower()
        assert "pip install weasyprint" not in hint
    assert "probe failed" in hint  # honest error detail surfaced


def test_probe_import_error_falls_back_to_pip_hint(monkeypatch):
    monkeypatch.setattr(doctor, "_weasyprint_render_probe", lambda: (False, "import error: No module named 'weasyprint'"))
    env = doctor.check_environment()
    cap = env["capabilities"]["render_document"]
    assert cap["active"] is False
    assert "pip install weasyprint" in cap["install_hint"]


def test_all_ready_false_when_weasy_broken(monkeypatch):
    """setup_notice consumes all_ready — a broken render must flip it to False."""
    monkeypatch.setattr(doctor, "_weasyprint_render_probe", lambda: (False, "OSError: libpango missing"))
    env = doctor.check_environment()
    # only force everything else active if this host lacks something; check relative behavior instead:
    caps = env["capabilities"]
    assert caps["render_document"]["active"] is False
    non_optional_inactive = [k for k, v in caps.items() if "(optional)" not in v["type"] and not v["active"]]
    if non_optional_inactive == ["render_document"]:
        assert env["all_ready"] is False

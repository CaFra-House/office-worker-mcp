"""Tests unitarios para fixes de plataforma v0.9.8:
- FIX 1: Detección de LibreOffice fuera de PATH (rutas estándar por plataforma).
- FIX 2: Resolución de directorio de skills de Hermes por plataforma (Windows %LOCALAPPDATA% vs Unix ~/.hermes).
- FIX 3: Stdout purity en el probe de WeasyPrint (warnings redirigidos a stderr).
"""
from __future__ import annotations
import glob
import os
import platform
import sys
import types
from pathlib import Path
import pytest

from office_worker.core.doctor import find_libreoffice_binary, _weasyprint_render_probe, _WEASYPROBE_CACHE
from office_worker.core.pdf_tools import convert_office_to_pdf
from office_worker.skills import _default_hermes_skills_dir, install_skill


@pytest.fixture(autouse=True)
def _clear_doctor_cache():
    _WEASYPROBE_CACHE.clear()
    yield
    _WEASYPROBE_CACHE.clear()


# ============================================================================
# FIX 1: Detección de LibreOffice fuera de PATH
# ============================================================================

def test_find_libreoffice_binary_via_which(monkeypatch):
    """Si shutil.which encuentra soffice o libreoffice, se usa esa ruta."""
    monkeypatch.setattr("shutil.which", lambda cmd: "/custom/bin/soffice" if cmd == "soffice" else None)
    res = find_libreoffice_binary()
    assert res == "/custom/bin/soffice"


def test_find_libreoffice_binary_windows_standard_path(monkeypatch):
    """En Windows, si which() falla, detecta rutas estándar de LibreOffice en Program Files."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "Windows")

    expected = r"C:\Program Files\LibreOffice\program\soffice.exe"

    def fake_isfile(path):
        return path == expected

    monkeypatch.setattr(os.path, "isfile", fake_isfile)
    monkeypatch.setattr(os, "access", lambda path, mode: True)
    res = find_libreoffice_binary()
    assert res == os.path.abspath(expected)


def test_find_libreoffice_binary_macos_standard_path(monkeypatch):
    """En macOS, si which() falla, detecta la ruta estándar en /Applications."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "Darwin")

    expected = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    monkeypatch.setattr(os.path, "isfile", lambda path: path == expected)
    monkeypatch.setattr(os, "access", lambda path, mode: True)

    res = find_libreoffice_binary()
    assert res == os.path.abspath(expected)


def test_find_libreoffice_binary_linux_opt_glob(monkeypatch):
    """En Linux, si which() falla, detecta binarios en /opt/libreoffice*/program/soffice."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "Linux")
    expected = "/opt/libreoffice24.2/program/soffice"
    monkeypatch.setattr(glob, "glob", lambda pat: [expected] if "libreoffice" in pat else [])
    monkeypatch.setattr(os.path, "isfile", lambda path: path == expected)
    monkeypatch.setattr(os, "access", lambda path, mode: True)

    res = find_libreoffice_binary()
    assert res == os.path.abspath(expected)


def test_find_libreoffice_binary_returns_none_when_not_found(monkeypatch):
    """Si which falla y ninguna ruta estándar existe, retorna None."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(os.path, "isfile", lambda path: False)
    monkeypatch.setattr(glob, "glob", lambda pat: [])

    assert find_libreoffice_binary() is None


def test_convert_office_to_pdf_raises_when_libreoffice_missing(monkeypatch, tmp_path):
    """convert_office_to_pdf levanta FileNotFoundError con mensaje claro si LibreOffice no está."""
    monkeypatch.setattr("office_worker.core.pdf_tools.find_libreoffice_binary", lambda: None)
    dummy_in = tmp_path / "dummy.docx"
    dummy_in.write_text("fake docx", encoding="utf-8")
    dummy_out = tmp_path / "dummy.pdf"

    with pytest.raises(FileNotFoundError, match="LibreOffice \\(soffice\\) no está instalado o disponible en PATH"):
        convert_office_to_pdf(str(dummy_in), str(dummy_out))


# ============================================================================
# FIX 2: Directorio de skills de Hermes por plataforma
# ============================================================================

def test_default_hermes_skills_dir_windows(monkeypatch, tmp_path):
    """En Windows, resuelve a %LOCALAPPDATA%\\hermes\\skills."""
    fake_localappdata = str(tmp_path / "fake_localappdata")
    monkeypatch.setattr(platform, "system", lambda: "Windows")
    monkeypatch.setenv("LOCALAPPDATA", fake_localappdata)

    res = _default_hermes_skills_dir()
    expected = Path(fake_localappdata) / "hermes" / "skills"
    assert res == expected


def test_default_hermes_skills_dir_unix(monkeypatch, tmp_path):
    """En Linux y macOS, resuelve a ~/.hermes/skills."""
    fake_home = tmp_path / "fake_home"
    monkeypatch.setattr(platform, "system", lambda: "Linux")
    monkeypatch.setattr(Path, "home", lambda: fake_home)

    res = _default_hermes_skills_dir()
    expected = fake_home / ".hermes" / "skills"
    assert res == expected

    monkeypatch.setattr(platform, "system", lambda: "Darwin")
    res_mac = _default_hermes_skills_dir()
    assert res_mac == expected


def test_install_skill_windows_default(monkeypatch, tmp_path):
    """install_skill en Windows sin dest_dir instala en %LOCALAPPDATA%\\hermes\\skills."""
    fake_localappdata = tmp_path / "AppData" / "Local"
    fake_localappdata.mkdir(parents=True)
    monkeypatch.setattr(platform, "system", lambda: "Windows")
    monkeypatch.setenv("LOCALAPPDATA", str(fake_localappdata))

    res = install_skill("office-worker")
    assert res["status"] == "ok"
    installed_file = fake_localappdata / "hermes" / "skills" / "office-worker" / "SKILL.md"
    assert installed_file.exists()


def test_install_skill_explicit_dest_override(tmp_path):
    """install_skill respeta el parámetro dest_dir explícito en cualquier plataforma."""
    custom_dest = tmp_path / "custom_skills_folder"
    res = install_skill("office-worker", dest_dir=custom_dest)
    assert res["status"] == "ok"
    installed_file = custom_dest / "office-worker" / "SKILL.md"
    assert installed_file.exists()


# ============================================================================
# FIX 3: Stdout purity defensivo en probe WeasyPrint
# ============================================================================

def test_weasyprint_render_probe_stdout_purity(monkeypatch, capfd):
    """Cualquier print ejecutado durante el import/render del probe va a stderr, dejando stdout limpio."""
    class FakeHTML:
        def __init__(self, string=None):
            print("WARNING: C library loaded with minor quirks on stdout")

        def write_pdf(self, target):
            print("NOTICE: rendering in progress on stdout")
            target.write(b"%PDF-fake")

    fake_weasy = types.ModuleType("weasyprint")
    fake_weasy.HTML = FakeHTML
    monkeypatch.setitem(sys.modules, "weasyprint", fake_weasy)

    works, detail = _weasyprint_render_probe()
    assert works is True
    assert detail == ""

    out, err = capfd.readouterr()
    # Stdout DEBE estar completamente limpio
    assert out == "", f"stdout contaminado con: {out!r}"
    # Los prints deben haber ido a stderr
    assert "WARNING: C library loaded" in err
    assert "NOTICE: rendering in progress" in err


# ============================================================================
# AUDIT ROUND 2 FIXES: FIX B1 (CalVer, os.access, BSD), FIX B4 (stderr None), FIX A1 (pdf extra)
# ============================================================================

def test_find_libreoffice_binary_calver_ordering(monkeypatch):
    """CalVer numérico elige libreoffice24.2 sobre libreoffice7.6 (lexicográfico fallaría)."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "Linux")
    cands = [
        "/opt/libreoffice7.6/program/soffice",
        "/opt/libreoffice24.2/program/soffice",
    ]
    monkeypatch.setattr(glob, "glob", lambda pat: cands if "libreoffice" in pat else [])
    monkeypatch.setattr(os.path, "isfile", lambda path: path in cands)
    monkeypatch.setattr(os, "access", lambda path, mode: True)

    res = find_libreoffice_binary()
    assert res == "/opt/libreoffice24.2/program/soffice"


def test_find_libreoffice_binary_discarded_when_not_executable(monkeypatch):
    """Si un candidato existe pero os.access retorna False (no ejecutable), se descarta."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "Linux")
    expected = "/opt/libreoffice24.2/program/soffice"
    monkeypatch.setattr(glob, "glob", lambda pat: [expected] if "libreoffice" in pat else [])
    monkeypatch.setattr(os.path, "isfile", lambda path: path == expected)
    monkeypatch.setattr(os, "access", lambda path, mode: False)

    assert find_libreoffice_binary() is None


def test_find_libreoffice_binary_posix_bsd(monkeypatch):
    """En sistemas BSD/POSIX no Linux, detecta /usr/local/bin/soffice."""
    monkeypatch.setattr("shutil.which", lambda cmd: None)
    monkeypatch.setattr(platform, "system", lambda: "FreeBSD")
    expected = "/usr/local/bin/soffice"
    monkeypatch.setattr(os.path, "isfile", lambda path: path == expected)
    monkeypatch.setattr(os, "access", lambda path, mode: True)

    res = find_libreoffice_binary()
    assert res == expected


def test_weasyprint_render_probe_stderr_none_fallback(monkeypatch):
    """Si sys.stderr es None, el probe no falla y redirige stdout a StringIO."""
    class FakeHTML:
        def __init__(self, string=None):
            print("warning on stdout")

        def write_pdf(self, target):
            target.write(b"%PDF-fake")

    fake_weasy = types.ModuleType("weasyprint")
    fake_weasy.HTML = FakeHTML
    monkeypatch.setitem(sys.modules, "weasyprint", fake_weasy)
    monkeypatch.setattr(sys, "stderr", None)

    works, detail = _weasyprint_render_probe()
    assert works is True
    assert detail == ""


def test_pyproject_weasyprint_optional_pdf_extra():
    """Verifica que weasyprint no esté en dependencias base y sí en optional-dependencies[pdf]."""
    import tomllib
    pyproject_path = Path(__file__).resolve().parents[1] / "pyproject.toml"
    with open(pyproject_path, "rb") as f:
        data = tomllib.load(f)

    project = data["project"]
    base_deps = project.get("dependencies", [])
    optional_deps = project.get("optional-dependencies", {})
    classifiers = project.get("classifiers", [])

    assert not any("weasyprint" in dep for dep in base_deps), "weasyprint no debe estar en dependencies base"
    assert any("jinja2" in dep for dep in base_deps), "jinja2 debe permanecer en dependencies base"
    assert "pdf" in optional_deps, "extra [pdf] debe existir en optional-dependencies"
    assert any("weasyprint>=60" in dep for dep in optional_deps["pdf"]), "extra [pdf] debe requerir weasyprint>=60"
    assert "Operating System :: OS Independent" not in classifiers, "classifier OS Independent debe eliminarse"

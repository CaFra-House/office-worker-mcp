"""Packaged skills for The Office Worker.

Provides discovery and installation functions to copy official agent skills into
~/.hermes/skills/<slug>/SKILL.md idempotently.
"""
from __future__ import annotations
import os
import shutil
from pathlib import Path

SKILLS_DIR = Path(__file__).parent

SKILLS_CATALOG = {
    "office-worker": {
        "slug": "office-worker",
        "file": "SKILL.md",
        "title": "The Office Worker",
        "description": "Full document clerk workflows (PDF, Word, Excel, PowerPoint) with hard anti-loop rules.",
    },
    "google-drive-gmail": {
        "slug": "google-drive-gmail",
        "file": "google_drive_gmail.SKILL.md",
        "title": "Google Drive & Gmail Integration",
        "description": "Optional workflow to upload generated documents to Drive or draft Gmail via workspace-mcp.",
    },
}

# Alias map
ALIASES = {
    "default": "office-worker",
    "office_worker": "office-worker",
    "google": "google-drive-gmail",
    "google_drive_gmail": "google-drive-gmail",
}


def list_packaged_skills() -> list[dict]:
    """Retorna la lista de skills empaquetadas oficiales disponibles."""
    skills = []
    for key, info in SKILLS_CATALOG.items():
        src_file = SKILLS_DIR / info["file"]
        skills.append({
            "name": key,
            "slug": info["slug"],
            "title": info["title"],
            "description": info["description"],
            "exists": src_file.exists(),
            "path": str(src_file) if src_file.exists() else None,
        })
    return skills


def resolve_packaged_skill(name: str = "office-worker") -> Path:
    """Resuelve y verifica la ruta física a una skill empaquetada.

    Lanza FileNotFoundError si la skill no existe en el catálogo o en disco.
    """
    key = ALIASES.get(name.lower().strip(), name.lower().strip())
    if key not in SKILLS_CATALOG:
        valid = list(SKILLS_CATALOG.keys()) + list(ALIASES.keys())
        raise KeyError(f"Skill desconocida '{name}'. Skills válidas: {valid}")

    info = SKILLS_CATALOG[key]
    file_path = SKILLS_DIR / info["file"]
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo de skill empaquetada no encontrado: {file_path}")
    return file_path


def install_skill(name: str = "office-worker", dest_dir: str | Path | None = None) -> dict:
    """Copia la skill empaquetada a ~/.hermes/skills/<slug>/SKILL.md de forma idempotente.

    - name: 'office-worker', 'google-drive-gmail', o 'all'.
    - dest_dir: directorio raíz de destino (por defecto ~/.hermes/skills).

    Devuelve dict con status y ruta(s) de instalación.
    """
    base_target = Path(dest_dir) if dest_dir else Path(os.path.expanduser("~/.hermes/skills"))

    if name.lower().strip() == "all":
        installed = []
        for s_name in SKILLS_CATALOG:
            res = install_skill(s_name, dest_dir=base_target)
            installed.append(res)
        return {"status": "ok", "installed": installed}

    key = ALIASES.get(name.lower().strip(), name.lower().strip())
    src_file = resolve_packaged_skill(key)
    info = SKILLS_CATALOG[key]

    target_dir = base_target / info["slug"]
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / "SKILL.md"

    shutil.copy2(src_file, target_file)

    return {
        "status": "ok",
        "skill": key,
        "slug": info["slug"],
        "installed_path": str(target_file.resolve()),
        "bytes": target_file.stat().st_size,
    }

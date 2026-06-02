# -*- coding: utf-8 -*-
"""Управление конфигурацией проектов."""
import json
from pathlib import Path

ROOT = Path(__file__).parent.parent
PROJECTS_DIR = ROOT / "projects"
REPORTS_DIR = ROOT / "reports"

PROJECTS_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)


def load_project(name: str) -> dict | None:
    path = PROJECTS_DIR / f"{name}.json"
    if not path.exists():
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_project(name: str, data: dict) -> Path:
    path = PROJECTS_DIR / f"{name}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path

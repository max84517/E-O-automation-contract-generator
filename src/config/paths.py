"""Centralised path constants for the E&O Contract Generator.

All paths are derived from the project root so the project is fully
portable — no hard-coded absolute paths anywhere else in the codebase.

Layout expected at runtime
--------------------------
project_root/
├── data/
│   └── E&O summary table.xlsx
├── template/
│   ├── Settlement Form_Draft.docx
│   └── Settlement Form_Draft_Elan & SYNA.docx
├── output/
└── src/
    └── config/
        └── paths.py   ← this file
"""
from __future__ import annotations

import sys
from pathlib import Path


def _get_root() -> Path:
    """Return the project root regardless of execution context.

    * Normal Python run  : root = src/config/../../..  (three levels up)
    * PyInstaller bundle : root = directory containing the .exe
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent


ROOT: Path = _get_root()

DATA_DIR: Path = ROOT / "data"
TEMPLATE_DIR: Path = ROOT / "template"
OUTPUT_DIR: Path = ROOT / "output"

EXCEL_FILE: Path = DATA_DIR / "E&O summary table.xlsx"
TEMPLATE_GENERAL: Path = TEMPLATE_DIR / "Settlement Form_Draft.docx"
TEMPLATE_ELAN_SYNA: Path = TEMPLATE_DIR / "Settlement Form_Draft_Elan & SYNA.docx"

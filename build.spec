# -*- mode: python ; coding: utf-8 -*-
"""Bootstrap PyInstaller spec for the current Phase 0 runtime."""

from __future__ import annotations

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


ROOT = Path(__file__).resolve().parent


def existing_data_dir(relative_path: str) -> tuple[str, str] | None:
    path = ROOT / relative_path
    if not path.exists():
        return None
    return (str(path), relative_path)


datas = []
for relative_path in (
    "mapping_tables",
    # Future phases may add versioned runtime assets here.
    "templates",
    "themes",
):
    entry = existing_data_dir(relative_path)
    if entry is not None:
        datas.append(entry)

datas.extend(collect_data_files("customtkinter"))

hiddenimports = [
    "customtkinter",
    "openpyxl.cell._writer",
]

excludes = [
    "xlwings",
    "pytest",
    "pip",
    "setuptools",
]


a = Analysis(
    ["main.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SpreadAutomation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="SpreadAutomation",
)

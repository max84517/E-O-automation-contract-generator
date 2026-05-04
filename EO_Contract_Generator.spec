# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for E&O Settlement Contract Generator.

Produces a onedir build.  Deploy the entire dist/EO_Contract_Generator/
folder together with data/ and template/ placed next to the .exe:

    EO_Contract_Generator/
    ├── EO_Contract_Generator.exe   ← launcher
    ├── _internal/                  ← runtime libs (keep with exe)
    ├── data/
    │   └── E&O summary table.xlsx
    ├── template/
    │   ├── Settlement Form_Draft.docx
    │   └── Settlement Form_Draft_Elan & SYNA.docx
    └── output/                     ← auto-created on first run
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # customtkinter ships JSON/image assets that must be bundled
        *collect_data_files('customtkinter'),
    ],
    hiddenimports=[
        *collect_submodules('customtkinter'),
        'openpyxl',
        'openpyxl.cell._writer',
        'num2words',
        'docx',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EO_Contract_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # no black terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EO_Contract_Generator',
)

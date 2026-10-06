# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect playwright driver node.exe, cli.js, and package contents
playwright_datas = collect_data_files('playwright')

datas = playwright_datas + [
    ('Ananya Bill.xlsm', '.'),
]

hiddenimports = [
    'PyQt6',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
    'reportlab',
    'reportlab.lib',
    'reportlab.platypus',
    'reportlab.pdfgen',
    'reportlab.lib.colors',
    'reportlab.lib.pagesizes',
    'reportlab.lib.styles',
    'openpyxl',
    'openpyxl.reader.excel',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.styles',
    'openpyxl.cell',
    'playwright',
    'playwright.sync_api',
    'state_manager',
    'contacts_manager',
    'excel_manager',
    'pdf_generator',
    'whatsapp_automator',
    'whatsapp_dispatcher',
] + collect_submodules('reportlab')

a = Analysis(
    ['src/app_gui.py'],
    pathex=['src'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyinstaller_runtime_hook.py'],
    excludes=['numpy', 'pandas', 'scipy', 'matplotlib', 'tkinter', 'torch', 'transformers', 'IPython', 'pytest'],
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
    name='AnanyaInvoiceAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
    name='AnanyaInvoiceAutomation',
)

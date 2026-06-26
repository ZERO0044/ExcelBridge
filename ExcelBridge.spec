# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for ExcelBridge
# Build: pyinstaller ExcelBridge.spec

import sys
from pathlib import Path

block_cipher = None

# Collect customtkinter data and icon files
added_files = [('_icons', '_icons')]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'customtkinter',
        'openpyxl',
        'xlrd',
        'excel_reader',
        'excel_writer',
        'rule_engine',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelBridge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # 不显示命令行窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',         # 应用图标
)

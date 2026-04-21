# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Windows build

import os
import sys

block_cipher = None

# 取得專案根目錄
proj_dir = os.path.dirname(os.path.abspath(SPECPATH))

# 檢查圖示檔案是否存在
icon_file = os.path.join(proj_dir, 'app_icon.ico')
if not os.path.exists(icon_file):
    print(f"WARNING: Icon file not found: {icon_file}")
    icon_file = None

a = Analysis(
    ['SearchingPro.py'],
    pathex=[proj_dir],
    binaries=[],
    datas=[('app_icon.ico', '.')] if os.path.exists(icon_file) else [],
    hiddenimports=[
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'psutil',
        'sqlite3',
        'concurrent.futures',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
        'PyPDF2',
        'docx',
        'openpyxl',
        'pptx',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SearchingPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file if icon_file else None,
)

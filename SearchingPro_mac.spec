# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for macOS build

import os

block_cipher = None

proj_dir = os.path.dirname(os.path.abspath(SPECPATH))

a = Analysis(
    ['SearchingPro.py'],
    pathex=[proj_dir],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'psutil',
        'sqlite3',
        'concurrent.futures',
        'datetime',
        'json',
        'os',
        're',
        'sys',
        'time',
        'shutil',
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
    name='SearchingPro',
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
    name='SearchingPro',
)

app = BUNDLE(
    coll,
    name='SearchingPro.app',
    icon='SearchingPro.icns',
    bundle_identifier='com.searchingpro.app',
    info_plist={
        'CFBundleName': 'SearchingPro',
        'CFBundleDisplayName': 'Searching Pro',
        'CFBundleGetInfoString': 'Advanced File Searcher',
        'CFBundleIdentifier': 'com.searchingpro.app',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': 'True',
        'NSRequiresAquaSystemAppearance': 'False',
        'LSMinimumSystemVersion': '10.13.0',
        'NSHumanReadableCopyright': 'Copyright © 2024 SearchingPro. All rights reserved.',
        'CFBundleDocumentTypes': [],
        'LSApplicationCategoryType': 'public.app-category.utilities',
    },
)

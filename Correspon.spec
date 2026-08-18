# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

block_cipher = None

hidden_imports = [
    'uvicorn',
    'uvicorn.logging',
    'uvicorn.loops',
    'uvicorn.loops.auto',
    'uvicorn.protocols',
    'uvicorn.protocols.http',
    'uvicorn.protocols.http.auto',
    'uvicorn.lifespan',
    'uvicorn.lifespan.on',
    'fastapi',
    'starlette',
    'pydantic',
    'webview',
    'clr_loader',
    'pythonnet',
    'docxtpl',
    'docx',
    'mammoth',
    'pandas',
    'openpyxl',
    'jinja2',
    'jinja2.ext',
    'pythoncom',
    'win32com',
    'win32com.client',
    'win32com.client.dynamic',
    'win32com.client.gencache',
]

excludes = [
    'pandas.tests',
    'scipy',
    'sklearn',
    'matplotlib',
    'tkinter',
    'IPython',
    'notebook',
    'sphinx',
    'pytest',
]

datas = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('sample_data', 'sample_data'),
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name='Correspon',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Correspon',
)

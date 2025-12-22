# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for CodeExtractPro v2.0 - Release Build (No Console)
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect customtkinter data files
datas = collect_data_files('customtkinter')

# Hidden imports
hiddenimports = [
    'customtkinter',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'PIL._tkinter_finder',
    'ast',
    'json',
    'csv',
    'html',
    'threading',
    'webbrowser',
    'concurrent.futures',
    'dataclasses',
    'enum',
    'pathlib',
    'typing',
]

# Add oletools if available
try:
    import oletools
    hiddenimports.extend(collect_submodules('oletools'))
except ImportError:
    pass

# Add win32com if available (Windows only)
if sys.platform == 'win32':
    try:
        import win32com
        hiddenimports.extend(['win32com', 'win32com.client', 'pythoncom', 'pywintypes'])
    except ImportError:
        pass

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'IPython', 'jupyter'],
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
    name='CodeExtractPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico' if os.path.exists('assets/icon.ico') else None,
    version='version_info.txt' if os.path.exists('version_info.txt') else None,
)

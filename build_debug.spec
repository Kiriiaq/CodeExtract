# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for CodeExtractPro v1.0 - Debug Build (With Console)
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
    # Application modules
    'src',
    'src.core',
    'src.core.config_manager',
    'src.core.export_manager',
    'src.core.logging_system',
    'src.core.workflow',
    'src.modules',
    'src.modules.vba_extractor',
    'src.modules.python_analyzer',
    'src.modules.folder_scanner',
    'src.modules.vba_optimizer',
    'src.modules.report_generator',
    'src.ui',
    'src.ui.main_window',
    'src.utils',
    'src.utils.widgets',
    'src.utils.helpers',
    'core',
    'core.config_manager',
    'core.export_manager',
    'core.logging_system',
    'modules',
    'modules.vba_extractor',
    'modules.python_analyzer',
    'modules.folder_scanner',
    'modules.vba_optimizer',
    'utils',
    'utils.widgets',
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
    pathex=['.', 'src'],
    binaries=[],
    datas=datas + [('src', 'src'), ('assets', 'assets')],
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
    name='CodeExtractPro_Debug',
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # No compression for easier debugging
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Console window for debug output
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico' if os.path.exists('assets/icon.ico') else None,
)

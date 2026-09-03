# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

dnd_datas = collect_data_files('tkinterdnd2')


a = Analysis(
    ['app_automatizacion_optimizada.py'],
    pathex=[],
    binaries=[],
    datas=[('ultraball.ico', '.')] + dnd_datas,
    hiddenimports=['pythoncom', 'pywintypes', 'win32com', 'win32com.client', 'tkinterdnd2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='app_automatizacion_optimizada',
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
    icon=['ultraball.ico'],
)

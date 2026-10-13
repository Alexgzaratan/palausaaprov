# -*- mode: python ; coding: utf-8 -*-
# Lanzador mínimo que ocupa la ruta histórica dist/app_automatizacion_optimizada.exe y abre el programa real.
# Compilar:  python -m PyInstaller --noconfirm --workpath build_lanzador lanzador.spec
a = Analysis(
    ['lanzador.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'unittest', 'email', 'http', 'xml', 'pydoc_data'],
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
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon=['ultraball.ico'],
)

# -*- mode: python ; coding: utf-8 -*-
# Por defecto genera una carpeta (dist/programa/), que arranca en 1-3 s. El nombre es corto a propósito:
# con uno largo, alguna ruta interna de Playwright supera los 260 caracteres de Windows.
# Con la variable de entorno PALAUSA_ONEFILE=1 genera un único .exe (arranca en ~10 s).
# El lanzador (lanzador.spec) mantiene en dist/ el .exe histórico que usan los accesos directos y la tarea programada.
import os

from PyInstaller.utils.hooks import collect_data_files

dnd_datas = collect_data_files('tkinterdnd2')
SOLO_CARPETA = os.environ.get('PALAUSA_ONEFILE') != '1'


a = Analysis(
    ['app_automatizacion_optimizada.py'],
    pathex=[],
    binaries=[],
    datas=[('ultraball.ico', '.')] + dnd_datas,
    hiddenimports=['pythoncom', 'pywintypes', 'win32com', 'win32com.client', 'tkinterdnd2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Paquetes que pandas/otros arrastran pero la aplicación no usa (reducen el tamaño y el tiempo de arranque)
        'matplotlib', 'mpl_toolkits', 'contourpy', 'kiwisolver',
        'PySide6', 'shiboken6',
        'pytest', '_pytest', 'pluggy', 'iniconfig',
        'IPython', 'pygments',
        'setuptools', 'pkg_resources', '_distutils_hack',
        'xlsxwriter',
    ],
    noarchive=False,
    optimize=0,
)


def _es_chromium_completo(destino):
    # La aplicación solo lanza el navegador en modo headless (chromium_headless_shell).
    # El Chromium completo que trae Playwright (~330 MB) no se usa nunca.
    return '/.local-browsers/chromium-' in destino.replace('\\', '/')


a.binaries = [b for b in a.binaries if not _es_chromium_completo(b[0])]
a.datas = [d for d in a.datas if not _es_chromium_completo(d[0])]

pyz = PYZ(a.pure)

if SOLO_CARPETA:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='app_automatizacion_optimizada',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        icon=['ultraball.ico'],
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='programa')
else:
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

"""Cambio de versión en el disco: lo usa el lanzador antes de abrir el programa.

Estructura de la instalación:
    <carpeta>/app_automatizacion_optimizada.exe   lanzador
    <carpeta>/programa/                           programa (se sustituye entero al actualizar)
    <carpeta>/config_...json, logs, base de datos datos del usuario (nunca se tocan)

La app deja la versión nueva en programa_nuevo/ con la marca .listo y, si cambia, el lanzador nuevo
como app_automatizacion_optimizada.exe.nuevo. Solo usa la biblioteca estándar básica (el lanzador
se compila sin red ni correo)."""
import os
import shutil
import sys
import time
from pathlib import Path

NOMBRE_PROGRAMA = "programa"
NOMBRE_NUEVO = "programa_nuevo"
NOMBRE_ANTERIOR = "programa_anterior"
NOMBRE_EJECUTABLE = "app_automatizacion_optimizada.exe"
MARCA_LISTO = ".listo"
SUFIJO_NUEVO = ".nuevo"
SUFIJO_VIEJO = ".viejo"


def ruta_larga(ruta):
    """Ruta con prefijo \\\\?\\ en Windows: algunas rutas internas del navegador rozan los 260 caracteres."""
    ruta = os.path.abspath(str(ruta))
    if os.name != "nt" or ruta.startswith("\\\\?\\"):
        return ruta
    return "\\\\?\\" + ruta


def borrar_carpeta(ruta):
    if os.path.exists(ruta_larga(ruta)):
        shutil.rmtree(ruta_larga(ruta), ignore_errors=True)
    return not os.path.exists(ruta_larga(ruta))


def version_preparada(carpeta):
    """Versión descargada y lista para aplicarse al reiniciar, o None."""
    marca = Path(carpeta) / NOMBRE_NUEVO / MARCA_LISTO
    try:
        return marca.read_text(encoding="utf-8").strip() or None
    except OSError:
        return None


def aplicar_actualizacion(carpeta, espera_max=15.0, intervalo=0.5):
    """Sustituye programa/ por programa_nuevo/ (la anterior queda en programa_anterior/).
    Espera a que se cierre el programa si sigue abierto. Si algo falla deja la versión que había."""
    base = Path(carpeta)
    nuevo, actual, anterior = base / NOMBRE_NUEVO, base / NOMBRE_PROGRAMA, base / NOMBRE_ANTERIOR
    if version_preparada(base) is None:
        return False
    if not borrar_carpeta(anterior):
        return False
    if actual.exists():
        fin = time.monotonic() + espera_max
        while True:
            try:
                os.rename(actual, anterior)
                break
            except OSError:
                if time.monotonic() > fin:
                    return False  # el programa sigue abierto: se intentará en el próximo arranque
                time.sleep(intervalo)
    try:
        os.rename(nuevo, actual)
    except OSError:
        if anterior.exists() and not actual.exists():
            os.rename(anterior, actual)
        return False
    try:
        (actual / MARCA_LISTO).unlink()
    except OSError:
        pass
    return True


def actualizar_lanzador(carpeta, ejecutable_actual=None):
    """Pone el lanzador nuevo si viene en la actualización. Windows permite renombrar un .exe en uso,
    así que el lanzador puede sustituirse a sí mismo; el viejo se borra en el siguiente arranque."""
    base = Path(carpeta)
    ejecutable = Path(ejecutable_actual or base / NOMBRE_EJECUTABLE)
    viejo = ejecutable.with_name(ejecutable.name + SUFIJO_VIEJO)
    nuevo = ejecutable.with_name(ejecutable.name + SUFIJO_NUEVO)
    try:
        viejo.unlink()
    except OSError:
        pass
    if not nuevo.exists():
        return False
    try:
        if ejecutable.exists():
            os.rename(ejecutable, viejo)
        os.rename(nuevo, ejecutable)
        return True
    except OSError:
        if viejo.exists() and not ejecutable.exists():
            os.rename(viejo, ejecutable)
        return False


def al_arrancar_lanzador(carpeta):
    """Todo lo que hace el lanzador antes de abrir el programa; nunca impide que el programa se abra."""
    try:
        aplicar_actualizacion(carpeta)
    except Exception:
        pass
    try:
        if getattr(sys, "frozen", False):
            actualizar_lanzador(carpeta, sys.executable)
    except Exception:
        pass

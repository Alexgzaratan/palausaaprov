"""Rutas de datos de la aplicación, ancladas a la carpeta del programa y no al directorio de trabajo."""
import os
import shutil
import sys
from pathlib import Path


def carpeta_programa():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _es_escribible(carpeta):
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        prueba = carpeta / ".escritura_prueba"
        prueba.write_text("ok", encoding="utf-8")
        prueba.unlink()
        return True
    except OSError:
        return False


def _carpeta_datos():
    forzada = os.environ.get("PALAUSA_DATOS")
    if forzada:
        Path(forzada).mkdir(parents=True, exist_ok=True)
        return Path(forzada)
    base = carpeta_programa()
    if _es_escribible(base):
        return base
    alternativa = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "PalausaAprovisionamiento"
    alternativa.mkdir(parents=True, exist_ok=True)
    return alternativa


CARPETA_DATOS = _carpeta_datos()


def ruta_datos(nombre, migrar=True):
    """Ruta absoluta de un fichero o carpeta de datos junto al programa.

    Si todavía no existe ahí pero sí en el directorio de trabajo (versiones antiguas que
    guardaban en relativo), se copia una sola vez para no perder la configuración previa."""
    destino = CARPETA_DATOS / nombre
    if migrar and not destino.exists():
        antiguo = Path.cwd() / nombre
        try:
            if antiguo.exists() and antiguo.resolve() != destino.resolve():
                if antiguo.is_dir():
                    shutil.copytree(antiguo, destino)
                else:
                    shutil.copy2(antiguo, destino)
        except OSError:
            pass
    return str(destino)

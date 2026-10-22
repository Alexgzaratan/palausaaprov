"""Actualizaciones desde GitHub Releases: aviso de versión nueva, descarga y preparación.

La app descarga el zip de la última versión publicada (publicar_version.py) y lo deja en programa_nuevo/;
al reiniciar, el lanzador lo pone en su sitio (cambio_version.py). Los datos del usuario no se tocan."""
import json
import os
import re
import shutil
import sys
import urllib.request
import zipfile
from pathlib import Path

import cambio_version
from version import REPOSITORIO, VERSION

URL_ULTIMA_VERSION = f"https://api.github.com/repos/{REPOSITORIO}/releases/latest"
AGENTE = f"PalausaAprovisionamiento/{VERSION}"
CARPETA_DESCARGA = "_actualizacion"


def version_tupla(texto):
    numeros = [int(n) for n in re.findall(r"\d+", str(texto))[:3]]
    return tuple(numeros + [0] * (3 - len(numeros)))


def es_mas_nueva(candidata, actual=VERSION):
    return version_tupla(candidata) > version_tupla(actual)


def carpeta_instalacion():
    """Carpeta con el lanzador y programa/ si la app es una instalación actualizable; si no, None
    (por ejemplo al ejecutar el .py, en los tests o en un .exe de un solo archivo)."""
    if not getattr(sys, "frozen", False):
        return None
    carpeta_exe = Path(sys.executable).resolve().parent
    base = carpeta_exe.parent
    if carpeta_exe.name != cambio_version.NOMBRE_PROGRAMA or not (base / cambio_version.NOMBRE_EJECUTABLE).exists():
        return None
    return base


def _peticion(url):
    return urllib.request.Request(url, headers={"User-Agent": AGENTE, "Accept": "application/vnd.github+json"})


def buscar_actualizacion(actual=VERSION, abrir=urllib.request.urlopen, espera=6):
    """Datos de la última versión publicada si es más nueva que la actual; None si no hay ninguna."""
    with abrir(_peticion(URL_ULTIMA_VERSION), timeout=espera) as respuesta:
        datos = json.load(respuesta)
    version = str(datos.get("tag_name", "")).lstrip("vV")
    if not version or datos.get("draft") or datos.get("prerelease") or not es_mas_nueva(version, actual):
        return None
    paquete = next((a for a in datos.get("assets", []) if str(a.get("name", "")).lower().endswith(".zip")), None)
    if paquete is None:
        return None
    return {"version": version, "notas": (datos.get("body") or "").strip(), "url": paquete["browser_download_url"],
            "tamano": int(paquete.get("size") or 0), "nombre": paquete["name"]}


def descargar(url, destino, tamano_esperado=0, progreso=None, abrir=urllib.request.urlopen):
    """Descarga por partes a un .parcial y solo lo renombra si llega completo."""
    parcial = str(destino) + ".parcial"
    with abrir(_peticion(url), timeout=30) as respuesta, open(parcial, "wb") as archivo:
        total = tamano_esperado or int(respuesta.headers.get("Content-Length") or 0)
        leido = 0
        while True:
            bloque = respuesta.read(1 << 20)
            if not bloque:
                break
            archivo.write(bloque)
            leido += len(bloque)
            if progreso and total:
                progreso(min(leido / total, 1.0))
    if tamano_esperado and os.path.getsize(parcial) != tamano_esperado:
        os.remove(parcial)
        raise ValueError("la descarga llegó incompleta")
    os.replace(parcial, destino)


def preparar_desde_zip(ruta_zip, carpeta, version):
    """Extrae programa/ del zip a programa_nuevo/ y el lanzador a <lanzador>.nuevo.
    La marca .listo se escribe al final: sin ella el lanzador no aplica nada."""
    base = Path(carpeta)
    nuevo = base / cambio_version.NOMBRE_NUEVO
    if not cambio_version.borrar_carpeta(nuevo):
        raise OSError(f"no se pudo vaciar {nuevo}")
    prefijo = cambio_version.NOMBRE_PROGRAMA + "/"
    with zipfile.ZipFile(ruta_zip) as paquete:
        nombres = paquete.namelist()
        if prefijo + cambio_version.NOMBRE_EJECUTABLE not in nombres:
            raise ValueError("el paquete no contiene el programa")
        for info in paquete.infolist():
            nombre = info.filename
            if ".." in Path(nombre).parts or nombre.startswith(("/", "\\")) or ":" in nombre:
                raise ValueError(f"ruta no válida en el paquete: {nombre}")
            if nombre.startswith(prefijo) and nombre != prefijo:
                destino = nuevo / nombre[len(prefijo):]
            elif nombre == cambio_version.NOMBRE_EJECUTABLE:
                destino = base / (cambio_version.NOMBRE_EJECUTABLE + cambio_version.SUFIJO_NUEVO)
            else:
                continue
            if info.is_dir():
                os.makedirs(cambio_version.ruta_larga(destino), exist_ok=True)
                continue
            os.makedirs(cambio_version.ruta_larga(destino.parent), exist_ok=True)
            with paquete.open(info) as origen, open(cambio_version.ruta_larga(destino), "wb") as salida:
                shutil.copyfileobj(origen, salida, 1 << 20)
    (nuevo / cambio_version.MARCA_LISTO).write_text(version, encoding="utf-8")


def descargar_y_preparar(info, carpeta, progreso=None, abrir=urllib.request.urlopen):
    temporal = Path(carpeta) / CARPETA_DESCARGA
    temporal.mkdir(exist_ok=True)
    ruta_zip = temporal / info["nombre"]
    try:
        descargar(info["url"], ruta_zip, info.get("tamano", 0), progreso, abrir)
        preparar_desde_zip(ruta_zip, carpeta, info["version"])
    finally:
        cambio_version.borrar_carpeta(temporal)

"""Lectura/escritura segura de la configuración y registro de errores."""
import json
import os
import shutil
import threading
import traceback
from datetime import datetime

from rutas_app import ruta_datos

CONFIG_FILE = ruta_datos("config_palausa_aprovisionamiento.json")
LOG_DETALLADO = ruta_datos("error_log_detallado.txt", migrar=False)
LOG_MAX_BYTES = 2 * 1024 * 1024

_bloqueo_config = threading.Lock()


def registrar_error_detallado(error_msg, traceback_str, ref=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        if os.path.exists(LOG_DETALLADO) and os.path.getsize(LOG_DETALLADO) > LOG_MAX_BYTES:
            os.replace(LOG_DETALLADO, LOG_DETALLADO + ".1")
        with open(LOG_DETALLADO, "a", encoding="utf-8") as f:
            f.write(f"--- ERROR [{timestamp}] ---\n")
            if ref: f.write(f"Referencia: {ref}\n")
            f.write(f"Mensaje: {error_msg}\n")
            f.write(f"Traceback:\n{traceback_str}\n")
            f.write("-" * 30 + "\n\n")
    except OSError:
        pass


def _leer_json(ruta):
    with open(ruta, "r", encoding="utf-8") as f:
        datos = json.load(f)
    if not isinstance(datos, dict):
        raise ValueError("La configuración no es un objeto JSON")
    return datos


def cargar_configuracion():
    """Lee la configuración. Si el fichero está corrupto recupera la última copia buena (.bak)."""
    for ruta in (CONFIG_FILE, CONFIG_FILE + ".bak"):
        if os.path.exists(ruta):
            try:
                return _leer_json(ruta)
            except Exception as e:
                registrar_error_detallado(str(e), traceback.format_exc(), ref=f"cargar_configuracion:{os.path.basename(ruta)}")
    return {}


def guardar_configuracion(config_dict):
    """Escritura atómica: fichero temporal + reemplazo, conservando la versión anterior como .bak."""
    with _bloqueo_config:
        temporal = CONFIG_FILE + ".tmp"
        try:
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(config_dict, f, indent=4, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
            if os.path.exists(CONFIG_FILE):
                try:
                    _leer_json(CONFIG_FILE)
                    shutil.copy2(CONFIG_FILE, CONFIG_FILE + ".bak")
                except Exception:
                    pass
            os.replace(temporal, CONFIG_FILE)
        except Exception as e:
            registrar_error_detallado(str(e), traceback.format_exc(), ref="guardar_configuracion")
            try:
                os.remove(temporal)
            except OSError:
                pass

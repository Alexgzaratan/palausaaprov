import os
import sys
import tempfile
from pathlib import Path

import pytest

# Los tests nunca deben tocar la configuración ni los logs reales
os.environ.setdefault("PALAUSA_DATOS", tempfile.mkdtemp(prefix="palausa_tests_"))

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

CONFIG_BASE = {
    "usuario_madd": "DF1", "ruta_excel_mopar": "C:/mopar.xls", "email_companeros_avisos": "a@x.es; b@x.es",
    "precio_min": "150", "umbral_rotacion": "0.12", "num_hilos": 4, "avisos_envio_automatico": False,
    "modo_apariencia": "Light", "mopar_numero": 7,
}


@pytest.fixture(scope="session")
def ventana_unica(tmp_path_factory):
    # Tk no admite bien crear y destruir varias ventanas seguidas en un mismo proceso: se reutiliza una para toda la sesión
    import configuracion
    import app_automatizacion_optimizada as app_mod
    carpeta = tmp_path_factory.mktemp("cfg")
    original = configuracion.CONFIG_FILE
    configuracion.CONFIG_FILE = str(carpeta / "config.json")
    configuracion.guardar_configuracion(dict(CONFIG_BASE))
    ventana = app_mod.AppIntegrada()
    yield ventana
    ventana.destroy()
    configuracion.CONFIG_FILE = original


@pytest.fixture
def app(ventana_unica, monkeypatch):
    import configuracion
    import app_automatizacion_optimizada as app_mod
    configuracion.guardar_configuracion(dict(CONFIG_BASE))
    ventana_unica.config.clear()
    ventana_unica.config.update(configuracion.cargar_configuracion())
    ventana_unica.reconstruir_interfaz()
    ventana_unica.tabview.set("Automatización")
    ventana_unica.construir_pestana("Automatización")
    ventana_unica.actualizar_paneles_por_pestana()
    mensajes = []
    monkeypatch.setattr(app_mod.messagebox, "showwarning", lambda *a, **k: mensajes.append(a))
    ventana_unica.mensajes = mensajes
    return ventana_unica

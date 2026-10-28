"""Instancia única, código de salida de la tarea programada y reanudar una publicación cortada."""
import sys
import threading

import pytest

import app_automatizacion_optimizada as app_mod
import instancia_unica
import publicar_version

MUTEX_PRUEBA = "Local\\PalausaAprovisionamiento_prueba"


def en_otro_hilo(funcion):
    # El mutex pertenece al hilo que lo toma: otro hilo hace de «segunda instancia»
    resultado = []
    hilo = threading.Thread(target=lambda: resultado.append(funcion()))
    hilo.start()
    hilo.join()
    return resultado[0]


@pytest.mark.skipif(sys.platform != "win32", reason="mutex de Windows")
def test_segunda_instancia_no_arranca_y_queda_libre_al_cerrar():
    primera = instancia_unica.tomar(MUTEX_PRUEBA, espera_ms=0)
    assert primera not in (None, True)
    try:
        assert en_otro_hilo(lambda: instancia_unica.tomar(MUTEX_PRUEBA, espera_ms=100)) is None
    finally:
        instancia_unica.soltar(primera)

    def tomar_y_soltar():
        handle = instancia_unica.tomar(MUTEX_PRUEBA, espera_ms=100)
        instancia_unica.soltar(handle)
        return handle

    assert en_otro_hilo(tomar_y_soltar) not in (None, True)


@pytest.mark.skipif(sys.platform != "win32", reason="mutex de Windows")
def test_espera_a_que_se_cierre_la_anterior():
    # Reinicio tras actualizar: la versión nueva arranca mientras la vieja termina de cerrarse
    listo, soltar = threading.Event(), threading.Event()

    def anterior():
        handle = instancia_unica.tomar(MUTEX_PRUEBA, espera_ms=0)
        listo.set()
        soltar.wait(5)
        instancia_unica.soltar(handle)

    hilo = threading.Thread(target=anterior)
    hilo.start()
    listo.wait(5)
    threading.Timer(0.3, soltar.set).start()
    nueva = instancia_unica.tomar(MUTEX_PRUEBA, espera_ms=3000)
    hilo.join()
    assert nueva not in (None, True)
    instancia_unica.soltar(nueva)


CONFIG_MADD = {"usuario_madd": "u", "password_madd": "p", "ruta_carpeta_facturas_madd": "c", "email_facturas_madd": "e"}


@pytest.mark.parametrize("config,error,codigo", [
    (CONFIG_MADD, None, 0),
    (CONFIG_MADD, "no carga la web", 1),
    ({"usuario_madd": "u"}, None, 1),  # configuración incompleta
])
def test_codigo_de_salida_de_la_tarea_programada(monkeypatch, config, error, codigo):
    monkeypatch.setattr(app_mod, "cargar_configuracion", lambda: dict(config))
    monkeypatch.setattr(app_mod, "registrar_error_detallado", lambda *a, **k: None)

    def tarea(*a, resultado_callback=None, **k):
        resultado_callback({"facturas": [], "error": error})

    monkeypatch.setattr(app_mod, "tarea_facturas_madd", tarea)
    assert app_mod.ejecutar_tarea_madd_standalone() == codigo


@pytest.fixture
def publicacion(monkeypatch, tmp_path):
    ordenes = []
    monkeypatch.setattr(publicar_version, "SALIDA", tmp_path)
    monkeypatch.setattr(publicar_version, "ejecutar", lambda orden, **k: ordenes.append([str(p) for p in orden]))
    monkeypatch.setattr(publicar_version, "ruta_gh", lambda: "gh")
    return ordenes, tmp_path


def test_reanuda_si_falta_la_release(monkeypatch, publicacion):
    ordenes, carpeta = publicacion
    (carpeta / "Palausa_Aprovisionamiento_9.9.9.zip").write_bytes(b"zip")
    monkeypatch.setattr(publicar_version, "release_existe", lambda etiqueta: False)
    publicar_version.reanudar("9.9.9", "v9.9.9", "notas")
    assert ordenes[0] == ["git", "push", "origin", "main", "v9.9.9"]
    assert ordenes[1][:4] == ["gh", "release", "create", "v9.9.9"] and ordenes[1][4].endswith("Palausa_Aprovisionamiento_9.9.9.zip")


def test_no_reanuda_si_ya_esta_publicada_o_falta_el_paquete(monkeypatch, publicacion):
    ordenes, carpeta = publicacion
    monkeypatch.setattr(publicar_version, "release_existe", lambda etiqueta: True)
    with pytest.raises(SystemExit, match="ya está publicada"):
        publicar_version.reanudar("9.9.9", "v9.9.9", "notas")
    monkeypatch.setattr(publicar_version, "release_existe", lambda etiqueta: False)
    with pytest.raises(SystemExit, match="no el paquete"):
        publicar_version.reanudar("9.9.9", "v9.9.9", "notas")
    assert ordenes == []

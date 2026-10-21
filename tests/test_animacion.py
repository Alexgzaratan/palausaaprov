import time

import animacion_inicio
import configuracion


def test_fases_de_la_animacion():
    inicio = animacion_inicio.estado(0)
    assert inicio["escala"] < 1 and inicio["apertura"] == 0 and inicio["opacidad"] == 1
    assert animacion_inicio.estado(0.45)["angulo"] != 0  # se agita
    assert animacion_inicio.estado(0.6)["angulo"] == 0   # pausa entre sacudidas
    aperturas = [animacion_inicio.estado(t / 100)["apertura"] for t in range(80, 130)]
    assert aperturas == sorted(aperturas) and aperturas[-1] == 1
    final = animacion_inicio.estado(animacion_inicio.DURACION)
    assert final["opacidad"] == 0 and final["destello"] == 1


def test_mezclar_colores():
    assert animacion_inicio.mezclar("#000000", "#ffffff", 0) == "#000000"
    assert animacion_inicio.mezclar("#000000", "#ffffff", 1) == "#ffffff"
    assert animacion_inicio.mezclar("#000000", "#ffffff", 0.5) == "#808080"


def test_cada_fotograma_se_dibuja_con_el_tamano_fijo(app):
    bola = animacion_inicio.Ultraball(app.imagen_logo(), 120, "#1a1613", "#ffdf8d")
    for t in (0, 0.1, 0.45, 0.95, 1.1, 1.5):
        fotograma = bola.dibujar(animacion_inicio.estado(t))
        assert fotograma.size == (bola.lado, bola.lado)


def esperar_fin(app, animacion, limite=5):
    fin = time.time() + limite
    while not animacion.terminada and time.time() < fin:
        app.update()
        time.sleep(0.01)


def test_la_animacion_termina_y_deja_la_app_visible(app):
    terminada = []
    animacion = app.mostrar_animacion_inicio()
    animacion.al_terminar = lambda: terminada.append(True)
    assert animacion.capa.winfo_exists()
    esperar_fin(app, animacion)
    assert terminada == [True]
    assert not animacion.capa.winfo_exists()
    assert float(app.attributes("-alpha")) == 1.0


def test_un_clic_salta_la_animacion(app):
    animacion = app.mostrar_animacion_inicio()
    animacion.capa.event_generate("<Button-1>")
    app.update()
    assert animacion.terminada
    assert float(app.attributes("-alpha")) == 1.0


def test_si_falla_la_app_se_muestra_igualmente(app, monkeypatch):
    monkeypatch.setattr(animacion_inicio, "Ultraball", lambda *a, **k: 1 / 0)
    animacion = app.mostrar_animacion_inicio()
    assert animacion.terminada
    assert float(app.attributes("-alpha")) == 1.0


def test_se_puede_desactivar_en_configuracion(app):
    app.tabview.set("Configuración")
    app.al_cambiar_pestana()
    assert app.animacion_var.get() is True
    app.animacion_var.set(False)
    app.guardar_configuracion_ui_internal()
    assert configuracion.cargar_configuracion()["animacion_inicio"] is False
    assert app.mostrar_animacion_inicio() is None

"""Cabecera con degradado según la hora y check verde al terminar una tarea."""
from datetime import datetime

import efectos


def test_momento_del_dia_y_degradado():
    assert [efectos.momento_del_dia(datetime(2026, 11, 16, h)) for h in (7, 12, 18, 23, 3)] == ["amanecer", "mediodia", "tarde", "noche", "noche"]
    imagen = efectos.imagen_degradado(400, 20, efectos.DEGRADADOS["tarde"], 200)
    assert imagen.size == (400, 20)
    assert imagen.getpixel((399, 19)) == efectos._rgb(efectos.DEGRADADOS["tarde"][2])  # liso bajo las fichas
    assert imagen.getpixel((60, 19)) != imagen.getpixel((399, 19))


def test_la_cabecera_es_el_degradado_y_lleva_las_fichas(app):
    lienzo = app.cabecera_lienzo
    assert lienzo is not None and app.frame_resumen is lienzo.marco_fichas
    app.update()
    lienzo.repintar(forzar=True)
    assert lienzo.estado["foto"] is not None and lienzo.estado["corte"] < lienzo.estado["ancho"]
    assert app.ficha_informe.winfo_exists()
    # Si la cabecera se quedó con otro momento del día, al revisar la hora toma los colores del actual
    lienzo.estado["momento"] = "noche" if efectos.momento_del_dia() != "noche" else "tarde"
    lienzo.derecha.configure(fg_color="#000000")
    lienzo.revisar_hora()
    assert lienzo.derecha.cget("fg_color") == efectos.DEGRADADOS[efectos.momento_del_dia()][2] and app.ficha_informe.winfo_exists()


def test_celebrar_pone_el_check_y_devuelve_el_boton_a_su_sitio(app, monkeypatch):
    app.config["animacion_inicio"] = True
    boton = app.btn_auto
    texto, color = boton.cget("text"), boton.cget("fg_color")
    monkeypatch.setattr(efectos, "CELEBRA_TRAZO_MS", 1)
    pendientes = []
    monkeypatch.setattr(boton, "after", lambda ms, funcion: pendientes.append(funcion))
    assert app.celebrar_tarea("auto", "stock completa")
    assert not app.celebrar_tarea("auto")  # ya está celebrando
    while pendientes and boton.cget("text") == "":
        pendientes.pop(0)()
    assert boton.cget("text") == texto and boton.cget("fg_color") == color and not boton._celebrando
    assert not app.celebrar_tarea("auto", "detenida") and not app.celebrar_tarea("auto", "stock con incidencias")
    app.config["animacion_inicio"] = False
    assert not app.celebrar_tarea("auto")

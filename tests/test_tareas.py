import json
from datetime import date

import pytest

import tareas

HOY = date(2026, 10, 22)


# ---------- fechas ----------

@pytest.mark.parametrize("texto,esperada", [
    ("", None), ("hoy", HOY), ("Mañana", date(2026, 10, 23)), ("25/10", date(2026, 10, 25)),
    ("3-11", date(2026, 11, 3)), ("5/1", date(2027, 1, 5)),  # un dd/mm ya pasado es del año que viene
    ("22/10", HOY), ("1/2/27", date(2027, 2, 1)), ("01.02.2028", date(2028, 2, 1)),
])
def test_interpretar_fecha(texto, esperada):
    assert tareas.interpretar_fecha(texto, HOY) == esperada


@pytest.mark.parametrize("texto", ["el lunes", "32/10", "29/02/2027", "10"])
def test_fechas_que_no_se_entienden(texto):
    with pytest.raises(ValueError):
        tareas.interpretar_fecha(texto, HOY)


def test_estado_y_texto_de_la_fecha():
    assert tareas.estado_fecha("", HOY) == "" and tareas.texto_fecha("", HOY) == ""
    assert tareas.estado_fecha("2026-10-21", HOY) == "vencida" and tareas.texto_fecha("2026-10-21", HOY) == "Ayer"
    assert tareas.estado_fecha("2026-10-22", HOY) == "hoy" and tareas.texto_fecha("2026-10-22", HOY) == "Hoy"
    assert tareas.texto_fecha("2026-10-23", HOY) == "Mañana"
    assert tareas.texto_fecha("2026-10-18", HOY) == "Venció el 18/10"
    assert tareas.texto_fecha("2026-11-02", HOY) == "Vence el 02/11"
    assert tareas.texto_fecha("2027-01-05", HOY) == "Vence el 05/01/2027"


# ---------- lista ----------

@pytest.fixture
def lista(tmp_path):
    return tareas.ListaTareas(str(tmp_path / tareas.ARCHIVO))


def test_anadir_y_guardar(lista):
    assert lista.anadir("   ") is None
    tarea = lista.anadir("  Llamar a   Roberto  ", date(2026, 10, 25), importante=True)
    assert tarea["texto"] == "Llamar a Roberto" and tarea["fecha"] == "2026-10-25" and tarea["importante"]
    otra = tareas.ListaTareas(lista.ruta)  # se ha guardado en el archivo
    assert [t["texto"] for t in otra.tareas] == ["Llamar a Roberto"]


def test_orden_importantes_fecha_y_completadas_al_final(lista):
    sin_fecha = lista.anadir("sin fecha")
    tarde = lista.anadir("tarde", date(2026, 12, 1))
    pronto = lista.anadir("pronto", date(2026, 10, 24))
    importante = lista.anadir("importante", importante=True)
    hecha = lista.anadir("hecha")
    lista.marcar(hecha["id"], True)
    assert [t["texto"] for t in lista.ordenadas()] == ["importante", "pronto", "tarde", "sin fecha", "hecha"]
    lista.marcar(hecha["id"], False)
    assert lista.buscar(hecha["id"])["completada"] == ""
    assert lista.pendientes() == 5


def test_editar_cambiar_y_borrar(lista):
    tarea = lista.anadir("Pedir presupuesto")
    lista.editar(tarea["id"], "Pedir presupuesto a Mopar")
    assert lista.editar(tarea["id"], "  ") is None  # un texto vacío no borra la tarea
    lista.cambiar_fecha(tarea["id"], date(2026, 10, 20))
    lista.cambiar_importante(tarea["id"])
    guardada = tareas.ListaTareas(lista.ruta).buscar(tarea["id"])
    assert guardada["texto"] == "Pedir presupuesto a Mopar" and guardada["fecha"] == "2026-10-20" and guardada["importante"]
    assert lista.vencidas(HOY) == 1
    lista.cambiar_fecha(tarea["id"], None)
    assert lista.vencidas(HOY) == 0
    assert lista.borrar(tarea["id"]) and not lista.tareas
    assert not lista.borrar("no-existe")


def test_borrar_completadas(lista):
    for texto in ("a", "b", "c"):
        lista.anadir(texto)
    for tarea in lista.tareas[:2]:
        lista.marcar(tarea["id"], True)
    assert lista.borrar_completadas() == 2
    assert [t["texto"] for t in lista.tareas] == ["c"]


def test_archivo_danado_se_aparta_sin_perderlo(tmp_path):
    ruta = tmp_path / tareas.ARCHIVO
    ruta.write_text("{esto no es json", encoding="utf-8")
    lista = tareas.ListaTareas(str(ruta))
    assert lista.tareas == []
    assert (tmp_path / (tareas.ARCHIVO + ".danado")).read_text(encoding="utf-8") == "{esto no es json"
    lista.anadir("nueva")
    assert json.loads(ruta.read_text(encoding="utf-8"))["tareas"][0]["texto"] == "nueva"


# ---------- pestaña y cabecera ----------

@pytest.fixture
def app_tareas(app, tmp_path):
    app.lista_tareas = tareas.ListaTareas(str(tmp_path / tareas.ARCHIVO))
    app.tabview.set("Comparador")
    app.al_cambiar_pestana()
    app.pintar_tareas()
    return app


def textos_de_la_lista(app):
    """Textos visibles de la lista (etiquetas de customtkinter, no sus piezas internas de Tk)."""
    textos = []
    for widget in app.lista_tareas_frame.winfo_children():
        for pieza in [widget] + widget.winfo_children():
            if type(pieza).__name__ == "CTkLabel":
                textos.append(pieza.cget("text"))
    return textos


def tecla_enlazada(campo, tecla):
    """La tecla está asociada al campo. (No se simula la pulsación: sin la ventana en primer plano,
    Windows no le da el foco del teclado y la pulsación simulada no llegaría al campo.)"""
    return bool(campo._entry.bind(tecla))


def test_anadir_desde_la_pestana_con_intro(app_tareas):
    app = app_tareas
    assert "No tienes tareas pendientes" in textos_de_la_lista(app)[0]
    app.nueva_tarea_entry.insert(0, "Revisar pedido S39P")
    app.nueva_fecha_entry.insert(0, "hoy")
    app.alternar_nueva_importante()
    assert tecla_enlazada(app.nueva_tarea_entry, "<Return>") and tecla_enlazada(app.nueva_fecha_entry, "<Return>")
    assert app.anadir_tarea() == "break"  # lo que hace Intro
    app.update()
    tarea = app.lista_tareas.tareas[0]
    assert tarea["texto"] == "Revisar pedido S39P" and tarea["importante"] and tarea["fecha"] == date.today().isoformat()
    assert app.nueva_tarea_entry.get() == "" and app._nueva_importante is False
    assert "Revisar pedido S39P" in textos_de_la_lista(app) and "Hoy" in textos_de_la_lista(app)


def test_fecha_no_valida_avisa_y_no_anade(app_tareas):
    app = app_tareas
    app.nueva_tarea_entry.insert(0, "Algo")
    app.nueva_fecha_entry.insert(0, "el lunes")
    app.anadir_tarea()
    assert app.lista_tareas.tareas == []
    assert app.avisos.historial[-1][0] == "aviso"


def test_marcar_editar_y_borrar_desde_la_pestana(app_tareas):
    app = app_tareas
    tarea = app.lista_tareas.anadir("Primera")
    app.pintar_tareas()
    app.marcar_tarea(tarea["id"], True)
    app.update()
    assert app.lista_tareas.buscar(tarea["id"])["hecha"]
    assert any("Borrar completadas (1)" == h.cget("text") for h in app.pie_tareas.winfo_children() if hasattr(h, "cget"))
    etiqueta = next(h for h in app.lista_tareas_frame.winfo_children()[0].winfo_children() if getattr(h, "cget", None) and h.cget("text") == "Primera")
    entrada = app.editar_tarea_en_linea(etiqueta, app.lista_tareas.buscar(tarea["id"]), "texto")
    entrada.delete(0, "end")
    entrada.insert(0, "Primera corregida")
    assert tecla_enlazada(entrada, "<Return>") and tecla_enlazada(entrada, "<FocusOut>")
    entrada.guardar_edicion()
    app.update()
    assert app.lista_tareas.buscar(tarea["id"])["texto"] == "Primera corregida"
    app.borrar_tareas_completadas()
    assert app.lista_tareas.tareas == []


def test_esc_al_editar_cancela_sin_detener_tareas(app_tareas, monkeypatch):
    app = app_tareas
    tarea = app.lista_tareas.anadir("No cambiar")
    app.pintar_tareas()
    detenidas = []
    monkeypatch.setattr(app, "detener_tarea_actual", lambda *a: detenidas.append(1))
    etiqueta = next(h for h in app.lista_tareas_frame.winfo_children()[0].winfo_children() if getattr(h, "cget", None) and h.cget("text") == "No cambiar")
    entrada = app.editar_tarea_en_linea(etiqueta, tarea, "texto")
    entrada.insert("end", " (cambiado)")
    assert tecla_enlazada(entrada, "<Escape>")
    assert entrada.cancelar_edicion() == "break"  # «break»: Esc no sigue hasta el atajo global de detener
    app.update()
    assert app.lista_tareas.buscar(tarea["id"])["texto"] == "No cambiar"
    assert detenidas == []  # Esc no ha llegado al atajo global de detener


def test_ficha_de_tareas_en_la_cabecera(app_tareas):
    app = app_tareas
    assert not hasattr(app, "ficha_tareas") or not app.ficha_tareas.winfo_exists()
    app.lista_tareas.anadir("Una")
    app.lista_tareas.anadir("Vencida", date(2020, 1, 1))
    app.pintar_tareas()
    textos = [h.cget("text") for h in app.ficha_tareas.winfo_children()]
    assert "2 tareas pendientes · 1 vencida" in textos
    app.tabview.set("Automatización")
    app.al_cambiar_pestana()
    app.ir_a_tareas()
    assert app.tabview.get() == "Comparador"

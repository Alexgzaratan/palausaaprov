import openpyxl
import pytest

import configuracion
import estilo
import app_automatizacion_optimizada as app_mod


# ---------- validaciones ----------

def test_correos_invalidos():
    assert app_mod.correos_invalidos("a@x.es; b@y.com, c@z.org") == []
    assert app_mod.correos_invalidos("") == []
    assert app_mod.correos_invalidos("a@x.es; sin_arroba; otro@") == ["sin_arroba", "otro@"]


def test_extraer_avisos_avisa_si_falta_la_cabecera(tmp_path):
    ruta = tmp_path / "hoja.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["hoja", "cualquier cosa"])
    wb.save(ruta)
    registro = []
    app_mod.extraer_avisos(str(ruta), lambda texto, tag=None: registro.append(texto))
    assert any("formato esperado" in t for t in registro)


# ---------- paleta ----------

def test_paletas_tienen_las_mismas_claves_y_sin_repetir_roles():
    assert set(estilo.PALETA_CLARA) == set(estilo.PALETA_OSCURA)
    assert estilo.paleta(True) is estilo.PALETA_OSCURA and estilo.paleta(False) is estilo.PALETA_CLARA


# ---------- reparto en columnas ----------

def buscar_contenedor(widget):
    """Contenedor de tarjetas de una pestaña (el marco desplazable donde se registran)."""
    if hasattr(widget, "_tarjetas"):
        return widget
    for hijo in widget.winfo_children():
        encontrado = buscar_contenedor(hijo)
        if encontrado is not None:
            return encontrado
    return None


def test_tarjetas_en_una_o_dos_columnas_segun_el_ancho(app):
    app.tabview.set("Automatización")
    contenedor = buscar_contenedor(app.tab_auto)
    assert len(contenedor._tarjetas) == 2
    app.repartir_tarjetas(contenedor, estilo.ANCHO_DOS_COLUMNAS + 100)
    assert {t.grid_info()["column"] for t in contenedor._tarjetas} == {0, 1}
    app.repartir_tarjetas(contenedor, estilo.ANCHO_DOS_COLUMNAS - 100)
    assert {t.grid_info()["column"] for t in contenedor._tarjetas} == {0}
    assert [t.grid_info()["row"] for t in contenedor._tarjetas] == [0, 1]


def test_tarjeta_impar_final_ocupa_todo_el_ancho(app):
    app.tabview.set("Pedidos")
    app.al_cambiar_pestana()
    contenedor = buscar_contenedor(app.tab_pedidos)
    app.repartir_tarjetas(contenedor, estilo.ANCHO_DOS_COLUMNAS + 100)
    ultima = contenedor._tarjetas[-1].grid_info()
    assert int(ultima["columnspan"]) == 2 and int(ultima["column"]) == 0


# ---------- panel de actividad ----------

def test_panel_se_oculta_en_reclamaciones_y_vuelve_al_salir(app):
    assert app.panel_actividad.winfo_manager() == "pack"
    for nombre in ("Reclamaciones", "Facturas", "Reclamaciones", "Configuración", "Automatización"):
        app.tabview.set(nombre)
        app.al_cambiar_pestana()
        esperado = "" if nombre == "Reclamaciones" else "pack"
        assert app.panel_actividad.winfo_manager() == esperado, nombre


def test_registro_plegable_y_persistente(app):
    assert app.registro_visible is True
    app.alternar_registro()
    assert app.registro_visible is False and not app.frame_log.winfo_manager()
    assert configuracion.cargar_configuracion()["registro_visible"] is False
    app.reconstruir_interfaz()
    assert app.registro_visible is False
    app.alternar_registro()
    assert app.registro_visible is True and app.frame_log.winfo_manager() == "pack"


# ---------- atajos ----------

def test_ctrl_numero_cambia_de_pestana(app):
    app.ir_a_pestana(5)
    assert app.tabview.get() == "Facturas" and "Facturas" in app.pestanas_construidas
    app.ir_a_pestana(7)
    assert app.tabview.get() == "Configuración"
    app.ir_a_pestana(99)
    assert app.tabview.get() == "Configuración"


def test_ctrl_intro_no_actua_si_la_tarea_ya_esta_en_curso(app, monkeypatch):
    llamadas = []
    monkeypatch.setattr(app, "iniciar_auto", lambda *a, **k: llamadas.append(1))
    app.tabview.set("Automatización")
    app.accion_principal()
    assert llamadas == [1]
    monkeypatch.setattr(app_mod, "automation_running", True)
    app.accion_principal()
    assert llamadas == [1]


def test_escape_sin_tareas_no_hace_nada(app, monkeypatch):
    preguntas = []
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: preguntas.append(a) or False)
    app.detener_tarea_actual()
    assert preguntas == []
    monkeypatch.setattr(app_mod, "automation_running", True)
    app.detener_tarea_actual()
    assert len(preguntas) == 1


def test_atajos_registrados(app):
    atajos = app.bind_all()
    assert "<Control-Key-Return>" in atajos
    assert "<Key-Escape>" in atajos
    assert all(f"<Control-Key-{n}>" in atajos for n in range(1, 9))


# ---------- ventana ----------

def test_tamano_inicial_proporcional_a_la_pantalla(app):
    ancho, alto, *_ = app.winfo_geometry().replace("x", "+").split("+")
    assert 900 <= int(ancho) <= 1500 and 680 <= int(alto) <= 1000

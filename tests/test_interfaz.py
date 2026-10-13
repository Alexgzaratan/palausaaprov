import pytest

import configuracion
import app_automatizacion_optimizada as app_mod


def test_solo_se_construye_la_primera_pestana_al_arrancar(app):
    assert app.pestanas_construidas == {"Automatización"}
    assert hasattr(app, "btn_auto")
    assert not hasattr(app, "btn_madd")


def test_cambiar_de_pestana_la_construye_con_los_valores_guardados(app):
    app.tabview.set("Facturas")
    app.al_cambiar_pestana()
    assert "Facturas" in app.pestanas_construidas
    assert app.usuario_madd_entry.get() == "DF1"


def test_guardar_con_pestanas_sin_construir_no_pierde_valores(app):
    app.guardar_configuracion_ui_internal()
    guardado = configuracion.cargar_configuracion()
    assert guardado["usuario_madd"] == "DF1"
    assert guardado["ruta_excel_mopar"] == "C:/mopar.xls"
    assert guardado["email_companeros_avisos"] == "a@x.es; b@x.es"
    assert guardado["num_hilos"] == 4
    assert guardado["avisos_envio_automatico"] is False
    assert guardado["modo_apariencia"] == "Light"
    assert guardado["mopar_numero"] == 7


def test_lo_editado_en_una_pestana_construida_se_guarda(app):
    app.tabview.set("Facturas")
    app.al_cambiar_pestana()
    app.usuario_madd_entry.delete(0, "end")
    app.usuario_madd_entry.insert(0, "NUEVO")
    app.guardar_configuracion_ui_internal()
    assert configuracion.cargar_configuracion()["usuario_madd"] == "NUEVO"


def test_cambio_de_tema_reconstruye_y_conserva_pestana_y_datos(app):
    app.tabview.set("Configuración")
    app.al_cambiar_pestana()
    app.modo_option.set("Oscuro")
    app.cambiar_modo_apariencia()
    assert app.config["modo_apariencia"] == "Dark"
    assert app.color_superficie == "#1e293b"
    assert app.tabview.get() == "Configuración"
    assert app.pestanas_construidas == {"Automatización", "Configuración"}
    assert not hasattr(app, "usuario_madd_entry")
    app.tabview.set("Facturas")
    app.al_cambiar_pestana()
    assert app.usuario_madd_entry.get() == "DF1"
    assert configuracion.cargar_configuracion()["modo_apariencia"] == "Dark"


def test_no_se_cambia_el_tema_con_una_tarea_en_curso(app, monkeypatch):
    app.tabview.set("Configuración")
    app.al_cambiar_pestana()
    monkeypatch.setattr(app_mod, "avisos_running", True)
    app.modo_option.set("Oscuro")
    app.cambiar_modo_apariencia()
    assert app.config["modo_apariencia"] == "Light"
    assert app.modo_option.get() == "Claro"
    assert app.mensajes

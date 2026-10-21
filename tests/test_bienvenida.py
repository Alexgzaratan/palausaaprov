from datetime import datetime

import app_automatizacion_optimizada as app_mod


def test_saludo_segun_la_hora():
    assert app_mod.texto_saludo("", datetime(2026, 10, 21, 8)) == "Buenos días"
    assert app_mod.texto_saludo("Álex", datetime(2026, 10, 21, 16)) == "Buenas tardes, Álex"
    assert app_mod.texto_saludo("  ", datetime(2026, 10, 21, 22)) == "Buenas noches"


def test_fecha_larga_en_castellano():
    assert app_mod.fecha_larga(datetime(2026, 10, 21)) == "Miércoles, 21 de octubre"


def test_resumen_del_dia(tmp_path, monkeypatch):
    hoy = datetime(2026, 10, 21, 9)
    carpeta = tmp_path / "hojas"
    hoja = app_mod.ruta_hoja_pedidos_del_dia(str(carpeta), hoy)
    log = tmp_path / "log.txt"
    log.write_text("--- CORREO [2026-10-21 07:01:30] ---\nFecha objetivo: 2026-10-20\nEstado: enviado\n", encoding="utf-8")
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(log))
    config = {"ruta_carpeta_hojas": str(carpeta), "usuario_madd": "X"}
    carpeta.mkdir()
    assert app_mod.resumen_del_dia(config, hoy) == [("calendario", "Hoja de pedidos de hoy pendiente", False), ("check", "Facturas enviadas a las 07:01", True)]
    hoja.parent.mkdir(parents=True)
    hoja.write_bytes(b"x")
    log.write_text("--- CORREO [2026-10-21 07:01:30] ---\nEstado: PENDIENTE en la bandeja de salida de Outlook\n", encoding="utf-8")
    assert app_mod.resumen_del_dia(config, hoy) == [("check", "Hoja de pedidos de hoy creada", True), ("aviso", "Facturas pendientes en Outlook", False)]
    assert app_mod.resumen_del_dia({}, hoy) == []
    sabado = datetime(2026, 10, 24, 9)
    assert all("Facturas" not in texto for _, texto, _ in app_mod.resumen_del_dia(config, sabado))


def test_cabecera_muestra_saludo_y_cambia_de_tema(app):
    import estilo
    assert app.tema_actual() in estilo.TEMAS
    app.tabview.set("Configuración")
    app.al_cambiar_pestana()
    app.tema_option.set("Salvia")
    app.cambiar_modo_apariencia()
    assert app.config["tema_color"] == "Salvia"
    assert app.color_principal == estilo.TEMAS["Salvia"]["claro"]["principal"]

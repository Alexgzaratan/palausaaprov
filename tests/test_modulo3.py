"""Facturas MADD: días sin enviar, resumen semanal, subcarpetas por mes y tarea programada."""
import json
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

import app_automatizacion_optimizada as app_mod
import historial_madd
import tarea_windows

REGISTRO = """--- EJECUCIÓN [2026-10-27 07:12:35] ---
Fecha objetivo: 2026-10-24
Nº facturas descargadas: 11
Sin errores.
------------------------------

--- CORREO [2026-10-27 07:12:40] ---
Fecha objetivo: 2026-10-24
Estado: enviado
Destinatarios: a@ejemplo.test | adjuntos: 11
------------------------------

--- EJECUCIÓN [2026-10-28 07:02:42] ---
Fecha objetivo: 2026-10-27
Nº facturas descargadas: 0
Error: Page.goto: net::ERR_CONNECTION_TIMED_OUT (paso: acceso/login)
------------------------------

--- CORREO [2026-10-28 07:03:00] ---
Fecha objetivo: 2026-10-27
Estado: aviso de fallo: enviado
Destinatarios: aviso@ejemplo.test
------------------------------

--- CORREO [2026-10-28 09:00:00] ---
Fecha objetivo: 2026-10-23
Estado: PENDIENTE en la bandeja de salida de Outlook
Destinatarios: a@ejemplo.test | adjuntos: 5
------------------------------
"""


@pytest.fixture
def registro(tmp_path):
    ruta = tmp_path / "log_facturas_madd.txt"
    ruta.write_text(REGISTRO, encoding="utf-8")
    return ruta


def test_lee_el_registro(registro):
    bloques = historial_madd.leer_registro(registro)
    assert [b["tipo"] for b in bloques] == ["ejecucion", "correo", "ejecucion", "correo", "correo"]
    assert bloques[2]["error"].startswith("Page.goto") and bloques[1]["adjuntos"] == 11
    # el aviso de fallo no cuenta como enviado; lo que sigue en la bandeja de salida sí (saldrá solo)
    assert historial_madd.dias_enviados(bloques) == {date(2026, 10, 24), date(2026, 10, 23)}
    assert historial_madd.leer_registro(registro.parent / "no_existe.txt") == []


def test_dias_pendientes():
    enviados = {date(2026, 10, 23), date(2026, 10, 24)}  # viernes 23 y sábado 24 (por un reenvío)
    # jueves 29: faltan lunes 26, martes 27 y miércoles 28 (el fin de semana no cuenta)
    assert historial_madd.dias_pendientes(date(2026, 10, 28), enviados) == [date(2026, 10, 26), date(2026, 10, 27), date(2026, 10, 28)]
    assert historial_madd.dias_pendientes(date(2026, 10, 24), enviados) == []  # ya enviado
    assert historial_madd.dias_pendientes(date(2026, 10, 28), set()) == [date(2026, 10, 28)]  # sin historial no se va hacia atrás
    largo = historial_madd.dias_pendientes(date(2026, 12, 31), {date(2026, 10, 1)})
    assert len(largo) == historial_madd.MAX_DIAS_PENDIENTES and largo[-1] == date(2026, 12, 31)


def test_resumen_de_la_semana(registro):
    bloques = historial_madd.leer_registro(registro)
    assert historial_madd.resumen_semana(bloques, date(2026, 10, 29)) == (16, 2, 1)
    assert historial_madd.texto_resumen_semana(bloques, date(2026, 10, 29)) == "Esta semana: 16 facturas en 2 correos · 1 fallo"
    assert historial_madd.texto_resumen_semana(bloques, date(2026, 11, 4)).startswith("Esta semana todavía no")


def test_facturas_en_subcarpetas_por_mes(tmp_path):
    dia = date(2026, 10, 28)
    (tmp_path / "Factura_111_2026-10-28.pdf").write_bytes(b"vieja")  # descargada antes del cambio
    nueva = historial_madd.ruta_nueva_factura(tmp_path, "222", dia)
    assert nueva == tmp_path / "2026-10" / "Factura_222_2026-10-28.pdf"
    nueva.parent.mkdir()
    nueva.write_bytes(b"nueva")
    (tmp_path / "2026-10" / "Factura_333_2026-10-27.pdf").write_bytes(b"otro dia")
    assert [r.name for r in app_mod.facturas_locales_del_dia(str(tmp_path), dia)] == ["Factura_111_2026-10-28.pdf", "Factura_222_2026-10-28.pdf"]
    assert historial_madd.factura_existente(tmp_path, "Factura_111_2026-10-28.pdf", dia) == tmp_path / "Factura_111_2026-10-28.pdf"
    assert historial_madd.factura_existente(tmp_path, "Factura_999_2026-10-28.pdf", dia) is None
    assert historial_madd.fecha_de_factura(nueva) == dia


def test_varios_dias_en_una_sesion_y_un_correo_por_dia(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(tmp_path / "log.txt"))
    monkeypatch.setattr(app_mod, "registrar_ultima_ejecucion", lambda *a, **k: None)
    sesiones, correos = [], []

    def descarga(usuario, password, carpeta, fechas, log):
        sesiones.append(set(fechas))
        return {"facturas": [str(tmp_path / "2026-10" / f"Factura_1_{d.isoformat()}.pdf") for d in sorted(fechas) if d.day != 27],
                "error": None, "paso_fallido": None}

    monkeypatch.setattr(app_mod, "ejecutar_descarga_facturas_madd", descarga)
    monkeypatch.setattr(app_mod, "_madd_enviar_correo_resultado", lambda email, fecha, rutas, log: correos.append((fecha, len(rutas))))
    dias = [date(2026, 10, 28), date(2026, 10, 26), date(2026, 10, 27)]
    app_mod.tarea_facturas_madd("u", "p", str(tmp_path), "a@ejemplo.test", "", fechas=dias)
    assert sesiones == [set(dias)]  # una sola sesión en MADD
    assert correos == [(date(2026, 10, 26), 1), (date(2026, 10, 27), 0), (date(2026, 10, 28), 1)]  # el 27 sin facturas: aviso «no hay»
    bloques = historial_madd.leer_registro(tmp_path / "log.txt")
    assert [(b["fecha"].day, b["facturas"]) for b in bloques] == [(26, 1), (27, 0), (28, 1)]


def test_la_tarea_diaria_recupera_los_dias_sin_enviar(registro, monkeypatch):
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(registro))
    monkeypatch.setattr(app_mod, "cargar_configuracion", lambda: {"usuario_madd": "u", "password_madd": "p", "ruta_carpeta_facturas_madd": "c",
                                                                 "email_facturas_madd": "e", "filtro_dia_facturas_madd": "Ayer"})
    monkeypatch.setattr(app_mod, "fecha_objetivo_madd", lambda filtro="Ayer", hoy=None: date(2026, 10, 28))
    llamadas = []
    monkeypatch.setattr(app_mod, "tarea_facturas_madd", lambda *a, resultado_callback=None, fechas=None, **k: (llamadas.append(fechas), resultado_callback({"error": None})))
    assert app_mod.ejecutar_tarea_madd_standalone() == 0
    assert llamadas == [[date(2026, 10, 26), date(2026, 10, 27), date(2026, 10, 28)]]


# ---------- tarea programada ----------

ESTADO = {"ejecutable": r"C:\Palausa\programa\app_automatizacion_optimizada.exe", "argumentos": "--tarea-madd",
          "inicio": "2026-10-16T06:00:00", "dias": 62, "ultima": "2026-10-28T07:00:15", "resultado": 0,
          "proxima": "2026-10-30T06:00:00", "perdidas": 1, "estado": "Ready"}


def test_resumen_de_la_tarea():
    texto, problema = tarea_windows.resumen(ESTADO, datetime(2026, 10, 29, 10, 0))
    assert texto == "Tarea de Windows: laborables a las 06:00 · última 28/10 07:00 ✓ · hoy no se ha ejecutado" and problema
    texto, problema = tarea_windows.resumen(ESTADO, datetime(2026, 10, 29, 6, 10))  # aún dentro del margen
    assert texto.endswith("próxima 30/10 06:00") and not problema
    assert not tarea_windows.no_se_ejecuto_hoy(ESTADO, datetime(2026, 10, 31, 12, 0))  # sábado: no toca
    fallida = dict(ESTADO, ultima="2026-10-29T06:00:05", resultado=1)
    assert tarea_windows.resumen(fallida, datetime(2026, 10, 29, 10, 0)) == ("Tarea de Windows: laborables a las 06:00 · última 29/10 06:00 ✗ (código 1) · próxima 30/10 06:00", True)
    assert tarea_windows.resumen(None) == ("Tarea de Windows sin activar: las facturas no se envían solas.", True)


def test_hora_con_zona_horaria():
    assert tarea_windows.hora_local("2026-10-16T06:00:00").strftime("%H:%M") == "06:00"
    assert tarea_windows.hora_local("2026-10-16T06:00:00+01:00", date(2026, 12, 1)) is not None
    assert tarea_windows.hora_local("basura") is None


def test_detecta_la_ruta_antigua_y_la_cambia_al_lanzador():
    lanzador = Path(r"C:\Palausa\app_automatizacion_optimizada.exe")
    assert tarea_windows.apunta_al_programa(ESTADO, lanzador)
    assert not tarea_windows.apunta_al_programa(dict(ESTADO, ejecutable=str(lanzador)), lanzador)
    ordenes = []
    hecho = tarea_windows.APUNTAR_REAL(lanzador, ejecutar=lambda orden, **k: ordenes.append(orden[-1]) or SimpleNamespace(returncode=0))
    assert hecho and "Set-ScheduledTask -TaskName 'PalausaFacturasMADD'" in ordenes[0]
    assert r"-Execute 'C:\Palausa\app_automatizacion_optimizada.exe'" in ordenes[0] and "-WorkingDirectory 'C:\\Palausa'" in ordenes[0]


def test_leer_estado_interpreta_la_respuesta():
    respuesta = SimpleNamespace(stdout=json.dumps(ESTADO), returncode=0)
    assert tarea_windows.LEER_ESTADO_REAL(ejecutar=lambda *a, **k: respuesta) == ESTADO
    assert tarea_windows.LEER_ESTADO_REAL(ejecutar=lambda *a, **k: SimpleNamespace(stdout="null", returncode=0)) is None
    assert tarea_windows.LEER_ESTADO_REAL(ejecutar=lambda *a, **k: SimpleNamespace(stdout="error", returncode=1)) is None


def test_la_pestana_muestra_el_estado_y_el_resumen(app, monkeypatch, registro):
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(registro))
    app.tabview.set("Facturas")
    app.construir_pestana("Facturas")
    app.mostrar_estado_tarea_madd(ESTADO)
    assert app.estado_tarea_madd_label.cget("text").startswith("Tarea de Windows: laborables a las 06:00")
    assert app.resumen_semana_madd_label.cget("text").startswith("Esta semana")


def test_al_abrir_avisa_de_los_dias_sin_enviar(app, monkeypatch, registro):
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(registro))
    monkeypatch.setattr(app_mod, "fecha_objetivo_madd", lambda filtro="Ayer", hoy=None: date(2026, 10, 28))
    monkeypatch.setattr(tarea_windows, "no_se_ejecuto_hoy", lambda estado, ahora=None: True)
    monkeypatch.setattr(app, "refrescar_estado_tarea_madd", lambda al_terminar=None: al_terminar(ESTADO))
    app.config.update({"usuario_madd": "u", "password_madd": "p", "email_facturas_madd": "a@ejemplo.test", "ruta_carpeta_facturas_madd": "c"})
    app.comprobar_facturas_pendientes()
    aviso = app.avisos.activos[-1]
    assert aviso.texto == "Facturas MADD sin enviar: 26/10, 27/10, 28/10." and aviso.boton_accion.cget("text") == "Enviar ahora"
    lanzadas = []
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(app_mod, "tarea_facturas_madd", lambda *a, fechas=None, **k: lanzadas.append(fechas))
    assert app.enviar_facturas_pendientes()
    app_mod.madd_running = False
    import time
    time.sleep(0.3)
    assert lanzadas == [[date(2026, 10, 26), date(2026, 10, 27), date(2026, 10, 28)]]


def test_sin_tarea_programada_no_avisa(app, monkeypatch, registro):
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(registro))
    monkeypatch.setattr(app, "refrescar_estado_tarea_madd", lambda al_terminar=None: al_terminar(None))
    app.config.update({"usuario_madd": "u", "password_madd": "p", "email_facturas_madd": "a@ejemplo.test", "ruta_carpeta_facturas_madd": "c"})
    antes = len(app.avisos.activos)
    app.comprobar_facturas_pendientes()
    assert len(app.avisos.activos) == antes

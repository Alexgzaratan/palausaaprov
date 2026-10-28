import os
import time
from datetime import date, datetime

import pytest

import app_automatizacion_optimizada as app_mod
import configuracion
import efectos

HOY = date(2026, 10, 28)


def crear(ruta, contenido="x", modificado=None):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(contenido)
    if modificado:
        marca = datetime(*modificado).timestamp()
        os.utime(ruta, (marca, marca))
    return ruta


# ---------- informe diario anterior ----------

def test_fecha_en_el_nombre():
    assert app_mod.fecha_en_nombre("INFORMES ALMACEN 27-10-2026.xlsx") == date(2026, 10, 27)
    assert app_mod.fecha_en_nombre("Penurias 31-07-26.csv") == date(2026, 7, 31)
    assert app_mod.fecha_en_nombre("informe_01.02.2027.xlsx") == date(2027, 2, 1)
    assert app_mod.fecha_en_nombre("INFORME.xlsx") is None
    assert app_mod.fecha_en_nombre("31-02-2026.xlsx") is None


def test_informe_anterior_es_el_mas_reciente_antes_de_hoy(tmp_path):
    for nombre in ("INFORMES ALMACEN 23-10-2026.xlsx", "INFORMES ALMACEN 27-10-2026.xlsx", "INFORMES ALMACEN 26-10-2026.xlsx",
                   "INFORMES ALMACEN 28-10-2026.xlsx",  # el de hoy no es «el anterior»
                   "~$INFORMES ALMACEN 27-10-2026.xlsx", "notas 27-10-2026.txt"):
        crear(tmp_path / nombre)
    assert app_mod.informe_diario_anterior(tmp_path, HOY).name == "INFORMES ALMACEN 27-10-2026.xlsx"


def test_sin_fechas_en_el_nombre_usa_la_fecha_de_modificacion(tmp_path):
    crear(tmp_path / "informe viejo.xlsx", modificado=(2026, 10, 20, 9))
    crear(tmp_path / "informe de ayer.xlsx", modificado=(2026, 10, 27, 17))
    crear(tmp_path / "informe de hoy.xlsx", modificado=(2026, 10, 28, 8))
    assert app_mod.informe_diario_anterior(tmp_path, HOY).name == "informe de ayer.xlsx"


def test_carpeta_vacia_o_inexistente(tmp_path):
    assert app_mod.informe_diario_anterior(tmp_path, HOY) is None
    assert app_mod.informe_diario_anterior(tmp_path / "no_existe", HOY) is None


# ---------- botones ----------

@pytest.fixture
def abiertos(monkeypatch):
    lista = []
    monkeypatch.setattr(app_mod, "abrir_ruta_en_sistema", lambda ruta: lista.append(str(ruta)) or True)
    return lista


def test_icono_abre_el_excel_de_stock(app, tmp_path, abiertos):
    excel = crear(tmp_path / "S40P.xlsx")
    app.excel_auto_entry.delete(0, "end")
    app.excel_auto_entry.insert(0, str(excel))
    fila = app.excel_auto_entry.master
    icono = next(h for h in fila.winfo_children() if hasattr(h, "_pista"))
    icono.invoke()
    assert abiertos == [str(excel)]
    app.excel_auto_entry.delete(0, "end")
    app.excel_auto_entry.insert(0, str(tmp_path / "movido.xlsx"))
    icono.invoke()
    assert abiertos == [str(excel)]
    assert app.avisos.historial[-1] == ("aviso", "No se encuentra movido.xlsx: puede que se haya movido o borrado.")


def test_solo_el_excel_de_stock_tiene_icono_de_abrir(app):
    for campo in (app.excel_quiter_entry, app.excel_rot_entry, app.folder_rot_entry):
        assert not any(hasattr(h, "_pista") for h in campo.master.winfo_children())


def test_abrir_hoja_de_hoy(app, tmp_path, abiertos):
    app.tabview.set("Pedidos")
    app.al_cambiar_pestana()
    app.carpeta_hojas_entry.delete(0, "end")
    app.carpeta_hojas_entry.insert(0, str(tmp_path))
    app.btn_abrir_hoja.invoke()
    assert abiertos == [] and "Todavía no hay hoja de hoy" in app.avisos.historial[-1][1]
    hoja = crear(app_mod.ruta_hoja_pedidos_del_dia(str(tmp_path)))
    app.btn_abrir_hoja.invoke()
    assert abiertos == [str(hoja)]


def test_abrir_informe_anterior(app, tmp_path, abiertos, monkeypatch):
    app.tabview.set("Comparador")
    app.al_cambiar_pestana()
    app.config["ruta_carpeta_informes"] = ""
    ayer = crear(tmp_path / "INFORMES ALMACEN 01-01-2026.xlsx")
    # Sin carpeta configurada: se elige una vez y queda guardada
    monkeypatch.setattr(app_mod.filedialog, "askdirectory", lambda **k: str(tmp_path))
    app.btn_informe_anterior.invoke()
    assert abiertos == [str(ayer)]
    assert configuracion.cargar_configuracion()["ruta_carpeta_informes"] == str(tmp_path)
    assert app.pista_informe_anterior() == "Abrir INFORMES ALMACEN 01-01-2026.xlsx"
    # Ya no vuelve a preguntar
    monkeypatch.setattr(app_mod.filedialog, "askdirectory", lambda **k: pytest.fail("no debe volver a preguntar"))
    app.btn_informe_anterior.invoke()
    assert abiertos == [str(ayer)] * 2


def test_carpeta_de_informes_en_configuracion(app, tmp_path):
    app.tabview.set("Configuración")
    app.al_cambiar_pestana()
    app.carpeta_informes_entry.delete(0, "end")
    app.carpeta_informes_entry.insert(0, str(tmp_path))
    app.guardar_configuracion_ui_internal()
    assert configuracion.cargar_configuracion()["ruta_carpeta_informes"] == str(tmp_path)


# ---------- pista al pasar el ratón ----------

def test_pista_aparece_y_se_va(app, monkeypatch):
    boton = app.btn_auto
    efectos.poner_pista(app, boton, lambda: "Texto de la pista", espera_ms=10)
    boton._canvas.event_generate("<Enter>")
    fin = time.time() + 1
    while boton._pista["ventana"] is None and time.time() < fin:
        app.update()
        time.sleep(0.01)
    ventana = boton._pista["ventana"]
    assert ventana is not None and ventana.winfo_children()[0].cget("text") == "Texto de la pista"
    monkeypatch.setattr(efectos, "puntero_dentro", lambda widget: False)
    boton._canvas.event_generate("<Leave>")
    fin = time.time() + 1
    while boton._pista["ventana"] is not None and time.time() < fin:
        app.update()
        time.sleep(0.01)
    assert boton._pista["ventana"] is None and not ventana.winfo_exists()

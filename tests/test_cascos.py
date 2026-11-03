"""Cruce de cascos con ventas (regla de la skill cruce-cascos-ventas)."""
import datetime
import os

import openpyxl
import pandas as pd
import pytest

import app_automatizacion_optimizada as app_mod
import cascos_cruce as cc


@pytest.fixture
def archivos(tmp_path):
    ventas = tmp_path / "VENTAS 8049.xlsx"
    with pd.ExcelWriter(ventas) as w:
        pd.DataFrame({"Fecha": ["05/08/2026", "06/08/2026", "07/08/2026", "31/08/2026", None],
                      "Documento": ["ALBARAN4126 103585", "ALBARAN4126 103600", "ALBARAN4126 103700", "R4126 555", "*** TOTAL"],
                      "Refer.": [11301237, 11301300, 11301400, 0, None]}).to_excel(w, sheet_name="Informe", index=False)
    cascos = tmp_path / "CASCOS_PROVISIONES.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RECAMBIOS PALAUSA"
    ws.append(["FECHA  RECEPCION PLACA", "CLIENTE", "Nº ALBARÁN"])
    filas = [
        (datetime.datetime(2026, 8, 5), 8049, "ALBARAN4126 103585"),  # cruza -> 11301237
        (datetime.datetime(2026, 8, 6), 8049, "N4126103600"),         # cruza -> 11301300
        (datetime.datetime(2026, 8, 7), 8049, 11301400),              # ya es una Refer.
        (datetime.datetime(2026, 8, 8), 8049, "N412693742"),          # anterior al informe (93742 < 103585)
        (datetime.datetime(2026, 8, 9), 8049, 11200001),              # Refer. antigua
        (datetime.datetime(2026, 8, 9), 8049, "4106101958"),          # mal tecleado
        (datetime.datetime(2026, 8, 10), 7000, "ALBARAN4126 103700"),  # otro cliente: no se toca
        (datetime.datetime(2026, 7, 20), 8049, "ALBARAN4126 103700"),  # antes de la fecha desde: no se toca
    ]
    for fila in filas:
        ws.append(list(fila))
    wb.create_sheet("OTRA")["A1"] = "intacta"
    wb.save(cascos)
    return str(ventas), str(cascos), tmp_path


def test_cruce_como_la_skill(archivos):
    ventas, cascos, tmp = archivos
    antes = os.path.getmtime(cascos)
    salida = tmp / "resultado.xlsx"
    r = cc.cruzar(ventas, cascos, str(salida))
    assert r["desde"] == datetime.date(2026, 8, 1) and r["ya_estaban"] == 1
    assert [(f, nuevo) for f, _, _, nuevo in r["cambiados"]] == [(2, 11301237), (3, 11301300)]
    assert [(f, m) for f, _, _, m in r["sin_cruce"]] == [(5, "albarán anterior al informe de ventas"), (6, "Refer. antigua (8 cifras)"), (7, "revisar a mano")]
    ws = openpyxl.load_workbook(salida)["RECAMBIOS PALAUSA"]
    assert [ws.cell(f, 3).value for f in range(2, 10)] == [11301237, 11301300, 11301400, "N412693742", 11200001, "4106101958", "ALBARAN4126 103700", "ALBARAN4126 103700"]
    assert ws.cell(5, 3).fill.start_color.rgb == "00FFFF00" and ws.cell(8, 3).fill.fill_type is None
    assert openpyxl.load_workbook(salida)["OTRA"]["A1"].value == "intacta"
    assert os.path.getmtime(cascos) == antes  # el original no se toca
    texto = cc.resumen(r)
    assert texto.startswith("Desde el 01/08/2026: 2 albaranes cambiados por su Refer., 1 ya estaban bien y 3 sin cruce") and "fila 7 (09/08): 4106101958" in texto


def test_fecha_desde_y_otro_cliente(archivos):
    ventas, cascos, tmp = archivos
    r = cc.cruzar(ventas, cascos, str(tmp / "r.xlsx"), cliente="7000", desde=datetime.date(2026, 8, 1))
    assert [(f, nuevo) for f, _, _, nuevo in r["cambiados"]] == [(8, 11301400)]
    r = cc.cruzar(ventas, cascos, str(tmp / "r2.xlsx"), desde=datetime.date(2026, 7, 1))
    assert (9, 11301400) in [(f, nuevo) for f, _, _, nuevo in r["cambiados"]]  # ahora sí entra la fila de julio


def test_nunca_sobrescribe_el_original(archivos):
    ventas, cascos, _ = archivos
    with pytest.raises(ValueError, match="no puede sobrescribir"):
        cc.cruzar(ventas, cascos, cascos)
    assert cc.ruta_salida(cascos, datetime.date(2026, 11, 3)).name == "CASCOS_PROVISIONES (cruzado 03-11-2026).xlsx"


def test_tarjeta_en_el_comparador(app, archivos, monkeypatch):
    ventas, cascos, _ = archivos
    app.tabview.set("Comparador")
    app.construir_pestana("Comparador")
    assert app.cliente_cascos_entry.get() == "8049"
    for campo, valor in ((app.cascos_entry, cascos), (app.ventas_cascos_entry, ventas)):
        campo.delete(0, "end")
        campo.insert(0, valor)
    resumenes = []
    monkeypatch.setattr(app_mod.messagebox, "showinfo", lambda titulo, texto, **k: resumenes.append(texto))
    monkeypatch.setattr(app_mod.threading, "Thread", lambda target, daemon=None: type("H", (), {"start": lambda self: target()})())
    monkeypatch.setattr(app, "after", lambda ms, funcion, *a: funcion(*a))
    monkeypatch.setattr(app, "excel_abierto", lambda *a: False)
    assert app.cruzar_cascos()
    assert resumenes and "2 albaranes cambiados" in resumenes[0] and cc.ruta_salida(cascos).exists()

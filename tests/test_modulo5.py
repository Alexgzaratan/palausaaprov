"""Mopar y Stock FIAT: número nuevo por envío, nunca sobrescribir, letras más allá de la Z, aviso de Excel ya
convertido y vista previa."""
import os
from datetime import datetime

import json

import pytest

import app_automatizacion_optimizada as app_mod
import configuracion
import ficheros_mopar


# Los .xls se simulan: una lista de filas en JSON que lee un xlrd falso (no hace falta xlwt para escribirlos)
class HojaFalsa:
    def __init__(self, filas):
        self.filas = filas
        self.nrows = len(filas)
        self.ncols = max((len(f) for f in filas), default=0)

    def cell_value(self, r, c):
        fila = self.filas[r]
        return fila[c] if c < len(fila) else ""


class LibroFalso:
    def __init__(self, ruta):
        self.hoja = HojaFalsa(json.loads(open(ruta, encoding="utf-8").read()))

    def sheet_by_index(self, indice):
        return self.hoja


@pytest.fixture(autouse=True)
def xls_falso(monkeypatch):
    monkeypatch.setattr(ficheros_mopar.xlrd, "open_workbook", LibroFalso)


def guardar(ruta, filas):
    open(ruta, "w", encoding="utf-8").write(json.dumps(filas))
    return str(ruta)


def excel_mopar(ruta, filas):
    cabecera = ["Referencia", "Descripción", "Cant", "Cant. asig.", "Ref. pedido", "Cliente"]
    return guardar(ruta, [cabecera] + [[ref, "", float(cantidad), "", "", cliente] for ref, cantidad, cliente in filas])


def excel_stock_fiat(ruta, filas):
    return guardar(ruta, [[2.0, ref, "", cantidad if isinstance(cantidad, str) else float(cantidad)] for ref, cantidad in filas])


@pytest.fixture
def entorno(tmp_path, monkeypatch):
    monkeypatch.setattr(configuracion, "CONFIG_FILE", str(tmp_path / "config.json"))
    monkeypatch.setattr(ficheros_mopar, "REGISTRO", str(tmp_path / "registro.json"))
    configuracion.guardar_configuracion({})
    carpeta = tmp_path / "MOPAR"
    carpeta.mkdir()
    return tmp_path, carpeta


def mopar(ruta, carpeta):
    registro, generados = [], []
    app_mod.generar_fichero_mopar_task(ruta, str(carpeta), log_callback=lambda t, tag=None: registro.append(t), numero_callback=generados.append)
    return registro, generados


def test_dos_envios_el_mismo_dia_llevan_numeros_distintos_y_no_se_pisan(entorno):
    tmp, carpeta = entorno
    (carpeta / "F54A.txt").write_text("de ayer", encoding="utf-8")
    manana = excel_mopar(tmp / "m1.xls", [(f"2/9812345{i:03d}", 2, "CLI 8049") for i in range(24)])
    tarde = excel_mopar(tmp / "m2.xls", [("1629123280", 1, "8050")])
    assert mopar(manana, carpeta)[1] == ["F55A.txt", "F55B.txt", "F55C.txt"]
    assert mopar(tarde, carpeta)[1] == ["F56.txt"]
    assert (carpeta / "F54A.txt").read_text(encoding="utf-8") == "de ayer"
    primera = (carpeta / "F55A.txt").read_text(encoding="utf-8").splitlines()
    assert len(primera) == 10 and primera[0] == "0009812345000" + "0000002" + "00008049"
    assert (carpeta / "F56.txt").read_text(encoding="utf-8") == "0001629123280000000100008050"
    assert configuracion.cargar_configuracion()["mopar_numero"] == 56


def test_si_los_enviados_se_movieron_la_numeracion_sigue(entorno):
    tmp, carpeta = entorno
    configuracion.guardar_configuracion({"mopar_numero": 60})  # carpeta vacía: los ficheros ya se movieron
    assert mopar(excel_mopar(tmp / "m.xls", [("1", 1, "1")]), carpeta)[1] == ["F61.txt"]


def test_nunca_sobrescribe(entorno, monkeypatch):
    tmp, carpeta = entorno
    (carpeta / "F7B.txt").write_text("no tocar", encoding="utf-8")
    with pytest.raises(FileExistsError):
        ficheros_mopar.escribir(["x"] * 15, carpeta, 7)
    assert sorted(os.listdir(carpeta)) == ["F7B.txt"] and (carpeta / "F7B.txt").read_text(encoding="utf-8") == "no tocar"  # ni siquiera F7A


def test_letras_despues_de_la_z():
    assert [ficheros_mopar.letra(i) for i in (0, 25, 26, 27, 51, 52)] == ["A", "Z", "AA", "AB", "AZ", "BA"]
    nombres = ficheros_mopar.nombres_ficheros(9, 275)
    assert len(nombres) == 28 and len(set(nombres)) == 28 and nombres[-2:] == ["F9AA.txt", "F9AB.txt"]
    assert ficheros_mopar.nombres_ficheros(9, 10) == ["F9.txt"] and ficheros_mopar.texto_ficheros(nombres[:3]) == "F9A–F9C"


def test_stock_fiat_genera_igual_y_con_su_propia_numeracion(entorno):
    tmp, carpeta = entorno
    ruta = excel_stock_fiat(tmp / "s.xls", [("000000485860", 5), ("", 1), ("9812", "x")])
    generados = []
    app_mod.generar_fichero_stock_fiat_task(ruta, str(carpeta), log_callback=lambda *a: None, numero_callback=generados.append)
    assert generados == ["F1.txt"]
    assert (carpeta / "F1.txt").read_text(encoding="utf-8").splitlines() == ["0000000485860000000500000000", "0000000009812000000000000000"]
    assert configuracion.cargar_configuracion()["stock_fiat_numero"] == 1 and "mopar_numero" not in configuracion.cargar_configuracion()


def test_vista_previa_y_aviso_de_excel_ya_convertido(entorno):
    tmp, carpeta = entorno
    ruta = excel_mopar(tmp / "m.xls", [(f"R{i}", 1, "1") for i in range(12)])
    lineas = ficheros_mopar.lineas_mopar(ruta)
    texto, repetido = ficheros_mopar.texto_confirmacion("mopar", lineas, str(carpeta), 5)
    assert not repetido and texto.startswith("Se generarán 2 ficheros Mopar (F5A–F5B) con 12 líneas")
    mopar(ruta, carpeta)
    texto, repetido = ficheros_mopar.texto_confirmacion("mopar", lineas, str(carpeta), 6)
    assert repetido and f"ya se convirtió en F1A–F1B hoy a las {datetime.now():%H}" in texto and "(F6A–F6B)" in texto
    assert ficheros_mopar.ya_generado("stock_fiat", lineas) is None  # cada generador lleva su registro


def test_la_pestana_pregunta_antes_de_generar_y_lista_los_de_stock_fiat(app, entorno, monkeypatch):
    tmp, carpeta = entorno
    ruta = excel_stock_fiat(tmp / "s.xls", [("1", 1)])
    app.tabview.set("Stock FIAT")
    app.construir_pestana("Stock FIAT")
    for campo, valor in ((app.excel_stock_fiat_entry, ruta), (app.folder_stock_fiat_entry, str(carpeta))):
        campo.delete(0, "end")
        campo.insert(0, valor)
    preguntas = []
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda titulo, texto, **k: preguntas.append(texto) or False)
    app.iniciar_stock_fiat()
    assert preguntas and "Se generará 1 fichero Stock FIAT (F1)" in preguntas[0] and os.listdir(carpeta) == []
    app.agregar_mopar_generado("F1.txt", "stock_fiat_generados_box")
    assert app.stock_fiat_generados_box.get("1.0", "end").strip() == "F1.txt"

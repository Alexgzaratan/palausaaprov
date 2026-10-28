import openpyxl
import pytest

import app_automatizacion_optimizada as app_mod


def excel(tmp_path, filas):
    ruta = tmp_path / "clientes.xlsx"
    wb = openpyxl.Workbook()
    for fila in filas:
        wb.active.append(fila)
    wb.save(ruta)
    return str(ruta)


@pytest.mark.parametrize("filas", [
    [["CLIENTE", "CORREO"], [8049, "taller@ejemplo.test"], [8050, None]],
    [["CLIENTE", "NOMBRE", "CORREO"], [8049, "", "taller@ejemplo.test"], [8050, "", None]],
    [["CLIENTE", "EMAIL"], ["8049", "taller@ejemplo.test"], [8050.0, ""]],
    [["Nº", "DIRECCIÓN"], [8049, "taller@ejemplo.test"], [8050, None]],  # cabecera sin reconocer, dos columnas
])
def test_solo_numero_y_correo(tmp_path, filas):
    clientes = app_mod.cargar_clientes_excel(excel(tmp_path, filas))
    assert clientes == {"8049": {"nombre": "", "email": "taller@ejemplo.test"}, "8050": {"nombre": "", "email": ""}}


def test_sigue_valiendo_con_el_nombre_del_taller(tmp_path):
    filas = [["CLIENTE", "NOMBRE", "CORREO"], [8049, "TALLER UNO", "a@ejemplo.test;b@ejemplo.test"]]
    assert app_mod.cargar_clientes_excel(excel(tmp_path, filas)) == {"8049": {"nombre": "TALLER UNO", "email": "a@ejemplo.test;b@ejemplo.test"}}
    sin_cabecera = [["", "", ""], [8049, "TALLER UNO", "a@ejemplo.test"]]
    assert app_mod.cargar_clientes_excel(excel(tmp_path, sin_cabecera))["8049"] == {"nombre": "TALLER UNO", "email": "a@ejemplo.test"}


def test_correo_delante_del_nombre(tmp_path):
    filas = [["CLIENTE", "CORREO", "NOMBRE"], [8049, "a@ejemplo.test", "TALLER UNO"]]
    assert app_mod.cargar_clientes_excel(excel(tmp_path, filas))["8049"] == {"nombre": "TALLER UNO", "email": "a@ejemplo.test"}


def test_nombre_en_el_resumen():
    assert app_mod.nombre_cliente_aviso({"cliente": "8049", "nombre": ""}) == "Cliente 8049"
    assert app_mod.nombre_cliente_aviso({"cliente": "8049", "nombre": "TALLER UNO"}) == "TALLER UNO (8049)"

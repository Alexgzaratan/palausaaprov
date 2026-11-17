"""Referencias del Excel urgente que ya salieron como NO SUMINISTRA y orden de las tarjetas del Comparador."""
from datetime import date

import openpyxl
import pytest

import app_automatizacion_optimizada as app_mod
import avisos_clientes as ac


@pytest.fixture
def registro(tmp_path):
    r = ac.RegistroAvisos(tmp_path / "avisos.json")
    r.anotar([
        {"hoja": "27-10-2026", "cliente": "8049", "tipo": "no_suministra", "referencias": ["1434 C8", "U12498CP"], "cuando": "2026-10-27T17:00:00", "estado": "enviado"},
        {"hoja": "28-10-2026", "cliente": "172553", "tipo": "no_suministra", "referencias": ["1434C8"], "cuando": "2026-10-28T17:00:00", "estado": "borrador"},
        {"hoja": "28-10-2026", "cliente": "8049", "tipo": "sustitucion", "referencias": ["9800 → 9801"], "cuando": "2026-10-28T17:00:00", "estado": "enviado"},
        {"hoja": "01-06-2026", "cliente": "8049", "tipo": "no_suministra", "referencias": ["VIEJA1"], "cuando": "2026-06-01T17:00:00", "estado": "enviado"},
    ])
    return r


def test_coincidencias_con_el_excel_urgente(registro):
    lineas = [["2/1434C8", 2, 8049], ["9800", 1, 8049], ["U12498CP", 1, 8050], ["VIEJA1", 1, 1], ["1434 C8", 5, 9]]
    res = ac.no_suministra_previas(registro, lineas, hoy=date(2026, 11, 3))
    assert [(c["referencia"], c["cantidad"]) for c in res] == [("1434C8", 2), ("U12498CP", 1)]  # sin repetir; la de junio es de hace más de 60 días
    assert res[0]["fechas"] == [date(2026, 10, 27), date(2026, 10, 28)] and res[0]["clientes"] == ["172553", "8049"]
    assert ac.texto_fechas(res[0]["fechas"]) == "27/10 y 28/10" and ac.texto_fechas([date(2026, 1, 2)]) == "02/01"
    assert ac.texto_fechas([date(2026, 10, d) for d in (1, 2, 3)]) == "01/10, 02/10 y 03/10"


def test_dialogo_solo_avisa_y_decide(app, registro):
    res = ac.no_suministra_previas(registro, [["1434C8", 2, 1]], hoy=date(2026, 11, 3))
    decisiones = []
    d = ac.DialogoNoSuministra(app, res, decisiones.append)
    assert list(d.filas) == ["1434C8"] and not hasattr(d, "botones")  # sin ticket: esas piezas ya no se fabrican
    d.decidir(True)
    assert decisiones == [True] and d.cerrado
    d.decidir(False)
    assert decisiones == [True]  # solo se decide una vez


def test_generar_pedidos_avisa_antes_de_generar(app, registro, monkeypatch, tmp_path):
    monkeypatch.setattr(app_mod, "REGISTRO_AVISOS", registro)
    monkeypatch.setattr(ac, "date", type("D", (date,), {"today": staticmethod(lambda: date(2026, 11, 3))}))
    ruta = tmp_path / "URGENTE.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["REF", "D", "CANT", "", "", "CLIENTE"])
    wb.active.append(["2/1434 C8", "x", 3, None, None, 8049])
    wb.save(ruta)
    app.tabview.set("Pedidos")
    app.construir_pestana("Pedidos")
    for campo, valor in ((app.excel_pedidos_entry, str(ruta)), (app.folder_pedidos_entry, str(tmp_path))):
        campo.delete(0, "end")
        campo.insert(0, valor)
    monkeypatch.setattr(app, "excel_abierto", lambda *a: False)
    seguidos = []
    monkeypatch.setattr(app, "continuar_pedidos", lambda ruta, carpeta: seguidos.append(ruta))
    app.iniciar_pedidos()
    dialogo = app.dialogo_no_suministra
    assert seguidos == [] and list(dialogo.filas) == ["1434C8"]  # no se genera nada hasta decidir
    dialogo.decidir(True)
    assert seguidos == [str(ruta)]


def test_tareas_donde_estaba_el_comparador(app):
    app.tabview.set("Comparador")
    app.construir_pestana("Comparador")
    app.update_idletasks()
    tarjetas = [w for w in app.lista_tareas_frame.master.master.winfo_children()]
    orden = sorted(((app.lista_tareas_frame.master, "tareas"), (app.btn_cascos.master.master, "cascos"), (app.btn_comparador.master.master, "comparador")),
                   key=lambda t: (int(t[0].grid_info().get("row", 0)), int(t[0].grid_info().get("column", 0))) if t[0].grid_info() else (t[0].winfo_y(), t[0].winfo_x()))
    assert [nombre for _, nombre in orden] == ["tareas", "cascos", "comparador"]


def plantilla_hoja(ruta):
    """Como HOJA DE PEDIDOS.xlsx: pestañas PSA y FIAT, cabecera en la fila 4, filas 5-50 con formato y B:D combinadas desde la 10."""
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    for i, nombre in enumerate(("PSA", "FIAT")):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = nombre
        ws["B2"], ws["B4"], ws["E4"], ws["F4"], ws["G4"], ws["H4"] = "PEDIDO Nº ", "REFERENCIAS A PEDIR", "ESTADO", "SUSTITUCIÓN", "CANTIDAD", "CLIENTE DMS"
        for fila in range(5, 51):
            ws.cell(fila, 2).fill = PatternFill("solid", fgColor="FFDDEBF7")
            ws.cell(fila, 10).fill = PatternFill("solid", fgColor="FFD9D9D9")
            if 10 <= fila <= 47:
                ws.merge_cells(f"B{fila}:D{fila}")
        lista = DataValidation(type="list", formula1='"FUENLAMOTOR,NO SUMINISTRA,DFS"')
        lista.add("E5:E50")
        ws.add_data_validation(lista)
    wb.save(ruta)
    return ruta


def test_apuntar_en_la_hoja_como_lo_hace_alex(tmp_path):
    ruta = plantilla_hoja(tmp_path / "05-11-2026.xlsx")
    wb = openpyxl.load_workbook(ruta)
    wb["PSA"]["B5"], wb["PSA"]["E5"] = "1427J6", "PLACANORTE"  # una línea que ya había escrito
    wb.save(ruta)
    coincidencias = [{"referencia": "1434C8", "original": "1434 C8", "cantidad": 2, "cliente": 8049},
                     {"referencia": "1427J6", "original": "1427J6", "cantidad": 1, "cliente": "160533"},
                     {"referencia": "U12498CP", "original": "U12498CP", "cantidad": "3", "cliente": "152919.0"}]
    assert app_mod.apuntar_no_suministra_en_hoja(ruta, coincidencias) == (["1434C8", "U12498CP"], ["1427J6"], [])
    ws = openpyxl.load_workbook(ruta)["PSA"]
    assert [ws.cell(6, c).value for c in (2, 5, 7, 8, 10)] == ["1434 C8", "NO SUMINISTRA", 2, 8049, "OK"]
    assert [ws.cell(7, c).value for c in (2, 5, 7, 8, 10)] == ["U12498CP", "NO SUMINISTRA", 3, 152919, "OK"]
    combinadas = {str(m) for m in ws.merged_cells.ranges}
    assert {"B6:D6", "B7:D7", "B10:D10"} <= combinadas and "B5:D5" not in combinadas  # como hace Álex, sin tocar las que no rellena
    assert ws["B6"].fill.start_color.rgb == "FFDDEBF7" and ws.data_validations.dataValidation  # formato y lista de ESTADO intactos
    assert openpyxl.load_workbook(ruta)["FIAT"]["B5"].value is None
    # otra vez lo mismo: no se repite nada
    assert app_mod.apuntar_no_suministra_en_hoja(ruta, coincidencias) == ([], ["1434C8", "1427J6", "U12498CP"], [])


def test_hoja_llena(tmp_path):
    ruta = plantilla_hoja(tmp_path / "h.xlsx")
    wb = openpyxl.load_workbook(ruta)
    for fila in range(5, 51):
        wb["PSA"].cell(fila, 2).value = f"R{fila}"
    wb.save(ruta)
    assert app_mod.apuntar_no_suministra_en_hoja(ruta, [{"referencia": "X1", "cantidad": 1, "cliente": 1}]) == ([], [], ["X1"])


def test_al_decidir_se_apuntan_en_la_hoja_de_hoy(app, registro, monkeypatch, tmp_path):
    monkeypatch.setattr(app_mod, "REGISTRO_AVISOS", registro)
    monkeypatch.setattr(ac, "date", type("D", (date,), {"today": staticmethod(lambda: date(2026, 11, 3))}))
    plantilla = plantilla_hoja(tmp_path / "HOJA DE PEDIDOS.xlsx")
    carpeta = tmp_path / "HOJAS"
    carpeta.mkdir()
    monkeypatch.setitem(app.config, "ruta_carpeta_hojas", str(carpeta))
    monkeypatch.setitem(app.config, "ruta_plantilla_hojas", str(plantilla))
    ruta = tmp_path / "URGENTE.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["REF", "D", "CANT", "", "", "CLIENTE"])
    wb.active.append(["2/1434 C8", "x", 3, None, None, "160533 TALLER DE PRUEBA SL"])
    wb.active.append(["2/9800111", "y", 1, None, None, "8049 OTRO TALLER"])  # esta sí se pide
    wb.save(ruta)
    app.tabview.set("Pedidos")
    app.construir_pestana("Pedidos")
    for campo, valor in ((app.excel_pedidos_entry, str(ruta)), (app.folder_pedidos_entry, str(tmp_path))):
        campo.delete(0, "end")
        campo.insert(0, valor)
    monkeypatch.setattr(app, "excel_abierto", lambda *a: False)
    monkeypatch.setattr(app, "continuar_pedidos", lambda ruta, carpeta: None)
    bloqueada = [True]
    monkeypatch.setattr(app_mod, "archivo_bloqueado", lambda r: bloqueada[0])
    preguntas = []

    def reintentar(*a, **k):  # la persona cierra la hoja en Excel y pulsa «Reintentar»
        preguntas.append(a[1])
        bloqueada[0] = False
        return True

    monkeypatch.setattr(app_mod.messagebox, "askretrycancel", reintentar)
    app.iniciar_pedidos()
    app.dialogo_no_suministra.decidir(False)  # aunque se cancele el pedido, en la hoja quedan apuntadas
    hoja = app_mod.ruta_hoja_pedidos_del_dia(carpeta)
    assert preguntas and "está abierta en Excel" in preguntas[0]
    ws = openpyxl.load_workbook(hoja)["PSA"]
    assert [ws.cell(5, c).value for c in (2, 5, 7, 8, 10)] == ["1434 C8", "NO SUMINISTRA", 3, 160533, "OK"]
    # y sale del Excel urgente: solo queda la que sí se pide, con copia de antes
    urgente = openpyxl.load_workbook(ruta).active
    assert [urgente.cell(f, 1).value for f in range(1, urgente.max_row + 1)] == ["REF", "2/9800111"]


def test_codigo_cliente_sin_nombre():
    assert app_mod.codigo_cliente("160533 TALLER DE PRUEBA SL") == 160533
    assert app_mod.codigo_cliente("152919.0") == 152919 and app_mod.codigo_cliente(8049) == 8049
    assert app_mod.codigo_cliente("") is None and app_mod.codigo_cliente("SIN CODIGO") == "SIN CODIGO"


def test_quitar_espacios_de_la_hoja(tmp_path):
    ruta = plantilla_hoja(tmp_path / "17-11-2026.xlsx")
    wb = openpyxl.load_workbook(ruta)
    wb["PSA"]["B5"], wb["PSA"]["F5"], wb["PSA"]["H5"] = "98 123 456 80", "98 123 456 99 ", "160533"
    wb["PSA"]["B6"], wb["FIAT"]["B5"], wb["FIAT"]["G5"] = "1434C8", " 5180 1234 ", 3
    wb.save(ruta)
    assert app_mod.quitar_espacios_hoja(ruta) == 3
    wb = openpyxl.load_workbook(ruta)
    assert [wb["PSA"]["B5"].value, wb["PSA"]["F5"].value, wb["PSA"]["B6"].value, wb["FIAT"]["B5"].value] == ["9812345680", "9812345699", "1434C8", "51801234"]
    assert wb["PSA"]["H5"].value == "160533" and wb["FIAT"]["G5"].value == 3  # solo referencias
    assert wb["PSA"].data_validations.dataValidation and "B10:D10" in {str(m) for m in wb["PSA"].merged_cells.ranges}
    assert app_mod.quitar_espacios_hoja(ruta) == 0

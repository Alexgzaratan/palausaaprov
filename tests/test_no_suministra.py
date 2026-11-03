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

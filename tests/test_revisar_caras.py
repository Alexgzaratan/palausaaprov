import re

import openpyxl
import pandas as pd
import pytest

import app_automatizacion_optimizada as app_mod
import piezas_caras

# (referencia, pedidas, precio unitario, resultado, media)
FILAS = [
    ("CARA", 1, 1850.0, "Pieza disponible", 0.5),     # unidad > 1600 y quedaría en blanco -> azul
    ("CARA_ROJA", 1, 2400.0, "Pieza disponible", 0.0),  # baja rotación: manda el rojo
    ("CARA_AMARILLA", 1, 1700.0, "Sin stock", 0.5),     # no disponible: manda el amarillo
    ("MUCHAS", 20, 150.0, "Pieza disponible", 0.5),     # total 3000 pero la unidad es barata -> blanco
    ("JUSTO", 1, 1600.0, "Pieza disponible", 0.5),      # igual al límite: no lo supera -> blanco
    ("BARATA", 1, 90.0, "Pieza disponible", 0.5),       # total < precio mínimo -> sin marcar
]


@pytest.fixture
def archivos(tmp_path):
    quiter = tmp_path / "S40P.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["Marca", "Referencia", "Descripción", "Pedidas", "P. compra", "Total"] + [None] * 11 + [None])
    for ref, pedidas, unitario, resultado, _ in FILAS:
        fila = [None] * 18
        fila[0], fila[1], fila[3], fila[4], fila[5], fila[17] = 2, ref, pedidas, unitario, round(pedidas * unitario, 2), resultado
        wb.active.append(fila)
    wb.save(quiter)
    propuesta = tmp_path / "S40PR.xlsx"
    pd.DataFrame({"Articulo": [f[0] for f in FILAS], "Media": [f[4] for f in FILAS]}).to_excel(propuesta, sheet_name="Informe", index=False)
    return quiter, propuesta, tmp_path


def filtrar(quiter, propuesta, carpeta, limite=1600.0):
    registro = []
    app_mod.filtrar_por_rotacion_task(str(quiter), str(propuesta), str(carpeta), 0.12, "FF0000", 150.0,
                                      log_callback=lambda t, tag=None: registro.append(t), limite_revisar=limite)
    return registro, carpeta / "S40P_ROTACION.xlsx"


def test_regla_de_piezas_caras():
    clasificar = app_mod.clasificar_fila_rotacion
    assert clasificar(1850, 0.5, "pieza disponible", 0.12, 150, 1850, 1600) == "revisar"
    assert clasificar(2400, 0.0, "pieza disponible", 0.12, 150, 2400, 1600) == "rojo"
    assert clasificar(1700, 0.5, "sin stock", 0.12, 150, 1700, 1600) == "amarillo"
    assert clasificar(3000, 0.5, "", 0.12, 150, 150, 1600) == "blanco"  # se mira la unidad, no el total
    assert clasificar(1850, 0.5, "", 0.12, 150, 1850, 0) == "blanco"    # 0 = no marcar
    assert clasificar(1850, 0.5, "", 0.12, 150, None, 1600) == "blanco"


def test_el_filtro_pinta_de_azul_solo_las_caras_que_quedarian_en_blanco(archivos):
    quiter, propuesta, carpeta = archivos
    registro, salida = filtrar(quiter, propuesta, carpeta)
    ws = openpyxl.load_workbook(salida).active
    colores = {ws.cell(row=f, column=2).value: (ws.cell(row=f, column=2).fill.start_color.rgb if ws.cell(row=f, column=2).fill.fill_type else None)
               for f in range(1, ws.max_row + 1)}
    assert colores == {"CARA": "009DC3E6", "CARA_ROJA": "00FF0000", "CARA_AMARILLA": "00FFFF00", "MUCHAS": None, "JUSTO": None, "BARATA": None}
    assert "Piezas en azul para revisar (más de 1600 EUR la unidad): 1." in " ".join(registro)
    resumen = next(t for t in registro if "Resumen rotacion" in t)
    assert re.search(r"rojas por umbral 1, amarillas por disponibilidad 1, .* EUR 1", resumen)


def test_sin_limite_no_hay_azules(archivos):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta, limite=0)
    assert app_mod.piezas_caras_rotacion(str(salida)) == []


def test_la_vista_previa_cuenta_las_azules_igual_que_el_filtro(archivos):
    quiter, propuesta, carpeta = archivos
    previsto = app_mod.contar_rotacion(app_mod.leer_filas_quiter(str(quiter)), app_mod.leer_medias_propuesta(str(propuesta)), 0.12, 150.0, 1600.0)
    assert previsto == {"rojo": 1, "amarillo": 1, "revisar": 1, "precio": 1, "blanco": 2}


def test_datos_de_las_piezas_caras(archivos):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    assert app_mod.piezas_caras_rotacion(str(salida)) == [
        {"linea": 1, "referencia": "CARA", "descripcion": "", "pedidas": 1.0, "unitario": 1850.0, "total": 1850.0}]
    marcadas, total = app_mod.lineas_marcadas_rotacion(str(salida))
    assert [ref for _, ref, _ in marcadas] == ["CARA_ROJA", "CARA_AMARILLA"] and total == 6  # la azul no va a poner a 0


def color_de(salida, referencia):
    ws = openpyxl.load_workbook(salida).active
    celda = next(ws.cell(row=f, column=2) for f in range(1, ws.max_row + 1) if ws.cell(row=f, column=2).value == referencia)
    return celda.fill.start_color.rgb if celda.fill.fill_type else None


@pytest.mark.parametrize("decision,color", [(piezas_caras.PEDIR, None), (piezas_caras.A_CERO, "00FF0000")])
def test_decidir_cambia_el_color_en_el_excel(archivos, decision, color):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    resultado = app_mod.decidir_piezas_caras_en_excel(str(salida), {1: decision, 2: piezas_caras.PEDIR})  # la 2 ya es roja: no se toca
    assert resultado == ((1, 0) if decision == piezas_caras.PEDIR else (0, 1))
    assert color_de(salida, "CARA") == color and color_de(salida, "CARA_ROJA") == "00FF0000"
    assert app_mod.piezas_caras_rotacion(str(salida)) == []  # ya no queda nada en azul para la skill
    marcadas, _ = app_mod.lineas_marcadas_rotacion(str(salida))
    assert ("CARA" in [ref for _, ref, _ in marcadas]) == (decision == piezas_caras.A_CERO)


def test_formato_euros():
    assert piezas_caras.euros(1850) == "1.850,00 €"
    assert piezas_caras.euros(12345.5) == "12.345,50 €"
    assert piezas_caras.euros(99.9) == "99,90 €"


@pytest.mark.parametrize("escrito,valor", [("1600", 1600.0), ("1.600", 1600.0), ("1.600,50", 1600.5), ("1600,5", 1600.5),
                                           ("2000 €", 2000.0), ("", 0.0), ("0", 0.0), ("1.5", 1.5), ("mucho", None)])
def test_importe_para_revisar(app, escrito, valor):
    app.limite_revisar_entry.delete(0, "end")
    app.limite_revisar_entry.insert(0, escrito)
    assert app.limite_revisar() == valor


def preparar_campos(app, quiter, carpeta):
    for campo, valor in ((app.excel_quiter_entry, quiter), (app.folder_rot_entry, carpeta), (app.limite_revisar_entry, "1600")):
        campo.delete(0, "end")
        campo.insert(0, str(valor))


def test_ventana_de_decision(app, archivos):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    preparar_campos(app, quiter, carpeta)
    terminado = []
    dialogo = app.decidir_piezas_caras(str(salida), app_mod.piezas_caras_rotacion(str(salida)), al_terminar=terminado.append)
    app.update()
    assert dialogo.ventana.winfo_exists() and dialogo.btn_aplicar.cget("state") == "disabled"  # hasta decidir todas
    dialogo.elegir(1, piezas_caras.A_CERO)
    assert dialogo.btn_aplicar.cget("state") == "normal"
    assert dialogo.botones[1][piezas_caras.A_CERO].cget("fg_color") == app.color_peligro
    assert dialogo.aplicar()
    assert terminado == [True] and not dialogo.ventana.winfo_exists()
    assert color_de(salida, "CARA") == "00FF0000"
    assert app.avisos.historial[-1] == ("ok", "Piezas caras: 0 para pedir y 1 a 0. S40P_ROTACION.xlsx actualizado.")


def test_mas_tarde_deja_la_azul_y_avisa(app, archivos):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    preparar_campos(app, quiter, carpeta)
    terminado = []
    dialogo = app.decidir_piezas_caras(str(salida), app_mod.piezas_caras_rotacion(str(salida)), al_terminar=terminado.append)
    dialogo.mas_tarde()
    assert terminado == [False] and color_de(salida, "CARA") == "009DC3E6"
    aviso = app.avisos.activos[-1]
    assert "Quedan 1 pieza cara en azul" in aviso.texto and aviso.boton_accion.cget("text") == "Decidir"


def test_con_el_excel_abierto_la_ventana_sigue(app, archivos, monkeypatch):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    preparar_campos(app, quiter, carpeta)

    def abierto(*a):
        raise PermissionError("en uso")

    monkeypatch.setattr(app_mod, "decidir_piezas_caras_en_excel", abierto)
    dialogo = app.decidir_piezas_caras(str(salida), app_mod.piezas_caras_rotacion(str(salida)))
    dialogo.todas(piezas_caras.PEDIR)
    assert dialogo.aplicar() is False and dialogo.ventana.winfo_exists() and app.mensajes
    dialogo.mas_tarde()


def test_copiar_lista_pregunta_antes_las_azules(app, archivos):
    quiter, propuesta, carpeta = archivos
    _, salida = filtrar(quiter, propuesta, carpeta)
    preparar_campos(app, quiter, carpeta)
    assert app.copiar_lista_quiter() is None  # no copia con piezas en azul: abre la ventana
    dialogo = app.dialogo_piezas_caras
    dialogo.todas(piezas_caras.PEDIR)
    dialogo.aplicar()
    texto = app.clipboard_get()
    assert texto.splitlines() == ["rotación S40P: 2 de 6 líneas a poner a 0", "Lín. 2 · CARA_ROJA (rojo)", "Lín. 3 · CARA_AMARILLA (amarillo)"]


def test_al_terminar_el_filtro_se_decide_antes_de_contar(app, archivos, monkeypatch):
    quiter, propuesta, carpeta = archivos
    filtrar(quiter, propuesta, carpeta)
    preparar_campos(app, quiter, carpeta)
    capturas = []
    monkeypatch.setattr(app, "mostrar_captura", lambda *datos: capturas.append(datos))
    app._inicio_rotacion = 0
    app._rotacion_del_flujo = True
    app.finalizar_filtro_rotacion()
    assert capturas == []  # espera a la decisión
    app.dialogo_piezas_caras.todas(piezas_caras.A_CERO)
    app.dialogo_piezas_caras.aplicar()
    assert capturas == [(3, 6, "S40P")]  # la cara puesta a 0 ya cuenta
    assert app.etiquetas_ultima_vez["rotacion"].cget("text").endswith("S40P: 3 de 6 líneas")

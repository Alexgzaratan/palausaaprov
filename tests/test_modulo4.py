"""Pedidos todo o nada, avisos sin repetir, clientes sin correo, textos editables e historial por taller."""
import csv
import threading
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

import app_automatizacion_optimizada as app_mod
import avisos_clientes


# ---------- pedidos ----------

def urgente(ruta, lineas):
    wb = openpyxl.Workbook()
    wb.active.append(["REFERENCIA", "DESC", "CANTIDAD", "x", "y", "CLIENTE"])
    for ref, cantidad, cliente in lineas:
        wb.active.append([ref, "d", cantidad, None, None, cliente])
    wb.save(ruta)
    return str(ruta)


@pytest.fixture
def pedidos(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "REGISTRO_PEDIDOS", str(tmp_path / "pedidos_generados.json"))
    monkeypatch.setattr(app_mod, "crear_copia_seguridad", lambda ruta: str(tmp_path / "backup.xlsx"))
    carpeta = tmp_path / "PEDIDOS"
    carpeta.mkdir()
    return tmp_path, carpeta


def generar(ruta, carpeta, stop_event=None):
    registro, numeros = [], []
    app_mod.generar_pedidos_task(ruta, str(carpeta), log_callback=lambda t, tag=None: registro.append(t), stop_event=stop_event, numero_callback=numeros.append)
    return registro, numeros


def csvs(carpeta):
    return sorted(p.name for p in carpeta.rglob("*") if p.is_file())


def test_pedidos_se_generan_y_se_recuerda_el_excel(pedidos):
    tmp, carpeta = pedidos
    ruta = urgente(tmp / "URGENTE.xlsx", [(f"2/REF{i}", 1, 8049) for i in range(150)])
    assert app_mod.pedido_ya_generado(ruta) is None
    lineas = app_mod.leer_lineas_pedido(ruta)
    registro, numeros = generar(ruta, carpeta)
    assert numeros == ["5000", "5001"] and csvs(carpeta) == ["5000.csv", "5001.csv"]
    with open(carpeta / "5000-5099" / "5000.csv", encoding="utf-8") as f:
        primeras = list(csv.reader(f, delimiter=";"))
    assert len(primeras) == 100 and primeras[0] == ["REF0", "1", "8049"]  # sin el «2/»
    assert openpyxl.load_workbook(ruta).active.max_row == 1 and openpyxl.load_workbook(ruta).active["A1"].value is None  # Excel limpio
    # si se vuelve a pegar lo mismo, la app lo reconoce
    urgente(tmp / "URGENTE.xlsx", [(f"2/REF{i}", 1, 8049) for i in range(150)])
    anterior = app_mod.pedido_ya_generado(ruta)
    assert anterior["numeros"] == ["5000", "5001"] and anterior["lineas"] == 150
    assert app_mod.firma_lineas_pedido(lineas) == anterior["firma"]
    assert app_mod.texto_numeros_pedido(anterior["numeros"]) == "5000-5001"


def test_detener_a_mitad_no_deja_pedidos_ni_limpia_el_excel(pedidos, monkeypatch):
    tmp, carpeta = pedidos
    ruta = urgente(tmp / "URGENTE.xlsx", [(f"REF{i}", 1, 8049) for i in range(250)])
    parar = threading.Event()
    original = app_mod.obtener_ruta_para_contador

    def ruta_y_parar(base, contador):  # se pulsa Detener después de escribir el primer CSV
        if contador == 5001:
            parar.set()
        return original(base, contador)

    monkeypatch.setattr(app_mod, "obtener_ruta_para_contador", ruta_y_parar)
    registro, numeros = generar(ruta, carpeta, parar)
    assert numeros == [] and csvs(carpeta) == []  # ni pedidos ni ficheros a medias
    assert openpyxl.load_workbook(ruta).active.max_row == 251
    assert any("no se ha generado ningún pedido" in t for t in registro)
    assert app_mod.pedido_ya_generado(ruta) is None


def test_un_fallo_a_mitad_no_deja_pedidos(pedidos, monkeypatch):
    tmp, carpeta = pedidos
    ruta = urgente(tmp / "URGENTE.xlsx", [(f"REF{i}", 1, 8049) for i in range(250)])
    original = app_mod.obtener_ruta_para_contador

    def falla(base, contador):
        if contador == 5002:
            raise OSError("disco lleno")
        return original(base, contador)

    monkeypatch.setattr(app_mod, "obtener_ruta_para_contador", falla)
    registro, numeros = generar(ruta, carpeta)
    assert numeros == [] and csvs(carpeta) == []
    assert openpyxl.load_workbook(ruta).active.max_row == 251
    # y la numeración sigue donde estaba: el siguiente pedido vuelve a ser el 5000
    monkeypatch.setattr(app_mod, "obtener_ruta_para_contador", original)
    assert generar(ruta, carpeta)[1] == ["5000", "5001", "5002"]


def test_la_app_pregunta_si_el_excel_ya_se_convirtio(app, pedidos, monkeypatch):
    tmp, carpeta = pedidos
    ruta = urgente(tmp / "URGENTE.xlsx", [("REF1", 1, 8049)])
    generar(ruta, carpeta)
    urgente(tmp / "URGENTE.xlsx", [("REF1", 1, 8049)])
    app.tabview.set("Pedidos")
    app.construir_pestana("Pedidos")
    for campo, valor in ((app.excel_pedidos_entry, ruta), (app.folder_pedidos_entry, str(carpeta))):
        campo.delete(0, "end")
        campo.insert(0, valor)
    preguntas = []
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda titulo, texto, **k: preguntas.append(texto) or False)
    monkeypatch.setattr(app, "excel_abierto", lambda *a: False)
    app.iniciar_pedidos()
    assert preguntas and "ya se convirtió en los pedidos 5000 hoy" in preguntas[0] and "se pedirán dos veces" in preguntas[0]
    assert openpyxl.load_workbook(ruta).active.max_row == 2  # no se ha tocado


# ---------- avisos ----------

class Correo:
    def __init__(self, buzon):
        self.buzon, self.To, self.Subject, self.HTMLBody, self.estado = buzon, "", "", "", None
        self.EntryID = f"ID{len(buzon.correos)}"
        buzon.correos.append(self)

    GetInspector = None

    def Send(self):
        self.estado = "enviado"

    def Save(self):
        self.estado = "borrador"


class Outlook:
    def __init__(self):
        self.correos = []

    def CreateItem(self, tipo):
        return Correo(self)

    def GetNamespace(self, nombre):
        return self

    def GetItemFromID(self, ident):
        return next(c for c in self.correos if c.EntryID == ident)

    def enviados(self):
        return [(c.To, c.Subject) for c in self.correos if c.estado == "enviado"]


@pytest.fixture
def outlook(monkeypatch):
    import win32com.client
    falso = Outlook()
    monkeypatch.setattr(win32com.client, "Dispatch", lambda nombre: falso)
    return falso


def hoja_pedido(ruta, lineas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["E2"] = "S99P"
    ws["B4"] = "REFERENCIAS A PEDIR"
    for i, (ref, estado, sustitucion, cliente) in enumerate(lineas, start=5):
        ws.cell(row=i, column=2).value = ref
        ws.cell(row=i, column=5).value = estado
        ws.cell(row=i, column=6).value = sustitucion
        ws.cell(row=i, column=8).value = cliente
    wb.save(ruta)
    return str(ruta)


def clientes_excel(ruta, filas, cabecera=("CLIENTE", "CORREO")):
    wb = openpyxl.Workbook()
    wb.active.append(list(cabecera))
    for fila in filas:
        wb.active.append(list(fila))
    wb.save(ruta)
    return str(ruta)


def avisar(hoja, clientes, registro, textos=None):
    resultados = []
    app_mod.generar_avisos_no_suministro_task(hoja, clientes, enviar_automaticamente=True, log_callback=lambda t, tag=None: None,
                                              resultado_callback=resultados.append, registro=registro, textos=textos)
    return resultados[-1]


def test_repetir_los_avisos_solo_envia_lo_nuevo(tmp_path, outlook):
    registro = avisos_clientes.RegistroAvisos(tmp_path / "avisos.json")
    clientes = clientes_excel(tmp_path / "clientes.xlsx", [(8049, "taller@ejemplo.test"), (8050, "otro@ejemplo.test")])
    lineas = [("A1", "NO SUMINISTRA", None, 8049), ("B1", None, "B2", 8050)]
    hoja = hoja_pedido(tmp_path / "28-10-2026.xlsx", lineas)
    avisar(hoja, clientes, registro)
    assert len(outlook.enviados()) == 2
    assert registro.resumen_hoja("28-10-2026")[0] == 2
    # se repite con la misma hoja y una línea nueva para el 8049: solo sale esa
    hoja_pedido(tmp_path / "28-10-2026.xlsx", lineas + [("A2", "NO SUMINISTRA", None, 8049)])
    resultado = avisar(hoja, clientes, registro)
    assert len(outlook.enviados()) == 3
    nuevo = outlook.correos[-1]
    assert "A2" in nuevo.HTMLBody and "A1" not in nuevo.HTMLBody and nuevo.To == "taller@ejemplo.test"
    assert sorted((a["cliente"], a["omitidas"]) for a in resultado["ya_avisados"]) == [("8049", 1), ("8050", 1)]
    # y una tercera vez no sale nada
    avisar(hoja, clientes, registro)
    assert len(outlook.enviados()) == 3


def test_cliente_que_no_esta_en_el_excel(tmp_path, outlook):
    registro = avisos_clientes.RegistroAvisos(tmp_path / "avisos.json")
    clientes = clientes_excel(tmp_path / "clientes.xlsx", [(8049, "")])
    hoja = hoja_pedido(tmp_path / "28-10-2026.xlsx", [("A1", "NO SUMINISTRA", None, 8049), ("C1", "NO SUMINISTRA", None, 9999)])
    resultado = avisar(hoja, clientes, registro)
    sin = {e["cliente"]: e for e in resultado["no_suministra"]["sin_correo"]}
    assert sin["9999"]["motivo"] == "no_en_excel" and sin["8049"]["motivo"] is None
    assert all(e["borrador_id"] for e in sin.values()) and outlook.enviados() == []
    assert registro.ya_avisadas("28-10-2026", "9999", "no_suministra") == set()  # un borrador no cuenta como avisado


def test_escribir_el_correo_lo_guarda_en_el_excel_y_envia_el_borrador(app, tmp_path, outlook, monkeypatch):
    monkeypatch.setattr(app_mod, "REGISTRO_AVISOS", avisos_clientes.RegistroAvisos(tmp_path / "avisos.json"))
    clientes = clientes_excel(tmp_path / "clientes.xlsx", [(8049, "")])
    hoja = hoja_pedido(tmp_path / "28-10-2026.xlsx", [("A1", "NO SUMINISTRA", None, 8049), ("C1", "NO SUMINISTRA", None, 9999)])
    resultado = avisar(hoja, clientes, app_mod.REGISTRO_AVISOS)
    app.tabview.set("Pedidos")
    app.construir_pestana("Pedidos")
    app.clientes_avisos_entry.delete(0, "end")
    app.clientes_avisos_entry.insert(0, clientes)
    lanzados = []
    monkeypatch.setattr(app_mod.threading, "Thread", lambda target, args=(), daemon=None: type("H", (), {"start": lambda self: lanzados.append(target(*args))})())
    dialogo = app.pedir_correos_clientes(resultado["no_suministra"]["sin_correo"], True)
    assert set(dialogo.campos) == {"8049", "9999"}
    dialogo.campos["9999"].insert(0, "nuevo@ejemplo.test")
    dialogo.campos["8049"].insert(0, "mal-escrito")
    assert dialogo.guardar() is False  # correo no válido: no se cierra
    dialogo.campos["8049"].delete(0, "end")
    dialogo.campos["8049"].insert(0, "taller@ejemplo.test")
    assert dialogo.guardar()
    assert sorted(outlook.enviados()) == [("nuevo@ejemplo.test", "Aviso referencia no suministrada 28-10-2026"), ("taller@ejemplo.test", "Aviso referencia no suministrada 28-10-2026")]
    assert len(outlook.correos) == 2  # se han enviado los borradores, no correos nuevos
    assert app_mod.cargar_clientes_excel(clientes) == {"8049": {"nombre": "", "email": "taller@ejemplo.test"}, "9999": {"nombre": "", "email": "nuevo@ejemplo.test"}}
    assert app_mod.REGISTRO_AVISOS.ya_avisadas("28-10-2026", "9999", "no_suministra") == {"C1"}


def test_guardar_correos_en_excel_de_tres_columnas(tmp_path):
    ruta = clientes_excel(tmp_path / "c.xlsx", [(8049, "TALLER UNO", "")], cabecera=("CLIENTE", "NOMBRE", "CORREO"))
    assert avisos_clientes.guardar_correos_clientes(ruta, {"8049": "a@ejemplo.test", "7000": "b@ejemplo.test"}, app_mod._normalizar_cliente_dms) == (1, 1)
    assert app_mod.cargar_clientes_excel(ruta) == {"8049": {"nombre": "TALLER UNO", "email": "a@ejemplo.test"}, "7000": {"nombre": "", "email": "b@ejemplo.test"}}


# ---------- textos ----------

def test_textos_editables(tmp_path, outlook):
    assert avisos_clientes.comprobar_texto("no_suministra_una", "Hola {referencia} del {fecha}") is None
    assert "no existe el hueco {referencias}" in avisos_clientes.comprobar_texto("no_suministra_una", "Hola {referencias}")
    assert "sin cerrar" in avisos_clientes.comprobar_texto("asunto_sustitucion", "Aviso {fecha")
    assert "vacío" in avisos_clientes.comprobar_texto("asunto_sustitucion", "  ")
    propios = avisos_clientes.textos({"textos_avisos": {"no_suministra_una": "Buenos días. La {referencia} ({fecha}) no llega.", "asunto_no_suministra": "Pieza {fecha}"}})
    assert propios["sustitucion_una"] == avisos_clientes.TEXTOS_POR_DEFECTO["sustitucion_una"]
    registro = avisos_clientes.RegistroAvisos(tmp_path / "avisos.json")
    clientes = clientes_excel(tmp_path / "clientes.xlsx", [(8049, "taller@ejemplo.test")])
    avisar(hoja_pedido(tmp_path / "28-10-2026.xlsx", [("A1", "NO SUMINISTRA", None, 8049)]), clientes, registro, propios)
    correo = outlook.correos[-1]
    assert correo.Subject == "Pieza 28-10-2026" and "Buenos días. La A1 (28-10-2026) no llega." in correo.HTMLBody


def test_ventana_de_textos_guarda_solo_lo_cambiado(app, monkeypatch):
    guardado = []
    monkeypatch.setattr(app_mod, "guardar_configuracion", lambda datos: guardado.append(dict(datos)))
    dialogo = app.editar_textos_avisos()
    dialogo.campos["asunto_no_suministra"].delete(0, "end")
    dialogo.campos["asunto_no_suministra"].insert(0, "Pieza {fech}")
    assert dialogo.guardar() is False and not guardado  # hueco mal escrito
    dialogo.campos["asunto_no_suministra"].delete(0, "end")
    dialogo.campos["asunto_no_suministra"].insert(0, "Pieza {fecha}")
    assert dialogo.guardar()
    assert guardado[-1]["textos_avisos"] == {"asunto_no_suministra": "Pieza {fecha}"}
    assert avisos_clientes.textos(app.config)["asunto_no_suministra"] == "Pieza {fecha}"


# ---------- historial ----------

def test_historial_por_taller(tmp_path, app):
    registro = avisos_clientes.RegistroAvisos(tmp_path / "avisos.json")
    registro.anotar([
        {"hoja": "27-10-2026", "cliente": "8049", "nombre": "TALLER", "tipo": "no_suministra", "referencias": ["A1", "A2"], "cuando": "2026-10-27T17:00:00", "estado": "enviado"},
        {"hoja": "28-10-2026", "cliente": "8049", "nombre": "TALLER", "tipo": "sustitucion", "referencias": ["B1 → B2"], "cuando": "2026-10-28T17:00:00", "estado": "enviado"},
        {"hoja": "28-10-2026", "cliente": "8050", "nombre": "", "tipo": "no_suministra", "referencias": ["C1"], "cuando": "2026-10-28T17:00:00", "estado": "enviado"},
        {"hoja": "28-10-2026", "cliente": "8051", "nombre": "", "tipo": "no_suministra", "referencias": ["D1"], "cuando": "2026-10-28T17:00:00", "estado": "borrador"},
        {"hoja": "30-09-2026", "cliente": "8050", "nombre": "", "tipo": "no_suministra", "referencias": ["E1"], "cuando": "2026-09-30T17:00:00", "estado": "enviado"},
    ])
    filas = registro.historial_mes(2026, 10)
    assert [(f["cliente"], f["avisos"], f["referencias"], f["no_suministra"], f["sustituciones"]) for f in filas] == [("8049", 2, 3, 2, 1), ("8050", 1, 1, 1, 0)]
    assert filas[0]["ultimo"] == datetime(2026, 10, 28, 17, 0)
    assert registro.referencias_mes("8050", 2026, 9) == 1 and registro.meses() == [(2026, 10), (2026, 9)]
    dialogo = avisos_clientes.DialogoHistorialAvisos(app, registro, hoy=datetime(2026, 10, 29))
    assert dialogo.total.cget("text") == "2 talleres · 3 correos · 4 referencias"
    dialogo.selector.set("septiembre 2026")
    dialogo.mostrar()
    assert dialogo.total.cget("text") == "1 taller · 1 correo · 1 referencia"
    dialogo.ventana.destroy()

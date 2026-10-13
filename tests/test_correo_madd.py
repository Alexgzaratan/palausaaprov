import time
from datetime import date, datetime
from types import SimpleNamespace

import pytest

import app_automatizacion_optimizada as app_mod


class FakeOutbox:
    def __init__(self, items):
        self.Items = items


class FakeNamespace:
    def __init__(self, items):
        self.items = items
        self.sincronizado = False
        self.SyncObjects = SimpleNamespace(Item=lambda n: SimpleNamespace(Start=self._sync))

    def _sync(self):
        self.sincronizado = True

    def GetDefaultFolder(self, tipo):
        assert tipo == 4  # bandeja de salida
        return FakeOutbox(self.items)


class FakeOutlook:
    def __init__(self, items):
        self.ns = FakeNamespace(items)

    def GetNamespace(self, nombre):
        return self.ns


def correo(asunto, creado):
    return SimpleNamespace(Subject=asunto, CreationTime=datetime.fromtimestamp(creado))


# ---------- detectar Outlook ----------

def test_outlook_en_ejecucion(monkeypatch):
    monkeypatch.setattr(app_mod.subprocess, "run", lambda *a, **k: SimpleNamespace(stdout=b"OUTLOOK.EXE   1234 Console  2  100.000 KB"))
    assert app_mod.outlook_en_ejecucion() is True
    monkeypatch.setattr(app_mod.subprocess, "run", lambda *a, **k: SimpleNamespace(stdout=b"INFORMACION: no hay tareas ejecutandose"))
    assert app_mod.outlook_en_ejecucion() is False


# ---------- abrir Outlook ----------

def test_si_outlook_esta_cerrado_se_abre_y_se_espera_a_que_responda(monkeypatch):
    abierto = []
    intentos = {"n": 0}
    monkeypatch.setattr(app_mod, "outlook_en_ejecucion", lambda: False)
    monkeypatch.setattr(app_mod.os, "startfile", lambda ruta: abierto.append(ruta), raising=False)
    monkeypatch.setattr(app_mod.time, "sleep", lambda s: None)
    outlook = FakeOutlook([])

    def dispatch(nombre):
        intentos["n"] += 1
        if intentos["n"] < 3:
            raise OSError("Outlook todavía arrancando")
        return outlook

    import win32com.client
    monkeypatch.setattr(win32com.client, "Dispatch", dispatch)
    registro = []
    resultado = app_mod.asegurar_outlook_abierto(lambda t, tag=None: registro.append(t))
    assert resultado is outlook
    assert abierto == ["outlook.exe"]
    assert intentos["n"] == 3
    assert any("cerrado" in t for t in registro)


def test_si_outlook_ya_esta_abierto_no_se_vuelve_a_abrir(monkeypatch):
    abierto = []
    monkeypatch.setattr(app_mod, "outlook_en_ejecucion", lambda: True)
    monkeypatch.setattr(app_mod.os, "startfile", lambda ruta: abierto.append(ruta), raising=False)
    import win32com.client
    monkeypatch.setattr(win32com.client, "Dispatch", lambda nombre: FakeOutlook([]))
    app_mod.asegurar_outlook_abierto(lambda *a, **k: None)
    assert abierto == []


def test_si_outlook_no_responde_se_lanza_un_error_claro(monkeypatch):
    monkeypatch.setattr(app_mod, "outlook_en_ejecucion", lambda: True)
    monkeypatch.setattr(app_mod.time, "sleep", lambda s: None)
    reloj = iter(range(0, 1000, 50))
    monkeypatch.setattr(app_mod.time, "time", lambda: next(reloj))
    import win32com.client

    def dispatch(nombre):
        raise OSError("no responde")

    monkeypatch.setattr(win32com.client, "Dispatch", dispatch)
    with pytest.raises(RuntimeError, match="no respondió"):
        app_mod.asegurar_outlook_abierto(lambda *a, **k: None, espera_max=100)


# ---------- esperar a que salga el correo ----------

def test_el_correo_sale_cuando_desaparece_de_la_bandeja_de_salida(monkeypatch):
    ahora = time.time()
    items = [correo("Facturas Stellantis del 2026-10-13", ahora), correo("Otro asunto", ahora)]
    outlook = FakeOutlook(items)
    pausas = []

    def dormir(segundos):
        pausas.append(segundos)
        if len(pausas) == 2:
            items.pop(0)  # Outlook lo envía

    monkeypatch.setattr(app_mod.time, "sleep", dormir)
    assert app_mod.esperar_salida_de_correo(outlook, "Facturas Stellantis del 2026-10-13", ahora) is True
    assert len(pausas) == 2


def test_un_correo_antiguo_con_el_mismo_asunto_no_cuenta(monkeypatch):
    ahora = time.time()
    outlook = FakeOutlook([correo("Facturas Stellantis del 2026-10-13", ahora - 3600)])
    monkeypatch.setattr(app_mod.time, "sleep", lambda s: None)
    assert app_mod.esperar_salida_de_correo(outlook, "Facturas Stellantis del 2026-10-13", ahora) is True


def test_si_no_sale_se_empuja_el_envio_y_se_devuelve_pendiente(monkeypatch):
    ahora = time.time()
    outlook = FakeOutlook([correo("Facturas", ahora)])
    reloj = {"t": ahora}
    monkeypatch.setattr(app_mod.time, "time", lambda: reloj["t"])
    monkeypatch.setattr(app_mod.time, "sleep", lambda s: reloj.__setitem__("t", reloj["t"] + s))
    assert app_mod.esperar_salida_de_correo(outlook, "Facturas", ahora, espera_max=60, intervalo=5) is False
    assert outlook.ns.sincronizado is True


# ---------- registro y flujo ----------

def test_el_estado_del_correo_queda_en_el_registro(tmp_path, monkeypatch):
    ruta = tmp_path / "log.txt"
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(ruta))
    app_mod.registrar_estado_correo_madd(date(2026, 10, 13), "PENDIENTE en la bandeja de salida de Outlook", "Destinatarios: a@x.es")
    texto = ruta.read_text(encoding="utf-8")
    assert "Fecha objetivo: 2026-10-13" in texto and "PENDIENTE" in texto and "a@x.es" in texto


@pytest.mark.parametrize("estado,esperado", [("enviado", "Estado: enviado"), ("pendiente", "Estado: PENDIENTE")])
def test_el_correo_de_facturas_registra_si_salio_o_no(tmp_path, monkeypatch, estado, esperado):
    ruta = tmp_path / "log.txt"
    monkeypatch.setattr(app_mod, "LOG_FACTURAS_MADD", str(ruta))
    capturado = {}

    def falso(destinatarios, asunto, cuerpo, adjuntos, log):
        capturado.update(destinatarios=destinatarios, asunto=asunto, adjuntos=adjuntos)
        return estado

    monkeypatch.setattr(app_mod, "enviar_con_outlook_abierto", falso)
    app_mod._madd_enviar_correo_resultado("a@x.es", date(2026, 10, 13), ["f1.pdf", "f2.pdf"], lambda *a, **k: None)
    assert capturado["asunto"] == "Facturas Stellantis del 2026-10-13" and capturado["adjuntos"] == ["f1.pdf", "f2.pdf"]
    assert esperado in ruta.read_text(encoding="utf-8")

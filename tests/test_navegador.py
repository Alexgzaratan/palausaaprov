"""Ventanas visibles en la app compilada: si falta el Chromium de Playwright, Edge o Chrome del sistema."""
import pytest

import navegador


class Chromium:
    def __init__(self, faltan):
        self.faltan, self.llamadas = faltan, []

    def launch(self, headless, channel=None):
        self.llamadas.append(channel)
        if channel in self.faltan:
            raise Exception("BrowserType.launch: Executable doesn't exist at chrome.exe")
        return f"navegador {channel}"


class P:
    def __init__(self, faltan):
        self.chromium = Chromium(faltan)


def test_usa_el_incluido_si_esta():
    p = P(faltan=())
    assert navegador.lanzar(p, True) == "navegador None" and p.chromium.llamadas == [None]


def test_sin_chromium_incluido_abre_edge():
    p = P(faltan=(None,))
    assert navegador.lanzar(p, True) == "navegador msedge" and p.chromium.llamadas == [None, "msedge"]


def test_sin_edge_abre_chrome_y_sin_ninguno_avisa():
    assert navegador.lanzar(P(faltan=(None, "msedge")), True) == "navegador chrome"
    with pytest.raises(RuntimeError, match="instala Microsoft Edge o Google Chrome"):
        navegador.lanzar(P(faltan=(None, "msedge", "chrome")), True)


def test_otros_errores_no_se_esconden():
    class Roto(Chromium):
        def launch(self, headless, channel=None):
            raise Exception("Target closed")
    p = P(())
    p.chromium = Roto(())
    with pytest.raises(Exception, match="Target closed"):
        navegador.lanzar(p, True)


def test_comprobar_ventanas(tmp_path):
    class ConRuta(Chromium):
        executable_path = str(tmp_path / "no_existe.exe")

        def launch(self, headless, channel=None):
            navegador_falso = super().launch(headless, channel)
            return type("N", (), {"close": lambda self: None, "nombre": navegador_falso})()
    p = P(())
    p.chromium = ConRuta(faltan=("msedge",))
    assert navegador.comprobar_ventanas(p) == "chrome"
    (tmp_path / "no_existe.exe").write_bytes(b"")
    assert navegador.comprobar_ventanas(p) == "Chromium incluido"

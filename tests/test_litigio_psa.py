"""Litigio de calidad en PSA Parts contra un portal falso (SAP ITS simplificado) y los tres pasos del formulario."""
import base64
import json
import threading
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import pytest

import litigio_psa as lp
import reclamaciones_power_line as pl

MENU = """<html><body>
<a href="#" onclick="sub('fact');return false">FACTURA</a> <a href="#" onclick="sub('lit');return false">LITIGIO</a>
<div id="fact" style="display:none"><a href="#" onclick="lista();return false">Lista de líneas de facturas</a></div>
<div id="lit" style="display:none"><a href="#" onclick="crear();return false">Crear litigio con referencia a fact.</a></div>
<div id="vista"></div>
<script>
var LINEAS = [['1018', '0382 NL', 2], ['1019', '16 100 178 80', 1]];
function sub(id){ document.getElementById(id).style.display = 'block'; }
function v(){ return document.getElementById('vista'); }
function lista(){
  v().innerHTML = '<input name="p_matnr[1]"><input name="date_f[1]" value="01.10.2026"><input name="date_t[1]" value="31.10.2026">'
    + '<a id="keyboard_ok" href="#" onclick="ejecutar();return false">Ejecutar programa</a><div id="res"></div>';
}
function ejecutar(){
  var mat = document.querySelector('[name="p_matnr[1]"]').value, desde = document.querySelector('[name="date_f[1]"]').value;
  fetch('/lineas?mat=' + encodeURIComponent(mat) + '&desde=' + desde).then(r => r.json()).then(filas => {
    var h = '<table><tr><td>FechaFact.</td><td>ClVt</td><td>N° pedido cliente</td><td>Doc.ventas</td><td>Pos.</td><td>Material</td><td>Factura</td><td>Pos.</td><td>Ctd.fact.</td></tr>';
    filas.forEach(f => { h += '<tr>' + f.map(c => '<td>' + c + '</td>').join('') + '</tr>'; });
    document.getElementById('res').innerHTML = h + '</table>'; });
}
function crear(){
  v().innerHTML = '<input name="v_vbeln[1]"><input type="radio" name="%%RB" value="v_fact_nr[1]" checked><input type="radio" name="%%RB" value="v_qualite[1]">'
    + '<a href="#" onclick="siguiente();return false">Siguiente</a>';
}
function siguiente(){
  var calidad = document.querySelector('[value="v_qualite[1]"]').checked;
  window.factura = document.querySelector('[name="v_vbeln[1]"]').value;
  var ops = calidad ? [['Q01','PIEZA DETERIORADA'],['Q05','DEFECTO TECNICO'],['N05','ERROR DE CATALOGO PIEZAS CODIFICADAS']] : [['F01','NO RECIBIDO']];
  v().innerHTML = '<p><b><font color="red">Atencion!!! La seleccion del documento rebasa las 00100 lineas visibles</font></b></p><div id="msg"></div>'
    + '<select name="vbak-augru[1]"><option value=""></option>' + ops.map(o => '<option value="' + o[0] + '">' + o[1] + '</option>').join('') + '</select>'
    + '<a href="#" onclick="buscar();return false">Buscar</a><input name="vbak-bstnk[1]"><input name="v_matnr[1]">'
    + '<a href="#" onclick="control();return false">Control</a><a href="#" onclick="guardar();return false">Guardar</a><table id="lineas"></table>';
  pintar(LINEAS);
}
function pintar(filas){
  document.getElementById('lineas').innerHTML = '<tr><td>Pos.</td><td>Material</td><td>Ctd</td><td>Cantid litigio</td><td>Status</td></tr>'
    + filas.map((f, i) => '<tr><td>' + f[0] + '</td><td>' + f[1] + '</td><td>' + f[2] + '</td><td><input name="t_vbrp-zmeng_w[' + (i + 1) + ']"></td><td id="st' + f[0] + '"></td></tr>').join('');
}
function buscar(){
  var m = document.querySelector('[name="v_matnr[1]"]').value.replace(/\\s/g, '');
  pintar(LINEAS.slice().sort((a, b) => (b[1].replace(/\\s/g, '') === m) - (a[1].replace(/\\s/g, '') === m)));
  document.querySelector('[name="v_matnr[1]"]').value = '';
}
function marcadas(){ return [...document.querySelectorAll('#lineas tr')].slice(1).map(tr => [tr.cells[0].innerText, +tr.cells[2].innerText, tr.querySelector('input').value]).filter(f => f[2]); }
function control(){
  marcadas().forEach(f => { if (+f[2] > f[1]) { document.getElementById('st' + f[0]).innerText = 'E';
    document.getElementById('msg').innerHTML = '<font color="red">Cantidad en litigio superior a la facturada</font>'; } });
  location.hash = 'control' + Date.now();
}
function guardar(){
  fetch('/guardar', {method: 'POST', body: JSON.stringify({factura: window.factura, motivo: document.querySelector('select').value, lineas: marcadas()})})
    .then(r => r.text()).then(n => { v().innerHTML = '<p>Litig calid S/retorn ' + n + ' se ha grabado</p><p>I Aviso informativo de prueba</p>'; });
}
</script></body></html>"""

FACTURAS = [  # FechaFact, ClVt, pedido cliente, Doc.ventas, Pos., Material, Factura, Pos., Ctd
    ["21.10.2026", "ZSTO", "S38P", "2063765885", "10", "0382 NL", "1048893677", "1018", "2"],
    ["05.10.2026", "ZSTO", "S36P", "2063700001", "20", "0382 NL", "1048800001", "0040", "1"],
]


class Portal:
    def __init__(self):
        self.guardados, self.busquedas = [], []
        portal = self

        class Manejador(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def responder(self, cuerpo, tipo="text/html; charset=utf-8", estado=200, cabeceras=()):
                datos = cuerpo.encode("utf-8")
                self.send_response(estado)
                self.send_header("Content-Type", tipo)
                self.send_header("Content-Length", str(len(datos)))
                for k, v in cabeceras:
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(datos)

            def autorizado(self):
                return self.headers.get("Authorization") == "Basic " + base64.b64encode(b"u:p").decode()

            def do_GET(self):
                if not self.autorizado():
                    return self.responder("no", estado=401, cabeceras=[("WWW-Authenticate", 'Basic realm="psa"')])
                url = urlparse(self.path)
                if url.path.startswith("/lineas"):
                    q = parse_qs(url.query)
                    portal.busquedas.append(q)
                    mat = q.get("mat", [""])[0].replace(" ", "")
                    return self.responder(json.dumps([f for f in FACTURAS if f[5].replace(" ", "") == mat]), "application/json")
                if "zmenu" in url.path:
                    return self.responder(MENU)
                return self.responder("<html><body>sesion</body></html>")

            def do_POST(self):
                cuerpo = self.rfile.read(int(self.headers.get("Content-Length") or 0)).decode("utf-8")
                portal.guardados.append(json.loads(cuerpo))
                return self.responder("131172099", "text/plain")

        self.servidor = ThreadingHTTPServer(("127.0.0.1", 0), Manejador)
        self.servidor.daemon_threads = True
        threading.Thread(target=self.servidor.serve_forever, daemon=True).start()
        self.url = f"http://127.0.0.1:{self.servidor.server_port}"


@pytest.fixture
def portal():
    p = Portal()
    yield p
    p.servidor.shutdown()


def litigio(portal, **kw):
    kw.setdefault("hoy", date(2026, 10, 29))
    registro = []
    return lp.LitigioPSA("u", "p", kw.pop("referencia", "0382NL"), url=portal.url, log=lambda t, tag=None: registro.append(t), **kw), registro


def test_crea_el_litigio_con_la_factura_mas_reciente(portal):
    vistos = []
    t, registro = litigio(portal, confirmar=lambda datos: vistos.append(datos) or True)
    resultado = t.ejecutar()
    assert resultado == {"litigio": "131172099", "doc_ventas": "2063765885", "factura": "1048893677", "pos": "1018",
                         "fecha_factura": date(2026, 10, 21), "avisos": ["I Aviso informativo de prueba"]}
    assert portal.guardados == [{"factura": "1048893677", "motivo": "Q01", "lineas": [["1018", 2, "1"]]}]
    assert portal.busquedas[0]["desde"] == ["01.09.2026"]  # día 1 del mes anterior
    assert vistos[0]["motivo"] == "Q01 · PIEZA DETERIORADA" and vistos[0]["cantidad"] == 1


def test_sin_confirmar_no_se_guarda(portal):
    t, _ = litigio(portal, confirmar=lambda datos: False)
    with pytest.raises(lp.LitigioNoCreado, match="cancelado"):
        t.ejecutar()
    assert portal.guardados == []


def test_otro_motivo_y_cantidad(portal):
    t, _ = litigio(portal, motivo="Q05", cantidad=2, confirmar=lambda d: True)
    t.ejecutar()
    assert portal.guardados[0]["motivo"] == "Q05" and portal.guardados[0]["lineas"] == [["1018", 2, "2"]]


def test_control_con_errores_no_guarda(portal):
    t, _ = litigio(portal, cantidad=5, confirmar=lambda d: True)
    with pytest.raises(RuntimeError, match="Cantidad en litigio superior"):
        t.ejecutar()
    assert portal.guardados == []


def test_sin_factura_reciente_avisa_y_no_crea_nada(portal):
    t, _ = litigio(portal, hoy=date(2026, 11, 20), confirmar=lambda d: True)
    with pytest.raises(lp.LitigioNoCreado, match="últimos 10 días.*21/10 factura 1048893677"):
        t.ejecutar()
    t, _ = litigio(portal, referencia="9999ZZ", confirmar=lambda d: True)
    with pytest.raises(lp.LitigioNoCreado, match="no hay facturas de 9999ZZ"):
        t.ejecutar()
    assert portal.guardados == []


def test_reglas():
    lineas = [{"fecha": date(2026, 10, 21), "material": "0382 NL", "factura": "1", "pos_factura": "10", "doc_ventas": "A"},
              {"fecha": date(2026, 10, 25), "material": "0382 NL", "factura": "2", "pos_factura": "20", "doc_ventas": "B"},
              {"fecha": date(2026, 10, 28), "material": "1111 AA", "factura": "3", "pos_factura": "30", "doc_ventas": "C"}]
    assert lp.elegir_linea(lineas, "2/0382NL", date(2026, 10, 29))["factura"] == "2"
    assert lp.elegir_linea(lineas, "0382NL", date(2026, 11, 10)) is None  # más de 10 días
    assert lp.desde_para_busqueda(date(2026, 1, 15)) == date(2025, 12, 1)
    assert lp.codigo_motivo("Q04 · MAL EMBASADA ETIQ NO CORRESPONDE C/RECIB") == "Q04" and lp.texto_motivo("Q01") == "Q01 · PIEZA DETERIORADA"
    assert [c for c, _ in lp.MOTIVOS] == ["Q01", "Q02", "Q03", "Q04", "Q05", "N05"]


def test_no_se_reclama_dos_veces_la_misma_linea(tmp_path):
    ruta = str(tmp_path / "litigios.json")
    pl.anotar_litigio_psa({"litigio": "131172099", "factura": "1048893677", "pos": "1018", "doc_ventas": "2063765885"}, "0382NL", ruta)
    assert pl.litigio_anterior("1048893677", "1018", ruta)["litigio"] == "131172099"
    assert pl.litigio_anterior("1048893677", "1019", ruta) is None


# ---------- los tres pasos en el formulario ----------

@pytest.fixture
def formulario(app, monkeypatch, tmp_path):
    monkeypatch.setattr(pl, "LITIGIOS_PSA", str(tmp_path / "litigios.json"))
    app.tabview.set("Reclamaciones")
    app.construir_pestana("Reclamaciones")
    modulo = app.reclamaciones_power_line
    modulo.configuracion.update({"usuario": "u", "password": "p"})
    return modulo


def test_paso_1_rellena_litigio_y_pedido(formulario, monkeypatch):
    creados = []

    class Falso:
        def __init__(self, usuario, password, referencia, motivo, cantidad, confirmar=None, log=None):
            creados.append((referencia, motivo, cantidad))

        def ejecutar(self):
            return {"litigio": "131172099", "doc_ventas": "2063765885", "factura": "1048893677", "pos": "1018", "fecha_factura": date(2026, 10, 21), "avisos": []}

    monkeypatch.setattr(pl.litigio_psa, "LitigioPSA", Falso)
    monkeypatch.setattr(pl.threading, "Thread", lambda target, args=(), daemon=None: type("H", (), {"start": lambda self: target(*args)})())
    monkeypatch.setattr(formulario.app, "after", lambda ms, funcion, *a: funcion(*a))
    formulario.referencia_entry.insert(0, "2/0382 NL")
    formulario.motivo_option.set(lp.texto_motivo("Q05"))
    assert formulario.crear_litigio_psa()
    assert creados == [("0382NL", "Q05", 1)]
    assert formulario.litigio_entry.get() == "131172099" and formulario.pedido_entry.get() == "2063765885" and formulario.referencia_entry.get() == "0382NL"
    assert pl.litigio_anterior("1048893677", "1018")["litigio"] == "131172099"


def test_paso_3_solo_tras_generar_el_documento(formulario, monkeypatch, tmp_path):
    assert formulario.btn_paso3.cget("state") == "disabled"
    monkeypatch.setattr(pl.messagebox, "showinfo", lambda *a, **k: None)
    word = tmp_path / "Litigio 131172099.docx"
    word.write_bytes(b"w")
    formulario._finalizar_generacion(str(word), None, 7, "131172099")
    assert formulario.btn_paso3.cget("state") == "normal"
    enviados = []
    monkeypatch.setattr(formulario, "enviar_a_power_line", lambda *a: enviados.append(a) or True)
    assert formulario.enviar_ultimo_a_power_line() and enviados == [(7, "131172099", str(word))]

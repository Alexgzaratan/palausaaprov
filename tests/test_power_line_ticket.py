"""Ticket de Power-Line contra un portal falso en localhost que imita al real: acceso en dos pasos, buscadores
(jQuery UI autocomplete), criterio «Número de reclamacion», «Enviar», Nº TKT y adjuntos con el clip."""
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pytest

import power_line_ticket as plt
import reclamaciones_power_line as pl

LOGIN = """<html><body><div id="paso">
<input id="identifierInput" name="subject"><a id="signOnButton" href="#" onclick="paso2();return false">Next</a></div>
<script>
function paso2(){ var u=document.getElementById('identifierInput').value;
  document.getElementById('paso').innerHTML='<input id="username" value="'+u+'"><input id="password" type="password"><a id="signOnButton" href="#" onclick="entrar();return false">Sign On</a>'; }
function entrar(){ var c=document.getElementById('password').value;
  fetch('/login', {method:'POST', body: JSON.stringify({u: document.getElementById('username').value, c: c})}).then(r => r.text()).then(t => {
    if (t === 'ok') location.href = '/app.html'; else { portalIntentos = 1; document.getElementById('password').value = ''; document.body.insertAdjacentHTML('beforeend', '<p class=\"ping-error\">Credenciales incorrectas</p>'); } }); }
</script></body></html>"""

APP = """<html><body><div>Pagina principal · Servicio de asistencia técnica</div><div id="vista"></div>
<ul class="ui-autocomplete" style="display:none"></ul>
<div class="ui-dialog" id="adjuntos" style="display:none"><span>Attachments</span><input type="file" class="upload-input hide" multiple style="display:none">
  <div id="lista"></div><button onclick="document.getElementById('adjuntos').style.display='none'">Cerrar</button></div>
<script>
var OPCIONES = {
  principal: [['123456X', '123456X - EMPRESA PRUEBA, SL']],
  shipToParty: [['123456X01', '123456X01 - EMPRESA PRUEBA, SL'], ['123456X03', '123456X03 - EMPRESA PRUEBA, SL']],
  ticketTheme: [['Litiges', 'Reclamaciones / Devoluciones'], ['Litiges', 'Litiges/retours (OPEL)'], ['Commandes', 'Pedidos'], ['Delais', 'Fecha estimada de entrega']],
  ticketType: []
};
// tipos de cada tema y criterios de cada tipo, como en el portal real
var TIPOS = {
  'Reclamaciones / Devoluciones': ['Anulacion de una reclamacion por entrega posterior', 'Progreso / Seguimiento de una reclamacion'],
  'Pedidos': ['Cancelacion pedido pieza codificada', 'Cancelacion linea de pedido o pedido completo', 'Motivo de cancelacion de un pedido de pieza codificada'],
  'Fecha estimada de entrega': ['Solicitud de plazo de entrega para un aviso con prioridad 1,2,3', 'Solicitud de plazo para un faltante sin aviso de reclamamción']
};
var CRITERIOS = {
  'Progreso / Seguimiento de una reclamacion': ['Número de reclamacion'],
  'Motivo de cancelacion de un pedido de pieza codificada': ['Número de pedido de SAP', 'Referencia'],
  'Cancelacion linea de pedido o pedido completo': ['Número de pedido de SAP', 'Referencia', 'Cantidad', 'Número de linea de pedido de SAP'],
  'Solicitud de plazo para un faltante sin aviso de reclamamción': ['Número de pedido de SAP', 'Referencia', 'Cantidad']
};
function campo(nombre){ return '<div x-field="'+nombre+'"><input oninput="buscar(this, \\''+nombre+'\\')" onkeydown="if (event.key === \\'ArrowDown\\') buscar(this, \\''+nombre+'\\', true)"></div>'; }
function vista(){
  var v = document.getElementById('vista');
  if (location.hash === '#/ds/helpdesk.ticket/list/1') v.innerHTML = '<button onclick="formulario()"><i class="fa fa-plus"></i></button>';
}
function formulario(){
  document.getElementById('vista').innerHTML = '<div x-field="ticketSeq">Nº Ticket</div>' + campo('principal') + campo('shipToParty') + campo('ticketTheme') + campo('ticketType')
    + '<div id="criterios"></div><div x-field="description"><div class="wysiwyg-editor" contenteditable="true"></div></div>'
    + '<a x-field="sendTicket" class="btn" href="#" onclick="enviar();return false">Enviar</a>'
    + '<button onclick="document.getElementById(\\'adjuntos\\').style.display=\\'block\\'"><i class="fa fa-paperclip"></i></button>';
}
function buscar(entrada, nombre, todas){
  var ul = document.querySelector('ul.ui-autocomplete'), texto = entrada.value.toLowerCase();
  var hay = OPCIONES[nombre].filter(o => todas || (o[0].toLowerCase().startsWith(texto) && texto.length > 2));
  ul.innerHTML = hay.length ? '' : '<li>No results found</li>';
  hay.forEach(o => { var li = document.createElement('li'); li.className = 'ui-menu-item'; li.textContent = o[1];
    li.onclick = () => { entrada.value = o[1]; ul.style.display = 'none';
      if (nombre === 'ticketTheme') OPCIONES.ticketType = (TIPOS[o[1]] || []).map(t => ['', t]);
      if (nombre === 'ticketType') criterio(o[1]); }; ul.appendChild(li); });
  ul.style.display = 'block';
}
function criterio(tipo){ document.getElementById('criterios').innerHTML = (CRITERIOS[tipo] || []).map(c =>
  '<div class="row-fluid has-labels"><span x-field="ticketReferenceType"><label>Tipo de criterio</label> <a>' + c + '</a></span> Valor <div x-field="value"><input></div></div>').join(''); }
function valores(){ var r = {criterios: {}};
  document.querySelectorAll('[x-field] input').forEach(i => { if (i.parentNode.getAttribute('x-field') !== 'value') r[i.parentNode.getAttribute('x-field')] = i.value; });
  document.querySelectorAll('.row-fluid.has-labels').forEach(f => r.criterios[f.querySelector('[x-field="ticketReferenceType"] a').textContent] = f.querySelector('input').value);
  r.description = document.querySelector('[x-field="description"] [contenteditable]').innerText; return r; }
function enviar(){ fetch('/enviado', {method:'POST', body: JSON.stringify(valores())}).then(r => r.text()).then(n => { document.querySelector('[x-field="ticketSeq"]').textContent = 'Nº Ticket' + n; }); }
document.querySelector('input.upload-input').onchange = function(){
  var f = this.files[0]; document.getElementById('lista').innerHTML += '<span>' + f.name + '</span>';
  fetch('/adjunto', {method:'POST', body: f.name}); };
// «la persona» pulsa Enviar cuando el portal se lo indica y el ticket ya está rellenado
setInterval(() => fetch('/pulsar').then(r => r.text()).then(t => {
  var v = document.querySelector('.row-fluid.has-labels input');
  if (t === 'si' && v && v.value && !window.pulsado) { window.pulsado = true; enviar(); } }), 300);
window.onhashchange = vista; vista();
</script></body></html>"""


class Portal:
    def __init__(self):
        self.enviados, self.adjuntos, self.pulsar = [], [], False
        self.intentos_login, self.fallos_login = 0, 0
        self.sesion, self.errores_internos = False, 0
        portal = self

        class Manejador(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def responder(self, cuerpo, tipo="text/html; charset=utf-8"):
                datos = cuerpo.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", tipo)
                self.send_header("Content-Length", str(len(datos)))
                self.end_headers()
                self.wfile.write(datos)

            def do_GET(self):
                if self.path.startswith("/app.html"):
                    if portal.errores_internos > 0:  # como el real: a veces, tras el acceso, «internal error»
                        portal.errores_internos -= 1
                        return self.responder("<html><body><h1>Internal Error</h1></body></html>")
                    return self.responder(APP)
                if self.path == "/" and portal.sesion:  # con la sesión abierta el enlace lleva directo a la aplicación
                    return self.responder("<html><body><script>location.href = \"/app.html\"</script></body></html>")
                if self.path == "/pulsar":
                    return self.responder("si" if portal.pulsar else "no", "text/plain")
                return self.responder(LOGIN)

            def do_POST(self):
                cuerpo = self.rfile.read(int(self.headers.get("Content-Length") or 0)).decode("utf-8")
                if self.path == "/login":
                    datos = json.loads(cuerpo)
                    portal.intentos_login += 1
                    if portal.fallos_login > 0:
                        portal.fallos_login -= 1
                        return self.responder("no", "text/plain")
                    correcto = (datos["u"], datos["c"]) == ("usuario", "clave")
                    portal.sesion = portal.sesion or correcto
                    return self.responder("ok" if correcto else "no", "text/plain")
                if self.path == "/enviado":
                    portal.enviados.append(json.loads(cuerpo))
                    return self.responder(f"TKT99900{len(portal.enviados)}", "text/plain")
                if self.path == "/adjunto":
                    portal.adjuntos.append(cuerpo)
                return self.responder("ok", "text/plain")

        self.servidor = ThreadingHTTPServer(("127.0.0.1", 0), Manejador)
        self.servidor.daemon_threads = True
        threading.Thread(target=self.servidor.serve_forever, daemon=True).start()
        self.url = f"http://127.0.0.1:{self.servidor.server_port}/"


@pytest.fixture
def portal(monkeypatch):
    portal = Portal()
    # La lista de tickets del portal falso está en /app.html#/ds/...
    monkeypatch.setattr(plt, "RUTA_LISTA_TICKETS", "app.html#/ds/helpdesk.ticket/list/1")
    yield portal
    portal.servidor.shutdown()


@pytest.fixture
def word(tmp_path):
    ruta = tmp_path / "Litigio 131172031.docx"
    ruta.write_bytes(b"word")
    return ruta


def datos(word, **cambios):
    base = {"litigio": "131172031", "ruta_word": str(word), "cuenta": "123456X - EMPRESA PRUEBA, SL", "cliente": "123456X03 - EMPRESA PRUEBA, SL",
            "tema": plt.TEMA_POR_DEFECTO, "tipo": plt.TIPO_POR_DEFECTO}
    return dict(base, **cambios)


def ticket(portal, word, **kw):
    registro = []
    extra = {k: kw.pop(k) for k in list(kw) if k in ("litigio", "cuenta", "cliente", "tema", "tipo")}
    return plt.TicketPowerLine("usuario", "clave", datos(word, **extra), log=lambda t, tag=None: registro.append(t), url=portal.url, visible=False, **kw), registro


ESPERADO = {"principal": "123456X - EMPRESA PRUEBA, SL", "shipToParty": "123456X03 - EMPRESA PRUEBA, SL", "ticketTheme": "Reclamaciones / Devoluciones",
            "ticketType": "Progreso / Seguimiento de una reclamacion", "criterios": {"Número de reclamacion": "131172031"}, "description": ""}


def test_envio_automatico_rellena_envia_y_adjunta(portal, word):
    t, registro = ticket(portal, word, enviar_solo=True)
    assert t.ejecutar() == "TKT999001" and t.error_adjunto is None
    assert portal.enviados == [ESPERADO]
    assert portal.adjuntos == ["Litigio 131172031.docx"]


def test_con_confirmacion_espera_a_que_la_persona_pulse_enviar(portal, word):
    portal.pulsar = True  # «la persona» pulsa Enviar en cuanto el ticket está rellenado
    t, registro = ticket(portal, word, enviar_solo=False, espera_confirmacion=30)
    assert t.ejecutar() == "TKT999001"
    assert portal.enviados == [ESPERADO] and portal.adjuntos == ["Litigio 131172031.docx"]
    assert any("pulsa «Enviar»" in r for r in registro)


def test_si_nadie_pulsa_enviar_no_se_envia_nada(portal, word):
    t, _ = ticket(portal, word, enviar_solo=False, espera_confirmacion=3)
    with pytest.raises(plt.TicketCancelado):
        t.ejecutar()
    assert portal.enviados == [] and portal.adjuntos == []


def test_un_valor_que_no_existe_se_explica(portal, word):
    t, _ = ticket(portal, word, enviar_solo=True, cliente="123456X09 - OTRA, SL")
    with pytest.raises(ValueError, match="no se encontró «123456X09 - OTRA, SL»"):
        t.ejecutar()
    assert portal.enviados == []


def test_contrasena_rechazada_se_reintenta_y_luego_avisa(portal, word, monkeypatch):
    monkeypatch.setattr(plt, "PAUSA_REINTENTO_ACCESO", 0.1)
    registro = []
    t = plt.TicketPowerLine("usuario", "mal", datos(word), url=portal.url, visible=False, enviar_solo=True, log=lambda texto, tag=None: registro.append(texto))
    with pytest.raises(plt.CredencialesRechazadas, match="Credenciales incorrectas.*3 intentos"):
        t.ejecutar()
    assert sum("Se reintenta" in r for r in registro) == 2 and portal.intentos_login == 3 and portal.enviados == []


def test_si_el_acceso_falla_una_vez_el_reintento_entra(portal, word, monkeypatch):
    monkeypatch.setattr(plt, "PAUSA_REINTENTO_ACCESO", 0.1)
    portal.fallos_login = 1  # la web rechaza la primera vez aunque la contraseña es buena
    t, registro = ticket(portal, word, enviar_solo=True)
    assert t.ejecutar() == "TKT999001" and portal.intentos_login == 2


def test_tipo_de_ticket_se_elige_del_desplegable_sin_escribir(portal, word):
    t, registro = ticket(portal, word, enviar_solo=True, tipo="Otro tipo que ya no existe")
    assert t.ejecutar() == "TKT999001"
    assert portal.enviados[0]["ticketType"] == "Progreso / Seguimiento de una reclamacion"  # la segunda opción
    assert any("se elige la opción 2" in r for r in registro)


def test_internal_error_tras_el_acceso_se_vuelve_a_abrir_el_enlace(portal, word):
    portal.errores_internos = 2
    t, registro = ticket(portal, word, enviar_solo=True)
    assert t.ejecutar() == "TKT999001" and portal.intentos_login == 1  # sin volver a escribir la contraseña
    assert sum("internal error" in r for r in registro) == 2


@pytest.mark.parametrize("clave", ["plazo_faltante", "motivo_cancelacion", "cancelacion_linea"])
def test_plantillas_rellenan_sus_criterios_por_nombre_y_no_adjuntan_nada(portal, tmp_path, clave):
    nombre, tema, tipo, criterios = plt.PLANTILLAS[clave]
    valores = {c: f"V{i}" for i, c in enumerate(criterios)}
    datos = {"cuenta": "123456X - EMPRESA PRUEBA, SL", "cliente": "123456X03 - EMPRESA PRUEBA, SL", "tema": tema, "tipo": tipo,
             "criterios": valores, "descripcion": "Texto de prueba"}
    t = plt.TicketPowerLine("usuario", "clave", datos, url=portal.url, visible=False, enviar_solo=True)
    assert t.ejecutar() == "TKT999001" and t.error_adjunto is None
    enviado = portal.enviados[0]
    assert (enviado["ticketTheme"], enviado["ticketType"]) == (tema, tipo)
    assert enviado["criterios"] == valores and enviado["description"] == "Texto de prueba"
    assert portal.adjuntos == []  # sin Word no se adjunta nada


def test_criterio_que_no_existe_se_explica(portal, word):
    t, _ = ticket(portal, word, enviar_solo=True)
    t.datos["criterios"] = {"Número de bastidor": "X"}
    with pytest.raises(ValueError, match="no tiene el criterio «Número de bastidor»"):
        t.ejecutar()
    assert portal.enviados == []


def test_busquedas_para():
    assert plt.busquedas_para("123456X - EMPRESA, SL") == ["123456X", "123456X - EMPRESA, SL"]
    assert plt.busquedas_para(plt.TEMA_POR_DEFECTO)[:2] == ["Litiges", plt.TEMA_POR_DEFECTO]
    assert plt.numero_ticket("Nº TicketTKT1612254") == "TKT1612254" and plt.numero_ticket("Nº Ticket") is None


def test_el_historial_guarda_el_ticket(tmp_path):
    historial = pl.HistorialReclamaciones(str(tmp_path / "r.db"))
    registro_id = historial.guardar("131172031", "", "9812", "w.docx", None)
    historial.guardar_ticket(registro_id, "TKT1612254")
    fila = historial.buscar("TKT16")[0]
    assert fila[1] == "131172031" and fila[7] == "TKT1612254"


def test_base_de_datos_antigua_recibe_la_columna(tmp_path):
    import sqlite3
    ruta = str(tmp_path / "vieja.db")
    con = sqlite3.connect(ruta)
    con.execute("CREATE TABLE reclamaciones_power_line (id INTEGER PRIMARY KEY AUTOINCREMENT, litigio TEXT NOT NULL, fecha TEXT NOT NULL, "
                "pedido_sap TEXT, referencia TEXT, ruta_word TEXT NOT NULL, ruta_pdf TEXT)")
    con.execute("INSERT INTO reclamaciones_power_line (litigio, fecha, ruta_word) VALUES ('1', 'x', 'w')")
    con.commit()
    con.close()
    assert pl.HistorialReclamaciones(ruta).buscar("1")[0][7] is None


def test_configuracion_del_ticket(app):
    app.tabview.set("Configuración")
    app.construir_pestana("Configuración")
    modulo = app.reclamaciones_power_line
    assert modulo.tema_power_line_entry.get() == plt.TEMA_POR_DEFECTO and modulo.tipo_power_line_entry.get() == plt.TIPO_POR_DEFECTO
    modulo.cuenta_power_line_entry.insert(0, "123456X - EMPRESA PRUEBA, SL")
    modulo.envio_automatico_var.set(True)
    config = modulo.obtener_configuracion()
    assert config["power_line_cuenta"] == "123456X - EMPRESA PRUEBA, SL" and config["power_line_envio_automatico"] is True
    assert modulo.valores_ticket()["tema"] == plt.TEMA_POR_DEFECTO


# ---------- pestaña «Tickets Power-Line» ----------

@pytest.fixture
def pestana_tickets(app, monkeypatch, tmp_path):
    monkeypatch.setattr(pl, "TICKETS_POWER_LINE", str(tmp_path / "tickets.json"))
    app.tabview.set("Reclamaciones")
    app.construir_pestana("Reclamaciones")
    modulo = app.reclamaciones_power_line
    modulo.configuracion.update({"power_line_cuenta": "123456X - EMPRESA PRUEBA, SL", "power_line_cliente": "123456X03 - EMPRESA PRUEBA, SL",
                                 "usuario_madd": "u", "password_madd": "p"})
    return modulo


def test_los_campos_cambian_con_la_plantilla(pestana_tickets):
    m = pestana_tickets
    assert list(m.criterios_entries) == ["Número de pedido de SAP", "Referencia", "Cantidad"]  # plazo de un faltante
    m.criterios_entries["Referencia"].insert(0, "9812")
    m.plantilla_ticket_option.set("Cancelación de línea o pedido completo")
    m._pintar_criterios()
    assert list(m.criterios_entries) == ["Número de pedido de SAP", "Referencia", "Cantidad", "Número de linea de pedido de SAP"]
    assert m.criterios_entries["Referencia"].get() == "9812"  # lo común se conserva


def test_crear_ticket_con_plantilla(pestana_tickets, monkeypatch):
    m = pestana_tickets
    avisos = []
    monkeypatch.setattr(pl.messagebox, "showwarning", lambda *a, **k: avisos.append(a[1]))
    assert m.crear_ticket_plantilla() is False and "Rellena" in avisos[-1]
    lanzados = []
    monkeypatch.setattr(m, "_lanzar_ticket", lambda datos, que, registro_id=None, plantilla=None: lanzados.append((datos, que, plantilla)) or True)
    m.plantilla_ticket_option.set("Motivo de cancelación de pieza codificada")
    m._pintar_criterios()
    m.criterios_entries["Número de pedido de SAP"].insert(0, "2063765885")
    m.criterios_entries["Referencia"].insert(0, "9812345680")
    m.descripcion_ticket_text.insert("1.0", "Pieza ya no necesaria")
    assert m.crear_ticket_plantilla()
    datos, que, plantilla = lanzados[0]
    assert (datos["tema"], datos["tipo"]) == ("Pedidos", "Motivo de cancelacion de un pedido de pieza codificada")
    assert datos["criterios"] == {"Número de pedido de SAP": "2063765885", "Referencia": "9812345680"} and datos["descripcion"] == "Pieza ya no necesaria"
    assert "ruta_word" not in datos and plantilla == "Motivo de cancelación de pieza codificada"


def test_los_tickets_creados_se_listan(pestana_tickets):
    pl.anotar_ticket("TKT999001", "Plazo para un faltante sin aviso", {"Número de pedido de SAP": "1", "Referencia": "R", "Cantidad": "2"})
    pestana_tickets.refrescar_tickets_recientes()
    assert "TKT999001  Plazo para un faltante sin aviso: 1 · R · 2" in pestana_tickets.tickets_recientes_box.get("1.0", "end")

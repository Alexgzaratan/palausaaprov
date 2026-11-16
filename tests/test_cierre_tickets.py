"""Cerrar y valorar los tickets resueltos de Power-Line contra un portal falso que imita al real: lista
«Tickets resueltos» (SlickGrid), comentarios del ticket y encuesta de «Cerrar» (Calidad de servicio,
¿La respuesta resolvió su problema?, Confirmar)."""
import pytest

import power_line_ticket as plt
import reclamaciones_power_line as pl
from test_power_line_ticket import Portal

APP_RESUELTOS = """<html><body><div>Pagina principal · Servicio de asistencia técnica</div><div id="vista"></div>
<div class="ui-dialog" id="encuesta" style="display:none"><span>Encuesta de satisfacción</span>
  <span x-field="rating"><ul>
    <li><label><input type="radio" name="r" value="0.00"><span>Nada satisfactorio</span></label></li>
    <li><label><input type="radio" name="r" value="1.00"><span>Insatisfactorio</span></label></li>
    <li><label><input type="radio" name="r" value="2.00"><span>Satisfactorio</span></label></li>
    <li><label><input type="radio" name="r" value="3.00"><span>Totalmente satisfactorio</span></label></li></ul></span>
  <span x-field="isProblemSolved"><label><input type="radio" name="s" value="true"><span>Sí</span></label>
    <label><input type="radio" name="s" value="false"><span>No</span></label></span>
  <span x-field="satisfactionComment"><textarea></textarea></span>
  <button x-field="confirm" name="confirm" onclick="confirmar()">Confirmar</button>
  <button onclick="document.getElementById('encuesta').style.display='none'">X</button></div>
<script>
var TICKETS = {
  TKT1000001: ['Hola Alex', 'Litigio aceptado, proceder a la destrucción de la pieza', 'Saludos'],
  TKT1000002: ['Buenos días', 'El abono se hará en la próxima factura']
};
var abierto = null, cerrados = {};
function lista(){
  var filas = Object.keys(TICKETS).filter(n => !cerrados[n]).map(n =>
    '<div class="slick-row" ondblclick="ficha(\\'' + n + '\\')"><div class="slick-cell">' + n + '</div><div class="slick-cell">Resuelto</div></div>').join('');
  // como el real: la rejilla aparece y las filas llegan un poco después
  document.getElementById('vista').innerHTML = '<div class="slick-viewport" style="min-height:200px"></div>';
  setTimeout(() => { document.querySelector('.slick-viewport').innerHTML = filas; }, 400);
}
function ficha(n){
  abierto = n;
  history.pushState(null, '', '#/ds/helpdesk.ticket.resolved/edit/' + n);  // como el real: la ficha tiene su propia dirección
  var c = '<div class="o2m-list-row">Actualizado por ALEX encendido 05/11/2026 - 11:35 Status : New</div>'
        + '<div class="o2m-list-row">Actualizado por GESTOR encendido 05/11/2026 - 11:41\\nStatus : In progress » Resolved\\n'
        + TICKETS[n].join('\\n') + '\\nPD - Gracias por cerrar y valorar los tickets de su placa.</div>';
  document.getElementById('vista').innerHTML = '<div x-field="ticketSeq">Nº Ticket' + n + '</div>'
    + '<div x-field="ticketCommentList">' + c + '</div>'
    + '<a x-field="closeTicket" name="closeTicket" href="#" onclick="abrirEncuesta();return false">Cerrar</a>';
  document.querySelectorAll('.o2m-list-row').forEach(r => r.innerHTML = r.innerHTML.split('\\n').join('<br>'));
}
function abrirEncuesta(){ document.querySelectorAll('#encuesta input').forEach(i => i.checked = false); document.getElementById('encuesta').style.display = 'block'; }
function confirmar(){
  var r = document.querySelector('#encuesta [x-field="rating"] input:checked'), s = document.querySelector('#encuesta [x-field="isProblemSolved"] input:checked');
  if (!r || !s) { alert('Please rate the quality of service'); return; }
  fetch('/cerrado', {method:'POST', body: JSON.stringify({ticket: abierto, rating: r.value, resuelto: s.value})}).then(() => {
    cerrados[abierto] = true; document.getElementById('encuesta').style.display = 'none';
    document.querySelector('[x-field="closeTicket"]').style.display = 'none'; });
}
function vista(){ if (location.hash === '#/ds/helpdesk.ticket.resolved/list/1') lista(); }
// «la persona» pulsa Confirmar cuando el portal se lo indica y la encuesta está marcada
setInterval(() => fetch('/pulsar').then(r => r.text()).then(t => {
  var e = document.getElementById('encuesta');
  if (t === 'si' && e.style.display === 'block' && document.querySelector('#encuesta input:checked')) confirmar(); }), 300);
window.onhashchange = vista; vista();
</script></body></html>"""


@pytest.fixture
def portal(monkeypatch):
    portal = Portal(app=APP_RESUELTOS)
    monkeypatch.setattr(plt, "RUTA_TICKETS_RESUELTOS", "app.html#/ds/helpdesk.ticket.resolved/list/1")
    yield portal
    portal.servidor.shutdown()


def cierre(portal, **kw):
    registro = []
    return plt.CierreTickets("usuario", "clave", log=lambda t, tag=None: registro.append(t), url=portal.url, visible=False, **kw), registro


def test_confirmando_sola_cierra_todos_con_totalmente_satisfactorio_y_si(portal):
    tarea, _ = cierre(portal, confirmar_solo=True)
    resultados = tarea.ejecutar()
    assert [(r["ticket"], r["cerrado"], r["error"]) for r in resultados] == [("TKT1000001", True, None), ("TKT1000002", True, None)]
    assert portal.cerrados == [{"ticket": "TKT1000001", "rating": "3.00", "resuelto": "true"}, {"ticket": "TKT1000002", "rating": "3.00", "resuelto": "true"}]
    assert resultados[0]["respuesta"].splitlines()[:2] == ["Hola Alex", "Litigio aceptado, proceder a la destrucción de la pieza"]


def test_sin_confirmar_espera_a_que_la_persona_pulse_confirmar(portal):
    portal.pulsar = True
    tarea, registro = cierre(portal)
    resultados = tarea.ejecutar()
    assert all(r["cerrado"] for r in resultados) and len(portal.cerrados) == 2
    assert any("pulsa «Confirmar»" in linea for linea in registro)


def test_si_nadie_confirma_no_se_cierra_nada(portal):
    tarea, _ = cierre(portal, espera_confirmacion=2)
    resultados = tarea.ejecutar()
    assert [r["cerrado"] for r in resultados] == [False, False] and portal.cerrados == []


def test_limpiar_respuesta():
    texto = "Actualizado por GESTOR encendido 05/11/2026 - 11:41\n\nStatus : In progress » Resolved\n\nHola Alex\n\nLitigio aceptado\n"
    assert plt.limpiar_respuesta(texto) == "Hola Alex\nLitigio aceptado"


def test_resumen_y_registro_de_respuestas(tmp_path):
    resultados = [{"ticket": "TKT1", "litigio": "131209474", "respuesta": "Hola\nLitigio aceptado\nPD - Gracias", "cerrado": True, "error": None},
                  {"ticket": "TKT2", "litigio": None, "respuesta": "", "cerrado": False, "error": None}]
    texto = pl.texto_resueltos(resultados)
    assert texto.startswith("1 de 2 tickets cerrado y valorado") and "TKT1 (litigio 131209474): cerrado ✓" in texto
    assert "«Hola Litigio aceptado»" in texto and "TKT2: sin cerrar" in texto and "Gracias" not in texto
    ruta = str(tmp_path / "respuestas.json")
    pl.anotar_respuestas(resultados, ruta)
    assert [(r["ticket"], r["cerrado"]) for r in pl.leer_respuestas(ruta)] == [("TKT1", True), ("TKT2", False)]
    assert pl.texto_resueltos([]) == "No hay tickets resueltos pendientes de cerrar."


def test_el_historial_encuentra_el_litigio_del_ticket(tmp_path):
    historial = pl.HistorialReclamaciones(str(tmp_path / "h.db"))
    registro = historial.guardar("131209474", "SAP1", "REF", "w.docx", None)
    historial.guardar_ticket(registro, "TKT1622729")
    assert historial.litigio_de_ticket("TKT1622729") == "131209474" and historial.litigio_de_ticket("TKT0") is None


def test_boton_de_la_pestana_tickets(app, monkeypatch, tmp_path):
    import app_automatizacion_optimizada as app_mod
    monkeypatch.setattr(pl, "RESPUESTAS_POWER_LINE", str(tmp_path / "respuestas.json"))
    app.tabview.set("Reclamaciones")
    app.construir_pestana("Reclamaciones")
    m = app.reclamaciones_power_line
    m.configuracion.update({"usuario_madd": "u", "password_madd": "p", "power_line_envio_automatico": False})
    usados = []

    class Falso:
        def __init__(self, usuario, password, confirmar_solo=False, log=None, visible=True, solo_leer=False):
            usados.append((confirmar_solo, visible, solo_leer))
            self.solo_leer = solo_leer

        def ejecutar(self):
            return [{"ticket": "TKT1622729", "respuesta": "Litigio aceptado", "cerrado": not self.solo_leer, "error": None}]

    monkeypatch.setattr(plt, "CierreTickets", Falso)
    monkeypatch.setattr(pl.threading, "Thread", lambda target, daemon=None: type("H", (), {"start": lambda self: target()})())
    monkeypatch.setattr(app, "after", lambda ms, funcion, *a: funcion(*a))
    monkeypatch.setattr(m.historial, "litigio_de_ticket", lambda ticket: "131209474")
    mostrados = []
    monkeypatch.setattr(pl.messagebox, "showinfo", lambda titulo, texto, **k: mostrados.append(texto))
    monkeypatch.setattr(pl.efectos, "animar_envio", lambda *a, **k: None)
    assert m.cerrar_tickets_resueltos()
    assert usados == [(False, True, False)]  # sin el envío automático, «Confirmar» lo pulsa la persona
    assert "TKT1622729 (litigio 131209474): cerrado ✓" in mostrados[0] and "«Litigio aceptado»" in mostrados[0]
    assert pl.leer_respuestas()[-1]["litigio"] == "131209474"
    assert m.btn_cerrar_resueltos.cget("text") == "Cerrar tickets resueltos" and not m._enviando_power_line

    assert "TKT1622729 · litigio 131209474 · cerrado" in m.respuestas_box.get("1.0", "end")


def test_ver_respuestas_sin_cerrar_ni_abrir_ventana(app, monkeypatch, tmp_path):
    monkeypatch.setattr(pl, "RESPUESTAS_POWER_LINE", str(tmp_path / "respuestas.json"))
    app.tabview.set("Reclamaciones")
    app.construir_pestana("Reclamaciones")
    m = app.reclamaciones_power_line
    m.configuracion.update({"usuario_madd": "u", "password_madd": "p"})
    assert "Pulsa «Ver respuestas»" in m.respuestas_box.get("1.0", "end")
    usados = []

    class Falso:
        def __init__(self, usuario, password, confirmar_solo=False, log=None, visible=True, solo_leer=False):
            usados.append((visible, solo_leer))

        def ejecutar(self):
            return [{"ticket": "TKT1622729", "respuesta": "Hola Alex\nLitigio aceptado, proceder a la destrucción de la pieza\nPD - Gracias", "cerrado": False, "error": None}]

    monkeypatch.setattr(plt, "CierreTickets", Falso)
    monkeypatch.setattr(pl.threading, "Thread", lambda target, daemon=None: type("H", (), {"start": lambda self: target()})())
    monkeypatch.setattr(app, "after", lambda ms, funcion, *a: funcion(*a))
    monkeypatch.setattr(m.historial, "litigio_de_ticket", lambda ticket: "131209474")
    ventanas = []
    monkeypatch.setattr(pl.messagebox, "showinfo", lambda *a, **k: ventanas.append(a))
    assert m.ver_respuestas_resueltos()
    assert usados == [(False, True)] and ventanas == []
    caja = m.respuestas_box.get("1.0", "end")
    assert "TKT1622729 · litigio 131209474 · pendiente de cerrar" in caja and "proceder a la destrucción de la pieza" in caja and "Gracias" not in caja
    assert m.btn_ver_respuestas.cget("text") == "Ver respuestas"


def test_solo_leer_no_pulsa_cerrar(portal):
    tarea, _ = cierre(portal, solo_leer=True)
    resultados = tarea.ejecutar()
    assert [r["ticket"] for r in resultados] == ["TKT1000001", "TKT1000002"] and portal.cerrados == []
    assert "El abono se hará en la próxima factura" in resultados[1]["respuesta"]


def test_al_abrir_avisa_arriba_de_los_tickets_resueltos_sin_cerrar_nada(app, monkeypatch, tmp_path):
    monkeypatch.setattr(pl, "RESPUESTAS_POWER_LINE", str(tmp_path / "respuestas.json"))
    m = app.reclamaciones_power_line
    m.configuracion.update({"usuario_madd": "u", "password_madd": "p"})
    usados = []

    class Falso:
        def __init__(self, usuario, password, confirmar_solo=False, log=None, visible=True, solo_leer=False):
            usados.append((visible, solo_leer))

        def ejecutar(self):
            return [{"ticket": "TKT1", "respuesta": "Aceptado", "cerrado": False, "error": None},
                    {"ticket": "TKT2", "respuesta": "Aceptado", "cerrado": False, "error": None}]

    monkeypatch.setattr(plt, "CierreTickets", Falso)
    monkeypatch.setattr(pl.threading, "Thread", lambda target, daemon=None: type("H", (), {"start": lambda self: target()})())
    monkeypatch.setattr(app, "after", lambda ms, funcion, *a: funcion(*a))
    monkeypatch.setattr(m.historial, "litigio_de_ticket", lambda ticket: None)
    avisos = []
    monkeypatch.setattr(app, "avisar", lambda texto, *a, **k: avisos.append(texto))
    assert m.comprobar_resueltos_al_abrir()
    assert usados == [(False, True)]  # sin ventana y solo leyendo
    assert m.resueltos_pendientes == 2 and not m._enviando_power_line
    assert avisos == ["Hay 2 tickets resueltos en Power-Line sin cerrar."]
    assert app.ficha_resueltos.winfo_exists()
    app.ir_a_tickets_resueltos()
    assert app.tabview.get() == "Reclamaciones" and m.panel.get() == "Tickets Power-Line"

    # Desactivado en Configuración, o si Power-Line falla: no hace ni dice nada
    m.configuracion["power_line_avisar_resueltos"] = False
    assert not m.comprobar_resueltos_al_abrir() and len(usados) == 1
    m.configuracion["power_line_avisar_resueltos"] = True
    monkeypatch.setattr(Falso, "ejecutar", lambda self: (_ for _ in ()).throw(RuntimeError("sin conexión")))
    assert m.comprobar_resueltos_al_abrir() and len(avisos) == 1 and not m._enviando_power_line

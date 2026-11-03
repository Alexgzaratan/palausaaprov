"""Ticket de seguimiento de una reclamación en Power-Line (Axelor) con Playwright.

Pasos (comprobados en el portal el 29/10/2026):
  1. Inicio de sesión de Stellantis en idfed.mpsa.com: usuario -> «Next» -> contraseña -> «Sign On».
  2. «Ticket para responsables» > «Tickets» > «+»: formulario de ticket nuevo.
  3. Cuenta Solicitante, Cliente destinatario, Tema y Tipo (buscadores; el tema y el tipo se buscan con su nombre
     interno en francés, p. ej. «Litiges» o «Avancement», y se muestran en español). Al elegir el tipo aparece el
     criterio «Número de reclamacion»: su Valor es el Nº de litigio. La descripción se deja vacía.
  4. «Enviar»: lo pulsa la persona (por defecto) o la app si el envío automático está activado.
  5. Ya enviado (aparece el Nº TKT...), se adjunta el Word del litigio con el clip.
Los valores de la cuenta y del cliente viven en la configuración de cada usuario, no en el código.

Plantillas (comprobadas el 03/11/2026): cada tipo de ticket añade sus filas de criterio («Tipo de criterio» +
«Valor»); se rellenan por el nombre del criterio. El tema y el tipo se eligen abriendo el desplegable."""
import re
import time

URL_POWERLINE = "https://powerline.axelor.com/"
RUTA_LISTA_TICKETS = "#/ds/helpdesk.ticket/list/1"
TEMA_POR_DEFECTO = "Reclamaciones / Devoluciones"
TIPO_POR_DEFECTO = "Progreso / Seguimiento de una reclamacion"
# Nombre interno (francés) con el que el buscador encuentra cada opción
BUSQUEDAS_CONOCIDAS = {
    TEMA_POR_DEFECTO.lower(): ("Litiges",),
    TIPO_POR_DEFECTO.lower(): ("Avancement",),
}
CRITERIO_RECLAMACION = "Número de reclamacion"
# clave: (nombre en la app, tema, tipo, criterios). Los textos son los de Power-Line, erratas incluidas.
PLANTILLAS = {
    "seguimiento": ("Seguimiento de una reclamación", TEMA_POR_DEFECTO, TIPO_POR_DEFECTO, (CRITERIO_RECLAMACION,)),
    "plazo_faltante": ("Plazo para un faltante sin aviso", "Fecha estimada de entrega",
                       "Solicitud de plazo para un faltante sin aviso de reclamamción", ("Número de pedido de SAP", "Referencia", "Cantidad")),
    "motivo_cancelacion": ("Motivo de cancelación de pieza codificada", "Pedidos",
                           "Motivo de cancelacion de un pedido de pieza codificada", ("Número de pedido de SAP", "Referencia")),
    "cancelacion_linea": ("Cancelación de línea o pedido completo", "Pedidos", "Cancelacion linea de pedido o pedido completo",
                          ("Número de pedido de SAP", "Referencia", "Cantidad", "Número de linea de pedido de SAP")),
}
ESPERA_CONFIRMACION = 20 * 60  # segundos que se espera a que la persona pulse «Enviar»
ESPERA_INICIO_SESION = 90      # por si hay que completar a mano algún paso del acceso
INTENTOS_ACCESO = 3            # el acceso de Stellantis falla a veces sin motivo: 1 intento + 2 reintentos
PAUSA_REINTENTO_ACCESO = 5
RECARGAS_ERROR_INTERNO = 3     # veces que se vuelve a abrir el enlace si sale «internal error» tras el acceso


class TicketCancelado(Exception):
    """La persona cerró la ventana o no pulsó «Enviar» a tiempo: no se ha enviado nada."""


class CredencialesRechazadas(RuntimeError):
    def __str__(self):
        detalle = super().__str__()
        return "Stellantis ha rechazado el usuario o la contraseña de Power-Line" + (f" ({detalle})" if detalle else "")


def busquedas_para(valor):
    """Textos que se prueban en un buscador para encontrar 'valor': el código («123456X» de «123456X - EMPRESA, SL»),
    el nombre en francés si se conoce, el propio valor y su primera palabra."""
    valor = str(valor or "").strip()
    candidatos = []
    if " - " in valor:
        candidatos.append(valor.split(" - ", 1)[0].strip())
    candidatos += list(BUSQUEDAS_CONOCIDAS.get(valor.lower(), ()))
    candidatos += [valor, valor.split(" ")[0]]
    vistos = []
    for c in candidatos:
        if c and c not in vistos:
            vistos.append(c)
    return vistos


def numero_ticket(texto):
    encontrado = re.search(r"TKT\d+", texto or "")
    return encontrado.group(0) if encontrado else None


class TicketPowerLine:
    """datos: litigio, ruta_word, cuenta, cliente, tema, tipo. enviar_solo=True pulsa «Enviar» sin esperar a nadie."""

    def __init__(self, usuario, password, datos, enviar_solo=False, log=lambda texto, tag=None: None, url=URL_POWERLINE,
                 espera_confirmacion=ESPERA_CONFIRMACION, visible=True):
        self.usuario, self.password = usuario, password
        self.datos = datos
        self.enviar_solo = enviar_solo
        self.log = log
        self.url = url
        self.espera_confirmacion = espera_confirmacion
        self.visible = visible

    # ---------- pasos ----------

    def iniciar_sesion(self, page):
        """El acceso de Stellantis a veces rechaza o se cuelga sin motivo: se reintenta desde el principio hasta
        INTENTOS_ACCESO veces, con una pausa entre intentos."""
        ultimo = None
        for intento in range(1, INTENTOS_ACCESO + 1):
            try:
                self._intentar_acceso(page)
                return
            except TicketCancelado:
                raise
            except Exception as e:
                if page.is_closed():
                    raise TicketCancelado("se cerró la ventana")
                ultimo = e
                if intento < INTENTOS_ACCESO:
                    self.log(f"No se pudo entrar en Power-Line (intento {intento} de {INTENTOS_ACCESO}): {str(e).splitlines()[0]}. Se reintenta...", "error")
                    time.sleep(PAUSA_REINTENTO_ACCESO)
        if isinstance(ultimo, CredencialesRechazadas):
            raise CredencialesRechazadas(f"{ultimo.args[0] if ultimo.args else ''} en los {INTENTOS_ACCESO} intentos; revísalos en Configuración".strip())
        raise ultimo

    def _en_power_line(self, page):
        try:
            return page.locator("text=/Servicio de asistencia|Pagina principal|Bandeja de entrada/").count() > 0
        except Exception:
            return False

    def _intentar_acceso(self, page):
        page.goto(self.url, timeout=60000)
        identificador = page.locator("#identifierInput")
        try:
            identificador.wait_for(state="visible", timeout=20000)
            identificador.fill(self.usuario)
            page.click("#signOnButton")
        except Exception:
            pass  # sesión ya abierta o pantalla distinta: se sigue más abajo
        clave = page.locator("#password")
        try:
            clave.wait_for(state="visible", timeout=30000)
        except Exception:
            clave = None
        if clave is not None:
            usuario = page.locator("#username")
            if usuario.count() and usuario.first.is_visible() and not usuario.first.input_value():
                usuario.first.fill(self.usuario)
            clave.fill(self.password)
            page.click("#signOnButton")
            self._comprobar_acceso_rechazado(page)
        self.log("Esperando a entrar en Power-Line...")
        fin = time.monotonic() + ESPERA_INICIO_SESION
        recargas = 0
        while time.monotonic() < fin:
            if page.is_closed():
                raise TicketCancelado("se cerró la ventana")
            if self._en_power_line(page):
                return
            if self._error_interno(page) and recargas < RECARGAS_ERROR_INTERNO:
                # Tras el acceso Power-Line a veces muestra «internal error»: basta con volver a abrir el enlace
                recargas += 1
                self.log(f"Power-Line ha dado «internal error»: se vuelve a abrir {self.url} ({recargas} de {RECARGAS_ERROR_INTERNO}).", "error")
                page.goto(self.url, timeout=60000)
            time.sleep(1)
        raise RuntimeError("no se pudo entrar en Power-Line")

    @staticmethod
    def _error_interno(page):
        try:
            return page.locator("text=/internal (server )?error/i").count() > 0
        except Exception:
            return False

    def _comprobar_acceso_rechazado(self, page, segundos=15):
        """Si tras «Submit» vuelve la pantalla de contraseña (vacía), el acceso se ha rechazado."""
        fin = time.monotonic() + segundos
        time.sleep(1.5)
        while time.monotonic() < fin:
            if self._en_power_line(page):
                return
            clave = page.locator("#password")
            try:
                if clave.count() and clave.first.is_visible() and not clave.first.input_value():
                    mensaje = page.locator(".ping-error, .error, [class*=error]").first
                    texto = mensaje.inner_text(timeout=1000).strip() if mensaje.count() else ""
                    raise CredencialesRechazadas(texto)
            except CredencialesRechazadas:
                raise
            except Exception:
                pass
            time.sleep(1)

    def abrir_ticket_nuevo(self, page):
        page.goto(self.url.rstrip("/") + "/" + RUTA_LISTA_TICKETS, timeout=60000)
        nuevo = page.locator("button:has(i.fa-plus)").first
        nuevo.wait_for(state="visible", timeout=30000)
        nuevo.click()
        page.locator('[x-field="principal"] input').first.wait_for(state="visible", timeout=30000)

    @staticmethod
    def _esperar_valor(entrada, valor, segundos=5):
        fin = time.monotonic() + segundos
        while time.monotonic() < fin:
            if entrada.input_value().strip().lower() == valor.strip().lower():
                return True
            time.sleep(0.2)
        return False

    def elegir(self, page, campo, valor):
        """Escribe en el buscador y elige la opción que coincide exactamente con 'valor'."""
        entrada = page.locator(f'[x-field="{campo}"] input').first
        exacto = re.compile(r"^\s*" + re.escape(valor) + r"\s*$", re.I)
        for texto in busquedas_para(valor):
            entrada.click()
            entrada.fill("")
            entrada.type(texto, delay=30)
            opcion = page.locator("ul.ui-autocomplete:visible li.ui-menu-item").filter(has_text=exacto).first
            try:
                opcion.wait_for(state="visible", timeout=6000)
            except Exception:
                continue
            opcion.click()
            if self._esperar_valor(entrada, valor):
                return
        raise ValueError(f"no se encontró «{valor}» en el campo {campo} de Power-Line")

    def elegir_del_desplegable(self, page, campo, valor, posicion=2):
        """Tipo de ticket: no se escribe nada, se abre el desplegable y se elige 'valor' (hoy es la opción nº 'posicion').
        Si ese texto no aparece se elige la opción de esa posición y se avisa en el registro."""
        entrada = page.locator(f'[x-field="{campo}"] input').first
        lista = page.locator("ul.ui-autocomplete:visible li.ui-menu-item:not(.tag-select-action)")
        for abrir in ("ArrowDown", "ArrowDown", "escribir"):
            try:
                entrada.click()
                if abrir == "ArrowDown":
                    entrada.press("ArrowDown")
                else:
                    entrada.fill("")
                    entrada.type(busquedas_para(valor)[0], delay=30)
                lista.first.wait_for(state="visible", timeout=8000)
            except Exception:
                continue
            exacta = lista.filter(has_text=re.compile(r"^\s*" + re.escape(valor) + r"\s*$", re.I)).first
            if exacta.count():
                exacta.click()
            elif posicion and lista.count() >= posicion:
                elegida = lista.nth(posicion - 1)
                texto = elegida.inner_text().strip()
                self.log(f"En «Tipo de ticket» no aparece «{valor}»: se elige la opción {posicion}, «{texto}».", "error")
                elegida.click()
                valor = texto
            else:
                continue
            if self._esperar_valor(entrada, valor):
                return
        raise ValueError(f"no se pudo elegir «{valor}» en el desplegable {campo} de Power-Line")

    def criterios(self):
        """{nombre del criterio: valor}. Con 'litigio' (seguimiento de reclamación) es el «Número de reclamacion»."""
        if self.datos.get("criterios"):
            return dict(self.datos["criterios"])
        return {CRITERIO_RECLAMACION: str(self.datos["litigio"])}

    def rellenar_criterio(self, page, nombre, valor):
        # El nombre va en un enlace dentro del bloque «Tipo de criterio»
        etiqueta = page.locator('[x-field="ticketReferenceType"] a').filter(has_text=re.compile(r"^\s*" + re.escape(nombre) + r"\s*$", re.I))
        fila = page.locator("div.row-fluid.has-labels").filter(has=etiqueta).first
        try:
            fila.wait_for(state="visible", timeout=20000)  # las filas aparecen al elegir el tipo
        except Exception:
            raise ValueError(f"el ticket no tiene el criterio «{nombre}»")
        entrada = fila.locator('[x-field="value"] input').first
        entrada.click()
        entrada.fill(str(valor))
        entrada.press("Tab")

    def rellenar(self, page):
        for campo, clave in (("principal", "cuenta"), ("shipToParty", "cliente")):
            self.elegir(page, campo, self.datos[clave])
        try:
            self.elegir_del_desplegable(page, "ticketTheme", self.datos["tema"], posicion=None)
        except ValueError:
            self.elegir(page, "ticketTheme", self.datos["tema"])  # por si el desplegable no se abre: buscando
        # El tipo del seguimiento es hoy la 2.ª opción: si su texto cambiara se elige esa posición (los demás, solo por texto)
        posicion = None if self.datos.get("criterios") else 2
        self.elegir_del_desplegable(page, "ticketType", self.datos["tipo"], posicion=posicion)
        for nombre, valor in self.criterios().items():
            self.rellenar_criterio(page, nombre, valor)
        descripcion = str(self.datos.get("descripcion") or "").strip()
        if descripcion:
            editor = page.locator('[x-field="description"] [contenteditable="true"], [x-field="description"].wysiwyg-editor, [x-field="description"] .wysiwyg-editor').first
            # De una vez (letra a letra se perdía texto: el portal refresca el formulario mientras se escribe) y comprobándolo
            for _ in range(3):
                editor.click()
                editor.fill(descripcion)
                time.sleep(1)
                if " ".join(editor.inner_text().split()) == " ".join(descripcion.split()):
                    break
            else:
                raise ValueError("no se pudo escribir la descripción completa en Power-Line")

    def leer_numero(self, page):
        try:
            return numero_ticket(page.locator('[x-field="ticketSeq"]').first.inner_text(timeout=2000))
        except Exception:
            return None

    def enviar(self, page):
        page.locator('[x-field="sendTicket"]').first.click()
        aceptar = page.locator(".ui-dialog:visible button:has-text('OK')").first
        try:
            aceptar.wait_for(state="visible", timeout=3000)
            aceptar.click()
        except Exception:
            pass
        return self.esperar_numero(page, 90)

    def esperar_numero(self, page, segundos):
        fin = time.monotonic() + segundos
        while time.monotonic() < fin:
            if page.is_closed():
                raise TicketCancelado("se cerró la ventana sin enviar el ticket")
            numero = self.leer_numero(page)
            if numero:
                return numero
            time.sleep(1)
        return None

    def adjuntar_word(self, page, ruta_word):
        page.locator("button:has(i.fa-paperclip)").first.click()
        dialogo = page.locator(".ui-dialog:visible").filter(has_text="Attachments").first
        dialogo.wait_for(state="visible", timeout=20000)
        dialogo.locator("input.upload-input").first.set_input_files(str(ruta_word))
        nombre = re.sub(r".*[\\/]", "", str(ruta_word))
        dialogo.locator(f"text={nombre}").first.wait_for(state="visible", timeout=90000)
        dialogo.locator("button:has-text('Cerrar')").first.click()

    # ---------- todo junto ----------

    def ejecutar(self):
        """Devuelve el Nº de ticket. Lanza TicketCancelado si la persona no lo envía y RuntimeError/ValueError si
        algo falla antes de enviar. Si falla al adjuntar, el ticket ya está enviado: se devuelve con el error en
        self.error_adjunto para avisar de que hay que subir el Word a mano."""
        from playwright.sync_api import sync_playwright
        self.error_adjunto = None
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=not self.visible)
            try:
                page = navegador.new_context(no_viewport=True).new_page()
                self.log("Entrando en Power-Line...")
                self.iniciar_sesion(page)
                self.log("Creando el ticket...")
                self.abrir_ticket_nuevo(page)
                self.rellenar(page)
                if self.enviar_solo:
                    self.log("Enviando el ticket...")
                    numero = self.enviar(page)
                else:
                    self.log("Ticket rellenado: revísalo en la ventana de Power-Line y pulsa «Enviar».", "success")
                    numero = self.esperar_numero(page, self.espera_confirmacion)
                if not numero:
                    raise TicketCancelado("no se pulsó «Enviar»")
                if not self.datos.get("ruta_word"):
                    self.log(f"Ticket {numero} enviado.", "success")
                    return numero
                self.log(f"Ticket {numero} enviado. Adjuntando el Word...")
                try:
                    self.adjuntar_word(page, self.datos["ruta_word"])
                except Exception as e:
                    self.error_adjunto = str(e).splitlines()[0]
                return numero
            finally:
                try:
                    navegador.close()
                except Exception:
                    pass

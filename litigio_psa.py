"""Litigio de calidad en PSA Parts (parts.citroen.com, SAP ITS) con Playwright, como la skill reclamacion-calidad-psa.

Pantallas (comprobadas el 29/10/2026, hasta «Control»):
  - Menú en español: /sap/bc/gui/sap/its/zmenu?sap-language=ES (tras abrir sesión en /zzlogpl0).
  - FACTURA > «Lista de líneas de facturas»: material p_matnr[1], fechas date_f[1]/date_t[1], «Ejecutar programa»
    (#keyboard_ok). Columnas: FechaFact., ClVt, N° pedido cliente, Doc.ventas, Pos., Material, Factura, Pos., Ctd.fact...
  - LITIGIO > «Crear litigio con referencia a fact.»: factura v_vbeln[1], tipo (radio v_qualite[1] = Calidad), «Siguiente».
  - Líneas: motivo vbak-augru[1], buscar material v_matnr[1] + «Buscar» (la línea salta arriba), cantidad
    t_vbrp-zmeng_w[n] de esa fila, «Control» y «Guardar» -> «Litig calid S/retorn NNNNNNNNN se ha grabado».
«Guardar» crea el litigio y no se puede deshacer: antes se pide confirmación (salvo que se desactive)."""
import re
import time
from datetime import date, timedelta

URL_BASE = "https://parts.citroen.com"
RUTA_SESION = "/zzlogpl0"
RUTA_MENU = "/sap/bc/gui/sap/its/zmenu?sap-language=ES"
MAX_DIAS_FACTURA = 10
MOTIVOS = (  # motivos de un litigio de Calidad, tal como los da PARTS
    ("Q01", "PIEZA DETERIORADA"),
    ("Q02", "DETERIORADA EMBALAJE AUSENTE/INSUF"),
    ("Q03", "KIT INCOMPLETO"),
    ("Q04", "MAL EMBASADA ETIQ NO CORRESPONDE C/RECIB"),
    ("Q05", "DEFECTO TECNICO"),
    ("N05", "ERROR DE CATALOGO PIEZAS CODIFICADAS"),
)
MOTIVO_POR_DEFECTO = "Q01"


class LitigioNoCreado(Exception):
    """No se ha guardado ningún litigio (sin factura válida, error en Control, o la persona no confirmó)."""


def normalizar_ref(texto):
    return re.sub(r"\s+", "", str(texto or "")).replace("2/", "").upper()


def texto_motivo(codigo):
    return next((f"{c} · {t}" for c, t in MOTIVOS if c == codigo), codigo)


def codigo_motivo(texto):
    """'Q01 · PIEZA DETERIORADA' -> 'Q01'."""
    return str(texto or "").split("·")[0].strip() or MOTIVO_POR_DEFECTO


def desde_para_busqueda(hoy=None):
    """Día 1 del mes anterior (como en la skill)."""
    hoy = hoy or date.today()
    primero = hoy.replace(day=1)
    return (primero - timedelta(days=1)).replace(day=1)


def leer_fecha(texto):
    d, m, a = (int(x) for x in texto.strip().split("."))
    return date(a, m, d)


def elegir_linea(lineas, referencia, hoy=None, max_dias=MAX_DIAS_FACTURA):
    """La línea facturada más reciente de esa referencia, y como mucho de hace 'max_dias'; None si no hay."""
    hoy = hoy or date.today()
    suyas = [l for l in lineas if normalizar_ref(l["material"]) == normalizar_ref(referencia)]
    validas = [l for l in suyas if 0 <= (hoy - l["fecha"]).days <= max_dias]
    return max(validas, key=lambda l: (l["fecha"], l["factura"], l["pos_factura"])) if validas else None


def resumen_lineas(lineas, limite=6):
    return "; ".join(f"{l['fecha']:%d/%m} factura {l['factura']} (Doc.ventas {l['doc_ventas']})" for l in sorted(lineas, key=lambda l: l["fecha"], reverse=True)[:limite])


class LitigioPSA:
    """confirmar(datos) -> bool se llama antes de «Guardar» (None = guardar sin preguntar)."""

    def __init__(self, usuario, password, referencia, motivo=MOTIVO_POR_DEFECTO, cantidad=1, confirmar=None,
                 log=lambda texto, tag=None: None, url=URL_BASE, visible=False, hoy=None):
        self.usuario, self.password = usuario, password
        self.referencia = normalizar_ref(referencia)
        self.motivo = motivo
        self.cantidad = cantidad
        self.confirmar = confirmar
        self.log = log
        self.url = url.rstrip("/")
        self.visible = visible
        self.hoy = hoy or date.today()

    # ---------- utilidades del portal (SAP ITS: lento y a veces hay que repetir el clic) ----------

    @staticmethod
    def pulsar_hasta(page, objetivo, comprobar, intentos=3, espera=10000):
        """Pulsa 'objetivo' (selector o locator) hasta que 'comprobar()' sea cierto."""
        for _ in range(intentos):
            (page.locator(objetivo).first if isinstance(objetivo, str) else objetivo).click()
            fin = time.monotonic() + espera / 1000
            while time.monotonic() < fin:
                try:
                    if comprobar():
                        return True
                except Exception:
                    pass
                time.sleep(0.5)
        return False

    def abrir_menu(self, page):
        page.goto(self.url + RUTA_SESION, timeout=60000)
        page.wait_for_load_state("domcontentloaded")
        time.sleep(2)
        page.goto(self.url + RUTA_MENU, timeout=60000)
        self.enlace(page, "FACTURA").wait_for(state="visible", timeout=30000)

    @staticmethod
    def enlace(page, texto):
        """Enlace del menú por su texto exacto, sin distinguir mayúsculas (el portal escribe «Factura» y lo muestra en mayúsculas
        con estilos); «Lista de facturas» no cuenta."""
        return page.locator("a").filter(has_text=re.compile(r"^\s*" + re.escape(texto) + r"\s*$", re.I)).first

    def ir_a(self, page, menu, opcion, campo):
        """Menú (FACTURA, LITIGIO) > opción, hasta que aparezca 'campo'."""
        submenu = page.locator(f"a:has-text('{opcion}')").first
        ok = self.pulsar_hasta(page, self.enlace(page, menu), lambda: submenu.is_visible(), espera=6000) and \
            self.pulsar_hasta(page, submenu, lambda: page.locator(campo).first.is_visible())
        if not ok:
            raise RuntimeError(f"PSA Parts no abre {menu} > {opcion}")

    # ---------- paso 2: buscar la factura ----------

    def lineas_facturadas(self, page):
        self.ir_a(page, "FACTURA", "Lista de líneas de facturas", "input[name='p_matnr[1]']")
        page.fill("input[name='p_matnr[1]']", self.referencia)
        page.fill("input[name='date_f[1]']", f"{desde_para_busqueda(self.hoy):%d.%m.%Y}")
        hay_tabla = lambda: page.locator("text=Doc.ventas").count() > 0 or page.locator("text=/No se ha encontrado|ningún|ninguna/i").count() > 0
        self.pulsar_hasta(page, "#keyboard_ok", hay_tabla, intentos=3, espera=12000)
        filas = page.evaluate("""() => {
            const tabla = [...document.querySelectorAll('table')].find(t => [...t.rows].some(r => r.innerText.includes('Doc.ventas')));
            if (!tabla) return [];
            const filas = [...tabla.rows].map(r => [...r.cells].map(c => c.innerText.trim()));
            const i = filas.findIndex(f => f.includes('Doc.ventas'));
            return filas.slice(i).filter(f => f.length > 1);
        }""")
        if not filas:
            return []
        cabecera, lineas = filas[0], []
        i_doc, i_fact, i_mat = cabecera.index("Doc.ventas"), cabecera.index("Factura"), cabecera.index("Material")
        i_fecha = next(i for i, c in enumerate(cabecera) if c.lower().startswith("fechafact"))
        for f in filas[1:]:
            try:
                lineas.append({"fecha": leer_fecha(f[i_fecha]), "doc_ventas": f[i_doc], "pos_pedido": f[i_doc + 1], "material": f[i_mat],
                               "factura": f[i_fact], "pos_factura": f[i_fact + 1]})
            except (ValueError, IndexError):
                continue
        return lineas

    # ---------- paso 3: crear el litigio ----------

    def crear(self, page, linea):
        self.ir_a(page, "LITIGIO", "Crear litigio con referencia a fact", "input[name='v_vbeln[1]']")
        page.fill("input[name='v_vbeln[1]']", linea["factura"])
        page.check("input[type=radio][value='v_qualite[1]']")
        motivo = "select[name='vbak-augru[1]']"
        if not self.pulsar_hasta(page, "a:has-text('Siguiente')", lambda: page.locator(motivo).first.is_visible()):
            raise RuntimeError("PSA Parts no pasa a las líneas de la factura")
        opciones = page.locator(motivo + " option").evaluate_all("os => os.map(o => o.value)")
        if self.motivo not in opciones:  # al marcar Calidad salen los motivos Q..; si no, no se marcó bien
            raise RuntimeError(f"el motivo {self.motivo} no está disponible: el litigio no es de Calidad")
        page.select_option(motivo, self.motivo)
        page.fill("input[name='v_matnr[1]']", self.referencia)
        posicion = page.locator("td:first-child").filter(has_text=re.compile(r"^\s*" + re.escape(linea["pos_factura"]) + r"\s*$"))
        fila = page.locator("tr").filter(has=page.locator("input[name^='t_vbrp-zmeng_w']")).filter(has=posicion).first
        self.pulsar_hasta(page, "a:has-text('Buscar')", lambda: fila.is_visible(), espera=8000)
        if not fila.is_visible() or normalizar_ref(self.referencia) not in normalizar_ref(fila.inner_text()):
            raise RuntimeError(f"no aparece la línea {linea['pos_factura']} de {self.referencia} en la factura {linea['factura']}")
        fila.locator("input[name^='t_vbrp-zmeng_w']").first.fill(str(self.cantidad))
        url_antes = page.url
        self.pulsar_hasta(page, "a:has-text('Control')", lambda: page.url != url_antes, espera=10000)
        errores = self.errores(page, fila)
        if errores:
            raise RuntimeError("Control con errores: " + " · ".join(errores))

    @staticmethod
    def errores(page, fila):
        """Mensajes de error tras «Control»: el Status de la línea y los textos de error de la página."""
        mensajes = []
        try:
            estado = fila.locator("td").last.inner_text().strip()
            if estado:
                mensajes.append(f"status de la línea: {estado}")
        except Exception:
            pass
        # El portal pone sus mensajes en rojo (bloque «Message d'erreur»); el de las 100 líneas visibles es solo informativo
        for texto in page.locator("font[color=red]").all_inner_texts():
            if texto.strip() and "rebasa las" not in texto:
                mensajes.append(texto.strip())
        return mensajes

    def guardar(self, page):
        hecho = lambda: page.locator("text=/se ha grabado/i").count() > 0
        if not self.pulsar_hasta(page, "a:has-text('Guardar')", hecho, espera=15000):
            raise RuntimeError("PSA Parts no ha confirmado el litigio (no aparece «se ha grabado»): revísalo en Lista de litigios")
        texto = page.locator("text=/se ha grabado/i").first.inner_text()
        numero = re.search(r"(\d{6,})\s+se ha grabado", texto) or re.search(r"(\d{6,})", texto)
        avisos = [l.strip() for l in page.locator("body").inner_text().splitlines() if re.match(r"^\s*I\b", l)]
        return (numero.group(1) if numero else None), avisos

    # ---------- todo junto ----------

    def ejecutar(self):
        """{litigio, doc_ventas, factura, pos, fecha_factura, avisos}. LitigioNoCreado si no se guarda."""
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=not self.visible)
            try:
                page = navegador.new_context(http_credentials={"username": self.usuario, "password": self.password}).new_page()
                self.log("Entrando en PSA Parts...")
                self.abrir_menu(page)
                self.log(f"Buscando las facturas de {self.referencia}...")
                lineas = self.lineas_facturadas(page)
                linea = elegir_linea(lineas, self.referencia, self.hoy)
                if linea is None:
                    if lineas:
                        raise LitigioNoCreado(f"ninguna factura de {self.referencia} es de los últimos {MAX_DIAS_FACTURA} días. Encontradas: {resumen_lineas(lineas)}")
                    raise LitigioNoCreado(f"no hay facturas de {self.referencia} desde el {desde_para_busqueda(self.hoy):%d/%m/%Y}")
                self.log(f"Factura {linea['factura']} del {linea['fecha']:%d/%m} (Doc.ventas {linea['doc_ventas']}, línea {linea['pos_factura']}). Preparando el litigio...")
                self.crear(page, linea)
                datos = dict(linea, referencia=self.referencia, motivo=texto_motivo(self.motivo), cantidad=self.cantidad)
                if self.confirmar is not None and not self.confirmar(datos):
                    raise LitigioNoCreado("no se ha guardado: cancelado antes de «Guardar»")
                self.log("Guardando el litigio...")
                numero, avisos = self.guardar(page)
                if not numero:
                    raise RuntimeError("el litigio se ha grabado pero no se ha podido leer su número: búscalo en Lista de litigios")
                return {"litigio": numero, "doc_ventas": linea["doc_ventas"], "factura": linea["factura"], "pos": linea["pos_factura"],
                        "fecha_factura": linea["fecha"], "avisos": avisos}
            finally:
                try:
                    navegador.close()
                except Exception:
                    pass

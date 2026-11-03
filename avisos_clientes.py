"""Avisos a clientes (no suministra / sustituciones): registro de lo enviado para no repetirlo, historial por
taller, textos de los correos editables y alta de correos en el Excel de clientes. Las ventanas de esta parte
(correos que faltan, textos e historial) también están aquí."""
import json
import os
import re
import threading
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import customtkinter as ctk
import openpyxl

import estilo

NO_SUMINISTRA = "no_suministra"
SUSTITUCION = "sustitucion"
NOMBRE_TIPO = {NO_SUMINISTRA: "no suministra", SUSTITUCION: "sustitución"}
MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")

# ---------- textos de los correos ----------

TEXTOS_POR_DEFECTO = {
    "asunto_no_suministra": "Aviso referencia no suministrada {fecha}",
    "no_suministra_una": "Buenas tardes.\n\nLa referencia {referencia} correspondiente al pedido del día {fecha} no suministra.\n\nGracias por su atención,",
    "no_suministra_varias": "Buenas tardes.\n\nLas siguientes referencias correspondientes al pedido del día {fecha} no suministran:\n\n{referencias}\n\nGracias por su atención,",
    "asunto_sustitucion": "Aviso referencia sustituida {fecha}",
    "sustitucion_una": "Buenas tardes.\n\nLa referencia {original} correspondiente al pedido del día {fecha} ha sido sustituida por la referencia {sustituta}.\n\nGracias por su atención,",
    "sustitucion_varias": "Buenas tardes.\n\nLas siguientes referencias correspondientes al pedido del día {fecha} han sido sustituidas:\n\n{referencias}\n\nGracias por su atención,",
}
# Qué huecos se pueden usar en cada texto (y con qué valor de ejemplo se comprueba al guardar)
HUECOS = {
    "asunto_no_suministra": {"fecha": "28-10-2026"},
    "no_suministra_una": {"fecha": "28-10-2026", "referencia": "9812345680"},
    "no_suministra_varias": {"fecha": "28-10-2026", "referencias": "- 9812345680\n- 1629123280"},
    "asunto_sustitucion": {"fecha": "28-10-2026"},
    "sustitucion_una": {"fecha": "28-10-2026", "original": "9812345680", "sustituta": "9812345699"},
    "sustitucion_varias": {"fecha": "28-10-2026", "referencias": "- 9812345680 → 9812345699"},
}
TITULOS_TEXTOS = {
    "asunto_no_suministra": "Asunto · no suministra",
    "no_suministra_una": "No suministra · una referencia",
    "no_suministra_varias": "No suministra · varias referencias",
    "asunto_sustitucion": "Asunto · sustitución",
    "sustitucion_una": "Sustitución · una referencia",
    "sustitucion_varias": "Sustitución · varias referencias",
}


def textos(config):
    """Textos en uso: los guardados en la configuración (clave textos_avisos) o los de siempre."""
    guardados = (config or {}).get("textos_avisos") or {}
    return {clave: guardados.get(clave) or valor for clave, valor in TEXTOS_POR_DEFECTO.items()}


def comprobar_texto(clave, texto):
    """None si el texto es válido; si no, el motivo (hueco que no existe o llaves sin cerrar)."""
    if not texto.strip():
        return f"«{TITULOS_TEXTOS[clave]}» está vacío."
    try:
        texto.format(**HUECOS[clave])
    except KeyError as e:
        disponibles = ", ".join("{" + h + "}" for h in HUECOS[clave])
        return f"En «{TITULOS_TEXTOS[clave]}» no existe el hueco {{{e.args[0]}}}. Puedes usar: {disponibles}."
    except (ValueError, IndexError):
        return f"En «{TITULOS_TEXTOS[clave]}» hay una llave {{ }} sin cerrar o vacía."
    return None


def cuerpo_no_suministra(referencias, fecha, textos_avisos=None):
    t = textos_avisos or TEXTOS_POR_DEFECTO
    if len(referencias) == 1:
        return t["no_suministra_una"].format(fecha=fecha, referencia=referencias[0])
    return t["no_suministra_varias"].format(fecha=fecha, referencias="\n".join(f"- {ref}" for ref in referencias))


def cuerpo_sustitucion(sustituciones, fecha, textos_avisos=None):
    t = textos_avisos or TEXTOS_POR_DEFECTO
    if len(sustituciones) == 1:
        original, final = sustituciones[0]
        return t["sustitucion_una"].format(fecha=fecha, original=original, sustituta=final)
    return t["sustitucion_varias"].format(fecha=fecha, referencias="\n".join(f"- {o} → {f}" for o, f in sustituciones))


def asunto(tipo, fecha, textos_avisos=None):
    return (textos_avisos or TEXTOS_POR_DEFECTO)[f"asunto_{tipo}"].format(fecha=fecha)


# ---------- registro de avisos (avisos_enviados.json) ----------

class RegistroAvisos:
    """Lista de avisos gestionados: {hoja, cliente, nombre, tipo, referencias, cuando, estado ('enviado' o
    'borrador'), email}. Solo lo enviado cuenta como avisado."""

    def __init__(self, ruta):
        self.ruta = str(ruta)
        self._bloqueo = threading.Lock()

    def cargar(self):
        try:
            with open(self.ruta, encoding="utf-8") as f:
                datos = json.load(f)
            return datos if isinstance(datos, list) else []
        except (OSError, ValueError):
            return []

    def anotar(self, nuevos):
        if not nuevos:
            return
        with self._bloqueo:
            datos = self.cargar() + list(nuevos)
            temporal = f"{self.ruta}.{os.getpid()}.tmp"
            with open(temporal, "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=1)
            os.replace(temporal, self.ruta)

    def ya_avisadas(self, hoja, cliente, tipo):
        return {str(ref) for r in self.cargar() if r.get("estado") == "enviado" and r.get("hoja") == hoja
                and r.get("cliente") == cliente and r.get("tipo") == tipo for ref in r.get("referencias", [])}

    def resumen_hoja(self, hoja):
        """(correos enviados de esa hoja, hora del último) o (0, None)."""
        enviados = [r for r in self.cargar() if r.get("hoja") == hoja and r.get("estado") == "enviado"]
        if not enviados:
            return 0, None
        return len(enviados), max(datetime.fromisoformat(r["cuando"]) for r in enviados)

    def del_mes(self, anio, mes):
        return [r for r in self.cargar() if r.get("estado") == "enviado" and _mes(r) == (anio, mes)]

    def referencias_mes(self, cliente, anio, mes):
        return sum(len(r.get("referencias", [])) for r in self.del_mes(anio, mes) if r.get("cliente") == cliente)

    def historial_mes(self, anio, mes):
        """Por taller: [{cliente, nombre, avisos, referencias, no_suministra, sustituciones, ultimo}], de más a menos referencias."""
        por_cliente = defaultdict(lambda: {"avisos": 0, "referencias": 0, NO_SUMINISTRA: 0, SUSTITUCION: 0, "nombre": "", "ultimo": None})
        for r in self.del_mes(anio, mes):
            fila = por_cliente[r["cliente"]]
            fila["avisos"] += 1
            fila["referencias"] += len(r.get("referencias", []))
            fila[r.get("tipo", NO_SUMINISTRA)] += len(r.get("referencias", []))
            fila["nombre"] = r.get("nombre") or fila["nombre"]
            cuando = datetime.fromisoformat(r["cuando"])
            fila["ultimo"] = max(fila["ultimo"], cuando) if fila["ultimo"] else cuando
        filas = [{"cliente": cliente, "nombre": d["nombre"], "avisos": d["avisos"], "referencias": d["referencias"],
                  "no_suministra": d[NO_SUMINISTRA], "sustituciones": d[SUSTITUCION], "ultimo": d["ultimo"]}
                 for cliente, d in por_cliente.items()]
        return sorted(filas, key=lambda f: (-f["referencias"], f["cliente"]))

    def meses(self):
        return sorted({_mes(r) for r in self.cargar() if r.get("estado") == "enviado"}, reverse=True)


def normalizar_ref(referencia):
    """«2/1434 C8» -> «1434C8» (las referencias de la hoja de pedido pueden llevar espacios)."""
    return re.sub(r"\s+", "", str(referencia or "")).replace("2/", "").upper()


def fecha_del_aviso(registro):
    """Día del pedido (por el nombre de la hoja dd-mm-aaaa) o, si no se puede, el día en que se avisó."""
    encontrado = re.search(r"(\d{2})-(\d{2})-(\d{4})", str(registro.get("hoja") or ""))
    if encontrado:
        try:
            return date(int(encontrado.group(3)), int(encontrado.group(2)), int(encontrado.group(1)))
        except ValueError:
            pass
    return datetime.fromisoformat(registro["cuando"]).date()


def no_suministra_previas(registro, lineas, dias=60, hoy=None):
    """Referencias de 'lineas' ([referencia, cantidad, cliente], como el Excel urgente) que ya salieron como NO SUMINISTRA en
    los avisos de los últimos 'dias': [{referencia, cantidad, fechas, clientes}] en el orden del Excel."""
    hoy = hoy or date.today()
    vistas = defaultdict(lambda: {"fechas": set(), "clientes": set()})
    for r in registro.cargar():
        if r.get("tipo") != NO_SUMINISTRA:
            continue
        fecha = fecha_del_aviso(r)
        if (hoy - fecha).days > dias:
            continue
        for ref in r.get("referencias", []):
            vistas[normalizar_ref(ref)]["fechas"].add(fecha)
            vistas[normalizar_ref(ref)]["clientes"].add(str(r.get("cliente", "")))
    resultado, ya = [], set()
    for referencia, cantidad, *_ in lineas:
        clave = normalizar_ref(referencia)
        if clave in vistas and clave not in ya:
            ya.add(clave)
            resultado.append({"referencia": clave, "cantidad": cantidad, "fechas": sorted(vistas[clave]["fechas"]),
                              "clientes": sorted(c for c in vistas[clave]["clientes"] if c)})
    return resultado


def texto_fechas(fechas):
    textos = [f"{f:%d/%m}" for f in fechas]
    return textos[0] if len(textos) == 1 else ", ".join(textos[:-1]) + " y " + textos[-1]


def _mes(registro):
    cuando = datetime.fromisoformat(registro["cuando"])
    return cuando.year, cuando.month


def nombre_mes(anio, mes):
    return f"{MESES[mes - 1]} {anio}"


# ---------- Excel de clientes ----------

def columnas_clientes(cabecera, columnas):
    """Índices (nombre, email) de clientes.xlsx. Se buscan por la cabecera (CORREO/EMAIL, NOMBRE/TALLER); si no se
    reconoce: A=num_cliente, B=email con dos columnas, o A=num_cliente, B=nombre, C=email con tres."""
    titulos = [str(valor or "").strip().upper() for valor in cabecera]
    email = next((i for i, t in enumerate(titulos) if i > 0 and ("CORREO" in t or "MAIL" in t)), None)
    nombre = next((i for i, t in enumerate(titulos) if i > 0 and ("NOMBRE" in t or "TALLER" in t)), None)
    if email is None:
        email = 1 if columnas <= 2 else 2
        if nombre is None and columnas > 2:
            nombre = 1
    return nombre, email


def guardar_correos_clientes(ruta, correos, normalizar=str):
    """Escribe {num_cliente: email} en el Excel de clientes: rellena el correo si el cliente ya está y, si no,
    añade una fila al final. Lanza PermissionError si el Excel está abierto. Devuelve (actualizados, añadidos)."""
    wb = openpyxl.load_workbook(ruta)
    try:
        ws = wb.active
        cabecera = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        _, col_email = columnas_clientes(cabecera, ws.max_column)
        filas = {}
        for fila in range(2, ws.max_row + 1):
            numero = normalizar(ws.cell(row=fila, column=1).value)
            if numero:
                filas[numero] = fila
        actualizados = anadidos = 0
        for cliente, email in correos.items():
            if cliente in filas:
                ws.cell(row=filas[cliente], column=col_email + 1).value = email
                actualizados += 1
            else:
                nueva = ws.max_row + 1
                ws.cell(row=nueva, column=1).value = int(cliente) if str(cliente).isdigit() else cliente
                ws.cell(row=nueva, column=col_email + 1).value = email
                anadidos += 1
        wb.save(ruta)
        return actualizados, anadidos
    finally:
        wb.close()


# ---------- ventanas ----------

def _ventana(app, titulo, ancho=None):
    ventana = ctk.CTkToplevel(app)
    ventana.title(titulo)
    ventana.configure(fg_color=app.color_panel_claro)
    ventana.transient(app)
    tarjeta = ctk.CTkFrame(ventana, fg_color=app.color_superficie, corner_radius=estilo.RADIO_TARJETA, border_width=1, border_color=app.color_borde,
                           **({"width": ancho} if ancho else {}))
    tarjeta.pack(padx=16, pady=16, fill="both", expand=True)
    return ventana, tarjeta


def _centrar(app, ventana):
    ventana.update_idletasks()
    x = app.winfo_rootx() + (app.winfo_width() - ventana.winfo_width()) // 2
    y = app.winfo_rooty() + max(40, (app.winfo_height() - ventana.winfo_height()) // 3)
    ventana.geometry(f"+{max(0, x)}+{max(0, y)}")


class DialogoCorreosClientes:
    """Clientes cuyo aviso se quedó sin destinatario (no están en el Excel o no tienen correo). Se escribe el
    correo, se guarda en el Excel de clientes y al_guardar({cliente: email}) envía o completa sus borradores;
    si al_guardar devuelve False (p. ej. Excel abierto) la ventana sigue abierta."""

    def __init__(self, app, entradas, al_guardar, envio_automatico=True, validar=lambda texto: []):
        self.app = app
        self.al_guardar = al_guardar
        self.validar = validar
        self.campos = {}
        self.cerrado = False
        por_cliente = defaultdict(list)
        for entrada in entradas:
            por_cliente[entrada["cliente"]].append(entrada)
        self.por_cliente = dict(por_cliente)
        ventana, tarjeta = _ventana(app, "Clientes sin correo")
        self.ventana = ventana
        ventana.protocol("WM_DELETE_WINDOW", self.cerrar)
        ventana.bind("<Escape>", lambda evento: (self.cerrar(), "break")[1])
        cuantos = len(self.por_cliente)
        ctk.CTkLabel(tarjeta, text=f"{cuantos} cliente{'s' if cuantos != 1 else ''} sin correo", font=estilo.fuente_titulo(),
                     text_color=app.color_texto).pack(anchor="w", padx=estilo.PAD_X, pady=(14, 2))
        accion = "se envían sus avisos" if envio_automatico else "se completan sus borradores"
        app.crear_etiqueta(tarjeta, f"Sus avisos están en Borradores sin destinatario. Escribe el correo (varios separados por ;): se guarda en el Excel de clientes y {accion}.",
                           secundario=True, wraplength=560, justify="left").pack(anchor="w", padx=estilo.PAD_X, pady=(0, 10))
        lista = ctk.CTkScrollableFrame(tarjeta, fg_color="transparent", height=min(320, 78 * cuantos)) if cuantos > 4 else ctk.CTkFrame(tarjeta, fg_color="transparent")
        lista.pack(padx=estilo.PAD_X - 4, fill="both", expand=True)
        for cliente, suyas in self.por_cliente.items():
            fila = ctk.CTkFrame(lista, fg_color=app.color_superficie_alt, corner_radius=10)
            fila.pack(fill="x", pady=(0, 6))
            campo = app.crear_entrada(fila, width=240, placeholder_text="correo@taller.es")
            campo.pack(side="right", padx=10, pady=8)
            self.campos[cliente] = campo
            texto = ctk.CTkFrame(fila, fg_color="transparent")
            texto.pack(side="left", fill="x", expand=True, padx=(12, 6), pady=8)
            motivo = "no está en el Excel de clientes" if any(e.get("motivo") == "no_en_excel" for e in suyas) else "sin correo en el Excel de clientes"
            nombre = next((e.get("nombre") for e in suyas if e.get("nombre")), "")
            app.crear_etiqueta(texto, f"Cliente {cliente}" + (f" · {nombre}" if nombre else ""), negrita=True).pack(anchor="w")
            detalle = " · ".join(f"{NOMBRE_TIPO[e['tipo']]}: {', '.join(e['detalle'])}" for e in suyas)
            app.crear_etiqueta(texto, f"{motivo} — {detalle}", estilo.TAM_PEQUENO, secundario=True, wraplength=300, justify="left").pack(anchor="w")
        pie = ctk.CTkFrame(tarjeta, fg_color="transparent")
        pie.pack(padx=estilo.PAD_X, pady=(8, 14), fill="x")
        self.btn_guardar = app.crear_boton_principal(pie, "Guardar y enviar" if envio_automatico else "Guardar en los borradores", self.guardar, alto=estilo.ALTO_BOTON_COMPACTO)
        self.btn_guardar.pack(side="right")
        app.crear_boton_secundario(pie, "Ahora no", self.cerrar, alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right", padx=(0, 6))
        _centrar(app, ventana)
        ventana.after(150, lambda: (not self.cerrado and ventana.winfo_exists()) and (ventana.lift(), ventana.focus_force()))

    def correos(self):
        return {cliente: campo.get().strip() for cliente, campo in self.campos.items() if campo.get().strip()}

    def guardar(self):
        correos = self.correos()
        if not correos:
            self.app.avisar("Escribe al menos un correo, o pulsa «Ahora no».")
            return False
        invalidos = [direccion for email in correos.values() for direccion in self.validar(email)]
        if invalidos:
            self.app.avisar("Estos correos no parecen válidos: " + ", ".join(invalidos), "aviso")
            return False
        if self.al_guardar(correos) is False:
            return False
        self.cerrar()
        return True

    def cerrar(self):
        if self.cerrado:
            return
        self.cerrado = True
        try:
            self.ventana.destroy()
        except Exception:
            pass


class DialogoTextosAvisos:
    """Editar los textos de los avisos. al_guardar(dict) los guarda (None = volver a los de siempre)."""

    def __init__(self, app, actuales, al_guardar):
        self.app = app
        self.al_guardar = al_guardar
        self.campos = {}
        ventana, tarjeta = _ventana(app, "Textos de los avisos")
        self.ventana = ventana
        ventana.bind("<Escape>", lambda evento: (self.ventana.destroy(), "break")[1])
        ctk.CTkLabel(tarjeta, text="Textos de los avisos a clientes", font=estilo.fuente_titulo(), text_color=app.color_texto).pack(anchor="w", padx=estilo.PAD_X, pady=(14, 2))
        app.crear_etiqueta(tarjeta, "Los huecos entre llaves se rellenan solos: {fecha} es la fecha de la hoja, {referencia} o {referencias} las piezas "
                           "(la lista ya va con un guion por línea) y {original} / {sustituta} la pieza sustituida. La firma de Outlook se añade al final.",
                           secundario=True, wraplength=620, justify="left").pack(anchor="w", padx=estilo.PAD_X, pady=(0, 8))
        cuerpo = ctk.CTkScrollableFrame(tarjeta, fg_color="transparent", width=640, height=430)
        cuerpo.pack(padx=estilo.PAD_X - 4, fill="both", expand=True)
        for clave, titulo in TITULOS_TEXTOS.items():
            app.crear_etiqueta(cuerpo, titulo, negrita=True).pack(anchor="w", pady=(6, 2))
            if clave.startswith("asunto_"):
                campo = app.crear_entrada(cuerpo)
                campo.insert(0, actuales[clave])
                campo.pack(fill="x")
            else:
                campo = ctk.CTkTextbox(cuerpo, height=110, font=estilo.fuente(), fg_color=app.color_superficie_alt, text_color=app.color_texto,
                                       border_width=1, border_color=app.color_borde, wrap="word")
                campo.insert("1.0", actuales[clave])
                campo.pack(fill="x")
            self.campos[clave] = campo
        pie = ctk.CTkFrame(tarjeta, fg_color="transparent")
        pie.pack(padx=estilo.PAD_X, pady=(10, 14), fill="x")
        app.crear_boton_secundario(pie, "Volver a los de siempre", self.restaurar, alto=estilo.ALTO_BOTON_COMPACTO).pack(side="left")
        app.crear_boton_principal(pie, "Guardar", self.guardar, alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right")
        app.crear_boton_secundario(pie, "Cancelar", self.ventana.destroy, alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right", padx=(0, 6))
        _centrar(app, ventana)

    def valores(self):
        return {clave: (campo.get() if clave.startswith("asunto_") else campo.get("1.0", "end-1c")).strip() for clave, campo in self.campos.items()}

    def guardar(self):
        valores = self.valores()
        for clave, texto in valores.items():
            error = comprobar_texto(clave, texto)
            if error:
                self.app.avisar(error, "aviso")
                return False
        self.al_guardar({clave: texto for clave, texto in valores.items() if texto != TEXTOS_POR_DEFECTO[clave]} or None)
        self.ventana.destroy()
        return True

    def restaurar(self):
        for clave, campo in self.campos.items():
            if clave.startswith("asunto_"):
                campo.delete(0, "end")
                campo.insert(0, TEXTOS_POR_DEFECTO[clave])
            else:
                campo.delete("1.0", "end")
                campo.insert("1.0", TEXTOS_POR_DEFECTO[clave])


class DialogoHistorialAvisos:
    """Avisos enviados por taller en un mes."""

    def __init__(self, app, registro, hoy=None):
        self.app = app
        self.registro = registro
        hoy = hoy or datetime.now()
        self.meses = registro.meses() or [(hoy.year, hoy.month)]
        if (hoy.year, hoy.month) not in self.meses:
            self.meses.insert(0, (hoy.year, hoy.month))
        ventana, tarjeta = _ventana(app, "Historial de avisos")
        self.ventana = ventana
        ventana.bind("<Escape>", lambda evento: (self.ventana.destroy(), "break")[1])
        cabecera = ctk.CTkFrame(tarjeta, fg_color="transparent")
        cabecera.pack(fill="x", padx=estilo.PAD_X, pady=(14, 4))
        ctk.CTkLabel(cabecera, text="Avisos enviados por taller", font=estilo.fuente_titulo(), text_color=app.color_texto).pack(side="left")
        self.selector = app.crear_menu(cabecera, [nombre_mes(*m) for m in self.meses], comando=lambda valor: self.mostrar())
        self.selector.set(nombre_mes(*self.meses[0]))
        self.selector.pack(side="right")
        self.total = app.crear_etiqueta(tarjeta, "", secundario=True)
        self.total.pack(anchor="w", padx=estilo.PAD_X, pady=(0, 6))
        self.lista = ctk.CTkScrollableFrame(tarjeta, fg_color="transparent", width=560, height=360)
        self.lista.pack(padx=estilo.PAD_X - 4, pady=(0, 14), fill="both", expand=True)
        self.mostrar()
        _centrar(app, ventana)

    def mes_elegido(self):
        nombres = [nombre_mes(*m) for m in self.meses]
        return self.meses[nombres.index(self.selector.get())] if self.selector.get() in nombres else self.meses[0]

    def mostrar(self):
        for hijo in self.lista.winfo_children():
            hijo.destroy()
        filas = self.registro.historial_mes(*self.mes_elegido())
        referencias = sum(f["referencias"] for f in filas)
        correos = sum(f["avisos"] for f in filas)
        self.total.configure(text=f"{len(filas)} taller{'es' if len(filas) != 1 else ''} · {correos} correo{'s' if correos != 1 else ''} · "
                                  f"{referencias} referencia{'s' if referencias != 1 else ''}" if filas else "Ningún aviso enviado este mes.")
        app = self.app
        for f in filas:
            fila = ctk.CTkFrame(self.lista, fg_color=app.color_superficie_alt, corner_radius=10)
            fila.pack(fill="x", pady=(0, 6))
            app.crear_etiqueta(fila, f"{f['referencias']} ref.", negrita=True).pack(side="right", padx=12, pady=8)
            texto = ctk.CTkFrame(fila, fg_color="transparent")
            texto.pack(side="left", fill="x", expand=True, padx=(12, 6), pady=6)
            app.crear_etiqueta(texto, f"Cliente {f['cliente']}" + (f" · {f['nombre']}" if f["nombre"] else ""), negrita=True).pack(anchor="w")
            partes = [f"{f['avisos']} correo{'s' if f['avisos'] != 1 else ''}"]
            if f["no_suministra"]:
                partes.append(f"{f['no_suministra']} no suministra")
            if f["sustituciones"]:
                partes.append(f"{f['sustituciones']} sustituida{'s' if f['sustituciones'] != 1 else ''}")
            partes.append(f"último {f['ultimo']:%d/%m}")
            app.crear_etiqueta(texto, " · ".join(partes), estilo.TAM_PEQUENO, secundario=True).pack(anchor="w")


class DialogoNoSuministra:
    """Antes de generar los pedidos urgentes: referencias que ya salieron como NO SUMINISTRA. Cada una tiene «Ticket de
    plazo» (al_ticket(coincidencia) prepara el ticket de Power-Line). al_decidir(True) genera igualmente; False cancela."""

    def __init__(self, app, coincidencias, al_ticket, al_decidir):
        self.app = app
        self.al_ticket = al_ticket
        self.al_decidir = al_decidir
        self.cerrado = False
        self.botones = {}
        ventana, tarjeta = _ventana(app, "Referencias que no suministran")
        self.ventana = ventana
        ventana.protocol("WM_DELETE_WINDOW", lambda: self.decidir(False))
        ventana.bind("<Escape>", lambda evento: (self.decidir(False), "break")[1])
        cuantas = len(coincidencias)
        ctk.CTkLabel(tarjeta, text=f"{cuantas} referencia{'s' if cuantas != 1 else ''} ya salió como NO SUMINISTRA" if cuantas == 1 else
                     f"{cuantas} referencias ya salieron como NO SUMINISTRA", font=estilo.fuente_titulo(), text_color=app.color_texto).pack(anchor="w", padx=estilo.PAD_X, pady=(14, 2))
        app.crear_etiqueta(tarjeta, "Según los avisos enviados a los talleres. Puedes pedir plazo a Power-Line antes de volver a pedirlas.",
                           secundario=True, wraplength=560, justify="left").pack(anchor="w", padx=estilo.PAD_X, pady=(0, 10))
        lista = ctk.CTkScrollableFrame(tarjeta, fg_color="transparent", height=min(300, 70 * cuantas)) if cuantas > 4 else ctk.CTkFrame(tarjeta, fg_color="transparent")
        lista.pack(padx=estilo.PAD_X - 4, fill="both", expand=True)
        for c in coincidencias:
            fila = ctk.CTkFrame(lista, fg_color=app.color_superficie_alt, corner_radius=10)
            fila.pack(fill="x", pady=(0, 6))
            boton = app.crear_boton_secundario(fila, "Ticket de plazo", lambda c=c: self.ticket(c), alto=estilo.ALTO_BOTON_COMPACTO, width=130)
            boton.pack(side="right", padx=10, pady=8)
            self.botones[c["referencia"]] = boton
            texto = ctk.CTkFrame(fila, fg_color="transparent")
            texto.pack(side="left", fill="x", expand=True, padx=(12, 6), pady=8)
            cantidad = f" · {c['cantidad']} ud" if str(c.get("cantidad") or "").strip() else ""
            app.crear_etiqueta(texto, f"{c['referencia']}{cantidad}", negrita=True).pack(anchor="w")
            clientes = f" (cliente{'s' if len(c['clientes']) != 1 else ''} {', '.join(c['clientes'])})" if c["clientes"] else ""
            app.crear_etiqueta(texto, f"NO SUMINISTRA el {texto_fechas(c['fechas'])}{clientes}", estilo.TAM_PEQUENO, secundario=True,
                               wraplength=360, justify="left").pack(anchor="w")
        pie = ctk.CTkFrame(tarjeta, fg_color="transparent")
        pie.pack(padx=estilo.PAD_X, pady=(8, 14), fill="x")
        app.crear_boton_principal(pie, "Generar igualmente", lambda: self.decidir(True), alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right")
        app.crear_boton_secundario(pie, "Cancelar", lambda: self.decidir(False), alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right", padx=(0, 6))
        _centrar(app, ventana)

    def ticket(self, coincidencia):
        self.al_ticket(coincidencia)
        boton = self.botones.get(coincidencia["referencia"])
        if boton is not None and boton.winfo_exists():
            boton.configure(text="Ticket preparado ✓")

    def decidir(self, generar):
        if self.cerrado:
            return
        self.cerrado = True
        try:
            self.ventana.destroy()
        except Exception:
            pass
        self.al_decidir(generar)

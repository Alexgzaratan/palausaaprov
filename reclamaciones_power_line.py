"""Módulo aislado para crear y consultar reclamaciones de Power Line."""

import atexit
import os
import shutil
import sqlite3
import tempfile
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from PIL import Image, ImageOps
from tkinterdnd2 import DND_FILES, TkinterDnD

import efectos
import estilo
import litigio_psa
import power_line_ticket
from configuracion import registrar_error_detallado
from rutas_app import ruta_datos
import json as _json

try:  # fotos HEIC/HEIF (iPhone): si la librería no está, simplemente no se admiten
    from pillow_heif import register_heif_opener
    register_heif_opener()
    ADMITE_HEIC = True
except Exception:
    ADMITE_HEIC = False


DATABASE_FILE = ruta_datos("reclamaciones_power_line.db")
LITIGIOS_PSA = ruta_datos("litigios_psa.json", migrar=False)
TICKETS_POWER_LINE = ruta_datos("tickets_power_line.json", migrar=False)  # tickets creados con las plantillas  # litigios creados desde la app (para no reclamar dos veces la misma línea)
PREFIJO_ADJUNTOS = "power_line_adjuntos_"
EXTENSIONES_FOTO = "*.jpg *.jpeg *.png *.bmp *.tif *.tiff" + (" *.heic *.heif" if ADMITE_HEIC else "")
PHOTO_FIELDS = (
    ("bulto", "Fotografía bulto completo"),
    ("deterioro", "Fotografía detalle del deterioro"),
    ("etiqueta", "Fotografía detalle de la etiqueta"),
)
TEXT_BOOKMARKS = {
    "pedido": ("PedidoSAP", "PEDIDO_SAP", "Pedido_SAP"),
    "referencia": ("Referencia", "REFERENCIA"),
    "descripcion": ("Descripcion", "DESCRIPCION", "Descripción"),
}
PHOTO_BOOKMARKS = {
    "bulto": ("FotoBultoCompleto", "FOTO_BULTO_COMPLETO", "FotoBulto"),
    "deterioro": ("FotoDetalleDeterioro", "FOTO_DETALLE_DETERIORO", "FotoDeterioro"),
    "etiqueta": ("FotoDetalleEtiqueta", "FOTO_DETALLE_ETIQUETA", "FotoEtiqueta"),
}


def leer_litigios_psa(ruta=None):
    try:
        with open(ruta or LITIGIOS_PSA, encoding="utf-8") as f:
            datos = _json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, ValueError):
        return []


def anotar_litigio_psa(resultado, referencia, ruta=None):
    ruta = ruta or LITIGIOS_PSA
    datos = leer_litigios_psa(ruta)[-499:] + [{"litigio": resultado["litigio"], "referencia": referencia, "factura": resultado["factura"],
                                               "pos": resultado["pos"], "doc_ventas": resultado["doc_ventas"], "cuando": datetime.now().isoformat(timespec="seconds")}]
    temporal = ruta + ".tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        _json.dump(datos, f, ensure_ascii=False, indent=1)
    os.replace(temporal, ruta)


def leer_tickets(ruta=None):
    try:
        with open(ruta or TICKETS_POWER_LINE, encoding="utf-8") as f:
            datos = _json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, ValueError):
        return []


def anotar_ticket(numero, plantilla, criterios, ruta=None):
    ruta = ruta or TICKETS_POWER_LINE
    datos = leer_tickets(ruta)[-199:] + [{"ticket": numero, "plantilla": plantilla, "criterios": criterios, "cuando": datetime.now().isoformat(timespec="seconds")}]
    temporal = ruta + ".tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        _json.dump(datos, f, ensure_ascii=False, indent=1)
    os.replace(temporal, ruta)


def litigio_anterior(factura, pos, ruta=None):
    """El litigio ya creado desde la app para esa factura y línea, o None."""
    return next((r for r in reversed(leer_litigios_psa(ruta)) if r.get("factura") == factura and r.get("pos") == pos), None)


def registrar_error_reclamacion(error, detalle, referencia):
    """Al registro común de la app (con rotación y sin fallar si no se puede escribir)."""
    registrar_error_detallado(f"Reclamaciones Power Line: {detalle}", error, ref=referencia or "-")


def limpiar_adjuntos_antiguos(horas=24, carpeta_temporal=None):
    """Borra las carpetas temporales de fotos que dejaron sesiones anteriores. Devuelve cuántas se borraron."""
    base = Path(carpeta_temporal or tempfile.gettempdir())
    limite = time.time() - horas * 3600
    borradas = 0
    try:
        candidatas = list(base.glob(PREFIJO_ADJUNTOS + "*"))
    except OSError:
        return 0
    for carpeta in candidatas:
        try:
            if carpeta.is_dir() and carpeta.stat().st_mtime < limite:
                shutil.rmtree(carpeta, ignore_errors=True)
                borradas += not carpeta.exists()
        except OSError:
            continue
    return borradas


def abrir_imagen(origen):
    """Imagen con la orientación correcta: las fotos del móvil en vertical guardan el giro en el EXIF."""
    imagen = Image.open(origen)
    try:
        girada = ImageOps.exif_transpose(imagen)
    except Exception:
        return imagen
    if girada is not imagen:
        imagen.close()
    return girada


class HistorialReclamaciones:
    def __init__(self, database_path=DATABASE_FILE):
        self.database_path = database_path
        conexion = self._conexion()
        try:
            with conexion:
                conexion.execute("""
                    CREATE TABLE IF NOT EXISTS reclamaciones_power_line (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        litigio TEXT NOT NULL,
                        fecha TEXT NOT NULL,
                        pedido_sap TEXT,
                        referencia TEXT,
                        ruta_word TEXT NOT NULL,
                        ruta_pdf TEXT
                    )
                """)
                columnas = {fila[1] for fila in conexion.execute("PRAGMA table_info(reclamaciones_power_line)")}
                if "ticket" not in columnas:  # nº de ticket de Power-Line (TKT...), añadido después
                    conexion.execute("ALTER TABLE reclamaciones_power_line ADD COLUMN ticket TEXT")
        finally:
            conexion.close()

    def _conexion(self):
        return sqlite3.connect(self.database_path)

    def guardar(self, litigio, pedido_sap, referencia, ruta_word, ruta_pdf):
        conexion = self._conexion()
        try:
            with conexion:
                cursor = conexion.execute(
                    "INSERT INTO reclamaciones_power_line "
                    "(litigio, fecha, pedido_sap, referencia, ruta_word, ruta_pdf) VALUES (?, ?, ?, ?, ?, ?)",
                    (litigio, datetime.now().strftime("%d/%m/%Y %H:%M"), pedido_sap, referencia, ruta_word, ruta_pdf),
                )
                return cursor.lastrowid
        finally:
            conexion.close()

    def guardar_ticket(self, registro_id, ticket):
        conexion = self._conexion()
        try:
            with conexion:
                conexion.execute("UPDATE reclamaciones_power_line SET ticket = ? WHERE id = ?", (ticket, registro_id))
        finally:
            conexion.close()

    def buscar(self, texto=""):
        filtro = f"%{texto.strip()}%"
        conexion = self._conexion()
        try:
            return conexion.execute(
                "SELECT id, litigio, fecha, pedido_sap, referencia, ruta_word, ruta_pdf, ticket "
                "FROM reclamaciones_power_line WHERE litigio LIKE ? OR referencia LIKE ? OR pedido_sap LIKE ? OR ticket LIKE ? "
                "ORDER BY id DESC",
                (filtro, filtro, filtro, filtro),
            ).fetchall()
        finally:
            conexion.close()

    def eliminar(self, registro_id):
        conexion = self._conexion()
        try:
            with conexion:
                conexion.execute("DELETE FROM reclamaciones_power_line WHERE id = ?", (registro_id,))
        finally:
            conexion.close()


class DocumentoPowerLine:
    """Rellena una copia de plantilla mediante Word para no alterar su formato."""

    def __init__(self, plantilla, salida, datos, fotos, generar_pdf):
        self.plantilla = Path(plantilla)
        self.salida = Path(salida)
        self.datos = datos
        self.fotos = fotos
        self.generar_pdf = generar_pdf

    @staticmethod
    def _preparar_foto(origen, carpeta_temporal, clave):
        destino = Path(carpeta_temporal) / f"{clave}.jpg"
        with abrir_imagen(origen) as imagen:
            imagen = imagen.convert("RGB")
            imagen.thumbnail((1920, 1920), Image.Resampling.LANCZOS)
            imagen.save(destino, "JPEG", quality=82, optimize=True)
        return str(destino)

    @staticmethod
    def _rango_marcador(documento, nombres):
        for nombre in nombres:
            if documento.Bookmarks.Exists(nombre):
                return documento.Bookmarks.Item(nombre).Range
        return None

    @staticmethod
    def _buscar_marcador(documento, etiquetas):
        for etiqueta in etiquetas:
            for historia in range(1, 8):
                try:
                    rango = documento.StoryRanges(historia)
                except Exception:
                    continue
                while rango is not None:
                    encontrado = rango.Duplicate
                    buscador = encontrado.Find
                    buscador.ClearFormatting()
                    buscador.Text = etiqueta
                    if buscador.Execute():
                        return encontrado
                    try:
                        rango = rango.NextStoryRange
                    except Exception:
                        rango = None
        return None

    @staticmethod
    def _texto_en_rango(documento, valor, nombres, etiquetas):
        rango = DocumentoPowerLine._rango_marcador(documento, nombres)
        reemplazar = rango is not None
        if rango is None:
            rango = DocumentoPowerLine._buscar_marcador(documento, etiquetas)
        if rango is None:
            raise ValueError(f"No se encontró el campo de plantilla: {etiquetas[0]}")
        if reemplazar or rango.Text.strip().startswith(("{{", "<<")):
            rango.Text = valor
            return
        rango.Collapse(0)
        rango.InsertAfter(f" {valor}")

    @staticmethod
    def _imagen_en_rango(documento, ruta_imagen, nombres, etiquetas):
        rango = DocumentoPowerLine._rango_marcador(documento, nombres)
        reemplazar = rango is not None
        if rango is None:
            rango = DocumentoPowerLine._buscar_marcador(documento, etiquetas)
        if rango is None:
            raise ValueError(f"No se encontró el espacio de plantilla: {etiquetas[0]}")
        if reemplazar or rango.Text.strip().startswith(("{{", "<<")):
            rango.Text = ""
        else:
            rango.Collapse(0)
            rango.InsertAfter("\r")
            rango.Collapse(0)
            rango.ParagraphFormat.Alignment = 1
        imagen = documento.InlineShapes.AddPicture(str(ruta_imagen), False, True, rango)
        ancho_original, alto_original = float(imagen.Width), float(imagen.Height)
        max_ancho, max_alto = 260.0, 190.0
        escala = min(max_ancho / ancho_original, max_alto / alto_original, 1.0)
        imagen.LockAspectRatio = False
        imagen.Width = round(ancho_original * escala)
        imagen.Height = round(alto_original * escala)

    def generar(self):
        if not self.plantilla.is_file():
            raise FileNotFoundError("La plantilla Word configurada no existe.")
        self.salida.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="power_line_") as temporal:
            fotos = {clave: self._preparar_foto(ruta, temporal, clave) for clave, ruta in self.fotos.items()}
            shutil.copy2(self.plantilla, self.salida)
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            word = None
            documento = None
            try:
                try:
                    word = win32com.client.DispatchEx("Word.Application")
                except Exception as error:
                    raise RuntimeError("No se pudo iniciar Microsoft Word. Comprueba que Microsoft Word está instalado y puede abrirse normalmente.") from error
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    documento = word.Documents.Open(str(self.salida))
                except Exception as error:
                    raise RuntimeError("No se pudo abrir la copia de la plantilla Word.") from error
                self._texto_en_rango(documento, self.datos["pedido"], TEXT_BOOKMARKS["pedido"], ("{{PEDIDO_SAP}}", "<<PEDIDO_SAP>>", "Nº de pedido SAP:", "Nº pedido SAP:"))
                self._texto_en_rango(documento, self.datos["referencia"], TEXT_BOOKMARKS["referencia"], ("{{REFERENCIA}}", "<<REFERENCIA>>", "Referencia:"))
                self._texto_en_rango(documento, self.datos["descripcion"], TEXT_BOOKMARKS["descripcion"], ("{{DESCRIPCION}}", "<<DESCRIPCION>>", "Descripción del problema:"))
                for clave, _ in PHOTO_FIELDS:
                    etiqueta = "{{FOTO_" + clave.upper() + "}}"
                    etiquetas_foto = {
                        "bulto": (etiqueta, "<<FOTO_BULTO>>", "Fotografía bulto Completo"),
                        "deterioro": (etiqueta, "<<FOTO_DETERIORO>>", "Fotografía Detalle del Deterioro"),
                        "etiqueta": (etiqueta, "<<FOTO_ETIQUETA>>", "Fotografía Detalle de la Etiqueta"),
                    }
                    self._imagen_en_rango(documento, fotos[clave], PHOTO_BOOKMARKS[clave], etiquetas_foto[clave])
                documento.Save()
                ruta_pdf = None
                if self.generar_pdf:
                    ruta_pdf = str(self.salida.with_suffix(".pdf"))
                    documento.ExportAsFixedFormat(ruta_pdf, 17)
                return str(self.salida), ruta_pdf
            finally:
                if documento is not None:
                    try:
                        documento.Close(SaveChanges=True)
                    except Exception:
                        pass
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()


class ModuloReclamacionesPowerLine:
    def __init__(self, app, configuracion):
        self.app = app
        self.configuracion = configuracion
        self.historial = HistorialReclamaciones()
        self.fotos = {clave: "" for clave, _ in PHOTO_FIELDS}
        self.previews = {}
        self.imagenes_preview = {}
        self._carpeta_adjuntos = None  # se crea al adjuntar la primera foto y se borra al cerrar la app
        self.dnd_activo = False
        threading.Thread(target=limpiar_adjuntos_antiguos, daemon=True).start()

    @property
    def carpeta_adjuntos(self):
        if self._carpeta_adjuntos is None:
            self._carpeta_adjuntos = Path(tempfile.mkdtemp(prefix=PREFIJO_ADJUNTOS))
            atexit.register(shutil.rmtree, str(self._carpeta_adjuntos), True)
        return self._carpeta_adjuntos

    def construir(self, tab):
        try:
            TkinterDnD.require(self.app)
            self.dnd_activo = True
        except Exception:
            self.dnd_activo = False
        panel = ctk.CTkTabview(tab, fg_color=self.app.color_panel_claro, corner_radius=10, border_width=0,
                               segmented_button_selected_color=self.app.color_principal,
                               segmented_button_selected_hover_color=self.app.color_hover,
                               segmented_button_fg_color=self.app.color_panel_medio,
                               segmented_button_unselected_color=self.app.color_panel_medio,
                               text_color=self.app.color_texto)
        panel.configure(command=lambda: self.app.resaltar_pestana_activa(panel))
        panel.pack(padx=8, pady=8, fill="both", expand=True)
        self.tab_formulario = panel.add("Formulario")
        self.tab_historial = panel.add("Historial")
        self.tab_tickets = panel.add("Tickets Power-Line")
        self.app.resaltar_pestana_activa(panel)
        self._construir_formulario()
        self._construir_historial()
        self._construir_tickets()

    def _construir_formulario(self):
        app = self.app
        contenedor = ctk.CTkScrollableFrame(self.tab_formulario, fg_color="transparent")
        contenedor.pack(padx=2, pady=2, fill="both", expand=True)
        card = app.crear_tarjeta_tab(contenedor, "Nueva reclamación", "Tres pasos: crear el litigio en PARTS, generar el Word con las fotos y enviarlo a Power-Line.", icono="aviso")
        self.referencia_entry = self._campo(card, "Referencia")
        paso1 = ctk.CTkFrame(card, fg_color="transparent")
        paso1.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        self.motivo_option = app.crear_menu(paso1, [litigio_psa.texto_motivo(c) for c, _ in litigio_psa.MOTIVOS], width=330)
        self.motivo_option.set(litigio_psa.texto_motivo(self.configuracion.get("psa_motivo") or litigio_psa.MOTIVO_POR_DEFECTO))
        self.motivo_option.pack(side="left")
        efectos.poner_pista(app, self.motivo_option, "Motivo del litigio de Calidad en PARTS")
        app.crear_etiqueta(paso1, "Cant.", secundario=True).pack(side="left", padx=(10, 4))
        self.cantidad_entry = app.crear_entrada(paso1, width=46, justify="center")
        self.cantidad_entry.insert(0, "1")
        self.cantidad_entry.pack(side="left")
        self.btn_litigio = app.crear_boton_principal(paso1, "1 · Crear litigio en PARTS", self.crear_litigio_psa, alto=estilo.ALTO_CONTROL)
        self.btn_litigio.pack(side="right")
        efectos.poner_pista(app, self.btn_litigio, "Busca la factura más reciente de la referencia (máx. 10 días) y crea el litigio de Calidad")
        self.litigio_entry = self._campo(card, "Nº Litigio *")
        self.pedido_entry = self._campo(card, "Nº Pedido SAP")
        self._crear_selector_destino(card)
        app.crear_etiqueta(card, "Descripción del problema", negrita=True).pack(anchor="w", padx=estilo.PAD_X, pady=(2, 4))
        self.descripcion_text = ctk.CTkTextbox(card, height=56, corner_radius=estilo.RADIO, border_width=1, border_color=app.color_borde,
                                               fg_color=app.color_campo, text_color=app.color_texto, font=estilo.fuente())
        self.descripcion_text.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        app.crear_etiqueta(card, "Fotografías *", negrita=True).pack(anchor="w", padx=estilo.PAD_X, pady=(2, 4))
        fotos_frame = ctk.CTkFrame(card, fg_color="transparent")
        fotos_frame.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        for clave, etiqueta in PHOTO_FIELDS:
            self._crear_selector_foto(fotos_frame, clave, etiqueta)
        self.pdf_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(card, text="Generar también PDF", variable=self.pdf_var, fg_color=app.color_principal,
                        hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_X, pady=(0, estilo.PAD_Y))
        acciones = ctk.CTkFrame(card, fg_color="transparent")
        acciones.pack(padx=estilo.PAD_X, pady=(0, 14), fill="x")
        self.btn_generar = app.crear_boton_principal(acciones, "2 · Generar documento", self.generar_documento)
        self.btn_generar.pack(fill="x")
        self.btn_paso3 = app.crear_boton_principal(acciones, "3 · Enviar a Power-Line", self.enviar_ultimo_a_power_line)
        self.btn_paso3.configure(state="disabled")
        self.btn_paso3.pack(fill="x", pady=(8, 0))
        efectos.poner_pista(app, self.btn_paso3, "Crea el ticket de seguimiento con el Nº de litigio y le adjunta el Word recién generado")
        app.crear_boton_secundario(acciones, "Limpiar formulario", self.limpiar_formulario).pack(fill="x", pady=(8, 0))

    def _crear_selector_destino(self, card):
        app = self.app
        destino = ctk.CTkFrame(card, fg_color=app.color_superficie_alt, corner_radius=10)
        destino.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        app.crear_etiqueta(destino, "Carpeta de destino", negrita=True).pack(anchor="w", padx=estilo.PAD_FILA, pady=(estilo.PAD_FILA - 4, 4))
        destino_fila = ctk.CTkFrame(destino, fg_color="transparent")
        destino_fila.pack(padx=estilo.PAD_FILA, pady=(0, estilo.PAD_FILA - 2), fill="x")
        self.destino_entry = app.crear_entrada(destino_fila)
        self.destino_entry.insert(0, self.configuracion.get("power_line_ruta_guardado", ""))
        self.destino_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        app.crear_boton_secundario(destino_fila, "Elegir carpeta", self.cambiar_carpeta, width=122).pack(side="right")

    def _campo(self, parent, texto):
        self.app.crear_etiqueta(parent, texto, negrita=True).pack(anchor="w", padx=estilo.PAD_X, pady=(2, 4))
        entrada = self.app.crear_entrada(parent)
        entrada.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        return entrada

    def _crear_selector_foto(self, parent, clave, etiqueta):
        app = self.app
        marco = ctk.CTkFrame(parent, fg_color=app.color_superficie_alt, corner_radius=10)
        marco.pack(fill="x", pady=3)
        vista = ctk.CTkLabel(marco, text="", width=54)
        vista.pack(side="left", padx=(8, 0), pady=5)
        etiqueta_estado = app.crear_etiqueta(marco, etiqueta + "\nArrastra desde el explorador o un adjunto del correo", estilo.TAM_PEQUENO, justify="left", anchor="w")
        etiqueta_estado.pack(side="left", padx=12, pady=8, fill="x", expand=True)
        boton = app.crear_boton_secundario(marco, "Seleccionar", lambda: self.seleccionar_foto(clave), alto=estilo.ALTO_BOTON_COMPACTO, width=94)
        boton.pack(side="right", padx=(4, 8), pady=8)
        eliminar = ctk.CTkButton(marco, text="×", width=30, height=estilo.ALTO_BOTON_COMPACTO, corner_radius=estilo.RADIO, fg_color=app.color_peligro,
                                 hover_color=app.color_peligro_hover, command=lambda: self.eliminar_foto(clave))
        eliminar.pack(side="right", padx=2, pady=8)
        self.previews[clave] = (vista, etiqueta_estado, boton, eliminar, etiqueta)
        if self.dnd_activo:
            marco.drop_target_register(DND_FILES)
            marco.dnd_bind("<<Drop>>", lambda event, foto=clave: self._soltar_foto(event, foto))

    def seleccionar_foto(self, clave):
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", EXTENSIONES_FOTO), ("Todos los archivos", "*.*")])
        if ruta:
            self.establecer_foto(clave, ruta)

    def _soltar_foto(self, event, clave):
        rutas = self.app.tk.splitlist(event.data)
        if rutas and Path(rutas[0]).is_file():
            self.establecer_foto(clave, rutas[0])
        else:
            messagebox.showwarning("Adjunto no disponible", "No se pudo leer el adjunto arrastrado. Guárdalo desde el correo y vuelve a arrastrarlo o selecciónalo.")
        return "copy"

    def establecer_foto(self, clave, ruta):
        try:
            with Image.open(ruta) as imagen:
                imagen.verify()
        except Exception:
            messagebox.showerror("Error", "El archivo seleccionado no es una imagen válida.")
            return
        extension = Path(ruta).suffix.lower() or ".jpg"
        copia = self.carpeta_adjuntos / f"{clave}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{extension}"
        shutil.copy2(ruta, copia)
        self.fotos[clave] = str(copia)
        vista, etiqueta, _, _, titulo = self.previews[clave]
        with abrir_imagen(ruta) as imagen:
            miniatura = imagen.convert("RGB")
            miniatura.thumbnail((50, 50), Image.Resampling.LANCZOS)
            preview = ctk.CTkImage(light_image=miniatura.copy(), dark_image=miniatura.copy(), size=miniatura.size)
        self.imagenes_preview[clave] = preview
        vista.configure(image=preview)
        etiqueta.configure(text=f"{titulo}\n{Path(ruta).name}  ✓")

    def eliminar_foto(self, clave):
        self.fotos[clave] = ""
        vista, etiqueta, _, _, titulo = self.previews[clave]
        self.imagenes_preview.pop(clave, None)
        vista.configure(image=None)
        etiqueta.configure(text=titulo + "\nArrastra una imagen o selecciónala")

    def limpiar_formulario(self):
        if not messagebox.askyesno("Limpiar formulario", "¿Borrar Nº Litigio, Pedido, Referencia, descripción y las fotografías adjuntadas?"):
            return
        self.litigio_entry.delete(0, "end")
        self.pedido_entry.delete(0, "end")
        self.referencia_entry.delete(0, "end")
        self.descripcion_text.delete("1.0", "end")
        for clave, _ in PHOTO_FIELDS:
            self.eliminar_foto(clave)
        self.pdf_var.set(False)
        self.cantidad_entry.delete(0, "end")
        self.cantidad_entry.insert(0, "1")
        self._ultimo_generado = None
        self.btn_paso3.configure(state="disabled")

    def cambiar_carpeta(self):
        ruta = filedialog.askdirectory(title="Selecciona la carpeta de guardado")
        if ruta:
            self.configuracion["power_line_ruta_guardado"] = ruta
            if hasattr(self, "destino_entry"):
                self.destino_entry.delete(0, "end")
                self.destino_entry.insert(0, ruta)
            if hasattr(self, "guardado_entry"):
                self.guardado_entry.delete(0, "end")
                self.guardado_entry.insert(0, ruta)
            self.app.guardar_configuracion_ui_internal()
            self.app.actualizar_estado("carpeta de Power Line actualizada", ruta)

    def _construir_historial(self):
        app = self.app
        card = app.crear_tarjeta_tab(self.tab_historial, "Historial de reclamaciones", "Busca, abre o elimina registros; eliminar no borra los documentos.", icono="reloj")
        fila = ctk.CTkFrame(card, fg_color="transparent")
        fila.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        self.busqueda_entry = app.crear_entrada(fila, placeholder_text="Nº litigio, referencia, pedido SAP o ticket...")
        self.busqueda_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.busqueda_entry.bind("<KeyRelease>", lambda _event: self.actualizar_historial())
        app.crear_boton_secundario(fila, "Buscar", self.actualizar_historial, width=76).pack(side="right")
        tabla_frame = ctk.CTkFrame(card, fg_color=app.color_campo, corner_radius=estilo.RADIO, border_width=1, border_color=app.color_borde)
        tabla_frame.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="both", expand=True)
        estilo_tabla = ttk.Style()
        estilo_tabla.theme_use("clam")
        estilo_tabla.configure("PowerLine.Treeview", background=app.color_campo, fieldbackground=app.color_campo,
                               foreground=app.color_texto, rowheight=28, borderwidth=0, font=estilo.fuente(estilo.TAM_PEQUENO))
        estilo_tabla.configure("PowerLine.Treeview.Heading", background=app.color_panel_medio, foreground=app.color_texto,
                               font=estilo.fuente(estilo.TAM_PEQUENO, True), relief="flat", borderwidth=0)
        estilo_tabla.map("PowerLine.Treeview", background=[("selected", app.color_principal)], foreground=[("selected", "#ffffff")])
        columnas = ("litigio", "fecha", "pedido", "referencia", "ticket")
        self.tabla = ttk.Treeview(tabla_frame, columns=columnas, show="headings", height=7, style="PowerLine.Treeview")
        for clave, texto, ancho in (("litigio", "Nº Litigio", 115), ("fecha", "Fecha", 120), ("pedido", "Pedido SAP", 120), ("referencia", "Referencia", 150),
                                    ("ticket", "Ticket Power-Line", 130)):
            self.tabla.heading(clave, text=texto)
            self.tabla.column(clave, width=ancho, anchor="w")
        self.tabla.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        barra = ttk.Scrollbar(tabla_frame, orient="vertical", command=self.tabla.yview)
        barra.pack(side="right", fill="y", pady=4)
        self.tabla.configure(yscrollcommand=barra.set)
        acciones = ctk.CTkFrame(card, fg_color="transparent")
        acciones.pack(padx=estilo.PAD_X, pady=(0, 14), fill="x")
        app.crear_boton_secundario(acciones, "Abrir documento", self.abrir_documento).pack(side="left")
        app.crear_boton_secundario(acciones, "Abrir carpeta", self.abrir_carpeta).pack(side="left", padx=8)
        self.btn_enviar_power_line = app.crear_boton_principal(acciones, "Enviar a Power-Line", self.enviar_seleccionado_a_power_line, alto=estilo.ALTO_CONTROL)
        self.btn_enviar_power_line.pack(side="left")
        efectos.poner_pista(self.app, self.btn_enviar_power_line, "Crea el ticket de seguimiento con el Nº de litigio y le adjunta el Word")
        ctk.CTkButton(acciones, text="Eliminar registro", height=estilo.ALTO_CONTROL, corner_radius=estilo.RADIO, fg_color=app.color_peligro, hover_color=app.color_peligro_hover, command=self.eliminar_registro).pack(side="right")
        self.actualizar_historial()

    def actualizar_historial(self):
        if not hasattr(self, "tabla"):
            return
        for fila in self.tabla.get_children():
            self.tabla.delete(fila)
        for registro in self.historial.buscar(self.busqueda_entry.get()):
            self.tabla.insert("", "end", iid=str(registro[0]), values=(registro[1], registro[2], registro[3], registro[4], registro[7] or ""))

    def _registro_seleccionado(self):
        seleccion = self.tabla.selection()
        if not seleccion:
            messagebox.showwarning("Aviso", "Selecciona un registro del historial.")
            return None
        registro_id = int(seleccion[0])
        return next((fila for fila in self.historial.buscar("") if fila[0] == registro_id), None)

    def abrir_documento(self):
        registro = self._registro_seleccionado()
        if registro:
            try:
                os.startfile(registro[5])
            except OSError:
                messagebox.showwarning("Aviso", "No se pudo abrir el documento.")

    def abrir_carpeta(self):
        registro = self._registro_seleccionado()
        if registro:
            try:
                os.startfile(str(Path(registro[5]).parent))
            except OSError:
                messagebox.showwarning("Aviso", "No se pudo abrir la carpeta.")

    def eliminar_registro(self):
        registro = self._registro_seleccionado()
        if registro and messagebox.askyesno("Confirmar", "¿Eliminar el registro? Los archivos no se borrarán."):
            self.historial.eliminar(registro[0])
            self.actualizar_historial()

    def _validar(self):
        if not self.litigio_entry.get().strip():
            messagebox.showerror("Error", "El Nº Litigio es obligatorio.")
            return False
        faltantes = [etiqueta for clave, etiqueta in PHOTO_FIELDS if not self.fotos[clave]]
        if faltantes:
            messagebox.showerror("Error", "Debes adjuntar las tres fotografías requeridas.")
            return False
        plantilla = self.configuracion.get("power_line_ruta_plantilla", "")
        if not plantilla or not Path(plantilla).is_file():
            messagebox.showerror("Error", "Configura una plantilla Word válida en Configuración.")
            return False
        return True

    def generar_documento(self):
        if not self._validar():
            return
        ruta_guardado = self.destino_entry.get().strip() if hasattr(self, "destino_entry") else self.configuracion.get("power_line_ruta_guardado", "")
        if not ruta_guardado:
            ruta_guardado = filedialog.askdirectory(title="Selecciona la carpeta de guardado")
            if not ruta_guardado:
                return
            self.configuracion["power_line_ruta_guardado"] = ruta_guardado
            if hasattr(self, "destino_entry"):
                self.destino_entry.delete(0, "end")
                self.destino_entry.insert(0, ruta_guardado)
            if hasattr(self, "guardado_entry"):
                self.guardado_entry.delete(0, "end")
                self.guardado_entry.insert(0, ruta_guardado)
            self.app.guardar_configuracion_ui_internal()
        self.configuracion["power_line_ruta_guardado"] = ruta_guardado
        if hasattr(self, "guardado_entry"):
            self.guardado_entry.delete(0, "end")
            self.guardado_entry.insert(0, ruta_guardado)
        self.app.guardar_configuracion_ui_internal()
        datos = {"litigio": self.litigio_entry.get().strip(), "pedido": self.pedido_entry.get().strip(),
                 "referencia": self.referencia_entry.get().strip(), "descripcion": self.descripcion_text.get("1.0", "end-1c").strip()}
        salida = Path(ruta_guardado) / f"Litigio {datos['litigio']}.docx"
        self.btn_generar.configure(state="disabled")
        efectos.marcar_ocupado(self.btn_generar, "Generando")
        generar_pdf = bool(self.pdf_var.get())  # las variables de Tk se leen aquí, no desde el hilo
        threading.Thread(target=self._generar_en_segundo_plano, args=(datos, salida, generar_pdf), daemon=True).start()

    def _generar_en_segundo_plano(self, datos, salida, generar_pdf=False):
        try:
            datos_documento = {clave: datos[clave] for clave in ("pedido", "referencia", "descripcion")}
            word, pdf = DocumentoPowerLine(self.configuracion["power_line_ruta_plantilla"], salida, datos_documento, self.fotos.copy(), generar_pdf).generar()
            registro_id = self.historial.guardar(datos["litigio"], datos["pedido"], datos["referencia"], word, pdf)
            self.app.after(0, lambda: self._finalizar_generacion(word, pdf, registro_id, datos["litigio"]))
        except Exception as error:
            detalle = self._detalle_error(error)
            try:
                registrar_error_reclamacion(traceback.format_exc(), detalle, datos.get("referencia"))
            finally:
                self.app.after(0, lambda mensaje=detalle: self._error_generacion(mensaje))

    @staticmethod
    def _detalle_error(error):
        mensaje = str(error).strip()
        if mensaje and mensaje.lower() != "none":
            return mensaje
        argumentos = getattr(error, "args", ())
        detalle = " | ".join(str(argumento) for argumento in argumentos if argumento not in (None, ""))
        if detalle:
            return detalle
        return f"Error interno de Word ({type(error).__name__}). Consulta el registro detallado para más información."

    def _finalizar_generacion(self, word, pdf, registro_id=None, litigio=None):
        self.btn_generar.configure(state="normal", text="2 · Generar documento")
        if registro_id is not None and litigio:
            self._ultimo_generado = (registro_id, litigio, word)
            self.btn_paso3.configure(state="normal")
        self.actualizar_historial()
        self.app.actualizar_estado("reclamación Power Line generada", word)
        if self.configuracion.get("power_line_abrir_documento", False):
            os.startfile(word)
        if self.configuracion.get("power_line_abrir_carpeta", False):
            os.startfile(str(Path(word).parent))
        mensaje = "Documento Word generado correctamente."
        if pdf:
            mensaje += " También se ha generado el PDF."
        if registro_id is not None and litigio:
            mensaje += "\n\nPaso 3: «Enviar a Power-Line» crea el ticket y le adjunta este Word."
        messagebox.showinfo("Éxito", mensaje)

    def enviar_ultimo_a_power_line(self):
        if not getattr(self, "_ultimo_generado", None):
            messagebox.showwarning("Power-Line", "Primero genera el documento (paso 2).")
            return False
        return self.enviar_a_power_line(*self._ultimo_generado)

    # ---------- paso 1: litigio en PSA Parts (litigio_psa.py) ----------

    def crear_litigio_psa(self):
        if getattr(self, "_creando_litigio", False):
            return False
        referencia = litigio_psa.normalizar_ref(self.referencia_entry.get())
        if not referencia:
            messagebox.showwarning("Litigio", "Escribe la referencia que quieres reclamar.")
            return False
        try:
            cantidad = int(self.cantidad_entry.get().strip() or "1")
            assert cantidad > 0
        except (ValueError, AssertionError):
            messagebox.showwarning("Litigio", "La cantidad debe ser un número entero mayor que 0.")
            return False
        usuario, password = self.configuracion.get("usuario", ""), self.configuracion.get("password", "")
        if not usuario or not password:
            messagebox.showwarning("Litigio", "Faltan el usuario y la contraseña del portal Citroën (pestaña Configuración).")
            return False
        if self.litigio_entry.get().strip() and not messagebox.askyesno("Litigio", f"El formulario ya tiene el litigio {self.litigio_entry.get().strip()}. ¿Crear otro para {referencia}?"):
            return False
        motivo = litigio_psa.codigo_motivo(self.motivo_option.get())
        self.configuracion["psa_motivo"] = motivo
        sin_confirmar = bool(self.obtener_configuracion()["psa_litigio_sin_confirmacion"])
        self._creando_litigio = True
        self.btn_litigio.configure(state="disabled")
        efectos.marcar_ocupado(self.btn_litigio, "Creando")
        log = lambda texto, tag=None: self.app.after(0, self.app.agregar_log, f"[LITIGIO PSA] {texto}", tag)
        threading.Thread(target=self._litigio_en_segundo_plano, args=(usuario, password, referencia, motivo, cantidad, sin_confirmar, log), daemon=True).start()
        return True

    def _confirmar_en_pantalla(self, datos):
        """Desde el hilo de trabajo: muestra la confirmación en la ventana y espera la respuesta."""
        respuesta, hecho = [], threading.Event()

        def preguntar():
            anterior = datos.get("anterior")
            aviso = (f"⚠️ Esta línea ya se reclamó desde la app el {datetime.fromisoformat(anterior['cuando']):%d/%m} (litigio {anterior['litigio']}).\n\n"
                     if anterior else "")
            respuesta.append(messagebox.askyesno(
                "Guardar litigio en PARTS",
                aviso + f"Referencia {datos['referencia']}\nFactura {datos['factura']} del {datos['fecha']:%d/%m/%Y} · línea {datos['pos_factura']}\n"
                f"Doc.ventas {datos['doc_ventas']}\nMotivo {datos['motivo']} · cantidad {datos['cantidad']}\n\n"
                "«Guardar» crea el litigio y no se puede deshacer. ¿Guardar?", icon="warning", default="no" if anterior else "yes"))
            hecho.set()

        self.app.after(0, preguntar)
        hecho.wait()
        return respuesta[0]

    def _litigio_en_segundo_plano(self, usuario, password, referencia, motivo, cantidad, sin_confirmar, log):
        resultado, error = None, None
        try:
            def confirmar(datos):
                datos["anterior"] = litigio_anterior(datos["factura"], datos["pos_factura"])
                if sin_confirmar:
                    if datos["anterior"]:
                        log(f"Esa línea ya se reclamó (litigio {datos['anterior']['litigio']}): no se guarda sin confirmación.", "error")
                        return False
                    return True
                return self._confirmar_en_pantalla(datos)

            resultado = litigio_psa.LitigioPSA(usuario, password, referencia, motivo, cantidad, confirmar=confirmar, log=log).ejecutar()
            anotar_litigio_psa(resultado, referencia)
        except litigio_psa.LitigioNoCreado as e:
            error = ("aviso", str(e))
        except Exception as e:
            error = ("error", str(e).splitlines()[0] if str(e) else type(e).__name__)
            registrar_error_reclamacion(traceback.format_exc(), f"Litigio PSA: {error[1]}", referencia)
        self.app.after(0, lambda: self._litigio_terminado(referencia, resultado, error))

    def _litigio_terminado(self, referencia, resultado, error):
        self._creando_litigio = False
        self.btn_litigio.configure(state="normal", text="1 · Crear litigio en PARTS")
        if resultado:
            for campo, valor in ((self.litigio_entry, resultado["litigio"]), (self.pedido_entry, resultado["doc_ventas"])):
                campo.delete(0, "end")
                campo.insert(0, valor)
            self.referencia_entry.delete(0, "end")
            self.referencia_entry.insert(0, referencia)
            avisos = f" Avisos: {' · '.join(resultado['avisos'])}" if resultado.get("avisos") else ""
            self.app.agregar_log(f"[LITIGIO PSA] ✅ Litigio {resultado['litigio']} creado (factura {resultado['factura']}, Doc.ventas {resultado['doc_ventas']}).{avisos}", "success")
            self.app.avisar(f"Litigio {resultado['litigio']} creado en PARTS. Añade las fotos y la descripción y genera el documento.", "ok")
        elif error and error[0] == "aviso":
            self.app.agregar_log(f"[LITIGIO PSA] {error[1]}", "error")
            messagebox.showwarning("Litigio no creado", f"No se ha creado ningún litigio:\n{error[1]}")
        elif error:
            self.app.agregar_log(f"[LITIGIO PSA] ❌ {error[1]}", "error")
            messagebox.showwarning("Litigio", f"No se pudo crear el litigio de {referencia}:\n{error[1]}\n\nComprueba en PARTS (Lista de litigios) que no se haya grabado.")

    # ---------- ticket en Power-Line (power_line_ticket.py) ----------

    def credenciales_power_line(self):
        """Usuario y contraseña de Stellantis para Power-Line; si no se han puesto, los de MADD."""
        config = self.obtener_configuracion()
        usuario = config["power_line_usuario"].strip() or self.configuracion.get("usuario_madd", "")
        password = config["power_line_password"] or (self.configuracion.get("password_madd", "") if not config["power_line_usuario"].strip() else "")
        return usuario, password

    def valores_ticket(self):
        config = self.obtener_configuracion()
        return {"cuenta": config["power_line_cuenta"].strip(), "cliente": config["power_line_cliente"].strip(),
                "tema": config["power_line_tema"].strip() or power_line_ticket.TEMA_POR_DEFECTO,
                "tipo": config["power_line_tipo"].strip() or power_line_ticket.TIPO_POR_DEFECTO}

    def enviar_seleccionado_a_power_line(self):
        registro = self._registro_seleccionado()
        if not registro:
            return False
        if registro[7] and not messagebox.askyesno("Power-Line", f"El litigio {registro[1]} ya tiene el ticket {registro[7]}. ¿Crear otro?"):
            return False
        return self.enviar_a_power_line(registro[0], registro[1], registro[5])

    def enviar_a_power_line(self, registro_id, litigio, ruta_word):
        """Ticket de seguimiento de la reclamación (con el Word adjunto)."""
        if not ruta_word or not Path(ruta_word).is_file():
            messagebox.showwarning("Power-Line", f"No se encuentra el Word del litigio {litigio}:\n{ruta_word}\n\nSin él no se puede adjuntar al ticket.")
            return False
        return self._lanzar_ticket(dict(self.valores_ticket(), litigio=litigio, ruta_word=ruta_word), f"del litigio {litigio}", registro_id=registro_id)

    def _lanzar_ticket(self, datos, que, registro_id=None, plantilla=None):
        """Comprobaciones comunes y envío en segundo plano. 'que' describe el ticket en los avisos («del litigio 131...»)."""
        if getattr(self, "_enviando_power_line", False):
            self.app.avisar("Ya se está creando un ticket en Power-Line.")
            return False
        if not datos["cuenta"] or not datos["cliente"]:
            messagebox.showwarning("Power-Line", "Falta la Cuenta Solicitante o el Cliente destinatario de Power-Line. Rellénalos en Configuración > Reclamaciones Power Line.")
            return False
        usuario, password = self.credenciales_power_line()
        if not usuario or not password:
            messagebox.showwarning("Power-Line", "Falta el usuario o la contraseña de Stellantis: ponlos en Configuración > Reclamaciones Power Line (o los de MADD en la pestaña Facturas).")
            return False
        enviar_solo = bool(self.obtener_configuracion()["power_line_envio_automatico"])
        self._enviando_power_line = True
        for nombre in ("btn_enviar_power_line", "btn_paso3", "btn_crear_ticket"):
            boton = getattr(self, nombre, None)
            if boton is not None and boton.winfo_exists():
                boton.configure(state="disabled")
                efectos.marcar_ocupado(boton, "Enviando")
        aviso = "se enviará solo" if enviar_solo else "revísalo y pulsa «Enviar» en la ventana de Power-Line"
        self.app.avisar(f"Creando el ticket {que} en Power-Line: {aviso}.")
        log = lambda texto, tag=None: self.app.after(0, self.app.agregar_log, f"[POWER-LINE] {texto}", tag)
        threading.Thread(target=self._ticket_en_segundo_plano, args=(registro_id, datos, enviar_solo, log, que, plantilla), daemon=True).start()
        return True

    def _ticket_en_segundo_plano(self, registro_id, datos, enviar_solo, log, que=None, plantilla=None):
        que = que or f"del litigio {datos.get('litigio')}"
        error, numero, error_adjunto = None, None, None
        try:
            usuario, password = self.credenciales_power_line()
            ticket = power_line_ticket.TicketPowerLine(usuario, password, datos, enviar_solo=enviar_solo, log=log)
            numero = ticket.ejecutar()
            error_adjunto = ticket.error_adjunto
            if registro_id is not None:
                self.historial.guardar_ticket(registro_id, numero)
            if plantilla:
                anotar_ticket(numero, plantilla, datos.get("criterios", {}))
        except power_line_ticket.TicketCancelado:
            error = "cancelado"
        except Exception as e:
            error = str(e).splitlines()[0] if str(e) else type(e).__name__
            registrar_error_reclamacion(traceback.format_exc(), f"Ticket Power-Line: {error}", datos.get("litigio") or plantilla)
        con_word = bool(datos.get("ruta_word"))
        self.app.after(0, lambda: self._ticket_terminado(que, numero, error, error_adjunto, con_word))

    def _ticket_terminado(self, que, numero, error, error_adjunto, con_word=True):
        self._enviando_power_line = False
        for nombre, texto in (("btn_enviar_power_line", "Enviar a Power-Line"), ("btn_paso3", "3 · Enviar a Power-Line"), ("btn_crear_ticket", "Crear ticket")):
            boton = getattr(self, nombre, None)
            if boton is not None and boton.winfo_exists():
                boton.configure(state="normal", text=texto)
        self.actualizar_historial()
        self.refrescar_tickets_recientes()
        if numero and not error_adjunto:
            con = " con el Word" if con_word else ""
            self.app.agregar_log(f"[POWER-LINE] ✅ Ticket {numero} enviado{con} ({que}).", "success")
            self.app.avisar(f"Ticket {numero} enviado a Power-Line{con} ({que}).", "ok")
        elif numero:
            self.app.agregar_log(f"[POWER-LINE] ⚠️ Ticket {numero} enviado, pero no se pudo adjuntar el Word: {error_adjunto}", "error")
            messagebox.showwarning("Power-Line", f"El ticket {numero} se ha enviado, pero no se pudo adjuntar el Word ({error_adjunto}).\n\nAdjúntalo a mano con el clip del ticket.")
        elif error == "cancelado":
            self.app.agregar_log(f"[POWER-LINE] Ticket {que} no enviado (ventana cerrada o sin pulsar «Enviar»).", "error")
            self.app.avisar(f"No se ha enviado el ticket {que}.", "aviso")
        else:
            self.app.agregar_log(f"[POWER-LINE] ❌ No se pudo crear el ticket {que}: {error}", "error")
            messagebox.showwarning("Power-Line", f"No se pudo crear el ticket {que}:\n{error}")

    # ---------- tickets con plantilla (pestaña «Tickets Power-Line») ----------

    PLANTILLAS_PESTANA = ("plazo_faltante", "motivo_cancelacion", "cancelacion_linea")

    def _construir_tickets(self):
        app = self.app
        contenedor = ctk.CTkScrollableFrame(self.tab_tickets, fg_color="transparent")
        contenedor.pack(padx=2, pady=2, fill="both", expand=True)
        card = app.crear_tarjeta_tab(contenedor, "Nuevo ticket en Power-Line", "Elige la plantilla y rellena sus datos: la cuenta, el cliente, el tema y el tipo se ponen solos.", icono="correo")
        nombres = [power_line_ticket.PLANTILLAS[c][0] for c in self.PLANTILLAS_PESTANA]
        fila = ctk.CTkFrame(card, fg_color="transparent")
        fila.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        app.crear_etiqueta(fila, "Plantilla", negrita=True, width=180, anchor="w").pack(side="left")
        self.plantilla_ticket_option = app.crear_menu(fila, nombres, comando=lambda valor: self._pintar_criterios())
        self.plantilla_ticket_option.set(nombres[0])
        self.plantilla_ticket_option.pack(side="left", fill="x", expand=True)
        self.criterios_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.criterios_frame.pack(padx=estilo.PAD_X, fill="x")
        self.criterios_entries = {}
        app.crear_etiqueta(card, "Descripción (opcional)", negrita=True).pack(anchor="w", padx=estilo.PAD_X, pady=(6, 4))
        self.descripcion_ticket_text = ctk.CTkTextbox(card, height=60, corner_radius=estilo.RADIO, border_width=1, border_color=app.color_borde,
                                                      fg_color=app.color_campo, text_color=app.color_texto, font=estilo.fuente())
        self.descripcion_ticket_text.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        self.btn_crear_ticket = app.crear_boton_principal(card, "Crear ticket", self.crear_ticket_plantilla)
        self.btn_crear_ticket.pack(padx=estilo.PAD_X, pady=(0, 14), fill="x")
        recientes = app.crear_tarjeta_tab(contenedor, "Últimos tickets creados", "Los tickets creados desde esta pestaña.", icono="reloj")
        self.tickets_recientes_box = ctk.CTkTextbox(recientes, height=150, corner_radius=estilo.RADIO, border_width=1, border_color=app.color_borde,
                                                    fg_color=app.color_campo, text_color=app.color_texto, font=estilo.fuente(estilo.TAM_PEQUENO))
        self.tickets_recientes_box.pack(padx=estilo.PAD_X, pady=(0, 14), fill="x")
        self._pintar_criterios()
        self.refrescar_tickets_recientes()

    def plantilla_elegida(self):
        nombre = self.plantilla_ticket_option.get()
        return next(c for c in self.PLANTILLAS_PESTANA if power_line_ticket.PLANTILLAS[c][0] == nombre)

    def _pintar_criterios(self):
        anteriores = {}
        for nombre, entrada in self.criterios_entries.items():
            try:
                anteriores[nombre] = entrada.get()
            except Exception:
                pass  # la pestaña se reconstruyó: esa caja ya no existe
        for hijo in self.criterios_frame.winfo_children():
            hijo.destroy()
        self.criterios_entries = {}
        for nombre in power_line_ticket.PLANTILLAS[self.plantilla_elegida()][3]:
            fila = ctk.CTkFrame(self.criterios_frame, fg_color="transparent")
            fila.pack(pady=3, fill="x")
            self.app.crear_etiqueta(fila, nombre, width=180, anchor="w").pack(side="left")
            entrada = self.app.crear_entrada(fila)
            if anteriores.get(nombre):
                entrada.insert(0, anteriores[nombre])  # al cambiar de plantilla se conservan los datos comunes
            entrada.pack(side="left", fill="x", expand=True)
            self.criterios_entries[nombre] = entrada

    def crear_ticket_plantilla(self):
        clave = self.plantilla_elegida()
        nombre, tema, tipo, criterios = power_line_ticket.PLANTILLAS[clave]
        valores = {c: self.criterios_entries[c].get().strip() for c in criterios}
        vacios = [c for c, v in valores.items() if not v]
        if vacios:
            messagebox.showwarning("Power-Line", "Rellena: " + ", ".join(vacios) + ".")
            return False
        config = self.valores_ticket()
        datos = {"cuenta": config["cuenta"], "cliente": config["cliente"], "tema": tema, "tipo": tipo, "criterios": valores,
                 "descripcion": self.descripcion_ticket_text.get("1.0", "end-1c").strip()}
        resumen = " · ".join(valores.values())
        return self._lanzar_ticket(datos, f"«{nombre}» ({resumen})", plantilla=nombre)

    def refrescar_tickets_recientes(self):
        caja = getattr(self, "tickets_recientes_box", None)
        if caja is None or not caja.winfo_exists():
            return
        lineas = [f"{datetime.fromisoformat(t['cuando']):%d/%m %H:%M}  {t['ticket']}  {t['plantilla']}: {' · '.join(t['criterios'].values())}"
                  for t in reversed(leer_tickets()[-15:])]
        caja.configure(state="normal")
        caja.delete("1.0", "end")
        caja.insert("1.0", "\n".join(lineas) or "Todavía no se ha creado ningún ticket desde aquí.")
        caja.configure(state="disabled")

    def _error_generacion(self, detalle):
        self.btn_generar.configure(state="normal", text="2 · Generar documento")
        messagebox.showerror("Error", f"No se pudo generar la reclamación.\n{detalle}")

    def construir_ajustes(self, parent):
        app = self.app
        marco = ctk.CTkFrame(parent, fg_color=app.color_superficie_alt, corner_radius=10)
        marco.pack(padx=estilo.PAD_X, pady=(0, estilo.PAD_Y), fill="x")
        app.crear_etiqueta(marco, "Reclamaciones Power Line", negrita=True).pack(anchor="w", padx=estilo.PAD_FILA, pady=(estilo.PAD_FILA, 4))
        self.plantilla_entry = self._fila_ajuste(marco, "Ruta de la plantilla Word", "power_line_ruta_plantilla", self._seleccionar_plantilla)
        self.guardado_entry = self._fila_ajuste(marco, "Ruta por defecto de guardado", "power_line_ruta_guardado", self.cambiar_carpeta)
        self.abrir_documento_var = ctk.BooleanVar(value=bool(self.configuracion.get("power_line_abrir_documento", False)))
        self.abrir_carpeta_var = ctk.BooleanVar(value=bool(self.configuracion.get("power_line_abrir_carpeta", False)))
        ctk.CTkCheckBox(marco, text="Abrir documento al finalizar", variable=self.abrir_documento_var, fg_color=app.color_principal, hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_FILA, pady=4)
        ctk.CTkCheckBox(marco, text="Abrir carpeta al finalizar", variable=self.abrir_carpeta_var, fg_color=app.color_principal, hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_FILA, pady=4)
        self.litigio_sin_confirmacion_var = ctk.BooleanVar(value=bool(self.configuracion.get("psa_litigio_sin_confirmacion", False)))
        ctk.CTkCheckBox(marco, text="Guardar el litigio en PARTS sin pedir confirmación (nunca si esa línea ya se reclamó)", variable=self.litigio_sin_confirmacion_var,
                        fg_color=app.color_principal, hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_FILA, pady=4)
        app.crear_etiqueta(marco, "Ticket en Power-Line (se entra con el usuario de Stellantis de MADD)", negrita=True).pack(anchor="w", padx=estilo.PAD_FILA, pady=(10, 2))
        self.usuario_power_line_entry = self._fila_texto(marco, "Usuario Stellantis", "power_line_usuario", "", "Vacío = el mismo de MADD")
        self.password_power_line_entry = self._fila_texto(marco, "Contraseña", "power_line_password", "", "Vacío = la misma de MADD", ocultar=True)
        self.cuenta_power_line_entry = self._fila_texto(marco, "Cuenta Solicitante", "power_line_cuenta", "", "Código - nombre, tal como sale en Power-Line")
        self.cliente_power_line_entry = self._fila_texto(marco, "Cliente destinatario", "power_line_cliente", "", "Código - nombre, tal como sale en Power-Line")
        self.tema_power_line_entry = self._fila_texto(marco, "Tema del ticket", "power_line_tema", power_line_ticket.TEMA_POR_DEFECTO)
        self.tipo_power_line_entry = self._fila_texto(marco, "Tipo de ticket", "power_line_tipo", power_line_ticket.TIPO_POR_DEFECTO)
        self.envio_automatico_var = ctk.BooleanVar(value=bool(self.configuracion.get("power_line_envio_automatico", False)))
        ctk.CTkCheckBox(marco, text="Enviar el ticket sin pedir confirmación (la app pulsa «Enviar» y adjunta el Word sola)", variable=self.envio_automatico_var,
                        fg_color=app.color_principal, hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_FILA, pady=(4, estilo.PAD_FILA))

    def _fila_texto(self, parent, texto, clave, defecto, ejemplo="", ocultar=False):
        fila = ctk.CTkFrame(parent, fg_color="transparent")
        fila.pack(padx=estilo.PAD_FILA, pady=3, fill="x")
        self.app.crear_etiqueta(fila, texto, width=180, anchor="w").pack(side="left")
        opciones = {"placeholder_text": ejemplo} if ejemplo else {}
        if ocultar:
            opciones["show"] = "•"
        entrada = self.app.crear_entrada(fila, **opciones)
        valor = self.configuracion.get(clave) or defecto
        if valor:
            entrada.insert(0, valor)
        entrada.pack(side="left", fill="x", expand=True, padx=6)
        return entrada

    def _fila_ajuste(self, parent, texto, clave, comando):
        fila = ctk.CTkFrame(parent, fg_color="transparent")
        fila.pack(padx=estilo.PAD_FILA, pady=5, fill="x")
        self.app.crear_etiqueta(fila, texto, width=180, anchor="w").pack(side="left")
        entrada = self.app.crear_entrada(fila)
        entrada.insert(0, self.configuracion.get(clave, ""))
        entrada.pack(side="left", fill="x", expand=True, padx=6)
        self.app.crear_boton_secundario(fila, "Examinar", comando, width=80).pack(side="right")
        return entrada

    def _seleccionar_plantilla(self):
        ruta = filedialog.askopenfilename(filetypes=[("Plantilla Word", "*.docx *.dotx")])
        if ruta:
            self.plantilla_entry.delete(0, "end")
            self.plantilla_entry.insert(0, ruta)

    def obtener_configuracion(self):
        return {
            "power_line_ruta_plantilla": self.plantilla_entry.get() if hasattr(self, "plantilla_entry") else self.configuracion.get("power_line_ruta_plantilla", ""),
            "power_line_ruta_guardado": self.guardado_entry.get() if hasattr(self, "guardado_entry") else self.configuracion.get("power_line_ruta_guardado", ""),
            "power_line_abrir_documento": self.abrir_documento_var.get() if hasattr(self, "abrir_documento_var") else bool(self.configuracion.get("power_line_abrir_documento", False)),
            "power_line_abrir_carpeta": self.abrir_carpeta_var.get() if hasattr(self, "abrir_carpeta_var") else bool(self.configuracion.get("power_line_abrir_carpeta", False)),
            "power_line_usuario": self._valor("usuario_power_line_entry", "power_line_usuario", ""),
            "power_line_password": self._valor("password_power_line_entry", "power_line_password", ""),
            "power_line_cuenta": self._valor("cuenta_power_line_entry", "power_line_cuenta", ""),
            "power_line_cliente": self._valor("cliente_power_line_entry", "power_line_cliente", ""),
            "power_line_tema": self._valor("tema_power_line_entry", "power_line_tema", power_line_ticket.TEMA_POR_DEFECTO),
            "power_line_tipo": self._valor("tipo_power_line_entry", "power_line_tipo", power_line_ticket.TIPO_POR_DEFECTO),
            "power_line_envio_automatico": self.envio_automatico_var.get() if hasattr(self, "envio_automatico_var") else bool(self.configuracion.get("power_line_envio_automatico", False)),
            "psa_litigio_sin_confirmacion": self.litigio_sin_confirmacion_var.get() if hasattr(self, "litigio_sin_confirmacion_var") else bool(self.configuracion.get("psa_litigio_sin_confirmacion", False)),
            "psa_motivo": litigio_psa.codigo_motivo(self.motivo_option.get()) if hasattr(self, "motivo_option") else self.configuracion.get("psa_motivo", litigio_psa.MOTIVO_POR_DEFECTO),
        }

    def _valor(self, atributo, clave, defecto):
        entrada = getattr(self, atributo, None)
        if entrada is not None:
            try:
                return entrada.get()
            except Exception:
                pass
        return self.configuracion.get(clave) or defecto

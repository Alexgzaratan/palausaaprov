"""Módulo aislado para crear y consultar reclamaciones de Power Line."""

import os
import shutil
import sqlite3
import tempfile
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from PIL import Image
from tkinterdnd2 import DND_FILES, TkinterDnD

import efectos
import estilo
from rutas_app import ruta_datos


DATABASE_FILE = ruta_datos("reclamaciones_power_line.db")
ERROR_LOG_FILE = ruta_datos("error_log_detallado.txt", migrar=False)
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


def registrar_error_reclamacion(error, detalle, referencia):
    with open(ERROR_LOG_FILE, "a", encoding="utf-8") as archivo:
        archivo.write(
            f"--- ERROR RECLAMACIONES POWER LINE [{datetime.now():%Y-%m-%d %H:%M:%S}] ---\n"
            f"Referencia: {referencia or '-'}\nMensaje: {detalle}\nTraceback:\n{error}\n{'-' * 30}\n\n"
        )


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
        finally:
            conexion.close()

    def _conexion(self):
        return sqlite3.connect(self.database_path)

    def guardar(self, litigio, pedido_sap, referencia, ruta_word, ruta_pdf):
        conexion = self._conexion()
        try:
            with conexion:
                conexion.execute(
                    "INSERT INTO reclamaciones_power_line "
                    "(litigio, fecha, pedido_sap, referencia, ruta_word, ruta_pdf) VALUES (?, ?, ?, ?, ?, ?)",
                    (litigio, datetime.now().strftime("%d/%m/%Y %H:%M"), pedido_sap, referencia, ruta_word, ruta_pdf),
                )
        finally:
            conexion.close()

    def buscar(self, texto=""):
        filtro = f"%{texto.strip()}%"
        conexion = self._conexion()
        try:
            return conexion.execute(
                "SELECT id, litigio, fecha, pedido_sap, referencia, ruta_word, ruta_pdf "
                "FROM reclamaciones_power_line WHERE litigio LIKE ? OR referencia LIKE ? OR pedido_sap LIKE ? "
                "ORDER BY id DESC",
                (filtro, filtro, filtro),
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
        with Image.open(origen) as imagen:
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
        self.carpeta_adjuntos = Path(tempfile.mkdtemp(prefix="power_line_adjuntos_"))
        self.dnd_activo = False

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
        self.app.resaltar_pestana_activa(panel)
        self._construir_formulario()
        self._construir_historial()

    def _construir_formulario(self):
        app = self.app
        contenedor = ctk.CTkScrollableFrame(self.tab_formulario, fg_color="transparent")
        contenedor.pack(padx=2, pady=2, fill="both", expand=True)
        card = app.crear_tarjeta_tab(contenedor, "Nueva reclamación", "Los campos marcados con * son obligatorios.", icono="aviso")
        self.litigio_entry = self._campo(card, "Nº Litigio *")
        self.pedido_entry = self._campo(card, "Nº Pedido SAP")
        self.referencia_entry = self._campo(card, "Referencia")
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
        self.btn_generar = app.crear_boton_principal(acciones, "Generar documento", self.generar_documento)
        self.btn_generar.pack(fill="x")
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
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("Todos los archivos", "*.*")])
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
        with Image.open(ruta) as imagen:
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
        self.busqueda_entry = app.crear_entrada(fila, placeholder_text="Nº litigio, referencia o pedido SAP...")
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
        columnas = ("litigio", "fecha", "pedido", "referencia")
        self.tabla = ttk.Treeview(tabla_frame, columns=columnas, show="headings", height=7, style="PowerLine.Treeview")
        for clave, texto, ancho in (("litigio", "Nº Litigio", 115), ("fecha", "Fecha", 120), ("pedido", "Pedido SAP", 120), ("referencia", "Referencia", 150)):
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
        ctk.CTkButton(acciones, text="Eliminar registro", height=estilo.ALTO_CONTROL, corner_radius=estilo.RADIO, fg_color=app.color_peligro, hover_color=app.color_peligro_hover, command=self.eliminar_registro).pack(side="right")
        self.actualizar_historial()

    def actualizar_historial(self):
        if not hasattr(self, "tabla"):
            return
        for fila in self.tabla.get_children():
            self.tabla.delete(fila)
        for registro in self.historial.buscar(self.busqueda_entry.get()):
            self.tabla.insert("", "end", iid=str(registro[0]), values=(registro[1], registro[2], registro[3], registro[4]))

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
        threading.Thread(target=self._generar_en_segundo_plano, args=(datos, salida), daemon=True).start()

    def _generar_en_segundo_plano(self, datos, salida):
        try:
            datos_documento = {clave: datos[clave] for clave in ("pedido", "referencia", "descripcion")}
            word, pdf = DocumentoPowerLine(self.configuracion["power_line_ruta_plantilla"], salida, datos_documento, self.fotos.copy(), self.pdf_var.get()).generar()
            self.historial.guardar(datos["litigio"], datos["pedido"], datos["referencia"], word, pdf)
            self.app.after(0, lambda: self._finalizar_generacion(word, pdf))
        except Exception as error:
            detalle = self._detalle_error(error)
            registrar_error_reclamacion(traceback.format_exc(), detalle, datos.get("referencia"))
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

    def _finalizar_generacion(self, word, pdf):
        self.btn_generar.configure(state="normal", text="Generar documento")
        self.actualizar_historial()
        self.app.actualizar_estado("reclamación Power Line generada", word)
        if self.configuracion.get("power_line_abrir_documento", False):
            os.startfile(word)
        if self.configuracion.get("power_line_abrir_carpeta", False):
            os.startfile(str(Path(word).parent))
        mensaje = "Documento Word generado correctamente."
        if pdf:
            mensaje += " También se ha generado el PDF."
        messagebox.showinfo("Éxito", mensaje)

    def _error_generacion(self, detalle):
        self.btn_generar.configure(state="normal", text="Generar documento")
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
        ctk.CTkCheckBox(marco, text="Abrir carpeta al finalizar", variable=self.abrir_carpeta_var, fg_color=app.color_principal, hover_color=app.color_hover, text_color=app.color_texto, font=estilo.fuente()).pack(anchor="w", padx=estilo.PAD_FILA, pady=(4, estilo.PAD_FILA))

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
        }

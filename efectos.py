"""Efectos de la interfaz: avisos que no interrumpen, botones con indicador de actividad, tarjetas que
responden al ratón y arrastrar archivos a los campos. Todo con customtkinter y temporizadores after()."""
import time
import tkinter

import customtkinter as ctk

import estilo
from animacion_inicio import mezclar

GIRO = "◐◓◑◒"


def escala(widget):
    try:
        return ctk.ScalingTracker.get_widget_scaling(widget)
    except Exception:
        return 1.0


def puntero_dentro(widget):
    x, y = widget.winfo_pointerxy()
    izquierda, arriba = widget.winfo_rootx(), widget.winfo_rooty()
    return izquierda <= x < izquierda + widget.winfo_width() and arriba <= y < arriba + widget.winfo_height()


def _suave(x):
    return 1 - (1 - x) ** 3


# ---------- avisos ----------

class GestorAvisos:
    """Avisos que entran deslizando abajo a la derecha (encima del panel de actividad) y se van solos."""
    MARGEN = 20
    SEPARACION = 8
    DURACION_MS = 4000
    MAXIMO = 4
    FUERA = 420  # desplazamiento a la derecha para entrar y salir de la ventana

    def __init__(self, app):
        self.app = app
        self.activos = []
        self.historial = []  # (tipo, texto) de todos los avisos mostrados

    def color(self, tipo):
        app = self.app
        return {"ok": app.color_ok, "aviso": app.color_texto_error}.get(tipo, app.color_principal)

    def mostrar(self, texto, tipo="info", duracion=None, accion=None):
        """accion: (texto del botón, función), p. ej. ("Abrir", abrir_resultado); con acción dura más."""
        self.historial.append((tipo, texto))
        app = self.app
        self.activos = [m for m in self.activos if m.winfo_exists() and not getattr(m, "_cerrando", False)]
        while len(self.activos) >= self.MAXIMO:
            self.cerrar(self.activos[0])
            self.activos = self.activos[1:]
        color = self.color(tipo)
        marco = ctk.CTkFrame(app, fg_color=app.color_superficie, corner_radius=12, border_width=1, border_color=mezclar(app.color_borde, color, 0.55))
        ctk.CTkFrame(marco, fg_color=color, width=4, height=24, corner_radius=2).pack(side="left", fill="y", padx=(10, 0), pady=10)
        if estilo.FUENTE_ICONOS:
            icono = {"ok": "check", "aviso": "aviso"}.get(tipo, "info")
            ctk.CTkLabel(marco, text=estilo.ICONOS[icono], font=estilo.fuente_icono(16), text_color=color).pack(side="left", padx=(10, 0))
        boton = None
        if accion:
            texto_boton, funcion = accion

            def ejecutar(m=marco, f=funcion):
                self.cerrar(m)
                f()

            boton = ctk.CTkButton(marco, text=texto_boton, command=ejecutar, width=64, height=28, corner_radius=8, font=estilo.fuente(estilo.TAM_PEQUENO, True),
                                  fg_color=app.color_principal, hover_color=app.color_hover, text_color="#ffffff")  # el del tema: contraste AA con blanco
            boton.pack(side="right", padx=(0, 12), pady=10)
            duracion = duracion or 7000
        ctk.CTkLabel(marco, text=texto, font=estilo.fuente(), text_color=app.color_texto, wraplength=300, justify="left", anchor="w").pack(side="left", padx=(10, 12 if boton else 16), pady=10)
        for widget in [marco] + marco.winfo_children():
            if widget is not boton:
                widget.bind("<Button-1>", lambda evento, m=marco: self.cerrar(m))
        marco.boton_accion = boton
        marco.texto = texto
        marco._desplazamiento = self.FUERA
        marco._y = 0
        self.activos.append(marco)
        self.recolocar()
        marco.lift()
        self._deslizar(marco, 0)
        app.after(duracion or self.DURACION_MS, lambda: self.cerrar(marco))
        return marco

    def _altura_base(self):
        panel = getattr(self.app, "panel_actividad", None)
        try:
            if panel is not None and panel.winfo_exists() and panel.winfo_manager():
                return panel.winfo_height() / escala(self.app) + 12 + 10
        except Exception:
            pass
        return self.MARGEN

    def _colocar(self, marco):
        marco.place(relx=1.0, rely=1.0, anchor="se", x=-self.MARGEN + marco._desplazamiento, y=-marco._y)

    def recolocar(self):
        """El más reciente abajo; los anteriores se apilan encima."""
        self.app.update_idletasks()
        y = self._altura_base()
        for marco in reversed(self.activos):
            if not marco.winfo_exists():
                continue
            marco._y = y
            self._colocar(marco)
            y += marco.winfo_reqheight() / escala(self.app) + self.SEPARACION

    def _deslizar(self, marco, destino, al_terminar=None, pasos=9, paso=0, origen=None):
        if not marco.winfo_exists():
            return
        origen = marco._desplazamiento if origen is None else origen
        paso += 1
        marco._desplazamiento = origen + (destino - origen) * _suave(paso / pasos)
        self._colocar(marco)
        if paso < pasos:
            marco.after(18, lambda: self._deslizar(marco, destino, al_terminar, pasos, paso, origen))
        elif al_terminar:
            al_terminar()

    def cerrar(self, marco):
        if not marco.winfo_exists() or getattr(marco, "_cerrando", False):
            return
        marco._cerrando = True

        def quitar():
            if marco in self.activos:
                self.activos.remove(marco)
            marco.destroy()
            self.recolocar()

        self._deslizar(marco, self.FUERA, quitar, pasos=7)


# ---------- botones ocupados ----------

def marcar_ocupado(boton, texto, intervalo_ms=150):
    """Muestra un círculo girando y el tiempo transcurrido ("◐ Analizando · 0:42") hasta que otro código
    cambia el texto del botón (al terminar la tarea), momento en que se para solo."""
    inicio = time.monotonic()
    estado = {"texto": None, "vuelta": 0}

    def paso():
        try:
            if not boton.winfo_exists():
                return
            if estado["texto"] is not None and boton.cget("text") != estado["texto"]:
                return
            segundos = int(time.monotonic() - inicio)
            nuevo = f"{GIRO[estado['vuelta'] % len(GIRO)]}  {texto} · {segundos // 60}:{segundos % 60:02d}"
            estado["vuelta"] += 1
            estado["texto"] = nuevo
            boton.configure(text=nuevo)
            boton.after(intervalo_ms, paso)
        except Exception:
            pass

    paso()


# ---------- tarjetas que responden al ratón ----------

NIVEL_RESALTE = 0.7  # mezcla entre el borde neutro (0) y el color del tema (1)


def activar_resalte(app, tarjeta, pasos=6, intervalo_ms=20, vigilancia_ms=200):
    """Al entrar el ratón, el borde se funde hacia el color del tema; vuelve al salir de la tarjeta entera
    (se vigila la posición del puntero para no parpadear al pasar por encima de sus controles)."""
    estado = {"dentro": False, "nivel": 0.0, "animacion": None}

    def pintar():
        tarjeta.configure(border_color=mezclar(app.color_borde, app.color_principal, estado["nivel"]))

    def animar(objetivo):
        if estado["animacion"] is not None:
            try:
                tarjeta.after_cancel(estado["animacion"])
            except Exception:
                pass
        incremento = NIVEL_RESALTE / pasos

        def paso():
            estado["animacion"] = None
            if not tarjeta.winfo_exists():
                return
            diferencia = objetivo - estado["nivel"]
            if abs(diferencia) <= incremento:
                estado["nivel"] = objetivo
                pintar()
                return
            estado["nivel"] += incremento if diferencia > 0 else -incremento
            pintar()
            estado["animacion"] = tarjeta.after(intervalo_ms, paso)

        paso()

    def vigilar():
        if not tarjeta.winfo_exists() or not estado["dentro"]:
            return
        if puntero_dentro(tarjeta):
            tarjeta.after(vigilancia_ms, vigilar)
        else:
            estado["dentro"] = False
            animar(0.0)

    def entrar(evento=None):
        if estado["dentro"]:
            return
        estado["dentro"] = True
        animar(NIVEL_RESALTE)
        tarjeta.after(vigilancia_ms, vigilar)

    tarjeta.bind("<Enter>", entrar, add="+")
    tarjeta._resalte = estado
    tarjeta._entrar_resalte = entrar
    return estado


# ---------- pistas al pasar el ratón ----------

def poner_pista(app, widget, texto, espera_ms=450):
    """Rótulo pequeño bajo el control al dejar el ratón encima. `texto` puede ser una función, para que
    diga siempre lo actual (p. ej. el nombre del archivo que se va a abrir)."""
    estado = {"ventana": None, "espera": None}

    def mostrar():
        estado["espera"] = None
        if estado["ventana"] is not None or not widget.winfo_exists():
            return
        contenido = texto() if callable(texto) else texto
        if not contenido:
            return
        ventana = tkinter.Toplevel(widget)
        ventana.overrideredirect(True)
        ventana.attributes("-topmost", True)
        # Tamaño negativo = píxeles, como los textos de customtkinter (en positivo Tk usa puntos y sale más grande)
        fuente = (estilo.FUENTE, -round(estilo.TAM_PEQUENO * escala(widget)))
        tkinter.Label(ventana, text=contenido, bg=app.color_texto, fg=app.color_superficie, font=fuente,
                      padx=round(8 * escala(widget)), pady=round(4 * escala(widget))).pack()
        ventana.update_idletasks()
        x = widget.winfo_rootx() + (widget.winfo_width() - ventana.winfo_width()) // 2
        x = max(0, min(x, widget.winfo_screenwidth() - ventana.winfo_width()))
        ventana.geometry(f"+{x}+{widget.winfo_rooty() + widget.winfo_height() + 6}")
        estado["ventana"] = ventana

    def ocultar(evento=None):
        if estado["espera"] is not None:
            try:
                widget.after_cancel(estado["espera"])
            except Exception:
                pass
            estado["espera"] = None
        if estado["ventana"] is not None:
            try:
                estado["ventana"].destroy()
            except Exception:
                pass
            estado["ventana"] = None

    def entrar(evento=None):
        if estado["ventana"] is None and estado["espera"] is None:
            estado["espera"] = widget.after(espera_ms, mostrar)

    def salir(evento=None):
        # Entre el fondo y el texto de un botón también hay «salidas»: solo se oculta si el ratón ya no está encima
        widget.after(60, lambda: None if (widget.winfo_exists() and puntero_dentro(widget)) else ocultar())

    widget.bind("<Enter>", entrar, add="+")
    widget.bind("<Leave>", salir, add="+")
    widget.bind("<Button-1>", ocultar, add="+")
    widget._pista = estado
    widget._mostrar_pista, widget._ocultar_pista = mostrar, ocultar
    return estado


# ---------- arrastrar archivos a los campos ----------

def preparar_arrastre(ventana):
    """Carga tkdnd en la ventana; False si no está disponible (se sigue pudiendo usar «Examinar»)."""
    try:
        from tkinterdnd2 import TkinterDnD
        TkinterDnD.require(ventana)
        return True
    except Exception:
        return False


def activar_soltar(app, entrada, al_soltar):
    """El campo acepta archivos soltados desde el explorador; su borde se ilumina mientras están encima."""
    from tkinterdnd2 import COPY, DND_FILES

    def restaurar():
        con_foco = app.focus_get() is entrada._entry
        entrada.configure(border_color=app.color_principal if con_foco else app.color_borde_campo, border_width=2 if con_foco else 1)

    def encima(evento):
        entrada.configure(border_color=app.color_principal, border_width=2)
        return COPY

    def fuera(evento):
        restaurar()
        return COPY

    def soltar(evento):
        restaurar()
        al_soltar(evento.data)
        return COPY

    for widget in (entrada._entry, entrada._canvas):
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<DropEnter>>", encima)
        widget.dnd_bind("<<DropPosition>>", lambda evento: COPY)
        widget.dnd_bind("<<DropLeave>>", fuera)
        widget.dnd_bind("<<Drop>>", soltar)

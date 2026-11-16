"""Efectos de la interfaz: avisos que no interrumpen, botones con indicador de actividad, tarjetas que
responden al ratón y arrastrar archivos a los campos. Todo con customtkinter y temporizadores after()."""
import math
import time
from datetime import datetime
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

def _color(widget, nombre):
    try:
        return widget._apply_appearance_mode(widget.cget(nombre))
    except Exception:
        return None


def marcar_ocupado(boton, texto, intervalo_ms=150):
    """Muestra un círculo girando y el tiempo transcurrido ("◐ Analizando · 0:42") hasta que otro código
    cambia el texto del botón (al terminar la tarea), momento en que se para solo.
    Devuelve poner(detalle, fraccion): el paso en curso sustituye al texto ("◐ Fila 12.000 de 40.000 · 0:05") y aparece
    una barrita fina dentro del botón, con el avance (0 a 1) o en movimiento si fraccion es None. Se puede llamar
    desde otro hilo: solo guarda los valores y el temporizador del botón los pinta."""
    inicio = time.monotonic()
    estado = {"texto": None, "vuelta": 0, "detalle": None, "fraccion": None, "barra": None}

    def poner(detalle=None, fraccion=None):
        estado["detalle"], estado["fraccion"] = detalle, fraccion

    def quitar_barra():
        barra = estado["barra"]
        estado["barra"] = None
        if barra is not None:
            try:
                barra.stop()
                barra.destroy()
            except Exception:
                pass

    def pintar_barra():
        barra = estado["barra"]
        if barra is None:
            fondo, letra = _color(boton, "fg_color"), _color(boton, "text_color")
            pista = mezclar(fondo, letra, 0.3) if fondo and letra and fondo.startswith("#") and letra.startswith("#") else None
            barra = ctk.CTkProgressBar(boton, height=3, corner_radius=2, border_width=0, progress_color=letra or "#ffffff",
                                       fg_color=pista or "transparent", mode="determinate")
            barra.place(relx=0.5, rely=1.0, relwidth=0.8, anchor="s", y=-4)
            estado["barra"] = barra
            estado["modo"] = None
        fraccion = estado["fraccion"]
        modo = "indeterminate" if fraccion is None else "determinate"
        if estado.get("modo") != modo:
            barra.stop()
            barra.configure(mode=modo)
            if modo == "indeterminate":
                barra.start()
            estado["modo"] = modo
        if fraccion is not None:
            barra.set(max(0.0, min(1.0, fraccion)))

    def paso():
        try:
            if not boton.winfo_exists():
                return
            if estado["texto"] is not None and boton.cget("text") != estado["texto"]:
                quitar_barra()
                return
            segundos = int(time.monotonic() - inicio)
            nuevo = f"{GIRO[estado['vuelta'] % len(GIRO)]}  {estado['detalle'] or texto} · {segundos // 60}:{segundos % 60:02d}"
            estado["vuelta"] += 1
            estado["texto"] = nuevo
            boton.configure(text=nuevo)
            if estado["detalle"] is not None:
                pintar_barra()
            boton.after(intervalo_ms, paso)
        except Exception:
            pass

    paso()
    poner.estado = estado
    return poner


# ---------- campo que falta ----------

SACUDIDA = (6, -6, 4, -4, 2, -2, 0)  # desplazamientos en píxeles, uno cada SACUDIDA_MS
SACUDIDA_MS = 35
ROJO_MS = 2500  # el borde sigue en rojo este tiempo (o hasta que se escribe en el campo)


def sacudir(app, entrada):
    """El campo tiembla un instante con el borde en rojo y recibe el foco: «aquí falta algo». Solo se mueve si está
    colocado con pack (cambia el margen izquierdo y el derecho a la vez, sin mover a sus vecinos)."""
    if not entrada.winfo_exists():
        return False
    entrada.configure(border_color=app.color_texto_error, border_width=2)
    entrada.focus_set()
    margen = None
    if entrada.winfo_manager() == "pack":
        padx = entrada.pack_info().get("padx", 0)
        partes = [int(float(p)) for p in (padx if isinstance(padx, (tuple, list)) else str(padx).split())] or [0]
        margen = (partes[0], partes[-1])
    if getattr(entrada, "_sacudiendo", False):
        return True
    entrada._sacudiendo = True

    def mover(i=0):
        if not entrada.winfo_exists():
            return
        if margen is not None:
            izquierda, derecha = margen
            d = max(-izquierda, min(derecha, SACUDIDA[i])) if i < len(SACUDIDA) else 0
            entrada.pack_configure(padx=(izquierda + d, derecha - d))
        if i < len(SACUDIDA):
            entrada.after(SACUDIDA_MS, lambda: mover(i + 1))
        else:
            entrada._sacudiendo = False

    def quitar_rojo():
        if entrada.winfo_exists() and app.focus_get() is not getattr(entrada, "_entry", None):
            entrada.configure(border_color=app.color_borde_campo, border_width=1)
        elif entrada.winfo_exists():
            entrada.configure(border_color=app.color_principal, border_width=2)  # con el foco, como los demás campos

    mover()
    entrada.after(ROJO_MS, quitar_rojo)
    return True


# ---------- animación de envío ----------

ENVIO_DURACION = 1.1   # segundos de animación (avión + check)
ENVIO_VISIBLE_MS = 2600  # tiempo total en pantalla antes de irse sola
INTERVALO_ENVIO_MS = 16


def _rebote(x):
    c1 = 1.70158
    return 1 + (c1 + 1) * (x - 1) ** 3 + c1 * (x - 1) ** 2


def _fase(t, inicio, fin):
    return 0.0 if t <= inicio else 1.0 if t >= fin else (t - inicio) / (fin - inicio)


def _rectangulo_redondeado(lienzo, x1, y1, x2, y2, r, **opciones):
    puntos = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return lienzo.create_polygon(puntos, smooth=True, **opciones)


def animar_envio(app, titulo, subtitulo=""):
    """Tarjeta en el centro de la ventana: un avión de papel cruza dejando estela y aparece un check que se dibuja.
    Todo en un lienzo (sin bordes cortados). Se va sola o con un clic. Devuelve el lienzo, o None con el
    interruptor de animaciones apagado."""
    if not getattr(app, "config", {}).get("animacion_inicio", True):
        return None
    s = escala(app)
    ancho, alto = round(280 * s), round(196 * s)
    fondo, color, ok = app.color_superficie, app.color_principal, app.color_ok
    debajo = getattr(app, "color_panel_claro", fondo)
    lienzo = tkinter.Canvas(app, width=ancho, height=alto, bg=debajo, highlightthickness=0, bd=0, cursor="hand2")
    lienzo.place(relx=0.5, rely=0.45, anchor="center")
    tkinter.Misc.lift(lienzo)  # Canvas.lift sube elementos del lienzo, no el widget
    familia = estilo.fuente()[0]
    inicio = time.monotonic()
    zona_alto = 120 * s  # parte de arriba: avión y check
    cx, cy, radio = ancho / 2, 14 * s + zona_alto / 2, 34 * s

    def cerrar(evento=None):
        if lienzo.winfo_exists():
            lienzo.destroy()

    def avion(x, y, angulo, tam):
        c, si = math.cos(angulo), math.sin(angulo)
        puntos = [(1.0, 0.0), (-0.8, -0.65), (-0.45, 0.0), (-0.8, 0.65)]
        return [v for px, py in puntos for v in (x + (px * c - py * si) * tam, y + (px * si + py * c) * tam)]

    def punto(p):
        izquierda, arriba = 30 * s, 14 * s
        w = ancho - 60 * s
        return izquierda + p * w, arriba + zona_alto * 0.8 - math.sin(p * math.pi * 0.9) * zona_alto * 0.55 - p * zona_alto * 0.15

    def fotograma():
        if not lienzo.winfo_exists():
            return
        t = time.monotonic() - inicio
        lienzo.delete("all")
        _rectangulo_redondeado(lienzo, 1, 1, ancho - 2, alto - 2, 18 * s, fill=fondo, outline=mezclar(app.color_borde, ok, 0.6), width=1)
        lienzo.create_text(cx, alto - 52 * s, text=titulo, fill=app.color_texto, font=(familia, -round((estilo.TAM_PEQUENO + 4) * s), "bold"))
        if subtitulo:
            lienzo.create_text(cx, alto - 28 * s, text=subtitulo, fill=app.color_texto_sec, font=(familia, -round((estilo.TAM_TEXTO + 1) * s)))
        # 1) el avión cruza en arco de abajo-izquierda a arriba-derecha, con estela punteada
        vuelo = _suave(_fase(t, 0.0, 0.65))
        if vuelo < 1:
            estela = [v for i in range(13) for v in punto(vuelo * i / 12)]
            lienzo.create_line(*estela, fill=mezclar(fondo, color, 0.45), width=max(1, round(2 * s)), dash=(3, 4), smooth=True)
            x, y = punto(vuelo)
            x2, y2 = punto(min(1.0, vuelo + 0.02))
            lienzo.create_polygon(*avion(x, y, math.atan2(y2 - y, x2 - x), 13 * s), fill=color, outline="")
        # 2) círculo que rebota y check que se dibuja
        aparece = _fase(t, 0.5, 0.85)
        if aparece > 0:
            r = round(radio * max(0.0, _rebote(aparece)))
            lienzo.create_oval(round(cx) - r, round(cy) - r, round(cx) + r, round(cy) + r, fill=mezclar(fondo, ok, 0.18), outline=ok, width=max(2, round(2.5 * s)))
            trazo = _fase(t, 0.7, ENVIO_DURACION)
            if trazo > 0:
                a, b, c = (cx - 14 * s, cy + 1 * s), (cx - 4 * s, cy + 11 * s), (cx + 16 * s, cy - 11 * s)
                p = _suave(trazo) * 2
                fin1 = min(p, 1.0)
                puntos = [a, (a[0] + (b[0] - a[0]) * fin1, a[1] + (b[1] - a[1]) * fin1)]
                if p > 1:
                    puntos.append((b[0] + (c[0] - b[0]) * (p - 1), b[1] + (c[1] - b[1]) * (p - 1)))
                lienzo.create_line(*[v for xy in puntos for v in xy], fill=ok, width=round(4 * s), capstyle="round", joinstyle="round")
        if t < ENVIO_DURACION + 0.05:
            lienzo.after(INTERVALO_ENVIO_MS, fotograma)

    lienzo.bind("<Button-1>", cerrar)
    fotograma()
    lienzo.after(ENVIO_VISIBLE_MS, cerrar)
    return lienzo


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


# ---------- cabecera con degradado según la hora ----------

DEGRADADOS = {  # tres colores de izquierda a derecha; el último es además el fondo liso de la zona de fichas
    "amanecer": ("#f6a15c", "#e96b8a", "#7a6ad8"),
    "mediodia": ("#2f7fe6", "#2f9fd8", "#23958f"),
    "tarde": ("#e8793a", "#c8456e", "#5b3fa8"),
    "noche": ("#141c3a", "#27306b", "#4a3f8f"),
}
CABECERA_ALTO = 74
CABECERA_REVISION_MS = 5 * 60 * 1000  # cada cuánto se mira si ha cambiado el momento del día


def momento_del_dia(ahora=None):
    hora = (ahora or datetime.now()).hour
    return "amanecer" if 6 <= hora < 10 else "mediodia" if 10 <= hora < 16 else "tarde" if 16 <= hora < 21 else "noche"


def _rgb(color):
    return tuple(int(color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))


def imagen_degradado(ancho, alto, colores, corte):
    """Degradado horizontal de los tres colores hasta el píxel 'corte' (liso con el tercero desde ahí), con una luz
    suave arriba a la izquierda. Imagen PIL en RGB."""
    from PIL import Image, ImageChops
    c1, c2, c3 = (_rgb(c) for c in colores)
    corte = max(1, corte)
    fila = []
    for x in range(ancho):
        p = x / corte
        a, b, t = (c1, c2, p * 2) if p < 0.5 else (c2, c3, (p - 0.5) * 2) if p < 1 else (c3, c3, 0)
        fila.append(tuple(round(u + (v - u) * t) for u, v in zip(a, b)))
    imagen = Image.new("RGB", (ancho, 1))
    imagen.putdata(fila)
    imagen = imagen.resize((ancho, alto))
    luz = ImageChops.invert(Image.radial_gradient("L")).resize((alto * 7, alto * 4)).point(lambda v: max(0, v - 90) * 0.5)  # a cero antes del borde: sin escalón
    imagen.paste((255, 255, 255), (-alto * 2, -alto * 2), luz)
    return imagen


def crear_cabecera_degradada(app, titulo, subtitulo, logo=None, pista=""):
    """Banda de cabecera pintada en un lienzo: degradado según la hora, logo, saludo y fecha a la izquierda y, a la
    derecha, un marco liso (lienzo.derecha) para las fichas. lienzo.repintar() la ajusta al ancho y a las fichas."""
    from PIL import ImageTk
    s = escala(app)
    alto = round(CABECERA_ALTO * s)
    estado = {"momento": momento_del_dia(), "ancho": 0, "corte": 0, "foto": None}

    def colores():
        return DEGRADADOS[estado["momento"]]

    lienzo = tkinter.Canvas(app, height=alto, highlightthickness=0, bd=0, bg=colores()[2])
    derecha = ctk.CTkFrame(lienzo, fg_color=colores()[2], corner_radius=0)
    lienzo.derecha = derecha
    lienzo.marco_fichas = ctk.CTkFrame(derecha, fg_color="transparent", width=1, height=1)
    lienzo.marco_fichas.pack(anchor="e")
    if pista:
        ctk.CTkLabel(derecha, text=pista, font=estilo.fuente(estilo.TAM_PEQUENO), text_color="#eef1f7", height=16).pack(anchor="e", pady=(6, 0))
    fondo = lienzo.create_image(0, 0, anchor="nw")
    margen = round(20 * s)
    x_texto = margen
    if logo is not None:
        lado = round(38 * s)
        estado["logo"] = logo.convert("RGBA").resize((lado, lado))
        x_texto += lado + round(14 * s)
    familia_titulo, familia = estilo.fuente_titulo()[0], estilo.fuente()[0]
    lienzo.create_text(x_texto, alto / 2 - round(11 * s), anchor="w", text=titulo, fill="#ffffff", font=(familia_titulo, -round(estilo.TAM_APP * 1.35 * s), "bold"))
    lienzo.create_text(x_texto, alto / 2 + round(15 * s), anchor="w", text=subtitulo, fill="#eef1f7", font=(familia, -round(estilo.TAM_TEXTO * 1.35 * s)))
    ventana = lienzo.create_window(0, alto / 2, anchor="e", window=derecha)

    def repintar(evento=None, forzar=False):
        try:
            if not lienzo.winfo_exists():
                return
            ancho = lienzo.winfo_width()
            if ancho <= 1:
                return
            derecha.update_idletasks()
            corte = max(round(ancho * 0.25), min(round(ancho * 0.6), ancho - derecha.winfo_reqwidth() - margen * 2))
            if not forzar and (ancho, corte) == (estado["ancho"], estado["corte"]):
                return
            estado["ancho"], estado["corte"] = ancho, corte
            imagen = imagen_degradado(ancho, alto, colores(), corte)
            if estado.get("logo") is not None:
                imagen.paste(estado["logo"], (margen, (alto - estado["logo"].height) // 2), estado["logo"])
            estado["foto"] = ImageTk.PhotoImage(imagen, master=lienzo)  # se guarda: sin referencia Tk la borra
            lienzo.itemconfigure(fondo, image=estado["foto"])
            lienzo.coords(ventana, ancho - margen, alto / 2)
        except Exception:
            pass  # decorativo: si falla queda el color liso

    def revisar_hora():
        try:
            if not lienzo.winfo_exists():
                return
            nuevo = momento_del_dia()
            if nuevo != estado["momento"]:
                estado["momento"] = nuevo
                lienzo.configure(bg=colores()[2])
                derecha.configure(fg_color=colores()[2])
                repintar(forzar=True)
                if lienzo.al_cambiar is not None:
                    lienzo.al_cambiar()  # las fichas redondeadas toman el fondo nuevo
            lienzo.after(CABECERA_REVISION_MS, revisar_hora)
        except Exception:
            pass

    lienzo.repintar, lienzo.estado, lienzo.al_cambiar, lienzo.revisar_hora = repintar, estado, None, revisar_hora
    lienzo.bind("<Configure>", repintar)
    lienzo.after(CABECERA_REVISION_MS, revisar_hora)
    return lienzo


# ---------- celebración al terminar una tarea ----------

CELEBRA_TRAZO_MS = 380   # lo que tarda en dibujarse el check
CELEBRA_VISIBLE_MS = 1400  # tiempo en verde antes de volver a ser el botón de siempre
CELEBRA_PASO_MS = 16


def celebrar(app, boton):
    """El botón se pone verde y se dibuja un check blanco en su centro; al rato vuelve a su texto y su color.
    Decorativo: no hace nada con las animaciones apagadas, si el botón ya no existe o si ya está celebrando."""
    try:
        if not getattr(app, "config", {}).get("animacion_inicio", True) or boton is None or not boton.winfo_exists():
            return False
        if getattr(boton, "_celebrando", False):
            return False
        texto, fondo, encima = boton.cget("text"), boton.cget("fg_color"), boton.cget("hover_color")
        origen, verde = _color(boton, "fg_color"), app.color_ok
        if not (isinstance(origen, str) and origen.startswith("#") and verde.startswith("#")):
            return False
        boton._celebrando = True
        s = escala(boton)
        lado = round(26 * s)
        boton.configure(text="", fg_color=origen, hover_color=origen)
        lienzo = tkinter.Canvas(boton, width=lado, height=lado, bg=origen, highlightthickness=0, bd=0)
        lienzo.place(relx=0.5, rely=0.5, anchor="center")
        puntos = [(0.18 * lado, 0.54 * lado), (0.42 * lado, 0.76 * lado), (0.84 * lado, 0.28 * lado)]
        tramo1 = math.dist(puntos[0], puntos[1])
        total = tramo1 + math.dist(puntos[1], puntos[2])
        inicio = time.monotonic()
        pintado = [origen]

        def restaurar():
            boton._celebrando = False
            try:
                lienzo.destroy()
                if boton.winfo_exists():
                    cambios = {}
                    if boton.cget("text") == "":  # si otra tarea ya ha puesto su texto, no se pisa
                        cambios["text"] = texto
                    if _color(boton, "fg_color") == pintado[0]:  # solo si el color sigue siendo el que puso la celebración
                        cambios.update(fg_color=fondo, hover_color=encima)
                    if cambios:
                        boton.configure(**cambios)
            except Exception:
                pass

        def paso():
            try:
                if not boton.winfo_exists() or boton.cget("text") != "":
                    return restaurar()
                t = min(1.0, (time.monotonic() - inicio) * 1000 / CELEBRA_TRAZO_MS)
                color = mezclar(origen, verde, min(1.0, t * 2))
                boton.configure(fg_color=color, hover_color=color)
                pintado[0] = color
                lienzo.configure(bg=color)
                lienzo.delete("all")
                largo = total * _suave(t)
                linea = [*puntos[0]]
                if largo <= tramo1:
                    f = largo / tramo1
                    linea += [puntos[0][0] + (puntos[1][0] - puntos[0][0]) * f, puntos[0][1] + (puntos[1][1] - puntos[0][1]) * f]
                else:
                    f = (largo - tramo1) / (total - tramo1)
                    linea += [*puntos[1], puntos[1][0] + (puntos[2][0] - puntos[1][0]) * f, puntos[1][1] + (puntos[2][1] - puntos[1][1]) * f]
                lienzo.create_line(*linea, fill="#ffffff", width=max(2, round(3 * s)), capstyle="round", joinstyle="round")
                if t < 1.0:
                    boton.after(CELEBRA_PASO_MS, paso)
                else:
                    boton.after(CELEBRA_VISIBLE_MS, restaurar)
            except Exception:
                restaurar()

        paso()
        return True
    except Exception:
        return False

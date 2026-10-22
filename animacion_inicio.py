"""Animación de apertura: la ultraball aparece, se agita, se abre con un destello y deja paso a la app.

Se dibuja en una ventana sin bordes colocada encima del contenido de la ventana principal, con su mismo
color de fondo; al terminar se desvanece y deja ver la app. Un clic la salta."""
import math
import time
import tkinter

from PIL import Image, ImageDraw, ImageFilter, ImageTk

DURACION = 1.6        # segundos
INTERVALO_MS = 15     # ~60 fotogramas por segundo; si el equipo va lento se saltan fotogramas, no se alarga
ESPERA_MAXIMA_MS = 6000  # salvaguarda: pase lo que pase, la app se muestra
AMARILLO_BOLA = "#ffbe1a"


def _fase(t, inicio, fin):
    """Progreso de 0 a 1 de una fase de la animación."""
    if t <= inicio:
        return 0.0
    if t >= fin:
        return 1.0
    return (t - inicio) / (fin - inicio)


def _salida_suave(x):
    return 1 - (1 - x) ** 3


def _entrada_suave(x):
    return x ** 2


def _rebote(x):
    """Crece un poco por encima del tamaño final y vuelve (ease-out-back)."""
    c1 = 1.70158
    return 1 + (c1 + 1) * (x - 1) ** 3 + c1 * (x - 1) ** 2


def mezclar(color_a, color_b, proporcion):
    """Color intermedio entre dos colores #rrggbb (proporcion 0 = a, 1 = b)."""
    a = [int(color_a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(color_b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{round(x + (y - x) * proporcion):02x}" for x, y in zip(a, b))


def estado(t):
    """Estado de la animación en el segundo t: escala, giro, apertura de la tapa, luz, destello y opacidad."""
    aparecer = _fase(t, 0.0, 0.30)
    angulo = 0.0
    for inicio in (0.34, 0.62):  # dos sacudidas, con una pausa entre ellas
        x = _fase(t, inicio, inicio + 0.22)
        if 0 < x < 1:
            angulo = 16 * math.sin(x * 2 * math.pi) * (1 - 0.3 * x)
    return {
        "aparecer": aparecer,
        "escala": 0.55 + 0.45 * _rebote(aparecer) if aparecer < 1 else 1.0,
        "angulo": angulo,
        "apertura": _salida_suave(_fase(t, 0.90, 1.14)),
        "luz": _fase(t, 0.92, 1.12),
        "destello": _salida_suave(_fase(t, 1.04, 1.40)),
        "texto": aparecer * (1 - _salida_suave(_fase(t, 1.04, 1.40))),
        "opacidad": 1 - _entrada_suave(_fase(t, 1.18, DURACION)),
    }


DURACION_CAPTURA = 2.4


def estado_captura(t):
    """«¡Capturado!»: la bola cerrada se agita tres veces, hace clic, saltan estrellas y se desvanece."""
    aparecer = _fase(t, 0.0, 0.25)
    angulo = 0.0
    for inicio in (0.30, 0.62, 0.94):
        x = _fase(t, inicio, inicio + 0.24)
        if 0 < x < 1:
            angulo = 14 * math.sin(x * 2 * math.pi)
    clic = _fase(t, 1.22, 1.45)
    escala = 0.55 + 0.45 * _rebote(aparecer) if aparecer < 1 else 1.0
    return {
        "aparecer": aparecer,
        "escala": escala + 0.05 * math.sin(math.pi * clic),
        "angulo": angulo,
        "apertura": 0.0, "luz": 0.0, "destello": 0.0,
        "clic": clic,
        "estrellas": _salida_suave(_fase(t, 1.26, 1.95)),
        "texto": _salida_suave(_fase(t, 1.28, 1.50)),
        "opacidad": _fase(t, 0.0, 0.18) * (1 - _entrada_suave(_fase(t, 2.0, DURACION_CAPTURA))),
    }


def _estrella(cx, cy, radio):
    puntos = []
    for i in range(10):
        r = radio if i % 2 == 0 else radio * 0.45
        angulo = -math.pi / 2 + i * math.pi / 5
        puntos.append((cx + r * math.cos(angulo), cy + r * math.sin(angulo)))
    return puntos


class Ultraball:
    """Dibuja cada fotograma de la bola a partir de la imagen del icono, partida en tapa y base."""

    def __init__(self, imagen, tam, fondo, color_luz):
        self.tam = tam
        self.lado = int(tam * 2)
        self.fondo = fondo
        self.completa = imagen.convert("RGBA").resize((tam, tam), Image.LANCZOS)
        mitad = tam // 2
        self.tapa = self.completa.crop((0, 0, tam, mitad))
        self.base = self.completa.crop((0, mitad, tam, tam))
        # Halo de luz: elipse difuminada, se escala y se tiñe en cada fotograma
        mascara = Image.new("L", (160, 160), 0)
        ImageDraw.Draw(mascara).ellipse((48, 48, 112, 112), fill=255)  # margen amplio para que el difuminado no se corte
        self.halo = mascara.filter(ImageFilter.GaussianBlur(16))
        self.color_luz = tuple(int(color_luz.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))

    def centro(self):
        """Posición del centro de la bola dentro del fotograma (algo baja para que quepa la tapa al abrirse)."""
        return self.lado / 2, self.lado * 0.58

    def _pegar(self, lienzo, imagen, x, y):
        lienzo.paste(imagen, (round(x), round(y)), imagen)

    def dibujar(self, e):
        lienzo = Image.new("RGBA", (self.lado, self.lado), self.fondo)
        cx, cy = self.centro()
        tam = self.tam
        if e["apertura"] <= 0:
            bola = self.completa
            if e["escala"] != 1.0:
                lado = max(1, round(tam * e["escala"]))
                bola = bola.resize((lado, lado), Image.BICUBIC)
            if e["angulo"]:
                bola = bola.rotate(e["angulo"], resample=Image.BICUBIC, expand=True)
            self._pegar(lienzo, bola, cx - bola.width / 2, cy - bola.height / 2)
            if e.get("clic", 0) > 0 or e.get("estrellas", 0) > 0:
                lienzo.alpha_composite(self._clic_y_estrellas(e.get("clic", 0), e.get("estrellas", 0)))
        else:
            p = e["apertura"]
            base_y = cy + 0.05 * tam * p
            self._pegar(lienzo, self.base, cx - tam / 2, base_y)
            # La tapa sube y se aplana al inclinarse hacia atrás
            alto_tapa = max(1, round(self.tapa.height * (1 - 0.45 * p)))
            tapa = self.tapa.resize((tam, alto_tapa), Image.BICUBIC)
            borde_tapa = cy - 0.42 * tam * p
            self._pegar(lienzo, tapa, cx - tam / 2, borde_tapa - alto_tapa)
            if e["luz"] > 0:
                # El halo crece con el destello sin salirse del fotograma (lado = 2 x tam)
                ancho = max(1, round(tam * (0.8 + 0.7 * e["luz"] + 0.4 * e["destello"])))
                alto = max(1, round(tam * (0.3 + 0.7 * e["luz"] + 0.6 * e["destello"])))
                mascara = self.halo.resize((ancho, alto), Image.BILINEAR).point(lambda v: round(v * e["luz"]))
                luz = Image.new("RGBA", (ancho, alto), self.color_luz + (255,))
                luz.putalpha(mascara)
                self._pegar(lienzo, luz, cx - ancho / 2, (borde_tapa + base_y) / 2 - alto / 2)
        return lienzo.convert("RGB")

    def _clic_y_estrellas(self, clic, estrellas):
        """Capa transparente con el anillo del clic (sale del botón) y las estrellas que saltan hacia arriba."""
        capa = Image.new("RGBA", (self.lado, self.lado), (0, 0, 0, 0))
        dibujo = ImageDraw.Draw(capa)
        cx, cy = self.centro()
        tam = self.tam
        if 0 < clic < 1:
            radio = tam * (0.12 + 0.26 * clic)
            dibujo.ellipse((cx - radio, cy - radio, cx + radio, cy + radio), outline=self.color_luz + (round(255 * (1 - clic)),),
                           width=max(2, round(tam * 0.035)))
        if 0 < estrellas < 1:
            dorado = tuple(int(AMARILLO_BOLA.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
            opacidad = round(255 * (1 - estrellas ** 3))
            for grados in (-160, -125, -90, -55, -20):
                angulo = math.radians(grados)
                distancia = tam * (0.25 + 0.55 * estrellas)
                x = cx + math.cos(angulo) * distancia
                y = cy - 0.1 * tam + math.sin(angulo) * distancia
                dibujo.polygon(_estrella(x, y, tam * 0.085 * (1 - 0.35 * estrellas)), fill=dorado + (opacidad,),
                               outline=(176, 112, 0, opacidad))  # contorno dorado oscuro: se ve sobre fondo claro y oscuro
        return capa


class AnimacionApertura:
    """Capa sin bordes encima del contenido de la ventana, con la bola animada y uno o dos textos debajo.
    Sirve para la apertura (tapa la ventana mientras se coloca) y para «¡Capturado!» (aparece sobre la app)."""

    def __init__(self, ventana, fondo, color_texto, color_luz, al_terminar=None, funcion_estado=estado, duracion=DURACION,
                 ocultar_ventana=True, titulo="Palausa Aprovisionamiento", subtitulo=""):
        self.ventana = ventana
        self.al_terminar = al_terminar
        self.fondo = fondo
        self.color_texto = color_texto
        self.color_luz = color_luz
        self.funcion_estado = funcion_estado
        self.duracion = duracion
        self.ocultar_ventana = ocultar_ventana
        self.titulo = titulo
        self.subtitulo = subtitulo
        self.terminada = False
        self.capa = None
        self._foto = None
        self._geometria = None
        self._opacidad = None

    def iniciar(self, imagen, fuente, fuente_subtitulo=None):
        ventana, fondo = self.ventana, self.fondo
        if self.ocultar_ventana:
            ventana.attributes("-alpha", 0.0)  # la ventana principal se coloca sin verse
            ventana.update()
            ventana.lift()
            ventana.focus_force()
        else:
            ventana.update_idletasks()
        self.capa = tkinter.Toplevel(ventana)
        self.capa.overrideredirect(True)
        self.capa.configure(bg=fondo)
        if not self.ocultar_ventana:
            self.capa.attributes("-alpha", 0.0)
        self.capa.attributes("-topmost", True)
        self.lienzo = tkinter.Canvas(self.capa, bg=fondo, highlightthickness=0, bd=0, cursor="hand2")
        self.lienzo.pack(fill="both", expand=True)
        self.seguir_ventana()
        ancho, alto = self._geometria[:2]
        self.bola = Ultraball(imagen, max(110, min(240, round(min(ancho, alto) * 0.26))), fondo, self.color_luz)
        self.centro = (ancho / 2, alto * 0.45)
        self.imagen = self.lienzo.create_image(0, 0)
        self.texto = self.lienzo.create_text(self.centro[0], self.centro[1] + self.bola.tam * 0.78, text=self.titulo, font=fuente, fill=fondo)
        self.texto_secundario = None
        if self.subtitulo:
            self.texto_secundario = self.lienzo.create_text(self.centro[0], self.centro[1] + self.bola.tam * 1.02, text=self.subtitulo,
                                                            font=fuente_subtitulo or fuente, fill=fondo)
        self.capa.bind("<Button-1>", lambda evento: self.terminar())
        self.pintar(self.funcion_estado(0))
        self.capa.update()
        if self.ocultar_ventana:
            ventana.attributes("-alpha", 1.0)  # ya está tapada por la animación
        self.inicio = time.perf_counter()
        self.capa.after(INTERVALO_MS, self.paso)
        ventana.after(ESPERA_MAXIMA_MS, self.terminar)

    def seguir_ventana(self):
        """Mantiene la animación encima del contenido de la ventana principal (por si se mueve)."""
        geometria = (self.ventana.winfo_width(), self.ventana.winfo_height(), self.ventana.winfo_rootx(), self.ventana.winfo_rooty())
        if geometria != self._geometria:
            self._geometria = geometria
            self.capa.geometry("{}x{}+{}+{}".format(*geometria))

    def pintar(self, e):
        cx, cy = self.centro
        self._foto = ImageTk.PhotoImage(self.bola.dibujar(e))
        bx, by = self.bola.centro()
        self.lienzo.coords(self.imagen, cx - bx + self.bola.lado / 2, cy - by + self.bola.lado / 2)
        self.lienzo.itemconfigure(self.imagen, image=self._foto)
        self.lienzo.itemconfigure(self.texto, fill=mezclar(self.fondo, self.color_texto, e["texto"]))
        if self.texto_secundario is not None:
            self.lienzo.itemconfigure(self.texto_secundario, fill=mezclar(self.fondo, self.color_texto, e["texto"] * 0.85))
        opacidad = round(max(0.0, min(1.0, e["opacidad"])), 3)
        if opacidad != self._opacidad and (opacidad < 1 or self._opacidad is not None):
            self.capa.attributes("-alpha", opacidad)
        self._opacidad = opacidad

    def paso(self):
        if self.terminada:
            return
        try:
            t = time.perf_counter() - self.inicio
            if t >= self.duracion:
                self.terminar()
                return
            self.seguir_ventana()
            self.pintar(self.funcion_estado(t))
            self.capa.after(INTERVALO_MS, self.paso)
        except Exception:
            self.terminar()

    def terminar(self):
        if self.terminada:
            return
        self.terminada = True
        try:
            self.ventana.attributes("-alpha", 1.0)
        except Exception:
            pass
        if self.capa is not None:
            try:
                self.capa.destroy()
            except Exception:
                pass
        if self.al_terminar:
            self.al_terminar()


def reproducir(ventana, imagen, fondo, color_texto, color_luz, fuente, al_terminar=None):
    """Lanza la animación de apertura; si algo falla, la app se muestra directamente sin animación."""
    animacion = AnimacionApertura(ventana, fondo, color_texto, color_luz, al_terminar)
    try:
        animacion.iniciar(imagen, fuente)
    except Exception:
        animacion.terminar()
    return animacion


def reproducir_captura(ventana, imagen, fondo, color_texto, color_luz, titulo, subtitulo, fuente, fuente_subtitulo, al_terminar=None):
    """«¡Capturado!» sobre la app ya abierta (al terminar bien el flujo completo); un clic la salta."""
    animacion = AnimacionApertura(ventana, fondo, color_texto, color_luz, al_terminar, funcion_estado=estado_captura,
                                  duracion=DURACION_CAPTURA, ocultar_ventana=False, titulo=titulo, subtitulo=subtitulo)
    try:
        animacion.iniciar(imagen, fuente, fuente_subtitulo)
    except Exception:
        animacion.terminar()
    return animacion

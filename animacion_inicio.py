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
        "opacidad": 1 - _entrada_suave(_fase(t, 1.18, DURACION)),
    }


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


class AnimacionApertura:
    def __init__(self, ventana, fondo, color_texto, color_luz, al_terminar=None):
        self.ventana = ventana
        self.al_terminar = al_terminar
        self.fondo = fondo
        self.color_texto = color_texto
        self.color_luz = color_luz
        self.terminada = False
        self.capa = None
        self._foto = None
        self._geometria = None

    def iniciar(self, imagen, fuente):
        ventana, fondo = self.ventana, self.fondo
        ventana.attributes("-alpha", 0.0)  # la ventana principal se coloca sin verse
        ventana.update()
        ventana.lift()
        ventana.focus_force()
        self.capa = tkinter.Toplevel(ventana)
        self.capa.overrideredirect(True)
        self.capa.configure(bg=fondo)
        self.capa.attributes("-topmost", True)
        self.lienzo = tkinter.Canvas(self.capa, bg=fondo, highlightthickness=0, bd=0, cursor="hand2")
        self.lienzo.pack(fill="both", expand=True)
        self.seguir_ventana()
        ancho, alto = self._geometria[:2]
        self.bola = Ultraball(imagen, max(110, min(240, round(min(ancho, alto) * 0.26))), fondo, self.color_luz)
        self.centro = (ancho / 2, alto * 0.45)
        self.imagen = self.lienzo.create_image(0, 0)
        self.texto = self.lienzo.create_text(self.centro[0], self.centro[1] + self.bola.tam * 0.78, text="Palausa Aprovisionamiento",
                                             font=fuente, fill=fondo)
        self.capa.bind("<Button-1>", lambda evento: self.terminar())
        self.pintar(estado(0))
        self.capa.update()
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
        self.lienzo.itemconfigure(self.texto, fill=mezclar(self.fondo, self.color_texto, e["aparecer"] * (1 - e["destello"])))
        if e["opacidad"] < 1:
            self.capa.attributes("-alpha", max(0.0, e["opacidad"]))

    def paso(self):
        if self.terminada:
            return
        try:
            t = time.perf_counter() - self.inicio
            if t >= DURACION:
                self.terminar()
                return
            self.seguir_ventana()
            self.pintar(estado(t))
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
    """Lanza la animación; si algo falla, la app se muestra directamente sin animación."""
    animacion = AnimacionApertura(ventana, fondo, color_texto, color_luz, al_terminar)
    try:
        animacion.iniciar(imagen, fuente)
    except Exception:
        animacion.terminar()
    return animacion

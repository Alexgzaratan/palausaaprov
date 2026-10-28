"""Ventana para decidir las piezas caras (en azul) del filtro de rotación: pedirlas (pasan a blanco) o
ponerlas a 0 (pasan a rojo). Así el Excel que llega a Quiter solo tiene rojo, amarillo y blanco."""
import customtkinter as ctk

import estilo

PEDIR = "pedir"
A_CERO = "cero"


def euros(importe):
    """1850.5 -> '1.850,50 €'."""
    return f"{importe:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".") + " €"


def numero(valor):
    try:
        return float(str(valor).replace(",", "."))
    except (TypeError, ValueError):
        return None


class DialogoPiezasCaras:
    """piezas: [{"linea", "referencia", "descripcion", "pedidas", "unitario", "total"}].
    al_decidir(decisiones) recibe {línea: PEDIR | A_CERO}, o None si se deja para más tarde; si devuelve False
    (no se pudo guardar, p. ej. Excel abierto) la ventana sigue abierta para volver a intentarlo."""

    def __init__(self, app, piezas, pedido, limite, al_decidir):
        self.app = app
        self.piezas = piezas
        self.al_decidir = al_decidir
        self.elecciones = {}
        self.botones = {}
        self.cerrado = False
        ventana = ctk.CTkToplevel(app)
        self.ventana = ventana
        ventana.title(f"Piezas caras en {pedido}")
        ventana.configure(fg_color=app.color_panel_claro)
        ventana.transient(app)
        ventana.protocol("WM_DELETE_WINDOW", self.mas_tarde)
        ventana.bind("<Escape>", lambda evento: (self.mas_tarde(), "break")[1])

        tarjeta = ctk.CTkFrame(ventana, fg_color=app.color_superficie, corner_radius=estilo.RADIO_TARJETA, border_width=1, border_color=app.color_borde)
        tarjeta.pack(padx=16, pady=16, fill="both", expand=True)
        cuantas = len(piezas)
        ctk.CTkLabel(tarjeta, text=f"{cuantas} pieza{'s' if cuantas != 1 else ''} cara{'s' if cuantas != 1 else ''} en {pedido}", font=estilo.fuente_titulo(),
                     text_color=app.color_texto).pack(anchor="w", padx=estilo.PAD_X, pady=(14, 2))
        app.crear_etiqueta(tarjeta, f"Su precio por unidad supera {euros(limite)}. «Pedir» la deja en blanco y «Poner a 0» la pasa a rojo para Quiter.",
                           secundario=True, wraplength=560, justify="left").pack(anchor="w", padx=estilo.PAD_X, pady=(0, 10))

        lista = ctk.CTkScrollableFrame(tarjeta, fg_color="transparent", height=min(320, 76 * cuantas)) if cuantas > 4 else ctk.CTkFrame(tarjeta, fg_color="transparent")
        lista.pack(padx=estilo.PAD_X - 4, fill="both", expand=True)
        for pieza in piezas:
            self._fila(lista, pieza)

        pie = ctk.CTkFrame(tarjeta, fg_color="transparent")
        pie.pack(padx=estilo.PAD_X, pady=(8, 14), fill="x")
        app.crear_boton_secundario(pie, "Pedir todas", lambda: self.todas(PEDIR), alto=estilo.ALTO_BOTON_COMPACTO).pack(side="left")
        app.crear_boton_secundario(pie, "Todas a 0", lambda: self.todas(A_CERO), alto=estilo.ALTO_BOTON_COMPACTO).pack(side="left", padx=(6, 0))
        self.btn_aplicar = app.crear_boton_principal(pie, "Aplicar", self.aplicar, alto=estilo.ALTO_BOTON_COMPACTO)
        self.btn_aplicar.configure(width=110, state="disabled")
        self.btn_aplicar.pack(side="right")
        app.crear_boton_secundario(pie, "Más tarde", self.mas_tarde, alto=estilo.ALTO_BOTON_COMPACTO).pack(side="right", padx=(0, 6))

        ventana.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - ventana.winfo_width()) // 2
        y = app.winfo_rooty() + max(40, (app.winfo_height() - ventana.winfo_height()) // 3)
        ventana.geometry(f"+{max(0, x)}+{max(0, y)}")
        ventana.after(150, self._tomar_foco)

    def _tomar_foco(self):
        if self.cerrado or not self.ventana.winfo_exists():
            return
        try:
            self.ventana.lift()
            self.ventana.focus_force()
            self.ventana.grab_set()  # hasta decidir (o «Más tarde») no se usa la ventana principal
        except Exception:
            pass

    def _fila(self, padre, pieza):
        app = self.app
        fila = ctk.CTkFrame(padre, fg_color=app.color_superficie_alt, corner_radius=10)
        fila.pack(fill="x", pady=(0, 6))
        botones = ctk.CTkFrame(fila, fg_color="transparent")
        botones.pack(side="right", padx=10, pady=8)
        linea = pieza["linea"]
        self.botones[linea] = {}
        for opcion, texto in ((PEDIR, "Pedir"), (A_CERO, "Poner a 0")):
            boton = app.crear_boton_secundario(botones, texto, lambda o=opcion, l=linea: self.elegir(l, o), alto=estilo.ALTO_BOTON_COMPACTO, width=96)
            boton.pack(side="left", padx=(6 if opcion == A_CERO else 0, 0))
            self.botones[linea][opcion] = boton
        texto = ctk.CTkFrame(fila, fg_color="transparent")
        texto.pack(side="left", fill="x", expand=True, padx=(12, 6), pady=8)
        app.crear_etiqueta(texto, f"Lín. {linea} · {pieza['referencia']}", negrita=True).pack(anchor="w")
        unitario, total = pieza.get("unitario"), pieza.get("total")
        detalle = [str(pieza.get("descripcion") or "").strip()]
        if unitario is not None:
            pedidas = pieza.get("pedidas")
            cantidad = f"{pedidas:g} ud × " if isinstance(pedidas, (int, float)) else ""
            detalle.append(f"{cantidad}{euros(unitario)}" + (f" = {euros(total)}" if total is not None else ""))
        app.crear_etiqueta(texto, " · ".join(d for d in detalle if d), estilo.TAM_PEQUENO, secundario=True, wraplength=330, justify="left").pack(anchor="w")

    def elegir(self, linea, opcion):
        self.elecciones[linea] = opcion
        app = self.app
        for clave, boton in self.botones[linea].items():
            if clave == opcion:  # la elegida, en el color del tema (o rojo si va a 0)
                color, hover = (app.color_peligro, app.color_peligro_hover) if opcion == A_CERO else (app.color_principal, app.color_hover)
                boton.configure(fg_color=color, hover_color=hover, text_color="#ffffff", border_color=color)
            else:
                boton.configure(fg_color=app.color_panel_medio, hover_color=app.color_panel_medio_hover, text_color=app.color_texto, border_color=app.color_borde)
        self.btn_aplicar.configure(state="normal" if len(self.elecciones) == len(self.piezas) else "disabled")

    def todas(self, opcion):
        for pieza in self.piezas:
            self.elegir(pieza["linea"], opcion)

    def aplicar(self):
        if self.cerrado or len(self.elecciones) < len(self.piezas):
            return False
        if self.al_decidir(dict(self.elecciones)) is False:
            return False
        self._cerrar()
        return True

    def mas_tarde(self):
        if self.cerrado:
            return
        self._cerrar()
        self.al_decidir(None)

    def _cerrar(self):
        self.cerrado = True
        try:
            self.ventana.grab_release()
        except Exception:
            pass
        try:
            self.ventana.destroy()
        except Exception:
            pass

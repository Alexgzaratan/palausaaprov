"""Paletas, tipografías, iconos, tamaños y espaciados de la interfaz: un único sitio para mantener un aspecto uniforme."""

# Tipografía: se elige la mejor disponible al arrancar (configurar_fuentes). Valores por defecto para Windows 10.
FUENTE = "Segoe UI"
FUENTE_TITULOS = "Segoe UI Semibold"
FUENTE_MONO = "Consolas"
FUENTE_ICONOS = None  # sin fuente de iconos no se muestran iconos

TAM_APP = 20
TAM_TITULO = 16
TAM_BOTON = 13
TAM_TEXTO = 12
TAM_PEQUENO = 11   # tamaño mínimo
TAM_ICONO = 18

# Controles (esquinas algo más redondeadas: aspecto más amable)
ALTO_CONTROL = 36
ALTO_BOTON = 40
ALTO_BOTON_COMPACTO = 32
RADIO = 10
RADIO_TARJETA = 14

# Espaciados
PAD_X = 16          # margen horizontal dentro de las tarjetas
PAD_Y = 10          # separación vertical entre controles
PAD_FILA = 12       # relleno interior de una fila de ajuste
PAD_TARJETA = 8     # separación entre tarjetas

# Anchura mínima del contenido para repartir las tarjetas en dos columnas
ANCHO_DOS_COLUMNAS = 1100

# Iconos de Segoe Fluent Icons (Windows 11) / Segoe MDL2 Assets (Windows 10), comprobados visualmente
ICONOS = {
    "buscar": "", "filtro": "", "carrito": "", "calendario": "", "correo": "",
    "comparar": "", "documento": "", "coche": "", "factura": "", "reloj": "",
    "ajustes": "", "aviso": "", "check": "", "carpeta": "", "info": "",
    "estrella": "", "estrella_llena": "", "cerrar": "", "tareas": "", "abrir": "",
}


def configurar_fuentes(familias_disponibles):
    """Usa Segoe UI Variable, Cascadia Mono y Segoe Fluent Icons si están instaladas (Windows 11)."""
    global FUENTE, FUENTE_TITULOS, FUENTE_MONO, FUENTE_ICONOS
    familias = set(familias_disponibles)
    if "Segoe UI Variable Text" in familias:
        FUENTE = "Segoe UI Variable Text"
    if "Segoe UI Variable Display" in familias:
        FUENTE_TITULOS = "Segoe UI Variable Display"
    if "Cascadia Mono" in familias:
        FUENTE_MONO = "Cascadia Mono"
    for candidata in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
        if candidata in familias:
            FUENTE_ICONOS = candidata
            break


def fuente(tam=TAM_TEXTO, negrita=False, mono=False):
    familia = FUENTE_MONO if mono else FUENTE
    return (familia, tam, "bold") if negrita else (familia, tam)


def fuente_titulo(tam=TAM_TITULO):
    return (FUENTE_TITULOS, tam, "bold")


def fuente_icono(tam=TAM_ICONO):
    return (FUENTE_ICONOS, tam)


# ---------- temas de color ----------
# Neutros cálidos (lino) y verdosos (salvia); cada tema tiene variante clara y oscura.
_NEUTROS_LINO_CLARO = {
    "superficie": "#fffdf9",        # fondo de tarjetas y paneles
    "superficie_alt": "#f5efe6",    # fondo de sub-secciones dentro de una tarjeta
    "campo": "#fffefb",             # fondo de campos de texto
    "borde": "#e9e0d4",             # separadores y tarjetas (decorativo)
    "borde_campo": "#8f8477",       # borde de campos de texto: 3:1 frente al fondo
    "texto": "#2a2420",
    "texto_sec": "#675a4e",
    "texto_placeholder": "#76695c",
    "panel_medio": "#f0e9df",       # fondo de botones secundarios
    "panel_medio_hover": "#e6ddd0",
    "panel_claro": "#f8f3ec",       # fondo general de la ventana
    "ok": "#15803d",                # texto de éxito en los registros
    "texto_error": "#c2261d",       # texto de error en los registros
    "peligro": "#c2261d",           # fondo de botones de detener/eliminar
    "peligro_hover": "#9f1f18",
    "enfasis": "#1c1714",           # texto de énfasis (estado, porcentaje)
}
_NEUTROS_LINO_OSCURO = {
    "superficie": "#26211d", "superficie_alt": "#302a25", "campo": "#1f1b18",
    "borde": "#3d362f", "borde_campo": "#8f8376",
    "texto": "#f3ede4", "texto_sec": "#b8ab9c", "texto_placeholder": "#a69a8c",
    "panel_medio": "#38312b", "panel_medio_hover": "#453c34", "panel_claro": "#1a1613",
    "ok": "#4ade80", "texto_error": "#f87171", "peligro": "#c2261d", "peligro_hover": "#9f1f18", "enfasis": "#fbf7f1",
}
_NEUTROS_SALVIA_CLARO = {
    "superficie": "#fcfdfb", "superficie_alt": "#eef3ec", "campo": "#fdfefc",
    "borde": "#dbe4d8", "borde_campo": "#7f8b78",
    "texto": "#1f2a23", "texto_sec": "#55655a", "texto_placeholder": "#63705f",
    "panel_medio": "#e9efe6", "panel_medio_hover": "#dde6d9", "panel_claro": "#f3f6f1",
    "ok": "#15803d", "texto_error": "#c2261d", "peligro": "#c2261d", "peligro_hover": "#9f1f18", "enfasis": "#16201a",
}
_NEUTROS_SALVIA_OSCURO = {
    "superficie": "#1f2622", "superficie_alt": "#28302b", "campo": "#19201c",
    "borde": "#343e37", "borde_campo": "#7f8f84",
    "texto": "#e9efe8", "texto_sec": "#a8b6aa", "texto_placeholder": "#97a69a",
    "panel_medio": "#2f3832", "panel_medio_hover": "#3b463e", "panel_claro": "#151a17",
    "ok": "#4ade80", "texto_error": "#f87171", "peligro": "#c2261d", "peligro_hover": "#9f1f18", "enfasis": "#f5f8f4",
}

TEMAS = {
    "Lino": {
        "claro": {**_NEUTROS_LINO_CLARO, "principal": "#2f5fc4", "hover": "#244c9e"},
        "oscuro": {**_NEUTROS_LINO_OSCURO, "principal": "#2f5fc4", "hover": "#244c9e"},
    },
    "Terracota": {
        "claro": {**_NEUTROS_LINO_CLARO, "principal": "#b4532a", "hover": "#93431f"},
        "oscuro": {**_NEUTROS_LINO_OSCURO, "principal": "#b4532a", "hover": "#93431f"},
    },
    "Salvia": {
        "claro": {**_NEUTROS_SALVIA_CLARO, "principal": "#2f6b4f", "hover": "#24563f"},
        "oscuro": {**_NEUTROS_SALVIA_OSCURO, "principal": "#2f6b4f", "hover": "#24563f"},
    },
}
TEMA_POR_DEFECTO = "Terracota"

# Compatibilidad: paletas del tema por defecto
PALETA_CLARA = TEMAS[TEMA_POR_DEFECTO]["claro"]
PALETA_OSCURA = TEMAS[TEMA_POR_DEFECTO]["oscuro"]


# Parejas (texto, fondo, mínimo) que deben cumplir contraste WCAG AA en todas las paletas; las comprueban los tests
PAREJAS_CONTRASTE = (
    ("texto", "superficie", 4.5), ("texto", "superficie_alt", 4.5), ("texto", "campo", 4.5), ("texto", "panel_medio", 4.5),
    ("texto_sec", "superficie", 4.5), ("texto_sec", "superficie_alt", 4.5), ("texto_sec", "panel_claro", 4.5),
    ("texto_placeholder", "campo", 4.5), ("ok", "campo", 4.5), ("texto_error", "campo", 4.5), ("enfasis", "superficie", 4.5),
    ("#ffffff", "principal", 4.5), ("#ffffff", "hover", 4.5), ("#ffffff", "peligro", 4.5),
    ("borde_campo", "campo", 3.0), ("borde_campo", "superficie", 3.0), ("borde_campo", "superficie_alt", 3.0),
)


def luminancia(color):
    color = color.lstrip("#")
    canales = [int(color[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    lineal = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in canales]
    return 0.2126 * lineal[0] + 0.7152 * lineal[1] + 0.0722 * lineal[2]


def contraste(a, b):
    claro, oscuro = sorted((luminancia(a), luminancia(b)), reverse=True)
    return (claro + 0.05) / (oscuro + 0.05)


def paleta(oscuro, tema=TEMA_POR_DEFECTO):
    return TEMAS.get(tema, TEMAS[TEMA_POR_DEFECTO])["oscuro" if oscuro else "claro"]

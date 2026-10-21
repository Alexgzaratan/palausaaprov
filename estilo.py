"""Paletas, tamaños y espaciados de la interfaz: un único sitio para mantener un aspecto uniforme."""

FUENTE = "Segoe UI"
FUENTE_MONO = "Consolas"

# Tipografía (tamaño mínimo 11)
TAM_APP = 20
TAM_TITULO = 16
TAM_BOTON = 13
TAM_TEXTO = 12
TAM_PEQUENO = 11

# Controles
ALTO_CONTROL = 36
ALTO_BOTON = 40
ALTO_BOTON_COMPACTO = 32
RADIO = 8
RADIO_TARJETA = 12

# Espaciados
PAD_X = 16          # margen horizontal dentro de las tarjetas
PAD_Y = 10          # separación vertical entre controles
PAD_FILA = 12       # relleno interior de una fila de ajuste
PAD_TARJETA = 8     # separación entre tarjetas

# Anchura mínima del contenido para repartir las tarjetas en dos columnas
ANCHO_DOS_COLUMNAS = 1100


def fuente(tam=TAM_TEXTO, negrita=False, mono=False):
    familia = FUENTE_MONO if mono else FUENTE
    return (familia, tam, "bold") if negrita else (familia, tam)


PALETA_CLARA = {
    "principal": "#2563eb",         # acento primario (acciones principales)
    "hover": "#1d4ed8",
    "superficie": "#ffffff",        # fondo de tarjetas y paneles
    "superficie_alt": "#f1f4f9",    # fondo de sub-secciones dentro de una tarjeta
    "campo": "#ffffff",             # fondo de campos de texto
    "borde": "#dbe1ea",             # separadores y tarjetas (decorativo)
    "borde_campo": "#838c9c",       # borde de campos de texto: 3:1 frente al fondo
    "texto": "#1e2733",
    "texto_sec": "#5b6779",         # 4.5:1 incluso sobre las filas grises
    "texto_placeholder": "#6b7587",
    "panel_medio": "#eef1f6",       # fondo de botones secundarios
    "panel_medio_hover": "#dfe4ec",
    "panel_claro": "#f5f6f9",       # fondo general de la ventana
    "ok": "#15803d",                # texto de éxito en los registros
    "texto_error": "#dc2626",       # texto de error en los registros
    "peligro": "#dc2626",           # fondo de botones de detener/eliminar
    "peligro_hover": "#b91c1c",
    "enfasis": "#0f172a",           # texto de énfasis (estado, porcentaje)
}

PALETA_OSCURA = {
    "principal": "#2563eb",         # el blanco encima cumple 4.5:1
    "hover": "#1d4ed8",
    "superficie": "#1e293b",
    "superficie_alt": "#26374d",
    "campo": "#16202f",
    "borde": "#334155",
    "borde_campo": "#7b8aa0",
    "texto": "#e2e8f0",
    "texto_sec": "#94a3b8",
    "texto_placeholder": "#8291a7",
    "panel_medio": "#2b3a4f",
    "panel_medio_hover": "#374763",
    "panel_claro": "#0f172a",
    "ok": "#22c55e",
    "texto_error": "#f87171",
    "peligro": "#dc2626",
    "peligro_hover": "#b91c1c",
    "enfasis": "#f1f5f9",
}


# Parejas (texto, fondo, mínimo) que deben cumplir contraste WCAG AA en ambas paletas; las comprueban los tests
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


def paleta(oscuro):
    return PALETA_OSCURA if oscuro else PALETA_CLARA

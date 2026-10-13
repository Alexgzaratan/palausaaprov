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
    "borde": "#dbe1ea",
    "texto": "#1e2733",
    "texto_sec": "#64748b",
    "panel_medio": "#eef1f6",       # fondo de botones secundarios
    "panel_medio_hover": "#dfe4ec",
    "panel_claro": "#f5f6f9",       # fondo general de la ventana
    "ok": "#16a34a",
    "peligro": "#dc2626",
    "peligro_hover": "#b91c1c",
    "enfasis": "#0f172a",           # texto de énfasis (estado, porcentaje)
}

PALETA_OSCURA = {
    "principal": "#3b82f6",
    "hover": "#2563eb",
    "superficie": "#1e293b",
    "superficie_alt": "#26374d",
    "campo": "#16202f",
    "borde": "#334155",
    "texto": "#e2e8f0",
    "texto_sec": "#94a3b8",
    "panel_medio": "#2b3a4f",
    "panel_medio_hover": "#374763",
    "panel_claro": "#0f172a",
    "ok": "#22c55e",
    "peligro": "#f87171",
    "peligro_hover": "#ef4444",
    "enfasis": "#f1f5f9",
}


def paleta(oscuro):
    return PALETA_OSCURA if oscuro else PALETA_CLARA

import pytest

import estilo


@pytest.mark.parametrize("nombre,paleta", [("clara", estilo.PALETA_CLARA), ("oscura", estilo.PALETA_OSCURA)])
def test_contraste_wcag_aa_en_ambas_paletas(nombre, paleta):
    fallos = []
    for texto, fondo, minimo in estilo.PAREJAS_CONTRASTE:
        color_texto = paleta.get(texto, texto)
        color_fondo = paleta.get(fondo, fondo)
        ratio = estilo.contraste(color_texto, color_fondo)
        if ratio < minimo:
            fallos.append(f"{texto} sobre {fondo}: {ratio:.2f} < {minimo}")
    assert not fallos, f"paleta {nombre}: " + "; ".join(fallos)


def test_contraste_conocido():
    assert round(estilo.contraste("#000000", "#ffffff"), 1) == 21.0
    assert round(estilo.contraste("#ffffff", "#ffffff"), 1) == 1.0


def test_pestana_activa_con_texto_blanco_y_las_demas_normales(app):
    app.ir_a_pestana(1)
    botones = app.tabview._segmented_button._buttons_dict
    assert botones["Pedidos"].cget("text_color") == "#ffffff"
    assert botones["Automatización"].cget("text_color") == app.color_texto
    app.ir_a_pestana(0)
    assert botones["Automatización"].cget("text_color") == "#ffffff"
    assert botones["Pedidos"].cget("text_color") == app.color_texto


def test_el_campo_con_foco_se_resalta(app):
    campo = app.excel_auto_entry
    assert campo.cget("border_color") == app.color_borde_campo
    campo._entry.event_generate("<FocusIn>")
    app.update()
    assert campo.cget("border_color") == app.color_principal and campo.cget("border_width") == 2
    campo._entry.event_generate("<FocusOut>")
    app.update()
    assert campo.cget("border_color") == app.color_borde_campo and campo.cget("border_width") == 1


def test_subtitulo_se_ajusta_al_ancho_de_la_tarjeta(app):
    etiqueta = app.crear_etiqueta(app.tab_auto, "texto largo " * 20, wraplength=470)
    app.ajustar_subtitulo(etiqueta, 1200)
    ancho_ancho = etiqueta._ancho_ajuste
    app.ajustar_subtitulo(etiqueta, 500)
    assert ancho_ancho > etiqueta._ancho_ajuste >= 200
    etiqueta.destroy()

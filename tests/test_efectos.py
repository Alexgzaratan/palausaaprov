import time

import pytest

import animacion_inicio
import app_automatizacion_optimizada as app_mod
import efectos


def bombear(app, segundos):
    fin = time.time() + segundos
    while time.time() < fin:
        app.update()
        time.sleep(0.01)


# ---------- avisos ----------

def test_aviso_aparece_y_se_va_solo(app):
    aviso = app.avisar("Configuración guardada.", "ok", duracion=300)
    bombear(app, 0.1)
    assert aviso.winfo_exists() and aviso.winfo_ismapped()
    assert app.avisos.historial[-1] == ("ok", "Configuración guardada.")
    bombear(app, 0.7)
    assert not aviso.winfo_exists()


def test_un_clic_cierra_el_aviso(app):
    aviso = app.avisar("Hola", duracion=10000)
    bombear(app, 0.2)
    aviso._canvas.event_generate("<Button-1>")  # el fondo visible de un CTkFrame es su lienzo
    bombear(app, 0.4)
    assert not aviso.winfo_exists()


def test_los_avisos_se_apilan_sin_taparse(app):
    primero = app.avisar("Primero", duracion=10000)
    segundo = app.avisar("Segundo", duracion=10000)
    bombear(app, 0.3)
    assert primero.winfo_rooty() + primero.winfo_height() <= segundo.winfo_rooty()  # el más reciente abajo
    for aviso in (primero, segundo):
        app.avisos.cerrar(aviso)
    bombear(app, 0.3)


def test_avisos_quedan_por_encima_del_panel_de_actividad(app):
    aviso = app.avisar("Encima del panel", duracion=10000)
    bombear(app, 0.3)
    assert aviso.winfo_rooty() + aviso.winfo_height() <= app.panel_actividad.winfo_rooty()
    app.avisos.cerrar(aviso)
    bombear(app, 0.3)


def test_guardar_configuracion_avisa_sin_ventana(app, monkeypatch):
    monkeypatch.setattr(app_mod.messagebox, "showinfo", lambda *a, **k: pytest.fail("no debe abrirse una ventana"))
    app.guardar_config_ui()
    assert app.avisos.historial[-1] == ("ok", "Configuración guardada.")


# ---------- arrastrar archivos ----------

def test_soltar_un_excel_en_su_campo(app, tmp_path):
    excel = tmp_path / "con espacio" / "S40P.xlsx"
    excel.parent.mkdir()
    excel.write_text("x")
    campo = app.excel_auto_entry
    assert app.soltar_en_campo(campo, "{" + str(excel) + "}", "archivo", (".xlsx",))
    assert campo.get() == str(excel).replace("\\", "/")


def test_soltar_algo_que_no_encaja_no_cambia_el_campo(app, tmp_path):
    campo = app.excel_auto_entry
    campo.delete(0, "end")
    campo.insert(0, "C:/antes.xlsx")
    pdf = tmp_path / "factura.pdf"
    pdf.write_text("x")
    assert not app.soltar_en_campo(campo, str(pdf), "archivo", (".xlsx",))
    assert campo.get() == "C:/antes.xlsx"
    assert app.avisos.historial[-1] == ("aviso", "Aquí va un Excel (.xlsx).")
    assert not app.soltar_en_campo(app.folder_rot_entry, str(pdf), "carpeta")
    assert app.avisos.historial[-1][0] == "aviso"
    assert app.soltar_en_campo(app.folder_rot_entry, str(tmp_path), "carpeta")
    assert app.folder_rot_entry.get() == str(tmp_path).replace("\\", "/")


def test_los_campos_aceptan_arrastrar(app):
    if not app.arrastre_activo:
        pytest.skip("tkdnd no disponible en este equipo")
    enlazados = app.excel_rot_entry._entry.bind()
    assert any("Drop" in evento for evento in enlazados)


# ---------- botones ocupados ----------

def test_boton_ocupado_muestra_giro_y_tiempo_y_se_para_solo(app):
    boton = app.btn_rotacion
    efectos.marcar_ocupado(boton, "Analizando", intervalo_ms=30)
    primero = boton.cget("text")
    assert primero[0] in efectos.GIRO and primero.endswith("Analizando · 0:00")
    bombear(app, 0.1)
    assert boton.cget("text")[0] != primero[0] or boton.cget("text") != primero  # gira
    boton.configure(text="Resaltar por rotación")  # la tarea termina
    bombear(app, 0.15)
    assert boton.cget("text") == "Resaltar por rotación"


# ---------- tarjetas ----------

def test_tarjeta_se_resalta_y_vuelve_al_salir(app, monkeypatch):
    tarjeta = app.btn_rotacion.master
    while not hasattr(tarjeta, "_resalte"):
        tarjeta = tarjeta.master
    dentro = {"valor": True}
    monkeypatch.setattr(efectos, "puntero_dentro", lambda widget: dentro["valor"])
    tarjeta._entrar_resalte()
    bombear(app, 0.3)
    assert tarjeta._resalte["nivel"] == efectos.NIVEL_RESALTE
    assert tarjeta.cget("border_color") == animacion_inicio.mezclar(app.color_borde, app.color_principal, efectos.NIVEL_RESALTE)
    bombear(app, 0.3)
    assert tarjeta._resalte["dentro"]  # sigue dentro (por encima de sus controles): no parpadea
    dentro["valor"] = False
    bombear(app, 0.6)
    assert tarjeta._resalte["nivel"] == 0 and tarjeta.cget("border_color") == app.color_borde


# ---------- «¡Capturado!» ----------

def test_fases_de_la_captura():
    assert animacion_inicio.estado_captura(0)["opacidad"] == 0
    giros = [animacion_inicio.estado_captura(t / 100)["angulo"] for t in range(30, 120)]
    assert max(giros) > 10 and min(giros) < -10  # se agita
    assert animacion_inicio.estado_captura(1.3)["clic"] > 0
    medio = animacion_inicio.estado_captura(1.6)
    assert 0 < medio["estrellas"] < 1 and medio["texto"] == 1 and medio["opacidad"] == 1
    assert animacion_inicio.estado_captura(animacion_inicio.DURACION_CAPTURA)["opacidad"] == 0


def test_fotogramas_de_la_captura(app):
    bola = animacion_inicio.Ultraball(app.imagen_logo(), 120, "#1a1613", "#ffdf8d")
    for t in (0, 0.5, 1.3, 1.6, 2.2):
        assert bola.dibujar(animacion_inicio.estado_captura(t)).size == (bola.lado, bola.lado)


def test_captura_se_muestra_y_termina(app):
    animacion = app.mostrar_captura(24, 46, "S39P")
    assert animacion.capa.winfo_exists() and animacion.subtitulo == "24 de 46 líneas a poner a 0  ·  S39P"
    fin = time.time() + 5
    while not animacion.terminada and time.time() < fin:
        app.update()
        time.sleep(0.01)
    assert animacion.terminada and not animacion.capa.winfo_exists()
    assert float(app.attributes("-alpha")) == 1.0


def test_captura_desactivada_con_las_animaciones(app):
    app.config["animacion_inicio"] = False
    assert app.mostrar_captura(1, 2, "S1") is None

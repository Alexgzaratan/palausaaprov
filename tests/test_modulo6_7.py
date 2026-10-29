"""Comparador (formato conservado, nombre por fecha, recuento, cabecera sin marcar, «Últimos dos») y
Reclamaciones Power Line (fotos giradas, HEIC, carpetas temporales y registro de errores)."""
import os
import time
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import Font
from PIL import Image

import app_automatizacion_optimizada as app_mod
import reclamaciones_power_line as pl


# ---------- comparador ----------

def excel(ruta, filas, segunda_hoja=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    for fila in filas:
        ws.append(fila)
    ws["A1"].font = Font(bold=True)
    ws.column_dimensions["B"].width = 33
    ws["C2"].number_format = "0.00"
    if segunda_hoja:
        wb.create_sheet("OTRA")["A1"] = "se conserva"
    wb.save(ruta)
    return str(ruta)


def test_compara_conservando_el_formato_y_sin_marcar_la_cabecera(tmp_path):
    cabecera = ["REF", "DESCRIPCION", "CANT"]
    viejo = excel(tmp_path / "Penurias 09-07-2026.xlsx", [cabecera, ["A1", "x", 1], ["B2", "y", 2]])
    nuevo = excel(tmp_path / "Penurias 10-07-2026.xlsx", [cabecera, ["A1", "x", 1], ["C3", "z", 3], [None, None, None], ["B2", "y", 9]], segunda_hoja=True)
    salida = app_mod.ruta_resultado_comparacion(tmp_path / "salida", nuevo)
    assert Path(salida).name == "Penurias 10-07-2026 comparado.xlsx"
    resultados, registro = [], []
    app_mod.comparar_excel_task(viejo, nuevo, salida, "FFFF00", log_callback=lambda t, tag=None: registro.append(t), resultado_callback=resultados.append)
    assert resultados == [(1, 3)]  # solo A1 se repite entera; la fila vacía no cuenta
    assert any("1 de 3 filas ya estaban" in t for t in registro)
    wb = openpyxl.load_workbook(salida)
    ws = wb.active
    marcadas = [r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).fill.fill_type == "solid"]
    assert marcadas == [2]  # ni la cabecera (igual en los dos) ni las distintas
    assert ws["A1"].font.bold and ws.column_dimensions["B"].width == 33 and ws["C2"].number_format == "0.00"
    assert wb.sheetnames == ["Sheet", "OTRA"] and wb["OTRA"]["A1"].value == "se conserva"


def test_dos_ultimos_excel(tmp_path):
    for nombre in ("Penurias 08-07-2026.xlsx", "Penurias 10-07-2026.xlsx", "Penurias 09-07-2026.xlsx",
                   "Penurias 10-07-2026 comparado.xlsx", "~$Penurias 10-07-2026.xlsx", "notas.txt"):
        (tmp_path / nombre).write_bytes(b"x")
    anterior, ultimo = app_mod.dos_ultimos_excel(tmp_path)
    assert (anterior.name, ultimo.name) == ("Penurias 09-07-2026.xlsx", "Penurias 10-07-2026.xlsx")
    assert app_mod.dos_ultimos_excel(tmp_path / "no_existe") is None


def test_boton_ultimos_dos_rellena_los_campos(app, tmp_path):
    for nombre in ("Penurias 09-07-2026.xlsx", "Penurias 10-07-2026.xlsx"):
        (tmp_path / nombre).write_bytes(b"x")
    app.tabview.set("Comparador")
    app.construir_pestana("Comparador")
    app.excel_new_entry.delete(0, "end")
    app.excel_new_entry.insert(0, str(tmp_path / "Penurias 01-07-2026.xlsx"))  # basta con que exista su carpeta
    app.usar_dos_ultimos_excel()
    assert Path(app.excel_old_entry.get()).name == "Penurias 09-07-2026.xlsx" and Path(app.excel_new_entry.get()).name == "Penurias 10-07-2026.xlsx"


# ---------- Power Line ----------

def foto_vertical(ruta):
    """Foto de 40x20 guardada «tumbada» con la orientación 6 en el EXIF (girar 90°), como hacen los móviles."""
    imagen = Image.new("RGB", (40, 20), "red")
    exif = Image.Exif()
    exif[0x0112] = 6
    imagen.save(ruta, "JPEG", exif=exif)
    return str(ruta)


def test_las_fotos_se_enderezan_segun_el_exif(tmp_path):
    ruta = foto_vertical(tmp_path / "movil.jpg")
    with pl.abrir_imagen(ruta) as imagen:
        assert imagen.size == (20, 40)
    preparada = pl.DocumentoPowerLine._preparar_foto(ruta, tmp_path, "bulto")
    with Image.open(preparada) as imagen:
        assert imagen.size == (20, 40)


def test_admite_fotos_heic(tmp_path):
    if not pl.ADMITE_HEIC:
        pytest.skip("pillow-heif no instalado")
    ruta = tmp_path / "iphone.heic"
    Image.new("RGB", (30, 10), "blue").save(ruta, "HEIF")
    preparada = pl.DocumentoPowerLine._preparar_foto(str(ruta), tmp_path, "etiqueta")
    with Image.open(preparada) as imagen:
        assert imagen.format == "JPEG" and imagen.size == (30, 10)
    assert "*.heic" in pl.EXTENSIONES_FOTO


def test_limpia_las_carpetas_temporales_antiguas(tmp_path):
    antigua, reciente, ajena = tmp_path / (pl.PREFIJO_ADJUNTOS + "viejo"), tmp_path / (pl.PREFIJO_ADJUNTOS + "nuevo"), tmp_path / "otra_cosa"
    for carpeta in (antigua, reciente, ajena):
        carpeta.mkdir()
        (carpeta / "f.jpg").write_bytes(b"x")
    hace_dos_dias = time.time() - 48 * 3600
    os.utime(antigua, (hace_dos_dias, hace_dos_dias))
    os.utime(ajena, (hace_dos_dias, hace_dos_dias))
    assert pl.limpiar_adjuntos_antiguos(carpeta_temporal=tmp_path) == 1
    assert not antigua.exists() and reciente.exists() and ajena.exists()


def test_la_carpeta_temporal_solo_se_crea_al_adjuntar(app, monkeypatch, tmp_path):
    creadas = []
    original = pl.tempfile.mkdtemp
    monkeypatch.setattr(pl.tempfile, "mkdtemp", lambda prefix="": creadas.append(prefix) or original(prefix=prefix, dir=tmp_path))
    modulo = pl.ModuloReclamacionesPowerLine(app, {})
    assert creadas == []
    assert modulo.carpeta_adjuntos.is_dir() and modulo.carpeta_adjuntos == modulo.carpeta_adjuntos and creadas == [pl.PREFIJO_ADJUNTOS]


def test_los_errores_van_al_registro_comun(monkeypatch):
    registrados = []
    monkeypatch.setattr(pl, "registrar_error_detallado", lambda mensaje, traza, ref=None: registrados.append((mensaje, ref)))
    pl.registrar_error_reclamacion("Traceback...", "Word no responde", "9812345680")
    assert registrados == [("Reclamaciones Power Line: Word no responde", "9812345680")]


def test_si_falla_el_registro_el_boton_no_se_queda_bloqueado(app, monkeypatch):
    modulo = pl.ModuloReclamacionesPowerLine(app, {"power_line_ruta_plantilla": "no_existe.docx"})
    llamadas = []
    monkeypatch.setattr(pl, "registrar_error_reclamacion", lambda *a: (_ for _ in ()).throw(OSError("disco")))
    monkeypatch.setattr(app, "after", lambda ms, funcion, *a: llamadas.append(funcion))
    with pytest.raises(OSError):
        modulo._generar_en_segundo_plano({"pedido": "", "referencia": "R", "descripcion": "", "litigio": "1"}, Path("x.docx"), False)
    assert llamadas  # el aviso de error (que vuelve a activar el botón) se programa igualmente

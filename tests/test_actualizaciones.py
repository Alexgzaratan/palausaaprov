import io
import json
import os
import zipfile

import pytest

import actualizaciones
import cambio_version


class Respuesta(io.BytesIO):
    def __init__(self, datos, cabeceras=None):
        super().__init__(datos)
        self.headers = cabeceras or {}

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()


def release(etiqueta, zips=("Palausa_Aprovisionamiento.zip",), **extra):
    datos = {"tag_name": etiqueta, "body": "Novedades", "assets": [{"name": n, "browser_download_url": f"https://x/{n}", "size": 10} for n in zips]}
    datos.update(extra)
    return lambda peticion, timeout=None: Respuesta(json.dumps(datos).encode())


def test_comparar_versiones():
    assert actualizaciones.es_mas_nueva("1.0.1", "1.0.0")
    assert actualizaciones.es_mas_nueva("v1.10.0", "1.9.9")
    assert not actualizaciones.es_mas_nueva("1.0.0", "1.0.0")
    assert not actualizaciones.es_mas_nueva("0.9", "1.0.0")


def test_buscar_actualizacion():
    info = actualizaciones.buscar_actualizacion("1.0.0", abrir=release("v1.2.0"))
    assert info["version"] == "1.2.0" and info["notas"] == "Novedades" and info["url"].endswith(".zip")
    assert actualizaciones.buscar_actualizacion("1.2.0", abrir=release("v1.2.0")) is None
    assert actualizaciones.buscar_actualizacion("1.0.0", abrir=release("v1.2.0", zips=())) is None
    assert actualizaciones.buscar_actualizacion("1.0.0", abrir=release("v1.2.0", prerelease=True)) is None


def test_descarga_completa_e_incompleta(tmp_path):
    destino = tmp_path / "a.zip"
    progreso = []
    actualizaciones.descargar("https://x/a.zip", destino, 5, progreso.append, abrir=lambda p, timeout=None: Respuesta(b"12345"))
    assert destino.read_bytes() == b"12345" and progreso[-1] == 1.0
    with pytest.raises(ValueError):
        actualizaciones.descargar("https://x/a.zip", tmp_path / "b.zip", 9, abrir=lambda p, timeout=None: Respuesta(b"123"))
    assert not (tmp_path / "b.zip").exists() and not (tmp_path / "b.zip.parcial").exists()


def instalacion(base):
    """Instalación como la del usuario: lanzador, programa/ y sus datos al lado."""
    (base / "programa").mkdir(parents=True)
    (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).write_text("programa viejo")
    (base / "programa" / "solo_en_la_vieja.txt").write_text("x")
    (base / cambio_version.NOMBRE_EJECUTABLE).write_text("lanzador viejo")
    (base / "config_palausa_aprovisionamiento.json").write_text('{"usuario": "yo"}')
    (base / "log_facturas_madd.txt").write_text("historial")


def paquete(ruta, con_programa=True, extra=None):
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr(cambio_version.NOMBRE_EJECUTABLE, "lanzador nuevo")
        z.writestr("LEEME.txt", "léeme")
        if con_programa:
            z.writestr("programa/" + cambio_version.NOMBRE_EJECUTABLE, "programa nuevo")
            z.writestr("programa/_internal/a/b.dll", "dll")
        for nombre, contenido in (extra or {}).items():
            z.writestr(nombre, contenido)
    return ruta


def test_actualizacion_completa_conserva_los_datos(tmp_path):
    base = tmp_path / "Palausa"
    instalacion(base)
    actualizaciones.preparar_desde_zip(paquete(tmp_path / "p.zip"), base, "1.1.0")
    assert cambio_version.version_preparada(base) == "1.1.0"
    assert (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).read_text() == "programa viejo"  # aún no se ha aplicado
    assert cambio_version.aplicar_actualizacion(base)
    assert (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).read_text() == "programa nuevo"
    assert (base / "programa" / "_internal" / "a" / "b.dll").exists()
    assert not (base / "programa" / "solo_en_la_vieja.txt").exists()
    assert (base / "programa_anterior" / "solo_en_la_vieja.txt").exists()  # copia de la versión anterior
    assert not (base / "programa" / cambio_version.MARCA_LISTO).exists() and cambio_version.version_preparada(base) is None
    assert (base / "config_palausa_aprovisionamiento.json").read_text() == '{"usuario": "yo"}'
    assert (base / "log_facturas_madd.txt").read_text() == "historial"
    assert not (base / "LEEME.txt").exists()
    lanzador = base / cambio_version.NOMBRE_EJECUTABLE
    assert cambio_version.actualizar_lanzador(base, lanzador)
    assert lanzador.read_text() == "lanzador nuevo"
    cambio_version.actualizar_lanzador(base, lanzador)  # el siguiente arranque borra el viejo
    assert not (base / (cambio_version.NOMBRE_EJECUTABLE + cambio_version.SUFIJO_VIEJO)).exists()


def test_sin_actualizacion_preparada_no_se_toca_nada(tmp_path):
    base = tmp_path / "Palausa"
    instalacion(base)
    assert not cambio_version.aplicar_actualizacion(base)
    assert not cambio_version.actualizar_lanzador(base)
    assert (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).read_text() == "programa viejo"


def test_paquete_sin_programa_o_con_rutas_peligrosas(tmp_path):
    base = tmp_path / "Palausa"
    instalacion(base)
    with pytest.raises(ValueError):
        actualizaciones.preparar_desde_zip(paquete(tmp_path / "a.zip", con_programa=False), base, "1.1.0")
    with pytest.raises(ValueError):
        actualizaciones.preparar_desde_zip(paquete(tmp_path / "b.zip", extra={"programa/../../fuera.txt": "x"}), base, "1.1.0")
    assert cambio_version.version_preparada(base) is None
    assert not (tmp_path / "fuera.txt").exists()


def test_si_el_programa_sigue_abierto_espera_y_luego_lo_deja_para_otra_vez(tmp_path, monkeypatch):
    base = tmp_path / "Palausa"
    instalacion(base)
    actualizaciones.preparar_desde_zip(paquete(tmp_path / "p.zip"), base, "1.1.0")
    renombrar = os.rename
    intentos = []

    def ocupado(origen, destino):
        if os.path.basename(str(origen)) == "programa":
            intentos.append(1)
            raise PermissionError("en uso")
        return renombrar(origen, destino)

    monkeypatch.setattr(cambio_version.os, "rename", ocupado)
    assert not cambio_version.aplicar_actualizacion(base, espera_max=0.2, intervalo=0.05)
    assert len(intentos) > 1
    assert (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).read_text() == "programa viejo"
    assert cambio_version.version_preparada(base) == "1.1.0"  # se aplicará en el próximo arranque
    monkeypatch.setattr(cambio_version.os, "rename", renombrar)
    assert cambio_version.aplicar_actualizacion(base)


def test_si_falla_al_poner_la_nueva_vuelve_la_anterior(tmp_path, monkeypatch):
    base = tmp_path / "Palausa"
    instalacion(base)
    actualizaciones.preparar_desde_zip(paquete(tmp_path / "p.zip"), base, "1.1.0")
    renombrar = os.rename

    def falla_la_nueva(origen, destino):
        if os.path.basename(str(origen)) == cambio_version.NOMBRE_NUEVO:
            raise PermissionError("bloqueado")
        return renombrar(origen, destino)

    monkeypatch.setattr(cambio_version.os, "rename", falla_la_nueva)
    assert not cambio_version.aplicar_actualizacion(base)
    assert (base / "programa" / cambio_version.NOMBRE_EJECUTABLE).read_text() == "programa viejo"


def test_descargar_y_preparar_borra_la_descarga(tmp_path):
    base = tmp_path / "Palausa"
    instalacion(base)
    datos = paquete(tmp_path / "p.zip").read_bytes()
    info = {"version": "1.1.0", "url": "https://x/p.zip", "tamano": len(datos), "nombre": "p.zip"}
    actualizaciones.descargar_y_preparar(info, base, abrir=lambda p, timeout=None: Respuesta(datos))
    assert cambio_version.version_preparada(base) == "1.1.0"
    assert not (base / actualizaciones.CARPETA_DESCARGA).exists()


def test_en_desarrollo_no_hay_instalacion_actualizable():
    assert actualizaciones.carpeta_instalacion() is None


def test_boton_de_actualizacion_en_la_cabecera(app):
    assert not hasattr(app, "btn_actualizacion") or not app.btn_actualizacion.winfo_exists()
    app.mostrar_aviso_actualizacion({"version": "9.0.0", "notas": "", "url": "u", "tamano": 0, "nombre": "p.zip"})
    assert app.btn_actualizacion.cget("text") == "Actualizar a 9.0.0"
    app.version_preparada = "9.0.0"
    app.actualizar_resumen()
    assert app.btn_actualizacion.cget("text") == "Reiniciar para actualizar a 9.0.0"
    app.version_preparada = None
    app.actualizacion_disponible = None
    app.actualizar_resumen()

"""Historial de las facturas MADD a partir de log_facturas_madd.txt: qué días se han enviado, cuáles faltan
(la tarea no se ejecutó o falló) y el resumen de la semana. También dónde se guarda cada factura: en
subcarpetas por mes (AAAA-MM); las antiguas siguen en la carpeta principal y se buscan en los dos sitios."""
import re
from datetime import date, datetime, timedelta
from pathlib import Path

MAX_DIAS_PENDIENTES = 10  # como mucho se recuperan los 10 últimos días laborables sin enviar

_CABECERA = re.compile(r"^--- (EJECUCIÓN|CORREO) \[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] ---")


def leer_registro(ruta):
    """Bloques del registro como dicts: tipo ('ejecucion'/'correo'), cuando, fecha (objetivo), y según el tipo
    facturas y error, o estado y adjuntos. Las líneas que no se entienden se ignoran."""
    bloques = []
    try:
        lineas = Path(ruta).read_text(encoding="utf-8").splitlines()
    except OSError:
        return bloques
    actual = None
    for linea in lineas:
        cabecera = _CABECERA.match(linea)
        if cabecera:
            actual = {"tipo": "ejecucion" if cabecera.group(1) == "EJECUCIÓN" else "correo",
                      "cuando": datetime.strptime(cabecera.group(2), "%Y-%m-%d %H:%M:%S"), "fecha": None}
            bloques.append(actual)
            continue
        if actual is None or ":" not in linea:
            continue
        clave, valor = (parte.strip() for parte in linea.split(":", 1))
        if clave == "Fecha objetivo":
            try:
                actual["fecha"] = date.fromisoformat(valor)
            except ValueError:
                pass
        elif clave == "Nº facturas descargadas":
            actual["facturas"] = int(valor) if valor.isdigit() else 0
        elif clave == "Error":
            actual["error"] = valor
        elif clave == "Estado":
            actual["estado"] = valor
        elif clave == "Destinatarios":
            adjuntos = re.search(r"adjuntos:\s*(\d+)", valor)
            actual["adjuntos"] = int(adjuntos.group(1)) if adjuntos else 0
    return [b for b in bloques if b["fecha"] is not None]


def es_envio(bloque):
    """Correo de facturas que ha salido o está en la bandeja de salida (saldrá solo). Los avisos de fallo no cuentan."""
    estado = bloque.get("estado", "")
    return bloque["tipo"] == "correo" and (estado.startswith("enviado") or estado.startswith("PENDIENTE"))


def dias_enviados(bloques):
    return {b["fecha"] for b in bloques if es_envio(b)}


def laborables_entre(desde, hasta):
    """Días laborables d con desde < d <= hasta."""
    dias, dia = [], desde + timedelta(days=1)
    while dia <= hasta:
        if dia.weekday() < 5:
            dias.append(dia)
        dia += timedelta(days=1)
    return dias


def dias_pendientes(objetivo, enviados, maximo=MAX_DIAS_PENDIENTES):
    """Días laborables hasta 'objetivo' (incluido) que faltan desde el último envío. Sin historial no se
    recupera nada hacia atrás: solo 'objetivo' si no se ha enviado."""
    anteriores = [d for d in enviados if d <= objetivo]
    if not anteriores:
        return [] if objetivo in enviados else [objetivo]
    return [d for d in laborables_entre(max(anteriores), objetivo) if d not in enviados][-maximo:]


def resumen_semana(bloques, hoy=None):
    """(facturas enviadas, correos, ejecuciones con fallo) de esta semana, de lunes a hoy."""
    hoy = hoy or date.today()
    lunes = hoy - timedelta(days=hoy.weekday())
    de_la_semana = [b for b in bloques if lunes <= b["cuando"].date() <= hoy]
    envios = [b for b in de_la_semana if es_envio(b)]
    fallos = sum(1 for b in de_la_semana if b["tipo"] == "ejecucion" and b.get("error"))
    return sum(b.get("adjuntos", 0) for b in envios), len(envios), fallos


def texto_resumen_semana(bloques, hoy=None):
    facturas, correos, fallos = resumen_semana(bloques, hoy)
    if not correos and not fallos:
        return "Esta semana todavía no se ha enviado ningún correo de facturas."
    texto = f"Esta semana: {facturas} factura{'s' if facturas != 1 else ''} en {correos} correo{'s' if correos != 1 else ''}"
    return texto + f" · {fallos} fallo{'s' if fallos != 1 else ''}" if fallos else texto + " · sin fallos"


def texto_dias(dias):
    return ", ".join(d.strftime("%d/%m") for d in dias)


# ---------- dónde está cada factura ----------

def carpeta_del_mes(carpeta, fecha):
    return Path(carpeta) / f"{fecha:%Y-%m}"


def ruta_nueva_factura(carpeta, numero, fecha):
    return carpeta_del_mes(carpeta, fecha) / f"Factura_{numero}_{fecha.isoformat()}.pdf"


def factura_existente(carpeta, nombre, fecha):
    """La factura ya descargada, en su subcarpeta del mes o en la carpeta principal (las antiguas); o None."""
    for ruta in (carpeta_del_mes(carpeta, fecha) / nombre, Path(carpeta) / nombre):
        if ruta.exists():
            return ruta
    return None


def facturas_del_dia(carpeta, fecha):
    """PDF ya descargados de ese día (Factura_<número>_AAAA-MM-DD.pdf), en la subcarpeta del mes o en la principal."""
    patron = f"Factura_*_{fecha.isoformat()}.pdf"
    try:
        encontradas = {r.name: r for r in Path(carpeta).glob(patron)}
        encontradas.update({r.name: r for r in carpeta_del_mes(carpeta, fecha).glob(patron)})
    except OSError:
        return []
    return sorted(encontradas.values(), key=lambda r: r.name)


def fecha_de_factura(ruta):
    coincidencia = re.search(r"_(\d{4}-\d{2}-\d{2})\.pdf$", str(ruta))
    return date.fromisoformat(coincidencia.group(1)) if coincidencia else None

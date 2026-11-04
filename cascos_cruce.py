"""Cruce de CASCOS_PROVISIONES con el informe de ventas de un cliente (como la skill cruce-cascos-ventas).

En CASCOS_PROVISIONES (hoja «RECAMBIOS PALAUSA») el Nº ALBARÁN de algunos clientes se escribe con el número de documento
y no con la Refer. Se corrige cruzándolo con el informe de ventas de Quiter (hoja «Informe»):
  - Al Nº ALBARÁN y al Documento se les quita todo lo que no sea dígito: «ALBARAN4126 103585» = «4126103585».
  - Si cruzan, el Nº ALBARÁN pasa a ser la Refer. (número entero). Si ya es una Refer. del informe, se deja.
  - Si no cruza, se deja igual y la celda se pinta en amarillo.
  - Solo filas del cliente con fecha (columna A) igual o posterior a «desde», y solo la columna Nº ALBARÁN.
El original no se toca: el resultado se guarda en la carpeta de cruces, por año y mes (ver ruta_salida).

Lo mismo sirve para el Excel de PIEZA NUEVA (última hoja con columnas «Cliente» y «Nº Albarán», p. ej.
«PIEZA_NUEVA _2025»). Ahí una celda puede llevar varios albaranes separados por «/»: se cruza cada uno."""
import datetime
import re
from pathlib import Path

HOJA_CASCOS = "RECAMBIOS PALAUSA"
HOJA_VENTAS = "Informe"
CLIENTE_POR_DEFECTO = "8049"
AMARILLO = "FFFF00"
# tipo: (prefijo del archivo de salida, hoja fija o None = la última que tenga las columnas, columna cliente, columna albarán)
TIPOS = {
    "cascos": ("CASCOS", HOJA_CASCOS, "CLIENTE", "Nº ALBARÁN"),
    "pieza_nueva": ("PIEZA NUEVA", None, "CLIENTE", "Nº ALBARÁN"),
}


def _columnas(ws, col_cliente, col_albaran):
    """(columna cliente, columna albarán) en base 1 comparando la cabecera sin mayúsculas; None si no están."""
    cabecera = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[1]]
    try:
        return cabecera.index(col_cliente.upper()) + 1, cabecera.index(col_albaran.upper()) + 1
    except ValueError:
        return None


def elegir_hoja(wb, tipo):
    _, hoja, col_cliente, col_albaran = TIPOS[tipo]
    if hoja is not None:
        if hoja not in wb.sheetnames:
            raise ValueError(f"el Excel de cascos no tiene la hoja «{hoja}»")
        ws = wb[hoja]
        if _columnas(ws, col_cliente, col_albaran) is None:
            raise ValueError("el Excel de cascos no tiene las columnas «CLIENTE» y «Nº ALBARÁN»")
        return ws
    candidatas = [ws for ws in wb.worksheets if ws.max_row > 1 and _columnas(ws, col_cliente, col_albaran)]
    if not candidatas:
        raise ValueError("el Excel no tiene ninguna hoja con datos y columnas «Cliente» y «Nº Albarán»")
    return candidatas[-1]  # la del año en curso es la última


def digitos(valor):
    return re.sub(r"\D", "", str(valor or ""))


def leer_ventas(ruta_ventas):
    """({clave del documento: Refer.}, fecha más antigua). Solo las líneas de albarán (sin facturas ni la fila de total)."""
    import pandas as pd
    hojas = pd.ExcelFile(ruta_ventas).sheet_names
    v = pd.read_excel(ruta_ventas, sheet_name=HOJA_VENTAS if HOJA_VENTAS in hojas else 0, header=0).dropna(subset=["Documento"])
    v = v[v["Documento"].astype(str).str.startswith("ALBARAN")]
    mapa = {digitos(d): int(r) for d, r in zip(v["Documento"], v["Refer."]) if str(r).strip() not in ("", "nan")}
    fecha_min = pd.to_datetime(v["Fecha"], dayfirst=True, errors="coerce").min() if "Fecha" in v.columns else None
    return mapa, (None if fecha_min is None or pd.isna(fecha_min) else fecha_min.to_pydatetime())


def motivo_sin_cruce(valor, claves):
    """Por qué no cruza: albarán anterior al informe, Refer. antigua (8 cifras) o a revisar a mano."""
    clave = digitos(valor)
    if re.fullmatch(r"\d{8}", str(valor).strip()):
        return "Refer. antigua (8 cifras)"
    # La clave es serie (4 cifras, p. ej. 4126) + nº de documento: si el número es menor que el más bajo del informe
    # para esa serie, es un albarán de antes del periodo del informe
    minimos = {}
    for k in claves:
        if len(k) > 4 and k.isdigit():
            minimos[k[:4]] = min(minimos.get(k[:4], int(k[4:])), int(k[4:]))
    if len(clave) > 4 and clave[:4] in minimos and int(clave[4:]) < minimos[clave[:4]]:
        return "albarán anterior al informe de ventas"
    return "revisar a mano"


MESES = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
SUBCARPETA_POR_DEFECTO = ("ABONO CASCOS", "CRUCES 8049")


def carpeta_por_defecto():
    """Escritorio\\ABONO CASCOS\\CRUCES 8049 si existe la carpeta ABONO CASCOS en el escritorio; si no, ''."""
    escritorio = Path.home() / "Desktop"
    base = escritorio / SUBCARPETA_POR_DEFECTO[0]
    return str(base / SUBCARPETA_POR_DEFECTO[1]) if base.is_dir() else ""


def ruta_salida(carpeta_destino, cliente=CLIENTE_POR_DEFECTO, hoy=None, tipo="cascos"):
    """carpeta\\AAAA\\MES\\CASCOS <cliente> dd-mm-aaaa.xlsx (PIEZA NUEVA ... para ese tipo), como las hojas de pedido.
    Si ya hay uno de ese día, «(2)», «(3)»..."""
    hoy = hoy or datetime.date.today()
    carpeta = Path(carpeta_destino) / str(hoy.year) / MESES[hoy.month - 1]
    base = f"{TIPOS[tipo][0]} {str(cliente).strip()} {hoy:%d-%m-%Y}"
    ruta, n = carpeta / f"{base}.xlsx", 2
    while ruta.exists():
        ruta, n = carpeta / f"{base} ({n}).xlsx", n + 1
    return ruta


def cruzar_celda(valor, mapa, refs):
    """(nuevo valor, partes sin cruce). Cada albarán separado por «/» se cambia por su Refer. si cruza; las que ya son
    una Refer. del informe se dejan. Un solo albarán queda como número; varios, como «Refer./Refer.»."""
    partes = [p.strip() for p in str(valor).strip().split("/") if p.strip()]
    nuevas, sin = [], []
    for parte in partes:
        if re.fullmatch(r"\d+", parte) and int(parte) in refs:
            nueva = int(parte)
        elif digitos(parte) in mapa:
            nueva = mapa[digitos(parte)]
        else:
            nueva = parte
            sin.append(parte)
        if nueva not in nuevas:  # «albarán/Refer.» del mismo documento queda en una sola Refer.
            nuevas.append(nueva)
    if len(nuevas) == 1:
        return nuevas[0], sin
    return "/".join(str(n) for n in nuevas), sin


def cruzar(ruta_ventas, ruta_cascos, salida, cliente=CLIENTE_POR_DEFECTO, desde=None, tipo="cascos"):
    """Hace el cruce y guarda el resultado en 'salida'. Devuelve {desde, cambiados, ya_estaban, sin_cruce: [(fila, fecha, valor, motivo)], hoja}.
    'desde' (date/datetime): si no se da, el día 1 del mes de la fecha más antigua del informe.
    tipo: «cascos» (CASCOS_PROVISIONES) o «pieza_nueva» (Excel de PIEZA NUEVA)."""
    import openpyxl
    from openpyxl.styles import PatternFill
    if Path(salida).resolve() == Path(ruta_cascos).resolve():
        raise ValueError("el resultado no puede sobrescribir el CASCOS_PROVISIONES original")
    mapa, fecha_min = leer_ventas(ruta_ventas)
    if not mapa:
        raise ValueError("el informe de ventas no tiene líneas de albarán (columna «Documento» que empiece por ALBARAN)")
    refs = set(mapa.values())
    if desde is None:
        if fecha_min is None:
            raise ValueError("no se puede saber desde qué fecha cruzar: indica la fecha")
        desde = datetime.datetime(fecha_min.year, fecha_min.month, 1)
    elif not isinstance(desde, datetime.datetime):
        desde = datetime.datetime(desde.year, desde.month, desde.day)
    cliente = str(cliente).strip()

    wb = openpyxl.load_workbook(ruta_cascos)
    try:
        ws = elegir_hoja(wb, tipo)
        col_cliente, col_albaran = _columnas(ws, *TIPOS[tipo][2:])
        amarillo = PatternFill("solid", fgColor=AMARILLO)
        cambiados, sin_cruce, ya = [], [], 0
        for fila in range(2, ws.max_row + 1):
            if str(ws.cell(fila, col_cliente).value).strip().replace(".0", "") != cliente:
                continue
            fecha = ws.cell(fila, 1).value
            if not isinstance(fecha, datetime.datetime) or fecha < desde:
                continue
            celda = ws.cell(fila, col_albaran)
            valor = celda.value
            if valor is None:
                continue
            nuevo, sin = cruzar_celda(valor, mapa, refs)
            igual = str(nuevo) == str(valor).strip()
            if sin:
                sin_cruce.append((fila, fecha.date(), valor, motivo_sin_cruce(sin[0], mapa)))
                celda.fill = amarillo
            elif igual:
                ya += 1
            else:
                cambiados.append((fila, fecha.date(), valor, nuevo))
            if not igual:  # si no cruza nada, la celda se queda como estaba (número o texto)
                celda.value = nuevo
        Path(salida).parent.mkdir(parents=True, exist_ok=True)
        wb.save(salida)
        hoja = ws.title
    finally:
        wb.close()
    return {"desde": desde.date(), "cambiados": cambiados, "ya_estaban": ya, "sin_cruce": sin_cruce, "hoja": hoja}


def resumen(resultado):
    sin = resultado["sin_cruce"]
    lineas = [f"Desde el {resultado['desde']:%d/%m/%Y}: {len(resultado['cambiados'])} albaranes cambiados por su Refer., "
              f"{resultado['ya_estaban']} ya estaban bien y {len(sin)} sin cruce (en amarillo)."]
    por_motivo = {}
    for _, _, _, motivo in sin:
        por_motivo[motivo] = por_motivo.get(motivo, 0) + 1
    lineas += [f"  · {n} {motivo}" for motivo, n in sorted(por_motivo.items(), key=lambda x: -x[1])]
    a_mano = [f"fila {f} ({d:%d/%m}): {v}" for f, d, v, m in sin if m == "revisar a mano"][:15]
    if a_mano:
        lineas += ["", "A revisar a mano:"] + [f"  {x}" for x in a_mano]
    if por_motivo.get("albarán anterior al informe de ventas", 0) >= 5:
        lineas += ["", "Hay bastantes albaranes anteriores al informe: saca un informe de ventas más largo."]
    return "\n".join(lineas)

"""Ficheros TXT de Mopar y Stock FIAT: lectura del Excel, numeración F<n> (un número nuevo en cada envío),
nombres con letra (A..Z, AA, AB...), escritura sin sobrescribir nunca y registro de lo generado para avisar
si el mismo Excel se convierte dos veces."""
import hashlib
import json
import os
import re
from datetime import datetime
from pathlib import Path

import xlrd

from configuracion import cargar_configuracion, guardar_configuracion
from rutas_app import ruta_datos

LINEAS_POR_FICHERO = 10
REGISTRO = ruta_datos("ficheros_mopar_generados.json", migrar=False)


# ---------- lectura del Excel (cada línea: referencia 13 + cantidad 7 + ubicación 8) ----------

def _linea(referencia, cantidad, ubicacion):
    return f"{referencia.replace('2/', '').zfill(13)[:13]}{str(cantidad).zfill(7)[:7]}{ubicacion.zfill(8)[:8]}"


def _cantidad(valor, fila, avisos):
    try:
        return int(float(valor))
    except Exception:
        avisos.append(f"Cantidad no válida en fila {fila}, se usará 0.")
        return 0


def lineas_mopar(ruta_excel, avisos=None):
    """Excel de Mopar con cabecera: columnas «Referencia», «Cant» y «Cliente» (los dígitos del cliente son la ubicación).
    Lanza ValueError si falta alguna columna."""
    avisos = avisos if avisos is not None else []
    ws = xlrd.open_workbook(ruta_excel).sheet_by_index(0)
    cabecera = [str(ws.cell_value(0, i)).strip() for i in range(ws.ncols)] if ws.nrows else []
    for columna in ("Referencia", "Cant", "Cliente"):
        if columna not in cabecera:
            raise ValueError(f"No se encuentra la columna '{columna}' en el Excel.")
    col_ref, col_cant, col_cliente = (cabecera.index(c) for c in ("Referencia", "Cant", "Cliente"))
    lineas = []
    for r in range(1, ws.nrows):
        try:
            referencia = str(ws.cell_value(r, col_ref)).strip()
            cantidad = _cantidad(ws.cell_value(r, col_cant), r + 1, avisos)
            cliente = "".join(filter(str.isdigit, str(ws.cell_value(r, col_cliente)).strip()))
            lineas.append(_linea(referencia, cantidad, cliente))
        except Exception as e:
            avisos.append(f"Error en fila {r + 1}: {e}")
    return lineas


def lineas_stock_fiat(ruta_excel, avisos=None):
    """Excel de Stock FIAT sin cabecera: referencia en la columna B y cantidad en la D; ubicación 00000000."""
    avisos = avisos if avisos is not None else []
    ws = xlrd.open_workbook(ruta_excel).sheet_by_index(0)
    if ws.ncols < 4:
        raise ValueError("El Excel debe tener al menos columnas hasta la D.")

    def valor_limpio(valor):
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()

    lineas = []
    for r in range(ws.nrows):
        try:
            referencia = valor_limpio(ws.cell_value(r, 1))
            if not referencia:
                continue
            lineas.append(_linea(referencia, _cantidad(ws.cell_value(r, 3), r + 1, avisos), ""))
        except Exception as e:
            avisos.append(f"Error en fila {r + 1}: {e}")
    return lineas


# ---------- numeración y nombres ----------

def letra(indice):
    """0 -> A, 25 -> Z, 26 -> AA, 27 -> AB... (nunca se repite una letra dentro del mismo número)."""
    texto = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        texto = chr(ord("A") + resto) + texto
    return texto


def nombres_ficheros(numero, total_lineas):
    if total_lineas <= LINEAS_POR_FICHERO:
        return [f"F{numero}.txt"]
    grupos = (total_lineas + LINEAS_POR_FICHERO - 1) // LINEAS_POR_FICHERO
    return [f"F{numero}{letra(i)}.txt" for i in range(grupos)]


def texto_ficheros(nombres):
    """['F55A.txt', 'F55B.txt', 'F55C.txt'] -> 'F55A–F55C'; un solo fichero -> 'F55'."""
    sin_extension = [Path(n).stem for n in nombres]
    return sin_extension[0] if len(sin_extension) == 1 else f"{sin_extension[0]}–{sin_extension[-1]}"


def numero_mas_alto(carpeta):
    try:
        numeros = [int(m.group(1)) for n in os.listdir(carpeta) if (m := re.match(r"F(\d+)", n))]
    except OSError:
        return 0
    return max(numeros, default=0)


def siguiente_numero(carpeta, clave, config=None):
    """Número del próximo envío: uno más que el mayor entre la carpeta y el último usado (por si los ficheros
    ya enviados se mueven de la carpeta)."""
    config = cargar_configuracion() if config is None else config
    try:
        ultimo = int(config.get(f"{clave}_numero") or 0)
    except (TypeError, ValueError):
        ultimo = 0
    return max(numero_mas_alto(carpeta), ultimo) + 1


def reservar_numero(carpeta, clave):
    config = cargar_configuracion()
    numero = siguiente_numero(carpeta, clave, config)
    config[f"{clave}_fecha"] = datetime.now().strftime("%Y-%m-%d")
    config[f"{clave}_numero"] = numero
    guardar_configuracion(config)
    return numero


def escribir(lineas, carpeta, numero):
    """Escribe los ficheros de un envío. Si alguno ya existiera no se escribe nada (FileExistsError)."""
    nombres = nombres_ficheros(numero, len(lineas))
    rutas = [Path(carpeta) / n for n in nombres]
    existentes = [r.name for r in rutas if r.exists()]
    if existentes:
        raise FileExistsError(f"ya existe {', '.join(existentes)}: no se sobrescribe nada")
    for i, ruta in enumerate(rutas):
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("\n".join(lineas[i * LINEAS_POR_FICHERO:(i + 1) * LINEAS_POR_FICHERO]))
    return nombres


# ---------- registro: ¿este Excel ya se convirtió? ----------

def firma(lineas):
    return hashlib.sha256("\n".join(lineas).encode("utf-8")).hexdigest()


def _leer_registro(ruta=None):
    try:
        with open(ruta or REGISTRO, encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, ValueError):
        return []


def ya_generado(clave, lineas, ruta=None):
    """El registro del envío anterior con exactamente las mismas líneas, o None."""
    huella = firma(lineas)
    return next((r for r in reversed(_leer_registro(ruta)) if r.get("clave") == clave and r.get("firma") == huella), None)


def anotar(clave, lineas, nombres, ruta=None):
    ruta = ruta or REGISTRO
    datos = _leer_registro(ruta)[-199:] + [{"clave": clave, "firma": firma(lineas), "ficheros": nombres, "lineas": len(lineas),
                                             "cuando": datetime.now().isoformat(timespec="seconds")}]
    temporal = ruta + ".tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=1)
    os.replace(temporal, ruta)


def texto_confirmacion(clave, lineas, carpeta, numero, nombre="Mopar", ahora=None):
    """Texto de la vista previa antes de generar (con el aviso si ese Excel ya se convirtió)."""
    nombres = nombres_ficheros(numero, len(lineas))
    cuantos = len(nombres)
    texto = (f"Se generará{'n' if cuantos != 1 else ''} {cuantos} fichero{'s' if cuantos != 1 else ''} {nombre} "
             f"({texto_ficheros(nombres)}) con {len(lineas)} línea{'s' if len(lineas) != 1 else ''} en {carpeta}.")
    anterior = ya_generado(clave, lineas)
    if anterior:
        cuando = datetime.fromisoformat(anterior["cuando"])
        dia = "hoy" if cuando.date() == (ahora or datetime.now()).date() else f"el {cuando:%d/%m}"
        texto = (f"⚠️ Este Excel ya se convirtió en {texto_ficheros(anterior['ficheros'])} {dia} a las {cuando:%H:%M}. "
                 f"Si lo generas otra vez, esas líneas se enviarán dos veces.\n\n" + texto)
    return texto + "\n\n¿Generar?", anterior is not None

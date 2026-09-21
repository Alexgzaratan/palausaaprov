import customtkinter as ctk
import threading
import os
import sys
import openpyxl
from openpyxl import load_workbook
import csv
import re
import json
import ctypes
import pandas as pd
import time
import traceback
import shutil
import xlrd
import string
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright
from tkinter import filedialog, messagebox
from openpyxl.styles import PatternFill
from concurrent.futures import ThreadPoolExecutor
from reclamaciones_power_line import ModuloReclamacionesPowerLine

# --- CONFIGURACIÓN DE WINDOWS PARA ICONO EN BARRA DE TAREAS ---
try:
    myappid = 'Palausa.Aprovisionamiento'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

# --- CONFIGURACIÓN GLOBAL ---
CONFIG_FILE = "config_palausa_aprovisionamiento.json"
ICONO_NOMBRE = "ultraball.ico"
ERROR_FOLDER = "errores_capturas"
LOG_DETALLADO = "error_log_detallado.txt"

# Bloqueo para escritura segura en Excel desde múltiples hilos
excel_lock = threading.Lock()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ICONO_PATH = resource_path(ICONO_NOMBRE)

if not os.path.exists(ERROR_FOLDER):
    os.makedirs(ERROR_FOLDER)

# Variables globales para control de ejecución
automation_running = False
pedidos_running = False
comparador_running = False
filtro_rotacion_running = False
mopar_running = False
stock_fiat_running = False
hoja_pedidos_running = False
avisos_running = False
madd_running = False
stop_event_auto = threading.Event()
stop_event_pedidos = threading.Event()
stop_event_comparador = threading.Event()
stop_event_filtro = threading.Event()
stop_event_mopar = threading.Event()
stop_event_stock_fiat = threading.Event()

def tarea_activa():
    """Nombre de la tarea en curso que escribe sobre un Excel, o None si no hay ninguna."""
    if automation_running:
        return "Automatizacion"
    if pedidos_running:
        return "Generacion de pedidos"
    if comparador_running:
        return "Comparador de Excel"
    if mopar_running:
        return "Fichero MOPAR"
    if stock_fiat_running:
        return "Stock FIAT"
    return None

def cargar_configuracion():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            registrar_error_detallado(str(e), traceback.format_exc(), ref="cargar_configuracion")
            return {}
    return {}

def guardar_configuracion(config_dict):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config_dict, f, indent=4)
    except Exception as e:
        registrar_error_detallado(str(e), traceback.format_exc(), ref="guardar_configuracion")

def registrar_error_detallado(error_msg, traceback_str, ref=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_DETALLADO, "a", encoding="utf-8") as f:
        f.write(f"--- ERROR [{timestamp}] ---\n")
        if ref: f.write(f"Referencia: {ref}\n")
        f.write(f"Mensaje: {error_msg}\n")
        f.write(f"Traceback:\n{traceback_str}\n")
        f.write("-" * 30 + "\n\n")

def abrir_ruta_en_sistema(ruta):
    if not ruta or not os.path.exists(ruta):
        return False
    try:
        os.startfile(ruta)
        return True
    except Exception:
        return False

def crear_copia_seguridad(ruta_archivo, carpeta_destino="backups_pedidos"):
    Path(carpeta_destino).mkdir(parents=True, exist_ok=True)
    nombre_base = Path(ruta_archivo).stem
    extension = Path(ruta_archivo).suffix
    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_backup = Path(carpeta_destino) / f"{nombre_base}_backup_{marca_tiempo}{extension}"
    shutil.copy2(ruta_archivo, ruta_backup)
    return str(ruta_backup)

def obtener_numero_mopar_del_dia(carpeta_salida, prefijo_clave="mopar"):
    config = cargar_configuracion()
    clave_fecha = f"{prefijo_clave}_fecha"
    clave_numero = f"{prefijo_clave}_numero"
    hoy = datetime.now().strftime("%Y-%m-%d")
    ultimo_dia = config.get(clave_fecha)
    ultimo_numero = config.get(clave_numero)

    if ultimo_dia == hoy and ultimo_numero is not None:
        return int(ultimo_numero)

    ultimo_f = 0
    for f in os.listdir(carpeta_salida):
        m = re.match(r"F(\d+)", f)
        if m:
            ultimo_f = max(ultimo_f, int(m.group(1)))

    proximo_f = ultimo_f + 1
    config[clave_fecha] = hoy
    config[clave_numero] = proximo_f
    guardar_configuracion(config)
    return proximo_f

# ---------------------------------------------------
# LÓGICA DE GESTIÓN DE CARPETAS POR RANGOS (PARA PEDIDOS)
# ---------------------------------------------------

def obtener_ultimo_contador_y_ruta(carpeta_base, extension=".csv"):
    if not os.path.exists(carpeta_base):
        os.makedirs(carpeta_base, exist_ok=True)
    ultimo_contador = 4999
    subcarpetas = [d for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    subcarpetas_rango = []
    for d in subcarpetas:
        match = re.match(r"(\d+)-(\d+)", d)
        if match:
            subcarpetas_rango.append((int(match.group(1)), d))
    if subcarpetas_rango:
        subcarpetas_rango.sort()
        ultimo_rango_inicio, nombre_subcarpeta = subcarpetas_rango[-1]
        ruta_subcarpeta = os.path.join(carpeta_base, nombre_subcarpeta)
        regex_archivo = r"(\d+)" + re.escape(extension)
        for f in os.listdir(ruta_subcarpeta):
            m = re.match(regex_archivo, f)
            if m:
                ultimo_contador = max(ultimo_contador, int(m.group(1)))
    else:
        regex_archivo = r"(\d+)" + re.escape(extension)
        for f in os.listdir(carpeta_base):
            m = re.match(regex_archivo, f)
            if m:
                ultimo_contador = max(ultimo_contador, int(m.group(1)))
    return ultimo_contador

def obtener_ruta_para_contador(carpeta_base, contador):
    rango_inicio = (contador // 100) * 100
    rango_fin = rango_inicio + 99
    nombre_rango = f"{rango_inicio}-{rango_fin}"
    ruta_rango = os.path.join(carpeta_base, nombre_rango)
    if not os.path.exists(ruta_rango):
        os.makedirs(ruta_rango, exist_ok=True)
    return ruta_rango

# ---------------------------------------------------
# 1. LÓGICA DE FILTRO DE ROTACIÓN
# ---------------------------------------------------

def filtrar_por_disponibilidad_task(ruta_excel_quiter, ruta_salida_folder, color_resaltado="FFFF00", log_callback=None, progreso_callback=None, finalizado_callback=None):
    def log(texto, tag=None):
        if log_callback: log_callback(f"[ROTACION] {texto}", tag)
    log("Iniciando analisis de disponibilidad...")
    try:
        wb = load_workbook(ruta_excel_quiter)
        ws = wb.active
        fill_color = PatternFill(start_color=color_resaltado, end_color=color_resaltado, fill_type="solid")
        total_filas = max(ws.max_row - 1, 1)
        filas_resaltadas = 0
        for i in range(2, ws.max_row + 1):
            resultado_busqueda = ws.cell(row=i, column=18).value
            resultado_normalizado = "" if resultado_busqueda is None else " ".join(str(resultado_busqueda).strip().lower().split())
            if resultado_normalizado and resultado_normalizado not in ("pieza disponible", "disponible"):
                for cell in ws[i]:
                    cell.fill = fill_color
                filas_resaltadas += 1
            if progreso_callback: progreso_callback((i - 1) / total_filas)

        nombre_base = os.path.splitext(os.path.basename(ruta_excel_quiter))[0]
        nombre_archivo = f"{nombre_base}_ROTACION.xlsx"
        ruta_final = os.path.join(ruta_salida_folder, nombre_archivo)

        ws.delete_rows(1)
        ws.delete_cols(8, 10)

        wb.save(ruta_final)
        wb.close()
        log(f"Analisis completado. Filas resaltadas: {filas_resaltadas}. Guardado en: {ruta_final}", "success")
        if progreso_callback: progreso_callback(1)
    except Exception as e:
        log(f"Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        if finalizado_callback: finalizado_callback()

def filtrar_por_rotacion_task(ruta_excel_quiter, ruta_propuesta_xls, ruta_salida_folder, umbral_media, color_resaltado, log_callback=None, progreso_callback=None, finalizado_callback=None):
    def log(texto, tag=None):
        if log_callback: log_callback(f"[ROTACIÓN] {texto}", tag)
    log(f"📊 Iniciando análisis de rotación...")
    try:
        log("📖 Leyendo archivo de rotación...")
        try:
            df_propuesta = pd.read_excel(ruta_propuesta_xls, sheet_name='Informe')
        except Exception as e:
            log(f"❌ Error al leer el archivo Excel: {e}", "error")
            registrar_error_detallado(str(e), traceback.format_exc())
            return
        columnas = [str(c).strip() for c in df_propuesta.columns]
        col_articulo = next((c for c in columnas if c.lower() in ["articulo", "artículo", "referencia", "ref", "código", "codigo"]), None)
        col_media = next((c for c in columnas if c.lower() in ["media", "consumo", "promedio"]), None)
        if not col_articulo or not col_media:
            log(f"❌ Error: No se encuentran las columnas necesarias.", "error")
            return
        dict_medias = {}
        for _, fila in df_propuesta.iterrows():
            try:
                ref = str(fila[col_articulo]).strip().replace('2/', '')
                dict_medias[ref] = float(fila[col_media])
            except Exception:
                continue
        log(f"✅ Cargadas {len(dict_medias)} referencias.", "success")
        wb = load_workbook(ruta_excel_quiter)
        ws = wb.active
        fill_color = PatternFill(start_color=color_resaltado, end_color=color_resaltado, fill_type="solid")
        fill_no_disponible = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        total_filas = max(ws.max_row - 1, 1)
        filas_baja_rotacion = 0
        filas_no_disponibles = 0
        for i in range(2, ws.max_row + 1):
            ref = str(ws.cell(row=i, column=2).value).strip().replace('2/', '')
            resultado_busqueda = ws.cell(row=i, column=18).value
            resultado_normalizado = "" if resultado_busqueda is None else " ".join(str(resultado_busqueda).strip().lower().split())
            if dict_medias.get(ref, 0.0) < umbral_media:
                for cell in ws[i]: cell.fill = fill_color
                filas_baja_rotacion += 1
            elif resultado_normalizado and resultado_normalizado not in ("pieza disponible", "disponible"):
                for cell in ws[i]: cell.fill = fill_no_disponible
                filas_no_disponibles += 1
            if progreso_callback: progreso_callback((i - 1) / total_filas)
        
        # AJUSTE v13: Nombre original + rotacion
        nombre_base = os.path.splitext(os.path.basename(ruta_excel_quiter))[0]
        nombre_archivo = f"{nombre_base}_ROTACION.xlsx"
        ruta_final = os.path.join(ruta_salida_folder, nombre_archivo)

        ws.delete_rows(1)
        ws.delete_cols(8, 10)
        
        wb.save(ruta_final)
        wb.close()
        log(f"Resumen rotacion: rojas por umbral {filas_baja_rotacion}, amarillas por disponibilidad {filas_no_disponibles}.", "success")
        log(f"✅ Análisis completado. Guardado en: {ruta_final}", "success")
        if progreso_callback: progreso_callback(1)
    except Exception as e:
        log(f"❌ Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 2. LÓGICA DE COMPARADOR DE EXCEL
# ---------------------------------------------------

def comparar_excel_task(ruta_excel_antiguo, ruta_excel_nuevo, ruta_salida_excel, color_resaltado, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None):
    global comparador_running
    def log(texto, tag=None):
        if log_callback: log_callback(f"[COMPARACIÓN] {texto}", tag)
    log(f"🔍 Iniciando comparación...")
    comparador_running = True
    try:
        wb_antiguo = load_workbook(ruta_excel_antiguo, data_only=True)
        ws_antiguo = wb_antiguo.active
        lineas_antiguo = set(tuple(fila) for fila in ws_antiguo.iter_rows(values_only=True) if any(fila))
        wb_nuevo = load_workbook(ruta_excel_nuevo, data_only=True)
        ws_nuevo = wb_nuevo.active
        wb_salida = openpyxl.Workbook()
        ws_salida = wb_salida.active
        fill_color = PatternFill(start_color=color_resaltado, end_color=color_resaltado, fill_type="solid")
        total_lineas = ws_nuevo.max_row
        for i, row in enumerate(ws_nuevo.iter_rows(values_only=True)):
            if stop_event and stop_event.is_set(): break
            ws_salida.append(row)
            if tuple(row) in lineas_antiguo:
                for cell in ws_salida[ws_salida.max_row]: cell.fill = fill_color
            if progreso_callback: progreso_callback((i + 1) / total_lineas)
        if not (stop_event and stop_event.is_set()):
            Path(os.path.dirname(ruta_salida_excel)).mkdir(parents=True, exist_ok=True)
            wb_salida.save(ruta_salida_excel)
            log(f"✅ Resultado guardado en: {ruta_salida_excel}", "success")
    except Exception as e:
        log(f"❌ Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        comparador_running = False
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 3. LÓGICA DE AUTOMATIZACIÓN MULTI-HILO (WORKER)
# ---------------------------------------------------

def worker_playwright(worker_id, sub_filas, ruta_excel, usuario, password, precio_minimo, log_callback, update_global_progress, stop_event):
    def log(texto, tag=None):
        log_callback(f"[HILO {worker_id}] {texto}", tag)

    def esperar_y_hacer_click(locator, timeout=15000):
        locator.wait_for(state="visible", timeout=timeout)
        locator.click()

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--disable-gpu", "--disable-dev-shm-usage"])
            try:
                context = browser.new_context(http_credentials={"username": usuario, "password": password})
                context.route("**/*.{png,jpg,jpeg,gif,svg,woff,woff2}", lambda route: route.abort())
                page = context.new_page()
                url_acceso = "https://parts.citroen.com/zzlogpl0"
                try:
                    page.goto(url_acceso, timeout=45000)
                    page.wait_for_selector("text=/Espa.ol/", timeout=20000)
                    page.locator("text=/Espa.ol/").first.click()
                    page.wait_for_selector("text=/Referencia/", timeout=25000)
                except Exception as e:
                    log(f"Error de acceso/login: {e}", "error")
                    registrar_error_detallado(str(e), traceback.format_exc())
                    return
                target_frame = None
                for f in page.frames:
                    try:
                        if f.locator("text=/Referencia/").count() > 0:
                            target_frame = f
                            break
                    except Exception:
                        continue
                if not target_frame:
                    log("No se encontró el frame de trabajo.", "error")
                    return

                def preparar_formulario():
                    btn_referencia = target_frame.locator("text=Referencia").first
                    btn_disponibilidad = target_frame.locator("text=Mostrar disponibilidad pieza").first
                    esperar_y_hacer_click(btn_referencia, timeout=20000)
                    esperar_y_hacer_click(btn_disponibilidad, timeout=20000)
                    campo_ref = target_frame.locator('input[name="mara-matnr[1]"]').first
                    campo_ref.wait_for(state="visible", timeout=20000)
                    campo_ref.click()
                    campo_ref.fill("")
                    return campo_ref

                for fila_idx, referencia, valor_precio in sub_filas:
                    if stop_event.is_set():
                        break
                    try:
                        precio_num = float(str(valor_precio).replace(",", "."))
                        if precio_num < precio_minimo:
                            update_global_progress()
                            continue
                    except Exception:
                        update_global_progress()
                        continue
                    log(f"Consultando: {referencia}")
                    try:
                        ultimo_error = None
                        for intento in range(2):
                            try:
                                campo_ref = preparar_formulario()
                                campo_ref.fill(str(referencia))
                                radio = target_frame.locator('input[type="radio"]').nth(1)
                                radio.wait_for(state="attached", timeout=10000)
                                radio.click(force=True)
                                btn_ejecutar = target_frame.locator("text=Ejecutar programa").first
                                esperar_y_hacer_click(btn_ejecutar, timeout=15000)
                                break
                            except Exception as intento_error:
                                ultimo_error = intento_error
                                if intento == 0:
                                    log(f"Reintentando {referencia}: el formulario no estaba listo.", "error")
                                    try:
                                        btn_otro = target_frame.locator("text=Otro Articulo").first
                                        if btn_otro.is_visible(timeout=3000):
                                            btn_otro.click()
                                    except Exception:
                                        pass
                                    page.wait_for_timeout(1200)
                                    continue
                                raise ultimo_error
                        selector = "tr.moduloTableauBlanc td"
                        target_frame.locator(selector).first.wait_for(state="visible", timeout=15000)
                        resultado = target_frame.locator(selector).first.inner_text().strip()
                        with excel_lock:
                            wb = openpyxl.load_workbook(ruta_excel)
                            ws = wb.active
                            ws.cell(row=fila_idx, column=18).value = resultado
                            wb.save(ruta_excel)
                            wb.close()
                        log(f"OK {referencia}: {resultado}")
                        try:
                            btn_otro = target_frame.locator("text=Otro Articulo").first
                            esperar_y_hacer_click(btn_otro, timeout=10000)
                            page.wait_for_timeout(700)
                        except Exception:
                            pass
                    except Exception as e:
                        log(f"Error en {referencia}: {e}", "error")
                    update_global_progress()
            finally:
                browser.close()
    except Exception as e:
        log(f"Error critico en hilo: {e}", "error")

def ejecutar_automatizacion_task(ruta_excel, usuario, password, precio_minimo, num_hilos, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None):
    global automation_running
    automation_running = True
    if stop_event:
        stop_event.clear()
    try:
        wb = openpyxl.load_workbook(ruta_excel)
        sheet = wb.active
        filas_trabajo = []
        for r in range(2, sheet.max_row + 1):
            ref = sheet.cell(row=r, column=2).value
            pre = sheet.cell(row=r, column=6).value
            if ref:
                filas_trabajo.append((r, ref, pre))
        total_tareas = len(filas_trabajo)
        if total_tareas == 0:
            if finalizado_callback:
                finalizado_callback()
            return

        procesadas = 0
        cierre_notificado = False
        progreso_lock = threading.Lock()

        def update_progress():
            nonlocal procesadas, cierre_notificado
            global automation_running
            with progreso_lock:
                procesadas += 1
                progreso = procesadas / total_tareas
                deberia_cerrar = procesadas >= total_tareas and not cierre_notificado
                if deberia_cerrar:
                    cierre_notificado = True
            if progreso_callback:
                progreso_callback(progreso)
            if deberia_cerrar:
                automation_running = False
                if finalizado_callback:
                    finalizado_callback()

        chunk_size = (total_tareas // num_hilos) + 1
        chunks = [filas_trabajo[i:i + chunk_size] for i in range(0, total_tareas, chunk_size)]
        log_callback(f"[SISTEMA] Iniciando con {len(chunks)} hilos paralelos...")
        with ThreadPoolExecutor(max_workers=num_hilos) as executor:
            for i, chunk in enumerate(chunks):
                executor.submit(worker_playwright, i + 1, chunk, ruta_excel, usuario, password, precio_minimo, log_callback, update_progress, stop_event)
        if not stop_event.is_set():
            log_callback("Automatizacion MULTI-HILO completada con exito.", "success")
    except Exception as e:
        log_callback(f"Error general: {e}", "error")
    finally:
        automation_running = False
        if finalizado_callback and not cierre_notificado:
            finalizado_callback()

# ---------------------------------------------------
# 4. LÓGICA DE GENERACIÓN DE PEDIDOS (CON CARPETAS POR RANGO) - v13
# ---------------------------------------------------

def generar_pedidos_task(ruta_excel, carpeta_pedidos, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None, numero_callback=None):
    global pedidos_running
    pedidos_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[PEDIDOS] {texto}", tag) # AJUSTE v13: Prefijo [PEDIDOS]
    try:
        ruta_backup = crear_copia_seguridad(ruta_excel)
        log(f"Backup creado: {ruta_backup}", "success")
        # 1. LIMPIEZA DE REFERENCIAS (2/) Y GUARDADO INTERMEDIO
        wb = load_workbook(ruta_excel)
        ws = wb.active
        for fila in range(2, ws.max_row + 1):
            valor = ws[f"A{fila}"].value
            if valor:
                valor_str = str(valor)
                if "2/" in valor_str:
                    ws[f"A{fila}"] = valor_str.replace("2/", "")
        wb.save(ruta_excel)
        
        # 2. LECTURA DE DATOS FILTRADOS (SOLO COLUMNAS A, C, F)
        wb_data = load_workbook(ruta_excel, data_only=True)
        ws_data = wb_data.active
        
        datos = []
        for row in ws_data.iter_rows(min_row=2, values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                ref = row[0] if len(row) > 0 else ""
                cant = row[2] if len(row) > 2 else ""
                cli = row[5] if len(row) > 5 else ""
                datos.append([ref, cant, cli])
        
        if not datos:
            log("⚠️ El Excel está vacío o no tiene datos válidos.", "error")
            return

        # 3. GESTIÓN DE CARPETAS Y GENERACIÓN DE CSV
        contador = obtener_ultimo_contador_y_ruta(carpeta_pedidos, ".csv")
        
        for i in range(0, len(datos), 100):
            if stop_event and stop_event.is_set(): break
            bloque = datos[i:i+100]
            
            contador += 1
            ruta_rango = obtener_ruta_para_contador(carpeta_pedidos, contador)
            ruta_csv = os.path.join(ruta_rango, f"{contador}.csv")
            
            with open(ruta_csv, "w", newline="", encoding="utf-8") as f:
                csv.writer(f, delimiter=";").writerows(bloque)
            
            log(f"✅ Pedido generado: {contador}.csv (en {os.path.basename(ruta_rango)})", "success")
            if numero_callback: numero_callback(str(contador))
            if progreso_callback: progreso_callback((i + len(bloque)) / len(datos))
        
        # 4. BORRADO FÍSICO FINAL DEL EXCEL
        if not (stop_event and stop_event.is_set()):
            wb_clean = load_workbook(ruta_excel)
            ws_clean = wb_clean.active
            if ws_clean.max_row > 0:
                ws_clean.delete_rows(1, ws_clean.max_row)
            wb_clean.save(ruta_excel)
            log("✅ Pedidos generados (Ref, Cant, Cliente) y Excel limpiado.", "success")
            
    except Exception as e:
        log(f"❌ Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        pedidos_running = False
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 4bis. LÓGICA DE HOJA DE PEDIDOS DIARIA (COPIA EN LIMPIO DE LA PLANTILLA)
# ---------------------------------------------------

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

def ruta_hoja_pedidos_del_dia(carpeta_hojas, fecha=None):
    """Ruta AÑO/MES/DD-MM-AAAA.xlsx que le corresponde a una fecha dentro de la carpeta de hojas."""
    fecha = fecha or datetime.now()
    carpeta_mes = Path(carpeta_hojas) / str(fecha.year) / MESES_ES[fecha.month]
    nombre = fecha.strftime("%d-%m-%Y") + ".xlsx"
    return carpeta_mes / nombre

def generar_hoja_pedidos_task(ruta_plantilla, carpeta_hojas, log_callback=None, progreso_callback=None, finalizado_callback=None, forzar=False, fecha=None):
    global hoja_pedidos_running
    hoja_pedidos_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[HOJA PEDIDOS] {texto}", tag)
    log("Generando hoja de pedidos del día a partir de la plantilla...")
    try:
        if not ruta_plantilla or not os.path.exists(ruta_plantilla):
            log("❌ No se encuentra la plantilla de hoja de pedidos.", "error")
            return
        ruta_final = ruta_hoja_pedidos_del_dia(carpeta_hojas, fecha)
        if progreso_callback: progreso_callback(0.2)
        if ruta_final.exists() and not forzar:
            log(f"La hoja de hoy ya existía, no se ha sobrescrito: {ruta_final}", "success")
            if progreso_callback: progreso_callback(1)
            return
        ruta_final.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ruta_plantilla, ruta_final)
        if progreso_callback: progreso_callback(1)
        log(f"✅ Hoja de pedidos generada: {ruta_final}", "success")
    except Exception as e:
        log(f"❌ Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc(), ref="generar_hoja_pedidos_task")
    finally:
        hoja_pedidos_running = False
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 4bis-2. AVISOS A CLIENTES DE REFERENCIAS QUE NO SUMINISTRAN
# ---------------------------------------------------

def _normalizar_texto_estado(valor):
    return " ".join(str(valor).strip().upper().split())

def _normalizar_cliente_dms(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()

def cargar_clientes_excel(ruta_clientes):
    """Lee clientes.xlsx (A=num_cliente, B=nombre, C=email, varios separados por ';') -> {num_cliente: {"nombre", "email"}}."""
    clientes = {}
    wb = openpyxl.load_workbook(ruta_clientes, data_only=True)
    ws = wb.active
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not fila or fila[0] is None:
            continue
        num = _normalizar_cliente_dms(fila[0])
        if not num:
            continue
        nombre = str(fila[1]).strip() if len(fila) > 1 and fila[1] not in (None, "") else ""
        email = str(fila[2]).strip() if len(fila) > 2 and fila[2] not in (None, "") else ""
        clientes[num] = {"nombre": nombre, "email": email}
    wb.close()
    return clientes

def extraer_no_suministradas(ruta_hoja, log=lambda *a, **k: None):
    """Recorre todas las pestañas de la hoja de pedido y agrupa por cliente DMS las referencias con ESTADO 'NO SUMINISTRA'."""
    wb = openpyxl.load_workbook(ruta_hoja, data_only=True)
    agrupado = {}
    for ws in wb.worksheets:
        pedido = ws.cell(row=2, column=5).value
        encontradas = 0
        for fila in range(5, ws.max_row + 1):
            referencia = ws.cell(row=fila, column=2).value
            estado = ws.cell(row=fila, column=5).value
            cliente = ws.cell(row=fila, column=8).value
            if not referencia or not estado:
                continue
            if "NO SUMINISTRA" not in _normalizar_texto_estado(estado):
                continue
            cliente_norm = _normalizar_cliente_dms(cliente)
            if not cliente_norm:
                continue
            agrupado.setdefault(cliente_norm, []).append(str(referencia).strip())
            encontradas += 1
        if encontradas:
            log(f"Pestaña '{ws.title}' (pedido {pedido}): {encontradas} referencia(s) NO SUMINISTRA.")
    wb.close()
    return agrupado

def construir_cuerpo_aviso(referencias, fecha_str):
    if len(referencias) == 1:
        return (
            "Buenas tardes.\n\n"
            f"La referencia {referencias[0]} correspondiente al pedido del día {fecha_str} no suministra.\n\n"
            "Gracias por su atención,"
        )
    lineas = "\n".join(f"- {ref}" for ref in referencias)
    return (
        "Buenas tardes.\n\n"
        f"Las siguientes referencias correspondientes al pedido del día {fecha_str} no suministran:\n\n"
        f"{lineas}\n\n"
        "Gracias por su atención,"
    )

def crear_borrador_outlook(outlook, destinatarios, asunto, cuerpo_texto):
    """Crea un borrador en Outlook de escritorio (no lo envía) respetando la firma por defecto del usuario.
    'outlook' es la aplicación Outlook ya conectada (COM ya inicializado en el hilo llamante)."""
    mail = outlook.CreateItem(0)  # olMailItem
    mail.Subject = asunto
    if destinatarios:
        mail.To = destinatarios
    mail.GetInspector  # fuerza a Outlook a cargar la firma por defecto de "mensaje nuevo"
    firma_html = mail.HTMLBody or ""
    cuerpo_html = cuerpo_texto.replace("\n", "<br>")
    mail.HTMLBody = f"<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\">{cuerpo_html}</div><br>{firma_html}"
    mail.Save()

def generar_avisos_no_suministro_task(ruta_hoja, ruta_clientes, log_callback=None, progreso_callback=None, finalizado_callback=None, resultado_callback=None):
    global avisos_running
    avisos_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[AVISOS] {texto}", tag)
    log("Analizando la hoja de pedido en busca de referencias que no suministran...")
    resultado = {"generados": [], "sin_correo": [], "error": None}
    try:
        match = re.search(r"\d{2}-\d{2}-\d{4}", Path(ruta_hoja).stem)
        fecha_str = match.group(0) if match else Path(ruta_hoja).stem
        try:
            clientes = cargar_clientes_excel(ruta_clientes)
        except Exception as e:
            log(f"❌ No se pudo leer el Excel de clientes: {e}", "error")
            registrar_error_detallado(str(e), traceback.format_exc(), ref="generar_avisos_no_suministro_task_clientes")
            resultado["error"] = str(e)
            return
        try:
            agrupado = extraer_no_suministradas(ruta_hoja, log)
        except Exception as e:
            log(f"❌ No se pudo leer la hoja de pedido: {e}", "error")
            registrar_error_detallado(str(e), traceback.format_exc(), ref="generar_avisos_no_suministro_task_hoja")
            resultado["error"] = str(e)
            return
        if not agrupado:
            log("No se han encontrado referencias con estado NO SUMINISTRA.", "success")
            return
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as e:
                log(f"❌ No se pudo conectar con Outlook: {e}", "error")
                registrar_error_detallado(str(e), traceback.format_exc(), ref="generar_avisos_no_suministro_task_outlook")
                resultado["error"] = str(e)
                return
            total = len(agrupado)
            hecho = 0
            for cliente_num, referencias in agrupado.items():
                info = clientes.get(cliente_num)
                nombre = info["nombre"] if info else ""
                email = info["email"] if info else ""
                asunto = f"Aviso referencia no suministrada {fecha_str}"
                cuerpo = construir_cuerpo_aviso(referencias, fecha_str)
                try:
                    crear_borrador_outlook(outlook, email, asunto, cuerpo)
                except Exception as e:
                    log(f"❌ Error creando el borrador para el cliente {cliente_num}: {e}", "error")
                    registrar_error_detallado(str(e), traceback.format_exc(), ref=f"generar_avisos_no_suministro_task_{cliente_num}")
                    hecho += 1
                    if progreso_callback: progreso_callback(hecho / total)
                    continue
                entrada = {"cliente": cliente_num, "nombre": nombre, "email": email, "referencias": referencias}
                if email:
                    resultado["generados"].append(entrada)
                    log(f"✅ Borrador creado para {nombre or cliente_num} ({len(referencias)} ref.)", "success")
                else:
                    resultado["sin_correo"].append(entrada)
                    log(f"⚠️ Cliente {cliente_num} sin email en el Excel de clientes; borrador creado sin destinatario.", "error")
                hecho += 1
                if progreso_callback: progreso_callback(hecho / total)
            log(f"✨ Proceso completado: {len(resultado['generados'])} borrador(es) con destinatario, {len(resultado['sin_correo'])} sin email.", "success")
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        resultado["error"] = str(e)
        log(f"❌ Error: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc(), ref="generar_avisos_no_suministro_task")
    finally:
        avisos_running = False
        if resultado_callback: resultado_callback(resultado)
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 5. LÓGICA DE FICHERO MOPAR (CON NOMBRES FXXXX + LETRAS) - v13
# ---------------------------------------------------

def generar_fichero_mopar_task(ruta_excel, carpeta_salida, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None, numero_callback=None):
    global mopar_running
    mopar_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[MOPAR] {texto}", tag)
    log(f"🚀 Iniciando generación de ficheros MOPAR...")
    try:
        log("📖 Leyendo archivo Excel...")
        try:
            wb = xlrd.open_workbook(ruta_excel)
            ws = wb.sheet_by_index(0)
        except Exception as e:
            log(f"❌ Error al abrir el Excel: {e}", "error")
            return
        headers = [str(ws.cell_value(0, i)).strip() for i in range(ws.ncols)]
        try:
            col_ref = headers.index("Referencia")
            col_cant = headers.index("Cant")
            col_cliente = headers.index("Cliente")
        except ValueError as e:
            col_faltante = str(e).split()[-1]
            log(f"❌ Error: No se encuentra la columna {col_faltante} en el Excel.", "error")
            return
        total_filas = ws.nrows - 1
        if total_filas <= 0:
            log("⚠️ El Excel no tiene datos de artículos.", "error")
            return
        lineas_fichero = []
        log(f"🔍 Procesando {total_filas} filas...")
        for r in range(1, ws.nrows):
            if stop_event and stop_event.is_set(): break
            try:
                ref_raw = str(ws.cell_value(r, col_ref)).strip()
                ref_clean = ref_raw.replace("2/", "")
                referencia = ref_clean.zfill(13)[:13]
                try:
                    cant_val = int(float(ws.cell_value(r, col_cant)))
                except Exception:
                    cant_val = 0
                    log(f"⚠️ Cantidad no válida en fila {r+1}, se usará 0.", "error")
                cantidad = str(cant_val).zfill(7)[:7]
                cliente_raw = str(ws.cell_value(r, col_cliente)).strip()
                cliente_nums = "".join(filter(str.isdigit, cliente_raw))
                ubicacion = cliente_nums.zfill(8)[:8]
                linea = f"{referencia}{cantidad}{ubicacion}"
                lineas_fichero.append(linea)
            except Exception as e:
                log(f"⚠️ Error en fila {r+1}: {e}", "error")
            if progreso_callback: progreso_callback(r / total_filas)
        
        if not (stop_event and stop_event.is_set()):
            # AJUSTE v13: Lógica FXXXX + Letras
            if not os.path.exists(carpeta_salida):
                os.makedirs(carpeta_salida, exist_ok=True)

            proximo_f = obtener_numero_mopar_del_dia(carpeta_salida)
            letras = string.ascii_uppercase # A, B, C...
            archivos_creados = 0
            
            for i in range(0, len(lineas_fichero), 10):
                grupo = lineas_fichero[i:i+10]
                # Si solo hay un archivo, no lleva letra. Si hay varios, llevan A, B, C...
                if len(lineas_fichero) <= 10:
                    nombre_fichero = f"F{proximo_f}.txt"
                else:
                    indice_letra = i // 10
                    letra = letras[indice_letra % len(letras)]
                    nombre_fichero = f"F{proximo_f}{letra}.txt"
                
                ruta_final = os.path.join(carpeta_salida, nombre_fichero)
                with open(ruta_final, "w", encoding="utf-8") as f:
                    f.write("\n".join(grupo))
                
                log(f"✅ Generado MOPAR: {nombre_fichero}", "success")
                if numero_callback: numero_callback(nombre_fichero)
                archivos_creados += 1

            log(f"✨ Proceso completado. Se han generado {archivos_creados} archivos con prefijo F{proximo_f}.", "success")
    except Exception as e:
        log(f"❌ Error crítico: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        mopar_running = False
        if finalizado_callback: finalizado_callback()

# ---------------------------------------------------
# 5bis. LÓGICA DE FACTURAS STELLANTIS MADD (DESCARGA DIARIA)
# ---------------------------------------------------
# Estructura de la web confirmada por el usuario el 2026-09-21 (login SSO en
# idfed.mpsa.com, selector de perfil ng-select, tabla de documentos "eDoc-Partners").
# Solo lectura y descarga: nunca se pulsan los botones de estado (leído/impreso/etc).

MADD_URL_DOCUMENTOS = "https://www.madd.stellantis.com/documents"
MADD_PERFIL_TEXTO = "Piezas de recambio"
LOG_FACTURAS_MADD = "log_facturas_madd.txt"
MADD_PAGINAS_MAX = 10  # 100 docs/página -> hasta 1000 documentos más recientes revisados

def _dia_laborable_anterior(fecha):
    """Día laborable anterior a 'fecha', saltando sábados y domingos (los lunes devuelve el viernes)."""
    anterior = fecha - timedelta(days=1)
    while anterior.weekday() >= 5:  # 5=sábado, 6=domingo
        anterior -= timedelta(days=1)
    return anterior

def registrar_log_facturas_madd(fecha_objetivo, resultado):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FACTURAS_MADD, "a", encoding="utf-8") as f:
        f.write(f"--- EJECUCIÓN [{timestamp}] ---\n")
        f.write(f"Fecha objetivo: {fecha_objetivo.isoformat()}\n")
        f.write(f"Nº facturas descargadas: {len(resultado.get('facturas', []))}\n")
        if resultado.get("error"):
            f.write(f"Error: {resultado['error']} (paso: {resultado.get('paso_fallido') or 'desconocido'})\n")
        else:
            f.write("Sin errores.\n")
        f.write("-" * 30 + "\n\n")

def _madd_login(page, usuario, password):
    page.goto(MADD_URL_DOCUMENTOS, timeout=45000)
    page.wait_for_selector("#username", timeout=30000)
    page.fill("#username", usuario)
    page.fill("#password", password)
    page.click("#signOnButton")
    page.wait_for_load_state("networkidle", timeout=30000)

def _madd_seleccionar_perfil(page):
    selector_perfil = page.locator('ng-select[formcontrolname="profile"]')
    # Si ya hay un perfil activo de una sesión previa, esta pantalla no aparece.
    if selector_perfil.count() == 0 or not selector_perfil.first.is_visible():
        return
    selector_perfil.first.click()
    opcion = page.locator(".ng-dropdown-panel .ng-option", has_text=MADD_PERFIL_TEXTO).first
    opcion.wait_for(state="visible", timeout=10000)
    opcion.click()
    boton_confirmar = page.locator('button:has(i.bi-check-circle):not([disabled])').first
    boton_confirmar.wait_for(state="visible", timeout=10000)
    boton_confirmar.click()
    page.wait_for_selector("tr.border-bottom", timeout=20000)

def _madd_procesar_facturas_del_dia(page, fechas_objetivo, carpeta_destino, log):
    """Recorre las páginas de resultados, filtra solo por los títulos reales 'Factura ...' (ignora
    cualquier otro elemento/fila duplicada del DOM) y descarga cada factura en el momento en que se
    encuentra, para no depender de índices que quedarían obsoletos tras cambiar de página."""
    try:
        page.select_option('select[name="pageSize"]', label="100 elementos por página")
        page.wait_for_timeout(1000)
    except Exception:
        log("No se pudo ampliar el tamaño de página, se sigue con el valor por defecto.", "error")
    patron = re.compile(r"Factura\s+(\S+)\s+(\d{2}/\d{2}/\d{4})")
    facturas_descargadas = []
    for pagina in range(1, MADD_PAGINAS_MAX + 1):
        titulos = page.locator("a").filter(has_text=re.compile(r"^Factura\s"))
        total_titulos = titulos.count()
        log(f"Página {pagina}/{MADD_PAGINAS_MAX}: {total_titulos} factura(s) en la página, revisando...")
        coincidencias_pagina = 0
        for i in range(total_titulos):
            enlace = titulos.nth(i)
            try:
                texto_titulo = enlace.inner_text(timeout=3000).strip()
            except Exception:
                continue
            match = patron.search(texto_titulo)
            if not match:
                continue
            numero, fecha_str = match.group(1), match.group(2)
            dia, mes, anio = fecha_str.split("/")
            fecha_factura = datetime(int(anio), int(mes), int(dia)).date()
            if fecha_factura not in fechas_objetivo:
                continue
            coincidencias_pagina += 1
            nombre_archivo = f"Factura_{numero}_{fecha_factura.isoformat()}.pdf"
            ruta_destino = Path(carpeta_destino) / nombre_archivo
            if ruta_destino.exists():
                facturas_descargadas.append(str(ruta_destino))
                log(f"Ya existía {nombre_archivo}, no se ha vuelto a descargar.", "success")
                continue
            try:
                fila = enlace.locator("xpath=ancestor::tr[1]")
                boton_descarga = fila.locator("td.text-center button").first
                with page.expect_download(timeout=30000) as info_descarga:
                    boton_descarga.click()
                descarga = info_descarga.value
                descarga.save_as(str(ruta_destino))
                facturas_descargadas.append(str(ruta_destino))
                log(f"✅ Descargada: {nombre_archivo}", "success")
            except Exception as e:
                log(f"❌ Error descargando {nombre_archivo}: {e}", "error")
                raise
        log(f"Página {pagina}/{MADD_PAGINAS_MAX} revisada. {coincidencias_pagina} coincidencia(s) en esta página.")
        if pagina >= MADD_PAGINAS_MAX:
            break
        boton_siguiente = page.locator('a[aria-label="Next"]')
        if boton_siguiente.count() == 0:
            log("No hay botón 'Next', se detiene la paginación.")
            break
        li_padre = boton_siguiente.first.locator("xpath=..")
        if "disabled" in (li_padre.get_attribute("class") or ""):
            log("Botón 'Next' deshabilitado (última página), se detiene la paginación.")
            break
        log(f"Pasando a la página {pagina + 1}...")
        boton_siguiente.first.click()
        page.wait_for_timeout(1500)
    return facturas_descargadas

def ejecutar_descarga_facturas_madd(usuario, password, carpeta_destino, fechas_objetivo, log):
    Path(carpeta_destino).mkdir(parents=True, exist_ok=True)
    resultado = {"facturas": [], "error": None, "paso_fallido": None}
    ultimo_error = None
    for intento in range(3):
        resultado["paso_fallido"] = None
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True, args=["--disable-gpu", "--disable-dev-shm-usage"])
                try:
                    context = browser.new_context(accept_downloads=True)
                    page = context.new_page()
                    try:
                        _madd_login(page, usuario, password)
                    except Exception:
                        resultado["paso_fallido"] = "acceso/login"
                        raise
                    try:
                        _madd_seleccionar_perfil(page)
                    except Exception:
                        resultado["paso_fallido"] = "cambio de perfil"
                        raise
                    try:
                        resultado["facturas"] = _madd_procesar_facturas_del_dia(page, fechas_objetivo, carpeta_destino, log)
                    except Exception:
                        resultado["paso_fallido"] = resultado["paso_fallido"] or "listado/descarga de documentos"
                        raise
                    return resultado
                finally:
                    browser.close()
        except Exception as e:
            ultimo_error = e
            log(f"⚠️ Intento {intento + 1}/3 fallido en '{resultado['paso_fallido']}': {e}", "error")
            registrar_error_detallado(str(e), traceback.format_exc(), ref="facturas_madd")
            if intento < 2:
                time.sleep(5)
    resultado["error"] = str(ultimo_error)
    return resultado

def enviar_correo_outlook_madd(outlook, destinatarios, asunto, cuerpo_texto, adjuntos=None):
    """Crea y ENVÍA (Send, no Save) un correo desde Outlook de escritorio, respetando la firma por defecto."""
    mail = outlook.CreateItem(0)  # olMailItem
    mail.Subject = asunto
    if destinatarios:
        mail.To = destinatarios
    mail.GetInspector  # fuerza a Outlook a cargar la firma por defecto de "mensaje nuevo"
    firma_html = mail.HTMLBody or ""
    cuerpo_html = cuerpo_texto.replace("\n", "<br>")
    mail.HTMLBody = f"<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\">{cuerpo_html}</div><br>{firma_html}"
    if adjuntos:
        for ruta in adjuntos:
            mail.Attachments.Add(str(ruta))
    mail.Send()

def _madd_enviar_correo_resultado(email_destino, fecha_objetivo, rutas_facturas, log):
    fecha_str = fecha_objetivo.strftime("%Y-%m-%d")
    asunto = f"Facturas Stellantis del {fecha_str}"
    if rutas_facturas:
        cuerpo = f"Adjunto las facturas del día {fecha_str}. Total: {len(rutas_facturas)} facturas."
    else:
        cuerpo = f"No hay facturas del día {fecha_str}"
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        enviar_correo_outlook_madd(outlook, email_destino, asunto, cuerpo, adjuntos=rutas_facturas or None)
        log(f"✅ Correo enviado a {email_destino} ({len(rutas_facturas)} factura(s) adjuntas).", "success")
    finally:
        pythoncom.CoUninitialize()

def _madd_enviar_aviso_fallo(email_aviso, fecha_objetivo, resultado, log):
    if not email_aviso:
        log("No hay correo de aviso de fallo configurado; no se envía ningún aviso.", "error")
        return
    fecha_str = fecha_objetivo.strftime("%Y-%m-%d")
    paso = resultado.get("paso_fallido") or "desconocido"
    error = resultado.get("error") or "error desconocido"
    asunto = f"[AVISO] Fallo en descarga de facturas Stellantis del {fecha_str}"
    cuerpo = f"La descarga automática de facturas Stellantis del día {fecha_str} ha fallado tras 3 intentos.\n\nPaso que falló: {paso}\nError: {error}"
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        enviar_correo_outlook_madd(outlook, email_aviso, asunto, cuerpo)
        log(f"✅ Aviso de fallo enviado a {email_aviso}.", "error")
    finally:
        pythoncom.CoUninitialize()

def tarea_facturas_madd(usuario, password, carpeta_destino, email_destino, email_aviso_fallo, filtro_dia="Ayer", log_callback=None, finalizado_callback=None, resultado_callback=None):
    global madd_running
    madd_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[FACTURAS MADD] {texto}", tag)
    log("Iniciando descarga de facturas Stellantis MADD...")
    hoy = datetime.now().date()
    fecha_objetivo = hoy if filtro_dia == "Hoy" else _dia_laborable_anterior(hoy)
    resultado = {"facturas": [], "error": None, "paso_fallido": None}
    try:
        resultado = ejecutar_descarga_facturas_madd(usuario, password, carpeta_destino, {fecha_objetivo}, log)
        registrar_log_facturas_madd(fecha_objetivo, resultado)
        if resultado.get("error"):
            log(f"❌ Fallo tras 3 intentos en el paso '{resultado.get('paso_fallido')}': {resultado['error']}", "error")
            _madd_enviar_aviso_fallo(email_aviso_fallo, fecha_objetivo, resultado, log)
            return
        _madd_enviar_correo_resultado(email_destino, fecha_objetivo, resultado["facturas"], log)
        log(f"✨ Proceso completado: {len(resultado['facturas'])} factura(s) del {fecha_objetivo.isoformat()}.", "success")
    except Exception as e:
        resultado["error"] = str(e)
        resultado["paso_fallido"] = resultado.get("paso_fallido") or "envío de correo"
        log(f"❌ Error crítico: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc(), ref="tarea_facturas_madd")
        registrar_log_facturas_madd(fecha_objetivo, resultado)
        try:
            _madd_enviar_aviso_fallo(email_aviso_fallo, fecha_objetivo, resultado, log)
        except Exception as e2:
            log(f"❌ No se pudo enviar el aviso de fallo: {e2}", "error")
    finally:
        madd_running = False
        if resultado_callback: resultado_callback(resultado)
        if finalizado_callback: finalizado_callback()

def ejecutar_tarea_madd_standalone():
    """Punto de entrada para la tarea programada de Windows (--tarea-madd), sin GUI."""
    config = cargar_configuracion()
    usuario = config.get("usuario_madd", "")
    password = config.get("password_madd", "")
    carpeta = config.get("ruta_carpeta_facturas_madd", "")
    email_destino = config.get("email_facturas_madd", "")
    email_aviso = config.get("email_aviso_fallo_madd", "")
    filtro_dia = config.get("filtro_dia_facturas_madd", "Ayer")
    if not usuario or not password or not carpeta or not email_destino:
        registrar_error_detallado("Configuración incompleta para la tarea MADD (usuario/password/carpeta/email).", "", ref="madd_standalone")
        return
    tarea_facturas_madd(usuario, password, carpeta, email_destino, email_aviso, filtro_dia=filtro_dia, log_callback=None)

# ---------------------------------------------------
# 6. INTERFAZ GRÁFICA
# ---------------------------------------------------

def generar_fichero_stock_fiat_task(ruta_excel, carpeta_salida, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None):
    global stock_fiat_running
    stock_fiat_running = True
    def log(texto, tag=None):
        if log_callback: log_callback(f"[STOCK FIAT] {texto}", tag)
    log("Iniciando generacion de ficheros Stock FIAT...")
    try:
        log("Leyendo archivo Excel...")
        try:
            wb = xlrd.open_workbook(ruta_excel)
            ws = wb.sheet_by_index(0)
        except Exception as e:
            log(f"Error al abrir el Excel: {e}", "error")
            return
        if ws.ncols < 4:
            log("Error: El Excel debe tener al menos columnas hasta la D.", "error")
            return

        col_ref = 1
        col_cant = 3
        total_filas = ws.nrows
        if total_filas <= 0:
            log("El Excel no tiene datos de articulos.", "error")
            return

        def valor_limpio(valor):
            if isinstance(valor, float) and valor.is_integer():
                return str(int(valor))
            return str(valor).strip()

        lineas_fichero = []
        log(f"Procesando {total_filas} filas...")
        for fila_idx in range(ws.nrows):
            if stop_event and stop_event.is_set(): break
            try:
                ref_raw = valor_limpio(ws.cell_value(fila_idx, col_ref))
                if not ref_raw:
                    continue
                ref_clean = ref_raw.replace("2/", "")
                referencia = ref_clean.zfill(13)[:13]
                try:
                    cant_val = int(float(ws.cell_value(fila_idx, col_cant)))
                except Exception:
                    cant_val = 0
                    log(f"Cantidad no válida en fila {fila_idx+1}, se usará 0.", "error")
                cantidad = str(cant_val).zfill(7)[:7]
                ubicacion = "00000000"
                linea = f"{referencia}{cantidad}{ubicacion}"
                lineas_fichero.append(linea)
            except Exception as e:
                log(f"Error en fila {fila_idx+1}: {e}", "error")
            if progreso_callback: progreso_callback((fila_idx + 1) / total_filas)

        if not (stop_event and stop_event.is_set()):
            if not os.path.exists(carpeta_salida):
                os.makedirs(carpeta_salida, exist_ok=True)

            proximo_f = obtener_numero_mopar_del_dia(carpeta_salida, prefijo_clave="stock_fiat")
            letras = string.ascii_uppercase
            archivos_creados = 0

            for i in range(0, len(lineas_fichero), 10):
                grupo = lineas_fichero[i:i+10]
                if len(lineas_fichero) <= 10:
                    nombre_fichero = f"F{proximo_f}.txt"
                else:
                    indice_letra = i // 10
                    letra = letras[indice_letra % len(letras)]
                    nombre_fichero = f"F{proximo_f}{letra}.txt"

                ruta_final = os.path.join(carpeta_salida, nombre_fichero)
                with open(ruta_final, "w", encoding="utf-8") as f:
                    f.write("\n".join(grupo))

                log(f"Generado Stock FIAT: {nombre_fichero}", "success")
                archivos_creados += 1

            log(f"Proceso completado. Se han generado {archivos_creados} archivos con prefijo F{proximo_f}.", "success")
    except Exception as e:
        log(f"Error critico: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        stock_fiat_running = False
        if finalizado_callback: finalizado_callback()

class AppIntegrada(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.config = cargar_configuracion()
        self.aplicar_tema_inicial()
        self.title("Palausa Aprovisionamiento")
        self.geometry("1000x980")
        self.minsize(1000, 760)
        self.resizable(True, True)
        if os.path.exists(ICONO_PATH):
            try: self.iconbitmap(ICONO_PATH)
            except Exception: pass
        self.setup_ui()
        self.comprobar_hoja_pedidos_al_inicio()

    def aplicar_tema_inicial(self):
        # Modo y tema fijos: interfaz clara y consistente en toda la aplicación.
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")
        self.accent_color = None
        self.hover_color = None

    def aplicar_estilos_ui(self):
        # Paleta única, neutra y profesional (sustituye al antiguo sistema de temas/regiones).
        self.color_principal = "#2563eb"       # Acento primario (acciones principales)
        self.color_hover = "#1d4ed8"
        self.color_superficie = "#ffffff"      # Fondo de tarjetas/paneles
        self.color_superficie_alt = "#f1f4f9"  # Fondo de sub-secciones dentro de una tarjeta
        self.color_borde = "#dbe1ea"
        self.color_texto = "#1e2733"
        self.color_texto_sec = "#64748b"
        self.color_panel_oscuro = "#eef2f8"    # Secciones destacadas de fondo claro
        self.color_panel_medio = "#eef1f6"     # Fondo de botones secundarios
        self.color_panel_medio_hover = "#dfe4ec"
        self.color_panel_claro = "#f5f6f9"     # Fondo general de la ventana
        self.color_acento_azul = "#0ea5e9"
        self.color_ok = "#16a34a"
        self.color_warn = "#d97706"
        self.color_peligro = "#dc2626"
        self.color_peligro_hover = "#b91c1c"
        self.color_pokedex_luz = "#0f172a"     # Texto de énfasis (estado, porcentaje)

    def crear_tarjeta_tab(self, tab, titulo, subtitulo):
        card = ctk.CTkFrame(tab, fg_color=self.color_superficie, corner_radius=12, border_width=1, border_color=self.color_borde)
        card.pack(padx=20, pady=16, fill="both", expand=True)
        ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 17, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(22, 2))
        ctk.CTkLabel(card, text=subtitulo, font=("Segoe UI", 12), text_color=self.color_texto_sec).pack(anchor="w", padx=24, pady=(0, 14))
        ctk.CTkFrame(card, fg_color=self.color_borde, height=1).pack(fill="x", padx=24, pady=(0, 16))
        return card

    def crear_fila_archivo(self, parent, label_text, entry_attr, config_key, placeholder, button_text, command):
        ctk.CTkLabel(parent, text=label_text, font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(padx=24, pady=(0, 16), fill="x")
        entry = ctk.CTkEntry(row, placeholder_text=placeholder, height=38, corner_radius=8, border_width=1, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        entry.insert(0, self.config.get(config_key, ""))
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        setattr(self, entry_attr, entry)
        btn = ctk.CTkButton(row, text=button_text, width=92, height=38, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, border_width=1, border_color=self.color_borde, command=command)
        btn.pack(side="right")
        return entry

    def crear_boton_principal(self, parent, text, command):
        return ctk.CTkButton(parent, text=text, height=44, corner_radius=8, font=("Segoe UI", 13, "bold"), fg_color=self.color_principal, text_color="#ffffff", hover_color=self.color_hover, command=command)

    def setup_ui(self):
        self.aplicar_estilos_ui()
        acc_color = self.color_principal
        self.ultima_ruta_generada = ""
        self.configure(fg_color=self.color_panel_claro)

        cabecera = ctk.CTkFrame(self, fg_color=self.color_superficie, corner_radius=0, border_width=0)
        cabecera.pack(fill="x")
        cabecera_int = ctk.CTkFrame(cabecera, fg_color="transparent")
        cabecera_int.pack(fill="x", padx=24, pady=16)
        ctk.CTkLabel(cabecera_int, text="Palausa Aprovisionamiento", font=("Segoe UI", 20, "bold"), text_color=self.color_texto).pack(anchor="w")
        ctk.CTkLabel(cabecera_int, text="Herramienta interna de gestión de stock y pedidos", font=("Segoe UI", 12), text_color=self.color_texto_sec).pack(anchor="w", pady=(2, 0))
        ctk.CTkFrame(self, fg_color=self.color_borde, height=1).pack(fill="x")

        self.tabview = ctk.CTkTabview(
            self,
            height=460,
            corner_radius=10,
            segmented_button_selected_color=acc_color,
            segmented_button_selected_hover_color=self.color_hover,
            segmented_button_fg_color=self.color_panel_medio,
            segmented_button_unselected_color=self.color_panel_medio,
            text_color=self.color_texto,
            fg_color=self.color_panel_claro,
            border_width=0,
        )
        self.tabview.pack(pady=(16, 8), padx=22, fill="x")
        self.tab_auto = self.tabview.add("Automatización")
        self.tab_pedidos = self.tabview.add("Pedidos")
        self.tab_comparador = self.tabview.add("Comparador")
        self.tab_mopar = self.tabview.add("Fichero Mopar")
        self.tab_stock_fiat = self.tabview.add("Stock FIAT")
        self.tab_madd = self.tabview.add("Facturas MADD")
        self.tab_reclamaciones_power_line = self.tabview.add("Reclamaciones")
        self.tab_config = self.tabview.add("Configuración")
        self.build_tab_auto(acc_color)
        self.build_tab_pedidos(acc_color)
        self.build_tab_comparador(acc_color)
        self.build_tab_mopar(acc_color)
        self.build_tab_stock_fiat(acc_color)
        self.build_tab_madd(acc_color)
        self.reclamaciones_power_line = ModuloReclamacionesPowerLine(self, self.config)
        self.reclamaciones_power_line.construir(self.tab_reclamaciones_power_line)
        self.build_tab_config(acc_color)
        self.tabview.configure(command=self.actualizar_paneles_por_pestana)

        self.frame_progreso = ctk.CTkFrame(self, fg_color=self.color_superficie, corner_radius=10, border_width=1, border_color=self.color_borde)
        self.frame_progreso.pack(pady=(0, 8), padx=22, fill="x")
        ctk.CTkLabel(self.frame_progreso, text="Progreso", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=16, pady=(12, 4))
        barra_row = ctk.CTkFrame(self.frame_progreso, fg_color="transparent")
        barra_row.pack(padx=16, pady=(0, 12), fill="x")
        self.progreso_var = ctk.DoubleVar()
        self.barra_prog = ctk.CTkProgressBar(barra_row, variable=self.progreso_var, height=12, corner_radius=999, progress_color=self.color_principal, fg_color=self.color_superficie_alt)
        self.barra_prog.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.porcentaje_label = ctk.CTkLabel(barra_row, text="0%", font=("Segoe UI", 13, "bold"), text_color=self.color_pokedex_luz)
        self.porcentaje_label.pack(side="right")

        self.frame_estado = ctk.CTkFrame(self, fg_color=self.color_superficie, corner_radius=10, border_width=1, border_color=self.color_borde)
        self.frame_estado.pack(pady=(0, 8), padx=22, fill="x")
        ctk.CTkLabel(self.frame_estado, text="Estado", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=16, pady=(12, 4))
        self.estado_label = ctk.CTkLabel(self.frame_estado, text="Estado: listo", font=("Segoe UI", 12, "bold"), text_color=self.color_pokedex_luz)
        self.estado_label.pack(anchor="w", padx=16, pady=(0, 2))
        self.ruta_label = ctk.CTkLabel(self.frame_estado, text="Última salida: -", font=("Consolas", 10), text_color=self.color_texto_sec, anchor="w", justify="left")
        self.ruta_label.pack(anchor="w", padx=16, pady=(0, 8))
        self.frame_acciones = ctk.CTkFrame(self.frame_estado, fg_color="transparent")
        self.frame_acciones.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(self.frame_acciones, text="Abrir última salida", height=30, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, command=self.abrir_ultima_salida).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Abrir errores", height=30, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, command=lambda: self.abrir_directorio(ERROR_FOLDER)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Abrir log", height=30, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, command=self.abrir_log_detallado).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Limpiar log", height=30, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, command=self.limpiar_log).pack(side="left")

        self.frame_log = ctk.CTkFrame(self, fg_color=self.color_superficie, corner_radius=10, border_width=1, border_color=self.color_borde)
        self.frame_log.pack(pady=(0, 16), padx=22, fill="both", expand=True)
        ctk.CTkLabel(self.frame_log, text="Registro de actividad", font=("Segoe UI", 13, "bold"), text_color=self.color_texto).pack(anchor="w", padx=16, pady=(12, 8))
        self.log_box = ctk.CTkTextbox(self.frame_log, font=("Consolas", 12), border_width=1, border_color=self.color_borde, corner_radius=8, fg_color="#ffffff", text_color=self.color_texto)
        self.log_box.pack(padx=16, pady=(0, 16), fill="both", expand=True)
        self.log_box.tag_config("success", foreground=self.color_ok)
        self.log_box.tag_config("error", foreground=self.color_peligro)

    def actualizar_paneles_por_pestana(self):
        ocultar_paneles = self.tabview.get() == "Reclamaciones"
        paneles = (self.frame_progreso, self.frame_estado, self.frame_log)
        if ocultar_paneles:
            for panel in paneles:
                panel.pack_forget()
            self.tabview.pack_configure(fill="both", expand=True, pady=(16, 16))
            return
        self.tabview.pack_configure(fill="x", expand=False, pady=(16, 8))
        if not self.frame_progreso.winfo_manager():
            self.frame_progreso.pack(pady=(0, 8), padx=22, fill="x", before=self.frame_estado)
        if not self.frame_estado.winfo_manager():
            self.frame_estado.pack(pady=(0, 8), padx=22, fill="x", before=self.frame_log)
        if not self.frame_log.winfo_manager():
            self.frame_log.pack(pady=(0, 16), padx=22, fill="both", expand=True)

    def build_tab_auto(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_auto, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Consulta de disponibilidad", "Carga tu Excel y ejecuta la consulta paralela contra el portal.")
        self.crear_fila_archivo(card, "Excel de stock", "excel_auto_entry", "ruta_excel_auto", "Excel de stock...", "Examinar", lambda: self.seleccionar_archivo(self.excel_auto_entry, [("Excel files", "*.xlsx")]))
        row_input = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_input.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_input, text="Precio mínimo (EUR)", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.precio_entry = ctk.CTkEntry(row_input, width=120, height=36, justify="center", corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.precio_entry.insert(0, self.config.get("precio_min", "0"))
        self.precio_entry.pack(side="right", padx=14, pady=12)
        self.btn_auto = self.crear_boton_principal(card, "Iniciar automatización", self.toggle_auto)
        self.btn_auto.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card, text="Referencias consultadas", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.auto_log_box = self.crear_caja_log(card, alto=140)
        self.auto_log_box.pack(padx=24, pady=(0, 12), fill="x")
        ctk.CTkButton(card, text="Abrir carpeta del Excel", height=36, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, border_width=1, border_color=self.color_borde, command=self.abrir_carpeta_auto).pack(padx=24, pady=(0, 24), fill="x")

        card_rotacion = self.crear_tarjeta_tab(contenedor, "Filtro de rotación", "Marca referencias de baja rotación en base al umbral y la propuesta de pedido.")
        self.crear_fila_archivo(card_rotacion, "Excel de Quiter", "excel_quiter_entry", "ruta_excel_quiter", "Excel de Quiter (DMS)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_quiter_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card_rotacion, "Archivo de propuesta", "excel_rot_entry", "ruta_excel_rotacion", "Archivo PROPUESTAPEDIDO.xls...", "Examinar", lambda: self.seleccionar_archivo(self.excel_rot_entry, [("Excel files", "*.xls *.xlsx")]))
        self.crear_fila_archivo(card_rotacion, "Carpeta destino", "folder_rot_entry", "ruta_folder_rotacion", "Carpeta destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_rot_entry))
        row_umbral = ctk.CTkFrame(card_rotacion, fg_color=self.color_superficie_alt, corner_radius=10)
        row_umbral.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_umbral, text="Umbral mínimo", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.umbral_entry = ctk.CTkEntry(row_umbral, width=120, justify="center", height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.umbral_entry.insert(0, self.config.get("umbral_rotacion", "0.08"))
        self.umbral_entry.pack(side="right", padx=14, pady=12)
        self.btn_rotacion = self.crear_boton_principal(card_rotacion, "Resaltar por rotación", self.iniciar_filtro_rotacion)
        self.btn_rotacion.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card_rotacion, text="Registro de rotación", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.rotacion_log_box = self.crear_caja_log(card_rotacion, alto=110)
        self.rotacion_log_box.pack(padx=24, pady=(0, 12), fill="x")
        ctk.CTkButton(card_rotacion, text="Abrir carpeta destino", height=36, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, border_width=1, border_color=self.color_borde, command=self.abrir_carpeta_rotacion).pack(padx=24, pady=(0, 24), fill="x")

    def build_tab_pedidos(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_pedidos, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Generación de pedidos", "Convierte el Excel urgente en lotes CSV y guarda copia de seguridad antes de limpiar.")
        self.crear_fila_archivo(card, "Excel urgente", "excel_pedidos_entry", "ruta_excel_pedidos", "Excel URGENTE.xlsx...", "Examinar", lambda: self.seleccionar_archivo(self.excel_pedidos_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_pedidos_entry", "ruta_folder_pedidos", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_pedidos_entry))
        self.btn_pedidos = self.crear_boton_principal(card, "Generar pedidos", self.toggle_pedidos)
        self.btn_pedidos.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card, text="Números de pedido generados", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.pedidos_generados_box = ctk.CTkTextbox(card, height=90, font=("Consolas", 12), border_width=1, border_color=self.color_borde, corner_radius=8, fg_color="#ffffff", text_color=self.color_texto)
        self.pedidos_generados_box.pack(padx=24, pady=(0, 24), fill="x")
        self.pedidos_generados_box.configure(state="disabled")

        card_hoja = self.crear_tarjeta_tab(contenedor, "Hoja de pedidos del día", "Crea una copia en blanco de la plantilla, guardada con la fecha de hoy dentro de HOJAS DE PEDIDO.")
        self.crear_fila_archivo(card_hoja, "Plantilla de hoja", "plantilla_hoja_entry", "ruta_plantilla_hojas", "HOJA DE PEDIDOS.xlsx...", "Examinar", lambda: self.seleccionar_archivo(self.plantilla_hoja_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card_hoja, "Carpeta HOJAS DE PEDIDO", "carpeta_hojas_entry", "ruta_carpeta_hojas", "Carpeta HOJAS DE PEDIDO...", "Carpeta", lambda: self.seleccionar_carpeta(self.carpeta_hojas_entry))
        self.btn_hoja_pedidos = self.crear_boton_principal(card_hoja, "Generar hoja de hoy", self.iniciar_hoja_pedidos)
        self.btn_hoja_pedidos.pack(padx=24, pady=(4, 24), fill="x")

        card_avisos = self.crear_tarjeta_tab(contenedor, "Avisos de no suministro", "Lee la hoja de pedido ya rellenada y crea un borrador en Outlook por cliente con las referencias en NO SUMINISTRA. Nunca se envían solos.")
        self.crear_fila_archivo(card_avisos, "Hoja de pedido a procesar", "hoja_avisos_entry", "ruta_hoja_avisos", "Hoja dd-mm-aaaa.xlsx ya rellenada...", "Examinar", lambda: self.seleccionar_archivo(self.hoja_avisos_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card_avisos, "Excel de clientes", "clientes_avisos_entry", "ruta_excel_clientes", "clientes.xlsx...", "Examinar", lambda: self.seleccionar_archivo(self.clientes_avisos_entry, [("Excel files", "*.xlsx")]))
        self.btn_avisos = self.crear_boton_principal(card_avisos, "Generar avisos de no suministro", self.iniciar_avisos_no_suministro)
        self.btn_avisos.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card_avisos, text="Registro de avisos", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.avisos_log_box = self.crear_caja_log(card_avisos, alto=110)
        self.avisos_log_box.pack(padx=24, pady=(0, 24), fill="x")

    def build_tab_comparador(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_comparador, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Comparador de Excel", "Resalta las líneas repetidas y genera un archivo comparado en la carpeta elegida.")
        self.crear_fila_archivo(card, "Excel antiguo", "excel_old_entry", "ruta_excel_antiguo", "Excel antiguo...", "Examinar", lambda: self.seleccionar_archivo(self.excel_old_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Excel nuevo", "excel_new_entry", "ruta_excel_nuevo", "Excel nuevo...", "Examinar", lambda: self.seleccionar_archivo(self.excel_new_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_comp_entry", "ruta_folder_comparador", "Carpeta destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_comp_entry))
        row_color = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_color.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_color, text="Color de resaltado", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.color_optionmenu = ctk.CTkOptionMenu(row_color, values=["Amarillo", "Rojo", "Verde", "Azul", "Naranja"], height=36, corner_radius=8, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.color_optionmenu.set(self.config.get("color_resaltado", "Amarillo"))
        self.color_optionmenu.pack(side="right", padx=14, pady=12)
        self.btn_comparador = self.crear_boton_principal(card, "Comparar Excels", self.toggle_comparador)
        self.btn_comparador.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_mopar(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_mopar, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Generador Mopar", "Transforma el Excel en ficheros TXT Mopar con numeración diaria persistente.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_mopar_entry", "ruta_excel_mopar", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_mopar_entry, [("Excel files (.xls)", "*.xls")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_mopar_entry", "ruta_folder_mopar", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_mopar_entry))
        self.btn_mopar = self.crear_boton_principal(card, "Generar fichero Mopar", self.toggle_mopar)
        self.btn_mopar.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card, text="Ficheros Mopar generados", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.mopar_generados_box = ctk.CTkTextbox(card, height=90, font=("Consolas", 12), border_width=1, border_color=self.color_borde, corner_radius=8, fg_color="#ffffff", text_color=self.color_texto)
        self.mopar_generados_box.pack(padx=24, pady=(0, 24), fill="x")
        self.mopar_generados_box.configure(state="disabled")

    def build_tab_stock_fiat(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_stock_fiat, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Stock FIAT", "Transforma el Excel en ficheros TXT con el mismo formato final de Mopar.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_stock_fiat_entry", "ruta_excel_stock_fiat", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_stock_fiat_entry, [("Excel files (.xls)", "*.xls")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_stock_fiat_entry", "ruta_folder_stock_fiat", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_stock_fiat_entry))
        self.btn_stock_fiat = self.crear_boton_principal(card, "Generar Stock FIAT", self.toggle_stock_fiat)
        self.btn_stock_fiat.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_madd(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_madd, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Facturas Stellantis MADD", "Descarga diaria de facturas desde el portal MADD (www.madd.stellantis.com) y envía un correo con los PDF adjuntos.")
        frame_creds_madd = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        frame_creds_madd.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(frame_creds_madd, text="Usuario MADD", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=0, column=0, padx=18, pady=12, sticky="w")
        self.usuario_madd_entry = ctk.CTkEntry(frame_creds_madd, width=220, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.usuario_madd_entry.insert(0, self.config.get("usuario_madd", ""))
        self.usuario_madd_entry.grid(row=0, column=1, padx=18, pady=12, sticky="ew")
        ctk.CTkLabel(frame_creds_madd, text="Contraseña MADD", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=1, column=0, padx=18, pady=12, sticky="w")
        self.password_madd_entry = ctk.CTkEntry(frame_creds_madd, width=220, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto, show="*")
        self.password_madd_entry.insert(0, self.config.get("password_madd", ""))
        self.password_madd_entry.grid(row=1, column=1, padx=18, pady=12, sticky="ew")
        frame_creds_madd.grid_columnconfigure(1, weight=1)
        self.crear_fila_archivo(card, "Carpeta de facturas", "carpeta_facturas_madd_entry", "ruta_carpeta_facturas_madd", "Carpeta donde guardar las facturas...", "Carpeta", lambda: self.seleccionar_carpeta(self.carpeta_facturas_madd_entry))
        row_email = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_email.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_email, text="Correo diario de facturas", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=16, pady=(10, 2))
        self.email_facturas_madd_entry = ctk.CTkEntry(row_email, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.email_facturas_madd_entry.insert(0, self.config.get("email_facturas_madd", ""))
        self.email_facturas_madd_entry.pack(padx=16, pady=(0, 10), fill="x")
        row_email_fallo = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_email_fallo.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_email_fallo, text="Correo de aviso si falla", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=16, pady=(10, 2))
        self.email_aviso_fallo_madd_entry = ctk.CTkEntry(row_email_fallo, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.email_aviso_fallo_madd_entry.insert(0, self.config.get("email_aviso_fallo_madd", ""))
        self.email_aviso_fallo_madd_entry.pack(padx=16, pady=(0, 10), fill="x")
        row_filtro = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_filtro.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_filtro, text="Facturas de", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.filtro_dia_madd_option = ctk.CTkOptionMenu(row_filtro, values=["Ayer", "Hoy"], height=36, corner_radius=8, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.filtro_dia_madd_option.set(self.config.get("filtro_dia_facturas_madd", "Ayer"))
        self.filtro_dia_madd_option.pack(side="right", padx=14, pady=12)
        self.btn_madd = self.crear_boton_principal(card, "Ejecutar ahora (prueba manual)", self.iniciar_facturas_madd)
        self.btn_madd.pack(padx=24, pady=(4, 16), fill="x")
        ctk.CTkLabel(card, text="Registro de la ejecución", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=24, pady=(0, 6))
        self.madd_log_box = self.crear_caja_log(card, alto=140)
        self.madd_log_box.pack(padx=24, pady=(0, 24), fill="x")

        card_prog = self.crear_tarjeta_tab(contenedor, "Programación diaria", "Cuando compruebes que la ejecución manual funciona bien, activa la tarea de Windows para que corra sola todos los días laborables a las 08:00.")
        fila_prog = ctk.CTkFrame(card_prog, fg_color="transparent")
        fila_prog.pack(padx=24, pady=(0, 20), fill="x")
        self.btn_activar_tarea_madd = ctk.CTkButton(fila_prog, text="Activar tarea programada (08:00 L-V)", height=40, corner_radius=8, fg_color=acc_color, hover_color=self.color_hover, command=self.activar_tarea_programada_madd)
        self.btn_activar_tarea_madd.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.btn_desactivar_tarea_madd = ctk.CTkButton(fila_prog, text="Desactivar", height=40, corner_radius=8, fg_color=self.color_panel_medio, text_color=self.color_texto, hover_color=self.color_panel_medio_hover, border_width=1, border_color=self.color_borde, command=self.desactivar_tarea_programada_madd)
        self.btn_desactivar_tarea_madd.pack(side="right")

    def build_tab_config(self, acc_color):
        contenedor = ctk.CTkScrollableFrame(self.tab_config, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        card = self.crear_tarjeta_tab(contenedor, "Configuración general", "Ajusta credenciales, rendimiento y apariencia visual de la herramienta.")
        frame_creds = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        frame_creds.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(frame_creds, text="Usuario", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=0, column=0, padx=18, pady=12, sticky="w")
        self.user_entry = ctk.CTkEntry(frame_creds, width=220, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.user_entry.insert(0, self.config.get("usuario", ""))
        self.user_entry.grid(row=0, column=1, padx=18, pady=12, sticky="ew")
        ctk.CTkLabel(frame_creds, text="Contraseña", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=1, column=0, padx=18, pady=12, sticky="w")
        self.pass_entry = ctk.CTkEntry(frame_creds, width=220, height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto, show="*")
        self.pass_entry.insert(0, self.config.get("password", ""))
        self.pass_entry.grid(row=1, column=1, padx=18, pady=12, sticky="ew")
        frame_creds.grid_columnconfigure(1, weight=1)
        frame_perf = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        frame_perf.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(frame_perf, text="Hilos paralelos", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.hilos_option = ctk.CTkOptionMenu(frame_perf, values=["1", "2", "3", "4", "5"], height=36, corner_radius=8, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.hilos_option.set(str(self.config.get("num_hilos", "3")))
        self.hilos_option.pack(side="right", padx=14, pady=12)
        self.reclamaciones_power_line.construir_ajustes(card)
        ctk.CTkButton(card, text="Guardar configuración", height=42, corner_radius=8, fg_color=acc_color, hover_color=self.color_hover, command=self.guardar_config_ui).pack(padx=24, pady=(4, 24), fill="x")

    def seleccionar_archivo(self, entry, tipos):
        ruta = filedialog.askopenfilename(filetypes=tipos)
        if ruta:
            entry.delete(0, "end")
            entry.insert(0, ruta)

    def seleccionar_carpeta(self, entry):
        ruta = filedialog.askdirectory()
        if ruta:
            entry.delete(0, "end")
            entry.insert(0, ruta)

    def agregar_log(self, texto, tag=None):
        self.log_box.insert("end", f"{texto}\n", tag)
        self.log_box.see("end")

    def actualizar_progreso(self, valor):
        self.progreso_var.set(valor)
        self.porcentaje_label.configure(text=f"{int(valor * 100)}%")

    def actualizar_estado(self, texto, ruta=""):
        self.estado_label.configure(text=f"Estado: {texto}")
        if ruta:
            self.ultima_ruta_generada = ruta
            self.ruta_label.configure(text=f"Última salida: {ruta}")

    def limpiar_log(self):
        self.log_box.delete("1.0", "end")

    def crear_caja_log(self, parent, alto=110):
        caja = ctk.CTkTextbox(parent, height=alto, font=("Consolas", 12), border_width=1, border_color=self.color_borde, corner_radius=8, fg_color="#ffffff", text_color=self.color_texto)
        caja.tag_config("success", foreground=self.color_ok)
        caja.tag_config("error", foreground=self.color_peligro)
        caja.configure(state="disabled")
        return caja

    def agregar_a_caja(self, caja, texto, tag=None):
        caja.configure(state="normal")
        caja.insert("end", f"{texto}\n", tag)
        caja.configure(state="disabled")
        caja.see("end")

    def limpiar_caja(self, caja):
        caja.configure(state="normal")
        caja.delete("1.0", "end")
        caja.configure(state="disabled")

    def log_callback_auto(self, texto, tag=None):
        self.agregar_log(texto, tag)
        self.agregar_a_caja(self.auto_log_box, texto, tag)

    def log_callback_rotacion(self, texto, tag=None):
        self.agregar_log(texto, tag)
        self.agregar_a_caja(self.rotacion_log_box, texto, tag)

    def abrir_carpeta_auto(self):
        ruta = self.widget_get("excel_auto_entry")
        if not ruta:
            messagebox.showinfo("Información", "Selecciona primero un Excel de stock.")
            return
        self.abrir_directorio(os.path.dirname(ruta) if os.path.isfile(ruta) else ruta)

    def abrir_carpeta_rotacion(self):
        ruta = self.widget_get("folder_rot_entry")
        if not ruta:
            messagebox.showinfo("Información", "Selecciona primero una carpeta de destino.")
            return
        self.abrir_directorio(ruta)

    def agregar_pedido_generado(self, numero):
        self.pedidos_generados_box.configure(state="normal")
        self.pedidos_generados_box.insert("end", f"{numero}\n")
        self.pedidos_generados_box.configure(state="disabled")
        self.pedidos_generados_box.see("end")

    def limpiar_pedidos_generados(self):
        self.pedidos_generados_box.configure(state="normal")
        self.pedidos_generados_box.delete("1.0", "end")
        self.pedidos_generados_box.configure(state="disabled")

    def agregar_mopar_generado(self, nombre_fichero):
        self.mopar_generados_box.configure(state="normal")
        self.mopar_generados_box.insert("end", f"{nombre_fichero}\n")
        self.mopar_generados_box.configure(state="disabled")
        self.mopar_generados_box.see("end")

    def limpiar_mopar_generados(self):
        self.mopar_generados_box.configure(state="normal")
        self.mopar_generados_box.delete("1.0", "end")
        self.mopar_generados_box.configure(state="disabled")

    def abrir_log_detallado(self):
        if not os.path.exists(LOG_DETALLADO):
            Path(LOG_DETALLADO).touch()
        self.abrir_directorio(LOG_DETALLADO)

    def abrir_directorio(self, ruta):
        if not abrir_ruta_en_sistema(ruta):
            messagebox.showwarning("Aviso", "No se pudo abrir la ruta seleccionada.")

    def abrir_ultima_salida(self):
        if not self.ultima_ruta_generada:
            messagebox.showinfo("Información", "Todavía no hay una salida generada.")
            return
        if os.path.isfile(self.ultima_ruta_generada):
            ruta_a_abrir = os.path.dirname(self.ultima_ruta_generada)
        else:
            ruta_a_abrir = self.ultima_ruta_generada
        self.abrir_directorio(ruta_a_abrir)

    def widget_get(self, attr_name, default=""):
        widget = getattr(self, attr_name, None)
        if not widget:
            return default
        try:
            return widget.get()
        except Exception:
            return default

    def validar_archivo(self, ruta, mensaje):
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("Aviso", mensaje)
            return False
        return True

    def validar_carpeta(self, ruta, mensaje):
        if not ruta:
            messagebox.showwarning("Aviso", mensaje)
            return False
        Path(ruta).mkdir(parents=True, exist_ok=True)
        return True

    def guardar_config_ui(self, silent=False):
        self.guardar_configuracion_ui_internal()
        if not silent:
            messagebox.showinfo("Éxito", "Configuración guardada.")

    def toggle_auto(self):
        global automation_running
        if automation_running:
            stop_event_auto.set()
            self.btn_auto.configure(text="Deteniendo...", state="disabled")
        else:
            self.iniciar_auto()

    def iniciar_auto(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        try:
            self.guardar_configuracion_ui_internal()
            ruta = self.widget_get("excel_auto_entry")
            user = self.widget_get("user_entry")
            pwd = self.widget_get("pass_entry")
            hilos = int(self.widget_get("hilos_option", "3"))
            precio = float(self.widget_get("precio_entry", "0"))
        except Exception as e:
            registrar_error_detallado(str(e), traceback.format_exc())
            messagebox.showerror("Error", "La automatizacion no pudo iniciarse por un problema interno de la interfaz.")
            return
        if precio < 0:
            messagebox.showwarning("Aviso", "Introduce un precio minimo valido.")
            return
        if not self.validar_archivo(ruta, "Selecciona un Excel valido para la automatizacion."):
            return
        if not user or not pwd:
            messagebox.showwarning("Aviso", "Completa usuario y contrasena antes de continuar.")
            return
        self.actualizar_estado("automatización en curso", ruta)
        self.btn_auto.configure(text="Detener", fg_color=self.color_peligro, hover_color=self.color_peligro_hover)
        self.limpiar_caja(self.auto_log_box)
        stop_event_auto.clear()
        threading.Thread(target=ejecutar_automatizacion_task, args=(ruta, user, pwd, precio, hilos, lambda t, tag=None: self.after(0, self.log_callback_auto, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_auto), stop_event_auto), daemon=True).start()

    def finalizar_auto(self):
        acc_color = self.color_principal
        self.actualizar_estado("automatización finalizada", self.excel_auto_entry.get())
        self.btn_auto.configure(text="Iniciar automatización", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def toggle_pedidos(self):
        global pedidos_running
        if not pedidos_running:
            self.iniciar_pedidos()
        else:
            stop_event_pedidos.set()

    def iniciar_pedidos(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_pedidos_entry.get()
        folder = self.folder_pedidos_entry.get()
        if not self.validar_archivo(ruta, "Selecciona un Excel válido para generar pedidos."):
            return
        if not self.validar_carpeta(folder, "Selecciona una carpeta de destino para los pedidos."):
            return
        self.actualizar_estado("generando pedidos", folder)
        self.btn_pedidos.configure(text="Detener", fg_color=self.color_peligro, hover_color=self.color_peligro_hover)
        self.limpiar_pedidos_generados()
        stop_event_pedidos.clear()
        threading.Thread(target=generar_pedidos_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_pedidos), stop_event_pedidos, lambda n: self.after(0, self.agregar_pedido_generado, n)), daemon=True).start()

    def finalizar_pedidos(self):
        acc_color = self.color_principal
        self.actualizar_estado("pedidos generados", self.folder_pedidos_entry.get())
        self.btn_pedidos.configure(text="Generar pedidos", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def iniciar_hoja_pedidos(self, forzar=False, automatico=False):
        global hoja_pedidos_running
        if hoja_pedidos_running:
            return
        self.guardar_configuracion_ui_internal()
        plantilla = self.widget_get("plantilla_hoja_entry")
        carpeta = self.widget_get("carpeta_hojas_entry")
        if not self.validar_archivo(plantilla, "Selecciona la plantilla HOJA DE PEDIDOS.xlsx."):
            return
        if not self.validar_carpeta(carpeta, "Selecciona la carpeta HOJAS DE PEDIDO."):
            return
        if not forzar and ruta_hoja_pedidos_del_dia(carpeta).exists():
            if automatico:
                return
            if not messagebox.askyesno("La hoja de hoy ya existe", "Ya existe una hoja de pedidos generada hoy. ¿Quieres sustituirla por una copia en blanco de la plantilla? Se perderá lo que hayas escrito en ella."):
                return
            forzar = True
        self.actualizar_estado("generando hoja de pedidos del día")
        self.btn_hoja_pedidos.configure(text="Generando...", state="disabled")
        threading.Thread(
            target=generar_hoja_pedidos_task,
            args=(plantilla, carpeta),
            kwargs={
                "log_callback": lambda t, tag=None: self.after(0, self.agregar_log, t, tag),
                "progreso_callback": lambda v: self.after(0, self.actualizar_progreso, v),
                "finalizado_callback": lambda: self.after(0, self.finalizar_hoja_pedidos, carpeta),
                "forzar": forzar,
            },
            daemon=True,
        ).start()

    def finalizar_hoja_pedidos(self, carpeta):
        acc_color = self.color_principal
        ruta = ruta_hoja_pedidos_del_dia(carpeta)
        self.actualizar_estado("hoja de pedidos lista", str(ruta))
        self.btn_hoja_pedidos.configure(text="Generar hoja de hoy", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def comprobar_hoja_pedidos_al_inicio(self):
        plantilla = self.config.get("ruta_plantilla_hojas", "")
        carpeta = self.config.get("ruta_carpeta_hojas", "")
        if not plantilla or not carpeta or not os.path.exists(plantilla) or not os.path.isdir(carpeta):
            return
        if ruta_hoja_pedidos_del_dia(carpeta).exists():
            return
        self.agregar_log("[HOJA PEDIDOS] No existe la hoja de hoy, generándola automáticamente...")
        threading.Thread(
            target=generar_hoja_pedidos_task,
            args=(plantilla, carpeta),
            kwargs={
                "log_callback": lambda t, tag=None: self.after(0, self.agregar_log, t, tag),
                "finalizado_callback": lambda: self.after(0, self.finalizar_hoja_pedidos, carpeta),
            },
            daemon=True,
        ).start()

    def log_callback_avisos(self, texto, tag=None):
        self.agregar_log(texto, tag)
        self.agregar_a_caja(self.avisos_log_box, texto, tag)

    def iniciar_avisos_no_suministro(self):
        global avisos_running
        if avisos_running:
            return
        self.guardar_configuracion_ui_internal()
        ruta_hoja = self.widget_get("hoja_avisos_entry")
        ruta_clientes = self.widget_get("clientes_avisos_entry")
        if not self.validar_archivo(ruta_hoja, "Selecciona la hoja de pedido (dd-mm-aaaa.xlsx) ya rellenada."):
            return
        if not self.validar_archivo(ruta_clientes, "Selecciona el Excel de clientes (clientes.xlsx)."):
            return
        if not messagebox.askyesno("Generar avisos", "Se creará un borrador en Outlook por cada cliente con referencias en NO SUMINISTRA. Se guardan en Borradores sin enviarse. ¿Continuar?"):
            return
        self.actualizar_estado("generando avisos de no suministro")
        self.btn_avisos.configure(text="Generando...", state="disabled")
        self.limpiar_caja(self.avisos_log_box)
        threading.Thread(
            target=generar_avisos_no_suministro_task,
            args=(ruta_hoja, ruta_clientes),
            kwargs={
                "log_callback": lambda t, tag=None: self.after(0, self.log_callback_avisos, t, tag),
                "progreso_callback": lambda v: self.after(0, self.actualizar_progreso, v),
                "finalizado_callback": lambda: self.after(0, self.finalizar_avisos_no_suministro),
                "resultado_callback": lambda r: self.after(0, self.mostrar_resultado_avisos, r),
            },
            daemon=True,
        ).start()

    def finalizar_avisos_no_suministro(self):
        acc_color = self.color_principal
        self.actualizar_estado("avisos de no suministro procesados")
        self.btn_avisos.configure(text="Generar avisos de no suministro", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def mostrar_resultado_avisos(self, resultado):
        if resultado.get("error"):
            return
        generados = resultado.get("generados", [])
        sin_correo = resultado.get("sin_correo", [])
        if not generados and not sin_correo:
            return
        lineas = [f"Borradores generados: {len(generados)}"]
        for entrada in generados:
            lineas.append(f"  • {entrada['nombre'] or entrada['cliente']} ({entrada['cliente']}): {', '.join(entrada['referencias'])}")
        if sin_correo:
            lineas.append("")
            lineas.append(f"Sin email en el Excel de clientes ({len(sin_correo)}), borrador creado sin destinatario:")
            for entrada in sin_correo:
                lineas.append(f"  • Cliente {entrada['cliente']}: {', '.join(entrada['referencias'])}")
        messagebox.showinfo("Avisos de no suministro", "\n".join(lineas))

    def toggle_comparador(self):
        global comparador_running
        if not comparador_running:
            self.iniciar_comparador()
        else:
            stop_event_comparador.set()

    def iniciar_comparador(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        self.guardar_configuracion_ui_internal()
        ruta_ant = self.excel_old_entry.get()
        ruta_nue = self.excel_new_entry.get()
        folder = self.folder_comp_entry.get()
        color = self.color_optionmenu.get()
        colores_hex = {"Amarillo": "FFFF00", "Rojo": "FF0000", "Verde": "00FF00", "Azul": "0000FF", "Naranja": "FFA500"}
        if not self.validar_archivo(ruta_ant, "Selecciona el Excel antiguo."):
            return
        if not self.validar_archivo(ruta_nue, "Selecciona el Excel nuevo."):
            return
        if not self.validar_carpeta(folder, "Selecciona una carpeta para el resultado comparado."):
            return
        ruta_salida = str(Path(folder) / "resultado_comparacion.xlsx")
        self.actualizar_estado("comparando excels", ruta_salida)
        self.btn_comparador.configure(text="Detener", fg_color=self.color_peligro, hover_color=self.color_peligro_hover)
        stop_event_comparador.clear()
        threading.Thread(target=comparar_excel_task, args=(ruta_ant, ruta_nue, ruta_salida, colores_hex[color], lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_comparador), stop_event_comparador), daemon=True).start()

    def finalizar_comparador(self):
        acc_color = self.color_principal
        self.actualizar_estado("comparación finalizada", str(Path(self.folder_comp_entry.get()) / "resultado_comparacion.xlsx"))
        self.btn_comparador.configure(text="Comparar Excels", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def iniciar_filtro_rotacion(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        self.guardar_configuracion_ui_internal()
        ruta_quiter = self.excel_quiter_entry.get()
        ruta_rot = self.excel_rot_entry.get()
        folder_salida = self.folder_rot_entry.get()
        try:
            umbral = float(self.umbral_entry.get())
        except Exception:
            messagebox.showwarning("Aviso", "Introduce un umbral de rotación válido.")
            return
        if not self.validar_archivo(ruta_quiter, "Selecciona el Excel de Quiter."):
            return
        if not self.validar_archivo(ruta_rot, "Selecciona el archivo de propuesta."):
            return
        if not self.validar_carpeta(folder_salida, "Selecciona una carpeta de salida para rotación."):
            return
        ruta_estimacion = str(Path(folder_salida) / f"{Path(ruta_quiter).stem}_ROTACION.xlsx")
        self.actualizar_estado("analizando rotación", ruta_estimacion)
        self.btn_rotacion.configure(state="disabled", text="Analizando...")
        self.limpiar_caja(self.rotacion_log_box)
        threading.Thread(target=filtrar_por_rotacion_task, args=(ruta_quiter, ruta_rot, folder_salida, umbral, "FF0000", lambda t, tag=None: self.after(0, self.log_callback_rotacion, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_filtro_rotacion)), daemon=True).start()

    def finalizar_filtro_rotacion(self):
        ruta_quiter = self.excel_quiter_entry.get()
        if self.folder_rot_entry.get() and ruta_quiter:
            ruta_final = str(Path(self.folder_rot_entry.get()) / f"{Path(ruta_quiter).stem}_ROTACION.xlsx")
            self.actualizar_estado("rotación completada", ruta_final)
        self.btn_rotacion.configure(state="normal", text="Resaltar por rotación")

    def toggle_mopar(self):
        global mopar_running
        if mopar_running:
            stop_event_mopar.set()
            self.btn_mopar.configure(text="Deteniendo...", state="disabled")
        else:
            self.iniciar_mopar()

    def iniciar_mopar(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_mopar_entry.get()
        folder = self.folder_mopar_entry.get()
        if not self.validar_archivo(ruta, "Por favor selecciona un archivo Excel válido."):
            return
        if not self.validar_carpeta(folder, "Por favor selecciona una carpeta de destino."):
            return
        self.actualizar_estado("generando fichero Mopar", folder)
        if not os.path.exists(ruta):
            messagebox.showwarning("Aviso", "Por favor selecciona un archivo Excel válido.")
            return
        if not os.path.exists(folder):
            messagebox.showwarning("Aviso", "Por favor selecciona una carpeta de destino.")
            return
        self.btn_mopar.configure(text="Detener", fg_color=self.color_peligro, hover_color=self.color_peligro_hover)
        self.limpiar_mopar_generados()
        stop_event_mopar.clear()
        threading.Thread(target=generar_fichero_mopar_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_mopar), stop_event_mopar, lambda n: self.after(0, self.agregar_mopar_generado, n)), daemon=True).start()

    def toggle_stock_fiat(self):
        global stock_fiat_running
        if stock_fiat_running:
            stop_event_stock_fiat.set()
            self.btn_stock_fiat.configure(text="Deteniendo...", state="disabled")
        else:
            self.iniciar_stock_fiat()

    def iniciar_stock_fiat(self):
        activa = tarea_activa()
        if activa:
            messagebox.showwarning("Aviso", f"Ya hay una tarea en curso ({activa}). Espera a que termine antes de iniciar otra.")
            return
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_stock_fiat_entry.get()
        folder = self.folder_stock_fiat_entry.get()
        if not self.validar_archivo(ruta, "Por favor selecciona un archivo Excel válido para Stock FIAT."):
            return
        if not self.validar_carpeta(folder, "Por favor selecciona una carpeta de destino para Stock FIAT."):
            return
        self.actualizar_estado("generando Stock FIAT", folder)
        if not os.path.exists(ruta):
            messagebox.showwarning("Aviso", "Por favor selecciona un archivo Excel válido para Stock FIAT.")
            return
        if not os.path.exists(folder):
            messagebox.showwarning("Aviso", "Por favor selecciona una carpeta de destino para Stock FIAT.")
            return
        self.btn_stock_fiat.configure(text="Detener", fg_color=self.color_peligro, hover_color=self.color_peligro_hover)
        stop_event_stock_fiat.clear()
        threading.Thread(target=generar_fichero_stock_fiat_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_stock_fiat), stop_event_stock_fiat), daemon=True).start()

    def log_callback_madd(self, texto, tag=None):
        self.agregar_log(texto, tag)
        self.agregar_a_caja(self.madd_log_box, texto, tag)

    def iniciar_facturas_madd(self):
        global madd_running
        if madd_running:
            return
        self.guardar_configuracion_ui_internal()
        usuario = self.widget_get("usuario_madd_entry")
        password = self.widget_get("password_madd_entry")
        carpeta = self.widget_get("carpeta_facturas_madd_entry")
        email_destino = self.widget_get("email_facturas_madd_entry")
        email_aviso = self.widget_get("email_aviso_fallo_madd_entry")
        filtro_dia = self.widget_get("filtro_dia_madd_option", "Ayer")
        if not usuario or not password:
            messagebox.showwarning("Aviso", "Completa el usuario y la contraseña de MADD antes de continuar.")
            return
        if not self.validar_carpeta(carpeta, "Selecciona la carpeta donde guardar las facturas."):
            return
        if not email_destino:
            messagebox.showwarning("Aviso", "Introduce el correo de destino de las facturas.")
            return
        if not messagebox.askyesno("Ejecutar prueba manual", "Esto iniciará sesión en MADD, descargará las facturas del día seleccionado y ENVIARÁ un correo real con los adjuntos (o el aviso de 'sin facturas'). ¿Continuar?"):
            return
        self.actualizar_estado("descargando facturas MADD")
        self.btn_madd.configure(text="Ejecutando...", state="disabled")
        self.limpiar_caja(self.madd_log_box)
        threading.Thread(
            target=tarea_facturas_madd,
            args=(usuario, password, carpeta, email_destino, email_aviso),
            kwargs={
                "filtro_dia": filtro_dia,
                "log_callback": lambda t, tag=None: self.after(0, self.log_callback_madd, t, tag),
                "finalizado_callback": lambda: self.after(0, self.finalizar_facturas_madd),
            },
            daemon=True,
        ).start()

    def finalizar_facturas_madd(self):
        acc_color = self.color_principal
        self.actualizar_estado("descarga de facturas MADD finalizada", self.widget_get("carpeta_facturas_madd_entry"))
        self.btn_madd.configure(text="Ejecutar ahora (prueba manual)", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def activar_tarea_programada_madd(self):
        if not getattr(sys, "frozen", False):
            messagebox.showwarning("Aviso", "La tarea programada solo se puede activar desde el .exe compilado (dist/), no ejecutando el código fuente con python.")
            return
        if not messagebox.askyesno("Activar tarea programada", "Esto crea una tarea en el Programador de tareas de Windows que ejecutará la descarga de facturas MADD todos los días laborables a las 08:00, aunque la app esté cerrada. ¿Confirmas que ya has probado la ejecución manual y ha funcionado bien?"):
            return
        ruta_exe = sys.executable
        comando = f'schtasks /create /tn "PalausaFacturasMADD" /tr "\\"{ruta_exe}\\" --tarea-madd" /sc weekly /d MON,TUE,WED,THU,FRI /st 08:00 /f'
        try:
            resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)
            if resultado.returncode == 0:
                messagebox.showinfo("Éxito", "Tarea programada activada: se ejecutará todos los días laborables a las 08:00.")
            else:
                messagebox.showerror("Error", f"No se pudo crear la tarea programada:\n{resultado.stderr or resultado.stdout}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la tarea programada: {e}")

    def desactivar_tarea_programada_madd(self):
        try:
            resultado = subprocess.run('schtasks /delete /tn "PalausaFacturasMADD" /f', shell=True, capture_output=True, text=True)
            if resultado.returncode == 0:
                messagebox.showinfo("Hecho", "Tarea programada desactivada.")
            else:
                messagebox.showinfo("Aviso", "No había ninguna tarea programada activa (o ya se había desactivado).")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo desactivar la tarea: {e}")

    def guardar_configuracion_ui_internal(self):
        config = {
            "usuario": self.widget_get("user_entry"),
            "password": self.widget_get("pass_entry"),
            "num_hilos": int(self.widget_get("hilos_option", "3")),
            "ruta_excel_auto": self.widget_get("excel_auto_entry"),
            "ruta_excel_pedidos": self.widget_get("excel_pedidos_entry"),
            "ruta_folder_pedidos": self.widget_get("folder_pedidos_entry"),
            "ruta_plantilla_hojas": self.widget_get("plantilla_hoja_entry"),
            "ruta_carpeta_hojas": self.widget_get("carpeta_hojas_entry"),
            "ruta_hoja_avisos": self.widget_get("hoja_avisos_entry"),
            "ruta_excel_clientes": self.widget_get("clientes_avisos_entry"),
            "precio_min": self.widget_get("precio_entry", "0"),
            "ruta_excel_antiguo": self.widget_get("excel_old_entry"),
            "ruta_excel_nuevo": self.widget_get("excel_new_entry"),
            "ruta_folder_comparador": self.widget_get("folder_comp_entry"),
            "ruta_excel_quiter": self.widget_get("excel_quiter_entry"),
            "ruta_excel_rotacion": self.widget_get("excel_rot_entry"),
            "ruta_folder_rotacion": self.widget_get("folder_rot_entry"),
            "umbral_rotacion": self.widget_get("umbral_entry", "0.08"),
            "ruta_excel_mopar": self.widget_get("excel_mopar_entry"),
            "ruta_folder_mopar": self.widget_get("folder_mopar_entry"),
            "ruta_excel_stock_fiat": self.widget_get("excel_stock_fiat_entry"),
            "ruta_folder_stock_fiat": self.widget_get("folder_stock_fiat_entry"),
            "usuario_madd": self.widget_get("usuario_madd_entry"),
            "password_madd": self.widget_get("password_madd_entry"),
            "ruta_carpeta_facturas_madd": self.widget_get("carpeta_facturas_madd_entry"),
            "email_facturas_madd": self.widget_get("email_facturas_madd_entry", ""),
            "email_aviso_fallo_madd": self.widget_get("email_aviso_fallo_madd_entry", ""),
            "filtro_dia_facturas_madd": self.widget_get("filtro_dia_madd_option", "Ayer"),
            "color_resaltado": self.widget_get("color_optionmenu", "Amarillo"),
            "modo_apariencia": self.widget_get("modo_option", self.config.get("modo_apariencia", "Dark")),
            "mopar_fecha": self.config.get("mopar_fecha"),
            "mopar_numero": self.config.get("mopar_numero"),
            "stock_fiat_fecha": self.config.get("stock_fiat_fecha"),
            "stock_fiat_numero": self.config.get("stock_fiat_numero")
        }
        if hasattr(self, "reclamaciones_power_line"):
            config.update(self.reclamaciones_power_line.obtener_configuracion())
            self.config.update(config)
        guardar_configuracion(config)

    def finalizar_mopar(self):
        acc_color = self.color_principal
        self.actualizar_estado("fichero Mopar generado", self.folder_mopar_entry.get())
        self.btn_mopar.configure(text="Generar fichero Mopar", fg_color=acc_color, hover_color=self.color_hover, state="normal")

    def finalizar_stock_fiat(self):
        acc_color = self.color_principal
        self.actualizar_estado("Stock FIAT generado", self.folder_stock_fiat_entry.get())
        self.btn_stock_fiat.configure(text="Generar Stock FIAT", fg_color=acc_color, hover_color=self.color_hover, state="normal")

if __name__ == "__main__":
    if "--tarea-madd" in sys.argv:
        ejecutar_tarea_madd_standalone()
        sys.exit(0)
    app = AppIntegrada()
    app.mainloop()

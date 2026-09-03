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
from datetime import datetime
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

def generar_pedidos_task(ruta_excel, carpeta_pedidos, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None):
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
# 5. LÓGICA DE FICHERO MOPAR (CON NOMBRES FXXXX + LETRAS) - v13
# ---------------------------------------------------

def generar_fichero_mopar_task(ruta_excel, carpeta_salida, log_callback=None, progreso_callback=None, finalizado_callback=None, stop_event=None):
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
                archivos_creados += 1
            
            log(f"✨ Proceso completado. Se han generado {archivos_creados} archivos con prefijo F{proximo_f}.", "success")
    except Exception as e:
        log(f"❌ Error crítico: {e}", "error")
        registrar_error_detallado(str(e), traceback.format_exc())
    finally:
        mopar_running = False
        if finalizado_callback: finalizado_callback()

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
        self.tab_rotacion = self.tabview.add("Filtro Rotación")
        self.tab_mopar = self.tabview.add("Fichero Mopar")
        self.tab_stock_fiat = self.tabview.add("Stock FIAT")
        self.tab_reclamaciones_power_line = self.tabview.add("Reclamaciones")
        self.tab_config = self.tabview.add("Configuración")
        self.build_tab_auto(acc_color)
        self.build_tab_pedidos(acc_color)
        self.build_tab_comparador(acc_color)
        self.build_tab_rotacion(acc_color)
        self.build_tab_mopar(acc_color)
        self.build_tab_stock_fiat(acc_color)
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
        card = self.crear_tarjeta_tab(self.tab_auto, "Consulta de disponibilidad", "Carga tu Excel y ejecuta la consulta paralela contra el portal.")
        self.crear_fila_archivo(card, "Excel de stock", "excel_auto_entry", "ruta_excel_auto", "Excel de stock...", "Examinar", lambda: self.seleccionar_archivo(self.excel_auto_entry, [("Excel files", "*.xlsx")]))
        row_input = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_input.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_input, text="Precio mínimo (EUR)", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.precio_entry = ctk.CTkEntry(row_input, width=120, height=36, justify="center", corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.precio_entry.insert(0, self.config.get("precio_min", "0"))
        self.precio_entry.pack(side="right", padx=14, pady=12)
        self.btn_auto = self.crear_boton_principal(card, "Iniciar automatización", self.toggle_auto)
        self.btn_auto.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_pedidos(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_pedidos, "Generación de pedidos", "Convierte el Excel urgente en lotes CSV y guarda copia de seguridad antes de limpiar.")
        self.crear_fila_archivo(card, "Excel urgente", "excel_pedidos_entry", "ruta_excel_pedidos", "Excel URGENTE.xlsx...", "Examinar", lambda: self.seleccionar_archivo(self.excel_pedidos_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_pedidos_entry", "ruta_folder_pedidos", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_pedidos_entry))
        self.btn_pedidos = self.crear_boton_principal(card, "Generar pedidos", self.toggle_pedidos)
        self.btn_pedidos.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_comparador(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_comparador, "Comparador de Excel", "Resalta las líneas repetidas y genera un archivo comparado en la carpeta elegida.")
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

    def build_tab_rotacion(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_rotacion, "Filtro de rotación", "Marca referencias de baja rotación en base al umbral y la propuesta de pedido.")
        self.crear_fila_archivo(card, "Excel de Quiter", "excel_quiter_entry", "ruta_excel_quiter", "Excel de Quiter (DMS)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_quiter_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Archivo de propuesta", "excel_rot_entry", "ruta_excel_rotacion", "Archivo PROPUESTAPEDIDO.xls...", "Examinar", lambda: self.seleccionar_archivo(self.excel_rot_entry, [("Excel files", "*.xls *.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_rot_entry", "ruta_folder_rotacion", "Carpeta destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_rot_entry))
        row_umbral = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=10)
        row_umbral.pack(padx=24, pady=(0, 16), fill="x")
        ctk.CTkLabel(row_umbral, text="Umbral mínimo", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.umbral_entry = ctk.CTkEntry(row_umbral, width=120, justify="center", height=36, corner_radius=8, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.umbral_entry.insert(0, self.config.get("umbral_rotacion", "0.08"))
        self.umbral_entry.pack(side="right", padx=14, pady=12)
        self.btn_rotacion = self.crear_boton_principal(card, "Resaltar por rotación", self.iniciar_filtro_rotacion)
        self.btn_rotacion.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_mopar(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_mopar, "Generador Mopar", "Transforma el Excel en ficheros TXT Mopar con numeración diaria persistente.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_mopar_entry", "ruta_excel_mopar", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_mopar_entry, [("Excel files (.xls)", "*.xls")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_mopar_entry", "ruta_folder_mopar", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_mopar_entry))
        self.btn_mopar = self.crear_boton_principal(card, "Generar fichero Mopar", self.toggle_mopar)
        self.btn_mopar.pack(padx=24, pady=(4, 24), fill="x")

    def build_tab_stock_fiat(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_stock_fiat, "Stock FIAT", "Transforma el Excel en ficheros TXT con el mismo formato final de Mopar.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_stock_fiat_entry", "ruta_excel_stock_fiat", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_stock_fiat_entry, [("Excel files (.xls)", "*.xls")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_stock_fiat_entry", "ruta_folder_stock_fiat", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_stock_fiat_entry))
        self.btn_stock_fiat = self.crear_boton_principal(card, "Generar Stock FIAT", self.toggle_stock_fiat)
        self.btn_stock_fiat.pack(padx=24, pady=(4, 24), fill="x")

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
        stop_event_auto.clear()
        threading.Thread(target=ejecutar_automatizacion_task, args=(ruta, user, pwd, precio, hilos, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_auto), stop_event_auto), daemon=True).start()

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
        stop_event_pedidos.clear()
        threading.Thread(target=generar_pedidos_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_pedidos), stop_event_pedidos), daemon=True).start()

    def finalizar_pedidos(self):
        acc_color = self.color_principal
        self.actualizar_estado("pedidos generados", self.folder_pedidos_entry.get())
        self.btn_pedidos.configure(text="Generar pedidos", fg_color=acc_color, hover_color=self.color_hover, state="normal")

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
        threading.Thread(target=filtrar_por_rotacion_task, args=(ruta_quiter, ruta_rot, folder_salida, umbral, "FF0000", lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_filtro_rotacion)), daemon=True).start()

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
        stop_event_mopar.clear()
        threading.Thread(target=generar_fichero_mopar_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_mopar), stop_event_mopar), daemon=True).start()

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

    def guardar_configuracion_ui_internal(self):
        config = {
            "usuario": self.widget_get("user_entry"),
            "password": self.widget_get("pass_entry"),
            "num_hilos": int(self.widget_get("hilos_option", "3")),
            "ruta_excel_auto": self.widget_get("excel_auto_entry"),
            "ruta_excel_pedidos": self.widget_get("excel_pedidos_entry"),
            "ruta_folder_pedidos": self.widget_get("folder_pedidos_entry"),
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
    app = AppIntegrada()
    app.mainloop()

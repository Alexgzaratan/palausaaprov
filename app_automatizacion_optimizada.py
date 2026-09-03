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
from PIL import Image
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
REGION_ICON_FOLDER = Path("assets") / "region_icons"

# Colores del Tema Lila (Cyber-Lila)
COLOR_LILA = "#9b59b6"
COLOR_LILA_HOVER = "#8e44ad"
COLOR_FONDO_NEGRO = "#1a1a1a"

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

def cargar_configuracion():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except: return {}
    return {}

def guardar_configuracion(config_dict):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config_dict, f, indent=4)
    except: pass

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

def obtener_numero_mopar_del_dia(carpeta_salida):
    config = cargar_configuracion()
    hoy = datetime.now().strftime("%Y-%m-%d")
    ultimo_dia = config.get("mopar_fecha")
    ultimo_numero = config.get("mopar_numero")

    if ultimo_dia == hoy and ultimo_numero is not None:
        return int(ultimo_numero)

    ultimo_f = 999
    for f in os.listdir(carpeta_salida):
        m = re.match(r"F(\d+)", f)
        if m:
            ultimo_f = max(ultimo_f, int(m.group(1)))

    proximo_f = ultimo_f + 1
    config["mopar_fecha"] = hoy
    config["mopar_numero"] = proximo_f
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
            except: continue
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
            context = browser.new_context()
            context.route("**/*.{png,jpg,jpeg,gif,svg,woff,woff2}", lambda route: route.abort())
            page = context.new_page()
            url_acceso = f"https://{usuario}:{password}@parts.citroen.com/zzlogpl0"
            try:
                page.goto(url_acceso, timeout=45000)
                page.wait_for_selector("text=/Espa.ol/", timeout=20000)
                page.locator("text=/Espa.ol/").first.click()
                page.wait_for_selector("text=/Referencia/", timeout=25000)
            except Exception as e:
                log(f"Error de acceso/login: {e}", "error")
                registrar_error_detallado(str(e), traceback.format_exc())
                browser.close()
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
                log("No se encontr? el frame de trabajo.", "error")
                browser.close()
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

        def update_progress():
            nonlocal procesadas, cierre_notificado
            global automation_running
            procesadas += 1
            progreso = procesadas / total_tareas
            if progreso_callback:
                progreso_callback(progreso)
            if procesadas >= total_tareas and not cierre_notificado:
                cierre_notificado = True
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
                except:
                    cant_val = 0
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
            
            # Buscar el último número FXXXX en la carpeta
            ultimo_f = 999
            for f in os.listdir(carpeta_salida):
                m = re.match(r"F(\d+)", f)
                if m:
                    ultimo_f = max(ultimo_f, int(m.group(1)))
            
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
                except:
                    cant_val = 0
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

            proximo_f = obtener_numero_mopar_del_dia(carpeta_salida)
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
        self.title("PALAUSA APROVISIONAMIENTO PRO (MULTI-HILO)")
        self.geometry("850x950")
        self.resizable(False, False)
        if os.path.exists(ICONO_PATH):
            try: self.iconbitmap(ICONO_PATH)
            except: pass
        self.setup_ui()

    def aplicar_tema_inicial(self):
        modo = self.config.get("modo_apariencia", "Dark")
        tema = "Sinnoh"
        self.config["tema_color"] = "Sinnoh"
        ctk.set_appearance_mode(modo)
        mapa_temas = {
            "Azul": "blue",
            "Verde": "green",
            "Azul Oscuro": "dark-blue",
            "Kanto": "blue",
            "Johto": "green",
            "Hoenn": "blue",
            "Sinnoh": "dark-blue",
            "Unova": "dark-blue",
        }
        tema_base = mapa_temas.get(tema, "blue")
        ctk.set_default_color_theme(tema_base)
        self.accent_color = None
        self.hover_color = None

    def aplicar_estilos_ui(self):
        tema = "Sinnoh"
        paletas = {
            "Kanto": {
                "principal": "#d83a3a",
                "hover": "#a92a2a",
                "superficie": "#f1f4f8",
                "superficie_alt": "#dce4ec",
                "borde": "#7b8ca3",
                "texto": "#16212f",
                "texto_sec": "#506277",
                "panel_oscuro": "#1b2735",
                "panel_medio": "#243447",
                "panel_claro": "#eef3f8",
                "acento_azul": "#4aa3ff",
                "luz": "#ffd24a",
            },
            "Johto": {
                "principal": "#c59a2f",
                "hover": "#98751f",
                "superficie": "#f7f3e8",
                "superficie_alt": "#e8debf",
                "borde": "#8f8570",
                "texto": "#2a2418",
                "texto_sec": "#6e644e",
                "panel_oscuro": "#2d2a24",
                "panel_medio": "#4b4438",
                "panel_claro": "#f3ecdb",
                "acento_azul": "#9fd3ff",
                "luz": "#f6e27a",
            },
            "Hoenn": {
                "principal": "#b83a20",
                "hover": "#842414",
                "superficie": "#f8efe8",
                "superficie_alt": "#efd5c7",
                "borde": "#9a5a3c",
                "texto": "#2a170f",
                "texto_sec": "#7c4c38",
                "panel_oscuro": "#2f140d",
                "panel_medio": "#5b2415",
                "panel_claro": "#f6e6dd",
                "acento_azul": "#ff7a2f",
                "luz": "#ffd166",
            },
            "Sinnoh": {
                "principal": "#6f7fd8",
                "hover": "#515fab",
                "superficie": "#f1f2f8",
                "superficie_alt": "#dde1f0",
                "borde": "#7b84a2",
                "texto": "#1a2031",
                "texto_sec": "#5c6683",
                "panel_oscuro": "#1f2438",
                "panel_medio": "#313a57",
                "panel_claro": "#eceff8",
                "acento_azul": "#8ed0ff",
                "luz": "#f2d06b",
            },
            "Unova": {
                "principal": "#3b3b3b",
                "hover": "#252525",
                "superficie": "#f3f3f3",
                "superficie_alt": "#dfdfdf",
                "borde": "#8a8a8a",
                "texto": "#171717",
                "texto_sec": "#5a5a5a",
                "panel_oscuro": "#111111",
                "panel_medio": "#232323",
                "panel_claro": "#ededed",
                "acento_azul": "#66d1ff",
                "luz": "#d8d8d8",
            },
        }
        paleta = paletas.get(tema, paletas["Kanto"])
        self.color_principal = paleta["principal"]
        self.color_hover = paleta["hover"]
        self.color_superficie = paleta["superficie"]
        self.color_superficie_alt = paleta["superficie_alt"]
        self.color_borde = paleta["borde"]
        self.color_texto = paleta["texto"]
        self.color_texto_sec = paleta["texto_sec"]
        self.color_panel_oscuro = paleta["panel_oscuro"]
        self.color_panel_medio = paleta["panel_medio"]
        self.color_panel_claro = paleta["panel_claro"]
        self.color_acento_azul = paleta["acento_azul"]
        self.color_ok = "#2ecc71"
        self.color_warn = "#f39c12"
        self.color_peligro = "#d9534f"
        self.color_pokedex_luz = paleta["luz"]

    def crear_tarjeta_tab(self, tab, titulo, subtitulo):
        card = ctk.CTkFrame(tab, fg_color=self.color_superficie, corner_radius=18, border_width=2, border_color=self.color_borde)
        card.pack(padx=28, pady=18, fill="both", expand=True)
        header = ctk.CTkFrame(card, fg_color=self.color_panel_medio, corner_radius=14)
        header.pack(padx=14, pady=(14, 8), fill="x")
        led_row = ctk.CTkFrame(header, fg_color="transparent")
        led_row.pack(anchor="w", padx=14, pady=(10, 4))
        ctk.CTkLabel(led_row, text="●", text_color=self.color_principal, font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0, 4))
        ctk.CTkLabel(led_row, text="●", text_color=self.color_acento_azul, font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0, 4))
        ctk.CTkLabel(led_row, text="●", text_color=self.color_pokedex_luz, font=("Segoe UI", 14, "bold")).pack(side="left")
        ctk.CTkLabel(header, text=titulo.upper(), font=("Segoe UI", 18, "bold"), text_color="#f7fbff").pack(anchor="w", padx=14, pady=(0, 2))
        ctk.CTkLabel(header, text=subtitulo, font=("Segoe UI", 12), text_color="#b9c7d5").pack(anchor="w", padx=14, pady=(0, 12))
        return card

    def crear_fila_archivo(self, parent, label_text, entry_attr, config_key, placeholder, button_text, command):
        ctk.CTkLabel(parent, text=label_text, font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(anchor="w", padx=22, pady=(6, 6))
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(padx=22, pady=(0, 12), fill="x")
        entry = ctk.CTkEntry(row, placeholder_text=placeholder, height=38, corner_radius=12, border_width=2, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        entry.insert(0, self.config.get(config_key, ""))
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        setattr(self, entry_attr, entry)
        btn = ctk.CTkButton(row, text=button_text, width=78, height=38, corner_radius=12, fg_color=self.color_panel_medio, text_color="#f7fbff", hover_color="#17222e", command=command)
        btn.pack(side="right")
        return entry

    def crear_boton_principal(self, parent, text, command):
        return ctk.CTkButton(parent, text=text, height=46, corner_radius=14, font=("Segoe UI", 14, "bold"), fg_color=self.color_principal, text_color="#fff7f7", border_width=1, border_color="#7a1d1d", hover_color=self.color_hover, command=command)

    def obtener_icono_region_path(self):
        tema = self.config.get("tema_color", "Kanto").lower()
        ruta = REGION_ICON_FOLDER / f"{tema}.png"
        return ruta if ruta.exists() else None

    def setup_ui(self):
        self.aplicar_estilos_ui()
        acc_color = self.color_principal
        self.ultima_ruta_generada = ""
        self.configure(fg_color="#cfd9e3")
        hero = ctk.CTkFrame(self, fg_color=self.color_principal, corner_radius=22, border_width=2, border_color="#58648f")
        hero.pack(pady=(6, 2), padx=22, fill="x")
        hero_top = ctk.CTkFrame(hero, fg_color="transparent")
        hero_top.pack(fill="x", padx=16, pady=(8, 8))
        hero_lights = ctk.CTkFrame(hero_top, fg_color="transparent")
        hero_lights.pack(side="left")
        ctk.CTkLabel(hero_lights, text="◉", font=("Segoe UI", 24, "bold"), text_color="#8ed0ff").pack(side="left")
        ctk.CTkLabel(hero_lights, text="●", font=("Segoe UI", 18, "bold"), text_color=self.color_pokedex_luz).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(hero_lights, text="●", font=("Segoe UI", 18, "bold"), text_color="#ffffff").pack(side="left", padx=(6, 0))
        hero_title = ctk.CTkFrame(hero_top, fg_color="transparent")
        hero_title.pack(side="left", expand=True, padx=(18, 0))
        ctk.CTkLabel(hero_title, text="PALAUSA APROVISIONAMIENTO PRO", font=("Segoe UI", 22, "bold"), text_color="#fffaf7").pack(anchor="w")
        ruta_icono_region = self.obtener_icono_region_path()
        if ruta_icono_region:
            self.region_icon_img = ctk.CTkImage(light_image=Image.open(ruta_icono_region), dark_image=Image.open(ruta_icono_region), size=(72, 72))
            ctk.CTkLabel(hero_top, image=self.region_icon_img, text="").pack(side="right")
        self.tabview = ctk.CTkTabview(
            self,
            width=810,
            height=350,
            corner_radius=18,
            segmented_button_selected_color=acc_color,
            segmented_button_selected_hover_color=self.color_hover,
            fg_color=self.color_panel_claro,
            border_width=2,
            border_color=self.color_borde
        )
        self.tabview.pack(pady=2, padx=22, fill="x")
        self.tab_auto = self.tabview.add("Automatizacion")
        self.tab_pedidos = self.tabview.add("Pedidos")
        self.tab_comparador = self.tabview.add("Comparador")
        self.tab_rotacion = self.tabview.add("Filtro Rotacion")
        self.tab_mopar = self.tabview.add("Fichero Mopar")
        self.tab_stock_fiat = self.tabview.add("Stock FIAT")
        self.tab_reclamaciones_power_line = self.tabview.add("Reclamaciones")
        self.tab_config = self.tabview.add("Configuracion")
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
        self.frame_progreso = ctk.CTkFrame(self, fg_color=self.color_panel_medio, corner_radius=18, border_width=2, border_color=self.color_borde)
        self.frame_progreso.pack(pady=(2, 1), padx=22, fill="x")
        ctk.CTkLabel(self.frame_progreso, text="Estado de escaneo", font=("Segoe UI", 12, "bold"), text_color="#f4f6fb").pack(anchor="w", padx=16, pady=(10, 2))
        barra_row = ctk.CTkFrame(self.frame_progreso, fg_color="transparent")
        barra_row.pack(padx=16, pady=(0, 10), fill="x")
        self.progreso_var = ctk.DoubleVar()
        self.barra_prog = ctk.CTkProgressBar(barra_row, variable=self.progreso_var, height=16, corner_radius=999, progress_color=self.color_acento_azul, fg_color="#41556b")
        self.barra_prog.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.porcentaje_label = ctk.CTkLabel(barra_row, text="0%", font=("Segoe UI", 14, "bold"), text_color=self.color_pokedex_luz)
        self.porcentaje_label.pack(side="right")
        self.frame_estado = ctk.CTkFrame(self, fg_color=self.color_panel_medio, corner_radius=16, border_width=2, border_color=self.color_borde)
        self.frame_estado.pack(pady=(1, 1), padx=22, fill="x")
        ctk.CTkLabel(self.frame_estado, text="Centro de control", font=("Segoe UI", 13, "bold"), text_color="#f4f6fb").pack(anchor="w", padx=14, pady=(8, 1))
        self.estado_label = ctk.CTkLabel(self.frame_estado, text="Estado: listo", font=("Segoe UI", 12, "bold"), text_color=self.color_pokedex_luz)
        self.estado_label.pack(anchor="w", padx=14, pady=(1, 1))
        self.ruta_label = ctk.CTkLabel(self.frame_estado, text="Ultima salida: -", font=("Consolas", 10), text_color="#cfdae5", anchor="w", justify="left")
        self.ruta_label.pack(anchor="w", padx=14, pady=(0, 6))
        self.frame_acciones = ctk.CTkFrame(self.frame_estado, fg_color="transparent")
        self.frame_acciones.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Abrir ultima salida", height=30, corner_radius=10, fg_color="#35506a", text_color="#f8fbff", hover_color="#24384a", command=self.abrir_ultima_salida).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Abrir errores", height=30, corner_radius=10, fg_color="#35506a", text_color="#f8fbff", hover_color="#24384a", command=lambda: self.abrir_directorio(ERROR_FOLDER)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Abrir log", height=30, corner_radius=10, fg_color="#35506a", text_color="#f8fbff", hover_color="#24384a", command=self.abrir_log_detallado).pack(side="left", padx=(0, 8))
        ctk.CTkButton(self.frame_acciones, text="Limpiar log", height=30, corner_radius=10, fg_color="#35506a", text_color="#f8fbff", hover_color="#24384a", command=self.limpiar_log).pack(side="left")
        self.frame_log = ctk.CTkFrame(self, fg_color=self.color_panel_medio, corner_radius=18, border_width=2, border_color=self.color_borde)
        self.frame_log.pack(pady=(2, 10), padx=22, fill="both", expand=True)
        ctk.CTkLabel(self.frame_log, text="Registro de actividad", font=("Segoe UI", 14, "bold"), text_color="#f4f6fb").pack(anchor="w", padx=16, pady=(10, 6))
        self.log_box = ctk.CTkTextbox(self.frame_log, width=800, height=500, font=("Consolas", 13), border_width=2, border_color="#5a6d84", corner_radius=16, fg_color="#f8fbff", text_color="#102030")
        self.log_box.pack(padx=14, pady=(0, 14), fill="both", expand=True)
        self.log_box.tag_config("success", foreground="#2ecc71")
        self.log_box.tag_config("error", foreground="#e74c3c")

    def actualizar_paneles_por_pestana(self):
        ocultar_paneles = self.tabview.get() == "Reclamaciones"
        paneles = (self.frame_progreso, self.frame_estado, self.frame_log)
        if ocultar_paneles:
            for panel in paneles:
                panel.pack_forget()
            self.tabview.pack_configure(fill="both", expand=True, pady=(2, 10))
            return
        self.tabview.pack_configure(fill="x", expand=False, pady=2)
        if not self.frame_progreso.winfo_manager():
            self.frame_progreso.pack(pady=(2, 1), padx=22, fill="x", before=self.frame_estado)
        if not self.frame_estado.winfo_manager():
            self.frame_estado.pack(pady=(1, 1), padx=22, fill="x", before=self.frame_log)
        if not self.frame_log.winfo_manager():
            self.frame_log.pack(pady=(2, 10), padx=22, fill="both", expand=True)

    def setup_tab_auto(self, acc_color):
        ctk.CTkLabel(self.tab_auto, text="Consulta de Disponibilidad Citroen (Multi-Hilo)", font=("Segoe UI", 16, "bold")).pack(pady=10)
        row_file = ctk.CTkFrame(self.tab_auto, fg_color="transparent")
        row_file.pack(pady=10, fill="x", padx=50)
        self.excel_auto_entry = ctk.CTkEntry(row_file, placeholder_text="Excel de stock...", height=30)
        self.excel_auto_entry.insert(0, self.config.get("ruta_excel_auto", ""))
        self.excel_auto_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_file = ctk.CTkButton(row_file, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_auto_entry, [("Excel files", "*.xlsx")]))
        if self.hover_color:
            btn_file.configure(hover_color=self.hover_color)
        btn_file.pack(side="right")
        row_input = ctk.CTkFrame(self.tab_auto, fg_color="transparent")
        row_input.pack(pady=10)
        ctk.CTkLabel(row_input, text="Precio minimo (EUR):").grid(row=0, column=0, padx=10)
        self.precio_entry = ctk.CTkEntry(row_input, width=100, height=30, justify="center")
        self.precio_entry.insert(0, self.config.get("precio_min", "0"))
        self.precio_entry.grid(row=0, column=1, padx=5)
        self.btn_auto = ctk.CTkButton(self.tab_auto, text="INICIAR AUTOMATIZACION", height=45, font=("Segoe UI", 14, "bold"), fg_color=acc_color, command=self.toggle_auto)
        if self.hover_color:
            self.btn_auto.configure(hover_color=self.hover_color)
        self.btn_auto.pack(pady=15, padx=100, fill="x")

    def setup_tab_pedidos(self):
        acc_color = self.accent_color if self.accent_color else "#28a745"
        ctk.CTkLabel(self.tab_pedidos, text="Generacion de CSV Urgentes", font=("Segoe UI", 16, "bold")).pack(pady=10)
        row_file = ctk.CTkFrame(self.tab_pedidos, fg_color="transparent")
        row_file.pack(pady=10, fill="x", padx=50)
        self.excel_pedidos_entry = ctk.CTkEntry(row_file, placeholder_text="Excel URGENTE.xlsx...", height=30)
        self.excel_pedidos_entry.insert(0, self.config.get("ruta_excel_pedidos", ""))
        self.excel_pedidos_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_file = ctk.CTkButton(row_file, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_pedidos_entry, [("Excel files", "*.xlsx")]))
        btn_file.pack(side="right")
        row_folder = ctk.CTkFrame(self.tab_pedidos, fg_color="transparent")
        row_folder.pack(pady=10, fill="x", padx=50)
        self.folder_pedidos_entry = ctk.CTkEntry(row_folder, placeholder_text="Carpeta de destino...", height=30)
        self.folder_pedidos_entry.insert(0, self.config.get("ruta_folder_pedidos", ""))
        self.folder_pedidos_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_folder = ctk.CTkButton(row_folder, text="Carpeta", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_carpeta(self.folder_pedidos_entry))
        btn_folder.pack(side="right")
        self.btn_pedidos = ctk.CTkButton(self.tab_pedidos, text="GENERAR PEDIDOS", height=45, font=("Segoe UI", 14, "bold"), fg_color=acc_color, command=self.toggle_pedidos)
        self.btn_pedidos.pack(pady=15, padx=100, fill="x")

    def setup_tab_comparador(self, acc_color):
        ctk.CTkLabel(self.tab_comparador, text="Comparar Archivos Excel", font=("Segoe UI", 16, "bold")).pack(pady=10)
        row_excel_old = ctk.CTkFrame(self.tab_comparador, fg_color="transparent")
        row_excel_old.pack(pady=5, fill="x", padx=50)
        self.excel_old_entry = ctk.CTkEntry(row_excel_old, placeholder_text="Excel Antiguo...", height=30)
        self.excel_old_entry.insert(0, self.config.get("ruta_excel_antiguo", ""))
        self.excel_old_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_excel_old, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_old_entry, [("Excel files", "*.xlsx")])).pack(side="right")
        row_excel_new = ctk.CTkFrame(self.tab_comparador, fg_color="transparent")
        row_excel_new.pack(pady=5, fill="x", padx=50)
        self.excel_new_entry = ctk.CTkEntry(row_excel_new, placeholder_text="Excel Nuevo...", height=30)
        self.excel_new_entry.insert(0, self.config.get("ruta_excel_nuevo", ""))
        self.excel_new_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_excel_new, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_new_entry, [("Excel files", "*.xlsx")])).pack(side="right")
        row_folder_comp = ctk.CTkFrame(self.tab_comparador, fg_color="transparent")
        row_folder_comp.pack(pady=5, fill="x", padx=50)
        self.folder_comp_entry = ctk.CTkEntry(row_folder_comp, placeholder_text="Carpeta destino...", height=30)
        self.folder_comp_entry.insert(0, self.config.get("ruta_folder_comparador", ""))
        self.folder_comp_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_folder_comp, text="Carpeta", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_carpeta(self.folder_comp_entry)).pack(side="right")
        row_color = ctk.CTkFrame(self.tab_comparador, fg_color="transparent")
        row_color.pack(pady=10)
        self.color_optionmenu = ctk.CTkOptionMenu(row_color, values=["Amarillo", "Rojo", "Verde", "Azul", "Naranja"], height=30, fg_color=acc_color, button_color=acc_color)
        self.color_optionmenu.set(self.config.get("color_resaltado", "Amarillo"))
        self.color_optionmenu.grid(row=0, column=1, padx=5)
        self.btn_comparador = ctk.CTkButton(self.tab_comparador, text="COMPARAR EXCELS", height=45, font=("Segoe UI", 14, "bold"), fg_color=acc_color, command=self.toggle_comparador)
        self.btn_comparador.pack(pady=10, padx=100, fill="x")

    def setup_tab_rotacion(self, acc_color):
        ctk.CTkLabel(self.tab_rotacion, text="Analisis de Consumos (Rotacion)", font=("Segoe UI", 16, "bold")).pack(pady=10)
        row_quiter = ctk.CTkFrame(self.tab_rotacion, fg_color="transparent")
        row_quiter.pack(pady=5, fill="x", padx=50)
        self.excel_quiter_entry = ctk.CTkEntry(row_quiter, placeholder_text="Excel de Quiter (DMS)...", height=30)
        self.excel_quiter_entry.insert(0, self.config.get("ruta_excel_quiter", ""))
        self.excel_quiter_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_quiter, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_quiter_entry, [("Excel files", "*.xlsx")])).pack(side="right")
        row_rot = ctk.CTkFrame(self.tab_rotacion, fg_color="transparent")
        row_rot.pack(pady=5, fill="x", padx=50)
        self.excel_rot_entry = ctk.CTkEntry(row_rot, placeholder_text="Archivo PROPUESTAPEDIDO.xls...", height=30)
        self.excel_rot_entry.insert(0, self.config.get("ruta_excel_rotacion", ""))
        self.excel_rot_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_rot, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_rot_entry, [("Excel files", "*.xls *.xlsx")])).pack(side="right")
        row_folder_rot = ctk.CTkFrame(self.tab_rotacion, fg_color="transparent")
        row_folder_rot.pack(pady=5, fill="x", padx=50)
        self.folder_rot_entry = ctk.CTkEntry(row_folder_rot, placeholder_text="Carpeta destino...", height=30)
        self.folder_rot_entry.insert(0, self.config.get("ruta_folder_rotacion", ""))
        self.folder_rot_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_folder_rot, text="Carpeta", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_carpeta(self.folder_rot_entry)).pack(side="right")
        row_umbral = ctk.CTkFrame(self.tab_rotacion, fg_color="transparent")
        row_umbral.pack(pady=10)
        ctk.CTkLabel(row_umbral, text="Umbral media minima:").grid(row=0, column=0, padx=10)
        self.umbral_entry = ctk.CTkEntry(row_umbral, width=80, justify="center")
        self.umbral_entry.insert(0, self.config.get("umbral_rotacion", "0.08"))
        self.umbral_entry.grid(row=0, column=1)
        self.btn_rotacion = ctk.CTkButton(self.tab_rotacion, text="RESALTAR POR ROTACION", height=45, font=("Segoe UI", 14, "bold"), fg_color=acc_color, command=self.iniciar_filtro_rotacion)
        self.btn_rotacion.pack(pady=15, padx=100, fill="x")

    def setup_tab_mopar(self, acc_color):
        ctk.CTkLabel(self.tab_mopar, text="Conversion a Fichero MOPAR", font=("Segoe UI", 16, "bold")).pack(pady=10)
        row_file = ctk.CTkFrame(self.tab_mopar, fg_color="transparent")
        row_file.pack(pady=10, fill="x", padx=50)
        self.excel_mopar_entry = ctk.CTkEntry(row_file, placeholder_text="Excel de entrada (.xls)...", height=30)
        self.excel_mopar_entry.insert(0, self.config.get("ruta_excel_mopar", ""))
        self.excel_mopar_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_file, text="Archivo", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_archivo(self.excel_mopar_entry, [("Excel files", "*.xls *.xlsx")])).pack(side="right")
        row_folder = ctk.CTkFrame(self.tab_mopar, fg_color="transparent")
        row_folder.pack(pady=10, fill="x", padx=50)
        self.folder_mopar_entry = ctk.CTkEntry(row_folder, placeholder_text="Carpeta de destino...", height=30)
        self.folder_mopar_entry.insert(0, self.config.get("ruta_folder_mopar", ""))
        self.folder_mopar_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row_folder, text="Carpeta", width=70, height=30, fg_color=acc_color, command=lambda: self.seleccionar_carpeta(self.folder_mopar_entry)).pack(side="right")
        self.btn_mopar = ctk.CTkButton(self.tab_mopar, text="GENERAR FICHERO MOPAR", height=45, font=("Segoe UI", 14, "bold"), fg_color=acc_color, command=self.toggle_mopar)
        self.btn_mopar.pack(pady=15, padx=100, fill="x")

    def setup_tab_config(self, acc_color):
        ctk.CTkLabel(self.tab_config, text="Credenciales", font=("Segoe UI", 16, "bold")).pack(pady=10)
        frame_creds = ctk.CTkFrame(self.tab_config)
        frame_creds.pack(pady=5, padx=50, fill="x")
        ctk.CTkLabel(frame_creds, text="Usuario:").grid(row=0, column=0, padx=20, pady=10)
        self.user_entry = ctk.CTkEntry(frame_creds, width=200, height=30)
        self.user_entry.insert(0, self.config.get("usuario", ""))
        self.user_entry.grid(row=0, column=1, padx=20, pady=10)
        ctk.CTkLabel(frame_creds, text="Contrasena:").grid(row=1, column=0, padx=20, pady=10)
        self.pass_entry = ctk.CTkEntry(frame_creds, width=200, height=30, show="*")
        self.pass_entry.insert(0, self.config.get("password", ""))
        self.pass_entry.grid(row=1, column=1, padx=20, pady=10)

        ctk.CTkLabel(self.tab_config, text="Rendimiento", font=("Segoe UI", 16, "bold")).pack(pady=(15, 5))
        frame_perf = ctk.CTkFrame(self.tab_config)
        frame_perf.pack(pady=5, padx=50, fill="x")
        ctk.CTkLabel(frame_perf, text="Numero de hilos paralelos:").grid(row=0, column=0, padx=20, pady=10)
        self.hilos_option = ctk.CTkOptionMenu(frame_perf, values=["1", "2", "3", "4", "5"], height=30, fg_color=acc_color, button_color=acc_color)
        self.hilos_option.set(str(self.config.get("num_hilos", "3")))
        self.hilos_option.grid(row=0, column=1, padx=20, pady=10)

        ctk.CTkLabel(self.tab_config, text="Personalizacion", font=("Segoe UI", 16, "bold")).pack(pady=(15, 5))
        frame_temas = ctk.CTkFrame(self.tab_config)
        frame_temas.pack(pady=5, padx=50, fill="x")
        ctk.CTkLabel(frame_temas, text="Modo:").grid(row=0, column=0, padx=20, pady=5)
        self.modo_option = ctk.CTkOptionMenu(frame_temas, values=["Dark", "Light", "System"], height=30, fg_color=acc_color, button_color=acc_color)
        self.modo_option.set(self.config.get("modo_apariencia", "Dark"))
        self.modo_option.grid(row=0, column=1, padx=20, pady=5)
        ctk.CTkLabel(frame_temas, text="Tema:").grid(row=1, column=0, padx=20, pady=5)
        self.tema_option = ctk.CTkOptionMenu(frame_temas, values=["Azul", "Verde", "Azul Oscuro"], height=30, fg_color=acc_color, button_color=acc_color)
        tema_actual = self.config.get("tema_color", "Azul")
        if tema_actual not in ["Azul", "Verde", "Azul Oscuro"]:
            tema_actual = "Azul"
        self.tema_option.set(tema_actual)
        self.tema_option.grid(row=1, column=1, padx=20, pady=5)
        ctk.CTkButton(self.tab_config, text="GUARDAR CONFIGURACION", height=40, fg_color=acc_color, command=self.guardar_config_ui).pack(pady=20)

    def build_tab_auto(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_auto, "Consulta de disponibilidad", "Carga tu Excel y ejecuta la consulta paralela contra el portal.")
        self.crear_fila_archivo(card, "Excel de stock", "excel_auto_entry", "ruta_excel_auto", "Excel de stock...", "Examinar", lambda: self.seleccionar_archivo(self.excel_auto_entry, [("Excel files", "*.xlsx")]))
        row_input = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        row_input.pack(padx=22, pady=(2, 14), fill="x")
        ctk.CTkLabel(row_input, text="Precio minimo (EUR)", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.precio_entry = ctk.CTkEntry(row_input, width=120, height=36, justify="center", corner_radius=12, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.precio_entry.insert(0, self.config.get("precio_min", "0"))
        self.precio_entry.pack(side="right", padx=14, pady=12)
        self.btn_auto = self.crear_boton_principal(card, "EJECUTAR AUTOMATIZACION PARALELA", self.toggle_auto)
        self.btn_auto.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_pedidos(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_pedidos, "Generacion de pedidos", "Convierte el Excel urgente en lotes CSV y guarda copia de seguridad antes de limpiar.")
        self.crear_fila_archivo(card, "Excel urgente", "excel_pedidos_entry", "ruta_excel_pedidos", "Excel URGENTE.xlsx...", "Examinar", lambda: self.seleccionar_archivo(self.excel_pedidos_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_pedidos_entry", "ruta_folder_pedidos", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_pedidos_entry))
        self.btn_pedidos = self.crear_boton_principal(card, "GENERAR PEDIDOS", self.toggle_pedidos)
        self.btn_pedidos.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_comparador(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_comparador, "Comparador de Excel", "Resalta las lineas repetidas y genera un archivo comparado en la carpeta elegida.")
        self.crear_fila_archivo(card, "Excel antiguo", "excel_old_entry", "ruta_excel_antiguo", "Excel antiguo...", "Examinar", lambda: self.seleccionar_archivo(self.excel_old_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Excel nuevo", "excel_new_entry", "ruta_excel_nuevo", "Excel nuevo...", "Examinar", lambda: self.seleccionar_archivo(self.excel_new_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_comp_entry", "ruta_folder_comparador", "Carpeta destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_comp_entry))
        row_color = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        row_color.pack(padx=22, pady=(2, 14), fill="x")
        ctk.CTkLabel(row_color, text="Color de resaltado", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.color_optionmenu = ctk.CTkOptionMenu(row_color, values=["Amarillo", "Rojo", "Verde", "Azul", "Naranja"], height=36, corner_radius=12, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.color_optionmenu.set(self.config.get("color_resaltado", "Amarillo"))
        self.color_optionmenu.pack(side="right", padx=14, pady=12)
        self.btn_comparador = self.crear_boton_principal(card, "COMPARAR EXCELS", self.toggle_comparador)
        self.btn_comparador.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_rotacion(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_rotacion, "Filtro de rotacion", "Marca referencias de baja rotacion en base al umbral y la propuesta de pedido.")
        self.crear_fila_archivo(card, "Excel de Quiter", "excel_quiter_entry", "ruta_excel_quiter", "Excel de Quiter (DMS)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_quiter_entry, [("Excel files", "*.xlsx")]))
        self.crear_fila_archivo(card, "Archivo de propuesta", "excel_rot_entry", "ruta_excel_rotacion", "Archivo PROPUESTAPEDIDO.xls...", "Examinar", lambda: self.seleccionar_archivo(self.excel_rot_entry, [("Excel files", "*.xls *.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_rot_entry", "ruta_folder_rotacion", "Carpeta destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_rot_entry))
        row_umbral = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        row_umbral.pack(padx=22, pady=(2, 14), fill="x")
        ctk.CTkLabel(row_umbral, text="Umbral minimo", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.umbral_entry = ctk.CTkEntry(row_umbral, width=120, justify="center", height=36, corner_radius=12, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.umbral_entry.insert(0, self.config.get("umbral_rotacion", "0.08"))
        self.umbral_entry.pack(side="right", padx=14, pady=12)
        self.btn_rotacion = self.crear_boton_principal(card, "RESALTAR POR ROTACION", self.iniciar_filtro_rotacion)
        self.btn_rotacion.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_mopar(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_mopar, "Generador Mopar", "Transforma el Excel en ficheros TXT Mopar con numeracion diaria persistente.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_mopar_entry", "ruta_excel_mopar", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_mopar_entry, [("Excel files", "*.xls *.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_mopar_entry", "ruta_folder_mopar", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_mopar_entry))
        self.btn_mopar = self.crear_boton_principal(card, "GENERAR FICHERO MOPAR", self.toggle_mopar)
        self.btn_mopar.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_stock_fiat(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_stock_fiat, "Stock FIAT", "Transforma el Excel en ficheros TXT con el mismo formato final de Mopar.")
        self.crear_fila_archivo(card, "Excel de entrada", "excel_stock_fiat_entry", "ruta_excel_stock_fiat", "Excel de entrada (.xls)...", "Examinar", lambda: self.seleccionar_archivo(self.excel_stock_fiat_entry, [("Excel files", "*.xls *.xlsx")]))
        self.crear_fila_archivo(card, "Carpeta destino", "folder_stock_fiat_entry", "ruta_folder_stock_fiat", "Carpeta de destino...", "Carpeta", lambda: self.seleccionar_carpeta(self.folder_stock_fiat_entry))
        self.btn_stock_fiat = self.crear_boton_principal(card, "GENERAR STOCK FIAT", self.toggle_stock_fiat)
        self.btn_stock_fiat.pack(padx=22, pady=(4, 22), fill="x")

    def build_tab_config(self, acc_color):
        card = self.crear_tarjeta_tab(self.tab_config, "Configuracion general", "Ajusta credenciales, rendimiento y apariencia visual de la herramienta.")
        frame_creds = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        frame_creds.pack(padx=22, pady=(0, 14), fill="x")
        ctk.CTkLabel(frame_creds, text="Usuario", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=0, column=0, padx=18, pady=12, sticky="w")
        self.user_entry = ctk.CTkEntry(frame_creds, width=220, height=36, corner_radius=12, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto)
        self.user_entry.insert(0, self.config.get("usuario", ""))
        self.user_entry.grid(row=0, column=1, padx=18, pady=12, sticky="ew")
        ctk.CTkLabel(frame_creds, text="Contrasena", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=1, column=0, padx=18, pady=12, sticky="w")
        self.pass_entry = ctk.CTkEntry(frame_creds, width=220, height=36, corner_radius=12, border_color=self.color_borde, fg_color="#ffffff", text_color=self.color_texto, show="*")
        self.pass_entry.insert(0, self.config.get("password", ""))
        self.pass_entry.grid(row=1, column=1, padx=18, pady=12, sticky="ew")
        frame_creds.grid_columnconfigure(1, weight=1)
        frame_perf = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        frame_perf.pack(padx=22, pady=(0, 14), fill="x")
        ctk.CTkLabel(frame_perf, text="Hilos paralelos", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).pack(side="left", padx=(16, 12), pady=12)
        self.hilos_option = ctk.CTkOptionMenu(frame_perf, values=["1", "2", "3", "4", "5"], height=36, corner_radius=12, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.hilos_option.set(str(self.config.get("num_hilos", "3")))
        self.hilos_option.pack(side="right", padx=14, pady=12)
        frame_temas = ctk.CTkFrame(card, fg_color=self.color_superficie_alt, corner_radius=14)
        frame_temas.pack(padx=22, pady=(0, 14), fill="x")
        ctk.CTkLabel(frame_temas, text="Region", font=("Segoe UI", 12, "bold"), text_color=self.color_texto).grid(row=0, column=0, padx=18, pady=12, sticky="w")
        self.tema_option = ctk.CTkOptionMenu(frame_temas, values=["Sinnoh"], height=36, corner_radius=12, fg_color=acc_color, button_color=acc_color, button_hover_color=self.color_hover)
        self.tema_option.set("Sinnoh")
        self.tema_option.grid(row=0, column=1, padx=18, pady=12, sticky="ew")
        frame_temas.grid_columnconfigure(1, weight=1)
        self.reclamaciones_power_line.construir_ajustes(card)
        ctk.CTkButton(card, text="GUARDAR CONFIGURACION", height=42, corner_radius=14, fg_color=acc_color, hover_color=self.color_hover, command=self.guardar_config_ui).pack(padx=22, pady=(4, 22), fill="x")

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
            self.ruta_label.configure(text=f"?ltima salida: {ruta}")

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
            messagebox.showinfo("Informacion", "Todavia no hay una salida generada.")
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
            messagebox.showinfo("Exito", "Configuracion guardada.")

    def toggle_auto(self):
        global automation_running
        if automation_running:
            stop_event_auto.set()
            self.btn_auto.configure(text="DETENIENDO...", state="disabled")
        else:
            self.iniciar_auto()

    def iniciar_auto(self):
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
        self.actualizar_estado("automatizacion en curso", ruta)
        self.btn_auto.configure(text="DETENER", fg_color="red")
        stop_event_auto.clear()
        threading.Thread(target=ejecutar_automatizacion_task, args=(ruta, user, pwd, precio, hilos, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_auto), stop_event_auto), daemon=True).start()

    def finalizar_auto(self):
        acc_color = self.color_principal
        self.actualizar_estado("automatizacion finalizada", self.excel_auto_entry.get())
        self.btn_auto.configure(text="INICIAR AUTOMATIZACION", fg_color=acc_color, state="normal")

    def toggle_pedidos(self):
        global pedidos_running
        if not pedidos_running:
            self.iniciar_pedidos()
        else:
            stop_event_pedidos.set()

    def iniciar_pedidos(self):
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_pedidos_entry.get()
        folder = self.folder_pedidos_entry.get()
        if not self.validar_archivo(ruta, "Selecciona un Excel valido para generar pedidos."):
            return
        if not self.validar_carpeta(folder, "Selecciona una carpeta de destino para los pedidos."):
            return
        self.actualizar_estado("generando pedidos", folder)
        self.btn_pedidos.configure(text="DETENER", fg_color="red")
        stop_event_pedidos.clear()
        threading.Thread(target=generar_pedidos_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_pedidos), stop_event_pedidos), daemon=True).start()

    def finalizar_pedidos(self):
        acc_color = self.color_principal
        self.actualizar_estado("pedidos generados", self.folder_pedidos_entry.get())
        self.btn_pedidos.configure(text="GENERAR PEDIDOS", fg_color=acc_color, state="normal")

    def toggle_comparador(self):
        global comparador_running
        if not comparador_running:
            self.iniciar_comparador()
        else:
            stop_event_comparador.set()

    def iniciar_comparador(self):
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
        self.btn_comparador.configure(text="DETENER", fg_color="red")
        stop_event_comparador.clear()
        threading.Thread(target=comparar_excel_task, args=(ruta_ant, ruta_nue, ruta_salida, colores_hex[color], lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_comparador), stop_event_comparador), daemon=True).start()

    def finalizar_comparador(self):
        acc_color = self.color_principal
        self.actualizar_estado("comparacion finalizada", str(Path(self.folder_comp_entry.get()) / "resultado_comparacion.xlsx"))
        self.btn_comparador.configure(text="COMPARAR EXCELS", fg_color=acc_color, state="normal")

    def iniciar_filtro_rotacion(self):
        self.guardar_configuracion_ui_internal()
        ruta_quiter = self.excel_quiter_entry.get()
        ruta_rot = self.excel_rot_entry.get()
        folder_salida = self.folder_rot_entry.get()
        try:
            umbral = float(self.umbral_entry.get())
        except Exception:
            messagebox.showwarning("Aviso", "Introduce un umbral de rotacion valido.")
            return
        if not self.validar_archivo(ruta_quiter, "Selecciona el Excel de Quiter."):
            return
        if not self.validar_archivo(ruta_rot, "Selecciona el archivo de propuesta."):
            return
        if not self.validar_carpeta(folder_salida, "Selecciona una carpeta de salida para rotacion."):
            return
        ruta_estimacion = str(Path(folder_salida) / f"{Path(ruta_quiter).stem}_ROTACION.xlsx")
        self.actualizar_estado("analizando rotacion", ruta_estimacion)
        self.btn_rotacion.configure(state="disabled", text="ANALIZANDO...")
        threading.Thread(target=filtrar_por_rotacion_task, args=(ruta_quiter, ruta_rot, folder_salida, umbral, "FF0000", lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_filtro_rotacion)), daemon=True).start()

    def finalizar_filtro_rotacion(self):
        ruta_quiter = self.excel_quiter_entry.get()
        if self.folder_rot_entry.get() and ruta_quiter:
            ruta_final = str(Path(self.folder_rot_entry.get()) / f"{Path(ruta_quiter).stem}_ROTACION.xlsx")
            self.actualizar_estado("rotacion completada", ruta_final)
        self.btn_rotacion.configure(state="normal", text="RESALTAR POR ROTACION")

    def toggle_mopar(self):
        global mopar_running
        if mopar_running:
            stop_event_mopar.set()
            self.btn_mopar.configure(text="DETENIENDO...", state="disabled")
        else:
            self.iniciar_mopar()

    def iniciar_mopar(self):
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_mopar_entry.get()
        folder = self.folder_mopar_entry.get()
        if not self.validar_archivo(ruta, "Por favor selecciona un archivo Excel valido."):
            return
        if not self.validar_carpeta(folder, "Por favor selecciona una carpeta de destino."):
            return
        self.actualizar_estado("generando fichero Mopar", folder)
        if not os.path.exists(ruta):
            messagebox.showwarning("Aviso", "Por favor selecciona un archivo Excel valido.")
            return
        if not os.path.exists(folder):
            messagebox.showwarning("Aviso", "Por favor selecciona una carpeta de destino.")
            return
        self.btn_mopar.configure(text="DETENER", fg_color="red")
        stop_event_mopar.clear()
        threading.Thread(target=generar_fichero_mopar_task, args=(ruta, folder, lambda t, tag=None: self.after(0, self.agregar_log, t, tag), lambda v: self.after(0, self.actualizar_progreso, v), lambda: self.after(0, self.finalizar_mopar), stop_event_mopar), daemon=True).start()

    def toggle_stock_fiat(self):
        global stock_fiat_running
        if stock_fiat_running:
            stop_event_stock_fiat.set()
            self.btn_stock_fiat.configure(text="DETENIENDO...", state="disabled")
        else:
            self.iniciar_stock_fiat()

    def iniciar_stock_fiat(self):
        self.guardar_configuracion_ui_internal()
        ruta = self.excel_stock_fiat_entry.get()
        folder = self.folder_stock_fiat_entry.get()
        if not self.validar_archivo(ruta, "Por favor selecciona un archivo Excel valido para Stock FIAT."):
            return
        if not self.validar_carpeta(folder, "Por favor selecciona una carpeta de destino para Stock FIAT."):
            return
        self.actualizar_estado("generando Stock FIAT", folder)
        if not os.path.exists(ruta):
            messagebox.showwarning("Aviso", "Por favor selecciona un archivo Excel valido para Stock FIAT.")
            return
        if not os.path.exists(folder):
            messagebox.showwarning("Aviso", "Por favor selecciona una carpeta de destino para Stock FIAT.")
            return
        self.btn_stock_fiat.configure(text="DETENER", fg_color="red")
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
            "tema_color": self.widget_get("tema_option", self.config.get("tema_color", "Azul")),
            "mopar_fecha": self.config.get("mopar_fecha"),
            "mopar_numero": self.config.get("mopar_numero")
        }
        if hasattr(self, "reclamaciones_power_line"):
            config.update(self.reclamaciones_power_line.obtener_configuracion())
            self.config.update(config)
        guardar_configuracion(config)

    def finalizar_mopar(self):
        acc_color = self.color_principal
        self.actualizar_estado("fichero Mopar generado", self.folder_mopar_entry.get())
        self.btn_mopar.configure(text="📄 GENERAR FICHERO MOPAR", fg_color=acc_color, state="normal")

    def finalizar_stock_fiat(self):
        acc_color = self.color_principal
        self.actualizar_estado("Stock FIAT generado", self.folder_stock_fiat_entry.get())
        self.btn_stock_fiat.configure(text="GENERAR STOCK FIAT", fg_color=acc_color, state="normal")

if __name__ == "__main__":
    app = AppIntegrada()
    app.mainloop()

"""Lanzador mínimo: mantiene la ruta histórica dist\app_automatizacion_optimizada.exe.

Los accesos directos y la tarea programada de Windows siguen apuntando a ese archivo; este lanzador
abre el programa real (carpeta app_automatizacion_optimizada\\) reenviando los argumentos."""
import os
import subprocess
import sys

carpeta = os.path.dirname(os.path.abspath(sys.executable))
destino = os.path.join(carpeta, "programa", "app_automatizacion_optimizada.exe")
argumentos = sys.argv[1:]

if not os.path.exists(destino):
    sys.exit(2)

if argumentos:
    # Ejecución desatendida (tarea programada): se espera al programa para devolver su código de salida
    sys.exit(subprocess.call([destino] + argumentos, cwd=carpeta))
subprocess.Popen([destino], cwd=carpeta)

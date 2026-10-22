"""Lanzador mínimo: mantiene la ruta histórica dist\app_automatizacion_optimizada.exe.

Los accesos directos y la tarea programada de Windows siguen apuntando a ese archivo; este lanzador
abre el programa real (carpeta programa\\) reenviando los argumentos. Antes, si la app dejó preparada
una versión nueva, la pone en su sitio (cambio_version.py)."""
import os
import subprocess
import sys

import cambio_version

carpeta = os.path.dirname(os.path.abspath(sys.executable))
cambio_version.al_arrancar_lanzador(carpeta)
destino = os.path.join(carpeta, cambio_version.NOMBRE_PROGRAMA, cambio_version.NOMBRE_EJECUTABLE)
argumentos = sys.argv[1:]

if not os.path.exists(destino):
    sys.exit(2)

if argumentos:
    # Ejecución desatendida (tarea programada): se espera al programa para devolver su código de salida
    sys.exit(subprocess.call([destino] + argumentos, cwd=carpeta))
subprocess.Popen([destino], cwd=carpeta)

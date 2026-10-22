"""Publica una versión nueva de la app en GitHub Releases, de donde la descargan las demás instalaciones.

Uso:
    python publicar_version.py 1.1.0 --notas "Qué cambia en esta versión"
    python publicar_version.py 1.1.0 --notas "..." --sin-publicar   (solo prueba: compila y crea el zip)

Pasos: comprueba que el repositorio está limpio, pone la versión en version.py, pasa las pruebas,
compila programa y lanzador, comprueba que el navegador incluido arranca, crea el zip, hace commit y
etiqueta vX.Y.Z, sube a GitHub y crea la release con el zip. Si un paso falla, se detiene ahí.
Necesita GitHub CLI con sesión iniciada (gh auth login)."""
import argparse
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
DIST = RAIZ / "dist"
SALIDA = RAIZ / "releases"
ARCHIVO_VERSION = RAIZ / "version.py"
EJECUTABLE = "app_automatizacion_optimizada.exe"

LEEME = """Palausa Aprovisionamiento {version}
==============================

INSTALACIÓN (solo la primera vez)
1. Descomprime este zip en una carpeta de ruta corta, por ejemplo C:\\Palausa
   (no en una ruta muy larga: el navegador incluido tiene archivos con nombres largos).
2. Abre app_automatizacion_optimizada.exe. Si Windows muestra «Windows protegió su PC»,
   pulsa «Más información» y «Ejecutar de todas formas» (la app no está firmada).
3. En la pestaña Configuración rellena tu usuario y contraseña del portal, y en Facturas
   y Pedidos los correos que uses. Para enviar correos hace falta Outlook instalado.
4. Si quieres, crea un acceso directo a app_automatizacion_optimizada.exe en el escritorio.

ACTUALIZACIONES
La app avisa sola cuando hay una versión nueva (botón «Actualizar a ...» arriba a la derecha).
Al aceptar se descarga y se instala al reiniciar. Tu configuración, tus registros y tu base de
datos están fuera de la carpeta «programa» y no se tocan nunca.

No borres ni muevas la carpeta «programa»: es la app.
"""


def ejecutar(orden, **kw):
    print("  $", " ".join(str(p) for p in orden), flush=True)
    resultado = subprocess.run([str(p) for p in orden], cwd=RAIZ, **kw)
    if resultado.returncode != 0:
        sys.exit(f"ERROR: falló el paso anterior (código {resultado.returncode}).")
    return resultado


def salida_de(orden):
    return subprocess.run([str(p) for p in orden], cwd=RAIZ, capture_output=True, text=True, encoding="utf-8").stdout.strip()


def ruta_gh():
    return shutil.which("gh") or r"C:\Program Files\GitHub CLI\gh.exe"


def version_tupla(texto):
    return tuple(int(n) for n in texto.split("."))


def version_actual():
    return re.search(r'VERSION = "([^"]+)"', ARCHIVO_VERSION.read_text(encoding="utf-8")).group(1)


def poner_version(version):
    texto = ARCHIVO_VERSION.read_text(encoding="utf-8")
    ARCHIVO_VERSION.write_text(re.sub(r'VERSION = "[^"]+"', f'VERSION = "{version}"', texto), encoding="utf-8")


def programa_abierto():
    lista = salida_de(["tasklist", "/FI", f"IMAGENAME eq {EJECUTABLE}", "/NH"])
    return EJECUTABLE.lower() in lista.lower()


def crear_zip(version):
    SALIDA.mkdir(exist_ok=True)
    destino = SALIDA / f"Palausa_Aprovisionamiento_{version}.zip"
    if destino.exists():
        destino.unlink()
    programa = DIST / "programa"
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as paquete:
        paquete.write(DIST / EJECUTABLE, EJECUTABLE)
        paquete.writestr("LEEME.txt", LEEME.format(version=version))
        for archivo in sorted(programa.rglob("*")):
            if archivo.is_file():
                paquete.write(archivo, "programa/" + archivo.relative_to(programa).as_posix())
    return destino


def main():
    parser = argparse.ArgumentParser(description="Publica una versión nueva en GitHub Releases.")
    parser.add_argument("version", help="número de versión, por ejemplo 1.1.0")
    parser.add_argument("--notas", required=True, help="qué cambia en esta versión (lo verá quien actualice)")
    parser.add_argument("--sin-publicar", action="store_true", help="compila y crea el zip, pero no sube nada")
    args = parser.parse_args()
    version, etiqueta = args.version, f"v{args.version}"

    print("1. Comprobaciones")
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        sys.exit("ERROR: la versión debe tener la forma 1.2.3")
    if version_tupla(version) < version_tupla(version_actual()):
        sys.exit(f"ERROR: {version} es anterior a la versión actual {version_actual()}")
    if salida_de(["git", "status", "--porcelain", "--untracked-files=no"]):
        sys.exit("ERROR: hay cambios sin guardar en git. Haz commit antes de publicar.")
    if salida_de(["git", "branch", "--show-current"]) != "main":
        sys.exit("ERROR: hay que publicar desde la rama main.")
    if salida_de(["git", "tag", "-l", etiqueta]):
        sys.exit(f"ERROR: la versión {version} ya está publicada (existe la etiqueta {etiqueta}).")
    if programa_abierto():
        sys.exit("ERROR: cierra la app antes de compilar (la carpeta dist\\programa está en uso).")
    if not args.sin_publicar and subprocess.run([ruta_gh(), "auth", "status"], capture_output=True).returncode != 0:
        sys.exit("ERROR: GitHub CLI no tiene sesión iniciada. Ejecuta: gh auth login")

    anterior = version_actual()
    poner_version(version)
    try:
        print("2. Pruebas")
        ejecutar([sys.executable, "-m", "pytest", "-q", "tests"])
        print("3. Compilación")
        ejecutar([sys.executable, "-m", "PyInstaller", "--noconfirm", "app_automatizacion_optimizada.spec"])
        ejecutar([sys.executable, "-m", "PyInstaller", "--noconfirm", "--workpath", "build_lanzador", "lanzador.spec"])
        print("4. Navegador incluido")
        diagnostico = DIST / "diagnostico_navegador.txt"
        if diagnostico.exists():
            diagnostico.unlink()
        ejecutar([DIST / "programa" / EJECUTABLE, "--probar-navegador"])
        if not diagnostico.exists() or not diagnostico.read_text(encoding="utf-8").startswith("OK"):
            sys.exit("ERROR: el navegador incluido no arranca en la versión compilada.")
        print("5. Paquete")
        paquete = crear_zip(version)
        print(f"  {paquete} ({paquete.stat().st_size / 1048576:.0f} MB)")
    except BaseException:
        poner_version(anterior)
        raise
    if args.sin_publicar:
        poner_version(anterior)
        print("Prueba terminada: no se ha subido nada (version.py vuelve a su valor).")
        return

    print("6. Git")
    if version != anterior:
        ejecutar(["git", "add", "version.py"])
        ejecutar(["git", "commit", "-q", "-m", f"Versión {version}\n\n{args.notas}"])
    ejecutar(["git", "tag", "-a", etiqueta, "-m", f"Versión {version}"])
    ejecutar(["git", "push", "origin", "main", etiqueta])
    print("7. Release en GitHub")
    ejecutar([ruta_gh(), "release", "create", etiqueta, paquete, "--title", f"Palausa Aprovisionamiento {version}", "--notes", args.notas])
    print(f"Publicada la versión {version}. Las demás instalaciones la verán al abrir la app.")


if __name__ == "__main__":
    main()

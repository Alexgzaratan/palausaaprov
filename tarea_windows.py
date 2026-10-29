"""Estado de la tarea programada de facturas MADD en el Programador de tareas de Windows, y corrección de su
ruta para que ejecute el lanzador (así la tarea aplica las actualizaciones y usa siempre la versión instalada)."""
import json
import subprocess
from datetime import datetime, timedelta

NOMBRE_TAREA = "PalausaFacturasMADD"
DIAS_SEMANA = ("L", "M", "X", "J", "V", "S", "D")
# DaysOfWeek del disparador semanal: domingo=1, lunes=2, martes=4... sábado=64
_BITS_DIAS = {6: 1, 0: 2, 1: 4, 2: 8, 3: 16, 4: 32, 5: 64}


def _powershell(script, ejecutar=subprocess.run):
    return ejecutar(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], capture_output=True, text=True,
                    timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def _comillas(texto):
    return "'" + str(texto).replace("'", "''") + "'"


def leer_estado(nombre=NOMBRE_TAREA, ejecutar=subprocess.run):
    """Datos de la tarea o None si no existe (o no se puede leer)."""
    script = (
        f"$t = Get-ScheduledTask -TaskName {_comillas(nombre)} -ErrorAction SilentlyContinue; if (-not $t) {{ 'null'; exit }}; "
        "$i = $t | Get-ScheduledTaskInfo; $d = $t.Triggers | Select-Object -First 1; "
        "$f = { param($x) if ($x -and $x.Year -gt 2000) { $x.ToString('s') } else { $null } }; "
        "[pscustomobject]@{ ejecutable = $t.Actions[0].Execute; argumentos = $t.Actions[0].Arguments; inicio = $d.StartBoundary; "
        "dias = $d.DaysOfWeek; ultima = (& $f $i.LastRunTime); resultado = $i.LastTaskResult; proxima = (& $f $i.NextRunTime); "
        "perdidas = $i.NumberOfMissedRuns; estado = [string]$t.State } | ConvertTo-Json -Compress"
    )
    try:
        salida = _powershell(script, ejecutar)
        datos = json.loads((salida.stdout or "").strip() or "null")
    except Exception:
        return None
    return datos if isinstance(datos, dict) else None


def hora_local(inicio, hoy=None):
    """Hora (time) a la que se dispara hoy, según el StartBoundary ('2026-10-16T06:00:00+01:00' o sin zona)."""
    try:
        momento = datetime.fromisoformat(str(inicio))
    except ValueError:
        return None
    hoy = hoy or datetime.now().date()
    momento = momento.replace(year=hoy.year, month=hoy.month, day=hoy.day)
    return (momento.astimezone().time() if momento.tzinfo else momento.time()).replace(second=0, microsecond=0)


def dia_programado(estado, dia):
    dias = estado.get("dias")
    if not isinstance(dias, int) or not dias:
        return True  # disparador diario u otro tipo: se considera que toca
    return bool(dias & _BITS_DIAS[dia.weekday()])


def texto_dias(estado):
    dias = estado.get("dias")
    if dias == 62:
        return "laborables"
    if not isinstance(dias, int) or not dias:
        return "todos los días"
    return "".join(letra for numero, letra in enumerate(DIAS_SEMANA) if dias & _BITS_DIAS[numero])


def no_se_ejecuto_hoy(estado, ahora=None, margen=timedelta(minutes=30)):
    """True si hoy tocaba, ya ha pasado la hora (con margen) y la última ejecución es de otro día."""
    ahora = ahora or datetime.now()
    hora = hora_local(estado.get("inicio"), ahora.date())
    if hora is None or not dia_programado(estado, ahora.date()):
        return False
    if ahora < datetime.combine(ahora.date(), hora) + margen:
        return False
    ultima = estado.get("ultima")
    try:
        return ultima is None or datetime.fromisoformat(ultima).date() < ahora.date()
    except ValueError:
        return True


def resumen(estado, ahora=None):
    """(texto, hay_problema) para la pestaña de facturas."""
    if estado is None:
        return "Tarea de Windows sin activar: las facturas no se envían solas.", True
    ahora = ahora or datetime.now()
    hora = hora_local(estado.get("inicio"), ahora.date())
    partes = [f"Tarea de Windows: {texto_dias(estado)}" + (f" a las {hora:%H:%M}" if hora else "")]
    if (estado.get("estado") or "").lower() == "disabled":
        partes.append("deshabilitada")
        return " · ".join(partes), True
    problema = False
    if estado.get("ultima"):
        ultima = datetime.fromisoformat(estado["ultima"])
        correcta = estado.get("resultado") in (0, None)
        partes.append(f"última {ultima:%d/%m %H:%M} " + ("✓" if correcta else f"✗ (código {estado.get('resultado')})"))
        problema = not correcta
    if no_se_ejecuto_hoy(estado, ahora):
        partes.append("hoy no se ha ejecutado")
        problema = True
    elif estado.get("proxima"):
        partes.append(f"próxima {datetime.fromisoformat(estado['proxima']):%d/%m %H:%M}")
    return " · ".join(partes), problema


def apunta_al_programa(estado, lanzador):
    """True si la tarea ejecuta el programa directamente (programa\\...exe) en vez del lanzador."""
    ejecutable = str(estado.get("ejecutable") or "").strip('"').lower()
    return bool(ejecutable) and ejecutable != str(lanzador).lower() and ejecutable.endswith("\\programa\\" + lanzador.name.lower())


def apuntar_al_lanzador(lanzador, nombre=NOMBRE_TAREA, ejecutar=subprocess.run):
    """Cambia solo la acción de la tarea (horario y ajustes se conservan). True si se pudo."""
    script = (f"$a = New-ScheduledTaskAction -Execute {_comillas(lanzador)} -Argument '--tarea-madd' -WorkingDirectory {_comillas(lanzador.parent)}; "
              f"Set-ScheduledTask -TaskName {_comillas(nombre)} -Action $a | Out-Null")
    try:
        return _powershell(script, ejecutar).returncode == 0
    except Exception:
        return False

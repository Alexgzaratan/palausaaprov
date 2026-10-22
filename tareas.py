"""Lista de tareas pendientes (pestaña Comparador): datos, fechas límite y guardado en JSON.

El archivo vive junto a la configuración (fuera de programa/), así que las actualizaciones no lo tocan."""
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta

ARCHIVO = "tareas_pendientes.json"


def interpretar_fecha(texto, hoy=None):
    """'' -> None; 'hoy', 'mañana', 'dd/mm', 'dd/mm/aa' o 'dd/mm/aaaa' (también con '-' o '.') -> date.
    Un 'dd/mm' ya pasado este año se entiende del año que viene. Lanza ValueError si no se entiende."""
    hoy = hoy or date.today()
    texto = (texto or "").strip().lower()
    if not texto:
        return None
    if texto == "hoy":
        return hoy
    if texto in ("mañana", "manana"):
        return hoy + timedelta(days=1)
    partes = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2}|\d{4}))?", texto)
    if not partes:
        raise ValueError("Escribe la fecha como dd/mm o dd/mm/aaaa (o «hoy», «mañana»).")
    dia, mes, anio = int(partes.group(1)), int(partes.group(2)), partes.group(3)
    try:
        if anio:
            return date(int(anio) + (2000 if len(anio) == 2 else 0), mes, dia)
        fecha = date(hoy.year, mes, dia)
        return fecha if fecha >= hoy else date(hoy.year + 1, mes, dia)
    except ValueError:
        raise ValueError(f"La fecha {texto} no existe.")


def estado_fecha(fecha_iso, hoy=None):
    """'vencida', 'hoy', 'futura' o '' (sin fecha)."""
    if not fecha_iso:
        return ""
    hoy = hoy or date.today()
    fecha = date.fromisoformat(fecha_iso)
    return "vencida" if fecha < hoy else ("hoy" if fecha == hoy else "futura")


def texto_fecha(fecha_iso, hoy=None):
    if not fecha_iso:
        return ""
    hoy = hoy or date.today()
    fecha = date.fromisoformat(fecha_iso)
    dias = (fecha - hoy).days
    if dias == 0:
        return "Hoy"
    if dias == 1:
        return "Mañana"
    if dias == -1:
        return "Ayer"
    return f"{'Venció' if dias < 0 else 'Vence'} el {fecha:%d/%m}" + (f"/{fecha:%Y}" if fecha.year != hoy.year else "")


class ListaTareas:
    def __init__(self, ruta):
        self.ruta = ruta
        self.tareas = self._cargar()

    def _cargar(self):
        try:
            with open(self.ruta, encoding="utf-8") as f:
                datos = json.load(f)
            return [t for t in datos.get("tareas", []) if isinstance(t, dict) and t.get("texto")]
        except FileNotFoundError:
            return []
        except (OSError, ValueError):
            # Archivo dañado: se aparta (no se pierde) y se empieza con la lista vacía
            try:
                os.replace(self.ruta, self.ruta + ".danado")
            except OSError:
                pass
            return []

    def guardar(self):
        temporal = self.ruta + ".tmp"
        with open(temporal, "w", encoding="utf-8") as f:
            json.dump({"version": 1, "tareas": self.tareas}, f, ensure_ascii=False, indent=1)
        os.replace(temporal, self.ruta)

    def buscar(self, identificador):
        return next((t for t in self.tareas if t["id"] == identificador), None)

    def anadir(self, texto, fecha=None, importante=False):
        texto = " ".join((texto or "").split())
        if not texto:
            return None
        tarea = {"id": uuid.uuid4().hex[:10], "texto": texto, "hecha": False, "importante": bool(importante),
                 "fecha": fecha.isoformat() if fecha else "", "creada": datetime.now().isoformat(timespec="seconds"), "completada": ""}
        self.tareas.append(tarea)
        self.guardar()
        return tarea

    def _cambiar(self, identificador, **valores):
        tarea = self.buscar(identificador)
        if tarea is None:
            return None
        tarea.update(valores)
        self.guardar()
        return tarea

    def marcar(self, identificador, hecha):
        return self._cambiar(identificador, hecha=bool(hecha), completada=datetime.now().isoformat(timespec="seconds") if hecha else "")

    def editar(self, identificador, texto):
        texto = " ".join((texto or "").split())
        return self._cambiar(identificador, texto=texto) if texto else None

    def cambiar_fecha(self, identificador, fecha):
        return self._cambiar(identificador, fecha=fecha.isoformat() if fecha else "")

    def cambiar_importante(self, identificador):
        tarea = self.buscar(identificador)
        return self._cambiar(identificador, importante=not tarea["importante"]) if tarea else None

    def borrar(self, identificador):
        antes = len(self.tareas)
        self.tareas = [t for t in self.tareas if t["id"] != identificador]
        if len(self.tareas) != antes:
            self.guardar()
        return len(self.tareas) != antes

    def borrar_completadas(self):
        antes = len(self.tareas)
        self.tareas = [t for t in self.tareas if not t["hecha"]]
        if len(self.tareas) != antes:
            self.guardar()
        return antes - len(self.tareas)

    def ordenadas(self):
        """Pendientes primero: importantes, luego por fecha límite (las que no tienen, al final) y por antigüedad.
        Después las completadas, la más reciente primero."""
        pendientes = sorted((t for t in self.tareas if not t["hecha"]),
                            key=lambda t: (not t["importante"], t["fecha"] == "", t["fecha"], t["creada"]))
        hechas = sorted((t for t in self.tareas if t["hecha"]), key=lambda t: t["completada"], reverse=True)
        return pendientes + hechas

    def pendientes(self):
        return sum(1 for t in self.tareas if not t["hecha"])

    def vencidas(self, hoy=None):
        return sum(1 for t in self.tareas if not t["hecha"] and estado_fecha(t["fecha"], hoy) == "vencida")

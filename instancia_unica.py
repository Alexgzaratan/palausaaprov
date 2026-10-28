"""Una sola ventana de la app a la vez (mutex con nombre de Windows).

Si ya hay una abierta, la segunda la trae al frente y se cierra. La tarea programada (--tarea-madd) no
pasa por aquí y puede ejecutarse aunque la app esté abierta. Al reiniciar para actualizar, la versión
nueva espera unos segundos a que la anterior termine de cerrarse: el mutex queda libre al salir."""
import ctypes
import sys

NOMBRE_MUTEX = "Local\\PalausaAprovisionamiento"
TITULO_VENTANA = "Palausa Aprovisionamiento"
ERROR_ALREADY_EXISTS = 183
WAIT_OBJECT_0 = 0x0
WAIT_ABANDONED = 0x80
SW_RESTORE = 9


def tomar(nombre=NOMBRE_MUTEX, espera_ms=3000):
    """Handle del mutex si esta es la única instancia (hay que conservarlo hasta salir); None si ya hay otra.
    Fuera de Windows, o si la API falla, no bloquea nada."""
    if sys.platform != "win32":
        return True
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_ulong)
        kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
        handle = kernel32.CreateMutexW(None, True, nombre)
        if not handle:
            return True
        if ctypes.get_last_error() != ERROR_ALREADY_EXISTS:
            return handle
        # Ya existe: si la otra instancia se está cerrando (reinicio tras actualizar) se queda libre enseguida
        if kernel32.WaitForSingleObject(handle, espera_ms) in (WAIT_OBJECT_0, WAIT_ABANDONED):
            return handle
        kernel32.CloseHandle(handle)
        return None
    except Exception:
        return True


def soltar(handle):
    if sys.platform != "win32" or handle in (None, True):
        return
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.ReleaseMutex.argtypes = (ctypes.c_void_p,)
    kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
    kernel32.ReleaseMutex(handle)
    kernel32.CloseHandle(handle)


def traer_al_frente(titulo=TITULO_VENTANA):
    """Muestra la ventana ya abierta; si no la encuentra, avisa de que la app ya está abierta."""
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.FindWindowW.restype = ctypes.c_void_p
        user32.IsIconic.argtypes = (ctypes.c_void_p,)
        user32.ShowWindow.argtypes = (ctypes.c_void_p, ctypes.c_int)
        user32.SetForegroundWindow.argtypes = (ctypes.c_void_p,)
        ventana = user32.FindWindowW(None, titulo)
        if ventana:
            if user32.IsIconic(ventana):
                user32.ShowWindow(ventana, SW_RESTORE)
            user32.SetForegroundWindow(ventana)
            return True
        user32.MessageBoxW(None, f"{titulo} ya está abierta.", titulo, 0x40)
    except Exception:
        pass
    return False

"""Abrir Chromium con Playwright también en la app compilada.

La versión compilada solo incluye el Chromium oculto (headless shell): el completo pesa ~330 MB y se deja fuera
(ver app_automatizacion_optimizada.spec). Las ventanas visibles (litigio en PARTS, ticket de Power-Line) usan entonces
el Microsoft Edge de Windows o, si no está, Google Chrome. Desde el código fuente se sigue usando el de Playwright."""

CANALES_DEL_SISTEMA = ("msedge", "chrome")


def falta_ejecutable(error):
    return "Executable doesn't exist" in str(error)


def comprobar_ventanas(p):
    """Con qué se abrirán las ventanas visibles: «Chromium incluido», «msedge» o «chrome» (se prueba sin enseñar
    nada). RuntimeError si con ninguno."""
    import os
    if os.path.exists(p.chromium.executable_path):
        return "Chromium incluido"
    for canal in CANALES_DEL_SISTEMA:
        try:
            p.chromium.launch(headless=True, channel=canal).close()
            return canal
        except Exception:
            continue
    raise RuntimeError("no hay navegador para las ventanas visibles (PARTS, Power-Line): falta Edge o Chrome")


def lanzar(p, visible, **opciones):
    """p.chromium.launch(headless=not visible); si falta el Chromium incluido, el Edge o el Chrome instalados."""
    try:
        return p.chromium.launch(headless=not visible, **opciones)
    except Exception as e:
        if not falta_ejecutable(e):
            raise
        primero = e
    for canal in CANALES_DEL_SISTEMA:
        try:
            return p.chromium.launch(headless=not visible, channel=canal, **opciones)
        except Exception as e:
            if not falta_ejecutable(e) and "is not found" not in str(e):
                raise
    raise RuntimeError("no se encuentra ningún navegador para abrir la ventana: instala Microsoft Edge o Google Chrome") from primero

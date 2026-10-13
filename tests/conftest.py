import os
import sys
import tempfile
from pathlib import Path

# Los tests nunca deben tocar la configuración ni los logs reales
os.environ.setdefault("PALAUSA_DATOS", tempfile.mkdtemp(prefix="palausa_tests_"))

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

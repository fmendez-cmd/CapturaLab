# INFRAESTRUCTURA/config_rutas.py
import os
import tempfile
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Lectura automática del .env del proyecto
ENV_PATH = os.path.join(BASE_DIR, ".env")
if os.path.exists(ENV_PATH):
    with open(ENV_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                os.environ[key.strip()] = value.strip()

DRIVE_PLANTILLAS_URL = os.getenv("DRIVE_PLANTILLAS_URL", "")
DRIVE_MAPAS_URL = os.getenv("DRIVE_MAPAS_URL", "")

# Salida final: Documentos del usuario actual.
DOCUMENTOS_DIR = str(Path.home() / "Documents")

# Únicamente trabajo temporal. No es caché ni almacenamiento persistente.
TEMP_TRABAJO_DIR = os.path.join(tempfile.gettempdir(), "GeotestLab")
TEMP_PLANTILLAS_DIR = os.path.join(TEMP_TRABAJO_DIR, "plantillas")
os.makedirs(TEMP_PLANTILLAS_DIR, exist_ok=True)

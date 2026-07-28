# utils/rutas.py
"""
Constantes de rutas centralizadas para todo el proyecto.
Si cambia la estructura de carpetas, solo se modifica aquí.
"""
from pathlib import Path

# Raíz del proyecto (carpeta que contiene app.py)
BASE_DIR = Path(__file__).resolve().parent.parent

# Carpeta de datos
DATOS_DIR = BASE_DIR / "datos"

# Archivos de datos
RUTA_MAPEO            = str(DATOS_DIR / "mapeo_plantillas.json")
RUTA_PLANTILLAS       = str(DATOS_DIR / "Hojas_de_verificacion.xlsx")
RUTA_ANALIZADORES     = DATOS_DIR / "analizadores_bel.xlsx"
RUTA_OAUTH_CONFIG     = DATOS_DIR / "google_oauth_client.json"
RUTA_OAUTH_TOKEN      = DATOS_DIR / "google_drive_token.json"

# Carpeta de reportes temporales
RUTA_REPORTES         = str(BASE_DIR / "reportes") + "/"

# Archivo de secrets de Streamlit
RUTA_STREAMLIT_SECRETS = BASE_DIR / ".streamlit" / "secrets.toml"
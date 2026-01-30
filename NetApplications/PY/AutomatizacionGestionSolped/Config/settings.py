# config/settings.py

import os
from dotenv import load_dotenv
from pathlib import Path

# Cargar .env
load_dotenv()

# Ruta base del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent


def get_env_variable(key: str, required: bool = True):
    value = os.getenv(key)

    if required and not value:
        raise EnvironmentError(f"La variable '{key}' no está definida en .env")

    return value


# ========= CONFIG SAP ==========
SAP_CONFIG = {
    "user": get_env_variable("SAP_USUARIO"),
    "password": get_env_variable("SAP_PASSWORD"),
    "mandante": get_env_variable("SAP_MANDANTE"),
    "sistema": get_env_variable("SAP_SISTEMA"),
    "idioma": get_env_variable("SAP_IDIOMA"),
    "logon_path": get_env_variable("SAP_LOGON_PATH"),
}

# ========= Database ==========
DB_CONFIG = {
    "host": get_env_variable("SERVERDB"),
    "database": get_env_variable("NAMEDB"),
    "user": get_env_variable("USERDB"),
    "password": get_env_variable("PASSWORDDB"),
    "schema": get_env_variable("SCHEMA"),
}

# ========= RUTAS =========
RUTAS = {
    "PathLog": get_env_variable("PATHLOG"),
    "PathLogError": get_env_variable("PATHLOGERROR"),
    "PathResultados": get_env_variable("PATHRESULTADOS"),
    "PathReportes": get_env_variable("PATHREPORTES"),
    "PathInsumo": get_env_variable("PATHINSUMO"),
    # "PathTexto": get_env_variable("PATHTEXTO_SAP"),
    # "PathRuta": get_env_variable("PATHRUTA_SAP"),
    # Archivo de configuración de correos
    "ArchivoCorreos": os.path.join(BASE_DIR, "Insumo", "EnvioCorreos.xlsx"),
    # Rutas de archivos
    "PathInsumos": os.path.join(BASE_DIR, "Insumo"),
    "PathSalida": os.path.join(BASE_DIR, "Salida"),
    "PathTemp": os.path.join(BASE_DIR, "Temp"),
    "PathResultado": os.path.join(BASE_DIR, "Resultado"),
}

# Crear carpetas si no existen
for key, path in RUTAS.items():
    if key.startswith("Path") and key not in [
        "PathLog",
        "PathLogError",
        "ArchivoCorreos",
    ]:
        os.makedirs(path, exist_ok=True)

# Crear carpeta de logs
os.makedirs(os.path.dirname(RUTAS["PathLog"]), exist_ok=True)

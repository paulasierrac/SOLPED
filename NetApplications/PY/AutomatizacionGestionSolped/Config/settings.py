# config/settings.py

import os
from dotenv import load_dotenv
from pathlib import Path
from Config.InitConfig import inConfig

# Cargar .env
load_dotenv()



# Ruta base del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent

def LeerVariableEntorno(key: str, required: bool = True):
    value = os.getenv(key)

    if required and not value:
        raise EnvironmentError(f"La variable '{key}' no está definida en .env")

    return value

# ========= CONEXION BASE DE DATOS ==========
DB_CONFIG = {
    "host": LeerVariableEntorno("SERVERDB"),
    "database": LeerVariableEntorno("NAMEDB"),
    "user": LeerVariableEntorno("USERDB"),
    "password": LeerVariableEntorno("PASSWORDDB"),
    "schema": LeerVariableEntorno("SCHEMA"),
    
}

# ========= CONFIG SAP ==========
SAP_CONFIG = {
    "user": LeerVariableEntorno("SAP_USUARIO"),
    "password": LeerVariableEntorno("SAP_PASSWORD"),
}

# ========= CONFIG EMAIL ==========
CONFIG_EMAIL = {
    "smtp_server": LeerVariableEntorno("EMAIL_SMTP_SERVER"),
    "smtp_port": LeerVariableEntorno("EMAIL_SMTP_PORT"),
    "email": LeerVariableEntorno("EMAIL_USER"),
    "password": LeerVariableEntorno("EMAIL_PASSWORD"),  # IMPORTANTE: Cambiar por variable de entorno en producción
}

# ========= RUTAS =========
RUTAS = {
    "PathLog":inConfig("PathLog"),
    "PathLogError":inConfig("PathErrorLog"),   
    #"PathLogError": LeerVariableEntorno("PATHLOGERROR"),
    "PathResultados": LeerVariableEntorno("PATHRESULTADOS"),
    "PathReportes": LeerVariableEntorno("PATHREPORTES"),
    "PathInsumo": LeerVariableEntorno("PATHINSUMO"),
    "PathTexto": LeerVariableEntorno("PATHTEXTO_SAP"),
    "PathRuta": LeerVariableEntorno("PATHRUTA_SAP"),
    #"PathTempFileServer": LeerVariableEntorno("SAP_TEMP_PATH"),
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


# config/settings.py

import os
from dotenv import load_dotenv

# Cargar variables desde .env
load_dotenv()


def get_env_variable(key: str, required: bool = True):
    """
    Obtiene una variable de entorno. 
    Si `required=True` y no existe, lanza un error más claro.
    """
    value = os.getenv(key)

    if required and not value:
        raise EnvironmentError(f"La variable de entorno '{key}' no está definida en el archivo .env")

    return value


# === Configuración SAP ===
SAP_CONFIG = {
    "user": get_env_variable("SAP_USUARIO"),
    "password": get_env_variable("SAP_PASSWORD"),
    "mandante": get_env_variable("SAP_MANDANTE"),
    "sistema": get_env_variable("SAP_SISTEMA"),
    "idioma": get_env_variable("SAP_IDIOMA"),
    "logon_path": get_env_variable("SAP_LOGON_PATH"),
}


RUTAS = {
    "PathLog": get_env_variable("PATHLOG"),
    "PathLogError": get_env_variable("PATHLOGERROR")
}

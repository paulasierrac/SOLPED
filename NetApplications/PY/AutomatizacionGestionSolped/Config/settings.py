# config/settings.py
import os
#from dotenv import load_dotenv

# Cargar el archivo .env
#load_dotenv()

SAP_CONFIG = {
    "user": os.getenv("SAP_USER"),
    "password": os.getenv("SAP_PASSWORD"),
    "mandante": os.getenv("SAP_MANDANTE"),
    "sistema": os.getenv("SAP_SISTEMA"),
    "idioma": os.getenv("SAP_IDIOMA"),
    "logon_path": os.getenv("SAP_LOGON_PATH"),
}
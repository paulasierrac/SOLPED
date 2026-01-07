# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from requests import session
from HU.HU01_LoginSAP import ObtenerSesionActiva,conectar_sap
from Funciones.ValidacionM21N import SapTextEditor,AbrirSolped
from Funciones.GeneralME53N import AbrirTransaccion
import pyautogui  # Asegúrate de tener pyautogui instalado
import time

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


import re
from typing import List, Optional


def Main_Login():
    try:

        #session = conectar_sap( SAP_CONFIG["sistema"], SAP_CONFIG["mandante"],SAP_CONFIG["user"], SAP_CONFIG["password"], SAP_CONFIG["idioma"] )
        session = ObtenerSesionActiva()
        AbrirTransaccion(session, "ME2L")
        # codigo para pruebas
        #abrirSolped(session, 1300139390, 7)
        #abrirSolped(session, 1300139391, 9)
        #AbrirSolped(session, 1300139392, 4)
        #abrirSolped(session, 1300139393, 7)
        #abrirSolped(session, 1300139394, 7)
        
       

        
      

    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

if __name__ == "__main__":
    Main_Login()



# ============================================
# HU05: Descarga de Orden de Compra (OC) en ME9F
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Descarga la OC generada desde la transacción ME9F.
# Ultima modificacion: 08/12/2023
# Propiedad de Colsubsidio
# Cambios: Estructura y logs.
# ============================================

import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import traceback
import pyautogui
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from Funciones.GeneralME53N import AbrirTransaccion
from Funciones.ValidacionM21N import esperar_sap_listo

def EjecutarHU05(session, orden_de_compra):
    """
    Ejecuta la Historia de Usuario 05: Descarga de OC desde ME9F.
    """
    task_name = "HU05_DescargaOCME9F"

    try:
        WriteLog(
            mensaje=f"Inicia HU05 para la Orden de Compra: {orden_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        
        if not session:
            raise ValueError("Sesion SAP no valida.")

        if not orden_de_compra:
            raise ValueError("El número de Orden de Compra es inválido o no fue proporcionado.")

        # Abrir transacción ME9F
        AbrirTransaccion(session, "ME9F")
        esperar_sap_listo(session)   
          

        session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = orden_de_compra
        session.findById("wnd[0]").sendVKey(0)
        
        # Presionar el botón de ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1)
        
        # Seleccionar la línea y "Message Output"
        session.findById("wnd[0]/usr/chk[1,5]").selected = True
        pyautogui.hotkey("shift", "f5") # Botón "Message Output"

        #Adicionar codigo para guardar el PDF resultante, hilo treads para manejo de la ventana emergente 

        WriteLog(
            mensaje=f"Procesamiento en ME9F completado para la OC: {orden_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU05: {e} | {error_text}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise

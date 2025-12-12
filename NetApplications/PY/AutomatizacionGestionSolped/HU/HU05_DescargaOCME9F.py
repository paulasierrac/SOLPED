# ============================================
# HU04: Descarga de Orden de Compra (OC) en ME9F
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

def EjecutarHU04(session, orden_de_compra):
    """
    Ejecuta la Historia de Usuario 04: Descarga de OC desde ME9F.
    """
    task_name = "HU04_DescargaOCME9F"

    try:
        WriteLog(
            mensaje=f"Inicia HU04 para la Orden de Compra: {orden_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        if not session:
            raise ValueError("Sesion SAP no valida.")

        if not orden_de_compra:
            raise ValueError("El número de Orden de Compra es inválido o no fue proporcionado.")

        # Abrir transacción ME9F
        AbrirTransaccion(session, "/nME9F")
        
        WriteLog(
            mensaje="Transacción ME9F abierta con éxito.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = orden_de_compra
        session.findById("wnd[0]").sendVKey(0)
        
        # Presionar el botón de ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1)
        
        # Seleccionar la línea y "Message Output"
        session.findById("wnd[0]/usr/chk[1,5]").selected = True
        pyautogui.hotkey("shift", "f5") # Botón "Message Output"

        WriteLog(
            mensaje=f"Procesamiento en ME9F completado para la OC: {orden_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU04: {e} | {error_text}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise

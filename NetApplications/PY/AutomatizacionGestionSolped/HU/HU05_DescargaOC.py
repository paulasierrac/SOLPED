# ============================================
# HU05: Descarga de Orden de Compra (OC) en ME9F
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Descarga la OC generada desde la transacción ME9F.
# Ultima modificacion: 08/12/2023
# Propiedad de Colsubsidio
# Cambios: Estructura y logs.
# ============================================

import pyperclip
from requests import session
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import traceback
import pyautogui
from Funciones.GuiShellFunciones import set_GuiTextField_text
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from Funciones.GeneralME53N import AbrirTransaccion
from Funciones.ValidacionM21N import esperar_sap_listo
from Funciones.ControlHU import control_hu

def EjecutarHU05(session, ordenes_de_compra):
    """
    Ejecuta la Historia de Usuario 05: Descarga de OC desde ME9F.
    """
    task_name = "HU05_DescargaOC"

    try:
        WriteLog(
            mensaje=f"Inicia HU05 para la Orden de Compra: {ordenes_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        if not session:
            raise ValueError("Sesion SAP no valida.")

        if not ordenes_de_compra:
            raise ValueError(
                "El número de Orden de Compra es inválido o no fue proporcionado."
            )

        # Abrir transacción ME9F
        AbrirTransaccion(session, "ME2L")
        esperar_sap_listo(session)

        # Alcance de la lista
        session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "ALV"

        session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press()

        # Definir la lista de órdenes de compra
        ordenes_de_compra = [
            "4200339200",
            "4200339201",
            "4200339202",
            "4200339203",
            "4200339204",
            "4200339205",
            "4200339206",
        ]
        # Convertir la lista a una cadena de texto (por ejemplo, separada por saltos de línea)
        for i in range(len(ordenes_de_compra)):
            set_GuiTextField_text(session, f"SLOW_I[1,{i}]", ordenes_de_compra[i])

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Presionar el botón de ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1)

        # Seleccionar la línea y "Message Output"
        # session.findById("wnd[0]/usr/chk[1,5]").selected = True
        # pyautogui.hotkey("shift", "f5") # Botón "Message Output"

        # Adicionar codigo para guardar el PDF resultante, hilo treads para manejo de la ventana emergente

        WriteLog(
            mensaje=f"Procesamiento en ME9F completado para la OC: {ordenes_de_compra}",
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

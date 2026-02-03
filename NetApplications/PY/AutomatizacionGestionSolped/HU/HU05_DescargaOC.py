# ============================================
# HU05: Descarga de Orden de Compra (OC) en ME9F
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Descarga la OC generada desde la transacción ME9F.
# Ultima modificacion: 08/12/2023
# Propiedad de Colsubsidio
# Cambios: Estructura y logs.
# ============================================

from datetime import datetime
import pyperclip
from requests import session
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import traceback
import pyautogui
from funciones.GuiShellFunciones import ProcesarTablaMejorada
from funciones.EscribirLog import WriteLog
from config.settings import RUTAS
from funciones.GeneralME53N import AbrirTransaccion
from funciones.ValidacionM21N import esperar_sap_listo

def EjecutarHU05(session, ordenes_de_compra: list):
    """
    Ejecuta la Historia de Usuario 05: Descarga de OC desde ME9F.
    """
    task_name = "HU05_DescargaOC"

    try:

        if not session:
            raise ValueError("Sesion SAP no valida.")

        # Abrir transacción ME9F
        AbrirTransaccion(session, "ME2L")
        esperar_sap_listo(session)   
          
        # Alcance de la lista
        session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV"
        # Presionar Enter
        session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press()

        # Ingresar las órdenes de compra en la tabla
        for i in range(len(ordenes_de_compra)):
            ventanaobj = session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]")
            ventanaobj.text = ordenes_de_compra[i]
          
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        # Presionar el botón de ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(0.5)
        session.findById("wnd[0]/tbar[1]/btn[45]").press()  # Botón de lista de opciones / Fichero local crtl + shift + F9
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Botón de exportar
        # Seleccionar la línea y "Message Output"

        # === Fecha ===
        ahora = datetime.now()
        fecha_hora = ahora.strftime("%d/%m/%Y %H:%M:%S")
        fecha_archivo = ahora.strftime("%Y%m%d_%H%M%S")
        #Guardar el archivo txt en la ruta especificada
        ruta_guardar = rf"{RUTAS["PathInsumo"]}"
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta_guardar
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = rf"LiberadasOC_{fecha_archivo}.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar
     
        archivo = rf"LiberadasOC_{fecha_archivo}.txt"
        df_Ocliberadas = ProcesarTablaMejorada(archivo)

        print (df_Ocliberadas)

        


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

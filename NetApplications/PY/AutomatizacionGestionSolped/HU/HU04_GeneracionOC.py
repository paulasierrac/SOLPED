# ============================================
# HU05: Generacion de Orden de Compra desde Solped
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Genera Orden de Compra a partir de las Solicitudes de Pedido validadas.
# Ultima modificacion: 27/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import re
import subprocess
import time
import os
from Config.settings import RUTAS
from Funciones.ValidacionM21N import BorrarTextosDesdeSolped,obtener_numero_oc,ejecutar_accion_sap,buscar_y_clickear,limpiar_id_sap,ejecutar_creacion_hijo
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion,procesarTablaME5A
import traceback
import pyautogui  # Asegúrate de tener pyautogui instalado



def EjecutarHU04(session, archivo):
    
    task_name = "HU5_GeneracionOC"
    """
    Ejecuta la Historia de Usuario 05 encargada de la
    generacion de OC desde la transacción ME21N.
    """
    try:
        WriteLog(
            mensaje=f"Inicia HU05 para el archivo {archivo}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        # ============================
        # Abrir transacción ME21N
        # ============================
        AbrirTransaccion(session, "ME21N")
        print("Transacción ME21N abierta con éxito.")
        time.sleep(0.5)

        # ============================
        # Limpiar textos Solped
        # ============================
        df_solpeds = procesarTablaME5A(archivo)

        if df_solpeds.empty:
            WriteLog(
                mensaje=f"No se encontraron Solpeds para procesar en el archivo {archivo}.",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return

        solpeds_unicas = df_solpeds['PurchReq'].unique()
        
        for solped in solpeds_unicas:
            # Contar los items para la solped actual
            item_count = df_solpeds[df_solpeds['PurchReq'] == solped].shape[0]
            
            WriteLog(
                mensaje=f"Procesando Solped: {solped} con {item_count} item(s).",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            
            BorrarTextosDesdeSolped(session, solped, item_count)

        # Después de procesar todas las solpeds y (presumiblemente) guardar la OC.
        orden_de_compra = obtener_numero_oc(session)
        WriteLog(
            mensaje=f"Se generó la Orden de Compra: {orden_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )


        WriteLog(
            mensaje=f"HU05 finalizada correctamente para archivo {archivo}.",
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





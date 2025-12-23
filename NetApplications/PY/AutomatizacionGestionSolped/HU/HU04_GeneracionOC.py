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
from Funciones.ValidacionM21N import BorrarTextosDesdeSolped,obtener_numero_oc
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion,procesarTablaME5A
import traceback
import pyautogui  # Asegúrate de tener pyautogui instalado

def EjecutarHU04(session, archivo):
    
    task_name = "HU4_GeneracionOC"
    """
    Ejecuta la Historia de Usuario 04 encargada de la
    generacion de OC desde la transacción ME21N.
    """
    try:
        WriteLog(
            mensaje=f"HU04 Inicia para el archivo {archivo}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        

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
        print(f"Solpeds únicas a procesar: {solpeds_unicas}")
        
        for solped in solpeds_unicas[1:]:  # Saltar la primera solped si es necesario (Encabezados)
             # --- Validación de Solped ---
            if (
                not solped                      # None o vacío
                or not str(solped).isdigit()    # Debe ser numérica
                #or len(str(solped)) != 10       # Longitud típica SAP (ej: 1300139274)
                or solped not in df_solpeds['PurchReq'].values  # Debe existir en el DF
            ):
                WriteLog(
                    mensaje=f"Solped inválida u omitida: {solped}",
                    estado="WARNING",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )
                continue  # Saltar a la siguiente solped
            # Contar los items para la solped actual
            item_count = df_solpeds[df_solpeds['PurchReq'] == solped].shape[0]
            
            WriteLog(
                mensaje=f"Procesando Solped: {solped} con {item_count} item(s).",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            print(f"Session actual: {session}")
            print(f"procesando solped: {solped} de items: {item_count}")
            BorrarTextosDesdeSolped(session, solped, item_count)
            orden_de_compra = obtener_numero_oc(session)
            # Después de procesar todas las solpeds y (presumiblemente) guardar la OC.        
            WriteLog(
                mensaje=f"Se generó la Orden de Compra: {orden_de_compra}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )


        WriteLog(
            mensaje=f"HU04 finalizada correctamente para archivo {archivo}.",
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





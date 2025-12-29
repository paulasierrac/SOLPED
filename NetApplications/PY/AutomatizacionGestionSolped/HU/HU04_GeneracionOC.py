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
from Funciones.ValidacionM21N import select_GuiTab, obtener_numero_oc,set_GuiComboBox_key,cambiar_grupo_compra, validar_y_ajustar_solped,abrirSolped
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
        #print(f"Solpeds a procesar: {solpeds_unicas}")
        WriteLog(
                    mensaje=f"listado de Solped cargadas : {solpeds_unicas}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )
        
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
                mensaje=f"Procesando Solped: {solped} de items: {item_count} .",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"]
                #path_log=f"{RUTAS["PathLog"]}StevInforme.txt", # revisar ruta para hacer el informe 
                
            )
            #print(f"procesando solped: {solped} de items: {item_count}")
            AbrirTransaccion(session, "ME21N")
            #navegacion por SAP que permite abrir Solped 
            abrirSolped(session, solped, item_count)
            #se selecciona la clase de docuemnto ZRCR, revisar alcance si es necesario cambiar a otra clase dependiendo de algun criterio
            set_GuiComboBox_key(session, "TOPLINE-BSART", "ZRCR")
            #se ingresa a la pestaña  Dat.org. de cabecera, asegurándonos de que esté visible
            select_GuiTab(session, "TABHDT9") 
            # Se cambia el grupo de compra dependiendo de la org de compra
            cambiar_grupo_compra(session)
            # Seleccionar la pestaña de textos, asegurándonos de que esté visible
            select_GuiTab(session, "TABIDT14")
            # Valores y textos se validan y ajustan 
            validar_y_ajustar_solped(session, item_count)
            #***********///////////**************///////////********
            # Se debe remplazar con guardar OC 
            #////////////*******///////////////*******
            # Salir para pruebas 
            pyautogui.press("F12")
            time.sleep(1)
            pyautogui.hotkey("TAB")
            time.sleep(0.5)
            pyautogui.hotkey("enter")
            # Salir para pruebas 
            #***********///////////**************///////////********  
            
            orden_de_compra = obtener_numero_oc(session)
            # Después de procesar todas las solpeds y (presumiblemente) guardar la OC.        
            WriteLog(
                mensaje=f" para la solped : {solped} Se generó la Orden de Compra: {orden_de_compra}",
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





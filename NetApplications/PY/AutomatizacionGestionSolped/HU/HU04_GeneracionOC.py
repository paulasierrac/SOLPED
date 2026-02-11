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
from Funciones.ControlHU import control_hu
from Funciones.GuiShellFunciones import (
    esperar_sap_listo,
    obtener_numero_oc,
    ProcesarTabla,
    SetGuiComboBoxkey,
    CambiarGrupoCompra,
)
from Funciones.ValidacionM21N import (
    SelectGuiTab,
    ValidarAjustarSolped,
    AbrirSolped,
    MostrarCabecera,
)
from Funciones.EscribirInforme import WriteInformeOperacion
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion
import traceback
import pyautogui  # Asegúrate de tener pyautogui instaladoi
from Funciones.ControlHU import control_hu


def EjecutarHU04(session, archivo):

    task_name = "HU4_GeneracionOC"
    """
    Ejecuta la Historia de Usuario 04 encargada de la
    generacion de OC desde la transacción ME21N.
    """
    try:
        control_hu(task_name=task_name, estado=0)
        WriteLog(
            mensaje=f"HU04 Inicia para el archivo {archivo}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ============================
        # Limpiar textos Solped
        # ============================
        # Cambiar Funcion por # df_solpeds = ProcesarTablaME5A(archivo)
        df_solpeds = ProcesarTabla(archivo)

        if df_solpeds.empty:
            WriteLog(
                mensaje=f"No se encontraron Solpeds para procesar en el archivo {archivo}.",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return

        solpeds_unicas = df_solpeds["PurchReq"].unique()
        print(f"Solpeds a procesar: {solpeds_unicas}")
        WriteLog(
            mensaje=f"listado de Solped cargadas : {solpeds_unicas}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        for (
            solped
        ) in solpeds_unicas:  # Saltar la primera solped si es necesario (Encabezados)
            # --- Validación de Solped ---
            if (
                not solped  # None o vacío
                or not str(solped).isdigit()  # Debe ser numérica
                # or len(str(solped)) != 10       # Longitud típica SAP (ej: 1300139274)
                or solped not in df_solpeds["PurchReq"].values  # Debe existir en el DF
            ):
                WriteLog(
                    mensaje=f"Solped inválida u omitida: {solped}",
                    estado="WARNING",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )
                continue  # Saltar a la siguiente solped
            # Contar los items para la solped actual
            item_count = df_solpeds[df_solpeds["PurchReq"] == solped].shape[0]

            WriteLog(
                mensaje=f"Procesando Solped: {solped} de items: {item_count} .",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            acciones = []
            # print(f"procesando solped: {solped} de items: {item_count}")
            AbrirTransaccion(session, "ME21N")
            # navegacion por SAP que permite abrir Solped

            AbrirSolped(session, solped, item_count)

            # se selecciona la clase de docuemnto ZRCR, revisar alcance si es necesario cambiar a otra clase dependiendo de algun criterio
            SetGuiComboBoxkey(session, "TOPLINE-BSART", "ZRCR")

            esperar_sap_listo(session)
            pyautogui.hotkey("ctrl", "F2")
            # se ingresa a la pestaña  Dat.org. de cabecera, asegurándonos de que esté visible
            SelectGuiTab(session, "TABHDT9")
            # Se cambia el grupo de compra dependiendo de la org de compra, y se guardan acciones
            acciones.extend(CambiarGrupoCompra(session))
            # Seleccionar la pestaña de textos, asegurándonos de que esté visible
            esperar_sap_listo(session)
            # time.sleep(0.5)
            pyautogui.hotkey("ctrl", "F4")
            SelectGuiTab(session, "TABIDT14")
            # Valores y textos se validan y ajustan
            acciones.extend(ValidarAjustarSolped(session, item_count))

            # *********************************
            # Se debe remplazar con guardar OC
            # *********************************
            # /Salir para pruebas
            pyautogui.press("F12")
            time.sleep(1)
            pyautogui.hotkey("TAB")
            time.sleep(0.5)
            pyautogui.hotkey("enter")
            # /Salir para pruebas
            # *********************************

            # Obtener el numero de la orden de compra generada desde la barra de estado.
            orden_de_compra = obtener_numero_oc(session)

            # Stev: validar si se debe hacer algo mas con la OC generada

            ruta = WriteInformeOperacion(
                item_count=item_count,
                solped=solped,
                orden_compra=orden_de_compra,
                acciones=acciones,
                estado="EXITOSO",
                bot_name="Resock",
                task_name=task_name,
                path_informes=r".\Salida",
                observaciones="Proceso ejecutado sin errores.",
            )

            print(f"Informe generado en: {ruta}")

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
        control_hu(task_name=task_name, estado=100)


    except Exception as e:
        control_hu(task_name=task_name, estado=99)
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU04: {e}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise

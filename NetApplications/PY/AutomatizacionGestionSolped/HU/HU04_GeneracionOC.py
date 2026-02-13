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
from Config.InicializarConfig import inConfig
from Config.settings import RUTAS
from Funciones.GuiShellFunciones import (
    EsperarSAPListo,
    ObtenerNumeroOC,
    ProcesarTabla,
    SetGuiComboBoxkey,
    CambiarGrupoCompra,
)
from Funciones.ValidacionME21N import (
    SelectGuiTab,
    ValidarAjustarSolped,
    AbrirSolped,
    MostrarCabecera,
)
from Funciones.EscribirInforme import EscribirIformeOperacion
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion
import traceback
import pyautogui  # Asegúrate de tener pyautogui instaladoi
from Funciones.ControlHU import ControlHU

from Repositories.Consultas import Querys


def EjecutarHU04(session, archivo):

    nombreTarea = "HU4_GeneracionOC"
    """
    Ejecuta la Historia de Usuario 04 encargada de la
    generacion de OC desde la transacción ME21N.
    """
    try:
        ControlHU(nombreTarea, estado=0)
        WriteLog(
            mensaje=f"HU04 Inicia para el archivo {archivo}",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        # ============================
        # Limpiar textos Solped
        # ============================
        # Cambiar Funcion por # dfSolpeds = ProcesarTablaME5A(archivo)
        dfSolpeds = ProcesarTabla(archivo)
        """
        #STEV: se trata de llenar el data frame desde la base de datos pero falla 

        query = Querys("GestionSolped")
        dfSolpeds = query.fetch_all(tabla="expsolped03")
        print(dfSolpeds)
        """
        if dfSolpeds.empty:
            WriteLog(
                mensaje=f"No se encontraron Solpeds para procesar en el archivo {archivo}.",
                estado="WARNING",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
            return

        solpedsUnicas = dfSolpeds["PurchReq"].unique()
        print(f"Solpeds a procesar: {solpedsUnicas}")
        WriteLog(
            mensaje=f"listado de Solped cargadas : {solpedsUnicas}",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        for (
            solped
        ) in solpedsUnicas:  # Saltar la primera solped si es necesario (Encabezados)
            # --- Validación de Solped ---
            if (
                not solped  # None o vacío
                or not str(solped).isdigit()  # Debe ser numérica
                # or len(str(solped)) != 10       # Longitud típica SAP (ej: 1300139274)
                or solped not in dfSolpeds["PurchReq"].values  # Debe existir en el DF
            ):
                WriteLog(
                    mensaje=f"Solped inválida u omitida: {solped}",
                    estado="WARNING",
                    nombreTarea=nombreTarea,
                    rutaRegistro=inConfig("PathLog"),
                )
                continue  # Saltar a la siguiente solped
            # Contar los items para la solped actual
            itemCount = dfSolpeds[dfSolpeds["PurchReq"] == solped].shape[0]

            WriteLog(
                mensaje=f"Procesando Solped: {solped} de items: {itemCount} .",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
            acciones = []

            # print(f"procesando solped: {solped} de items: {itemCount}")
            AbrirTransaccion(
                session,
                "ME21N",
            )
            EsperarSAPListo(session)
            # navegacion por SAP que permite abrir Solped
            posiciones = ["10", "40", "50", "60"]
            AbrirSolped(session, solped, itemCount, posiciones)

            # se selecciona la clase de docuemnto ZRCR, revisar alcance si es necesario cambiar a otra clase dependiendo de algun criterio
            SetGuiComboBoxkey(session, "TOPLINE-BSART", "ZRCR")

            EsperarSAPListo(session)

            # se ingresa a la pestaña  Dat.org. de cabecera, asegurándonos de que esté visible
            pyautogui.hotkey("ctrl", "F2")
            SelectGuiTab(session, "TABHDT9")
            # Se cambia el grupo de compra dependiendo de la org de compra, y se guardan acciones
            acciones.extend(CambiarGrupoCompra(session))
            # Seleccionar la pestaña de textos, asegurándonos de que esté visible
            EsperarSAPListo(session)
            # time.sleep(0.5)
            # pestaña textos
            pyautogui.hotkey("ctrl", "F4")
            SelectGuiTab(session, "TABIDT14")
            # Valores y textos se validan y ajustan
            acciones.extend(ValidarAjustarSolped(session, itemCount))

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
            ordenDeCompra = ObtenerNumeroOC(session)

            # Stev: validar si se debe hacer algo mas con la OC generada

            ruta = EscribirIformeOperacion(
                itemCount=itemCount,
                solped=solped,
                ordenCompra=ordenDeCompra,
                acciones=acciones,
                estado="EXITOSO",
                botName="Resock",
                nombreTarea=nombreTarea,
                pathInformes=r".\Salida",
                observaciones="Proceso ejecutado sin errores.",
            )

            print(f"Informe generado en: {ruta}")

            # Después de procesar todas las solpeds y (presumiblemente) guardar la OC.
            WriteLog(
                mensaje=f" para la solped : {solped} Se generó la Orden de Compra: {ordenDeCompra}",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )

        WriteLog(
            mensaje=f"HU04 finalizada correctamente para archivo {archivo}.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        ControlHU(nombreTarea, estado=100)

    except Exception as e:
        ControlHU(nombreTarea, estado=99)
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU04: {e}",
            estado="ERROR",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        raise

# ================================
# GestionSOLPED – HU00: DespliegueAmbiente
# Autor: Steven Navarro - NetApplications
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste ruta base dinámica + estándar Colsubsidio
# ================================


import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import subprocess
import os
from Config.InicializarConfig import inConfig
from Config.settings import RUTAS, SAP_CONFIG 
from Funciones.ValidacionME21N import ventana_abierta

import pyautogui

from Config.InicializarConfig import inConfig
from Config.settings import RUTAS, SAP_CONFIG
from Funciones.ValidacionME21N import ventana_abierta
from Funciones.ControlHU import control_hu


def AbrirSAPLogon():
    """Abre SAP Logon si no está ya abierto."""
    #SAP_CONFIG = get_sap_config()
    try:
        # WriteLog | INFO | INICIA abrir_sap_logon

        win32com.client.GetObject("SAPGUI")
        print("INFO | SAP Logon ya se encuentra abierto")

        # WriteLog | INFO | FINALIZA abrir_sap_logon
        return True
    except:
        # Si no está abierto, se lanza el ejecutable
        #"logon_path": LeerVariableEntorno("SAP_LOGON_PATH"),
        subprocess.Popen(inConfig("SapRutaLogon"))
        time.sleep(5)  # Esperar a que abra SAP Logon
        return False


def ConectarSAP(conexion, mandante, usuario, password, idioma="ES"):

    abrir_sap = AbrirSAPLogon()
    time.sleep(3)
    if abrir_sap:
        print(" SAP Logon 750 ya se encuentra abierto")
    else:
        print(" SAP Logon 750 abierto ")

    try:
        # WriteLog | INFO | INICIA ConectarSAP
        task_name = "HU01_LoginSAP"
        control_hu(task_name, estado=0)

        abrir_sap = AbrirSAPLogon()
        time.sleep(3)

        if abrir_sap:
            print("INFO | SAP Logon ya se encuentra abierto")
        else:
            print("INFO | SAP Logon abierto correctamente")

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            raise Exception("No se pudo obtener objeto SAPGUI")

        application = sap_gui_auto.GetScriptingEngine

        connection = None
        for item in application.Connections:
            if item.Description.strip().upper() == conexion.strip().upper():
                connection = item
                break

        if not connection:
            print(f"INFO | Abriendo nueva conexion a {conexion}")
            connection = application.OpenConnection(conexion, True)
            time.sleep(3)
        else:
            print(f"INFO | Conexion existente encontrada con {conexion}")

        if connection.Children.Count > 0:
            session = connection.Children(0)
            print("INFO | Sesion existente reutilizada")
        else:
            session = connection.Children(0).CreateSession()
            print("INFO | Nueva sesion creada")

        # Login
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        session.findById("wnd[0]").sendVKey(0)

        print("INFO | Credenciales enviadas correctamente")

        if ventana_abierta(session, "Copyrigth"):
            pyautogui.press("enter")
            print("INFO | Ventana Copyrigth cerrada")

        try:
            if validarLoginDiag(
                ruta_imagen=rf".\img\logindiag.png",
                confidence=0.5,
                intentos=20,
                espera=0.5,
            ):
                print("INFO | Ventana loginDiag superada correctamente")
        except Exception as e:
            print(f"no se encontro ventana Copyrigth en login {e}")

        if ventana_abierta(session, "Info de licencia en entrada al sistema múltiple"):
            
            print("entro a la funcion click")
            time.sleep(20)  
            pyautogui.click()
            pyautogui.press("enter")
               
            try:
                if validarLoginDiag(
                    ruta_imagen=rf".\img\Infodelicenciaenentradaalsistemamultiple.png",
                    confidence=0.8,
                    intentos=20,
                    espera=0.5
                ):  
                    pyautogui.click()
                    print("encontro la imagen ")
                    print("Ventana info de licencia inesperada superada correctamente")
            except Exception as e:
                print(f"no se encontro ventana Copyrigth en login {e}")
        return session

    except Exception as e:
        control_hu(task_name, estado=99)
        print(f"ERROR | Error al conectar a SAP: {e}")
        # WriteLog | ERROR | Error grave ConectarSAP
        return None


# ============================================================
# ObtenerSesionActiva
# Autor: Automatizacion RPA
# Descripcion: Obtiene una sesion activa de SAP
# ============================================================
def ObtenerSesionActiva():

    try:
        # WriteLog | INFO | INICIA ObtenerSesionActiva

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        for conn in application.Connections:
            if conn.Children.Count > 0:
                session = conn.Children(0)
                print(f"INFO | Sesion encontrada en conexión: {conn.Description}")

                # WriteLog | INFO | FINALIZA ObtenerSesionActiva
                return session

        print("WARN | No se encontró ninguna sesion activa")
        return None

    except Exception as e:
        print(f"ERROR | Error al obtener la sesion activa: {e}")
        # WriteLog | ERROR | Error ObtenerSesionActiva
        return None


# ============================================================
# validarLoginDiag
# Autor: Automatizacion RPA
# Descripcion: Valida ventana de dialogo por imagen
# ============================================================
def validarLoginDiag(ruta_imagen, confidence=0.5, intentos=3, espera=0.5):

    for intento in range(intentos):
        try:
            pos = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
            if pos:
                pyautogui.press("enter")
                return True

        except Exception as e:
            print(f"WARN | Error buscando imagen loginDiag: {e}")

        time.sleep(espera)

    print(f"WARN | No se encontró la ventana login diag: {ruta_imagen}")
    return False

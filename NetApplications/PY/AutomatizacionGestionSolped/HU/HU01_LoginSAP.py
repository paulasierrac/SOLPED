# ================================
# GestionSOLPED – HU00: DespliegueAmbiente
# Autor: Steven Navarro - NetApplications
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste ruta base dinámica + estándar Colsubsidio
# ================================
import traceback

import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import subprocess

import pyautogui

from Config.InicializarConfig import inConfig
from Config.settings import RUTAS, SAP_CONFIG 
from Funciones.ValidacionME21N import ventanaAbierta
from Funciones.EscribirLog import WriteLog
from Funciones.ControlHU import ControlHU



def AbrirSAPLogon():
    """Abre SAP Logon si no está ya abierto."""
    try:
        win32com.client.GetObject("SAPGUI")
        return True
    except:
        # Si no está abierto, se lanza el ejecutable
        subprocess.Popen(inConfig("SapRutaLogon"))
        time.sleep(2)  # Esperar a que abra SAP Logon
        return False


def ConectarSAP(conexion, mandante, usuario, password, idioma="ES"):

    try:
        nombreTarea = "HU01_LoginSAP"
        ControlHU(nombreTarea, estado=0)
        abrirSap = AbrirSAPLogon()
        if abrirSap:
            WriteLog( mensaje="SAP Logon ya se encuentra abierto",estado="WARN", nombreTarea="Abrir SAP Logon", )
        else:
            WriteLog( mensaje="SAP Logon 750 abierto",estado="INFO", nombreTarea="Abrir SAP Logon", )
      
        
        sapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not sapGuiAuto:
            raise Exception("No se pudo obtener objeto SAPGUI")

        application = sapGuiAuto.GetScriptingEngine

        connection = None
        for item in application.Connections:
            if item.Description.strip().upper() == conexion.strip().upper():
                connection = item
                break

        if not connection:
            WriteLog( mensaje=f"Abriendo nueva conexion a {conexion}",estado="INFO", nombreTarea="Abrir SAP Logon", )
            connection = application.OpenConnection(conexion, True)
        else:
            WriteLog( mensaje=f"Conexion existente encontrada con {conexion}",estado="INFO", nombreTarea="Abrir SAP Logon", )

        if connection.Children.Count > 0:
            session = connection.Children(0)
            WriteLog( mensaje=f"Sesion existente reutilizada",estado="INFO", nombreTarea="Abrir SAP Logon", )
        else:
            session = connection.Children(0).CreateSession()
            WriteLog( mensaje=f"Nueva sesion creada",estado="INFO", nombreTarea="Abrir SAP Logon", )


        # Login
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        session.findById("wnd[0]").sendVKey(0)

        WriteLog( mensaje=f"Credenciales enviadas correctamente",estado="INFO", nombreTarea="Abrir SAP Logon", )

        if ventanaAbierta(session, "Copyrigth"):
            pyautogui.press("enter")
            WriteLog( mensaje=f"Ventana Copyrigth cerrada",estado="INFO", nombreTarea="Abrir SAP Logon", )
        
        #Time sleep para el multisecion alcazar a Seleccionar continuar 
        time.sleep(5)

        return session

    except Exception as e:
        #traceback.print_exc()
        ControlHU(nombreTarea, estado=99)
        WriteLog( mensaje=f"Error al conectar a SAP: {e}",estado="ERROR", nombreTarea="Abrir SAP Logon", )
        # WriteLog | ERROR | Error grave ConectarSAP
        return None


# ============================================================
# ObtenerSesionActiva
# Autor: Automatizacion RPA
# Descripcion: Obtiene una sesion activa de SAP
# ============================================================
def ObtenerSesionActiva():

    """Obtiene una sesión SAP ya iniciada (con usuario logueado)."""
    try:
        sapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = sapGuiAuto.GetScriptingEngine

        for conn in application.Connections:
            if conn.Children.Count > 0:
                session = conn.Children(0)
                WriteLog( mensaje=f"Sesion encontrada en conexión: {conn.Description}",estado="INFO", nombreTarea="Abrir SAP Logon", )
                return session
        WriteLog( mensaje=f"No se encontró ninguna sesion activa",estado="WARN", nombreTarea="Abrir SAP Logon", )
        return None

    except Exception as e:
        WriteLog( mensaje=f"Error al obtener la sesion activa: {e}",estado="ERROR", nombreTarea="Abrir SAP Logon", )
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
            WriteLog( mensaje=f"buscando imagen de la ruta:{ruta_imagen} Error:{e}",estado="ERROR", nombreTarea="Abrir SAP Logon", )
        time.sleep(espera)
    WriteLog( mensaje=f"imagen de la ruta:{ruta_imagen} no encontrada",estado="WARN", nombreTarea="Abrir SAP Logon", )
    return False

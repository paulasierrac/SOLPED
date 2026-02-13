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
            WriteLog( mensaje="SAP Logon ya se encuentra abierto",estado="WARN", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)
        else:
            WriteLog( mensaje="SAP Logon 750 abierto",estado="INFO", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)
      
        
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
            WriteLog( mensaje=f"Abriendo nueva conexion a {conexion}",estado="INFO", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)
            connection = application.OpenConnection(conexion, True)
        else:
            WriteLog( mensaje=f"Conexion existente encontrada con {conexion}",estado="INFO", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)

        if connection.Children.Count > 0:
            session = connection.Children(0)
            WriteLog( mensaje=f"Sesion existente reutilizada",estado="INFO", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)
        else:
            session = connection.Children(0).CreateSession()
            WriteLog( mensaje=f"Nueva sesion creada",estado="INFO", nombreTarea="Abrir SAP Logon", rutaRegistro=inConfig("PathLog"),)


        # Login
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        session.findById("wnd[0]").sendVKey(0)

        print("INFO | Credenciales enviadas correctamente")

        if ventanaAbierta(session, "Copyrigth"):
            pyautogui.press("enter")
            print("INFO | Ventana Copyrigth cerrada")

        try:
            if validarLoginDiag(
                ruta_imagen=rf".\img\logindiag.png",
                confidence=0.5,
                intentos= int (inConfig("ReIntentos")),
                espera=0.5,
            ):
                print("INFO | Ventana loginDiag superada correctamente")
        except Exception as e:
            print(f"no se encontro ventana Copyrigth en login {e}")

        if ventanaAbierta(session, "Info de licencia en entrada al sistema múltiple"):
            
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
        ControlHU(nombreTarea, estado=99)
        print(f"ERROR | Error al conectar a SAP: {e}")
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

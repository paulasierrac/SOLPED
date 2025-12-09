import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
from Config.settings import RUTAS, SAP_CONFIG

# import pyautogui


def abrir_sap_logon():
    """Abre SAP Logon si no está ya abierto."""
    try:
        # Verificar si SAP ya está abierto
        sapgui = win32com.client.GetObject("SAPGUI")
        return True
    except:
        # Si no está abierto, se lanza el ejecutable
        subprocess.Popen(SAP_CONFIG["logon_path"])
        time.sleep(5)  # Esperar a que abra SAP Logon
        return False


def conectar_sap(conexion, mandante, usuario, password, idioma="ES"):

    abrir_sap = abrir_sap_logon()
    if abrir_sap:
        print(" SAP Logon 750 ya se encuentra abierto")
    else:
        print(" SAP Logon 750 abierto ")

    try:
        print("Iniciando conexion con SAP...")

        # 1️⃣ Obtener objeto SAPGUI
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            raise Exception(
                "No se pudo obtener el objeto SAPGUI. Asegúrate de que SAP Logon esté instalado y el scripting habilitado."
            )

        application = sap_gui_auto.GetScriptingEngine  # motor de Scripting

        # 2️⃣ Buscar conexión activa
        # application.Connections → lista de conexiones (entradas en SAP Logon).
        connection = None
        for item in application.Connections:
            if item.Description.strip().upper() == conexion.strip().upper():
                connection = item
                break

        # 3️⃣ Si no existe conexión abierta, abrir una nueva
        if not connection:
            print(f"Abriendo nueva conexion a {conexion}...")
            connection = application.OpenConnection(conexion, True)
            time.sleep(3)  # Esperar que abra
        else:
            print(f"✅ Conexion existente encontrada con {conexion}.")

        # 4️⃣ Verificar sesión
        if connection.Children.Count > 0:
            session = connection.Children(0)
            print("Sesion existente reutilizada.")
        else:
            session = connection.Children(0).CreateSession()
            print(" Nueva sesion creada.")

        # 5️⃣ Si la pantalla está en login, ingresar credenciales
        # if "RSYST-BNAME" in session.findById("wnd[0]/usr").Text:
        #     print("🧩 Ingresando credenciales...")
        # if password is None:
        #     password = getpass.getpass("Contraseña SAP: ")
        # Ingresar datos de login
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        session.findById("wnd[0]").sendVKey(0)
        print(" Conectado correctamente a SAP.")

        return session

    except Exception as e:
        print(f" Error al conectar a SAP: {e}")
        return None


def ObtenerSesionActiva():
    """Obtiene una sesión SAP ya iniciada (con usuario logueado)."""
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        # Buscar una conexión activa con sesión
        for conn in application.Connections:
            if conn.Children.Count > 0:
                session = conn.Children(0)
                print(f" Sesion encontrada en conexión: {conn.Description}")
                return session

        print(" No se encontró ninguna sesion activa.")
        return None

    except Exception as e:
        print(f" Error al obtener la sesion activa: {e}")
        return None

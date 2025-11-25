# Prueba nuemro 2 
import win32com.client
import subprocess
import time
from Config.settings import SAP_CONFIG
from HU.HU1_LoginSAP import obtener_sesion_activa

subprocess.Popen(SAP_CONFIG["logon_path"])
time.sleep(5)
sapgui = win32com.client.GetObject("SAPGUI")
application = sapgui.GetScriptingEngine 
connection = application.OpenConnection(SAP_CONFIG["sistema"], True)
session= connection.Children(0)

#session = obtener_sesion_activa()

try:
        #campo = session.findById("wnd[0]/usr/pwdRSYST-BCODE")
        #campo.text = "sT1f%4L*"
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = SAP_CONFIG["mandante"]
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_CONFIG["user"]
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_CONFIG["password"]
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = SAP_CONFIG["idioma"]
        session.findById("wnd[0]").sendVKey(0)
        print(f"Campo modificado correctamente.{SAP_CONFIG["password"]}")
except Exception as e:
        print(f"❌ No se pudo escribir en el campo password: {e}")


if session:
    print("Conexion establecida, listo para ejecutar transacciones.")
else:
    print("No se pudo establecer la conexión.")


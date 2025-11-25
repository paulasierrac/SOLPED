import win32com.client
import time
import pyautogui
import subprocess


def ObtenerSesionActiva():
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        return session
    except:
        print("No fue posible obtener la sesión activa.")
        return None


session = ObtenerSesionActiva()

if session:
    try:
        # 1) Obtener el objeto del editor
        editor = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        # 1. Tomar el texto completo del editor
        texto = editor.text
        print(texto)
        # 2. Guardarlo directamente en un archivo
        path = r"C:\Users\CGRPA009\Documents\texto_sap.txt"
        with open(path, "w", encoding="utf-8") as f:
            f.write(texto)

        print("Botón presionado correctamente.")
    except Exception as e:
        print(f"Error al presionar el botón: {e}")

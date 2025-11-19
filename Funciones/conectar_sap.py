import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
#import pyautogui

def abrir_sap_logon():
    """Abre SAP Logon si no está ya abierto."""
    try:
        # Verificar si SAP ya está abierto
        sapgui = win32com.client.GetObject("SAPGUI")
        return True
    except:
        # Si no está abierto, se lanza el ejecutable
        subprocess.Popen(r'"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"')
        time.sleep(5)  # Esperar a que abra SAP Logon
        return False
    
def conectar_sap(conexion, mandante, usuario, password, idioma="ES"):
    try:
        print("Iniciando conexion con SAP...")

        # 1️⃣ Obtener objeto SAPGUI
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            raise Exception("No se pudo obtener el objeto SAPGUI. Asegúrate de que SAP Logon esté instalado y el scripting habilitado.")

        application = sap_gui_auto.GetScriptingEngine #motor de Scripting

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
        if "RSYST-BNAME" in session.findById("wnd[0]/usr").Text:
            print("🧩 Ingresando credenciales...")
        if password is None:
            password = getpass.getpass("Contraseña SAP: ")
 

        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        time.sleep(10)
        #pyautogui.press('Tab')
        #pyautogui.write(password, interval=0.1)
        #session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        #pyautogui.press('Enter')
        #session.findById("wnd[0]").sendVKey(0)
        #session.findById("wnd[0]").resizeWorkingPane
        #103, 16, false
        print(" Conectado correctamente a SAP.")  
      
        return session

    except Exception as e:
        print(f" Error al conectar a SAP: {e}")
        return None

def obtener_sesion_activa():
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

def descarga_solpedME5A(session, transaccion):
    if session:
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        print("Transacción ME5A abierta con éxito.")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "ALV"
        session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ZSUA"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ZSOL"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "ZSUB"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "ZSU3"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").setFocus
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").text = "03"
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").caretPosition = 2
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        #session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()

        ruta_guardar = r"C:\Users\CGRPA042\Desktop\AutomatizacionGestionSolped\export_solped.txt"
        if os.path.exists(ruta_guardar):
            os.remove(ruta_guardar)

        #session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]").select()

        # Seleccionar tipo de exportación (generalmente “Spreadsheet”)
        #session.findById("wnd[1]/usr/radRB_OTHERS").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === GUARDAR ARCHIVO ===
        time.sleep(1)
        #session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(ruta_guardar)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\CGRPA042\Desktop\AutomatizacionGestionSolped"
        #session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(ruta_guardar)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # “Guardar”
        print("✅ Archivo exportado y guardado en:", ruta_guardar)

        



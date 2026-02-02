# ============================================
# Función Local: DescargarSolpedME5A
# Autor: Tu Nombre - Configurador RPA
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import pyautogui
from config.settings import RUTAS
from funciones.GeneralME53N import AbrirTransaccion
import win32com.client
import time
import os


def DescargarSolpedME5A(session, estado):

    if not session:
        raise ValueError("Sesión SAP no válida.")
    
    # Ruta destino – ejemplo estándar Colsubsidio
    #ruta_guardar = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped{estado}.txt"
    ruta_guardar = rf"{RUTAS["PathInsumo"]}\expSolped{estado}.txt"
    # ============================
    # Abrir transacción ME5A
    # ============================
    AbrirTransaccion(session, "ME5A")
    print("Transacción ME5A abierta con éxito.")
    session.findById("wnd[0]").maximize()

    # ============================
    # Visual.lista Solicitudes de pedido 
    # ============================

    # Alcance de la lista
    session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "ALV"

    # Clase de documento
    session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press()

    # Tabla de selección
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,0]"
    ).text = "ZSUA"

    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,1]"
    ).text = "ZSOL"

    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,2]"
    ).text = "ZSUB"

    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,3]"
    ).text = "ZSU3"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,3]"
    ).setFocus
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
        "ctxtRSCSEL_255-SLOW_I[1,3]"
    ).caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Aceptar selección
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Ejecutar

    # ============================
    # Aplicar Filtro de Estado 03 o 05 
    # ============================
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").text = estado
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").caretPosition = 2
    session.findById("wnd[0]").sendVKey(0)

    # Ejecutar F8
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Exportar
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    time.sleep(5)
    # ============================
    # Guardar archivo , revisar rutas relativas 
    # ============================

    #ruta_guardar = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped{estado}.txt"
    ruta_guardar = rf"{RUTAS["PathInsumo"]}\expSolped{estado}.txt"  

    if os.path.exists(ruta_guardar):
        os.remove(ruta_guardar)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = (
        #r"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo"
        rf"{RUTAS["PathInsumo"]}"
    )

    ruta_guardar = rf"{RUTAS["PathInsumo"]}\expSolped{estado}.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = rf"expSolped{estado}.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar
    time.sleep(1)

    # Salir de SAP 
    session.findById("wnd[0]").sendVKey(12)
    time.sleep(1)
    pyautogui.press("f3")
    time.sleep(1)
    pyautogui.press("f12")
    print(
        f"Archivo exportado correctamente: {ruta_guardar}"
    )  # luego reemplazar con WriteLog

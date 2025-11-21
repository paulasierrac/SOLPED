import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
#import pyautogui

def descarga_solpedME5A(session, estado):
    if session:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME5A"
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
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").text = estado
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").caretPosition = 2
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        #session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()

        ruta_guardar = fr"C:\NetApplications\PY\AutomatizacionGestionSolped\Input\expSolped{estado}.txt"

        if os.path.exists(ruta_guardar):
            os.remove(ruta_guardar)

        #session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]").select()

        # Seleccionar tipo de exportación (generalmente “Spreadsheet”)
        #session.findById("wnd[1]/usr/radRB_OTHERS").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === GUARDAR ARCHIVO ===
        time.sleep(1)
        #session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(ruta_guardar)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\NetApplications\PY\AutomatizacionGestionSolped\Input"
        #session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(ruta_guardar)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fr"expSolped{estado}.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # “Guardar”
        print("Archivo exportado y guardado en:", ruta_guardar)
        session.findById("wnd[0]").sendVKey(12)
        time.sleep(0.5)

# ============================================
# Función Local: DescargarSolpedME5A
# Autor: Tu Nombre - Configurador RPA
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client  
import time
import os

def DescargarSolpedME5A(session, estado):

    if not session:
        raise ValueError("Sesión SAP no válida.")

    # Ruta destino – ejemplo estándar Colsubsidio
    ruta_guardar = fr"C:\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped{estado}.txt"

    # ============================
    # Abrir transacción ME5A
    # ============================
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nME5A"
    session.findById("wnd[0]").sendVKey(0)
    print("Transacción ME5A abierta con éxito.")
    session.findById("wnd[0]").maximize()

    # ============================
    # Configuración de variante ALV
    # ============================
    session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "ALV"

    # Tipo de documento
    session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press()

    # Tabla de selección
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,0]").text = "ZSUA"

    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,1]").text = "ZSOL"

    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,2]").text = "ZSUB"

    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,3]").text = "ZSU3"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,3]").setFocus
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
                     "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                     "ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Aceptar selección
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Ejecutar

    # ============================
    # Aplicar Filtro de Estado
    # ============================
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").text = estado
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").caretPosition = 2
    session.findById("wnd[0]").sendVKey(0)

    # Ejecutar ALV
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Exportar
    session.findById("wnd[0]/tbar[1]/btn[45]").press()

    # ============================
    # Guardar archivo
    # ============================
    
    ruta_guardar = fr"C:\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped{estado}.txt"
    if os.path.exists(ruta_guardar):
        os.remove(ruta_guardar)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\NetApplications\PY\AutomatizacionGestionSolped\Insumo"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fr"expSolped{estado}.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar
    time.sleep(1)
    
    print("Archivo exportado y guardado en:", ruta_guardar)
    session.findById("wnd[0]").sendVKey(12)
    time.sleep(0.5)
    
    print(f"Archivo exportado correctamente: {ruta_guardar}")  # luego reemplazar con WriteLog
# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU01_LoginSAP import conectar_sap
from HU.HU05_GeneracionOC import GenerarOCDesdeSolped


# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Main_Pruebas3():
    try:
  
        session = conectar_sap(
         SAP_CONFIG["sistema"], 
         SAP_CONFIG["mandante"],
         SAP_CONFIG["user"],
         SAP_CONFIG["password"],
         SAP_CONFIG["idioma"]
         )
        
          
        
        
        
        #GenerarOCDesdeSolped(session, "1300139102", 2)  # Reemplaza con la Solped real:  1300139102, 2  1300139269 , 6
        #GenerarOCDesdeSolped(session, "1300139269", 6)  # Reemplaza con la Solped real:  1300139102, 2  1300139269 , 6   1300138077, 10
        GenerarOCDesdeSolped(session, "1300138077", 10)
       
        # Solped compartidas por el grupo
        #GenerarOCDesdeSolped(session, "1300139390", 7)   #no tiene dato organizacion de compra 
        #GenerarOCDesdeSolped(session, "1300139391", 9)
        #GenerarOCDesdeSolped(session, "1300139392", 4)  # no tiene registros 
        #GenerarOCDesdeSolped(session, "1300139393", 7)
        #GenerarOCDesdeSolped(session, "1300139394", 7)




    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL: {e} | {error_text}",
            estado="ERROR",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLogError"],
        )
        raise


if __name__ == "__main__":
    Main_Pruebas3()

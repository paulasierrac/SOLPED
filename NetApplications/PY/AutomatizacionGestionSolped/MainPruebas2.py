# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU01_LoginSAP import conectar_sap
from HU.HU04_GeneracionOC import EjecutarHU04
from Funciones.ValidacionM21N import BorrarTextosDesdeSolped
from HU.HU05_DescargaOCME9F import EjecutarHU04


# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Main_Pruebas2():
    try:  
        session = conectar_sap(
         SAP_CONFIG["sistema"], 
         SAP_CONFIG["mandante"],
         SAP_CONFIG["user"],
         SAP_CONFIG["password"],
         SAP_CONFIG["idioma"]
         )
          
        #EjecutarHU05(session)
        print(session)

      
        #solpeds = [("1300139102", 2),("1300139269", 6),("1300138077", 10),("1300139272", 10),("1300136848", 83)]
        # Solped compartidas por el grupo
        ListaOC = ["42003400000","4200340002","4200340003","4200340004","4200340005",]
        for orden in ListaOC:
           EjecutarHU04(session, orden)

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
    Main_Pruebas2()


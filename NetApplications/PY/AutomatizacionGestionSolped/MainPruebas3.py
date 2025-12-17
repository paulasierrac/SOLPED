# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU01_LoginSAP import conectar_sap
from Funciones.ValidacionM21N import BorrarTextosDesdeSolped

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
          
        #EjecutarHU05(session)
        print(session)

      
        #solpeds = [("1300139102", 2),("1300139269", 6),("1300138077", 10),("1300139272", 10),("1300136848", 83)]
        # Solped compartidas por el grupo
        solpeds = [("1300139394", 7),("1300139391", 9),("1300139392", 4),("1300139393", 7),("1300139390", 7)]
        #solpeds = [("1300139269", 6)]

        for solped, posicion in solpeds:
            BorrarTextosDesdeSolped(session, solped, posicion)

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

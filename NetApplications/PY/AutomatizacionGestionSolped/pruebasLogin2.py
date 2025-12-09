# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import ObtenerSesionActiva, conectar_sap, abrir_sap_logon
from HU.HU02_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU05_GeneracionOC import GenerarOCDesdeSolped
from HU.HU04_DescargaOCME9F import descarga_OCME9F

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS, SAP_CONFIG


def Main_Pruebas3():
    try:

        session = conectar_sap(
            SAP_CONFIG["sistema"],
            SAP_CONFIG["mandante"],
            SAP_CONFIG["user"],
            SAP_CONFIG["password"],
            SAP_CONFIG["idioma"],
        )

        GenerarOCDesdeSolped(
            session, "1300139102", 2
        )  # Reemplaza con la Solped real:  1300139102, 2  1300139269 , 6
        # GenerarOCDesdeSolped(session, "1300139269", 6)  # Reemplaza con la Solped real:  1300139102, 2  1300139269 , 6   1300138077, 10
        # GenerarOCDesdeSolped(session, "1300138077", 10)
        # GenerarOCDesdeSolped(session, "1300177338", 13)

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

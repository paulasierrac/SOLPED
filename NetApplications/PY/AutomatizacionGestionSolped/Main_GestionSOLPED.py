# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU1_LoginSAP import ObtenerSesionActiva
from HU.HU2_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS


def Main_GestionSOLPED():
    try:
        # Despliegue de ambiente
        EjecutarHU00()

        # Capturar sesion SAP
        session = ObtenerSesionActiva()

        # HU2 Descarga ME5A
        WriteLog(
            mensaje="Inicia HU02",
            estado="INFO",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLog"],
        )
        EjecutarHU02(session)

        # HU2 validacion ME5AN
        WriteLog(
            mensaje="Inicia HU03",
            estado="INFO",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLog"],
        )
        EjecutarHU03(session)
        # WriteLog(activar_log=config["ActivarLog"],path_log=config["PathLog"],mensaje="Inicia HU02",estado="INFO")

        # EjecutarHU02(config)

        # WriteLog( activar_log=config["ActivarLog"],path_log=config["PathLog"],mensaje="Finaliza automatización",estado="INFO")

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
    Main_GestionSOLPED()

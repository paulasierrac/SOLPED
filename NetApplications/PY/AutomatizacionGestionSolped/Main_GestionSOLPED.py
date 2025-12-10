# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot encargado de ejecutar las historias de usuario
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios:
#   - Reemplazo de print() por WriteLog
#   - Cumplimiento estricto de estándar Colsubsidio
#   - Manejo de excepciones y log por día
# ================================

from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import (
    ObtenerSesionActiva,
    conectar_sap,
)
from HU.HU02_DescargaME5A import (
    EjecutarHU02,
)
from HU.HU03_ValidacionME53N import EjecutarHU03
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
    EnviarNotificacionCorreo,
    EnviarCorreoPersonalizado,
    NotificarRevisionManualSolped,
)
from Config.settings import RUTAS, SAP_CONFIG
import traceback


def Main_GestionSolped():
    try:
        task_name = "Main_GestionSOLPED"

        # ================================
        # Inicio de Main
        # ================================
        WriteLog(
            mensaje="Inicio ejecución Main GestionSolped.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 1. Despliegue de ambiente
        # ================================
        WriteLog(
            mensaje="Inicia HU00_DespliegueAmbiente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        # EjecutarHU00()

        # ================================
        # 2. Obtener sesión SAP
        # ================================
        WriteLog(
            mensaje="Obteniendo sesión SAP...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        session = conectar_sap(
            SAP_CONFIG["sistema"],
            SAP_CONFIG["mandante"],
            SAP_CONFIG["user"],
            SAP_CONFIG["password"],
            "EN",
        )

        session = ObtenerSesionActiva()

        WriteLog(
            mensaje="Sesión SAP obtenida correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 3. Ejecutar HU02 – Descarga ME5A
        # ================================
        WriteLog(
            mensaje="Inicia HU02 - Descarga ME5A.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # EjecutarHU02(session)

        WriteLog(
            mensaje="HU02 finalizada correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 4. Ejecutar HU03 – Validación ME53N
        # ================================
        archivos_validar = ["expSolped03.txt", "expSolped05.txt"]

        for archivo in archivos_validar:
            WriteLog(
                mensaje=f"Inicia HU03 - Validación ME53N para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            EjecutarHU03(session, archivo)

            WriteLog(
                mensaje=f"HU03 finalizada correctamente para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Notificación de finalización HU02 con archivo descargado (código 2)

        # ================================
        # Fin de Main
        # ================================
        WriteLog(
            mensaje="Main GestionSolped finalizado correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        error_stack = traceback.format_exc()
        WriteLog(
            mensaje=f"Error Global en Main: {e} | {error_stack}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise


if __name__ == "__main__":
    Main_GestionSolped()

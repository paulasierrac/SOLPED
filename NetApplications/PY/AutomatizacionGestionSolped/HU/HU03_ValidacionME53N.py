# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión inicial
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import time
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from Funciones.ValidacionME53N import ValidacionME53N


def EjecutarHU03(session):
    """session: objeto de SAP GUI
    Realiza la verificacion del SOLPED"""

    try:
        WriteLog(
            mensaje="Inicia HU03",
            estado="INFO",
            task_name="HU03_ValidacionME53N",
            path_log=RUTAS["PathLog"],
        )
        numero_solped = 1300139306
        ValidacionME53N(session, numero_solped)
        return True

    except Exception as e:
        WriteLog(
            mensaje=f"Error en HU03_BuscarSolpedME53N: {e}",
            estado="ERROR",
            task_name="HU03_ValidacionME53N",
            path_log=RUTAS["PathLogError"],
        )

        return False

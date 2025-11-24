# =========================================
# NombreDeLaIniciativa – HUxx: BuscarSolpedME53N
# Autor: TuNombre, Empresa, Rol
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión inicial
# =========================================
import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import time
from Funciones.EscribirLog import WriteLog

def BuscarSolpedME53N(session, config, numero_solped):
    """
    session: objeto de SAP GUI
    config: diccionario In_Config con parámetros
    numero_solped: número de la solped a consultar
    """

    try:
        WriteLog(
            activar_log=config["ActivarLog"],
            path_log=config["PathLog"],
            mensaje="Inicia HUxx - BuscarSolpedME53N",
            estado="INFO"
        )

        # Validar sesión SAP
        if session is None:
            WriteLog(
                activar_log=config["ActivarLog"],
                path_log=config["PathLog"],
                mensaje="Sesión SAP no disponible",
                estado="ERROR"
            )
            raise Exception("Sesión SAP no disponible")

        # Abrir transacción ME53N
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME53N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            activar_log=config["ActivarLog"],
            path_log=config["PathLog"],
            mensaje="Transacción ME53N abierta",
            estado="INFO"
        )

        # Ingresar número de SOLPED
        campo_solped = "wnd[0]/usr/ctxtRM06E-BANFN"
        session.findById(campo_solped).text = numero_solped
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            activar_log=config["ActivarLog"],
            path_log=config["PathLog"],
            mensaje=f"Solped {numero_solped} consultada exitosamente",
            estado="INFO"
        )

        return True

    except Exception as e:
        WriteLog(
            activar_log=config.get("ActivarLog", True),
            path_log=config.get("PathLog", "Audit/Logs/error_hu.txt"),
            mensaje=f"Error en HUxx_BuscarSolpedME53N: {e}",
            estado="ERROR"
        )
        return False


# def buscar_SolpedME53N(session):        
#     if session:
#         session.findById("wnd[0]/tbar[0]/okcd").text = ""
#         session.findById("wnd[0]").sendVKey(0)
#         print("Transacción ME5A abierta con éxito.")


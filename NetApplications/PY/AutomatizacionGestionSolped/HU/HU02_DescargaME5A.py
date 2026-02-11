# ============================================
# HU02: Descargar Solicitudes de Pedido ME5A
# Autor: Henry - Configurador RPA
# Descripcion: Descarga las solicitudes de pedido filtradas por estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
from Funciones.DescargarSolpedME5A import DescargarSolpedME5A
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS
from Funciones.ControlHU import control_hu


def EjecutarHU02(session):
    """
    session: objeto de SAP GUI

    Ejecuta la Historia de Usuario 02 encargada de la
    descarga de SOLPED desde la transacción ME5A.
    """
    try:
        task_name = "HU02_DescargaME5A"
        control_hu(task_name=task_name, estado=0)
        WriteLog(
            mensaje="Inicia HU02",
            estado="INFO",
            task_name="HU2_DescargaME5A",
            path_log=RUTAS["PathLog"],
        )
        estado = "03"
        DescargarSolpedME5A(session, estado)
        estado = "05"
        DescargarSolpedME5A(session, estado)
        control_hu(task_name=task_name, estado=100)
    except Exception as e:
        control_hu(task_name=task_name, estado=99)
        WriteLog(
            mensaje=f"ERROR GLOBAL: {e}",
            estado="ERROR",
            task_name="HU2_DescargaME5A",
            path_log=RUTAS["PathLogError"],
        )
        raise

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
from Funciones.ControlHU import ControlHU
import traceback
from Config.settings import RUTAS


def EjecutarHU02(session):
    """
    session: objeto de SAP GUI

    Ejecuta la Historia de Usuario 02 encargada de la
    descarga de SOLPED desde la transacción ME5A.
    """
    try:
        nombreTarea = "HU02_DescargaME5A"
        ControlHU(nombreTarea, estado=0)
        WriteLog(
            mensaje="Inicia HU02",
            estado="INFO",
            nombreTarea="HU2_DescargaME5A",
            rutaRegistro=inConfig("PathLog"),
        )
        estado = "03"
        DescargarSolpedME5A(session, estado)
        estado = "05"
        DescargarSolpedME5A(session, estado)
        ControlHU(nombreTarea, estado=100)
    except Exception as e:
        ControlHU(nombreTarea, estado=99)
        WriteLog(
            mensaje=f"ERROR GLOBAL: {e}",
            estado="ERROR",
            nombreTarea="HU2_DescargaME5A",
            rutaRegistro=RUTAS["PathLogError"],
        )
        raise

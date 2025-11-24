# ============================================
# HU02: Descargar Solicitudes de Pedido ME5A
# Autor: Tu Nombre - Configurador RPA
# Descripcion: Descarga las solicitudes de pedido filtradas por estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
from Funciones.DescargarSolpedME5A import DescargarSolpedME5A
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS

def EjecutarHU02(session):
    """
    Ejecuta la Historia de Usuario 02 encargada de la
    descarga de SOLPED desde la transacción ME5A.
    """
    try:
        estado = "03"
        DescargarSolpedME5A(session, estado) 
        estado = "05"
        DescargarSolpedME5A(session, estado)
        
    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(mensaje=f"ERROR GLOBAL: {e} | {error_text}",estado="ERROR",task_name="Main_GestionSOLPED",path_log=RUTAS["PathLogError"])
        raise

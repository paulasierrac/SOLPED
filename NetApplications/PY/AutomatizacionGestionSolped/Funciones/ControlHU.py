from Repositories.ControlHU import ControlHURepo
import socket
import re

def extraer_hu(iNombreHU: str) -> int:
    match = re.match(r'HU(\d+)', iNombreHU.upper())
    if not match:
        raise ValueError(f"Nombre de HU invalido: {iNombreHU}")
    return int(match.group(1))

def control_hu(iNombreHU: str, estado: int):
    
    iHuId = extraer_hu(iNombreHU)

    if estado == 0:
        activa = 1
    elif estado in (99, 100):
        activa = 0
    else:
        activa = 1
    
    maquina = socket.gethostname()

    ControlHURepo.actualizar_estado_hu(
        iHuId=iHuId,
        iNombreHU=iNombreHU,
        estado=estado,
        activa=activa,
        maquina=maquina
    )

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

    #repo = ControlHURepo(schema="GestionSolped")
    repo = ControlHURepo(schema=None)


    Prueba1 = repo.ActualizarEstadoHU(iHuId=iHuId, iNombreHU=iNombreHU, estado=estado, activa=activa, maquina=maquina)

    if Prueba1:
        print(f"Verdadero HU Actualizada {iHuId}")
    else:
        print(f"Falso HU Actualizada {iHuId}")
 
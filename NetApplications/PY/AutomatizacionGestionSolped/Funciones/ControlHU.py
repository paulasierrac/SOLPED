from repositories.ControlHU import ControlHURepository
import socket
import re

def extraer_hu(nombre_hu: str) -> int:
    match = re.match(r'HU(\d+)', nombre_hu.upper())
    if not match:
        raise ValueError(f"Nombre de HU invalido: {nombre_hu}")
    return int(match.group(1))

def control_hu(nombre_hu: str, estado: int):
    
    hu_id = extraer_hu(nombre_hu)

    if estado == 0:
        activa = 1
    elif estado in (99, 100):
        activa = 0
    else:
        activa = 1
    
    maquina = socket.gethostname()

    ControlHURepository.upsert_control_hu(
        hu_id=hu_id,
        nombre_hu=nombre_hu,
        estado=estado,
        activa=activa,
        maquina=maquina
    )
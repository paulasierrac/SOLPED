import os
from datetime import datetime
import socket
import getpass


def EscribirIformeOperacion(
    itemCount  :int,
    solped: str,
    ordenCompra: str,
    acciones: list,
    estado: str,
    botName: str,
    nombreTarea: str,
    pathInformes: str,
    observaciones: str = ""
):
    """
    Genera un informe de negocio para envío por correo

    solped        : Número de SOLPED
    ordenCompra  : Número de Orden de Compra generada
    itemCount    : Cantidad de posiciones en la SOLPED
    acciones      : Lista de acciones realizadas (strings)
    estado        : EXITOSO | PARCIAL | ERROR
    botName      : Nombre del Bot RPA (ej. 'Resock')
    nombreTarea     : HU o proceso (ej. 'HU05_GeneracionOC')
    pathInformes : Carpeta donde se guardará el informe
    observaciones : Texto libre opcional
    """

    # === Fecha ===
    ahora = datetime.now()
    fecha_hora = ahora.strftime("%d/%m/%Y %H:%M:%S")
    fechaArchivo = ahora.strftime("%Y%m%d_%H%M%S")

    # === Sistema ===
    nombre_maquina = socket.gethostname()
    usuario = getpass.getuser()

    # === Asegurar carpeta ===
    os.makedirs(pathInformes, exist_ok=True)

    # === Nombre archivo ===
    nombreArchivo = (
        f"Informe_{botName}_SOLPED_{solped}_OC_{ordenCompra}_{fechaArchivo}.txt"
    )
    rutaArchivo = os.path.join(pathInformes, nombreArchivo)

    # === Construcción del informe ===
    contenido = []
    contenido.append("INFORME AUTOMATIZACIÓN – MODIFICACIÓN SOLPED / GENERACIÓN OC\n")
    contenido.append("=" * 65 + "\n\n")

    contenido.append(f"Fecha ejecución   : {fecha_hora}\n")
    contenido.append(f"Usuario ejecución : {usuario}\n")
    contenido.append(f"Equipo            : {nombre_maquina}\n")
    contenido.append(f"Bot RPA           : {botName}\n")
    contenido.append(f"Proceso           : {nombreTarea}\n\n")
    contenido.append(f"SOLPED            : {solped}\n")
    contenido.append(f"ORDEN DE COMPRA   : {ordenCompra}\n")
    contenido.append(f"POSICIONES SOLPED : {itemCount}\n")
    contenido.append(f"Estado final      : {estado}\n\n")

    contenido.append("Acciones realizadas:\n")
    for acc in acciones:
        contenido.append(f" - {acc}\n")

    if observaciones:
        contenido.append("\nObservaciones:\n")
        contenido.append(f"{observaciones}\n")

    contenido.append("\n" + "=" * 65 + "\n")
    contenido.append("Informe generado automáticamente por RPA.\n")

    # === Guardar ===
    with open(rutaArchivo, "w", codificacion="utf-8") as f:
        f.writelines(contenido)

    return rutaArchivo
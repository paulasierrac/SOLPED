import os
import datetime
import socket
import getpass


def WriteInformeOperacion(
    solped: str,
    orden_compra: str,
    acciones: list,
    estado: str,
    bot_name: str,
    task_name: str,
    path_informes: str,
    observaciones: str = ""
):
    """
    Genera un informe de negocio para envío por correo

    solped        : Número de SOLPED
    orden_compra  : Número de Orden de Compra generada
    acciones      : Lista de acciones realizadas (strings)
    estado        : EXITOSO | PARCIAL | ERROR
    bot_name      : Nombre del Bot RPA (ej. 'Resock')
    task_name     : HU o proceso (ej. 'HU05_GeneracionOC')
    path_informes : Carpeta donde se guardará el informe
    observaciones : Texto libre opcional
    """

    # === Fecha ===
    ahora = datetime.datetime.now()
    fecha_hora = ahora.strftime("%d/%m/%Y %H:%M:%S")
    fecha_archivo = ahora.strftime("%Y%m%d_%H%M%S")

    # === Sistema ===
    nombre_maquina = socket.gethostname()
    usuario = getpass.getuser()

    # === Asegurar carpeta ===
    os.makedirs(path_informes, exist_ok=True)

    # === Nombre archivo ===
    nombre_archivo = (
        f"Informe_{bot_name}_SOLPED_{solped}_OC_{orden_compra}_{fecha_archivo}.txt"
    )
    ruta_archivo = os.path.join(path_informes, nombre_archivo)

    # === Construcción del informe ===
    contenido = []
    contenido.append("INFORME AUTOMATIZACIÓN – MODIFICACIÓN SOLPED / GENERACIÓN OC\n")
    contenido.append("=" * 65 + "\n\n")

    contenido.append(f"Fecha ejecución   : {fecha_hora}\n")
    contenido.append(f"Usuario ejecución : {usuario}\n")
    contenido.append(f"Equipo            : {nombre_maquina}\n")
    contenido.append(f"Bot RPA           : {bot_name}\n")
    contenido.append(f"Proceso           : {task_name}\n\n")

    contenido.append(f"SOLPED            : {solped}\n")
    contenido.append(f"ORDEN DE COMPRA   : {orden_compra}\n")
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
    with open(ruta_archivo, "w", encoding="utf-8") as f:
        f.writelines(contenido)

    return ruta_archivo
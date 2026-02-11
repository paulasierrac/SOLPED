# ================================
# Funcion: WriteLog
# Autor: Paula Sierra - NetApplications
# Descripcion: Registrar eventos en archivo log con estructura estándar
# Ultima modificacion: 11/02/2026
# Propiedad de Colsubsidio
# Cambios:
# - Se agregan validaciones
# - Se corrige creación de carpetas
# - Se normaliza estructura línea
# ================================

import datetime
import os
import getpass
import socket
from Config.InicializarConfig import inConfig


def WriteLog(mensaje: str, estado: str, taskName: str, pathLog: str):
    """
    mensaje  : Texto del log
    estado   : INFO, DEBUG, WARN, ERROR
    taskName: Nombre de HU o Main
    pathLog : Ruta de carpeta Logs o archivo .log
    """

    try:
        # ==========================================================
        # 1. Validaciones básicas
        # ==========================================================
        if not mensaje:
            mensaje = "Mensaje vacío"

        if not taskName:
            taskName = "TaskNoDefinida"

        estado = estado.upper()
        if estado not in ["INFO", "DEBUG", "WARN", "ERROR"]:
            estado = "INFO"

        # ==========================================================
        # 2. Fecha
        # ==========================================================
        ahora = datetime.datetime.now()
        fecha_linea = ahora.strftime("%d/%m/%Y %H:%M:%S")
        fechaArchivo = ahora.strftime("%Y%m%d")

        # ==========================================================
        # 3. Datos del sistema
        # ==========================================================
        nombre_maquina = socket.gethostname()
        usuario = getpass.getuser()

        # ==========================================================
        # 4. Determinar ruta final
        # ==========================================================
        base, extension = os.path.splitext(pathLog)

        if extension:
            carpeta_logs = os.path.dirname(pathLog)
            ruta_archivo = pathLog
        else:
            carpeta_logs = pathLog
            nombreArchivo = f"Log{nombre_maquina}{usuario}{fechaArchivo}.log"
            ruta_archivo = os.path.join(carpeta_logs, nombreArchivo)

        os.makedirs(carpeta_logs, exist_ok=True)

        # ==========================================================
        # 5. Construcción línea estándar
        # ==========================================================
        linea = (
            f"{fecha_linea} | "
            f"{estado} | "
            f"{mensaje} | "
            f"{inConfig('CodigoRobot')} | "
            f"{taskName} | "
            "\n"
        )

        # ==========================================================
        # 6. Escritura
        # ==========================================================
        with open(ruta_archivo, "a", encoding="utf-8") as f:
            f.write(linea)

    except Exception as e:
        # Nunca debe romper el bot por logging
        print(f"ERROR | Fallo WriteLog: {e}")

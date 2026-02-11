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
from Config.init_config import in_config


def WriteLog(mensaje: str, estado: str, task_name: str, path_log: str):
    """
    mensaje  : Texto del log
    estado   : INFO, DEBUG, WARN, ERROR
    task_name: Nombre de HU o Main
    path_log : Ruta de carpeta Logs o archivo .log
    """

    try:
        # ==========================================================
        # 1. Validaciones básicas
        # ==========================================================
        if not mensaje:
            mensaje = "Mensaje vacío"

        if not task_name:
            task_name = "TaskNoDefinida"

        estado = estado.upper()
        if estado not in ["INFO", "DEBUG", "WARN", "ERROR"]:
            estado = "INFO"

        # ==========================================================
        # 2. Fecha
        # ==========================================================
        ahora = datetime.datetime.now()
        fecha_linea = ahora.strftime("%d/%m/%Y %H:%M:%S")
        fecha_archivo = ahora.strftime("%Y%m%d")

        # ==========================================================
        # 3. Datos del sistema
        # ==========================================================
        nombre_maquina = socket.gethostname()
        usuario = getpass.getuser()

        # ==========================================================
        # 4. Determinar ruta final
        # ==========================================================
        base, extension = os.path.splitext(path_log)

        if extension:
            carpeta_logs = os.path.dirname(path_log)
            ruta_archivo = path_log
        else:
            carpeta_logs = path_log
            nombre_archivo = f"Log{nombre_maquina}{usuario}{fecha_archivo}.log"
            ruta_archivo = os.path.join(carpeta_logs, nombre_archivo)

        os.makedirs(carpeta_logs, exist_ok=True)

        # ==========================================================
        # 5. Construcción línea estándar
        # ==========================================================
        linea = (
            f"{fecha_linea} | "
            f"{estado} | "
            f"{mensaje} | "
            f"{in_config('CodigoRobot')} | "
            f"{task_name} | "
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

# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot encargado de ejecutar las historias de usuario
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios:
#   - Reemplazo de print() por WriteLog
#   - Cumplimiento estricto de estándar Colsubsidio
#   - Manejo de excepciones y log por día
# ================================

from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU1_LoginSAP import ObtenerSesionActiva
from HU.HU2_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
    EnviarNotificacionCorreo,
    EnviarCorreoPersonalizado,
    NotificarRevisionManualSolped,
)
from Config.settings import RUTAS
import traceback


def Main_GestionSolped():
    try:
        task_name = "Main_GestionSOLPED"

        # ================================
        # Inicio de Main
        # ================================
        WriteLog(
            mensaje="Inicio ejecución Main GestionSolped.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # # Enviar correo de inicio (código 1)
        # # EnviarNotificacionCorreo(codigo_correo=1, task_name=task_name)
        # archivo_descargado = rf"{RUTAS['PathReportes']}/Reporte_1300139268_10.txt"
        # # Enviar correo de inicio (código 2 adjunto)
        # EnviarNotificacionCorreo(
        #     codigo_correo=54, task_name=task_name, adjuntos=[archivo_descargado]
        # )

        # exito_personalizado = EnviarCorreoPersonalizado(
        #     destinatario="soporte_critico@netapplications.com.co",
        #     asunto="Alerta Crítica: El servicio X ha fallado",
        #     cuerpo=(
        #         "<h1>Error Inesperado</h1>"
        #         "<p>El proceso de sincronización ha fallado en la etapa de validación de datos.</p>"
        #         "<p><strong>Revisar logs en:</strong> \\\\servidor\\logs\\errores.txt</p>"
        #     ),
        #     task_name=task_name,
        #     adjuntos=["C:/Archivos/log_error_20251204.txt"],
        #     cc=["paula.sierra@netapplications.com.co"],
        # )

        # if exito_personalizado:
        #     print(f"Notificación enviada exitosamente exito_personalizado.")
        # else:
        #     print(f"Fallo al enviar la notificación exito_personalizado.")

        # NUMERO_SOLPED = "8000012345"
        # DESTINOS = ["usuario.revision@empresa.com", "supervisor@empresa.com"]
        # RAZONES_VALIDACION = (
        #     "1. El centro de costo asignado no es válido para el tipo de material.\n"
        #     "2. La cantidad solicitada supera el límite sin aprobación especial."
        # )

        # # Llamada a la función
        # exito_notificacion = NotificarRevisionManualSolped(
        #     destinatarios=DESTINOS,
        #     numero_solped=NUMERO_SOLPED,
        #     validaciones=RAZONES_VALIDACION,
        # )

        # exito_notificacion = NotificarRevisionManualSolped(
        #     destinatarios=["usuario.revision@empresa.com", "supervisor@empresa.com"],
        #     numero_solped="8000012345",
        #     validaciones=(
        #         "1. El centro de costo asignado no es válido para el tipo de material.\n"
        #         "2. La cantidad solicitada supera el límite sin aprobación especial."
        #     ),
        # )

        # if exito_notificacion:
        #     print(f"Notificación enviada exitosamente para SOLPED {NUMERO_SOLPED}.")
        # else:
        #     print(f"Fallo al enviar la notificación para SOLPED {NUMERO_SOLPED}.")
        # ================================
        # 1. Despliegue de ambiente
        # ================================
        WriteLog(
            mensaje="Inicia HU00_DespliegueAmbiente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        # EjecutarHU00()

        # ================================
        # 2. Obtener sesión SAP
        # ================================
        WriteLog(
            mensaje="Obteniendo sesión SAP...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        session = ObtenerSesionActiva()

        WriteLog(
            mensaje="Sesión SAP obtenida correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 3. Ejecutar HU02 – Descarga ME5A
        # ================================
        WriteLog(
            mensaje="Inicia HU02 - Descarga ME5A.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # EjecutarHU02(session)

        WriteLog(
            mensaje="HU02 finalizada correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 4. Ejecutar HU03 – Validación ME53N
        # ================================
        archivos_validar = ["expSolped05.txt", "expSolped03.txt"]

        for archivo in archivos_validar:
            WriteLog(
                mensaje=f"Inicia HU03 - Validación ME53N para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            EjecutarHU03(session, archivo)

            WriteLog(
                mensaje=f"HU03 finalizada correctamente para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Notificación de finalización HU02 con archivo descargado (código 2)

        # ================================
        # Fin de Main
        # ================================
        WriteLog(
            mensaje="Main GestionSolped finalizado correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        error_stack = traceback.format_exc()
        WriteLog(
            mensaje=f"Error Global en Main: {e} | {error_stack}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise


if __name__ == "__main__":
    Main_GestionSolped()

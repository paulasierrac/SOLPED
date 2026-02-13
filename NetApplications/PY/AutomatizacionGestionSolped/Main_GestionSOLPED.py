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

from Funciones.EscribirLog import WriteLog
from Funciones.EmailSender import EnviarNotificacionCorreo
#from Funciones.GeneralME53N import AppendHipervinculoObservaciones

from Config.settings import RUTAS, SAP_CONFIG
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import ConectarSAP, ObtenerSesionActiva
from HU.HU02_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU04_GeneracionOC import EjecutarHU04
from HU.HU05_DescargaOC import EjecutarHU05

from Config.InicializarConfig import inConfig
from Config.settings import RUTAS, SAP_CONFIG



def Main_GestionSolped():
    try:
        nombreTarea = "Main_GestionSOLPED"

        EjecutarHU00()

        # ================================  

        # Inicio de Main
        # ================================

        # Enviar correo de inicio
        WriteLog(mensaje="Inicio ejecución Main GestionSolped.", estado="INFO", nombreTarea=nombreTarea,  rutaRegistro=inConfig("PathLog"),)

        #EnviarNotificacionCorreo(codigoCorreo=1, nombreTarea=nombreTarea)

        # ================================
        # 1. Despliegue de ambiente
        # ================================sssss
        WriteLog(
            mensaje="Inicia HU00_DespliegueAmbiente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        
        WriteLog(
            mensaje="Finaliza HU00_DespliegueAmbiente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        # ================================
        # 2. Obtener sesión SAP
        # ================================
        WriteLog(
            mensaje="Inicia HU01_LoginSAP.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        session = ConectarSAP(inConfig("SapSistema"),inConfig("SapMandante") ,SAP_CONFIG["user"],SAP_CONFIG["password"],)

        #session = ObtenerSesionActiva()

        WriteLog(
            mensaje="Finaliza HU01_LoginSAP.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        # ================================
        # 3. Ejecutar HU02 – Descarga ME5A
        # ================================
        WriteLog(
            mensaje="Inicia HU02 - Descarga ME5A.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        
        #EjecutarHU02(session)

        WriteLog(
            mensaje="HU02 finalizada correctamente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        # ================================
        # 4. Ejecutar HU03 – Validación Solped ME53N
        # ================================
        # archivos_validar = ["expSolped03.txt","expSolped03 copy.txt"]
        archivos_validar = [
            "expSolped03.txt"
        ]  # Dos solped para prueba 1300139393  1300139394
        WriteLog(
                mensaje=f"Inicia HU03 - Validación ME53N para archivo.",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
        # for archivo in archivos_validar:
        #     EjecutarHU03(session, archivo)

        WriteLog(
                mensaje=f"Finaliza HU03 - Validación ME53N para archivo.",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
        
        # ================================
        # 5. Ejecutar HU04 – Generación de OC
        # ================================
        WriteLog(
            mensaje="Inicia HU04 - Generación OC.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        archivos_validar = ["expSolped03 copy.txt"] 
        for archivo in archivos_validar:
            EjecutarHU04(session, archivo)

        WriteLog(
            mensaje="HU04 - Generación OC finalizada correctamente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        
        # ================================
        # 5. Ejecutar HU05 – Descarga de OC
        # ================================
        WriteLog(
            mensaje="Inicia HU05 - Descarga OC generadas.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        #EjecutarHU05(session)

        WriteLog(
            mensaje="HU05 - Descarga OC finalizada correctamente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        
        # ================================
        # 5. Ejecutar HU06 – Envío de OC por correo
        # ================================
        WriteLog(
            mensaje="Inicia HU06 - Envío OC por correo.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        # EjecutarHU06(session)

        WriteLog(
            mensaje="HU06 - Envío OC por correo finalizada correctamente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )
        
        EnviarNotificacionCorreo(codigoCorreo=2, nombreTarea=nombreTarea, adjuntos=[])

        WriteLog(
            mensaje="Main GestionSolped finalizado correctamente.",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

    except Exception as e:
        WriteLog(
            mensaje=f"Error Global en Main: {e}",
            estado="ERROR",
            nombreTarea=nombreTarea,
            rutaRegistro=RUTAS["PathLogError"],
        )
        raise


if __name__ == "__main__":
    Main_GestionSolped()

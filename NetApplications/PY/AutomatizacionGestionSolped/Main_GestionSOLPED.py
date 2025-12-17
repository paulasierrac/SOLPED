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
from HU.HU01_LoginSAP import (
    ObtenerSesionActiva,
    conectar_sap,
)
from HU.HU02_DescargaME5A import (
    EjecutarHU02,
)
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU04_GeneracionOC import EjecutarHU04
from Funciones.EscribirLog import WriteLog
from Funciones.ValidacionM21N import (
    leer_solpeds_desde_archivo,
    BorrarTextosDesdeSolped

 
)
from Config.settings import RUTAS, SAP_CONFIG
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

        # ================================
        # 1. Despliegue de ambiente
        # ================================
        WriteLog(
            mensaje="Inicia HU00_DespliegueAmbiente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        #EjecutarHU00()

        # ================================
        # 2. Obtener sesión SAP
        # ================================
        WriteLog(
            mensaje="Obteniendo sesión SAP...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        session = conectar_sap(SAP_CONFIG["sistema"],SAP_CONFIG["mandante"],SAP_CONFIG["user"],SAP_CONFIG["password"],)
        #session = ObtenerSesionActiva()

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

        #EjecutarHU02(session)

        WriteLog(
            mensaje="HU02 finalizada correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 4. Ejecutar HU03 – Validación Solped ME53N 
        # ================================
        archivos_validar = ["expSolped05.txt", "expSolped03.txt"]

        for archivo in archivos_validar:
            WriteLog(
                mensaje=f"Inicia HU03 - Validación ME53N para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            #EjecutarHU03(session, archivo)

            WriteLog(
                mensaje=f"HU03 finalizada correctamente para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Notificación de finalización HU02 con archivo descargado (código 2)

        # ================================
        # 5. Ejecutar HU04 – Creacion de OC
        # ================================

            WriteLog(
                mensaje="HU04 - Creacion de OC desde ME21N.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            archivos_validar = ["expSolped05 1.txt"]
            EjecutarHU04(session, archivos_validar)
            

            WriteLog(
                mensaje=f"HU05 finalizada correctamente para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )


        # Finalizacion de HU4 generacion de OC 

        # ================================
        # 5. Ejecutar HU05 – Descarga de OC y envio de correo 
        # ================================

        archivos_validar = ["expSolped05 1.txt", "expSolped05.txt"]
        
        for archivo in archivos_validar:
            WriteLog(
                mensaje=f"Inicia HU05 - Descarga de OC y envio de correo  {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            ruta = rf"{RUTAS["PathInsumo"]}{archivo}"
            print("Esta es la ruta: ", ruta)
            dataSolpeds = leer_solpeds_desde_archivo(ruta)
                         
            for solped, info in dataSolpeds.items():
                print(f"Solped {solped} tiene {info['items']} items")
                #Cambiar por funcion de descarga de OC 


            WriteLog(
                mensaje=f"HU05 finalizada correctamente para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )


        # Finalizacion de HU5 generacion de OC 

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

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

from time import time




from funciones.GeneralME53N import (
    EnviarNotificacionCorreo,
    EnviarCorreoPersonalizado,
    NotificarRevisionManualSolped,
    convertir_txt_a_excel,
    NotificarRevisionManualSolped,
)
from funciones.EscribirLog import WriteLog

from config.settings import RUTAS, SAP_CONFIG
from config.initconfig import in_config

from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import conectar_sap
from HU.HU02_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU04_GeneracionOC import EjecutarHU04
from HU.HU05_DescargaOC import EjecutarHU05

from funciones.GuiShellFunciones import leer_solpeds_desde_archivo
from config.settings import RUTAS, SAP_CONFIG
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
        EjecutarHU00()

        # ================================
        # 2. Obtener sesión SAP
        # ================================
        WriteLog(
            mensaje="Obteniendo sesión SAP...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )


        session = conectar_sap(in_config("SAP_SISTEMA"),in_config("SAP_MANDANTE") ,SAP_CONFIG["user"],SAP_CONFIG["password"],)
     

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
        ordenes_de_compra = ["4200339200", "4200339201", "4200339202", "4200339203", "4200339204", "4200339205", "4200339206"]
        EjecutarHU05(session,ordenes_de_compra)

        WriteLog(
            mensaje="HU02 finalizada correctamente.",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ================================
        # 4. Ejecutar HU03 – Validación Solped ME53N
        # ================================
        #archivos_validar = ["expSolped03.txt","expSolped03 copy.txt"]
        archivos_validar = ["expSolped03.txt"] # Dos solped para prueba 1300139393  1300139394

        for archivo in archivos_validar:
            WriteLog(
                mensaje=f"Inicia HU03 - Validación ME53N para archivo {archivo}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            #EjecutarHU03(session, archivo)
            # convertir_txt_a_excel(archivo)

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
        # TODO - revisar si es necesario EL LOG DE INICIO HU04 por cada archivo o solo una vez

            # WriteLog(
            #     mensaje="Inicia HU04 - Creacion de OC desde ME21N.",
            #     estado="INFO",
            #     task_name=task_name,
            #     path_log=RUTAS["PathLog"],
            # )

            #archivos_validar = ["expSolped05 1.txt"] # 1300139271,1300139272
            archivos_validar = ["expSolped03 copy.txt"] # CAMBIAR A 05 PARA SOLPED LIBERADAS
            #archivos_validar = ["expSolped03.txt"] # CAMBIAR A 05 PARA SOLPED LIBERADAS
            #archivos_validar = ["expSolped03.txt"] # Dos solped para prueba 1300139393  1300139394 / se daño 


            for archivo in archivos_validar:
                EjecutarHU04(session, archivo)


            # WriteLog(
            #     mensaje=f"HU04 finalizada correctamente para archivo {archivo}.",
            #     estado="INFO",
            #     task_name=task_name,
            #     path_log=RUTAS["PathLog"],
            # )

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
                # Cambiar por funcion de descarga de OC
            dataOC = leer_solpeds_desde_archivo(ruta)
            print(type(dataOC))
                         
            #for OrdenCompra, info in dataOC.items():
                #print(f"OrdenCompra {OrdenCompra} tiene {info['items']} items")
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

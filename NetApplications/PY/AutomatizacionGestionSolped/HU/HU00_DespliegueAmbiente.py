# ================================
# GestionSOLPED – HU00: DespliegueAmbiente
# Autor: Paula Sierra - NetApplications
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 11/02/2026
# Propiedad de Colsubsidio
# Cambios:
# - Se agrega manejo de errores
# - Se agrega trazabilidad
# - Se validan rutas
# ================================

import os
import json
import random

from Config.InicializarConfig import initConfig, inConfig
from Funciones.FuncionesExcel import ExcelService
from Repositories.TicketInsumo import TicketInsumoRepo 


#from Config.initconfig import initConfig


def EjecutarHU00():

    """
    Prepara el entorno: valida carpetas, carga parámetros y estructura inicial.
    """
    try : 
            
        # ==========================================================
        # 1. Ruta base del proyecto (importante)
        # ==========================================================
        ruta_base = os.path.dirname(os.path.abspath(__file__))  # ruta de HU00
        ruta_base = os.path.abspath(os.path.join(ruta_base, ".."))
        # Sube un nivel para quedar en /AutomatizacionGestionSolped

        # ==========================================================
        # 2. Definir las carpetas obligatorias según estándar
        # ==========================================================
        carpetas = [
            "Audit/Logs",
            "Audit/Screenshots",
            "Temp",
            "Insumo",
            "Resultado",
            "Funciones",
            "HU",
        ]

        for carpeta in carpetas:
            ruta_completa = os.path.join(ruta_base, carpeta)

            if not os.path.exists(ruta_completa):
                os.makedirs(ruta_completa)

        # ==========================================================
        # 3. Cargar parámetros desde o BD
        # ==========================================================
        initConfig()
    
        # ==========================================================
        # 4. Cargar Ecxel con hojas que van a ser las tablas de parametros en la BD
        # ==========================================================


        # try : 
        #     idTablaticket= random.randint(1, 10)
        #     TicketInsumoRepo.crearPCTicketInsumo( "Inicio cargue de insumo " ) 0
        #     rutaParametros = os.path.join(inConfig("PathInsumo"),"Parametros SAMIR.xlsx")
        #     ExcelService.ejecutar_bulk_desde_excel(rutaParametros)
        #     TicketInsumoRepo.crearPCTicketInsumo( "Finalizo cargue de insumo " ) 100 
        # except: 
        #     TicketInsumoRepo.crearPCTicketInsumo( error  "Finalizo cargue de insumo " ) 99 

        ruta_config = os.path.join(ruta_base, "Config.json")

        if os.path.exists(ruta_config):
            with open(ruta_config, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = {}

        return config

    except Exception as e:
         print("Error")

         
    #     WriteLog(
    #         mensaje=f"Error Global en Main: {e} | {error_stack}",
    #         estado="ERROR",
    #         task_name=task_name,
    #         path_log=RUTAS["PathLogError"],
    #     )
    #     raise

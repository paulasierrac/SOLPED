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
from Funciones.FuncionesExcel import ServicioExcel
from Repositories.TicketInsumo import TicketInsumoRepo 

def EjecutarHU00():

    """
    Prepara el entorno: valida carpetas, carga parámetros y estructura inicial.
    """
    try : 
        
        #Probando GIT 
        # ==========================================================
        # 1. Ruta base del proyecto (importante)
        # ==========================================================
        rutaBase = os.path.dirname(os.path.abspath(__file__))  # ruta de HU00
        rutaBase = os.path.abspath(os.path.join(rutaBase, ".."))
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
            rutaCompleta = os.path.join(rutaBase, carpeta)

            if not os.path.exists(rutaCompleta):
                os.makedirs(rutaCompleta)

        # ==========================================================
        # 3. Cargar parámetros desde o BD
        # ==========================================================
        initConfig()
    
        # ==========================================================
        # 4. Cargar Ecxel con hojas que van a ser las tablas de parametros en la BD
        # ==========================================================


        try : 
                 
            TicketInsumoRepo.crearPCTicketInsumo( estado=0, observaciones= "Cargue de insumo")
            rutaParametros = os.path.join(inConfig("PathInsumo"),"Parametros SAMIR.xlsx")
            ServicioExcel.ejecutarBulkDesdeExcel(rutaParametros)
            TicketInsumoRepo.crearPCTicketInsumo( estado=100, observaciones= "Cargue de insumo")
        except: 
            TicketInsumoRepo.crearPCTicketInsumo( error= 99, observaciones="Carge de insumo " )

        rutaConfig = os.path.join(rutaBase, "Config.json")

        if os.path.exists(rutaConfig):
            with open(rutaConfig, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = {}

        return config

    except Exception as e:
         print("Error en despliege ")

         
    #     WriteLog(
    #         mensaje=f"Error Global en Main: {e} | {error_stack}",
    #         estado="ERROR",
    #         nombreTarea=nombreTarea,
    #         rutaRegistro=RUTAS["PathLogError"],
    #     )
    #     raise

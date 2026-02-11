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
from Funciones.ControlHU import control_hu
from Config.init_config import in_config,init_config


def EjecutarHU00():
    """
    Prepara el entorno: valida carpetas, carga parámetros y estructura inicial.
    """

    task_name = "HU00_DespliegueAmbiente"

    try:
        # WriteLog | INFO | INICIA HU00

        # ==========================================================
        # 0. Cargar parámetros iniciales
        # ==========================================================
        init_config()
        control_hu(task_name=task_name, estado=0)

        # ==========================================================
        # 1. Obtener ruta base del proyecto
        # ==========================================================
        ruta_base = os.path.dirname(os.path.abspath(__file__))
        ruta_base = os.path.abspath(os.path.join(ruta_base, ".."))

        if not os.path.exists(ruta_base):
            raise Exception(f"Ruta base no existe: {ruta_base}")

        # WriteLog | DEBUG | Ruta base: ruta_base

        # ==========================================================
        # 2. Definir carpetas obligatorias
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
                # WriteLog | INFO | Carpeta creada: ruta_completa
            else:
                # WriteLog | DEBUG | Carpeta existente: ruta_completa
                pass

        # ==========================================================
        # 3. Cargar archivo config.json si existe
        # ==========================================================
        ruta_config = os.path.join(ruta_base, "config.json")

        if os.path.exists(ruta_config):
            try:
                with open(ruta_config, "r", encoding="utf-8") as f:
                    config = json.load(f)
                # WriteLog | INFO | config.json cargado correctamente
            except Exception as e:
                # WriteLog | ERROR | Error leyendo config.json
                print(f"ERROR | Error leyendo config.json: {e}")
                config = {}
        else:
            # WriteLog | WARN | No existe config.json
            config = {}

        rutaParametros = os.path.join(in_config("PathInsumo"),"Parametros SAMIR.xlsx")
        ExcelService.ejecutar_bulk_desde_excel(rutaParametros)


    

        ruta_config = os.path.join(ruta_base, "Config.json")

        if os.path.exists(ruta_config):
            with open(ruta_config, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = {}

        # WriteLog | INFO | FINALIZA HU00
        control_hu(task_name=task_name, estado=100)
        return config

    except Exception as e:
        # WriteLog | ERROR | Error grave en HU00
        control_hu(task_name=task_name, estado=99)
        print(f"ERROR | Error en HU00_DespliegueAmbiente: {e}")
        return {}

# ================================
# NombreDeLaIniciativa – HU00: DespliegueAmbiente
# Autor: TuNombre, Empresa, Rol
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión inicial
# ================================

import json
import os
import win32com.client


def EjecutarHU00():
    # # Leer parámetros desde archivo (equivalente BD tabla Parámetros)
    # with open("config.json", "r", encoding="utf-8") as f:
    #     config = json.load(f)

    # Validar carpetas obligatorias
    for carpeta in ["Audit/Logs", "Audit/Screenshots", "Temp", "Insumo", "Resultado", "Funciones", "HU"]:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)


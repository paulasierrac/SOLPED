# ================================
# GestionSOLPED – HU00: DespliegueAmbiente
# Autor: Paula Sierra - NetApplications
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste ruta base dinámica + estándar Colsubsidio
# ================================

import os
import json


def EjecutarHU00():
    """
    Prepara el entorno: valida carpetas, carga parámetros y estructura inicial.
    """

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
    # 3. (Opcional) Cargar parámetros desde config.json o BD
    # ==========================================================
    ruta_config = os.path.join(ruta_base, "config.json")

    if os.path.exists(ruta_config):
        with open(ruta_config, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {}

    return config

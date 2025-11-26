# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión inicial
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import time
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
    AbrirTransaccion,
    ColsultarSolped,
    procesarTablaME5A,
    ObtenerItemTextME53N,
    ObtenerItemsME53N,
    TablaItemsDataFrame,
)
from Config.settings import RUTAS
from Funciones.ValidacionME53N import ValidacionME53N


def EjecutarHU03(session):

    try:
        WriteLog(
            mensaje="Inicia HU03",
            estado="INFO",
            task_name="HU03_ValidacionME53N",
            path_log=RUTAS["PathLog"],
        )

        dfsolpeds = procesarTablaME5A("expSolped03.txt")
        solped_unicos = dfsolpeds["PurchReq"].unique().tolist()
        solped_unicos.pop(0)
        print(solped_unicos)
        AbrirTransaccion(session, "ME53N")

        for solped in solped_unicos:
            ColsultarSolped(session, solped)
            dtItems = ObtenerItemsME53N(session, solped)

            print(dtItems)

            num_filas = dtItems.shape[0]
            print(f"Solped {solped} tiene {num_filas} filas")

            lista_dicts = dtItems.to_dict(orient="records")

            for i, fila in enumerate(lista_dicts):
                # fila es un diccionario con los datos de cada ítem
                texto = ObtenerItemTextME53N(session, solped)
                print(texto)

            # texto = ObtenerItemTextME53N(session, solped)
            # estado = "Con Aplica" if texto.strip() else "Sin Texto"
            # dfsolpeds.loc[dfsolpeds["PurchReq"] == solped, "Estado"] = estado

        return True

    except Exception as e:
        WriteLog(
            mensaje=f"Error en HU03_BuscarSolpedME53N: {e}",
            estado="ERROR",
            task_name="HU03_ValidacionME53N",
            path_log=RUTAS["PathLogError"],
        )
        return False

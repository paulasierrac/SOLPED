# ============================================
# HU05: Descarga de Orden de Compra (OC) en ME9F
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Descarga la OC generada desde la transacción ME9F.
# Ultima modificacion: 08/12/2023
# Propiedad de Colsubsidio
# Cambios: Estructura y logs.
# ============================================

from datetime import datetime
import pandas as pd
import pyperclip

from requests import session
import time
import traceback

from Config.settings import RUTAS, DB_CONFIG
from Config.InicializarConfig import inConfig

from Funciones.GuiShellFunciones import ProcesarTablaMejorada
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion
from Funciones.ValidacionME21N import EsperarSAPListo


from Funciones.FuncionesExcel import ServicioExcel


def EjecutarHU05(session, ordenesDeCompra: list):
    """
    Ejecuta la Historia de Usuario 05: Descarga de OC desde ME9F.
    """
    taskName = "HU05_DescargaOC"

    try:

        if not session:
            raise ValueError("Sesion SAP no valida.")
        
        AbrirTransaccion(session, "ME2L")
        EsperarSAPListo(session)   
        session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV" # Alcance de la lista
        session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press() # Presionar Enter
        # Alistar Texto para pegar desde el portapapeles, estándar de Windows \r\n (Carriage Return + Line Feed).
        textoParaCopiar = '\r\n'.join(ordenesDeCompra)
        pyperclip.copy(textoParaCopiar)
        EsperarSAPListo(session)
        #Boton Pegar desde el portapapeles
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        EsperarSAPListo(session) 
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        # Presionar el botón de ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(0.5)
        session.findById("wnd[0]/tbar[1]/btn[45]").press()  # Botón de lista de opciones / Fichero local crtl + shift + F9
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Botón de exportar
        # Seleccionar la línea y "Message Output"

        # === Fecha ===
        ahora = datetime.now()
        #fecha_hora = ahora.strftime("%d/%m/%Y %H:%M:%S")
        fechaArchivo = ahora.strftime("%Y%m%d_%H%M%S")
        #Guardar el archivo txt en la ruta especificada
        rutaGuardar = rf"{inConfig("PathTemp")}"
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = rutaGuardar
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = rf"LiberadasOC_{fechaArchivo}.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar



        archivo = rf"LiberadasOC_{fechaArchivo}.txt"

        dfOcliberadas = ProcesarTablaMejorada(archivo)
        #print (dfOcliberadas)
        dfOcliberadas.columns = [col.strip() for col in dfOcliberadas.columns]

        print(dfOcliberadas)

        # Definir las columnas deseadas
        columnasInteres = ["Doc.compr.", "EstadLib"]

        # Crear el nuevo DataFrame validando que las columnas existan
        if all(col in dfOcliberadas.columns for col in columnasInteres):
            dfFiltrado = dfOcliberadas[columnasInteres].copy()
            print("Nuevo DataFrame creado exitosamente.")
        else:
            # Caso alternativo: Si las columnas tienen nombres ligeramente distintos
            print(f"Columnas encontradas en el archivo: {list(dfOcliberadas.columns)}")
            # Intento de búsqueda por coincidencia parcial si falla la exacta
            colDoc = next((c for c in dfOcliberadas.columns if "Doc.compr" in c), None)
            colEst = next((c for c in dfOcliberadas.columns if "EstadLib" in c), None)
            
            if colDoc and colEst:
                dfFiltrado = dfOcliberadas[[colDoc, colEst]].copy()
                dfFiltrado.columns = ["Doc.compr.", "EstadLib"] # Renombrar para estandarizar
        
        print(dfFiltrado)
        # Guardar el DataFrame filtrado en un archivo Excel
        dfFiltrado.to_excel(rf"{inConfig("PathTemp")}\OC_Liberadas.xlsx", index=False)
        #Sube el Excel a la base de datos
        ServicioExcel.ejecutarBulkDesdeExcel(rf"{inConfig("PathTemp")}\OC_Liberadas.xlsx")

        WriteLog(
            mensaje=f"Procesamiento en ME9F completado para la OC: {ordenesDeCompra}",
            estado="INFO",
            taskName=taskName,
            pathLog=RUTAS["PathLog"],
        )

    except Exception as e:
        errorText = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU05: {e} | {errorText}",
            estado="ERROR",
            taskName=taskName,
            pathLog=RUTAS["PathLogError"],
        )
        raise

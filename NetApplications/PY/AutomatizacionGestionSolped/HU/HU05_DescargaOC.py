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
from Config.InitConfig import inConfig

from Funciones.GuiShellFunciones import ProcesarTablaMejorada
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import AbrirTransaccion
from Funciones.ValidacionME21N import EsperarSAPListo


from Funciones.FuncionesExcel import ExcelService


def EjecutarHU05(session, ordenes_de_compra: list):
    """
    Ejecuta la Historia de Usuario 05: Descarga de OC desde ME9F.
    """
    task_name = "HU05_DescargaOC"

    try:

        if not session:
            raise ValueError("Sesion SAP no valida.")
        
        AbrirTransaccion(session, "ME2L")
        EsperarSAPListo(session)   
        session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV" # Alcance de la lista
        session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press() # Presionar Enter
        # Alistar Texto para pegar desde el portapapeles, estándar de Windows \r\n (Carriage Return + Line Feed).
        texto_para_copiar = '\r\n'.join(ordenes_de_compra)
        pyperclip.copy(texto_para_copiar)
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
        fecha_archivo = ahora.strftime("%Y%m%d_%H%M%S")
        #Guardar el archivo txt en la ruta especificada
        ruta_guardar = rf"{inConfig("PathTemp")}"
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta_guardar
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = rf"LiberadasOC_{fecha_archivo}.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar



        archivo = rf"LiberadasOC_{fecha_archivo}.txt"

        df_Ocliberadas = ProcesarTablaMejorada(archivo)
        #print (df_Ocliberadas)
        df_Ocliberadas.columns = [col.strip() for col in df_Ocliberadas.columns]

        print(df_Ocliberadas)

        # Definir las columnas deseadas
        columnas_interes = ["Doc.compr.", "EstadLib"]

        # Crear el nuevo DataFrame validando que las columnas existan
        if all(col in df_Ocliberadas.columns for col in columnas_interes):
            df_filtrado = df_Ocliberadas[columnas_interes].copy()
            print("Nuevo DataFrame creado exitosamente.")
        else:
            # Caso alternativo: Si las columnas tienen nombres ligeramente distintos
            print(f"Columnas encontradas en el archivo: {list(df_Ocliberadas.columns)}")
            # Intento de búsqueda por coincidencia parcial si falla la exacta
            col_doc = next((c for c in df_Ocliberadas.columns if "Doc.compr" in c), None)
            col_est = next((c for c in df_Ocliberadas.columns if "EstadLib" in c), None)
            
            if col_doc and col_est:
                df_filtrado = df_Ocliberadas[[col_doc, col_est]].copy()
                df_filtrado.columns = ["Doc.compr.", "EstadLib"] # Renombrar para estandarizar
        
        print(df_filtrado)
        # Guardar el DataFrame filtrado en un archivo Excel
        df_filtrado.to_excel(rf"{inConfig("PathTemp")}\OC_Liberadas.xlsx", index=False)
        #Sube el Excel a la base de datos
        ExcelService.ejecutar_bulk_desde_excel(rf"{inConfig("PathTemp")}\OC_Liberadas.xlsx")

        WriteLog(
            mensaje=f"Procesamiento en ME9F completado para la OC: {ordenes_de_compra}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL en HU05: {e} | {error_text}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        raise

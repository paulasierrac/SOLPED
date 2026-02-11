# ============================================
# Función Local: SAPFuncionesME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Archivo Base funciones necesarias de SAP en la transaccion ME53N
# Ultima modificacion: 02/02/2026
# Propiedad de Colsubsidio
# Cambios:
# ============================================
import traceback
import win32com.client
import time
import os
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import ObtenerTextoDelPortapapeles
from Config.settings import RUTAS
import pandas as pd
import datetime
import re
import win32clipboard
import pyautogui
import chardet
from datetime import datetime
from typing import Dict, List, Tuple
import smtplib
import os
from Funciones.EmailSender import EmailSender
from typing import List, Union
import sys
from openpyxl import load_workbook
from Funciones.ValidacionME53N import (
    DeterminarEstadoFinal,
    ExtraerDatosTexto,
    GenerarObservaciones,
    GenerarReporteValidacion,
    ProcesarYValidarItem,
    extraerDatosReporte,
    AppendHipervinculoObservaciones,
    obtenerFilaExpSolped,
    LimpiarNumeroRobusto,
    ObtenerValorDesdeFila,
)


def ObtenerItemTextME53N(session, numero_solped, numero_item):
    """session: objeto de SAP GUI
    numero_solped: numero de SOLPED
    numero_item: numero del item actual
    Realiza la extraccion del texto del editor SAP"""

    try:
        WriteLog(
            mensaje=f"ObtenerItemTextME53N {numero_solped} Item {numero_item}",
            estado="INFO",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:
            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemTextME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # ---------------- Capturar Texto----------------
        # 1) Obtener el objeto del editor
        editor = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        # 2) Asegurar que el editor tiene el foco
        editor.SetFocus()
        time.sleep(0.5)

        # 3) Seleccionar TODO el texto
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)

        # 4) Copiar al portapapeles
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)

        # 5) Obtener texto del portapapeles con codificacion correcta
        texto_completo = ObtenerTextoDelPortapapeles()

        # 6) Limpiar caracteres problematicos si los hay
        texto_limpio = texto_completo.encode("utf-8", errors="replace").decode("utf-8")

        identificador = f"\n=====Solped: {numero_solped} Item: {numero_item} Registro: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n"

        # 7. Guardar texto en archivo de log
        # path = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\texto_ITEMsap.txt"
        path = rf"{RUTAS["PathInsumos"]}\texto_ITEMsap.txt"
        with open(path, "a", encoding="utf-8") as f:
            f.write(identificador)
            f.write(texto_limpio + "\n")
            f.write("-" * 80 + "\n")

        # 8. Navegar al siguiente item
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/"
            "btn%#AUTOTEXT002"
        ).press()
        time.sleep(0.5)

        return texto_limpio

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemTextME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLogError"],
        )
        return ""


def ProcesarTablaME5A(name, dias=None):
    """name: nombre del txt a utilizar
    return data frame
    Procesa txt estructura ME5A y devuelve un df con manejo de columnas dinamico.
    dias: int|None -> número de días a mantener (si None, no aplica filtro por fecha)"""

    try:
        WriteLog(
            mensaje=f"Procesar archivo nombre {name}",
            estado="INFO",
            task_name="ProcesarTablaME5A",
            path_log=RUTAS["PathLog"],
        )

        # path = f".\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{RUTAS["PathInsumos"]}\{name}"

        # INTENTAR LEER CON DIFERENTES CODIFICACIONES
        lineas = []
        codificaciones = ["latin-1", "cp1252", "iso-8859-1", "utf-8"]

        for codificacion in codificaciones:
            try:
                with open(path, "r", encoding=codificacion) as f:
                    lineas = f.readlines()
                print(f"EXITO: Archivo leido con codificacion {codificacion}")
                break
            except UnicodeDecodeError as e:
                print(f"ERROR con {codificacion}: {e}")
                continue
            except Exception as e:
                print(f"ERROR con {codificacion}: {e}")
                continue

        if not lineas:
            print("ERROR: No se pudo leer el archivo con ninguna codificacion")
            return pd.DataFrame()

        # Filtrar solo lineas de datos
        filas = [l for l in lineas if l.startswith("|") and not l.startswith("|---")]

        # DETECTAR ESTRUCTURA DE COLUMNAS DINAMICAMENTE
        if not filas:
            print("No se encontraron filas de datos en el archivo")
            return pd.DataFrame()

        # Analizar la primera fila para determinar estructura
        primera_fila = filas[0].strip().split("|")[1:-1]  # Quitar | inicial y final
        primera_fila = [p.strip() for p in primera_fila]

        num_columnas = len(primera_fila)
        print(f"Estructura detectada: {num_columnas} columnas")
        print(f"   Encabezados: {primera_fila}")

        # DEFINIR COLUMNAS BASE SEGUN ESTRUCTURA
        if num_columnas == 14:
            # Estructura original (sin Estado ni Observaciones)
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
            ]
            columnas_extra = ["Estado", "Observaciones"]

        elif num_columnas == 15:
            # Verificar si la columna 15 es "Estado" o "Observaciones"
            ultima_columna = primera_fila[-1].lower()
            if "estado" in ultima_columna:
                # Estructura con Estado pero sin Observaciones
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Estado",
                ]
                columnas_extra = ["Observaciones"]
            else:
                # Estructura con Observaciones pero sin Estado
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Observaciones",
                ]
                columnas_extra = ["Estado"]

        elif num_columnas == 16:
            # Estructura completa con Estado y Observaciones
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
                "Estado",
                "Observaciones",
            ]
            columnas_extra = []
        else:
            print(f"ERROR: Estructura no soportada: {num_columnas} columnas")
            return pd.DataFrame()

        # PROCESAR TODAS LAS FILAS
        filas_proc = []
        for i, fila in enumerate(filas):
            partes = fila.strip().split("|")[1:-1]
            partes = [p.strip() for p in partes]

            # Validar que tenga el numero correcto de columnas
            if len(partes) == num_columnas:
                filas_proc.append(partes)
            elif len(partes) == num_columnas + 1 and partes[-1] == "":
                # Caso: columna extra vacia al final
                filas_proc.append(partes[:num_columnas])
                if i < 3:  # Solo log primeras filas
                    print(f"   ADVERTENCIA Fila {i+1}: Columna extra vacia removida")
            else:
                print(
                    f"   ERROR Fila {i+1} ignorada: {len(partes)} columnas vs {num_columnas} esperadas"
                )
                if i == 0:  # Solo mostrar detalle para primera fila
                    print(f"      Contenido: {partes}")
                continue

        # CREAR DATAFRAME
        df = pd.DataFrame(filas_proc, columns=columnas_base)

        # AGREGAR COLUMNAS FALTANTES
        for col_extra in columnas_extra:
            if col_extra not in df.columns:
                df[col_extra] = ""
                print(f"EXITO: Columna '{col_extra}' agregada al DataFrame")

        # FILTRAR: Si la primera fila es encabezado, eliminarla
        primera_fila_es_encabezado = any(
            col in df.iloc[0].values if not df.empty else False
            for col in [
                "Purch.Req.",
                "Item",
                "Req.Date",
                "Short Text",
                "PurchReq",
                "Estado",
                "Observaciones",
            ]
        )

        if not df.empty and primera_fila_es_encabezado:
            df = df.iloc[1:].reset_index(drop=True)
            print("EXITO: Fila de encabezado removida")

        print(f"EXITO: Archivo procesado: {len(df)} filas de datos")
        print(f"   - Columnas: {list(df.columns)}")

        if not df.empty:
            print(f"   - SOLPEDs: {df['PurchReq'].nunique()}")
            if "Estado" in df.columns:
                print(f"   - Estados unicos: {df['Estado'].value_counts().to_dict()}")

        # Normalizar formato fecha
        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        if dias is not None:
            hoy = pd.Timestamp.today().normalize()
            limite = hoy - pd.Timedelta(days=int(dias))
            filas_antes = len(df)
            df = df[df["ReqDate_fmt"] >= limite].reset_index(drop=True)
            filas_despues = len(df)
            print(
                f"EXITO: Filtrado por ReqDate últimos {dias} días -> {filas_despues}/{filas_antes}"
            )
        else:
            print("INFO: No se aplicó filtro por ReqDate (dias=None)")

        # opcional: eliminar columna auxiliar
        df.drop(columns=["ReqDate_fmt"], inplace=True)

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ProcesarTablaME5A: {e}",
            estado="ERROR",
            task_name="ProcesarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR en ProcesarTablaME5A: {e}")
        traceback.print_exc()
        return pd.DataFrame()


def TablaItemsDataFrame(name) -> pd.DataFrame:
    """name: nombre del archivo a consultar
    Convierte tabla de items a df con deteccion automatica de codificacion"""

    try:
        WriteLog(
            mensaje=f"Nombre de archivo {name}",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )

        # path = rf"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\TablasME53N\{name}"
        path = rf"{RUTAS["PathInsumos"]}\TablasME53N\{name}"

        # ========== DETECCION DE CODIFICACION ==========
        encoding = DetectarCodificacion(path)

        # 1. Leer archivo con la codificacion correcta
        try:
            with open(path, "r", encoding=encoding) as f:
                texto = f.read()
        except Exception as e:
            # Si falla, intentar con otras codificaciones comunes
            print(f"Error con {encoding}, intentando alternativas...")
            for enc in ["latin-1", "cp1252", "iso-8859-1", "utf-8"]:
                try:
                    with open(path, "r", encoding=enc) as f:
                        texto = f.read()
                    print(f"EXITO con {enc}")
                    encoding = enc
                    break
                except:
                    continue

        # 2. Separar por lineas
        lineas = texto.splitlines()

        # 3. Filtrar lineas que forman parte de la tabla
        tabla = [
            l
            for l in lineas
            if l.strip().startswith("|") and l.strip().endswith("|") and "---" not in l
        ]

        if not tabla:
            raise ValueError("No se encontro ninguna tabla SAP dentro del archivo.")

        # 4. Eliminar lineas de guiones largos (separadores)
        tabla = [
            l for l in tabla if not re.match(r"^-{5,}", l.replace("|", "").strip())
        ]

        # 5. Extraer encabezado
        encabezado_raw = tabla[0]
        columnas = [c.strip() for c in encabezado_raw.split("|")[1:-1]]

        # ========== SOLUCIONAR COLUMNAS DUPLICADAS ==========
        columnas_unicas = []
        contador = {}
        for col in columnas:
            if col in contador:
                contador[col] += 1
                columnas_unicas.append(f"{col}_{contador[col]}")
            else:
                contador[col] = 0
                columnas_unicas.append(col)

        # 6. Procesar filas de datos
        filas = []
        for fila in tabla[1:]:
            partes = [c.strip() for c in fila.split("|")[1:-1]]
            if len(partes) == len(columnas_unicas):  # validar integridad
                filas.append(partes)

        # 7. Convertir a DataFrame
        df = pd.DataFrame(filas, columns=columnas_unicas)

        WriteLog(
            mensaje=f"DataFrame conversion correcta. Codificacion: {encoding}",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: DataFrame conversion correcta")
        print(f"  - Filas: {df.shape[0]}")
        print(f"  - Columnas: {df.shape[1]}")
        print(f"  - Codificacion: {encoding}")

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en TablaItemsDataFrame: {e}",
            estado="ERROR",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR: {e}")
        return pd.DataFrame()


def DetectarCodificacion(path):
    """Detecta automaticamente la codificacion del archivo"""
    try:
        with open(path, "rb") as f:
            rawdata = f.read()

        resultado = chardet.detect(rawdata)
        encoding = resultado["encoding"]
        confidence = resultado["confidence"]

        print(f"Codificacion detectada: {encoding} (confianza: {confidence*100:.1f}%)")
        return encoding
    except Exception as e:
        print(f"Error detectando codificacion: {e}")
        return "utf-8"


def ObtenerItemsME53N(session, numero_solped):
    """session: objeto de SAP GUI
    numero_solped: numero de solicitud
    Obtiene los items de SOLPED y los pasa a un df"""

    try:
        WriteLog(
            mensaje=f"Solped {numero_solped} a obtener items",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:
            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemsME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # ========== EXPORTAR TABLA ==========
        grid = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        )
        grid.setFocus()
        time.sleep(0.5)

        # 1. Abrir menu contexto "Exportar"
        grid.pressToolbarContextButton("&MB_EXPORT")
        time.sleep(0.5)

        # 2. Seleccionar "Exportar → Hoja de calculo (PC)"
        grid.selectContextMenuItem("&PC")
        time.sleep(0.3)

        # 3. Confirmar ventana de exportar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.3)

        # 4. Escribir ruta de guardado
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = (
            # r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\TablasME53N"
            rf"{RUTAS["PathInsumos"]}\TablasME53N"
        )
        time.sleep(0.2)

        # 5. Nombre del archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = (
            f"TablaSolped{numero_solped}.txt"
        )
        time.sleep(0.2)

        # 6. Guardar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)  # Esperar a que se guarde

        # ========== CONVERTIR A DATAFRAME ==========
        df = TablaItemsDataFrame(f"TablaSolped{numero_solped}.txt")

        if df is None or df.empty:
            raise Exception("DataFrame vacio despues de conversion")

        WriteLog(
            mensaje=f"Solped {numero_solped} convertido a DF con exito",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: Solped {numero_solped} convertido a DF con exito")

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemsME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR en ObtenerItemsME53N: {e}")
        return pd.DataFrame()


def GuardarTablaME5A(df, name):
    """Guarda el DataFrame de vuelta al TXT con formato de tabla"""
    try:
        # path = f"C:\\Users\\CGRPA009\\Documents\\SOLPED-main\\SOLPED\\NetApplications\\PY\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{RUTAS["PathInsumos"]}\{name}"

        # ASEGURAR QUE TIENE LAS COLUMNAS NECESARIAS
        columnas_requeridas = ["Estado", "Observaciones"]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
                print(f"ADVERTENCIA: Columna '{col}' agregada para guardado")

        # Calcular anchos de columna basados en contenido
        anchos = {}
        for col in df.columns:
            max_contenido = df[col].astype(str).str.len().max() if not df.empty else 0
            anchos[col] = max(len(col), max_contenido) + 2

        # Crear linea separadora
        separador = "-" * (sum(anchos.values()) + len(df.columns) + 1)

        # Crear encabezado
        encabezado_partes = [str(col).ljust(anchos[col]) for col in df.columns]
        encabezado = "|" + "|".join(encabezado_partes) + "|"

        # Crear filas
        filas_txt = []
        for _, fila in df.iterrows():
            partes = []
            for col in df.columns:
                valor = str(fila[col])
                # Alinear a la derecha numeros, izquierda texto
                if (
                    col in ["Item", "Quantity"]
                    or valor.replace(".", "").replace("-", "").isdigit()
                ):
                    texto_valor = valor.rjust(anchos[col])
                else:
                    texto_valor = valor.ljust(anchos[col])
                partes.append(texto_valor)
            fila_txt = "|" + "|".join(partes) + "|"
            filas_txt.append(fila_txt)

        # Escribir archivo
        with open(path, "w", encoding="utf-8") as f:
            f.write(separador + "\n")
            f.write(encabezado + "\n")
            f.write(separador + "\n")
            for fila in filas_txt:
                f.write(fila + "\n")

        WriteLog(
            mensaje=f"Archivo {name} actualizado con exito - {len(df)} filas",
            estado="INFO",
            task_name="GuardarTablaME5A",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: Archivo guardado: {len(df)} filas, {len(df.columns)} columnas")
        return True

    except Exception as e:
        WriteLog(
            mensaje=f"Error al guardar {name}: {e}",
            estado="ERROR",
            task_name="GuardarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR al guardar archivo: {e}")
        return False


def ParsearTablaAttachments(contenido: str) -> list:
    """
    Parsea la tabla de attachments exportada de SAP y extrae información estructurada

    Formato esperado:
    AttachmentFor1300139391
    -----------------------------------------------------------------------------------
    |Icon|Title                                             |Creator Name  |Created On|
    -----------------------------------------------------------------------------------
    |    |COTIZACIONPEÑALISAQUIMICOSPARAPISCINASJUNIO2024_20|OSCAR VILLABON|09.12.2025|
    |    |COTIZACIONPEÑALISAQUIMICOSPARAPISCINASJUNIO2024_20|OSCAR VILLABON|09.12.2025|
    -----------------------------------------------------------------------------------

    Args:
        contenido: Texto de la tabla exportada

    Returns:
        list: Lista de diccionarios con {title, creator, date}
    """
    attachments = []

    try:
        lineas = contenido.strip().split("\n")

        # Buscar línea de encabezado (contiene "Title", "Creator", "Created")
        header_idx = -1
        for i, linea in enumerate(lineas):
            if "|" in linea and (
                "Title" in linea or "Título" in linea or "Titulo" in linea
            ):

                header_idx = i
                break

        if header_idx == -1:
            WriteLog(
                mensaje="No se encontró encabezado en tabla de attachments",
                estado="WARNING",
                task_name="ParsearTablaAttachments",
                path_log=RUTAS["PathLog"],
            )
            return attachments

        # Procesar filas de datos (después del encabezado y línea de guiones)
        for linea in lineas[
            header_idx + 2 :
        ]:  # +2 para saltar encabezado y línea de guiones
            # Ignorar líneas vacías o separadores
            if not linea.strip() or linea.strip().startswith("-"):
                continue

            # Debe contener pipes
            if "|" not in linea:
                continue

            # Dividir por pipes y limpiar espacios
            partes = [p.strip() for p in linea.split("|")]

            # Filtrar partes vacías del inicio y final (| al inicio y final de cada línea)
            # Formato típico: |    |COTIZACION...|OSCAR VILLABON|09.12.2025|
            # Esto da: ['', '', 'COTIZACION...', 'OSCAR VILLABON', '09.12.2025', '']
            partes_validas = [p for p in partes if p]

            # Debe tener exactamente 3 columnas de datos: Title, Creator Name, Created On
            # (Icon está vacío, así que solo contamos las que tienen datos)
            if len(partes_validas) >= 3:
                # Última estructura válida: [Title, Creator, Date]
                title = partes_validas[0]
                creator = partes_validas[1]
                date = partes_validas[2]

                # Validar que no sea una línea de encabezado repetida
                if title.lower() in ["title", "título", "icon"]:
                    continue
                if creator.lower() in ["creator", "creador", "creator name"]:
                    continue

                # Validar que tenga contenido real
                if not title or not creator or not date:
                    continue

                attachments.append({"title": title, "creator": creator, "date": date})

    except Exception as e:
        WriteLog(
            mensaje=f"Error parseando tabla de attachments: {e}",
            estado="ERROR",
            task_name="ParsearTablaAttachments",
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()

    return attachments


def ValidarAttachmentList(session, numero_solped):
    """
    Valida si la SOLPED tiene Attachment List y extrae la información

    Args:
        session: Objeto de SAP GUI
        numero_solped: Número de SOLPED a validar

    Returns:
        tuple: (tiene_attachments: bool, contenido_tabla: str, observaciones: str)
    """
    try:
        WriteLog(
            mensaje=f"Validando Attachment List para SOLPED {numero_solped}",
            estado="INFO",
            task_name="ValidarAttachmentList",
            path_log=RUTAS["PathLog"],
        )

        # 1. Presionar botón de GOS Toolbox
        try:
            session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
            time.sleep(0.5)
        except Exception as e:
            WriteLog(
                mensaje=f"Error al abrir GOS Toolbox: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )
            return False, "", "No se pudo abrir el menú de servicios GOS"

        # 2. Presionar botón VIEW_ATTA (View Attachments)
        try:
            session.findById("wnd[0]/shellcont[1]/shell").pressButton("VIEW_ATTA")
            time.sleep(1)
        except Exception as e:
            WriteLog(
                mensaje=f"No hay attachments disponibles: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )
            # Cerrar menú GOS si quedó abierto
            try:
                session.findById("wnd[0]/shellcont[1]").close()
            except:
                pass
            return False, "", "No se encontró lista de adjuntos (Attachment List vacía)"

        # 3. Verificar si se abrió la ventana "Service: Attachment list"
        try:
            # Intentar acceder al objeto de la tabla de attachments
            tabla_attachments = session.findById(
                "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell"
            )
            time.sleep(0.5)
        except Exception as e:
            WriteLog(
                mensaje=f"Ventana de Attachment List no encontrada: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )
            # Cerrar menú GOS
            try:
                session.findById("wnd[0]/shellcont[1]").close()
            except:
                pass
            return False, "", "Ventana de adjuntos no disponible"

        # 4. Exportar contenido de la tabla al portapapeles
        try:
            # Abrir menú de exportación
            tabla_attachments.pressToolbarContextButton("&MB_EXPORT")
            time.sleep(0.5)

            # Seleccionar "Hoja de cálculo" (opción PC)
            tabla_attachments.selectContextMenuItem("&PC")
            time.sleep(0.5)

            # Seleccionar formato "Spreadsheet" (radio button)
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]"
            ).select()
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]"
            ).setFocus()
            time.sleep(0.3)

            # Confirmar exportación (botón OK)
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(1)

        except Exception as e:
            WriteLog(
                mensaje=f"Error al exportar Attachment List: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )
            # Intentar cerrar ventanas abiertas
            try:
                session.findById(
                    "wnd[1]/tbar[0]/btn[12]"
                ).press()  # Cancelar ventana exportación
            except:
                pass
            try:
                session.findById("wnd[0]/shellcont[1]").close()  # Cerrar menú GOS
            except:
                pass
            return False, "", "Error al exportar lista de adjuntos"

        # 5. Obtener contenido del portapapeles
        time.sleep(0.5)
        contenido_portapapeles = ObtenerTextoDelPortapapeles()

        if not contenido_portapapeles or not contenido_portapapeles.strip():
            WriteLog(
                mensaje=f"Portapapeles vacío después de exportar attachments",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )
            # Cerrar ventanas
            try:
                session.findById(
                    "wnd[1]/tbar[0]/btn[12]"
                ).press()  # Cerrar ventana de Attachment List
            except:
                pass
            try:
                session.findById("wnd[0]/shellcont[1]").close()  # Cerrar menú GOS
            except:
                pass
            return False, "", "No se pudo copiar contenido de attachments"

        # 6. Cerrar ventana de Attachment List (botón 12 = Cerrar)
        try:
            session.findById("wnd[1]/tbar[0]/btn[12]").press()
            time.sleep(0.3)
        except Exception as e:
            WriteLog(
                mensaje=f"Advertencia al cerrar ventana de attachments: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )

        # 7. Cerrar menú GOS Toolbox (shellcont[1])
        try:
            session.findById("wnd[0]/shellcont[1]").close()
            time.sleep(0.3)
        except Exception as e:
            WriteLog(
                mensaje=f"Advertencia al cerrar menú GOS: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )

        # 8. Analizar contenido obtenido
        lineas = contenido_portapapeles.strip().split("\n")
        lineas_validas = [
            l.strip() for l in lineas if l.strip() and not l.strip().startswith("-")
        ]

        # Parsear attachments estructurados
        attachments_parseados = ParsearTablaAttachments(contenido_portapapeles)
        num_attachments = len(attachments_parseados)

        if num_attachments == 0:
            # Aún así cerrar todo antes de retornar
            return (
                False,
                contenido_portapapeles,
                "Lista de adjuntos vacía (sin archivos)",
            )

        # 9. Guardar contenido en archivo de log
        identificador = f"\n===== SOLPED: {numero_solped} - Attachment List - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n"
        path_log_attachments = rf"{RUTAS['PathInsumos']}\attachment_lists.txt"

        try:
            with open(path_log_attachments, "a", encoding="utf-8") as f:
                f.write(identificador)
                f.write(contenido_portapapeles)
                f.write("\n" + "-" * 80 + "\n")
        except Exception as e:
            WriteLog(
                mensaje=f"Advertencia: No se pudo guardar log de attachments: {e}",
                estado="WARNING",
                task_name="ValidarAttachmentList",
                path_log=RUTAS["PathLog"],
            )

        WriteLog(
            mensaje=f"SOLPED {numero_solped}: {num_attachments} attachment(s) encontrado(s)",
            estado="INFO",
            task_name="ValidarAttachmentList",
            path_log=RUTAS["PathLog"],
        )

        # Construir observación detallada con nombres de archivos
        observaciones_exito = f"✅ {num_attachments} archivo(s) adjunto(s)"

        if num_attachments <= 3:
            # Mostrar nombres si son pocos archivos
            nombres = [a["title"][:40] for a in attachments_parseados[:3]]
            observaciones_exito += f": {', '.join(nombres)}"
            if any(len(a["title"]) > 40 for a in attachments_parseados[:3]):
                observaciones_exito += "..."

        return True, contenido_portapapeles, observaciones_exito

    except Exception as e:
        error_trace = traceback.format_exc()
        WriteLog(
            mensaje=f"Error inesperado en ValidarAttachmentList: {e}\n{error_trace}",
            estado="ERROR",
            task_name="ValidarAttachmentList",
            path_log=RUTAS["PathLogError"],
        )
        # Intentar cerrar cualquier ventana abierta
        try:
            session.findById("wnd[1]/tbar[0]/btn[12]").press()
        except:
            pass
        try:
            session.findById("wnd[0]/shellcont[1]").close()
        except:
            pass
        return False, "", f"Error al validar attachments: {str(e)[:100]}"

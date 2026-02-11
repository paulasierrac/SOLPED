# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N (Versión Final v3 CORREGIDA)
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N y genera reporte consolidado
# Ultima modificacion: 09/12/2025 - Versión 3 CORREGIDA
# Propiedad de Colsubsidio
# Cambios v3 CORREGIDA:
#   - Corregido error con DataFrame (usar .empty en lugar de not df)
#   - Corregido parámetro ActualizarEstado (nuevo_estado en lugar de estado)
#   - Integrada información del texto del item en el reporte
#   - Delimitador punto y coma (;)
#   - Adjuntos desde SAP con WriteLog
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import time
import traceback
import pandas as pd
from datetime import datetime
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
    AbrirTransaccion,
    ColsultarSolped,
    ProcesarTablaME5A,
    ObtenerItemTextME53N,
    ObtenerItemsME53N,
    TablaItemsDataFrame,
    TraerSAPAlFrenteOpcion,
    ActualizarEstado,
    ActualizarEstadoYObservaciones,
    ProcesarYValidarItem,
    GuardarTablaME5A,
    NotificarRevisionManualSolped,
    ValidarAttachmentList,
    ParsearTablaAttachments,
    GenerarReporteAttachments,
)
from Config.settings import RUTAS


def GenerarReporteConsolidadoFinal(df_resultados, nombre_archivo_salida):
    """
    Genera un archivo de reporte consolidado en formato CSV delimitado por punto y coma (;)
    con todos los campos procesados.

    Args:
        df_resultados: DataFrame con todos los datos procesados
        nombre_archivo_salida: Nombre del archivo de salida (sin extensión)

    Returns:
        str: Ruta completa del archivo generado
    """
    try:
        # Definir columnas del reporte en el orden solicitado
        columnas_reporte = [
            # Datos de expSolped03.txt
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
            # Información de adjuntos
            "Adjuntos",
            "Nombre de Adjunto",
            # Datos de TablaSolped (ME53N)
            "Material_ME53N",
            "Short Text_ME53N",
            "Quantity_ME53N",
            "Un",
            "Valn Price",
            "Crcy",
            "Total Val.",
            "Deliv.Date",
            "Fix. Vend.",
            "Plant",
            "PGr_ME53N",
            "POrg",
            "Matl Group",
            # Datos extraídos del texto del item
            "Id",
            "PurchReq_Texto",
            "Item_Texto",
            "Razon Social:",
            "NIT:",
            "Correo:",
            "Concepto:",
            "Cantidad:",
            "Valor Unitario:",
            "Valor Total:",
            "Responsable:",
            "CECO:",
            # Resultados de validaciones
            "CAMPOS OBLIGATORIOS SAP ME53N:",
            "DATOS EXTRAIDOS Faltantes",
            "DATOS EXTRAIDOS DEL TEXTO",
            "DATOS EXTRAIDOS DEL TEXTO faltantes",
            "CANTIDAD",
            "VALOR_UNITARIO",
            "VALOR_TOTAL",
            "CONCEPTO",
            "VALIDACIONES",
            "Estado",
            "Observaciones",
        ]

        # Crear DataFrame para el reporte con las columnas correctas
        df_reporte = pd.DataFrame(columns=columnas_reporte)

        # Procesar cada fila de resultados
        for idx, row in df_resultados.iterrows():
            fila_reporte = {}

            # Copiar campos directamente del DataFrame original
            for col in columnas_reporte:
                if col in df_resultados.columns:
                    fila_reporte[col] = row.get(col, "")
                else:
                    fila_reporte[col] = ""

            # Agregar fila al reporte
            df_reporte = pd.concat(
                [df_reporte, pd.DataFrame([fila_reporte])], ignore_index=True
            )

        # Generar nombre de archivo con fecha actual
        fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_completo = f"reporte_{fecha_actual}.txt"
        ruta_completa = os.path.join(RUTAS["PathReportes"], nombre_completo)

        # Guardar archivo delimitado por PUNTO Y COMA (;)
        df_reporte.to_csv(
            ruta_completa,
            sep=";",  # Delimitador punto y coma
            index=False,
            encoding="utf-8-sig",
            na_rep="",
        )

        print(f"\n{'='*80}")
        WriteLog(
            mensaje=f"Generación de reporte consolidado iniciada",
            estado="INFO",
            task_name="GenerarReporteConsolidado",
            path_log=RUTAS["PathLog"],
        )

        print(f"REPORTE CONSOLIDADO GENERADO EXITOSAMENTE")
        WriteLog(
            mensaje=f"Reporte consolidado generado exitosamente: {nombre_completo}",
            estado="INFO",
            task_name="GenerarReporteConsolidado",
            path_log=RUTAS["PathLog"],
        )

        print(f"📁 Archivo: {nombre_completo}")
        print(f"📂 Ubicación: {ruta_completa}")
        print(f"📊 Total de registros: {len(df_reporte)}")
        print(f"🔹 Delimitador: punto y coma (;)")
        WriteLog(
            mensaje=f"Reporte generado con {len(df_reporte)} registros en {ruta_completa} (delimitador: ;)",
            estado="INFO",
            task_name="GenerarReporteConsolidado",
            path_log=RUTAS["PathLog"],
        )

        print(f"{'='*80}\n")

        return ruta_completa

    except Exception as e:
        print(f"Error al generar reporte consolidado: {e}")
        WriteLog(
            mensaje=f"Error al generar reporte consolidado: {e}",
            estado="ERROR",
            task_name="GenerarReporteConsolidado",
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return None


def NormalizarNombreColumna(nombre):
    """
    Normaliza los nombres de columnas para manejar variaciones
    (ej: "Total Val." vs "Total Value")
    """
    normalizaciones = {
        "Total Value": "Total Val.",
        "Val. Price": "Valn Price",
        "Value": "Val.",
    }

    for original, normalizado in normalizaciones.items():
        if original in nombre:
            return nombre.replace(original, normalizado)

    return nombre


def ObtenerValorTabla(df_tabla, columna, valor_default=""):
    """
    Obtiene un valor de la tabla manejando nombres de columnas variables
    """
    # Intentar primero el nombre exacto
    if columna in df_tabla.columns:
        return df_tabla[columna].iloc[0] if len(df_tabla) > 0 else valor_default

    # Intentar variaciones comunes
    variaciones = {
        "Total Val.": ["Total Value", "Total Val", "TotalValue"],
        "Valn Price": ["Val. Price", "Value Price", "Val Price"],
        "Short Text": ["ShortText", "Short_Text"],
    }

    if columna in variaciones:
        for variacion in variaciones[columna]:
            if variacion in df_tabla.columns:
                return (
                    df_tabla[variacion].iloc[0] if len(df_tabla) > 0 else valor_default
                )

    return valor_default


def EjecutarHU03(session, nombre_archivo):
    try:
        task_name = "HU03_ValidacionME53N"

        # === Inicio HU03 ===
        print("=" * 80)
        print("INICIO HU03 - Validación ME53N con generación de reporte consolidado v3")
        print("=" * 80)
        WriteLog(
            mensaje="Inicio HU03 - Validación ME53N con generación de reporte consolidado v3 (delimitador: ;)",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Traer SAP al frente
        TraerSAPAlFrenteOpcion()

        # Leer el archivo con las SOLPEDs a procesar
        df_solpeds = ProcesarTablaME5A(nombre_archivo)
        GuardarTablaME5A(df_solpeds, nombre_archivo)

        if df_solpeds.empty:
            print("ERROR: No se pudo cargar el archivo o esta vacio")
            WriteLog(
                mensaje="El archivo expSolped03.txt está vacío o no se pudo cargar",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            return False

        # === Validación de columnas ===
        columnas_requeridas = ["Estado", "Observaciones"]
        for columna in columnas_requeridas:
            if columna not in df_solpeds.columns:
                print(
                    f"ERROR: Columna requerida '{columna}' no encontrada en el DataFrame"
                )
                WriteLog(
                    mensaje=f"No se encontró la columna requerida: {columna}",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                return False

        # ============================================================
        # INICIALIZAR DATAFRAME PARA REPORTE CONSOLIDADO
        # ============================================================
        columnas_reporte = [
            # Datos de expSolped03.txt
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
            # Información de adjuntos
            "Adjuntos",
            "Nombre de Adjunto",
            # Datos de TablaSolped (ME53N)
            "Material_ME53N",
            "Short Text_ME53N",
            "Quantity_ME53N",
            "Un",
            "Valn Price",
            "Crcy",
            "Total Val.",
            "Deliv.Date",
            "Fix. Vend.",
            "Plant",
            "PGr_ME53N",
            "POrg",
            "Matl Group",
            # Datos extraídos del texto del item
            "Id",
            "PurchReq_Texto",
            "Item_Texto",
            "Razon Social:",
            "NIT:",
            "Correo:",
            "Concepto:",
            "Cantidad:",
            "Valor Unitario:",
            "Valor Total:",
            "Responsable:",
            "CECO:",
            # Resultados de validaciones
            "CAMPOS OBLIGATORIOS SAP ME53N:",
            "DATOS EXTRAIDOS Faltantes",
            "DATOS EXTRAIDOS DEL TEXTO",
            "DATOS EXTRAIDOS DEL TEXTO faltantes",
            "CANTIDAD",
            "VALOR_UNITARIO",
            "VALOR_TOTAL",
            "CONCEPTO",
            "VALIDACIONES",
            "Estado",
            "Observaciones",
        ]

        df_reporte_consolidado = pd.DataFrame(columns=columnas_reporte)

        # === Limpieza de SOLPEDs válidas ===
        solped_unicos = df_solpeds["PurchReq"].unique().tolist()

        # Filtrar SOLPEDs validas (excluir encabezados)
        solped_unicos_filtradas = []
        for solped in solped_unicos:
            solped_str = str(solped).strip()

            # Excluir encabezados y valores no validos
            if (
                solped_str
                and solped_str not in ["Purch.Req.", "PurchReq", "Purch.Req", ""]
                and not any(
                    header in solped_str for header in ["Purch.Req", "PurchReq"]
                )
                and solped_str.replace(".", "").isdigit()
            ):

                solped_limpia = solped_str.replace(".", "")
                if solped_limpia.isdigit():
                    solped_unicos_filtradas.append(solped_limpia)
                else:
                    solped_unicos_filtradas.append(solped_str)
            else:
                print(f"EXCLUIDO: '{solped_str}' (no es una SOLPED valida)")
                WriteLog(
                    mensaje=f"SOLPED excluida: '{solped_str}' (no es válida)",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

        solped_unicos = solped_unicos_filtradas

        if not solped_unicos:
            print("ERROR: No se encontraron SOLPEDs validas para procesar")
            WriteLog(
                mensaje="No se encontraron SOLPEDs válidas para procesar",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            return False

        print(f"Procesando {len(solped_unicos)} SOLPEDs unicas...")
        WriteLog(
            mensaje=f"Procesando {len(solped_unicos)} SOLPEDs únicas",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Informacion inicial del archivo
        print("RESUMEN INICIAL DEL ARCHIVO:")
        print(f"   - Total filas: {len(df_solpeds)}")
        print(f"   - SOLPEDs unicas validas: {len(solped_unicos)}")
        WriteLog(
            mensaje=f"Resumen inicial - Total filas: {len(df_solpeds)}, SOLPEDs únicas: {len(solped_unicos)}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Mostrar distribucion inicial de estados
        if "Estado" in df_solpeds.columns:
            estados_iniciales = df_solpeds["Estado"].value_counts()
            print(f"   - Distribucion inicial de estados:")
            for estado, count in estados_iniciales.items():
                print(f"     {estado}: {count}")
        print()

        # Abrir transaccion ME53N en SAP
        AbrirTransaccion(session, "ME53N")

        # Contadores para resumen final
        contadores = {
            "total_solpeds": len(solped_unicos),
            "procesadas_exitosamente": 0,
            "con_errores": 0,
            "sin_items": 0,
            "items_procesados": 0,
            "items_validados": 0,
            "items_sin_texto": 0,
            "items_verificar_manual": 0,
            "notificaciones_enviadas": 0,
            "notificaciones_fallidas": 0,
            "rechazadas_sin_attachments": 0,
        }

        # ========================================================
        # MODO DESARROLLO - REDIRIGIR CORREOS
        # ========================================================
        MODO_DESARROLLO = True  # Cambiar a False en producción
        EMAIL_DESARROLLO = "paula.sierra@netapplications.com.co"

        if MODO_DESARROLLO:
            print(f"\n{'='*60}")
            print(f" MODO DESARROLLO ACTIVO")
            print(f"📧 Todos los correos se enviarán a: {EMAIL_DESARROLLO}")
            print(f"{'='*60}\n")
            WriteLog(
                mensaje=f"MODO DESARROLLO: Correos redirigidos a {EMAIL_DESARROLLO}",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Almacenar SOLPEDs que requirieron revisión para reporte final
        solpeds_con_problemas = []

        # Procesar cada SOLPED
        for solped in solped_unicos:
            print(f"\n{'='*80}")
            print(f"PROCESANDO SOLPED: {solped}")
            print(f"{'='*80}")
            WriteLog(
                mensaje=f"Iniciando procesamiento de SOLPED: {solped}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            # Variables para notificación
            correos_responsables = []
            resumen_validaciones = []
            requiere_notificacion = False
            solped_rechazada_por_attachments = False

            try:
                # Consultar SOLPED en SAP
                resultado_consulta = ColsultarSolped(session, solped)

                if not resultado_consulta:
                    print(f"No se pudo consultar la SOLPED {solped}")
                    WriteLog(
                        mensaje=f"No se pudo consultar la SOLPED {solped} en SAP",
                        estado="ERROR",
                        task_name=task_name,
                        path_log=RUTAS["PathLogError"],
                    )
                    ActualizarEstado(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Error consulta",
                    )
                    contadores["con_errores"] += 1
                    continue

                # ============================================================
                # OBTENER ADJUNTOS DESDE VALIDACIÓN SAP
                # ============================================================
                print(f"\n--- Validando adjuntos en SAP para SOLPED {solped} ---")
                WriteLog(
                    mensaje=f"Iniciando validación de adjuntos en SAP para SOLPED {solped}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

                tiene_attachments, contenido_attachments, observaciones_attachments = (
                    ValidarAttachmentList(session, solped)
                )

                # Parsear attachments para información detallada
                attachments_lista = []
                if contenido_attachments:
                    attachments_lista = ParsearTablaAttachments(contenido_attachments)

                cantidad_adjuntos = 0
                nombres_adjuntos = []

                if tiene_attachments:
                    cantidad_adjuntos = len(attachments_lista)
                    nombres_adjuntos = [att["title"] for att in attachments_lista]

                    print(f"Encontrados {cantidad_adjuntos} adjuntos en SAP")
                    if cantidad_adjuntos <= 5:
                        for i, nombre in enumerate(nombres_adjuntos, 1):
                            print(f"   {i}. {nombre[:60]}")
                    else:
                        for i, nombre in enumerate(nombres_adjuntos[:3], 1):
                            print(f"   {i}. {nombre[:60]}")
                        print(f"   ... y {cantidad_adjuntos - 3} más")

                    # GUARDAR SOLO EN WRITELOG (NO EN ARCHIVO)
                    WriteLog(
                        mensaje=f"SOLPED {solped}: {cantidad_adjuntos} adjunto(s) encontrado(s) en SAP\n"
                        f"Nombres: {', '.join(nombres_adjuntos)}\n"
                        f"Contenido completo:\n{contenido_attachments}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                    print(f"SOLPED {solped} tiene attachments - Continuando validación")
                else:
                    print(f"No se encontraron adjuntos en SAP para SOLPED {solped}")
                    WriteLog(
                        mensaje=f"SOLPED {solped}: Sin adjuntos en SAP. Observación: {observaciones_attachments}",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                    # MARCAR SOLPED COMO RECHAZADA POR FALTA DE ADJUNTOS
                    print(f"\nSOLPED {solped} SERÁ RECHAZADA: Sin archivos adjuntos")
                    print(
                        f" Continuando con validaciones de items para reporte completo..."
                    )

                    contadores["rechazadas_sin_attachments"] += 1
                    solped_rechazada_por_attachments = True
                    requiere_notificacion = True

                adjuntos_info = str(cantidad_adjuntos)
                nombres_adjuntos_str = (
                    ", ".join(nombres_adjuntos) if nombres_adjuntos else ""
                )

                # === OBTENER ITEMS DE LA SOLPED ===
                dtItems = ObtenerItemsME53N(session, solped)

                # CORREGIDO: Usar .empty en lugar de not dtItems
                if dtItems is None or (
                    isinstance(dtItems, pd.DataFrame) and dtItems.empty
                ):
                    print(f"SOLPED {solped} sin items o no se pudieron obtener")
                    WriteLog(
                        mensaje=f"SOLPED {solped} sin items o no se pudieron obtener",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    ActualizarEstado(
                        df_solpeds, nombre_archivo, solped, nuevo_estado="Sin items"
                    )
                    contadores["sin_items"] += 1
                    continue

                # Obtener lista de items
                if isinstance(dtItems, pd.DataFrame):
                    items_info = dtItems["Item"].tolist()
                else:
                    items_info = dtItems

                print(f"Se encontraron {len(items_info)} items en la SOLPED {solped}")
                WriteLog(
                    mensaje=f"SOLPED {solped}: {len(items_info)} items encontrados",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

                # === OBTENER TABLA DE DATOS ME53N ===
                df_tabla_me53n = TablaItemsDataFrame(session, solped)

                if df_tabla_me53n.empty:
                    print(
                        f"No se pudo obtener la tabla de datos ME53N para SOLPED {solped}"
                    )
                    WriteLog(
                        mensaje=f"No se pudo obtener tabla ME53N para SOLPED {solped}",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                else:
                    # Normalizar nombres de columnas
                    df_tabla_me53n.columns = [
                        NormalizarNombreColumna(col) for col in df_tabla_me53n.columns
                    ]
                    print(f"Tabla ME53N obtenida con {len(df_tabla_me53n)} registros")
                    WriteLog(
                        mensaje=f"Tabla ME53N obtenida para SOLPED {solped}: {len(df_tabla_me53n)} registros",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                # Procesar cada item
                items_ok = 0
                items_con_problemas = 0
                items_sin_texto = 0

                for numero_item in items_info:
                    print(f"\n--- Procesando Item {numero_item} de SOLPED {solped} ---")
                    WriteLog(
                        mensaje=f"Procesando Item {numero_item} de SOLPED {solped}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    contadores["items_procesados"] += 1

                    # IMPORTANTE: Generar Id como concatenación directa de PurchReq + Item
                    item_str = str(numero_item).strip()
                    solped_str = str(solped).strip()
                    id_item = f"{solped_str}{item_str}"

                    print(
                        f"📌 ID generado: {id_item} (SOLPED: {solped_str} + Item: {item_str})"
                    )
                    WriteLog(
                        mensaje=f"ID generado: {id_item} (SOLPED: {solped_str} + Item: {item_str})",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                    # Obtener datos del item desde expSolped03
                    mask_item = (
                        df_solpeds["PurchReq"].astype(str).str.strip() == solped_str
                    ) & (df_solpeds["Item"].astype(str).str.strip() == item_str)

                    datos_expsolped = {}
                    if mask_item.any():
                        fila_item = df_solpeds[mask_item].iloc[0]
                        datos_expsolped = {
                            "PurchReq": solped_str,
                            "Item": item_str,
                            "ReqDate": str(fila_item.get("ReqDate", "")).strip(),
                            "Material": str(fila_item.get("Material", "")).strip(),
                            "Created": str(fila_item.get("Created", "")).strip(),
                            "ShortText": str(fila_item.get("ShortText", "")).strip(),
                            "PO": str(fila_item.get("PO", "")).strip(),
                            "Quantity": str(fila_item.get("Quantity", "")).strip(),
                            "Plnt": str(fila_item.get("Plnt", "")).strip(),
                            "PGr": str(fila_item.get("PGr", "")).strip(),
                            "Blank1": str(fila_item.get("Blank1", "")).strip(),
                            "D": str(fila_item.get("D", "")).strip(),
                            "Requisnr": str(fila_item.get("Requisnr", "")).strip(),
                            "ProcState": str(fila_item.get("ProcState", "")).strip(),
                        }

                    # Agregar información de adjuntos (desde SAP)
                    datos_expsolped["Adjuntos"] = adjuntos_info
                    datos_expsolped["Nombre de Adjunto"] = nombres_adjuntos_str

                    # Obtener datos de la tabla ME53N para este item
                    datos_me53n = {}
                    if not df_tabla_me53n.empty:
                        mask_tabla = (
                            df_tabla_me53n["Item"].astype(str).str.strip() == item_str
                        )
                        if mask_tabla.any():
                            fila_tabla = df_tabla_me53n[mask_tabla].iloc[0]

                            datos_me53n = {
                                "Material_ME53N": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Material", ""
                                    )
                                ).strip(),
                                "Short Text_ME53N": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Short Text", ""
                                    )
                                ).strip(),
                                "Quantity_ME53N": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Quantity", ""
                                    )
                                ).strip(),
                                "Un": str(
                                    ObtenerValorTabla(fila_tabla.to_frame().T, "Un", "")
                                ).strip(),
                                "Valn Price": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Valn Price", ""
                                    )
                                ).strip(),
                                "Crcy": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Crcy", ""
                                    )
                                ).strip(),
                                "Total Val.": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Total Val.", ""
                                    )
                                ).strip(),
                                "Deliv.Date": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Deliv.Date", ""
                                    )
                                ).strip(),
                                "Fix. Vend.": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Fix. Vend.", ""
                                    )
                                ).strip(),
                                "Plant": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Plant", ""
                                    )
                                ).strip(),
                                "PGr_ME53N": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "PGr", ""
                                    )
                                ).strip(),
                                "POrg": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "POrg", ""
                                    )
                                ).strip(),
                                "Matl Group": str(
                                    ObtenerValorTabla(
                                        fila_tabla.to_frame().T, "Matl Group", ""
                                    )
                                ).strip(),
                            }

                    # ============================================================
                    # OBTENER Y VALIDAR TEXTO DEL ITEM
                    # ============================================================
                    texto = ObtenerItemTextME53N(session, solped, numero_item)

                    datos_validacion = {}

                    if texto and texto.strip():
                        items_sin_texto += 0

                        # VALIDACION COMPLETA DEL TEXTO
                        resultado = ProcesarYValidarItem(
                            session,
                            solped,
                            numero_item,
                            texto,
                            dtItems,
                            tiene_attachments,
                            observaciones_attachments,
                            attachments_lista,
                        )

                        # Extraer componentes del resultado (tupla de 5 elementos)
                        datos_texto = resultado[0] if len(resultado) > 0 else {}
                        validaciones = resultado[1] if len(resultado) > 1 else {}
                        reporte = resultado[2] if len(resultado) > 2 else ""
                        estado_final = resultado[3] if len(resultado) > 3 else ""
                        observaciones = resultado[4] if len(resultado) > 4 else ""

                        # Construir datos de validación
                        datos_validacion = {
                            "Id": id_item,
                            "PurchReq_Texto": solped_str,
                            "Item_Texto": item_str,
                            "Razon Social:": str(
                                datos_texto.get("razon_social", "")
                            ).strip(),
                            "NIT:": str(datos_texto.get("nit", "")).strip(),
                            "Correo:": str(datos_texto.get("correo", "")).strip(),
                            "Concepto:": str(
                                datos_texto.get("concepto_compra", "")
                            ).strip(),
                            "Cantidad:": str(datos_texto.get("cantidad", "")).strip(),
                            "Valor Unitario:": str(
                                datos_texto.get("valor_unitario", "")
                            ).strip(),
                            "Valor Total:": str(
                                datos_texto.get("valor_total", "")
                            ).strip(),
                            "Responsable:": str(
                                datos_texto.get("responsable", "")
                            ).strip(),
                            "CECO:": str(datos_texto.get("ceco", "")).strip(),
                            "CAMPOS OBLIGATORIOS SAP ME53N:": str(
                                validaciones.get("campos_obligatorios", "")
                            ).strip(),
                            "DATOS EXTRAIDOS Faltantes": str(
                                validaciones.get("datos_extraidos_faltantes", "")
                            ).strip(),
                            "DATOS EXTRAIDOS DEL TEXTO": str(
                                validaciones.get("datos_extraidos_texto", "")
                            ).strip(),
                            "DATOS EXTRAIDOS DEL TEXTO faltantes": str(
                                validaciones.get("datos_texto_faltantes", "")
                            ).strip(),
                            "CANTIDAD": str(validaciones.get("cantidad", "")).strip(),
                            "VALOR_UNITARIO": str(
                                validaciones.get("valor_unitario", "")
                            ).strip(),
                            "VALOR_TOTAL": str(
                                validaciones.get("valor_total", "")
                            ).strip(),
                            "CONCEPTO": str(validaciones.get("concepto", "")).strip(),
                            "VALIDACIONES": str(
                                validaciones.get("resumen_validaciones", "")
                            ).strip(),
                            "Estado": str(estado_final).strip(),
                            "Observaciones": str(observaciones).strip(),
                        }

                        # Verificar si el item requiere revisión
                        if "validado" in estado_final.lower():
                            items_ok += 1
                            contadores["items_validados"] += 1
                            print(f"Item {numero_item} validado correctamente")
                            WriteLog(
                                mensaje=f"Item {numero_item} de SOLPED {solped} validado correctamente",
                                estado="INFO",
                                task_name=task_name,
                                path_log=RUTAS["PathLog"],
                            )
                        else:
                            items_con_problemas += 1
                            contadores["items_verificar_manual"] += 1
                            requiere_notificacion = True
                            print(f"Item {numero_item} requiere revisión manual")
                            WriteLog(
                                mensaje=f"Item {numero_item} de SOLPED {solped} requiere revisión manual",
                                estado="WARNING",
                                task_name=task_name,
                                path_log=RUTAS["PathLog"],
                            )

                            # Extraer correos responsables
                            if datos_texto.get("responsable"):
                                correos_resp = str(
                                    datos_texto.get("responsable", "")
                                ).split(",")
                                correos_responsables.extend(
                                    [c.strip() for c in correos_resp if c.strip()]
                                )
                    else:
                        # Item sin texto
                        items_sin_texto += 1
                        contadores["items_sin_texto"] += 1
                        print(f"Item {numero_item} sin texto")
                        WriteLog(
                            mensaje=f"Item {numero_item} de SOLPED {solped} sin texto",
                            estado="WARNING",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )

                        # Datos mínimos
                        datos_validacion = {
                            "Id": id_item,
                            "PurchReq_Texto": solped_str,
                            "Item_Texto": item_str,
                            "Estado": "Sin texto",
                            "Observaciones": "Item sin texto para validar",
                        }

                    # Consolidar todos los datos del item
                    datos_completos_item = {}
                    datos_completos_item.update(datos_expsolped)
                    datos_completos_item.update(datos_me53n)
                    datos_completos_item.update(datos_validacion)

                    # Agregar al DataFrame de reporte consolidado
                    df_reporte_consolidado = pd.concat(
                        [df_reporte_consolidado, pd.DataFrame([datos_completos_item])],
                        ignore_index=True,
                    )

                    print(f"Item {numero_item} agregado al reporte con Id: {id_item}")
                    WriteLog(
                        mensaje=f"Item {numero_item} agregado al reporte consolidado con Id: {id_item}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                # === ACTUALIZAR ESTADO FINAL DE LA SOLPED ===
                if solped_rechazada_por_attachments:
                    estado_final = "Rechazada - Sin Attachments"
                    obs_final = f"RECHAZADA por falta de adjuntos | Items: {items_ok} validados, {items_con_problemas} requieren revisión, {items_sin_texto} sin texto"
                elif items_con_problemas > 0:
                    estado_final = f"{items_ok}/{len(items_info)} validados"
                    obs_final = f"{items_con_problemas} items requieren revisión manual"
                else:
                    estado_final = f"{items_ok}/{len(items_info)} validados"
                    obs_final = "Todos los items validados correctamente"

                ActualizarEstadoYObservaciones(
                    df_solpeds,
                    nombre_archivo,
                    solped,
                    nuevo_estado=estado_final,
                    observaciones=obs_final,
                )

                contadores["procesadas_exitosamente"] += 1
                print(f"\n✅ SOLPED {solped} procesada: {estado_final}")
                WriteLog(
                    mensaje=f"SOLPED {solped} procesada exitosamente: {estado_final}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

            except Exception as e:
                print(f"Error procesando SOLPED {solped}: {e}")
                WriteLog(
                    mensaje=f"Error procesando SOLPED {solped}: {e}\n{traceback.format_exc()}",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                traceback.print_exc()
                ActualizarEstado(
                    df_solpeds, nombre_archivo, solped, nuevo_estado="Error"
                )
                contadores["con_errores"] += 1

        # === GUARDAR ARCHIVO ACTUALIZADO ===
        GuardarTablaME5A(df_solpeds, nombre_archivo)

        # ============================================================
        # GENERAR ARCHIVO DE REPORTE CONSOLIDADO
        # ============================================================
        print(f"\n{'='*80}")
        print("GENERANDO REPORTE CONSOLIDADO...")
        print(f"{'='*80}")
        WriteLog(
            mensaje="Iniciando generación de reporte consolidado",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        ruta_reporte = GenerarReporteConsolidado(
            GenerarReporteConsolidado, nombre_archivo
        )

        if ruta_reporte:
            WriteLog(
                mensaje=f"Reporte consolidado generado exitosamente: {ruta_reporte}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
        else:
            WriteLog(
                mensaje="Error al generar reporte consolidado",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )

        # === RESUMEN FINAL ===
        print(f"\n{'='*80}")
        print("RESUMEN FINAL DE PROCESAMIENTO")
        print(f"{'='*80}")
        print(
            f"📊 SOLPEDs procesadas: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}"
        )
        print(f"📋 Items procesados: {contadores['items_procesados']}")
        print(f"Items validados: {contadores['items_validados']}")
        print(f"Items para verificación: {contadores['items_verificar_manual']}")
        print(f"🚫 Items sin texto: {contadores['items_sin_texto']}")
        print(
            f"SOLPEDs rechazadas (sin adjuntos): {contadores['rechazadas_sin_attachments']}"
        )
        if ruta_reporte:
            print(f"📄 Reporte consolidado: {os.path.basename(ruta_reporte)}")
        print(f"{'='*80}\n")

        WriteLog(
            mensaje=f"Resumen final - SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items: {contadores['items_procesados']}, Validados: {contadores['items_validados']}, "
            f"Revisar: {contadores['items_verificar_manual']}, Sin texto: {contadores['items_sin_texto']}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        WriteLog(
            mensaje=f"HU03 completado exitosamente",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        return True

    except Exception as e:
        print(f"Error en EjecutarHU03: {e}")
        WriteLog(
            mensaje=f"Error crítico en EjecutarHU03: {e}\n{traceback.format_exc()}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return False

# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N (Versión FINAL v3 - LÓGICA ORIGINAL RESTAURADA)
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N y genera reporte consolidado
# Ultima modificacion: 09/12/2025 - Versión 3 FINAL
# Propiedad de Colsubsidio
# Cambios v3:
#   - DataFrame del reporte creado AL PRINCIPIO
#   - Cada item se agrega INMEDIATAMENTE al reporte
#   - SIEMPRE valida items (aunque no tenga adjuntos)
#   - Lógica original de procesamiento restaurada
#   - Delimitador punto y coma (;)
#   - Adjuntos guardados SOLO en WriteLog
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import os
import traceback
import pandas as pd
from datetime import datetime
from funciones.EscribirLog import WriteLog
from funciones.GeneralME53N import (
    AbrirTransaccion,
    ColsultarSolped,
    procesarTablaME5A,
    ObtenerItemTextME53N,
    ObtenerItemsME53N,
    TablaItemsDataFrame,
    TraerSAPAlFrente_Opcion,
    ActualizarEstado,
    ActualizarEstadoYObservaciones,
    ProcesarYValidarItem,
    GuardarTablaME5A,
    ValidarAttachmentList,
    GenerarReporteAttachments,
    ParsearTablaAttachments,
    extraerDatosReporte,
)
from config.settings import RUTAS


def EjecutarHU03(session, nombre_archivo):
    try:
        task_name = "HU03_ValidacionME53N"

        print("=" * 80)
        print("INICIO HU03 - Validación ME53N con reporte consolidado incremental")
        print("=" * 80)
        WriteLog(
            mensaje="Inicio HU03 v3 (reporte incremental, lógica original)",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Traer SAP al frente
        TraerSAPAlFrente_Opcion()

        # Leer archivo
        df_solpeds = procesarTablaME5A(nombre_archivo)
        GuardarTablaME5A(df_solpeds, nombre_archivo)

        if df_solpeds.empty:
            print("ERROR: Archivo vacío")
            WriteLog(
                mensaje="Archivo expSolped03.txt vacío",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            return False

        # ============================================================
        # ✅ CREAR DATAFRAME DEL REPORTE AL PRINCIPIO
        # ============================================================
        columnas_reporte = [
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
            "Adjuntos",
            "Nombre de Adjunto",
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

        # ============================================================
        # MAPEO DINÁMICO DE CAMPOS ME53N (MULTI-IDIOMA / MULTI-LAYOUT)
        # ============================================================

        mapeoCamposReporte = {
            "Material_ME53N": ["Material", "Mat.", "Artículo"],
            "Short Text_ME53N": ["Short Text", "Texto breve", "Descripción"],
            "Quantity_ME53N": ["Cantidad", "Quantity", "Qty", "Menge"],
            "Un": ["Un", "UM", "UoM"],
            "Valn Price": [
                "PrecioVal.",
                "Valn Price",
                "Val. Price",
                "Precio valoración",
            ],
            "Crcy": ["Mon.", "Currency", "Crcy"],
            "Total Val.": ["Valor tot.", "Total Val.", "Total Value"],
            "Deliv.Date": ["Fe.entrega", "Delivery Date", "Deliv.Date"],
            "Fix. Vend.": ["ProvFijo", "Proveedor fijo", "Fix. Vend."],
            "Plant": ["Centro", "Plant"],
            "PGr_ME53N": ["GCp", "Grupo compras", "Purch. Group", "PGr"],
            "POrg": ["OrgC", "Purch. Org", "POrg"],
            "Matl Group": ["Gpo.artíc.", "Grupo artículo", "Matl Group"],
            "Pedido": ["Pedido", "PO", "Purchase Order"],
        }
        print(
            f"✅ DataFrame del reporte inicializado ({len(columnas_reporte)} columnas)"
        )
        WriteLog(
            mensaje=f"DataFrame del reporte inicializado: {len(columnas_reporte)} columnas",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Validar columnas
        columnas_requeridas = ["Estado", "Observaciones"]
        for columna in columnas_requeridas:
            if columna not in df_solpeds.columns:
                print(f"ERROR: Columna '{columna}' no encontrada")
                WriteLog(
                    mensaje=f"Columna requerida '{columna}' no encontrada",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                return False

        # Limpiar SOLPEDs
        solped_unicos = df_solpeds["PurchReq"].unique().tolist()
        solped_unicos_filtradas = []

        for solped in solped_unicos:
            solped_str = str(solped).strip()
            if (
                solped_str
                and solped_str not in ["Purch.Req.", "PurchReq", "Purch.Req", ""]
                and not any(
                    header in solped_str for header in ["Purch.Req", "PurchReq"]
                )
                and solped_str.replace(".", "").isdigit()
            ):
                solped_limpia = solped_str.replace(".", "")
                solped_unicos_filtradas.append(
                    solped_limpia if solped_limpia.isdigit() else solped_str
                )

        solped_unicos = solped_unicos_filtradas

        if not solped_unicos:
            print("ERROR: No hay SOLPEDs válidas")
            WriteLog(
                mensaje="No se encontraron SOLPEDs válidas",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            return False

        print(f"Procesando {len(solped_unicos)} SOLPEDs...")
        WriteLog(
            mensaje=f"Procesando {len(solped_unicos)} SOLPEDs",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Abrir transacción ME53N
        AbrirTransaccion(session, "ME53N")

        # Contadores
        contadores = {
            "total_solpeds": len(solped_unicos),
            "procesadas_exitosamente": 0,
            "con_errores": 0,
            "sin_items": 0,
            "items_procesados": 0,
            "items_validados": 0,
            "items_sin_texto": 0,
            "items_verificar_manual": 0,
            "rechazadas_sin_attachments": 0,
        }

        # ============================================================
        # PROCESAR CADA SOLPED
        # ============================================================
        for solped in solped_unicos:
            print(f"\n{'='*80}")
            print(f"PROCESANDO SOLPED: {solped}")
            print(f"{'='*80}")
            WriteLog(
                mensaje=f"Procesando SOLPED: {solped}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            solped_rechazada_por_attachments = False

            try:
                # Consultar SOLPED
                if not ColsultarSolped(session, solped):
                    print(f"❌ Error consultando SOLPED {solped}")
                    WriteLog(
                        mensaje=f"Error consultando SOLPED {solped}",
                        estado="ERROR",
                        task_name=task_name,
                        path_log=RUTAS["PathLogError"],
                    )
                    ActualizarEstado(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="❌ Error consulta",
                    )
                    contadores["con_errores"] += 1
                    continue

                # ============================================================
                # VALIDAR ATTACHMENT LIST
                # ============================================================
                print(f"\n--- Validando Attachment List ---")
                WriteLog(
                    mensaje=f"Validando attachments SOLPED {solped}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

                tiene_attachments, contenido_attachments, obs_attachments = (
                    ValidarAttachmentList(session, solped)
                )

                # Parsear attachments
                attachments_lista = (
                    ParsearTablaAttachments(contenido_attachments)
                    if contenido_attachments
                    else []
                )

                # Generar reporte
                reporte_attachments = GenerarReporteAttachments(
                    solped, tiene_attachments, contenido_attachments, obs_attachments
                )
                print(reporte_attachments)

                # ✅ Guardar SOLO en WriteLog
                if tiene_attachments and contenido_attachments:
                    WriteLog(
                        mensaje=f"SOLPED {solped} - Attachments:\n{reporte_attachments}\n\nContenido:\n{contenido_attachments}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    print(f"✅ Info adjuntos guardada en WriteLog")
                else:
                    WriteLog(
                        mensaje=f"SOLPED {solped}: Sin attachments. {obs_attachments}",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    print(f"⚠️ SOLPED {solped} sin adjuntos")

                # ✅ MARCAR SI NO TIENE ATTACHMENTS (pero CONTINUAR)
                if not tiene_attachments:
                    print(f"\n❌ SOLPED {solped} SERÁ RECHAZADA: Sin archivos adjuntos")
                    print(f"⚠️  CONTINUANDO con validaciones para reporte completo...")
                    WriteLog(
                        mensaje=f"SOLPED {solped} rechazada por falta de adjuntos (continúa procesamiento)",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    contadores["rechazadas_sin_attachments"] += 1
                    solped_rechazada_por_attachments = True
                else:
                    print(f"✅ SOLPED {solped} tiene attachments")

                # Preparar info adjuntos
                cantidad_adjuntos = len(attachments_lista)
                nombres_adjuntos = [att["title"] for att in attachments_lista]
                adjuntos_info = str(cantidad_adjuntos)
                nombres_adjuntos_str = (
                    ", ".join(nombres_adjuntos) if nombres_adjuntos else ""
                )

                # ============================================================
                # ✅ OBTENER ITEMS (LÓGICA ORIGINAL)
                # ============================================================
                dtItems = ObtenerItemsME53N(session, solped)
                print("📋 Columnas dtItems:", list(dtItems.columns))
                WriteLog(
                    mensaje=f"Columnas dtItems SOLPED {solped}: {list(dtItems.columns)}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

                if dtItems is None or dtItems.empty:
                    print(f"⚠️ SOLPED {solped} sin items")
                    WriteLog(
                        mensaje=f"SOLPED {solped} sin items",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Sin Items",
                        observaciones="No se encontraron items en SAP",
                    )
                    contadores["sin_items"] += 1
                    continue

                dtItems["Item"] = dtItems["Pos."].astype(str).str.strip()

                print(f"✅ Items encontrados: {dtItems.shape[0]}")
                WriteLog(
                    mensaje=f"SOLPED {solped}: {dtItems.shape[0]} items",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

                # ✅ Convertir a lista (LÓGICA ORIGINAL)
                lista_dicts = dtItems.to_dict(orient="records")

                # Filtrar totales
                if lista_dicts:
                    ultima_fila = lista_dicts[-1]
                    if (
                        ultima_fila.get("Status", "").strip() == "*"
                        or ultima_fila.get("Item", "").strip() == ""
                        or ultima_fila.get("Material", "").strip() == ""
                    ):
                        lista_dicts.pop()
                        print(f"Fila de total eliminada")

                # ============================================================
                # ✅ PROCESAR CADA ITEM (LÓGICA ORIGINAL RESTAURADA)
                # ============================================================
                contador_con_texto = 0
                contador_validados = 0
                contador_verificar_manual = 0
                contador_sin_texto_solped = 0

                for fila in lista_dicts:
                    numero_item = str(fila.get("Item", "")).strip()

                    # Validación fuerte del Item SAP
                    if not numero_item.isdigit():
                        WriteLog(
                            mensaje=f"Item inválido detectado en SOLPED {solped}: '{numero_item}'. Fila ignorada.",
                            estado="WARNING",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )
                        continue  # 👈 SOLO aquí

                    contadores["items_procesados"] += 1

                    print(f"\n--- Procesando Item {numero_item} ---")
                    WriteLog(
                        mensaje=f"Procesando Item {numero_item} de SOLPED {solped}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                    # Marcar como procesando
                    ActualizarEstado(
                        df_solpeds, nombre_archivo, solped, numero_item, "Procesando"
                    )

                    time.sleep(0.5)

                    # Generar Id
                    item_str = str(numero_item).strip()
                    solped_str = str(solped).strip()
                    id_item = f"{solped_str}{numero_item}"

                    print(f"📌 ID: {id_item}")

                    # ============================================================
                    # OBTENER DATOS PARA EL REPORTE
                    # ============================================================

                    # Datos de expSolped03
                    mask_item = (
                        df_solpeds["PurchReq"].astype(str).str.strip() == solped_str
                    ) & (df_solpeds["Item"].astype(str).str.strip() == item_str)

                    datos_expsolped = {
                        "PurchReq": solped_str,
                        "Item": item_str,
                        "ReqDate": "",
                        "Material": "",
                        "Created": "",
                        "ShortText": "",
                        "PO": "",
                        "Quantity": "",
                        "Plnt": "",
                        "PGr": "",
                        "Blank1": "",
                        "D": "",
                        "Requisnr": "",
                        "ProcState": "",
                    }

                    if mask_item.any():
                        fila_item = df_solpeds[mask_item].iloc[0]
                        for campo in [
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
                        ]:
                            datos_expsolped[campo] = str(
                                fila_item.get(campo, "")
                            ).strip()

                    # Agregar adjuntos
                    datos_expsolped["Adjuntos"] = adjuntos_info
                    datos_expsolped["Nombre de Adjunto"] = nombres_adjuntos_str

                    # Datos de tabla ME53N (desde el DataFrame dtItems)
                    datos_me53n = extraerDatosReporte(
                        fila=fila, df=dtItems, mapeo=mapeoCamposReporte
                    )

                    # ============================================================
                    # VALIDAR CAMPOS OBLIGATORIOS SAP ME53N
                    # ============================================================

                    CAMPOS_OBLIGATORIOS = [
                        "Material_ME53N",
                        "Quantity_ME53N",
                        "Valn Price",
                        "Deliv.Date",
                        "Plant",
                        "PGr_ME53N",
                        "POrg",
                        "Fix. Vend.",
                    ]

                    faltantes = [
                        campo
                        for campo in CAMPOS_OBLIGATORIOS
                        if not datos_me53n.get(campo)
                    ]

                    if faltantes:
                        resultado_campos_obligatorios = (
                            "Faltan campos obligatorios: " + ", ".join(faltantes)
                        )
                    else:
                        resultado_campos_obligatorios = "Campos obligatorios completos"

                    # ============================================================
                    # ✅ OBTENER TEXTO Y VALIDAR (LÓGICA ORIGINAL)
                    # ============================================================
                    texto = ObtenerItemTextME53N(session, solped, numero_item)

                    datos_validacion = {
                        "Id": id_item,
                        "PurchReq_Texto": solped_str,
                        "Item_Texto": item_str,
                        "Razon Social:": "",
                        "NIT:": "",
                        "Correo:": "",
                        "Concepto:": "",
                        "Cantidad:": "",
                        "Valor Unitario:": "",
                        "Valor Total:": "",
                        "Responsable:": "",
                        "CECO:": "",
                        "CAMPOS OBLIGATORIOS SAP ME53N:": resultado_campos_obligatorios,
                        "DATOS EXTRAIDOS Faltantes": "",
                        "DATOS EXTRAIDOS DEL TEXTO": "",
                        "DATOS EXTRAIDOS DEL TEXTO faltantes": "",
                        "CANTIDAD": "",
                        "VALOR_UNITARIO": "",
                        "VALOR_TOTAL": "",
                        "CONCEPTO": "",
                        "VALIDACIONES": "",
                        "Estado": "",
                        "Observaciones": "",
                    }

                    # ✅ VALIDAR SIEMPRE SI HAY TEXTO
                    if texto and texto.strip():
                        contador_con_texto += 1

                        print(f"✅ Item con texto - Validando...")
                        WriteLog(
                            mensaje=f"Item {numero_item}: Texto encontrado, validando",
                            estado="INFO",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )

                        # ✅ VALIDACIÓN COMPLETA (LÓGICA ORIGINAL)
                        (
                            datos_texto,
                            validaciones,
                            reporte,
                            estado_final,
                            observaciones,
                        ) = ProcesarYValidarItem(
                            session,
                            solped,
                            numero_item,
                            texto,
                            dtItems,
                            tiene_attachments,
                            obs_attachments,
                            attachments_lista,
                        )

                        # Actualizar datos de validación
                        datos_validacion.update(
                            {
                                "Razon Social:": str(
                                    datos_texto.get("razon_social", "")
                                ).strip(),
                                "NIT:": str(datos_texto.get("nit", "")).strip(),
                                "Correo:": str(datos_texto.get("correo", "")).strip(),
                                "Concepto:": str(
                                    datos_texto.get("concepto_compra", "")
                                ).strip(),
                                "Cantidad:": str(
                                    datos_texto.get("cantidad", "")
                                ).strip(),
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
                                "DATOS EXTRAIDOS Faltantes": str(
                                    validaciones.get("datos_extraidos_faltantes", "")
                                ).strip(),
                                "DATOS EXTRAIDOS DEL TEXTO": str(
                                    validaciones.get("datos_extraidos_texto", "")
                                ).strip(),
                                "DATOS EXTRAIDOS DEL TEXTO faltantes": str(
                                    validaciones.get("datos_texto_faltantes", "")
                                ).strip(),
                                "CANTIDAD": str(
                                    validaciones.get("cantidad", "")
                                ).strip(),
                                "VALOR_UNITARIO": str(
                                    validaciones.get("valor_unitario", "")
                                ).strip(),
                                "VALOR_TOTAL": str(
                                    validaciones.get("valor_total", "")
                                ).strip(),
                                "CONCEPTO": str(
                                    validaciones.get("concepto", "")
                                ).strip(),
                                "VALIDACIONES": str(
                                    validaciones.get("resumen_validaciones", "")
                                ).strip(),
                                "Estado": str(estado_final).strip(),
                                "Observaciones": str(observaciones).strip(),
                            }
                        )

                        if "validado" in estado_final.lower():
                            contador_validados += 1
                            contadores["items_validados"] += 1
                            print(f"✅ Item validado")
                        else:
                            contador_verificar_manual += 1
                            contadores["items_verificar_manual"] += 1
                            print(f"⚠️ Item requiere revisión")

                        WriteLog(
                            mensaje=f"Item {numero_item} validado: {estado_final}",
                            estado="INFO",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )
                    else:
                        contador_sin_texto_solped += 1
                        contadores["items_sin_texto"] += 1
                        datos_validacion["Estado"] = "⚠️ Sin texto"
                        datos_validacion["Observaciones"] = "Item sin texto"
                        print(f"⚠️ Item sin texto")
                        WriteLog(
                            mensaje=f"Item {numero_item}: Sin texto",
                            estado="WARNING",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )

                    # ============================================================
                    # ✅ AGREGAR ITEM AL REPORTE INMEDIATAMENTE
                    # ============================================================
                    datos_completos_item = {}
                    datos_completos_item.update(datos_expsolped)
                    datos_completos_item.update(datos_me53n)
                    datos_completos_item.update(datos_validacion)

                    df_reporte_consolidado = pd.concat(
                        [df_reporte_consolidado, pd.DataFrame([datos_completos_item])],
                        ignore_index=True,
                    )

                    print(f"✅ Item agregado al reporte (ID: {id_item})")
                    WriteLog(
                        mensaje=f"Item {numero_item} agregado al reporte (ID: {id_item})",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                # ============================================================
                # ACTUALIZAR ESTADO FINAL DE LA SOLPED
                # ============================================================
                if solped_rechazada_por_attachments:
                    estado_final = "❌ Rechazada - Sin Attachments"
                    obs_final = f"❌ RECHAZADA por falta de adjuntos | Items: {contador_validados} validados, {contador_verificar_manual} requieren revisión, {contador_sin_texto_solped} sin texto"
                elif contador_verificar_manual > 0:
                    estado_final = f"⚠️ Verificar manualmente"
                    obs_final = f"⚠️ {contador_verificar_manual}/{len(lista_dicts)} items requieren revisión + Attachments OK"
                else:
                    estado_final = (
                        f"✅ {contador_validados}/{len(lista_dicts)} validados"
                    )
                    obs_final = "Todos los items validados correctamente"

                ActualizarEstadoYObservaciones(
                    df_solpeds,
                    nombre_archivo,
                    solped,
                    nuevo_estado=estado_final,
                    observaciones=obs_final,
                )

                contadores["procesadas_exitosamente"] += 1
                print(f"\n✅ SOLPED {solped} completada: {estado_final}")
                WriteLog(
                    mensaje=f"SOLPED {solped} completada: {estado_final}",
                    estado="INFO",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

            except Exception as e:
                print(f"❌ Error procesando SOLPED {solped}: {e}")
                WriteLog(
                    mensaje=f"Error en SOLPED {solped}: {e}\n{traceback.format_exc()}",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                traceback.print_exc()
                ActualizarEstado(
                    df_solpeds, nombre_archivo, solped, nuevo_estado="❌ Error"
                )
                contadores["con_errores"] += 1

        # Guardar archivo actualizado
        GuardarTablaME5A(df_solpeds, nombre_archivo)

        # ============================================================
        # ✅ GUARDAR REPORTE CONSOLIDADO AL FINAL
        # ============================================================
        mapeoColumnasReporteES = {
            "Material_ME53N": "Material",
            "Short Text_ME53N": "Texto breve",
            "Quantity_ME53N": "Cantidad",
            "Un": "UM",
            "Valn Price": "Precio valoración",
            "Crcy": "Moneda",
            "Total Val.": "Valor Total",
            "Deliv.Date": "Fecha entrega",
            "Fix. Vend.": "Proveedor fijo",
            "Plant": "Centro",
            "PGr_ME53N": "Grupo compras",
            "POrg": "Organización compras",
            "Matl Group": "Grupo artículo",
            "Pedido": "Pedido",
        }
        print(f"\n{'='*80}")
        print("GUARDANDO REPORTE CONSOLIDADO...")
        print(f"{'='*80}")
        WriteLog(
            mensaje="Guardando reporte consolidado",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        try:
            fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_reporte = f"reporte_{fecha_actual}.txt"
            ruta_reporte = os.path.join(RUTAS["PathReportes"], nombre_reporte)


            # ============================================================
            # RENOMBRAR COLUMNAS A ESPAÑOL (SEGÚN SAP) PARA REPORTE FINAL
            # ============================================================

            df_reporte_consolidado = df_reporte_consolidado.rename(
                columns=mapeoColumnasReporteES
            )

            print("✅ Columnas del reporte renombradas a español (SAP)")
            WriteLog(
                mensaje="Columnas del reporte renombradas a español según SAP",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            df_reporte_consolidado.to_csv(
                ruta_reporte, sep=";", index=False, encoding="utf-8-sig", na_rep=""
            )

            print(f"✅ REPORTE CONSOLIDADO GUARDADO")
            print(f"📁 Archivo: {nombre_reporte}")
            print(f"📂 Ubicación: {ruta_reporte}")
            print(f"📊 Registros: {len(df_reporte_consolidado)}")
            print(f"🔹 Delimitador: ;")

            WriteLog(
                mensaje=f"Reporte consolidado guardado: {nombre_reporte} ({len(df_reporte_consolidado)} registros)",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        except Exception as e:
            print(f"❌ Error guardando reporte: {e}")
            WriteLog(
                mensaje=f"Error guardando reporte: {e}",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            traceback.print_exc()

        # RESUMEN FINAL
        print(f"\n{'='*80}")
        print("RESUMEN FINAL")
        print(f"{'='*80}")
        print(
            f"📊 SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}"
        )
        print(f"📋 Items: {contadores['items_procesados']}")
        print(f"✅ Validados: {contadores['items_validados']}")
        print(f"⚠️ Verificar: {contadores['items_verificar_manual']}")
        print(f"🚫 Sin texto: {contadores['items_sin_texto']}")
        print(f"❌ Rechazadas: {contadores['rechazadas_sin_attachments']}")
        print(f"📄 Reporte: {nombre_reporte}")
        print(f"{'='*80}\n")

        WriteLog(
            mensaje=f"HU03 completado - {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']} SOLPEDs, "
            f"{contadores['items_procesados']} items, {contadores['items_validados']} validados",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        return True

    except Exception as e:
        print(f"❌ Error crítico: {e}")
        WriteLog(
            mensaje=f"Error crítico: {e}\n{traceback.format_exc()}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return False

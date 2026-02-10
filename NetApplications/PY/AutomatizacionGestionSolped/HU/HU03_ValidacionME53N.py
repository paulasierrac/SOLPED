# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 06/02/2026
# Propiedad de Colsubsidio
# Cambios:
#   - Versión con validación completa y uso correcto de validaciones
#   - Notificaciones automáticas a responsables de Colsubsidio
#   - FIX: Corrección de datos faltantes en reporte final
#   - FIX: Corrección de duplicados en reporte
#   - FIX: Manejo robusto de errores de conversión
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import traceback
from funciones.EmailSender import EnviarNotificacionCorreo
from funciones.ReporteFinalME53N import (
    construir_fila_reporte_final,
    generar_reporte_final_excel,
)
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
    NotificarRevisionManualSolped,
    ValidarAttachmentList,
    GenerarReporteAttachments,
    ParsearTablaAttachments,
    convertir_txt_a_excel,
    AppendHipervinculoObservaciones,
    obtener_fila_expsolped,
    limpiar_numero_robusto,
    obtener_valor_desde_fila,
)
from config.settings import RUTAS


def EjecutarHU03(session, nombre_archivo):
    try:
        # ==========================
        # CONFIGURACIÓN DEL PROCESO
        # ==========================
        task_name = "HU03_ValidacionME53N"
        TraerSAPAlFrente_Opcion()
        # === Inicio HU03 ===
        WriteLog(
            mensaje="Inicio HU03 - Validación ME53N (Versión Corregida)",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Leer el archivo con las SOLPEDs a procesar
        df_solpeds = procesarTablaME5A(nombre_archivo)
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

        solped_unicos = solped_unicos_filtradas

        if not solped_unicos:
            print("ERROR: No se encontraron SOLPEDs validas para procesar")
            return False

        print(f"Procesando {len(solped_unicos)} SOLPEDs unicas...")

        # Informacion inicial del archivo
        print("RESUMEN INICIAL DEL ARCHIVO:")
        print(f"   - Total filas: {len(df_solpeds)}")
        print(f"   - SOLPEDs unicas validas: {len(solped_unicos)}")

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
            print(f"MODO DESARROLLO ACTIVO")
            print(f"Todos los correos se enviarán a: {EMAIL_DESARROLLO}")
            print(f"{'='*60}\n")
            WriteLog(
                mensaje=f"MODO DESARROLLO: Correos redirigidos a {EMAIL_DESARROLLO}",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Almacenar SOLPEDs que requirieron revisión para reporte final
        solpeds_con_problemas = []

        # fila reporte final
        filas_reporte_final = []

        # ========================================================
        # PROCESAR CADA SOLPED
        # ========================================================
        for solped in solped_unicos:
            print(f"\n{'='*80}")
            print(f"PROCESANDO SOLPED: {solped}")
            print(f"{'='*80}")

            # Variables para notificación
            correos_responsables = []
            resumen_validaciones = []
            requiere_notificacion = False

            try:
                # 1. Marcar SOLPED como "En Proceso"
                resultado_estado = ActualizarEstado(
                    df_solpeds, nombre_archivo, solped, nuevo_estado="En Proceso"
                )

                if not resultado_estado:
                    print(
                        f"ADVERTENCIA: No se pudo actualizar estado de SOLPED {solped}"
                    )
                    continue

                # 2. Consultar SOLPED en SAP
                resultado_consulta = ColsultarSolped(session, solped)
                if not resultado_consulta:
                    print(f"ERROR: No se pudo consultar SOLPED {solped} en SAP")
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Error Consulta",
                        observaciones="No se pudo consultar en SAP",
                    )
                    contadores["con_errores"] += 1
                    continue

                time.sleep(0.5)

                # ========================================================
                # 3. VALIDAR ATTACHMENT LIST
                # ========================================================
                print(f"\n--- Validando Attachment List ---")

                tiene_attachments, contenido_attachments, obs_attachments = (
                    ValidarAttachmentList(session, solped)
                )

                # Parsear attachments para información detallada
                attachments_lista = (
                    ParsearTablaAttachments(contenido_attachments)
                    if contenido_attachments
                    else []
                )

                # Generar reporte de attachments
                reporte_attachments = GenerarReporteAttachments(
                    solped, tiene_attachments, contenido_attachments, obs_attachments
                )
                print(reporte_attachments)

                # Guardar reporte de attachments SOLO si tiene adjuntos
                if attachments_lista:
                    path_reporte_attach = (
                        f"{RUTAS['PathReportes']}\\Attachments_{solped}.txt"
                    )
                    try:
                        with open(path_reporte_attach, "w", encoding="utf-8") as f:
                            f.write(reporte_attachments)
                        print(f"Reporte de attachments guardado: {path_reporte_attach}")
                    except Exception as e:
                        print(
                            f"Advertencia: No se pudo guardar reporte de attachments: {e}"
                        )
                else:
                    print(
                        f"No se genera archivo de adjuntos para SOLPED {solped} (sin archivos)"
                    )
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Sin Adjuntos",
                        observaciones="No cuenta con lista de Adjuntos",
                    )

                # MARCAR SI NO TIENE ATTACHMENTS (pero continuar validación)
                solped_rechazada_por_attachments = False

                if not attachments_lista:
                    print(f"\nSOLPED {solped} SERÁ RECHAZADA: Sin archivos adjuntos")
                    print(
                        f"Continuando con validaciones de items para reporte completo..."
                    )

                    contadores["rechazadas_sin_attachments"] += 1
                    solped_rechazada_por_attachments = True
                    requiere_notificacion = True

                    # Agregar a resumen de validaciones
                    resumen_validaciones.append(
                        f"\nMOTIVO DE RECHAZO PRINCIPAL\n"
                        f"   No cuenta con Attachment List\n"
                        f"   Acción requerida: Adjuntar documentación soporte\n"
                        f"   {obs_attachments}\n"
                        f"   Aunque se complete el resto de validaciones, la SOLPED queda RECHAZADA\n"
                    )

                else:
                    print(f"SOLPED {solped} tiene attachments - Continuando validación")

                    # Agregar info detallada de attachments a validaciones
                    info_attachments = (
                        f"\n📎 ATTACHMENT LIST ({len(attachments_lista)} archivo(s))\n"
                    )
                    info_attachments += f"   {obs_attachments}\n"

                    if attachments_lista:
                        info_attachments += f"\n   Archivos adjuntos:\n"
                        for i, attach in enumerate(
                            attachments_lista[:5], 1
                        ):  # Máximo 5 en resumen
                            info_attachments += f"   {i}. {attach['title'][:50]}\n"
                            info_attachments += f"      Creado por: {attach['creator']} - {attach['date']}\n"

                        if len(attachments_lista) > 5:
                            info_attachments += f"   ... y {len(attachments_lista) - 5} archivo(s) más\n"

                    resumen_validaciones.append(info_attachments)

                # 4. Obtener items de esta SOLPED
                dtItems = ObtenerItemsME53N(session, solped)

                # ============================================
                # DEBUG: Verificar columnas y datos de ME53N
                # ============================================
                if MODO_DESARROLLO and dtItems is not None and not dtItems.empty:
                    print(f"\n🔍 DEBUG - Columnas de dtItems:")
                    for col in dtItems.columns:
                        print(f"  '{col}'")

                    print(f"\n🔍 DEBUG - Primera fila de dtItems:")
                    primera_fila = dtItems.iloc[0].to_dict()
                    for key, val in primera_fila.items():
                        if any(
                            palabra in str(key)
                            for palabra in [
                                "Precio",
                                "Valor",
                                "Price",
                                "Total",
                                "Cantidad",
                            ]
                        ):
                            print(f"  {key}: '{val}' (tipo: {type(val).__name__})")
                # ============================================

                if dtItems is None or dtItems.empty:
                    contadores["sin_items"] += 1
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Sin Items",
                        observaciones="No se encontraron items en SAP",
                    )
                    print(f"ADVERTENCIA: SOLPED {solped}: Sin items en SAP")
                    continue

                print(f"Items encontrados en SAP: {dtItems.shape[0]}")

                # 5. Convertir a lista de diccionarios y filtrar totales
                lista_dicts = dtItems.to_dict(orient="records")

                # Filtrar: Eliminar la ultima fila si es un total
                if lista_dicts:
                    ultima_fila = lista_dicts[-1]
                    if (
                        ultima_fila.get("Status", "").strip() == "*"
                        or ultima_fila.get("Item", "").strip() == ""
                        or ultima_fila.get("Material", "").strip() == ""
                    ):
                        lista_dicts.pop()
                        print(f"Fila de total eliminada")

                # ========================================================
                # 6. PROCESAR CADA ITEM
                # ========================================================
                contador_con_texto = 0
                contador_validados = 0
                contador_verificar_manual = 0
                items_procesados_en_solped = len(lista_dicts)

                for i, fila in enumerate(lista_dicts):

                    print(f"\n🧩 ITEM {i} - Detalle de columnas ME53N:")
                    for col, val in fila.items():
                        print(f"  {col}: {str(val)[:50]}")

                    numero_item = fila.get("Pos.", str(i)).strip()
                    contadores["items_procesados"] += 1

                    print(f"\n--- Procesando Item {numero_item} ---")

                    # ============================================
                    # NUEVO: Obtener datos de expSolped03.txt
                    # ============================================
                    fila_exp = obtener_fila_expsolped(df_solpeds, solped, numero_item)

                    if fila_exp:
                        print(f"✅ Datos expSolped encontrados para item {numero_item}")
                        if MODO_DESARROLLO:
                            print(f"🔍 Valores clave de expSolped:")
                            print(f"  PurchReq: {fila_exp.get('PurchReq', 'N/A')}")
                            print(f"  ReqDate: {fila_exp.get('ReqDate', 'N/A')}")
                            print(f"  Created: {fila_exp.get('Created', 'N/A')}")
                            print(f"  ShortText: {fila_exp.get('ShortText', 'N/A')}")
                            print(f"  Quantity: {fila_exp.get('Quantity', 'N/A')}")
                    else:
                        print(
                            f"⚠️ No se encontraron datos expSolped para item {numero_item}"
                        )
                        fila_exp = {}
                    # ============================================

                    # ============================================
                    # NUEVO: Obtener datos específicos de ME53N
                    # ============================================
                    fila_me53n = fila  # Por defecto usar los mismos datos

                    if dtItems is not None and not dtItems.empty:
                        try:
                            # Buscar la fila correspondiente en dtItems
                            mascara = (
                                dtItems["Pos."].astype(str).str.strip()
                                == str(numero_item).strip()
                            )
                            filas_encontradas = dtItems[mascara]

                            if not filas_encontradas.empty:
                                fila_me53n = filas_encontradas.iloc[0].to_dict()
                                print(
                                    f"✅ Datos ME53N encontrados para item {numero_item}"
                                )

                                if MODO_DESARROLLO:
                                    print(f"🔍 Valores clave de ME53N:")
                                    print(
                                        f"  Cantidad: {fila_me53n.get('Cantidad', 'N/A')}"
                                    )
                                    print(
                                        f"  PrecioVal.: {fila_me53n.get('PrecioVal.', 'N/A')}"
                                    )
                                    print(
                                        f"  Valor tot.: {fila_me53n.get('Valor tot.', 'N/A')}"
                                    )
                            else:
                                print(
                                    f"⚠️ No se encontraron datos ME53N para item {numero_item}"
                                )
                        except Exception as e:
                            print(f"⚠️ Error buscando datos ME53N: {e}")
                    # ============================================

                    # Marcar item como "Procesando"
                    ActualizarEstado(
                        df_solpeds, nombre_archivo, solped, numero_item, "Procesando"
                    )

                    time.sleep(0.5)

                    # Obtener texto del editor SAP
                    texto = ObtenerItemTextME53N(session, solped, numero_item)

                    # Procesar y validar el texto
                    if texto and texto.strip():
                        contador_con_texto += 1

                        # VALIDACION COMPLETA DEL TEXTO
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
                            dtItems,  # ← CORREGIDO: DataFrame completo (no dict)
                            tiene_attachments,
                            obs_attachments,
                            attachments_lista,
                        )

                        # ========================================================
                        # CAPTURAR CORREOS DE COLSUBSIDIO PARA NOTIFICACIÓN
                        # ========================================================
                        responsable = datos_texto.get("responsable_compra", "")
                        if responsable and "@colsubsidio.com" in responsable.lower():
                            # Puede venir con múltiples correos separados por coma
                            correos_encontrados = [
                                email.strip()
                                for email in responsable.split(",")
                                if "@colsubsidio.com" in email.lower()
                            ]
                            correos_responsables.extend(correos_encontrados)
                            print(
                                f"Correo responsable detectado: {', '.join(correos_encontrados)}"
                            )

                        # Imprimir resumen de validacion DETALLADO
                        print(f"RESULTADO VALIDACION:")
                        print(f"  Estado: {estado_final}")
                        print(f"  Observaciones: {observaciones}")

                        # Mostrar resumen de validaciones
                        print(f"  Validaciones contra SAP:")
                        if "resumen" in validaciones:
                            print(f"    - {validaciones['resumen']}")

                        # Mostrar campos obligatorios
                        if "campos_obligatorios" in validaciones:
                            obligatorios = validaciones["campos_obligatorios"]
                            print(
                                f"    - Campos obligatorios: {obligatorios['presentes']}/{obligatorios['total']} presentes"
                            )
                            if obligatorios["faltantes"]:
                                print(
                                    f"    - Faltantes: {', '.join(obligatorios['faltantes'])}"
                                )

                        # Mostrar campos clave extraidos
                        campos_clave = [
                            "razon_social",
                            "nit",
                            "concepto_compra",
                            "cantidad",
                            "valor_total",
                        ]
                        print(f"  Campos clave extraidos:")
                        for campo in campos_clave:
                            if datos_texto.get(campo):
                                valor = datos_texto[campo]
                                if len(str(valor)) > 50:
                                    valor = str(valor)[:50] + "..."
                                print(f"    {campo}: {valor}")
                            else:
                                print(f"    {campo}: NO ENCONTRADO")

                        # Mostrar detalles de validaciones especificas
                        campos_validacion = [
                            "cantidad",
                            "valor_unitario",
                            "valor_total",
                            "fecha_entrega",
                            "concepto",
                        ]
                        print(f"  Detalles de validacion:")
                        for campo in campos_validacion:
                            if campo in validaciones and validaciones[campo].get(
                                "texto"
                            ):
                                estado_val = (
                                    "COINCIDE"
                                    if validaciones[campo]["match"]
                                    else "NO COINCIDE"
                                )
                                print(f"    {campo}: {estado_val}")
                                print(f"      Texto: {validaciones[campo]['texto']}")
                                print(f"      Tabla: {validaciones[campo]['tabla']}")
                                if validaciones[campo].get("diferencia"):
                                    print(
                                        f"      Diferencia: {validaciones[campo]['diferencia']}"
                                    )

                        # Guardar reporte detallado en archivo
                        path_reporte = f"{RUTAS['PathReportes']}\\Reporte_{solped}_{numero_item}.txt"
                        try:
                            with open(path_reporte, "w", encoding="utf-8") as f:
                                f.write(reporte)
                            print(f"Reporte guardado: {path_reporte}")
                        except Exception as e:
                            print(f"ADVERTENCIA: Error al guardar reporte: {e}")

                        # Actualizar estado y observaciones en el archivo principal
                        ActualizarEstadoYObservaciones(
                            df_solpeds,
                            nombre_archivo,
                            solped,
                            numero_item,
                            estado_final,
                            observaciones,
                        )
                    # DEBUG TEMPORAL
                    if MODO_DESARROLLO:
                        print(
                            f"\n🔍 DEBUG - Verificando datos antes de construir fila:"
                        )
                        print(f"  fila_exp keys: {list(fila_exp.keys())}")
                        print(f"  fila_me53n keys: {list(fila_me53n.keys())}")
                        print(f"  datos_texto keys: {list(datos_texto.keys())}")

                        print(
                            f"\n  fila_exp['PurchReq']: {fila_exp.get('PurchReq', 'FALTA')}"
                        )
                        print(f"  fila_exp['Item']: {fila_exp.get('Item', 'FALTA')}")
                        print(
                            f"  datos_texto['razon_social']: {datos_texto.get('razon_social', 'FALTA')}"
                        )
                        print(
                            f"  datos_texto['nit']: {datos_texto.get('nit', 'FALTA')}"
                        )
                        # ========================================================
                        # FILTRO CRÍTICO: evitar fila TOTAL / sin item válido
                        # ========================================================
                        if (
                            not numero_item
                            or not str(numero_item).strip().isdigit()
                            or str(numero_item).strip() in ["", "0"]
                        ):
                            print(
                                f"Fila ignorada (item inválido o total): '{numero_item}'"
                            )
                            continue

                        # ========================================================
                        # CONSTRUIR FILA PARA REPORTE FINAL (UNA SOLA VEZ)
                        # ========================================================
                        fila_reporte = construir_fila_reporte_final(
                            solped=solped,
                            item=numero_item,
                            datos_exp=fila_exp,  # ← CORREGIDO: Datos de expSolped03.txt
                            datos_adjuntos={
                                "cantidad": len(attachments_lista),
                                "nombres": ", ".join(
                                    [a["title"] for a in attachments_lista]
                                ),
                            },
                            datos_me53n=fila_me53n,  # ← CORREGIDO: Datos de ME53N
                            datos_texto=datos_texto,
                            resultado_validaciones={
                                "faltantes_me53n": validaciones.get(
                                    "campos_obligatorios", {}
                                ).get("faltantes"),
                                "faltantes_texto": validaciones.get("faltantes_texto"),
                                "cantidad": validaciones.get("cantidad", {}).get(
                                    "match"
                                ),
                                "valor_unitario": validaciones.get(
                                    "valor_unitario", {}
                                ).get("match"),
                                "valor_total": validaciones.get("valor_total", {}).get(
                                    "match"
                                ),
                                "concepto": validaciones.get("concepto", {}).get(
                                    "match"
                                ),
                                "estado": estado_final,
                                "observaciones": observaciones,
                            },
                        )

                        filas_reporte_final.append(fila_reporte)
                        print(f"📊 Fila agregada al reporte para item {numero_item}")
                        # ========================================================
                        # CONSTRUIR RESUMEN PARA NOTIFICACIÓN (DETALLADO POR ITEM)
                        # ========================================================

                        if estado_final != "Aprobado":

                            requiere_notificacion = True

                            item_info = f"\n📋 ITEM {numero_item}\n"
                            item_info += f"Estado: {estado_final}\n"
                            item_info += f"Observaciones: {observaciones}\n\n"

                            # -------- ME53N --------
                            faltantes_me53n = validaciones.get(
                                "campos_obligatorios", {}
                            ).get("faltantes", [])
                            if faltantes_me53n:
                                item_info += (
                                    f"- ME53N faltantes: {', '.join(faltantes_me53n)}\n"
                                )
                            else:
                                item_info += "- ME53N faltantes: Ninguno\n"

                            # -------- TEXTO --------
                            faltantes_texto = validaciones.get("faltantes_texto", [])
                            if faltantes_texto:
                                item_info += (
                                    f"- Texto faltantes: {', '.join(faltantes_texto)}\n"
                                )
                            else:
                                item_info += "- Texto faltantes: Ninguno\n"

                            # -------- VALIDACIONES --------
                            def estado_ok(flag):
                                return "OK" if flag else "ERROR"

                            item_info += "\nValidaciones:\n"
                            item_info += f"  Cantidad: {estado_ok(validaciones.get('cantidad', {}).get('match', False))}\n"
                            item_info += f"  Valor Unitario: {estado_ok(validaciones.get('valor_unitario', {}).get('match', False))}\n"
                            item_info += f"  Valor Total: {estado_ok(validaciones.get('valor_total', {}).get('match', False))}\n"
                            item_info += f"  Concepto: {estado_ok(validaciones.get('concepto', {}).get('match', False))}\n"

                            resumen_validaciones.append(item_info)

                        # Contar segun el resultado
                        if estado_final == "Aprobado":
                            contador_validados += 1
                            contadores["items_validados"] += 1
                            print(
                                f"EXITO: Item {numero_item} VALIDADO para orden de compra"
                            )
                        else:
                            contador_verificar_manual += 1
                            contadores["items_verificar_manual"] += 1
                            print(
                                f"ADVERTENCIA: Item {numero_item} requiere verificacion manual"
                            )

                    else:
                        # Sin texto en el editor
                        contadores["items_sin_texto"] += 1
                        observaciones_item = (
                            "Texto no encontrado en el editor SAP - No se puede validar"
                        )
                        ActualizarEstadoYObservaciones(
                            df_solpeds,
                            nombre_archivo,
                            solped,
                            numero_item,
                            "Sin Texto",
                            observaciones_item,
                        )
                        print(f"Item {numero_item}: Sin texto - No validado")

                        # También requiere notificación
                        requiere_notificacion = True
                        resumen_validaciones.append(
                            f"\n📋 ITEM {numero_item}:\n"
                            f"   Estado: Sin Texto\n"
                            f"   Observaciones: {observaciones_item}\n"
                        )

                # ========================================================
                # 7. ESTADO FINAL DE LA SOLPED (considerando attachments)
                # ========================================================
                if solped_rechazada_por_attachments:
                    # SOLPED rechazada por falta de attachments (independiente de items)
                    estado_final_solped = "Rechazada"
                    observaciones_solped = (
                        f"RECHAZADA por falta de adjuntos - "
                        f"Items: {contador_validados} validados, "
                        f"{contador_verificar_manual} requieren revisión, "
                        f"{items_procesados_en_solped - contador_con_texto} sin texto"
                    )

                elif contador_validados == items_procesados_en_solped:
                    estado_final_solped = "Aprobado"
                    observaciones_solped = f"Todos validados ({contador_validados} de {items_procesados_en_solped}) + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1
                    requiere_notificacion = False

                elif contador_verificar_manual > 0:
                    estado_final_solped = "Pendiente"
                    observaciones_solped = f"{contador_verificar_manual} de {items_procesados_en_solped} items requieren revisión + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1

                else:
                    estado_final_solped = "Rechazada"
                    observaciones_solped = "No se pudo procesar correctamente"
                    contadores["con_errores"] += 1

                ActualizarEstadoYObservaciones(
                    df_solpeds,
                    nombre_archivo,
                    solped,
                    nuevo_estado=estado_final_solped,
                    observaciones=observaciones_solped,
                )

                print(f"\n{'='*60}")
                if solped_rechazada_por_attachments:
                    print(f"SOLPED {solped} RECHAZADA (Sin Attachments)")
                else:
                    print(f"SOLPED {solped} completada")
                print(f"  Estado final: {estado_final_solped}")
                print(f"  Observaciones: {observaciones_solped}")
                print(f"  Total filas en reporte: {len(filas_reporte_final)}")
                print(f"{'='*60}")

                # ========================================================
                # ENVIAR NOTIFICACIÓN SI ES NECESARIO (UNA POR SOLPED)
                # ========================================================
                if requiere_notificacion and correos_responsables:
                    # Eliminar duplicados de correos
                    correos_unicos = list(set(correos_responsables))

                    # ========================================================
                    # MODO DESARROLLO - REDIRIGIR CORREOS
                    # ========================================================
                    if MODO_DESARROLLO:
                        correos_originales = correos_unicos.copy()
                        correos_unicos = [EMAIL_DESARROLLO]
                        print(f"\n{'='*60}")
                        print(f"NOTIFICACIÓN (MODO DESARROLLO)")
                        print(f"{'='*60}")
                        print(
                            f"Destinatarios originales: {', '.join(correos_originales)}"
                        )
                        print(f"Redirigido a: {EMAIL_DESARROLLO}")
                    else:
                        print(f"\n{'='*60}")
                        print(f"ENVIANDO NOTIFICACIÓN DE REVISIÓN MANUAL")
                        print(f"{'='*60}")
                        print(f"Destinatarios: {', '.join(correos_unicos)}")

                    # Construir texto completo de validaciones
                    texto_validaciones = f"SOLPED: {solped}\n"

                    # Agregar info de modo desarrollo
                    if MODO_DESARROLLO:
                        texto_validaciones += f"\nMODO DESARROLLO - CORREO DE PRUEBA\n"
                        texto_validaciones += f"Destinatarios originales: {', '.join(correos_originales)}\n"
                        texto_validaciones += f"{'='*60}\n\n"

                    texto_validaciones += f"Estado Final: {estado_final_solped}\n"
                    texto_validaciones += f"Total Items: {items_procesados_en_solped}\n"
                    texto_validaciones += f"Items Validados: {contador_validados}\n"
                    texto_validaciones += (
                        f"Items Requieren Revisión: {contador_verificar_manual}\n"
                    )
                    texto_validaciones += f"Items Sin Texto: {items_procesados_en_solped - contador_con_texto}\n"
                    texto_validaciones += f"\n{'='*60}\n"
                    texto_validaciones += f"DETALLE POR ITEM:\n"
                    texto_validaciones += "".join(resumen_validaciones)

                    # Enviar notificación
                    try:
                        exito_notificacion = NotificarRevisionManualSolped(
                            destinatarios=correos_unicos,
                            numero_solped=solped,
                            validaciones=texto_validaciones,
                            task_name=task_name,
                        )

                        if exito_notificacion:
                            if MODO_DESARROLLO:
                                print(
                                    f"[DESARROLLO] Correo enviado a {EMAIL_DESARROLLO}"
                                )
                                print(f"   (Original: {', '.join(correos_originales)})")
                            else:
                                print(
                                    f"Notificación enviada correctamente a {len(correos_unicos)} destinatario(s)"
                                )
                            contadores["notificaciones_enviadas"] += 1

                            # Guardar info para reporte final
                            solpeds_con_problemas.append(
                                {
                                    "solped": solped,
                                    "estado": estado_final_solped,
                                    "tiene_attachments": tiene_attachments,
                                    "obs_attachments": obs_attachments,
                                    "attachments_detalle": (
                                        attachments_lista[:10]
                                        if attachments_lista
                                        else []
                                    ),
                                    "items_total": items_procesados_en_solped,
                                    "items_ok": contador_validados,
                                    "items_revisar": contador_verificar_manual,
                                    "items_sin_texto": items_procesados_en_solped
                                    - contador_con_texto,
                                    "responsables": (
                                        correos_originales
                                        if MODO_DESARROLLO
                                        else correos_unicos
                                    ),
                                    "detalle": resumen_validaciones,
                                }
                            )
                        else:
                            print(f"Error al enviar notificación")
                            contadores["notificaciones_fallidas"] += 1

                    except Exception as e_notif:
                        print(f"Error al enviar notificación: {e_notif}")
                        contadores["notificaciones_fallidas"] += 1
                        WriteLog(
                            mensaje=f"Error al enviar notificación para SOLPED {solped}: {e_notif}",
                            estado="WARNING",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )

                    print(f"{'='*60}\n")

                elif requiere_notificacion and not correos_responsables:
                    mensaje_advertencia = f"SOLPED {solped} requiere revisión pero NO se encontró correo @colsubsidio.com"

                    if MODO_DESARROLLO:
                        print(f"[DESARROLLO] {mensaje_advertencia}")
                        print(f"   Se enviaría notificación genérica en producción")
                    else:
                        print(f"{mensaje_advertencia}")

                    WriteLog(
                        mensaje=f"SOLPED {solped}: Requiere revisión pero sin correo de responsable",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

                    # Guardar para reporte final aunque no tenga responsable
                    solpeds_con_problemas.append(
                        {
                            "solped": solped,
                            "estado": estado_final_solped,
                            "items_total": items_procesados_en_solped,
                            "items_ok": contador_validados,
                            "items_revisar": contador_verificar_manual,
                            "items_sin_texto": items_procesados_en_solped
                            - contador_con_texto,
                            "responsables": [],
                            "detalle": resumen_validaciones,
                        }
                    )

            except Exception as e:
                contadores["con_errores"] += 1
                observaciones_error = f"Error durante procesamiento: {str(e)[:100]}"
                ActualizarEstadoYObservaciones(
                    df_solpeds,
                    nombre_archivo,
                    solped,
                    nuevo_estado="Error",
                    observaciones=observaciones_error,
                )
                print(f"ERROR procesando {solped}: {e}")
                WriteLog(
                    mensaje=f"Error procesando SOLPED {solped}: {e}",
                    estado="ERROR",
                    task_name="EjecutarHU03",
                    path_log=RUTAS["PathLogError"],
                )
                traceback.print_exc()
                continue

        # 8. Mostrar resumen final del proceso
        print(f"\n{'='*80}")
        print("PROCESO COMPLETADO - RESUMEN FINAL")
        print(f"{'='*80}")

        # Resumen detallado
        print(f"\nESTADISTICAS DEL PROCESO:")
        print(f"  SOLPEDs totales: {contadores['total_solpeds']}")
        print(
            f"  SOLPEDs procesadas exitosamente: {contadores['procesadas_exitosamente']}"
        )
        print(f"  SOLPEDs con errores: {contadores['con_errores']}")
        print(f"  SOLPEDs sin items: {contadores['sin_items']}")
        print(
            f"  SOLPEDs rechazadas sin attachments: {contadores['rechazadas_sin_attachments']}"
        )
        print(f"  Items procesados: {contadores['items_procesados']}")
        print(f"  Items validados para OC: {contadores['items_validados']}")
        print(f"  Items para verificar manual: {contadores['items_verificar_manual']}")
        print(f"  Items sin texto: {contadores['items_sin_texto']}")
        print(f"\nNOTIFICACIONES:")
        print(f"  Notificaciones enviadas: {contadores['notificaciones_enviadas']}")
        print(f"  Notificaciones fallidas: {contadores['notificaciones_fallidas']}")
        print(f"\nREPORTE FINAL:")
        print(f"  Total filas en reporte: {len(filas_reporte_final)}")

        # Recargar archivo para mostrar estados finales
        df_final = procesarTablaME5A(nombre_archivo)
        if not df_final.empty and "Estado" in df_final.columns:
            print("\nDISTRIBUCION FINAL DE ESTADOS:")
            resumen = df_final["Estado"].value_counts()
            for estado, cantidad in resumen.items():
                print(f"  {estado}: {cantidad}")

            # Mostrar algunas observaciones comunes
            if (
                "Observaciones" in df_final.columns
                and not df_final["Observaciones"].isna().all()
            ):
                print(f"\nOBSERVACIONES MAS FRECUENTES:")
                obs_count = df_final["Observaciones"].value_counts().head(5)
                for obs, count in obs_count.items():
                    if obs and str(obs).strip():
                        print(f"  '{obs[:50]}...': {count}")

        print("\n")
        WriteLog(
            mensaje=f"HU03 completado exitosamente. "
            f"SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items validados: {contadores['items_validados']}/{contadores['items_procesados']}, "
            f"Notificaciones enviadas: {contadores['notificaciones_enviadas']}, "
            f"Filas en reporte: {len(filas_reporte_final)}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # ========================================================
        # GENERAR ARCHIVO FINAL
        # ========================================================
        if filas_reporte_final:

            print("📊 Generando reporte final consolidado ME53N...")
            WriteLog(
                mensaje="Generando reporte final consolidado ME53N",
                estado="INFO",
                task_name="HU03_ValidacionME53N",
                path_log=RUTAS["PathLog"],
            )

            path_reporte = generar_reporte_final_excel(filas_reporte_final)

            if path_reporte:
                print(f"✅ Reporte final generado correctamente: {path_reporte}")
                WriteLog(
                    mensaje=f"Reporte final generado correctamente: {path_reporte}",
                    estado="OK",
                    task_name="HU03_ValidacionME53N",
                    path_log=RUTAS["PathLog"],
                )
            else:
                print("⚠️ No se pudo generar el reporte final")
                WriteLog(
                    mensaje="No se pudo generar el reporte final",
                    estado="WARNING",
                    task_name="HU03_ValidacionME53N",
                    path_log=RUTAS["PathLog"],
                )

        else:
            print("⚠️ No hay filas para generar el reporte final")
            WriteLog(
                mensaje="No hay filas para generar el reporte final",
                estado="WARNING",
                task_name="HU03_ValidacionME53N",
                path_log=RUTAS["PathLog"],
            )

        # Convertir a Excel y agregar hipervínculos
        convertir_txt_a_excel(nombre_archivo)
        archivo_descargado = rf"{RUTAS['PathInsumos']}/expSolped03.xlsx"
        AppendHipervinculoObservaciones(
            ruta_excel=archivo_descargado, carpeta_reportes=RUTAS["PathReportes"]
        )

        # Enviar correo de finalización
        EnviarNotificacionCorreo(
            codigo_correo=10, task_name=task_name, adjuntos=[archivo_descargado]
        )

        return True

    except Exception as e:
        WriteLog(
            mensaje=f"Error en EjecutarHU03: {e}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return False

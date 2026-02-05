# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 08/12/2025
# Propiedad de Colsubsidio
# Cambios:
#   - Versión con validación completa y uso correcto de validaciones
#   - Notificaciones automáticas a responsables de Colsubsidio
# =========================================
import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
import time
import traceback
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
    EnviarNotificacionCorreo,
    AppendHipervinculoObservaciones,
)
from config.settings import RUTAS


def EjecutarHU03(session, nombre_archivo):
    try:
        # ==========================
        # CONFIGURACIÓN DEL PROCESO
        # ==========================
        task_name = "HU03_ValidacionME53N"
        # Controla si el proceso debe detener validaciones cuando NO hay adjuntos
        CANCELAR_SI_NO_HAY_ADJUNTOS = (
            True  # ← ponlo en False si quieres seguir validando
        )

        # === Inicio HU03 ===
        WriteLog(
            mensaje="Inicio HU03 - Validación ME53N",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Traer SAP al frente
        # TraerSAPAlFrente_Opcion()

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
                # 3. VALIDAR ATTACHMENT LIST (NUEVA VALIDACIÓN)
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

                # Guardar reporte de attachments
                # Guardar reporte de attachments SOLO si tiene adjuntos
                if tiene_attachments and contenido_attachments:
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

                if not tiene_attachments:
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

                # 3. Obtener items de esta SOLPED
                dtItems = ObtenerItemsME53N(session, solped)

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

                # 4. Convertir a lista de diccionarios y filtrar totales
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
                # 5. PROCESAR CADA ITEM
                # ========================================================
                contador_con_texto = 0
                contador_validados = 0
                contador_verificar_manual = 0
                items_procesados_en_solped = len(lista_dicts)

                for i, fila in enumerate(lista_dicts):
                    numero_item = fila.get("Item", str(i)).strip()
                    contadores["items_procesados"] += 1

                    print(f"\n--- Procesando Item {numero_item} ---")

                    # Marcar item como "Procesando"
                    ActualizarEstado(
                        df_solpeds, nombre_archivo, solped, numero_item, "Procesando"
                    )

                    time.sleep(0.5)

                    # Obtener texto del editor SAP
                    texto = ObtenerItemTextME53N(session, solped, numero_item)
                    # print(texto)

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
                            dtItems,
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
                                if len(valor) > 50:
                                    valor = valor[:50] + "..."
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
                            if campo in validaciones and validaciones[campo]["texto"]:
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

                        # ========================================================
                        # CONSTRUIR RESUMEN PARA NOTIFICACIÓN
                        # ========================================================
                        if estado_final != "Registro validado para orden de compra":
                            requiere_notificacion = True

                            # Construir texto de validación del item
                            item_info = f"\n ITEM {numero_item}:\n"
                            item_info += f"   Estado: {estado_final}\n"
                            item_info += f"   Observaciones: {observaciones}\n"

                            # Agregar campos clave
                            if datos_texto.get("nit"):
                                item_info += f"   NIT: {datos_texto['nit']}\n"
                            if datos_texto.get("razon_social"):
                                item_info += (
                                    f"   Razón Social: {datos_texto['razon_social']}\n"
                                )
                            if datos_texto.get("concepto_compra"):
                                concepto_corto = datos_texto["concepto_compra"][:100]
                                item_info += f"   Concepto: {concepto_corto}...\n"

                            # Agregar problemas de validación
                            if validaciones.get("campos_obligatorios", {}).get(
                                "faltantes"
                            ):
                                item_info += f"   Campos faltantes: {', '.join(validaciones['campos_obligatorios']['faltantes'])}\n"

                            resumen_validaciones.append(item_info)

                        # Contar segun el resultado
                        if estado_final == "Registro validado para orden de compra":
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
                # 6. ESTADO FINAL DE LA SOLPED (considerando attachments)
                # ========================================================
                if solped_rechazada_por_attachments:
                    # SOLPED rechazada por falta de attachments (independiente de items)
                    estado_final_solped = "Rechazada"
                    observaciones_solped = (
                        f"RECHAZADA por falta de adjuntos | "
                        f"Items: {contador_validados} validados, "
                        f"{contador_verificar_manual} requieren revisión, "
                        f"{items_procesados_en_solped - contador_con_texto} sin texto"
                    )
                    # Ya fue contada en rechazadas_sin_attachments

                elif contador_validados == items_procesados_en_solped:
                    estado_final_solped = "Registro validado para orden de compra"
                    observaciones_solped = f"Todos validados ({contador_validados} de {items_procesados_en_solped}) + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1
                    requiere_notificacion = False

                elif contador_verificar_manual > 0:
                    estado_final_solped = "Verificar manualmente"
                    observaciones_solped = f"{contador_verificar_manual} de {items_procesados_en_solped} items requieren revisión + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1

                else:
                    estado_final_solped = "Sin procesar"
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
                                    ),  # Máximo 10 para el reporte
                                    "items_total": (
                                        items_procesados_en_solped
                                        if "items_procesados_en_solped" in locals()
                                        else 0
                                    ),
                                    "items_ok": (
                                        contador_validados
                                        if "contador_validados" in locals()
                                        else 0
                                    ),
                                    "items_revisar": (
                                        contador_verificar_manual
                                        if "contador_verificar_manual" in locals()
                                        else 0
                                    ),
                                    "items_sin_texto": (
                                        (
                                            items_procesados_en_solped
                                            - contador_con_texto
                                        )
                                        if "items_procesados_en_solped" in locals()
                                        and "contador_con_texto" in locals()
                                        else 0
                                    ),
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
                continue

        # 7. Mostrar resumen final del proceso
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
        print(f"  Items procesados: {contadores['items_procesados']}")
        print(f"  Items validados para OC: {contadores['items_validados']}")
        print(f"  Items para verificar manual: {contadores['items_verificar_manual']}")
        print(f"  Items sin texto: {contadores['items_sin_texto']}")
        print(f"\nNOTIFICACIONES:")
        print(f"  Notificaciones enviadas: {contadores['notificaciones_enviadas']}")
        print(f"  Notificaciones fallidas: {contadores['notificaciones_fallidas']}")

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
            f"Notificaciones enviadas: {contadores['notificaciones_enviadas']}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Ruta del archivo a convertir

        convertir_txt_a_excel(nombre_archivo)
        archivo_descargado = rf"{RUTAS['PathInsumos']}/expSolped03.xlsx"
        AppendHipervinculoObservaciones(
            ruta_excel=archivo_descargado, carpeta_reportes=RUTAS["PathReportes"]
        )

        # Enviar correo de inicio (código 2 adjunto)
        EnviarNotificacionCorreo(
            codigo_correo=54, task_name=task_name, adjuntos=[archivo_descargado]
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

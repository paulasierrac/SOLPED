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
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
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
)
from Config.settings import RUTAS


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
        TraerSAPAlFrente_Opcion()

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
            print(f"⚠️  MODO DESARROLLO ACTIVO")
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
                # ✅ 3. VALIDAR ATTACHMENT LIST (NUEVA VALIDACIÓN)
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
                        f"⚠️ No se genera archivo de adjuntos para SOLPED {solped} (sin archivos)"
                    )
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Sin Adjuntos",
                        observaciones="No cuenta con lista de Adjuntos",
                    )

                # ⚠️ MARCAR SI NO TIENE ATTACHMENTS (pero continuar validación)
                solped_rechazada_por_attachments = False

                if not tiene_attachments:
                    print(f"\n❌ SOLPED {solped} SERÁ RECHAZADA: Sin archivos adjuntos")
                    print(
                        f"⚠️  Continuando con validaciones de items para reporte completo..."
                    )

                    contadores["rechazadas_sin_attachments"] += 1
                    solped_rechazada_por_attachments = True
                    requiere_notificacion = True

                    # Agregar a resumen de validaciones
                    resumen_validaciones.append(
                        f"\n🚫 MOTIVO DE RECHAZO PRINCIPAL\n"
                        f"   ❌ No cuenta con Attachment List\n"
                        f"   Acción requerida: Adjuntar documentación soporte\n"
                        f"   {obs_attachments}\n"
                        f"   ⚠️ Aunque se complete el resto de validaciones, la SOLPED queda RECHAZADA\n"
                    )

                else:
                    print(
                        f"✅ SOLPED {solped} tiene attachments - Continuando validación"
                    )

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
                                f"📧 Correo responsable detectado: {', '.join(correos_encontrados)}"
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
                                item_info += f"   ⚠️ Campos faltantes: {', '.join(validaciones['campos_obligatorios']['faltantes'])}\n"

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
                    estado_final_solped = "Rechazada - Sin Attachments"
                    observaciones_solped = (
                        f"❌ RECHAZADA por falta de adjuntos | "
                        f"Items: {contador_validados} validados, "
                        f"{contador_verificar_manual} requieren revisión, "
                        f"{items_procesados_en_solped - contador_con_texto} sin texto"
                    )
                    # Ya fue contada en rechazadas_sin_attachments

                elif contador_validados == items_procesados_en_solped:
                    estado_final_solped = "Registro validado para orden de compra"
                    observaciones_solped = f"✅ Todos validados ({contador_validados}/{items_procesados_en_solped}) + Attachments OK"
                    contadores["procesadas_exitosamente"] += 1
                    requiere_notificacion = False

                elif contador_verificar_manual > 0:
                    estado_final_solped = "Verificar manualmente"
                    observaciones_solped = f"⚠️ {contador_verificar_manual}/{items_procesados_en_solped} items requieren revisión + Attachments OK"
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
                    print(f"❌ SOLPED {solped} RECHAZADA (Sin Attachments)")
                else:
                    print(f"✅ SOLPED {solped} completada")
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
                        print(f"📧 NOTIFICACIÓN (MODO DESARROLLO)")
                        print(f"{'='*60}")
                        print(
                            f"Destinatarios originales: {', '.join(correos_originales)}"
                        )
                        print(f"Redirigido a: {EMAIL_DESARROLLO}")
                    else:
                        print(f"\n{'='*60}")
                        print(f"📧 ENVIANDO NOTIFICACIÓN DE REVISIÓN MANUAL")
                        print(f"{'='*60}")
                        print(f"Destinatarios: {', '.join(correos_unicos)}")

                    # Construir texto completo de validaciones
                    texto_validaciones = f"SOLPED: {solped}\n"

                    # Agregar info de modo desarrollo
                    if MODO_DESARROLLO:
                        texto_validaciones += (
                            f"\n⚠️ MODO DESARROLLO - CORREO DE PRUEBA\n"
                        )
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
                                    f"✅ [DESARROLLO] Correo enviado a {EMAIL_DESARROLLO}"
                                )
                                print(f"   (Original: {', '.join(correos_originales)})")
                            else:
                                print(
                                    f"✅ Notificación enviada correctamente a {len(correos_unicos)} destinatario(s)"
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
                            print(f"❌ Error al enviar notificación")
                            contadores["notificaciones_fallidas"] += 1

                    except Exception as e_notif:
                        print(f"❌ Error al enviar notificación: {e_notif}")
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
                        print(f"⚠️  [DESARROLLO] {mensaje_advertencia}")
                        print(f"   Se enviaría notificación genérica en producción")
                    else:
                        print(f"⚠️  {mensaje_advertencia}")

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

        # ========================================================
        # ENVIAR REPORTE FINAL GENERAL A NETAPPLICATIONS
        # ========================================================
        if solpeds_con_problemas:
            print(f"\n{'='*80}")
            if MODO_DESARROLLO:
                print(f"📧 GENERANDO REPORTE FINAL CONSOLIDADO (MODO DESARROLLO)")
            else:
                print(f"📧 GENERANDO REPORTE FINAL CONSOLIDADO")
            print(f"{'='*80}")

            # Construir reporte consolidado
            reporte_final = f"REPORTE CONSOLIDADO - VALIDACIÓN ME53N\n"
            reporte_final += f"Fecha: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            reporte_final += f"Archivo procesado: {nombre_archivo}\n"

            if MODO_DESARROLLO:
                reporte_final += f"\n{'='*60}\n"
                reporte_final += f"⚠️ MODO DESARROLLO ACTIVO\n"
                reporte_final += f"Este es un correo de prueba.\n"
                reporte_final += (
                    f"En producción se enviaría a los destinatarios reales.\n"
                )
                reporte_final += f"{'='*60}\n"

            reporte_final += f"\n{'='*60}\n"
            reporte_final += f"RESUMEN GENERAL:\n"
            reporte_final += f"{'='*60}\n"
            reporte_final += (
                f"Total SOLPEDs procesadas: {contadores['total_solpeds']}\n"
            )
            reporte_final += f"SOLPEDs con problemas: {len(solpeds_con_problemas)}\n"
            reporte_final += f"SOLPEDs exitosas: {contadores['procesadas_exitosamente'] - len(solpeds_con_problemas)}\n"
            reporte_final += (
                f"Total items procesados: {contadores['items_procesados']}\n"
            )
            reporte_final += f"Items validados: {contadores['items_validados']}\n"
            reporte_final += (
                f"Items para verificar: {contadores['items_verificar_manual']}\n"
            )
            reporte_final += f"Items sin texto: {contadores['items_sin_texto']}\n"
            reporte_final += (
                f"Notificaciones enviadas: {contadores['notificaciones_enviadas']}\n"
            )
            reporte_final += (
                f"Notificaciones fallidas: {contadores['notificaciones_fallidas']}\n\n"
            )

            reporte_final += f"{'='*60}\n"
            reporte_final += f"DETALLE DE SOLPEDS CON PROBLEMAS:\n"
            reporte_final += f"{'='*60}\n\n"

            for idx, solped_info in enumerate(solpeds_con_problemas, 1):
                reporte_final += f"\n{idx}. SOLPED: {solped_info['solped']}\n"
                reporte_final += f"   Estado: {solped_info['estado']}\n"
                reporte_final += f"   Items Total: {solped_info['items_total']}\n"
                reporte_final += f"   Items Validados: {solped_info['items_ok']}\n"
                reporte_final += (
                    f"   Items Requieren Revisión: {solped_info['items_revisar']}\n"
                )
                reporte_final += (
                    f"   Items Sin Texto: {solped_info['items_sin_texto']}\n"
                )

                if solped_info["responsables"]:
                    if MODO_DESARROLLO:
                        reporte_final += f"   Responsables (no notificados - modo desarrollo): {', '.join(solped_info['responsables'])}\n"
                    else:
                        reporte_final += f"   Responsables notificados: {', '.join(solped_info['responsables'])}\n"
                else:
                    reporte_final += f"   ⚠️ Sin responsable identificado\n"

                reporte_final += f"\n   DETALLE DE ITEMS:\n"
                reporte_final += "".join(solped_info["detalle"])
                reporte_final += f"\n   {'-'*60}\n"

            # Enviar correo consolidado a NetApplications
            try:
                from Funciones.GeneralME53N import EnviarCorreoPersonalizado

                if MODO_DESARROLLO:
                    asunto_final = f"[DESARROLLO] 📊 Reporte Consolidado - {len(solpeds_con_problemas)} SOLPEDs requieren atención"
                else:
                    asunto_final = f"📊 Reporte Consolidado Validación ME53N - {len(solpeds_con_problemas)} SOLPEDs"

                cuerpo_final = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; background-color: #f4f4f4; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f4; padding: 20px;">
        <tr>
            <td align="center">
                <table width="700" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                """

                # Banner de modo desarrollo
                if MODO_DESARROLLO:
                    cuerpo_final += f"""
                    <!-- Banner Desarrollo -->
                    <tr>
                        <td style="background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 20px 40px;">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td width="50">
                                        <div style="font-size: 36px;">⚠️</div>
                                    </td>
                                    <td>
                                        <h3 style="margin: 0 0 5px 0; color: #856404; font-size: 18px;">MODO DESARROLLO</h3>
                                        <p style="margin: 0; color: #856404; font-size: 14px;">
                                            Este es un correo de prueba. Los destinatarios reales NO fueron notificados.
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    """

                cuerpo_final += f"""
                    <!-- Header -->
                    <tr>
                        <td style="background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); padding: 40px 40px 30px 40px;">
                            <h1 style="margin: 0; color: #ffffff; font-size: 32px; font-weight: 600; text-align: center;">
                                📊 Reporte Consolidado
                            </h1>
                            <p style="margin: 10px 0 0 0; color: #ecf0f1; font-size: 16px; text-align: center; opacity: 0.95;">
                                Validación ME53N - {time.strftime('%d/%m/%Y %H:%M')}
                            </p>
                        </td>
                    </tr>
                    
                    <!-- Resumen Ejecutivo -->
                    <tr>
                        <td style="padding: 35px 40px 25px 40px;">
                            <h2 style="margin: 0 0 20px 0; color: #2c3e50; font-size: 22px; border-bottom: 3px solid #3498db; padding-bottom: 10px;">
                                Resumen Ejecutivo
                            </h2>
                            
                            <!-- Estadísticas en Grid -->
                            <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 25px;">
                                <tr>
                                    <!-- Total Procesadas -->
                                    <td width="33%" style="padding: 20px 15px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; text-align: center;">
                                        <div style="color: #ffffff; font-size: 38px; font-weight: bold; margin-bottom: 8px;">
                                            {contadores['total_solpeds']}
                                        </div>
                                        <div style="color: #ffffff; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">
                                            SOLPEDs<br>Procesadas
                                        </div>
                                    </td>
                                    
                                    <td width="2%"></td>
                                    
                                    <!-- Con Problemas -->
                                    <td width="33%" style="padding: 20px 15px; background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); border-radius: 8px; text-align: center;">
                                        <div style="color: #ffffff; font-size: 38px; font-weight: bold; margin-bottom: 8px;">
                                            {len(solpeds_con_problemas)}
                                        </div>
                                        <div style="color: #ffffff; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">
                                            Requieren<br>Atención
                                        </div>
                                    </td>
                                    
                                    <td width="2%"></td>
                                    
                                    <!-- Items Validados -->
                                    <td width="33%" style="padding: 20px 15px; background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); border-radius: 8px; text-align: center;">
                                        <div style="color: #ffffff; font-size: 38px; font-weight: bold; margin-bottom: 8px;">
                                            {contadores['items_validados']}
                                        </div>
                                        <div style="color: #ffffff; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">
                                            Items<br>Validados
                                        </div>
                                    </td>
                                </tr>
                            </table>
                            
                            <!-- Métricas Adicionales -->
                            <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f8f9fa; border-radius: 8px; padding: 20px;">
                                <tr>
                                    <td width="50%" style="padding: 10px; border-right: 1px solid #dee2e6;">
                                        <div style="color: #6c757d; font-size: 13px; margin-bottom: 5px;">Items Procesados</div>
                                        <div style="color: #2c3e50; font-size: 24px; font-weight: bold;">{contadores['items_procesados']}</div>
                                    </td>
                                    <td width="50%" style="padding: 10px;">
                                        <div style="color: #6c757d; font-size: 13px; margin-bottom: 5px;">Items Verificar Manual</div>
                                        <div style="color: #e74c3c; font-size: 24px; font-weight: bold;">{contadores['items_verificar_manual']}</div>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="50%" style="padding: 10px; border-right: 1px solid #dee2e6; border-top: 1px solid #dee2e6;">
                                        <div style="color: #6c757d; font-size: 13px; margin-bottom: 5px;">Items Sin Texto</div>
                                        <div style="color: #f39c12; font-size: 24px; font-weight: bold;">{contadores['items_sin_texto']}</div>
                                    </td>
                                    <td width="50%" style="padding: 10px; border-top: 1px solid #dee2e6;">
                                        <div style="color: #6c757d; font-size: 13px; margin-bottom: 5px;">Notificaciones Enviadas</div>
                                        <div style="color: #27ae60; font-size: 24px; font-weight: bold;">{contadores['notificaciones_enviadas']}</div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- SOLPEDs con Problemas -->
                    <tr>
                        <td style="padding: 0 40px 35px 40px;">
                            <h2 style="margin: 0 0 20px 0; color: #2c3e50; font-size: 22px; border-bottom: 3px solid #e74c3c; padding-bottom: 10px;">
                                🚨 SOLPEDs que Requieren Atención
                            </h2>
                """

                # Lista de SOLPEDs con problemas
                for idx, solped_info in enumerate(solpeds_con_problemas, 1):
                    # Determinar color según severidad
                    if solped_info["items_revisar"] > solped_info["items_ok"]:
                        color_borde = "#e74c3c"  # Rojo - Más problemas
                        color_fondo = "#ffebee"
                    elif solped_info["items_sin_texto"] > 0:
                        color_borde = "#f39c12"  # Naranja - Advertencia
                        color_fondo = "#fff3e0"
                    else:
                        color_borde = "#3498db"  # Azul - Info
                        color_fondo = "#e3f2fd"

                    cuerpo_final += f"""
                            <!-- SOLPED {idx} -->
                            <div style="background-color: {color_fondo}; border-left: 4px solid {color_borde}; border-radius: 6px; padding: 20px; margin-bottom: 20px;">
                                <table width="100%" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td>
                                            <div style="display: flex; align-items: center; margin-bottom: 12px;">
                                                <span style="background-color: {color_borde}; color: #ffffff; padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; margin-right: 12px;">
                                                    #{idx}
                                                </span>
                                                <span style="color: #2c3e50; font-size: 20px; font-weight: bold;">
                                                    SOLPED {solped_info['solped']}
                                                </span>
                                            </div>
                                        </td>
                                        <td align="right">
                                            <span style="background-color: {color_borde}; color: #ffffff; padding: 8px 16px; border-radius: 20px; font-size: 13px; font-weight: 600;">
                                                {solped_info['estado']}
                                            </span>
                                        </td>
                                    </tr>
                                </table>
                                
                                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 15px;">
                                    <tr>
                                        <td width="25%" style="padding: 8px 0;">
                                            <div style="color: #7f8c8d; font-size: 12px; margin-bottom: 4px;">Total Items</div>
                                            <div style="color: #2c3e50; font-size: 18px; font-weight: bold;">{solped_info['items_total']}</div>
                                        </td>
                                        <td width="25%" style="padding: 8px 0;">
                                            <div style="color: #7f8c8d; font-size: 12px; margin-bottom: 4px;">✓ Validados</div>
                                            <div style="color: #27ae60; font-size: 18px; font-weight: bold;">{solped_info['items_ok']}</div>
                                        </td>
                                        <td width="25%" style="padding: 8px 0;">
                                            <div style="color: #7f8c8d; font-size: 12px; margin-bottom: 4px;">⚠ Revisar</div>
                                            <div style="color: #e74c3c; font-size: 18px; font-weight: bold;">{solped_info['items_revisar']}</div>
                                        </td>
                                        <td width="25%" style="padding: 8px 0;">
                                            <div style="color: #7f8c8d; font-size: 12px; margin-bottom: 4px;">Sin Texto</div>
                                            <div style="color: #f39c12; font-size: 18px; font-weight: bold;">{solped_info['items_sin_texto']}</div>
                                        </td>
                                    </tr>
                                </table>
                                
                                <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid rgba(0,0,0,0.1);">
                    """

                    if solped_info["responsables"]:
                        responsables_str = ", ".join(solped_info["responsables"])
                        if MODO_DESARROLLO:
                            cuerpo_final += f"""
                                    <div style="color: #856404; font-size: 13px;">
                                        <strong>⚠️ Responsable (no notificado - desarrollo):</strong><br>
                                        <span style="font-family: monospace; background-color: rgba(255,255,255,0.7); padding: 4px 8px; border-radius: 3px; display: inline-block; margin-top: 5px;">
                                            {responsables_str}
                                        </span>
                                    </div>
                            """
                        else:
                            cuerpo_final += f"""
                                    <div style="color: #27ae60; font-size: 13px;">
                                        <strong>✓ Notificado a:</strong> <span style="font-family: monospace;">{responsables_str}</span>
                                    </div>
                            """
                    else:
                        cuerpo_final += f"""
                                    <div style="color: #e74c3c; font-size: 13px;">
                                        <strong>⚠️ Sin responsable identificado</strong> - Se requiere asignación manual
                                    </div>
                        """

                    cuerpo_final += """
                                </div>
                            </div>
                    """

                cuerpo_final += f"""
                        </td>
                    </tr>
                    
                    <!-- Información del Adjunto -->
                    <tr>
                        <td style="padding: 0 40px 35px 40px;">
                            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; padding: 25px; text-align: center;">
                                <div style="font-size: 48px; margin-bottom: 10px;">📎</div>
                                <h3 style="margin: 0 0 10px 0; color: #ffffff; font-size: 18px;">Reporte Detallado Adjunto</h3>
                                <p style="margin: 0; color: #ffffff; font-size: 14px; opacity: 0.9;">
                                    El archivo adjunto contiene información detallada de cada item, validaciones y observaciones completas.
                                </p>
                            </div>
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #2c3e50; padding: 30px 40px; text-align: center;">
                            <p style="margin: 0 0 10px 0; color: #ecf0f1; font-size: 14px;">
                                <strong>Sistema de Validación Automática</strong>
                            </p>
                            <p style="margin: 0; color: #95a5a6; font-size: 12px;">
                                © {time.strftime('%Y')} Colsubsidio | NetApplications<br>
                                Este correo fue generado automáticamente
                            </p>
                        </td>
                    </tr>
                    
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
                """

                # Guardar reporte detallado en archivo
                path_reporte_final = f"{RUTAS['PathReportes']}\\Reporte_Consolidado_{time.strftime('%Y%m%d_%H%M%S')}.txt"
                with open(path_reporte_final, "w", encoding="utf-8") as f:
                    f.write(reporte_final)

                print(f"Reporte consolidado guardado: {path_reporte_final}")

                # Enviar correo
                exito_final = EnviarCorreoPersonalizado(
                    destinatario=(
                        EMAIL_DESARROLLO
                        if MODO_DESARROLLO
                        else "paula.sierra@netapplications.com.co"
                    ),
                    asunto=asunto_final,
                    cuerpo=cuerpo_final,
                    task_name=task_name,
                    adjuntos=[path_reporte_final],
                )

                if exito_final:
                    if MODO_DESARROLLO:
                        print(
                            f"✅ [DESARROLLO] Reporte consolidado enviado a {EMAIL_DESARROLLO}"
                        )
                    else:
                        print(f"✅ Reporte consolidado enviado a NetApplications")
                else:
                    print(f"❌ Error al enviar reporte consolidado")

            except Exception as e_final:
                print(f"❌ Error al generar/enviar reporte final: {e_final}")
                WriteLog(
                    mensaje=f"Error al enviar reporte consolidado: {e_final}",
                    estado="WARNING",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )

            print(f"{'='*80}\n")
        else:
            print(
                f"\n✅ Todas las SOLPEDs validadas correctamente - No se requiere reporte consolidado\n"
            )

        # ======================================================
        # === Finalización HU03 ===
        # ======================================================
        WriteLog(
            mensaje=f"HU03 completado exitosamente. "
            f"SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items validados: {contadores['items_validados']}/{contadores['items_procesados']}, "
            f"Notificaciones enviadas: {contadores['notificaciones_enviadas']}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
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

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
#   - UPDATE: WriteLog optimizado solo en puntos esenciales
# =========================================
import time
import traceback
from Funciones.ControlHU import control_hu
from Funciones.EmailSender import EnviarNotificacionCorreo
from Funciones.ReporteFinalME53N import (
    ConstruirFilaReporteFinal,
    GenerarReporteFinalExcel,
)
from Funciones.EscribirLog import WriteLog
from Funciones.GeneralME53N import (
    AbrirTransaccion,
    ColsultarSolped,
    TraerSAPAlFrenteOpcion,
    ActualizarEstado,
    ActualizarEstadoYObservaciones,
    NotificarRevisionManualSolped,
    GenerarReporteAttachments,
    ConvertirTxtAExcel,
    AppendHipervinculoObservaciones,
    obtenerFilaExpSolped,
)
from Funciones.SAPFuncionesME53N import (
    ProcesarTablaME5A,
    ObtenerItemTextME53N,
    ObtenerItemsME53N,
    GuardarTablaME5A,
    ValidarAttachmentList,
    ParsearTablaAttachments,
)

from Config.settings import RUTAS
from Funciones.FuncionesExcel import ExcelService
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


def EjecutarHU03(session, nombre_archivo):
    try:
        task_name = "HU03_ValidacionME53N"
        control_hu(task_name, estado=0)

        TraerSAPAlFrenteOpcion()

        WriteLog(
            mensaje="Inicio HU03 - Validación ME53N",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Leer el archivo con las SOLPEDs a procesar
        df_solpeds = ProcesarTablaME5A(nombre_archivo)
        GuardarTablaME5A(df_solpeds, nombre_archivo)

        if df_solpeds.empty:
            WriteLog(
                mensaje="El archivo expSolped03.txt está vacío o no se pudo cargar",
                estado="ERROR",
                task_name=task_name,
                path_log=RUTAS["PathLogError"],
            )
            return False

        # Validación de columnas
        columnas_requeridas = ["Estado", "Observaciones"]
        for columna in columnas_requeridas:
            if columna not in df_solpeds.columns:
                WriteLog(
                    mensaje=f"Columna requerida '{columna}' no encontrada",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                return False

        # Limpieza de SOLPEDs válidas
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
                if solped_limpia.isdigit():
                    solped_unicos_filtradas.append(solped_limpia)
                else:
                    solped_unicos_filtradas.append(solped_str)

        solped_unicos = solped_unicos_filtradas

        if not solped_unicos:
            WriteLog(
                mensaje="No se encontraron SOLPEDs válidas para procesar",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return False

        WriteLog(
            mensaje=f"Procesando {len(solped_unicos)} SOLPEDs - Total filas: {len(df_solpeds)}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

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

        # Modo desarrollo
        MODO_DESARROLLO = True
        EMAIL_DESARROLLO = "paula.sierra@netapplications.com.co"

        if MODO_DESARROLLO:
            WriteLog(
                mensaje=f"MODO DESARROLLO: Correos redirigidos a {EMAIL_DESARROLLO}",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        solpeds_con_problemas = []
        filas_reporte_final = []

        # PROCESAR CADA SOLPED
        for solped in solped_unicos:
            correos_responsables = []
            resumen_validaciones = []
            requiere_notificacion = False

            try:
                # Marcar SOLPED como "En Proceso"
                resultado_estado = ActualizarEstado(
                    df_solpeds, nombre_archivo, solped, nuevo_estado="En Proceso"
                )

                if not resultado_estado:
                    continue

                # Consultar SOLPED en SAP
                resultado_consulta = ColsultarSolped(session, solped)
                if not resultado_consulta:
                    WriteLog(
                        mensaje=f"No se pudo consultar SOLPED {solped} en SAP",
                        estado="ERROR",
                        task_name=task_name,
                        path_log=RUTAS["PathLogError"],
                    )
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

                # VALIDAR ATTACHMENT LIST
                tiene_attachments, contenido_attachments, obs_attachments = (
                    ValidarAttachmentList(session, solped)
                )

                attachments_lista = (
                    ParsearTablaAttachments(contenido_attachments)
                    if contenido_attachments
                    else []
                )

                reporte_attachments = GenerarReporteAttachments(
                    solped, tiene_attachments, contenido_attachments, obs_attachments
                )

                # Guardar reporte de attachments SOLO si tiene adjuntos
                if attachments_lista:
                    path_reporte_attach = (
                        f"{RUTAS['PathReportes']}\\Attachments_{solped}.txt"
                    )
                    try:
                        with open(path_reporte_attach, "w", encoding="utf-8") as f:
                            f.write(reporte_attachments)
                    except Exception as e:
                        pass
                else:
                    ActualizarEstadoYObservaciones(
                        df_solpeds,
                        nombre_archivo,
                        solped,
                        nuevo_estado="Sin Adjuntos",
                        observaciones="No cuenta con lista de Adjuntos",
                    )

                # MARCAR SI NO TIENE ATTACHMENTS
                solped_rechazada_por_attachments = False

                if not attachments_lista:
                    contadores["rechazadas_sin_attachments"] += 1
                    solped_rechazada_por_attachments = True
                    requiere_notificacion = True

                    resumen_validaciones.append(
                        f"\nMOTIVO DE RECHAZO PRINCIPAL\n"
                        f"   No cuenta con Attachment List\n"
                        f"   Acción requerida: Adjuntar documentación soporte\n"
                        f"   {obs_attachments}\n"
                    )
                else:
                    info_attachments = (
                        f"\n📎 ATTACHMENT LIST ({len(attachments_lista)} archivo(s))\n"
                    )
                    info_attachments += f"   {obs_attachments}\n"

                    if attachments_lista:
                        info_attachments += f"\n   Archivos adjuntos:\n"
                        for i, attach in enumerate(attachments_lista[:5], 1):
                            info_attachments += f"   {i}. {attach['title'][:50]}\n"
                            info_attachments += f"      Creado por: {attach['creator']} - {attach['date']}\n"

                        if len(attachments_lista) > 5:
                            info_attachments += f"   ... y {len(attachments_lista) - 5} archivo(s) más\n"

                    resumen_validaciones.append(info_attachments)

                # Obtener items de esta SOLPED
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
                    continue

                # Convertir a lista de diccionarios y filtrar totales
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

                # PROCESAR CADA ITEM
                contador_con_texto = 0
                contador_validados = 0
                contador_verificar_manual = 0
                items_procesados_en_solped = len(lista_dicts)

                for i, fila in enumerate(lista_dicts):
                    numero_item = fila.get("Pos.", str(i)).strip()
                    contadores["items_procesados"] += 1

                    # Obtener datos de expSolped03.txt
                    fila_exp = obtenerFilaExpSolped(df_solpeds, solped, numero_item)
                    if not fila_exp:
                        fila_exp = {}

                    # Obtener datos específicos de ME53N
                    fila_me53n = fila

                    if dtItems is not None and not dtItems.empty:
                        try:
                            mascara = (
                                dtItems["Pos."].astype(str).str.strip()
                                == str(numero_item).strip()
                            )
                            filas_encontradas = dtItems[mascara]

                            if not filas_encontradas.empty:
                                fila_me53n = filas_encontradas.iloc[0].to_dict()
                        except Exception as e:
                            pass

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
                            dtItems,
                            tiene_attachments,
                            obs_attachments,
                            attachments_lista,
                        )

                        # CAPTURAR CORREOS DE COLSUBSIDIO
                        responsable = datos_texto.get("responsable_compra", "")
                        if responsable and "@colsubsidio.com" in responsable.lower():
                            correos_encontrados = [
                                email.strip()
                                for email in responsable.split(",")
                                if "@colsubsidio.com" in email.lower()
                            ]
                            correos_responsables.extend(correos_encontrados)

                        # Guardar reporte detallado
                        path_reporte = f"{RUTAS['PathReportes']}\\Reporte_{solped}_{numero_item}.txt"
                        try:
                            with open(path_reporte, "w", encoding="utf-8") as f:
                                f.write(reporte)
                        except Exception as e:
                            pass

                        # Actualizar estado y observaciones
                        ActualizarEstadoYObservaciones(
                            df_solpeds,
                            nombre_archivo,
                            solped,
                            numero_item,
                            estado_final,
                            observaciones,
                        )

                        # FILTRO CRÍTICO: evitar fila TOTAL
                        if (
                            not numero_item
                            or not str(numero_item).strip().isdigit()
                            or str(numero_item).strip() in ["", "0"]
                        ):
                            continue

                        # CONSTRUIR FILA PARA REPORTE FINAL
                        fila_reporte = ConstruirFilaReporteFinal(
                            solped=solped,
                            item=numero_item,
                            datos_exp=fila_exp,
                            datos_adjuntos={
                                "cantidad": len(attachments_lista),
                                "nombres": ", ".join(
                                    [a["title"] for a in attachments_lista]
                                ),
                            },
                            datos_me53n=fila_me53n,
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

                        # CONSTRUIR RESUMEN PARA NOTIFICACIÓN
                        if estado_final != "Aprobado":
                            requiere_notificacion = True

                            item_info = f"\n📋 ITEM {numero_item}\n"
                            item_info += f"Estado: {estado_final}\n"
                            item_info += f"Observaciones: {observaciones}\n\n"

                            faltantes_me53n = validaciones.get(
                                "campos_obligatorios", {}
                            ).get("faltantes", [])
                            if faltantes_me53n:
                                item_info += (
                                    f"- ME53N faltantes: {', '.join(faltantes_me53n)}\n"
                                )
                            else:
                                item_info += "- ME53N faltantes: Ninguno\n"

                            faltantes_texto = validaciones.get("faltantes_texto", [])
                            if faltantes_texto:
                                item_info += (
                                    f"- Texto faltantes: {', '.join(faltantes_texto)}\n"
                                )
                            else:
                                item_info += "- Texto faltantes: Ninguno\n"

                            def estado_ok(flag):
                                return "OK" if flag else "ERROR"

                            item_info += "\nValidaciones:\n"
                            item_info += f"  Cantidad: {estado_ok(validaciones.get('cantidad', {}).get('match', False))}\n"
                            item_info += f"  Valor Unitario: {estado_ok(validaciones.get('valor_unitario', {}).get('match', False))}\n"
                            item_info += f"  Valor Total: {estado_ok(validaciones.get('valor_total', {}).get('match', False))}\n"
                            item_info += f"  Concepto: {estado_ok(validaciones.get('concepto', {}).get('match', False))}\n"

                            resumen_validaciones.append(item_info)

                        # Contar según resultado
                        if estado_final == "Aprobado":
                            contador_validados += 1
                            contadores["items_validados"] += 1
                        else:
                            contador_verificar_manual += 1
                            contadores["items_verificar_manual"] += 1

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

                        requiere_notificacion = True
                        resumen_validaciones.append(
                            f"\n📋 ITEM {numero_item}:\n"
                            f"   Estado: Sin Texto\n"
                            f"   Observaciones: {observaciones_item}\n"
                        )

                # ESTADO FINAL DE LA SOLPED
                if solped_rechazada_por_attachments:
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

                # ENVIAR NOTIFICACIÓN SI ES NECESARIO
                if requiere_notificacion and correos_responsables:
                    correos_unicos = list(set(correos_responsables))

                    if MODO_DESARROLLO:
                        correos_originales = correos_unicos.copy()
                        correos_unicos = [EMAIL_DESARROLLO]

                    # Construir texto completo de validaciones
                    texto_validaciones = f"SOLPED: {solped}\n"

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

                    try:
                        exito_notificacion = NotificarRevisionManualSolped(
                            destinatarios=correos_unicos,
                            numero_solped=solped,
                            validaciones=texto_validaciones,
                            task_name=task_name,
                        )

                        if exito_notificacion:
                            contadores["notificaciones_enviadas"] += 1

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
                            contadores["notificaciones_fallidas"] += 1

                    except Exception as e_notif:
                        contadores["notificaciones_fallidas"] += 1
                        WriteLog(
                            mensaje=f"Error al enviar notificación para SOLPED {solped}: {e_notif}",
                            estado="WARNING",
                            task_name=task_name,
                            path_log=RUTAS["PathLog"],
                        )

                elif requiere_notificacion and not correos_responsables:
                    WriteLog(
                        mensaje=f"SOLPED {solped}: Requiere revisión pero sin correo de responsable",
                        estado="WARNING",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )

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
                WriteLog(
                    mensaje=f"Error procesando SOLPED {solped}: {e}",
                    estado="ERROR",
                    task_name=task_name,
                    path_log=RUTAS["PathLogError"],
                )
                continue

        # Resumen final del proceso
        WriteLog(
            mensaje=f"PROCESO COMPLETADO - SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items validados: {contadores['items_validados']}/{contadores['items_procesados']}, "
            f"Notificaciones: {contadores['notificaciones_enviadas']}, "
            f"Rechazadas sin attachments: {contadores['rechazadas_sin_attachments']}, "
            f"Filas reporte: {len(filas_reporte_final)}",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # GENERAR ARCHIVO FINAL
        if filas_reporte_final:
            WriteLog(
                mensaje="Generando reporte final consolidado ME53N",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

            path_reporte = GenerarReporteFinalExcel(filas_reporte_final)

            if path_reporte:
                WriteLog(
                    mensaje=f"Reporte final generado: {path_reporte}",
                    estado="OK",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )
            else:
                WriteLog(
                    mensaje="No se pudo generar el reporte final",
                    estado="WARNING",
                    task_name=task_name,
                    path_log=RUTAS["PathLog"],
                )
        else:
            WriteLog(
                mensaje="No hay filas para generar el reporte final",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Convertir a Excel y agregar hipervínculos
        ConvertirTxtAExcel(nombre_archivo)
        archivo_descargado = rf"{RUTAS['PathInsumos']}/expSolped03.xlsx"
        AppendHipervinculoObservaciones(
            ruta_excel=archivo_descargado, carpeta_reportes=RUTAS["PathReportes"]
        )

        # Sube el Excel a la base de datos
        ExcelService.ejecutar_bulk_desde_excel(rf"{path_reporte}")

        # Enviar correo de finalización
        EnviarNotificacionCorreo(
            codigo_correo=3, task_name=task_name, adjuntos=[path_reporte]
        )

        control_hu(task_name, estado=100)
        return True

    except Exception as e:
        control_hu(task_name, estado=99)
        WriteLog(
            mensaje=f"Error en EjecutarHU03: {e}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return False

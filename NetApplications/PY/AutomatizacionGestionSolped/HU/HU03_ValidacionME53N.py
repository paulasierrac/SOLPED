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
from Funciones.ControlHU import ControlHU
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
from Funciones.FuncionesExcel import ServicioExcel
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


def EjecutarHU03(session, nombreArchivo):
    try:
        nombreTarea = "HU03_ValidacionME53N"
        ControlHU(nombreTarea, estado=0)

        TraerSAPAlFrenteOpcion()

        WriteLog(
            mensaje="Inicio HU03 - Validación ME53N",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=RUTAS["PathLog"],
        )

        # Leer el archivo con las SOLPEDs a procesar
        dfSolpeds = ProcesarTablaME5A(nombreArchivo)
        GuardarTablaME5A(dfSolpeds, nombreArchivo)

        if dfSolpeds.empty:
            WriteLog(
                mensaje="El archivo expSolped03.txt está vacío o no se pudo cargar",
                estado="ERROR",
                nombreTarea=nombreTarea,
                rutaRegistro=RUTAS["PathLogError"],
            )
            return False

        # Validación de columnas
        columnasRequeridas = ["Estado", "Observaciones"]
        for columna in columnasRequeridas:
            if columna not in dfSolpeds.columns:
                WriteLog(
                    mensaje=f"Columna requerida '{columna}' no encontrada",
                    estado="ERROR",
                    nombreTarea=nombreTarea,
                    rutaRegistro=RUTAS["PathLogError"],
                )
                return False

        # Limpieza de SOLPEDs válidas
        solpedUnicos = dfSolpeds["PurchReq"].unique().tolist()

        solpedUnicosFiltradas = []
        for solped in solpedUnicos:
            solpedStr = str(solped).strip()

            if (
                solpedStr
                and solpedStr not in ["Purch.Req.", "PurchReq", "Purch.Req", ""]
                and not any(
                    header in solpedStr for header in ["Purch.Req", "PurchReq"]
                )
                and solpedStr.replace(".", "").isdigit()
            ):
                solpedLimpia = solpedStr.replace(".", "")
                if solpedLimpia.isdigit():
                    solpedUnicosFiltradas.append(solpedLimpia)
                else:
                    solpedUnicosFiltradas.append(solpedStr)

        solpedUnicos = solpedUnicosFiltradas

        if not solpedUnicos:
            WriteLog(
                mensaje="No se encontraron SOLPEDs válidas para procesar",
                estado="WARNING",
                nombreTarea=nombreTarea,
                rutaRegistro=RUTAS["PathLog"],
            )
            return False

        WriteLog(
            mensaje=f"Procesando {len(solpedUnicos)} SOLPEDs - Total filas: {len(dfSolpeds)}",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=RUTAS["PathLog"],
        )

        # Abrir transaccion ME53N en SAP
        AbrirTransaccion(session, "ME53N")

        # Contadores para resumen final
        contadores = {
            "total_solpeds": len(solpedUnicos),
            "procesadas_exitosamente": 0,
            "con_errores": 0,
            "sin_items": 0,
            "items_procesados": 0,
            "items_validados": 0,
            "items_sin_texto": 0,
            "items_verificar_manual": 0,
            "notificaciones_enviadas": 0,
            "notificaciones_fallidas": 0,
            "rechazadas_sin_archivosAdjuntos": 0,
        }

        # Modo desarrollo
        MODO_DESARROLLO = True
        EMAIL_DESARROLLO = "paula.sierra@netapplications.com.co"

        if MODO_DESARROLLO:
            WriteLog(
                mensaje=f"MODO DESARROLLO: Correos redirigidos a {EMAIL_DESARROLLO}",
                estado="WARNING",
                nombreTarea=nombreTarea,
                rutaRegistro=RUTAS["PathLog"],
            )

        solpedsConProblemas = []
        filasReporteFinal = []

        # PROCESAR CADA SOLPED
        for solped in solpedUnicos:
            correosResponsables = []
            resumenValidaciones = []
            requiereNotificacion = False

            try:
                # Marcar SOLPED como "En Proceso"
                resultadoEstado = ActualizarEstado(
                    dfSolpeds, nombreArchivo, solped, nuevoEstado="En Proceso"
                )

                if not resultadoEstado:
                    continue

                # Consultar SOLPED en SAP
                resultadoConsulta = ColsultarSolped(session, solped)
                if not resultadoConsulta:
                    WriteLog(
                        mensaje=f"No se pudo consultar SOLPED {solped} en SAP",
                        estado="ERROR",
                        nombreTarea=nombreTarea,
                        rutaRegistro=RUTAS["PathLogError"],
                    )
                    ActualizarEstadoYObservaciones(
                        dfSolpeds,
                        nombreArchivo,
                        solped,
                        nuevoEstado="Error Consulta",
                        observaciones="No se pudo consultar en SAP",
                    )
                    contadores["con_errores"] += 1
                    continue

                time.sleep(0.5)

                # VALIDAR ATTACHMENT LIST
                tieneAttachments, contenidoAttachments, obsAttachments = (
                    ValidarAttachmentList(session, solped)
                )

                archivosAdjuntosLista = (
                    ParsearTablaAttachments(contenidoAttachments)
                    if contenidoAttachments
                    else []
                )

                reporteAttachments = GenerarReporteAttachments(
                    solped, tieneAttachments, contenidoAttachments, obsAttachments
                )

                # Guardar reporte de archivosAdjuntos SOLO si tiene adjuntos
                if archivosAdjuntosLista:
                    pathReporteAttach = (
                        f"{RUTAS['PathReportes']}\\Attachments_{solped}.txt"
                    )
                    try:
                        with open(pathReporteAttach, "w", encoding="utf-8") as f:
                            f.write(reporteAttachments)
                    except Exception as e:
                        pass
                else:
                    ActualizarEstadoYObservaciones(
                        dfSolpeds,
                        nombreArchivo,
                        solped,
                        nuevoEstado="Sin Adjuntos",
                        observaciones="No cuenta con lista de Adjuntos",
                    )

                # MARCAR SI NO TIENE ATTACHMENTS
                solpedRechazadaPorAttachments = False

                if not archivosAdjuntosLista:
                    contadores["rechazadas_sin_archivosAdjuntos"] += 1
                    solpedRechazadaPorAttachments = True
                    requiereNotificacion = True

                    resumenValidaciones.append(
                        f"\nMOTIVO DE RECHAZO PRINCIPAL\n"
                        f"   No cuenta con Attachment List\n"
                        f"   Acción requerida: Adjuntar documentación soporte\n"
                        f"   {obsAttachments}\n"
                    )
                else:
                    infoAttachments = (
                        f"\n📎 ATTACHMENT LIST ({len(archivosAdjuntosLista)} archivo(s))\n"
                    )
                    infoAttachments += f"   {obsAttachments}\n"

                    if archivosAdjuntosLista:
                        infoAttachments += f"\n   Archivos adjuntos:\n"
                        for i, attach in enumerate(archivosAdjuntosLista[:5], 1):
                            infoAttachments += f"   {i}. {attach['title'][:50]}\n"
                            infoAttachments += f"      Creado por: {attach['creator']} - {attach['date']}\n"

                        if len(archivosAdjuntosLista) > 5:
                            infoAttachments += f"   ... y {len(archivosAdjuntosLista) - 5} archivo(s) más\n"

                    resumenValidaciones.append(infoAttachments)

                # Obtener items de esta SOLPED
                dtItems = ObtenerItemsME53N(session, solped)

                if dtItems is None or dtItems.empty:
                    contadores["sin_items"] += 1
                    ActualizarEstadoYObservaciones(
                        dfSolpeds,
                        nombreArchivo,
                        solped,
                        nuevoEstado="Sin Items",
                        observaciones="No se encontraron items en SAP",
                    )
                    continue

                # Convertir a lista de diccionarios y filtrar totales
                listaDicts = dtItems.to_dict(orient="records")

                # Filtrar: Eliminar la ultima fila si es un total
                if listaDicts:
                    ultimaFila = listaDicts[-1]
                    if (
                        ultimaFila.get("Status", "").strip() == "*"
                        or ultimaFila.get("Item", "").strip() == ""
                        or ultimaFila.get("Material", "").strip() == ""
                    ):
                        listaDicts.pop()

                # PROCESAR CADA ITEM
                contadorConTexto = 0
                contadorValidados = 0
                contadorVerificarManual = 0
                itemsProcesadosEnSolped = len(listaDicts)

                for i, fila in enumerate(listaDicts):
                    numeroItem = fila.get("Pos.", str(i)).strip()
                    contadores["items_procesados"] += 1

                    # Obtener datos de expSolped03.txt
                    filaExp = obtenerFilaExpSolped(dfSolpeds, solped, numeroItem)
                    if not filaExp:
                        filaExp = {}

                    # Obtener datos específicos de ME53N
                    filaMe53n = fila

                    if dtItems is not None and not dtItems.empty:
                        try:
                            mascara = (
                                dtItems["Pos."].astype(str).str.strip()
                                == str(numeroItem).strip()
                            )
                            filasEncontradas = dtItems[mascara]

                            if not filasEncontradas.empty:
                                filaMe53n = filasEncontradas.iloc[0].to_dict()
                        except Exception as e:
                            pass

                    # Marcar item como "Procesando"
                    ActualizarEstado(
                        dfSolpeds, nombreArchivo, solped, numeroItem, "Procesando"
                    )

                    time.sleep(0.5)

                    # Obtener texto del editor SAP
                    texto = ObtenerItemTextME53N(session, solped, numeroItem)

                    # Procesar y validar el texto
                    if texto and texto.strip():
                        contadorConTexto += 1

                        # VALIDACION COMPLETA DEL TEXTO
                        (
                            datosTexto,
                            validaciones,
                            reporte,
                            estadoFinal,
                            observaciones,
                        ) = ProcesarYValidarItem(
                            session,
                            solped,
                            numeroItem,
                            texto,
                            dtItems,
                            tieneAttachments,
                            obsAttachments,
                            archivosAdjuntosLista,
                        )

                        # CAPTURAR CORREOS DE COLSUBSIDIO
                        responsable = datosTexto.get("responsable_compra", "")
                        if responsable and "@colsubsidio.com" in responsable.lower():
                            correosEncontrados = [
                                email.strip()
                                for email in responsable.split(",")
                                if "@colsubsidio.com" in email.lower()
                            ]
                            correosResponsables.extend(correosEncontrados)

                        # Guardar reporte detallado
                        pathReporte = f"{RUTAS['PathReportes']}\\Reporte_{solped}_{numeroItem}.txt"
                        try:
                            with open(pathReporte, "w", encoding="utf-8") as f:
                                f.write(reporte)
                        except Exception as e:
                            pass

                        # Actualizar estado y observaciones
                        ActualizarEstadoYObservaciones(
                            dfSolpeds,
                            nombreArchivo,
                            solped,
                            numeroItem,
                            estadoFinal,
                            observaciones,
                        )

                        # FILTRO CRÍTICO: evitar fila TOTAL
                        if (
                            not numeroItem
                            or not str(numeroItem).strip().isdigit()
                            or str(numeroItem).strip() in ["", "0"]
                        ):
                            continue

                        # CONSTRUIR FILA PARA REPORTE FINAL
                        filaReporte = ConstruirFilaReporteFinal(
                            solped=solped,
                            item=numeroItem,
                            datos_exp=filaExp,
                            datosAdjuntos={
                                "cantidad": len(archivosAdjuntosLista),
                                "nombres": ", ".join(
                                    [a["title"] for a in archivosAdjuntosLista]
                                ),
                            },
                            datosMe53n=filaMe53n,
                            datosTexto=datosTexto,
                            resultadoValidaciones={
                                "faltantesMe53n": validaciones.get(
                                    "campos_obligatorios", {}
                                ).get("faltantes"),
                                "faltantesTexto": validaciones.get("faltantesTexto"),
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
                                "estado": estadoFinal,
                                "observaciones": observaciones,
                            },
                        )

                        filasReporteFinal.append(filaReporte)

                        # CONSTRUIR RESUMEN PARA NOTIFICACIÓN
                        if estadoFinal != "Aprobado":
                            requiereNotificacion = True

                            itemInfo = f"\n📋 ITEM {numeroItem}\n"
                            itemInfo += f"Estado: {estadoFinal}\n"
                            itemInfo += f"Observaciones: {observaciones}\n\n"

                            faltantesMe53n = validaciones.get(
                                "campos_obligatorios", {}
                            ).get("faltantes", [])
                            if faltantesMe53n:
                                itemInfo += (
                                    f"- ME53N faltantes: {', '.join(faltantesMe53n)}\n"
                                )
                            else:
                                itemInfo += "- ME53N faltantes: Ninguno\n"

                            faltantesTexto = validaciones.get("faltantesTexto", [])
                            if faltantesTexto:
                                itemInfo += (
                                    f"- Texto faltantes: {', '.join(faltantesTexto)}\n"
                                )
                            else:
                                itemInfo += "- Texto faltantes: Ninguno\n"

                            def estadoOk(flag):
                                return "OK" if flag else "ERROR"

                            itemInfo += "\nValidaciones:\n"
                            itemInfo += f"  Cantidad: {estadoOk(validaciones.get('cantidad', {}).get('match', False))}\n"
                            itemInfo += f"  Valor Unitario: {estadoOk(validaciones.get('valor_unitario', {}).get('match', False))}\n"
                            itemInfo += f"  Valor Total: {estadoOk(validaciones.get('valor_total', {}).get('match', False))}\n"
                            itemInfo += f"  Concepto: {estadoOk(validaciones.get('concepto', {}).get('match', False))}\n"

                            resumenValidaciones.append(itemInfo)

                        # Contar según resultado
                        if estadoFinal == "Aprobado":
                            contadorValidados += 1
                            contadores["items_validados"] += 1
                        else:
                            contadorVerificarManual += 1
                            contadores["items_verificar_manual"] += 1

                    else:
                        # Sin texto en el editor
                        contadores["items_sin_texto"] += 1
                        observacionesItem = (
                            "Texto no encontrado en el editor SAP - No se puede validar"
                        )
                        ActualizarEstadoYObservaciones(
                            dfSolpeds,
                            nombreArchivo,
                            solped,
                            numeroItem,
                            "Sin Texto",
                            observacionesItem,
                        )

                        requiereNotificacion = True
                        resumenValidaciones.append(
                            f"\n📋 ITEM {numeroItem}:\n"
                            f"   Estado: Sin Texto\n"
                            f"   Observaciones: {observacionesItem}\n"
                        )

                # ESTADO FINAL DE LA SOLPED
                if solpedRechazadaPorAttachments:
                    estadoFinalSolped = "Rechazada"
                    observacionesSolped = (
                        f"RECHAZADA por falta de adjuntos - "
                        f"Items: {contadorValidados} validados, "
                        f"{contadorVerificarManual} requieren revisión, "
                        f"{itemsProcesadosEnSolped - contadorConTexto} sin texto"
                    )

                elif contadorValidados == itemsProcesadosEnSolped:
                    estadoFinalSolped = "Aprobado"
                    observacionesSolped = f"Todos validados ({contadorValidados} de {itemsProcesadosEnSolped}) + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1
                    requiereNotificacion = False

                elif contadorVerificarManual > 0:
                    estadoFinalSolped = "Pendiente"
                    observacionesSolped = f"{contadorVerificarManual} de {itemsProcesadosEnSolped} items requieren revisión + Contiene Adjuntos"
                    contadores["procesadas_exitosamente"] += 1

                else:
                    estadoFinalSolped = "Rechazada"
                    observacionesSolped = "No se pudo procesar correctamente"
                    contadores["con_errores"] += 1

                ActualizarEstadoYObservaciones(
                    dfSolpeds,
                    nombreArchivo,
                    solped,
                    nuevoEstado=estadoFinalSolped,
                    observaciones=observacionesSolped,
                )

                # ENVIAR NOTIFICACIÓN SI ES NECESARIO
                if requiereNotificacion and correosResponsables:
                    correosUnicos = list(set(correosResponsables))

                    if MODO_DESARROLLO:
                        correosOriginales = correosUnicos.copy()
                        correosUnicos = [EMAIL_DESARROLLO]

                    # Construir texto completo de validaciones
                    textoValidaciones = f"SOLPED: {solped}\n"

                    if MODO_DESARROLLO:
                        textoValidaciones += f"\nMODO DESARROLLO - CORREO DE PRUEBA\n"
                        textoValidaciones += f"Destinatarios originales: {', '.join(correosOriginales)}\n"
                        textoValidaciones += f"{'='*60}\n\n"

                    textoValidaciones += f"Estado Final: {estadoFinalSolped}\n"
                    textoValidaciones += f"Total Items: {itemsProcesadosEnSolped}\n"
                    textoValidaciones += f"Items Validados: {contadorValidados}\n"
                    textoValidaciones += (
                        f"Items Requieren Revisión: {contadorVerificarManual}\n"
                    )
                    textoValidaciones += f"Items Sin Texto: {itemsProcesadosEnSolped - contadorConTexto}\n"
                    textoValidaciones += f"\n{'='*60}\n"
                    textoValidaciones += f"DETALLE POR ITEM:\n"
                    textoValidaciones += "".join(resumenValidaciones)

                    try:
                        exitoNotificacion = NotificarRevisionManualSolped(
                            destinatarios=correosUnicos,
                            numeroSolped=solped,
                            validaciones=textoValidaciones,
                            nombreTarea=nombreTarea,
                        )

                        if exitoNotificacion:
                            contadores["notificaciones_enviadas"] += 1

                            solpedsConProblemas.append(
                                {
                                    "solped": solped,
                                    "estado": estadoFinalSolped,
                                    "tieneAttachments": tieneAttachments,
                                    "obsAttachments": obsAttachments,
                                    "archivosAdjuntos_detalle": (
                                        archivosAdjuntosLista[:10]
                                        if archivosAdjuntosLista
                                        else []
                                    ),
                                    "items_total": itemsProcesadosEnSolped,
                                    "items_ok": contadorValidados,
                                    "items_revisar": contadorVerificarManual,
                                    "items_sin_texto": itemsProcesadosEnSolped
                                    - contadorConTexto,
                                    "responsables": (
                                        correosOriginales
                                        if MODO_DESARROLLO
                                        else correosUnicos
                                    ),
                                    "detalle": resumenValidaciones,
                                }
                            )
                        else:
                            contadores["notificaciones_fallidas"] += 1

                    except Exception as e_notif:
                        contadores["notificaciones_fallidas"] += 1
                        WriteLog(
                            mensaje=f"Error al enviar notificación para SOLPED {solped}: {e_notif}",
                            estado="WARNING",
                            nombreTarea=nombreTarea,
                            rutaRegistro=RUTAS["PathLog"],
                        )

                elif requiereNotificacion and not correosResponsables:
                    WriteLog(
                        mensaje=f"SOLPED {solped}: Requiere revisión pero sin correo de responsable",
                        estado="WARNING",
                        nombreTarea=nombreTarea,
                        rutaRegistro=RUTAS["PathLog"],
                    )

                    solpedsConProblemas.append(
                        {
                            "solped": solped,
                            "estado": estadoFinalSolped,
                            "items_total": itemsProcesadosEnSolped,
                            "items_ok": contadorValidados,
                            "items_revisar": contadorVerificarManual,
                            "items_sin_texto": itemsProcesadosEnSolped
                            - contadorConTexto,
                            "responsables": [],
                            "detalle": resumenValidaciones,
                        }
                    )

            except Exception as e:
                contadores["con_errores"] += 1
                observacionesError = f"Error durante procesamiento: {str(e)[:100]}"
                ActualizarEstadoYObservaciones(
                    dfSolpeds,
                    nombreArchivo,
                    solped,
                    nuevoEstado="Error",
                    observaciones=observacionesError,
                )
                WriteLog(
                    mensaje=f"Error procesando SOLPED {solped}: {e}",
                    estado="ERROR",
                    nombreTarea=nombreTarea,
                    rutaRegistro=RUTAS["PathLogError"],
                )
                continue

        # Resumen final del proceso
        WriteLog(
            mensaje=f"PROCESO COMPLETADO - SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items validados: {contadores['items_validados']}/{contadores['items_procesados']}, "
            f"Notificaciones: {contadores['notificaciones_enviadas']}, "
            f"Rechazadas sin archivosAdjuntos: {contadores['rechazadas_sin_archivosAdjuntos']}, "
            f"Filas reporte: {len(filasReporteFinal)}",
            estado="INFO",
            nombreTarea=nombreTarea,
            rutaRegistro=RUTAS["PathLog"],
        )

        # GENERAR ARCHIVO FINAL
        if filasReporteFinal:
            WriteLog(
                mensaje="Generando reporte final consolidado ME53N",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=RUTAS["PathLog"],
            )

            pathReporte = GenerarReporteFinalExcel(filasReporteFinal)

            if pathReporte:
                WriteLog(
                    mensaje=f"Reporte final generado: {pathReporte}",
                    estado="OK",
                    nombreTarea=nombreTarea,
                    rutaRegistro=RUTAS["PathLog"],
                )
            else:
                WriteLog(
                    mensaje="No se pudo generar el reporte final",
                    estado="WARNING",
                    nombreTarea=nombreTarea,
                    rutaRegistro=RUTAS["PathLog"],
                )
        else:
            WriteLog(
                mensaje="No hay filas para generar el reporte final",
                estado="WARNING",
                nombreTarea=nombreTarea,
                rutaRegistro=RUTAS["PathLog"],
            )

        # Convertir a Excel y agregar hipervínculos
        ConvertirTxtAExcel(nombreArchivo)
        archivoDescargado = rf"{RUTAS['PathInsumos']}/expSolped03.xlsx"
        AppendHipervinculoObservaciones(
            rutaExcel=archivoDescargado, carpetaReportes=RUTAS["PathReportes"]
        )

        # Sube el Excel a la base de datos
        ServicioExcel.ejecutarBulkDesdeExcel(rf"{pathReporte}")

        # Enviar correo de finalización
        EnviarNotificacionCorreo(
            codigoCorreo=3, nombreTarea=nombreTarea, adjuntos=[pathReporte]
        )

        ControlHU(nombreTarea, estado=100)
        return True

    except Exception as e:
        ControlHU(nombreTarea, estado=99)
        WriteLog(
            mensaje=f"Error en EjecutarHU03: {e}",
            estado="ERROR",
            nombreTarea=nombreTarea,
            rutaRegistro=RUTAS["PathLogError"],
        )
        traceback.print_exc()
        return False

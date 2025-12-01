# =========================================
# NombreDeLaIniciativa – HU03: ValidacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta la búsqueda de una SOLPED en la transacción ME53N
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión con validación completa y uso correcto de validaciones
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
)
from Config.settings import RUTAS


def EjecutarHU03(session, nombre_archivo):
    try:
        task_name = "HU03_ValidacionME53N"

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
        }

        # Procesar cada SOLPED
        for solped in solped_unicos:
            print(f"\n{'='*80}")
            print(f"PROCESANDO SOLPED: {solped}")
            print(f"{'='*80}")

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

                # 5. Procesar cada item individualmente
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
                    print(texto)
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
                            session, solped, numero_item, texto, dtItems
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
                                estado = (
                                    "COINCIDE"
                                    if validaciones[campo]["match"]
                                    else "NO COINCIDE"
                                )
                                print(f"    {campo}: {estado}")
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
                        observaciones = (
                            "Texto no encontrado en el editor SAP - No se puede validar"
                        )
                        ActualizarEstadoYObservaciones(
                            df_solpeds,
                            nombre_archivo,
                            solped,
                            numero_item,
                            "Sin Texto",
                            observaciones,
                        )
                        print(f"Item {numero_item}: Sin texto - No validado")

                # 6. Actualizar estado final de la SOLPED completa
                if (
                    contador_validados == items_procesados_en_solped
                    and items_procesados_en_solped > 0
                ):
                    # Todos los items validados exitosamente
                    estado_final_solped = "Registro validado para orden de compra"
                    observaciones_solped = f"Todos los items validados ({contador_validados}/{items_procesados_en_solped})"
                    contadores["procesadas_exitosamente"] += 1

                elif contador_validados > 0 and contador_verificar_manual == 0:
                    # Algunos validados, otros sin texto
                    estado_final_solped = "Parcialmente validado"
                    observaciones_solped = f"{contador_validados}/{items_procesados_en_solped} items validados, {items_procesados_en_solped - contador_validados} sin texto"
                    contadores["procesadas_exitosamente"] += 1

                elif contador_verificar_manual > 0:
                    # Items que requieren verificacion manual
                    estado_final_solped = "Verificar manualmente"
                    observaciones_solped = f"{contador_verificar_manual}/{items_procesados_en_solped} items requieren verificacion manual"
                    contadores["procesadas_exitosamente"] += 1

                else:
                    # Ningun item procesado correctamente
                    estado_final_solped = "Sin procesar"
                    observaciones_solped = "No se pudo procesar ningun item"
                    contadores["con_errores"] += 1

                ActualizarEstadoYObservaciones(
                    df_solpeds,
                    nombre_archivo,
                    solped,
                    nuevo_estado=estado_final_solped,
                    observaciones=observaciones_solped,
                )

                print(f"\nEXITO: SOLPED {solped} completada")
                print(f"  Estado final: {estado_final_solped}")
                print(f"  Resumen: {observaciones_solped}")

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

        # ======================================================
        # === Finalización HU03 ===
        # ======================================================
        WriteLog(
            mensaje=f"HU03 completado exitosamente. "
            f"SOLPEDs: {contadores['procesadas_exitosamente']}/{contadores['total_solpeds']}, "
            f"Items validados: {contadores['items_validados']}/{contadores['items_procesados']}",
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

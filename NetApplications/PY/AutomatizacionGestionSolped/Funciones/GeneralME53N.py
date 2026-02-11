# ============================================
# Función Local: GeneralME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Archivo Base funciones necesarias transaccion ME53N
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Correcciones en ObtenerItemTextME53N y campos concepto_compra
# ============================================
import traceback
import win32com.client
import time
import os
from Funciones.EscribirLog import WriteLog
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


# Configurar encoding para consola de Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def EliminarArchivoSiExiste(ruta_archivo):
    try:
        if os.path.exists(ruta_archivo):
            WriteLog(f"Eliminando archivo: {ruta_archivo}", "INFO")
            os.remove(ruta_archivo)
            WriteLog(f"Archivo eliminado correctamente: {ruta_archivo}", "INFO")
        else:
            WriteLog(f"No existe archivo para eliminar: {ruta_archivo}", "INFO")
    except Exception as e:
        WriteLog(f"Error al eliminar archivo {ruta_archivo} | Error: {str(e)}", "ERROR")


def ConvertirTxtAExcel(archivo):
    """
    Convierte un archivo TXT delimitado por pipes (|) a Excel.

    Parámetros:
    -----------
    ruta_archivo_txt : str
        Ruta completa del archivo TXT a convertir

    Retorna:
    --------
    str : Ruta del archivo Excel generado

    Ejemplo:
    --------
    >>> ConvertirTxtAExcel('datos.txt')
    'datos.xlsx'
    """

    try:
        print(f"Leyendo archivo: {archivo}")

        ruta_archivo_txt = rf"{RUTAS["PathInsumos"]}\{archivo}"
        # Leer el archivo
        with open(ruta_archivo_txt, "r", encoding="utf-8") as f:
            lineas = f.readlines()

        # Filtrar líneas que contienen datos (excluir líneas de separadores)
        lineas_validas = []
        for linea in lineas:
            linea_limpia = linea.strip()
            # Verificar que tenga pipes y no sea solo guiones
            if (
                "|" in linea_limpia
                and not linea_limpia.replace("-", "").replace("|", "").strip() == ""
            ):
                lineas_validas.append(linea_limpia)

        if len(lineas_validas) < 2:
            raise ValueError("El archivo no contiene suficientes datos")

        print(f"Lineas validas encontradas: {len(lineas_validas)}")

        # Procesar encabezados (primera línea válida)
        # NO filtrar campos vacíos, mantener todas las posiciones
        encabezados_raw = lineas_validas[0].split("|")
        # Eliminar solo el primer y último elemento si están vacíos (bordes del pipe)
        if encabezados_raw and encabezados_raw[0].strip() == "":
            encabezados_raw = encabezados_raw[1:]
        if encabezados_raw and encabezados_raw[-1].strip() == "":
            encabezados_raw = encabezados_raw[:-1]
        encabezados = [campo.strip() for campo in encabezados_raw]

        print(f"\nColumnas encontradas: {len(encabezados)}")
        for i, col in enumerate(encabezados, 1):
            print(f"  {i}. {col}")

        # Procesar datos (resto de líneas)
        datos_procesados = []
        for i, linea in enumerate(lineas_validas[1:], start=2):
            campos_raw = linea.split("|")
            # Eliminar solo el primer y último elemento si están vacíos (bordes del pipe)
            if campos_raw and campos_raw[0].strip() == "":
                campos_raw = campos_raw[1:]
            if campos_raw and campos_raw[-1].strip() == "":
                campos_raw = campos_raw[:-1]
            # Mantener TODAS las posiciones, incluso las vacías
            campos = [campo.strip() for campo in campos_raw]

            # Asegurar que tenga el mismo número de columnas
            if len(campos) != len(encabezados):
                print(
                    f"  Advertencia fila {i}: {len(campos)} columnas (esperadas: {len(encabezados)})"
                )
                # Ajustar el tamaño
                if len(campos) < len(encabezados):
                    campos.extend([""] * (len(encabezados) - len(campos)))
                else:
                    campos = campos[: len(encabezados)]

            datos_procesados.append(campos)

        # Crear DataFrame
        df = pd.DataFrame(datos_procesados, columns=encabezados)

        print(f"\nDataFrame creado: {len(df)} filas x {len(df.columns)} columnas")

        # Generar nombre del archivo Excel
        ruta_excel = ruta_archivo_txt.rsplit(".", 1)[0] + ".xlsx"

        # Guardar a Excel con formato
        print(f"\nGuardando archivo Excel...")
        with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")

            # Ajustar ancho de columnas
            worksheet = writer.sheets["Datos"]
            for idx, col in enumerate(df.columns):
                # Calcular ancho máximo
                max_length = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                max_length = min(max_length, 60)

                # Calcular letra de columna
                if idx < 26:
                    col_letter = chr(65 + idx)
                else:
                    col_letter = chr(64 + idx // 26) + chr(65 + idx % 26)

                worksheet.column_dimensions[col_letter].width = max_length

        print(f"\n[OK] Archivo convertido exitosamente!")
        print(f"Ubicacion: {ruta_excel}")
        return ruta_excel

    except FileNotFoundError:
        print(f"[ERROR] No se encontro el archivo '{ruta_archivo_txt}'")
        raise
    except Exception as e:
        print(f"[ERROR] Error al convertir el archivo: {str(e)}")
        import traceback

        traceback.print_exc()
        raise


def GenerarReporteAttachments(
    solped: str, tiene_attachments: bool, contenido: str, observaciones: str
) -> str:
    """
    Genera un reporte formateado de la validación de attachments

    Args:
        solped: Número de SOLPED
        tiene_attachments: Si tiene attachments o no
        contenido: Contenido de la tabla exportada
        observaciones: Observaciones de la validación

    Returns:
        str: Reporte formateado
    """
    reporte = f"\n{'='*80}\n"
    reporte += f"VALIDACIÓN ATTACHMENT LIST - SOLPED: {solped}\n"
    reporte += f"{'='*80}\n\n"

    reporte += f"Estado: {'CON ADJUNTOS' if tiene_attachments else 'SIN ADJUNTOS'}\n"
    reporte += f"Observaciones: {observaciones}\n\n"

    attachments = ParsearTablaAttachments(contenido)

    if attachments:
        # Parsear tabla de attachments

        if attachments:
            reporte += f"ARCHIVOS ADJUNTOS ENCONTRADOS ({len(attachments)}):\n"
            reporte += f"{'-'*80}\n"

            for i, attach in enumerate(attachments, 1):
                reporte += f"\n{i}. Archivo: {attach['title']}\n"
                reporte += f"   Creado por: {attach['creator']}\n"
                reporte += f"   Fecha: {attach['date']}\n"

            reporte += f"\n{'-'*80}\n"

        # Agregar tabla completa como referencia
        reporte += f"\nTABLA COMPLETA EXPORTADA DE SAP:\n"
        reporte += f"{'-'*80}\n"
        reporte += contenido
        reporte += f"\n{'-'*80}\n"

    else:
        reporte += f"⚠️ SOLPED RECHAZADA: No cuenta con archivos adjuntos\n"
        reporte += (
            f"Acción requerida: Adjuntar documentación soporte antes de continuar\n"
        )

    reporte += f"\n{'='*80}\n"
    return reporte


def NotificarRevisionManualSolped(
    destinatarios: Union[str, List[str]],
    numero_solped: Union[int, str],
    validaciones: str,
    task_name: str = "RevisionManualSolped",
) -> bool:
    """
    Envía una notificación de revisión manual para un SOLPED específico,
    formateando automáticamente el asunto y el cuerpo.

    Args:
        destinatarios: Un email (str) o una lista de emails (List[str]).
        numero_solped: El número de la solicitud de pedido (SOLPED).
        validaciones: Texto que contiene las razones de la validación.
        task_name: Nombre de la tarea para los logs.

    Returns:
        bool: True si el envío fue exitoso, False en caso contrario.
    """

    # 1. Preparar el Asunto
    asunto_template = f"El Solped {numero_solped} Necesita revisión manual"

    # 2. Preparar el Cuerpo del Mensaje (Formato HTML)
    cuerpo_template = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #CC0000;">Solicitud de Revisión Manual Requerida</h2>
                <p>El Solped <strong>{numero_solped}</strong> necesita ser validado por las siguientes razones:</p>
                
                <div style="border: 1px solid #ddd; padding: 15px; margin: 15px 0; background-color: #f9f9f9;">
                    <div style="padding: 10px; margin: 10px 0; background-color: #f4f4f4; border-radius: 6px;">
                        {convertir_validaciones_a_lista(validaciones)}
                    </div>
                </div>

                <p>Por favor, ingrese al sistema para realizar las correcciones o ajustes necesarios.</p>
                <br>
                <p>Atentamente,<br>Sistema de Notificaciones</p>
            </body>
        </html>
    """

    # Asegurar que destinatarios sea una lista si viene como string
    if isinstance(destinatarios, str):
        destinatario_principal = destinatarios
        cc_list = None
    else:
        # Usamos el primer elemento como destinatario principal y el resto como CC (o podrías ajustar esta lógica)
        if destinatarios:
            destinatario_principal = destinatarios[0]
            cc_list = destinatarios[1:] if len(destinatarios) > 1 else None
        else:
            # Manejar el caso de lista vacía si fuera necesario
            print("Error: La lista de destinatarios está vacía.")
            return False

    # 3. Llamar a la función de envío personalizada
    return EnviarCorreoPersonalizado(
        destinatario=destinatario_principal,
        asunto=asunto_template,
        cuerpo=cuerpo_template,
        task_name=task_name,
        cc=cc_list,
        adjuntos=None,  # No se esperan adjuntos para esta notificación
    )


def EnviarCorreoPersonalizado(
    destinatario: str,
    asunto: str,
    cuerpo: str,
    task_name: str = "EnvioPersonalizado",
    adjuntos: list = None,
    cc: list = None,
    bcc: list = None,
) -> bool:
    """
    Envía un correo electrónico con estructura personalizada, sin usar el Excel de correos.

    Args:
        destinatario: Email del destinatario (cadena de texto).
        asunto: Asunto del correo (cadena de texto).
        cuerpo: Cuerpo del mensaje (puede ser HTML).
        task_name: Nombre de la tarea para logs.
        adjuntos: Lista de rutas de archivos a adjuntar (opcional).
        cc: Lista de correos en copia (opcional).
        bcc: Lista de correos en copia oculta (opcional).

    Returns:
        bool: True si se envió correctamente, False en caso contrario.
    """
    try:
        WriteLog(
            mensaje=f"Preparando envío personalizado para {destinatario}...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Log de adjuntos
        if adjuntos:
            WriteLog(
                mensaje=f"Adjuntos a enviar: {', '.join(adjuntos)}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Crear EmailSender con configuración por defecto
        sender = EmailSender()

        # Llamar al método de envío directo de la clase EmailSender
        exito = sender.enviar_correo(
            destinatario=destinatario,
            asunto=asunto,
            cuerpo=cuerpo,
            cc=cc,
            bcc=bcc,
            adjuntos=adjuntos,
        )

        if exito:
            WriteLog(
                mensaje=f"Correo personalizado enviado exitosamente a {destinatario}.",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return True
        else:
            WriteLog(
                mensaje=f"Fallo al enviar el correo personalizado a {destinatario}.",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return False

    except Exception as e:
        error_stack = traceback.format_exc()
        WriteLog(
            mensaje=f"Error fatal en el envío personalizado: {e} | {error_stack}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        return False


def TraerSAPAlFrenteOpcion():
    """Usar Alt+Tab para traer SAP al frente"""
    try:
        pyautogui.hotkey("alt", "tab")
        time.sleep(0.5)
        print("SAP traido al frente (Opcion - Alt+Tab)")
    except Exception as e:
        print(f"Error en Opcion 4: {e}")


def convertir_validaciones_a_lista(texto):
    """
    Convierte el bloque de texto de validaciones en una lista HTML <ul><li>.
    Cada item debe comenzar con '📋 ITEM' u otro marcador detectable.
    """
    lineas = [l.strip() for l in texto.split("\n") if l.strip()]

    lista_html = "<ul style='font-size:14px; line-height:1.5;'>"

    for linea in lineas:
        # Detectar inicio de item
        if linea.startswith("-ITEM"):
            lista_html += f"<li><strong>{linea}</strong></li>"
        else:
            lista_html += f"<li>{linea}</li>"

    lista_html += "</ul>"

    return lista_html


def ObtenerTextoDelPortapapeles():
    """Obtener texto del portapapeles con manejo correcto de codificacion"""
    try:
        # Abrir portapapeles
        win32clipboard.OpenClipboard()
        try:
            # Obtener texto con CF_UNICODETEXT (maneja mejor caracteres especiales)
            texto = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            return texto if texto else ""
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Error al leer portapapeles: {e}")
        return ""


def AbrirTransaccion(session, transaccion):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida"""

    try:
        WriteLog(
            mensaje=f"Abrir Transaccion {transaccion}",
            estado="INFO",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="AbrirTransaccion",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # Abrir transaccion dinamica
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje=f"Transaccion {transaccion} abierta",
            estado="INFO",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLog"],
        )
        print(f"Transaccion {transaccion} abierta")
        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en AbrirTransaccion: {e}",
            estado="ERROR",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLogError"],
        )

        return False


def ColsultarSolped(session, numero_solped):
    """session: objeto de SAP GUI
    numero_solped:  numero de SOLPED a consultar
    Realiza la verificacion del SOLPED"""

    try:
        WriteLog(
            mensaje=f"Numero de SOLPED : {numero_solped}",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="ColsultarSolped",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # Boton de Otra consulta
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(0.3)
        # Escribir numero de solped
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN"
        ).text = numero_solped
        # Activar el radiobutton "Purch. Requisition"
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B"
        ).setFocus()
        # Seleccionar el radiobutton "Purch. Requisition"
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B"
        ).select()

        # Presionar el boton OK (btn[0])
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(3)

        WriteLog(
            mensaje=f"Solped {numero_solped} consultada exitosamente",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )

        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en ColsultarSolped: {e}",
            estado="ERROR",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLogError"],
        )

        return False


def ActualizarEstadoYObservaciones(
    df, nombre_archivo, purch_req, item=None, nuevo_estado="", observaciones=""
):
    """Actualiza el estado y observaciones en el DataFrame y guarda el archivo"""
    try:
        # ASEGURAR QUE EXISTE LA COLUMNA OBSERVACIONES
        if "Observaciones" not in df.columns:
            df["Observaciones"] = ""
            print("ADVERTENCIA: Columna 'Observaciones' creada en el DataFrame")

        # Crear mascara para filtrar
        if item is not None:
            # Actualizar item especifico
            mask = (df["PurchReq"] == str(purch_req)) & (
                df["Item"] == str(item).strip()
            )
        else:
            # Actualizar toda la SOLPED
            mask = df["PurchReq"] == str(purch_req)

        # Actualizar estado y observaciones
        if mask.sum() > 0:
            df.loc[mask, "Estado"] = nuevo_estado
            if observaciones:
                df.loc[mask, "Observaciones"] = observaciones
            # Guardar archivo actualizado
            GuardarTablaME5A(df, nombre_archivo)
            print(
                f"EXITO: Actualizado: {purch_req}" + (f" Item {item}" if item else "")
            )
            return True
        else:
            print(
                f"No se encontro PurchReq {purch_req}"
                + (f", Item {item}" if item else "")
            )
            return False

    except Exception as e:
        print(f"Error al actualizar estado y observaciones: {e}")
        return False


def ActualizarEstado(df, nombre_archivo, purch_req, item=None, nuevo_estado=""):
    """Actualiza el estado en el DataFrame y guarda el archivo"""
    try:
        # Crear mascara para filtrar
        if item is not None:
            # Actualizar item especifico
            mask = (df["PurchReq"] == str(purch_req)) & (
                df["Item"] == str(item).strip()
            )
        else:
            # Actualizar toda la SOLPED
            mask = df["PurchReq"] == str(purch_req)

        # Actualizar estado
        if mask.sum() > 0:
            df.loc[mask, "Estado"] = nuevo_estado
            # Guardar archivo actualizado
            GuardarTablaME5A(df, nombre_archivo)
            return True
        else:
            print(
                f"No se encontro PurchReq {purch_req}"
                + (f", Item {item}" if item else "")
            )
            return False

    except Exception as e:
        print(f"Error al actualizar estado: {e}")
        return False


def obtenerValorTabla(fila, posibles_nombres, default=0):
    """Intenta múltiples nombres de columna"""
    for nombre in posibles_nombres:
        if nombre in fila and fila[nombre] not in [None, "", " "]:
            return fila[nombre]
    return default


def FormatoMoneda(valor):
    """Convierte un número en formato moneda $xx,xxx.xx"""
    try:
        valor = float(valor)
        return f"${valor:,.2f}"
    except:
        return str(valor)


def limpiar_numero(valor):
    """
    Limpia y convierte un valor a número
    Maneja valores None, vacíos, y no numéricos
    """
    if valor is None or valor == "":
        return 0.0

    # Convertir a string
    valor_str = str(valor).strip()

    # Si es vacío después de strip
    if not valor_str:
        return 0.0

    # ============================================
    # NUEVO: Validar que es numérico
    # ============================================
    # Si contiene solo letras (como 'K', 'D'), retornar 0
    if valor_str.isalpha():
        return 0.0

    # Si es muy corto y no es numérico, retornar 0
    if len(valor_str) <= 2 and not any(c.isdigit() for c in valor_str):
        return 0.0
    # ============================================

    try:
        # Remover símbolos de moneda y separadores
        valor_limpio = valor_str.replace("$", "").replace("COP", "").replace("USD", "")
        valor_limpio = valor_limpio.replace(".", "").replace(",", ".")
        valor_limpio = valor_limpio.strip()

        # Convertir a float
        return float(valor_limpio)
    except (ValueError, AttributeError) as e:
        print(f"ERROR limpiando numero '{valor}': {e}")
        return 0.0

# ============================================
# Función Local: GeneralME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Archivo Base funciones necesarias transaccion ME53N
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Correcciones en ObtenerItemTextME53N y campos concepto_compra
# ============================================
import traceback
import time
import os
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from Config.InicializarConfig import inConfig
import pandas as pd
import pyautogui
from datetime import datetime
from typing import Dict, List, Tuple
import os
from Funciones.EmailSender import EmailSender
from typing import List, Union
import sys
from openpyxl import load_workbook
from Funciones.SAPFuncionesME53N import GuardarTablaME5A, ParsearTablaAttachments
import win32clipboard


# Configurar encoding para consola de Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def EliminarArchivoSiExiste(rutaArchivo):
    try:
        if os.path.exists(rutaArchivo):
            WriteLog(f"Eliminando archivo: {rutaArchivo}", "INFO")
            os.remove(rutaArchivo)
            WriteLog(f"Archivo eliminado correctamente: {rutaArchivo}", "INFO")
        else:
            WriteLog(f"No existe archivo para eliminar: {rutaArchivo}", "INFO")
    except Exception as e:
        WriteLog(f"Error al eliminar archivo {rutaArchivo} | Error: {str(e)}", "ERROR")


def ConvertirTxtAExcel(archivo):
    """
    Convierte un archivo TXT delimitado por pipes (|) a Excel.

    Parámetros:
    -----------
    rutaArchivoTxt : str
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

        rutaArchivoTxt = rf"{RUTAS["PathInsumos"]}\{archivo}"
        # Leer el archivo
        with open(rutaArchivoTxt, "r", encoding="utf-8") as f:
            lineas = f.readlines()

        # Filtrar líneas que contienen datos (excluir líneas de separadores)
        lineasValidas = []
        for linea in lineas:
            lineaLimpia = linea.strip()
            # Verificar que tenga pipes y no sea solo guiones
            if (
                "|" in lineaLimpia
                and not lineaLimpia.replace("-", "").replace("|", "").strip() == ""
            ):
                lineasValidas.append(lineaLimpia)

        if len(lineasValidas) < 2:
            raise ValueError("El archivo no contiene suficientes datos")

        print(f"Lineas validas encontradas: {len(lineasValidas)}")

        # Procesar encabezados (primera línea válida)
        # NO filtrar campos vacíos, mantener todas las posiciones
        encabezadosFila = lineasValidas[0].split("|")
        # Eliminar solo el primer y último elemento si están vacíos (bordes del pipe)
        if encabezadosFila and encabezadosFila[0].strip() == "":
            encabezadosFila = encabezadosFila[1:]
        if encabezadosFila and encabezadosFila[-1].strip() == "":
            encabezadosFila = encabezadosFila[:-1]
        encabezados = [campo.strip() for campo in encabezadosFila]

        print(f"\nColumnas encontradas: {len(encabezados)}")
        for i, col in enumerate(encabezados, 1):
            print(f"  {i}. {col}")

        # Procesar datos (resto de líneas)
        datosProcesados = []
        for i, linea in enumerate(lineasValidas[1:], start=2):
            camposFila = linea.split("|")
            # Eliminar solo el primer y último elemento si están vacíos (bordes del pipe)
            if camposFila and camposFila[0].strip() == "":
                camposFila = camposFila[1:]
            if camposFila and camposFila[-1].strip() == "":
                camposFila = camposFila[:-1]
            # Mantener TODAS las posiciones, incluso las vacías
            campos = [campo.strip() for campo in camposFila]

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

            datosProcesados.append(campos)

        # Crear DataFrame
        df = pd.DataFrame(datosProcesados, columns=encabezados)

        print(f"\nDataFrame creado: {len(df)} filas x {len(df.columns)} columnas")

        # Generar nombre del archivo Excel
        rutaExcel = rutaArchivoTxt.rsplit(".", 1)[0] + ".xlsx"

        # Guardar a Excel con formato
        print(f"\nGuardando archivo Excel...")
        with pd.ExcelWriter(rutaExcel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")

            # Ajustar ancho de columnas
            worksheet = writer.sheets["Datos"]
            for idx, col in enumerate(df.columns):
                # Calcular ancho máximo
                longitudMax = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                longitudMax = min(longitudMax, 60)

                # Calcular letra de columna
                if idx < 26:
                    col_letter = chr(65 + idx)
                else:
                    col_letter = chr(64 + idx // 26) + chr(65 + idx % 26)

                worksheet.column_dimensions[col_letter].width = longitudMax

        print(f"\n[OK] Archivo convertido exitosamente!")
        print(f"Ubicacion: {rutaExcel}")
        return rutaExcel

    except FileNotFoundError:
        print(f"[ERROR] No se encontro el archivo '{rutaArchivoTxt}'")
        raise
    except Exception as e:
        print(f"[ERROR] Error al convertir el archivo: {str(e)}")
        import traceback

        traceback.print_exc()
        raise


def GenerarReporteAttachments(
    solped: str, tieneAttachments: bool, contenido: str, observaciones: str
) -> str:
    """
    Genera un reporte formateado de la validación de archivosAdjuntos

    Args:
        solped: Número de SOLPED
        tieneAttachments: Si tiene archivosAdjuntos o no
        contenido: Contenido de la tabla exportada
        observaciones: Observaciones de la validación

    Returns:
        str: Reporte formateado
    """
    reporte = f"\n{'='*80}\n"
    reporte += f"VALIDACIÓN ATTACHMENT LIST - SOLPED: {solped}\n"
    reporte += f"{'='*80}\n\n"

    reporte += f"Estado: {'CON ADJUNTOS' if tieneAttachments else 'SIN ADJUNTOS'}\n"
    reporte += f"Observaciones: {observaciones}\n\n"

    archivosAdjuntos = ParsearTablaAttachments(contenido)

    if archivosAdjuntos:
        # Parsear tabla de archivosAdjuntos

        if archivosAdjuntos:
            reporte += f"ARCHIVOS ADJUNTOS ENCONTRADOS ({len(archivosAdjuntos)}):\n"
            reporte += f"{'-'*80}\n"

            for i, attach in enumerate(archivosAdjuntos, 1):
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
    numeroSolped: Union[int, str],
    validaciones: str,
    nombreTarea: str = "RevisionManualSolped",
) -> bool:
    """
    Envía una notificación de revisión manual para un SOLPED específico,
    formateando automáticamente el asunto y el cuerpo.

    Args:
        destinatarios: Un email (str) o una lista de emails (List[str]).
        numeroSolped: El número de la solicitud de pedido (SOLPED).
        validaciones: Texto que contiene las razones de la validación.
        nombreTarea: Nombre de la tarea para los logs.

    Returns:
        bool: True si el envío fue exitoso, False en caso contrario.
    """

    # 1. Preparar el Asunto
    asuntoTemplate = f"El Solped {numeroSolped} Necesita revisión manual"

    # 2. Preparar el Cuerpo del Mensaje (Formato HTML)
    cuerpoTemplate = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #CC0000;">Solicitud de Revisión Manual Requerida</h2>
                <p>El Solped <strong>{numeroSolped}</strong> necesita ser validado por las siguientes razones:</p>
                
                <div style="border: 1px solid #ddd; padding: 15px; margin: 15px 0; background-color: #f9f9f9;">
                    <div style="padding: 10px; margin: 10px 0; background-color: #f4f4f4; border-radius: 6px;">
                        {convertirValidacionesALista(validaciones)}
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
        destinatarioPrincipal = destinatarios
        ccList = None
    else:
        # Usamos el primer elemento como destinatario principal y el resto como CC (o podrías ajustar esta lógica)
        if destinatarios:
            destinatarioPrincipal = destinatarios[0]
            ccList = destinatarios[1:] if len(destinatarios) > 1 else None
        else:
            # Manejar el caso de lista vacía si fuera necesario
            print("Error: La lista de destinatarios está vacía.")
            return False

    # 3. Llamar a la función de envío personalizada
    return EnviarCorreoPersonalizado(
        destinatario=destinatarioPrincipal,
        asunto=asuntoTemplate,
        cuerpo=cuerpoTemplate,
        nombreTarea=nombreTarea,
        cc=ccList,
        adjuntos=None,  # No se esperan adjuntos para esta notificación
    )


def EnviarCorreoPersonalizado(
    destinatario: str,
    asunto: str,
    cuerpo: str,
    nombreTarea: str = "EnvioPersonalizado",
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
        nombreTarea: Nombre de la tarea para logs.
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
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
        )

        # Log de adjuntos
        if adjuntos:
            WriteLog(
                mensaje=f"Adjuntos a enviar: {', '.join(adjuntos)}",
                estado="INFO",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
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
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
            return True
        else:
            WriteLog(
                mensaje=f"Fallo al enviar el correo personalizado a {destinatario}.",
                estado="WARNING",
                nombreTarea=nombreTarea,
                rutaRegistro=inConfig("PathLog"),
            )
            return False

    except Exception as e:
        error_stack = traceback.format_exc()
        WriteLog(
            mensaje=f"Error fatal en el envío personalizado: {e} | {error_stack}",
            estado="ERROR",
            nombreTarea=nombreTarea,
            rutaRegistro=inConfig("PathLog"),
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


def convertirValidacionesALista(texto):
    """
    Convierte el bloque de texto de validaciones en una lista HTML <ul><li>.
    Cada item debe comenzar con '📋 ITEM' u otro marcador detectable.
    """
    lineas = [l.strip() for l in texto.split("\n") if l.strip()]

    listaHtml = "<ul style='font-size:14px; line-height:1.5;'>"

    for linea in lineas:
        # Detectar inicio de item
        if linea.startswith("-ITEM"):
            listaHtml += f"<li><strong>{linea}</strong></li>"
        else:
            listaHtml += f"<li>{linea}</li>"

    listaHtml += "</ul>"

    return listaHtml


def ObtenerTextoDelPortapapeles():
    """Obtener texto del portapapeles con manejo correcto de encoding"""
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
    """
    session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida

    """

    try:
        WriteLog(
            mensaje=f"Abrir Transaccion {transaccion}",
            estado="INFO",
            nombreTarea="AbrirTransaccion",
            rutaRegistro=inConfig("PathLog"),
        )

        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                nombreTarea="AbrirTransaccion",
                rutaRegistro=inConfig("PathLog"),
            )
            raise Exception("Sesion SAP no disponible")

        # Abrir transaccion dinamica
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje=f"Transaccion {transaccion} abierta",
            estado="INFO",
            nombreTarea="AbrirTransaccion",
            rutaRegistro=inConfig("PathLog"),
        )
        print(f"Transaccion {transaccion} abierta")
        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en AbrirTransaccion: {e}",
            estado="ERROR",
            nombreTarea="AbrirTransaccion",
            rutaRegistro=inConfig("PathLog"),
        )

        return False


def ColsultarSolped(session, numeroSolped):
    """session: objeto de SAP GUI
    numeroSolped:  numero de SOLPED a consultar
    Realiza la verificacion del SOLPED"""

    try:
        WriteLog(
            mensaje=f"Numero de SOLPED : {numeroSolped}",
            estado="INFO",
            nombreTarea="ColsultarSolped",
            rutaRegistro=inConfig("PathLog"),
        )

        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                nombreTarea="ColsultarSolped",
                rutaRegistro=inConfig("PathLog"),
            )
            raise Exception("Sesion SAP no disponible")

        # Boton de Otra consulta
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(0.3)
        # Escribir numero de solped
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN"
        ).text = numeroSolped
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
            mensaje=f"Solped {numeroSolped} consultada exitosamente",
            estado="INFO",
            nombreTarea="ColsultarSolped",
            rutaRegistro=inConfig("PathLog"),
        )

        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en ColsultarSolped: {e}",
            estado="ERROR",
            nombreTarea="ColsultarSolped",
            rutaRegistro=inConfig("PathLog"),
        )

        return False


def ActualizarEstadoYObservaciones(
    df, nombreArchivo, purchReq, item=None, nuevoEstado="", observaciones=""
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
            mascara = (df["PurchReq"] == str(purchReq)) & (
                df["Item"] == str(item).strip()
            )
        else:
            # Actualizar toda la SOLPED
            mascara = df["PurchReq"] == str(purchReq)

        # Actualizar estado y observaciones
        if mascara.sum() > 0:
            df.loc[mascara, "Estado"] = nuevoEstado
            if observaciones:
                df.loc[mascara, "Observaciones"] = observaciones
            # Guardar archivo actualizado
            GuardarTablaME5A(df, nombreArchivo)
            print(f"EXITO: Actualizado: {purchReq}" + (f" Item {item}" if item else ""))
            return True
        else:
            print(
                f"No se encontro PurchReq {purchReq}"
                + (f", Item {item}" if item else "")
            )
            return False

    except Exception as e:
        print(f"Error al actualizar estado y observaciones: {e}")
        return False


def ActualizarEstado(df, nombreArchivo, purchReq, item=None, nuevoEstado=""):
    """Actualiza el estado en el DataFrame y guarda el archivo"""
    try:
        # Crear mascara para filtrar
        if item is not None:
            # Actualizar item especifico
            mascara = (df["PurchReq"] == str(purchReq)) & (
                df["Item"] == str(item).strip()
            )
        else:
            # Actualizar toda la SOLPED
            mascara = df["PurchReq"] == str(purchReq)

        # Actualizar estado
        if mascara.sum() > 0:
            df.loc[mascara, "Estado"] = nuevoEstado
            # Guardar archivo actualizado
            GuardarTablaME5A(df, nombreArchivo)
            return True
        else:
            print(
                f"No se encontro PurchReq {purchReq}"
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


def LimpiarNumero(valor):
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

def TransformartxtMe5a(ruta_txt: str):

    if not os.path.exists(ruta_txt):
        raise FileNotFoundError(f"No existe el archivo: {ruta_txt}")

    # ===============================
    # 1. LEER Y LIMPIAR ARCHIVO
    # ===============================
    with open(ruta_txt, "r", encoding="latin-1", errors="replace") as f:
        lineas = f.readlines()

    lineasValidas = [
        linea.strip()
        for linea in lineas
        if linea.strip().startswith("|")
        and "|" in linea
        and not set(linea.strip()) == {"-"}
    ]

    if not lineasValidas:
        raise ValueError("El archivo no contiene líneas válidas con formato esperado.")

    columnas = [col.strip() for col in lineasValidas[0].split("|")[1:-1]]

    datos = []
    for linea in lineasValidas[1:]:
        valores = [v.strip() for v in linea.split("|")[1:-1]]
        if len(valores) == len(columnas):
            datos.append(valores)

    df = pd.DataFrame(datos, columns=columnas)

    # ===============================
    # 2. ELIMINAR COLUMNAS NO NECESARIAS
    # ===============================
    columnasEliminar = ["CDoc", "EL", "ProvFijo", "PrecioVal.", "Valor tot.", "Fondo"]

    df = df.drop(
        columns=[c for c in columnasEliminar if c in df.columns], errors="ignore"
    )

    # ===============================
    # 3. RENOMBRAR COLUMNAS A INGLÉS
    # ===============================
    df = df.rename(
        columns={
            "Sol.pedido": "PurchReq",
            "Pos.": "Item",
            "Fe.solic.": "ReqDate",
            "Creado por": "Created",
            "Texto breve": "ShortText",
            "Pedido": "PO",
            "Cantidad": "Quantity",
            "Ce.": "Plnt",
            "GCp": "PGr",
            "Gestores": "Requisnr",
            "Stat.trat.": "ProcState",
        }
    )

    # ===============================
    # 4. AGREGAR COLUMNA D (VACÍA)
    # ===============================
    if "Blank1" not in df.columns:
        df["Blank1"] = ""

    if "D" not in df.columns:
        df["D"] = ""
    # ===============================
    # EXTRAER NÚMERO DESDE NOMBRE TXT
    # ===============================
    nombreArchivo = os.path.basename(ruta_txt)
    numeroProcstate = "".join(filter(str.isdigit, nombreArchivo))

    if not numeroProcstate:
        raise ValueError("No se pudo extraer número del nombre del archivo.")

    numeroProcstate = numeroProcstate.zfill(2)

    # ===============================
    # 5. AGREGAR ProcState
    # ===============================
    df["ProcState"] = numeroProcstate

    # ===============================
    # 6. ORDEN FINAL
    # ===============================
    ordenFinal = [
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



    df = df[[col for col in ordenFinal if col in df.columns]]

    # ===============================
    # 7. GUARDAR ARCHIVO TRANSFORMADO
    # ===============================
    rutaSalida = ruta_txt

    anchos = {
        col: max(df[col].astype(str).map(len).max(), len(col)) + 3
        for col in df.columns
    }

    with open(rutaSalida, "w", encoding="utf-8") as f:

        f.write("-" * sum(anchos.values()) + "\n")

        encabezado = "|"
        for col in df.columns:
            encabezado += f"{col:<{anchos[col]}}|"
        f.write(encabezado + "\n")

        f.write("-" * sum(anchos.values()) + "\n")

        for _, row in df.iterrows():
            linea = "|"
            for col in df.columns:
                linea += f"{str(row[col]):<{anchos[col]}}|"
            f.write(linea + "\n")

        f.write("-" * sum(anchos.values()) + "\n")

    print(f"Archivo transformado correctamente: {rutaSalida}")
    return rutaSalida


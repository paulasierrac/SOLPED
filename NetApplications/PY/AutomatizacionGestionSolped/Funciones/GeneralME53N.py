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


def EnviarNotificacionCorreo(
    codigo_correo: int, task_name: str = "Notificacion", adjuntos: list = None
):
    """
    Envía notificaciones por correo según el código especificado

    Args:
        codigo_correo: Código del correo a enviar (1=Inicio, 2=Éxito, 3=Error, etc.)
        task_name: Nombre de la tarea para logs
        adjuntos: Lista de rutas de archivos a adjuntar (opcional)

    Returns:
        bool: True si se envió correctamente, False en caso contrario
    """
    try:
        WriteLog(
            mensaje=f"Enviando notificación con código {codigo_correo}...",
            estado="INFO",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )

        # Log de adjuntos si existen
        if adjuntos:
            WriteLog(
                mensaje=f"Adjuntos a enviar: {', '.join(adjuntos)}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )

        # Crear EmailSender con configuración por defecto
        sender = EmailSender()

        # Enviar correo según código con adjuntos
        resultados = sender.procesar_excel_y_enviar(
            archivo_excel=RUTAS.get(
                "ArchivoCorreos",
                # r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\EnvioCorreos.xlsx",
                rf"{RUTAS["ArchivoCorreos"]}",
            ),
            codigo_correo=codigo_correo,
            columna_codigo="codemailparameter",
            columna_destinatario="toemailparameter",  # Nombre correcto
            columna_asunto="asuntoemailparameter",  # Nombre correcto
            columna_cuerpo="bodyemailparameter",  # Nombre correcto
            columna_cc="ccemailparameter",  # Nombre correcto
            columna_bcc="bccemailparameter",  # Nombre correcto
            adjuntos_dinamicos=adjuntos,
        )

        if resultados["exitosos"] > 0:
            WriteLog(
                mensaje=f"Notificación enviada correctamente. Exitosos: {resultados['exitosos']}",
                estado="INFO",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return True
        else:
            WriteLog(
                mensaje=f"No se pudo enviar la notificación. Fallidos: {resultados['fallidos']}",
                estado="WARNING",
                task_name=task_name,
                path_log=RUTAS["PathLog"],
            )
            return False

    except Exception as e:
        error_stack = traceback.format_exc()
        WriteLog(
            mensaje=f"Error al enviar notificación: {e} | {error_stack}",
            estado="ERROR",
            task_name=task_name,
            path_log=RUTAS["PathLogError"],
        )
        return False


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


def TraerSAPAlFrente_Opcion():
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


def procesarTablaME5A(name, dias=None):
    """name: nombre del txt a utilizar
    return data frame
    Procesa txt estructura ME5A y devuelve un df con manejo de columnas dinamico.
    dias: int|None -> número de días a mantener (si None, no aplica filtro por fecha)"""

    try:
        WriteLog(
            mensaje=f"Procesar archivo nombre {name}",
            estado="INFO",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLog"],
        )

        # path = f".\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{RUTAS["PathInsumos"]}\{name}"

        # INTENTAR LEER CON DIFERENTES CODIFICACIONES
        lineas = []
        codificaciones = ["latin-1", "cp1252", "iso-8859-1", "utf-8"]

        for codificacion in codificaciones:
            try:
                with open(path, "r", encoding=codificacion) as f:
                    lineas = f.readlines()
                print(f"EXITO: Archivo leido con codificacion {codificacion}")
                break
            except UnicodeDecodeError as e:
                print(f"ERROR con {codificacion}: {e}")
                continue
            except Exception as e:
                print(f"ERROR con {codificacion}: {e}")
                continue

        if not lineas:
            print("ERROR: No se pudo leer el archivo con ninguna codificacion")
            return pd.DataFrame()

        # Filtrar solo lineas de datos
        filas = [l for l in lineas if l.startswith("|") and not l.startswith("|---")]

        # DETECTAR ESTRUCTURA DE COLUMNAS DINAMICAMENTE
        if not filas:
            print("No se encontraron filas de datos en el archivo")
            return pd.DataFrame()

        # Analizar la primera fila para determinar estructura
        primera_fila = filas[0].strip().split("|")[1:-1]  # Quitar | inicial y final
        primera_fila = [p.strip() for p in primera_fila]

        num_columnas = len(primera_fila)
        print(f"Estructura detectada: {num_columnas} columnas")
        print(f"   Encabezados: {primera_fila}")

        # DEFINIR COLUMNAS BASE SEGUN ESTRUCTURA
        if num_columnas == 14:
            # Estructura original (sin Estado ni Observaciones)
            columnas_base = [
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
            columnas_extra = ["Estado", "Observaciones"]

        elif num_columnas == 15:
            # Verificar si la columna 15 es "Estado" o "Observaciones"
            ultima_columna = primera_fila[-1].lower()
            if "estado" in ultima_columna:
                # Estructura con Estado pero sin Observaciones
                columnas_base = [
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
                    "Estado",
                ]
                columnas_extra = ["Observaciones"]
            else:
                # Estructura con Observaciones pero sin Estado
                columnas_base = [
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
                    "Observaciones",
                ]
                columnas_extra = ["Estado"]

        elif num_columnas == 16:
            # Estructura completa con Estado y Observaciones
            columnas_base = [
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
                "Estado",
                "Observaciones",
            ]
            columnas_extra = []
        else:
            print(f"ERROR: Estructura no soportada: {num_columnas} columnas")
            return pd.DataFrame()

        # PROCESAR TODAS LAS FILAS
        filas_proc = []
        for i, fila in enumerate(filas):
            partes = fila.strip().split("|")[1:-1]
            partes = [p.strip() for p in partes]

            # Validar que tenga el numero correcto de columnas
            if len(partes) == num_columnas:
                filas_proc.append(partes)
            elif len(partes) == num_columnas + 1 and partes[-1] == "":
                # Caso: columna extra vacia al final
                filas_proc.append(partes[:num_columnas])
                if i < 3:  # Solo log primeras filas
                    print(f"   ADVERTENCIA Fila {i+1}: Columna extra vacia removida")
            else:
                print(
                    f"   ERROR Fila {i+1} ignorada: {len(partes)} columnas vs {num_columnas} esperadas"
                )
                if i == 0:  # Solo mostrar detalle para primera fila
                    print(f"      Contenido: {partes}")
                continue

        # CREAR DATAFRAME
        df = pd.DataFrame(filas_proc, columns=columnas_base)

        # AGREGAR COLUMNAS FALTANTES
        for col_extra in columnas_extra:
            if col_extra not in df.columns:
                df[col_extra] = ""
                print(f"EXITO: Columna '{col_extra}' agregada al DataFrame")

        # FILTRAR: Si la primera fila es encabezado, eliminarla
        primera_fila_es_encabezado = any(
            col in df.iloc[0].values if not df.empty else False
            for col in [
                "Purch.Req.",
                "Item",
                "Req.Date",
                "Short Text",
                "PurchReq",
                "Estado",
                "Observaciones",
            ]
        )

        if not df.empty and primera_fila_es_encabezado:
            df = df.iloc[1:].reset_index(drop=True)
            print("EXITO: Fila de encabezado removida")

        print(f"EXITO: Archivo procesado: {len(df)} filas de datos")
        print(f"   - Columnas: {list(df.columns)}")

        if not df.empty:
            print(f"   - SOLPEDs: {df['PurchReq'].nunique()}")
            if "Estado" in df.columns:
                print(f"   - Estados unicos: {df['Estado'].value_counts().to_dict()}")

        # Normalizar formato fecha
        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        if dias is not None:
            hoy = pd.Timestamp.today().normalize()
            limite = hoy - pd.Timedelta(days=int(dias))
            filas_antes = len(df)
            df = df[df["ReqDate_fmt"] >= limite].reset_index(drop=True)
            filas_despues = len(df)
            print(
                f"EXITO: Filtrado por ReqDate últimos {dias} días -> {filas_despues}/{filas_antes}"
            )
        else:
            print("INFO: No se aplicó filtro por ReqDate (dias=None)")

        # opcional: eliminar columna auxiliar
        df.drop(columns=["ReqDate_fmt"], inplace=True)

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en procesarTablaME5A: {e}",
            estado="ERROR",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR en procesarTablaME5A: {e}")
        traceback.print_exc()
        return pd.DataFrame()


def GuardarTablaME5A(df, name):
    """Guarda el DataFrame de vuelta al TXT con formato de tabla"""
    try:
        # path = f"C:\\Users\\CGRPA009\\Documents\\SOLPED-main\\SOLPED\\NetApplications\\PY\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{RUTAS["PathInsumos"]}\{name}"

        # ASEGURAR QUE TIENE LAS COLUMNAS NECESARIAS
        columnas_requeridas = ["Estado", "Observaciones"]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
                print(f"ADVERTENCIA: Columna '{col}' agregada para guardado")

        # Calcular anchos de columna basados en contenido
        anchos = {}
        for col in df.columns:
            max_contenido = df[col].astype(str).str.len().max() if not df.empty else 0
            anchos[col] = max(len(col), max_contenido) + 2

        # Crear linea separadora
        separador = "-" * (sum(anchos.values()) + len(df.columns) + 1)

        # Crear encabezado
        encabezado_partes = [str(col).ljust(anchos[col]) for col in df.columns]
        encabezado = "|" + "|".join(encabezado_partes) + "|"

        # Crear filas
        filas_txt = []
        for _, fila in df.iterrows():
            partes = []
            for col in df.columns:
                valor = str(fila[col])
                # Alinear a la derecha numeros, izquierda texto
                if (
                    col in ["Item", "Quantity"]
                    or valor.replace(".", "").replace("-", "").isdigit()
                ):
                    texto_valor = valor.rjust(anchos[col])
                else:
                    texto_valor = valor.ljust(anchos[col])
                partes.append(texto_valor)
            fila_txt = "|" + "|".join(partes) + "|"
            filas_txt.append(fila_txt)

        # Escribir archivo
        with open(path, "w", encoding="utf-8") as f:
            f.write(separador + "\n")
            f.write(encabezado + "\n")
            f.write(separador + "\n")
            for fila in filas_txt:
                f.write(fila + "\n")

        WriteLog(
            mensaje=f"Archivo {name} actualizado con exito - {len(df)} filas",
            estado="INFO",
            task_name="GuardarTablaME5A",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: Archivo guardado: {len(df)} filas, {len(df.columns)} columnas")
        return True

    except Exception as e:
        WriteLog(
            mensaje=f"Error al guardar {name}: {e}",
            estado="ERROR",
            task_name="GuardarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR al guardar archivo: {e}")
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


def DetectarCodificacion(path):
    """Detecta automaticamente la codificacion del archivo"""
    try:
        with open(path, "rb") as f:
            rawdata = f.read()

        resultado = chardet.detect(rawdata)
        encoding = resultado["encoding"]
        confidence = resultado["confidence"]

        print(f"Codificacion detectada: {encoding} (confianza: {confidence*100:.1f}%)")
        return encoding
    except Exception as e:
        print(f"Error detectando codificacion: {e}")
        return "utf-8"


def TablaItemsDataFrame(name) -> pd.DataFrame:
    """name: nombre del archivo a consultar
    Convierte tabla de items a df con deteccion automatica de codificacion"""

    try:
        WriteLog(
            mensaje=f"Nombre de archivo {name}",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )

        # path = rf"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\TablasME53N\{name}"
        path = rf"{RUTAS["PathInsumos"]}\TablasME53N\{name}"

        # ========== DETECCION DE CODIFICACION ==========
        encoding = DetectarCodificacion(path)

        # 1. Leer archivo con la codificacion correcta
        try:
            with open(path, "r", encoding=encoding) as f:
                texto = f.read()
        except Exception as e:
            # Si falla, intentar con otras codificaciones comunes
            print(f"Error con {encoding}, intentando alternativas...")
            for enc in ["latin-1", "cp1252", "iso-8859-1", "utf-8"]:
                try:
                    with open(path, "r", encoding=enc) as f:
                        texto = f.read()
                    print(f"EXITO con {enc}")
                    encoding = enc
                    break
                except:
                    continue

        # 2. Separar por lineas
        lineas = texto.splitlines()

        # 3. Filtrar lineas que forman parte de la tabla
        tabla = [
            l
            for l in lineas
            if l.strip().startswith("|") and l.strip().endswith("|") and "---" not in l
        ]

        if not tabla:
            raise ValueError("No se encontro ninguna tabla SAP dentro del archivo.")

        # 4. Eliminar lineas de guiones largos (separadores)
        tabla = [
            l for l in tabla if not re.match(r"^-{5,}", l.replace("|", "").strip())
        ]

        # 5. Extraer encabezado
        encabezado_raw = tabla[0]
        columnas = [c.strip() for c in encabezado_raw.split("|")[1:-1]]

        # ========== SOLUCIONAR COLUMNAS DUPLICADAS ==========
        columnas_unicas = []
        contador = {}
        for col in columnas:
            if col in contador:
                contador[col] += 1
                columnas_unicas.append(f"{col}_{contador[col]}")
            else:
                contador[col] = 0
                columnas_unicas.append(col)

        # 6. Procesar filas de datos
        filas = []
        for fila in tabla[1:]:
            partes = [c.strip() for c in fila.split("|")[1:-1]]
            if len(partes) == len(columnas_unicas):  # validar integridad
                filas.append(partes)

        # 7. Convertir a DataFrame
        df = pd.DataFrame(filas, columns=columnas_unicas)

        WriteLog(
            mensaje=f"DataFrame conversion correcta. Codificacion: {encoding}",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: DataFrame conversion correcta")
        print(f"  - Filas: {df.shape[0]}")
        print(f"  - Columnas: {df.shape[1]}")
        print(f"  - Codificacion: {encoding}")

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en TablaItemsDataFrame: {e}",
            estado="ERROR",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR: {e}")
        return pd.DataFrame()


def ObtenerItemsME53N(session, numero_solped):
    """session: objeto de SAP GUI
    numero_solped: numero de solicitud
    Obtiene los items de SOLPED y los pasa a un df"""

    try:
        WriteLog(
            mensaje=f"Solped {numero_solped} a obtener items",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:
            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemsME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # ========== EXPORTAR TABLA ==========
        grid = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        )
        grid.setFocus()
        time.sleep(0.5)

        # 1. Abrir menu contexto "Exportar"
        grid.pressToolbarContextButton("&MB_EXPORT")
        time.sleep(0.5)

        # 2. Seleccionar "Exportar → Hoja de calculo (PC)"
        grid.selectContextMenuItem("&PC")
        time.sleep(0.3)

        # 3. Confirmar ventana de exportar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.3)

        # 4. Escribir ruta de guardado
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = (
            # r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\TablasME53N"
            rf"{RUTAS["PathInsumos"]}\TablasME53N"
        )
        time.sleep(0.2)

        # 5. Nombre del archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = (
            f"TablaSolped{numero_solped}.txt"
        )
        time.sleep(0.2)

        # 6. Guardar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)  # Esperar a que se guarde

        # ========== CONVERTIR A DATAFRAME ==========
        df = TablaItemsDataFrame(f"TablaSolped{numero_solped}.txt")

        if df is None or df.empty:
            raise Exception("DataFrame vacio despues de conversion")

        WriteLog(
            mensaje=f"Solped {numero_solped} convertido a DF con exito",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )
        print(f"EXITO: Solped {numero_solped} convertido a DF con exito")

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemsME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR en ObtenerItemsME53N: {e}")
        return pd.DataFrame()


def ObtenerItemTextME53N(session, numero_solped, numero_item):
    """session: objeto de SAP GUI
    numero_solped: numero de SOLPED
    numero_item: numero del item actual
    Realiza la extraccion del texto del editor SAP"""

    try:
        WriteLog(
            mensaje=f"ObtenerItemTextME53N {numero_solped} Item {numero_item}",
            estado="INFO",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesion SAP
        if session is None:
            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemTextME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesion SAP no disponible")

        # ---------------- Capturar Texto----------------
        # 1) Obtener el objeto del editor
        editor = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        # 2) Asegurar que el editor tiene el foco
        editor.SetFocus()
        time.sleep(0.5)

        # 3) Seleccionar TODO el texto
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)

        # 4) Copiar al portapapeles
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)

        # 5) Obtener texto del portapapeles con codificacion correcta
        texto_completo = ObtenerTextoDelPortapapeles()

        # 6) Limpiar caracteres problematicos si los hay
        texto_limpio = texto_completo.encode("utf-8", errors="replace").decode("utf-8")

        identificador = f"\n=====Solped: {numero_solped} Item: {numero_item} Registro: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n"

        # 7. Guardar texto en archivo de log
        # path = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\texto_ITEMsap.txt"
        path = rf"{RUTAS["PathInsumos"]}\texto_ITEMsap.txt"
        with open(path, "a", encoding="utf-8") as f:
            f.write(identificador)
            f.write(texto_limpio + "\n")
            f.write("-" * 80 + "\n")

        # 8. Navegar al siguiente item
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/"
            "btn%#AUTOTEXT002"
        ).press()
        time.sleep(0.5)

        return texto_limpio

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemTextME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLogError"],
        )
        return ""


def FormatoMoneda(valor):
    """Convierte un número en formato moneda $xx,xxx.xx"""
    try:
        valor = float(valor)
        return f"${valor:,.2f}"
    except:
        return str(valor)


def ValidarContraTabla(
    datos_texto: Dict, df_items: pd.DataFrame, item_num: str = ""
) -> Dict:
    """Compara los datos extraidos del texto con la tabla de items SAP"""
    validaciones = {
        "cantidad": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "valor_unitario": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "valor_total": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        # "fecha_entrega": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "concepto": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "campos_obligatorios": {"presentes": 0, "total": 0, "faltantes": []},
        "resumen": "",
        "campos_validados": 0,
        "total_campos": 0,
    }

    if df_items.empty:
        validaciones["resumen"] = "Tabla vacia - No se puede validar"
        return validaciones

    # Buscar el item especifico en el DataFrame
    item_df = df_items
    if item_num and "Item" in df_items.columns:
        item_df = df_items[
            df_items["Item"].astype(str).str.strip() == str(item_num).strip()
        ]
        if item_df.empty:
            item_df = df_items.iloc[[0]]  # usar primera fila si no encuentra

    if item_df.empty:
        validaciones["resumen"] = "Item no encontrado en tabla - No se puede validar"
        return validaciones

    # ------------------------------------------------------------------
    # VALIDACIÓN DE CAMPOS OBLIGATORIOS DE LA TABLA SAP ME53N
    # ------------------------------------------------------------------
    campos_obligatorios_me53n = [
        "Material",
        "Short Text",
        "Quantity",
        "Un",
        "Valn Price",
        "Crcy",
        "Total Val.",
        "Deliv.Date",
        "Fix. Vend.",
        "Plant",
        "PGr",
        "POrg",
        "Matl Group",
    ]

    faltantes_me53n = []

    fila = item_df.iloc[0]

    for campo in campos_obligatorios_me53n:
        if campo not in item_df.columns:
            faltantes_me53n.append(campo)
        else:
            val = fila.get(campo, "")
            if val is None or str(val).strip() in ["", "nan", "None"]:
                faltantes_me53n.append(campo)

    validaciones["campos_me53n"] = {
        "presentes": len(campos_obligatorios_me53n) - len(faltantes_me53n),
        "total": len(campos_obligatorios_me53n),
        "faltantes": faltantes_me53n,
    }

    # --- VALIDAR CANTIDAD ---
    if datos_texto["cantidad"]:
        cantidad_texto = LimpiarNumero(datos_texto["cantidad"])
        if "Quantity" in item_df.columns:
            # CORREGIDO: Asegurar que obtenemos un valor escalar
            cantidad_tabla_val = item_df["Quantity"].iloc[0]
            cantidad_tabla = LimpiarNumero(str(cantidad_tabla_val))
            validaciones["cantidad"]["texto"] = datos_texto["cantidad"]
            validaciones["cantidad"]["tabla"] = str(cantidad_tabla)
            validaciones["cantidad"]["match"] = (
                abs(cantidad_texto - cantidad_tabla) < 0.01
            )
            if not validaciones["cantidad"]["match"]:
                validaciones["cantidad"][
                    "diferencia"
                ] = f"Difiere en {abs(cantidad_texto - cantidad_tabla):.2f}"

    # --- VALIDAR VALOR UNITARIO ---
    if datos_texto["valor_unitario"]:
        valor_texto = LimpiarNumero(datos_texto["valor_unitario"])
        if "Valn Price" in item_df.columns:
            valor_tabla = LimpiarNumero(str(item_df["Valn Price"].iloc[0]))
            # validaciones["valor_unitario"]["texto"] = datos_texto["valor_unitario"]
            # validaciones["valor_unitario"]["tabla"] = str(valor_tabla)

            validaciones["valor_unitario"]["texto"] = FormatoMoneda(
                datos_texto["valor_unitario"]
            )
            validaciones["valor_unitario"]["tabla"] = FormatoMoneda(valor_tabla)

            # Tolerancia del 1%
            if valor_tabla > 0:
                diferencia_relativa = abs(valor_texto - valor_tabla) / valor_tabla
                validaciones["valor_unitario"]["match"] = diferencia_relativa < 0.01
            else:
                validaciones["valor_unitario"]["match"] = valor_texto == valor_tabla

            if not validaciones["valor_unitario"]["match"]:
                diferencia = abs(valor_texto - valor_tabla)
                validaciones["valor_unitario"][
                    "diferencia"
                ] = f"Difiere en {FormatoMoneda(diferencia)}"
                # validaciones["valor_unitario"][
                #     "diferencia"
                # ] = f"Difiere en ${diferencia:,.2f}"

    # --- VALIDAR VALOR TOTAL ---
    if datos_texto["valor_total"]:
        valor_texto = LimpiarNumero(datos_texto["valor_total"])

        # ACEPTAR AMBOS CAMPOS DE SAP
        columna_total = None
        if "Total Value" in item_df.columns:
            columna_total = "Total Value"
        elif "Total Val." in item_df.columns:
            columna_total = "Total Val."

        if columna_total:
            valor_tabla = LimpiarNumero(str(item_df[columna_total].iloc[0]))

            # validaciones["valor_total"]["texto"] = datos_texto["valor_total"]
            # validaciones["valor_total"]["tabla"] = str(valor_tabla)

            validaciones["valor_total"]["texto"] = FormatoMoneda(
                datos_texto["valor_total"]
            )
            validaciones["valor_total"]["tabla"] = FormatoMoneda(valor_tabla)

            # Tolerancia del 1%
            if valor_tabla > 0:
                diferencia_relativa = abs(valor_texto - valor_tabla) / valor_tabla
                validaciones["valor_total"]["match"] = diferencia_relativa < 0.01
            else:
                validaciones["valor_total"]["match"] = valor_texto == valor_tabla

            if not validaciones["valor_total"]["match"]:
                diferencia = abs(valor_texto - valor_tabla)
                validaciones["valor_total"][
                    "diferencia"
                ] = f"Difiere en {FormatoMoneda(diferencia)}"

                # validaciones["valor_total"][
                #     "diferencia"
                # ] = f"Difiere en ${diferencia:,.2f}"
        else:
            # Si no existe ninguna columna válida
            validaciones["valor_total"]["tabla"] = "Campo no encontrado en SAP"
            validaciones["valor_total"]["match"] = False

    # # --- VALIDAR FECHA DE ENTREGA ---
    # if datos_texto["fecha_prestacion"] and "Deliv.Date" in item_df.columns:
    #     fecha_texto = datos_texto["fecha_prestacion"]
    #     fecha_tabla = str(item_df["Deliv.Date"].iloc[0]) if not item_df.empty else ""
    #     validaciones["fecha_entrega"]["texto"] = fecha_texto
    #     validaciones["fecha_entrega"]["tabla"] = fecha_tabla
    #     validaciones["fecha_entrega"]["match"] = NormalizarFecha(
    #         fecha_texto
    #     ) == NormalizarFecha(fecha_tabla)
    #     if not validaciones["fecha_entrega"]["match"]:
    #         validaciones["fecha_entrega"]["diferencia"] = "Fechas no coinciden"

    # --- VALIDAR CONCEPTO ---
    if datos_texto["concepto_compra"] and "Short Text" in item_df.columns:
        concepto_texto = datos_texto["concepto_compra"].upper()
        concepto_tabla_val = item_df["Short Text"].iloc[0] if not item_df.empty else ""
        concepto_tabla = str(concepto_tabla_val).upper()

        validaciones["concepto"]["texto"] = datos_texto["concepto_compra"][:50] + (
            "..." if len(datos_texto["concepto_compra"]) > 50 else ""
        )
        validaciones["concepto"]["tabla"] = concepto_tabla[:50] + (
            "..." if len(concepto_tabla) > 50 else ""
        )

        # Verificar coincidencia de palabras clave
        palabras_texto = set(re.findall(r"\w+", concepto_texto))
        palabras_tabla = set(re.findall(r"\w+", concepto_tabla))
        coincidencias = len(palabras_texto & palabras_tabla)

        # CORREGIDO: Mejorar logica de validacion de palabras
        if palabras_texto and palabras_tabla:
            palabras_minimas = max(
                2, min(len(palabras_texto), len(palabras_tabla)) // 3
            )
            validaciones["concepto"]["match"] = coincidencias >= palabras_minimas
        else:
            validaciones["concepto"]["match"] = False

        if not validaciones["concepto"]["match"]:
            validaciones["concepto"][
                "diferencia"
            ] = f"Solo {coincidencias} palabras coinciden (minimo: {palabras_minimas if 'palabras_minimas' in locals() else 2})"

    # --- VALIDAR CAMPOS OBLIGATORIOS ---
    campos_obligatorios = {
        "nit": "NIT",
        "concepto_compra": "Concepto de Compra",
        "cantidad": "Cantidad",
        "valor_total": "Valor Total",
    }

    campos_presentes = 0
    campos_faltantes = []

    for campo, nombre in campos_obligatorios.items():
        if datos_texto.get(campo) and str(datos_texto[campo]).strip():
            campos_presentes += 1
        else:
            campos_faltantes.append(nombre)

    validaciones["campos_obligatorios"]["presentes"] = campos_presentes
    validaciones["campos_obligatorios"]["total"] = len(campos_obligatorios)
    validaciones["campos_obligatorios"]["faltantes"] = campos_faltantes

    # --- CALCULAR RESUMEN ---
    campos_para_validar = [
        "cantidad",
        "valor_unitario",
        "valor_total",
        # "fecha_entrega",
        "concepto",
    ]
    campos_validados = sum(
        [1 for campo in campos_para_validar if validaciones[campo]["match"]]
    )

    validaciones["campos_validados"] = campos_validados
    validaciones["total_campos"] = len(campos_para_validar)

    validaciones["resumen"] = (
        f"{campos_validados}/{len(campos_para_validar)} campos coinciden, "
        f"{campos_presentes}/{len(campos_obligatorios)} campos obligatorios presentes"
    )

    return validaciones


def LimpiarNumero(valor: str) -> float:
    """Convierte string con formato monetario a numero con mejor manejo de errores"""
    if not valor or valor == "N/A" or str(valor).strip() == "":
        return 0.0

    try:
        # Convertir a string y limpiar
        valor_str = str(valor).strip()

        # Eliminar simbolos monetarios y espacios
        valor_limpio = valor_str.replace("$", "").replace(" ", "").strip()

        # Detectar separador decimal
        # Si tiene tanto punto como coma, el ultimo es el decimal
        if "." in valor_limpio and "," in valor_limpio:
            if valor_limpio.rindex(".") > valor_limpio.rindex(","):
                # Punto es decimal (formato US: 1,000.50)
                valor_limpio = valor_limpio.replace(",", "")
            else:
                # Coma es decimal (formato EU: 1.000,50)
                valor_limpio = valor_limpio.replace(".", "").replace(",", ".")
        elif "," in valor_limpio:
            # Solo comas - podria ser miles o decimal
            if valor_limpio.count(",") == 1 and len(valor_limpio.split(",")[1]) == 2:
                # Es decimal (formato: 1000,50)
                valor_limpio = valor_limpio.replace(",", ".")
            else:
                # Es separador de miles (formato: 1,000 o 1,000,000)
                valor_limpio = valor_limpio.replace(",", "")
        elif "." in valor_limpio:
            # Solo puntos - podria ser miles o decimal
            if valor_limpio.count(".") == 1 and len(valor_limpio.split(".")[1]) == 2:
                # Es decimal (formato: 1000.50)
                pass  # Ya esta en formato correcto
            else:
                # Es separador de miles (formato: 1.000 o 1.000.000)
                valor_limpio = valor_limpio.replace(".", "")

        # Convertir a float
        return float(valor_limpio)

    except Exception as e:
        print(f"ERROR limpiando numero '{valor}': {e}")
        return 0.0


def NormalizarFecha(fecha: str) -> str:
    """Normaliza formato de fecha para comparacion"""
    if not fecha:
        return ""
    # Intentar parsear y normalizar
    for formato in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]:
        try:
            dt = datetime.strptime(fecha.strip(), formato)
            return dt.strftime("%Y-%m-%d")
        except:
            continue
    return fecha.strip()


def DeterminarEstadoFinal(datos_texto: Dict, validaciones: Dict) -> Tuple[str, str]:
    """
    Determina el estado final y observaciones basado en validaciones
    AJUSTADO: Maneja textos que solo son descripciones sin datos estructurados
    """
    # Cortar validación temprana para textos sin estructura
    if datos_texto.get("tipo_texto") == "solo_descripcion":
        return "Solo descripcion", "El texto solo contiene una descripción del producto"

    if datos_texto.get("tipo_texto") == "tabla_sap":
        return "Texto invalido", "El texto contiene una tabla SAP exportada"

    if datos_texto.get("tipo_texto") == "vacio":
        return "Sin Texto", "El item no tiene texto"

    campos_obligatorios_presentes = validaciones.get("campos_obligatorios", {}).get(
        "presentes", 0
    )
    total_campos_obligatorios = validaciones.get("campos_obligatorios", {}).get(
        "total", 4
    )
    campos_validados = validaciones.get("campos_validados", 0)

    # CASO 1: Texto vacio o muy corto
    concepto = datos_texto.get("concepto_compra", "")
    if not concepto or len(concepto.strip()) < 5:
        return "Sin Texto", "No se encontro texto en el item"

    # CASO 2: Texto es solo tabla de SAP (detectar por pipes y guiones)
    if concepto.count("|") > 10 and concepto.count("-") > 50:
        return (
            "Texto invalido",
            "El texto es una tabla de SAP exportada, no contiene informacion del proveedor",
        )

    # CASO 3: Texto es solo descripcion del producto (sin datos del proveedor)
    # Si NO tiene ningun campo obligatorio Y el texto es corto (menos de 200 chars)
    if campos_obligatorios_presentes == 0 and len(concepto) < 200:
        return (
            "Solo descripcion",
            f"Texto solo contiene descripcion del producto: {concepto[:50]}...",
        )

    # CASO 4: Texto tiene algunos datos pero incompletos
    if campos_obligatorios_presentes == 0 and len(concepto) >= 200:
        return (
            "Verificar manualmente",
            "Texto extenso pero sin campos estructurados (NIT, valores, etc)",
        )

    # CASO 5: Validacion normal - tiene campos estructurados
    if campos_obligatorios_presentes >= 3 and campos_validados >= 3:
        estado = "Registro validado para orden de compra"
        observaciones = "Validacion exitosa - Cumple requisitos minimos"
    elif campos_obligatorios_presentes >= 2:
        estado = "Verificar manualmente"
        observaciones = GenerarObservaciones(datos_texto, validaciones)
        if campos_validados < 2:
            estado = "Datos no coinciden con SAP"
    else:
        estado = "Falta informacion critica"
        observaciones = GenerarObservaciones(datos_texto, validaciones)

    # =============================================================
    #  VALIDACIÓN DE CAMPOS OBLIGATORIOS SAP (ME53N)
    #  SI FALTAN → ESTADO = "Verificar manualmente"
    # =============================================================
    campos_me = validaciones.get("campos_me53n", {})

    if campos_me and campos_me.get("faltantes"):
        faltantes = campos_me["faltantes"]

        # agregar a observaciones
        if observaciones:
            observaciones += " | "

        observaciones += f"Faltan campos SAP: {', '.join(faltantes)}"

        # cambiar estado final
        return "Verificar manualmente", observaciones

    return estado, observaciones


def ExtraerDatosTexto(texto: str) -> Dict:
    """Extrae campos estructurados del texto capturado
    AJUSTADO: Se agregan nuevos patrones manteniendo la lógica existente."""

    datos = {
        "razon_social": "",
        "nit": "",
        "correo": "",
        "empresa": "",
        "concepto_compra": "",
        "fecha_prestacion": "",
        "valor_unitario": "",
        "valor_total": "",
        "cantidad": "",
        "subtotal": "",
        "iva_impo": "",
        "total": "",
        "responsable_compra": "",
        "ceco": "",
        "telefono": "",
        "direccion_entrega": "",
        "tipo_texto": "desconocido",
    }

    if not texto or not texto.strip():
        datos["tipo_texto"] = "vacio"
        return datos

    texto_limpio = texto.strip()
    texto_upper = texto_limpio.upper()
    lineas = [linea.strip() for linea in texto_limpio.split("\n") if linea.strip()]

    # ------------------------------------------------------------
    # CLASIFICACIÓN DEL TEXTO (MANTENIDA TAL CUAL)
    # ------------------------------------------------------------

    if texto.count("|") > 10 and texto.count("-") > 50:
        datos["tipo_texto"] = "tabla_sap"
        datos["concepto_compra"] = "TABLA SAP EXPORTADA"
        return datos

    if len(texto_limpio) < 200 and not any(
        kw in texto_upper
        for kw in ["NIT", "RAZON SOCIAL", "VALOR", "CANTIDAD:", "PROVEEDOR"]
    ):
        datos["tipo_texto"] = "solo_descripcion"
        datos["concepto_compra"] = texto_limpio
        return datos

    if any(
        kw in texto_upper
        for kw in ["NIT", "RAZON SOCIAL", "PROVEEDOR:", "VALOR TOTAL", "CANTIDAD:"]
    ):
        datos["tipo_texto"] = "estructurado"
    else:
        datos["tipo_texto"] = "texto_simple"

    # ------------------------------------------------------------
    # EXTRACCIÓN DE CAMPOS (NUEVOS + EXISTENTES)
    # ------------------------------------------------------------

    # --- RAZON SOCIAL (NUEVO + mantiene lo que tenías) ---
    m = re.search(r"GENERAR ORDEN DE COMPRA A[:\s]*(.+?)\s*NIT", texto_upper)
    if m:
        datos["razon_social"] = m.group(1).strip()

    # fallback TUYO (línea con RAZON SOCIAL o PROVEEDOR)
    if not datos["razon_social"]:
        for linea in lineas:
            linea_upper = linea.upper()
            if any(kw in linea_upper for kw in ["RAZON SOCIAL", "PROVEEDOR"]):
                if ":" in linea:
                    datos["razon_social"] = linea.split(":", 1)[1].strip()
                    break

    # --- NIT (MANTENIDO + más flexible) ---
    patrones_nit = [r"NIT[\s:]*([0-9.\-]+)", r"IDENTIFICACION[\s:]*([0-9.\-]+)"]
    for patron in patrones_nit:
        m = re.search(patron, texto_upper)
        if m:
            datos["nit"] = m.group(1).strip()
            break

    # --- CORREOS (MEJORA) ---
    correos = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", texto)

    # Correo del proveedor (NO colsubsidio)
    correos_proveedor = [c for c in correos if "colsubsidio.com" not in c.lower()]
    if correos_proveedor:
        datos["correo"] = correos_proveedor[0].strip()

    # Correos responsables (@colsubsidio.com)
    correos_resp = [c.lower() for c in correos if "colsubsidio.com" in c.lower()]
    if correos_resp:
        datos["responsable_compra"] = ", ".join(correos_resp)

    # --- CONCEPTO (NUEVO + tu fallback) ---
    m = re.search(
        r"POR CONCEPTO DE[:\s]*(.+?)\s*(EMPRESA|FECHA|HORA|DIRECCION|CANTIDAD)",
        texto_upper,
    )
    if m:
        datos["concepto_compra"] = m.group(1).strip()

    # fallback TUYO
    if not datos["concepto_compra"] and lineas:
        datos["concepto_compra"] = lineas[0][:200]

    # --- CANTIDAD (MANTENIDO) ---
    m = re.search(r"CANTIDAD[\s:]*([0-9.,]+)", texto_upper)
    if m:
        datos["cantidad"] = m.group(1).strip()

    # --- VALORES (MANTENIDO) ---
    patron_valor = r"[\$]?\s*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)"

    # valor unitario
    for linea in lineas:
        if any(
            kw in linea.upper()
            for kw in ["VALOR UNITARIO", "VR UNITARIO", "PRECIO UNITARIO"]
        ):
            m = re.search(patron_valor, linea)
            if m:
                datos["valor_unitario"] = m.group(1).strip()
                break

    # valor total / subtotal
    for linea in lineas:
        if any(kw in linea.upper() for kw in ["VALOR TOTAL", "SUBTOTAL"]):
            m = re.search(patron_valor, linea)
            if m:
                datos["valor_total"] = m.group(1).strip()
                break

    # --- CECO (NUEVO) ---
    m = re.search(r"CECO[:\s]*([0-9]+)", texto_upper)
    if m:
        datos["ceco"] = m.group(1).strip()

    return datos


def GenerarObservaciones(datos_texto: Dict, validaciones: Dict) -> str:
    """Genera observaciones detalladas
    AJUSTADO: Incluye informacion sobre el tipo de texto"""

    observaciones = []

    # Agregar info sobre tipo de texto
    tipo_texto = datos_texto.get("tipo_texto", "desconocido")
    if tipo_texto == "solo_descripcion":
        return "Texto solo contiene descripcion del producto - No incluye datos del proveedor"
    elif tipo_texto == "tabla_sap":
        return "Error: Texto es una tabla de SAP, no informacion del proveedor"
    elif tipo_texto == "vacio":
        return "Item sin texto"
    elif tipo_texto == "texto_simple":
        observaciones.append("Texto no estructurado")

    # Campos obligatorios faltantes
    if "campos_obligatorios" in validaciones:
        campos_faltantes = validaciones["campos_obligatorios"].get("faltantes", [])
        if campos_faltantes:
            observaciones.append(f"Faltan: {', '.join(campos_faltantes)}")

    # Validaciones fallidas
    campos_validacion = {
        "cantidad": "Cantidad",
        "valor_unitario": "Valor unitario",
        "valor_total": "Valor total",
    }

    for campo, nombre in campos_validacion.items():
        if campo in validaciones and isinstance(validaciones[campo], dict):
            if validaciones[campo].get("texto") and not validaciones[campo].get(
                "match", False
            ):
                diferencia = validaciones[campo].get("diferencia", "")
                if diferencia:
                    observaciones.append(f"{nombre} {diferencia}")

    if not observaciones:
        return "Texto sin campos requeridos para validacion"

    return " | ".join(observaciones[:5])


def GenerarReporteValidacion(
    solped: str, item: str, datos_texto: Dict, validaciones: Dict
) -> str:
    """Genera un reporte legible de la validacion"""

    reporte = f"\n{'='*80}\n"
    reporte += f"REPORTE DE VALIDACION - SOLPED: {solped}, ITEM: {item}\n"
    reporte += f"{'='*80}\n\n"

    # ===================================================================================
    # 1. CAMPOS OBLIGATORIOS SAP ME53N
    # ===================================================================================
    if "campos_me53n" in validaciones:
        datos_me53n = validaciones["campos_me53n"]
        reporte += "CAMPOS OBLIGATORIOS SAP ME53N:\n"
        reporte += f"  Presentes: {datos_me53n['presentes']}/{datos_me53n['total']}\n"

        if datos_me53n["faltantes"]:
            reporte += "  Faltantes:\n"
            for campo in datos_me53n["faltantes"]:
                reporte += f"    - {campo}\n"
        else:
            reporte += "  ✓ Todos los campos obligatorios ME53N están presentes.\n"

        reporte += "\n"

    # ===================================================================================
    # 2. ADVERTENCIA POR TEXTO SIN ESTRUCTURA
    # ===================================================================================
    if datos_texto.get("tipo_texto") in ["solo_descripcion", "vacio", "tabla_sap"]:
        reporte += (
            "ADVERTENCIA IMPORTANTE:\n"
            "  Este ítem contiene solo descripción o no tiene estructura completa.\n"
            "  → No fue posible extraer todos los campos estructurados.\n"
            "  → Validar manualmente la información.\n"
            "  → Revisar cantidad, valor unitario, valor total y NIT directamente en SAP.\n\n"
        )

    # ===================================================================================
    # 3. DATOS EXTRAÍDOS DEL TEXTO
    # ===================================================================================
    reporte += "DATOS EXTRAIDOS DEL TEXTO:\n"
    reporte += f"  Razon Social: {datos_texto.get('razon_social') or 'No encontrado'}\n"
    reporte += f"  NIT: {datos_texto.get('nit') or 'No encontrado'}\n"
    reporte += f"  Correo: {datos_texto.get('correo') or 'No encontrado'}\n"
    reporte += (
        f"  Concepto: {datos_texto.get('concepto_compra')[:50] or 'No encontrado'}...\n"
    )
    reporte += f"  Cantidad: {datos_texto.get('cantidad') or 'No encontrado'}\n"
    reporte += (
        f"  Valor Unitario: {datos_texto.get('valor_unitario') or 'No encontrado'}\n"
    )
    reporte += f"  Valor Total: {datos_texto.get('valor_total') or 'No encontrado'}\n"
    reporte += (
        f"  Responsable: {datos_texto.get('responsable_compra') or 'No encontrado'}\n"
    )
    reporte += f"  CECO: {datos_texto.get('ceco') or 'No encontrado'}\n\n"

    # ===================================================================================
    # 4. CAMPOS OBLIGATORIOS EXTRAÍDOS DEL TEXTO
    # ===================================================================================
    if "campos_obligatorios" in validaciones:
        oblig = validaciones["campos_obligatorios"]
        reporte += "CAMPOS OBLIGATORIOS (Segun Texto Extraído):\n"
        reporte += f"  Presentes: {oblig['presentes']}/{oblig['total']}\n"

        if oblig["faltantes"]:
            reporte += "  Faltantes:\n"
            for campo in oblig["faltantes"]:
                reporte += f"    - {campo}\n"
        else:
            reporte += "  ✓ Todos los campos obligatorios del texto están presentes.\n"

        reporte += "\n"

    # ===================================================================================
    # 5. VALIDACIONES DETALLADAS
    # ===================================================================================
    reporte += "VALIDACIONES:\n"

    for campo, validacion in validaciones.items():

        if campo in ["resumen", "campos_obligatorios", "campos_me53n"]:
            continue  # ya fueron procesados

        if not isinstance(validacion, dict):
            continue

        if "match" not in validacion:
            continue

        estado = "EXITO" if validacion["match"] else "ERROR"
        reporte += f"  {estado} {campo.upper()}:\n"

        if validacion.get("texto"):
            reporte += f"      Texto: {validacion['texto']}\n"

        if validacion.get("tabla"):
            reporte += f"      Tabla: {validacion['tabla']}\n"

        if validacion.get("diferencia"):
            reporte += f"      {validacion['diferencia']}\n"

    # ===================================================================================
    # 6. RESUMEN FINAL
    # ===================================================================================
    reporte += f"\n{validaciones['resumen']}\n"
    reporte += f"{'='*80}\n"

    return reporte


def ProcesarYValidarItem(
    session, solped: str, item_num: str, texto: str, df_items: pd.DataFrame
) -> Tuple[Dict, Dict, str, str, str]:
    """
    Procesa un item: extrae datos, valida y genera reporte
    Returns: (datos_texto, validaciones, reporte, estado_final, observaciones)
    """

    # 1. Extraer datos del texto
    datos_texto = ExtraerDatosTexto(texto)

    # 2. Validar contra tabla (pasando el numero de item para busqueda especifica)
    validaciones = ValidarContraTabla(datos_texto, df_items, item_num)

    # 3. Determinar estado final y observaciones
    estado_final, observaciones = DeterminarEstadoFinal(datos_texto, validaciones)
    # Evitar generar reportes completos cuando el texto no tiene estructura
    if datos_texto.get("tipo_texto") in ["vacio", "solo_descripcion", "tabla_sap"]:
        observaciones = (
            f"Texto sin estructura completa ({datos_texto.get('tipo_texto')}). "
            "Solo contiene descripción."
        )

    # 4. Generar reporte
    reporte = GenerarReporteValidacion(solped, item_num, datos_texto, validaciones)

    return datos_texto, validaciones, reporte, estado_final, observaciones

# ============================================
# HU06: Organizar y enviar por correo Orden de Compra (OC)
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Organizar y enviar por correo Orden de Compra (OC)
# Ultima modificacion: 08/14/2026
# Propiedad de Colsubsidio
# Cambios: Estructura y logs.
# ============================================

import os
import re
import pandas as pd
import shutil
from PyPDF2 import PdfReader
from datetime import datetime
#from Funciones.GeneralME53N import EnviarNotificacionCorreo
from Funciones.EmailSender import EnviarNotificacionCorreo
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
import csv


# =============================
# CONFIGURACIÓN
# =============================

INPUT_DIR = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Ordenes de Compra"
OUTPUT_DIR = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\OC_Proveedores"
ERROR_DIR = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\pdf_no_procesados"
ADJURIDICO = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Juridicas"
ADESTANDAR = rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Estandar"


REGEX_OC = r"ORDEN\s+DE\s+COMPRA\s*(?:N[°ºo]?\.?)?\s*(\d{8,12})"

REGEX_SR = r"Sr\.\s*(?:[\s\S]{0,80}?)\:\s*([A-ZÁÉÍÓÚÑ\s]{5,60})"
REGEX_EMPRESA = r"EMPRESA\s*:\s*([A-ZÁÉÍÓÚÑ\s]{2,60})"
REGEX_CORREOS = r"(?:CORREO\s*ELECTRONICO|CORREO|E-MAIL)\s*[:\-]?\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})"

REGEX_RAZON_SOCIAL = r"RAZON\s+SOCIAL\s*(?:\n|\r|\s)*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"
REGEX_SOLICITA_OC = (
    r"SE\s+SOLICITA\s+GENERAR\s+OC\s+A\s*(?:\n|\r|\s)*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"
)
REGEX_PROVEEDOR_LINEA = r"PROVEEDOR\s*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"

rutaParametros = r"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Archivo_Parametros.xlsx"


# =============================
# UTILIDADES
# =============================


def generarEnviosYEeporte(consolidado):
    reportePath = os.path.join(
        OUTPUT_DIR, f"Informe_Envios_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    )

    with open(reportePath, mode="w", newline="", codificacion="utf-8") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(
            ["Proveedor", "Correos Destino", "Ordenes de Compra", "Estado Envío"]
        )

        for proveedor, datos in consolidado.items():
            adjuntosFinales = list(datos["rutas_archivos"])

            # Definir la ruta de adjuntos extra según el tipo
            rutaExtra = ADJURIDICO if datos["tipo"] == "Juridico" else ADESTANDAR
            # VALIDACIÓN: Si es una carpeta, recorre y añade cada archivo
            if os.path.isdir(rutaExtra):
                for archivo in os.listdir(rutaExtra):
                    rutaCompleta = os.path.join(rutaExtra, archivo)
                    # Solo añade si es un archivo (evita subcarpetas)
                    if os.path.isfile(rutaCompleta):
                        adjuntosFinales.append(rutaCompleta)
            else:
                # Si la ruta era un archivo individual, lo añade directo
                adjuntosFinales.append(rutaExtra)

            try:
                ocsStr = ", ".join(datos["ocs"])
                correosStr = ", ".join(datos["correos"])

                log(f"--- Enviando Correo Consolidado a: {proveedor} ---")
                log(f"OCs incluidas: {ocsStr}")
                log(f"Adjuntos: {len(datos['rutas_archivos'])} archivos")

                # Aquí iría tu lógica real de Outlook/SMTP enviando todos los datos['adjuntosFinales']

                EnviarNotificacionCorreo(
                    codigoCorreo=1,
                    nombreTarea="Prueba - Notificacion",
                    adjuntos=adjuntosFinales,
                )

                # Registro en informe
                writer.writerow([proveedor, correosStr, ocsStr, "Enviado"])

            except Exception as e:
                log(f"Error enviando a {proveedor}: {e}")
                writer.writerow([proveedor, correosStr, ocsStr, f"Error: {e}"])

    log(f"Informe generado en: {reportePath}")


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def extractTextFromPdf(pdfPath):
    reader = PdfReader(pdfPath)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text


def parseOc(text):
    match = re.search(REGEX_OC, text, re.IGNORECASE)
    if not match:
        raise ValueError("Número de OC no encontrado")
    return match.group(1)


def parseEmpresa(text):
    match = re.search(REGEX_EMPRESA, text, re.IGNORECASE)
    if not match:
        raise ValueError("Empresa no encontrada")

    return limpiarNombre(match.group(1))


def parseProveedor(text):
    patrones = [
        REGEX_SR,  # Bloque SAP con Sr. separado
        REGEX_RAZON_SOCIAL,  # RAZON SOCIAL:
        REGEX_SOLICITA_OC,  # SE SOLICITA GENERAR OC A:
        REGEX_PROVEEDOR_LINEA,  # PROVEEDOR:
        REGEX_EMPRESA,  # EMPRESA:
    ]

    for patron in patrones:
        match = re.search(patron, text, re.IGNORECASE)
        if match:
            return limpiarNombre(match.group(1))

    raise ValueError("Proveedor no encontrado en ningún formato")


def parseProveedorSr(text):
    match = re.search(REGEX_SR, text, re.IGNORECASE)
    if not match:
        raise ValueError("Proveedor (Sr.) no encontrado")

    return limpiarNombre(match.group(1))


def limpiarNombre(nombre):
    nombre = nombre.splitlines()[0]
    nombre = nombre.strip().upper()
    nombre = re.sub(r"\s{2,}", " ", nombre)
    return nombre


def safe_name(text):
    return re.sub(r"[<>:\"/\\|?*]", "", text)


def organizarPdf(pdfPath, oc, proveedorSr, empresa):
    proveedorSafe = safe_name(proveedorSr)
    empresaSafe = safe_name(empresa)

    destinoDir = os.path.join(OUTPUT_DIR, proveedorSafe)
    os.makedirs(destinoDir, exist_ok=True)

    nuevoNombre = f"OC_{oc}_{empresaSafe}.pdf"
    destinoPdf = os.path.join(destinoDir, nuevoNombre)

    shutil.move(pdfPath, destinoPdf)
    return destinoPdf


def enviarCorreoSimulado(correos, oc):
    log(f"Simulación envío OC {oc}")
    log("Correos detectados:")
    for correo in correos:
        print(f"  - {correo}")


def parseCorreos(text):
    correos = re.findall(REGEX_CORREOS, text, re.IGNORECASE)
    correosLimpios = list(set(c.lower() for c in correos))
    return correosLimpios


def obtenerTipoProveedor(textoPdf, dictTipos):
    # Buscamos patrones numéricos de NIT
    match = re.search(r"(?:NIT|Nit/C\.C\.)\s*[:\-]?\s*([\d\.\-]+)", textoPdf)

    if match:
        # Limpiamos el NIT: quitamos puntos, guiones y espacios
        nitSucio = match.group(1)
        nitLimpio = re.sub(r"\D", "", nitSucio)

        # Retornamos el tipo; si no existe, devolvemos 'No Registrado'
        return dictTipos.get(nitLimpio, "No Registrado")

    return "Sin NIT en PDF"


def EjecutarHU06():

    nombreTarea = "HU06_EnvioCorreoOC"
    # Diccionario para agrupar: { "Nombre Proveedor": { "correos": [], "ocs": [], "archivos": [] } }
    consolidadoProveedores = {}

    # Cargar la tabla de parámetros Hoja: Proveedores
    condfParametros = pd.read_excel(rutaParametros, sheet_name="Proveedores")

    # Limpieza y creación del diccionario de tipos
    # Aseguramos que el NIT sea string y no tenga decimales (.0)
    condfParametros["Nit"] = (
        condfParametros["Nit"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

    # Convertimos a diccionario para búsquedas ultra rápidas por NIT
    # { '890900076': 'Juridico', ... }
    dictTipos = dict(zip(condfParametros["Nit"], condfParametros["Tipo de proveedor"]))

    # ================================
    # Envio Correo OC
    # ================================
    WriteLog(
        mensaje="Inicio ejecución Envio Correo de OC.",
        estado="INFO",
        nombreTarea=nombreTarea,
        rutaRegistro=RUTAS["PathLog"],
    )

    # Asegurar que los directorios existen
    os.makedirs(ERROR_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(INPUT_DIR, exist_ok=True)

    for file in os.listdir(INPUT_DIR):
        # se asegura que sea un archivo .pdf
        if not file.lower().endswith(".pdf"):
            continue  # pasa a la siguiete iteracion si no es un pdf

        pdfPath = os.path.join(INPUT_DIR, file)

        try:
            log(f"Procesando: {file}")

            text = extractTextFromPdf(pdfPath)

            oc = parseOc(text)
            proveedor = parseProveedor(text)
            empresa = parseEmpresa(text)
            correos = parseCorreos(text)

            rutaFinal = organizarPdf(pdfPath, oc, proveedor, empresa)

            # Agrupar datos para el envío consolidado
            tipo = obtenerTipoProveedor(text, dictTipos)
            if proveedor not in consolidadoProveedores:
                consolidadoProveedores[proveedor] = {
                    "correos": correos,
                    "ocs": [],
                    "rutas_archivos": [],
                    "tipo": tipo,  # Guardamos si es Juridico o Persona Natural
                }

            consolidadoProveedores[proveedor]["ocs"].append(oc)
            consolidadoProveedores[proveedor]["rutas_archivos"].append(rutaFinal)

            log(f"Agrupado OK: OC {oc} para {proveedor}")

        except Exception as e:
            log(f"ERROR: {file} - {e}")
            shutil.move(pdfPath, os.path.join(ERROR_DIR, file))

    # ============================================
    # ENVÍO CONSOLIDADO E INFORME
    # ================================
    generarEnviosYEeporte(consolidadoProveedores)

    WriteLog(
        mensaje="Fin ejecución Envio Correo de OC.",
        estado="INFO",
        nombreTarea=nombreTarea,
        rutaRegistro=RUTAS["PathLog"],
    )
    log("Fin proceso RPA")

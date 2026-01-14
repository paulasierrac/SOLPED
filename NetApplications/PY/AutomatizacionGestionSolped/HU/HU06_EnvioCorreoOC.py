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
import shutil
from PyPDF2 import PdfReader
from datetime import datetime

# =============================
# CONFIGURACIÓN
# =============================

INPUT_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Ordenes de Compra"
OUTPUT_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\OC_Proveedores"
ERROR_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\pdf_no_procesados"

REGEX_OC = r"ORDEN\s+DE\s+COMPRA\s*(?:N[°ºo]?\.?)?\s*(\d{8,12})"
REGEX_PROVEEDOR = r"(?:Sr\.?\s*:|Proveedor\s*:|Vendor\s*:)\s*([A-ZÁÉÍÓÚÑ\s\.]{5,60})"
REGEX_PROVEEDOR_SOLICITA = r"SE\s+SOLICITA\s+GENERAR\s+OC\s+A\s*:\s*([A-ZÁÉÍÓÚÑ\s]{5,60})"
REGEX_PROVEEDOR_SR = r"Sr\.\s*[\r\n]+:\s*([A-ZÁÉÍÓÚÑ\s]{5,60})"
REGEX_EMPRESA = r"EMPRESA\s*:\s*([A-ZÁÉÍÓÚÑ\s]{2,60})"


# =============================
# UTILIDADES
# =============================

def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def parse_oc(text):
    match = re.search(REGEX_OC, text, re.IGNORECASE)
    if not match:
        raise ValueError("Número de OC no encontrado")
    return match.group(1)

def parse_proveedor(text):
    # 1. Caso explícito: "SE SOLICITA GENERAR OC A:"
    match = re.search(REGEX_PROVEEDOR_SOLICITA, text, re.IGNORECASE)
    if match:
        return limpiar_nombre(match.group(1))

    # 2. Caso SAP: Sr. en línea + nombre en siguiente
    match = re.search(REGEX_PROVEEDOR_SR, text, re.IGNORECASE)
    if match:
        return limpiar_nombre(match.group(1))

    # 3. Fallback: EMPRESA
    match = re.search(REGEX_EMPRESA, text, re.IGNORECASE)
    if match:
        return limpiar_nombre(match.group(1))

    raise ValueError("Proveedor no encontrado")


def limpiar_nombre(nombre):
    nombre = nombre.splitlines()[0]
    nombre = nombre.strip().upper()
    nombre = re.sub(r"\s{2,}", " ", nombre)
    return nombre

def safe_name(text):
    return re.sub(r"[<>:\"/\\|?*]", "", text)

def organizar_pdf(pdf_path, oc, proveedor):
    proveedor_safe = safe_name(proveedor)
    destino_dir = os.path.join(OUTPUT_DIR, proveedor_safe)
    os.makedirs(destino_dir, exist_ok=True)

    nuevo_nombre = f"OC_{oc}_{proveedor_safe}.pdf"
    destino_pdf = os.path.join(destino_dir, nuevo_nombre)

    shutil.move(pdf_path, destino_pdf)
    return destino_pdf

def enviar_correo_simulado(proveedor, oc, pdf_path):
    log(f"Correo enviado a proveedor [{proveedor}] con OC [{oc}] adjunta.")
    log(f"Adjunto: {pdf_path}")

# =============================
# MAIN
# =============================

def EjecutarHU06():
    log("Inicio proceso RPA - Órdenes de Compra")

    os.makedirs(ERROR_DIR, exist_ok=True)

    for file in os.listdir(INPUT_DIR):
        if not file.lower().endswith(".pdf"):
            continue

        pdf_path = os.path.join(INPUT_DIR, file)

        try:
            log(f"Procesando: {file}")

            text = extract_text_from_pdf(pdf_path)
            #print(text)
            oc = parse_oc(text)
            proveedor = parse_proveedor(text)

            nuevo_pdf = organizar_pdf(pdf_path, oc, proveedor)
            enviar_correo_simulado(proveedor, oc, nuevo_pdf)

            log(f"Procesado OK: OC {oc}")

        except Exception as e:
            log(f"ERROR: {file} - {e}")
            shutil.move(pdf_path, os.path.join(ERROR_DIR, file))
            raise

    log("Fin proceso RPA")

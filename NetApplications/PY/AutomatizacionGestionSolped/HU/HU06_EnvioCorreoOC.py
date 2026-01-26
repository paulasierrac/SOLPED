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
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
import csv


# =============================
# CONFIGURACIÓN
# =============================

INPUT_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Ordenes de Compra"
OUTPUT_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\OC_Proveedores"
ERROR_DIR = fr"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\pdf_no_procesados"

REGEX_OC = r"ORDEN\s+DE\s+COMPRA\s*(?:N[°ºo]?\.?)?\s*(\d{8,12})"

REGEX_SR = r"Sr\.\s*(?:[\s\S]{0,80}?)\:\s*([A-ZÁÉÍÓÚÑ\s]{5,60})"
REGEX_EMPRESA = r"EMPRESA\s*:\s*([A-ZÁÉÍÓÚÑ\s]{2,60})"
REGEX_CORREOS = r"(?:CORREO\s*ELECTRONICO|CORREO|E-MAIL)\s*[:\-]?\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})"

REGEX_RAZON_SOCIAL = r"RAZON\s+SOCIAL\s*(?:\n|\r|\s)*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"
REGEX_SOLICITA_OC = r"SE\s+SOLICITA\s+GENERAR\s+OC\s+A\s*(?:\n|\r|\s)*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"
REGEX_PROVEEDOR_LINEA = r"PROVEEDOR\s*:\s*([A-ZÁÉÍÓÚÑ\s]{5,80})"

ruta_parametros = r"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\Archivo_Parametros.xlsx"


# =============================
# UTILIDADES
# =============================



def generar_envios_y_reporte(consolidado):
    reporte_path = os.path.join(OUTPUT_DIR, f"Informe_Envios_{datetime.now().strftime('%Y%m%d_%H%M')}.csv")
    
    with open(reporte_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(['Proveedor', 'Correos Destino', 'Ordenes de Compra', 'Estado Envío'])

        for proveedor, datos in consolidado.items():
            adjuntos_finales = list(datos["rutas_archivos"])
    
            if datos["tipo"] == "Juridico":
                adjuntos_finales.append(RUTAS["AdjuntoJuridico"])
            else:
                # Adjuntos para otros (Persona Natural / No Registrado)
                adjuntos_finales.append(RUTAS["AdjuntoEstandar"])

            try:
                ocs_str = ", ".join(datos["ocs"])
                correos_str = ", ".join(datos["correos"])
                
                log(f"--- Enviando Correo Consolidado a: {proveedor} ---")
                log(f"OCs incluidas: {ocs_str}")
                log(f"Adjuntos: {len(datos['rutas_archivos'])} archivos")
                
                # Aquí iría tu lógica real de Outlook/SMTP enviando todos los datos['adjuntos_finales'] 
                
                # Registro en informe
                writer.writerow([proveedor, correos_str, ocs_str, 'Enviado'])
                
            except Exception as e:
                log(f"Error enviando a {proveedor}: {e}")
                writer.writerow([proveedor, correos_str, ocs_str, f'Error: {e}'])

    log(f"Informe generado en: {reporte_path}")

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

def parse_empresa(text):
    match = re.search(REGEX_EMPRESA, text, re.IGNORECASE)
    if not match:
        raise ValueError("Empresa no encontrada")

    return limpiar_nombre(match.group(1))


def parse_proveedor(text):
    patrones = [
        REGEX_SR,                # Bloque SAP con Sr. separado
        REGEX_RAZON_SOCIAL,      # RAZON SOCIAL:
        REGEX_SOLICITA_OC,       # SE SOLICITA GENERAR OC A:
        REGEX_PROVEEDOR_LINEA,   # PROVEEDOR:
        REGEX_EMPRESA            # EMPRESA:
    ]

    for patron in patrones:
        match = re.search(patron, text, re.IGNORECASE)
        if match:
            return limpiar_nombre(match.group(1))

    raise ValueError("Proveedor no encontrado en ningún formato")

def parse_proveedor_sr(text):
    match = re.search(REGEX_SR, text, re.IGNORECASE)
    if not match:
        raise ValueError("Proveedor (Sr.) no encontrado")

    return limpiar_nombre(match.group(1))

def limpiar_nombre(nombre):
    nombre = nombre.splitlines()[0]
    nombre = nombre.strip().upper()
    nombre = re.sub(r"\s{2,}", " ", nombre)
    return nombre

def safe_name(text):
    return re.sub(r"[<>:\"/\\|?*]", "", text)

def organizar_pdf(pdf_path, oc, proveedor_sr, empresa):
    proveedor_safe = safe_name(proveedor_sr)
    empresa_safe = safe_name(empresa)

    destino_dir = os.path.join(OUTPUT_DIR, proveedor_safe)
    os.makedirs(destino_dir, exist_ok=True)

    nuevo_nombre = f"OC_{oc}_{empresa_safe}.pdf"
    destino_pdf = os.path.join(destino_dir, nuevo_nombre)

    shutil.move(pdf_path, destino_pdf)
    return destino_pdf

def enviar_correo_simulado(correos, oc):
    log(f"Simulación envío OC {oc}")
    log("Correos detectados:")
    for correo in correos:
        print(f"  - {correo}")


def parse_correos(text):
    correos = re.findall(REGEX_CORREOS, text, re.IGNORECASE)
    correos_limpios = list(set(c.lower() for c in correos))
    return correos_limpios

def obtener_tipo_proveedor(texto_pdf, dict_tipos):
    # Buscamos patrones numéricos de NIT
    match = re.search(r"(?:NIT|Nit/C\.C\.)\s*[:\-]?\s*([\d\.\-]+)", texto_pdf)
    
    if match:
        # Limpiamos el NIT: quitamos puntos, guiones y espacios
        nit_sucio = match.group(1)
        nit_limpio = re.sub(r"\D", "", nit_sucio)
        
        # Retornamos el tipo; si no existe, devolvemos 'No Registrado'
        return dict_tipos.get(nit_limpio, "No Registrado")
    
    return "Sin NIT en PDF"



def EjecutarHU06():

    task_name = "HU06_EnvioCorreoOC"
    # Diccionario para agrupar: { "Nombre Proveedor": { "correos": [], "ocs": [], "archivos": [] } }
    consolidado_proveedores = {}

    # Cargar la tabla de parámetros Hoja: Proveedores
    df_parametros = pd.read_excel(ruta_parametros, sheet_name='Proveedores')

    # Limpieza y creación del diccionario de tipos
    # Aseguramos que el NIT sea string y no tenga decimales (.0)
    df_parametros['Nit'] = df_parametros['Nit'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    # Convertimos a diccionario para búsquedas ultra rápidas por NIT
    # { '890900076': 'Juridico', ... }
    dict_tipos = dict(zip(df_parametros['Nit'], df_parametros['Tipo de proveedor']))
    

    # ================================
    # Envio Correo OC
    # ================================
    WriteLog(
        mensaje="Inicio ejecución Envio Correo de OC.",
        estado="INFO",
        task_name=task_name,
        path_log=RUTAS["PathLog"],
    )


    # Asegurar que los directorios existen
    os.makedirs(ERROR_DIR, exist_ok=True) 
    os.makedirs(OUTPUT_DIR, exist_ok=True) 
    os.makedirs(INPUT_DIR, exist_ok=True) 

    for file in os.listdir(INPUT_DIR):
        # se asegura que sea un archivo .pdf 
        if not file.lower().endswith(".pdf"):
            continue  # pasa a la siguiete iteracion si no es un pdf 

        pdf_path = os.path.join(INPUT_DIR, file)

        try:
            log(f"Procesando: {file}")

            text = extract_text_from_pdf(pdf_path)

            oc = parse_oc(text)
            proveedor = parse_proveedor(text)
            empresa = parse_empresa(text)
            correos = parse_correos(text)

            ruta_final = organizar_pdf(
                pdf_path,
                oc,
                proveedor,
                empresa
            )

            # Agrupar datos para el envío consolidado
            tipo = obtener_tipo_proveedor(text, dict_tipos)
            if proveedor not in consolidado_proveedores:
                consolidado_proveedores[proveedor] = {
                    "correos": correos,
                    "ocs": [],
                    "rutas_archivos": [],
                    "tipo": tipo # Guardamos si es Juridico o Persona Natural
                }
            
            consolidado_proveedores[proveedor]["ocs"].append(oc)
            consolidado_proveedores[proveedor]["rutas_archivos"].append(ruta_final)

            log(f"Agrupado OK: OC {oc} para {proveedor}")

        except Exception as e:
            log(f"ERROR: {file} - {e}")
            shutil.move(pdf_path, os.path.join(ERROR_DIR, file))

    # ============================================
    # ENVÍO CONSOLIDADO E INFORME
    # ================================
    generar_envios_y_reporte(consolidado_proveedores)

    WriteLog(
        mensaje="Fin ejecución Envio Correo de OC.",
        estado="INFO",
        task_name=task_name,
        path_log=RUTAS["PathLog"],
        )
    log("Fin proceso RPA")



# ============================================
# Función Local: ValidacionHU3
# Autor: Paula Sierra - NetApplications
# Descripcion: Funciones de validacion
# Ultima modificacion: 02/02/2026
# Propiedad de Colsubsidio
# Cambios:
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

def ValidarContraTabla(
    datosTexto: Dict, df_items: pd.DataFrame, item_num: str = ""
) -> Dict:
    """
    Compara los datos extraidos del texto con la tabla de items SAP ME53N.
    Maneja variaciones de nombres de columnas según idioma/configuración SAP.
    """
    validaciones = {
        "cantidad": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "valor_unitario": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "valor_total": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "concepto": {"match": False, "texto": "", "tabla": "", "diferencia": ""},
        "campos_obligatorios": {"presentes": 0, "total": 0, "faltantes": []},
        "resumen": "",
        "campos_validados": 0,
        "total_campos": 0,
    }

    campos_obligatorios_me53n = {
        "Material": "Material",
        "Quantity": "Cantidad",
        "Valn Price": "Precio valoración",
        "Deliv.Date": "Fecha entrega",
        "Plant": "Centro",
        "PGr": "Grupo de compras",
        "POrg": "Organización de compras",
        "Fix. Vend.": "Proveedor fijo",
    }

    if df_items.empty:
        validaciones["resumen"] = "Tabla vacia - No se puede validar"
        return validaciones

    # ============================================================
    # FILTRAR ITEM
    # ============================================================
    item_df = df_items
    if item_num and "Pos." in df_items.columns:
        item_df = df_items[
            df_items["Pos."].astype(str).str.strip() == str(item_num).strip()
        ]

    if item_df.empty:
        item_df = df_items.iloc[[0]]

    fila = item_df.iloc[0]  # ← FILA REAL

    columnas_disponibles = [str(c).strip() for c in fila.index]
    # ==================================================================
    # MAPEO DE CAMPOS SAP ME53N
    # Basado en nombres reales de la transacción según idioma
    # ==================================================================

    alias_campos = {
        # MATERIAL (Campo técnico: MATNR)
        "Material": [
            # Español
            "Material",
            "Número de material",
            "Nº material",
            "Mat.",
            # Inglés
            "Material",
            "Material Number",
            "Material No",
            "Mat. No.",
            # Portugués
            "Material",
            "Número do material",
            "Nº do material",
            # Otros
            "Código",
            "Code",
            "Item Code",
        ],
        # TEXTO BREVE (Campo técnico: TXZ01)
        "Short Text": [
            # Español
            "Texto breve",
            "Descripción",
            "Desc.",
            "Denominación",
            # Inglés
            "Short Text",
            "Short text",
            "Description",
            "Item Text",
            "Text",
            # Portugués
            "Texto breve",
            "Descrição",
            "Texto resumido",
        ],
        # CANTIDAD (Campo técnico: MENGE)
        "Quantity": [
            # Español
            "Cantidad",
            "Cant.",
            "Ctd.",
            "Cantidad pedido",
            # Inglés
            "Quantity",
            "Qty",
            "Order Quantity",
            "PO Quantity",
            # Portugués
            "Quantidade",
            "Qtd",
            "Qtde",
            "Quantidade pedido",
        ],
        # UNIDAD DE MEDIDA (Campo técnico: MEINS)
        "Un": [
            # Español
            "Un",
            "UM",
            "Unidad",
            "UMed",
            "Unidad medida",
            # Inglés
            "Un",
            "Unit",
            "UoM",
            "Unit of Measure",
            "U.M.",
            # Portugués
            "Un",
            "Unidade",
            "UM",
            "Unidade de medida",
        ],
        # PRECIO VALORACIÓN (Campo técnico: BPREIS / NETPR)
        "Valn Price": [
            # Español
            "Precio valoración",
            "Precio val.",
            "PrecioVal.",
            "PrecioVal",
            "Val. Price",
            "Precio unit.",
            "Precio unitario",
            "Pr.unit.",
            "Precio",
            # Inglés
            "Valn Price",
            "Valuation Price",
            "Unit Price",
            "Price",
            "Net Price",
            # Portugués
            "Preço valorização",
            "Preço unit.",
            "Preço unitário",
        ],
        # MONEDA (Campo técnico: WAERS)
        "Crcy": [
            # Español
            "Moneda",
            "Mon.",
            "Divisa",
            "Crcy",
            # Inglés
            "Crcy",
            "Currency",
            "Curr",
            "Ccy",
            # Portugués
            "Moeda",
            "Crcy",
        ],
        # VALOR TOTAL (Campo técnico: NETWR)
        "Total Val.": [
            # Español
            "Valor total",
            "Total",
            "Importe",
            "Imp.total",
            "Total Val.",
            "Valor tot.",
            "Valor tot",
            "Valor neto",
            "Importe neto",
            # Inglés
            "Total Val.",
            "Total Value",
            "Net Value",
            "Total",
            "Amount",
            "Net Price",
            "Total Amount",
            "Total Price",
            # Portugués
            "Valor total",
            "Total",
            "Valor líquido",
        ],
        # FECHA DE ENTREGA (Campo técnico: EINDT)
        "Deliv.Date": [
            # Español
            "Fecha entrega",
            "Fech.entr.",
            "Fecha de entrega",
            "Deliv.Date",
            "Fecha",
            "F.entrega",
            # Inglés
            "Deliv.Date",
            "Delivery Date",
            "Del. Date",
            "Delivery",
            "Date",
            # Portugués
            "Data entrega",
            "Data de entrega",
            "Dt.entrega",
        ],
        # PROVEEDOR FIJO (Campo técnico: LIFNR)
        "Fix. Vend.": [
            # Español
            "Proveedor fijo",
            "Prov.fijo",
            "Proveedor",
            "Fix. Vend.",
            # Inglés
            "Fix. Vend.",
            "Fixed Vendor",
            "Vendor",
            "Supplier",
            "Fixed Supplier",
            # Portugués
            "Fornecedor fixo",
            "Fornecedor",
            "Forn.fixo",
        ],
        # CENTRO (Campo técnico: WERKS)
        "Plant": [
            # Español
            "Centro",
            "Planta",
            "Cen.",
            "Plant",
            # Inglés
            "Plant",
            "Center",
            "Plnt",
            # Portugués
            "Centro",
            "Planta",
        ],
        # GRUPO DE COMPRAS (Campo técnico: EKGRP)
        "PGr": [
            # Español
            "Grupo compras",
            "Gr.compras",
            "GC",
            "PGr",
            # Inglés
            "PGr",
            "Purchasing Group",
            "Purch. Group",
            "Pur. Group",
            # Portugués
            "Grupo compras",
            "Gr.compras",
        ],
        # ORGANIZACIÓN DE COMPRAS (Campo técnico: EKORG)
        "POrg": [
            # Español
            "Org.compras",
            "Organización compras",
            "Org. compras",
            "POrg",
            # Inglés
            "POrg",
            "Purchasing Organization",
            "Purch. Org",
            "Pur. Org",
            # Portugués
            "Org.compras",
            "Organização compras",
        ],
        # GRUPO DE ARTÍCULOS (Campo técnico: MATKL)
        "Matl Group": [
            # Español
            "Grupo artículos",
            "Gr.artículos",
            "Grupo material",
            "Matl Group",
            # Inglés
            "Matl Group",
            "Material Group",
            "Mat. Group",
            "Item Group",
            # Portugués
            "Grupo mercadorias",
            "Gr.mercadorias",
            "Grupo material",
        ],
    }

    # ==================================================================
    # FUNCIÓN DE BÚSQUEDA MEJORADA
    # ==================================================================
    def buscar_columna(campo_estandar):
        posibles = alias_campos.get(campo_estandar, [])

        for alias in posibles:
            for col in columnas_disponibles:
                if alias.lower().strip() == col.lower():
                    return col

        for alias in posibles:
            for col in columnas_disponibles:
                if alias.lower() in col.lower() or col.lower() in alias.lower():
                    return col

        return None

    # ============================================================
    # VALIDAR CAMPOS OBLIGATORIOS ME53N (NEGOCIO)
    # ============================================================

    faltantes = []

    for campo_estandar, nombre_negocio in campos_obligatorios_me53n.items():
        col = buscar_columna(campo_estandar)

        if not col:
            faltantes.append(nombre_negocio)
            continue

        val = fila.get(col, "")
        if val is None or str(val).strip() in ["", "nan", "None"]:
            faltantes.append(nombre_negocio)

    validaciones["campos_me53n"] = {
        "presentes": len(campos_obligatorios_me53n) - len(faltantes),
        "total": len(campos_obligatorios_me53n),
        "faltantes": faltantes,
    }

    # ==================================================================
    # VALIDAR CANTIDAD
    # ==================================================================
    if datosTexto.get("cantidad"):
        cantidad_txt = LimpiarNumeroRobusto(datosTexto["cantidad"])
        col = buscar_columna("Quantity")

        if col:
            cantidad_tabla = LimpiarNumeroRobusto(fila.get(col))
            validaciones["cantidad"]["texto"] = datosTexto["cantidad"]
            validaciones["cantidad"]["tabla"] = str(cantidad_tabla)
            diff = abs(cantidad_txt - cantidad_tabla)

            validaciones["cantidad"]["match"] = diff < 0.01
            if not validaciones["cantidad"]["match"]:
                validaciones["cantidad"]["diferencia"] = f"Difiere en {diff:.2f}"

    # ==================================================================
    # VALIDAR VALOR UNITARIO
    # ==================================================================
    if datosTexto.get("valor_unitario"):
        val_txt = LimpiarNumeroRobusto(datosTexto["valor_unitario"])
        col = buscar_columna("Valn Price")

        if col:
            val_tabla = LimpiarNumeroRobusto(fila.get(col))
            validaciones["valor_unitario"]["texto"] = FormatoMoneda(val_txt)
            validaciones["valor_unitario"]["tabla"] = FormatoMoneda(val_tabla)

            diff = abs(val_txt - val_tabla)
            validaciones["valor_unitario"]["match"] = diff < 0.01

            if not validaciones["valor_unitario"]["match"]:
                validaciones["valor_unitario"][
                    "diferencia"
                ] = f"Difiere en {FormatoMoneda(diff)}"

    # ==================================================================
    # VALIDAR VALOR TOTAL
    # ==================================================================
    if datosTexto.get("valor_total"):
        val_txt = LimpiarNumeroRobusto(datosTexto["valor_total"])
        col = buscar_columna("Total Val.")

        if col:
            val_tabla = LimpiarNumeroRobusto(fila.get(col))
            validaciones["valor_total"]["texto"] = FormatoMoneda(val_txt)
            validaciones["valor_total"]["tabla"] = FormatoMoneda(val_tabla)

            diff = abs(val_txt - val_tabla)
            validaciones["valor_total"]["match"] = diff < 0.01

            if not validaciones["valor_total"]["match"]:
                validaciones["valor_total"][
                    "diferencia"
                ] = f"Difiere en {FormatoMoneda(diff)}"
    # ==================================================================
    # VALIDAR CONCEPTO
    # ==================================================================
    if datosTexto.get("concepto_compra"):
        col = buscar_columna("Short Text")

        if col:
            txt = datosTexto["concepto_compra"].upper()
            tabla = str(fila.get(col)).upper()

            palabras_txt = set(re.findall(r"\w+", txt))
            palabras_tabla = set(re.findall(r"\w+", tabla))

            coincidencias = len(palabras_txt & palabras_tabla)

            validaciones["concepto"]["texto"] = txt[:50]
            validaciones["concepto"]["tabla"] = tabla[:50]
            validaciones["concepto"]["match"] = coincidencias >= 2

            if not validaciones["concepto"]["match"]:
                validaciones["concepto"][
                    "diferencia"
                ] = f"Solo {coincidencias} palabras coinciden"

    # ============================================================
    # CAMPOS OBLIGATORIOS TEXTO
    # ============================================================
    oblig = ["nit", "concepto_compra", "cantidad", "valor_total"]
    presentes = sum(1 for c in oblig if datosTexto.get(c))
    faltan = [c for c in oblig if not datosTexto.get(c)]

    validaciones["campos_obligatorios"] = {
        "presentes": presentes,
        "total": len(oblig),
        "faltantes": faltan,
    }

    # ============================================================
    # RESUMEN
    # ============================================================
    campos = ["cantidad", "valor_unitario", "valor_total", "concepto"]
    ok = sum(1 for c in campos if validaciones[c]["match"])

    validaciones["campos_validados"] = ok
    validaciones["total_campos"] = len(campos)

    validaciones["resumen"] = (
        f"{ok}/{len(campos)} campos coinciden, "
        f"{presentes}/{len(oblig)} campos obligatorios presentes"
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

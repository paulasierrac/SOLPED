import pandas as pd
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from datetime import datetime
from openpyxl.utils import get_column_letter
import os

# ======================================================
# COLUMNAS OFICIALES DEL REPORTE FINAL
# ======================================================
COLUMNAS_REPORTE_FINAL = [
    "ID_REPORTE",
    # Datos expSolped03.txt
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
    # Adjuntos
    "CantAdjuntos",
    "Nombre de Adjunto",
    # Datos ME53N
    "Material_ME53N",
    "Short Text_ME53N",
    "Quantity_ME53N",
    "Un",
    "Valn Price",
    "Crcy",
    "Total Val.",
    "Deliv.Date",
    "Fix. Vend.",
    "Plant",
    "PGr_ME53N",
    "POrg",
    "Matl Group",
    # Texto del ítem
    "Id",
    "PurchReq_Texto",
    "Item_Texto",
    "Razon Social:",
    "NIT:",
    "Correo:",
    "Concepto:",
    "Cantidad:",
    "Valor Unitario:",
    "Valor Total:",
    "Responsable:",
    "CECO:",
    # Resultados validación
    "CAMPOS OBLIGATORIOS FALTANTES ME53N",
    "DATOS EXTRAIDOS DEL TEXTO FALTANTES",
    "CANTIDAD",
    "VALOR_UNITARIO",
    "VALOR_TOTAL",
    "CONCEPTO",
    "Estado",
    "Observaciones",
]


# ======================================================
# CONSTRUCTOR DE FILA CONSOLIDADA
# ======================================================


def determinar_estado_reporte(
    tiene_adjuntos: bool,
    faltantes_me53n: list,
    faltantes_texto: list,
    resultado_validaciones: dict,
):
    """
    Estados permitidos:
    Aprobado | Rechazado | Pendiente
    """

    # 🔴 RECHAZADO
    if not tiene_adjuntos:
        return "Rechazado"

    # 🟡 PENDIENTE
    if (
        faltantes_me53n
        or faltantes_texto
        or not resultado_validaciones.get("cantidad", True)
        or not resultado_validaciones.get("valor_unitario", True)
        or not resultado_validaciones.get("valor_total", True)
        or not resultado_validaciones.get("concepto", True)
    ):
        return "Pendiente"

    # 🟢 APROBADO
    return "Aprobado"


def ConstruirFilaReporteFinal(
    solped,
    item,
    datos_exp,
    datos_adjuntos,
    datos_me53n,
    datos_texto,
    resultado_validaciones,
):
    """
    Construye una fila para el reporte final consolidado

    Args:
        solped: Número de SOLPED
        item: Número de item
        datos_exp: Dict con datos de expSolped03.txt
        datos_adjuntos: Dict con información de adjuntos
        datos_me53n: Dict con datos de ME53N (TablaSolped)
        datos_texto: Dict con datos extraídos del texto del editor
        resultado_validaciones: Dict con resultados de las validaciones

    Returns:
        Dict con todos los datos para una fila del reporte
    """

    # ========================================================
    # 1. CAMPOS OBLIGATORIOS ME53N
    # ========================================================

    campos_me53n_obligatorios = {
        "Material": datos_me53n.get("Material"),
        "Cantidad": datos_me53n.get("Cantidad"),
        "Precio valoración": datos_me53n.get("PrecioVal."),
        "Fecha entrega": datos_me53n.get("Fe.entrega"),
        "Centro": datos_me53n.get("Centro"),
        "Grupo de compras": datos_me53n.get("GCp"),
        "Organización de compras": datos_me53n.get("OrgC"),
        "Proveedor fijo": datos_me53n.get("ProvFijo"),
    }

    faltantes_me53n = [
        campo
        for campo, valor in campos_me53n_obligatorios.items()
        if valor is None or str(valor).strip() == ""
    ]

    faltantes_me53n_texto = ", ".join(faltantes_me53n)

    # ========================================================
    # 2. CAMPOS OBLIGATORIOS DEL TEXTO
    # ========================================================

    campos_texto_obligatorios = {
        "nit": datos_texto.get("nit"),
        "concepto": datos_texto.get("concepto_compra"),
        "cantidad": datos_texto.get("cantidad"),
        "valor_total": datos_texto.get("valor_total"),
    }

    faltantes_texto = [
        campo
        for campo, valor in campos_texto_obligatorios.items()
        if valor is None or str(valor).strip() == ""
    ]

    faltantes_texto_texto = ", ".join(faltantes_texto)

    # ========================================================
    # 3. NORMALIZAR ADJUNTOS
    # ========================================================

    cant_adj = datos_adjuntos.get("cantidad", 0)
    nombres_adj = datos_adjuntos.get("nombres", "")

    if cant_adj in [None, ""]:
        cant_adj = 0

    if nombres_adj is None:
        nombres_adj = ""

    # ========================================================
    # 4. DETERMINAR ESTADO FINAL
    # ========================================================

    estado_final = determinar_estado_reporte(
        tiene_adjuntos=cant_adj > 0,
        faltantes_me53n=faltantes_me53n,
        faltantes_texto=faltantes_texto,
        resultado_validaciones=resultado_validaciones,
    )
    # ========================================
    # CONSTRUIR FILA DEL REPORTE
    # ========================================
    fila = {
        # ID único del reporte
        "ID_REPORTE": f"{solped}{item}",
        # ========================================
        # DATOS DE expSolped03.txt
        # ========================================
        "PurchReq": datos_exp.get("PurchReq", ""),
        "Item": datos_exp.get("Item", ""),
        "ReqDate": datos_exp.get("ReqDate", ""),
        "Material": datos_exp.get("Material", ""),
        "Created": datos_exp.get("Created", ""),
        "ShortText": datos_exp.get("ShortText", ""),
        "PO": datos_exp.get("PO", ""),
        "Quantity": datos_exp.get("Quantity", 0),
        "Plnt": datos_exp.get("Plnt", ""),
        "PGr": datos_exp.get("PGr", ""),
        "Blank1": datos_exp.get("Blank1", ""),
        "D": datos_exp.get("D", ""),
        "Requisnr": datos_exp.get("Requisnr", ""),
        "ProcState": datos_exp.get("ProcState", ""),
        # ========================================
        # DATOS DE ADJUNTOS
        # ========================================
        "CantAdjuntos": datos_adjuntos.get("cantidad", 0),
        "Nombre de Adjunto": datos_adjuntos.get("nombres", ""),
        # ========================================
        # DATOS DE ME53N (TablaSolped)
        # ========================================
        "Material_ME53N": datos_me53n.get("Material", ""),
        "Short Text_ME53N": datos_me53n.get("Texto breve", ""),
        "Quantity_ME53N": datos_me53n.get("Cantidad", ""),
        "Un": datos_me53n.get("UM", ""),
        "Valn Price": datos_me53n.get("PrecioVal.", ""),
        "Crcy": datos_me53n.get("Mon.", ""),
        "Total Val.": datos_me53n.get("Valor tot.", ""),
        "Deliv.Date": datos_me53n.get("Fe.entrega", ""),
        "Fix. Vend.": datos_me53n.get("ProvFijo", ""),
        "Plant": datos_me53n.get("Centro", ""),
        "PGr_ME53N": datos_me53n.get("GCp", ""),
        "POrg": datos_me53n.get("OrgC", ""),
        "Matl Group": datos_me53n.get("Gpo.artíc.", ""),
        "Id": datos_me53n.get("Pos.", ""),
        # ========================================
        # DATOS EXTRAÍDOS DEL TEXTO
        # ========================================
        "PurchReq_Texto": datos_texto.get("numero_solped", ""),
        "Item_Texto": datos_texto.get("numero_item", ""),
        "Razon Social:": datos_texto.get("razon_social", ""),
        "NIT:": datos_texto.get("nit", ""),
        "Correo:": datos_texto.get("correo", ""),
        "Concepto:": datos_texto.get("concepto_compra", ""),
        "Cantidad:": datos_texto.get("cantidad", ""),
        "Valor Unitario:": datos_texto.get("valor_unitario", ""),
        "Valor Total:": datos_texto.get("valor_total", ""),
        "Responsable:": datos_texto.get("responsable_compra", ""),
        "CECO:": datos_texto.get("ceco", ""),
        # ========================================
        # RESULTADOS DE VALIDACIONES
        # ========================================
        "CAMPOS OBLIGATORIOS FALTANTES ME53N": faltantes_me53n_texto,
        "DATOS EXTRAIDOS DEL TEXTO FALTANTES": faltantes_texto_texto,
        "CANTIDAD": resultado_validaciones.get("cantidad", False),
        "VALOR_UNITARIO": resultado_validaciones.get("valor_unitario", False),
        "VALOR_TOTAL": resultado_validaciones.get("valor_total", False),
        "CONCEPTO": resultado_validaciones.get("concepto", False),
        "Estado": estado_final,
        "Observaciones": resultado_validaciones.get("observaciones", ""),
    }

    return fila


# ======================================================
# GENERADOR DEL EXCEL FINAL
# ======================================================
def GenerarReporteFinalExcel(filas_reporte):
    """
    Genera el archivo Excel con el reporte final consolidado

    Args:
        filas_reporte: Lista de diccionarios, cada uno representa una fila

    Returns:
        Ruta del archivo generado o None si hay error
    """
    try:
        if not filas_reporte:
            print("⚠️ No hay filas para generar el reporte")
            return None

        # Crear DataFrame
        df = pd.DataFrame(filas_reporte)

        # Generar nombre del archivo con timestamp
        timestamp = datetime.now().strftime("%Y%m%d")
        nombre_archivo = f"Reporte_final_consolidado_ME53N_{timestamp}.xlsx"
        ruta_completa = os.path.join(RUTAS["PathReportes"], nombre_archivo)

        # Asegurar que existe el directorio
        os.makedirs(RUTAS["PathReportes"], exist_ok=True)

        # Guardar Excel con formato
        with pd.ExcelWriter(ruta_completa, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Validación ME53N", index=False)

            # Obtener el worksheet para aplicar formato
            worksheet = writer.sheets["Validación ME53N"]

            # Ajustar ancho de columnas
            for idx, col in enumerate(df.columns, 1):
                # Calcular ancho basado en el contenido
                max_length = max(df[col].astype(str).apply(len).max(), len(str(col)))
                # Limitar el ancho máximo
                adjusted_width = min(max_length + 2, 50)
                col_letter = get_column_letter(idx)
                worksheet.column_dimensions[col_letter].width = adjusted_width

            # Congelar primera fila (encabezados)
            worksheet.freeze_panes = "A2"

        print(f"✅ Reporte Excel generado: {ruta_completa}")
        return ruta_completa

    except Exception as e:
        print(f"❌ Error generando reporte Excel: {e}")
        print("ERROR REAL:", e)
        import traceback

        traceback.print_exc()
        return None


def generar_reporte_resumen(filas_reporte):
    """
    Genera un resumen estadístico del reporte

    Args:
        filas_reporte: Lista de diccionarios con las filas del reporte

    Returns:
        Dict con estadísticas del reporte
    """
    if not filas_reporte:
        return {}

    df = pd.DataFrame(filas_reporte)

    resumen = {
        "total_items": len(df),
        "solpeds_unicas": df["PurchReq"].nunique() if "PurchReq" in df.columns else 0,
        "items_por_estado": (
            df["Estado"].value_counts().to_dict() if "Estado" in df.columns else {}
        ),
        "items_con_adjuntos": (
            len(df[df["CantAdjuntos"] > 0]) if "CantAdjuntos" in df.columns else 0
        ),
        "items_sin_adjuntos": (
            len(df[df["CantAdjuntos"] == 0]) if "CantAdjuntos" in df.columns else 0
        ),
    }

    # Validaciones
    if "CANTIDAD" in df.columns:
        resumen["cantidad_ok"] = df["CANTIDAD"].sum()
        resumen["cantidad_error"] = len(df) - df["CANTIDAD"].sum()

    if "VALOR_UNITARIO" in df.columns:
        resumen["valor_unitario_ok"] = df["VALOR_UNITARIO"].sum()
        resumen["valor_unitario_error"] = len(df) - df["VALOR_UNITARIO"].sum()

    if "VALOR_TOTAL" in df.columns:
        resumen["valor_total_ok"] = df["VALOR_TOTAL"].sum()
        resumen["valor_total_error"] = len(df) - df["VALOR_TOTAL"].sum()

    if "CONCEPTO" in df.columns:
        resumen["concepto_ok"] = df["CONCEPTO"].sum()
        resumen["concepto_error"] = len(df) - df["CONCEPTO"].sum()

    return resumen


def imprimir_resumen_reporte(filas_reporte):
    """
    Imprime un resumen del reporte en consola

    Args:
        filas_reporte: Lista de diccionarios con las filas del reporte
    """
    resumen = generar_reporte_resumen(filas_reporte)

    if not resumen:
        print("⚠️ No hay datos para generar resumen")
        return

    print(f"\n{'='*60}")
    print("RESUMEN DEL REPORTE FINAL")
    print(f"{'='*60}")
    print(f"Total items procesados: {resumen['total_items']}")
    print(f"SOLPEDs únicas: {resumen['solpeds_unicas']}")
    print(f"\nItems con adjuntos: {resumen['items_con_adjuntos']}")
    print(f"Items sin adjuntos: {resumen['items_sin_adjuntos']}")

    if "items_por_estado" in resumen and resumen["items_por_estado"]:
        print(f"\nDistribución por estado:")
        for estado, count in resumen["items_por_estado"].items():
            print(f"  {estado}: {count}")

    print(f"\nValidaciones:")
    print(f"  Cantidad OK: {resumen.get('cantidad_ok', 0)}/{resumen['total_items']}")
    print(
        f"  Valor Unitario OK: {resumen.get('valor_unitario_ok', 0)}/{resumen['total_items']}"
    )
    print(
        f"  Valor Total OK: {resumen.get('valor_total_ok', 0)}/{resumen['total_items']}"
    )
    print(f"  Concepto OK: {resumen.get('concepto_ok', 0)}/{resumen['total_items']}")
    print(f"{'='*60}\n")


# =========================================
# FUNCIONES DE UTILIDAD ADICIONALES
# =========================================


def validar_estructura_fila(fila):
    """
    Valida que una fila tenga la estructura mínima requerida

    Args:
        fila: Dict con los datos de la fila

    Returns:
        Tuple (es_valida, mensaje_error)
    """
    campos_requeridos = ["ID_REPORTE", "PurchReq", "Item", "Estado"]

    for campo in campos_requeridos:
        if campo not in fila:
            return False, f"Falta campo requerido: {campo}"
        if not fila[campo]:
            return False, f"Campo requerido vacío: {campo}"

    return True, "OK"


def limpiar_datos_fila(fila):
    """
    Limpia y normaliza los datos de una fila

    Args:
        fila: Dict con los datos de la fila

    Returns:
        Dict con los datos limpios
    """
    fila_limpia = {}

    for key, value in fila.items():
        # Convertir None a string vacío
        if value is None:
            fila_limpia[key] = ""
        # Limpiar strings
        elif isinstance(value, str):
            fila_limpia[key] = value.strip()
        # Mantener otros tipos
        else:
            fila_limpia[key] = value

    return fila_limpia


def exportar_a_csv(filas_reporte, nombre_archivo=None):
    """
    Exporta el reporte a CSV (adicional al Excel)

    Args:
        filas_reporte: Lista de diccionarios con las filas
        nombre_archivo: Nombre del archivo (opcional)

    Returns:
        Ruta del archivo generado o None si hay error
    """
    try:
        if not filas_reporte:
            return None

        df = pd.DataFrame(filas_reporte)

        if nombre_archivo is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"Reporte_final_ME53N_{timestamp}.csv"

        ruta_completa = os.path.join(RUTAS["PathReportes"], nombre_archivo)
        df.to_csv(ruta_completa, index=False, encoding="utf-8-sig")

        print(f"✅ Reporte CSV generado: {ruta_completa}")
        return ruta_completa

    except Exception as e:
        print(f"❌ Error generando CSV: {e}")
        return None

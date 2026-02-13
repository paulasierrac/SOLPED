import pandas as pd
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
from datetime import datetime
from openpyxl.utils import get_column_letter
import os
from Config.InicializarConfig import inConfig
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
    faltantesMe53n: list,
    faltantesTexto: list,
    resultadoValidaciones: dict,
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
        faltantesMe53n
        or faltantesTexto
        or not resultadoValidaciones.get("cantidad", True)
        or not resultadoValidaciones.get("valor_unitario", True)
        or not resultadoValidaciones.get("valor_total", True)
        or not resultadoValidaciones.get("concepto", True)
    ):
        return "Pendiente"

    # 🟢 APROBADO
    return "Aprobado"


def ConstruirFilaReporteFinal(
    solped,
    item,
    datos_exp,
    datosAdjuntos,
    datosMe53n,
    datosTexto,
    resultadoValidaciones,
):
    """
    Construye una fila para el reporte final consolidado

    Args:
        solped: Número de SOLPED
        item: Número de item
        datos_exp: Dict con datos de expSolped03.txt
        datosAdjuntos: Dict con información de adjuntos
        datosMe53n: Dict con datos de ME53N (TablaSolped)
        datosTexto: Dict con datos extraídos del texto del editor
        resultadoValidaciones: Dict con resultados de las validaciones

    Returns:
        Dict con todos los datos para una fila del reporte
    """

    # ========================================================
    # 1. CAMPOS OBLIGATORIOS ME53N
    # ========================================================

    campos_me53n_obligatorios = {
        "Material": datosMe53n.get("Material"),
        "Cantidad": datosMe53n.get("Cantidad"),
        "Precio valoración": datosMe53n.get("PrecioVal."),
        "Fecha entrega": datosMe53n.get("Fe.entrega"),
        "Centro": datosMe53n.get("Centro"),
        "Grupo de compras": datosMe53n.get("GCp"),
        "Organización de compras": datosMe53n.get("OrgC"),
        "Proveedor fijo": datosMe53n.get("ProvFijo"),
    }

    faltantesMe53n = [
        campo
        for campo, valor in campos_me53n_obligatorios.items()
        if valor is None or str(valor).strip() == ""
    ]

    faltantesMe53n_texto = ", ".join(faltantesMe53n)

    # ========================================================
    # 2. CAMPOS OBLIGATORIOS DEL TEXTO
    # ========================================================

    campos_texto_obligatorios = {
        "nit": datosTexto.get("nit"),
        "concepto": datosTexto.get("concepto_compra"),
        "cantidad": datosTexto.get("cantidad"),
        "valor_total": datosTexto.get("valor_total"),
    }

    faltantesTexto = [
        campo
        for campo, valor in campos_texto_obligatorios.items()
        if valor is None or str(valor).strip() == ""
    ]

    faltantesTexto_texto = ", ".join(faltantesTexto)

    # ========================================================
    # 3. NORMALIZAR ADJUNTOS
    # ========================================================

    cant_adj = datosAdjuntos.get("cantidad", 0)
    nombres_adj = datosAdjuntos.get("nombres", "")

    if cant_adj in [None, ""]:
        cant_adj = 0

    if nombres_adj is None:
        nombres_adj = ""

    # ========================================================
    # 4. DETERMINAR ESTADO FINAL
    # ========================================================

    estadoFinal = determinar_estado_reporte(
        tiene_adjuntos=cant_adj > 0,
        faltantesMe53n=faltantesMe53n,
        faltantesTexto=faltantesTexto,
        resultadoValidaciones=resultadoValidaciones,
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
        "CantAdjuntos": datosAdjuntos.get("cantidad", 0),
        "Nombre de Adjunto": datosAdjuntos.get("nombres", ""),
        # ========================================
        # DATOS DE ME53N (TablaSolped)
        # ========================================
        "Material_ME53N": datosMe53n.get("Material", ""),
        "Short Text_ME53N": datosMe53n.get("Texto breve", ""),
        "Quantity_ME53N": datosMe53n.get("Cantidad", ""),
        "Un": datosMe53n.get("UM", ""),
        "Valn Price": datosMe53n.get("PrecioVal.", ""),
        "Crcy": datosMe53n.get("Mon.", ""),
        "Total Val.": datosMe53n.get("Valor tot.", ""),
        "Deliv.Date": datosMe53n.get("Fe.entrega", ""),
        "Fix. Vend.": datosMe53n.get("ProvFijo", ""),
        "Plant": datosMe53n.get("Centro", ""),
        "PGr_ME53N": datosMe53n.get("GCp", ""),
        "POrg": datosMe53n.get("OrgC", ""),
        "Matl Group": datosMe53n.get("Gpo.artíc.", ""),
        "Id": datosMe53n.get("Pos.", ""),
        # ========================================
        # DATOS EXTRAÍDOS DEL TEXTO
        # ========================================
        "PurchReq_Texto": datosTexto.get("numeroSolped", ""),
        "Item_Texto": datosTexto.get("numeroItem", ""),
        "Razon Social:": datosTexto.get("razon_social", ""),
        "NIT:": datosTexto.get("nit", ""),
        "Correo:": datosTexto.get("correo", ""),
        "Concepto:": datosTexto.get("concepto_compra", ""),
        "Cantidad:": datosTexto.get("cantidad", ""),
        "Valor Unitario:": datosTexto.get("valor_unitario", ""),
        "Valor Total:": datosTexto.get("valor_total", ""),
        "Responsable:": datosTexto.get("responsable_compra", ""),
        "CECO:": datosTexto.get("ceco", ""),
        # ========================================
        # RESULTADOS DE VALIDACIONES
        # ========================================
        "CAMPOS OBLIGATORIOS FALTANTES ME53N": faltantesMe53n_texto,
        "DATOS EXTRAIDOS DEL TEXTO FALTANTES": faltantesTexto_texto,
        "CANTIDAD": resultadoValidaciones.get("cantidad", False),
        "VALOR_UNITARIO": resultadoValidaciones.get("valor_unitario", False),
        "VALOR_TOTAL": resultadoValidaciones.get("valor_total", False),
        "CONCEPTO": resultadoValidaciones.get("concepto", False),
        "Estado": estadoFinal,
        "Observaciones": resultadoValidaciones.get("observaciones", ""),
    }

    return fila


# ======================================================
# GENERADOR DEL EXCEL FINAL (ACTUALIZA POR ID_REPORTE)
# ======================================================
def GenerarReporteFinalExcel(filas_reporte):
    """
    Genera o actualiza el archivo Excel con el reporte final consolidado.
    Si el ID_REPORTE ya existe, actualiza la fila.
    Si no existe, la agrega.
    """

    try:
        if not filas_reporte:
            print("No hay filas para generar el reporte")
            return None

        dfNuevo = pd.DataFrame(filas_reporte)

        nombreArchivo = "ReporteFinal.xlsx"
        rutaCompleta = os.path.join(inConfig("PathResultado"), nombreArchivo)

        os.makedirs(inConfig("PathResultado"), exist_ok=True)

        # ==================================================
        # SI EXISTE EL ARCHIVO → ACTUALIZAR REGISTROS
        # ==================================================
        if os.path.exists(rutaCompleta):

            print("Archivo existente encontrado, actualizando registros...")

            dfExistente = pd.read_excel(rutaCompleta)

            # Si no existe columna ID_REPORTE, lo tratamos como nuevo
            if "ID_REPORTE" not in dfExistente.columns:
                dfFinal = pd.concat([dfExistente, dfNuevo], ignore_index=True)

            else:
                dfExistente["ID_REPORTE"] = dfExistente["ID_REPORTE"].astype(str)
                dfNuevo["ID_REPORTE"] = dfNuevo["ID_REPORTE"].astype(str)

                # Convertir existente en dict indexado por ID
                dfExistente = dfExistente.set_index("ID_REPORTE")

                for _, filaNueva in dfNuevo.iterrows():
                    idReporte = filaNueva["ID_REPORTE"]

                    # Si existe → actualizar
                    if idReporte in dfExistente.index:
                        for col in dfNuevo.columns:
                            dfExistente.at[idReporte, col] = filaNueva[col]
                    else:
                        # Si no existe → agregar
                        dfExistente.loc[idReporte] = filaNueva

                dfFinal = dfExistente.reset_index()

        else:
            print("No existe archivo previo, creando nuevo reporte...")
            dfFinal = dfNuevo

        # ==================================================
        # GUARDAR ARCHIVO FINAL
        # ==================================================
        with pd.ExcelWriter(rutaCompleta, engine="openpyxl") as writer:
            dfFinal.to_excel(writer, sheet_name="Validación ME53N", index=False)

            worksheet = writer.sheets["Validación ME53N"]

            # Ajustar ancho columnas
            for idx, col in enumerate(dfFinal.columns, 1):
                longitudMax = max(
                    dfFinal[col].astype(str).apply(len).max(), len(str(col))
                )

                adjusted_width = min(longitudMax + 2, 50)
                col_letter = get_column_letter(idx)
                worksheet.column_dimensions[col_letter].width = adjusted_width

            worksheet.freeze_panes = "A2"

        print(f"Reporte Excel actualizado correctamente: {rutaCompleta}")
        return rutaCompleta

    except Exception as e:
        print(f"Error generando reporte Excel: {e}")
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
        "solpedsUnicas": df["PurchReq"].nunique() if "PurchReq" in df.columns else 0,
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
    print(f"SOLPEDs únicas: {resumen['solpedsUnicas']}")
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


def exportar_a_csv(filas_reporte, nombreArchivo=None):
    """
    Exporta el reporte a CSV (adicional al Excel)

    Args:
        filas_reporte: Lista de diccionarios con las filas
        nombreArchivo: Nombre del archivo (opcional)

    Returns:
        Ruta del archivo generado o None si hay error
    """
    try:
        if not filas_reporte:
            return None

        df = pd.DataFrame(filas_reporte)

        if nombreArchivo is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombreArchivo = f"Reporte_final_ME53N_{timestamp}.csv"

        rutaCompleta = os.path.join(RUTAS["PathReportes"], nombreArchivo)
        df.to_csv(rutaCompleta, index=False, encoding="utf-8-sig")

        print(f"Reporte CSV generado: {rutaCompleta}")
        return rutaCompleta

    except Exception as e:
        print(f"Error generando CSV: {e}")
        return None

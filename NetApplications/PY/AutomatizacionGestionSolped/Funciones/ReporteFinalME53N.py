import pandas as pd
from funciones.EscribirLog import WriteLog
from config.settings import RUTAS
from datetime import datetime

# ======================================================
# COLUMNAS OFICIALES DEL REPORTE FINAL
# ======================================================
COLUMNAS_REPORTE_FINAL = [

    "ID_REPORTE",

    # Datos expSolped03.txt
    "PurchReq", "Item", "ReqDate", "Material", "Created", "ShortText",
    "PO", "Quantity", "Plnt", "PGr", "Blank1", "D", "Requisnr", "ProcState",

    # Adjuntos
    "CantAdjuntos", "Nombre de Adjunto",

    # Datos ME53N
    "Material_ME53N", "Short Text_ME53N", "Quantity_ME53N", "Un",
    "Valn Price", "Crcy", "Total Val.", "Deliv.Date", "Fix. Vend.",
    "Plant", "PGr_ME53N", "POrg", "Matl Group",

    # Texto del ítem
    "Id", "PurchReq_Texto", "Item_Texto", "Razon Social:", "NIT:",
    "Correo:", "Concepto:", "Cantidad:", "Valor Unitario:",
    "Valor Total:", "Responsable:", "CECO:",

    # Resultados validación
    "CAMPOS OBLIGATORIOS FALTANTES ME53N",
    "DATOS EXTRAIDOS DEL TEXTO FALTANTES",
    "CANTIDAD", "VALOR_UNITARIO", "VALOR_TOTAL", "CONCEPTO",
    "Estado", "Observaciones"
]

# ======================================================
# CONSTRUCTOR DE FILA CONSOLIDADA
# ======================================================
def construir_fila_reporte_final(
    solped,
    item,
    datos_exp,
    datos_adjuntos,
    datos_me53n,
    datos_texto,
    resultado_validaciones
):

    fila = {}

    # ID único del reporte
    fila["ID_REPORTE"] = f"{solped}.{item}"

    # ==============================
    # EXP SOLPED
    # ==============================
    for col in [
        "PurchReq", "Item", "ReqDate", "Material", "Created", "ShortText",
        "PO", "Quantity", "Plnt", "PGr", "Blank1", "D", "Requisnr", "ProcState"
    ]:
        fila[col] = datos_exp.get(col, "")

    # ==============================
    # ADJUNTOS
    # ==============================
    fila["CantAdjuntos"] = datos_adjuntos.get("cantidad", 0)
    fila["Nombre de Adjunto"] = datos_adjuntos.get("nombres", "")

    # ==============================
    # ME53N
    # ==============================
    fila.update({
        "Material_ME53N": datos_me53n.get("Material"),
        "Short Text_ME53N": datos_me53n.get("ShortText"),
        "Quantity_ME53N": datos_me53n.get("Quantity"),
        "Un": datos_me53n.get("Unidad"),
        "Valn Price": datos_me53n.get("Precio"),
        "Crcy": datos_me53n.get("Moneda"),
        "Total Val.": datos_me53n.get("Total"),
        "Deliv.Date": datos_me53n.get("FechaEntrega"),
        "Fix. Vend.": datos_me53n.get("ProveedorFijo"),
        "Plant": datos_me53n.get("Centro"),
        "PGr_ME53N": datos_me53n.get("GrupoCompras"),
        "POrg": datos_me53n.get("OrgCompras"),
        "Matl Group": datos_me53n.get("GrupoMaterial"),
    })

    # ==============================
    # TEXTO DEL ITEM
    # ==============================
    for campo in [
        "Id", "PurchReq", "Item", "Razon Social", "NIT",
        "Correo", "Concepto", "Cantidad", "Valor Unitario",
        "Valor Total", "Responsable", "CECO"
    ]:
        fila[f"{campo}:" if campo not in ["Id", "PurchReq", "Item"] else campo] = datos_texto.get(campo, "")

    fila["PurchReq_Texto"] = datos_texto.get("PurchReq", "")
    fila["Item_Texto"] = datos_texto.get("Item", "")

    # ==============================
    # RESULTADOS VALIDACIÓN
    # ==============================
    fila.update({
        "CAMPOS OBLIGATORIOS FALTANTES ME53N": resultado_validaciones.get("faltantes_me53n"),
        "DATOS EXTRAIDOS DEL TEXTO FALTANTES": resultado_validaciones.get("faltantes_texto"),
        "CANTIDAD": resultado_validaciones.get("cantidad"),
        "VALOR_UNITARIO": resultado_validaciones.get("valor_unitario"),
        "VALOR_TOTAL": resultado_validaciones.get("valor_total"),
        "CONCEPTO": resultado_validaciones.get("concepto"),
        "Estado": resultado_validaciones.get("estado"),
        "Observaciones": resultado_validaciones.get("observaciones"),
    })

    return fila


# ======================================================
# GENERADOR DEL EXCEL FINAL
# ======================================================
def generar_reporte_final_excel(filas):
    if not filas:
        return None

    df = pd.DataFrame(filas)
    df = df.reindex(columns=COLUMNAS_REPORTE_FINAL)

    nombre_archivo = f"Reporte_Final_ME53N_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path_salida = f"{RUTAS['PathReportes']}\\{nombre_archivo}"

    df.to_excel(path_salida, index=False)

    WriteLog(
        mensaje=f"Reporte final generado: {path_salida}",
        estado="INFO",
        task_name="ReporteFinalME53N",
        path_log=RUTAS["PathLog"],
    )

    return path_salida

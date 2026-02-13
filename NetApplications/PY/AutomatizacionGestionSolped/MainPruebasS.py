import pandas as pd
import os

from Funciones.EscribirLog import WriteLog
from Funciones.EmailSender import EnviarNotificacionCorreo

# from Funciones.GeneralME53N import AppendHipervinculoObservaciones

from Config.settings import RUTAS, SAP_CONFIG

# from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import ConectarSAP, ObtenerSesionActiva

# from HU.HU02_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03

# from HU.HU04_GeneracionOC import EjecutarHU04
# from HU.HU05_DescargaOC import EjecutarHU05

from Config.InicializarConfig import inConfig, initConfig
from Funciones.FuncionesExcel import ServicioExcel


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
    # EXTRAER NÚMERO DESDE NOMBRE TXT
    # ===============================
    nombre_archivo = os.path.basename(ruta_txt)

    numeroProcstate = "".join(filter(str.isdigit, nombre_archivo))

    if not numeroProcstate:
        raise ValueError("No se pudo extraer número del nombre del archivo.")

    numeroProcstate = numeroProcstate.zfill(2)
    # ===============================
    # 4. AGREGAR ProcState = "03"
    # ===============================
    df["ProcState"] = numeroProcstate

    # ===============================
    # 5. ORDEN FINAL
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
        "D",
        "Requisnr",
        "ProcState",
    ]

    df = df[[col for col in ordenFinal if col in df.columns]]

    # ===============================
    # 6. GUARDAR ARCHIVO TRANSFORMADO
    # ===============================
    # rutaSalida = ruta_txt.replace(".txt", "_transformado.txt")
    rutaSalida = ruta_txt

    anchos = {
        col: max(df[col].astype(str).map(len).max(), len(col)) + 3 for col in df.columns
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


if __name__ == "__main__":
    # MainSantiago()
    nombreTarea = "MainPruebasS"
    initConfig()
    # session = ConectarSAP(
    #     inConfig("SapSistema"),
    #     inConfig("SapMandante"),
    #     SAP_CONFIG["user"],
    #     SAP_CONFIG["password"],
    # )
    pathReporte = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Resultado\Reportes\ReporteFinal.xlsx"

    # Sube el Excel a la base de datos
    ServicioExcel.ejecutarBulkDesdeExcel(pathReporte)
    # TransformartxtMe5a(
    #     r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped05.txt"
    # )

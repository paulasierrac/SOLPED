from HU.HU01_LoginSAP import ObtenerSesionActiva
from HU.HU01_LoginSAP import ConectarSAP, ObtenerSesionActiva
from Config.InicializarConfig import inConfig
from Config.settings import RUTAS, SAP_CONFIG
import pandas as pd
import os


def transformar_txt_me5a(ruta_txt: str):
    """
    Transforma el TXT exportado de SAP ME5A al formato estándar requerido.
    - Elimina columnas innecesarias
    - Renombra columnas a inglés
    - Agrega columna ProcState con valor fijo "03"
    """

    if not os.path.exists(ruta_txt):
        raise FileNotFoundError(f"No existe el archivo: {ruta_txt}")

    # ===============================
    # 1. LEER Y LIMPIAR ARCHIVO
    # ===============================
    with open(ruta_txt, "r", encoding="utf-8") as f:
        lineas = f.readlines()

    lineas_validas = [
        linea.strip()
        for linea in lineas
        if linea.strip().startswith("|") and not set(linea.strip()) == {"-"}
    ]

    columnas = [col.strip() for col in lineas_validas[0].split("|")[1:-1]]

    datos = []
    for linea in lineas_validas[1:]:
        valores = [v.strip() for v in linea.split("|")[1:-1]]
        if len(valores) == len(columnas):
            datos.append(valores)

    df = pd.DataFrame(datos, columns=columnas)

    # ===============================
    # 2. ELIMINAR COLUMNAS NO NECESARIAS
    # ===============================
    columnas_eliminar = ["CDoc", "EL", "ProvFijo", "PrecioVal.", "Valor tot.", "Fondo"]

    df = df.drop(
        columns=[c for c in columnas_eliminar if c in df.columns], errors="ignore"
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
    # 4. AGREGAR ProcState = "03"
    # ===============================
    df["ProcState"] = "03"

    # ===============================
    # 5. ORDEN FINAL
    # ===============================
    orden_final = [
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

    df = df[[col for col in orden_final if col in df.columns]]

    # ===============================
    # 6. GUARDAR SOBRE EL MISMO ARCHIVO
    # ===============================
    ruta_salida = ruta_txt.replace(".txt", "_transformado.txt")

    with open(ruta_salida, "w", encoding="utf-8") as f:
        ancho_total = sum(
            [
                max(df[col].astype(str).map(len).max(), len(col)) + 3
                for col in df.columns
            ]
        )
        f.write("-" * ancho_total + "\n")

        encabezado = "|"
        for col in df.columns:
            encabezado += f"{col:<15}|"
        f.write(encabezado + "\n")

        f.write("-" * ancho_total + "\n")

        for _, row in df.iterrows():
            linea = "|"
            for col in df.columns:
                linea += f"{str(row[col]):<15}|"
            f.write(linea + "\n")

        f.write("-" * ancho_total + "\n")

    print(f"✅ Archivo transformado correctamente: {ruta_salida}")
    return ruta_salida


def MainSantiago():
    try:
        session = ObtenerSesionActiva()
        session = ConectarSAP(
            inConfig("SapSistema"),
            inConfig("SapMandante"),
            SAP_CONFIG["user"],
            SAP_CONFIG["password"],
        )

    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise


if __name__ == "__main__":
    MainSantiago()
    transformar_txt_me5a(
        r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped03.txt"
    )

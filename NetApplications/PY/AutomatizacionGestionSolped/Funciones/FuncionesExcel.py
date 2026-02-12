
#Funciones.FuncionesExcel.py

import os
import csv
import re
import unicodedata
import warnings
import pandas as pd

from Config.InicializarConfig import inConfig
from Repositories.Excel import ExcelRepo


class ServicioExcel:

    @staticmethod
    def limpiarExcel(
        rutaEntrada: str,
        columnasMapeo: dict,
        hoja: str | int = 0,
        header: int = 0
    ) -> str:
        """
        Limpia un archivo Excel dejando solo las columnas requeridas,
        las renombra y guarda el resultado en un archivo nuevo.

        :param rutaEntrada: Ruta del archivo Excel original
        :param rutaSalida: Ruta donde se guardará el Excel limpio
        :param columnasMapeo: Diccionario {col_original: col_final}
        :param hoja: Nombre o índice de la hoja (default 0)
        """
        nombreArchivo = os.path.splitext(os.path.basename(rutaEntrada))[0]

        # Leer archivo
        df = pd.read_excel(rutaEntrada, sheet_name=hoja, header=header)

        # Normalizar nombres de columnas
        df.columns = [ServicioExcel.normalizacionColumna(c) for c in df.columns]

        # Columnas que realmente existen en el Excel
        columnasExistentes = [
            col for col in columnasMapeo.keys() if col in df.columns
        ]

        # Advertencia si faltan columnas
        columnasFaltantes = set(columnasMapeo.keys()) - set(columnasExistentes)
        if columnasFaltantes:
            print(f"Advertencia: columnas faltantes -> {columnasFaltantes}")

        # Filtrar columnas necesarias
        dfLimpio = df[columnasExistentes]

        # Renombrar columnas
        dfLimpio = dfLimpio.rename(columns=columnasMapeo)

        # Guardar archivo limpio
        
        carpetaTemp = inConfig("PathTemp")
        rutaSalida = os.path.join(carpetaTemp, f"{nombreArchivo}limpio.xlsx")
        dfLimpio.to_excel(rutaSalida, index=False)

        print(f"Archivo limpio generado correctamente en: {rutaSalida}")

        return rutaSalida

    # -----------------------------
    # NORMALIZACIÓN DE NOMBRES
    # -----------------------------
    @staticmethod
    def normalizacionColumna(nombre: str) -> str:
        nombre = nombre.strip().lower()
        nombre = unicodedata.normalize('NFKD', nombre).encode('ascii', 'ignore').decode()
        nombre = re.sub(r'[^\w]', '_', nombre)
        nombre = re.sub(r'_+', '_', nombre)
        return nombre

    # -----------------------------
    # LIMPIEZA DE DATOS
    # -----------------------------
    @staticmethod
    def limpiarTexto(valor):
        if pd.isna(valor):
            return ""
        valor = str(valor)
        valor = unicodedata.normalize("NFKC", valor)
        valor = re.sub(r"[\x00-\x1F\x7F]", " ", valor)
        return valor.replace("\n", " ").replace("\r", " ").strip()

    @staticmethod
    def sanitizeText(value: str) -> str:
        if value is None:
            return "NULL"
        value = unicodedata.normalize("NFKC", str(value))
        value = re.sub(r"[\x00-\x1F\x7F]", " ", value)
        value = value.replace('"', "").strip()
        return value if value else "NULL"

    # -----------------------------
    # OBTENER COLUMNAS DEL EXCEL
    # -----------------------------
    @staticmethod
    def obtenerColumnasExcel(rutaExcel: str, header: int = 0) -> list[str]:
        df = pd.read_excel(
            rutaExcel,
            header=header,
            nrows=0,
            engine="openpyxl"
        )
        return [ServicioExcel.normalizacionColumna(c) for c in df.columns]
    
    
    # -----------------------------
    # EXCEL → CSV
    # -----------------------------
    @staticmethod
    def excelACSV(rutaExcel: str, header: int = 0) -> tuple[str, list]:

        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
        try:
            df = pd.read_excel(
                rutaExcel,
                header=header,
                dtype=str,
                engine="openpyxl"
            )

            df.columns = [ServicioExcel.normalizacionColumna(c) for c in df.columns]
            ordenColumnas = list(df.columns)

            df = df.map(ServicioExcel.limpiarTexto)

            nombreBase = os.path.splitext(os.path.basename(rutaExcel))[0]
            carpetaTemp = inConfig("PathTemp")
            rutaCSV = os.path.join(carpetaTemp, f"{nombreBase}.csv")

            df.to_csv(
                rutaCSV,
                sep=";",
                index=False,
                encoding="utf-8-sig"
            )
            return rutaCSV, ordenColumnas
        except Exception as e:
            print(f"Error convirtiendo Excel a CSV: {e}")
            raise

    # -----------------------------
    # CSV → TXT
    # -----------------------------
    @staticmethod
    def convertirTXT(csv_path: str) -> str:

        txt_path = os.path.splitext(csv_path)[0] + ".txt"

        with open(csv_path, "r", encoding="latin1", newline="") as csv_file, \
             open(txt_path, "w", encoding="utf-8", newline="\n") as txt_file:

            reader = csv.reader(csv_file)

            for row in reader:
                cleaned = [ServicioExcel.sanitizeText(v) for v in row]

                if all(v == "NULL" for v in cleaned):
                    continue

                txt_file.write(";".join(cleaned) + "\n")

        return txt_path

    # -----------------------------
    # ORQUESTADOR FINAL
    # -----------------------------
    @staticmethod
    def ejecutarBulkDesdeExcel(rutaExcel: str, header: int = 0):
        """
        Punto único de entrada:
        - El Excel define columnas
        - El nombre del archivo define la tabla
        """

        nombreTabla = ServicioExcel.normalizacionColumna(
            os.path.splitext(os.path.basename(rutaExcel))[0]
        )

        rutaCSV = None
        rutaTXT = None

        try:
            # 1. Columnas (orden exacto del Excel)
            ordenColumnas = ServicioExcel.obtenerColumnasExcel(rutaExcel, header)

            # 2. Excel → CSV
            rutaCSV, ordenColumnas = ServicioExcel.excelACSV(rutaExcel, header)

            # 3. CSV → TXT
            rutaTXT = ServicioExcel.convertirTXT(rutaCSV)

            # 4. Bulk con tabla temporal + final
            VariableExcelRepo = ExcelRepo(schema="GestionSolped")
            VariableExcelRepo.ejecutarBulkDinamico(
                rutaTXT=rutaTXT,
                tabla=nombreTabla,
                columnas=ordenColumnas
            )

        except Exception as e:
            print(f"Error ejecutando bulk desde Excel: {e}")
            raise

        finally:
            for f in (rutaCSV, rutaTXT):
                if f and os.path.exists(f):
                    os.remove(f)

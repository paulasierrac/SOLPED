from config.init_config import in_config
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from funciones.FuncionesExcel import ExcelService
import pandas as pd


def Prueba():
    EjecutarHU00()
    ruta_excel = r"\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped03.xlsx"
    df = pd.read_excel(ruta_excel)
    print(df)

    ExcelService.ejecutar_bulk_desde_excel(ruta_excel)


if __name__ == "__main__":
    Prueba()

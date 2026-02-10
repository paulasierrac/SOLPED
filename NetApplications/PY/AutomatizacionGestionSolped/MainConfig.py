from config.init_config import in_config
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from funciones.FuncionesExcel import ExcelService
import pandas as pd

def Prueba():
    EjecutarHU00()
    ruta_excel = r"\\192.168.50.169\RPA_SAMIR_GestionSolped\Temp\OC_Liberadas.xlsx"

    df = pd.read_excel(ruta_excel)
    print(df)

    ExcelService.excel_a_csv(ruta_excel, 0)

    ExcelService.ejecutar_bulk_desde_excel(ruta_excel)



if __name__ == "__main__":
    Prueba()
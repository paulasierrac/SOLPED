from Config.InicializarConfig import inConfig
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from Funciones.FuncionesExcel import ServicioExcel
import pandas as pd

def Prueba():
    EjecutarHU00()
    rutaExcel = r"\\192.168.50.169\RPA_SAMIR_GestionSolped\Temp\OC_Liberadas.xlsx"

    df = pd.read_excel(rutaExcel)
    print(df)

    ServicioExcel.excelACSV(rutaExcel, 0)

    ServicioExcel.ejecutarBulkDesdeExcel(rutaExcel)



if __name__ == "__main__":
    Prueba()
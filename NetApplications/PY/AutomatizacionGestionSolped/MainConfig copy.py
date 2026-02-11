from Config.init_config import inConfig
from HU.HU00_DespliegueAmbiente import EjecutarHU00


def Prueba():
    EjecutarHU00()

    parametro_ejemplo = inConfig("SAP_LOGON_PATH")
    print(f"El valor del parametro 'SAP_LOGON_PATH' es: {parametro_ejemplo}")


if __name__ == "__main__":
    Prueba()

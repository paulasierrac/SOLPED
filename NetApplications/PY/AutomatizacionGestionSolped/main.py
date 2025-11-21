# main_conectar_sap.py
from HU.HU1_LoginSAP import conectar_sap , abrir_sap_logon,obtener_sesion_activa
from HU.HU2_DescargaME5A import descarga_solpedME5A
import time


if __name__ == "__main__":
 
    session = obtener_sesion_activa()

    if session:
        print("Conexion establecida, listo para ejecutar transacciones.")
    else:
        print("No se pudo establecer la conexión.")
    
    estado = "03"
    descarga_solpedME5A(session,estado)
    estado = "05"
    descarga_solpedME5A(session,estado)    
# main_conectar_sap.py
from Funciones.conectar_sap import conectar_sap , abrir_sap_logon,obtener_sesion_activa,descarga_solpedME5A
from Funciones.pruebas  import obtener_secreto_keyvault
import time


if __name__ == "__main__":

    conexion = "ERP-CORPORATIVO-PRODUCCION"
    mandante = "410"
    #usuario = obtener_secreto_keyvault (secret_name="SAP-Usuario")
    usuario = "CGRPA065"
    #password = obtener_secreto_keyvault (secret_name="SAP-Pass")
    password = "sT1f%4L*" 

    idioma = "ES"

    print(usuario)
    print(password)

    abrir_sap = abrir_sap_logon()
    if abrir_sap:
        print(" SAP Logon 750 se encuentra abierto")
    else:
        print(" SAP Logon 750 abierto ")

    #session = conectar_sap(conexion, mandante, usuario, password, idioma)
    session = obtener_sesion_activa()

    if session:
        print("Conexion establecida, listo para ejecutar transacciones.")
    else:
        print("No se pudo establecer la conexión.")
    
    transaccion="/nME5A"
     
    descarga_solpedME5A(session,transaccion)
        
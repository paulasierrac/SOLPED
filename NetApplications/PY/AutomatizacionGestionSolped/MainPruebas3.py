# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from requests import session
from HU.HU01_LoginSAP import ObtenerSesionActiva
from Funciones.ValidacionM21N import obtener_numero_oc,select_GuiTab
from Funciones.GuiShellFunciones import SapTextEditor
from Funciones.GeneralME53N import AbrirTransaccion
import re

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG



def main():
    session = ObtenerSesionActiva()
    if not session:
        return


    try:

        select_GuiTab(session, "TABHDT9") 
                 

    
    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        print("Asegúrate de que estás en la pantalla correcta de SAP y el ID del editor es el correcto.")
if __name__ == "__main__":
    main()

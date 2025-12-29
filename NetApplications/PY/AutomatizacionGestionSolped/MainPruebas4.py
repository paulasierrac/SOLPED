# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from requests import session
from HU.HU01_LoginSAP import ObtenerSesionActiva,conectar_sap
from Funciones.ValidacionM21N import SapTextEditor,find_sap_object_in_usr
from Funciones.GeneralME53N import AbrirTransaccion
import pyautogui  # Asegúrate de tener pyautogui instalado
import time

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


import re
from typing import List, Optional



def Main_Pruebas4():
    try:

        session = ObtenerSesionActiva()

        # pruebas para lograr borrar textos desde la solped atraves del objeto session
        # TODO: la idea es que traiga el texto y valide con un if si tiene contenido (despues del F2) y si tiene contenido lo borre
        """
        session.findById(
            "wnd[0]/usr/
            "subSUB0:SAPLMEGUI:0010/" \
            "subSUB3:SAPLMEVIEWS:1100/" \
            "subSUB2:SAPLMEVIEWS:1200/" \
            "subSUB1:SAPLMEGUI:1301/" \
            "subSUB2:SAPLMEGUI:1303/" \
            "tabsITEM_DETAIL/tabpTABIDT14/
            "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/
            "subTEXTS:SAPLMMTE:0200/" \
            "cntlTEXT_TYPES_0200/shell").selectedNode = "F01"
        session.findById(
            "wnd[0]/usr/" \
            "subSUB0:SAPLMEGUI:0010/" \
            "subSUB3:SAPLMEVIEWS:1100/" \
            "subSUB2:SAPLMEVIEWS:1200/" \
            "subSUB1:SAPLMEGUI:1301/" \
            "subSUB2:SAPLMEGUI:1303/" \
            "tabsITEM_DETAIL/tabpTABIDT14/" \
            "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/" \
            "subTEXTS:SAPLMMTE:0200/" \
            "subEDITOR:SAPLMMTE:0201/" \
            "cntlTEXT_EDITOR_0201/shellcont/shell"
            ).setSelectionIndexes 125,125
        session.findById(
            "wnd[0]/usr/" \
            "subSUB0:SAPLMEGUI:0010/" \
            "subSUB3:SAPLMEVIEWS:1100/" \
            "subSUB2:SAPLMEVIEWS:1200/" \
            "subSUB1:SAPLMEGUI:1301/" \
            "subSUB2:SAPLMEGUI:1303/" \
            "tabsITEM_DETAIL/tabpTABIDT14/" \
            "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/" \
            "subTEXTS:SAPLMMTE:0200/" \
            "cntlTEXT_TYPES_0200/shell").selectedNode = "F02"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").setSelectionIndexes 47,47
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F03"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").setSelectionIndexes 0,0
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F04"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F05"

        """
        textoPosicionF = (
            "wnd[0]/usr/"
            "subSUB0:SAPLMEGUI:0010/" \
            "subSUB3:SAPLMEVIEWS:1100/" \
            "subSUB2:SAPLMEVIEWS:1200/" \
            "subSUB1:SAPLMEGUI:1301/" \
            "subSUB2:SAPLMEGUI:1303/" \
            "tabsITEM_DETAIL/tabpTABIDT14/"
            "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/" \
            "cntlTEXT_TYPES_0200/shell"
        )
        EDITOR_ID = (
            "wnd[0]/usr/"
            "subSUB0:SAPLMEGUI:0010/"
            "subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/"
            "subSUB2:SAPLMEGUI:1303/"
            "tabsITEM_DETAIL/tabpTABIDT14/"
            "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/"
            "subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )
        
        for i in range(2, 6):  # F02 a F05
            textPF = session.findById(textoPosicionF)
            nodo = f"F0{i}"
            textPF.selectedNode = nodo
            editxt = session.findById(EDITOR_ID)
            editor = SapTextEditor(session, EDITOR_ID)
            texto = editor.get_all_text()
            if texto :
                print("El texto no está vacío. Procediendo a borrarlo...")
                editxt.SetUnprotectedTextPart(0,".")
                print("Texto borrado exitosamente.")
    

    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

if __name__ == "__main__":
    Main_Pruebas4()



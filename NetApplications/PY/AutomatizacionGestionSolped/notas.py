# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU1_LoginSAP import ObtenerSesionActiva,conectar_sap,abrir_sap_logon
from HU.HU2_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU5_GeneracionOC import GenerarOCDesdeSolped
from HU.HU4_DescargaOCME9F import descarga_OCME9F

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Notas():
     # textoPocision3=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI"
    #                  ":1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
    #                  "subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F02"
    # textoPocision3.selectedNode = "F02"

    # Elimino el texto y lo repmplazo por un punto


    #for i in range(1):
    #foco Select a campo textos de cada pocision 
    session = ObtenerSesionActiva()

    # foco Select en la pestaña textos despues de imputacion 
    textoPestana = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:11"
                     "00/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:"
                     "SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14")
    textoPestana.select()
    
    textoPocision = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/"
                    "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI"
                    ":1303/tabsITEM_DETAIL/tabpTABIDT14")
    textoPocision.select()
    textoPocision1 = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS"
                     ":1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14"
                     "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE"
                     ":0201/cntlTEXT_EDITOR_0201/shellcont/shell")
    textoPocision1.setSelectionIndexes (0,0)
    # foco Select a campo texto Pedido de info
    textoPedidodeinfo=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                    "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/"
                    "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")
    textoPedidodeinfo.selectedNode = "F05"
    # Boton borrar texto
    botonBorrar=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1"
                    ":SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB"
                    ":SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201")
    botonBorrar.press()
    # Punto en el texto
    puntoentexto = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1"
                                    ":SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:"
                                    "SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")
    puntoentexto.text = "."

  
    # paso al siguiiente item de la solped
    botonabajo=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301"
                    "/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE"
                    ":0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")
    
    botonabajo.setSelectionIndexes (1,1)

    botonabajoPress=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2"
                                    ":SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002")
    botonabajoPress.press()


      #     def find_combo_dyn_6000(session):
      #     raiz = session.findById("wnd[0]/usr")
      
      #     def buscar(root):
      #         try:
      #             for child in root.Children:
      #                 # Condición: es un combo y su Id termina en ese sufijo
      #                 if (child.Type == "GuiComboBox" and
      #                     child.Id.endswith("subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST")):
      #                     return child
      
      #                 # Buscar recursivamente en hijos
      #                 result = buscar(child)
      #                 if result:
      #                     return result
      #         except:
      #             # Algunos objetos no tienen Children, ignoramos el error
      #             pass
      #         return None
      
      #     return buscar(raiz)
      
      # combo = find_combo_dyn_6000(session)
      
      # if combo:
      #     print("Encontrado:", combo.Id)
      #     # Ejemplo: seleccionar primer elemento
      #     if combo.Entries.Count > 0:
      #         combo.Key = combo.Entries(0).Key
      #         combo.setFocus()
      # else:
      #     print("No se encontró el combo DYN_6000-LIST")
  
    """
    time.sleep(1)
    #session.findById("wnd[0]").sendVKey(0)
    pyautogui.press("enter")
    time.sleep(1)

    # Navegar hasta el campo Solicitudes de pedido
    for i in range(6):   
      pyautogui.hotkey("down")
      #time.sleep(0.5)
    
    
    pyautogui.press("enter")
    time.sleep(0.5)
    # ingresa el numero de la solped que va a revisar 
    session.findById("wnd[0]/usr/ctxtSP$00026-LOW").text = solped
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    for i in range(2):  
      pyautogui.hotkey("shift", "TAB")
    pyautogui.hotkey("TAB")
    pyautogui.press("enter")
      #time.sleep(0.5)
   """ 



   # if "RSYST-BNAME" in session.findById("wnd[0]/usr").Text:
   #      print("🧩 Ingresando credenciales...")
   # if password is None:
   #    password = getpass.getpass("Contraseña SAP: ")
    #editor = session.findById("wnd[0]/tbar[1]/btn[9]")
    #editor.setFocus()
    #editor.press()  # Botón Resumen de documento no activo 
    #
    #editor.setFocus()
    #editor.press()  # Botón Resumen de documento no activo 
    #session.findById("wnd[0]/tbar[1]/btn[9]").press()  # Botón Resumen de documento no activo 
    #time.sleep(2)
    #session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").press("SELECT")
    #time.sleep(5)
    #session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem("8265D72160021FD0B2F42CA052588245NEW:REQ_QUERY")


    # entrada manual mientras se soluciona el tema del menu contextual
    
    #time.sleep(5)   
    #
   #  session.findById("wnd[0]/usr/ctxtSP$00026-LOW").setFocus()
   #  session.findById("wnd[0]/usr/ctxtSP$00026-LOW").caretPosition = (10)
   #  session.findById("wnd[0]").sendVKey (0)
   #  session.findById("wnd[0]/tbar[1]/btn[8]").press()
   #  time.sleep(5)
   #  session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").expandNode ("          1")
   #  session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").topNode = ("          1")

    # ============================
    # Ingresar datos de la Solped
    # ============================
   


      #    editor = session.findById(
      #       "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
      #       "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
      #       "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
      #       "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
      #       "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
      #       "cntlTEXT_EDITOR_0201/shellcont/shell"
      #   )
 
      #   # 2) Asegurar que el editor tiene el foco
      #   editor.SetFocus()
      #   time.sleep(0.5)
 
      #   # 3) Seleccionar TODO el texto
      #   pyautogui.hotkey("ctrl", "a")
      #   time.sleep(0.3)
 
      #   # 4) Copiar al portapapeles
      #   pyautogui.hotkey("ctrl", "c")
      #   time.sleep(0.5)
 
      #   # 5) Obtener texto del portapapeles con codificacion correcta
      #   texto_completo = ObtenerTextoDelPortapapeles()




if __name__ == "__main__":
    Notas()




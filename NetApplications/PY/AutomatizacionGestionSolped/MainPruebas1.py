# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from Funciones.ValidacionM21N import SapTextEditor, get_GuiTextField_text, press_GuiButton, set_GuiTextField_text,normalizar_precio_sap, validar_y_ajustar_solped,debug_sap_object
from HU.HU01_LoginSAP import ObtenerSesionActiva,conectar_sap,abrir_sap_logon

import re
from typing import List, Optional



def Main_Pruebas1():
    try:
        session = ObtenerSesionActiva()
        #validar_y_ajustar_solped(session, 7)
        textop=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/" \
        "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" \
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/" \
        "tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/" \
        "subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")
        """
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/" \
        "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" \
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/" \
        "tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/" \
        "subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F03"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F04"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F05"
        """
        
        print(type(textop))
        print(textop.Type)

        textop.selectedNode = "F03"
        debug_sap_object(textop, "GuiShell")


  
     
    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise


if __name__ == "__main__":
    Main_Pruebas1()




"""
        for i in range(item):
                       
            obj_btnDel = None
            selectsFs = [2,3,4,5]
            obj_tabstrip = ejecutar_creacion_hijo(session)

            if obj_tabstrip:
                nombre_pestaña_buscada = "Textos" # O "Invoice", "Entregas", etc.
                pestaña_encontrada = False
                for pestaña in obj_tabstrip.Children:
                    # pestaña.Text te da el nombre visible (ej: "Condiciones")
                    # pestaña.Name te da el ID técnico (ej: "TABIDT3")
                    if pestaña.Text == nombre_pestaña_buscada:
                        pestaña_encontrada = True

                        print(f"Pestaña '{nombre_pestaña_buscada}' seleccionada. (ID Técnico: {pestaña.Name})")
                        full_id_btnDel = limpiar_id_sap(pestaña.Id) + ruta_restante_btnDel
                        full_id_textoposicion = limpiar_id_sap(pestaña.Id) + ruta_restante_textoposicion
                        full_id_textoarea = limpiar_id_sap(pestaña.Id) + ruta_restante_textoarea
                        time.sleep(2)
                        pestaña.Select()

                        for i in selectsFs:
                            F0n = "F0" + str(i)
                        
                            # .selectedNode = "F02" Texto pedido de posicion   
                            obj_textoposicion = session.findById(full_id_textoposicion)
                            print(f"Texto posicion  '{obj_textoposicion.Id}' seleccionada. (ID Técnico: {obj_textoposicion.Name})")
                            obj_textoposicion.selectedNode = F0n
                            time.sleep(2)
                            #Boton Eliminar 
                            try:
                                obj_btnDel = session.findById(full_id_btnDel)
                                print(f"Bot+on Delete '{obj_btnDel.Id}' seleccionada. (ID Técnico: {obj_btnDel.Name})")
                                obj_btnDel.Press()
                                time.sleep(1)

                                # entrar a editar texto "."
                                obj_textoarea = session.findById(full_id_textoarea)
                                obj_textoarea.text = "."
                            except:
                                pass
                        time.sleep(20)   
                        ruta=rf".\img\abajo.png"
                        buscar_y_clickear(ruta, confidence=0.8, intentos=20, espera=0.5)
                        print("Preparando siguiente iteración...")
                        if not pestaña_encontrada:         print(...)
                        
                        break

                if not pestaña_encontrada:
                    print(f"No se encontró la pestaña llamada {nombre_pestaña_buscada}")
"""
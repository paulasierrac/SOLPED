from HU.HU01_LoginSAP import ObtenerSesionActiva
from Funciones.GuiShellFunciones import find_sap_object




def MainSantiago():
    try:
        session = ObtenerSesionActiva()

        # Scroll = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/" \
        # "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")

        # # Todo: Stev: bucle para revisar visibles en el grid de posiciones
        # filas_visibles = Scroll.VisibleRowCount
        # print("Filas visibles:", filas_visibles)
        
        # fila = 7  # Ejemplo de fila a seleccionar
       
        # Scroll = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/" \
        # "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")
        # Scroll.verticalScrollbar.position = 15

        varianteSeleccion = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]")
        

        print(type(varianteSeleccion))
        print(varianteSeleccion.Type)
        res = find_sap_object(varianteSeleccion, "GuiShell")
        print(type(res))
        print(res.Type)
        

        #varianteSeleccion.press()

        

    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

if __name__ == "__main__":
    MainSantiago()
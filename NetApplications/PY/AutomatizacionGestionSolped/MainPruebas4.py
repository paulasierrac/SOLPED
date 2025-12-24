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
from Funciones.ValidacionM21N import SapTextEditor
from Funciones.GeneralME53N import AbrirTransaccion
import pyautogui  # Asegúrate de tener pyautogui instalado
import time

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Main_Pruebas4():
    try:

        #session = conectar_sap( SAP_CONFIG["sistema"], SAP_CONFIG["mandante"],SAP_CONFIG["user"], SAP_CONFIG["password"], SAP_CONFIG["idioma"] )
        session = ObtenerSesionActiva()
        #AbrirTransaccion(session, "ME21N")
        # codigo para pruebas
        print(session)

        EDITOR_ID2=(
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
            "cntlTEXT_EDITOR_0201/shellcont/shell")


        """
        #Obtiene los valores de los campos de precio en la tabla de posiciones
        for fila in range(4):
            precio = get_GuiTextField_text(session, f"NETPR[10,{fila}]")
            print(f"Fila {fila+1}: {precio}")
        """
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

        editor_shell = session.findById(EDITOR_ID)
   
        print("ID encontrado:", editor_shell.Id)
        print("Tipo:", editor_shell.Type)
  
        editor = SapTextEditor(session, EDITOR_ID)

        print("Tipo:", type(editor))
        print("Texto original:",flush=True)
        texto = editor.get_all_text()
        print(texto,flush=True)


        reemplazos = {
                "VENTA SERVICIO": "V1",
                "VENTA PRODUCTO": "V1",
                "GASTO PROPIO SERVICIO": "C2",
                "GASTO PROPIO PRODUCTO": "C2",
                "SAA": "R3", #"SAA SERVICIO": "R3"
                "SAA PRODUCTO": "R3",
                "DAVIVIENDA": "HOLA MUNDO",
            }


        nuevo_texto, cambios = editor.replace_in_text(texto, reemplazos)

     
        print(type(nuevo_texto))
        

        print(f"texto modificado {nuevo_texto}")
        print(f"cambios realizados {cambios}") 

        #editor.set_text( nuevo_texto )  # Usa el método
        editor_shell.SetUnprotectedTextPart(0,"**//TEXTO MODIFICADO POR BOT RESOC//**")
        editor_shell.SetUnprotectedTextPart(1,nuevo_texto)

              
        #editor_shell.text = nuevo_texto
       
        """
        def normalizar(texto):
            return texto.replace("\r", "").replace("\xa0", " ")
        texto_norm = normalizar(texto)
        reemplazos = {
                "VENTA SERVICIO": "V1",
                "VENTA PRODUCTO": "V1",
                "GASTO PROPIO SERVICIO": "C2",
                "GASTO PROPIO PRODUCTO": "C2",
                "\nSAA\n": "\nR3\n",
            }
        nuevo_texto = texto_norm
        for b, r in reemplazos.items():
            nuevo_texto = nuevo_texto.replace(b, r)

        print(f"Nuevo texto: {nuevo_texto}") 
            
        print(f"Líneas modificadas: {cambios}")

        # print("Texto final:",flush=True)
        # texto = editor.get_all_text()
        # print(texto,flush=True)

        editor = SapTextEditor(session, EDITOR_ID)

        cambio = editor.reemplazar_linea_exacta("SAA", "R3")
        

        if cambio:
            print("Línea SAA reemplazada por R3")
        else:
            print("No se encontró la línea SAA")

        # def set_all_text(self, texto):
        #     self.shell.SetFocus()
        #     self.shell.SelectAll()
        #     self.shell.SetUnprotectedTextPart(texto)   

        print("Esperando antes de la siguiente modificación...")

        editor.set_all_text(editor.get_all_text().replace("\nDAVIVIENDA\n", "\nRWWWWWW\n"))


        if nuevo_texto != texto_norm:
            editor.set_all_text(nuevo_texto)

        """

      
                

    
  
        #grid.selectContextMenuItem("8265D72160021FD0B6F43226BAE842F8NEW:REQ_QUERY")
         #                          "8265D72160021FD0B6F5D4A4306A42D8NEW:REQ_QUERY"

        #grid.selectContextMenuItem(":REQ_QUERY")


        #grid.selectContextMenuItem("REQ")

#session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem "8265D72160021FD0B6F43226BAE842F8NEW:REQ_QUERY"

        #grid.selectContextMenuItem("EBAN")
        #session.sendVKey(0)  
        #grid.pressContextButton("SELECT")

        #grid.press()
        #grid.selectRow(6)  # Selecciona la primera fila
        # grid.pressContextButton("SELECT")
        # grid.selectRow(item_index)
        # grid.selectContextMenuItem("8265D72160021FD0B6F19C1CA23F42F6NEW:REQ_QUERY")

        #session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton ("SELECT")
        #pyautogui.press("s")
        #pyautogui.hotkey("enter")

          # #ejecutar_accion_sap(id_documento="0", ruta_vbs=rf".\scriptsVbs\clickptextos.vbs")

        # # Abre Ventana Solicitudes de Pedido M21N
        # #grid = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]")
        # buscar_y_clickear(rf".\img\vSeleccion.png", confidence=0.7, intentos=20, espera=0.5)
        # time.sleep(0.5)
        # pyautogui.press("s")

                # grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14")
        # print(type(grid))
        # print(grid.Type)
        # grid.select()






    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

if __name__ == "__main__":
    Main_Pruebas4()



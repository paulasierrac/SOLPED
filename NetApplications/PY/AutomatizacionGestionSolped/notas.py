# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU01_LoginSAP import ObtenerSesionActiva
# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Notas():

    def replace_in_text(self,texto: str, replacements: dict):
        """
        Reemplaza textos sobre un string completo.

        Args:
            texto (str): texto original
            replacements (dict): {"SAA": "R3", ...}

        Returns:
            nuevo_texto (str)
            cambios (int): número de líneas modificadas
        """
        lineas = texto.splitlines()
        cambios = 0
        nuevas_lineas = []

        for linea in lineas:
            nueva = linea
            for buscar, reemplazar in replacements.items():
                # reemplazo exacto por línea
                if linea.strip() == buscar:
                    nueva = reemplazar
                else:
                    nueva = nueva.replace(buscar, reemplazar)

            if nueva != linea:
                cambios += 1

            nuevas_lineas.append(nueva)

        return "\n".join(nuevas_lineas), cambios

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

    """ #obtener los textos de la pestana textos, se utiliza el editor de terxtos de SAP y se accede a el por medio de su ID
        # y con el metodo GetLineText se obtiene el texto de la primera linea (index 1)
        grid = session.findById("wnd[0]/usr/" 
        "subSUB0:SAPLMEGUI:0010/" 
        "subSUB3:SAPLMEVIEWS:1100/" 
        "subSUB2:SAPLMEVIEWS:1200/" 
        "subSUB1:SAPLMEGUI:1301/" 
        "subSUB2:SAPLMEGUI:1303/" 
        "tabsITEM_DETAIL/tabpTABIDT14/" 
        "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329"
        "/subTEXTS:SAPLMMTE:0200/" 
        "subEDITOR:SAPLMMTE:0201/" 
        "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        texto = grid.GetLineText(1)
        print("Texto en la pestaña de textos:")
        print(texto)

        
    """
    #Funciones que no se usarion, se dejan por si en el futuro se requieren para entendimiento de SAP
    """
        def ejecutar_accion_sap(id_documento="0", ruta_vbs=rf".\scriptsVbs\clickptextos.vbs"):
        # Asegúrate de poner la ruta correcta donde guardaste el código de arriba

        ruta_vbs = ruta_vbs


        if os.path.exists(ruta_vbs):
            try:
                # Enviamos el id_documento como argumento
                subprocess.run(["cscript", "//Nologo", ruta_vbs, str(id_documento)], check=True)
                print(f"Script ejecutado correctamente para el ID: {id_documento}")
            except subprocess.CalledProcessError as e:
                print(f"Error al ejecutar VBS: {e}")
        else:
            print("No se encuentra el archivo VBS")
    #NO SE USA
    def PressBuscarBoton(session):

        # Asumimos que ya tienes la sesión iniciada
        # SapGuiAuto = win32com.client.GetObject("SAPGUI")
        # ... session = ...
        # 1. Definir el contenedor padre estable (justo antes de donde cambia el número)
        padre_id = "wnd[0]/usr"
        obj_padre = session.findById(padre_id)

        # 2. Definir el patrón Regex para la parte cambiante
        # Buscamos "subSUB0:SAPLMEGUI:001" seguido de un dígito (0-9)
        patron = re.compile(r"subSUB0:SAPLMEGUI:001\d")

        # 3. Iterar sobre los hijos del padre para encontrar la coincidencia
        id_contenedor_encontrado = None

        for hijo in obj_padre.Children:
            # El hijo.Id devuelve la ruta completa, extraemos solo la parte final o comparamos todo
            if patron.search(hijo.Id):
                id_contenedor_encontrado = hijo.Id
                break
        if id_contenedor_encontrado:
            print(f"Contenedor variable encontrado: {id_contenedor_encontrado}")
            # 4. Reconstruir la ruta completa del botón
            # Esta es la parte de la ruta que va DESPUÉS del número cambiante
            resto_ruta = "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201"
            ruta_final_boton = id_contenedor_encontrado + resto_ruta
            try:
                boton = session.findById(ruta_final_boton)
                boton.Press()
                print("Botón presionado con éxito.")
                return True
            except Exception as e:
                print(f"Error al presionar el botón: {e}")
        else:
            print("No se encontró el contenedor que coincida con la Regex.")
            return False
    #NO SE USA
    def find_sap_control(session, parent_id, dynamic_regex, trailing_path, desired_action=None, value=None):

        # Busca un control SAP cuyo ID contiene una parte dinámica (SAPLMEGUI:0010/0015/etc.)
        # y ejecuta una acción específica (.press, asignar .text, etc.).

        # Args:
        #     session         : Objeto SAP GUI Scripting de la sesión actual.
        #     parent_id       : Punto inicial estable (ej: "wnd[0]/usr")
        #     dynamic_regex   : Patrón regex para identificar el contenedor variable.
        #                       Ej: r"subSUB0:SAPLMEGUI:001\d"
        #     trailing_path   : Ruta que viene DESPUÉS del contenedor dinámico.
        #     desired_action  : Acción a ejecutar: "press", "set_text", "focus", None
        #     value           : Valor para acciones como "set_text"

        # Returns:
        #     El control encontrado (GuiComponent) o None si falla.


        parent = session.findById(parent_id)
        patron = re.compile(dynamic_regex)
        dynamic_container = None

        # Buscar el contenedor que contiene la parte dinámica
        for child in parent.Children:
            if patron.search(child.Id):
                dynamic_container = child.Id
                break

        if dynamic_container is None:
            print("No se encontró un contenedor que coincida con el patrón dinámico.")
            return None

        ruta_final = dynamic_container + trailing_path

        try:
            control = session.findById(ruta_final)
        except:
            print(f"No se pudo encontrar el control final: {ruta_final}")
            return None

        # Ejecutar acción solicitada
        if desired_action == "press":
            try:
                control.press()
                print("Acción .press ejecutada con éxito.")
            except Exception as e:
                print(f"Error al ejecutar .press(): {e}")
                return None

        elif desired_action == "set_text":
            try:
                control.text = value
                print(f"Texto asignado correctamente: {value}")
            except Exception as e:
                print(f"Error al asignar texto: {e}")
                return None

        elif desired_action == "focus":
            try:
                control.setFocus()
                print("Control enfocado correctamente.")
            except Exception as e:
                print(f"Error al aplicar setFocus: {e}")
                return None

        elif desired_action is None:
            # Solo devolver el control sin hacer nada
            pass

        return control
    #NO SE USA
    def limpiar_id_sap(ruta_absoluta):
     
        #Toma una ruta larga tipo '/app/con[0]/ses[0]/wnd[0]/usr...'
        #y devuelve solo desde 'wnd[0]/usr...'
       
        if "/wnd[" in ruta_absoluta:
            # Dividimos el string en donde aparezca "/wnd["
            partes = ruta_absoluta.split("/wnd[")
            # partes[1] contendrá "0]/usr/..." así que le volvemos a pegar el prefijo "wnd["
            ruta_limpia = "wnd[" + partes[1]
            return ruta_limpia
        return ruta_absoluta # Si ya estaba limpia, la devuelve igual
    #NO SE USA
    def ejecutar_creacion_hijo(session):
    
    
    """






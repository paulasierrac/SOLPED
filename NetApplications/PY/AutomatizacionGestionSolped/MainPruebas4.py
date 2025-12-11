from Funciones.ValidacionM21N import ejecutar_creacion_hijo,buscar_y_clickear
from HU.HU01_LoginSAP import ObtenerSesionActiva


session=ObtenerSesionActiva()

for i in range(3):
                
    obj_btnDel = None
    selectsFs = [2,3,4,5]
    obj_tabstrip = ejecutar_creacion_hijo(session)
    ruta=rf"C:\Users\CGRPA042\Documents\Steven\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\img\abajo2.png"
    buscar_y_clickear(ruta, confidence=0.7, intentos=20, espera=0.5)








# import re

# import win32com.client
# from HU.HU01_LoginSAP import ObtenerSesionActiva 
# import time
# import pyautogui
# from Funciones.ValidacionM21N import  buscar_y_clickear



# def limpiar_id_sap(ruta_absoluta):
#     """
#     Toma una ruta larga tipo '/app/con[0]/ses[0]/wnd[0]/usr...'
#     y devuelve solo desde 'wnd[0]/usr...'
#     """
#     if "/wnd[" in ruta_absoluta:
#         # Dividimos el string en donde aparezca "/wnd["
#         partes = ruta_absoluta.split("/wnd[")
#         # partes[1] contendrá "0]/usr/..." así que le volvemos a pegar el prefijo "wnd["
#         ruta_limpia = "wnd[" + partes[1]
#         return ruta_limpia
#     return ruta_absoluta # Si ya estaba limpia, la devuelve igual



# session=ObtenerSesionActiva()
# # Patrón que busca el contenedor de pestañas 'tabsITEM_DETAIL'
# # El .* permite que haya cualquier cosa en el medio (incluyendo el cambio 0010/0020)
# # pero anclamos el final con 'tabsITEM_DETAIL' para ser precisos.
# patron_tabstrip = re.compile(r"wnd\[0\]/usr/.*SAPLMEGUI:\d{4}/.*/tabsITEM_DETAIL$")
 
# obj_tabstrip = None
# obj_btnDel = None

# # Recorremos recursivamente o buscamos inteligentemente. 
# # Dado que 'FindAllByName' no existe nativamente en SAP para rutas parciales,
# # lo mejor es localizar el padre 'usr' y buscar el hijo que coincida.
# user_area = session.findById("wnd[0]/usr")
 
# # Nota: Esto es simplificado. En estructuras profundas, a veces es mejor iterar 
# # sobre los hijos de 'usr' buscando cual contiene "SAPLMEGUI".
# for hijo in user_area.Children:
#     if "SAPLMEGUI" in hijo.Id:
#         # Una vez dentro del área variable, intentamos construir la ruta al tabstrip
#         # Ojo: Aquí asumimos la estructura interna fija después del cambio 0010/0020
#         # Tomamos el ID del hijo (ej: ...:0010) y le pegamos el resto de la ruta que SÍ es constante:
#         ruta_restante = "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL"
#         ruta_restante_btnDel = "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201"
#         ruta_restante_textoposicion="/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell"
#         ruta_restante_textoarea="/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell"

#         try:
#             full_id = hijo.Id + ruta_restante
#             full_id = limpiar_id_sap(full_id)
#             obj_tabstrip = session.findById(full_id)
#             print("id:!!!!!")
#             print(full_id)
#             break # ¡Encontrado!
#         except:
#             continue

# selectsFs = [2,3,4,5]

# if obj_tabstrip:
#     nombre_pestaña_buscada = "Textos" # O "Invoice", "Entregas", etc.
#     pestaña_encontrada = False
#     for pestaña in obj_tabstrip.Children:
#         # pestaña.Text te da el nombre visible (ej: "Condiciones")
#         # pestaña.Name te da el ID técnico (ej: "TABIDT3")
#         if pestaña.Text == nombre_pestaña_buscada:
#             pestaña_encontrada = True

#             print(f"Pestaña '{nombre_pestaña_buscada}' seleccionada. (ID Técnico: {pestaña.Name})")
#             full_id_btnDel = limpiar_id_sap(pestaña.Id) + ruta_restante_btnDel
#             full_id_textoposicion = limpiar_id_sap(pestaña.Id) + ruta_restante_textoposicion
#             full_id_textoarea = limpiar_id_sap(pestaña.Id) + ruta_restante_textoarea
#             time.sleep(2)
#             pestaña.Select()

#             for i in selectsFs:
#                 F0n = "F0" + str(i)
            
#                 # .selectedNode = "F02" Texto pedido de posicion   
#                 obj_textoposicion = session.findById(full_id_textoposicion)
#                 print(f"Texto posicion  '{obj_textoposicion.Id}' seleccionada. (ID Técnico: {obj_textoposicion.Name})")
#                 obj_textoposicion.selectedNode = F0n
#                 time.sleep(2)
#                 #Boton Eliminar 
#                 try:
#                     obj_btnDel = session.findById(full_id_btnDel)
#                     print(f"Bot+on Delete '{obj_btnDel.Id}' seleccionada. (ID Técnico: {obj_btnDel.Name})")
#                     obj_btnDel.Press()

#                     # entrar a editar texto "."
#                     obj_textoarea = session.findById(full_id_textoarea)
#                     obj_textoarea.text = "."
#                 except:
#                     pass
#                     # ruta=rf".\img\abajo.png"
#                     # buscar_y_clickear(ruta, confidence=0.8, intentos=20, espera=0.5)

#             break
#     if not pestaña_encontrada:
#         print(f"No se encontró la pestaña llamada {nombre_pestaña_buscada}")
# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import re
import subprocess
import time
import os
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
import pyautogui 


def boton_existe(session,id):
    try:
        session.findById(id)
        return True
    except:
        return False
    
def buscar_y_clickear(ruta_imagen, confidence=0.5, intentos=20, espera=0.5):
    """
    Busca una imagen en pantalla y hace click cuando la encuentra.

    Args:
        ruta_imagen (str): Ruta de la imagen a buscar.
        confidence (float): Confianza para el match (requiere OpenCV).
        intentos (int): Número de intentos antes de fallar.
        espera (float): Tiempo entre intentos en segundos.

    Returns:
        bool: True si hizo click, False si no encontró la imagen.
    """

    for _ in range(intentos):
        pos = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
        if pos:
            pyautogui.click(*pos)
            return True
        time.sleep(espera)

    print(f"[WARNING] No se encontró la imagen: {ruta_imagen}")
    return False

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


def find_sap_control(session, parent_id, dynamic_regex, trailing_path, desired_action=None, value=None):
    
    """
    Busca un control SAP cuyo ID contiene una parte dinámica (SAPLMEGUI:0010/0015/etc.)
    y ejecuta una acción específica (.press, asignar .text, etc.).

    Args:
        session         : Objeto SAP GUI Scripting de la sesión actual.
        parent_id       : Punto inicial estable (ej: "wnd[0]/usr")
        dynamic_regex   : Patrón regex para identificar el contenedor variable.
                          Ej: r"subSUB0:SAPLMEGUI:001\d"
        trailing_path   : Ruta que viene DESPUÉS del contenedor dinámico.
        desired_action  : Acción a ejecutar: "press", "set_text", "focus", None
        value           : Valor para acciones como "set_text"

    Returns:
        El control encontrado (GuiComponent) o None si falla.
    """

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


def limpiar_id_sap(ruta_absoluta):
    """
    Toma una ruta larga tipo '/app/con[0]/ses[0]/wnd[0]/usr...'
    y devuelve solo desde 'wnd[0]/usr...'
    """
    if "/wnd[" in ruta_absoluta:
        # Dividimos el string en donde aparezca "/wnd["
        partes = ruta_absoluta.split("/wnd[")
        # partes[1] contendrá "0]/usr/..." así que le volvemos a pegar el prefijo "wnd["
        ruta_limpia = "wnd[" + partes[1]
        return ruta_limpia
    return ruta_absoluta # Si ya estaba limpia, la devuelve igual

# def ejecutar_creacion_hijo(session):
#     user_area = session.findById("wnd[0]/usr")
#     for hijo in user_area.Children:
#         if "SAPLMEGUI" in hijo.Id:
#             # Una vez dentro del área variable, intentamos construir la ruta al tabstrip
#             # Ojo: Aquí asumimos la estructura interna fija después del cambio 0010/0020
#             # Tomamos el ID del hijo (ej: ...:0010) y le pegamos el resto de la ruta que SÍ es constante:
#             ruta_restante = "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL"
            
#             try:
#                 full_id = hijo.Id + ruta_restante
#                 full_id = limpiar_id_sap(full_id)
#                 obj_tabstrip = session.findById(full_id)
#                 print("id:!!!!!")
#                 print(full_id)
#                 return obj_tabstrip

#             except:
#                 continue


def ejecutar_creacion_hijo(session):
    # 1. Definir el área padre.
    # A veces incluso encontrar wnd[0]/usr falla si SAP está muy lag.
    try:
        user_area = session.findById("wnd[0]/usr")
    except:
        # Si falla de entrada, esperamos un poco y reintentamos una vez
        time.sleep(1)
        user_area = session.findById("wnd[0]/usr")
 
    ruta_restante = "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL"
    # 2. BUCLE DE RESILIENCIA (Reintentos)
    # Intentaremos leer los hijos hasta 3 veces antes de rendirnos.
    max_intentos = 3
    for intento in range(max_intentos):
        try:
            # Intentamos acceder a la propiedad .Children
            # Aquí es donde estaba fallando tu código
            hijos = user_area.Children
            for hijo in hijos:
                if "SAPLMEGUI" in hijo.Id:
                    try:
                        full_id = hijo.Id + ruta_restante
                        full_id = limpiar_id_sap(full_id)
                        # Verificamos si el objeto realmente existe con esa ruta
                        obj_tabstrip = session.findById(full_id)
                        # Si llegamos aquí, todo está bien
                        # print(f"Contenedor encontrado: {full_id}")
                        return obj_tabstrip
                    except:
                        continue
            # Si terminamos el for y no retornamos nada, es que no se encontró en este intento
            # pero no hubo error técnico.
            break
 
        except Exception as e:
            # Este bloque captura el error "Data necessary... not available"
            print(f"Intento {intento + 1}/{max_intentos} fallido esperando a SAP... ({e})")
            time.sleep(1.5) # Espera importante: Dale tiempo a SAP para terminar de pintar
            continue
 
    return None # Si fallaron los 3 intentos o no se encontró


def EncontrarDynpros(session):
    import re

    wnd = session.ActiveWindow
    dynpros = set()
    cola = [wnd]
    patron = re.compile(r":[A-Z0-9_]+:(\d{4})")

    while cola:
        nodo = cola.pop(0)

        try:
            nid = nodo.Id
            encontrados = patron.findall(nid)
            for d in encontrados:
                dynpros.add(d)
        except:
            pass

        try:
            count = nodo.Children.Count
        except:
            continue

        for i in range(count):
            try:
                child = nodo.Children(i)
                cola.append(child)
            except:
                pass
    lista_dynpros = sorted(dynpros)

    if lista_dynpros:
        primer_valor = lista_dynpros[0]
        return primer_valor
    else:
        return("No se encontraron dynpros.")
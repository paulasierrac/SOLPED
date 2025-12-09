# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
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
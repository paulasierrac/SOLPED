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
from pyautogui import ImageNotFoundException
from Funciones.GeneralME53N import AbrirTransaccion, ColsultarSolped, procesarTablaME5A, ActualizarEstadoYObservaciones


def boton_existe(session,id):
    try:
        session.findById(id)
        return True
    except:
        return False
    
# def buscar_y_clickear(ruta_imagen, confidence=0.5, intentos=20, espera=0.5):
#     """
#     Busca una imagen en pantalla y hace click cuando la encuentra.

#     Args:
#         ruta_imagen (str): Ruta de la imagen a buscar.
#         confidence (float): Confianza para el match (requiere OpenCV).
#         intentos (int): Número de intentos antes de fallar.
#         espera (float): Tiempo entre intentos en segundos.

#     Returns:
#         bool: True si hizo click, False si no encontró la imagen.
#     """

#     for _ in range(intentos):
#         pos = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
#         if pos:
#             pyautogui.click(*pos)
#             return True
#         time.sleep(espera)

#     print(f"[WARNING] No se encontró la imagen: {ruta_imagen}")
#     return False


def buscar_y_clickear(
    ruta_imagen,
    confidence=0.5,
    intentos=20,
    espera=0.5,
    fail_silently=True,
    log=True
):
    """
    Busca una imagen en pantalla y hace click cuando la encuentra,
    controlando el error si no aparece y permitiendo continuar el flujo.

    Args:
        ruta_imagen (str): Ruta de la imagen a buscar.
        confidence (float): Nivel de confianza para el match (OpenCV).
        intentos (int): Número máximo de intentos.
        espera (float): Tiempo de espera entre intentos (segundos).
        fail_silently (bool): Si True, no lanza excepción al fallar.
        log (bool): Si True, muestra mensajes de estado.

    Returns:
        bool: True si se encontró e hizo click, False si no.
    """

    for intento in range(1, intentos + 1):
        try:
            pos = pyautogui.locateCenterOnScreen(
                ruta_imagen,
                confidence=confidence
            )

            if pos:
                pyautogui.click(pos)
                if log:
                    print(f"[INFO] Imagen encontrada y clickeada: {ruta_imagen}")
                return True

        except ImageNotFoundException:
            # PyAutoGUI puede lanzar esta excepción en algunas versiones
            pass

        except Exception as e:
            if log:
                print(f"[ERROR] Error inesperado buscando imagen {ruta_imagen}: {e}")
            if not fail_silently:
                raise

        time.sleep(espera)

    if log:
        print(f"[WARNING] Imagen no encontrada tras {intentos} intentos: {ruta_imagen}")

    if not fail_silently:
        raise RuntimeError(f"No se encontró la imagen: {ruta_imagen}")

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

def BorrarTextosDesdeSolped(session, solped, item=2):
    
    esperar_sap_listo(session)
    AbrirTransaccion(session, "ME21N")
    print("Transacción ME21N abierta con éxito.")
    time.sleep(0.5)
    session.findById("wnd[0]").maximize()
    try:
        # Validación básica de sesión
        if not session:
            raise ValueError("Sesion SAP no valida.")
        esperar_sap_listo(session)
        #Click en carrito para el foco 
        ruta=rf".\img\carrito.png"
        buscar_y_clickear(ruta, confidence=0.5, intentos=20, espera=0.5)

        # Navegar hasta el campo Variante de seccion
        for i in range(
            6
        ):  # 29 veces desde menu(sin Shift), 7 desde proveedor, 12 desde org compras
            pyautogui.hotkey("shift", "TAB")
            time.sleep(0.5)
        pyautogui.press("enter")
        # Selecciona el campo Solicitudes de pedido en la lista
        time.sleep(0.5)
        pyautogui.press("s")
        time.sleep(0.5)

        # ingresa el numero de la solped que va a revisar
        esperar_sap_listo(session)
        session.findById("wnd[0]/usr/ctxtSP$00026-LOW").text = solped
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Navegar hasta la sol.pedido en la lista
        for i in range(2):
            pyautogui.hotkey("shift", "TAB")
        pyautogui.hotkey("TAB")

        # Despliga los itemns de la solped
        time.sleep(0.5)
        pyautogui.hotkey("right")
        time.sleep(0.5)
        pyautogui.hotkey("down")
        time.sleep(0.5)

        # Selecciona todos los items de la solped revisar variable item para ajustar
        with pyautogui.hold("shift"):
            pyautogui.press(
                "down", presses=item
            )  # Stev: cantidad de items a bajar articulos de la solped
            time.sleep(0.5)

        # enter en tomar pedido con articulos seleccionados (Click en tomar pedido )
        for i in range(5):
            pyautogui.hotkey("shift", "TAB")
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1)
        print("Esperando a click en pestana de texto y luego en info.......... ")
        time.sleep(1)
        ejecutar_accion_sap(id_documento="click pestaña texto e info ",ruta_vbs=rf".\scriptsVbs\clickptextos.vbs")
        time.sleep(10)


        # Definimos las rutas relativas (colas estáticas)
        ruta_restante_btnDel = "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201"
        ruta_restante_textoposicion = "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell"
        ruta_restante_textoarea = "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell"
        # Bucle principal de items (filas de la solped)

        for i in range(item):
            selectsFs = [2, 3, 4, 5]
            # --- CAMBIO CLAVE: Bucle interno de tipos de texto ---
            for j in selectsFs:
                print(f"--- Procesando tipo de texto F0{j} ---")
                # 1. PASO CRÍTICO: RE-DESCUBRIR LA PESTAÑA Y RE-CALCULAR IDs EN CADA VUELTA
                # Porque el .Press() anterior pudo haber cambiado el ID del contenedor padre (0010 vs 0015)
                obj_tabstrip = ejecutar_creacion_hijo(session)
                if not obj_tabstrip:
                    print("No se pudo encontrar el contenedor dinámico en esta iteración.")
                    break
                # Buscar la pestaña "Textos" de nuevo (su ID padre pudo cambiar)
                full_id_base_pestaña = ""
                pestaña_encontrada = False
                esperar_sap_listo(session)
                for pestaña in obj_tabstrip.Children:
                    if pestaña.Text == "Textos":
                        # Capturamos el ID limpio actual de la pestaña
                        full_id_base_pestaña = limpiar_id_sap(pestaña.Id)
                        pestaña_encontrada = True
                        # Aseguramos que esté seleccionada (importante tras un refresh)
                        try:
                            pestaña.Select()
                        except:
                            pass # A veces ya está seleccionada
                        break
                if not pestaña_encontrada:
                    print("Pestaña Textos no encontrada, saltando...")
                    continue
                # 2. CONSTRUIR RUTAS FRESCAS CON EL ID BASE ACTUAL
                # Ahora estamos seguros de que 'full_id_base_pestaña' es válido para ESTE momento
                current_id_textoposicion = full_id_base_pestaña + ruta_restante_textoposicion
                current_id_btnDel = full_id_base_pestaña + ruta_restante_btnDel
                current_id_textoarea = full_id_base_pestaña + ruta_restante_textoarea
                try:
                    # 3. SELECCIONAR NODO EN EL ÁRBOL
                    F0n = "F0" + str(j)
                    obj_textoposicion = session.findById(current_id_textoposicion)
                    obj_textoposicion.selectedNode = F0n
                    # Pequeña espera para que SAP cargue el texto asociado a ese nodo
                    time.sleep(1)
                    # 4. INTENTAR BORRAR
                    # Verificamos si existe el botón delete (a veces no hay texto y el botón se deshabilita o desaparece)
                    try:
                        obj_btnDel = session.findById(current_id_btnDel)
                        obj_btnDel.Press()
                        print(f"Texto F0{j} eliminado.")
                        # --- ESPERA OBLIGATORIA TRAS BORRAR ---
                        # Aquí SAP destruye y reconstruye la pantalla. 
                        # Esto es lo que rompe los IDs para la siguiente vuelta del 'for j'.
                        time.sleep(1.5)
                        # 5. EDITAR TEXTO (Poner el punto)
                        # Ojo: Como hubo refresh, debemos re-buscar el área de texto con el ID fresco
                        # Pero cuidado: a veces al borrar, el foco cambia. 
                        # Re-validamos el objeto antes de usarlo.
                        try:
                            obj_textoarea = session.findById(current_id_textoarea)
                            obj_textoarea.text = "."
                        except:
                            # Si falla aquí, es probable que necesitemos recalcular el ID de nuevo 
                            # o que el área de texto no esté lista.
                            pass
                    except Exception as e_btn:
                        # Si no encuentra el botón de borrar, es que no había texto o ya estaba vacío
                        # print(f"No se requiere borrar o botón no disponible: {e_btn}")
                        pass
                except Exception as e:
                    print(f"Error procesando texto F0{j}: {e}")
                    # Si falla algo grave, intentamos continuar con el siguiente tipo de texto
                    continue
            # --- FIN DEL BUCLE INTERNO ---
            # Lógica para pasar al siguiente item (flecha abajo visual con PyAutoGUI)
            print("Pasando al siguiente item de la Solped...")
            time.sleep(1)
            ruta_img = rf".\img\abajo.png"
            buscar_y_clickear(ruta_img, confidence=0.8, intentos=20, espera=0.5)

        # Salir de SAP 
        #session.findById("wnd[0]").sendVKey(12)
        esperar_sap_listo(session)
        time.sleep(1)
        pyautogui.hotkey("ctrl", "s") 
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.press("F12")
        time.sleep(1)
        esperar_sap_listo(session)

    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise

def leer_solpeds_desde_archivo(ruta_archivo):
    resultados = {}

    with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
        for linea in f:
            # Todas las líneas útiles empiezan con '|'
            if not linea.strip().startswith("|"):
                continue
            
            partes = [p.strip() for p in linea.split("|")]

            # Esperamos al menos 16 columnas por la estructura del archivo
            if len(partes) < 16:
                continue

            try:
                purch_req = partes[1]       # PurchReq
                estado    = partes[15]      # Estado
            except IndexError:
                continue  # Si alguna fila viene corrupta

            # Evitar encabezados
            if purch_req.lower().startswith("purch"):
                continue

            # Inicializar
            if purch_req not in resultados:
                resultados[purch_req] = {
                    "items": 0,
                    "estados": set()
                }

            # Sumar item
            resultados[purch_req]["items"] += 1
            resultados[purch_req]["estados"].add(estado)

    return resultados   

def obtener_numero_oc(session):
    """
    Obtiene el número de la Orden de Compra creada desde la barra de estado.
    """
    try:
        # El mensaje de éxito con el número de OC suele aparecer en la barra de estado.
        status_text = session.findById("wnd[0]/sbar").text
        # Usamos una expresión regular para buscar un número que sigue a un texto estándar.
        # "Standard PO created under the number 4500021244" -> Ejemplo
        match = re.search(r'(\d{10,})', status_text) # Busca 10 o más dígitos
        if match:
            numero_oc = match.group(1)
            print(f"Número de OC extraído: {numero_oc}")
            return numero_oc
        else:
            print("No se pudo encontrar el número de OC en la barra de estado.")
            return None
    except Exception as e:
        print(f"Error al obtener el número de OC: {e}")
        return None 

def esperar_sap_listo(session, timeout=10):
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            if not session.Busy:
                return True
        except:
            pass
        time.sleep(0.2)

    raise TimeoutError("SAP GUI no terminó de cargar (session.Busy)")

# ===============================================================================================
# INICIO DE CÓDIGO DE VALIDACIÓN INDEPENDIENTE PARA HU04
# ===============================================================================================

import pandas as pd
import chardet
import win32clipboard
from datetime import datetime
from typing import Dict, Tuple, List

# --- Funciones auxiliares de GUI y archivos (reimplementación para HU04) ---

def _DetectarCodificacion_HU04(path: str) -> str:
    try:
        with open(path, "rb") as f:
            rawdata = f.read()
        resultado = chardet.detect(rawdata)
        return resultado["encoding"]
    except Exception:
        return "utf-8"

def _TablaItemsDataFrame_HU04(name: str) -> pd.DataFrame:
    try:
        path = rf"{RUTAS['PathInsumos']}\TablasME53N\{name}"
        encoding = _DetectarCodificacion_HU04(path)
        
        with open(path, "r", encoding=encoding, errors='ignore') as f:
            lineas = f.read().splitlines()

        tabla = [l for l in lineas if l.strip().startswith("|") and "---" not in l]
        if not tabla: return pd.DataFrame()

        encabezado_raw = tabla[0]
        columnas = [c.strip() for c in encabezado_raw.split("|")[1:-1]]
        
        columnas_unicas = []
        contador = {}
        for col in columnas:
            if col in contador:
                contador[col] += 1
                columnas_unicas.append(f"{col}_{contador[col]}")
            else:
                contador[col] = 0
                columnas_unicas.append(col)

        filas = []
        for fila in tabla[1:]:
            partes = [c.strip() for c in fila.split("|")[1:-1]]
            if len(partes) == len(columnas_unicas):
                filas.append(partes)
        
        return pd.DataFrame(filas, columns=columnas_unicas)
    except Exception as e:
        print(f"ERROR en _TablaItemsDataFrame_HU04: {e}")
        return pd.DataFrame()

def _ObtenerItemsME53N_HU04(session, numero_solped: str) -> pd.DataFrame:
    try:
        grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
        grid.pressToolbarContextButton("&MB_EXPORT")
        grid.selectContextMenuItem("&PC")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = rf"{RUTAS['PathInsumos']}\TablasME53N"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"TablaSolped{numero_solped}_HU04.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)
        return _TablaItemsDataFrame_HU04(f"TablaSolped{numero_solped}_HU04.txt")
    except Exception as e:
        print(f"ERROR en _ObtenerItemsME53N_HU04: {e}")
        return pd.DataFrame()

def _ObtenerTextoDelPortapapeles_HU04() -> str:
    try:
        win32clipboard.OpenClipboard()
        texto = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return texto or ""
    except Exception:
        return ""

def _ObtenerItemTextME53N_HU04(session, numero_solped: str, numero_item: str) -> str:
    try:
        editor = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")
        editor.SetFocus()
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)
        texto_completo = _ObtenerTextoDelPortapapeles_HU04()
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
        time.sleep(0.5)
        return texto_completo
    except Exception as e:
        print(f"ERROR en _ObtenerItemTextME53N_HU04: {e}")
        return ""

# --- Funciones de procesamiento y validación (reimplementación para HU04) ---

def _LimpiarNumero_HU04(valor: str) -> float:
    if not valor or not isinstance(valor, str): return 0.0
    valor_limpio = valor.strip().replace("$", "").replace(" ", "")
    if "." in valor_limpio and "," in valor_limpio:
        if valor_limpio.rfind(".") > valor_limpio.rfind(","):
            valor_limpio = valor_limpio.replace(",", "")
        else:
            valor_limpio = valor_limpio.replace(".", "").replace(",", ".")
    elif "," in valor_limpio:
        valor_limpio = valor_limpio.replace(",", ".")
    try:
        return float(valor_limpio)
    except (ValueError, TypeError):
        return 0.0

def _ExtraerDatosTexto_HU04(texto: str) -> Dict:
    datos = { "nit": "", "concepto_compra": "", "cantidad": "", "valor_total": "", "codigo_operacion": [] }
    if not texto or not texto.strip(): return datos

    texto_upper = texto.upper()
    
    # --- Lógica de normalización específica de HU04 ---
    REEMPLAZOS = {
        "VENTA SERVICIO": "V1", "VENTA PRODUCTO": "V1",
        "GASTO PROPIO SERVICIO": "C2", "GASTO PROPIO PRODUCTO": "C2",
        "SAA SERVICIO": "R3", "SAA PRODUCTO": "R3",
    }
    codigos_encontrados = set()
    for keyword, code in REEMPLAZOS.items():
        if keyword in texto_upper:
            codigos_encontrados.add(code)
    datos["codigo_operacion"] = sorted(list(codigos_encontrados))

    # --- Extracción de campos ---
    patrones = {
        "nit": r"NIT[\s:]*([0-9.\-]+)",
        "concepto_compra": r"POR CONCEPTO DE[:\s]*(.+?)\s*(?:EMPRESA|FECHA|CANTIDAD|VALOR)",
        "cantidad": r"CANTIDAD[\s:]*([0-9.,]+)",
        "valor_total": r"VALOR TOTAL[\s:]*([\$]?\s*[0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)",
    }
    for campo, patron in patrones.items():
        m = re.search(patron, texto_upper)
        if m:
            datos[campo] = m.group(1).strip()
            
    return datos

def _ValidarContraTabla_HU04(datos_texto: Dict, df_items: pd.DataFrame, item_num: str) -> Dict:
    validaciones = { "cantidad": {"match": False}, "valor_total": {"match": False}, "resumen": "" }
    if df_items.empty:
        validaciones["resumen"] = "Tabla vacía"
        return validaciones

    item_df = df_items[df_items["Item"].astype(str).str.strip() == str(item_num).strip()]
    if item_df.empty:
        validaciones["resumen"] = "Item no encontrado"
        return validaciones

    fila_item = item_df.iloc[0]

    # Validar Cantidad
    if datos_texto["cantidad"] and "Quantity" in fila_item:
        cantidad_texto = _LimpiarNumero_HU04(datos_texto["cantidad"])
        cantidad_tabla = _LimpiarNumero_HU04(str(fila_item["Quantity"]))
        validaciones["cantidad"]["match"] = abs(cantidad_texto - cantidad_tabla) < 0.01

    # Validar Valor Total
    columna_total_sap = "Total Value" if "Total Value" in fila_item else "Total Val."
    if datos_texto["valor_total"] and columna_total_sap in fila_item:
        valor_texto = _LimpiarNumero_HU04(datos_texto["valor_total"])
        valor_tabla = _LimpiarNumero_HU04(str(fila_item[columna_total_sap]))
        if valor_tabla > 0:
            validaciones["valor_total"]["match"] = abs(valor_texto - valor_tabla) / valor_tabla < 0.01
        else:
            validaciones["valor_total"]["match"] = valor_texto == valor_tabla
            
    return validaciones

def _DeterminarEstadoFinal_HU04(datos_texto: Dict, validaciones: Dict) -> Tuple[str, str]:
    if not all(v.get("match") for v in validaciones.values() if isinstance(v, dict)):
        return "Datos no coinciden", "Los datos del texto no coinciden con los de la tabla SAP."
    if not datos_texto.get("codigo_operacion"):
        return "Sin Codigo Op", "No se encontró código de operación (V1, C2, R3) en el texto."
    return "Registro validado para orden de compra", "Validación de HU04 exitosa."

def _ProcesarYValidarItem_HU04(session, solped: str, item_num: str, texto: str, df_items: pd.DataFrame) -> Tuple[str, str]:
    datos_texto = _ExtraerDatosTexto_HU04(texto)
    validaciones = _ValidarContraTabla_HU04(datos_texto, df_items, item_num)
    estado_final, observaciones = _DeterminarEstadoFinal_HU04(datos_texto, validaciones)
    return estado_final, observaciones

# --- Función principal de orquestación para HU04 ---

def ValidarSolpedParaOC(session, task_name, solped, df_solpeds_para_actualizar, archivo):
    WriteLog(f"Iniciando validación tipo HU04 para SOLPED {solped}", "INFO", task_name, RUTAS["PathLog"])
    
    AbrirTransaccion(session, "ME53N")
    if not ColsultarSolped(session, solped):
        WriteLog(f"No se pudo consultar la SOLPED {solped} en ME53N.", "ERROR", task_name, RUTAS["PathLogError"])
        ActualizarEstadoYObservaciones(df_solpeds_para_actualizar, archivo, solped, nuevo_estado="Error Consulta ME53N", observaciones="No se pudo consultar en SAP para validación.")
        return False

    df_items = _ObtenerItemsME53N_HU04(session, solped)
    if df_items.empty:
        WriteLog(f"SOLPED {solped} no tiene ítems para validar.", "WARNING", task_name, RUTAS["PathLog"])
        ActualizarEstadoYObservaciones(df_solpeds_para_actualizar, archivo, solped, nuevo_estado="Sin Items", observaciones="No se encontraron items en SAP para validar.")
        return False

    lista_items = df_items.to_dict(orient="records")
    if lista_items and (lista_items[-1].get("Status", "").strip() == "*" or lista_items[-1].get("Item", "").strip() == ""):
        lista_items.pop()

    for item_row in lista_items:
        numero_item = item_row.get("Item", "").strip()
        print(f"--- Validando Item {numero_item} de SOLPED {solped} (HU04) ---")

        texto_item = _ObtenerItemTextME53N_HU04(session, solped, numero_item)
        
        estado_final, observaciones = _ProcesarYValidarItem_HU04(session, solped, numero_item, texto_item, df_items)

        if estado_final != "Registro validado para orden de compra":
            mensaje_error = f"Item {numero_item} no validado. Estado: {estado_final}. Obs: {observaciones}"
            WriteLog(f"SOLPED {solped}: {mensaje_error}", "ERROR", task_name, RUTAS["PathLogError"])
            
            ActualizarEstadoYObservaciones(
                df_solpeds_para_actualizar, archivo, solped, item=numero_item,
                nuevo_estado=estado_final, observaciones=observaciones
            )
            ActualizarEstadoYObservaciones(
                df_solpeds_para_actualizar, archivo, solped,
                nuevo_estado="Error de Validacion HU04", observaciones=f"Fallo en item {numero_item}: {estado_final}"
            )
            return False 
    
    return True
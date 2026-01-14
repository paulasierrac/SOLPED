# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import traceback
import pandas as pd
import re
import subprocess
import time
import os
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
import pyautogui
from pyautogui import ImageNotFoundException
from Funciones.GeneralME53N import AbrirTransaccion, ColsultarSolped, procesarTablaME5A, ActualizarEstadoYObservaciones
from typing import List, Optional

class SapTextEditor:
    """
    Wrapper para el editor de textos SAP (GuiShell - SAPLMMTE).
    Permite leer y modificar texto línea por línea de forma segura.
    #Stev: se prueban varios metodos, pero la mejor opcion es tomar todo el texto y luego setearlo todo de nuevo desde la linea 0
    # usando EditorTxt.SetUnprotectedTextPart(0,".")
    """

    def __init__(self, session, editor_id):
        """
        Args:
            session: sesión activa SAP GUI
            editor_id (str): ID completo del editor (hasta /shell)
        """
        self.session = session
        self.editor_id = editor_id
        self.shell = self._get_shell()

    def _get_shell(self):
        shell = self.session.findById(self.editor_id)
        if shell.Type != "GuiShell":
            raise Exception("El objeto encontrado no es un GuiShell (Text Editor SAP)")
        return shell

    # ------------------------------------------------------------------
    # LECTURA
    # ------------------------------------------------------------------

    def get_line(self, index):
        """Obtiene el texto completo de una línea."""
        try:
            return self.shell.GetLineText(index)
        except Exception:
            return None

    def get_all_text(self, max_lines=100):
        """
        Obtiene todo el texto del editor SAP sin saltos de línea finales
        ni líneas vacías generadas por el control.
        """
        lines = []


        for i in range(max_lines):
            try:
                line = self.get_line(i)

                if line is None:
                    break

                # Limpia caracteres invisibles pero conserva el contenido
                clean_line = line.rstrip()
                lines.append(clean_line)

            except Exception:
                break

        # Elimina líneas vacías finales
        while lines and lines[-1] == "":
            lines.pop()

        return "\n".join(lines)




    # ------------------------------------------------------------------
    # ESCRITURA
    # ------------------------------------------------------------------

    def set_editable_line(self, index, new_text):
        """Modifica únicamente la parte editable de una línea."""
        try:
            self.shell.SetUnprotectedTextPart(index, new_text)
            return True
        except Exception:
            return False

    def replace_in_text(self, texto: str, replacements: dict):
        """
        Reemplaza textos sobre un string completo, evitando líneas vacías iniciales
        y agregando un mensaje final con el total de cambios.

        Args:
            texto (str): texto original del editor SAP
            replacements (dict): {"SAA": "R3", ...}

        Returns:
            nuevo_texto (str)
            cambios (int)
        #Stev: cambia el texto segun un diccionario de reemplazos y retorna el texto modificado y la cantidad de cambios realizados
        """

        if not texto:
            return texto, 0

        # 1. Eliminar SOLO saltos de línea iniciales
        texto = texto.lstrip("\n")
        lineas = texto.splitlines()

        nuevas_lineas = []
        cambios = 0
        #CambioExacto=[]

        for linea in lineas:
            nueva = linea

            for buscar, reemplazar in replacements.items():
                # Reemplazo exacto por línea
                if linea.strip() == buscar:
                    nueva = reemplazar
                    print(f"[CAMBIO EXACTO] '{linea}' -> '{reemplazar}'")
                else:
                    nueva = nueva.replace(buscar, reemplazar)

            if nueva != linea:
                cambios += 1

            nuevas_lineas.append(nueva)

        # 2. Agregar mensaje final si hubo cambios
        if cambios > 0:
            nuevas_lineas.append("")
            nuevas_lineas.append(
                f"TEXTO MODIFICADO AUTOMÁTICAMENTE POR BOT RPA – CAMBIOS APLICADOS: {cambios}"
            )

        return "\n".join(nuevas_lineas), cambios

# fin class SapTextEditor:
# fin utilidades

# ===============================================================================================
# Funciones para obtener el ID de los objetos dinamicamnete dependiento del objeto padre
# devuelve el valor de la propiedad o ejecuta la accion deseada
# ===============================================================================================

def press_GuiButton(session, button_id):
    """
    Presiona un botón (GuiButton) en SAP GUI a partir de su ID lógico.

    Args:
        session: sesión SAP activa
        button_id (str): ID lógico del botón (ej: 'AUTOTEXT002')

    Ejemplo:
        press_GuiButton(session, "AUTOTEXT002")
    """

    if not button_id:
        raise ValueError("button_id es obligatorio")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"btn%#{button_id}"

    def buscar_boton(obj):
        try:
            if obj.Type == "GuiButton" and obj.Id.endswith(target_suffix):
                return obj

            for child in obj.Children:
                res = buscar_boton(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    boton = buscar_boton(usr)

    if not boton:
        raise Exception(f"No se encontró GuiButton con ID lógico: {button_id}")

    # Press() es seguro incluso si el botón ya fue usado
    boton.Press()

def SetGuiComboBoxkey(session, campo_id, key_value="ZRCR"):
    """
    Selecciona un valor en un GuiComboBox de SAP GUI usando un ID lógico.

    Args:
        session: sesión SAP activa
        campo_id (str): identificador lógico del campo (ej: 'TOPLINE-BSART')
        key_value (str): clave a seleccionar en el combo (ej: 'ZRCR')

    Raises:
        Exception si no se encuentra el GuiComboBox
    Ejemplo de uso:
        SetGuiComboBoxkey(session, "TOPLINE-BSART", "ZRCR")
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio")

    usr = session.findById("wnd[0]/usr")

    def buscar_combobox(obj):
        try:
            if (
                obj.Type == "GuiComboBox"
                and campo_id in obj.Id
            ):
                return obj

            for child in obj.Children:
                res = buscar_combobox(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    combo = buscar_combobox(usr)

    if not combo:
        raise Exception(f"No se encontró GuiComboBox con ID lógico: {campo_id}")

    # Selección por Key (forma correcta y estable)
    combo.Key = key_value

def set_GuiCTextField_text(session, campo_id, valor):
    """
    Setea un texto en un GuiCTextField (ME21N / cabecera)
    usando el identificador lógico SAP (ej: 'EKORG', 'EKGRP').

    Args:
        session: sesión activa de SAP GUI
        campo_id (str): ID lógico del campo (ej: 'EKORG')
        valor (str): texto a insertar

    Raises:
        ValueError: si los argumentos son inválidos
        Exception: si no se encuentra el campo
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio (ej: 'EKORG')")

    if valor is None:
        raise ValueError("valor no puede ser None")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"-{campo_id}"

    def buscar_ctextfield(obj):
        try:
            if (
                obj.Type == "GuiCTextField"
                and obj.Id.endswith(target_suffix)
            ):
                return obj

            for child in obj.Children:
                res = buscar_ctextfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    ctext = buscar_ctextfield(usr)

    if not ctext:
        raise Exception(f"No se encontró GuiCTextField con ID lógico: {campo_id}")

    # En SAP es buena práctica asegurar foco antes de escribir
    try:
        ctext.SetFocus()
    except Exception:
        pass

    ctext.Text = str(valor)

def get_GuiCTextField_text(session, campo_id):
    """
    Obtiene el texto (.Text.strip()) de un GuiCTextField en ME21N
    a partir del identificador lógico SAP (ej: 'EKORG', 'EKGRP').

    Args:
        session: sesión activa de SAP GUI
        campo_id (str): identificador lógico del campo SAP (ej: 'EKORG')

    Returns:
        str: texto del campo (strip)

    Raises:
        ValueError: si el argumento es inválido
        Exception: si no se encuentra el GuiCTextField
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio (ej: 'EKORG')")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"-{campo_id}"

    def buscar_ctextfield(obj):
        try:
            if (
                obj.Type == "GuiCTextField"
                and obj.Id.endswith(target_suffix)
            ):
                return obj

            for child in obj.Children:
                res = buscar_ctextfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    ctext = buscar_ctextfield(usr)

    if not ctext:
        raise Exception(f"No se encontró GuiCTextField con ID lógico: {campo_id}")

    return ctext.Text.strip()

#for fila in range(item):
#   precio = get_GuiTextField_text(session, f"NETPR[10,{fila}]")

def get_GuiTextField_text(session, campo_posicion):
    """
    Obtiene el texto de un GuiTextField dentro de un TableControl SAP
    usando una posición lógica (ej: 'NETPR[10,0]').

    Args:
        session: sesión SAP activa
        campo_posicion (str): campo con posición SAP (ej: 'NETPR[10,0]')

    Returns:
        str: texto del campo

    Raises:
        Exception si no se encuentra el objeto
    """

    if not campo_posicion:
        raise ValueError("campo_posicion es obligatorio")

    # Parsear NETPR[10,0]
    match = re.match(r"(.+)\[(\d+),(\d+)\]", campo_posicion)
    if not match:
        raise ValueError("Formato inválido. Use: NETPR[10,0]")

    campo, col, fila = match.groups()
    col = int(col)
    fila = int(fila)

    usr = session.findById("wnd[0]/usr")

    def buscar_textfield(obj):
        try:
            if (
                obj.Type == "GuiTextField"
                and campo in obj.Id
                and obj.Id.endswith(f"[{col},{fila}]")
            ):
                return obj

            for child in obj.Children:
                res = buscar_textfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    txt = buscar_textfield(usr)

    if not txt:
        raise Exception(f"No se encontró GuiTextField: {campo_posicion}")

    return txt.Text.strip()

def set_GuiTextField_text(session, campo_posicion, valor):
    """
    Setea el texto de un GuiTextField dentro de un TableControl SAP
    usando posición lógica (ej: 'NETPR[10,0]' o 'MENGE[6,0]').
    Compatible con M21N (MEPO1211).
    """

    if not campo_posicion:
        raise ValueError("campo_posicion es obligatorio")

    if valor is None:
        valor = ""

    # Parseo CAMPO[col,fila]
    match = re.fullmatch(r"([A-Z0-9_]+)\[(\d+),(\d+)\]", campo_posicion.upper())
    if not match:
        raise ValueError("Formato inválido. Use: NETPR[10,0] o MENGE[6,0]")

    campo, col, fila = match.groups()
    col = int(col)
    fila = int(fila)

    usr = session.findById("wnd[0]/usr")

    objetivo = f"-{campo}[{col},{fila}]"

    def buscar_textfield(obj):
        try:
            if (
                obj.Type == "GuiTextField"
                and objetivo in obj.Id
            ):
                return obj

            for child in obj.Children:
                res = buscar_textfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    txt = buscar_textfield(usr)

    if not txt:
        raise Exception(f"No se encontró GuiTextField SAP: {campo}[{col},{fila}]")

    # Seteo seguro (SAP-friendly)
    txt.SetFocus()
    txt.Text = str(valor)
    txt.CaretPosition = len(txt.Text)
    session.findById("wnd[0]").sendVKey(0)


def ventana_abierta(session, titulo_parcial):
    """
    Verifica si existe una ventana abierta cuyo título contenga el texto indicado.

    Args:
        session: sesión activa SAP GUI
        titulo_parcial (str): texto a buscar en el título (case-insensitive)

    Returns:
        bool
    """

    titulo_parcial = titulo_parcial.lower()

    for wnd in session.Children:
        try:
            if titulo_parcial in wnd.Text.lower():
                return True
        except Exception:
            pass

    return False

def SelectGuiTab(session, tab_id):
    """
    Selecciona una pestaña (GuiTab) del detalle de posición en ME21N.
    Args:
        session: sesión activa de SAP GUI
        tab_id (str): ID lógico de la pestaña (ej: 'TABIDT14')
    Ejemplos:
        SelectGuiTab(session, "TABIDT14")  # Textos
        SelectGuiTab(session, "TABIDT05")  # Entrega
    """

    if not tab_id:
        raise ValueError("tab_id es obligatorio")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"tabp{tab_id}"

    def buscar_tab(obj):
        try:
            if obj.Type == "GuiTab" and obj.Id.endswith(target_suffix):
                return obj
            for child in obj.Children:
                res = buscar_tab(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    tab = buscar_tab(usr)

    if not tab:
        raise Exception(f"No se encontró la pestaña GuiTab con ID :{tab_id}")
    # Select() es seguro incluso si ya está seleccionada
    tab.Select()

def boton_existe(session, id):
    """
    Verifica de forma segura si un objeto SAP existe a partir de su ID completo.

    Args:
        session: La sesión activa de SAP GUI.
        id (str): El ID completo del objeto a verificar.

    Returns:
        bool: True si el objeto existe, False en caso contrario.
    """
    try:
        session.findById(id)
        return True
    except:
        return False

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
    task_name = "HU4_GeneracionOC"

    for intento in range(1, intentos + 1):
        try:
            pos = pyautogui.locateCenterOnScreen(
                ruta_imagen,
                confidence=confidence
            )

            if pos:
                pyautogui.click(pos)
                #pyautogui.press("enter") # Descomentar si se quiere dar enter tras el click
                if log:
                    WriteLog(
                        mensaje=f"Imagen encontrada y clickeada: {ruta_imagen}",
                        estado="INFO",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                    #print(f"[INFO] Imagen encontrada y clickeada: {ruta_imagen}")
                return True

        except ImageNotFoundException:
            # PyAutoGUI puede lanzar esta excepción en algunas versiones
            #pyautogui.press("enter") # Descomentar si se quiere dar enter tras el click
            pass

        except Exception as e:
            if log:
                 WriteLog(
                        mensaje=f"Error inesperado buscando imagen {ruta_imagen}: {e}",
                        estado="ERROR",
                        task_name=task_name,
                        path_log=RUTAS["PathLog"],
                    )
                 #print(f"[ERROR] Error inesperado buscando imagen {ruta_imagen}: {e}")
            if not fail_silently:
                raise

        time.sleep(espera)

    if log:
        WriteLog(
            mensaje=f"Imagen no encontrada tras {intento} intentos: {ruta_imagen}",
            estado="WARNING",
            task_name=task_name,
            path_log=RUTAS["PathLog"],
        )
        #print(f"[WARNING] Imagen no encontrada tras {intento} intentos: {ruta_imagen}")

    if not fail_silently:
        raise RuntimeError(f"No se encontró la imagen: {ruta_imagen}")

    return False

# ===============================================================================================
# BORRA LOS TEXTOS DE LAS SOLPED QUE NO SE USAS DESPUES DE "texto posicion"  HU4 G OC
# ===============================================================================================
def obtener_valor(texto: str, contiene: List[str]) -> Optional[str]:
    """
    Busca un valor numérico en una línea que contenga
    alguna de las palabras clave especificadas, con o sin símbolo $.

    Args:
        texto (str): Texto multilínea donde buscar.
        contiene (List[str]): Palabras clave a buscar en la línea.

    Returns:
        Optional[str]: Valor numérico encontrado (como string) o None.
    """

    # Patrón: opcional $, números con separadores de miles
    patron = re.compile(r"(?:\$?\s*)(\d{1,3}(?:[.,]\d{3})*|\d+)")

    contiene_upper = [c.upper() for c in contiene]

    for linea in texto.splitlines():
        linea_upper = linea.upper()

        if any(c in linea_upper for c in contiene_upper):
            match = patron.search(linea)
            if match:
                # Normalizar valor (quita separadores)
                valor = match.group(1).replace(".", "").replace(",", "")
                return valor

    return None

def ValidarAjustarSolped(session,item=1):
    """
    Cambia los precios de la Solped segun el texto del "Texto pedido" (textPF.selectedNode ="F01")
    borra los textos adicionales que no se utilizan (textPF.selectedNode ="F02"), hasta el F05

    Args:
        session: sesión SAP activa
        item (int): posiciones que tiene la Solped

    Raises:
        Exception si no se encuentra el objeto
    """

    try:
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
        Scroll = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/" \
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")

        # Todo: Stev: bucle para revisar visibles en el grid de posiciones
        filas_visibles = Scroll.VisibleRowCount
        # Lista de acciones en SAP que sirve de informe
        acciones = []
        #Obtiene los valores de los campos de precio en la tabla de posiciones

        for fila in range(item):  #cambiar por item
            PosicionSolped = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST")
            PosicionSolped.key = f"   {fila+1}"
            Position=fila
            print(f"Posicion :{Position} Fila: {fila}")
            print("Filas visubles ", filas_visibles)
            # obtiene el Precio de la posicion
            precio = get_GuiTextField_text(session, f"NETPR[10,0]")
            precio = normalizar_precio_sap(precio)
            print(f"precio posicion {fila+1}0:{precio}")

            # Obtine la Cantidad en la Posicion
            CantidadPosicion = get_GuiTextField_text(session, f"MENGE[6,0]")
            #CantidadPosicion = normalizar_precio_sap(precio)
            print(f"Cantidad posicion {fila+1}0:{CantidadPosicion}")

            # obtiene el texto del objeto
            editor = SapTextEditor(session, EDITOR_ID)
            texto = editor.get_all_text()

            # Obtiene el valor en el texto
            claves = ["VALOR "] # str que busca en el texto
            preciotexto = obtener_valor(texto, claves)
            preciotexto = normalizar_precio_sap(preciotexto)
            print("este es el precio en los textos :", preciotexto)

            # Obtiene la cantidad en el texto
            claves = ["CANTIDAD"] # str que busca en el texto
            cantidadtexto = obtener_valor(texto, claves)
            print("esta es la cantidad en los textos :", cantidadtexto)
            #print( "Precio obtenido desde el texto: ",preciotexto)
            acciones.append(f"Precio en el texto de la posicion {fila+1}0: {preciotexto}")
            acciones.append(f"Cantid en el texto de la posicion {fila+1}0: {cantidadtexto}")

            # Comparacion de Valores de Cantidad
            if CantidadPosicion==cantidadtexto or cantidadtexto==None:
                if CantidadPosicion==cantidadtexto:
                    print(f"Cantidad coincideen la posicion : {fila+1}0: {CantidadPosicion} == {cantidadtexto}")
                    acciones.append(f"Cantidad coincideen la posicion : {fila+1}0: {CantidadPosicion} == {cantidadtexto}")
                elif cantidadtexto==None:
                    print(f"no hay cantidad en el texto de la posicion: {fila+1}0: {CantidadPosicion} == {cantidadtexto}")
                    acciones.append(f"no hay cantidad en el texto de la posicion: {fila+1}0: {CantidadPosicion} == {cantidadtexto}")
            else:
                set_GuiTextField_text(session, f"MENGE[6,0]", cantidadtexto)
                print(f"Se mofico posicion :{fila+1}0 Cantidad -> {CantidadPosicion} != {cantidadtexto}")
                acciones.append(f"Se mofico Cantidad en posicion {fila+1}0 = CP: {CantidadPosicion} != CT:{cantidadtexto}")


            # Comparacion de Valores de Cantidad TODO: que pasa si el precio texto es nulo

            if precio==preciotexto:
                print(f"Precio coincide en la posicion : {fila+1}0: {precio} == {preciotexto}")
                acciones.append(f"Precio en la posicion : {fila+1}0: {precio} == {preciotexto}")
            else:
                set_GuiTextField_text(session, f"NETPR[10,0]", preciotexto)
                print(f"Se mofico posicion :{fila+1}0 Precio -> {precio} != {preciotexto}")
                acciones.append(f"Se mofico Precio en posicion :{fila+1}0: {precio} != {preciotexto}")

            # Realiza los reemplazos en el texto

            reemplazos = {
                    "VENTA SERVICIO": "V1",
                    "VENTA PRODUCTO": "V1",
                    "GASTO PROPIO SERVICIO": "C2",
                    "GASTO PROPIO PRODUCTO": "C2",
                    "SAA": "R3", #"SAA SERVICIO": "R3"
                    "SAA PRODUCTO": "R3",
                }
            nuevo_texto, cambios = editor.replace_in_text(texto, reemplazos)

            acciones.append(f"Cambios realizados: {cambios} en la posicion :{fila+1}0 en el Texto ")
            editext=session.findById(EDITOR_ID)
            editext.SetUnprotectedTextPart(0,nuevo_texto)
            #Borra los textos de cada editor F02 en adelante
            for i in range(2, 6):  # F02 a F05
                textPF = session.findById(textoPosicionF)
                nodo = f"F0{i}"
                textPF.selectedNode = nodo
                editxt = session.findById(EDITOR_ID)
                #editor = SapTextEditor(session, EDITOR_ID)
                texto = editor.get_all_text()
                if texto :
                    #print("El texto no esta vacío. Procediendo a borrarlo... :"f"F0{i}")
                    editxt.SetUnprotectedTextPart(0,".")
                    acciones.append(f"Texto borrado en F0{i} en posicion {fila+1}0")

            # presiona el botón abajo siguiente posicion
            #press_GuiButton(session, "AUTOTEXT002")
            textPF.selectedNode ="F01"
            esperar_sap_listo(session)
            #time.sleep(0.5)
            # da scroll una posicion hacia abajo para no perder visual de los objetos en la tabla de SAP
            Scroll = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/" \
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")
            Scroll.verticalScrollbar.position = fila+1
            SelectGuiTab(session, "TABIDT14")
            print("Posicion visible despues del Scroll:")
            print(get_GuiTextField_text(session, f"EBELP[1,0]"))
            esperar_sap_listo(session)
        # Devuelve las accines ejecutadas en una lista 
        return acciones

    except Exception as e:
        #todo: canbiar por log
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

def AbrirSolped(session, solped, item=2):
    """
    Navega en la GUI de SAP para tomar una Solicitud de Pedido (SOLPED) específica
    y prepararla para la creación de una Orden de Compra.

    Args:
        session: La sesión activa de SAP GUI.
        solped (str): El número de la Solicitud de Pedido a procesar.
        item (int): El número de ítems o posiciones que contiene la SOLPED.

    Raises:
        TimeoutError: Si una ventana esperada de SAP no aparece en el tiempo definido.
        Exception: Captura y relanza errores generales durante la interacción con SAP.
    """
    try:
        print("SOLPED            :",solped)
        print("POSICIONES        :",item)

        esperar_sap_listo(session)
        # Click Variante de Seleccion y selecciona el campo Solicitudes de pedido en la lista
        timeout = time.time() + 25
        ventana= "Solicitudes de pedido"
        while not ventana_abierta(session, ventana):
            if time.time() > timeout:
                raise TimeoutError(f"No se abrió la ventana :{ventana}")
            buscar_y_clickear(rf".\img\vSeleccion.png", confidence=0.8, intentos=5, espera=0.5)
            esperar_sap_listo(session)
            time.sleep(2)
            pyautogui.press("s") # selecciona el campo Solicitudes de pedido en la lista

        # ingresa el numero de la solped que va a revisar  #Funciona perfecto
        esperar_sap_listo(session)
        session.findById("wnd[0]/usr/ctxtSP$00026-LOW").text = solped
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Navegar hasta la sol.pedido en la lista
        buscar_y_clickear(rf".\img\sol.pedido.png", confidence=0.8, intentos=20, espera=0.5)
        # Despliga los itemns de la solped
        time.sleep(0.5)
        pyautogui.hotkey("right")
        time.sleep(0.5)
        pyautogui.hotkey("down")
        time.sleep(0.5)

        # Selecciona todos los items de la solped revisar variable item para ajustar
        with pyautogui.hold("shift"):
            pyautogui.press("down", presses=item)  # Stev: cantidad de items a bajar articulos de la solped
            time.sleep(0.5)

        # Click en tomar pedido
        buscar_y_clickear(rf".\img\tomar.png", confidence=0.7, intentos=20, espera=0.5)

        #Docstring for MostrarCabecera
        MostrarCabecera()


    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise






def leer_solpeds_desde_archivo(ruta_archivo):
    """
    Lee un archivo de texto plano con formato de tabla (| separado) y extrae
    información de Solicitudes de Pedido (SOLPEDs), agrupando por número de SOLPED.

    Args:
        ruta_archivo (str): La ruta completa al archivo de texto a leer.

    Returns:
        dict: Un diccionario donde cada clave es un número de SOLPED y el valor
              es otro diccionario con el conteo de 'items' y un 'set' de 'estados'.
              Ej: {'10023456': {'items': 3, 'estados': {'Estado A', 'Estado B'}}}
    """
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
    """
    Espera hasta que la sesión de SAP GUI no esté ocupada (session.Busy es False).

    Args:
        session: La sesión activa de SAP GUI.
        timeout (int): Tiempo máximo de espera en segundos.

    Raises:
        TimeoutError: Si SAP sigue ocupado después del tiempo de espera.
    """
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            if not session.Busy:
                return True
        except:
            pass
        time.sleep(0.2)

    raise TimeoutError("SAP GUI no terminó de cargar (session.Busy)")

def CambiarGrupoCompra(session):
    """
    Cambia el Grupo de Compras ('EKGRP') basado en la Organización de Compras ('EKORG') actual.

    Args:
        session: La sesión activa de SAP GUI.

    Returns:
        list: Una lista de strings con las acciones realizadas.

    Raises:
        ValueError: Si la Organización de Compras actual no está en el mapa de condiciones.
    """
    # Obtener el valor actual de la organización de compra
    obj_orgCompra = get_GuiCTextField_text(session, "EKORG")
    if not obj_orgCompra:
        obj_orgCompra = obj_orgCompra.upper()

    #print(f"Valor de OrgCompra: {obj_orgCompra}")
    condiciones = {
        "s":"RCC",
        "S":"RCC",
        "":"RCC",
        "OC15": "RCC",
        "OC26": "HAB",
        "OC25": "HAB",
        "OC28": "AC2",
        "OC27": "AC2"
    }

    if obj_orgCompra not in condiciones:
        raise ValueError(f"Organización de compra '{obj_orgCompra}' no reconocida.")

    obj_grupoCompra = condiciones[obj_orgCompra]



    set_GuiCTextField_text(session, "EKGRP", obj_grupoCompra)
    #print(f"Grupo de compra actualizado a: {obj_grupoCompra}")
    acciones = []
    acciones.append(f"Valor de OrgCompra: {obj_orgCompra}")
    acciones.append(f"Grupo de compra actualizado a: {obj_grupoCompra}")
    return acciones

def normalizar_precio_sap(precio: str) -> int:
    """
    Convierte un precio SAP tipo '2.750.000,00' en entero 2750000
    para comparaciones confiables.
    """
    if not precio:
        return 0

    # Quitar separador de miles y decimales
    limpio = precio.replace(".", "").replace(",00", "")

    return int(limpio)

def MostrarCabecera():
    """
    Asegura que las secciones principales de la interfaz (Cabecera, Resumen, Detalle)
    estén visibles en la transacción ME21N para prevenir errores de "objeto no encontrado".
    """
    time.sleep(0.5)
    pyautogui.hotkey("ctrl","F2")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl","F3")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl","F4")

def ProcesarTabla(name, dias=None):
    """name: nombre del txt a utilizar
    return data frame
    Procesa txt estructura ME5A y devuelve un df con manejo de columnas dinamico.
    dias: int|None -> número de días a mantener (si None, no aplica filtro por fecha)"""

    try:
        WriteLog(
            mensaje=f"Procesar archivo nombre {name}",
            estado="INFO",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLog"],
        )

        # path = f".\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{RUTAS["PathInsumos"]}\{name}"

        # INTENTAR LEER CON DIFERENTES CODIFICACIONES
        lineas = []
        codificaciones = ["latin-1", "cp1252", "iso-8859-1", "utf-8"]

        for codificacion in codificaciones:
            try:
                with open(path, "r", encoding=codificacion) as f:
                    lineas = f.readlines()
                #print(f"EXITO: Archivo leido con codificacion {codificacion}")
                break
            except UnicodeDecodeError as e:
                print(f"ERROR con {codificacion}: {e}")
                continue
            except Exception as e:
                print(f"ERROR con {codificacion}: {e}")
                continue

        if not lineas:
            print("ERROR: No se pudo leer el archivo con ninguna codificacion")
            return pd.DataFrame()

        # Filtrar solo lineas de datos
        filas = [l for l in lineas if l.startswith("|") and not l.startswith("|---")]

        # DETECTAR ESTRUCTURA DE COLUMNAS DINAMICAMENTE
        if not filas:
            print("No se encontraron filas de datos en el archivo")
            return pd.DataFrame()

        # Analizar la primera fila para determinar estructura
        primera_fila = filas[0].strip().split("|")[1:-1]  # Quitar | inicial y final
        primera_fila = [p.strip() for p in primera_fila]

        num_columnas = len(primera_fila)
        #print(f"Estructura detectada: {num_columnas} columnas")
        #print(f"   Encabezados: {primera_fila}")

        # DEFINIR COLUMNAS BASE SEGUN ESTRUCTURA
        if num_columnas == 14:
            # Estructura original (sin Estado ni Observaciones)
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
            ]
            columnas_extra = ["Estado", "Observaciones"]

        elif num_columnas == 15:
            # Verificar si la columna 15 es "Estado" o "Observaciones"
            ultima_columna = primera_fila[-1].lower()
            if "estado" in ultima_columna:
                # Estructura con Estado pero sin Observaciones
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Estado",
                ]
                columnas_extra = ["Observaciones"]
            else:
                # Estructura con Observaciones pero sin Estado
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Observaciones",
                ]
                columnas_extra = ["Estado"]

        elif num_columnas == 16:
            # Estructura completa con Estado y Observaciones
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
                "Estado",
                "Observaciones",
            ]
            columnas_extra = []
        else:
            print(f"ERROR: Estructura no soportada: {num_columnas} columnas")
            return pd.DataFrame()

        # PROCESAR TODAS LAS FILAS
        filas_proc = []
        for i, fila in enumerate(filas):
            partes = fila.strip().split("|")[1:-1]
            partes = [p.strip() for p in partes]

            # Validar que tenga el numero correcto de columnas
            if len(partes) == num_columnas:
                filas_proc.append(partes)
            elif len(partes) == num_columnas + 1 and partes[-1] == "":
                # Caso: columna extra vacia al final
                filas_proc.append(partes[:num_columnas])
                if i < 3:  # Solo log primeras filas
                    print(f"   ADVERTENCIA Fila {i+1}: Columna extra vacia removida")
            else:
                print(
                    f"   ERROR Fila {i+1} ignorada: {len(partes)} columnas vs {num_columnas} esperadas"
                )
                if i == 0:  # Solo mostrar detalle para primera fila
                    print(f"      Contenido: {partes}")
                continue

        # CREAR DATAFRAME
        df = pd.DataFrame(filas_proc, columns=columnas_base)

        # AGREGAR COLUMNAS FALTANTES
        for col_extra in columnas_extra:
            if col_extra not in df.columns:
                df[col_extra] = ""
                print(f"EXITO: Columna '{col_extra}' agregada al DataFrame")

        # FILTRAR: Si la primera fila es encabezado, eliminarla
        primera_fila_es_encabezado = any(
            col in df.iloc[0].values if not df.empty else False
            for col in [
                "Purch.Req.",
                "Item",
                "Req.Date",
                "Short Text",
                "PurchReq",
                "Estado",
                "Observaciones",
            ]
        )

        if not df.empty and primera_fila_es_encabezado:
            df = df.iloc[1:].reset_index(drop=True)
            #print("EXITO: Fila de encabezado removida")

        #print(f"EXITO: Archivo procesado: {len(df)} filas de datos")
        #print(f"   - Columnas: {list(df.columns)}")

        if not df.empty:
            print(f"   - SOLPEDs: {df['PurchReq'].nunique()}")
            if "Estado" in df.columns:
                print(f"   - Estados unicos: {df['Estado'].value_counts().to_dict()}")

        # Normalizar formato fecha
        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        if dias is not None:
            hoy = pd.Timestamp.today().normalize()
            limite = hoy - pd.Timedelta(days=int(dias))
            filas_antes = len(df)
            df = df[df["ReqDate_fmt"] >= limite].reset_index(drop=True)
            filas_despues = len(df)
            print(
                f"EXITO: Filtrado por ReqDate últimos {dias} días -> {filas_despues}/{filas_antes}"
            )
        else:
            print("INFO: No se aplicó filtro por ReqDate (dias=None)")

        # opcional: eliminar columna auxiliar
        df.drop(columns=["ReqDate_fmt"], inplace=True)

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en procesarTablaME5A: {e}",
            estado="ERROR",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )
        print(f"ERROR en procesarTablaME5A: {e}")
        traceback.print_exc()
        return pd.DataFrame()

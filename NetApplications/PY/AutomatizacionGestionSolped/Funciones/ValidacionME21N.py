# ============================================
# Función Local: validacionME53N
# Autor: Steven Navarro - NetApplications
# Descripcion: Funciones 
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
from requests import session
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
from Funciones.Login import ObtenerSesionActiva
from Funciones.GuiShellFunciones import (EditorTextoSAP,
setGuiTextFieldText,              
ObtenerTextoCampoGuitextfield,
buscarObjetoPorIdParcial,
get_importesCondiciones,
obtenerValor,
extraerConcepto,
obtenerCorreos,
normalizarPrecioSap, 
clasificarConcepto,
EsperarSAPListo,
buscarYClickear, setSapTableScroll, 
ventanaAbierta,
SelectGuiTab,
MostrarCabecera,
ObtenerNumeroOC
)
from typing import List, Literal, Optional


def ValidarAjustarSolped(session, item=1):
    """
    Cambia los precios y las cantidades de la Solped segun el texto del "Texto pedido" (textPF.selectedNode ="F01")
    borra los textos adicionales que no se utilizan (textPF.selectedNode ="F02"), hasta el F05

    Args:
        session: sesión SAP activa
        item (int): posiciones que tiene la Solped

    Raises:
        Exception si no se encuentra el objeto
    """

    try:

        # Lista de acciones en SAP que sirve de informe
        acciones = []

        for fila in range(item):  # cambiar por item

            # Obtiene el Precio de la posicion
            # PrecioPosicion = ObtenerTextoCampoGuitextfield(session, f"NETPR[10,{fila}]") #Stev: implementar while para scroll, hacer dinamico el f"NETPR[10,{fila}]"
            PrecioPosicion = ObtenerTextoCampoGuitextfield(session, f"NETPR[10,0]")
            PrecioPosicion = normalizarPrecioSap(PrecioPosicion)

            # Obtine la Cantidad en la Posicion
            # CantidadPosicion = ObtenerTextoCampoGuitextfield(session, f"MENGE[6,{fila}]") #Stev: implementar while para scroll, hacer dinamico el f"MENGE[6,{fila}]"
            CantidadPosicion = ObtenerTextoCampoGuitextfield(session, f"MENGE[6,0]")
            # CantidadPosicion = normalizarPrecioSap(CantidadPosicion)

            # FechaPosicion = ObtenerTextoCampoGuitextfield(session, f"EEIND[9,0]")
            # #CantidadPosicion = normalizarPrecioSap(CantidadPosicion)

            # Selecbox de la posicion de la solped  ejemplo de guia :  1 [10] 80016676 , LAVADO MANTEL GRANDE 
            PosicionSolped = buscarObjetoPorIdParcial(session, "cmbDYN_6000-LIST")
            PosicionSolped.key = f"   {fila+1}"

            # Navega a la pestaña de textos
            EsperarSAPListo(session)
            SelectGuiTab(session, "TABIDT14")
            textPF1 = buscarObjetoPorIdParcial(session, "cntlTEXT_TYPES_0200/shell")
            textPF1.selectedNode = "F01" # Foco en primer Texto IMPORTANTE
            EsperarSAPListo(session)
            EDITOR_ID= buscarObjetoPorIdParcial(session, "cntlTEXT_EDITOR_0201/shellcont/shell")
        
            EsperarSAPListo(session)
            # obtiene el texto del objeto ├─ Leer textos
            editor = EditorTextoSAP(session, EDITOR_ID.id)
            texto = editor.TraerTodoElTexto()

            # Obtiene la FECHA: en el texto (Precio)
            claves = ["FECHA:"] # str que busca en el texto
            FechaTexto = obtenerValor(texto, claves)
            print(f"Fecha encontrada en el Texto {FechaTexto}")
            #preciotexto = normalizarPrecioSap(preciotexto)

            # Obtiene el valor en el texto (Precio)
            claves = ["VALOR"]  # str que busca en el texto
            preciotexto = obtenerValor(texto, claves)
            preciotexto = normalizarPrecioSap(preciotexto)

            # Obtiene la cantidad en el texto
            claves = ["CANTIDAD"]  # str que busca en el texto
            cantidadtexto = obtenerValor(texto, claves)

            # Obtiene impuestos en el texto
            claves = ["IMPUESTO QUE APLICA"]  # str que busca en el texto
            impuestostexto = obtenerValor(texto, claves)

            correosColdubsidio = obtenerCorreos(
                texto, "@colsubsidio.com"
            )  # ejemplo de uso de la funcion obtener correos
            acciones.append(
                f"Correos encontrados en el texto de la posicion {fila+1}: {correosColdubsidio}"
            )
            acciones.append(
                f"Impuestos encontrados en el texto de la posicion {fila+1}: {impuestostexto}"
            )

            concepto = extraerConcepto(texto)
            if concepto:
                tipo = clasificarConcepto(concepto)
                acciones.append(f"{concepto} => {tipo}")

            # Comparacion de Valores de Cantidad
            if (
                cantidadtexto
                and cantidadtexto.strip()
                and CantidadPosicion != cantidadtexto
            ):
                setGuiTextFieldText(session, f"MENGE[6,0]", cantidadtexto)
                # print(f"Se mofico posicion :{fila+1}0 Cantidad -> {CantidadPosicion} != {cantidadtexto}")
                acciones.append(
                    f"Posicion {fila+1}0 => CP: {CantidadPosicion} != CT:{cantidadtexto} ⚠️ Se Actualiza Cantidad"
                )

            # Comparacion de Valores de Precio
            if (
                preciotexto
                and str(preciotexto).strip()
                and PrecioPosicion != preciotexto
            ):
                setGuiTextFieldText(session, f"NETPR[10,0]", preciotexto)
                # print(f"Se mofico posicion :{fila+1}0 Precio -> {PrecioPosicion} != {preciotexto}")
                acciones.append(
                    f"Posicion {fila+1}0 => PP:{PrecioPosicion} != PT:{preciotexto}⚠️ Se Actualiza Precio"
                )

            # Realiza los reemplazos en el texto segun cuadro 
            reemplazos = {"VENTA SERVICIO": "V1","VENTA PRODUCTO": "V1","GASTO PROPIO SERVICIO": "C2","GASTO PROPIO PRODUCTO": "C2","SAA": "R3","SAA PRODUCTO": "R3"} #"SAA SERVICIO": "R3"
            nuevo_texto,cambios,cambioEcxacto = editor.RemplazarTextos(texto, reemplazos)

            # Si hay cambios, agrega a la lista de acciones
            if cambios > 0:
                acciones.append(
                    f"Cambios realizados: {cambios} en la posicion :{fila+1}0 en el Texto :{cambioEcxacto}"
                )

            #Borra los textos de cada editor F02 en adelante
            for i in range(2, 6):  # F02 a F05  2,6   F02 a F03 2,
                SelectGuiTab(session, "TABIDT14")
                nodo = f"F0{i}"               
                textPF = buscarObjetoPorIdParcial(session, "cntlTEXT_TYPES_0200/shell")
                textPF.selectedNode = nodo
                texto = editor.TraerTodoElTexto()
                if texto :
                    #print("El texto no esta vacío. Procediendo a borrarlo... :"f"F0{i}")
                    editxt=session.findById(EDITOR_ID.id)
                    editxt.SetUnprotectedTextPart(0,".")

            EsperarSAPListo(session)
            """
            #STEV: Codigo para recuperar impuesto saludable desde la pestaña Condiciones, por lentitud del bot se desactiva po ahora 2/9/2026
            
            # valorImpSaludable = get_importesCondiciones(session)
            # if valorImpSaludable:
            #     acciones.append(f"Impuesto Saludable en la posicion {fila+1}0: {valorImpSaludable}")
            """                   
            setSapTableScroll(session, "TC_1211", fila+1) # da scroll una posicion hacia abajo para no perder visual de los objetos en la tabla de SAP
            #print(f"Primera posicion visible : {ObtenerTextoCampoGuitextfield(session, f'EBELP[1,0]')}") # Muestra la primera posicion Visible despues del scroll 
            EsperarSAPListo(session)

        # Devuelve las acciones ejecutadas en una lista 
        return acciones

    except Exception as e:
        # todo: canbiar por log
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
          
        #EsperarSAPListo(session)
        # Click Variante de Seleccion y selecciona el campo Solicitudes de pedido en la lista
        timeout = time.time() + 25
        ventana = "Solicitudes de pedido"
        while not ventanaAbierta(session, ventana):
            if time.time() > timeout:
                raise TimeoutError(f"No se abrió la ventana :{ventana}")
            
            buscarYClickear(rf".\img\vSeleccion.png", confidence=0.8, intentos=5, espera=0.5)
            #session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton("SELECT")
            # VarianteSeleccion = buscarObjetoPorIdParcial(session, "/shell[0]")
            # VarianteSeleccion1= buscarObjetoPorIdParcial(session, "SELECT")
            # VarianteSeleccion.pressContextButton (VarianteSeleccion1.id)
            # EsperarSAPListo(session)
            # SolicitudesdePedido = buscarObjetoPorIdParcial(session, ":REQ_QUERY")
            # VarianteSeleccion.selectContextMenuItem (SolicitudesdePedido.id)
            #session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton("SELECT")
            time.sleep(2)
            pyautogui.press(
                "s"
            )  # selecciona el campo Solicitudes de pedido en la lista

        # ingresa el numero de la solped que va a revisar  #Funciona perfecto
        EsperarSAPListo(session)
        session.findById("wnd[0]/usr/ctxtSP$00026-LOW").text = solped
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Navegar hasta la sol.pedido en la lista
        buscarYClickear(
            rf".\img\sol.pedido.png", confidence=0.8, intentos=20, espera=0.5
        )
        # Despliga los itemns de la solped
        time.sleep(0.5)
        pyautogui.hotkey("right")
        time.sleep(0.5)
        pyautogui.hotkey("down")
        time.sleep(0.5)

        """
        # Selecciona todos los items de la solped revisar variable item para ajustar
        with pyautogui.hold("shift"):
            pyautogui.press(
                "down", presses=item
            )  # Stev: cantidad de items a bajar articulos de la solped
            time.sleep(0.5)
        """
        primerItem = 2 #desde donde se toman las pociciones TODO: que se pase por parametro, segun cliente con posiciones 
        ultimoItem = item + 2 # Ultima posicion tomada 
        for i in range(primerItem,ultimoItem):   # recordar que en range no incluye el ultimo 
            session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").selectNode(f"          {i}")

        EsperarSAPListo(session)
        # Click en tomar pedido
        #buscarYClickear(rf".\img\tomar.png", confidence=0.7, intentos=20, espera=0.5)
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressButton ("COPY")
        
        #Docstring for MostrarCabecera
        MostrarCabecera()

    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise

# ============================================
# HU05: Generacion de Orden de Compra desde Solped
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Genera Orden de Compra a partir de las Solicitudes de Pedido validadas.
# Ultima modificacion: 27/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import re
import subprocess
import time
import os
from Config.settings import RUTAS
from Funciones.ValidacionM21N import (
    PressBuscarBoton,
    buscar_y_clickear,
    ejecutar_accion_sap,
    limpiar_id_sap,
    ejecutar_creacion_hijo
)
from Funciones.GeneralME53N import AbrirTransaccion


import pyautogui  # Asegúrate de tener pyautogui instalado


def GenerarOCDesdeSolped(session, solped, item=2):
    try:
        # Validación básica de sesión
        if not session:
            raise ValueError("Sesion SAP no valida.")

        # ============================
        # Abrir transacción ME21N
        # ============================
        # Paso 1: abrir transacción ME21N
        AbrirTransaccion(session, "ME21N")
        print("Transacción ME21N abierta con éxito.")
        time.sleep(0.5)

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
    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise

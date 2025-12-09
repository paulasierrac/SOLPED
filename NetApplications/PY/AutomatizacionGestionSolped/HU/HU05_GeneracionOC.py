# ============================================
# HU05: Generacion de Orden de Compra desde Solped
# Autor: Steven Navarro - Configurador RPA
# Descripcion: Genera Orden de Compra a partir de las Solicitudes de Pedido validadas.
# Ultima modificacion: 27/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import subprocess
import time
import os
from Config.settings import RUTAS
from HU.HU1_LoginSAP import ObtenerSesionActiva
from Funciones.ValidacionM21N import (
    boton_existe,
    buscar_y_clickear,
    ejecutar_accion_sap,
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

        # Navegar hasta el campo Variante de seccion
        for i in range(
            7
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
        time.sleep(3)

        # ejecutar_accion_sap(id_documento="click pestaña texto e info ",ruta_vbs=rf".\scriptsVbs\clickptextos.vbs")

        # for para navegar por las posiciones de la solped
        for i in range(item):
            # lista de posiciones de la solped
            # ir a la pestaña textos
            print(
                f"Navegando por la posicion {i+1} de la Solped {solped} Esperando foco en pestana textos.."
            )
            ejecutar_accion_sap(
                id_documento="click pestana texto e info ",
                ruta_vbs=rf".\scriptsVbs\clickptextos.vbs",
            )

            # Validar si el boton borrar texto existe
            Boton = boton_existe(
                session,
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1"
                ":SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB"
                ":SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201",
            )
            if Boton:
                print("Boton Borrar presente")
                # presionar Boton borrar texto
                botonBorrar = session.findById(
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1"
                    ":SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB"
                    ":SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/btnDELETE_0201"
                )
                botonBorrar.press()
                # Punto en el texto
                time.sleep(1)
                puntoentexto = session.findById(
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1"
                    ":SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:"
                    "SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell"
                )
                puntoentexto.text = "."
                time.sleep(1)
                #
                # ejecutar_accion_sap(ruta_vbs=rf".\scriptsVbs\clickptextos.vbs")
                # ir a pedido info con un flecha abajo
                # pyautogui.press('down')
                # entrar a editar texto con ctrl + enter
                # pyautogui.hotkey("ctrl","enter")
                # time.sleep(1)

            else:
                print("Boton Borrar NO esta en esta vista")
                # paso al siguiente texto con flecha abajo
                # pyautogui.press('down')
                # entrar a editar texto con ctrl + enter
                # pyautogui.hotkey("ctrl","enter")
                time.sleep(1)

            # paso al siguiiente item de la solped
            ejecutar_accion_sap(
                id_documento="boton abajo",
                ruta_vbs=rf".\scriptsVbs\clickbotonabajo.vbs",
            )
            time.sleep(1)

    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise

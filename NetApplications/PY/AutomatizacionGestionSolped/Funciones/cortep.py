import win32com.client
import time
import pyautogui
import subprocess
import win32clipboard
import pyperclip
import win32gui
import win32con


def ObtenerSesionActiva():
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        return session
    except:
        print("No fue posible obtener la sesión activa.")
        return None


def ObtenerTextoDelPortapapeles():
    """Obtener texto del portapapeles con manejo correcto de codificación"""
    try:
        # Abrir portapapeles
        win32clipboard.OpenClipboard()
        try:
            # Obtener texto con CF_UNICODETEXT (maneja mejor caracteres especiales)
            texto = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            return texto if texto else ""
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Error al leer portapapeles: {e}")
        return ""


# ========== OPCIÓN 1: Usar pyautogui para simular Alt+Tab ==========
def TraerSAPAlFrente_Opcion1():
    """Usar Alt+Tab para traer SAP al frente"""
    try:
        pyautogui.hotkey("alt", "tab")
        time.sleep(0.5)
        print("✓ SAP traído al frente (Opción 1 - Alt+Tab)")
    except Exception as e:
        print(f"Error en Opción 4: {e}")


session = ObtenerSesionActiva()

if session:
    try:

        # Probar la opción más directa primero
        TraerSAPAlFrente_Opcion1()

        # 1) Obtener el objeto del editor
        editor = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        # 2) Asegurar que el editor tiene el foco
        editor.SetFocus()
        time.sleep(0.5)

        # 3) Seleccionar TODO el texto
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)

        # 4) Copiar al portapapeles
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)

        # 5) Obtener texto del portapapeles con codificación correcta
        texto_completo = ObtenerTextoDelPortapapeles()

        # 6) Limpiar caracteres problemáticos si los hay
        texto_limpio = texto_completo.encode("utf-8", errors="replace").decode("utf-8")

        print("=" * 80)
        print("TEXTO COMPLETO DEL EDITOR:")
        print("=" * 80)
        print(texto_limpio)
        print("=" * 80)

        # 7) Guardar en archivo con UTF-8 explícito
        path = r"C:\Users\CGRPA009\Documents\texto_sap.txt"
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(texto_limpio)
            print(f"\nArchivo guardado en: {path}")
        except Exception as e:
            print(f"Error al guardar archivo: {e}")

        # # 8) Restaurar posición (click al inicio)
        # pyautogui.hotkey("ctrl", "Home")
        # time.sleep(0.2)

        # # 9) Deshacer la selección (click en algún lado)
        # pyautogui.press("escape")

        print("\nProceso completado correctamente.")

        # # 2. Guardarlo directamente en un archivo
        # path = r"C:\Users\CGRPA009\Documents\texto_sap.txt"
        # with open(path, "w", encoding="utf-8") as f:
        #     f.write(texto)

        print("Botón presionado correctamente.")
    except Exception as e:
        print(f"Error al presionar el botón: {e}")

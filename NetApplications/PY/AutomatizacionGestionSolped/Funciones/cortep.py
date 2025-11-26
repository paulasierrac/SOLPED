import win32com.client
import time
import pyautogui
import subprocess


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


session = ObtenerSesionActiva()

if session:
    try:

        ruta = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\expSolped03.txt"
        salida = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\purch_req_unicos.txt"
        # Leer archivo con codificación robusta
        try:
            texto = open(ruta, "r", encoding="utf-8").read()
        except:
            texto = open(ruta, "r", encoding="latin-1").read()

        lineas = texto.splitlines()

        tabla = []

        for linea in lineas:
            # ignorar líneas de guiones
            if linea.startswith("-") or linea.strip() == "":
                continue

            # separar las columnas por el caracter |
            columnas = [col.strip() for col in linea.split("|") if col.strip() != ""]

            tabla.append(columnas)

        purch_reqs = []
        # Mostrar tabla (lista de listas)
        for fila in tabla:
            purch_req = fila[0]
            item = fila[1]
            req_date = fila[2]
            material = fila[3]
            created_by = fila[4]
            short_text = fila[5]
            # print(purch_req)

            purch_reqs.append(purch_req)

        # 3. Quitar duplicados
        purch_reqs_unicos = sorted(set(purch_reqs))  # ordenados y sin duplicados

        # 4. Guardarlos en un TXT
        with open(salida, "w", encoding="utf-8") as f:
            for pr in purch_reqs_unicos:
                f.write(pr + "\n")

        print("Proceso completado.")
        print("Total únicos:", len(purch_reqs_unicos))
        print("Archivo generado:", salida)

        # --------------------------------------Recorriendo Valores Unico----------------------------------------------------

        ruta_unicos = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\purch_req_unicos.txt"

        with open(ruta_unicos, "r", encoding="utf-8") as f:
            for linea in f:
                numero_solped = linea.strip()
                print("Procesando:", numero_solped)

        # # 1) Obtener el objeto del editor
        # editor = session.findById(
        #     "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
        #     "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
        #     "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
        #     "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
        #     "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
        #     "cntlTEXT_EDITOR_0201/shellcont/shell"
        # )

        # # 1. Tomar el texto completo del editor
        # texto = editor.text
        # print(texto)
        # # 2. Guardarlo directamente en un archivo
        # path = r"C:\Users\CGRPA009\Documents\texto_sap.txt"
        # with open(path, "w", encoding="utf-8") as f:
        #     f.write(texto)

        print("Botón presionado correctamente.")
    except Exception as e:
        print(f"Error al presionar el botón: {e}")

# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import time
import os
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS


def AbrirTransaccion(session, transaccion):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida"""

    try:
        WriteLog(
            mensaje="ValidacionME53N",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesión SAP
        if session is None:

            WriteLog(
                mensaje="Sesión SAP no disponible",
                estado="ERROR",
                task_name="ColsultarSolped",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesión SAP no disponible")

        # Abrir transacción ME53N
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME53N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje="Transacción ME53N abierta",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )
        print("Transacción ME53N abierta")
        # Ingresar número de SOLPED

        # Boton de Otra consulta
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        # Escribir numero de solped
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN"
        ).text = numero_solped

        # Activar el radiobutton "Purch. Requisition"
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B"
        ).setFocus()
        # Seleccionar el radiobutton "Purch. Requisition"
        session.findById(
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B"
        ).select()

        # Presionar el botón OK (btn[0])
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(1)

        WriteLog(
            mensaje="Solped {numero_solped} consultada exitosamente",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )
        # ---------------- Exportar tabla a txt----------------

        grid = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        )

        # 1. Abrir menú contexto "Exportar"
        grid.pressToolbarContextButton("&MB_EXPORT")

        # 2. Seleccionar "Exportar → Hoja de cálculo (PC)"
        grid.selectContextMenuItem("&PC")

        # 3. Confirmar ventana de exportar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # 4. Escribir ruta de guardado
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = (
            r"C:\Users\CGRPA009\Documents\SAP\SAP GUI\ruta"
        )

        # 5. Nombre del archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "TablaSolped.txt"

        # 6. Guardar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # ---------------- Capturar Texto----------------

        # 1) Obtener el objeto del editor
        editor = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/"
            "subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/"
            "cntlTEXT_EDITOR_0201/shellcont/shell"
        )

        # 1. Tomar el texto completo del editor
        texto = editor.text
        print(texto)
        # 2. Guardarlo directamente en un archivo
        path = r"C:\Users\CGRPA009\Documents\texto_sap.txt"
        with open(path, "w", encoding="utf-8") as f:
            f.write(texto)

        # item Abajo
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/"
            "btn%#AUTOTEXT002"
        ).press()
        # item Arriba
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/"
            "btn%#AUTOTEXT001"
        ).press()

        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en ColsultarSolped: {e}",
            estado="ERROR",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLogError"],
        )

        return False

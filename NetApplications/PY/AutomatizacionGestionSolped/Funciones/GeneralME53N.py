# ============================================
# Función Local: GeneralME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Archivo Base funciones necesarias transaccion ME53N
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
import win32com.client
import time
import os
from Funciones.EscribirLog import WriteLog
from Config.settings import RUTAS
import pandas as pd
import datetime
import re


def procesarTablaME5A(name):
    """name: nombre del txt a utilizar
    return data frame
    Procesa txt estructura ME5A y devuelve un txt"""

    try:
        WriteLog(
            mensaje=f"Procesar archivo nombre{name}",
            estado="INFO",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLog"],
        )

        path = f"C:\\Users\\CGRPA009\\Documents\\SOLPED-main\\SOLPED\\NetApplications\\PY\\AutomatizacionGestionSolped\\Insumo\\{name}"

        # Cargar texto como raw y limpiarlo
        with open(path, "r", encoding="utf-8") as f:
            lineas = f.readlines()

        # Filtrar solo líneas de datos
        filas = [l for l in lineas if l.startswith("|") and not l.startswith("|---")]

        filas_proc = []
        for fila in filas:
            partes = fila.strip().split("|")

            # Quitar primer y último elemento vacío
            partes = partes[1:-1]

            # Limpiar espacios en cada columna
            partes = [p.strip() for p in partes]

            # Validación: la fila DEBE tener 14 columnas exactas
            if len(partes) != 14:
                print("Fila con columnas inesperadas:", partes)
                continue

            filas_proc.append(partes)

        columnas = [
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

        df = pd.DataFrame(filas_proc, columns=columnas)
        df["Estado"] = ""

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en procesarTablaME5A: {e}",
            estado="ERROR",
            task_name="procesarTablaME5A",
            path_log=RUTAS["PathLogError"],
        )

        return False


def AbrirTransaccion(session, transaccion):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida"""

    try:
        WriteLog(
            mensaje=f"Abrir Transaccion {transaccion}",
            estado="INFO",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesión SAP
        if session is None:

            WriteLog(
                mensaje="Sesión SAP no disponible",
                estado="ERROR",
                task_name="AbrirTransaccion",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesión SAP no disponible")

        # Abrir transacción dinamica
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje=f"Transacción {transaccion} abierta",
            estado="INFO",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLog"],
        )
        print(f"Transacción {transaccion} abierta")
        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en AbrirTransaccion: {e}",
            estado="ERROR",
            task_name="AbrirTransaccion",
            path_log=RUTAS["PathLogError"],
        )

        return False


def ColsultarSolped(session, numero_solped):
    """session: objeto de SAP GUI
    numero_solped:  numero de SOLPED a consultar
    Realiza la verificacion del SOLPED"""

    try:
        WriteLog(
            mensaje=f"Numero de SOLPED : {numero_solped}",
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

        # Boton de Otra consulta
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(0.3)
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

        time.sleep(3)

        WriteLog(
            mensaje=f"Solped {numero_solped} consultada exitosamente",
            estado="INFO",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLog"],
        )

    except Exception as e:
        WriteLog(
            mensaje=f"Error en ColsultarSolped: {e}",
            estado="ERROR",
            task_name="ColsultarSolped",
            path_log=RUTAS["PathLogError"],
        )

        return False


def TablaItemsDataFrame(name) -> pd.DataFrame:
    """name: nombre se archivo a consultar
    Convierte tabla de items a df"""

    try:
        WriteLog(
            mensaje=f"Nombre de archivo {name}",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )

        # Abrir transacción dinamica
        path = rf"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\{name}"

        # 1. Leer archivo completo
        with open(path, "r", encoding="utf-8") as f:
            texto = f.read()

        # 2. Separar por líneas
        lineas = texto.splitlines()

        # 3. Filtrar solo las líneas que forman parte de la tabla
        # (líneas que empiezan y terminan con | )
        tabla = [
            l
            for l in lineas
            if l.strip().startswith("|") and l.strip().endswith("|") and "---" not in l
        ]

        if not tabla:
            raise ValueError("No se encontró ninguna tabla SAP dentro del archivo.")

        # 4. Eliminar líneas de guiones largos (separadores)
        tabla = [
            l for l in tabla if not re.match(r"^-{5,}", l.replace("|", "").strip())
        ]

        # 5. La primera fila válida es el encabezado
        encabezado_raw = tabla[0]
        columnas = [c.strip() for c in encabezado_raw.split("|")[1:-1]]

        # 6. Procesar las filas de datos
        filas = []
        for fila in tabla[1:]:
            partes = [c.strip() for c in fila.split("|")[1:-1]]
            if len(partes) == len(columnas):  # validar integridad
                filas.append(partes)

        # 7. Convertir a DataFrame
        df = pd.DataFrame(filas, columns=columnas)

        WriteLog(
            mensaje=f"DataFrame {df} conversion correcta",
            estado="INFO",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLog"],
        )
        print(f"DataFrame {df} conversion correcta")

        return df
    except Exception as e:
        WriteLog(
            mensaje=f"Error en TablaItemsDataFrame: {e}",
            estado="ERROR",
            task_name="TablaItemsDataFrame",
            path_log=RUTAS["PathLogError"],
        )


def ObtenerItemsME53N(session, numero_solped):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Obtiene los items de SOLPED y los pasa a un df"""

    try:
        WriteLog(
            mensaje=f"Solped{numero_solped} a obtener items",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesión SAP
        if session is None:

            WriteLog(
                mensaje="Sesión SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemsME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesión SAP no disponible")

        # ---------------- Exportar item a txt----------------

        grid = session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
            "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
            "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        )
        grid.setFocus()
        time.sleep(0.5)

        # 1. Abrir menú contexto "Exportar"
        grid.pressToolbarContextButton("&MB_EXPORT")
        time.sleep(0.5)
        # 2. Seleccionar "Exportar → Hoja de cálculo (PC)"
        grid.selectContextMenuItem("&PC")

        # 3. Confirmar ventana de exportar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # 4. Escribir ruta de guardado
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = (
            r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo"
        )

        # 5. Nombre del archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = (
            f"TablaSolped{numero_solped}.txt"
        )

        # 6. Guardar
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.1)
        # Convertir items de la SOLPED a un dataFrame
        df = TablaItemsDataFrame(f"TablaSolped{numero_solped}.txt")

        WriteLog(
            mensaje=f"Solped {numero_solped} convertido a Df con exito",
            estado="INFO",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLog"],
        )
        print(f"Solped {numero_solped} convertido a Df con exito")

        return df
    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemsME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemsME53N",
            path_log=RUTAS["PathLogError"],
        )


def ObtenerItemTextME53N(session, numero_solped):
    """session: objeto de SAP GUI
    Realiza la verificacion del SOLPED"""

    try:
        WriteLog(
            mensaje=f"ObtenerItemTextME53N {numero_solped}",
            estado="INFO",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLog"],
        )

        # Validar sesión SAP
        if session is None:

            WriteLog(
                mensaje="Sesión SAP no disponible",
                estado="ERROR",
                task_name="ObtenerItemTextME53N",
                path_log=RUTAS["PathLog"],
            )
            raise Exception("Sesión SAP no disponible")

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

        identificador = f"\n===== Registro: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n"
        # 1. Tomar el texto completo del editor
        texto = editor.text
        print(texto)
        # 2. Guardarlo directamente en un archivo
        path = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\texto_ITEMsap.txt"
        with open(path, "w", encoding="utf-8") as f:
            f.write(identificador)
            f.write(texto + "\n")
            f.write("-" * 80 + "\n")

        # item Abajo
        session.findById(
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/"
            "btn%#AUTOTEXT002"
        ).press()

        return texto
    except Exception as e:
        WriteLog(
            mensaje=f"Error en ObtenerItemTextME53N: {e}",
            estado="ERROR",
            task_name="ObtenerItemTextME53N",
            path_log=RUTAS["PathLogError"],
        )

        return False

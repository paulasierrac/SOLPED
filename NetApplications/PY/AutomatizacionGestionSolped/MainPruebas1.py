# ================================
# Main: GestionSolped
# Autor: Paula Sierra, Henry Navarro - NetApplications
# Descripcion: Main principal del Bot
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste inicial para cumplimiento de estándar
# ================================
from HU.HU00_DespliegueAmbiente import EjecutarHU00
from HU.HU1_LoginSAP import ObtenerSesionActiva,conectar_sap,abrir_sap_logon
from HU.HU2_DescargaME5A import EjecutarHU02
from HU.HU03_ValidacionME53N import EjecutarHU03
from HU.HU5_GeneracionOC import GenerarOCDesdeSolped

# from NetApplications.PY.AutomatizacionGestionSolped.HU.HU03_ValidacionME53N import buscar_SolpedME53N
from Funciones.EscribirLog import WriteLog
import traceback
from Config.settings import RUTAS,SAP_CONFIG


def Main_Pruebas1():
    try:
        # HU00 pruebas de desarrollo con login
        #EjecutarHU00()

        # HU1 Login - Inicio de sesion SAP 
        WriteLog(
            mensaje="Inicia HU01",
            estado="INFO",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLog"],
        )
        
        #session = ObtenerSesionActiva()  #//Intentar obtener sesión activa , para pruebas de desarrollo con login manual
        session = conectar_sap(
         SAP_CONFIG["sistema"], 
         SAP_CONFIG["mandante"],
         SAP_CONFIG["user"],
         SAP_CONFIG["password"],
         SAP_CONFIG["idioma"]
         )
        
        # HU5 Gereacion de OC 
        WriteLog(
            mensaje="Inicia HU05",
            estado="INFO",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLog"],
        )
       
       
        GenerarOCDesdeSolped(session, "1300139102")  # Reemplaza con la Solped real:  1300139102  1300139269
     
    except Exception as e:
        error_text = traceback.format_exc()
        WriteLog(
            mensaje=f"ERROR GLOBAL: {e} | {error_text}",
            estado="ERROR",
            task_name="Main_GestionSOLPED",
            path_log=RUTAS["PathLogError"],
        )
        raise


if __name__ == "__main__":
    Main_Pruebas1()




"""
        for i in range(item):
                       
            obj_btnDel = None
            selectsFs = [2,3,4,5]
            obj_tabstrip = ejecutar_creacion_hijo(session)

            if obj_tabstrip:
                nombre_pestaña_buscada = "Textos" # O "Invoice", "Entregas", etc.
                pestaña_encontrada = False
                for pestaña in obj_tabstrip.Children:
                    # pestaña.Text te da el nombre visible (ej: "Condiciones")
                    # pestaña.Name te da el ID técnico (ej: "TABIDT3")
                    if pestaña.Text == nombre_pestaña_buscada:
                        pestaña_encontrada = True

                        print(f"Pestaña '{nombre_pestaña_buscada}' seleccionada. (ID Técnico: {pestaña.Name})")
                        full_id_btnDel = limpiar_id_sap(pestaña.Id) + ruta_restante_btnDel
                        full_id_textoposicion = limpiar_id_sap(pestaña.Id) + ruta_restante_textoposicion
                        full_id_textoarea = limpiar_id_sap(pestaña.Id) + ruta_restante_textoarea
                        time.sleep(2)
                        pestaña.Select()

                        for i in selectsFs:
                            F0n = "F0" + str(i)
                        
                            # .selectedNode = "F02" Texto pedido de posicion   
                            obj_textoposicion = session.findById(full_id_textoposicion)
                            print(f"Texto posicion  '{obj_textoposicion.Id}' seleccionada. (ID Técnico: {obj_textoposicion.Name})")
                            obj_textoposicion.selectedNode = F0n
                            time.sleep(2)
                            #Boton Eliminar 
                            try:
                                obj_btnDel = session.findById(full_id_btnDel)
                                print(f"Bot+on Delete '{obj_btnDel.Id}' seleccionada. (ID Técnico: {obj_btnDel.Name})")
                                obj_btnDel.Press()
                                time.sleep(1)

                                # entrar a editar texto "."
                                obj_textoarea = session.findById(full_id_textoarea)
                                obj_textoarea.text = "."
                            except:
                                pass
                        time.sleep(20)   
                        ruta=rf".\img\abajo.png"
                        buscar_y_clickear(ruta, confidence=0.8, intentos=20, espera=0.5)
                        print("Preparando siguiente iteración...")
                        if not pestaña_encontrada:         print(...)
                        
                        break

                if not pestaña_encontrada:
                    print(f"No se encontró la pestaña llamada {nombre_pestaña_buscada}")
 """
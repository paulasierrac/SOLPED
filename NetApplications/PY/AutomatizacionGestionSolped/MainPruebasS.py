from HU.HU01_LoginSAP import ObtenerSesionActiva
from Funciones.ValidacionM21N import get_GuiCTextField_text,set_GuiCTextField_text
def cambiar_grupo_compra(session): 
    
    """
    orgCompra = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/" 
    "subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" 
    "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/" 
    "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG"
    
    grupoCompra = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/" 
    "subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" 
    "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/" 
    "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP"
    """

    # Obtener el valor actual de la organización de compra
    obj_orgCompra = get_GuiCTextField_text(session, "EKORG")

    print(f"Valor de OrgCompra: {obj_orgCompra}")

    condiciones = {
        "OC15": "RCC",
        "OC26": "HAB",
        "OC25": "HAB",
        "OC28": "AC2",
        "OC27": "AC2" 
    }

    if obj_orgCompra not in condiciones:
        raise ValueError(f"Organización de compra '{obj_orgCompra}' no reconocida.")
    
    obj_grupoCompra = condiciones[obj_orgCompra]

    print(obj_grupoCompra)

    #cambio = session.findById(grupoCompra)
    #cambio = get_GuiCTextField_text(session,"EKGRP")
    set_GuiCTextField_text(session, "EKGRP", obj_grupoCompra)


    print(f"Grupo de compra actualizado a: {obj_grupoCompra}")


def MainSantiago():
    try:
        session = ObtenerSesionActiva()
        if not session:
            return

        cambiar_grupo_compra(session)

    except Exception as e:
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

if __name__ == "__main__":
    MainSantiago()
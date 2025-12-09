import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
<<<<<<< HEAD
#import pyautogui


def descarga_OCME9F(session, estado):
    if session:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME9F"
        session.findById("wnd[0]").sendVKey(0)
        print("Transacción ME9F abierta con éxito.")
=======
import pyautogui


def descarga_OCME9F(session, ordende_compra):
    if not session:
     raise ValueError("Sesion SAP no valida.")

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nME9F"
    session.findById("wnd[0]").sendVKey(0)
    print("Transaccion ME9F abierta con exito.")
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = ordende_compra
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1)
    session.findById("wnd[0]/usr/chk[1,5]").selected = True
    #pyautogui.press("F8")
    pyautogui.hotkey("shift", "f5")


>>>>>>> and
       
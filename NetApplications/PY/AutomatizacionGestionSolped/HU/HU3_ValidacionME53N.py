import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
#import pyautogui

def buscar_SolpedME53N(session):        
    if session:
        session.findById("wnd[0]/tbar[0]/okcd").text = ""
        session.findById("wnd[0]").sendVKey(0)
        print("Transacción ME5A abierta con éxito.")


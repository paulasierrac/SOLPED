import win32com.client # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
#import pyautogui


def descarga_OCME9F(session, estado):
    if session:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME9F"
        session.findById("wnd[0]").sendVKey(0)
        print("Transacción ME9F abierta con éxito.")
       
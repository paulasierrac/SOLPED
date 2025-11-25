import win32com.client
import subprocess
import time

subprocess.Popen(r"C:\\Program Files (x86)\\SAP\\FrontEnd\\SapGui\\saplogon.exe")
time.sleep(5)
sapgui = win32com.client.GetObject("SAPGUI")
application = sapgui.GetScriptingEngine 
connection = application.OpenConnection("ERP-CORPORATIVO-CALIDAD", True)
session= connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "410"
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "CGRPA065"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = 'sT1f%4L*'
# for variable_temporal in 10:
#     password = session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = variable_temporal
#     print(f"letra : {password}")
session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "ES"
session.findById("wnd[0]").sendVKey(0)
print(" Conectado correctamente a SAP.") 

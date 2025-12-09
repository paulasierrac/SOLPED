If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
 
' --- LOGICA PARA RECIBIR PARAMETROS ---
Dim sapID
' Verificamos si Python nos mandó algo (Argumento 0)
If WScript.Arguments.Count > 0 Then
    sapID = WScript.Arguments(0)
Else
    ' Valor por defecto por si ejecutas el vbs manualmente para probar
    sapID = "00000000" 
End If
' --------------------------------------
 
session.findById("wnd[0]").maximize
 
' Aquí usamos la variable sapID en lugar del número fijo
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14").select
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F02"
 
' Opcional: Descomenta si necesitas dar Enter al final
' session.findById("wnd[0]").sendVKey 0
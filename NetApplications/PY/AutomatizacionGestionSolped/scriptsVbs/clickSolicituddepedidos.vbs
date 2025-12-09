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
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton "SELECT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem "8265D72160021FD0B5A702EC42E70296NEW:REQ_QUERY"
 
' Opcional: Descomenta si necesitas dar Enter al final
' session.findById("wnd[0]").sendVKey 0
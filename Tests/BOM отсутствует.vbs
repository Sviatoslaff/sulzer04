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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[22]").press
session.findById("wnd[1]/usr/ctxtVIQMEL-EQUNR").text = "12304567"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").text = "2001"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").setFocus
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").caretPosition = 4
session.findById("wnd[1]/usr/btnG_ENTER").press
session.findById("wnd[1]/usr/btnG_CANCEL").press

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
session.findById("wnd[1]/usr/ctxtVIQMEL-EQUNR").text = "476782"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").text = "2001"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").setFocus
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").caretPosition = 4
session.findById("wnd[1]/usr/btnG_ENTER").press
session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART").text = "ZAE"
session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART").caretPosition = 3
session.findById("wnd[2]/tbar[0]/btn[6]").press
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          1","999"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          1","999"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").clickLink "          1","999"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "204991297376"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[5]").press

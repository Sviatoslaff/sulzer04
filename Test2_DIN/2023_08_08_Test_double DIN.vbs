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
session.findById("wnd[1]/usr/ctxtVIQMEL-EQUNR").text = "460792"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").text = "2001"
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").setFocus
session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").caretPosition = 4
session.findById("wnd[1]/usr/btnG_ENTER").press
session.findById("wnd[2]/tbar[0]/btn[6]").press
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          1","999"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          1","999"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").clickLink "          1","999"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = "441.01"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").setFocus
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").topNode = "         48"
session.findById("wnd[0]/tbar[1]/btn[5]").press

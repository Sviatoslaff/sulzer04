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
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
session.findById("wnd[1]/usr/txtRV45A-POSNR").text = "250"
session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
session.findById("wnd[1]").sendVKey 0

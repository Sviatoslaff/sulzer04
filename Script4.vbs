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
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_ITEM").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").caretPosition = 9
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_ITEM").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = "123" + vbCr + ""
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 3,3
session.findById("wnd[0]/tbar[0]/btn[3]").press

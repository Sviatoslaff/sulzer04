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
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/ctxtRV45A-MABNR[1,0]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").key = "00000"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_APLCD").key = "3199"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").setFocus
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").key = "00000"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_APLCD").key = "3199"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").setFocus
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").key = "00000"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_APLCD").key = "3199"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").setFocus
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").key = "00000"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_APLCD").key = "3199"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").setFocus
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

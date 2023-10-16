
'Dim rowCount As Long 
'Dim row As Long

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "50004050"
session.findById("wnd[0]").sendVKey 0

'Set grid = session.findById(UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiCustomControl").Id)
tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id
Set grid = session.findById(tblArea)
'row = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").currentRow
row = grid.currentRow

session.findById(tblArea & "/ctxtRV45A-MABNR[1,"& row - 1 &"]").setFocus


tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id
Set grid = session.findById(tblArea)
sapRow = grid.currentRow                'Here is the current visible row of the QTN
'MsgBox "Row: " & sapRow


If sapRow > 7 Then
    rowCount = grid.RowCount
    goto_pos = session.findById(tblArea & "/txtVBAP-POSNR[0," & sapRow - 5 & "]").text
'    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
    session.findById("wnd[1]/usr/txtRV45A-POSNR").text = goto_pos
    session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 300

    tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id
    Set grid = session.findById(tblArea)
    sapRow = grid.currentRow                'Here is the current visible row of the QTN
    Set cell = grid.GetCell(sapRow + 6, 1)
    cell.setFocus()
    sapRow = grid.currentRow                'Here is the current visible row of the QTN

    MsgBox "Row: " & sapRow



MsgBox "new Row: " & sapRow

End If        

MsgBox lines


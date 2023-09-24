session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "50004061"
session.findById("wnd[0]").sendVKey 0

'Set grid = session.findById(UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiCustomControl").Id)
tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_ANGEBOT", "GuiTableControl").Id
Set grid = session.findById(tblArea)
'row = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").currentRow
row = grid.currentRow

session.findById(tblArea & "/ctxtRV45A-MABNR[1,"& row - 1 &"]").setFocus


saprow = 0
newsaprow = grid.currentRow
diff = newsaprow - saprow

lines = ""
For i = newsaprow To saprow Step -1
    if lines <> "" Then 
        lines = lines & ", "
    end if	
    lines = lines & session.findById(tblArea & "/txtVBAP-POSNR[0," & i & "]").text
Next
if lines <> "" Then
    lines = "[" & lines & "]"
End if	

MsgBox lines


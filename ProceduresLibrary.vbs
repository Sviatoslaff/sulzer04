' Places messages to systems 
' row - SAP row, obj - BOM, DIN or Article or both processed, res - result of action (empty or one or multiple lines)
' sw - software which needs update by info 
Sub InformUser(row, obj, res, sw, comment, ArticlesExcel, intRow, tblArea)
    
    If obj = cBOM Then
        If res = cEmpty Then
            typeText = "Please clarify pump serial number / Existence of BOM on factory"
        End If
    End If

    If obj = cDIN Then
        If res = cEmpty Then
            typeText = "There is not DIN number in pump BOM. Please check DIN number"
        ElseIf res = cOne Then
            typeText = "Ok. DIN info added"
        ElseIf res = cMulti Then
            typeText = "Please choose the exact Part number / Article nom "
        End If
    End If

    If obj = cArticle Then
        If res = cEmpty Then
            typeText = "There is not Part number / Article nom in pump BOM. Please check Part number / Article nom"
        ElseIf res = cOne Then
            typeText = "Ok. Article nom info added"
        ElseIf res = cMulti Then
            typeText = "Ok. Part number / Article nom info added. First DIN number added"
        End If
    End If

    If obj = cDINArt Then
        If res = cEmpty Then
            typeText = "There is not Part number / Article nom in pump BOM. Please check Part number / Article nom"
        ElseIf res = cOne Then
            typeText = "OK, DIN and Part number / Article nom info added"
        ElseIf res = cMulti Then
            typeText = "No Multi Case"
        End If
    End If
    
    If sw = cExcel Or sw = cBoth Then
        If comment <> "" Then
            comment = " Lines:" & comment
        End If
        ArticlesExcel.Cells(intRow, 10).Value = typeText & comment
    End If
    
    ' The focus on the SAP inquiry screen
    If (res = cEmpty) And (res = cSAP Or res = cBoth) Then

        tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id	
   
        session.findById(tblArea & "/ctxtRV45A-MABNR[1," & row & "]").text = "MISC"            'Should be MISC
        session.findById(tblArea & "/ctxtRV45A-KWMENG[12" & row & "]").text = ArticlesExcel.Cells(intRow, 8).Value
        Set btnArea = UserArea.findByName("/btnBT_ITEM").Id
'        btnArea = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050"
        Set btnItem = session.findById(btnArea)
        btnItem.press
        
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").Select
        txtArea = UserArea.findByName("SPLITTER_CONTAINER", "GuiCustomControl").Id
        Set textField = session.findById(txtArea & "/shellcont/shellcont/shell/shellcont[1]/shell")
        textField.text = typeText
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    End If
    
End Sub

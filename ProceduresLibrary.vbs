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
            typeText = "There are not DIN number and Part number / Article nom in pump BOM. Please check your request"
        ElseIf res = cOne Then
            typeText = "OK, DIN and Part number / Article nom info added"
        ElseIf res = cDINWrong Then
            typeText = "There Is Not this DIN In BOM, Part number / Article nom info added" & vbNewLine & _
            "(DIN Is Not the same you requested)"
        ElseIf res = cArtiWrong Then
            typeText = "There Is Not this Part number / Article nom In BOM," & vbNewLine & _
            "DIN info added (Part number / Article nom Is Not the same you requested)"
        ElseIf res = cDINArtWrong Then
            typeText = "DIN And Part number / Article nom Do Not correlate In BOM." & vbNewLine & _
            "Please choose one (DIN Or Part number / Article nom)"
        End If
    End If
    
    If sw = cExcel Or sw = cBoth Then
        If comment <> "" Then
            comment = " Lines:" & comment
        End If
        ArticlesExcel.Cells(intRow, 10).Value = typeText & comment
    End If
    
    ' The focus on the SAP inquiry screen
    If (res = cEmpty) And (sw = cSAP Or sw = cBoth) Then
        
        tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
        
        session.findById(tblArea & "/ctxtRV45A-MABNR[1," & row & "]").text = "MISC"            'Should be MISC
        session.findById(tblArea & "/txtVBAP-ZMENG[2," & row & "]").text = ArticlesExcel.Cells(intRow, 8).Value

        pressEnter()
        pressEnter()

        session.findById("wnd[0]/tbar[0]/btn[11]").press            'SAVE
        WScript.Sleep 300
        pressEnter()


        tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
        session.findById(tblArea & "/ctxtRV45A-MABNR[1," & row & "]").setFocus
        session.findById("wnd[0]").sendVKey 2       'Двойной клик по позиции

        ' Заполнение текста
        ' session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").Select       'Тексты
        ' txtArea = UserArea.findByName("SPLITTER_CONTAINER", "GuiCustomControl").Id
        ' Set textField = session.findById(txtArea & "/shellcont/shellcont/shell/shellcont[1]/shell")
        ' textField.text = typeText

        WScript.Sleep 200
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13").select       'Допданные 

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").key = "00000"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_APLCD").key = "3199"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cmbVBAP-ZZ_PMPTY").setFocus

        'session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        pressF3()


        
    End If
    
End Sub

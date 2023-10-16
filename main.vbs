Public Const cBOM = 7, cDIN = 1, cArticle = 2, cDINArt = 3        'object
Public Const cEmpty = 10, cOne = 11, cMulti = 12, cDINWrong = 13, cArtiWrong = 14, cDINArtWrong = 15               'result
Public Const cSAP = 20, cExcel = 21, cBoth = 22                    'user information direction 

'Запрашиваем файл QTN
Dim excelFile
excelFile = selectExcel()

'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)

qtn = ArticlesExcel.Cells(22, 4).Value
Dim arrParts()

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = qtn
session.findById("wnd[0]").sendVKey 0




' Считаем, что в 25 строке - начало таблицы для обработки
intRow = 25
' Цикл для каждой строки
On Error Resume Next
Do Until ArticlesExcel.Cells(intRow,9).Value = ""
    Err.Clear
    tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
    Set grid = session.findById(tblArea)
    sapRow = grid.currentRow                'Here is the current visible row of the QTN
MsgBox "Row: " & sapRow

    If sapRow > 7 Then
        rowCount = grid.RowCount
        goto_pos = session.findById(tblArea & "/txtVBAP-POSNR[0," & sapRow - 5 & "]").text
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
        session.findById("wnd[1]/usr/txtRV45A-POSNR").text = goto_pos
        session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
        session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 300

        tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
        Set grid = session.findById(tblArea)
        sapRow = grid.currentRow                'Here is the current visible row of the QTN
        Set cell = grid.GetCell(sapRow + 6, 1)
        cell.setFocus()
        sapRow = grid.currentRow                'Here is the current visible row of the QTN

'MsgBox "new Row: " & sapRow
    
    End If        

    equipment = ArticlesExcel.Cells(intRow, 9).Value
    Qty = ArticlesExcel.Cells(intRow, 8).Value
    MsgBox "Excel item to go: " & ArticlesExcel.Cells(intRow, 2).Value
    
    ' Call Equipment dialog and input equipment
    session.findById("wnd[0]/tbar[1]/btn[22]").press
    session.findById("wnd[1]/usr/ctxtVIQMEL-EQUNR").text = equipment
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").text = ArticlesExcel.Cells(21, 4).Value    'Planning Plant
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").setFocus
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").caretPosition = 4
    session.findById("wnd[1]/usr/btnG_ENTER").press
    
    WScript.Sleep 500
    
    If session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART",False) Is Nothing Then
        ' BOM не существует - выполняем заполнение текстом
        session.findById("wnd[1]/usr/btnG_CANCEL").press
        Call InformUser(sapRow, cBOM, cEmpty, cBoth, "", ArticlesExcel, intRow, tblArea)        'Обработка 
    Else
        
        ' BOM существует
        session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART").text = ArticlesExcel.Cells(2, 4).Value   'Order type
        session.findById("wnd[2]/tbar[0]/btn[6]").press            'Structure List
        
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          1","999"
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          1","999"
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").clickLink "          1","999"            'Бинокль
        
        DIN = ArticlesExcel.Cells(intRow, 6).Value
        isDIN = DIN <> "-.-"
        Article = ArticlesExcel.Cells(intRow, 7).Value
        isArticle = Article <> "-"
        
        Case21 = ((Not isDIN) And isArticle)
        Case22 = (isDIN And (Not isArticle))
        Case23 = (isDIN And isArticle)
        
        'Сценарий 2.1
        If (Case21) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = Article
            obj = cArticle
MsgBox "Case 2.1"
        End If
        
        'Сценарий 2.2
        If (Case22) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = DIN
            obj = cDIN
MsgBox "Case 2.2"
        End If
        
        'Сценарий 2.3
        If (Case23) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = Article
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = DIN
            obj = cDINArt
MsgBox "Case 2.3"
        End If
        
        session.findById("wnd[1]").sendVKey 0                'Нажали Enter в окне Find
        
        'Анализ в окне выбора Structure List
        Set Parts = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetSelectedNodes()
        ReDim arrParts(Parts.Count(), 3)
        If (Not (Parts Is Nothing)) Then
            For i = 0 To Parts.Count() - 1
                nodekey = Parts.Item(i)
                session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem nodekey,"1" 
                WScript.Sleep 300
                arrParts(i,1) = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(nodekey, "1")
                arrParts(i,2) = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(nodekey, "6") 'DIN
                MsgBox ("Values: "  & arrParts(i,1) & " " & arrParts(i,2) )
            Next
        End If
        
        
        
        session.findById("wnd[0]/tbar[1]/btn[5]").press        'Нажали Галку в Structure List
        
        'Анализ - вернулись ли в основное окно?
        tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
        If session.findById(tblArea, False) Is Nothing Then
            'Не вернулись
            session.findById("wnd[1]/tbar[0]/btn[0]").press        'На сообщении нажали галку
            pressF3()                                               'Вернулись в главное окно
            Call InformUser(sapRow, obj, cEmpty, cBoth, "", ArticlesExcel, intRow, tblArea)
            
        Else
            
            'Анализ - сколько строк вставилось 
            tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
            Set grid = session.findById(tblArea)
            newsaprow = grid.currentRow - 1
            diff = newsaprow - sapRow + 1
MsgBox "Positions inserted: " & diff & ": " & sapRow & ", " & newsaprow

            WScript.Sleep 300
            
            ' Находим номера - вставленные позиции
            lines = ""
            For i = newsaprow To saprow Step - 1
                If lines <> "" Then
                    lines = lines & ", "
                End If
                lines = lines & session.findById(tblArea & "/txtVBAP-POSNR[0," & i & "]").text
                session.findById(tblArea & "/txtVBAP-ZMENG[2," & i & "]").text = Qty
            Next
            If lines <> "" Then
                lines = "[" & lines & "]"
            End If

            If Case21 Or Case22 Then
                If (diff = 1) Then
                    Call InformUser(sapRow, obj, cOne, cExcel, lines, ArticlesExcel, intRow, tblArea)
                End If
                If (diff > 1) Then
                    Call InformUser(sapRow, obj, cMulti, cExcel, lines, ArticlesExcel, intRow, tblArea)
                End If
            Else    ' Case23 - Both DIN and Article provided
'                MsgBox "Case 23" 
                If UBound(arrParts) = 1 Then
                    If arrParts(0,2) = DIN And arrParts(0,1) = Article Then
                        MsgBox "1: " & arrParts(0,1) & " " & Article & " " & arrParts(0,2) & " " & DIN 
                        Call InformUser(sapRow, obj, cOne, cExcel, lines, ArticlesExcel, intRow, tblArea)
                    ElseIf arrParts(0,2) <> DIN Then
                        MsgBox "2: " & arrParts(0,1) & " " & Article & " " & arrParts(0,2) & " " & DIN 
                        Call InformUser(sapRow, obj, cDINWrong, cExcel, lines, ArticlesExcel, intRow, tblArea)
                    ElseIf arrParts(0,1) <> Article Then
                        MsgBox "3: " & arrParts(0,1) & " " & Article & " " & arrParts(0,2) & " " & DIN 
                        Call InformUser(sapRow, obj, cArtiWrong, cExcel, lines, ArticlesExcel, intRow, tblArea)
                    End If
                Else 
                    Call InformUser(sapRow, obj, cDINArtWrong, cExcel, lines, ArticlesExcel, intRow, tblArea)
                End If
            End If
        End If'Articles entered
    End If'BOM exists
    
    intRow = intRow + 1
Loop

objWorkbook.Save
objWorkbook.Close False
ArticlesExcel.Quit

session.findById("wnd[0]/tbar[0]/btn[11]").press
pressEnter()

MsgBox "The script finished!", vbSystemModal Or vbInformation

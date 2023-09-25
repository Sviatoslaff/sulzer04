Public Const cBOM = 7, cDIN = 1, cArticle = 2, cDINArt = 3        'object
Public Const cEmpty = 10, cOne = 11, cMulti = 12                'result
Public Const cSAP = 20, cExcel = 21, cBoth = 22                    'user information direction 

'Запрашиваем файл QTN
Dim excelFile
excelFile = selectExcel()

'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)

qtn = ArticlesExcel.Cells(22, 4).Value

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = qtn
session.findById("wnd[0]").sendVKey 0

tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
Set grid = session.findById(tblArea)


' Считаем, что в 25 строке - начало таблицы для обработки
intRow = 25
' Цикл для каждой строки
'On Error Resume Next
Do Until ArticlesExcel.Cells(intRow,9).Value = ""
    Err.Clear
    sapRow = grid.currentRow                'Here is the current row of the QTN
    equipment = ArticlesExcel.Cells(intRow, 9).Value
    Qty = ArticlesExcel.Cells(intRow, 8).Value
    
    ' Call Equipment dialog and input equipment
    session.findById("wnd[0]/tbar[1]/btn[22]").press
    session.findById("wnd[1]/usr/ctxtVIQMEL-EQUNR").text = equipment
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").text = ArticlesExcel.Cells(21, 4).Value    'Planning Plant
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").setFocus
    session.findById("wnd[1]/usr/ctxtVIQMEL-IWERK").caretPosition = 4
    session.findById("wnd[1]/usr/btnG_ENTER").press
    
    WScript.Delay 500
    
    If session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART",False) Is Nothing Then
		' BOM не существует - выполняем заполнение текстом
        session.findById("wnd[1]/usr/btnG_CANCEL").press
        Call InformUser(sapRow, cBOM, cEmpty, cBoth, "", ArticlesExcel, intRow, tblArea)        'Обработка 
    Else
        
        ' BOM существует
        session.findById("wnd[2]/usr/ctxtVBAKKOM-AUART").text = ArticlesExcel.Cells(21, 4).Value
        session.findById("wnd[2]/tbar[0]/btn[6]").press            'Structure List
        
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          1","999"
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          1","999"
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").clickLink "          1","999"            'Бинокль
        
        DIN = ArticlesExcel.Cells(intRow, 6).Value
        isDIN = DIN <> "-.-"
        Article = ArticlesExcel.Cells(intRow, 7).Value
        isArticle = Article <> "-"
        
        Case21 = (Not isDIN And isArticle)
        Case22 = (isDIN And Not isArticle)
        Case23 = (isDIN And isArticle)
        
        'Сценарий 2.1
        If (Case21) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = Article
            obj = cArticle
        End If
        
        'Сценарий 2.2
        If (Case22) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = DIN
            obj = cDIN
        End If
        
        'Сценарий 2.3
        If (Case23) Then
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = Article
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[1,21]").text = DIN
            obj = cDINArt
        End If
        
        session.findById("wnd[1]").sendVKey 0                'Нажали Enter в окне Find
        
        session.findById("wnd[0]/tbar[1]/btn[5]").press        'Нажали Галку в Structure List
        
        'Анализ - вернулись ли в основное окно?
        If session.findById(tblArea, False) Is Nothing Then
            'Не вернулись
            session.findById("wnd[1]/tbar[0]/btn[0]").press        'На сообщении нажали галку
            pressF3()                                            'Вернулись в главное окно
            Call InformUser(sapRow, obj, cEmpty, cBoth, "", ArticlesExcel, intRow, tblArea)
        End If
        
        'Анализ - сколько строк вставилось 
        newsaprow = grid.currentRow
        diff = newsaprow - saprow
        
        lines = ""
        For i = newsaprow To saprow Step - 1
            If lines <> "" Then
                lines = lines & ", "
            End If
            lines = lines & session.findById(tblArea & "/txtVBAP-POSNR[0," & i & "]").text
        Next
        If lines <> "" Then
            lines = "[" & lines & "]"
        End If
        
        If (diff = 1) Then
            Call InformUser(sapRow, obj, cOne, cExcel, lines, ArticlesExcel, intRow, tblArea)
        End If
        If (diff > 1) Then
            Call InformUser(sapRow, obj, cMulti, cExcel, lines, ArticlesExcel, intRow, tblArea)
        End If
        
    End If	'
    
    intRow = intRow + 1
Loop

objWorkbook.Save
objWorkbook.Close False
ArticlesExcel.Quit

MsgBox "The script finished!", vbSystemModal Or vbInformation

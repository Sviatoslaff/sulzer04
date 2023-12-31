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
Set allnodes = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetAllNodeKeys()
tree = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetTreeType()
' Set colheaders = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetColumnNames()

' For i = 0 to colheaders.Count()-1
'    txthead = colheaders.Item(i)
'    MsgBox "'" & txthead & "'"
' Next
Dim arrParts()
Set Parts = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetSelectedNodes()
MsgBox Parts.Count()
ReDim Preserve arrParts(Parts.Count(), 2)
If (Not (Parts Is Nothing)) Then
    For i = 0 To Parts.Count() - 1
        nodekey = Parts.Item(i)
        MsgBox (i & " " & Parts.Item(i)) 
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem nodekey, "1"  
        newsize = i+1
        MsgBox ("Sizes:" & UBound(arrParts) & " ")
        arrParts(i, 0) = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(nodekey, "1")
        arrParts(i, 1) = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(nodekey, "3") 'DIN
        MsgBox (i & " " & arrParts(i,0) & " " & arrParts(i,1) )
    Next
End If

' For i = 0 to allnodes.Count()-1
'    anode = allnodes.Item(i)
'    txtnode = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetNodeTextByKey(anode)
'    txt1 = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(anode, "1")
'    txt2 = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetItemText(anode, "2")
'    MsgBox "2: '" & txt2 & "'"
' Next
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          2","3"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          2","3"


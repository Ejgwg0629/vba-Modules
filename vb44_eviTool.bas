Attribute VB_Name = "vb44_eviTool"
'ToDo:
'    1. check()
Sub dataSplit()
    For Each aRng In Selection
        Set wRng = aRng
        dataArr = Split(aRng.Value, "|", -1)
        For i = LBound(dataArr) + 1 To UBound(dataArr) - 1
            wRng.Value = "'" & dataArr(i)
            Set wRng = wRng.Offset(0, 1)
        Next
    Next
End Sub


Private Type pl1Def
    lnNum  As Long
    level  As Integer
    name   As String
    format As String
    Length As Integer
    jaName As String
End Type

'
'0. initSet
'  1. readAsArr
'1. collect
'  1. isComment
'  2. readLine   ->   resList


Sub anaPl1Def()
    Set aRng = Selection.Cells(1)
    
    Dim re As New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = "^\s*(\d+)\s+(.*)"
    
    Dim astr As String
    astr = aRng.Value
    
    Set result = re.Execute(astr)
    
    If Not result.count > 0 Then
        aRng.Select
        
    End If
    
End Sub

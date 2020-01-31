Attribute VB_Name = "vb10_prepare"

Sub checkDir(dir As String)
    
End Sub

Function getParameter(ByRef dsName As String, ByRef maclName As String, ByRef separator As String) As Boolean
    getParameter = False
    
    dsName = UCase(ActiveCell.Value)
    If Not checkDs(dsName) Then
        Exit Sub
    End If
    
    maclName = UCase(ActiveCell.End(xlToRight).Value)
    If Not checkDs(maclName) Then
        Exit Sub
    End If
    
    separator = ActiveCell.End(xlToRight).End(xlToRight).Value
    If Not checkSeparator(separator) Then
        Exit Sub
    End If
    If separator = "" And maclName <> "" Then
        separator = "|"
    End If
    
    getParameter = True
End Function

Function checkDs(ds As String) As Boolean
    checkDs = True
    If ds = "" Then
        Call MsgBox("dsName cannot be null")
        checkDs = False
    End If
End Function

Function checkSeparator(separator As String) As Boolean
    checkSeparator = True
    If Len(separator) > 1 Then
        Call MsgBox("separator should be exactly 1 byte")
        checkSeparator = False
    End If
End Function

Function mangleDs(ByVal ds As String) As String
    If Left(ds, 1) <> "'" Then
        ds = "'" & ds
    End If
    If Right(ds, 1) <> "'" Then
        ds = ds & "'"
    End If
    mangleDs = ds
End Function

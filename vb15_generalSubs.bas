Attribute VB_Name = "vb15_generalSubs"
Public Sub deleteBlankRows()
    Dim aRng As Range
    Dim temp As Range
    Set aRng = ActiveCell
    cnt = Selection.count
    i = 1
    Do
        If Trim(aRng.Value) = "" Then
            Set temp = aRng
            Set aRng = aRng.Offset(1, 0)
            Rows(temp.Row).Delete
        Else
            Set aRng = aRng.Offset(1, 0)
        End If
        i = i + 1
    Loop Until i > cnt
End Sub

Attribute VB_Name = "vb88_colorScheme"
Sub colorScheme()
    Set csSht = ActiveWorkbook.ActiveSheet
    Cells.Columns.AutoFit
    Dim cRng As Range
    For Each cRng In csSht.UsedRange
        If regexMatch(cRng.Value, "^#[\da-f]{6}$") Then
            Call setColor(cRng, 1)
        ElseIf regexMatch(cRng.Value, "^\(?\d{1,3},\d{1,3},\d{1,3}\)?$") Then
            Call setColor(cRng, 2)
        End If
    Next
End Sub
 
Sub setColor(ByRef target As Range, ByVal rgbType As Integer)
    If rgbType = 1 Then
        b = CLng("&H" & Right(target.Value, 2))
        g = CLng("&H" & Mid(target.Value, Len(target.Value) - 3, 2))
        r = CLng("&H" & Mid(target.Value, Len(target.Value) - 5, 2))
        target.Interior.Color = RGB(r, g, b)
    ElseIf rgbType = 2 Then
        rgbValue = target.Value
        If Left(rgbValue, 1) = "(" Then
            rgbValue = Mid(rgbValue, 2)
        End If
        If Right(rgbValue, 1) = ")" Then
            rgbValue = Mid(rgbValue, 1, Len(rgbValue) - 1)
        End If
        rgbArr = Split(rgbValue, ",")
        target.Interior.Color = RGB(rgbArr(0), rgbArr(1), rgbArr(2))
    End If
End Sub

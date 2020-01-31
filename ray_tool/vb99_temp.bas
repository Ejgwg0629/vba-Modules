Attribute VB_Name = "vb99_temp"
Sub renum()
    Dim aRng As Range
    Dim nRng As Range
    num = InputBox("input a number")
    num = CInt(num)
    For Each aRng In Selection
        Set nRng = aRng.End(xlToRight)
        If rMatch(nRng.Value, "^\d+(-\d+)*") Then
            If InStr(1, nRng, ".") <> 0 Then
                pos1 = InStr(1, nRng, ".")
            Else
                pos1 = 99999999
            End If
            If InStr(1, nRng, "-") <> 0 Then
                pos2 = InStr(1, nRng, "-")
            Else
                pos2 = 99999999
            End If
            If pos1 < pos2 Then
                sep = pos1
            Else
                sep = pos2
            End If
            lstr = Mid(nRng, 1, sep - 1)
            rstr = Mid(nRng, sep)
            lstr = CInt(lstr) + num
            nRng.Value = CStr(lstr) & rstr
        End If
    Next
End Sub

Function eval(STR As String)
    Application.Volatile
    eval = Application.Evaluate(STR)
End Function
Sub Test()
    Set dstRange = ActiveCell
End Sub

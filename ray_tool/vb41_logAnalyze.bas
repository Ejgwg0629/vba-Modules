Attribute VB_Name = "vb41_logAnalyze"
'insert blanks pre&aft full-width characters
Sub blankIsrt()
    Dim STR As String
    Dim oStr As String
    STR = ActiveCell.Text
    Dim i As Integer: i = 1
    Do
        ch = Mid(STR, i, 1)
        If AscW(ch) < 0 Or AscW(ch) > 255 Then
            oStr = oStr & " "
            Do
                oStr = oStr & ch
                i = i + 1
                ch = Mid(STR, i, 1)
            Loop Until AscW(ch) > 0 And AscW(ch) < 255
            oStr = oStr & " "
        End If
        oStr = oStr & ch
        i = i + 1
    Loop Until i > Len(STR)
    
    ActiveCell.Value = oStr
End Sub

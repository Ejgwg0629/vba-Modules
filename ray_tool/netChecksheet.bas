Attribute VB_Name = "netChecksheet"
Sub deleteSheetForRerun()
    For Each sht In Sheets
        shtName = Left(sht.name, 4)
        If shtName = "NET_" Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub

Sub copySheet()
    Application.ScreenUpdating = False

    deleteSheetForRerun
    Dim srcRng As Range
    Dim fromsheet As Worksheet
    Set fromsheet = Sheets("from")
    Set srcRng = fromsheet.Range("a1")
    srcName = srcRng.Value
    i = 1
    While srcName <> ""
        Sheets("NET").Copy before:=Sheets("from")
        ActiveSheet.name = "NET_" & STR(i)
        
        ActiveSheet.Range("i5").Value = "NET：" & srcRng.Value
        Set srcRng = srcRng.Offset(1, 0)
        
        ActiveSheet.Range("k5").Value = "NET：" & srcRng.Value
        Set srcRng = srcRng.Offset(1, 0)
        
        ActiveSheet.Range("i30").Value = "NET：" & srcRng.Value
        Set srcRng = srcRng.Offset(1, 0)
                
        ActiveSheet.Range("k30").Value = "NET：" & srcRng.Value
        Set srcRng = srcRng.Offset(1, 0)
        
        srcName = srcRng.Value
        i = i + 1
    Wend
    Application.ScreenUpdating = True
End Sub

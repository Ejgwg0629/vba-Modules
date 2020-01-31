Attribute VB_Name = "joblib_req"
Sub copySheet()
    Application.ScreenUpdating = False
    deleteSheetForRerun
    Dim srcRng As Range
    Dim netsheet As Worksheet
    Set netsheet = Sheets("NET")
    Set srcRng = netsheet.Range("d7")
    srcName = srcRng.Value
    While srcName <> ""
        Sheets("ネット申請").Copy before:=Sheets("データ待ち申請")
        ActiveSheet.name = "ネット申請" & "_" & srcName
        
        setval srcRng, netsheet
        
        Set srcRng = srcRng.Offset(1, 0)
        srcName = srcRng.Value
    Wend
    Application.ScreenUpdating = True
End Sub

Sub deleteSheetForRerun()
    For Each sht In Sheets
        shtName = Left(sht.name, 3)
        If shtName = "移行_" Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub

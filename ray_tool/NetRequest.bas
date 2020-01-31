Attribute VB_Name = "NetRequest"
Sub copySheet()
    Application.ScreenUpdating = False
    Sheets("ネット申請").Range("s4") = Date
    deleteSheetForRerun
    Dim srcRng As Range
    Dim netsheet As Worksheet
    Set netsheet = Sheets("NET")
    Set srcRng = netsheet.Range("d7")
    srcName = srcRng.Value
    While srcName <> ""
        Sheets("ネット申請").Copy before:=Sheets("ネット申請")
        ActiveSheet.name = "ネット申請" & "_" & srcName
        
        setval srcRng, netsheet
        
        Set srcRng = srcRng.Offset(1, 0)
        srcName = srcRng.Value
    Wend
    Application.ScreenUpdating = True
End Sub

Private Sub setval(ByRef srcRng As Range, ByRef srcSht As Worksheet)
    Dim toRng As Range
    Dim buff() As Variant
    Dim timr As String: timr = srcSht.Range("o" & srcRng.Row)
    Dim netid As String: netid = srcSht.Range("d" & srcRng.Row)
    Dim prenetid As String: prenetid = srcSht.Range("n" & srcRng.Row)
    Dim jobid As String: jobid = srcSht.Range("y" & srcRng.Row)
    Dim netname As String: netname = srcSht.Range("e" & srcRng.Row)
    Dim schdid As String: schdid = srcSht.Range("i" & srcRng.Row)
    
    ActiveSheet.Range("s5").Value = srcSht.Range("c" & srcRng.Row)
    ActiveSheet.Range("s9").Value = srcSht.Range("e" & srcRng.Row)
    
    
    If timr <> "無" Then
        arr = Split(format(timr, "hh:mm"), ":")
        Set toRng = ActiveSheet.Range("s10")
        copyTime arr, toRng
    End If
    
    
    
    ReDim buff(Len(schdid) - 1)
    For i = 1 To Len(schdid)
        buff(i - 1) = Mid$(schdid, i, 1)
    Next
    Set toRng = ActiveSheet.Range("s48")
    copyArr buff, toRng
    
    

    ReDim buff(Len(netid) - 1)
    For i = 1 To Len(netid)
        buff(i - 1) = Mid$(netid, i, 1)
    Next
    Set toRng = ActiveSheet.Range("s8")
    copyArr buff, toRng
    
    
    
    ReDim buff(Len(jobid) - 1)
    For i = 1 To Len(jobid)
        buff(i - 1) = Mid$(jobid, i, 1)
    Next
    Set toRng = ActiveSheet.Range("v28")
    copyArr buff, toRng
    If jobid = "DUMMY" Then
        ActiveSheet.Range("as27") = "D"
    End If
    
    
    
    If prenetid <> "-" Then
        ReDim buff(Len(prenetid) - 1)
        For i = 1 To Len(prenetid)
            buff(i - 1) = Mid$(prenetid, i, 1)
        Next
        Set toRng = ActiveSheet.Range("s19")
        copyArr buff, toRng
    End If
End Sub

Sub copyTime(ByRef fromArr As Variant, ByRef toRng As Range)
    toRng.Value = fromArr(0)
    toRng.Offset(0, 1).Value = fromArr(1)
End Sub

Sub copyArr(ByRef fromArr() As Variant, ByRef toRng As Range)
    
    Set svrng = toRng
    For Each sgChr In fromArr
        svrng.Value = sgChr
        Set svrng = svrng.Offset(0, 1)
    Next
End Sub


Sub deleteSheetForRerun()
    For Each sht In Sheets
        shtName = Left(sht.name, 6)
        If shtName = "ネット申請_" Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub


Sub changeAllSheetViews()
    For Each sht In Worksheets
        If sht.Visible Then
            sht.Select
            ActiveWindow.View = xlNormalView
            ActiveWindow.Zoom = 145
        End If
    Next
End Sub

Sub deleteAllNames()
    For Each nm In ActiveSheet.Names
        nm.Delete
    Next
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next
End Sub

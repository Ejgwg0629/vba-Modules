Attribute VB_Name = "vb60_WBS"
Sub sheetView()
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    
    If ActiveSheet.ListObjects(1).ShowAutoFilter Then
        ActiveSheet.ListObjects(1).AutoFilter.ShowAllData
    End If
    
    ActiveSheet.Rows("1:6").Hidden = True
    ActiveSheet.Rows("8:26").Hidden = True
    
    ActiveWindow.SplitRow = 21
    ActiveWindow.FreezePanes = True

    ActiveSheet.ListObjects("ÉeÅ[ÉuÉã2").Range.AutoFilter Field:=5, Criteria1:="óã"
End Sub

Sub gradeTimeInput()
    For Each aRng In Selection
        If aRng.Column = 21 Then
            totalTime = Range("I" & aRng.Row).Value
            aRng.Offset(0, 0).Value = totalTime * 0.5
            aRng.Offset(0, 1).Value = totalTime * 0.2
            aRng.Offset(0, 2).Value = totalTime * 0.1
            aRng.Offset(0, 3).Value = totalTime * 0.2
        ElseIf aRng.Column = 25 Then
            totalTime = Range("I" & aRng.Row).Value
            aRng.Offset(0, 0).Value = totalTime * 0.22
            aRng.Offset(0, 1).Value = totalTime * 0.1
            aRng.Offset(0, 2).Value = totalTime * 0.15
            aRng.Offset(0, 3).Value = totalTime * 0.15
            aRng.Offset(0, 4).Value = totalTime * 0.18
            aRng.Offset(0, 5).Value = totalTime * 0.1
            aRng.Offset(0, 6).Value = totalTime * 0.05
            aRng.Offset(0, 7).Value = totalTime * 0.03
            aRng.Offset(0, 8).Value = totalTime * 0.02
        ElseIf aRng.Column = 34 Then
            totalTime = Range("I" & aRng.Row).Value
            aRng.Offset(0, 0).Value = totalTime * 0.25
            aRng.Offset(0, 1).Value = totalTime * 0.15
            aRng.Offset(0, 2).Value = totalTime * 0.15
            aRng.Offset(0, 3).Value = totalTime * 0.15
            aRng.Offset(0, 4).Value = totalTime * 0.2
            aRng.Offset(0, 5).Value = totalTime * 0.1
            aRng.Offset(0, 6).Value = totalTime * 0
        End If
    Next
End Sub

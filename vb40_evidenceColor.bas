Attribute VB_Name = "vb40_evidenceColor"
'***********************************************
' eviColor
'***********************************************
Private foldHeaderFlg
Private addStarFlg

Sub evidenceColor()
    foldHeaderFlg = True
    addStarFlg = False
    
    If Not Selection.count > 1 Then
        Range("C1").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.FormatConditions.Delete
    End If
    
    Application.ScreenUpdating = False
    Call frameGen
   
    Dim eRng As Range
    Dim eRngStr As String
    For Each eRng In Selection
        eRngStr = eRng.Value
        If InStr(1, eRngStr, "$", vbTextCompare) <> 0 Then
            Call compareColor1(eRng)
        End If
        If InStr(1, eRngStr, "\", vbTextCompare) <> 0 Then
            Call compareColor2(eRng)
        End If
        If InStr(1, eRngStr, "DSN=", 1) <> 0 Then
            'Call dsColor(eRng)
        End If
        If InStr(1, Left(eRngStr, 8), "IEF472I", vbTextCompare) <> 0 Then
            eRng.Resize(1, 22).Interior.Color = &HCEC7FF
            eRng.Font.Color = &H6009C
        End If
        If InStr(1, Left(eRngStr, 8), "IEF142I", vbTextCompare) <> 0 Or _
           InStr(1, Left(eRngStr, 8), "IEF272I", vbTextCompare) <> 0 Then
            Call stepResult(eRng)
        End If
        If Left(eRngStr, 10) = "1  ISRSUPC" And foldHeaderFlg Then
            Call compareHeader(eRng)
        End If
'        If rMatch(str, "(^\d+//PANNEWEXEC)|(^\d+//PANOLDEXEC)|(^\d+//NEWDDDDDISP=SHR)|(^\d+//OLDDDDDDISP=SHR)") Then
'            eRng.Resize(1, 18).Interior.Color = &H2EE2A6
'        End If

'        If InStr(1, Left(eRngStr, 9), "5668-910", vbTextCompare) <> 0 Then
'           Call compileContent
'        End If
    Next
    Application.ScreenUpdating = True
    Selection.Cells(1).Activate
End Sub

Sub frameGen()
    With Selection
        .Offset(0, -1).Borders(xlEdgeLeft).LineStyle = xlContinous
        .Offset(0, -1).Borders(xlEdgeLeft).ColorIndex = 0

        .Cells(1).Offset(0, -1).Resize(1, 46).Borders(xlEdgeTop).LineStyle = xlContinous
        .Cells(1).Offset(0, -1).Resize(1, 46).Borders(xlEdgeTop).ColorIndex = 0

        .Cells(1).Offset(0, 44).Resize(.count, 1).Borders(xlEdgeRight).LineStyle = xlContinous
        .Cells(1).Offset(0, 44).Resize(.count, 1).Borders(xlEdgeRight).ColorIndex = 0

        .Offset(0, -1).Resize(Selection.count, 46).Interior.Color = &HFFFFFF
        .Font.Color = &H0
    End With
End Sub

Sub compareColor1(ByRef eRng As Range)
    Dim STR As String
    STR = WorksheetFunction.Substitute(eRng.Value, " ", "")
    If regexMatch(STR, "(^(\$+)\|\1)|(^I-\$+\|)|(^\|D-\$+)|(^RN-\$+\|RO-\$+)") Then
        If addStarFlg Then
            eRng.Offset(0, -2).Value = "š"
        End If
        eRng.Offset(0, -1).Resize(1, 46).Interior.Color = &HCEC7FF
        eRng.Font.Color = &H6009C
    End If
End Sub

Sub compareColor2(ByRef eRng As Range)
    Dim STR As String
    STR = WorksheetFunction.Substitute(eRng.Value, " ", "")
    If regexMatch(STR, "(^(\\+)\|\1)|(^I-\\+\|)|(^\|D-\\+)|(^RN-\\+\|RO-\\+)") Then
        If addStarFlg Then
            eRng.Offset(0, -2).Value = "š"
        End If
        eRng.Offset(0, -1).Resize(1, 46).Interior.Color = &HCEC7FF
        eRng.Font.Color = &H6009C
    End If
End Sub

Sub dsColor(ByRef eRng As Range)
    Dim STR As String
    STR = WorksheetFunction.Substitute(eRng.Value, " ", "")

    If rMatch(STR, "//.*?DSN=TCRD\.SCS(\.UG)?71O03.*SMT") Then
        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent4
        
    ElseIf rMatch(STR, "//.*?DSN=TCRD\.SCS(\.UG)?71O03.*[BM]IS\d{1,2}") Then
        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent4
        
    ElseIf rMatch(STR, "//.*?DSN=TCRD\.SCS(\.UG)?71O03.*[BM]OS\d{1,2}") Then
        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent6
        
    ElseIf rMatch(STR, "//.*?DSN=TCRD\.SCS(\.UG)?71O03.*C[IO]S\d{1,2}") Then
        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent5
        
'    ElseIf rMatch(str, "//.*?DSN=TGD1\.UT\d{2}.*ST\d{2}\.H\d{2}") Then
'        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent3
'
'    ElseIf rMatch(str, "//.*?DSN=TRE1\.UT\dUT\d{2}.*ST\d{2}\.[ISROW]\d{3}") Then
'        eRng.Resize(1, 20).Interior.ThemeColor = xlThemeColorAccent2
    
    End If
End Sub

Sub stepResult(ByRef eRng As Range)
    Dim STR As String
    STR = WorksheetFunction.Substitute(eRng.Value, " ", "")

    If regexMatch(STR, "CONDCODE0000") Then
        eRng.Resize(1, 17).Interior.Color = &HCEEFC6
        eRng.Font.Color = &H6100
    ElseIf regexMatch(STR, "WASNOTEXECUTED") Then
        eRng.Resize(1, 17).Interior.Color = &H436AED
        eRng.Font.Color = &H5D1DA7
    Else
        eRng.Resize(1, 17).Interior.Color = &H9CEBFF
        eRng.Font.Color = &H659C
    End If
End Sub

Sub compileContent(ByRef eRng As Range)
    Dim STR As String
    STR = WorksheetFunction.Substitute(eRng.Value, " ", "")

    pgCtr = regexRepl(STR, ".*(\d+)$")
    If pgCtr > 2 And _
        (Mid(eRng.Offset(1, 0).Value, 5, 11) = "STMT LEV NT" _
         Or Left(eRng.Offset(1, 0).Value, 7) = "DCL NO.") Then
        Rows(eRng.Row).Hidden = True
        Rows(eRng.Row + 1).Hidden = True
    End If
End Sub

Sub compareHeader(ByRef eRng As Range)
    Dim STR As String
    STR = RTrim(eRng.Value)
    
    pgCtr = LTrim(Right(STR, 4))
    If pgCtr >= 2 And _
            Left(eRng.Offset(5, 0), 3) = " ID" Then
        Rows(eRng.Row & ":" & eRng.Row + 6).Hidden = True
    End If
    If Left(eRng.Offset(6, 0), 13) = "    ----+----" Then
        eRng.Offset(6, 0).Font.Color = &HA6A6A6
    End If
End Sub

Sub makeActiveColumnC()
'    ActiveSheet.Shapes.SelectAll
'    Selection.Placement = xlFreeFloating
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    
    Call vb42_style.pageRefresh
    Call evidenceColor
End Sub

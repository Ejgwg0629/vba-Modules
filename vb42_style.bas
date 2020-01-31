Attribute VB_Name = "vb42_style"
Sub pageRefresh()
    With Application
        .WindowState = xlNormal
        .Width = 1183
        .Height = 670
        .Calculation = xlCalculationAutomatic
    End With
    
    ActiveSheet.Cells.Select
    Selection.ColumnWidth = 3.4
    Columns("a").ColumnWidth = 1.7
    Columns("av").ColumnWidth = 1.7
    
    With Selection.Font
        .name = "HG恨集窶"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Columns("b:av").NumberFormatLocal = "@"
    ActiveWindow.Zoom = 100
    ActiveSheet.Range("A1").Select
    
    For Each aRng In ActiveSheet.Hyperlinks
        aRng.Range.Font.Color = RGB(0, 0, 255)
    Next
End Sub

Sub oldPageRefresh()
    Application.Width = 1183
    Application.Height = 670
    ActiveSheet.Cells.Select
    ActiveWindow.DisplayGridlines = False
    With Selection.Font
        .name = "HG恨集窶"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ' Selection.NumberFormatLocal = "@"
    ActiveWindow.Zoom = 100
    ActiveSheet.Range("A1").Select
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub shapeResize()
    Selection.ShapeRange.LockAspectRatio = msoFalse
    Selection.ShapeRange.Width = 610
    Selection.ShapeRange.Height = 580
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub

Sub shapePlacement()
    For Each sht In Worksheets
        For Each shp In sht.Shapes
            shp.Placement = xlFreeFloating
            'shp.Placement = xlMove
        Next
    Next
End Sub

Sub bDelete()
    i = ActiveCell.Offset(5, 0).Row
    Rows(i).Delete
    Rows(i).Delete
    Rows(i).Delete
    Rows(i + 1).Delete
    Rows(i + 1).Delete
    i = ActiveCell.Row
    Rows(i).Delete
    Rows(i).Delete
    Rows(i + 1).Delete
End Sub

Sub colorful()
    Selection.Interior.Color = ranColor()
End Sub

Sub stylish()
    Set aSht = ActiveSheet
    Application.ScreenUpdating = False
    For Each sht In Worksheets
        If sht.Visible Then
            sht.Select
            Cells.Select
            Dim fontNum
            fontNum = Int((3 * Rnd) + 1)
            Select Case fontNum
            Case 1
                Selection.Font.name = "HG娵恨集窶-PRO"
                If ActiveSheet.Shapes.count > 0 And ActiveSheet.Shapes.count > ActiveSheet.Comments.count Then
                    ActiveSheet.Shapes.SelectAll
                    If Selection.ShapeRange.TextFrame2.HasText Then
                        On Error GoTo nextSheet
                        With Selection.ShapeRange.TextFrame2.TextRange.Font
                            .NameComplexScript = "HG娵恨集窶-PRO"
                            .NameFarEast = "HG娵恨集窶-PRO"
                            .name = "HG娵恨集窶-PRO"
                        End With
                    End If
                End If
            Case 2
                Selection.Font.name = "Meiryo UI"
                If ActiveSheet.Shapes.count > 0 And ActiveSheet.Shapes.count > ActiveSheet.Comments.count Then
                    ActiveSheet.Shapes.SelectAll
                    If Selection.ShapeRange.TextFrame2.HasText Then
                        On Error GoTo nextSheet
                        With Selection.ShapeRange.TextFrame2.TextRange.Font
                            .NameComplexScript = "Meiryo UI"
                            .NameFarEast = "Meiryo UI"
                            .name = "Meiryo UI"
                        End With
                    End If
                End If
            Case 3
                Selection.Font.name = "HG恨集窶"
                If ActiveSheet.Shapes.count > 0 And ActiveSheet.Shapes.count > ActiveSheet.Comments.count Then
                    ActiveSheet.Shapes.SelectAll
                    If Selection.ShapeRange.TextFrame2.HasText Then
                        On Error GoTo nextSheet
                        With Selection.ShapeRange.TextFrame2.TextRange.Font
                            .NameComplexScript = "HG恨集窶"
                            .NameFarEast = "HG恨集窶"
                            .name = "HG恨集窶"
                        End With
                    End If
                End If
'            Case 4
'                Selection.Font.name = "HG惓灢彂懱-PRO"
'            Case 5
'                Selection.Font.name = "DFKai-SB"
'            Case 6
'                Selection.Font.name = "HG嫵壢彂懱"
            Case Else
                Selection.Font.name = "HG恨集窶"
                If ActiveSheet.Shapes.count > 0 And ActiveSheet.Shapes.count > ActiveSheet.Comments.count Then
                    ActiveSheet.Shapes.SelectAll
                    If Selection.ShapeRange.TextFrame2.HasText Then
                        On Error GoTo nextSheet
                        With Selection.ShapeRange.TextFrame2.TextRange.Font
                            .NameComplexScript = "HG恨集窶"
                            .NameFarEast = "HG恨集窶"
                            .name = "HG恨集窶"
                        End With
                    End If
                End If
            End Select
            ActiveWindow.DisplayGridlines = False
        End If
nextSheet:
        sht.Range("A1").Select
    Next
    Application.ScreenUpdating = True
    aSht.Select
End Sub

Sub borderChange()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("D24:T25").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("D26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Underline = xlUnderlineStyleSingle
    Range("G26").Select

End Sub


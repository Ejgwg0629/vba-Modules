Attribute VB_Name = "vb85_olAppointment"
' global configuration
Private Const henreiCol% = 3
Private Const seikyuCol% = 4
 
' * get first date cell in Range("A:A")
' * get the date cell of today
' * delete all old appointments
' * loop
'   * find a special day
'   * add it to calendar
Sub main()
    Dim todayCell As Range
    Set todayCell = getFirstDate()
    
    ' get todayCell to be todayCell
    If Not getTodayCell(todayCell) Then
        Exit Sub
    End If
    delOldAppointment
    Do While todayCell.Value <> ""
        If isSpecialDay(todayCell) Then
            Call addAppointment(todayCell)
        End If
        Set todayCell = todayCell.Offset(1, 0)
    Loop
End Sub
 
Function getFirstDate() As Range
    Set firstCell = Range("A1")
    Dim staPos As Integer
    staPos = 0
    Do While staPos < 15
        If IsDate(firstCell.Offset(staPos, 0)) Then
            Exit Do
        End If
        staPos = staPos + 1
    Loop
    Set getFirstDate = firstCell.Offset(staPos, 0)
End Function
 
Function getTodayCell(ByRef firstCell As Range) As Boolean
    Dim checked As Integer
    checked = 1
    Set aRng = firstCell.Offset(checked, 0)
    Do While checked < 5000 And aRng.Value <> ""
        If aRng - aRng.Offset(-1, 0) <> 1 Then
            Debug.Print "weird date order.."
            getTodayCell = False
            Exit Function
        End If
        If aRng = Date Then
            Set firstCell = aRng
        End If
        Set aRng = aRng.Offset(1, 0)
        checked = checked + 1
    Loop
    getTodayCell = True
End Function
 
Sub delOldAppointment()
    Set olAppointments = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(9).Items
    olAppointments.Sort "[Start]"
    olAppointments.IncludeRecurrences = False
    Set myAppointments = olAppointments.Restrict("[Start] >= '" & Date & "'")
    For i = myAppointments.count To 1 Step -1
        Set eachAppoint = myAppointments(i)
        If TimeValue(format(eachAppoint.Start, "hh:mm:ss")) >= TimeValue("13:00:00") And _
           TimeValue(format(eachAppoint.End, "hh:mm:ss")) <= TimeValue("13:30:00") And _
          (InStr(eachAppoint.Subject, "ä˙ì˙") <> 0 Or InStr(eachAppoint.Subject, "ï‘ñﬂ") <> 0 Or _
           InStr(eachAppoint.Subject, "êøãÅ") <> 0) Then
            eachAppoint.Delete
        End If
    Next
End Sub
 
Function isSpecialDay(ByVal todayCell As Range) As Boolean
    Set henrei = todayCell.Offset(0, henreiCol)
    Set seikyu = todayCell.Offset(0, seikyuCol)
    isSpecialDay = True
    If Not (InStr(henrei.Value, "ä˙ì˙") <> 0 And Day(todayCell) <> 2 And Day(todayCell) <> 17) Then
        If InStr(henrei.Value, "ï‘ñﬂ") = 0 Then
            If InStr(seikyu.Value, "êøãÅ") = 0 Then
                isSpecialDay = False
            End If
        End If
    End If
End Function
 
Function getSubject(ByVal currentCell As Range) As String
    Set henrei = currentCell.Offset(0, henreiCol)
    Set seikyu = currentCell.Offset(0, seikyuCol)
    If henrei.Value <> "" And (InStr(henrei.Value, "ä˙ì˙") <> 0 Or _
       InStr(henrei.Value, "ï‘ñﬂ") <> 0) Then
        getSubject = henrei.Value
    End If
    If seikyu.Value <> "" And InStr(seikyu.Value, "êøãÅ") <> 0 Then
        If getSubject <> "" Then
            getSubject = getSubject & " & " & seikyu.Value
        Else
            getSubject = seikyu.Value
        End If
    End If
    If getSubject = "" Then
        Debug.Print "weird, " & format(currentCell, "yyyy-mm-dd") & " is not a special day.."
    End If
End Function
 
Function nextWorkday(ByVal currentCell As Range) As String
    Do
        Set currentCell = currentCell.Offset(1, 0)
        Set henrei = currentCell.Offset(0, henreiCol)
    Loop Until henrei.Interior.Color = 16777215 And _
               henrei.MergeArea.count = 1
    nextWorkday = format(currentCell, "yyyy/mm/dd")
End Function
 
Sub addAppointment(ByVal currentCell As Range)
    Set olApp = CreateObject("Outlook.Application")
    ' olAppointmentItem <-- 1
    Set appointItem = olApp.CreateItem(1)
    With appointItem
        .Subject = getSubject(currentCell)
        .Location = "é©ê»"
        .Start = nextWorkday(currentCell) & " 13:00:00"
        .End = nextWorkday(currentCell) & " 13:30:00"
        .Categories = olApp.GetNamespace("MAPI").Categories.Item(9)
        .Save
    End With
End Sub
 

 
 
'
' so cool !!!
'
'#Const LateBind = True
'
'Const olMinimized As Long = 1
'Const olMaximized As Long = 2
'Const olFolderInbox As Long = 6
'
'#If LateBind Then
'
'Public Function OutlookApp( _
'    Optional WindowState As Long = olMinimized, _
'    Optional ReleaseIt As Boolean = False _
'    ) As Object
'    Static o As Object
'#Else
'Public Function OutlookApp( _
'    Optional WindowState As Outlook.OlWindowState = olMinimized, _
'    Optional ReleaseIt As Boolean _
') As Outlook.Application
'    Static o As Outlook.Application
'#End If
'On Error GoTo ErrHandler
'
'    Select Case True
'        Case o Is Nothing, Len(o.name) = 0
'            Set o = GetObject(, "Outlook.Application")
'            If o.Explorers.Count = 0 Then
'InitOutlook:
'                'Open inbox to prevent errors with security prompts
'                o.Session.GetDefaultFolder(olFolderCalendar).Display
'                o.ActiveExplorer.WindowState = WindowState
'            End If
'        Case ReleaseIt
'            Set o = Nothing
'    End Select
'    Set OutlookApp = o
'
'ExitProc:
'    Exit Function
'ErrHandler:
'    Select Case Err.Number
'        Case -2147352567
'            'User cancelled setup, silently exit
'            Set o = Nothing
'        Case 429, 462
'            Set o = GetOutlookApp()
'            If o Is Nothing Then
'                Err.Raise 429, "OutlookApp", "Outlook Application does not appear to be installed."
'            Else
'                Resume InitOutlook
'            End If
'        Case Else
'            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
'    End Select
'    Resume ExitProc
'    Resume
'End Function
'
'#If LateBind Then
'Private Function GetOutlookApp() As Object
'#Else
'Private Function GetOutlookApp() As Outlook.Application
'#End If
'On Error GoTo ErrHandler
'
'    Set GetOutlookApp = CreateObject("Outlook.Application")
'
'ExitProc:
'    Exit Function
'ErrHandler:
'    Select Case Err.Number
'        Case Else
'            'Do not raise any errors
'            Set GetOutlookApp = Nothing
'    End Select
'    Resume ExitProc
'    Resume
'End Function
'
'Sub MyMacroThatUseOutlook()
'    Dim OutApp  As Object
'    Set OutApp = OutlookApp()
'    'Automate OutApp as desired
'End Sub
'
'
'



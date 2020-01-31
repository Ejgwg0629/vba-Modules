Attribute VB_Name = "vb46_hostFileRecv"
Private Const fileDir$ = "C:\Users\scs-tlei\Desktop\DATA\"

Public Sub getFileFromHost()
    Call checkDir(fileDir)
    Dim dsName As String
    dsName = UCase(ActiveCell.Value)
    If Not checkDs(dsName) Then
        Debug.Print "weird dataset name.."
        Exit Sub
    End If
    Call doDownload(dsName)
    
    Call readFile(dsName, ActiveCell)
    Call vb40_evidenceColor.evidenceColor
End Sub

Public Sub getFileDividedFromHost()
    Call checkDir(fileDir)
    
    Dim dsName, maclName, separator As String
    If Not getParameter(dsName, maclName, separator) Then
        Exit Sub
    End If
    
    
End Sub

Sub checkDir(dir As String)
    
End Sub

Function getParameter(ByRef dsName As String, ByRef maclName As String, ByRef separator As String) As Boolean
    getParameter = False
    
    dsName = UCase(ActiveCell.Value)
    If Not checkDs(dsName) Then
        Exit Sub
    End If
    
    maclName = UCase(ActiveCell.End(xlToRight).Value)
    If Not checkDs(maclName) Then
        Exit Sub
    End If
    
    separator = ActiveCell.End(xlToRight).End(xlToRight).Value
    If Not checkSeparator(separator) Then
        Exit Sub
    End If
    If separator = "" And maclName <> "" Then
        separator = "|"
    End If
    
    getParameter = True
End Function

Function checkDs(ds As String) As Boolean
    checkDs = True
    If ds = "" Then
        Call MsgBox("dsName cannot be null")
        checkDs = False
    End If
End Function

Function checkSeparator(separator As String) As Boolean
    checkSeparator = True
    If Len(separator) > 1 Then
        Call MsgBox("separator should be exactly 1 byte")
        checkSeparator = False
    End If
End Function

Function mangleDs(ByVal ds As String) As String
    If Left(ds, 1) <> "'" Then
        ds = "'" & ds
    End If
    If Right(ds, 1) <> "'" Then
        ds = ds & "'"
    End If
    mangleDs = ds
End Function

Sub doDownload(ds As String)
    Set sess = CreateObject("PCOMM.autECLSession")
    'sess.setConnectionByName ("A")
    'Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLConnMgr = CreateObject("PCOMM.autECLConnMgr")
    Call autECLConnMgr.autECLConnList.Refresh
    If autECLConnMgr.autECLConnList.count <> 1 Then
        Call MsgBox("wried, the count of connections is not 1..")
    End If
    sess.SetConnectionByHandle (autECLConnMgr.autECLConnList(1).Handle)
    
    sess.autECLPS.SendKeys "[HOME]CMDE[ENTER]"
    sess.autECLOIA.WaitForInputReady
    Set mxfer = sess.autECLXfer
    
    If mxfer.Ready = False Then
        Set mxfer = Nothing
        Debug.Print "weird, mxfer not ready.."
        Exit Sub
    End If
    
    Call mxfer.ReceiveFile( _
        fileDir & ds, _
        mangleDs(ds), _
        "[JISCII CRLF SO NOCLEAR BLANK" _
    )
    
    Set autECLConnMgr = Nothing
    Set sess = Nothing
End Sub

Sub readFile(ds As String, fromCell As Range)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    filePath = fileDir & ds
    Set fStream = FSO.OpenTextFile(filePath, ForReading, False)
    
    Dim line As String
    Dim result() As String

    arraySize = 2999
    index = 0
    count = 0
    ReDim result(arraySize)
    While Not fStream.AtEndOfStream
        If index > arraySize Then
            fromCell.Resize(index, 1).Value2 = Application.Transpose(result)
            Set fromCell = fromCell.Offset(index, 0)
            
            count = count + index
            index = 0
            ReDim result(arraySize)
        End If
        line = fStream.ReadLine()
        line = replace(replace(replace(line, Chr(30), " "), Chr(31), " "), Chr(253), "!")
        result(index) = line

        index = index + 1
    Wend
    
    fStream.Close
    Set fStream = Nothing
    Set FSO = Nothing
    ReDim Preserve result(index - 1)
    
    fromCell.Resize(index, 1).Value2 = Application.Transpose(result)
    Set fromCell = fromCell.Offset(index, 0)
    Range(ActiveCell, fromCell).Select
    count = count + index
End Sub


Sub cmd()
    Set sess = CreateObject("PCOMM.autECLSession")
    'sess.setConnectionByName ("A")
    'Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLConnMgr = CreateObject("PCOMM.autECLConnMgr")
    Call autECLConnMgr.autECLConnList.Refresh
    If autECLConnMgr.autECLConnList.count <> 1 Then
        Call MsgBox("wried, the count of connections is not 1..")
    End If
    sess.SetConnectionByHandle (autECLConnMgr.autECLConnList(1).Handle)
    sess.autECLPS.SendKeys "[HOME]CMDE[ENTER]"
    sess.autECLOIA.WaitForInputReady
    Dim astr As String
    astr = """" & "TCRD.SCS71O03.JG5P606.DG5P105A.MIS02""    ""DG5P105""       ""|"""
    astr = "EXEC 'TCRD.UG71O03.REXX(TT)' '" & astr & "'"
    Debug.Print astr
    Call sess.autECLPS.SendKeys(astr)
    sess.autECLOIA.WaitForInputReady
    If (sess.autECLPS.waitForString("TEMP025 START", 1, 13, 100)) Then
        sess.autECLPS.SendKeys "[ENTER]"
        dsName = "UG71O03.TEMP.OUT01"
        ActiveCell.Value = "UG71O03.TEMP.OUT01"
        Call getFileFromHost
    Else

    End If
End Sub

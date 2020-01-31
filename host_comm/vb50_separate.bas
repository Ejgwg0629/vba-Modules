Attribute VB_Name = "vb50_separate"
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
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    filePath = fileDir & ds
    Set fStream = fso.OpenTextFile(filePath, ForReading, False)
    
    Dim line As String
    Dim result() As String

    arraySize = 2999
    Index = 0
    count = 0
    ReDim result(arraySize)
    While Not fStream.AtEndOfStream
        If Index > arraySize Then
            fromCell.Resize(Index, 1).Value2 = Application.Transpose(result)
            Set fromCell = fromCell.Offset(Index, 0)
            
            count = count + Index
            Index = 0
            ReDim result(arraySize)
        End If
        line = fStream.ReadLine()
        line = Replace(Replace(Replace(line, Chr(30), " "), Chr(31), " "), Chr(253), "!")
        result(Index) = line

        Index = Index + 1
    Wend
    
    fStream.Close
    Set fStream = Nothing
    Set fso = Nothing
    ReDim Preserve result(Index - 1)
    
    fromCell.Resize(Index, 1).Value2 = Application.Transpose(result)
    Set fromCell = fromCell.Offset(Index, 0)
    Range(ActiveCell, fromCell).Select
    count = count + Index
End Sub

Function makeCmdline()
    Dim astr As String
    astr = """" & "TCRD.SCS71O03.JG5P606.DG5P105A.MIS02""    ""DG5P105""       ""|"""
    astr = "EXEC 'TCRD.UG71O03.REXX(TT)' '" & astr & "'"
    Debug.Print astr
End Function

Function tsocmde(ByVal cmdline As String, Optional ByRef sess As Object)
    Set sess = CreateObject("PCOMM.autECLSession")
    'sess.setConnectionByName ("A")
    'Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLConnMgr = CreateObject("PCOMM.autECLConnMgr")
    Call autECLConnMgr.autECLConnList.Refresh
    If autECLConnMgr.autECLConnList.count <> 1 Then
        Call MsgBox("wried, count of connections is not 1..")
    End If
    Call sess.SetConnectionByHandle(autECLConnMgr.autECLConnList(1).Handle)
    Call sess.autECLOIA.WaitForInputReady
    Call sess.autECLPS.SendKeys("[HOME]CMDE[ENTER]")
    Call sess.autECLOIA.WaitForInputReady
    
    Call sess.autECLPS.SendKeys(cmdline)
    Call sess.autECLOIA.WaitForInputReady
    If (sess.autECLPS.waitForString("RAYTOOL ", 1, 2, , True, True)) Then
        sess.autECLPS.SendKeys "[ENTER]"
        
        ' getText
        dsName = "UG71O03.TEMP.OUT01"
        ActiveCell.Value = "UG71O03.TEMP.OUT01"
        Call getFileFromHost
    Else

    End If
End Function



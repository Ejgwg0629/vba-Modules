Private Const fileDir = "C:\Users\scs-tlei\Desktop\DATA\"
Private Const ForReading = 1

Private sess, autECLConnMgr, mxfer
Set sess = Nothing
Set autECLConnMgr = Nothing
Set mxfer = Nothing

run32()
call readFile()

Sub readFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = fileDir & "genmac.srl"
    Set fStream = fso.OpenTextFile(filePath, ForReading, False)
    
    Dim line
    While Not fStream.AtEndOfStream
        line = fStream.ReadLine()
        if len(replace(replace(line, " ", ""), chr(9), "")) <= 1 then
            ' line = replace(replace(line, " ", ""), chr(9), "")
            ' wscript.echo asc(line)
            Exit Sub
        end if
        dsPath = left(line, instr(line, " ")-1)
        dsn = mid(line, instr(line, "'"))
        call doDownload(dsn, dsPath)
    Wend
    
    fStream.Close
    Set fStream = Nothing
    Set fso = Nothing
End Sub


Sub doDownload(dsn, path)
    call prepareSession()
    sess.autECLPS.SendKeys "[HOME]CMDE[ENTER]"
    sess.autECLOIA.WaitForInputReady
    
    If mxfer.Ready = False Then
        Set mxfer = Nothing
        Debug.Print "weird, mxfer not ready.."
        Exit Sub
    End If
    
    Call mxfer.ReceiveFile( _
        path, _
        mangleDs(dsn), _
        "[JISCII CRLF SO NOCLEAR QUIET" _
    )
    
End Sub

Sub prepareSession()
    if Not sess is Nothing Then
        Exit Sub
    end if
    Set sess = CreateObject("PCOMM.autECLSession")
    'sess.setConnectionByName ("A")
    'Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLConnMgr = CreateObject("PCOMM.autECLConnMgr")
    Call autECLConnMgr.autECLConnList.Refresh
    If autECLConnMgr.autECLConnList.count <> 1 Then
        Call MsgBox("wried, the count of connections is not 1..")
    End If
    sess.SetConnectionByHandle (autECLConnMgr.autECLConnList(1).Handle)

    Set mxfer = sess.autECLXfer
End Sub


Function mangleDs(ds)
    If Left(ds, 1) <> "'" Then
        ds = "'" & ds
    End If
    If Right(ds, 1) <> "'" Then
        ds = ds & "'"
    End If
    mangleDs = ds
End Function


Sub Run32()
    'Author: Demon
    'Date: 2015/7/9
    'Website: http://demon.tw

    Dim strComputer, objWMIService, colItems, objItem, strSystemType
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
    
    For Each objItem in colItems
        strSystemType = objItem.SystemType
    Next
    
    If InStr(strSystemType, "x64") > 0 Then
        Dim fso, WshShell, strFullName
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set WshShell = CreateObject("WScript.Shell")
        strFullName = WScript.FullName
        If InStr(1, strFullName, "system32", 1) > 0 Then
            strFullName = Replace(strFullName, "system32", "SysWOW64", 1, 1, 1)
            WshShell.Run strFullName & " " &_
                """" & WScript.ScriptFullName & """", 10, False
            WScript.Quit
        End If
    End If
End Sub
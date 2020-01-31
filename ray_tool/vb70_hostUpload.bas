Attribute VB_Name = "vb70_hostUpload"

Sub doUpload()
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
    
    Call mxfer.SendFile( _
        "C:\Users\scs-tlei\Desktop\DATA\macgen.rexx", _
        "'TCRD.UG71O03.REXX(TT1)'", _
        "[JISCII CRLF SO NOCLEAR BLANK" _
    )
    
    Set autECLConnMgr = Nothing
    Set sess = Nothing
End Sub


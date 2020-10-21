Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
) As Long

Public Sub followLink()
    Dim olMail As Outlook.MailItem
    Dim strURL As String
    Dim lSuccess As Long

    On Error GoTo tryAnotherDoc
    Set wdDoc = Application.ActiveWindow.WordEditor

    If False Then
tryAnotherDoc:
        Err.Clear
        Set olInsp = Application.ActiveWindow().Selection(1).GetInspector
        Set wdDoc = olInsp.WordEditor
    End If

    strText = wdDoc.Application.Selection.Range.Text

    strURL = Replace(regexRepl(strText, "^[> ]+", ""), Chr(13), "")
    strURL = Mid(strURL, InStr(LCase(strURL), "notes://") + 8)
    serverName = Left(strURL, InStr(strURL, "/") - 1)
    strURL = "Notes://" & serverName & Left(Mid(strURL, InStr(strURL, "/")), 83)
    Debug.Print strURL
    If Len(strURL) = 101 Or Len(strURL) = 99 Or Len(strURL) = 70 Then
        lSuccess = ShellExecute(0, "Open", strURL)
    ElseIf Len(strURL) < 101 Then
        MsgBox ("Invalid notes link..")
    Else
        MsgBox ("It's impossible..")
    End If
End Sub

'********************************************************************************
'2nd parameter: pattern to match the string
'3rd parameter: pattern to replace
'********************************************************************************
Function regexRepl(ByVal target As String, ByVal findPat As String, _
        Optional ByVal replPat As String = "$1") As String
    Dim regexObj As New RegExp
    With regexObj
        .IgnoreCase = False
        .MultiLine = True
        .Global = True
        .Pattern = findPat
    End With
    If regexObj.Test(target) = True Then
        repledStr = regexObj.Replace(target, replPat)
    Else
        repledStr = target
    End If
    regexRepl = repledStr
End Function


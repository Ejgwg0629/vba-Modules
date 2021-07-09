Sub filterAcnMailer(Item As Outlook.MailItem)
    Dim ns As Outlook.NameSpace
    Dim MailDest As Outlook.folder
    Set ns = Application.GetNamespace("MAPI")
    Set regex = CreateObject("VBScript.RegExp")
    Debug.Print Item.SenderName
    regex.Global = True
    regex.Pattern = "(.*)"
    If regex.test(Item.Subject) Then


        'Set MailDest = ns.Folders("Personal Folders").Folders("one").Folders("a")
        'Item.Move MailDest
    End If
End Sub

Sub test()
    Dim ns As Outlook.NameSpace
    Dim MailDest As Outlook.folder
    Set ns = Application.GetNamespace("MAPI")

    Dim topFolder As Variant
    For Each eachFolder In ns.Folders

        If "\\taihang.lei@accenture.com" = eachFolder.FolderPath Then
            Set topFolder = eachFolder
        End If
    Next


    Debug.Print topFolder.FolderPath


'    Set myFolder = ns.GetDefaultFolder(29)
'    myFolder.Display
'
'    Set myItem = myFolder.Items(2)
'
'    myItem.Display
'    Debug.Print ns.CurrentUser
End Sub


Function GetFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer

    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If

    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = SubFolders.Item(FoldersArray(i))
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If

    ' Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function

GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function

Sub TestGetFolder()
    Dim folder As Outlook.folder
    Set folder = GetFolder("\\taihang.lei@accenture.com\FX\00_enterPJ")
    If Not (folder Is Nothing) Then
        folder.Display
    End If
End Sub

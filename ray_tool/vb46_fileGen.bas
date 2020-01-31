Attribute VB_Name = "vb46_fileGen"
'=============================================
'fileGen
'=============================================
Sub nameGen()
    Dim hRng As Range
    Set hRng = ActiveSheet.Range("b3")
 
    While hRng.Value <> ""
        lVal = regexRepl(hRng.Value, ".*\.(.*\..*)$")
        If lVal = "" Then
            lVal = hRng.Value
        End If
        hRng.Offset(0, 2).Value = lVal
        Set hRng = hRng.Offset(1, 0)
    Wend
    Set hRng = Nothing
End Sub
 
Sub downFileGen()
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    fN = IIf([d1].Value = "", "\down.srl", [d1.value])
    
    uDefinedP = [a1].Value
    fP = IIf(uDefinedP = "", CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Data\" _
                        , uDefinedP)
 
    If Not FSO.FolderExists(uDefinedP) Then
        If MsgBox("folder not exitst.." & Chr(13) & _
                  "make a new folder " & Chr(13) & _
                  fP, vbYesNo) = vbYes Then
            Call FSO.CreateFolder(fP)
        Else
            Exit Sub
        End If
    End If
 
    fullP = fP & fN
    
    Dim fStream As TextStream
    Set fStream = FSO.OpenTextFile(fullP, ForWriting, True)
    
    Dim hRng As Range
    Set hRng = ActiveSheet.Range("B3")
    While hRng.Value <> ""
        downStatement = fP & hRng.Offset(0, 2).Value & " ~text(" & hRng.Value & ")"
        fStream.WriteLine (downStatement)
        Set hRng = hRng.Offset(1, 0)
    Wend
    fStream.Close
    Set fStream = Nothing
    Set FSO = Nothing
End Sub


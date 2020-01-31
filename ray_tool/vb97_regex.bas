Attribute VB_Name = "vb97_regex"
'********************************************************************************
'A function that returns the first {matched string} with the given pattern
'from the Target range
'********************************************************************************
Function regexFind(ByRef target As Range, ByVal pat As String) As String
    Dim regexObj As New RegExp
    With regexObj
        .IgnoreCase = False
        .MultiLine = True
        .Global = True
        .Pattern = pat
    End With
    If regexObj.Test(target.Value) = True Then
        Set matchCol = regexObj.Execute(target.Value)
        foundStr = matchCol.Item(0).Value
    Else
        foundStr = ""
    End If
    regexFind = foundStr
End Function


'********************************************************************************
'2nd parameter: pattern to match the string
'3rd parameter: pattern to replace
'********************************************************************************
Function regexRepl(ByRef target As Range, ByVal findPat As String, _
        Optional ByVal replPat As String = "$1") As String
    Dim regexObj As New RegExp
    With regexObj
        .IgnoreCase = False
        .MultiLine = True
        .Global = True
        .Pattern = findPat
    End With
    If regexObj.Test(target.Value) = True Then
        repledStr = regexObj.replace(target.Value, replPat)
    Else
        repledStr = ""
    End If
    regexRepl = repledStr
End Function


'***********************************************
'test whether patt matches str
'***********************************************
Public Function regexMatch(ByVal STR As String, ByVal patt As String) As Boolean
    Dim REG As New RegExp
    REG.Global = True
    REG.IgnoreCase = False
    REG.MultiLine = True
    REG.Pattern = patt
    regexMatch = REG.Test(STR)
End Function


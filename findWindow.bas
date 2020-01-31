Attribute VB_Name = "findWindow"
Option Explicit

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function IsWindowVisible Lib "User32" (ByVal hWnd As Long) As Boolean
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Const GW_HWNDNEXT = 2

Private Sub Test()

    Dim lhWndP As Long
    If GetHandleFromPartialCaption(lhWndP, "Excel") = True Then
        If IsWindowVisible(lhWndP) = True Then
          MsgBox "Found VISIBLE Window Handle: " & lhWndP, vbOKOnly + vbInformation
        Else
          MsgBox "Found INVISIBLE Window Handle: " & lhWndP, vbOKOnly + vbInformation
        End If
    Else
        MsgBox "Window 'Excel' not found!", vbOKOnly + vbExclamation
    End If

End Sub

Private Function GetHandleFromPartialCaption(ByRef lWnd As Long, ByVal sCaption As String) As Boolean

    Dim lhWndP As Long
    Dim sStr As String
    GetHandleFromPartialCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        If InStr(1, sStr, sCaption) > 0 Then
            GetHandleFromPartialCaption = True
            lWnd = lhWndP
            Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop

End Function

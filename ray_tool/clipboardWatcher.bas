Attribute VB_Name = "clipboardWatcher"
Option Explicit

Private Declare PtrSafe Function AddClipboardFormatListener Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function RemoveClipboardFormatListener Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal format As Long) As Long
Private Declare PtrSafe Function MsgBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32.dll" (ByVal hWnd As LongPtr) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const WM_CLIPBOARDUPDATE As Long = &H31D
Private Const CF_BITMAP As Long = 2

Public subclassFlag As Boolean
Private myWndForm As LongPtr
Private preWndProc As Long
Private dstRange As Range
Private dstSheet As Worksheet

Public Sub subclassFormatListener()
    Set dstRange = ActiveCell
    Set dstSheet = ActiveSheet
    If dstRange Is Nothing Or dstSheet Is Nothing Then
        subclassFlag = False
        Exit Sub
    End If
    myWndForm = FindWindow("ThunderDFrame", CBEventCapture.Caption)
    preWndProc = SetWindowLong(myWndForm, GWL_WNDPROC, AddressOf myWndProc)
    Call AddClipboardFormatListener(myWndForm)
    subclassFlag = True
    Call MsgBoxTimeout(0, "clipboard watcher started..", "RayTool", vbInformation, 0, 600)
End Sub
 
Public Sub unsubclassFormatListener()
    If subclassFlag = True Then
        Call RemoveClipboardFormatListener(myWndForm)
        Call SetWindowLong(myWndForm, GWL_WNDPROC, preWndProc)
        Call MsgBoxTimeout(0, "clipboard watcher ended..", "RayTool", vbInformation, 0, 600)
        Unload CBEventCapture
        subclassFlag = False
    End If
End Sub
 
Public Function myWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If uMsg = WM_CLIPBOARDUPDATE Then
            If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
                onClipboardBitmap
            End If
    End If
    myWndProc = CallWindowProc(preWndProc, myWndForm, uMsg, wParam, lParam)
End Function
 
Public Sub onClipboardBitmap()
    Dim pasteshape As Object
    Dim cbFormats As Variant
    Dim cellCount As Double
    Dim rc As Long

    ActiveWindow.Zoom = 100
    cbFormats = Application.ClipboardFormats
    If cbFormats(1) = xlClipboardFormatBitmap Then
        dstSheet.Activate
        dstRange.Select
        dstSheet.Paste
        Set pasteshape = Application.Selection
        pasteshape.ShapeRange.LockAspectRatio = msoTrue
        'pasteshape.ShapeRange.Width = 600
        
        cellCount = pasteshape.ShapeRange.Height / Range("c1").RowHeight
        Set dstRange = dstRange.Offset(CInt(cellCount + 3), 0)
        dstRange.Select
        rc = UpdateWindow(Application.hWnd)
    End If
End Sub

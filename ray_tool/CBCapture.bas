Attribute VB_Name = "CBCapture"
' Module1
Option Explicit

  Private Declare Function AddClipboardFormatListener Lib "User32" (ByVal hWnd As Long) As Long
  Private Declare Function RemoveClipboardFormatListener Lib "User32" (ByVal hWnd As Long) As Long
  ' �N���b�v�{�[�h�ɏ����ꂽWnnProc�萔:   WM_DRAWCLIPBOARD = &H31D


Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardViewer Lib "user32.dll" (ByVal hWndNewViewer As LongPtr) As LongPtr
Private Declare PtrSafe Function ChangeClipboardChain Lib "user32.dll" (ByVal hWndRemove As LongPtr, ByVal hWndNewNext As LongPtr) As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal format As Long) As Long

Declare PtrSafe Function GetClipboardOwner Lib "User32" () As LongPtr
Declare PtrSafe Function GetClipboardViewer Lib "User32" () As LongPtr
Declare PtrSafe Function CountClipboardFormats Lib "User32" () As Long
Declare PtrSafe Function EnumClipboardFormats Lib "User32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function GetPriorityClipboardFormat Lib "User32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Declare PtrSafe Function GetOpenClipboardWindow Lib "User32" () As LongPtr

Private Const GWL_WNDPROC As Long = -4
 
'Private Const WM_DRAWCLIPBOARD As Long = &H308
Private Const WM_DRAWCLIPBOARD As Long = &H31D
Private Const WM_CHANGECBCHAIN As Long = &H30D
Private Const WM_NCHITTEST As Long = &H84
 
Private Const CF_BITMAP As Long = 2
 
Private Const ROW_HEIGHT As Double = 13.5
 
Private hWndForm As LongPtr
Private wpWindowProcOrg As Long
Private hWndNextViewer As LongPtr
Private firstFired As Boolean
 
Public Sub catchClipboard()
    hWndForm = FindWindow("ThunderDFrame", CBEvent.Caption)
    wpWindowProcOrg = SetWindowLong(hWndForm, GWL_WNDPROC, AddressOf WindowProc)
    firstFired = False
    hWndNextViewer = SetClipboardViewer(hWndForm)
End Sub
 
Public Sub releaseClipboard()
    Call ChangeClipboardChain(hWndForm, hWndNextViewer)
    Call SetWindowLong(hWndForm, GWL_WNDPROC, wpWindowProcOrg)
End Sub
 
Public Function WindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_DRAWCLIPBOARD
            If Not firstFired Then
                firstFired = True
            ElseIf IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
                pasteToSheet
            End If
            If hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = True
        Case WM_CHANGECBCHAIN
            If wParam = hWndNextViewer Then
                hWndNextViewer = lParam
            ElseIf hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_NCHITTEST
            WindowProc = 0
        Case Else
            'WindowProc = CallWindowProc(wpWindowProcOrg, hWndForm, uMsg, wParam, lParam)
    End Select
End Function
 
Public Sub pasteToSheet()
    Dim CB As Variant, x As Integer, y As Integer, iCount As Integer
    Dim wkb As String
    Dim wks As String
    Dim cb1 As New DataObject
    Dim buf As String
    Dim hei As Double
    Dim distance As Integer
    
    '�\�����摜�Ԃ̋���
    distance = 5
        
    '�������p�G���A
    buf = ""
    '�\��t���J�n�ʒu�擾
    x = ActiveCell.Column
    y = ActiveCell.Row
    '�\�t�V�[�g��
    wkb = ActiveWorkbook.name
    wks = ActiveSheet.name
    '�V�F�C�v��
    iCount = ActiveSheet.Shapes.count
    
    ActiveWindow.Zoom = 100
    
    '�N���b�v�{�[�h�̏�����
    With cb1
        .SetText buf
        .GetFromClipboard
    End With
    
        '�N���b�v�{�[�h���擾
        CB = Application.ClipboardFormats
        
    '�N���b�v�{�[�h����łȂ��ꍇ
    If CB(1) <> True Then
        '�I�u�W�F�N�g���摜�̏ꍇ
        If CB(1) = xlClipboardFormatBitmap Then
            '�\��t���ʒu�̎w��
            Windows(wkb).Activate
            Sheets(wks).Select
            ActiveSheet.Cells(y, x).Select
            '�\��t��
            ActiveSheet.Paste
            
            '�\������摜���k������
            ActiveWindow.SmallScroll Down:=12
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Height = Selection.ShapeRange.Height * 0.77
            
            '����\��t���ʒu�̎Z�o
            iCount = ActiveSheet.Shapes.count
            ActiveSheet.Shapes(iCount).Select
            hei = Selection.ShapeRange.Height
            iCount = iCount + 1
            y = y + (hei / 13.5) + distance
            
            '�N���b�v�{�[�h�̏�����
            With cb1
                .SetText buf
                .GetFromClipboard
            End With
        End If
    End If
End Sub





















  
Public Sub catchClipboard2()
    Dim result As Boolean
    hWndForm = FindWindow("ThunderDFrame", CBEvent.Caption)
    wpWindowProcOrg = SetWindowLong(hWndForm, GWL_WNDPROC, AddressOf WindowProc)
    'result = AddClipboardFormatListener(hWndForm)
    result = True
    If (Not result) Then
        Debug.Print "failed"
    End If
End Sub
Public Sub releaseClipboard2()
    Call RemoveClipboardFormatListener(hWndForm)
End Sub
Public Sub WndProc2(ByRef m As Long)
     
    If m.msg = WM_DRAWCLIPBOARD Then
      Debug.Print ("")
      If iData.GetDataPresent(DataFormats.Text) Then
        'Debug.Print(CType(iData.GetData(DataFormats.Text), String))
      End If
    End If
 
    MyBase.WndProc (m)
End Sub

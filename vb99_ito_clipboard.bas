Attribute VB_Name = "vb99_ito_clipboard"
Option Explicit

Dim Flag As Boolean

Sub �Ď��J�n()
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
    
    '�J�n���b�Z�[�W�\��
    MsgBox "�Ď����J�n���܂�"
    
    ActiveWindow.Zoom = 100
    
    '�t���O������
    Flag = True
    
    '�N���b�v�{�[�h�̏�����
    With cb1
        .SetText buf
        .PutInClipboard
        .GetFromClipboard
    End With
    
    Do While Flag
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
                    .PutInClipboard
                    .GetFromClipboard
                End With
            End If
        End If
        '���[�U�[�C�x���g����
        DoEvents
    Loop
    
    '��~���b�Z�[�W�\��
    MsgBox "�Ď����~���܂���"
End Sub

Sub �Ď���~()
    Flag = False
End Sub



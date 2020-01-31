Attribute VB_Name = "vb99_ito_clipboard"
Option Explicit

Dim Flag As Boolean

Sub 監視開始()
    Dim CB As Variant, x As Integer, y As Integer, iCount As Integer
    Dim wkb As String
    Dim wks As String
    Dim cb1 As New DataObject
    Dim buf As String
    Dim hei As Double
    Dim distance As Integer
    
    '貼りつける画像間の距離
    distance = 5
        
    '初期化用エリア
    buf = ""
    '貼り付け開始位置取得
    x = ActiveCell.Column
    y = ActiveCell.Row
    '貼付シート名
    wkb = ActiveWorkbook.name
    wks = ActiveSheet.name
    'シェイプ個数
    iCount = ActiveSheet.Shapes.count
    
    '開始メッセージ表示
    MsgBox "監視を開始します"
    
    ActiveWindow.Zoom = 100
    
    'フラグ初期化
    Flag = True
    
    'クリップボードの初期化
    With cb1
        .SetText buf
        .PutInClipboard
        .GetFromClipboard
    End With
    
    Do While Flag
        'クリップボードを取得
        CB = Application.ClipboardFormats
        
        'クリップボードが空でない場合
        If CB(1) <> True Then
            'オブジェクトが画像の場合
            If CB(1) = xlClipboardFormatBitmap Then
                '貼り付け位置の指定
                Windows(wkb).Activate
                Sheets(wks).Select
                ActiveSheet.Cells(y, x).Select
                '貼り付け
                ActiveSheet.Paste
                
                '貼りつけた画像を縮小する
                ActiveWindow.SmallScroll Down:=12
                Selection.ShapeRange.LockAspectRatio = msoTrue
                Selection.ShapeRange.Height = Selection.ShapeRange.Height * 0.77
                
                '次回貼り付け位置の算出
                iCount = ActiveSheet.Shapes.count
                ActiveSheet.Shapes(iCount).Select
                hei = Selection.ShapeRange.Height
                iCount = iCount + 1
                y = y + (hei / 13.5) + distance
                
                'クリップボードの初期化
                With cb1
                    .SetText buf
                    .PutInClipboard
                    .GetFromClipboard
                End With
            End If
        End If
        'ユーザーイベント制御
        DoEvents
    Loop
    
    '停止メッセージ表示
    MsgBox "監視を停止しました"
End Sub

Sub 監視停止()
    Flag = False
End Sub



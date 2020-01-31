VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBEventCapture 
   Caption         =   "CBEventCaption"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "CBEventCapture.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CBEventCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Load()
    MsgBox "userform loaded.."
End Sub

Private Sub UserForm_Activate()
    MsgBox "userform activate.."
End Sub

Private Sub UserForm_Initialize()
    'MsgBox "userform initialized.."
End Sub

Private Sub UserForm_Deactivate()
    'MsgBox "userform deactivated.."
End Sub

Private Sub UserForm_Terminate()
    'MsgBox "userform terminated.."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'MsgBox "userform Qeuryclosed.."
End Sub


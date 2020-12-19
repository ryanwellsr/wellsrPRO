VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPromo 
   Caption         =   "Special Message For You - wellsrPRO"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   OleObjectBlob   =   "ufPromo.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCancel_Click()
Application.Run "mCheckForMessages.ClosePromo"
Unload Me
End Sub

Private Sub cbOK_Click()
Application.Run "mCheckForMessages.LaunchPromo"
Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + Application.Height - Me.Height - 50
    Me.Left = Application.Left + Application.Width - Me.Width - 50
    HideTitleBar Me
End Sub


'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Application.Run "mCheckForMessages.ClosePromo"
'Unload Me
'End Sub

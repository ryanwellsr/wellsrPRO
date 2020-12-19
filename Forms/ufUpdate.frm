VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufUpdate 
   Caption         =   "Updating wellsrPRO..."
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "ufUpdate.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
#If IsMac = False Then
    'hide the title bar if you're working on a windows machine. Otherwise, just display it as you normally would
    Me.Height = Me.Height - 10
    mPrettyUserforms.HideTitleBar Me
#End If
End Sub

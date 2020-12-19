VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFeedback 
   Caption         =   "Send Feedback - wellsrPRO"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "ufFeedback.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbCancel_Click()
Unload Me
End Sub

Private Sub cbSubmit_Click()
On Error GoTo 100:
    cbSubmit.Caption = "Submitting..."
    Application.Run "Feedback.SubmitFeedback"
    Unload Me
    Exit Sub
Exit Sub
100:
MsgBox "Error loading the wellsrPRO Feedback Form" & vbNewLine & "Please visit wellsr.com for assistance.", vbCritical, "Error 1100"
Unload Me
End Sub

Private Sub UserForm_Initialize()
On Error GoTo 100
lblVersion = strVersion
lblExcel = Application.Name & " " & Application.Version & "." & Application.Build
lblOS = Application.OperatingSystem
'cbCancel.SetFocus
With tbFeedback
    .SetFocus
    .SelStart = 0
    .SelLength = Len(tbFeedback.Text)
End With
Exit Sub
100:
MsgBox "Error loading the wellsrPRO Feedback Form" & vbNewLine & "Please visit wellsr.com for assistance.", vbCritical, "Error 1100"
End Sub


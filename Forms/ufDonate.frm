VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufDonate 
   Caption         =   "Best Selling VBA Reference Guides"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
   OleObjectBlob   =   "ufDonate.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bDonate As Boolean

Private Sub cbNo_Click()
'Application.Run "mCheckForMessages.DonateNo"
Unload Me 'should trigger the query close macro
End Sub

Private Sub cbYes_Click()
Dim strAmount As String
bDonate = True
If obFund.Value = True Then
    strAmount = "fundamentals/"
ElseIf obIO.Value = True Then
    strAmount = "file-io/"
ElseIf obLogic.Value = True Then
    strAmount = "logic-and-loops/"
ElseIf obArrays.Value = True Then
    strAmount = "arrays/"
ElseIf obBundle.Value = True Then
    strAmount = "bundle/"
ElseIf obStrings.Value = True Then
    strAmount = "strings/"
End If
If strAmount = "" Then strAmount = "bundle/"
Application.Run "mCheckForMessages.DonateYes", strAmount
Unload Me
End Sub



Private Sub obArrays_Click()
LabelDesc.Caption = "This cheat sheet is full of premium tips and tricks for working with arrays. It features over 20 macro examples and is setup in a beautiful 2-page printable reference guide designed to help you declare, populate, sort, and filter your arrays."
cbYes.Caption = "GET CHEAT SHEET"
End Sub

Private Sub obBundle_Click()
LabelDesc.Caption = "With this best-selling Ultimate VBA Training Bundle, your VBA knowledge growth will be staggering. Combined, these five PDFs give you 180+ tips and macros covering the 100 most important topics in VBA. Save 48% with our most popular bundle."
cbYes.Caption = "GET THE BUNDLE"
End Sub

Private Sub obFund_Click()
LabelDesc.Caption = "Great for beginners, this cheat sheet includes over 30 useful VBA tips covering the 25 most important topics in VBA. It also crams in over 20 helpful VBA macros!"
cbYes.Caption = "GET CHEAT SHEET"
End Sub

Private Sub obIO_Click()
LabelDesc.Caption = "Split into three parts, this Reference Guide takes an in-depth look at the VBA Open Statement, the FileSystemObject, and Application.FileDialog boxes. It includes everything you need to know about File I/O, including multiple ways to read from files, write to files, append to files, and prepend to files."
cbYes.Caption = "GET CHEAT SHEET"
End Sub

Private Sub obLogic_Click()
LabelDesc.Caption = "Dive deep into IF THEN, FOR NEXT, DO WHILE and many more logic and loop topics with this cheat sheet featuring 30 VBA tips and over a dozen macro examples. With this cheat sheet, complex conditional statements and nested loops will be a breeze."
cbYes.Caption = "GET CHEAT SHEET"
End Sub

Private Sub obStrings_Click()
LabelDesc.Caption = "By describing every string manipulation function, including RegEx pattern recognition functions and custom ready-to-use functions, this is the largest cheat sheet we've ever created and is easily the most comprehensive reference guide devoted to VBA strings around."
cbYes.Caption = "GET CHEAT SHEET"
End Sub

Private Sub UserForm_Initialize()
HideTitleBar Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If bDonate = False Then
Application.Run "mCheckForMessages.DonateNo"
End If
bDonate = False
Unload Me
End Sub


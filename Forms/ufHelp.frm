VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufHelp 
   Caption         =   "wellsrPRO - The new best way to learn VBA"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   OleObjectBlob   =   "ufHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Private Sub LaunchWebsite(strurl As String)
On Error GoTo wellsrLaunchError
    Dim r As Long
    r = ShellExecute(0, "open", strurl, 0, 0, 1)
    If r = 5 Then 'if access denied, try this alternative
            r = ShellExecute(0, "open", "rundll32.exe", "url.dll,FileProtocolHandler " & strurl, 0, 1)
    End If
wellsrLaunchError:
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub




Private Sub Label141_Click()
Call LaunchWebsite("https://wellsr.com/vba/2016/excel/easily-extract-text-between-two-strings-with-vba/" & "?source=wellsrPROxl")
End Sub



Private Sub Label155_Click()
Call LaunchWebsite("https://wellsr.com/vba/2016/excel/replace-Nth-occurrence-of-substring-in-string-with-vba/" & "?source=wellsrPROxl")
End Sub





Private Sub Label170_Click()
Call LaunchWebsite("https://wellsr.com/vba/2016/excel/make-excel-talk-with-application-speech-speak-vba/" & "?source=wellsrPROxl")
End Sub






Private Sub Label2_Click()
Call LaunchWebsite("https://wellsr.com/vba/" & "?source=wellsrPROxl")
End Sub


Private Sub Label3_Click()
Call LaunchWebsite("https://wellsr.com/vba/" & "?source=wellsrPROxl")
End Sub

Private Sub Image1_Click()
Call LaunchWebsite("https://wellsr.com/vba/" & "?source=wellsrPROxl")
End Sub

Private Sub Label1_Click()
Call LaunchWebsite("https://wellsr.com/vba/" & "?source=wellsrPROxl")
End Sub



Private Sub lblBundle_Click()
Call LaunchWebsite("https://wellsr.com/vba/vba-cheat-sheets/bundle/" & "?source=wellsrPROxl")
End Sub

Private Sub lblCodeLibrary_Click()
Call LaunchWebsite("https://wellsr.com/vba/code-library/" & "?source=wellsrPROxl")
End Sub

Private Sub lblAddins_Click()
Call LaunchWebsite("https://wellsr.com/vba/vba-cheat-sheets/")
End Sub


Private Sub lblMoreHelp_Click()
Call LaunchWebsite("https://us13.campaign-archive.com/?u=9fa014ad47345eb4b70403d38&id=a2fc1b8ab7")
End Sub

Private Sub lbSpecialOffers_Click()
ufHelp.MultiPage1.Value = 1
End Sub

Private Sub lblAutoImport_Click()
ufHelp.MultiPage1.Value = 2
End Sub

Private Sub lblLibrary_Click()
ufHelp.MultiPage1.Value = 3
End Sub


Private Sub UserForm_Initialize()
Me.labelVersion.Caption = "Thank you for installing wellsrPRO Version " & strVersion
End Sub

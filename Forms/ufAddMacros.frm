VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAddMacros 
   Caption         =   "Create your own macro library"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   OleObjectBlob   =   "ufAddMacros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAddMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAdd_Click()
'user wants to add a macro
Dim sresult As String
Dim i As Integer
Dim iNextEntry As Integer
Dim iMaxMacros As Integer 'max number of personal macros a user can add
Dim iFreeFile As Integer
Dim strFilePath As String
Dim berror As Boolean
On Error GoTo errmessage:
'initialize variables
iMaxMacros = 100 'max number of personal macros a user can add

If Trim(Me.tbTitle) = "" Then
    Me.tbTitle.BackColor = RGB(255, 0, 0)
    berror = True
End If

If Trim(Me.tbMacro) = "" Then
    Me.tbMacro.BackColor = RGB(255, 0, 0)
    berror = True
End If

If berror = True Then
    MsgBox "Please fill in all required fields.", vbInformation, "Errors detected in fields"
    Exit Sub
Else
    Call ResetBackColor
End If
'Find out how many macros exist (or get the next available slot)
For i = 1 To iMaxMacros 'supports iMaxMacros number of personal macros
    sresult = INIfavorites.sManageSectionEntry(iniRead, "MyMacros", CStr(i), sINI_FILE)
    If Trim(sresult) = "" Or sresult = "Error" Then
        'if the entry is empty or you get an error, the entry doesn't exist
        iNextEntry = i
        Exit For
    End If
Next i

'write the new macro to a text file
strFilePath = sMACROS & "\macro" & iNextEntry & ".txt"
iFreeFile = FreeFile
Open strFilePath For Output As #iFreeFile
Print #iFreeFile, Me.tbMacro.Text
Close #iFreeFile

'write new macro to ini file
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "MyMacros", CStr(iNextEntry), sINI_FILE, Me.tbTitle)
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "File", CStr(iNextEntry), sINI_FILE, sMACROS & "\macro" & iNextEntry & ".txt")

MsgBox "The macro has been added to your library. To import macros from your library, click ""Import Macros"" in the wellsrPRO ribbon " & _
       "and select ""My Macros"" in the dropdown menu.", vbOKOnly, "Macro successfully saved"
'sresult = INIfavorites.sManageSectionEntry(iniWrite, "File", CStr(i), sINI_FILE, Me.lbFavorites.List(i, 1))
Me.rbEdit.Value = True
Me.cbSelectMacro.ListIndex = iNextEntry - 1
Exit Sub
errmessage:
MsgBox "Oh no! I was unable to save your macro, and I don't know why!" & vbNewLine & vbNewLine & _
       "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro not saved"
Err.Clear
On Error GoTo 0
End Sub

Private Sub cbCancel_Click()
Unload Me
End Sub

Private Sub cbDelete_Click()
Dim iEntry As Integer
If Me.cbSelectMacro.ListIndex <> -1 Then
    iEntry = Me.cbSelectMacro.List(Me.cbSelectMacro.ListIndex, 1) 'index number is stored in hidden second column if combobox
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "MyMacros", CStr(iEntry), sINI_FILE, " ")
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "File", CStr(iEntry), sINI_FILE, " ")
    If INIfavorites.FileExists(sMACROS & "\macro" & iEntry & ".txt") = True Then
        Kill sMACROS & "\macro" & iEntry & ".txt"
    End If
    'repopulate combo box
    Call PopulateMyMacros
End If
End Sub

Private Sub cbSelectMacro_Change()
'populate new titles and import the macro when the user selects a different macro from their library
Dim strTitle As String
Dim strFile As String
Dim iIndex As Integer, iFreeFile As Integer
Dim sLine As String
On Error GoTo ErrorEncountered:
If cbSelectMacro.ListIndex <> -1 Then
    strTitle = Me.cbSelectMacro.List(Me.cbSelectMacro.ListIndex, 0)
    iIndex = Me.cbSelectMacro.List(Me.cbSelectMacro.ListIndex, 1)
    strFile = INIfavorites.sManageSectionEntry(iniRead, "File", CStr(iIndex), sINI_FILE)
    'clear text in textboxes
    Me.tbMacro.Text = ""
    Me.tbTitle.Text = ""
    Me.tbTitle = strTitle
    
    If INIfavorites.FileExists(strFile) = False Then GoTo ErrorEncountered
    
    iFreeFile = FreeFile
    Open strFile For Input As #iFreeFile
'    Do Until EOF(iFreeFile)
'        Line Input #iFreeFile, sLine
'        Me.tbMacro.Text = Me.tbMacro.Text & sLine & vbLf
'    Loop
    Me.tbMacro.Text = Input(LOF(iFreeFile), iFreeFile)
    Close #iFreeFile
Else
    Me.tbMacro.Text = ""
    Me.tbTitle.Text = ""
End If
Exit Sub
ErrorEncountered:
MsgBox "Oh no! I was unable to find your saved macro, and I don't know why!" & vbNewLine & vbNewLine & _
       "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro not found"
Err.Clear
On Error GoTo 0
End Sub

Private Sub cbShare_Click()
'send macro to google docs
Me.cbShare.Caption = "Submitting..."
Application.Run "ShareMacros.SendMacro"
Me.cbShare.Caption = "Share with Community"
End Sub

Private Sub cbUpdate_Click()
'user wants to update an existing macro
Dim sresult As String
Dim iEntry As Integer
Dim iListIndex As Integer
Dim iFreeFile As Integer
Dim strFilePath As String
Dim berror As Boolean
On Error GoTo errmessage:
'initialize variables
iMaxMacros = 100 'max number of personal macros a user can add
If Me.cbSelectMacro.ListIndex = -1 Then
    MsgBox "Please select a macro to edit or select ""Add New Macro.""", vbInformation, "No Macro Selected"
    Exit Sub
End If
If Trim(Me.tbTitle) = "" Then
    Me.tbTitle.BackColor = RGB(255, 0, 0)
    berror = True
End If

If Trim(Me.tbMacro) = "" Then
    Me.tbMacro.BackColor = RGB(255, 0, 0)
    berror = True
End If

If berror = True Then
    MsgBox "Please fill in all required fields.", vbInformation, "Errors detected in fields"
    Exit Sub
Else
    Call ResetBackColor
End If

'determine selected listbox item
iEntry = Me.cbSelectMacro.List(Me.cbSelectMacro.ListIndex, 1) 'index number is stored in hidden second column if combobox
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "MyMacros", CStr(iEntry), sINI_FILE, Me.tbTitle)
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "File", CStr(iEntry), sINI_FILE, sMACROS & "\macro" & iEntry & ".txt")

'write the revised macro to a text file
strFilePath = sMACROS & "\macro" & iEntry & ".txt"
iFreeFile = FreeFile
Open strFilePath For Output As #iFreeFile
Print #iFreeFile, Me.tbMacro.Text
Close #iFreeFile

'repopulate combobox
iListIndex = Me.cbSelectMacro.ListIndex
Call PopulateMyMacros
Me.cbSelectMacro.ListIndex = iListIndex

'success
MsgBox "The macro in your library has been updated. To import macros from your library, click ""Import Macros"" in the wellsrPRO ribbon " & _
       "and select ""My Macros"" in the dropdown menu.", vbOKOnly, "Macro successfully saved"
Exit Sub
errmessage:
'failure
MsgBox "Oh no! I was unable to save your macro, and I don't know why!" & vbNewLine & vbNewLine & _
       "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro not saved"
Err.Clear
On Error GoTo 0
End Sub



Private Sub rbAdd_Change()
If ufAddMacros.rbAdd.Value = True Then
    With ufAddMacros
        .lblTitle = "Enter a title for your macro:"
        .lblMacro = "Paste your macro in the box below"
        If InStr(UCase(.Caption), "SHARE") = 0 Then
            .cbUpdate.Visible = False
            .cbDelete.Visible = False
            .cbAdd.Visible = True
        End If
        .lblSelectMacro.Visible = False
        .cbSelectMacro.Clear
        .cbSelectMacro.Visible = False
        .tbTitle.Text = ""
        .tbMacro.Text = ""
    End With
    Call ResetBackColor
End If
End Sub

Private Sub rbEdit_Change()
Dim sresult As String
Dim iMaxMacros As Integer 'max number of personal macros a user can add
iMaxMacros = 100 'max number of personal macros a user can add
With ufAddMacros
    If .rbEdit.Value = True Then
        If InStr(UCase(.Caption), "SHARE") = 0 Then
            .lblTitle = "Edit the title of your macro:"
            .lblMacro = "Edit your macro in the box below"
            .cbUpdate.Visible = True
            .cbDelete.Visible = True
            .cbAdd.Visible = False
        Else
            .lblTitle = "Enter the title of your macro:"
            .lblMacro = "You'll share the macro in the box below"
        End If
        .lblSelectMacro.Visible = True
        .cbSelectMacro.Visible = True
        .cbSelectMacro.ColumnWidths = "200;0"
    
        Call ResetBackColor
        'populate combobox
        Call PopulateMyMacros
    End If
End With
End Sub

Private Sub ResetBackColor()
    ufAddMacros.tbMacro.BackColor = RGB(255, 255, 255)
    ufAddMacros.tbTitle.BackColor = RGB(255, 255, 255)
End Sub

Private Sub PopulateMyMacros()
Dim sresult As String
Dim iMaxMacros As Integer
Dim i As Integer
iMaxMacros = 100
        With ufAddMacros.cbSelectMacro
            .Clear
            For i = 1 To iMaxMacros
                sresult = INIfavorites.sManageSectionEntry(iniRead, "MyMacros", CStr(i), sINI_FILE)
                If Trim(sresult) = "" Or sresult = "Error" Then
                    'do nothing
                Else
                    .AddItem
                    .List(.ListCount - 1, 0) = sresult
                    .List(.ListCount - 1, 1) = i
                End If
            Next i
            '.ListIndex = 0
        End With
End Sub


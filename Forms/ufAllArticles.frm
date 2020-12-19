VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAllArticles 
   Caption         =   "Automatically Import Macros from wellsr.com - wellsrPRO"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   OleObjectBlob   =   "ufAllArticles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAllArticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbGroups_Change()
On Error GoTo ErrorEncountered:
Dim i As Integer, lcount As Integer
For i = Me.lbArticles.ListCount - 1 To 0 Step -1
    'delete all articles
    Me.lbArticles.RemoveItem i
Next i
ufAllArticles.bLaunch.Enabled = True
If InStr(Me.cbGroups.List(Me.cbGroups.ListIndex), "Recent") <> 0 Then
    'populate listbox with list of most recent articles
    With Me.lbArticles
        For i = LBound(strArticles, 2) To UBound(strArticles, 2)
            If strArticles(1, i) <> "" Then
                .AddItem
                .List(i, 0) = strArticles(1, i)
                .List(i, 1) = strArticles(3, i)
            End If
        Next i
    End With
ElseIf InStr(Me.cbGroups.List(Me.cbGroups.ListIndex), "Favorites") <> 0 Then
    'populate with list of favorites
    If INIfavorites.FileExists(sINI_FILE) = True Then
        'populate array of favorites
        Application.Run "inifavorites.populatefavorites"
    End If
ElseIf InStr(Me.cbGroups.List(Me.cbGroups.ListIndex), "My Macros") <> 0 Then
    'populate with personal macro library list
    If INIfavorites.FileExists(sINI_FILE) = True Then
        'populate array of favorites
        Application.Run "INIPersonalLibrary.PopulateMyMacros_IMPORT"
    End If
    ufAllArticles.bLaunch.Enabled = False
Else
    'populate listbox with the appropriate category item
    With Me.lbArticles
        iCount = 0
        For i = LBound(strTutorials, 2) To UBound(strTutorials, 2)
            If Me.cbGroups.List(Me.cbGroups.ListIndex) = strTutorials(2, i) Then
                .AddItem
                .List(iCount, 0) = strTutorials(1, i)
                .List(iCount, 1) = strTutorials(3, i)
                iCount = iCount + 1
            End If
        Next i
    End With
End If
Me.lbArticles.ListIndex = 0
Exit Sub
ErrorEncountered:
MsgBox "Error encountered while trying to enumerate article list.", vbCritical, "Error Encountered"
End Sub

Private Sub bImport_Click()
If Me.rbNew = True Then
    'import to new module
    Application.Run "parsehtml.ImportMacroExamples", Me.lbArticles.List(Me.lbArticles.ListIndex, 1)
Else
    'import to existing module
    Application.Run "parsehtml.ImportMacroExamples", Me.lbArticles.List(Me.lbArticles.ListIndex, 1), Me.cbModules.List(Me.cbModules.ListIndex, 0)
End If
End Sub

Private Sub bCancel_Click()
Unload Me
End Sub

Private Sub bLaunch_Click()
Application.Run "myxml.XMLLaunchWebsite", Me.lbArticles.List(Me.lbArticles.ListIndex, 1) & "?source=wellsrPROxl"
End Sub








Private Sub rbExisting_Change()
cbModules.Enabled = True
lbSelectModules.Enabled = True
'populate the list of modules in the active workbook
Application.Run "parsehtml.CheckTrustAccess", "dummy"
If ParseHTML.VBAIsTrusted Then
    Application.Run "parsehtml.ListModules"
End If
End Sub

Private Sub rbNew_Change()
'delete all existing entries
    For i = Me.cbModules.ListCount - 1 To 0 Step -1
        'delete all articles
        Me.cbModules.RemoveItem i
    Next i
Me.cbModules.Value = ""
'disable combo box
cbModules.Enabled = False
lbSelectModules.Enabled = False
End Sub


Private Sub UserForm_Initialize()
Dim i As Integer

'populate items in listbox
If INIfavorites.FileExists(sINI_FILE) = True Then
    'populate array of favorites
    With Me.cbGroups
        .AddItem
        .List(.ListCount - 1) = "Favorites"
        Application.Run "inifavorites.populatefavorites"
        .AddItem
        .List(.ListCount - 1) = "My Macros"
    End With
End If

With Me.cbGroups
    .AddItem
    .List(.ListCount - 1) = "Most Recent Tutorials"
    For i = LBound(strGroups) To UBound(strGroups)
        .AddItem
        .List(.ListCount - 1) = strGroups(i)
    Next i
    .ListIndex = 0
End With

Me.lbArticles.ColumnWidths = "150;0"
'If InStr(Me.cbGroups.List(Me.cbGroups.ListIndex), "Recent") <> 0 Then
'    'populate listbox with list of most recent articles
'    With Me.lbArticles
'        For i = LBound(strArticles, 2) To UBound(strArticles, 2)
'            If strArticles(1, i) <> "" Then
'                .AddItem
'                .List(i, 0) = strArticles(1, i)
'                .List(i, 1) = strArticles(3, i)
'            End If
'        Next i
'    End With
'Else
'    'populate listbox with the appropriate category item
'End If
cbModules.Value = ""
cbModules.Enabled = False
lbSelectModules.Enabled = False
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFavorites 
   Caption         =   "Manage Favorites - wellsrPRO"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12465
   OleObjectBlob   =   "ufFavorites.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bSave_Click()
'Write the favorited items to the ini file
Dim sresult As String
Dim i As Integer

'overwrite all previously stored values as empty strings
For i = 0 To 50
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "Articles", CStr(i), sINI_FILE, " ")
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "URLs", CStr(i), sINI_FILE, " ")
Next i

For i = 0 To Me.lbFavorites.ListCount - 1
    'Write each value to the ini file
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "Articles", CStr(i), sINI_FILE, Me.lbFavorites.List(i, 0))
    sresult = INIfavorites.sManageSectionEntry(iniWrite, "URLs", CStr(i), sINI_FILE, Me.lbFavorites.List(i, 1))
Next i
MsgBox "Favorites successfully saved", vbInformation, "Favorites Saved"
Unload Me
Exit Sub
exitearly:
MsgBox "Error encountered while saving favorites", vbInformation, "Error Saving Favorites"
End Sub

Private Sub BTN_MoveSelectedLeft_Click()
'    Dim iCtr As Long

'add it to the left list
'    For iCtr = 0 To Me.lbFavorites.ListCount - 1
'        If Me.lbFavorites.Selected(iCtr) = True Then
'            Me.lbArticles.AddItem Me.lbFavorites.List(iCtr)
'        End If
'    Next iCtr

'delete it from the right list
'    For iCtr = Me.lbFavorites.ListCount - 1 To 0 Step -1
'        If Me.lbFavorites.Selected(iCtr) = True Then
'            Me.lbFavorites.RemoveItem iCtr
'        End If
'    Next iCtr
If Me.lbFavorites.ListIndex > -1 Then
    Me.lbFavorites.RemoveItem Me.lbFavorites.ListIndex
End If
End Sub

Private Sub BTN_MoveSelectedRight_Click()
    Dim iCtr As Long, i As Long

'add it to the right list
    For iCtr = 0 To Me.lbArticles.ListCount - 1
        If Me.lbArticles.Selected(iCtr) = True Then
            With Me.lbFavorites
                .AddItem
                .List(.ListCount - 1, 0) = Me.lbArticles.List(iCtr, 0)
                .List(.ListCount - 1, 1) = Me.lbArticles.List(iCtr, 1)
            End With

        End If
    Next iCtr

'delete it from the left list
'    For iCtr = Me.lbArticles.ListCount - 1 To 0 Step -1
'        If Me.lbArticles.Selected(iCtr) = True Then
'            Me.lbArticles.RemoveItem iCtr
'        End If
'    Next iCtr
End Sub


Private Sub cmdDown_Click()

MoveItem 1

End Sub

Private Sub cmdUp_Click()

MoveItem -1

End Sub

Private Sub MoveItem(lOffset As Long)
Dim aTemp() As String
Dim i As Long
On Error GoTo exitearly:
With Me.lbFavorites
    If .ListIndex > -1 Then
        ReDim aTemp(0 To .ColumnCount - 1)
        For i = 0 To .ColumnCount - 1
            aTemp(i) = .List(.ListIndex + lOffset, i)
            .List(.ListIndex + lOffset, i) = .List(.ListIndex, i)
            .List(.ListIndex, i) = aTemp(i)
        Next i
    End If
    .ListIndex = .ListIndex + lOffset
End With
exitearly:
Err.Clear
On Error GoTo 0
End Sub

Private Sub cbCancel_Click()
Unload Me
End Sub


Private Sub UserForm_Initialize()
Dim i As Integer
'
Application.Run "parsehtml.ListAllArticles", , True
'populate items in listbox
'If INIfavorites.FileExists(sINI_FILE) = True Then
'    'populate array of favorites
'    With Me.cbGroups
'        .AddItem
'        .List(.ListCount - 1) = "Favorites"
'        Application.Run "inifavorites.populatefavorites"
'    End With
'End If

'populate list of tutorials
With Me.cbGroups
    .AddItem
    .List(.ListCount - 1) = "Most Recent Tutorials"
    For i = LBound(strGroups) To UBound(strGroups)
        .AddItem
        .List(.ListCount - 1) = strGroups(i)
    Next i
    .ListIndex = 0
End With

'populate favorites
    If INIfavorites.FileExists(sINI_FILE) = True Then
        'populate array of favorites
        Application.Run "inifavorites.populatefavorites2"
    End If
Me.lbArticles.ColumnWidths = "150;0"
Me.lbFavorites.ColumnWidths = "150;0"
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
End Sub



Private Sub cbGroups_Change()
On Error GoTo ErrorEncountered:
Dim i As Integer, lcount As Integer
For i = Me.lbArticles.ListCount - 1 To 0 Step -1
    'delete all articles
    Me.lbArticles.RemoveItem i
Next i
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
'ElseIf InStr(Me.cbGroups.List(Me.cbGroups.ListIndex), "Favorites") <> 0 Then
'    'populate with list of favorites
'    If INIfavorites.FileExists(sINI_FILE) = True Then
'        'populate array of favorites
'        Application.Run "inifavorites.populatefavorites"
'    End If
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


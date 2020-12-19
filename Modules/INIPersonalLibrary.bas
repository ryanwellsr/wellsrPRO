Attribute VB_Name = "INIPersonalLibrary"
Private Sub PopulateMyMacros_IMPORT()
Dim sReturn As String
Dim sReturn2 As String
Dim i As Integer
Dim iCount As Integer
Dim iMaxMacros As Integer
iMaxMacros = 100

    iCount = -1
    ' Read the ini file - can store up to 50 favorites
    For i = 1 To iMaxMacros
        sReturn = INIfavorites.sManageSectionEntry(iniRead, "MyMacros", CStr(i), sINI_FILE)
        sReturn2 = INIfavorites.sManageSectionEntry(iniRead, "File", CStr(i), sINI_FILE)
            If sReturn <> "Error" And sReturn2 <> "Error" Then
                If Trim(sReturn) <> "" And Trim(sReturn2) <> "" Then
                    iCount = iCount + 1
                    With ufAllArticles.lbArticles
                        .AddItem
                        .List(.ListCount - 1, 0) = sReturn
                        .List(.ListCount - 1, 1) = sReturn2
                    End With
                End If
            End If
    Next i
    If iCount < 0 Then
        With ufAllArticles.lbArticles
            .AddItem
            .List(.ListCount - 1, 0) = "Your library is empty. Add macros with ""Manage My Macros"""
            .List(.ListCount - 1, 1) = "NONE"
        End With
    End If

End Sub


Private Sub DisplayShareMacrosUF()

'hide buttons and display modified userform
With ufAddMacros
    .FrameSelect.Caption = "What do you want to share?"
    .rbAdd.Caption = "Share New Macro"
    .rbEdit.Caption = "Share Existing Macro"
    .lblSelectMacro.Caption = "Which macro would you like to share?"
    .cbAdd.Visible = False
    .cbDelete.Visible = False
    .cbUpdate.Visible = False
    .Caption = "Share your macros with the wellsrPRO community"
    .lblMain = "Share your macros with the wellsrPRO community"
    .cbShare.Top = 468
    .cbShare.Left = 364
    .lblShare.Top = 468
    .lblShare.Left = 20
End With
    Load ufAddMacros
    ufAddMacros.Show
End Sub

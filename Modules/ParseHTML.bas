Attribute VB_Name = "ParseHTML"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Option Private Module
Public strTutorials() As String 'variable to store all the Excel tutorials on website
Public strGroups() As String
Private Sub ImportMacroExamples(strurl As String, Optional strmodule As String)
'Macro to import example macros from website and add them to a new module
'Add Reference to (1) MICROSOFT HTML OBJECT LIBRARY
'                 (3) MICROSOFT VISUAL BASIC FOR APPLICATIONS EXTENSIBILITY 5.3
'Change this so the user can search through all my VBA articles and import the examples
'  -A feature to list them all and to search through all of them would be great features.
'     -Can probably list them in categories as they appear on my /vba/excel/ page.
    Application.ScreenUpdating = False
CheckTrustAccess (strmodule)
If VBAIsTrusted Then
    If InStr(ufAllArticles.cbGroups.List(ufAllArticles.cbGroups.ListIndex), "My Macros") = 0 Then
        'import macros from wellsr.com
        If strmodule = "" Then
            ExtractCode (strurl)
        Else
            ExtractCode strurl, strmodule
        End If
    Else
        'import macros from user's personal macro library
        If strmodule = "" Then
            GetMacroFromLibrary (strurl)
        Else
            GetMacroFromLibrary strurl, strmodule
        End If
    End If
End If
    Application.ScreenUpdating = True
End Sub
 
Private Sub CheckTrustAccess(Optional strmodule As String)
Dim strStatus, strOpp, strCheck As String
Dim bEnabled As Boolean
If Not VBAIsTrusted Then
'ask the user if they want me to try to programatically toggle trust access. If I fail, give them directions.
    bEnabled = False
    strStatus = "DISABLE"
    strOpp = "ENABLE"
    v = MsgBox("Trust Access to the VBA Project Object Model is " & strStatus & "D." & Chr(10) & Chr(10) & _
     "Would you like me to attempt to " & strOpp & " it?", vbYesNo, strOpp & " Trust Access?")
Else
    Exit Sub 'because already enabled
End If
 
    If v = 6 Then
        'try to toggle trust
        Call ToggleTrust(bEnabled, strmodule)
        If VBAIsTrusted = bEnabled Then
            'if ToggleTrust fails to programatically toggle trust
            MsgBox "I failed to " & strOpp & " Trust Access." & Chr(10) & Chr(10) & _
            "To " & strOpp & " this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Check the box next to ""Trust Access to the VBA project object model""" & vbNewLine & vbNewLine & _
            "THIS STEP MUST BE COMPLETE IN ORDER TO AUTOMATICALLY IMPORT VBA TUTORIAL EXAMPLES", vbOKOnly, "Auto " & strOpp & " Failed"
            End
        Else
            MsgBox "I successfully " & strOpp & "D Trust Access." & Chr(10) & Chr(10) & _
            "To " & strStatus & " this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Toggle the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "Auto " & strOpp & " Succeeded"
            On Error Resume Next
            wellsrCustomRibbon.InvalidateControl ("tTrust")
            Err.Clear
            On Error GoTo 0
        End If
    Else
        MsgBox "I cannot automatically import the VBA tutorial examples until Trust Access is granted. " & vbNewLine & _
        "To manually " & strOpp & " Trust Access:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Toggle the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "How to " & strOpp & " Trust Access"
   End If
End Sub
 
Function VBAIsTrusted() As Boolean
Dim a1 As Integer
On Error GoTo Label1
a1 = ActiveWorkbook.VBProject.VBComponents.count
VBAIsTrusted = True
Exit Function
Label1:
VBAIsTrusted = False
End Function
 
Private Sub ToggleTrust(bEnabled As Boolean, Optional strmodule As String)
Dim b1 As Integer, i As Integer
Dim strkeys As String, bFormOpen As Boolean
Application.ScreenUpdating = True
'hide all open userforms
bformloaded = IsFormLoaded("ufAllArticles")
If bformloaded Then
    ufAllArticles.Hide
End If
On Error Resume Next
    Do While i <= 2 'try to sendkeys 3 times
        Sleep 100
    strkeys = "%tms%v{ENTER}"
        Call SendKeys(Trim(strkeys)) 'ST%V{ENTER}")
        DoEvents
        If VBAIsTrusted <> bEnabled Then Exit Do 'successfully toggled trust
        Sleep (100)
        i = i + 1
    Loop
    
'show the userforms that were previously opened
If bformloaded Then
    If VBAIsTrusted Then
        If strmodule <> "" Then
            ListModules
            ufAllArticles.Show
        End If
    End If
End If
Application.ScreenUpdating = False
End Sub
 
Public Sub ButtonToggleTrust(bEnabled As Boolean)
Dim i As Integer
Application.ScreenUpdating = True
On Error Resume Next
    Do While i <= 2 'try to sendkeys 3 times
        Sleep 100
    strkeys = "%tms%v{ENTER}"
        Call SendKeys(Trim(strkeys)) 'ST%V{ENTER}")
        DoEvents
        If VBAIsTrusted = bEnabled Then Exit Do 'successfully toggled trust
        Sleep (100)
        i = i + 1
    Loop
End Sub


Private Sub ExtractCode(postURL As String, Optional strmodule As String)
' DEVELOPER:    Ryan Wells (wellsr.com)
' DESCRIPTION:  Extract all the VBA code from an article at the specified URL on my site
'               -Import the macro examples to a new module
' DATE:         08/2017
' DETAILS:      Added in Version 1.0.0 of wellsrPRO add-in
'               Requires Reference to MICROSOFT HTML OBJECT LIBRARY
On Error GoTo FailedToLoad
    Dim xmlHTTPReq As Object 'New MSXML2.XMLHTTP
    Dim htmlDOC As New HTMLDocument
    Dim htmlDOC2 As New HTMLDocument
    Dim classtag As MSHTML.IHTMLElementCollection 'Find code snippets based on class
    Dim classtag2 As MSHTML.IHTMLElementCollection
    Dim Snippet As MSHTML.IHTMLElement 'For each code snippet
    Dim i As Long, iCount As Long
    Dim bNewModule As Boolean, berror As Boolean
    Dim strTemp As String
    Dim iLen As Integer
    Dim Now30 As Date
    Dim v As Variant
   
    Dim proj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim codemod As VBIDE.CodeModule
    Application.ScreenUpdating = False

    Set xmlHTTPReq = CreateObject("MSXML2.ServerXMLHTTP")
    
    With xmlHTTPReq
        .Open "POST", postURL, False
        .send
    End With
    Now30 = DateAdd("s", 30, Now())
    Do While xmlHTTPReq.readyState <> 4
        'Try for 30 seconds
        DoEvents
        If Now >= Now30 Then
            berror = True
            GoTo FailedToLoad
        End If
    Loop
    If xmlHTTPReq.Status <> 200 Then
        GoTo FailedToLoad:
    End If
    'if xmlhttpreq.status
    With htmlDOC
        .body.innerHTML = xmlHTTPReq.responseText
    End With
    
    With htmlDOC
        For iCount = 0 To .getElementsByClassName("highlight").Length - 1 'for each code block in html
            With .getElementsByClassName("highlight")(iCount)
                htmlDOC2.body.innerHTML = .outerHTML 'grab html from the identified class
                If CBool(htmlDOC2.getElementsByClassName("language-vb").Length) Then 'confirm it's a VB code
                    'confirmed it's a vb tag so add the code
                    If bNewModule = False Then
                        'create new module the first time
                        Set proj = ActiveWorkbook.VBProject
                        If strmodule = "" Then
                            Set comp = proj.VBComponents.Add(vbext_ct_StdModule)
                        Else
                            Set comp = proj.VBComponents(strmodule)
                        End If
                        Set codemod = comp.CodeModule
                        With codemod
                            strTemp = "'#              " & postURL
                            iLen = Len(strTemp)
                            If iLen < Len("'#              Sample Macros Automatically Imported From:") Then
                                iLen = Len("'#              Sample Macros Automatically Imported From:")
                            End If
                            .InsertLines .CountOfLines + 1, "'" & Replace(Space(iLen + 1), " ", "#")
                            .InsertLines .CountOfLines + 1, "'#   PRODUCT:   wellsrPRO " & strVersion & Space(iLen - Len("'#   PRODUCT:   wellsrPRO " & strVersion)) & " #"
                            .InsertLines .CountOfLines + 1, "'#   DEVELOPER: Ryan Wells (wellsr.com)" & Space(iLen - Len("'#   DEVELOPER: Ryan Wells (wellsr.com)")) & " #"
                            .InsertLines .CountOfLines + 1, "'#   DETAILS:   Imported on " & Format(Now(), "mm/dd/yyyy") & Space(iLen - Len("'#   DETAILS:   Imported on xx/xx/xxxx")) & " #"
                            .InsertLines .CountOfLines + 1, "'#   SOURCE:    Sample Macros Automatically Imported From:" & Space(iLen - 58) & " #"
                            
                            .InsertLines .CountOfLines + 1, strTemp & " " & Space(iLen - Len(strTemp)) & "#"
                            If iLen > Len("'#   NOTES:     Imported macros may need to be moved, renamed or deleted") Then
                                .InsertLines .CountOfLines + 1, "'#   NOTES:     Imported macros may need to be moved, renamed or deleted" & Space(iLen - Len("'#   NOTES:     Imported macros may need to be moved, renamed or deleted")) & " #"
                            Else
                                .InsertLines .CountOfLines + 1, "'#   NOTES:     Imported macros may need to be moved, renamed or deleted #"
                            End If
                            .InsertLines .CountOfLines + 1, "'#              in order to function properly with your project." & Space(iLen - Len("'#              in order to function properly with your project.")) & " #"
                            .InsertLines .CountOfLines + 1, "'" & Replace(Space(iLen + 1), " ", "#")
                            .InsertLines .CountOfLines + 1, ""
                        End With
                    End If
                    With codemod 'add macro to the new module
                        .InsertLines .CountOfLines + 1, htmlDOC.getElementsByClassName("highlight")(iCount).innertext
                        .InsertLines .CountOfLines + 1, "'-----------------------------------------------------------"
                        .InsertLines .CountOfLines + 1, Chr(10)
                        bNewModule = True
                    End With
                End If
                Set htmlDOC2 = Nothing
            End With
        Next iCount
    End With

FailedToLoad:
    Set xmlHTTPReq = Nothing
    Set htmlDOC = Nothing
    If bNewModule = True Then
        bNewModule = False
        v = MsgBox("Sample macros were successfully loaded to the module " & codemod.Name & "." & vbNewLine & vbNewLine & _
                  "Would you like me to automatically launch the Visual Basic Editor?", vbYesNo, "Sample Macros Loaded")
        If v = 6 Then
            If IsFormLoaded("ufAllArticles") Then
                Unload ufAllArticles
            End If
            Application.VBE.MainWindow.Visible = True
            Application.VBE.MainWindow.WindowState = vbext_ws_Maximize
        Else
            'Application.VBE.MainWindow.Visible = False
            If Application.VBE.MainWindow.Visible = True Then
                Application.VBE.MainWindow.WindowState = vbext_ws_Minimize
            End If
        End If
    Else
        If Err.Number = 0 Then
            If berror = False Then
                MsgBox "No exportable macros were detected in the tutorial. No macros were imported.", , "No Sample Macros Detected"
            Else
                MsgBox "Timed out - No response when trying to access website.", , "Timed out"
            End If
        Else
            MsgBox "An error occurred while trying to import sample macros." & vbNewLine & vbNewLine & _
                    "Error " & Err.Number & ": " & Err.Description, , "Failed to Load Sample Macros"
            Err.Clear
        End If
    End If
    Application.ScreenUpdating = True
On Error GoTo 0
End Sub

Private Sub GetMacroFromLibrary(FilePath As String, Optional strmodule As String)
'copy macros from the user's personal macro library to a module
Dim proj As VBIDE.VBProject
Dim comp As VBIDE.VBComponent
Dim codemod As VBIDE.CodeModule
Dim iFreeFile As Integer
Dim i As Integer
Dim v As Variant
On Error GoTo FailedToLoadFromLibrary:
Application.ScreenUpdating = False

If INIfavorites.FileExists(FilePath) And Trim(FilePath) <> "" Then
    'create new module the first time
    Set proj = ActiveWorkbook.VBProject
    If strmodule = "" Then
        Set comp = proj.VBComponents.Add(vbext_ct_StdModule)
    Else
        Set comp = proj.VBComponents(strmodule)
    End If
    Set codemod = comp.CodeModule
    With codemod
        'print header info
        strTemp = "'#############################################################"
        iLen = Len(strTemp)
        .InsertLines .CountOfLines + 1, strTemp
        .InsertLines .CountOfLines + 1, "'#        Personal Macro library managed by wellsrPRO        #"
        .InsertLines .CountOfLines + 1, "'#   PRODUCT:   wellsrPRO " & strVersion & Space(iLen - Len("'#   PRODUCT:   wellsrPRO " & strVersion) - 1) & "#"
        .InsertLines .CountOfLines + 1, "'#   DEVELOPER: Ryan Wells (wellsr.com)" & Space(iLen - Len("'#   DEVELOPER: Ryan Wells (wellsr.com)") - 1) & "#"
        .InsertLines .CountOfLines + 1, "'#   DETAILS:   Imported on " & Format(Now(), "mm/dd/yyyy") & Space(iLen - Len("'#   DETAILS:   Imported on xx/xx/xxxx") - 1) & "#"
        .InsertLines .CountOfLines + 1, strTemp
        .InsertLines .CountOfLines + 1, ""
    End With
    With codemod 'add macro to module
        iFreeFile = FreeFile
            Open FilePath For Input As #iFreeFile
            'Do Until EOF(iFreeFile)
            '    Line Input #iFreeFile, sLine
            '    .InsertLines .CountOfLines + 1, sLine
            'Loop
            .InsertLines .CountOfLines + 1, Input(LOF(iFreeFile), iFreeFile)
            Close #iFreeFile
    End With
ElseIf FilePath = "NONE" Then
    Exit Sub
Else
    MsgBox "Oh no! I was unable to find your saved macro, and I don't know why!" & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro not found"
End If
Application.ScreenUpdating = True

'success
    v = MsgBox("Macros were successfully loaded to the module " & codemod.Name & "." & vbNewLine & vbNewLine & _
              "Would you like me to automatically launch the Visual Basic Editor?", vbYesNo, "Library Macros Loaded")
    If v = 6 Then
        If IsFormLoaded("ufAllArticles") Then
            Unload ufAllArticles
        End If
        Application.VBE.MainWindow.Visible = True
        Application.VBE.MainWindow.WindowState = vbext_ws_Maximize
    Else
        'Application.VBE.MainWindow.Visible = False
        If Application.VBE.MainWindow.Visible = True Then
            Application.VBE.MainWindow.WindowState = vbext_ws_Minimize
        End If
    End If

Exit Sub
FailedToLoadFromLibrary:
'failure
    MsgBox "An error occurred while trying to import your personal macros." & vbNewLine & vbNewLine & _
            "Error " & Err.Number & ": " & Err.Description, , "Failed to Load Library Macros"
    Err.Clear

Application.ScreenUpdating = True
On Error GoTo 0
End Sub


Private Sub ListAllArticles(Optional NoMacros As Boolean, Optional NoDisplay As Boolean)
' DEVELOPER:    Ryan Wells (wellsr.com)
' DESCRIPTION:  Return the article titles and article URLs of all the articles on my website
'               -to populate a list to later import the tutorial macros to a module
' DATE:         08/2017
' DETAILS:      Added in Version 3.0.0 of wellsrTools add-in
'               Requires Reference to MICROSOFT HTML OBJECT LIBRARY
On Error GoTo FailedToLoad
    Dim xmlHTTPReq As Object 'New MSXML2.XMLHTTP
    Dim htmlDOC As New HTMLDocument
    Dim htmlDOC2 As New HTMLDocument
    Dim classtag As MSHTML.IHTMLElementCollection 'Find code snippets based on class
    Dim classtag2 As MSHTML.IHTMLElementCollection
    Dim Snippet As MSHTML.IHTMLElement 'For each code snippet
    Dim i As Long, iCount As Long, iRow As Long, jcount As Long
    Dim bNewModule As Boolean, berror As Boolean
    Dim Now30 As Date
    Application.ScreenUpdating = False
   
    Dim proj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim codemod As VBIDE.CodeModule
    Dim strGroup As String
    Dim strurl As String, strTitle As String
    
    postURL = "https://wellsr.com/vba/excel/" 'website that lists all my Excel articles
    ReDim strTutorials(1 To 3, 0)
    ReDim strGroups(0)
    
    Set xmlHTTPReq = CreateObject("MSXML2.ServerXMLHTTP")
    
    With xmlHTTPReq
        .Open "POST", postURL, False
        .send
    End With
    Now30 = DateAdd("s", 30, Now())
    Do While xmlHTTPReq.readyState <> 4
        'Try for 30 seconds
        DoEvents
        If Now >= Now30 Then
            berror = True
            GoTo FailedToLoad
        End If
    Loop
    If xmlHTTPReq.Status <> 200 Then
        GoTo FailedToLoad:
    End If
    'if xmlhttpreq.status
    With htmlDOC
        .body.innerHTML = xmlHTTPReq.responseText
    End With
    
    With htmlDOC
        For iCount = 0 To .getElementsByClassName("sectiongroup").Length - 1
            With .getElementsByClassName("sectiongroup")(iCount)
                htmlDOC2.body.innerHTML = .outerHTML 'grab html from the identified class
                If CBool(htmlDOC2.getElementsByClassName("sectiontitle").Length) Then
                    'confirmed it's a vb tag so add the code
                    strGroup = htmlDOC2.getElementsByClassName("sectiontitle")(0).innertext
                    ReDim Preserve strGroups(0 To jcount)
                    strGroups(jcount) = strGroup
                    jcount = jcount + 1
                End If
                For i = 0 To htmlDOC2.getElementsByClassName("sectionlink").Length - 1
                    With htmlDOC2.getElementsByClassName("sectionlink")(i)
                        'grab link text
                        strTitle = .innertext
                        'grab link url
                        strurl = UDFs.SuperMid(.outerHTML, "href=""", """>")
                        If Left(strurl, 1) = "/" Then strurl = "https://wellsr.com" & strurl
                        strurl = Replace(strurl, "about:", "https://wellsr.com")
                    End With
                    'With ufAllArticles.lbArticles
                    '    .AddItem
                    '    .List(iRow, 0) = strTitle
                    '    .List(iRow, 1) = "(" & strGroup & ")"
                    '    .List(iRow, 2) = strUrl
                    'End With
                    ReDim Preserve strTutorials(1 To 3, 0 To iRow)
                    strTutorials(1, iRow) = strTitle
                    strTutorials(2, iRow) = strGroup
                    strTutorials(3, iRow) = strurl
                    iRow = iRow + 1
                Next i
                Set htmlDOC2 = Nothing
            End With
        Next iCount
    End With

    Set xmlHTTPReq = Nothing
    Set htmlDOC = Nothing

If NoDisplay = False Then
    If NoMacros = True Then
        'shorten the form and move the close button
        ufAllArticles.Height = 256.5
        ufAllArticles.bLaunch.Top = 204
        ufAllArticles.bLaunch.Left = 126
        ufAllArticles.bCancel.Top = 204
        ufAllArticles.bCancel.Left = 210
        ufAllArticles.lblMain.Caption = "Full List of Excel VBA Tutorials"
        ufAllArticles.Caption = "List of all wellsr.com Excel VBA Tutorials"
        ufAllArticles.cbModules.Enabled = False
        ufAllArticles.bImport.Enabled = False
        ufAllArticles.rbExisting.Enabled = False
        ufAllArticles.rbNew.Enabled = False
        ufAllArticles.Frame1.Enabled = False
        ufAllArticles.cbGroups.ListIndex = (ufAllArticles.cbGroups.ListCount - 1)
    End If
    Load ufAllArticles
    ufAllArticles.Show
End If
    Set xmlHTTPReq = Nothing
    Set htmlDOC = Nothing
    Application.ScreenUpdating = True
    If IsFormLoaded("ufAllArticles") Then Unload ufAllArticles
    Err.Clear
    On Error GoTo 0
    Exit Sub
FailedToLoad:
    Set xmlHTTPReq = Nothing
    Set htmlDOC = Nothing
    Application.ScreenUpdating = True
    If IsFormLoaded("ufAllArticles") Then Unload ufAllArticles
    MsgBox "Unable to connect to wellsr.com. Please check your internet connection and try again." & vbNewLine & vbNewLine & _
           "If this message persists, your system may be blocking external connections." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unable to Connect"
    Err.Clear
On Error GoTo 0
End Sub



Private Function IsFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded



Private Sub ListModules()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim i As Integer, lcount As Integer
    For i = ufAllArticles.cbModules.ListCount - 1 To 0 Step -1
        'delete all articles
        ufAllArticles.cbModules.RemoveItem i
    Next i
    
    Set VBProj = ActiveWorkbook.VBProject
    i = 0
    ufAllArticles.cbModules.ColumnWidths = "75;75"
    For Each VBComp In VBProj.VBComponents
        With ufAllArticles.cbModules
            .AddItem
            .List(i, 0) = VBComp.Name
            .List(i, 1) = ComponentTypeToString(VBComp.Type)
        End With
        i = i + 1
    Next VBComp
    ufAllArticles.cbModules.ListIndex = 0
End Sub

    
Private Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function


Attribute VB_Name = "wellsrModRibbon"
Option Explicit
Public wellsrCustomRibbon As IRibbonUI
Public sINI_FILE As String         'store path for configuration ini file here
Public sDOWNLOADS As String        'path where auto-update downloads will be stored
Public sRESOURCES As String        'path where downloaded zip files will be unzipped will be stored
Public sMACROS As String           'path where personal user macros will be stored
'For dropdown menu
Public wellsrTSNavItems As Integer      'Number of items in ribbon dropdown
Public wellsrTSNavArray() As String     'Store names of visible sheets
'REW for Dropdown
Public iSelected As Integer       'Return the currently selected item in the dropdown menu
'REW for toggle button
Public bTrust As Boolean         'for if toggle trust is activated


Public Sub wellsrHandleRibbon(Control As IRibbonControl)
Dim strOrig As String
Dim v As Variant
On Error GoTo wellsrHandleError:
    Select Case Control.id
        'Tools buttons
        Case "bHelp"
            QuickLaunch.LaunchwellsrHelp
        Case "bEdit"
            If Not ActiveCell Is Nothing And iSelected > 0 Then
                strOrig = ActiveCell.Formula
                ActiveCell.Formula = "=" & wellsrTSNavArray(iSelected, 0) & "()"
                v = Application.Dialogs(xlDialogFunctionWizard).Show
                If v = False Then
                    ActiveCell.Formula = strOrig
                End If
            ElseIf iSelected = 0 Then
                MsgBox "Select a function from the ""Custom Functions"" dropdown menu.", , "Select Function - wellsrPRO"
            End If
        Case "bDonate"
            Load ufDonate
            ufDonate.Show
        Case "bMyMacros"
            Load ufAddMacros
            ufAddMacros.Show
        Case "bConsult"
            Application.Run "myXML.XMLLaunchWebsite", "https://ask.wellsr.com/" & "?source=wellsrPROxl"
        Case "bShareMacros"
            Application.Run "INIPersonalLibrary.DisplayShareMacrosUF"
    End Select
Exit Sub
wellsrHandleError:
MsgBox "Error encountered while processing the wellsrPRO add-in." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub

Public Sub wellsrRibbon(ribbon As IRibbonUI)
Dim wb As Workbook
On Error GoTo wellsrRibbonError:
    Set wellsrCustomRibbon = ribbon
    Application.Run "wellsrmodRibbon.wellsrpopulateDD"
    Application.Run "myXML.grabRSS"
    Application.Run "wellsrmodRibbon.CreateEnvironment"
    Application.Run "Setup.InitialSetup"
    Application.Run "mCheckForUpdates.CheckForUpdates"
    Application.Run "mCheckForMessages.CheckForDonation"
    Application.Run "mCheckForMessages.CheckForMessages1"
    On Error Resume Next
    wellsrCustomRibbon.InvalidateControl ("wellsrdrp_Navigate")
    On Error GoTo wellsrRibbonError:
Exit Sub
wellsrRibbonError:
MsgBox "Error encountered while loading the wellsrPRO add-in." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub


Private Sub CreateEnvironment()
Dim sFolder As String
Dim dirExists As Boolean

'make INI file if it doesn't exist
    sFolder = Environ("PROGRAMDATA") & "\wellsr"
    If Len(Dir(sFolder, vbDirectory)) = 0 Then
        MkDir sFolder
    End If
    sINI_FILE = sFolder & "\wellsrtools.ini"
    
'Also make downloads folder
    sDOWNLOADS = sFolder & "\downloads"
    If Len(Dir(sDOWNLOADS, vbDirectory)) = 0 Then
        MkDir sDOWNLOADS
    End If
'And make folder for unzipping
    sRESOURCES = sFolder & "\resources"
    If Len(Dir(sRESOURCES, vbDirectory)) = 0 Then
        MkDir sRESOURCES
    End If
'Make folder where the user's personal macros will be stored
    sMACROS = sFolder & "\MyMacros"
    If Len(Dir(sMACROS, vbDirectory)) = 0 Then
        MkDir sMACROS
    End If

'test if the old APPDATA infrastructure exists.
'If it does, copy files and folders from it to PROGRAMDATA
    If FileExists(Environ("APPDATA") & "\wellsr\wellsrtools.ini") And FileExists(sINI_FILE) = False Then
        Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        FileCopy Environ("APPDATA") & "\wellsr\wellsrtools.ini", sFolder & "\wellsrtools.ini"
        objFSO.copyFolder Environ("APPDATA") & "\wellsr\*", sFolder & "\", True
        Set objFSO = Nothing
    End If

'if INI file doesn't exist, add a few starter entries (version and messages)
'...
'...
End Sub

'----------------------------------------------------------------------------
'--------------- Everything below here is for the Dropdown menu on the ribbon
'----------------------------------------------------------------------------
Private Sub wellsrpopulateDD()
    Dim indexcount As Integer
    Dim strDesc() As String
    Dim LinterpDesc As String
    Dim CompareDesc As String
    Dim LocateDesc As String
    Dim TrapIntDesc As String
    Dim DrawArrowsDesc As String
    Dim ArrayMatchDesc As String
    Dim TimeToWordsDesc As String
    Dim SuperMidDesc As String
    Dim ReplaceNDesc As String
    Dim SayItDesc As String
    Dim RGBnDesc As String
    
    indexcount = 10
    If ActiveWorkbook Is Nothing Then GoTo 101:
    
    LinterpDesc = "2D Linear Interpolation function that automatically picks which range to interpolate between based on the closest KnownX value to the NewX value you want to interpolate for."
    CompareDesc = "Compares Cell1 to Cell2 and if identical returns ""-"" by default but a different optional match string can be given. If cells are different, the output will either be ""FALSE"" or will optionally show the delta between the values if numeric."
    LocateDesc = "Determines if a given value is located somewhere in a range."
    TrapIntDesc = "Approximates the integral (the area under the curve) of a data set by using the trapezoidal rule."
    DrawArrowsDesc = "This function draws arrows or lines from the middle of one cell to the middle of another. Custom endpoints and shape colors are suppported."
    ArrayMatchDesc = "Return the value in one range based on the relative position of a value in another range."
    TimeToWordsDesc = "Converts decimal times to words. Function is capable of converting hours, minutes and seconds from decimal format (10.5) to words (""10 Minutes 30 Seconds"")."
    SuperMidDesc = "Easily extract text between two strings with this VBA Function. This function can extract a substring between two characters, delimiters, words and more. The delimiters can be the same delimiter or unique strings."
    ReplaceNDesc = "Replace Nth Occurrence of substring in a string."
    SayItDesc = "Function that speaks the content of whatever cell you want."
    RGBnDesc = "Convert colors defined with RGB to a long data type. This can be paired with the DrawArrows function to draw arrows of different colors."

    ReDim wellsrTSNavArray(0 To indexcount, 1) '2 column in the array (change 2nd # to higher # for more columns)
       wellsrTSNavArray(0, 0) = "Select a Function"
       
       'linterp
       wellsrTSNavArray(1, 0) = "Linterp"
       wellsrTSNavArray(1, 1) = "KnownYs, KnownXs, NewX"
       ReDim strDesc(1 To 3)
       strDesc(1) = "1-dimensional range containing your known Y values."
       strDesc(2) = "1-dimensional range containing your known X values."
       strDesc(3) = "The value you want to linearly interpolate on."
       Call RegisterUDF(wellsrTSNavArray(1, 0), LinterpDesc, strDesc)
       
       'compare
       wellsrTSNavArray(2, 0) = "compare"
       wellsrTSNavArray(2, 1) = "Cell1, Cell2, [CaseSensitive], [delta], [MatchString]"
       ReDim strDesc(1 To 5)
       strDesc(1) = "First cell to compare."
       strDesc(2) = "cell to compare against Cell1."
       strDesc(3) = "Optional Boolean that if set to TRUE will perform a case-sensitive comparison of the two entered cells. Defaults is TRUE."
       strDesc(4) = "Optional Boolean that if set to TRUE will display the delta between Cell1 and Cell2. Default is FALSE."
       strDesc(5) = "Optional string the user can choose to display when Cell1 and Cell2 match. Default is "" - ""."
       Call RegisterUDF(wellsrTSNavArray(2, 0), CompareDesc, strDesc)
       
       'locate
       wellsrTSNavArray(3, 0) = "locate"
       wellsrTSNavArray(3, 1) = "LookFor, InRange, [MatchString], [CaseSensitive]"
       ReDim strDesc(1 To 4)
       strDesc(1) = "Cell or string you want to look for in the range."
       strDesc(2) = "Range you want to search to find the value given in the variable LookFor."
       strDesc(3) = "Optional boolean that if set to TRUE will perform a case-sensitive search for the variable LookFor in the range InRange. Default is TRUE."
       strDesc(4) = "Optional string the user can choose to display when LookFor is in InRange. Default is to display the location of the found value."
       Call RegisterUDF(wellsrTSNavArray(3, 0), LocateDesc, strDesc)
       
       'TrapIntegration
       wellsrTSNavArray(4, 0) = "TrapIntegration"
       wellsrTSNavArray(4, 1) = "KnownXs, KnownYs"
       ReDim strDesc(1 To 2)
       strDesc(1) = "1-dimensional range containing your known X values."
       strDesc(2) = "1-dimensional range containing your known Y values."
       Call RegisterUDF(wellsrTSNavArray(4, 0), TrapIntDesc, strDesc)
       
       'DrawArrows
       wellsrTSNavArray(5, 0) = "DrawArrows"
       wellsrTSNavArray(5, 1) = "FromRange, ToRange, [RGBcolor], [LineType]"
       ReDim strDesc(1 To 4)
       strDesc(1) = "Cell or Range where you want your arrow to begin."
       strDesc(2) = "Cell or Range where you want your arrow to end."
       strDesc(3) = "Optional Long variable representing the color you want your arrow to be. To convert an RGB color to a Long variable, use the function RGBn(r,g,b). Default is RGBn(228, 108, 10)."
       strDesc(4) = "Optional String representing how you want the endpoints of your arrow to be. Options are ""Double"" for an arrow with 2 sides, ""Line"" for a line with no arrow endpoints and ""Single"" (default) for an arrow pointing to ToRange."
       Call RegisterUDF(wellsrTSNavArray(5, 0), DrawArrowsDesc, strDesc)
       
       'ArrayMatch
       wellsrTSNavArray(6, 0) = "ArrayMatch"
       wellsrTSNavArray(6, 1) = "xy, array1, array2"
       ReDim strDesc(1 To 3)
       strDesc(1) = "A single cell within range array1."
       strDesc(2) = "The range that xy exists within."
       strDesc(3) = "The array you want to look up to find the value in the cell corresponding to the position of xy inside array1. The size of array1 and array2 must match."
       Call RegisterUDF(wellsrTSNavArray(6, 0), ArrayMatchDesc, strDesc)

       'TimeToWords
       wellsrTSNavArray(7, 0) = "TimeToWords"
       wellsrTSNavArray(7, 1) = "rng1, [units]"
       ReDim strDesc(1 To 2)
       strDesc(1) = "1-dimensional range or value you want to convert to words."
       strDesc(2) = "Base units your input is in. Can be ""h"", ""m"" or ""s"" for hours, minutes and seconds, respectively. Default is minutes."
       Call RegisterUDF(wellsrTSNavArray(7, 0), TimeToWordsDesc, strDesc)

       'SuperMid
       wellsrTSNavArray(8, 0) = "SuperMid"
       wellsrTSNavArray(8, 1) = "strMain, str1, str2, [reverse]"
       ReDim strDesc(1 To 4)
       strDesc(1) = "The main cell or string you want to extract the substring from."
       strDesc(2) = "Your first delimiter string."
       strDesc(3) = "Your second delimiter string."
       strDesc(4) = "Optional boolean that when set to True, an InStrRev search will occur to find the last instance of the substrings in your main string."
       Call RegisterUDF(wellsrTSNavArray(8, 0), SuperMidDesc, strDesc)

       'ReplaceN
       wellsrTSNavArray(9, 0) = "ReplaceN"
       wellsrTSNavArray(9, 1) = "str1, strFind, strReplace, N, [Count]"
       ReDim strDesc(1 To 5)
       strDesc(1) = "The main string you want to perform the replacements on."
       strDesc(2) = "The string you want to replace in your main string."
       strDesc(3) = "The new string you want to insert into your main string."
       strDesc(4) = "An integer representing the instance of the string strFind you want to replace."
       strDesc(5) = "Optional integer declaring how many replacements you want to make."
       Call RegisterUDF(wellsrTSNavArray(9, 0), ReplaceNDesc, strDesc)

       'SayIt
       wellsrTSNavArray(10, 0) = "SayIt"
       wellsrTSNavArray(10, 1) = "cell1, [bAsync], [bPurge]"
       ReDim strDesc(1 To 5)
       strDesc(1) = "The cell or string you want Excel to say."
       strDesc(2) = "Optional boolean that if set to True will allow Excel to finish saying your string in the background. Setting this value to True speeds up execution."
       strDesc(3) = "Optional boolean that if set to True will cause your function to stop saying anything it was in the process of saying and start saying your new string."
       Call RegisterUDF(wellsrTSNavArray(10, 0), SayItDesc, strDesc)

       'RGBn - don't include this in the ribbon
       ReDim strDesc(1 To 3)
       strDesc(1) = "Red (0 to 255)"
       strDesc(2) = "Green (0 to 255)"
       strDesc(3) = "Blue (0 to 255)"
       Call RegisterUDF("RGBn", RGBnDesc, strDesc)

       'wellsrCustomRibbon.InvalidateControl ("wellsrdrp_Navigate")
       wellsrTSNavItems = indexcount + 1
       Exit Sub
101:
    wellsrTSNavItems = 0
End Sub

 
'This gets called for every item in the dropdown
'This callback gets called during startup and can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub wellsrdrp_Navigate_getItemCount(Control As IRibbonControl, ByRef returnedVal)
'
  'The total number of items in the dropdown box (index one based)
  returnedVal = wellsrTSNavItems
'
End Sub
'
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub wellsrdrp_Navigate_getItemID(Control As IRibbonControl, index As Integer, ByRef id)
'
  'This is an internal id used by xl.  It must be unique
  id = "wellsrTSNavSheets" & index
'
End Sub
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub wellsrdrp_Navigate_getItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'store the shape name here
  returnedVal = wellsrTSNavArray(index, 0)
'
End Sub
'
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub wellsrdrp_Navigate_getSelectedItemIndex(Control As IRibbonControl, ByRef returnedVal)
' Just select the selected value
  returnedVal = iSelected
'
End Sub
'
 
Public Sub wellsrdrp_Navigate_getItemSupertip(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'store the sheet and style here as a supertip (hover)
  returnedVal = "Function: " & wellsrTSNavArray(index, 0) & vbNewLine & "Input: " & wellsrTSNavArray(index, 1)
End Sub
 
'This callback only fires when a user changes the dropdown
Public Sub wellsrTS_Navigate_OnAction(Control As IRibbonControl, id As String, index As Integer)
'Remember the selected value
  iSelected = index
  Application.Run "wellsrmodRibbon.wellsrpopulateDD"
  If iSelected > wellsrTSNavItems Then
    iSelected = wellsrTSNavItems
  End If
  On Error Resume Next
  wellsrCustomRibbon.InvalidateControl ("wellsrdrp_Navigate")   'Reset the dropdown control
End Sub
 

'---------------------------------------------------------------------------
'--------------- For enabling/disabling the dropdown and edit/delete buttons
'---------------------------------------------------------------------------
Public Sub wellsrdrp_Navigate_getEnable(Control As IRibbonControl, ByRef returnedVal)
    'enables/disables all items with XML getEnabled attribute set to "drp_Navigate_getEnable"
    'triggers if the add-in is loaded but no workbooks are open or if there are no shapes to edit
'MsgBox "test2"
Call wellsrpopulateDD
    If ActiveWorkbook Is Nothing Or wellsrTSNavItems = 0 Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub



'-------------------------------------
'-------------- Add Macro Descriptions
'-------------------------------------
Private Sub RegisterUDF(fx As String, s As String, strDesc() As String)
    On Error Resume Next
    Application.MacroOptions Macro:=fx, Description:=s, ArgumentDescriptions:=strDesc, Category:="wellsrPRO"
    Err.Clear
    On Error GoTo 0
End Sub


'Private Sub UnregisterUDF()
'    Application.MacroOptions Macro:="IFERROR", Description:=Empty, Category:=Empty
'End Sub



'-------------------------------------
'-------------- Everything below here is for the toggle trust access button
'-------------------------------------
Public Sub buttonToggle(Control As IRibbonControl, pressed As Boolean)
On Error GoTo 100:
    
    'If the toggle button is pressed, bTrust is set to true and the custom ribbon is reloaded
    bTrust = pressed
    'if the button is clicked, toggle trust access
    If VBAIsTrusted <> bTrust Then
        ParseHTML.ButtonToggleTrust (bTrust)
    End If
    'reload ribbon (updates label text/image and buttons enabled/disabled)
    If VBAIsTrusted <> bTrust Then
        'somehow I didn't toggle trust access. Display a message
        If VBAIsTrusted = True Then
            MsgBox "I failed to disable Trust Access." & Chr(10) & Chr(10) & _
            "To disable this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Check the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "Auto Disable Failed"
        Else
            MsgBox "I failed to enable Trust Access." & Chr(10) & Chr(10) & _
            "To enable this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Check the box next to ""Trust Access to the VBA project object model""" & vbNewLine & vbNewLine & _
            "THIS STEP MUST BE COMPLETE IN ORDER TO AUTOMATICALLY IMPORT VBA TUTORIAL EXAMPLES", vbOKOnly, "Auto Enable Failed"
        End If
        pressed = Not pressed
        bTrust = Not bTrust
    End If
    
    On Error Resume Next
    wellsrCustomRibbon.InvalidateControl ("tTrust")
    Err.Clear
    On Error GoTo 0
Exit Sub
100:
MsgBox "Unknown error while loading the wellsrPRO Ribbon.", vbCritical, "Error Encountered"
End Sub

Public Sub wxPressed(Control As IRibbonControl, ByRef pressed)
    'If the variable bToggle is true, press the button
    If VBAIsTrusted <> bTrust Then
        bTrust = Not bTrust
    End If
    
    If bTrust Then
        pressed = True
    Else
        pressed = False
    End If

    On Error Resume Next
    wellsrCustomRibbon.InvalidateControl ("tTrust")
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub wxLabel(Control As IRibbonControl, ByRef returnedVal)
    'Changes the label on any button with XML getLabel attribute set to "rxLabel"
    If bTrust Then
        returnedVal = "Trust Access Enabled"
    Else
        returnedVal = "Trust Access Disabled"
    End If
End Sub

Public Sub wxImage(Control As IRibbonControl, ByRef returnedVal)
    'Changes the image on any button with XML getImage attribute set to "rxImage"
    '   -This only works if using one of the default MS imageMso images
    '    list: http://soltechs.net/CustomUI/imageMso01.asp?gal=1&count=no
    If bTrust Then
        returnedVal = "AcceptInvitation"
    Else
        returnedVal = "DeclineInvitation"
    End If
End Sub


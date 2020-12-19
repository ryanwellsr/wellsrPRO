Attribute VB_Name = "myXML"
Option Explicit
Public XMLNavItems As Integer      'Number of items in XML ribbon dropdown
Public strArticles() As String     'Store names of articles here

'REW for Dropdown
Public iSelectedXML As Integer       'Return the currently selected item in the XML dropdown menu

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Private Sub GrabRSS()
'NOTE: Since the DOMDocument .Load method was failing for xmlobject on some machines (access denied),
'      I loaded the xml file with XMLHTTP, then I read the response text as the DOMDocument (avoids the LOAD function).
Dim xmlObject As Object
Dim xmlobject2 As Object
Dim i As Long, j As Long
Dim xmlNode As Object 'MSXML2.IXMLDOMNode
Dim xmlNode2 As Object 'MSXML2.IXMLDOMNode
Dim xmlNode1 As Object
Dim intlength As Long, intlength2 As Long
Dim iCount As Integer
Dim now10 As Date
'Dim strArticles() As String
                '1,x is title
                '2,x is description
                '3,x is link
                '4,x is published date (pubDate)
On Error GoTo xmlErr:
iCount = 0
'create msxml object
Set xmlObject = CreateObject("MSXML2.ServerXMLHTTP") 'New MSXML2.DOMDocument60 'CreateObject("MSXML2.DOMDocument.3.0")
Set xmlobject2 = CreateObject("MSXML2.DOMDocument") 'New MSXML2.DOMDocument60 'CreateObject("MSXML2.DOMDocument.3.0")
 
xmlObject.Open "POST", "https://wellsr.com/vba/feed.xml", False
xmlObject.send
 
 
'!!!This was failing for some machines with different security settings!!!
'!!!   xmlObject in this context was a DOMDocument (not the XMLHTTP)   !!!
'add new xml references and on error try loading with them
'xmlObject.async = False
'load the xml page
'xmlObject.Load ("https://wellsr.com/vba/feed.xml")
 
'Make sure the page is loaded
    now10 = DateAdd("s", 10, Now())
    Do While xmlObject.readyState <> 4
        'Try for 30 seconds
        DoEvents
        If Now >= now10 Then
            GoTo xmlErr:
        End If
    Loop
    If xmlObject.Status <> 200 Then
        GoTo xmlErr:
    End If
 
    'since the DOMDocument .Load method was failing for xmlobject,
    xmlobject2.LoadXML (xmlObject.responseText)
 
intlength = xmlobject2.ChildNodes.Length - 1
If intlength = -1 Then GoTo xmlErr:
For i = 0 To intlength
'    MsgBox xmlobject2.ChildNodes.Item(i).BaseName
    If xmlobject2.ChildNodes.Item(i).BaseName = "rss" Then
        Set xmlNode = xmlobject2.ChildNodes.Item(i)
        Exit For
    End If
Next i
'get the result node
intlength = xmlNode.ChildNodes.Length - 1
For i = 0 To intlength
'    MsgBox xmlNode.ChildNodes.Item(i).BaseName
    If xmlNode.ChildNodes.Item(i).BaseName = "channel" Then
        Set xmlNode1 = xmlNode.ChildNodes.Item(i)
        Exit For
    End If
Next i
'get the rate node
intlength = xmlNode1.ChildNodes.Length - 1
For i = 0 To intlength
        'Finally drilled down to the articles... Pull titles
        If xmlNode1.ChildNodes.Item(i).BaseName = "item" Then
            Set xmlNode2 = xmlNode1.ChildNodes.Item(i)
            intlength2 = xmlNode2.ChildNodes.Length - 1
            ReDim Preserve strArticles(1 To 4, 0 To iCount)
            For j = 0 To intlength2
                'pull article titles
                If xmlNode2.ChildNodes.Item(j).BaseName = "title" Then
                    strArticles(1, iCount) = RemoveHTML(xmlNode2.ChildNodes.Item(j).Text)
                ElseIf xmlNode2.ChildNodes.Item(j).BaseName = "description" Then
                    strArticles(2, iCount) = RemoveHTML(Mid(xmlNode2.ChildNodes.Item(j).Text, 1, 250) & "...")
                ElseIf xmlNode2.ChildNodes.Item(j).BaseName = "link" Then
                    strArticles(3, iCount) = xmlNode2.ChildNodes.Item(j).Text
                ElseIf xmlNode2.ChildNodes.Item(j).BaseName = "pubDate" Then
                    strArticles(4, iCount) = xmlNode2.ChildNodes.Item(j).Text
                End If
            Next j
            iCount = iCount + 1
        End If
Next i
 
XMLNavItems = UBound(strArticles, 2) + 1
Set xmlNode2 = Nothing
Set xmlNode1 = Nothing
Set xmlNode = Nothing
Set xmlObject = Nothing
Set xmlobject2 = Nothing
Exit Sub
 
xmlErr:
ReDim strArticles(1 To 4, 0 To 9) As String
strArticles(1, 0) = "Cannot retrieve feed - Check internet connection"
strArticles(2, 0) = "XML feed not found. Please check your internet connection. " & vbNewLine & vbNewLine & "This may also occur on Apple computers and Windows machines older than Windows 2000."
strArticles(3, 0) = "https://wellsr.com/vba/"
strArticles(4, 0) = "Error " & Err.Number & ": " & Err.Description
Err.Clear
For i = 0 To 9
    strArticles(3, i) = "https://wellsr.com/vba/"
Next i
XMLNavItems = 1
Set xmlNode2 = Nothing
Set xmlNode = Nothing
Set xmlObject = Nothing
Set xmlobject2 = Nothing
End Sub


Public Sub XMLHandleRibbon(Control As IRibbonControl)
    Select Case Control.id
        Case "bLaunch"
            On Error GoTo XMLHandleError:
            Call XMLLaunchWebsite(strArticles(3, iSelectedXML) & "?source=wellsrPROxl")
            On Error GoTo 0
        Case "bRandom"
            On Error GoTo XMLHandleError:
            'Application.Run "parsehtml.ImportMacroExamples", strArticles(3, iSelectedXML)
            Call XMLLaunchWebsite(strArticles(3, Int((9 - 0 + 1) * Rnd + 0)) & "?source=wellsrPROxl")
            On Error GoTo 0
        Case "bHome"
            On Error GoTo FailedToLaunchList:
            Application.Run "parsehtml.ListAllArticles", True 'only show articles
            On Error GoTo 0
        Case "bImport"
            On Error GoTo FailedToLaunchList:
            Application.Run "parsehtml.ListAllArticles", False 'show articles with option to import
            On Error GoTo 0
        Case "bFavorites"
            On Error GoTo FailedToLaunchList:
            Application.Run "QuickLaunch.LaunchufFavorites"
            On Error GoTo 0
        Case "bFeedback"
            On Error GoTo FailedToLoadFeedback:
            Application.Run "QuickLaunch.LaunchufFeedback"
            On Error GoTo 0
        Case "bUpdate"
            On Error GoTo FailedToLoadUpdates:
            Application.Run "mCheckForUpdates.CheckForUpdates", True
            On Error GoTo 0
    End Select
Exit Sub
XMLHandleError:
MsgBox "Error encountered while trying to launch VBA article." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
Exit Sub
FailedToImport:
MsgBox "Error encountered while trying to import macros." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
Exit Sub
FailedToLaunchList:
MsgBox "Error encountered while trying to display list of articles." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
Exit Sub
FailedToLoadFeedback:
MsgBox "Error encountered while trying to load the feedback form." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
Exit Sub
FailedToLoadUpdates:
MsgBox "Error encountered while checking for updates." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub


Private Sub XMLLaunchWebsite(strurl As String)
On Error GoTo XMLLaunchError
    Dim r As Long
    r = ShellExecute(0, "open", strurl, 0, 0, 1)
    If r = 5 Then 'if access denied, try this alternative
            r = ShellExecute(0, "open", "rundll32.exe", "url.dll,FileProtocolHandler " & strurl, 0, 1)
    End If
Exit Sub
XMLLaunchError:
MsgBox "Error encountered while trying to launch VBA article." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub


'---------------------------
'STRIP OUT HTML FROM STRINGS
'---------------------------
Private Function RemoveHTML(sString As String) As String
    On Error GoTo Error_Handler
    Dim oRegEx          As Object
 
    Set oRegEx = CreateObject("vbscript.regexp")
 
    With oRegEx
        'Patterns see: http://regexlib.com/Search.aspx?k=html%20tags
        '.Pattern = "<[^>]+>"    'basic html pattern
        .Pattern = "<!*[^<>]*>"    'html tags and comments
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With
 
    RemoveHTML = oRegEx.Replace(sString, "")
Error_Handler_Exit:
    On Error Resume Next
    Set oRegEx = Nothing
    Exit Function
 
Error_Handler:
RemoveHTML = sString
Set oRegEx = Nothing
End Function


'------------------------
'FOR CONTROLLING DROPDOWN
'------------------------

Public Sub XML_getItemCount(Control As IRibbonControl, ByRef returnedVal)

  'The total number of items in the dropdown box (index one based)
  returnedVal = XMLNavItems
'
End Sub
'
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub XML_getItemID(Control As IRibbonControl, index As Integer, ByRef id)
'
  'This is an internal id used by xl.  It must be unique
  id = "XMLNav" & index
'
End Sub
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub XML_getItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'store the shape name here
  returnedVal = strArticles(1, index)
'
End Sub
'
'This callback gets called during startup automatically.  It can be called by using
'invalidateControl (ie. CustomRibbon.InvalidateControl("drp_Navigate")
Public Sub XML_getSelectedItemIndex(Control As IRibbonControl, ByRef returnedVal)
' Just select the selected value
  returnedVal = iSelectedXML
'
End Sub
'
 
Public Sub XML_getItemSupertip(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'store the sheet and style here as a supertip (hover)
    returnedVal = strArticles(2, index) & vbNewLine & vbNewLine & strArticles(4, index) & vbNewLine & vbNewLine & strArticles(3, index)
End Sub
 
'This callback only fires when a user changes the dropdown
Public Sub XML_OnAction(Control As IRibbonControl, id As String, index As Integer)
'Remember the selected value
  iSelectedXML = index
'''    Application.Run "wellsrmodRibbon.wellsrpopulateDD"
'''    If iSelected > wellsrTSNavItems Then
'''      iSelected = wellsrTSNavItems
'''    End If
'''    On Error Resume Next
'''    wellsrCustomRibbon.Invalidate 'Control ("drp_Navigate")   'Reset the dropdown control
End Sub
 


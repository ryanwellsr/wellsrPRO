Attribute VB_Name = "mCheckForUpdates"
'--------------------------------------
'------------- Check for Product Updates and Messages
'--------------------------------------
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub CheckForUpdates(Optional bRibbonClick As Boolean)
Const strurl As String = "https://wellsr.com/wellsrPROconnect.txt"
'define keywords for pulling data from website
Const strVerString  As String = "#VERSION#"
Const strLinkString As String = "#LINK#"
Const strMessID     As String = "#MESSAGEID#"
Const strMessString As String = "#MESSAGE#"
Const strMessPath   As String = "#MESSAGEPATH#"
Const strMessStart  As String = "#MESSAGESTART#"
Const strMessEnd    As String = "#MESSAGEEND#"
Const strMessLimit  As String = "#MESSAGELIMIT#"

Dim strConnect() As String
Dim SubVersions() As String
Dim SubVersionsLatest() As String
Dim strDownloadLink As String
Dim i As Integer, j As Integer  'counters
Dim HttpRequest As Object
Dim HttpDownload As Object  'for downloading file
Dim oStream As Object       'for downloading file
Dim LatestVersion As String
Dim v As Variant
Dim strInstalled As String  'Location of currently installed add-in
Dim strNewXLAM As String    'path to the newly downloaded XLAM file
Dim strCurrPath As String   'path where the current active add-in is installed
Dim strRead As String       'Read the message variables from my website
Dim strReturn As String     'Result from printing to ini file
On Error Resume Next
Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")

HttpRequest.Open "POST", strurl, False
HttpRequest.send

If HttpRequest.Status = 200 Then
    'successfully connected
    strConnect() = Split(HttpRequest.responseText, Chr(10))
    For i = LBound(strConnect) To UBound(strConnect)
        strConnect(i) = Application.Clean(Trim(strConnect(i)))
        If InStr(strConnect(i), strVerString) <> 0 Then
            'you found the version keyword
            i = i + 1
            LatestVersion = Replace(Replace(strConnect(i), Chr(10), ""), Chr(13), "")
            If LatestVersion <> strVersion Then
                SubVersionsLatest = Split(LatestVersion, ".")
                SubVersions = Split(strVersion, ".")
                'Make sure the ve version on my website isn't somehow an older version (like if pulled from cache, perhaps or I screwed up)
                If SubVersionsLatest(0) > SubVersions(0) Then
                    GoTo NewerVersionDetected:
                ElseIf SubVersionsLatest(0) = SubVersions(0) And SubVersionsLatest(1) > SubVersions(1) Then
                    GoTo NewerVersionDetected:
                ElseIf SubVersionsLatest(0) = SubVersions(0) And SubVersionsLatest(1) = SubVersions(1) And SubVersionsLatest(2) > SubVersions(2) Then
                    GoTo NewerVersionDetected:
                End If
                GoTo VersionsMatch 'If you make it here, no newer version is detected
                    
NewerVersionDetected:
            If Trim(LatestVersion) <> "" Then
                    v = MsgBox("A newer version of wellsrPRO is available. Would you like to install it now?" & vbNewLine & vbNewLine & _
                             "Your Version:   " & strVersion & vbNewLine & _
                             "Latest Version: " & LatestVersion, vbYesNo, "Updates Available")
            End If
VersionsMatch:
            End If
            If bRibbonClick = True And IsEmpty(v) Then
                If Trim(LatestVersion) <> "" Then
                    MsgBox "Woohoo! You have the latest version of wellsrPRO installed." & vbNewLine & vbNewLine & _
                                 "Your Version:   " & strVersion & vbNewLine & _
                                 "Latest Version: " & LatestVersion, vbOKOnly, "No Updates Available"
                Else
                    MsgBox "Unable to check the latest version. Please check your internet connection and try again.", _
                            vbOKOnly, "Unable to check for updates"
                End If
            End If
        ElseIf InStr(strConnect(i), strLinkString) <> 0 And v = vbYes Then
            'you found the download link to the latest version and the user wants to download it
            '(Step 1) Display your Progress Bar
            ufUpdate.LabelProgress.Width = 0
            ufUpdate.Show
            FractionComplete 0.1, "Downloading Latest Version..."
            i = i + 1
            strDownloadLink = Replace(Replace(strConnect(i), Chr(10), ""), Chr(13), "")
            Set HttpDownload = CreateObject("MSXML2.ServerXMLHTTP")
            
            HttpDownload.Open "GET", strDownloadLink, False
            HttpDownload.send
            FractionComplete 0.2, "Downloading Latest Version..."
            If HttpDownload.Status = 200 Then
                'download zip file
                Set oStream = CreateObject("ADODB.Stream")
                oStream.Open
                oStream.Type = 1
                oStream.write HttpDownload.responseBody
                FractionComplete 0.4, "Downloading Latest Version..."
                oStream.SaveToFile sDOWNLOADS & "\wellsrPRO" & LatestVersion & ".zip", 2 ' 1 = no overwrite, 2 = overwrite
                oStream.Close
                Set oStream = Nothing
                FractionComplete 0.5, "Preparing For Installation..."
                Sleep 100
                'delete all the files in the resources folder
                On Error Resume Next
                Kill sRESOURCES & "\*.*"
                On Error GoTo 0
                'unzip the file
                FractionComplete 0.6, "Unpacking Download..."
                Call UnZip(sRESOURCES, sDOWNLOADS & "\wellsrPRO" & LatestVersion & ".zip")
                Sleep 100
                'launch vbscript to overwrite the currently installed xlam file (moves to location of current installation)
                FractionComplete 0.75, "Installing New Version..."
                ufUpdate.LabelRestart.Visible = True
                Sleep 500
                strInstalled = ThisWorkbook.FullName
                strCurrPath = ThisWorkbook.Path & "\"
                strNewXLAM = sRESOURCES & "wellsrPRO.xlam"
                Call LaunchVBS(strNewXLAM, strInstalled, strCurrPath)
                'Application.OnTime Now + TimeValue("00:00:02"), "'LaunchVBS """ & strNewXLAM & """,""" & strInstalled & """,""" & strCurrPath & "'"
                'uninstall this add-in since the one will be installed in 5 seconds
                'ThisWorkbook.Close (False)
            Else
                Unload ufUpdate
                v = MsgBox("I seem to be having trouble installing the latest version. Would you like to manually install?" & vbNewLine & vbNewLine & _
                    "Error Code: " & HttpDownload.Status, vbYesNo, "Updates Available")
                If v = vbYes Then
                    Application.Run "myXML.XMLLaunchWebsite", "https://ask.wellsr.com/assets/files/wellsrPRO.zip"
                End If
            End If
            Set HttpDownload = Nothing
        ElseIf InStr(strConnect(i), strMessID) <> 0 Then
            i = i + 1
            strRead = Replace(Replace(strConnect(i), Chr(10), ""), Chr(13), "")
            strReturn = INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageID1", sINI_FILE)
            If strRead <> strReturn Then
                'this is a new message!
                'Update ini file with the contents
                strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageID1", sINI_FILE, strRead)
                For j = i To UBound(strConnect)
                    If InStr(strConnect(j), strMessString) <> 0 Then
                        'You find the promo message
                        j = j + 1
                        strRead = Replace(Replace(strConnect(j), Chr(10), ""), Chr(13), "")
                        'print message to ini file
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "Message1", sINI_FILE, strRead)
                    ElseIf InStr(strConnect(j), strMessPath) <> 0 Then
                        'you found the promo link
                        j = j + 1
                        strRead = Replace(Replace(strConnect(j), Chr(10), ""), Chr(13), "")
                        'print message to ini file
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessagePath1", sINI_FILE, strRead)
                    ElseIf InStr(strConnect(j), strMessStart) <> 0 Then
                        'you found when the promo begins
                        j = j + 1
                        strRead = Replace(Replace(strConnect(j), Chr(10), ""), Chr(13), "")
                        'print message to ini file
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageStart1", sINI_FILE, CStr(strRead))
                    ElseIf InStr(strConnect(j), strMessEnd) <> 0 Then
                        'you found when the promo ends
                        j = j + 1
                        strRead = Replace(Replace(strConnect(j), Chr(10), ""), Chr(13), "")
                        'print message to ini file
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageEnd1", sINI_FILE, CStr(strRead))
                    ElseIf InStr(strConnect(j), strMessLimit) <> 0 Then
                        'you found the maximum number of reminders to send the user
                        j = j + 1
                        strRead = Replace(Replace(strConnect(j), Chr(10), ""), Chr(13), "")
                        'print message to ini file
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageLimit1", sINI_FILE, CStr(strRead))
                        strReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageCount1", sINI_FILE, "0")
                    End If
                Next j
            End If
            'go ahead and exit the loop early if the message ID matches what's already in the ini file
            Exit For
        End If
    Next i
End If
Err.Clear
On Error GoTo 0
End Sub

Private Sub LaunchVBS(strCopyFrom As String, strCopyTo As String, strCurrPath As String)
'Sub to launch the VBscript file that will update the version of the add-in
Shell "wscript.exe """ & sRESOURCES & "\update.vbs""" & " """ & strCopyFrom & """ """ & strCopyTo & """ """ & strCurrPath & """"
End Sub

 

Private Sub UnZip(strTargetPath As String, Fname As Variant)
'unzip file
Application.ScreenUpdating = False
    Dim oApp As Object
    Dim FileNameFolder As Variant
 
    If Right(strTargetPath, 1) <> Application.PathSeparator Then
        strTargetPath = strTargetPath & Application.PathSeparator
    End If
 
    FileNameFolder = strTargetPath

    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items
    Set oApp = Nothing
Application.ScreenUpdating = True
End Sub



Sub FractionComplete(pctdone As Single, strCaption As String)
With ufUpdate
    .LabelCaption.Caption = strCaption
    .LabelProgress.Width = pctdone * (.FrameProgress.Width)
End With
DoEvents
End Sub

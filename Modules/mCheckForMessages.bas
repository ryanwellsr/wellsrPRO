Attribute VB_Name = "mCheckForMessages"
'*******************************************************************************
' Declaration for Reading and Writing to an INI file.
'*******************************************************************************
Option Private Module
'++++++++++++++++++++++++++++++++++++++++++++++++++++
' API Functions for Reading and Writing to INI File
'++++++++++++++++++++++++++++++++++++++++++++++++++++

' Declare for reading INI files.
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
                                          
    ' Declare for writing INI files.
    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
                                          
    ' Declare for writing INI files.
    Private Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
#End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++
' Enumeration for sManageSectionEntry funtion
'++++++++++++++++++++++++++++++++++++++++++++++++++++


'*******************************************************************************
' End INI file declaration Section.
'*******************************************************************************
'donate info
Dim MessageDate As Date
Dim limit As Integer
Dim count As Integer
Dim message As String
'promo info
Dim MessagePath1 As String
Dim MessageCount1 As Integer
Dim MessageID1 As String
Dim Message1 As String
Dim MessageStart1 As Date
Dim MessageEnd1 As Date
Dim MessageLimit1 As Integer

Sub CheckForDonation()
'Check to see if date for default message requesting donation has elapsed yet
Dim today As Date
Dim v As Variant
Dim sReturn As String
On Error Resume Next
today = Now()
MessageDate = INIfavorites.sManageSectionEntry(iniRead, "PRESETS", "DonateTime", sINI_FILE)
limit = CInt(INIfavorites.sManageSectionEntry(iniRead, "PRESETS", "DonateLimit", sINI_FILE))
count = CInt(INIfavorites.sManageSectionEntry(iniRead, "PRESETS", "DonateCount", sINI_FILE))
message = INIfavorites.sManageSectionEntry(iniRead, "PRESETS", "Donate", sINI_FILE)

If today >= MessageDate And MessageDate <> "12:00:00 AM" Then
    'check to see if you've already displayed the message
    If count < limit Then
        'Display a pretty userform asking for a donation
        If count = 0 Then
            ufDonate.lblCTA = "HURRY: 2 Days Only"
        ElseIf count = 1 Then
            ufDonate.lblCTA = "LAST CHANCE: Final Day"
        ElseIf count = 2 Then
            ufDonate.lblCTA = "It's Back: Today Only"
        End If
        Load ufDonate
        ufDonate.Show
    End If
End If
Err.Clear
On Error GoTo 0
End Sub

Sub CheckForMessages1()
'Check for messages from online (and now stored in the ini file)
Dim today As Date
Dim v As Variant
Dim sReturn As String
On Error Resume Next
today = Now()
MessageID1 = INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageID1", sINI_FILE)
Message1 = INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "Message1", sINI_FILE)
MessagePath1 = INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessagePath1", sINI_FILE)
MessageStart1 = CDate(INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageStart1", sINI_FILE))
MessageEnd1 = CDate(INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageEnd1", sINI_FILE))
MessageLimit1 = CInt(INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageLimit1", sINI_FILE))
MessageCount1 = CInt(INIfavorites.sManageSectionEntry(iniRead, "MESSAGES", "MessageCount1", sINI_FILE))

If today >= MessageStart1 And today <= MessageEnd1 And MessageCount1 < MessageLimit1 Then
    'if you're in the promotional window and you haven't already shown the advertisement too many times, display form
    Load ufPromo
    ufPromo.lblMessage.Caption = Message1
    ufPromo.Show
End If
Err.Clear
On Error GoTo 0
End Sub

Private Sub LaunchPromo()
'User requested more information about promo
Dim sReturn As String
If MessagePath1 <> "" Then
'launch promo website
    Application.Run "myXML.XMLLaunchWebsite", MessagePath1
    'don't show the offer again for 2 days
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageStart1", sINI_FILE, DateAdd("d", 2, Now))
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageCount1", sINI_FILE, CStr(MessageCount1 + 1))
    Call InteractionAnalytics("PROMO", True, MessagePath1)
Else
    MsgBox "We apologize - We were not able to find more information about this offer.", vbInformation, "Offer Not Found"
    Call InteractionAnalytics("PROMO", True, "Requested more info but Promo URL not found")
End If
End Sub

Private Sub ClosePromo()
'user did not click the advertisement.
'Remind them again in 12 hours
        MessageCount1 = MessageCount1 + 1
        sReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageStart1", sINI_FILE, DateAdd("h", 18, Now))
        sReturn = INIfavorites.sManageSectionEntry(iniWrite, "MESSAGES", "MessageCount1", sINI_FILE, CStr(MessageCount1))
        Call InteractionAnalytics("PROMO", False, MessagePath1)
End Sub

Private Sub DonateYes(Quantity As String)
'User clicked the donate button
Dim sReturn As String
Dim today As Date
today = Now()
        If InStr(1, Quantity, "http") = 0 Then
            Application.Run "myXML.XMLLaunchWebsite", "https://wellsr.com/vba/vba-cheat-sheets/" & Quantity & "?source=wellsrPROxl"
            Call InteractionAnalytics("CheatSheet", True, "https://wellsr.com/vba/vba-cheat-sheets/" & Quantity)
        Else
            Application.Run "myXML.XMLLaunchWebsite", Quantity
            Call InteractionAnalytics("CheatSheet", True, Quantity)
        End If
If today >= MessageDate Then
    'only update the entry in the ini file if they've had the add-in installed for more than the required days
        count = count + 1
        If count = 1 Then
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateTime", sINI_FILE, DateAdd("d", 1, Date))
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateCount", sINI_FILE, CStr(count))
        Else
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateTime", sINI_FILE, DateAdd("d", 91, Date))
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateCount", sINI_FILE, CStr(count))
        End If
End If
End Sub

Private Sub DonateNo()
'User did not click the donate button
Dim sReturn As String
Dim today As Date
today = Now()
Call InteractionAnalytics("DONATE", False, "No Donation")
If today >= MessageDate Then
    'only update the entry in the ini file if they've had the add-in installed for more than the required days
        count = count + 1
        If count = 1 Then
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateTime", sINI_FILE, DateAdd("d", 1, Date))
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateCount", sINI_FILE, CStr(count))
        Else
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateTime", sINI_FILE, DateAdd("d", 91, Date))
            sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateCount", sINI_FILE, CStr(count))
        End If
End If
End Sub















Private Sub InteractionAnalytics(strType As String, bMoreInfo As Boolean, strpath As String) 'strType="PROMO" for promo, "DONATE" for donation request
'Submit Feedback to Google Sheets
Dim HttpRequest As Object
Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
Dim strurl As String, strOS As String
Dim strVersion2 As String, strExcel As String 'return versions
Dim strInteraction As String
Dim strserial As String, strUsername As String
Dim strMAC As String, strIP As String
On Error Resume Next
strserial = GetDriveSerialNumber
strUsername = Environ$("Username")
strMAC = GetMyMACAddress
strIP = GetMyPublicIP
On Error GoTo 100:
strVersion2 = strVersion
strExcel = Application.Name & " " & Application.Version & "." & Application.Build
#If Win64 Then
   strExcel = strExcel & " (64-bit)"
#Else
   strExcel = strExcel & " (32-bit)"
#End If
strOS = Application.OperatingSystem
strInteraction = strType
If strType = "PROMO" Then
    strInteraction = strInteraction & " " & MessageID1
End If
If bMoreInfo = True Then
    strclick = "CLICKED"
Else
    strclick = "CANCELED"
End If
'https://docs.google.com/forms/d/e/1FAIpQLScLtqenhbtF1mHDjp0ppq4oWToNdkl_I7ClcZ4TWnWFmROh5Q/viewform?usp=pp_url&
'entry.1470534562=strInteraction&
'entry.1893264008=strClick&
'entry.1465812465=strPath&
'entry.1809480882=strOS&
'entry.1238441535=strExcel&
'entry.2015729830=strVersion2&
'entry.34657985=strUsername&
'entry.1248424676=strserial&
'entry.1023086801=strMAC&
'entry.1330624033=strIP
strurl = "https://docs.google.com/forms/d/e/1FAIpQLScLtqenhbtF1mHDjp0ppq4oWToNdkl_I7ClcZ4TWnWFmROh5Q/formResponse?ifq&" + _
    "&entry.1470534562=" + CStr(strInteraction) + _
    "&entry.1893264008=" + CStr(strclick) + _
    "&entry.1465812465=" + CStr(strpath) + _
    "&entry.1809480882=" + CStr(strOS) + _
    "&entry.1238441535=" + CStr(strExcel) + _
    "&entry.2015729830=" + CStr(strVersion2) + _
    "&entry.34657985=" + CStr(strUsername) + _
    "&entry.1248424676=" + CStr(strserial) + _
    "&entry.1023086801=" + CStr(strMAC) + _
    "&entry.1330624033=" + CStr(strIP) + _
    "&submit=Submit&ifq"
    
HttpRequest.Open "POST", strurl, False
HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
HttpRequest.send
Set HttpRequest = Nothing
Exit Sub
100:
    Set HttpRequest = Nothing
End Sub


      Private Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
      
          Dim fso As Object, Drv As Object
          Dim driveserial As Long
          
          'Create a FileSystemObject object
          Set fso = CreateObject("Scripting.FileSystemObject")
          
          'Assign the current drive letter if not specified
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(Application.Path))
          End If
      
          With Drv
              If .IsReady Then
                  driveserial = Abs(.SerialNumber)
              Else    '"Drive Not Ready!"
                  driveserial = -1
              End If
          End With
          
          'Clean up
          Set Drv = Nothing
          Set fso = Nothing
          
          GetDriveSerialNumber = driveserial
          
      End Function


Private Function GetMyMACAddress() As String
On Error GoTo 100:
    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myMACAddress    As String
    
    'Set the computer.
    strComputer = "."
    
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then myMACAddress = objItem.MACAddress
        Exit For
    Next
    
    'Return the IP string.
    GetMyMACAddress = myMACAddress
    Exit Function
100:
    GetMyMACAddress = ""
End Function


Private Function GetMyPublicIP() As String
    Dim HttpRequest As Object
On Error GoTo 100:
    On Error Resume Next
    'Create the XMLHttpRequest object.
    Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")

    'Check if the object was created.
    If Err.Number <> 0 Then
        'Return error message.
        GetMyPublicIP = "Could not create the XMLHttpRequest object!"
        'Release the object and exit.
        Set HttpRequest = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    'Create the request - no special parameters required.
    HttpRequest.Open "POST", "http://myip.dnsomatic.com", False
    
    'Send the request to the site.
    HttpRequest.send
        
    'Return the result of the request (the IP string).
    GetMyPublicIP = HttpRequest.responseText
    Set HttpRequest = Nothing
    Exit Function
100:
    HttpRequest = ""
    Set HttpRequest = Nothing
End Function



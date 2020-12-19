Attribute VB_Name = "Setup"
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

Private Sub InitialSetup()
'Create ini file if it doesn't exist and prepopulate it otherwise
Dim sReturn As String
If INIfavorites.FileExists(sINI_FILE) = True Then
    sReturn = INIfavorites.sManageSectionEntry(iniRead, "Version", "Installed", sINI_FILE)
    If sReturn <> strVersion Then
        'new version probably installed
        'write a new entry to the ini file and briefly display userform stating wellsrPRO was updated
        sReturn = INIfavorites.sManageSectionEntry(iniWrite, "Version", "Installed", sINI_FILE, strVersion)
'        MsgBox "wellsrPRO was successfully updated to version " & strVersion, vbInformation, "Successfully Updated"
        Load ufHelp
        ufHelp.Show vbModeless
        'now upload data to database
        Call RecordInstallation
    End If
Else
    'no ini file so make one and write the version number
    'and display userform saying thank you for installing wellsrPRO
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "Version", "Installed", sINI_FILE, strVersion)
    'MsgBox "Thank you for installing wellsrPRO version " & strVersion, vbInformation, "Thank you"
        Load ufHelp
        ufHelp.Show vbModeless
    'Now put messages that display a request to donate after X days, asks Y times then no longer asks
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "Donate", sINI_FILE, "2 Days Only: Save BIG on our VBA Reference Guides")
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateTime", sINI_FILE, DateAdd("d", 7, Date))
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateLimit", sINI_FILE, "3")
    sReturn = INIfavorites.sManageSectionEntry(iniWrite, "PRESETS", "DonateCount", sINI_FILE, "0")
    Call RecordInstallation
End If

End Sub

'https://docs.google.com/forms/d/e/1FAIpQLScQE-_NbmZgBjldK7mVr31-srMXBE02yD1rIU-TMjgbvs5vqQ/viewform?usp=pp_url&
'entry.1809480882=test
'entry.1238441535=test
'entry.2015729830=test
'entry.34657985=test
'entry.1248424676=test
'entry.1023086801=test
'entry.1330624033=test



Private Sub RecordInstallation()
'Submit installation record to Google Sheets when the version changes
On Error Resume Next
Dim HttpRequest As Object
Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
Dim strurl As String, strOS As String
Dim strVersion2 As String, strExcel As String 'return versions
Dim strFeedback As String
Dim strserial As String, strUsername As String
Dim strMAC As String, strIP As String
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
strurl = "https://docs.google.com/forms/d/e/1FAIpQLScQE-_NbmZgBjldK7mVr31-srMXBE02yD1rIU-TMjgbvs5vqQ/formResponse?ifq&" + _
    "entry.1809480882=" + CStr(strOS) + _
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
'MsgBox (httpRequest.Status & vbNewLine & httpRequest.statusText & vbNewLine & httpRequest.responseText) '& vbNewLine & httpRequest.responseXML)
'httpRequest.Status
'httpRequest.statusText
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






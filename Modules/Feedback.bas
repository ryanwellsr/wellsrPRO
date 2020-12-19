Attribute VB_Name = "Feedback"
'https://docs.google.com/forms/d/e/1FAIpQLSf5w8bnO6zqF1Dh91LucNPDFI9687nOs5J4IftLAnsTMFSn1g/viewform?usp=pp_url&entry.1809480882=test&entry.1238441535=test&entry.2015729830=test&entry.1470534562=test&entry.34657985=test&entry.1248424676=test&entry.1023086801=test&entry.1330624033=test
'entry.1809480882=test&
'entry.1238441535=test&
'entry.2015729830=test&
'entry.1470534562=test&
'entry.34657985=test&
'entry.1248424676=test&
'entry.1023086801=test&
'entry.1330624033=test


Private Sub SubmitFeedback()
'Submit Feedback to Google Sheets
Dim HttpRequest As Object
Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
Dim strurl As String, strOS As String
Dim strVersion2 As String, strExcel As String 'return versions
Dim strFeedback As String
Dim strserial As String, strUsername As String
Dim strMAC As String, strIP As String
On Error Resume Next
strserial = GetDriveSerialNumber
strUsername = Environ$("Username")
strMAC = GetMyMACAddress
strIP = GetMyPublicIP
On Error GoTo 100:
strVersion2 = ufFeedback.lblVersion.Caption
strExcel = ufFeedback.lblExcel.Caption
#If Win64 Then
   strExcel = strExcel & " (64-bit)"
#Else
   strExcel = strExcel & " (32-bit)"
#End If
strOS = ufFeedback.lblOS.Caption
strFeedback = URLEncode(ufFeedback.tbFeedback.Text)
strurl = "https://docs.google.com/forms/d/e/1FAIpQLSf5w8bnO6zqF1Dh91LucNPDFI9687nOs5J4IftLAnsTMFSn1g/formResponse?ifq&" + _
    "entry.1809480882=" + CStr(strOS) + _
    "&entry.1238441535=" + CStr(strExcel) + _
    "&entry.2015729830=" + CStr(strVersion2) + _
    "&entry.1470534562=" + CStr(strFeedback) + _
    "&entry.34657985=" + CStr(strUsername) + _
    "&entry.1248424676=" + CStr(strserial) + _
    "&entry.1023086801=" + CStr(strMAC) + _
    "&entry.1330624033=" + CStr(strIP) + _
    "&submit=Submit&ifq"
    
HttpRequest.Open "POST", strurl, False
HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
HttpRequest.send
If HttpRequest.Status = 200 Then
    MsgBox "Thank you for helping to improve wellsrPRO!", , "Success"
ElseIf HttpRequest.Status = 413 Then
    MsgBox "Your feedback text appears to be too long. Try breaking it into smaller pieces and submitting again.", , "Submission Failed"
Else
    MsgBox "We were unable to submit your feedback at this time.", , "Submission Failed"
End If
Set HttpRequest = Nothing
'MsgBox (httpRequest.Status & vbNewLine & httpRequest.statusText & vbNewLine & httpRequest.responseText) '& vbNewLine & httpRequest.responseXML)
'httpRequest.Status
'httpRequest.statusText
Exit Sub
100:
    MsgBox "We were unable to submit your feedback at this time.", , "Submission Failed"
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



Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String
'NOTE: I need to change this to a utf-8 encoded version if I'm going to get it to work with submitted macros.
On Error Resume Next
  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
  Err.Clear
On Error GoTo 0
End Function


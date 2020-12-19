Attribute VB_Name = "ShareMacros"
'#If VBA7 Then
'    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#Else
'    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#End If

Private Sub SendMacro()
'Send the user's macro to to Google Sheets
Dim HttpRequest As Object
Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
Dim strurl As String
Dim strName As String
Dim strMacro As String
Dim strMacroPieces() As String
Dim strTitle As String
Dim bErrors As Boolean
Dim i As Long
On Error GoTo 100:
strMacro = URLEncode(ufAddMacros.tbMacro.Text)
strTitle = URLEncode(ufAddMacros.tbTitle.Text)
If Trim(strTitle) = "" Or Trim(strMacro) = "" Then
    MsgBox "Your submission can't be processed because the macro field and/or title field appears blank.", vbInformation, "Submission failed"
    Exit Sub
End If
strName = InputBox("Please enter your name so other members of the wellsrPRO community will know who to thank!", "Enter your name", Environ$("Username"))
If Trim(strName) = "" Then Exit Sub
strName = URLEncode(strName)
'https://docs.google.com/forms/d/e/1FAIpQLSf45vJUiLmFYriDBfHqe8q1AuLPRektQRRYd8q_c4VykGBGRQ/viewform?usp=pp_url&
'entry.907881126&
'entry.928872493&
'entry.1826750253
strurl = "https://docs.google.com/forms/d/e/1FAIpQLSf45vJUiLmFYriDBfHqe8q1AuLPRektQRRYd8q_c4VykGBGRQ/formResponse?ifq&" + _
    "entry.907881126=" + CStr(strName) + _
    "&entry.928872493=" + CStr(strTitle) + _
    "&entry.1826750253=" + CStr(strMacro) + _
    "&submit=Submit&ifq"
    
HttpRequest.Open "POST", strurl, False
HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
HttpRequest.send
If HttpRequest.Status = 200 Then
    MsgBox "Your macro has been submitted! It will be reviewed by the editors at wellsr.com and, once approved," & _
           " it will be available for others in the wellsrPRO community to import into their spreadsheets.", , "Macro submitted"
ElseIf HttpRequest.Status = 413 Or HttpRequest.Status = 400 Then
    'macro too long. Split it into pieces of 1500 characters and try again
    strMacroPieces = SplitString(ufAddMacros.tbMacro.Text, 1500)
    For i = LBound(strMacroPieces) To UBound(strMacroPieces)
        bErrors = False
        strMacroPieces(i) = URLEncode("' PART " & i & vbNewLine & strMacroPieces(i))
TryAgain:
        strurl = "https://docs.google.com/forms/d/e/1FAIpQLSf45vJUiLmFYriDBfHqe8q1AuLPRektQRRYd8q_c4VykGBGRQ/formResponse?ifq&" + _
            "entry.907881126=" + CStr(strName) + _
            "&entry.928872493=" + CStr(strTitle) + _
            "&entry.1826750253=" + CStr(strMacroPieces(i)) + _
            "&submit=Submit&ifq"
        HttpRequest.Open "POST", strurl, False
        HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        HttpRequest.send
        If HttpRequest.Status <> 200 And bErrors = False Then
            bErrors = True
            GoTo TryAgain:
        End If
    Next i
        If HttpRequest.Status = 200 Then
            MsgBox "Your macro has been submitted! It will be reviewed by the editors at wellsr.com and, once approved," & _
                   " it will be available for others in the wellsrPRO community to import into their spreadsheets.", , "Macro submitted"
        Else
            MsgBox "We were unable to submit your macro at this time.", , "Submission Failed"
        End If
Else
    MsgBox "We were unable to submit your macro at this time.", , "Submission Failed"
End If
Set HttpRequest = Nothing
Exit Sub
100:
    MsgBox "We were unable to submit your macro at this time.", , "Submission Failed"
    Set HttpRequest = Nothing
End Sub


Public Function SplitString(ByVal str As String, ByVal numOfChar As Long) As String()
    Dim sArr() As String
    Dim nCount As Long
    ReDim sArr((Len(str) - 1) \ numOfChar)
    Do While Len(str)
        sArr(nCount) = Left$(str, numOfChar)
        str = Mid$(str, numOfChar + 1)
        nCount = nCount + 1
    Loop
    SplitString = sArr
End Function


Attribute VB_Name = "INIfavorites"
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

Enum iniAction
    iniRead = 1
    iniWrite = 2
End Enum

'*******************************************************************************
' End INI file declaration Section.
'*******************************************************************************


Function FileExists(ByVal strpath As String) As Boolean
'Function returns true if file exists, false otherwise
    If Dir(strpath) > "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub PopulateFavorites()
    Dim sReturn As String
    Dim sReturn2 As String
    Dim i As Integer
    Dim iCount As Integer
    
    iCount = -1
    ' Read the ini file - can store up to 50 favorites
    For i = 0 To 50
        sReturn = INIfavorites.sManageSectionEntry(iniRead, "Articles", CStr(i), sINI_FILE)
        sReturn2 = INIfavorites.sManageSectionEntry(iniRead, "URLs", CStr(i), sINI_FILE)
            If sReturn <> "Error" And sReturn2 <> "Error" Then
                iCount = iCount + 1
                If Trim(sReturn) <> "" And Trim(sReturn2) <> "" Then
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
            .List(.ListCount - 1, 0) = "No favorites found. Consider adding a favorite"
            .List(.ListCount - 1, 1) = "https://wellsr.com"
        End With
    End If
 End Sub
 
 
 Sub PopulateFavorites2()
    Dim sReturn As String
    Dim sReturn2 As String
    Dim i As Integer

    ' Read the ini file - can store up to 100 favorites
    For i = 0 To 50
        sReturn = INIfavorites.sManageSectionEntry(iniRead, "Articles", CStr(i), sINI_FILE)
        sReturn2 = INIfavorites.sManageSectionEntry(iniRead, "URLs", CStr(i), sINI_FILE)
            If sReturn <> "Error" And sReturn2 <> "Error" Then
                iCount = iCount + 1
                If Trim(sReturn) <> "" And Trim(sReturn2) <> "" Then
                    With ufFavorites.lbFavorites
                        .AddItem
                        .List(.ListCount - 1, 0) = sReturn
                        .List(.ListCount - 1, 1) = sReturn2
                    End With
                End If
            End If
    Next i
 End Sub








Function sManageSectionEntry(inAction As iniAction, _
                             sSection As String, _
                             sKey As String, _
                             sIniFile As String, _
                             Optional sValue As String) As String
'*******************************************************************************
' Description:  This reads an INI file section/key combination and
'               returns the read value as a string.
'
' Author:       Scott Lyerly
' Contact:      scott_lyerly@tjx.com, or scott.c.lyerly@gmail.com
'
' Notes:        Requires "Private Declare Function GetPrivateProfileString" and
'               "WritePrivateProfileString" to be added in the declarations
'               at the top of the module.
'
' Name:                 Date:           Init:   Modification:
' sManageSectionEntry   26-Nov-2013     SCL     Original development
'
' Arguments:    inAction    The action to take in teh funciton, reading or writing to
'                           to the INI file. Uses the enumeration iniAction in the
'                           declarations section.
'               sSection    The seciton of the INI file to search
'               sKey        The key of the INI from which to retrieve a value
'               sIniFile    The name and directory location of the INI file
'               sValue      The value to be written to the INI file (if writing - optional)
'
' Returns:      string      The return string is one of three things:
'                           1) The value being sought from the INI file.
'                           2) The value being written to the INI file (should match
'                              the sValue parameter).
'                           3) The word "Error". This can be changed to whatever makes
'                              the most sense to the programmer using it.
'*******************************************************************************

On Error GoTo Err_ManageSectionEntry

    ' Variable declarations.
    Dim sRetBuf         As String
Dim iLenBuf         As Integer
Dim sFileName       As String
Dim sReturnValue    As String
Dim lRetVal         As Long
    
    ' Based on the inAction parameter, take action.
    If inAction = iniRead Then  ' If reading from the INI file.

        ' Set the return buffer to by 256 spaces. This should be enough to
        ' hold the value being returned from the INI file, but if not,
        ' increase the value.
        sRetBuf = Space(256)

        ' Get the size of the return buffer.
        iLenBuf = Len(sRetBuf)

        ' Read the INI Section/Key value into the return variable.
        sReturnValue = GetPrivateProfileString(sSection, _
                                               sKey, _
                                               "", _
                                               sRetBuf, _
                                               iLenBuf, _
                                               sIniFile)

        ' Trim the excess garbage that comes through with the variable.
        sReturnValue = Trim(Left(sRetBuf, sReturnValue))

        ' If we get a value returned, pass it back as the argument.
        ' Else pass "False".
        If Len(sReturnValue) > 0 Then
            sManageSectionEntry = sReturnValue
        Else
            sManageSectionEntry = "Error"
        End If
ElseIf inAction = iniWrite Then ' If writing to the INI file.

        ' Check to see if a value was passed in the sValue parameter.
        If Len(sValue) = 0 Then
            sManageSectionEntry = "Error"

        Else
            
            ' Write to the INI file and capture the value returned
            ' in the API function.
            lRetVal = WritePrivateProfileString(sSection, _
                                               sKey, _
                                               sValue, _
                                               sIniFile)

            ' Check to see if we had an error wrting to the INI file.
            If lRetVal = 0 Then sManageSectionEntry = "Error"

        End If
End If
    
Exit_Clean:
    Exit Function
    
Err_ManageSectionEntry:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Clean

End Function


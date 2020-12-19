'*************************************************************************************************
'** You cannot run update.vbs manually. Attempting to do so will generate an error. This script **
'**     is an internal script used by wellsrPRO to check for updates. You can ignore it. :-)    **
'*************************************************************************************************
'-----------------------------------------------
'Automatically update the wellsrPRO Excel Add-in
'-----------------------------------------------
SourceFile =WScript.Arguments.Item(0)
DestinationFile=WScript.Arguments.Item(1)
DestinationFolder=WScript.Arguments.Item(2)

'Close existing workbook
Set objXl = GetObject(, "Excel.Application")
on Error Resume Next
objXL.Workbooks("wellsrPRO.xlam").Close(False)

'Install new xlam
Set fso = CreateObject("Scripting.FileSystemObject")
    'Check to see if the file already exists in the destination folder
    If fso.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If Not fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is not read-only.  Safe to replace the file.
            fso.CopyFile SourceFile, DestinationFolder, True
        Else 
            'The file exists and is read-only.
            'Remove the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes - 1
            'Replace the file
            fso.CopyFile SourceFile, DestinationFolder, True
            'Reapply the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
        End If
    Else
        'The file does not exist in the destination folder.  Safe to copy file to this folder.
        fso.CopyFile SourceFile, DestinationFolder, True
    End If
Set fso = Nothing

objXL.Workbooks.Open(DestinationFile)

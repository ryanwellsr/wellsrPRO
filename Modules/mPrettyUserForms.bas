Attribute VB_Name = "mPrettyUserforms"
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong _
                           Lib "User32" Alias "GetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong _
                           Lib "User32" Alias "SetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar _
                           Lib "User32" ( _
                           ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function FindWindowA _
                           Lib "User32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As Long
#Else
    Private Declare Function GetWindowLong _
                           Lib "User32" Alias "GetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong _
                           Lib "User32" Alias "SetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar _
                           Lib "User32" ( _
                           ByVal hWnd As Long) As Long
    Private Declare Function FindWindowA _
                           Lib "User32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As Long
#End If
Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub


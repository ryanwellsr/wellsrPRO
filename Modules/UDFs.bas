Attribute VB_Name = "UDFs"
Function compare(ByVal cell1 As Range, ByVal Cell2 As Range, Optional CaseSensitive As Variant, Optional delta As Variant, Optional MatchString As Variant)
Attribute compare.VB_Description = "Compares Cell1 to Cell2 and if identical returns ""-"" by default but a different optional match string can be given. If cells are different, the output will either be ""FALSE"" or will optionally show the delta between the values if numeric."
Attribute compare.VB_ProcData.VB_Invoke_Func = " \n20"
'******************************************************************************
'***DEVELOPER:   Ryan Wells (wellsr.com)                                      *
'***DATE:        04/2016                                                      *
'***DESCRIPTION: Compares Cell1 to Cell2 and if identical, returns "-" by     *
'***             default but a different optional match string can be given.  *
'***             If cells are different, the output will either be "FALSE"    *
'***             or will optionally show the delta between the values if      *
'***             numeric.                                                     *
'***INPUT:       Cell1 - First cell to compare.                               *
'***             Cell2 - Cell to compare against Cell1.                       *
'***             CaseSensitive - Optional boolean that if set to TRUE, will   *
'***                             perform a case-sensitive comparison of the   *
'***                             two entered cells. Default is TRUE.          *
'***             delta - Optional boolean that if set to TRUE, will display   *
'***                     the delta between Cell1 and Cell2.                   *
'***             MatchString - Optional string the user can choose to display *
'***                           when Cell1 and Cell2 match. Default is "-"     *
'***OUTPUT:      The output will be "-", a custom string or a delta if the    *
'***             cells match and will be "FALSE" if the cells do not match.   *
'***EXAMPLES:     =compare(A1,B1,FALSE,TRUE,"match")                          *
'***              =compare(A1,B1)                                             *
'******************************************************************************
 
'------------------------------------------------------------------------------
'I. Declare variables
'------------------------------------------------------------------------------
Dim strMatch As String 'string to display if Cell1 and Cell2 match
 
'------------------------------------------------------------------------------
'II. Error checking
'------------------------------------------------------------------------------
'Error 0 - catch all error
On Error GoTo CompareError:
 
'Error 1 - MatchString is invalid
If IsMissing(MatchString) = False Then
    If IsError(CStr(MatchString)) Then
        compare = "Invalid Match String"
        Exit Function
    End If
End If
 
'Error 2 - Cell1 contains more than 1 cell
If IsArray(cell1) = True Then
    If cell1.count <> 1 Then
        compare = "Too many cells in variable Cell1."
        Exit Function
    End If
End If
 
'Error 3 - Cell2 contains more than 1 cell
If IsArray(Cell2) = True Then
    If Cell2.count <> 1 Then
        compare = "Too many cells in variable Cell2."
        Exit Function
    End If
End If
 
'Error 4 - delta is not a boolean
If IsMissing(delta) = False Then
    If delta <> CBool(True) And delta <> CBool(False) Then
        compare = "Delta flag must be a boolean (TRUE or FALSE)."
        Exit Function
    End If
End If
 
'Error 5 - CaseSensitive is not a boolean
If IsMissing(CaseSensitive) = False Then
    If CaseSensitive <> CBool(True) And CaseSensitive <> CBool(False) Then
        compare = "CaseSensitive flag must be a boolean (TRUE or FALSE)."
        Exit Function
    End If
End If

'------------------------------------------------------------------------------
'III. Initialize Variables
'------------------------------------------------------------------------------
If IsMissing(CaseSensitive) Then
    CaseSensitive = CBool(True)
ElseIf CaseSensitive = False Then
    CaseSensitive = CBool(False)
Else
    CaseSensitive = CBool(True)
End If

If IsMissing(MatchString) Then
    strMatch = "-"
Else
    strMatch = CStr(MatchString)
End If
 
If IsMissing(delta) Then
    delta = CBool(False)
ElseIf delta = False Then
    delta = CBool(False)
Else
    delta = CBool(True)
End If
 
'------------------------------------------------------------------------------
'IV. Check for matches
'------------------------------------------------------------------------------
If cell1 = Cell2 Then
    compare = strMatch
ElseIf CaseSensitive = False Then
    If UCase(cell1) = UCase(Cell2) Then
        compare = strMatch
    ElseIf delta = True And IsNumeric(cell1) And IsNumeric(Cell2) Then
        compare = cell1 - Cell2
    Else
        compare = CBool(False)
    End If
ElseIf cell1 <> Cell2 And delta = True Then
    If IsNumeric(cell1) And IsNumeric(Cell2) Then
        'No case sensitive check because if not numeric, doesn't matter.
        compare = cell1 - Cell2
    Else
        compare = CBool(False)
    End If
Else
    compare = CBool(False)
End If
Exit Function
 
'------------------------------------------------------------------------------
'V. Final Error Handling
'------------------------------------------------------------------------------
CompareError:
    compare = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function


Function Linterp(ByVal KnownYs As Range, ByVal KnownXs As Range, NewX As Variant) As Variant
Attribute Linterp.VB_Description = "2D Linear Interpolation function that automatically picks which range to interpolate between based on the closest KnownX value to the NewX value you want to interpolate for."
Attribute Linterp.VB_ProcData.VB_Invoke_Func = " \n20"
'******************************************************************************
'***DEVELOPER: Ryan Wells (wellsr.com) *
'***DATE: 03/2016 *
'***DESCRIPTION: 2D Linear Interpolation function that automatically picks *
'*** which range to interpolate between based on the closest *
'*** KnownX value to the NewX value you want to interpolate for. *
'***INPUT: KnownYs - 1D range containing your known Y values. *
'*** KnownXs - 1D range containing your known X values. *
'*** NewX - Cell or number with the X value you want to *
'*** interpolate for. *
'***OUTPUT: The output will be the linear interpolated Y value *
'*** corresponding to the NewX value the user selects. *
'***NOTES: i. KnownYs do not have to be sorted. If the values are *
'*** unsorted, the function will linearly interpolate between the *
'*** two closest values to your NewX (one above, one below). *
'*** ii. KnownXs and KnownYs must be the same dimensions. It is a *
'*** good practice to have the Xs and corresponding Ys beside *
'*** each other in Excel before using Linterp. *
'***FORMULA: Linterp=Y0 + (Y1-Y0)*(NewX-X0)/(X1-X0) *
'***EXAMPLE: =Linterp(A2:A4,B2:B4,C2) *
'******************************************************************************
 
'------------------------------------------------------------------------------
'0. Declare Variables and Initialize Variables
'------------------------------------------------------------------------------
Dim bYRows As Boolean   'Y values are selected in a row (Nx1)
Dim bXRows As Boolean   'X values are selected in a row (Nx1)
Dim DeltaHi As Double   'delta between NewX and KnownXs if Known > NewX
Dim DeltaLo As Double   'delta between NewX and KnownXs if Known < NewX
Dim iHi As Long         'Index position of the closest value above NewX
Dim iLo As Long         'Index position of the closest value below NewX
Dim i As Long           'dummy counter
Dim Y0 As Double, Y1 As Double 'Linear Interpolation Y variables
Dim X0 As Double, X1 As Double 'Linear Interpolation Y variables
iHi = 2147483647
iLo = -2147483648#
DeltaHi = 1.79769313486231E+308
DeltaLo = -1.79769313486231E+308
 
'------------------------------------------------------------------------------
'I. Preliminary Error Checking
'------------------------------------------------------------------------------
'Error 0 - catch all error
On Error GoTo InterpError:
'Error 1 - NewX more than 1 cell selected
If IsArray(NewX) = True Then
    If NewX.count <> 1 Then
        Linterp = "Too many cells in variable NewX."
        Exit Function
    End If
End If
 
'Error 2 - NewX is not a number
If IsNumeric(NewX) = False Then
    Linterp = "NewX is non-numeric."
    Exit Function
End If
 
'Error 3 - dimensions aren't even
If KnownYs.count <> KnownXs.count Or _
   KnownYs.Rows.count <> KnownXs.Rows.count Or _
   KnownYs.Columns.count <> KnownXs.Columns.count Then
    Linterp = "Known ranges are different dimensions."
    Exit Function
End If
 
'Error 4 - known Ys are not Nx1 or 1xN dimensions
If KnownYs.Rows.count <> 1 And KnownYs.Columns.count <> 1 Then
    Linterp = "Known Y's should be in a single column or a single row."
    Exit Function
End If
 
'Error 5 - known Xs are not Nx1 or 1xN dimensions
If KnownXs.Rows.count <> 1 And KnownXs.Columns.count <> 1 Then
    Linterp = "Known X's should be in a single column or a single row."
    Exit Function
End If
 
'Error 6 - Too few known Y cells
If KnownYs.Rows.count <= 1 And KnownYs.Columns.count <= 1 Then
    Linterp = "Known Y's range must be larger than 1 cell"
    Exit Function
End If
 
'Error 7 - Too few known X cells
If KnownXs.Rows.count <= 1 And KnownXs.Columns.count <= 1 Then
    Linterp = "Known X's range must be larger than 1 cell"
    Exit Function
End If
 
'Error 8 - Check for non-numeric KnownYs
If KnownYs.Rows.count > 1 Then
    bYRows = True
    For i = 1 To KnownYs.Rows.count
        If IsNumeric(KnownYs.Cells(i, 1)) = False Then
            Linterp = "One or all Known Y's are non-numeric."
            Exit Function
        End If
    Next i
ElseIf KnownYs.Columns.count > 1 Then
    bYRows = False
    For i = 1 To KnownYs.Columns.count
        If IsNumeric(KnownYs.Cells(1, i)) = False Then
            Linterp = "One or all KnownYs are non-numeric."
            Exit Function
        End If
    Next i
End If
 
'Error 9 - Check for non-numeric KnownXs
If KnownXs.Rows.count > 1 Then
    bXRows = True
    For i = 1 To KnownXs.Rows.count
        If IsNumeric(KnownXs.Cells(i, 1)) = False Then
            Linterp = "One or all Known X's are non-numeric."
            Exit Function
        End If
    Next i
ElseIf KnownXs.Columns.count > 1 Then
    bXRows = False
    For i = 1 To KnownXs.Columns.count
        If IsNumeric(KnownXs.Cells(1, i)) = False Then
            Linterp = "One or all Known X's are non-numeric."
            Exit Function
        End If
    Next i
End If
 
'------------------------------------------------------------------------------
'II. Check for nearest values from list of Known X's
'------------------------------------------------------------------------------
If bXRows = True Then 'check by rows
    For i = 1 To KnownXs.Rows.count 'loop through known Xs
        If KnownXs.Cells(i, 1) <> "" Then
            If KnownXs.Cells(i, 1) > NewX And KnownXs.Cells(i, 1) - NewX < DeltaHi Then 'determine DeltaHi
                DeltaHi = KnownXs.Cells(i, 1) - NewX
                iHi = i
            ElseIf KnownXs.Cells(i, 1) < NewX And KnownXs.Cells(i, 1) - NewX > DeltaLo Then 'determine DeltaLo
                DeltaLo = KnownXs.Cells(i, 1) - NewX
                iLo = i
            ElseIf KnownXs.Cells(i, 1) = NewX Then 'match. just report corresponding Y
                Linterp = KnownYs.Cells(i, 1)
                Exit Function
            End If
        End If
    Next i
Else ' check by columns
    For i = 1 To KnownXs.Columns.count 'loop through known Xs
        If KnownXs.Cells(1, i) <> "" Then
            If KnownXs.Cells(1, i) > NewX And KnownXs.Cells(1, i) - NewX < DeltaHi Then 'determine DeltaHi
                DeltaHi = KnownXs.Cells(1, i) - NewX
                iHi = i
            ElseIf KnownXs.Cells(1, i) < NewX And KnownXs.Cells(1, i) - NewX > DeltaLo Then 'determine DeltaLo
                DeltaLo = KnownXs.Cells(1, i) - NewX
                iLo = i
            ElseIf KnownXs.Cells(1, i) = NewX Then 'match. just report corresponding Y
                Linterp = KnownYs.Cells(1, i)
                Exit Function
            End If
        End If
    Next i
End If
 
'------------------------------------------------------------------------------
'III. Linear interpolate based on the closest cells in the range. Includes minor error handling
'------------------------------------------------------------------------------
If iHi = 2147483647 Or iLo = -2147483648# Then
    Linterp = "NewX is out of range. Cannot linearly interpolate with the given Knowns."
    Exit Function
End If
If bXRows = True Then
    Y0 = KnownYs.Cells(iLo, 1)
    Y1 = KnownYs.Cells(iHi, 1)
    X0 = KnownXs.Cells(iLo, 1)
    X1 = KnownXs.Cells(iHi, 1)
Else
    Y0 = KnownYs.Cells(1, iLo)
    Y1 = KnownYs.Cells(1, iHi)
    X0 = KnownXs.Cells(1, iLo)
    X1 = KnownXs.Cells(1, iHi)
End If
Linterp = Y0 + (Y1 - Y0) * (NewX - X0) / (X1 - X0)
Exit Function
 
'------------------------------------------------------------------------------
'IV. Final Error Handling
'------------------------------------------------------------------------------
InterpError:
    Linterp = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function

Function Locate(ByVal LookFor As Variant, ByVal InRange As Range, Optional MatchString As String, Optional CaseSensitive As Variant) As Variant
Attribute Locate.VB_Description = "Determines if a given value is located somewhere in a range."
Attribute Locate.VB_ProcData.VB_Invoke_Func = " \n20"
'******************************************************************************
'***DEVELOPER:   Ryan Wells (wellsr.com)                                      *
'***DATE:        04/2016                                                      *
'***DESCRIPTION: Function determines if a given value is present in a range.  *
'***INPUT:       LookFor - Cell or string you want to look for in the range.  *
'***             InRange - Range you want to search to find the value given   *
'***                       in the variable LookFor.                           *
'***             CaseSensitive - Optional boolean that if set to TRUE, will   *
'***                             perform a case-sensitive search for the      *
'***                             variable LookFor in range InRange. Default   *
'***                             is FALSE.                                    *
'***             MatchString - Optional string the user can choose to display *
'***                           when LookFor is in InRange. Default is to      *
'***                           display the location of the found value.       *
'***OUTPUT:      By default, the output will be the address (cell) where the  *
'***             value LookFor was found in the range InRange. If the value   *
'***             is not found, a message saying such will display.            *
'***EXAMPLES:     =locate(V3,A2:P80)                                          *
'***              =locate("Smith",A2:P80,"-",TRUE)                            *
'******************************************************************************
 
'------------------------------------------------------------------------------
'I. Declare variables
'------------------------------------------------------------------------------
Dim i As Long, j As Long 'counters
Dim bMatch As Boolean
 
'------------------------------------------------------------------------------
'II. Preliminary Error Checking
'------------------------------------------------------------------------------
'Error 0 - catch all error
On Error GoTo LocateError:
 
'Error 1 - LookFor has too many cells
If IsArray(LookFor) = True Then
    If LookFor.count <> 1 Then
        Locate = "Too many cells in variable LookFor."
        Exit Function
    End If
End If
 
'Error 2 - CaseSensitive is not a boolean
If IsMissing(CaseSensitive) = False Then
    If CaseSensitive <> CBool(True) And CaseSensitive <> CBool(False) Then
        Locate = "CaseSensitive flag must be a boolean (TRUE or FALSE)."
        Exit Function
    End If
End If
 
'Error 3 - MatchString is invalid
If IsMissing(MatchString) = False Then
    If IsError(CStr(MatchString)) Then
        Locate = "Invalid Match String"
        Exit Function
    End If
End If
 
'Error 4 - LookFor is not a valid string
If IsError(LookFor) = True Then
    Locate = "Error in variable LookFor"
    Exit Function
End If
 
'Error 5 - InRange is not valid
If IsError(InRange) = True Then
    Locate = "Error in variable InRange"
    Exit Function
End If
 
'------------------------------------------------------------------------------
'III. Initialize Variables
'------------------------------------------------------------------------------
If IsMissing(CaseSensitive) Then
    CaseSensitive = CBool(False)
ElseIf CaseSensitive = False Then
    CaseSensitive = CBool(False)
Else
    CaseSensitive = CBool(True)
End If
 
'------------------------------------------------------------------------------
'IV. Look for the variable "LookFor" in range "InRange"
'------------------------------------------------------------------------------
For i = 1 To InRange.Rows.count
    For j = 1 To InRange.Columns.count
        If CaseSensitive = CBool(True) Then
            If LookFor = InRange(i, j) Then
                If IsMissing(MatchString) = True Or MatchString = Empty Then
                    Locate = "=" & InRange(i, j).Address
               Else
                    Locate = MatchString
                End If
                bMatch = True
                Exit Function
            End If
        Else
            If UCase(CStr(LookFor)) = UCase(CStr(InRange(i, j))) Then
                If IsMissing(MatchString) = True Or MatchString = Empty Then
                    Locate = "=" & InRange(i, j).Address
                Else
                    Locate = MatchString
                End If
                bMatch = True
                Exit Function
            End If
        End If
    Next j
Next i
 
If bMatch = False Then
    Locate = "Cannot find """ & LookFor & """ in range " & InRange.Address
End If
Exit Function
 
'------------------------------------------------------------------------------
'V. Final Error Handling
'------------------------------------------------------------------------------
LocateError:
    Locate = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function




Function TrapIntegration(KnownXs As Variant, KnownYs As Variant) As Variant
Attribute TrapIntegration.VB_Description = "Approximates the integral (the area under the curve) of a data set by using the trapezoidal rule."
Attribute TrapIntegration.VB_ProcData.VB_Invoke_Func = " \n20"
'------------------------------------------------------------------------------------------------------
'---DESCRIPTION: Approximates the integral using trapezoidal rule.-------------------------------------
'---CREATED BY: Ryan Wells-----------------------------------------------------------------------------
'---INPUT: KnownXs is the range of x-values. KnownYs is the range of y-values.-------------------------
'---OUTPUT: The output will be the approximate area under the curve (integral).------------------------
'------------------------------------------------------------------------------------------------------

    Dim i As Integer
    Dim bYRows As Boolean, bXRows As Boolean

'------------------------------------------------------------------------------
'I. Preliminary Error Checking
'------------------------------------------------------------------------------
On Error GoTo TrapIntError:
    'Error 1 - Check if the X values are range.
    If Not TypeName(KnownXs) = "Range" Then
        TrapIntegration = "Invalid X-range"
        Exit Function
    End If

    'Error 2 - Check if the Y values are range.
    If Not TypeName(KnownYs) = "Range" Then
        TrapIntegration = "Invalid Y-range"
        Exit Function
    End If


    'Error 3 - dimensions aren't even
    If KnownYs.count <> KnownXs.count Or _
       KnownYs.Rows.count <> KnownXs.Rows.count Or _
       KnownYs.Columns.count <> KnownXs.Columns.count Then
        TrapIntegration = "Known ranges are different dimensions."
        Exit Function
    End If
     
    'Error 4 - known Ys are not Nx1 or 1xN dimensions
    If KnownYs.Rows.count <> 1 And KnownYs.Columns.count <> 1 Then
        TrapIntegration = "Known Y's should be in a single column or a single row."
        Exit Function
    End If
     
    'Error 5 - known Xs are not Nx1 or 1xN dimensions
    If KnownXs.Rows.count <> 1 And KnownXs.Columns.count <> 1 Then
        TrapIntegration = "Known X's should be in a single column or a single row."
        Exit Function
    End If
     
    'Error 6 - Check for non-numeric KnownYs
    If KnownYs.Rows.count > 1 Then
        bYRows = True
        For i = 1 To KnownYs.Rows.count
            If IsNumeric(KnownYs.Cells(i, 1)) = False Then
                TrapIntegration = "One or all Known Y's are non-numeric."
                Exit Function
            End If
        Next i
    ElseIf KnownYs.Columns.count > 1 Then
        bYRows = False
        For i = 1 To KnownYs.Columns.count
            If IsNumeric(KnownYs.Cells(1, i)) = False Then
                TrapIntegration = "One or all KnownYs are non-numeric."
                Exit Function
            End If
        Next i
    End If
     
    'Error 7 - Check for non-numeric KnownXs
    If KnownXs.Rows.count > 1 Then
        bXRows = True
        For i = 1 To KnownXs.Rows.count
            If IsNumeric(KnownXs.Cells(i, 1)) = False Then
                TrapIntegration = "One or all Known X's are non-numeric."
                Exit Function
            End If
        Next i
    ElseIf KnownXs.Columns.count > 1 Then
        bXRows = False
        For i = 1 To KnownXs.Columns.count
            If IsNumeric(KnownXs.Cells(1, i)) = False Then
                TrapIntegration = "One or all Known X's are non-numeric."
                Exit Function
            End If
        Next i
    End If

'------------------------------------------------------------------------------
'II. Perform Trapezoidal Integration
'------------------------------------------------------------------------------
    TrapIntegration = 0

    'Apply the trapezoid rule: (y(i+1) + y(i))*(x(i+1) - x(i))*1/2.
    'Use the absolute value in case of negative numbers.
    If bXRows = True Then
        For i = 1 To KnownXs.Rows.count - 1
            TrapIntegration = TrapIntegration + Abs(0.5 * (KnownXs.Cells(i + 1, 1) _
            - KnownXs.Cells(i, 1)) * (KnownYs.Cells(i, 1) + KnownYs.Cells(i + 1, 1)))
        Next i
    Else
        For i = 1 To KnownXs.Columns.count - 1
            TrapIntegration = TrapIntegration + Abs(0.5 * (KnownXs.Cells(1, i + 1) _
            - KnownXs.Cells(1, i)) * (KnownYs.Cells(1, i) + KnownYs.Cells(1, i + 1)))
        Next i
    End If
Exit Function

TrapIntError:
    TrapIntegration = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function


Function RGBn(r As Integer, g As Integer, b As Integer) As Variant
Attribute RGBn.VB_Description = "Convert colors defined with RGB to a long data type. This can be paired with the DrawArrows function to draw arrows of different colors."
Attribute RGBn.VB_ProcData.VB_Invoke_Func = " \n20"
'------------------------------------------------------------------------------------------------
'---Script: RGBn---------------------------------------------------------------------------------
'---Created by: Ryan Wells ----------------------------------------------------------------------
'---Date: 4/2016 --------------------------------------------------------------------------------
'---Description: This function converts 3 integers to a long data type and can be paired --------
'---             with DrawArrows to draw arrows of a different color.- --------------------------
'------------------------------------------------------------------------------------------------

'Convert RGB to LONG:
If r > 255 Or g > 255 Or b > 255 Or r < 0 Or g < 0 Or b < 0 Then
    RGBn = "Arguments should be >=0 and <=255"
    Exit Function
End If

 RGBn = RGB(r, g, b)
End Function


Function DrawArrows(FromRange As Range, ToRange As Range, Optional RGBcolor As Variant, Optional LineType As String) As Variant
Attribute DrawArrows.VB_Description = "This function draws arrows or lines from the middle of one cell to the middle of another. Custom endpoints and shape colors are suppported."
Attribute DrawArrows.VB_ProcData.VB_Invoke_Func = " \n20"
'---------------------------------------------------------------------------------------------------
'---Script: DrawArrows------------------------------------------------------------------------------
'---Created by: Ryan Wells -------------------------------------------------------------------------
'---Date: 10/2015-----------------------------------------------------------------------------------
'---Description: This function draws arrows or lines from the middle of one cell to the middle -----
'----------------of another. Custom endpoints and shape colors are suppported ----------------------
'---------------------------------------------------------------------------------------------------

Dim dleft1 As Double, dleft2 As Double
Dim dtop1 As Double, dtop2 As Double
Dim dheight1 As Double, dheight2 As Double
Dim dwidth1 As Double, dwidth2 As Double
dleft1 = FromRange.Left
dleft2 = ToRange.Left
dtop1 = FromRange.Top
dtop2 = ToRange.Top
dheight1 = FromRange.Height
dheight2 = ToRange.Height
dwidth1 = FromRange.Width
dwidth2 = ToRange.Width
 
ActiveSheet.Shapes.AddConnector(msoConnectorStraight, dleft1 + dwidth1 / 2, dtop1 + dheight1 / 2, dleft2 + dwidth2 / 2, dtop2 + dheight2 / 2).Select
'format line
With Selection.ShapeRange.Line
    .BeginArrowheadStyle = msoArrowheadNone
    .EndArrowheadStyle = msoArrowheadOpen
    .Weight = 1.75
    .Transparency = 0.5
    If UCase(LineType) = "DOUBLE" Then 'double arrows
        .BeginArrowheadStyle = msoArrowheadOpen
    ElseIf UCase(LineType) = "LINE" Then 'Line (no arows)
        .EndArrowheadStyle = msoArrowheadNone
    Else 'single arrow
        'defaults to an arrow with one head
    End If
    'color arrow
    If IsMissing(RGBcolor) = False Then
        .ForeColor.RGB = CLng(RGBcolor)     'custom color
    Else
        .ForeColor.RGB = RGB(228, 108, 10)   'orange (DEFAULT)
    End If
End With
 
DrawArrows = ""
End Function

Function ArrayMatch(ByVal xy As Range, ByVal array1 As Range, array2 As Range) As Variant
Attribute ArrayMatch.VB_Description = "Return the value in one range based on the relative position of a value in another range."
Attribute ArrayMatch.VB_ProcData.VB_Invoke_Func = " \n20"
'******************************************************************************
'***DEVELOPER:   Ryan Wells (wellsr.com)                                      *
'***DATE:        07/2016                                                      *
'***DESCRIPTION: Return the value in a different range based on the relative  *
'***             position of a value in one range.                            *
'***INPUT:       xy - A single cell within range array1.                      *
'***             array1 - The array where xy exists within.                   *
'***             array2 - The array you want to look up to find the value in  *
'***                      the cell corresponding to the position of xy inside *
'***                      array1. The size of array1 and array2 must match.   *
'***OUTPUT:      The output will be the value in array2 that is in the same   *
'***             position as xy is within array1. This is a simpler form of   *
'***             one of the dozens of index/match/lookup Excel formula        *
'***             combinations.                                                *
'***EXAMPLE:     =ArrayMatch(D10,B9:E11,B3:E5)                                *
'******************************************************************************
 
 
'------------------------------------------------------------------------------
'I. Preliminary Error Checking
'------------------------------------------------------------------------------
'Error 0 - catch all error
On Error GoTo ArrayMatchError:
'Error 1 - xy more than 1 cell selected
If IsArray(xy) = True Then
    If xy.count <> 1 Then
        ArrayMatch = "Too many cells in variable xy (argument 1). xy should be 1 cell."
        Exit Function
    End If
End If
 
'Error 2 - array dimensions aren't even
If array1.count <> array2.count Or _
   array1.Rows.count <> array2.Rows.count Or _
   array1.Columns.count <> array2.Columns.count Then
    ArrayMatch = "Lookup arrays are different dimensions (arguments 2 3)."
    Exit Function
End If
 
'Error 3 - Cell xy is not positioned inside array1.
If Application.Intersect(xy, array1) Is Nothing Then
    ArrayMatch = "Cell xy (argument 1) does not reside inside array1 (argument 2)."
    Exit Function
End If
 
'------------------------------------------------------------------------------
'II. Return Position Inside Other Range
'------------------------------------------------------------------------------
If Not Intersect(array1, xy) Is Nothing Then
    ArrayMatch = array2.Cells(Range(array1(1), xy).Rows.count, _
                              Range(array1(1), xy).Columns.count)
End If
 
Exit Function
'------------------------------------------------------------------------------
'III. Final Error Handling
'------------------------------------------------------------------------------
ArrayMatchError:
    ArrayMatch = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function

Function TimeToWords(rng1 As Variant, Optional units As String)
Attribute TimeToWords.VB_Description = "Converts decimal times to words. Function is capable of converting hours, minutes and seconds from decimal format (10.5) to words (""10 Minutes 30 Seconds"")."
Attribute TimeToWords.VB_ProcData.VB_Invoke_Func = " \n20"
'*******************************************************************************
'DEVELOPER: Ryan Wells (wellsr.com)                                            *
'DATE: 08/2016                                                                 *
'DESCRIPTION: Converts decimal times to words.                                 *
'             Function is capable of converting hours, minutes and seconds from*
'             decimal format (10.5) to words ("10 Minutes 30 Seconds").        *
'EXAMPLES: =TimeToWords(A1) to convert from minutes (default)                  *
'          =TimeToWords(A1,"h") to convert from hours                          *
'          =TimeToWords(A1,"s") to convert from seconds                        *
'          =TimeToWords(7.5) converts to "7 Minutes 30 Seconds"                *
'*******************************************************************************

Dim dSec As Double, dMin As Double, dHrs As Double
Dim strHrs As String, strMin As String, strSec As String
Dim dBase As Double
If IsNumeric(rng1) = False Then GoTo 101:

'Check if base unit is hours, minutes or seconds. Default is minutes
units = UCase(units)
dBase = rng1
If units = "H" Or units = "HOUR" Or units = "HR" Or units = "HOURS" Or units = "HRS" Then
    dMin = Application.WorksheetFunction.RoundDown(dBase * 60, 0)
    dSec = dBase * 60 - dMin
ElseIf units = "S" Or units = "SECOND" Or units = "SEC" Or units = "SECONDS" Or units = "SECS" Then
    dMin = Application.WorksheetFunction.RoundDown(dBase / 60, 0)
    dSec = dBase / 60 - dMin
ElseIf units = "M" Or units = "MINUTE" Or units = "MIN" Or units = "MINUTES" Or units = "MINS" Or units = "" Then
    dMin = Application.WorksheetFunction.RoundDown(dBase, 0)
    dSec = dBase - dMin
Else
    TimeToWords = "Invalid units entered"
    Exit Function
End If

'calculate hours, minutes, and seconds based on base unit provided
dSec = Round(dSec * 60, 0)
If dMin >= 60 Then
    dHrs = dMin / 60
    dMin = (dHrs - Application.WorksheetFunction.RoundDown(dHrs, 0)) * 60
    dHrs = Application.WorksheetFunction.RoundDown(dHrs, 0)
    dMin = Round(dMin, 0)
End If
 
'handle plural vs not
If dHrs = 1 Then
    strHrs = dHrs & " Hour "
ElseIf dHrs = 0 Then
    strHrs = ""
Else
    strHrs = dHrs & " Hours "
End If
 
If dSec = 1 Then
    strSec = dSec & " Second "
ElseIf dSec = 0 Then
    strSec = ""
Else
    strSec = dSec & " Seconds "
End If
 
If dMin = 1 Then
    strMin = dMin & " Minute "
ElseIf dMin = 0 Then
    strMin = ""
Else
    strMin = dMin & " Minutes "
End If
 
'final results
TimeToWords = strHrs & strMin & strSec
Exit Function

'handle errors
101:
TimeToWords = "Non-numeric value entered as an argument"
End Function






Function SuperMid(ByVal strMain As String, str1 As String, str2 As String, Optional reverse As Boolean) As String
Attribute SuperMid.VB_Description = "Easily extract text between two strings with this VBA Function. This function can extract a substring between two characters, delimiters, words and more. The delimiters can be the same delimiter or unique strings."
Attribute SuperMid.VB_ProcData.VB_Invoke_Func = " \n20"
'DESCRIPTION: Extract the portion of a string between the two substrings defined in str1 and str2.
'DEVELOPER: Ryan Wells (wellsr.com)
'HOW TO USE: - Pass the argument your main string and the 2 strings you want to find in the main string.
' - This function will extract the values between the end of your first string and the beginning
' of your next string.
' - If the optional boolean "reverse" is true, an InStrRev search will occur to find the last
' instance of the substrings in your main string.
Dim i As Integer, j As Integer, temp As Variant
On Error GoTo errhandler:
If reverse = True Then
    i = InStrRev(strMain, str1)
    j = InStrRev(strMain, str2)
    If Abs(j - i) < Len(str1) Then j = InStrRev(strMain, str2, i)
    If i = j Then 'try to search 2nd half of string for unique match
        j = InStrRev(strMain, str2, i - 1)
    End If
Else
    i = InStr(1, strMain, str1)
    j = InStr(1, strMain, str2)
    If Abs(j - i) < Len(str1) Then j = InStr(i + Len(str1), strMain, str2)
    If i = j Then 'try to search 2nd half of string for unique match
        j = InStr(i + 1, strMain, str2)
    End If
End If
If i = 0 And j = 0 Then GoTo errhandler:
If j = 0 Then j = Len(strMain) + Len(str2) 'just to make it arbitrarily large
If i = 0 Then i = Len(strMain) + Len(str1) 'just to make it arbitrarily large
If i > j And j <> 0 Then 'swap order
    temp = j
    j = i
    i = temp
    temp = str2
    str2 = str1
    str1 = temp
End If
i = i + Len(str1)
SuperMid = Mid(strMain, i, j - i)
Exit Function
errhandler:
MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & "Aborting", , "Strings not found"
End
End Function



Function ReplaceN(ByVal str1 As Variant, strFind As String, strReplace As String, N As Long, Optional count As Long) As String
Attribute ReplaceN.VB_Description = "Replace Nth Occurrence of substring in a string."
Attribute ReplaceN.VB_ProcData.VB_Invoke_Func = " \n20"
Dim i As Long, j As Long
Dim strM As String
strM = str1
If count <= 0 Then count = 1
For i = 1 To N - 1
    j = InStr(1, strM, strFind)
    strM = Mid(strM, j + Len(strFind), Len(strM))
Next i
If N <= 0 Then
    ReplaceN = str1
Else
    ReplaceN = Mid(str1, 1, Len(str1) - Len(strM)) & Replace(strM, strFind, strReplace, Start:=1, count:=count)
End If
End Function



Function SayIt(cell1 As Variant, Optional bAsync As Boolean, Optional bPurge As Boolean)
Attribute SayIt.VB_Description = "Function that speaks the content of whatever cell you want."
Attribute SayIt.VB_ProcData.VB_Invoke_Func = " \n20"
Application.Speech.Speak (cell1), bAsync, , bPurge
SayIt = "Speaking"
End Function

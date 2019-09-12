Attribute VB_Name = "MESHVSTU"
'File:      MESHUTIL.BAS

'
'
'Version:   1.01
'Date:      5.Nov.1997
'Author:    Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
    
    Type curve
        n As Integer
        X(1 To 100) As Double
        Y(1 To 100) As Double
    End Type

    Public MainForm As meshdraw

   'Global constants for use with PR_SetTextData
   '
    Global Const HELVETICA = 10
    Global Const BLOCK = 0
    Global Const LEFT_ = 1
    Global Const RIGHT_ = 4
    Global Const HORIZ_CENTER = 2
    Global Const TOP_ = 8
    Global Const BOTTOM_ = 32
    Global Const VERT_CENTER = 16
    Global Const CURRENT = -1

'XY data type to represent points
    Type XY
        X As Double
        Y As Double
    End Type
    
'PI as a Global Constant
    Global Const PI = 3.141592654

'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

'Store current layer and text setings to reduce DRAFIX code
'this value is checked in PR_SetLayer
    Global g_sCurrentLayer As String
    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont
    Global g_nCurrTextAngle

    Global g_sCallingApplication    As String
    Global g_sSleeveType As String  'Dummy used to kid PR_CalcRaglan ()
    Global g_sPathJOBST  As String
    Global g_sSide       As String
    Global g_sID         As String

   'Windows API Functions Declarations
'    Declare Function GetActiveWindow Lib "User" () As Integer
'    Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer
'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
'    Declare Function GetNumTasks Lib "Kernel" () As Integer
'    Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
   
   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4


   'Calculation constants
    Global Const INCH1_16 = 0.0625
    Global Const INCH1_8 = 0.125
    Global Const INCH3_16 = 0.1875

     

    Global xyAxilla        As XY

    Option Explicit





Function Arccos(X As Double) As Double
     Arccos = Atn(-X / Sqr(-X * X + 1)) + 1.5708
End Function

Function FN_CalcAngle(xyStart As XY, xyEnd As XY) As Double
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim X, Y, rAngle As Double
    
    X = xyEnd.X - xyStart.X
    Y = xyEnd.Y - xyStart.Y

   'Horizontal
    If X = 0 Then
        If Y > 0 Then
            FN_CalcAngle = 90
        Else
            FN_CalcAngle = 270
        End If
        Exit Function
    End If
    
   'Vertical (avoid divide by zero later)
    If Y = 0 Then
        If X > 0 Then
            FN_CalcAngle = 0
        Else
            FN_CalcAngle = 180
        End If
        Exit Function
    End If

   'All other cases
    rAngle = Atn(Y / X) * (180 / PI)         'Convert to degrees
    
    If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
    
    If Y > 0 Then
        FN_CalcAngle = rAngle
    Else
        FN_CalcAngle = rAngle + 180
    End If

End Function

Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

End Function

Static Function FN_CirLinInt(xyStart As XY, xyEnd As XY, xyCen As XY, nRad As Double, xyInt As XY) As Integer
   'Function to calculate the intersection between
   'a line and a circle.
   'Note:-
   '    Returns true if intersection found.
   '    The first intersection (only) is found.
   '    Ported from DRAFIX CAD DLG version.
   '

    Dim nSlope, nA, nB, nC, nK, nM, nCalcTmp, nRoot As Double
    Dim nSign As Integer

    nSlope = FN_CalcAngle(xyStart, xyEnd)

   'Horizontal Line
    If nSlope = 0 Or nSlope = 180 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.X + Sqr(nC) * nSign
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.Y = xyStart.Y
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Vertical Line
    If nSlope = 90 Or nSlope = 270 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.Y + Sqr(nC) * nSign
            If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                xyInt.Y = nRoot
                xyInt.X = xyStart.X
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Non-othogonal line
    If nSlope > 0 Then
        nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X)  'Slope
        nK = xyStart.Y - nM * xyStart.X                     'Y-Axis intercept
        nA = (1 + nM ^ 2)
        nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
        nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
        nCalcTmp = (nB ^ 2) - (4 * nC * nA)

        If (nCalcTmp < 0) Then
            FN_CirLinInt = False         'No Roots
            Exit Function
        End If
        nSign = 1
        While nSign > -2
            nRoot = (-nB + (Sqr(nCalcTmp) / nSign)) / (2 * nA)
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.Y = nM * nRoot + nK
                FN_CirLinInt = True
                Exit Function   'Return first root found
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False 'Should never get to here
    End If
    FN_CirLinInt = False

End Function

Static Function FN_EscapeSlashesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeSlashesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = "\" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeSlashesInString = sEscapedString
    sEscapedString = ""

End Function

Function FN_LinLinInt(xyLine1Start As XY, xyLine1End As XY, xyLine2Start As XY, xyLine2End As XY, xyInt As XY) As Integer
   'Function:
   '       BOOLEAN = FN_LinLinInt( xyLine1Start, xyLine1End, xyLine2Start, xyLine2End, xyInt);
   'Parameters:
   '       xyLine1Start = xyLine1Start.X, xyLine1Start.Y
   '       xyLine1End = xyLine1End.X, xyLine1End.Y
   '       xyLine2Start = xyLine2Start.X, xyLine2Start.Y
   '       xyLine2End = xyLine2End.X, xyLine2End.Y
   '
   'Returns:
   '       True if intersection found and lies on the line
   '       False if no intesection
   '       xyInt =  intersection
   '
    Dim nK1#, nK2#, nM1#, nM2, nSlope1#, nSlope2, nCase#, nY#, nX#

   'Initialy false
    FN_LinLinInt = False

   'Calculate slope of lines
    nCase = 0
    nSlope1 = FN_CalcAngle(xyLine1Start, xyLine1End)
    If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
    If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2
    
    nSlope2 = FN_CalcAngle(xyLine2Start, xyLine2End)
    If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
    If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

    Select Case nCase
     
     Case 0
       'Both lines are Non-Orthogonal Lines
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nM2 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        If (nM1 = nM2) Then Exit Function               'Parallel lines
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)   'Y-Axis intercept
        nK2 = xyLine2Start.Y - (nM2 * xyLine2Start.X)   'Y-Axis intercept
        If (nK1 = nK2) Then Exit Function
       'Find X
        nX = (nK2 - nK1) / (nM1 - nM2)
       'Find Y
        nY = (nM1 * nX) + nK1

     Case 1
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X)                           'Y-Axis intercept
        nY = xyLine1Start.Y
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1

     Case 2
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X)                           'Y-Axis intercept
        nX = xyLine1Start.X
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1

     Case 4
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)                           'Y-Axis intercept
        nY = xyLine2Start.Y
        
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1
    
     Case 5
       'Parallel orthogonal lines, no intersection possible
        Exit Function

     Case 6
       'Line1 is Vertical and the Line2 is Horizontal
        nX = xyLine1Start.X
        nY = xyLine2Start.Y

     Case 8
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)                           'Y-Axis intercept
        nX = xyLine2Start.X
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1
     
     Case 9
       'Line1 is Horizontal and the Line2 is Vertical
        nX = xyLine2Start.X
        nY = xyLine1Start.Y

     Case 10
       'Parallel orthogonal lines, no intersection possible
        Exit Function
     
     Case Else
        Exit Function
    
    End Select

   'Ensure that the points X and Y are on the lines
    xyInt.X = nX
    xyInt.Y = nY

   'Line 1
    If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
    If (nY < min(xyLine1Start.Y, xyLine1End.Y) Or nY > max(xyLine1Start.Y, xyLine1End.Y)) Then Exit Function
   
   'Line 2
    If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
    If (nY < min(xyLine2Start.Y, xyLine2End.Y) Or nY > max(xyLine2Start.Y, xyLine2End.Y)) Then Exit Function
   
    FN_LinLinInt = True

End Function

Function fnGetString(ByVal sString As String, iIndex As Integer, sDelimiter As String) As String
   'Function to return as a string the iIndexth item in a string
   'that using the given string sDelimiter as the delimiter.
   'EG
   '    sString = "Sam Spade Hello"
   '    sDelimiter = " " {SPACE}
   '    fnGetNumber( sString, 2) = "Spade"
   '
   'If the iIndexth item is not found then return "" to indicate an error.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

    If Len(sString) = 0 Then
        fnGetString = ""
        Exit Function
    End If

   'Prepare string
    sString = sString & sDelimiter 'Trailing sDelimiter as stopper for last item

   'Get iIndexth item
    For ii = 1 To iIndex
        iPos = InStr(sString, sDelimiter)
        If ii = iIndex Then
            sString = Left$(sString, iPos - 1)
            fnGetString = sString
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos))
            If Len(sString) = 0 Then
                fnGetString = ""
                Exit Function
            End If
        End If
    Next ii

   'The function should have exited befor this, however just in case
   '(iIndex = 0) we indicate an error,
    fnGetString = ""

End Function

Function max(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the maximum of two numbers
    If nFirst >= nSecond Then
        max = nFirst
    Else
        max = nSecond
    End If
End Function

Function min(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the minimum of two numbers
    If nFirst <= nSecond Then
        min = nFirst
    Else
        min = nSecond
    End If
End Function

Sub PR_AddDBValueToLast(sDBName As String, sDBValue As String)
   'The last entity is given by hEnt
   '
    Print #fNum, "if (hEnt) SetDBData( hEnt," & QQ & sDBName & QQ & CC & QQ; sDBValue & QQ & ");"

End Sub

Sub PR_AddEntityArc(xyCen As XY, xyArcStart As XY, xyArcEnd As XY)
   'Draws an arc with original parameters for AddEntity("arc",...) routine
    
    Dim nRad, nStartAng, nEndAng, nDeltaAng
    
    nRad = FN_CalcLength(xyCen, xyArcStart)
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
    nDeltaAng = Abs(nStartAng - nEndAng)
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
    
End Sub

Sub PR_AddFormatedDBValueToLast(sDBName As String, sLeadingText As String, nValue As Double, sTrailingText As String)
   'The last entity is given by hEnt
   ' Format(sDataType, nValue)
    Print #fNum, "if (hEnt) {"
    Print #fNum, "SetDBData( hEnt," & QQ; sDBName; QCQ; sLeadingText; QQ; "+ Format("; QQ; "length"; QC; nValue; ")+"; QQ; sTrailingText; QQ; ");"

    Print #fNum, "}"

End Sub

Sub PR_AddVertex(xyPoint As XY, nBulge As Double)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to add a Vertex to a polyline opened
   'by PR_StartPoly
   'Allow the use of a bulge factor, but start and end widths
   'will be 0.
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL,  are globals initialised by FN_Open
   '
    Print #fNum, "  AddVertex("; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; "0,0,"; nBulge; ");"

End Sub

Static Sub PR_CalcMidPoint(xyStart As XY, xyEnd As XY, xyMid As XY)
    
    Dim aAngle As Double, nLength As Double

    aAngle = FN_CalcAngle(xyStart, xyEnd)
    nLength = FN_CalcLength(xyStart, xyEnd)
    
    If nLength = 0 Then
        xyMid = xyStart 'Avoid overflow
    Else
        PR_CalcPolar xyStart, aAngle, nLength / 2, xyMid
    End If

End Sub

Static Sub PR_CalcPolar(xyStart As XY, ByVal nAngle As Double, ByVal nLength As Double, xyReturn As XY)
   'Procedure to return a point at a distance and an angle from a given point
   '
   'PI is a Global Constant

    Dim A As Double, B As Double

   'Ensure angles are +ve and between 0 to 360 degrees
    If nAngle > 360 Then nAngle = nAngle - 360
    If nAngle < 0 Then nAngle = nAngle + 360

    Select Case nAngle
        Case 0
            xyReturn.X = xyStart.X + nLength
            xyReturn.Y = xyStart.Y
        Case 180
            xyReturn.X = xyStart.X - nLength
            xyReturn.Y = xyStart.Y
        Case 90
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y + nLength
        Case 270
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y - nLength
        Case Else
           'Convert from degees to radians
            nAngle = nAngle * PI / 180
            B = Sin(nAngle) * nLength
            A = Cos(nAngle) * nLength
            xyReturn.X = xyStart.X + A
            xyReturn.Y = xyStart.Y + B
        End Select

End Sub

Static Sub PR_DrawArc(xyCen As XY, xyArcStart As XY, xyArcEnd As XY)
   'Draws an arc
   'Restrictions
   '    1. Arc must be 180 degrees or less
    
Dim nRad, nStartAng, nEndAng, nDeltaAng, nOpp, nAdj, nSign
Dim xyMidPoint As XY

' ORIGINAL CODE ********************
'    nRad = FN_CalcLength(xyCen, xyArcStart)
'
'    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
'
'    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
'
'    If Side > 0 Then nDeltaAng = Abs(nEndAng - nStartAng) Else nDeltaAng = nEndAng - nStartAng
' ORIGINAL CODE ********************
'

    nRad = FN_CalcLength(xyCen, xyArcStart)
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)

   'Direction of arc
    nSign = Sgn(nEndAng - nStartAng)
    If Abs(nEndAng - nStartAng) > 180 Then
        nSign = 0 - nSign
    End If
   
   'Included angle
    PR_CalcMidPoint xyArcStart, xyArcEnd, xyMidPoint
    nAdj = FN_CalcLength(xyCen, xyMidPoint)
    nOpp = FN_CalcLength(xyArcStart, xyArcEnd) / 2

    If nAdj = 0 Then
        nDeltaAng = 180
    Else
        nDeltaAng = (Atn(nOpp / nAdj) * 2) * (180 / PI)
    End If

    nDeltaAng = nDeltaAng * nSign
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawFitted(Profile As curve)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a FITTED curve through the points
   'given in Profile.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '
    Dim ii As Integer

   'Draw the profile
   '    If there is no vertex or only one vertex then exit.
   '    For two vertex draw as a polyline (this degenerates to a Single line).
   '    For three vertex draw as a polyline (as no fitted curve can be drawn
   '    by a macro).
   '
    Select Case Profile.n
        Case 0 To 1
            Exit Sub
        Case 3
            PR_DrawPoly Profile
            Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
           'Warn the user to smooth the curve
 '           Print #fNum, "Display ("; QQ; "message"; QCQ; "OKquestion"; QCQ; "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line"; QQ; ");"
        Case Else
            Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
            For ii = 1 To Profile.n
                Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.Y(ii))
            Next ii
            Print #fNum, ");"
            Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
    End Select

End Sub

Sub PR_DrawLine(xyStart As XY, xyFinish As XY)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw a LINE between two points.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Print #fNum, "hEnt = AddEntity("
    Print #fNum, QQ; "line"; QC
    Print #fNum, "xyStart.x+"; Str$(xyStart.X); CC; "xyStart.y+"; Str$(xyStart.Y); CC
    Print #fNum, "xyStart.x+"; Str$(xyFinish.X); CC; "xyStart.y+"; Str$(xyFinish.Y)
    Print #fNum, ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawMarker(xyPoint As XY)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; "0.125);"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawMarkerNamed(sName$, xyPoint As XY, nWidth, nHeight, nAngle)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; sName; QC; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; nWidth; CC; nHeight; CC; nAngle; ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawPoly(Profile As curve)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a POLYLINE through the points
   'given in Profile.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '
    Dim ii As Integer

   'Exit if nothing to draw
    If Profile.n <= 1 Then Exit Sub

    Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    For ii = 1 To Profile.n
    Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.Y(ii))
    Next ii
    Print #fNum, ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub


Sub PR_DrawRectangle(xyLLCorner As XY, xyURCorner As XY)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw a RECTANGLE using opposite corner points.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '

    Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polygon"; QQ
    Print #fNum, CC; "xyStart.x+"; Str$(xyLLCorner.X); CC; "xyStart.y+"; Str$(xyLLCorner.Y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyLLCorner.X); CC; "xyStart.y+"; Str$(xyURCorner.Y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyURCorner.X); CC; "xyStart.y+"; Str$(xyURCorner.Y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyURCorner.X); CC; "xyStart.y+"; Str$(xyLLCorner.Y)
    Print #fNum, ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawText(sText As Variant, xyInsert As XY, nHeight As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw TEXT at the given height.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
   '
   '
    Dim nWidth
    nWidth = nHeight * g_nCurrTextAspect
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.Y); CC; nWidth; CC; nHeight; CC; g_nCurrTextAngle; ");"
    
End Sub


Sub PR_EndPoly()
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to end a POLYLINE
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '    STRING  sID
   '
   'Note:-
   '    fNum, CC, QQ, NL,  are globals initialised by FN_Open
   '
    Print #fNum, "EndPoly();"
    Print #fNum, "hEnt = UID (""find"", UID (""getmax"")) ;"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_InsertSymbol(sSymbol As String, xyInsert As XY, nScale As Single, nRotation As Single)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to insert a SYMBOL.
   'Where:-
   '    sSymbol     Symbol name, must exist and be in the symbol library
   '    xyInsert    The insertion point
   '    nScale      Symbol scaling factor, 1 = No scaling
   '    nRotation   Symbol rotation about insertion point
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   'and
   '    The DRAFIX symbol library sPathJOBST + "\JOBST.SLB" exists
   '
   'Note:-
   '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
   '    g_sPathJOBST is path to JOBST CAD System
   '
    Print #fNum, "SetSymbolLibrary("; QQ; FN_EscapeSlashesInString(g_sPathJOBST); "\\JOBST.SLB"; QQ; ");"
    Print #fNum, "Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ");"
    Print #fNum, "hEnt = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC
    Print #fNum, "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.Y); CC
    Print #fNum, Str$(nScale); CC & Str$(nScale); CC; Str$(nRotation); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
   
End Sub

Sub PR_MakeXY(xyReturn As XY, X As Double, Y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.X = X
    xyReturn.Y = Y

End Sub

Sub PR_NamedHandle(sHandleName As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to retain the entity handle of a previously
   'added entity.
   '
   'Assumes that hEnt is the entity handle to the just inserted entity.
   '
    Print #fNum, "HANDLE "; sHandleName; ";"
    Print #fNum, sHandleName; " = hEnt;"

End Sub

Sub PR_Open(sDrafixFile As String, fNum)
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    Open sDrafixFile For Output As fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(13) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 32    'Bottom
    g_nCurrTextFont = 0         'BLOCK
    g_nCurrTextAngle = 0


   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//by Visual Basic, Mesh Drawing Routine"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;"
    Print #fNum, "XY     xyStart, xyOrigin, xyScale, xyO;"
    Print #fNum, "STRING sFileNo, sSide, sID, sPathJOBST;"
    Print #fNum, "ANGLE  aAngle;"

   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; "));"
    
   'Set xyStart to 0,0 as we are using absolute co-ords
    Print #fNum, "xyStart.y = 0; xyStart.x = 0;"
   
   'Display Hour Glass Symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
End Sub

Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable
   '
    Print #fNum, sLine

End Sub

Sub PR_PutNumberAssign(sVariableName As String, nAssignedNumber As Variant)
   'Procedure to put a number assignment
   'Adds a semi-colon
   '    fNum is global variable

    Print #fNum, sVariableName; "="; Str$(nAssignedNumber); ";"

End Sub

Sub PR_PutStringAssign(sVariableName As String, sAssignedString As Variant)
   'Procedure to put a string assignment
   'Encloses String in quotes and adds a semi-colon
   '    fNum is global variable

    Print #fNum, sVariableName; "="; QQ; sAssignedString; QQ; ";"

End Sub

Sub PR_SetLayer(sNewLayer As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the current LAYER.
   'For this to work it assumes that hLayer is defined in DRAFIX as
   'a HANDLE.
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_sCurrentLayer are globals initialised by FN_Open
   '
   'To reduce unessesary writing of DRAFIX code check that the new layer
   'is different from the Current layer, change only if it is different.
   '

    If g_sCurrentLayer = sNewLayer Then Exit Sub
    g_sCurrentLayer = sNewLayer
    
    Print #fNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; sNewLayer; QQ; ");"
    Print #fNum, "if ( hLayer != %badtable)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Sub PR_SetLineStyle(sStyle As String)
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; sStyle; QQ; "));"
End Sub

Sub PR_SetTextData(nHoriz As Variant, nVert As Variant, nHt As Variant, nAspect As Variant, nFont As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the TEXT default attributes, these are
   'based on the values in the arguments.  Where the value is -ve then this
   'attribute is not set.
   'where :-
   '    nHoriz      Horizontal justification (1=Left, 2=Cen, 4=Right)
   '    nVert       Verticalal justification (8=Top, 16=Cen, 32=Bottom)
   '    nHt         Text height
   '    nAspect     Text aspect ratio (heigth/width)
   '    nFont       Text font (0 to 18)
   '
   'N.B. No checking is done on the values given
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextHt, g_nCurrTextAspect,
   '    g_nCurrTextHorizJust, g_nCurrTextVertJust, g_nCurrTextFont
   '    are globals initialised by FN_Open
   '
  
    If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; nHoriz; ");"
    g_nCurrTextHorizJust = nHoriz
    End If

    If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; nVert; ");"
    g_nCurrTextVertJust = nVert
    End If

    If nHt >= 0 And g_nCurrTextHt <> nHt Then
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; nHt; ");"
    g_nCurrTextHt = nHt
    End If

    If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; nAspect; ");"
    g_nCurrTextAspect = nAspect
    End If

    If nFont >= 0 And g_nCurrTextFont <> nFont Then
    Print #fNum, "SetData("; QQ; "TextFont"; QC; nFont; ");"
    g_nCurrTextFont = nFont
    End If


End Sub

Static Sub PR_StartPoly()
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to start a POLYLINE the points will
   'be given by the PR_AddVertex and finished by PR_EndPoly
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '    STRING  sID
   '
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '

    Print #fNum, "StartPoly (""PolyLine"");"


End Sub


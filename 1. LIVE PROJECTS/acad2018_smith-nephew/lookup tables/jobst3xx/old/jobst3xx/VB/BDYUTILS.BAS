Attribute VB_Name = "BDYUTILS"

'File:      BDYUTILS.BAS
'Purpose:   Module of utilities for use by the body
'
'
'Version:   1.01
'Date:      30.Jun.95
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
   
 'The following are now declared as PUBLIC in MESHCALC.BAS
   'XY data type to represent points
 '   Type XY
 '       x As Double
 '       y As Double
 '   End Type
    
 '   Public Type BiArc
 '       xyStart     As XY
 '       xyTangent   As XY
 '       xyEnd       As XY
 '       xyR1        As XY
 '       xyR2        As XY
 '       nR1         As Double
 '       nR2         As Double
 '   End Type
   
 '   Type curve
 '       n As Integer
 '       x(1 To 100) As Double
 '       y(1 To 100) As Double
 '   End Type


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



    Option Explicit

Function Arccos(x As Double) As Double
     Arccos = Atn(-x / Sqr(-x * x + 1)) + 1.5708
End Function

Public Function FN_BiArcCurve(xyStart As XY, aStart As Double, xyEnd As XY, aEnd As Double, Profile As BiArc) As Integer
   'Procedure to fit two arcs with a common tangent
   'between two points at the tangent angles specified.
   'Returning a BiArc curve.
   '
   'Input
   '        xyStart     Start point (XY System)
   '        xyEnd       End Point   (XY System)
   '        aStart      Start Tangent Angle in degrees
   '        aEnd        End Tangent Angle in degrees
   '
   'Output
   '        Profile     BiArc in (XY System)
   '
   'Restrictions:-
   'Angles must be positive and < 360
   'The data must result in a NON-INFLECTING curve.
   '
   '
   'Notes:-
   'The BiArc is a structure that is used to represent
   'the curve.
   '
   '        Type BiArc
   '                xyStart     as XY
   '                xyTangent   as XY
   '                xyEnd       as XY
   '                xyR1        as XY
   '                xyR2        as XY
   '                nR1         as Double
   '                nR2         as Double
   '        end
   '
   'The point of this represntation is to make
   'the drawing of the curve as simple as possible,
   'without the need to make further calculations.
   '
   '
   'Acknowledgements:-
   '
   '    British Ship Research Association (BSRA)
   '    Technical Memorandum No. 388
   '
   '   "The Fitting of Smooth Curves by Circular
   '    Arcs and Straight Lines"
   '    K.M.Bolton, B.Sc., M.Sc., Grad.I.M.A.
   '    October 1970
   '
   'KNOWN BUGS
   'Needs much more work to make a more general robust tool.
   'Works OK for some special cases (cf THUMB Curves)
   'GG 22.Mar.96
   '
   '
    Dim aAxis#, nTmp#, nLength#
    Dim Theta1#, Theta2#, Phi1#, Phi2#
    Dim A#, B#, c#, d#, P#, S1#, S2#, C1#, C2#
    Dim Rs#, R1#, R2#, Rmin
    Dim uvR1 As XY, uvR2 As XY, uvOrigin As XY
'    Dim fError

   'Use a file for printing results for debug
   'Open file
'    fError = FreeFile
'    Open "C:\TMP\FIT_ERR.DAT" For Output As fError


   'Initially return false
    FN_BiArcCurve = False
    
   'Check for silly data and return
    If aStart = aEnd Then GoTo Error_Close
    If xyStart.y = xyEnd.y And xyStart.x = xyEnd.x Then GoTo Error_Close

   'Translate tangent angle in XY system to the UV system.
   'Where the U-Axis is specified by the line xyStart,xyEnd
   '
    aAxis = FN_CalcAngle(xyStart, xyEnd)
    nLength = FN_CalcLength(xyStart, xyEnd)
'Print #fError, "aAxis degrees ="; aAxis
'Print #fError, "nLength ="; nLength

    Theta1 = aStart - aAxis
    If Theta1 = 0 Then GoTo Error_Close 'Straight line

    If Theta1 < 0 Then Theta1 = 360 + Theta1
'Print #fError, "Theta1 degrees ="; Theta1
    Theta1 = Theta1 * (PI / 180)
'Print #fError, "Theta1="; Theta1

    Theta2 = aEnd - aAxis
    If Theta2 < 0 Then Theta2 = 360 + Theta2
'Print #fError, "Theta2 degrees="; Theta2
    Theta2 = Theta2 * (PI / 180)
'Print #fError, "Theta2="; Theta2

   
   'Check that it is non-inflecting
   'return an error (false) for the inflecting case
   'let the calling routine worry about handling it
    If (Theta1 > 0 And Theta1 < (PI / 2)) And (Theta2 > 0 And Theta2 < (PI / 2)) Then GoTo Error_Close
    If (Theta1 > (3 * (PI / 2)) And Theta1 < (2 * PI)) And (Theta2 > (3 * (PI / 2)) And Theta2 < (2 * PI)) Then GoTo Error_Close
    
   'Calculate acute unsigned tangent angles to the line in the UV
   'co-ordinate system
   ' Phi1 = Theta1
    If Theta1 < PI Then
        Phi1 = Theta1
    Else
        Phi1 = Abs((2 * PI) - Theta1)
    End If
    If Theta2 < PI Then
        Phi2 = Theta2
    Else
        Phi2 = Abs((2 * PI) - Theta2)
    End If

'Print #fError, "Phi1="; Phi1
'Print #fError, "Phi2="; Phi2


   'Calculate R1 and R2
    S1 = Abs(Sin(Theta1))
    C1 = (-Sin(Theta1) * Cos(Theta1)) / S1
'Print #fError, "S1="; S1; "C1="; C1

    S2 = Abs(Sin(Theta2))
    C2 = (Sin(Theta2) * Cos(Theta2)) / S2
'Print #fError, "S2="; S2; "C2="; C2

    P = FN_CalcLength(xyStart, xyEnd)
'Print #fError, "P="; P

    If Phi1 <> Phi2 Then
'Print #fError, "Phi1 <> Phi2"
        A = S1 + S2
        B = (S1 * S2) - (C1 * C2) + 1
        c = S2
        Rs = (P * c) / B
        nTmp = (c ^ 2) - (c * A) + (B / 2)
'Print #fError, "A="; A; "B="; B; "C="; C
'Print #fError, "Rs="; Rs
'Print #fError, "nTmp="; nTmp
       'As this is a root we check that it is not -ve
        If nTmp < 0 Then GoTo Error_Close
        nTmp = (P * Sqr(nTmp)) / B
'Print #fError, "nTmp="; nTmp
        If Phi1 > Phi2 Then
            R1 = Rs - nTmp
        Else
            R1 = Rs + nTmp
        End If
'Print #fError, "R1="; R1
        If R1 <= 0 Then GoTo Error_Close
        d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
'Print #fError, "D="; d
        R2 = R1 - d
'Print #fError, "R2="; R2
    Else
'Print #fError, "Phi1 = Phi2"
        A = 2 * Sin(Theta1)
        B = (Sin(Theta1) ^ 2) - (Cos(Theta1) ^ 2) + 1
        R1 = (P * A) / (2 * B)
        R2 = R1
'Print #fError, "A="; A; "B="; A;
'Print #fError, "R1="; R1
    End If

   'The radi R1 and R2 must be greater than the specified
   'minimum Rmin
    Rmin = nLength / 3
    If R1 < Rmin Then
'Print #fError, "R1 < Rmin"
        R1 = Rmin
        d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
'Print #fError, "D="; d
        R2 = R1 - d
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
    End If
    If R2 < Rmin Then
'Print #fError, "R2 < Rmin"
        R2 = Rmin
        R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
    End If

   'Check for errors
    If R1 < 0 Or R2 < 0 Then GoTo Error_Close

   'Using the calculated radi, create BiArc Curve
   'Start and end points of bi-arc curve
    Profile.xyStart = xyStart
    Profile.xyEnd = xyEnd

   'Centers of arcs
   'Get UV co-ordinates
    PR_MakeXY uvR1, R1 * S1, R1 * C1
'Print #fError, "uvR1="; uvR1.X, uvR1.Y
'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
    PR_MakeXY uvR2, P - (R2 * S2), R2 * C2
'Print #fError, "uvR2="; uvR2.X, uvR2.Y
'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
    PR_MakeXY uvOrigin, 0, 0
   'Translate to XY co-ordinates
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
   'Tangent point on arc
    If R1 < R2 Then
        PR_CalcPolar Profile.xyR1, FN_CalcAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent
    Else
        PR_CalcPolar Profile.xyR1, FN_CalcAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent
    End If
   'radi of arcs
    Profile.nR1 = R1
    Profile.nR2 = R2
    
   'Test that the Tangent point lies between the start and end points
    If FN_CalcLength(xyStart, Profile.xyTangent) > nLength Then GoTo Error_Close
    If FN_CalcLength(xyEnd, Profile.xyTangent) > nLength Then GoTo Error_Close

   'return true as we have a sucessful fit
'Close #fError
    FN_BiArcCurve = True

    Exit Function
    
Error_Close:
'Print #fError, "Error and close"
'Close #fError

End Function

Function FN_CalcAngle(xyStart As XY, xyEnd As XY) As Double
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim x, y, rAngle As Double
    
    x = xyEnd.x - xyStart.x
    y = xyEnd.y - xyStart.y

   'Horizontal
    If x = 0 Then
        If y > 0 Then
            FN_CalcAngle = 90
        Else
            FN_CalcAngle = 270
        End If
        Exit Function
    End If
    
   'Vertical (avoid divide by zero later)
    If y = 0 Then
        If x > 0 Then
            FN_CalcAngle = 0
        Else
            FN_CalcAngle = 180
        End If
        Exit Function
    End If

   'All other cases
    rAngle = Atn(y / x) * (180 / PI)         'Convert to degrees
    
    If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
    
    If y > 0 Then
        FN_CalcAngle = rAngle
    Else
        FN_CalcAngle = rAngle + 180
    End If

End Function

Function FN_CalcCirCirInt(xyCen1 As XY, nRad1 As Double, xyCen2 As XY, nRad2 As Double, xyInt1 As XY, xyInt2 As XY) As Integer
   'Function that will return
   '    TRUE    if two circles intersect
   '    FALSE   if two circles don't intersect
   '
   'The intersection points are returned in the values
   '
   '    xyInt1 & xyInt2
   '
   'with intersection with lowest X value as xyInt1
   '
    Dim nLength As Double, aAngle  As Double, aTheta As Double, nCosTheta  As Double
    Dim xyTmp As XY

    FN_CalcCirCirInt = False

   'Check that the circles can intersect
   'It is a theorem of plane geometry that no three real numbers
   'a, b and c can be the lenghts of the sides of a triangle unless
   'the sum of any two is greater than the third.
   'We use this as our main test of possible intersection, we also check for silly
   'data.

   'Test for silly data
    If xyCen1.x = xyCen2.x And xyCen1.y = xyCen2.y Then Exit Function
    If nRad1 <= 0 Or nRad2 <= 0 Then Exit Function

   'Test for intersection
    nLength = FN_CalcLength(xyCen1, xyCen2)

    If (nLength + nRad1 < nRad2) Or (nLength + nRad2 < nRad1) Or (nRad1 + nRad2 < nLength) Then
        Exit Function
    Else
        FN_CalcCirCirInt = True
    End If

   'Calculate intesection points
   '
   'Special case where circles touch (ie Intersect at one point only)
    
    
   'Angle between centers
   'Note: Length between centers from above
    aAngle = FN_CalcAngle(xyCen1, xyCen2)

   'Get angle w.r.t line between centers to the intersection point
   'use cosine rule
   '
    nCosTheta = -((nRad2 ^ 2 - (nLength ^ 2 + nRad1 ^ 2)) / (2 * nLength * nRad1))

    aTheta = Atn(-nCosTheta / Sqr(-(nCosTheta ^ 2) + 1)) + 1.5708

    aTheta = aTheta * (180 / PI) 'convert to degrees

    aAngle = aAngle - aTheta
    PR_CalcPolar xyCen1, aAngle, nRad1, xyInt1

    aAngle = aAngle + aTheta
    PR_CalcPolar xyCen1, aAngle, nRad1, xyInt2

    If xyInt2.x < xyInt1.x Then
        xyTmp = xyInt1
        xyInt1 = xyInt2
        xyInt2 = xyTmp
    End If

    
End Function

Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.x - xyStart.x) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)

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
        nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.x + Sqr(nC) * nSign
            If nRoot >= min(xyStart.x, xyEnd.x) And nRoot <= max(xyStart.x, xyEnd.x) Then
                xyInt.x = nRoot
                xyInt.y = xyStart.y
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
        nC = nRad ^ 2 - (xyStart.x - xyCen.x) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.y + Sqr(nC) * nSign
            If nRoot >= min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
                xyInt.y = nRoot
                xyInt.x = xyStart.x
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
        nM = (xyEnd.y - xyStart.y) / (xyEnd.x - xyStart.x)  'Slope
        nK = xyStart.y - nM * xyStart.x                     'Y-Axis intercept
        nA = (1 + nM ^ 2)
        nB = 2 * (-xyCen.x + (nM * nK) - (xyCen.y * nM))
        nC = (xyCen.x ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
        nCalcTmp = (nB ^ 2) - (4 * nC * nA)

        If (nCalcTmp < 0) Then
            FN_CirLinInt = False         'No Roots
            Exit Function
        End If
        nSign = 1
        While nSign > -2
            nRoot = (-nB + (Sqr(nCalcTmp) / nSign)) / (2 * nA)
            If nRoot >= min(xyStart.x, xyEnd.x) And nRoot <= max(xyStart.x, xyEnd.x) Then
                xyInt.x = nRoot
                xyInt.y = nM * nRoot + nK
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
        nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.x - xyLine1Start.x) 'Slope
        nM2 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.x - xyLine2Start.x) 'Slope
        If (nM1 = nM2) Then Exit Function               'Parallel lines
        nK1 = xyLine1Start.y - (nM1 * xyLine1Start.x)   'Y-Axis intercept
        nK2 = xyLine2Start.y - (nM2 * xyLine2Start.x)   'Y-Axis intercept
        If (nK1 = nK2) Then Exit Function
       'Find X
        nX = (nK2 - nK1) / (nM1 - nM2)
       'Find Y
        nY = (nM1 * nX) + nK1

     Case 1
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.x - xyLine2Start.x) 'Slope
        nK1 = xyLine2Start.y - (nM1 * xyLine2Start.x)                           'Y-Axis intercept
        nY = xyLine1Start.y
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1

     Case 2
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.x - xyLine2Start.x) 'Slope
        nK1 = xyLine2Start.y - (nM1 * xyLine2Start.x)                           'Y-Axis intercept
        nX = xyLine1Start.x
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1

     Case 4
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.x - xyLine1Start.x) 'Slope
        nK1 = xyLine1Start.y - (nM1 * xyLine1Start.x)                           'Y-Axis intercept
        nY = xyLine2Start.y
        
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1
    
     Case 5
       'Parallel orthogonal lines, no intersection possible
        Exit Function

     Case 6
       'Line1 is Vertical and the Line2 is Horizontal
        nX = xyLine1Start.x
        nY = xyLine2Start.y

     Case 8
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.x - xyLine1Start.x) 'Slope
        nK1 = xyLine1Start.y - (nM1 * xyLine1Start.x)                           'Y-Axis intercept
        nX = xyLine2Start.x
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1
     
     Case 9
       'Line1 is Horizontal and the Line2 is Vertical
        nX = xyLine2Start.x
        nY = xyLine1Start.y

     Case 10
       'Parallel orthogonal lines, no intersection possible
        Exit Function
     
     Case Else
        Exit Function
    
    End Select

   'Ensure that the points X and Y are on the lines
    xyInt.x = nX
    xyInt.y = nY

   'Line 1
    If (nX < min(xyLine1Start.x, xyLine1End.x) Or nX > max(xyLine1Start.x, xyLine1End.x)) Then Exit Function
    If (nY < min(xyLine1Start.y, xyLine1End.y) Or nY > max(xyLine1Start.y, xyLine1End.y)) Then Exit Function
   
   'Line 2
    If (nX < min(xyLine2Start.x, xyLine2End.x) Or nX > max(xyLine2Start.x, xyLine2End.x)) Then Exit Function
    If (nY < min(xyLine2Start.y, xyLine2End.y) Or nY > max(xyLine2Start.y, xyLine2End.y)) Then Exit Function
   
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
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.x); CC; "xyStart.y +"; Str$(xyCen.y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
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
    Print #fNum, "  AddVertex("; "xyStart.x+"; xyPoint.x; CC; "xyStart.y+"; xyPoint.y; CC; "0,0,"; nBulge; ");"

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
            xyReturn.x = xyStart.x + nLength
            xyReturn.y = xyStart.y
        Case 180
            xyReturn.x = xyStart.x - nLength
            xyReturn.y = xyStart.y
        Case 90
            xyReturn.x = xyStart.x
            xyReturn.y = xyStart.y + nLength
        Case 270
            xyReturn.x = xyStart.x
            xyReturn.y = xyStart.y - nLength
        Case Else
           'Convert from degees to radians
            nAngle = nAngle * PI / 180
            B = Sin(nAngle) * nLength
            A = Cos(nAngle) * nLength
            xyReturn.x = xyStart.x + A
            xyReturn.y = xyStart.y + B
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
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.x); CC; "xyStart.y +"; Str$(xyCen.y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawCircle(xyCen As XY, nRadius As Double)
   'Draw a circle at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "circle"; QC; "xyStart.x+"; xyCen.x; CC; "xyStart.y+"; xyCen.y; CC; nRadius; ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

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
                Print #fNum, CC; "xyStart.x+"; Str$(Profile.x(ii)); CC; "xyStart.y+"; Str$(Profile.y(ii))
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
    Print #fNum, "xyStart.x+"; Str$(xyStart.x); CC; "xyStart.y+"; Str$(xyStart.y); CC
    Print #fNum, "xyStart.x+"; Str$(xyFinish.x); CC; "xyStart.y+"; Str$(xyFinish.y)
    Print #fNum, ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawMarker(xyPoint As XY)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart.x+"; xyPoint.x; CC; "xyStart.y+"; xyPoint.y; CC; "0.125);"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawMarkerNamed(sName$, xyPoint As XY, nWidth, nHeight, nAngle)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; sName; QC; "xyStart.x+"; xyPoint.x; CC; "xyStart.y+"; xyPoint.y; CC; nWidth; CC; nHeight; CC; nAngle; ");"
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
    Print #fNum, CC; "xyStart.x+"; Str$(Profile.x(ii)); CC; "xyStart.y+"; Str$(Profile.y(ii))
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
    Print #fNum, CC; "xyStart.x+"; Str$(xyLLCorner.x); CC; "xyStart.y+"; Str$(xyLLCorner.y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyLLCorner.x); CC; "xyStart.y+"; Str$(xyURCorner.y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyURCorner.x); CC; "xyStart.y+"; Str$(xyURCorner.y)
    Print #fNum, CC; "xyStart.x+"; Str$(xyURCorner.x); CC; "xyStart.y+"; Str$(xyLLCorner.y)
    Print #fNum, ");"
    Print #fNum, "if (hEnt) SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub Pr_DrawText(sText As Variant, xyInsert As XY, nHeight As Variant)
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
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; "xyStart.x+"; Str$(xyInsert.x); CC; "xyStart.y+"; Str$(xyInsert.y); CC; nWidth; CC; nHeight; CC; g_nCurrTextAngle; ");"
    
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
    Print #fNum, "xyStart.x+"; Str$(xyInsert.x); CC; "xyStart.y+"; Str$(xyInsert.y); CC
    Print #fNum, Str$(nScale); CC & Str$(nScale); CC; Str$(nRotation); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
   
End Sub

Sub PR_MakeXY(xyReturn As XY, x As Double, y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.x = x
    xyReturn.y = y

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


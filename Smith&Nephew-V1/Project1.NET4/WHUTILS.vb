Option Strict Off
Option Explicit On
Module WHUTILS
    'Project:   WHDRAW1.MAK
    'File:      WHUTILS.BAS
    'Purpose:   Module of utilities for use by the waist
    '           height
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

    'XY data type to represent points
    Structure xy
        Dim X As Double
        Dim Y As Double
    End Structure

    Structure Curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim Y() As Double

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim X(100)
            'UPGRADE_WARNING: Lower bound of array Y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim Y(100)
        End Sub
    End Structure


    '   'Globals set by FN_Open
    '   Public CC As Object 'Comma
    '   Public QQ As Object 'Quote
    '   Public NL As Object 'Newline
    '   Public fNum As Object 'Macro file number
    '   Public QCQ As Object 'Quote Comma Quote
    '   Public QC As Object 'Quote Comma
    '   Public CQ As Object 'Comma Quote

    '   'Store current layer and text setings to reduce DRAFIX code
    '   'this value is checked in PR_SetLayer
    '   Public g_sCurrentLayer As String
    '   Public g_nCurrTextHt As Object
    '   Public g_nCurrTextAspect As Object
    '   Public g_nCurrTextHorizJust As Object
    '   Public g_nCurrTextVertJust As Object
    '   Public g_nCurrTextFont As Object
    '   Public g_nCurrTextAngle As Object
    '   Public g_sPathJOBST As String

    Structure BiArc
        Dim xyStart As xy
        Dim xyTangent As xy
        Dim xyEnd As xy
        Dim xyR1 As xy
        Dim xyR2 As xy
        Dim nR1 As Double
        Dim nR2 As Double
    End Structure

    'Global constants for use with PR_SetTextData
    '
    Public Const HELVETICA As Short = 10
    Public Const BLOCK As Short = 0
    Public Const LEFT_ As Short = 1
    Public Const RIGHT_ As Short = 4
    Public Const HORIZ_CENTER As Short = 2
    Public Const TOP_ As Short = 8
    Public Const BOTTOM_ As Short = 32
    Public Const VERT_CENTER As Short = 16
    Public Const CURRENT As Short = -1

    Public Const PI As Double = 3.141592654

    Public g_nCurrTextAngle As Double


    Function Arccos(ByRef X As Double) As Double
		Arccos = System.Math.Atan(-X / System.Math.Sqrt(-X * X + 1)) + 1.5708
	End Function

    Function FN_CalcAngle(ByRef xyStart As Project1.HEADNECK1.xy, ByRef xyEnd As Project1.HEADNECK1.xy) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, Y As Object
        Dim rAngle As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = xyEnd.X - xyStart.X
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Y = xyEnd.Y - xyStart.Y

        'Horizontal
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Y > 0 Then
                FN_CalcAngle = 90
            Else
                FN_CalcAngle = 270
            End If
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Y = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If X > 0 Then
                FN_CalcAngle = 0
            Else
                FN_CalcAngle = 180
            End If
            Exit Function
        End If

        'All other cases
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rAngle = System.Math.Atan(Y / X) * (180 / PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If

    End Function

    Function FN_CalcLength(ByRef xyStart As Project1.HEADNECK1.xy, ByRef xyEnd As Project1.HEADNECK1.xy) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

    End Function

    Function FN_CirLinInt(ByRef xyStart As Project1.HEADNECK1.xy, ByRef xyEnd As Project1.HEADNECK1.xy, ByRef xyCen As Project1.HEADNECK1.xy, ByRef nRad As Double, ByRef xyInt As Project1.HEADNECK1.xy) As Short
        'Function to calculate the intersection between
        'a line and a circle.
        'Note:-
        '    Returns true if intersection found.
        '    The first intersection (only) is found.
        '    Ported from DRAFIX CAD DLG version.
        '

        Static nM, nC, nA, nSlope, nB, nK, nCalcTmp As Object
        Static nRoot As Double
        Static nSign As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nSlope = FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 0 Or nSlope = 180 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.X, xyEnd.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.X, xyEnd.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = xyStart.Y
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Vertical Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 90 Or nSlope = 270 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.Y + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.Y, xyEnd.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.Y, xyEnd.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                    xyInt.Y = nRoot
                    xyInt.X = xyStart.X
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Non-othogonal line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X) 'Slope
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nK = xyStart.Y - nM * xyStart.X 'Y-Axis intercept
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nA = (1 + nM ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.X, xyEnd.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.X, xyEnd.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xyInt.Y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False

    End Function

    Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Static ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Static Char_Renamed As String
		Static sEscapedString As String
		
		FN_EscapeSlashesInString = ""
		
		For ii = 1 To Len(sAssignedString)
			'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Char_Renamed = Mid(sAssignedString, ii, 1)
			If Char_Renamed = "\" Then
				sEscapedString = sEscapedString & "\" & Char_Renamed
			Else
				sEscapedString = sEscapedString & Char_Renamed
			End If
		Next ii
		
		FN_EscapeSlashesInString = sEscapedString
		sEscapedString = ""
		
	End Function

    Function FN_LinLinInt(ByRef xyLine1Start As Project1.HEADNECK1.xy, ByRef xyLine1End As Project1.HEADNECK1.xy, ByRef xyLine2Start As Project1.HEADNECK1.xy, ByRef xyLine2End As Project1.HEADNECK1.xy, ByRef xyInt As Project1.HEADNECK1.xy) As Short
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
        Dim nY, nSlope1, nK2, nK1, nM1, nCase, nX As Double
        Dim nM2, nSlope2 As Object

        'Initialy false
        FN_LinLinInt = False

        'Calculate slope of lines
        nCase = 0
        nSlope1 = FN_CalcAngle(xyLine1Start, xyLine1End)
        If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
        If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2

        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nSlope2 = FN_CalcAngle(xyLine2Start, xyLine2End)
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

        Select Case nCase

            Case 0
                'Both lines are Non-Orthogonal Lines
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nM2 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (nM1 = nM2) Then Exit Function 'Parallel lines
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nK2 = xyLine2Start.Y - (nM2 * xyLine2Start.X) 'Y-Axis intercept
                If (nK1 = nK2) Then Exit Function
                'Find X
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nX = (nK2 - nK1) / (nM1 - nM2)
                'Find Y
                nY = (nM1 * nX) + nK1

            Case 1
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nY = xyLine1Start.Y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1

            Case 2
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nX = xyLine1Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1

            Case 4
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
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
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
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
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.X, xyLine1End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.X, xyLine1End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.Y, xyLine1End.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.Y, xyLine1End.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nY < min(xyLine1Start.Y, xyLine1End.Y) Or nY > max(xyLine1Start.Y, xyLine1End.Y)) Then Exit Function

        'Line 2
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.X, xyLine2End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.X, xyLine2End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.Y, xyLine2End.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.Y, xyLine2End.Y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nY < min(xyLine2Start.Y, xyLine2End.Y) Or nY > max(xyLine2Start.Y, xyLine2End.Y)) Then Exit Function

        FN_LinLinInt = True


    End Function

    Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
		' Returns the maximum of two numbers
		'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFirst >= nSecond Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			max = nFirst
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			max = nSecond
		End If
	End Function
	
	Function min(ByRef nFirst As Object, ByRef nSecond As Object) As Object
		' Returns the minimum of two numbers
		'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFirst <= nSecond Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			min = nFirst
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			min = nSecond
		End If
	End Function
	
	Sub PR_AddDBValueToLast(ByRef sDBName As String, ByRef sDBValue As String)
        'The last entity is given by hEnt
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "if (hEnt) SetDBData( hEnt," & ARMDIA1.QQ & sDBName & ARMDIA1.QQ & ARMDIA1.CC & ARMDIA1.QQ & sDBValue & ARMDIA1.QQ & ");")

    End Sub
	
	
	Sub PR_AddVertex(ByRef xyPoint As xy, ByRef nBulge As Double)
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
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "  AddVertex(" & "xyStart.x+" & xyPoint.X & Project1.HEADNECK1.CC & "xyStart.y+" & xyPoint.Y & Project1.HEADNECK1.CC & "0,0," & nBulge & ");")

    End Sub

    Sub PR_CalcMidPoint(ByRef xyStart As Project1.HEADNECK1.xy, ByRef xyEnd As Project1.HEADNECK1.xy, ByRef xyMid As Project1.HEADNECK1.xy)

        Static aAngle, nLength As Double

        aAngle = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)

        If nLength = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object xyMid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyMid = xyStart 'Avoid overflow
        Else
            PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
        End If

    End Sub

    Sub PR_CalcPolar(ByRef xyStart As Project1.HEADNECK1.xy, ByVal nAngle As Double, ByVal nLength As Double, ByRef xyReturn As Project1.HEADNECK1.xy)
        'Procedure to return a point at a distance and an angle from a given point
        '
        'PI is a Global Constant

        Static A, B As Double

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
                B = System.Math.Sin(nAngle) * nLength
                A = System.Math.Cos(nAngle) * nLength
                xyReturn.X = xyStart.X + A
                xyReturn.Y = xyStart.Y + B
        End Select

    End Sub

    Sub PR_DrawArc(ByRef xyCen As Project1.HEADNECK1.xy, ByRef xyArcStart As Project1.HEADNECK1.xy, ByRef xyArcEnd As Project1.HEADNECK1.xy)
        'Draws an arc
        'Restrictions
        '    1. Arc must be 180 degrees or less

        Static nAdj, nDeltaAng, nStartAng, nRad, nEndAng, nOpp, nSign As Object
        Static xyMidPoint As Project1.HEADNECK1.xy

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

        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nRad = FN_CalcLength(xyCen, xyArcStart)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = FN_CalcAngle(xyCen, xyArcStart)
        'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd)

        'Direction of arc
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nSign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nSign = System.Math.Sign(nEndAng - nStartAng)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If System.Math.Abs(nEndAng - nStartAng) > 180 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSign = 0 - nSign
        End If

        'Included angle
        PR_CalcMidPoint(xyArcStart, xyArcEnd, xyMidPoint)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj = FN_CalcLength(xyCen, xyMidPoint)
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp = FN_CalcLength(xyArcStart, xyArcEnd) / 2

        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nAdj = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nDeltaAng = 180
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object nAdj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nOpp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nDeltaAng = (System.Math.Atan(nOpp / nAdj) * 2) * (180 / PI)
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nSign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nDeltaAng = nDeltaAng * nSign

        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "arc" & ARMDIA1.QC & "xyStart.x +" & Str(xyCen.X) & ARMDIA1.CC & "xyStart.y +" & Str(xyCen.Y) & ARMDIA1.CC & Str(nRad) & ARMDIA1.CC & Str(nStartAng) & ARMDIA1.CC & Str(nDeltaAng) & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "SetDBData(hEnt," & ARMDIA1.QQ & "ID" & ARMDIA1.QQ & ",sID);")
    End Sub

    Sub PR_DrawFitted(ByRef Profile As Curve)
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
		Dim ii As Short
		
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
				PR_DrawPoly(Profile)
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID);")
                'Warn the user to smooth the curve
                'Not required by Glove
                '           Print #fNum, "Display ("; QQ; "message"; QCQ; "OKquestion"; QCQ; "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line"; QQ; ");"
            Case Else
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, "hEnt = AddEntity(" & Project1.HEADNECK1.QQ & "poly" & Project1.HEADNECK1.QCQ & "fitted" & Project1.HEADNECK1.QQ)
                For ii = 1 To Profile.n
                    'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PrintLine(Project1.CLTXFDIA1.fNum, Project1.HEADNECK1.CC & "xyStart.x+" & Str(Profile.X(ii)) & Project1.HEADNECK1.CC & "xyStart.y+" & Str(Profile.Y(ii)))
                Next ii
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, ");")
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID);")
        End Select
		
	End Sub
	
	Sub PR_DrawLine(ByRef xyStart As xy, ByRef xyFinish As xy)
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
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "hEnt = AddEntity(")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, Project1.HEADNECK1.QQ & "line" & Project1.HEADNECK1.QC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "xyStart.x+" & Str(xyStart.X) & Project1.HEADNECK1.CC & "xyStart.y+" & Str(xyStart.Y) & Project1.HEADNECK1.CC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "xyStart.x+" & Str(xyFinish.X) & Project1.HEADNECK1.CC & "xyStart.y+" & Str(xyFinish.Y))
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID);")


    End Sub
	
	Sub PR_DrawMarker(ByRef xyPoint As xy, ByRef sExtraIDdata As String)
        'Draw a Marker at the given point
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "hEnt = AddEntity(" & Project1.HEADNECK1.QQ & "marker" & Project1.HEADNECK1.QCQ & "xmarker" & Project1.HEADNECK1.QC & "xyStart.x+" & xyPoint.X & Project1.HEADNECK1.CC & "xyStart.y+" & xyPoint.Y & Project1.HEADNECK1.CC & "0.0625);")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID+" & Project1.HEADNECK1.QQ & sExtraIDdata & Project1.HEADNECK1.QQ & ");")

    End Sub
	
	Sub PR_DrawMarkerNamed(ByRef sName As String, ByRef xyPoint As xy, ByRef nWidth As Object, ByRef nHeight As Object, ByRef nAngle As Object)
        'Draw a Marker at the given point
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & Project1.HEADNECK1.QQ & "marker" & Project1.HEADNECK1.QCQ & sName & Project1.HEADNECK1.QC & "xyStart.x+" & xyPoint.X & ARMDIA1.CC & "xyStart.y+" & xyPoint.Y & ARMDIA1.CC & nWidth & ARMDIA1.CC & nHeight & ARMDIA1.CC & nAngle & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID);")

    End Sub
	
	Sub PR_DrawPoly(ByRef Profile As Curve)
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
		Dim ii As Short
		
		'Exit if nothing to draw
		If Profile.n <= 1 Then Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "poly" & ARMDIA1.QCQ & "polyline" & ARMDIA1.QQ)
        For ii = 1 To Profile.n
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, ARMDIA1.CC & "xyStart.x+" & Str(Profile.X(ii)) & ARMDIA1.CC & "xyStart.y+" & Str(Profile.Y(ii)))
        Next ii
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "SetDBData(hEnt," & ARMDIA1.QQ & "ID" & ARMDIA1.QQ & ",sID);")

    End Sub
	
	
	Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As xy, ByRef nHeight As Object)
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
		Dim nWidth As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nWidth = nHeight * ARMDIA1.g_nCurrTextAspect
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "AddEntity(" & ARMDIA1.QQ & "text" & ARMDIA1.QCQ & sText & ARMDIA1.QC & "xyStart.x+" & Str(xyInsert.X) & ARMDIA1.CC & "xyStart.y+" & Str(xyInsert.Y) & ARMDIA1.CC & nWidth & ARMDIA1.CC & nHeight & ARMDIA1.CC & g_nCurrTextAngle & ");")

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
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "EndPoly();")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = UID (""find"", UID (""getmax"")) ;")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "SetDBData(hEnt," & ARMDIA1.QQ & "ID" & ARMDIA1.QQ & ",sID);")

    End Sub
	
	Sub PR_InsertSymbol(ByRef sSymbol As String, ByRef xyInsert As xy, ByRef nScale As Single, ByRef nRotation As Single)
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
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "SetSymbolLibrary(" & ARMDIA1.QQ & FN_EscapeSlashesInString(Project1.HEADNECK1.g_sPathJOBST) & "\\JOBST.SLB" & Project1.HEADNECK1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "Symbol(" & Project1.HEADNECK1.QQ & "find" & Project1.HEADNECK1.QCQ & sSymbol & Project1.HEADNECK1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "hEnt = AddEntity(" & Project1.HEADNECK1.QQ & "symbol" & Project1.HEADNECK1.QCQ & sSymbol & Project1.HEADNECK1.QC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "xyStart.x+" & Str(xyInsert.X) & Project1.HEADNECK1.CC & "xyStart.y+" & Str(xyInsert.Y) & Project1.HEADNECK1.CC)
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, Str(nScale) & Project1.HEADNECK1.CC & Str(nScale) & ARMDIA1.CC & Str(nRotation) & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "SetDBData(hEnt," & Project1.HEADNECK1.QQ & "ID" & Project1.HEADNECK1.QQ & ",sID);")

    End Sub
	
	Sub PR_MakeXY(ByRef xyReturn As xy, ByRef X As Double, ByRef Y As Double)
		'Utility to return a point based on the X and Y values
		'given
		xyReturn.X = X
		xyReturn.Y = Y
		
	End Sub
	
	Sub PR_NamedHandle(ByRef sHandleName As String)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to retain the entity handle of a previously
        'added entity.
        '
        'Assumes that hEnt is the entity handle to the just inserted entity.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "HANDLE " & sHandleName & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, sHandleName & " = hEnt;")

    End Sub
	
	Sub PR_SetEntityData(ByRef sDataType As String, ByRef Param1 As Object, ByRef Param2 As Object, ByRef sParam3 As Object)
		'Procedure to set entity data
		'Assumes that entity pointed at by the DRAFIX entity handle hEnt
		'is the beast we want
		
		Select Case sDataType
			Case "layer"
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, "hLayer = Table(" & ARMDIA1.QQ & "find" & ARMDIA1.QCQ & "layer" & ARMDIA1.QCQ & Param1 & ARMDIA1.QQ & ");")
                'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(Project1.CLTXFDIA1.fNum, "if ( hLayer > %zero && hLayer != 32768)" & "SetEntityData(hEnt, " & ARMDIA1.QQ & "layer" & ARMDIA1.QQ & ",hLayer);")
        End Select
		
	End Sub
	
	Sub PR_SetLayer(ByRef sNewLayer As String)
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

        If ARMDIA1.g_sCurrentLayer = sNewLayer Then Exit Sub
        ARMDIA1.g_sCurrentLayer = sNewLayer

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.CLTXFDIA1.fNum, "hLayer = Table(" & ARMDIA1.QQ & "find" & ARMDIA1.QCQ & "layer" & ARMDIA1.QCQ & sNewLayer & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "if ( hLayer > %zero && hLayer != 32768)" & "Execute (" & ARMDIA1.QQ & "menu" & ARMDIA1.QCQ & "SetLayer" & ARMDIA1.QC & "hLayer);")

    End Sub
	
	Sub PR_SetLineStyle(ByRef sStyle As String)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "Execute (" & ARMDIA1.QQ & "menu" & ARMDIA1.QCQ & "SetStyle" & ARMDIA1.QC & "Table(" & ARMDIA1.QQ & "find" & ARMDIA1.QCQ & "style" & ARMDIA1.QCQ & sStyle & ARMDIA1.QQ & "));")
    End Sub
	
	Sub PR_SetTextData(ByRef nHoriz As Object, ByRef nVert As Object, ByRef nHt As Object, ByRef nAspect As Object, ByRef nFont As Object)
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

        'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nHoriz >= 0 And ARMDIA1.g_nCurrTextHorizJust <> nHoriz Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(Project1.ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextHorzJust" & ARMDIA1.QC & nHoriz & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextHorizJust = nHoriz
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nVert >= 0 And ARMDIA1.g_nCurrTextVertJust <> nVert Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(Project1.ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextVertJust" & ARMDIA1.QC & nVert & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextVertJust = nVert
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nHt >= 0 And ARMDIA1.g_nCurrTextHt <> nHt Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(Project1.ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextHeight" & ARMDIA1.QC & nHt & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextHt = nHt
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nAspect >= 0 And ARMDIA1.g_nCurrTextAspect <> nAspect Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(Project1.ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextAspect" & ARMDIA1.QC & nAspect & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextAspect = nAspect
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nFont >= 0 And ARMDIA1.g_nCurrTextFont <> nFont Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(Project1.ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextFont" & ARMDIA1.QC & nFont & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextFont = nFont
        End If


    End Sub
	
	Sub PR_StartPoly()
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

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "StartPoly (""PolyLine"");")


    End Sub
	
	Function round(ByVal nNumber As Double) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35) = 1
		'    round(1.55) = 2
		'    round(2.50) = 3
		'    round(-2.50) = -3
		'
		
		Static nInt, nSign As Short
		
		nSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		nInt = Int(nNumber)
		If (nNumber - nInt) >= 0.5 Then
			round = (nInt + 1) * nSign
		Else
			round = nInt * nSign
		End If
		
	End Function
End Module
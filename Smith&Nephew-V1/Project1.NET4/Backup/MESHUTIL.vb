Option Strict Off
Option Explicit On
Module MESHUTIL
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
	
	Public MainForm As meshdraw
	
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
	
	
	'PI as a Global Constant
	Public Const PI As Double = 3.141592654
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fNum As Object 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	'Store current layer and text setings to reduce DRAFIX code
	'this value is checked in PR_SetLayer
	Public g_sCurrentLayer As String
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	Public g_nCurrTextAngle As Object
	
	Public g_sSleeveType As String 'Dummy used to kid PR_CalcRaglan ()
	Public g_sPathJOBST As String
	Public g_sSide As String
	Public g_sID As String
	
	'Windows API Functions Declarations
	'    Declare Function GetActiveWindow Lib "User" () As Integer'
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
	Public Const INCH1_16 As Double = 0.0625
	Public Const INCH1_8 As Double = 0.125
	Public Const INCH3_16 As Double = 0.1875
	
	
	
	Function Arccos(ByRef X As Double) As Double
		Arccos = System.Math.Atan(-X / System.Math.Sqrt(-X * X + 1)) + 1.5708
	End Function
	
	Function FN_BiArcCurve(ByRef xyStart As XY, ByRef aStart As Double, ByRef xyEnd As XY, ByRef aEnd As Double, ByRef Profile As BiArc) As Short
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
		Dim nTmp, aAxis, nLength As Double
		Dim Phi1, Theta1, Theta2, Phi2 As Double
		Dim C1, S1, d, B, A, c, P, S2, C2 As Double
		Dim R1, Rs, R2 As Double
		Dim Rmin As Object
		Dim uvR2, uvR1, uvOrigin As XY
		'    Dim fError
		
		'Use a file for printing results for debug
		'Open file
		'    fError = FreeFile
		'    Open "C:\TMP\FIT_ERR.DAT" For Output As fError
		
		
		'Initially return false
		FN_BiArcCurve = False
		
		'Check for silly data and return
		If aStart = aEnd Then GoTo Error_Close
		If xyStart.y = xyEnd.y And xyStart.X = xyEnd.X Then GoTo Error_Close
		
		'Translate tangent angle in XY system to the UV system.
		'Where the U-Axis is specified by the line xyStart,xyEnd
		'
		aAxis = FN_CalCAngle(xyStart, xyEnd)
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
			Phi1 = System.Math.Abs((2 * PI) - Theta1)
		End If
		If Theta2 < PI Then
			Phi2 = Theta2
		Else
			Phi2 = System.Math.Abs((2 * PI) - Theta2)
		End If
		
		'Print #fError, "Phi1="; Phi1
		'Print #fError, "Phi2="; Phi2
		
		
		'Calculate R1 and R2
		S1 = System.Math.Abs(System.Math.Sin(Theta1))
		C1 = (-System.Math.Sin(Theta1) * System.Math.Cos(Theta1)) / S1
		'Print #fError, "S1="; S1; "C1="; C1
		
		S2 = System.Math.Abs(System.Math.Sin(Theta2))
		C2 = (System.Math.Sin(Theta2) * System.Math.Cos(Theta2)) / S2
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
			nTmp = (P * System.Math.Sqrt(nTmp)) / B
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
			A = 2 * System.Math.Sin(Theta1)
			B = (System.Math.Sin(Theta1) ^ 2) - (System.Math.Cos(Theta1) ^ 2) + 1
			R1 = (P * A) / (2 * B)
			R2 = R1
			'Print #fError, "A="; A; "B="; A;
			'Print #fError, "R1="; R1
		End If
		
		'The radi R1 and R2 must be greater than the specified
		'minimum Rmin
		'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rmin = nLength / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If R1 < Rmin Then
			'Print #fError, "R1 < Rmin"
			'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			R1 = Rmin
			d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
			'Print #fError, "D="; d
			R2 = R1 - d
			'Print #fError, "R1="; R1
			'Print #fError, "R2="; R2
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If R2 < Rmin Then
			'Print #fError, "R2 < Rmin"
			'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			R2 = Rmin
			'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
			'Print #fError, "R1="; R1
			'Print #fError, "R2="; R2
		End If
		
		'Check for errors
		If R1 < 0 Or R2 < 0 Then GoTo Error_Close
		
		'Using the calculated radi, create BiArc Curve
		'Start and end points of bi-arc curve
		'UPGRADE_WARNING: Couldn't resolve default property of object Profile.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profile.xyStart = xyStart
		'UPGRADE_WARNING: Couldn't resolve default property of object Profile.xyEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profile.xyEnd = xyEnd
		
		'Centers of arcs
		'Get UV co-ordinates
		PR_MakeXY(uvR1, R1 * S1, R1 * C1)
		'Print #fError, "uvR1="; uvR1.X, uvR1.Y
		'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
		PR_MakeXY(uvR2, P - (R2 * S2), R2 * C2)
		'Print #fError, "uvR2="; uvR2.X, uvR2.Y
		'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
		PR_MakeXY(uvOrigin, 0, 0)
		'Translate to XY co-ordinates
		PR_CalcPolar(xyStart, FN_CalCAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1)
		PR_CalcPolar(xyStart, FN_CalCAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2)
		'Print #fError, "R1="; R1
		'Print #fError, "R2="; R2
		'Tangent point on arc
		If R1 < R2 Then
			PR_CalcPolar(Profile.xyR1, FN_CalCAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent)
		Else
			PR_CalcPolar(Profile.xyR1, FN_CalCAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent)
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
	
	Function FN_CalCAngle(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Function to return the angle between two points in degrees
		'in the range 0 - 360
		'Zero is always 0 and is never 360
		
		Dim X, y As Object
		Dim rAngle As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = xyEnd.X - xyStart.X
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		y = xyEnd.y - xyStart.y
		
		'Horizontal
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If X = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If y > 0 Then
				FN_CalCAngle = 90
			Else
				FN_CalCAngle = 270
			End If
			Exit Function
		End If
		
		'Vertical (avoid divide by zero later)
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If y = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If X > 0 Then
				FN_CalCAngle = 0
			Else
				FN_CalCAngle = 180
			End If
			Exit Function
		End If
		
		'All other cases
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees
		
		If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
		
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If y > 0 Then
			FN_CalCAngle = rAngle
		Else
			FN_CalCAngle = rAngle + 180
		End If
		
	End Function
	
	Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Fuction to return the length between two points
		'Greatfull thanks to Pythagorus
		
		FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)
		
	End Function
	
	Function FN_CirLinInt(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyCen As XY, ByRef nRad As Double, ByRef xyInt As XY) As Short
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
		nSlope = FN_CalCAngle(xyStart, xyEnd)
		
		'Horizontal Line
		'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nSlope = 0 Or nSlope = 180 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nSlope = -1
			'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
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
					xyInt.y = xyStart.y
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
				nRoot = xyCen.y + System.Math.Sqrt(nC) * nSign
				'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If nRoot >= min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
					xyInt.y = nRoot
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
			nM = (xyEnd.y - xyStart.y) / (xyEnd.X - xyStart.X) 'Slope
			'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nK = xyStart.y - nM * xyStart.X 'Y-Axis intercept
			'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nA = (1 + nM ^ 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.y * nM))
			'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
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
					xyInt.y = nM * nRoot + nK
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
	
	Function FN_LinLinInt(ByRef xyLine1Start As XY, ByRef xyLine1End As XY, ByRef xyLine2Start As XY, ByRef xyLine2End As XY, ByRef xyInt As XY) As Short
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
		nSlope1 = FN_CalCAngle(xyLine1Start, xyLine1End)
		If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
		If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nSlope2 = FN_CalCAngle(xyLine2Start, xyLine2End)
		'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
		'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8
		
		Select Case nCase
			
			Case 0
				'Both lines are Non-Orthogonal Lines
				nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
				'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nM2 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
				'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (nM1 = nM2) Then Exit Function 'Parallel lines
				nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
				'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nK2 = xyLine2Start.y - (nM2 * xyLine2Start.X) 'Y-Axis intercept
				If (nK1 = nK2) Then Exit Function
				'Find X
				'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nX = (nK2 - nK1) / (nM1 - nM2)
				'Find Y
				nY = (nM1 * nX) + nK1
				
			Case 1
				'Line 1 is Horizontal or Line 2 is horizontal
				nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
				nK1 = xyLine2Start.y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
				nY = xyLine1Start.y
				'Solve for X at the given Y value
				nX = (nY - nK1) / nM1
				
			Case 2
				'Line 1 is Vertical or Line 2 is Vertical
				nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
				nK1 = xyLine2Start.y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
				nX = xyLine1Start.X
				'Solve for Y at the given X value
				nY = (nM1 * nX) + nK1
				
			Case 4
				'Line 1 is Horizontal or Line 2 is horizontal
				nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
				nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
				nY = xyLine2Start.y
				
				'Solve for X at the given Y value
				nX = (nY - nK1) / nM1
				
			Case 5
				'Parallel orthogonal lines, no intersection possible
				Exit Function
				
			Case 6
				'Line1 is Vertical and the Line2 is Horizontal
				nX = xyLine1Start.X
				nY = xyLine2Start.y
				
			Case 8
				'Line 1 is Vertical or Line 2 is Vertical
				nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
				nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
				nX = xyLine2Start.X
				'Solve for Y at the given X value
				nY = (nM1 * nX) + nK1
				
			Case 9
				'Line1 is Horizontal and the Line2 is Vertical
				nX = xyLine2Start.X
				nY = xyLine1Start.y
				
			Case 10
				'Parallel orthogonal lines, no intersection possible
				Exit Function
				
			Case Else
				Exit Function
				
		End Select
		
		'Ensure that the points X and Y are on the lines
		xyInt.X = nX
		xyInt.y = nY
		
		'Line 1
		'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.X, xyLine1End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.X, xyLine1End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
		'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.y, xyLine1End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.y, xyLine1End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (nY < min(xyLine1Start.y, xyLine1End.y) Or nY > max(xyLine1Start.y, xyLine1End.y)) Then Exit Function
		
		'Line 2
		'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.X, xyLine2End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.X, xyLine2End.X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
		'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.y, xyLine2End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.y, xyLine2End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (nY < min(xyLine2Start.y, xyLine2End.y) Or nY > max(xyLine2Start.y, xyLine2End.y)) Then Exit Function
		
		FN_LinLinInt = True
		
	End Function
	
	Function FN_Open(ByRef sDrafixFile As String) As Short
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_Open = fNum
		
		'Initialise String globals
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(13) 'The new line character
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes (")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCQ = QQ & CC & QQ
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QC = QQ & CC
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CQ = CC & QQ
		
		
		'Globals to reduced drafix code written to file
		g_sCurrentLayer = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHt = 0.125
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAspect = 0.6
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHorizJust = 1 'Left
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextVertJust = 32 'Bottom
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextFont = 0 'BLOCK
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, Mesh Drawing Routine")
		
		'Define DRAFIX variables
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY     xyStart, xyOrigin, xyScale, xyO;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sFileNo, sSide, sID, sPathJOBST;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "ANGLE  aAngle;")
		
		'Text data
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & g_nCurrTextHorizJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & g_nCurrTextVertJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & g_nCurrTextHt & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & g_nCurrTextAspect & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & g_nCurrTextFont & ");")
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
		
		'Set xyStart to 0,0 as we are using absolute co-ords
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.y = 0; xyStart.x = 0;")
		
		'Display Hour Glass Symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Drawing" & QQ & ");")
		
	End Function
	
	
	Function fnGetString(ByVal sString As String, ByRef iIndex As Short, ByRef sDelimiter As String) As String
		'Function to return as a string the iIndexth item in a string
		'that using the given string sDelimiter as the delimiter.
		'EG
		'    sString = "Sam Spade Hello"
		'    sDelimiter = " " {SPACE}
		'    fnGetNumber( sString, 2) = "Spade"
		'
		'If the iIndexth item is not found then return "" to indicate an error.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
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
				sString = Left(sString, iPos - 1)
				fnGetString = sString
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos))
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
		PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & sDBName & QQ & CC & QQ & sDBValue & QQ & ");")
		
	End Sub
	
	Sub PR_AddEntityArc(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
		'Draws an arc with original parameters for AddEntity("arc",...) routine
		
		Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nRad = FN_CalcLength(xyCen, xyArcStart)
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nStartAng = FN_CalCAngle(xyCen, xyArcStart)
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nEndAng = FN_CalCAngle(xyCen, xyArcEnd)
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDeltaAng = System.Math.Abs(nStartAng - nEndAng)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "arc" & QC & "xyStart.x +" & Str(xyCen.X) & CC & "xyStart.y +" & Str(xyCen.y) & CC & Str(nRad) & CC & Str(nStartAng) & CC & Str(nDeltaAng) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_AddFormatedDBValueToLast(ByRef sDBName As String, ByRef sLeadingText As String, ByRef nValue As Double, ByRef sTrailingText As String)
		'The last entity is given by hEnt
		' Format(sDataType, nValue)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) {")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hEnt," & QQ & sDBName & QCQ & sLeadingText & QQ & "+ Format(" & QQ & "length" & QC & nValue & ")+" & QQ & sTrailingText & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "}")
		
	End Sub
	
	Sub PR_AddVertex(ByRef xyPoint As XY, ByRef nBulge As Double)
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
		PrintLine(fNum, "  AddVertex(" & "xyStart.x+" & xyPoint.X & CC & "xyStart.y+" & xyPoint.y & CC & "0,0," & nBulge & ");")
		
	End Sub
	
	Sub PR_CalcMidPoint(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyMid As XY)
		
		Static aAngle, nLength As Double
		
		aAngle = FN_CalCAngle(xyStart, xyEnd)
		nLength = FN_CalcLength(xyStart, xyEnd)
		
		If nLength = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object xyMid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyMid = xyStart 'Avoid overflow
		Else
			PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
		End If
		
	End Sub
	
	Sub PR_CalcPolar(ByRef xyStart As XY, ByVal nAngle As Double, ByVal nLength As Double, ByRef xyReturn As XY)
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
				xyReturn.y = xyStart.y
			Case 180
				xyReturn.X = xyStart.X - nLength
				xyReturn.y = xyStart.y
			Case 90
				xyReturn.X = xyStart.X
				xyReturn.y = xyStart.y + nLength
			Case 270
				xyReturn.X = xyStart.X
				xyReturn.y = xyStart.y - nLength
			Case Else
				'Convert from degees to radians
				nAngle = nAngle * PI / 180
				B = System.Math.Sin(nAngle) * nLength
				A = System.Math.Cos(nAngle) * nLength
				xyReturn.X = xyStart.X + A
				xyReturn.y = xyStart.y + B
		End Select
		
	End Sub
	
	Sub PR_DrawArc(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
		'Draws an arc
		'Restrictions
		'    1. Arc must be 180 degrees or less
		
		Static nAdj, nDeltaAng, nStartAng, nRad, nEndAng, nOpp, nSign As Object
		Static xyMidPoint As XY
		
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
		nStartAng = FN_CalCAngle(xyCen, xyArcStart)
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nEndAng = FN_CalCAngle(xyCen, xyArcEnd)
		
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
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "arc" & QC & "xyStart.x +" & Str(xyCen.X) & CC & "xyStart.y +" & Str(xyCen.y) & CC & Str(nRad) & CC & Str(nStartAng) & CC & Str(nDeltaAng) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_DrawFitted(ByRef Profile As curve)
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
				PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
				'Warn the user to smooth the curve
				'           Print #fNum, "Display ("; QQ; "message"; QCQ; "OKquestion"; QCQ; "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line"; QQ; ");"
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
				For ii = 1 To Profile.n
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
				Next ii
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		End Select
		
	End Sub
	
	Sub PR_DrawLine(ByRef xyStart As XY, ByRef xyFinish As XY)
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
		PrintLine(fNum, "hEnt = AddEntity(")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, QQ & "line" & QC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyStart.X) & CC & "xyStart.y+" & Str(xyStart.y) & CC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyFinish.X) & CC & "xyStart.y+" & Str(xyFinish.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_DrawMarker(ByRef xyPoint As XY)
		'Draw a Marker at the given point
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart.x+" & xyPoint.X & CC & "xyStart.y+" & xyPoint.y & CC & "0.125);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_DrawMarkerNamed(ByRef sName As String, ByRef xyPoint As XY, ByRef nWidth As Object, ByRef nHeight As Object, ByRef nAngle As Object)
		'Draw a Marker at the given point
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & sName & QC & "xyStart.x+" & xyPoint.X & CC & "xyStart.y+" & xyPoint.y & CC & nWidth & CC & nHeight & CC & nAngle & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_DrawPoly(ByRef Profile As curve)
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
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "polyline" & QQ)
		For ii = 1 To Profile.n
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	
	Sub PR_DrawRectangle(ByRef xyLLCorner As XY, ByRef xyURCorner As XY)
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "polygon" & QQ)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, CC & "xyStart.x+" & Str(xyLLCorner.X) & CC & "xyStart.y+" & Str(xyLLCorner.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, CC & "xyStart.x+" & Str(xyLLCorner.X) & CC & "xyStart.y+" & Str(xyURCorner.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, CC & "xyStart.x+" & Str(xyURCorner.X) & CC & "xyStart.y+" & Str(xyURCorner.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, CC & "xyStart.x+" & Str(xyURCorner.X) & CC & "xyStart.y+" & Str(xyLLCorner.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object)
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
		nWidth = nHeight * g_nCurrTextAspect
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC & nWidth & CC & nHeight & CC & g_nCurrTextAngle & ");")
		
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
		PrintLine(fNum, "EndPoly();")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = UID (""find"", UID (""getmax"")) ;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_InsertSymbol(ByRef sSymbol As String, ByRef xyInsert As XY, ByRef nScale As Single, ByRef nRotation As Single)
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
		PrintLine(fNum, "SetSymbolLibrary(" & QQ & FN_EscapeSlashesInString(g_sPathJOBST) & "\\JOBST.SLB" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC)
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, Str(nScale) & CC & Str(nScale) & CC & Str(nRotation) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
	End Sub
	
	Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef y As Double)
		'Utility to return a point based on the X and Y values
		'given
		xyReturn.X = X
		xyReturn.y = y
		
	End Sub
	
	Sub PR_NamedHandle(ByRef sHandleName As String)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to retain the entity handle of a previously
		'added entity.
		'
		'Assumes that hEnt is the entity handle to the just inserted entity.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE " & sHandleName & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sHandleName & " = hEnt;")
		
	End Sub
	
	Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
		
	End Sub
	
	Sub PR_PutNumberAssign(ByRef sVariableName As String, ByRef nAssignedNumber As Object)
		'Procedure to put a number assignment
		'Adds a semi-colon
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAssignedNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & Str(nAssignedNumber) & ";")
		
	End Sub
	
	Sub PR_PutStringAssign(ByRef sVariableName As String, ByRef sAssignedString As Object)
		'Procedure to put a string assignment
		'Encloses String in quotes and adds a semi-colon
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & QQ & sAssignedString & QQ & ";")
		
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
		
		If g_sCurrentLayer = sNewLayer Then Exit Sub
		g_sCurrentLayer = sNewLayer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hLayer = Table(" & QQ & "find" & QCQ & "layer" & QCQ & sNewLayer & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if ( hLayer != %badtable)" & "Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "hLayer);")
		
	End Sub
	
	Sub PR_SetLineStyle(ByRef sStyle As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & sStyle & QQ & "));")
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
		If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & nHoriz & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHorizJust = nHoriz
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & nVert & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextVertJust = nVert
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nHt >= 0 And g_nCurrTextHt <> nHt Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & nHt & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHt = nHt
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & nAspect & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAspect = nAspect
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFont >= 0 And g_nCurrTextFont <> nFont Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & nFont & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextFont = nFont
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
		PrintLine(fNum, "StartPoly (""PolyLine"");")
		
		
	End Sub
End Module
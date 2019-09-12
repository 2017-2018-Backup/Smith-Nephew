Option Strict Off
Option Explicit On
Module GLVEXTEN
	'File:      GLVEXTEN.BAS
	'Purpose:   Extend glove to Elbow and Axilla
	'
	'Version:   1.01
	'Date:      17.Jan.96
	'Author:    Gary George
	'
	'Projects:  CADGLOVE.MAK
	'           MANGLOVE.MAK
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	'
	'   This module is designed to be common to both the
	'   CAD Glove and the Manual Glove.  Therefor we make use
	'   of the global level variables indicated by g_
	'
	'   This may make the structure of the program a little
	'   more obscure but the concept is that all procedures and
	'   related module variables are in the one module GLVEXTEN.BAS
	'
	'
	
	Public Structure TapeData
		Dim nCir As Double
		Dim iMMs As Short
		Dim iRed As Short
		Dim iGms As Short
		Dim sNote As String
		Dim iTapePos As Short
		Dim sTapeText As String
	End Structure
	
	
	Public Structure Curve
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
	
	
	'XY data type to represent points
	Public Structure XY
		Dim X As Double
		Dim Y As Double
	End Structure
	
	
	Public Structure BiArc
		Dim xyStart As XY
		Dim xyTangent As XY
		Dim xyEnd As XY
		Dim xyR1 As XY
		Dim xyR2 As XY
		Dim nR1 As Double
		Dim nR2 As Double
	End Structure
	
	Public Structure Finger
		Dim DIP As Double
		Dim PIP As Double
		'UPGRADE_NOTE: Len was upgraded to Len_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Len_Renamed As Double
		Dim WebHt As Double
		Dim Tip As Short
		Dim xyUlnar As XY
		Dim xyRadial As XY
	End Structure
	
	'Fingers
	Public Const DIP As Short = 1
	Public Const PIP As Short = 2
	Public Const THUMB As Short = 3
	Public Const THUMB_LEN As Short = 2
	Public Const FINGER_LEN As Short = 1
	Public Const g_sTapeText As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
	
	
	'Tapes etc
	Public Const WRIST As Short = 1
	Public Const PALM As Short = 2
	Public Const TAPE_ONE_HALF As Short = 3
	Public Const TAPE_THREE As Short = 4
	Public Const TAPE_FOUR_HALF As Short = 5
	Public Const TAPE_SIX As Short = 6
	Public Const TAPE_SEVEN_HALF As Short = 7
	Public Const TAPE_NINE As Short = 8
	
	'Misc
	Public Const EIGHTH As Double = 0.125
	Public Const SIXTEENTH As Double = 0.0625
	Public Const QUARTER As Double = 0.25
	
	'Reduction Chart constants and Variables
	Const NOFF_FINGER_RED As Short = 32
	Const NOFF_ARM_RED As Short = 76
	Const NOFF_LENGTH_RED As Short = 24
	Const FIGURED_MINIMUM As Short = 6
	
	'Hand Reduction charts
	Dim iFingerRed(3, NOFF_FINGER_RED) As Short
	Dim nArmRed(8, NOFF_ARM_RED) As Double
	Dim nLengthRed(2, NOFF_LENGTH_RED) As Double
	
	'Arm reduction chart
	Public Const LOW_MODULUS As Short = 160
	Public Const HIGH_MODULUS As Short = 340
	Const NOFF_MODULUS As Short = 19
	
	Dim g_iModulusIndex As Short
	Dim g_iPowernet(NOFF_MODULUS, 23) As Short
	
	
	'Arm variables
	Public Const NOFF_ARMTAPES As Short = 16
	Public Const ELBOW_TAPE As Short = 9
	Public Const ARM_PLAIN As Short = 0
	Public Const ARM_FLAP As Short = 1
	Public Const GLOVE_NORMAL As Short = 0
	Public Const GLOVE_ELBOW As Short = 1
	Public Const GLOVE_AXILLA As Short = 2
	
	
	
	
	Function FN_FigureFinger(ByRef iSelect As Short, ByRef nFingerCir As Double, ByRef iInsert As Short) As Double
		'procedure to return the Figured circumference of a finger
		'
		' iSelect    a constant, one of DIP, PIP, THUMB
		' nFingerCir Dimension in inches in
		'            1" >= nFingerCir <= 4.825"
		' iInsert    Insert range to be used
		'            1 >= iNinsert <= 6
		'            iInsert = 1, use looked up figure
		'            iInsert > 1, add 2 to DIP & PIP, add 1 to THUMB
		'                         to looked up figure
		'            iInsert = 6, subtract 1 from DIP, PIP, THUMB
		'                         to looked up figure
		'
		
		Dim iIndex, iValue As Short
		
		'Check data is within range
		FN_FigureFinger = 0#
		If iSelect < DIP Or iSelect > THUMB Then Exit Function
		If iInsert < 1 Or iInsert > 6 Then Exit Function
		If nFingerCir < 1 Or nFingerCir > 4.875 Then Exit Function
		
		'Convert nFingerCir to an index
		'Note integer division \
		iIndex = (((nFingerCir - 1) * 1000) \ 125) + 1
		
		'Look up value
		iValue = iFingerRed(iSelect, iIndex)
		
		If iInsert = 6 Then
			iValue = iValue + 1
		ElseIf iSelect = THUMB Then 
			iValue = iValue - (iInsert - 1)
		Else
			iValue = iValue - ((iInsert - 1) * 2)
		End If
		
		If iValue < FIGURED_MINIMUM Then iValue = FIGURED_MINIMUM
		
		'Convert from 16ths to inches
		FN_FigureFinger = iValue * 0.0625
		
	End Function
	
	
	
	
	Function FN_FigureLength(ByRef iSelect As Short, ByRef nLen As Double) As Double
		'Procedure to return the Figured length of a finger
		'or thumb
		'
		' iSelect    A constant, one of THUMB or FINGER
		' nLen       Dimension in inches in
		'            1" >= nLen <= 3.875
		
		Dim iIndex, iValue As Short
		
		'Check data is within range
		FN_FigureLength = 0#
		If iSelect <> THUMB_LEN And iSelect <> FINGER_LEN Then Exit Function
		If nLen >= 1 And nLen <= 3.875 Then
			'Convert nLen to an index
			'Note integer division \
			iIndex = (((nLen - 1) * 1000) \ 125) + 1
			'Look up value
			FN_FigureLength = nLengthRed(iSelect, iIndex)
			
		ElseIf nLen < 1 Then 
			'Return given value unmodified
			FN_FigureLength = nLen
			
		ElseIf nLen > 3.875 And iSelect = THUMB_LEN Then 
			FN_FigureLength = nLen - 0.125
			
		ElseIf nLen > 3.875 And iSelect = FINGER_LEN Then 
			FN_FigureLength = nLen - 0.25
			
		End If
		
		
	End Function
	
	Function FN_FigureTape(ByRef iSelect As Short, ByRef nTapeCir As Double) As Double
		'Procedure to return the Figured circumference of an ARM tape
		'This also is used to figure the Palm and wrist
		'
		' iSelect    A constant, one off
		'            WRIST, PALM, TAPE_ONE_HALF ... TAPE_NINE
		' nTapeCir   Dimension in inches in
		'            3.5" >= nTapeCir <= 12.875"
		
		'NOTES;
		'    The PALM and WRIST chart only goes up to 12.375
		'    if this value is exceeded then a value is extrapolated
		'    from the 12.875 value.
		'
		
		Dim iIndex, iValue As Short
		Dim iAdditional As Double
		
		'Check data is within range
		FN_FigureTape = 0#
		If iSelect < WRIST Or iSelect > TAPE_NINE Then Exit Function
		If nTapeCir < 3.5 Then Exit Function
		
		'Convert nTapeCir to an index
		'Note integer division \
		If nTapeCir > 12.875 Then
			'Extrapolate a value for this case
			'Get index of the last value in the table (ie 12.875)
			iIndex = (((12.875 - 3.5) * 1000) \ 125) + 1
			'Look up value
			'For each 1/2" greater than 12.875 add an eighth
			'to the returned value
			iAdditional = ((nTapeCir - 12.875) \ 0.5) + 1
			
			FN_FigureTape = nArmRed(iSelect, iIndex) + (iAdditional * 0.125)
			
		Else
			iIndex = (((nTapeCir - 3.5) * 1000) \ 125) + 1
			'Look up value
			FN_FigureTape = nArmRed(iSelect, iIndex)
		End If
		
	End Function
	
	Function FN_GetGrams(ByVal iReduction As Short) As Short
		'Function that looks up the "Powernet" grams / tension chart
		'and returns the Grams for a given reduction value
		'
		'Used to back calculate from a given reduction
		'
		' Input
		'        iReduction      Reduction given
		'
		' Module Level variables
		'        g_iModulusIndex
		'        g_iPowernet()
		'
		If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
			FN_GetGrams = -1
			Exit Function
		End If
		
		If iReduction < 10 Then iReduction = 10
		If iReduction > 32 Then iReduction = 32
		
		FN_GetGrams = g_iPowernet(g_iModulusIndex, (iReduction - 10) + 1)
		
		
	End Function
	
	Function FN_GetReduction(ByRef iGrams As Object) As Short
		'Function that looks up the "Powernet" grams / tension chart
		'and returns the reduction value
		'
		' Input
		'        iGrams      Grams calculated from the data
		'
		' Module Level variables
		'        g_iModulusIndex
		'        g_iPowernet()
		'
		If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
			FN_GetReduction = -1
			Exit Function
		End If
		
		Dim iValue, ii, iPrevValue As Short
		
		Select Case iGrams
			Case Is <= g_iPowernet(g_iModulusIndex, 1)
				FN_GetReduction = 10
				
			Case Is >= g_iPowernet(g_iModulusIndex, 23)
				FN_GetReduction = 32
				
			Case Else
				'Return value closest
				iPrevValue = 0
				For ii = 1 To 23
					iValue = g_iPowernet(g_iModulusIndex, ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object iGrams. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If iValue > iGrams Then Exit For
					iPrevValue = iValue
				Next ii
				
				'UPGRADE_WARNING: Couldn't resolve default property of object iGrams. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(iGrams - iPrevValue) < System.Math.Abs(iGrams - iValue) Then
					FN_GetReduction = ii + 8
				Else
					FN_GetReduction = ii + 9
				End If
				
		End Select
		
		
	End Function
	
	Function FN_LengthWristToEOS() As Double
		'Calculates the distance from the EOS to the
		'wrist based on values given in the dialogue
		'
		Static ii As Short
		Static nLen, nSpace As Double
		
		'Get the data from the dialogue
		PR_GetDlgAboveWrist()
		
		nLen = 0
		
		Select Case g_ExtendTo
			Case GLOVE_NORMAL
				For ii = 1 To g_iNumTapesWristToEOS - 1
					nSpace = 1.375
					'           If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
					'           If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
					nLen = nLen + nSpace
				Next ii
				If nLen = 0 Then
					'If no tapes after wrist then default to 0.75
					nLen = 0.75
				End If
			Case GLOVE_ELBOW
				For ii = 1 To g_iNumTapesWristToEOS - 1
					nSpace = 1.375
					If ii = 1 And g_nPleats(1) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
					If ii = 2 And g_nPleats(2) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
					If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
					If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
					nLen = nLen + nSpace
				Next ii
				
			Case GLOVE_AXILLA
				For ii = 1 To g_iNumTapesWristToEOS - 1
					nSpace = 1.375
					If ii = 1 And g_nPleats(1) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
					If ii = 2 And g_nPleats(2) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
					If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
					If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
					nLen = nLen + nSpace
				Next ii
				If g_EOSType = ARM_FLAP Then
				End If
		End Select
		
		FN_LengthWristToEOS = nLen
		
	End Function
	
	Function FN_ValidateExtensionData() As String
		'Procedure to validated the extension data
		Static sError, NL As String
		Static iNoPleats, ii As Short
		
		NL = Chr(13)
		sError = "" 'Initialise because of static
		
		'If normal glove then exit
		If g_ExtendTo = GLOVE_NORMAL Then
			FN_ValidateExtensionData = ""
			Exit Function
		End If
		
		'Check pleats
		iNoPleats = 0
		For ii = 1 To 4
			If g_nPleats(ii) > 0 Then iNoPleats = iNoPleats + 1
		Next ii
		If iNoPleats + 1 > g_iNumTapesWristToEOS Then
			sError = sError & "Number of pleats exceeds availble spaces between the arm tapes!" & NL & "Disable pleats by Double Clicking on pleat label." & NL
		End If
		
		'Check that calculate has been used
		If g_CalculatedExtension = False Then
			sError = sError & "The arm has not been calculated" & NL
		End If
		
		'Check on style D flaps
		If g_EOSType = ARM_FLAP And InStr(1, g_sFlapType, "D") > 0 And g_nWaistCir = 0 Then
			sError = sError & "A Waist circumference must be given for a D-Style flap" & NL
		End If
		
		'Return error message
		FN_ValidateExtensionData = sError
		
		
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
		If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then GoTo Error_Close
		
		'Translate tangent angle in XY system to the UV system.
		'Where the U-Axis is specified by the line xyStart,xyEnd
		'
		aAxis = ARMDIA1.FN_CalcAngle(xyStart, xyEnd)
		nLength = ARMDIA1.FN_CalcLength(xyStart, xyEnd)
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
		
		P = ARMDIA1.FN_CalcLength(xyStart, xyEnd)
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
		ARMDIA1.PR_MakeXY(uvR1, R1 * S1, R1 * C1)
		'Print #fError, "uvR1="; uvR1.X, uvR1.Y
		'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
		ARMDIA1.PR_MakeXY(uvR2, P - (R2 * S2), R2 * C2)
		'Print #fError, "uvR2="; uvR2.X, uvR2.Y
		'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
		ARMDIA1.PR_MakeXY(uvOrigin, 0, 0)
		'Translate to XY co-ordinates
		ARMDIA1.PR_CalcPolar(xyStart, ARMDIA1.FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1)
		ARMDIA1.PR_CalcPolar(xyStart, ARMDIA1.FN_CalcAngle(uvOrigin, uvR2) + aAxis, ARMDIA1.FN_CalcLength(uvOrigin, uvR2), Profile.xyR2)
		'Print #fError, "R1="; R1
		'Print #fError, "R2="; R2
		'Tangent point on arc
		If R1 < R2 Then
			ARMDIA1.PR_CalcPolar(Profile.xyR1, ARMDIA1.FN_CalcAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent)
		Else
			ARMDIA1.PR_CalcPolar(Profile.xyR1, ARMDIA1.FN_CalcAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent)
		End If
		'radi of arcs
		Profile.nR1 = R1
		Profile.nR2 = R2
		
		'Test that the Tangent point lies between the start and end points
		If ARMDIA1.FN_CalcLength(xyStart, Profile.xyTangent) > nLength Then GoTo Error_Close
		If ARMDIA1.FN_CalcLength(xyEnd, Profile.xyTangent) > nLength Then GoTo Error_Close
		
		'return true as we have a sucessful fit
		'Close #fError
		FN_BiArcCurve = True
		
		Exit Function
		
Error_Close: 
		'Print #fError, "Error and close"
		'Close #fError
		
	End Function
	
	Sub PR_CalcFingerBlendingProfile(ByRef xyTop As XY, ByRef xyBottom As XY, ByRef xyFinger As XY, ByRef Digit As Finger, ByRef ReturnProfile As Curve, ByRef xyThumbPalm As XY)
		'Procedure to return a smooth inflecting curve between two points
		'
		'The assumptions about the required curve are :-
		'    1. The Top Point is above the Bottom Point
		'    2. The curve leaves the Top point at 270
		'    3. The curve enters the bottom point at 270
		'    4. The curve is aproximated by two arcs
		'       that are tangential at the mid point
		'       between xyTop and xyBottom
		'    5. The curve inflects at the point in 4 above
		'    6. The radius of the arcs are the same
		'
		'The returned profile has the following features :-
		'If there is no missing finger then
		'    1. Decreases in Y from xyTop to xyBottom
		'    2. Has 7 vertex
		'    3. Vertex = 1 = xyTop
		'    4. Vertex = 7 = xyBottom
		'    5. Where xyTop.X = xyBottom.X returns a
		'       two point profile.
		'If there is a missing finger then
		'    1. Decreases in Y from xyTop to xyBottom
		'    2. Has 6 vertex
		'    3. Vertex = 1 = xyFinger
		'    4. Vertex = 6 = xyBottom
		'    4. Vertex = 2 to 5 create a fillet
		'       that can be manually edited
		'
		'Notes :-
		'    This is a fairly restricted routine designed
		'    to blend the profile above the wrist into the
		'    glove with a smooth inflecting curve.
		'
		'Modifications :-
		'    The point given in xyThumbPalm is the current point
		'    of the global variable xyThumbPalm(1) this is modified
		'    to the intersection of the blended curve at the height
		'    given by xyThumbPalm.y
		'
		'
		Static aInc, aAngle, nLength, rAngle, nRadius As Double
		Static xyPt1, xyCenter, xyMidPoint, xyPt2 As XY
		Static xyConstruct, xyInt As XY
		Static iDirection As Short
		Static ii, MirrorResult As Short
		
		nLength = ARMDIA1.FN_CalcLength(xyBottom, xyTop)
		aAngle = ARMDIA1.FN_CalcAngle(xyBottom, xyTop)
		
		'If the angle is less than 90 Degrees then we just carry on.
		'However if the angle is greater, then we mirror this angle in the
		'y axis and calculate the points as befor, then we mirror the result
		'
		MirrorResult = False
		'UPGRADE_WARNING: Couldn't resolve default property of object xyConstruct. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyConstruct = xyTop
		If aAngle > 90 Then
			MirrorResult = True
			aAngle = 90 - (aAngle - 90)
			xyConstruct.X = xyBottom.X - (xyTop.X - xyBottom.X)
		End If
		
		'Degenerate to a straight line if xyTop.X = xyBottom.X
		'Then exit the sub routine
		If System.Math.Abs(aAngle) = 90 Or aAngle = 270 Then
			ReturnProfile.n = 2
			ReturnProfile.X(1) = xyTop.X
			ReturnProfile.Y(1) = xyTop.Y
			ReturnProfile.X(2) = xyBottom.X
			ReturnProfile.Y(2) = xyBottom.Y
			GoTo MissingFingers
		Else
			ReturnProfile.n = 7
			ReturnProfile.X(1) = xyConstruct.X
			ReturnProfile.Y(1) = xyConstruct.Y
			ReturnProfile.X(7) = xyBottom.X
			ReturnProfile.Y(7) = xyBottom.Y
		End If
		
		'Get midpoint
		ARMDIA1.PR_CalcPolar(xyBottom, aAngle, nLength / 2, xyMidPoint)
		ReturnProfile.X(4) = xyMidPoint.X
		ReturnProfile.Y(4) = xyMidPoint.Y
		
		'Calculate radius
		rAngle = aAngle * (PI / 180) 'Convert angle to Radians
		nRadius = (nLength / 4) / System.Math.Cos(rAngle)
		
		'First arc points
		ARMDIA1.PR_MakeXY(xyCenter, xyConstruct.X - nRadius, xyConstruct.Y)
		aAngle = ARMDIA1.FN_CalcAngle(xyCenter, xyMidPoint)
		aInc = (360 - aAngle) / 3
		
		'Find Intersection
		If xyThumbPalm.Y > xyMidPoint.Y And xyThumbPalm.Y < xyTop.Y Then
			ARMDIA1.PR_MakeXY(xyPt1, xyCenter.X, xyThumbPalm.Y)
			ARMDIA1.PR_MakeXY(xyPt2, xyCenter.X + 10, xyThumbPalm.Y)
			If BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyThumbPalm = xyInt
				If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
			End If
		End If
		
		For ii = 3 To 2 Step -1
			aAngle = aAngle + aInc
			ARMDIA1.PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
			ReturnProfile.X(ii) = xyPt1.X
			ReturnProfile.Y(ii) = xyPt1.Y
		Next ii
		
		'Second arc points
		ARMDIA1.PR_MakeXY(xyCenter, xyBottom.X + nRadius, xyBottom.Y)
		aAngle = ARMDIA1.FN_CalcAngle(xyCenter, xyMidPoint)
		
		'Find Intersection
		If xyThumbPalm.Y <= xyMidPoint.Y And xyThumbPalm.Y > xyBottom.Y Then
			ARMDIA1.PR_MakeXY(xyPt1, xyCenter.X, xyThumbPalm.Y)
			ARMDIA1.PR_MakeXY(xyPt2, xyCenter.X - 10, xyThumbPalm.Y)
			If BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyThumbPalm = xyInt
				If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
			End If
		End If
		
		For ii = 5 To 6
			aAngle = aAngle + aInc
			ARMDIA1.PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
			ReturnProfile.X(ii) = xyPt1.X
			ReturnProfile.Y(ii) = xyPt1.Y
		Next ii
		
		'Mirroring the result simplifies the code above, as we only have
		'to code for the case where angle < 90
		'N.B. Mirroring in the Y axis along the line X = xyBottom.X
		If MirrorResult Then
			For ii = 1 To ReturnProfile.n
				ReturnProfile.X(ii) = xyBottom.X - (ReturnProfile.X(ii) - xyBottom.X)
			Next ii
		End If
		
MissingFingers: 
		
		Static nX, nY As Double
		
		If Digit.Len_Renamed = 0 And (Digit.Tip = 0 Or Digit.Tip = 10) Then
			
			nY = xyFinger.Y - ReturnProfile.Y(2)
			nX = xyFinger.X - ReturnProfile.X(1)
			iDirection = System.Math.Sign(xyFinger.X - ReturnProfile.X(1))
			nX = System.Math.Abs(nX)
			'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nRadius = BDYUTILS.min(nX, nY) - 0.05
			
			'Straight line
			ReturnProfile.n = 6
			ReturnProfile.X(1) = xyFinger.X
			ReturnProfile.Y(1) = xyFinger.Y
			ARMDIA1.PR_MakeXY(xyCenter, xyTop.X + (iDirection * nRadius), xyFinger.Y - nRadius)
			aAngle = 90
			aInc = (90 / 3) * iDirection
			For ii = 2 To 5
				ARMDIA1.PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
				ReturnProfile.X(ii) = xyPt1.X
				ReturnProfile.Y(ii) = xyPt1.Y
				aAngle = aAngle + aInc
			Next ii
			ReturnProfile.X(6) = xyBottom.X
			ReturnProfile.Y(6) = xyBottom.Y
			
		End If
		
	End Sub
	
	
	Sub PR_CalculateArmTapeReductions()
		'MsgBox "PR_CalculateArmTapeReductions"
		'Procedure to calculate the reductions
		'at each arm tape
		Dim ii As Short
		
		'Don't Calculate if normal glove
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then Exit Sub
		
		If CType(MainForm.Controls("cboFabric"), Object).Text = "" Then
			MsgBox("Fabric not given! ", 48, "Manual Glove - Calculate ARM Button")
			Exit Sub
		Else
			If UCase(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 1, 3)) <> "POW" Then
				MsgBox("Fabric chosen is not Powernet!", 48, "Manual Glove - Calculate ARM Button")
				Exit Sub
			End If
			If Val(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 5, 3)) < LOW_MODULUS Or Val(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 5, 3)) > HIGH_MODULUS Then
				MsgBox("Modulus of fabric chosen is not available on Gram / Tension reduction chart!", 48, "Manual Glove - Calculate ARM Button")
				Exit Sub
			End If
			
		End If
		
		
		'Update from dialogue
		PR_GetDlgAboveWrist()
		'PR_PrintDlgAboveWrist
		If g_iPressure < 0 Then
			MsgBox("Can't Calculate, missing Pressure", 48, "Manual Glove - Calculate ARM Button")
			Exit Sub
		End If
		
		If g_iNumTapesWristToEOS <= 1 Then
			MsgBox("Can't Calculate, Only one tape between given Wrist and EOS", 48, "Manual Glove - Calculate ARM Button")
			Exit Sub
		End If
		
		'Check that there are no holes in the data
		For ii = g_iWristPointer To g_iEOSPointer
			If g_nCir(ii) = 0 Then
				MsgBox("Can't Calculate, missing Tape Circumferences between given Wrist and EOS", 48, "Manual Glove - Calculate ARM Button")
				Exit Sub
			End If
		Next ii
		
		
		'Set the MMs based on the selected pressure
		PR_SetMMs((CType(MainForm.Controls("cboPressure"), Object).Text))
		
		'Set the modulus based on the fabric
		PR_SetModulusIndex((CType(MainForm.Controls("cboFabric"), Object).Text))
		
		'Calculate grams etc
		For ii = 1 To NOFF_ARMTAPES
			If ii >= g_iWristPointer And ii <= g_iEOSPointer Then
				
				g_iGms(ii) = ARMDIA1.round(g_nCir(ii) * g_iMMs(ii))
				g_iRed(ii) = FN_GetReduction(g_iGms(ii))
				
				'Adjust at wrist
				If ii = g_iWristPointer Then
					'The wrist must lie between 10 and 14
					'From the given reduction we back calculate the grams and mms
					If g_iRed(ii) > 14 Then g_iRed(ii) = 14
					g_iGms(ii) = FN_GetGrams(g_iRed(ii))
					g_iMMs(ii) = ARMDIA1.round(g_iGms(ii) / g_nCir(ii))
				End If
				
				'Adjust at EOS
				If ii = g_iEOSPointer Then
					'Always use a 10 reduction for plain EOS
					'For Flaps always use a 12
					'From the given reduction we back calculate the grams and mms
					If g_EOSType = ARM_PLAIN Then
						g_iRed(ii) = 10
					Else
						g_iRed(ii) = 12
					End If
					g_iGms(ii) = FN_GetGrams(g_iRed(ii))
					g_iMMs(ii) = ARMDIA1.round(g_iGms(ii) / g_nCir(ii))
				End If
				
				'Display Results
				PR_GrdInchesDisplay(ii - 1, g_nCir(ii))
				PR_GramRedDisplay(ii - 1, g_iGms(ii), g_iRed(ii))
				CType(MainForm.Controls("mms"), Object)(ii + 7).Text = CStr(g_iMMs(ii))
				
			Else
				'Blank the displayed values
				PR_GrdInchesDisplay(ii - 1, 0)
				PR_GramRedDisplay(ii - 1, 0, 0)
				CType(MainForm.Controls("mms"), Object)(ii + 7).Text = ""
			End If
		Next ii
		
		g_CalculatedExtension = True
	End Sub
	
	Sub PR_CalculateExtension(ByRef xyWristUlnar As XY, ByRef xyWristRadial As XY, ByRef nInsert As Double)
		
		'Procedure to calculate the POINTS
		'used to draw the extension of the glove above the
		'wrist
		'The Procedure also supplies the data that is to be
		'printed at each tape
		'N.B.
		'The wrist points will be modified if the wrist has been
		'calculated
		
		Dim nValue, nMidWristX, nSpacing As Double
		Dim iStartTape, ii, iLastTape As Short
		Dim nFiguredValue As Double
		Dim iVertex As Short
		
		nMidWristX = xyWristUlnar.X + (System.Math.Abs(xyWristRadial.X - xyWristUlnar.X) / 2)
		
		If g_iNumTapesWristToEOS = 0 Or g_iNumTapesWristToEOS = 1 Then
			'Extension ends at wrist
			UlnarProfile.n = 1
			UlnarProfile.X(1) = xyWristUlnar.X
			UlnarProfile.Y(1) = xyWristUlnar.Y - 0.75
			
			RadialProfile.n = 1
			RadialProfile.X(1) = xyWristRadial.X
			RadialProfile.Y(1) = xyWristRadial.Y - 0.75
			
			Exit Sub
			
		End If
		
		If g_ExtendTo = GLOVE_NORMAL Then
			'Glove only extends two or three tapes past wrist
			'Use the charts to get the reductions
			'
			'Although given as part of the Extension
			'The normal glove uses data from the Hand Part of the dialogue
			'
			'
			UlnarProfile.n = g_iNumTapesWristToEOS
			UlnarProfile.X(1) = xyWristUlnar.X
			UlnarProfile.Y(1) = xyWristUlnar.Y
			
			RadialProfile.n = g_iNumTapesWristToEOS
			RadialProfile.X(1) = xyWristRadial.X
			RadialProfile.Y(1) = xyWristRadial.Y
			
			'God knows why this is but it works, and I don't give a !"£$%^&*(
			'any more
			iStartTape = 2
			iLastTape = g_iNumTapesWristToEOS + 1
			
			For ii = iStartTape To iLastTape
				
				'NB we add 1 to ii as the charts have wrist at ii=1, palm at ii=2
				'and 1-1/2 tape at ii=3.
				'The palm being out of order is a pain but it made it easier to
				'enter the chart data (swings and roundabouts, sorry!)
				'
				'NB use of 1/2 scale
				nFiguredValue = FN_FigureTape(ii + 1, g_nCir(ii))
				
				Select Case g_iInsertStyle
					Case 0
						If g_OnFold Then
							nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
						Else
							nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
						End If
					Case 1
						nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
					Case 2, 3
						nValue = (nFiguredValue + (2 * EIGHTH)) / 2
				End Select
				
				If g_OnFold Then
					UlnarProfile.X(ii) = UlnarProfile.X(1)
					RadialProfile.X(ii) = UlnarProfile.X(1) + nValue
				Else
					UlnarProfile.X(ii) = nMidWristX - (nValue / 2)
					RadialProfile.X(ii) = nMidWristX + (nValue / 2)
				End If
				
				'Standard spacing
				nSpacing = 1.375
				
				'We are working from the top (wrist) down
				UlnarProfile.Y(ii) = UlnarProfile.Y(ii - 1) - nSpacing
				RadialProfile.Y(ii) = RadialProfile.Y(ii - 1) - nSpacing
				
			Next ii
			
		End If
		
		
		If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_AXILLA Then
			'As we have recalculated the wrist we need to start at the
			'wrist
			UlnarProfile.n = g_iNumTapesWristToEOS
			UlnarProfile.Y(1) = xyWristUlnar.Y
			
			RadialProfile.n = g_iNumTapesWristToEOS
			RadialProfile.Y(1) = xyWristRadial.Y
			
			iStartTape = g_iWristPointer
			iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)
			
			'Note: As we can allow the wrist to be any tape we need a
			'vertex counter
			iVertex = 1
			
			For ii = iStartTape To iLastTape
				'NB use of 1/2 scale
				nFiguredValue = (g_nCir(ii) * ((100 - g_iRed(ii)) / 100))
				Select Case g_iInsertStyle
					Case 0
						If g_OnFold Then
							nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
						Else
							nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
						End If
					Case 1
						nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
					Case 2, 3
						nValue = (nFiguredValue + (2 * EIGHTH)) / 2
				End Select
				
				If g_OnFold Then
					UlnarProfile.X(iVertex) = UlnarProfile.X(1)
					RadialProfile.X(iVertex) = UlnarProfile.X(1) + nValue
				Else
					UlnarProfile.X(iVertex) = nMidWristX - (nValue / 2)
					RadialProfile.X(iVertex) = nMidWristX + (nValue / 2)
				End If
				
				
				'Setup the notes for the vertex
				'We do this here as we have all the data and it simplifies
				'the drawing side
				TapeNote(iVertex).sTapeText = LTrim(Mid(g_sTapeText, ((ii + 1) * 3) + 1, 3))
				TapeNote(iVertex).iTapePos = ii
				TapeNote(iVertex).nCir = g_nCir(ii)
				TapeNote(iVertex).iGms = g_iGms(ii)
				TapeNote(iVertex).iRed = g_iRed(ii)
				TapeNote(iVertex).iMMs = g_iMMs(ii)
				
				
				'Standard spacing
				nSpacing = 1.375
				
				'Account all for pleats
				'Wrist (as we have started at the wrist we use, iStartTape + 1
				'and iStartTape + 2 in this case)
				If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
				If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
				'XXXXXXXXXX
				'XXXX Careful now! this won't work if only 3 or less tapes given
				'XXXXXXXXXX
				'Axilla
				If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
				If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
				
				'We are working from the top (wrist) down
				'iVertex = 1 is set before the For Loop (Values from wrist)
				If iVertex <> 1 Then
					UlnarProfile.Y(iVertex) = UlnarProfile.Y(iVertex - 1) - nSpacing
					RadialProfile.Y(iVertex) = RadialProfile.Y(iVertex - 1) - nSpacing
					If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
				End If
				
				'Increment vertex count
				iVertex = iVertex + 1
				
			Next ii
			
			'Reset given wrist points  xyWristUlnar and xyWristRadial
			ARMDIA1.PR_MakeXY(xyWristUlnar, UlnarProfile.X(1), UlnarProfile.Y(1))
			ARMDIA1.PR_MakeXY(xyWristRadial, RadialProfile.X(1), RadialProfile.Y(1))
		End If
		
	End Sub
	
	Sub PR_CalcWristBlendingProfile(ByRef xyTSt As XY, ByRef xyTEnd As XY, ByRef xyBSt As XY, ByRef xyBEnd As XY, ByRef ReturnProfile As Curve, ByRef xyThumbPalm As XY)
		'Procedure to return a smooth inflecting curve between two points
		'
		'N.B. Parameters are given in the order of decreasing Y
		'
		Static nR1, rAngle, aA2, nL3, nL1, nL2, aA1, aA3, aInc, nR2 As Double
		Static xyBotSt, xyTopSt, xyCenter, xyMidPoint, xyTopEnd, xyBotEnd As XY
		Static xyR1, xyR2 As XY
		Static nA, nThirdOfL2, aAngle As Double
		Static xyPt2, xyPt1, xyInt As XY
		Static TopIsArc, MirrorResult, ii, Direction, BottomIsArc As Short
		Static Intersection As Double
		'UPGRADE_WARNING: Lower bound of array xyTmp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static xyTmp(10) As XY
		Static nTol As Double
		
		'Do this as we can't use ByVal
		'UPGRADE_WARNING: Couldn't resolve default property of object xyTopSt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyTopSt = xyTSt
		'UPGRADE_WARNING: Couldn't resolve default property of object xyTopEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyTopEnd = xyTEnd
		'UPGRADE_WARNING: Couldn't resolve default property of object xyBotSt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyBotSt = xyBSt
		'UPGRADE_WARNING: Couldn't resolve default property of object xyBotEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyBotEnd = xyBEnd
		
		'If the angle is less than 90 Degrees then we just carry on.
		'However if the angle is greater, then we mirror this angle in the
		'y axis and calculate the points as befor, then we mirror the result
		'
		MirrorResult = False
		aA2 = ARMDIA1.FN_CalcAngle(xyBSt, xyTEnd)
		aA3 = ARMDIA1.FN_CalcAngle(xyBEnd, xyBSt)
		Direction = 1
		If (aA2 > 90) Or ((aA2 = 90) And (aA3 < 90)) Then
			MirrorResult = True
			Direction = -1
			aA2 = 90 - (aA2 - 90)
			xyTopSt.X = xyBEnd.X - (xyTopSt.X - xyBEnd.X)
			xyTopEnd.X = xyBEnd.X - (xyTopEnd.X - xyBEnd.X)
			xyBotSt.X = xyBEnd.X - (xyBotSt.X - xyBEnd.X)
		End If
		aA1 = ARMDIA1.FN_CalcAngle(xyTopEnd, xyTopSt)
		aA3 = ARMDIA1.FN_CalcAngle(xyBotEnd, xyBotSt)
		
		'Degenerate to a straight line
		'Then exit the sub routine
		If aA1 = aA2 And aA2 = aA3 And (aA1 = 90 Or aA1 = 270) Then
			ReturnProfile.n = 2
			ReturnProfile.X(1) = xyTSt.X
			ReturnProfile.Y(1) = xyTSt.Y
			ReturnProfile.X(2) = xyBEnd.X
			ReturnProfile.Y(2) = xyBEnd.Y
			Exit Sub
		End If
		
		nL1 = ARMDIA1.FN_CalcLength(xyTEnd, xyTSt)
		nL2 = ARMDIA1.FN_CalcLength(xyBSt, xyTEnd)
		nL3 = ARMDIA1.FN_CalcLength(xyBEnd, xyBSt)
		
		'Get Included angles & radius & Centers of Arcs
		nThirdOfL2 = nL2 / 3
		
		'Top Arc
		If aA1 <> aA2 Then
			'Calculate the points for the ARC for this section
			TopIsArc = True
			nA = ARMDIA1.FN_CalcLength(xyTopSt, xyBotSt)
			nA = ((nL1 ^ 2 + nL2 ^ 2) - nA ^ 2) / (2 * nL1 * nL2)
			rAngle = BDYUTILS.Arccos(nA) / 2
			nR1 = nThirdOfL2 * System.Math.Tan(rAngle)
			ARMDIA1.PR_CalcPolar(xyTopEnd, 90, nThirdOfL2, xyTmp(2))
			ARMDIA1.PR_CalcPolar(xyTopEnd, ARMDIA1.FN_CalcAngle(xyTopEnd, xyBotSt), nThirdOfL2, xyTmp(5))
			ARMDIA1.PR_CalcPolar(xyTmp(2), 180, nR1, xyR1)
			
			'Top Arc Points
			aAngle = ARMDIA1.FN_CalcAngle(xyR1, xyTmp(5))
			aInc = (360 - aAngle) / 3
			For ii = 4 To 3 Step -1
				aAngle = aAngle + aInc
				ARMDIA1.PR_CalcPolar(xyR1, aAngle, nR1, xyTmp(ii))
			Next ii
		Else
			'Calculate the points for the STRAIGHT LINE for this section
			'Note: we use the same noff points for consistancy
			'with respect to the editor and subsequent points
			TopIsArc = False
			ARMDIA1.PR_CalcPolar(xyTopEnd, 90, nThirdOfL2, xyTmp(2))
			ARMDIA1.PR_CalcPolar(xyTopEnd, 90, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(3))
			ARMDIA1.PR_CalcPolar(xyTopEnd, 270, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(4))
			ARMDIA1.PR_CalcPolar(xyTopEnd, 270, nThirdOfL2, xyTmp(5))
		End If
		
		If aA2 <> aA3 Then
			'Bottom arc
			BottomIsArc = True
			nA = ARMDIA1.FN_CalcLength(xyTopEnd, xyBotEnd)
			nA = ((nL2 ^ 2 + nL3 ^ 2) - nA ^ 2) / (2 * nL3 * nL2)
			rAngle = BDYUTILS.Arccos(nA) / 2
			nR2 = nThirdOfL2 * System.Math.Tan(rAngle)
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))
			
			'Establish
			aAngle = ARMDIA1.FN_CalcAngle(xyBotEnd, xyTopEnd)
			If aAngle < ARMDIA1.FN_CalcAngle(xyBotEnd, xyBotSt) Then
				aAngle = ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd) - 90
			Else
				aAngle = ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd) + 90
			End If
			ARMDIA1.PR_CalcPolar(xyTmp(6), aAngle, nR2, xyR2)
			
			'Check that the gap between the wrist point and the arc is less than
			'or equal to 0.0625 ie. 1/16"
			'If not then make it so.
			nTol = 0.0625
			nA = ARMDIA1.FN_CalcLength(xyBotSt, xyR2)
			If (nA - nR2) > nTol Then
				nR2 = nR2 - ((nA - nR2) - nTol)
				ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
				ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))
				'aAngle from above
				ARMDIA1.PR_CalcPolar(xyTmp(6), aAngle, nR2, xyR2)
			End If
			
			'Bottom arc points
			aAngle = ARMDIA1.FN_CalcAngle(xyR2, xyTmp(6))
			aInc = (ARMDIA1.FN_CalcAngle(xyR2, xyTmp(9)) - aAngle) / 3
			For ii = 7 To 8
				aAngle = aAngle + aInc
				ARMDIA1.PR_CalcPolar(xyR2, aAngle, nR2, xyTmp(ii))
			Next ii
			
		Else
			BottomIsArc = False
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyTopEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(7))
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyBotEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(8))
			ARMDIA1.PR_CalcPolar(xyBotSt, ARMDIA1.FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))
		End If
		
		'Find Intersection
		Intersection = False
		If xyThumbPalm.Y >= xyTmp(2).Y Then
			nA = xyThumbPalm.Y - xyTmp(2).Y
			If nA = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyThumbPalm = xyTmp(2)
			Else
				ARMDIA1.PR_MakeXY(xyThumbPalm, xyTmp(2).X + (nA / System.Math.Tan(aA1 * (PI / 180))), xyThumbPalm.Y)
			End If
			Intersection = True
			
		ElseIf xyThumbPalm.Y < xyTmp(2).Y And xyThumbPalm.Y > xyTmp(5).Y Then 
			If TopIsArc Then
				ARMDIA1.PR_MakeXY(xyPt1, xyR1.X, xyThumbPalm.Y)
				ARMDIA1.PR_MakeXY(xyPt2, xyR1.X + 10, xyThumbPalm.Y)
				If BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyR1, nR1, xyInt) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyThumbPalm = xyInt
					Intersection = True
				End If
			Else
				nA = xyThumbPalm.Y - xyTmp(5).Y
				If nA = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyThumbPalm = xyTmp(5)
				Else
					ARMDIA1.PR_MakeXY(xyThumbPalm, xyTmp(5).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.Y)
				End If
				Intersection = True
			End If
			
		ElseIf xyThumbPalm.Y <= xyTmp(5).Y And xyThumbPalm.Y >= xyTmp(6).Y Then 
			nA = xyThumbPalm.Y - xyTmp(6).Y
			If nA = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyThumbPalm = xyTmp(6)
			Else
				ARMDIA1.PR_MakeXY(xyThumbPalm, xyTmp(6).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.Y)
			End If
			Intersection = True
			
		ElseIf xyThumbPalm.Y < xyTmp(6).Y And xyThumbPalm.Y > xyTmp(9).Y Then 
			If BottomIsArc Then
				ARMDIA1.PR_MakeXY(xyPt1, xyTmp(6).X + nR2, xyThumbPalm.Y)
				ARMDIA1.PR_MakeXY(xyPt2, xyTmp(6).X - nR2, xyThumbPalm.Y)
				If BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyR2, nR2, xyInt) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyThumbPalm = xyInt
					Intersection = True
				End If
			Else
				nA = xyThumbPalm.Y - xyTmp(9).Y
				If nA = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyThumbPalm = xyTmp(9)
				Else
					ARMDIA1.PR_MakeXY(xyThumbPalm, xyTmp(9).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.Y)
				End If
				Intersection = True
			End If
		End If
		If Intersection And MirrorResult Then xyThumbPalm.X = xyBotEnd.X - (xyThumbPalm.X - xyBotEnd.X)
		
		ReturnProfile.n = 10
		ReturnProfile.X(1) = xyTopSt.X
		ReturnProfile.Y(1) = xyTopSt.Y
		ReturnProfile.X(10) = xyBotEnd.X
		ReturnProfile.Y(10) = xyBotEnd.Y
		
		For ii = 2 To 9
			ReturnProfile.X(ii) = xyTmp(ii).X
			ReturnProfile.Y(ii) = xyTmp(ii).Y
		Next ii
		
		
		'Mirroring the result simplifies the code above, as we only have
		'to code for the case where angle < 90
		'N.B. Mirroring in the Y axis along the line X = xyBottom.X
		If MirrorResult Then
			For ii = 1 To ReturnProfile.n
				ReturnProfile.X(ii) = xyBotEnd.X - (ReturnProfile.X(ii) - xyBotEnd.X)
			Next ii
		End If
		
	End Sub
	
	Sub PR_DisableFigureArm()
		
		Dim ii As Short
		
		For ii = 3 To 11
			CType(MainForm.Controls("lblArm"), Object)(ii).Enabled = False
		Next ii
		
		'Pleats
		g_nPleats(1) = Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text)
		g_nPleats(2) = Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text)
		g_nPleats(3) = Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text)
		g_nPleats(4) = Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text)
		
		For ii = 0 To 3
			CType(MainForm.Controls("lblPleat"), Object)(ii).Enabled = False
		Next ii
		
		
		CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = False
		CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = False
		'  MainForm!txtShoulderPleat1 = ""
		'  MainForm!txtShoulderPleat2 = ""
		CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = False
		CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = False
		'  MainForm!txtWristPleat1 = ""
		'  MainForm!txtWristPleat2 = ""
		
		CType(MainForm.Controls("frmCalculate"), Object).Enabled = False
		CType(MainForm.Controls("cboPressure"), Object).Enabled = False
		CType(MainForm.Controls("cmdCalculate"), Object).Enabled = False
		
	End Sub
	
	Sub PR_DisableGloveToAxilla()
		
		CType(MainForm.Controls("frmGloveToAxilla"), Object).Enabled = False
		CType(MainForm.Controls("optProximalTape"), Object)(0).Checked = False
		CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = False
		CType(MainForm.Controls("optProximalTape"), Object)(0).Enabled = False
		CType(MainForm.Controls("optProximalTape"), Object)(1).Enabled = False
		PR_ProximalTape_Click((0))
		
	End Sub
	
	Sub PR_DisplayTextInches(ByRef ctlText As System.Windows.Forms.Control, ByRef ctlCaption As System.Windows.Forms.Control)
		Static nLen As Double
		nLen = CADGLOVE1.FN_InchesValue(ctlText)
		If nLen <> -1 Then ctlCaption.Text = ARMEDDIA1.fnInchestoText(nLen)
	End Sub
	
	Sub PR_DrawShoulderFlaps(ByRef xyS As XY, ByRef xyE As XY)
		'Draws shoulder flaps
		'Extracted from ARMDIA.FRM
		'A load of £%&**!&*^$><@~, but it works (just about!)
		'The only reason to keep it, is that it is production proven
		'with the arm.
		'
		'We calculate the curves and point as they would be
		'attached to an arm curve.
		'We then rotate and mirror the points to fit the
		'drawing of a Glove to axilla
		'
		'UPGRADE_WARNING: Arrays in structure RaglanFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure ShoulderFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static ShoulderFlap, RaglanFlap As Curve
		Static LastTapeValue As Object
		Static xyPt1, xyPt2 As XY
		Static xyTopFlap4, xyTopFlap2, xyTopFlap1, xyTopFlap3, xyTopFlap5 As XY
		Static xyTopFlap9, xyTopFlap7, xyTopFlap6, xyTopFlap8, xyTopFlap10 As XY
		Static xyBotFlap4, xyBotFlap2, xyBotFlap1, xyBotFlap3, xyBotFlap5 As XY
		Static xyFlap4, xyFlap2, xyFlap1, xyFlap3, xyFlap5 As XY
		Static Alpha, Theta, Phi, Beta, Omega, Delta, BottomRadius, TopArcIncrement As Object
		Static TopRadius, LittleBit, opp, h1, x1, y1, adj, hyp, Change, FlapMarkerAngle As Object
		Static xyFlapText1 As XY
		Static xyFlapText2, xyCentre, xyMid, xyFlapMarker, xyStrapText As XY
		Static FlapLength, FlapMarkerLength As Object
		Static TopArcAngle, TopArcRadius, TopArcLength, ArcLength, CircleCircum, BottomArcLength, BottomArcRadius, BottomArcAngle As Object
		Static xfablen, TempMarker, xfabby, xfabdist As Object
		Static xyRaglan6, xyRaglan4, xyRaglan2, xyRaglan1, xyRaglan3, xyRaglan5, xyRaglan7 As XY
		Static TopRagLanAngle, RaglanBottom, RagAng2, Rag4, Rag2, Rag1, Rag3, RagAng1, RagAng3, RaglanTip, InnerRaglanTip As Object
		Static nTemplateAngle, nNotchOffset, nTemplateRadius, nNotchToTangent As Double
		Static ii As Short
		Static RagAng5, PhiExtra, RagAng4, aMarker As Double
		Static Strap As String
		Static aAngle, nLength As Double
		Static xyStrt, xyMidPoint, xyTmp As XY
		
		
		BDYUTILS.PR_CalcMidPoint(xyS, xyE, xyMidPoint)
		nLength = ARMDIA1.FN_CalcLength(xyS, xyE)
		aAngle = -90
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BottomArcAngle = 47.5
		'All this crap because you can't use ByVal on user defined types (Bastards!)
		If g_sSide = "Right" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object xyStrt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyStrt = xyE
			aMarker = 135
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object xyStrt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyStrt = xyS
			aMarker = 45
		End If
		
		
		' Check for Custom Flap Length
		If Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LastTapeValue = ARMEDDIA1.fnDisplayToInches(CDbl(CType(MainForm.Controls("txtCustFlapLength"), Object).Text))
		Else
			'Standard flap length
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LastTapeValue = g_nCir(g_iEOSPointer)
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LastTapeValue = (LastTapeValue / 3.14) * 0.92
		End If
		
		
		' Calculate Bottom part of curve with 5 Points
		If nLength >= 5.5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BottomRadius = LastTapeValue - 0.625
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanBottom = 3.5
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TopRagLanAngle = 17
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanTip = 1.9375
			'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			InnerRaglanTip = 2.173
			nNotchOffset = 0.875 '14/16 ths Raglan Template s287
		ElseIf (nLength < 5.5) And (nLength > 3) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BottomRadius = 2.8
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanBottom = 2.5
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TopRagLanAngle = 17
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanTip = 1.05
			'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			InnerRaglanTip = 1.22
			nNotchOffset = 0.75 '12/16 ths Raglan Template s289
		ElseIf (nLength <= 3) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BottomRadius = 2
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanBottom = 2.0625
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RaglanTip = 0.625
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TopRagLanAngle = 15
			'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			InnerRaglanTip = 0.727
			nNotchOffset = 0.6875 '11/16 ths, Raglan Template s290
		End If
		
		nTemplateRadius = 3.5
		nTemplateAngle = 40
		nNotchToTangent = 2.125
		
		'Calc length of bottom Arc
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BottomArcRadius = BottomRadius
		'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CircleCircum = PI * (2 * (BottomArcRadius))
		'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BottomArcLength = (BottomArcAngle / 360) * CircleCircum
		
		'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ShoulderFlap.X(1) = xyStrt.X + LastTapeValue
		ShoulderFlap.Y(1) = 0
		
		For ii = 1 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Theta = (ii * 11.875 * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Cos(Theta) * BottomRadius
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Sin(Theta) * BottomRadius
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LittleBit = BottomRadius - adj
			'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShoulderFlap.X(ii + 1) = xyStrt.X + LastTapeValue - LittleBit
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShoulderFlap.Y(ii + 1) = 0 + opp
		Next ii
		ARMDIA1.PR_MakeXY(xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.Y(5))
		
		'Calculate 6 angles and points for top part of curve
		ARMDIA1.PR_MakeXY(xyTopFlap1, xyStrt.X, xyStrt.Y + nLength)
		ShoulderFlap.X(10) = xyTopFlap1.X
		ShoulderFlap.Y(10) = xyTopFlap1.Y
		
		'midpoint
		'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		h1 = ARMDIA1.FN_CalcLength(xyTopFlap1, xyBotFlap5)
		'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		h1 = h1 / 2
		'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Omega = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Omega = 180 - Omega
		'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Omega = (Omega * PI) / 180
		'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		opp = System.Math.Abs(h1 * System.Math.Sin(Omega))
		'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		adj = System.Math.Abs(h1 * System.Math.Cos(Omega))
		'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xyMid, xyBotFlap5.X - adj, xyBotFlap5.Y + opp)
		
		'Centre of Top Arc
		'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Delta = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Delta = Delta - 90
		'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Delta = (Delta * PI) / 180
		'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hyp = 4 * h1
		'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		opp = System.Math.Abs(hyp * System.Math.Sin(Delta))
		'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		adj = System.Math.Abs(hyp * System.Math.Cos(Delta))
		'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xyCentre, xyMid.X + adj, xyMid.Y + opp)
		'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopRadius = ARMDIA1.FN_CalcLength(xyCentre, xyTopFlap1)
		'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopArcRadius = TopRadius
		'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Phi = ARMDIA1.FN_CalcAngle(xyCentre, xyTopFlap1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Phi = Phi - 180
		'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Alpha = ARMDIA1.FN_CalcAngle(xyCentre, xyBotFlap5)
		'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Alpha = Alpha - 180
		'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopArcAngle = Alpha - Phi
		'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopArcIncrement = (Alpha - Phi) / 5
		
		'Calc Flap Length for position of marker
		'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CircleCircum = PI * (2 * (TopArcRadius))
		'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopArcLength = (TopArcAngle / 360) * CircleCircum
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FlapLength = TopArcLength + BottomArcLength
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FlapMarkerLength = FlapLength / 2.7
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
		'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If FlapMarkerLength < TopArcLength Then
			'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
			'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlapMarkerAngle = FlapMarkerAngle + Phi
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(TopRadius * System.Math.Sin(FlapMarkerAngle))
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(TopRadius * System.Math.Cos(FlapMarkerAngle))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyFlapMarker, xyCentre.X - adj, xyCentre.Y - opp)
			'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf FlapMarkerLength > TopArcLength Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempMarker = FlapMarkerLength - TopArcLength
			'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempMarker = BottomArcLength - TempMarker
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(BottomArcRadius * System.Math.Sin(FlapMarkerAngle))
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(BottomArcRadius * System.Math.Cos(FlapMarkerAngle))
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LittleBit = BottomRadius - adj
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyFlapMarker, xyStrt.X + LastTapeValue - LittleBit, 0 + opp)
			'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf FlapMarkerLength = TopArcLength Then 
			ARMDIA1.PR_MakeXY(xyFlapMarker, xyBotFlap5.X, xyBotFlap5.Y)
		End If
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		For ii = 1 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PhiExtra = Phi + (ii * TopArcIncrement)
			PhiExtra = (PhiExtra * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(TopRadius * System.Math.Sin(PhiExtra))
			'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(TopRadius * System.Math.Cos(PhiExtra))
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShoulderFlap.X(10 - ii) = xyCentre.X - adj
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShoulderFlap.Y(10 - ii) = xyCentre.Y - opp
		Next ii
		
		
		If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
			'Copy standard shoulder from profile down
			'only go one vertex past the FlapMarker
			RaglanFlap.n = 0
			For ii = 10 To 1 Step -1
				RaglanFlap.n = RaglanFlap.n + 1
				RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
				RaglanFlap.Y(RaglanFlap.n) = ShoulderFlap.Y(ii)
				If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
			Next ii
			CADUTILS.PR_RotateCurve(xyStrt, aAngle, RaglanFlap)
			If g_sSide = "Left" Then CADUTILS.PR_MirrorCurveInYaxis(xyMidPoint.X, 0, RaglanFlap)
			ARMDIA1.PR_DrawFitted(RaglanFlap)
		Else
			ShoulderFlap.n = 10
			CADUTILS.PR_RotateCurve(xyStrt, aAngle, ShoulderFlap)
			If g_sSide = "Left" Then CADUTILS.PR_MirrorCurveInYaxis(xyMidPoint.X, 0, ShoulderFlap)
			ARMDIA1.PR_DrawFitted(ShoulderFlap)
		End If
		
		'Draw Raglan
		If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag1 = System.Math.Sqrt((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag1 = BottomRadius - Rag1
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan1, xyStrt.X + LastTapeValue - Rag1, nNotchOffset)
			ARMDIA1.PR_MakeXY(xyRaglan2, xyRaglan1.X - nNotchToTangent, 0)
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan2, xyRaglan1)
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag2 = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag2 = Rag2 / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng1 = (RagAng1 * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(Rag2 * System.Math.Sin(RagAng1))
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(Rag2 * System.Math.Cos(RagAng1))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan3, xyRaglan2.X + adj, xyRaglan2.Y + opp)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan3, xyRaglan1)
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng1 = 90 - RagAng1
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng1 = (RagAng1 * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag3 = ARMDIA1.FN_CalcLength(xyRaglan3, xyRaglan1)
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rag4 = System.Math.Sqrt((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(Rag4 * System.Math.Sin(RagAng1))
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(Rag4 * System.Math.Cos(RagAng1))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan4, xyRaglan3.X - adj, xyRaglan3.Y + opp)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng2 = (nTemplateAngle * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(RaglanBottom * System.Math.Sin(RagAng2))
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(RaglanBottom * System.Math.Cos(RagAng2))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan5, xyRaglan1.X + adj, xyRaglan1.Y + opp)
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RagAng3 = ARMDIA1.FN_CalcAngle(xyRaglan5, xyFlapMarker)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If RagAng3 > 180 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RagAng3 = RagAng3 - 180
				'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
				'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf RagAng3 <= 180 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RagAng3 = RagAng3 - 90
				'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RagAng4 = 90 - RagAng3
				'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
			End If
			
			RagAng5 = RagAng4 - 13
			RagAng4 = (RagAng4 * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(RaglanTip * System.Math.Sin(RagAng4))
			'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(RaglanTip * System.Math.Cos(RagAng4))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan6, xyRaglan5.X - adj, xyRaglan5.Y + opp)
			
			'Draw Front line of tip
			RagAng5 = (RagAng5 * PI) / 180
			'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			opp = System.Math.Abs(InnerRaglanTip * System.Math.Sin(RagAng5))
			'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			adj = System.Math.Abs(InnerRaglanTip * System.Math.Cos(RagAng5))
			'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyRaglan7, xyRaglan5.X - adj, xyRaglan5.Y + opp)
			
			
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan1)
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan2)
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan4)
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan5)
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan6)
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan7)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyTmp = xyFlapMarker
			CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyTmp)
			
			If g_sSide = "Left" Then
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan1)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan2)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan4)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan5)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan6)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan7)
				CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyTmp)
			End If
			
			ARMDIA1.PR_DrawArc(xyRaglan4, xyRaglan1, xyRaglan2)
			ARMDIA1.PR_DrawLine(xyRaglan1, xyRaglan5)
			ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan6)
			ARMDIA1.PR_DrawLine(xyRaglan6, xyTmp)
			ARMDIA1.PR_Setlayer("Notes")
			ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan7)
			
		End If
		
		ARMDIA1.PR_Setlayer("Notes")
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
		ARMDIA1.PR_MakeXY(xyStrapText, xyMidPoint.X, xyMidPoint.Y - (2 * EIGHTH))
		
		If Val(CType(MainForm.Controls("txtStrapLength"), Object).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Strap = "STRAP = " & Trim(ARMEDDIA1.fnInchestoText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)))) & "\" & QQ
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Strap = "STRAP = 24\" & QQ
		End If
		If Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text) > 0 Then
			'Display FRONT strap length if given
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Strap = "BACK  STRAP = " & Trim(ARMEDDIA1.fnInchestoText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)))) & "\" & QQ
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Strap = Strap & "\nFRONT STRAP = " & Trim(ARMEDDIA1.fnInchestoText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text)))) & "\" & QQ
		End If
		If Val(CType(MainForm.Controls("txtWaistCir"), Object).Text) > 0 Then
			'Display Waist Circumference for D-Style flaps only
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Strap = Strap & "\nWAIST = " & Trim(ARMEDDIA1.fnInchestoText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtWaistCir"), Object).Text)))) & "\" & QQ
		End If
		
		ARMDIA1.PR_DrawText(Strap, xyStrapText, 0.1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xfabby. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xfabby = CType(MainForm.Controls("cboFlaps"), Object).Text
		'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xfablen = Len(xfabby)
		'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xfabdist = (xfablen * 0.075) + 0.1
		'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xyFlapText1, xyFlapMarker.X - (xfabdist * 0.75), xyFlapMarker.Y - 0.5)
		CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyFlapText1)
		If g_sSide = "Left" Then CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyFlapText1)
		ARMDIA1.PR_DrawText((CType(MainForm.Controls("cboFlaps"), Object).Text), xyFlapText1, 0.1)
		
		'Flap Marker
		CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyFlapMarker)
		If g_sSide = "Left" Then CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyFlapMarker)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & "closed arrow" & QC & "xyStart.x +" & Str(xyFlapMarker.X) & CC & "xyStart.y +" & Str(xyFlapMarker.Y) & ",0.25,0.1," & aMarker & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
		'draw the fold line
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		ARMDIA1.PR_MakeXY(xyPt1, xyStrt.X, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xyPt2, xyStrt.X + LastTapeValue, 0)
		
		CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyPt1)
		CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyPt2)
		If g_sSide = "Left" Then
			CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyPt1)
			CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyPt2)
		End If
		
		If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
			ARMDIA1.PR_DrawLine(xyPt1, xyRaglan2)
		Else
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
		End If
		
	End Sub
	
	Sub PR_EnableFigureArm()
		
		'Converse of PR_DisableFigureArm
		
		Static ii As Short
		
		For ii = 3 To 11
			CType(MainForm.Controls("lblArm"), Object)(ii).Enabled = True
		Next ii
		
		For ii = 0 To 3
			CType(MainForm.Controls("lblPleat"), Object)(ii).Enabled = True
		Next ii
		
		CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = True
		CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = True
		If g_nPleats(1) > 0 Then CType(MainForm.Controls("txtWristPleat1"), Object).Text = CStr(g_nPleats(1))
		If g_nPleats(2) > 0 Then CType(MainForm.Controls("txtWristPleat2"), Object).Text = CStr(g_nPleats(2))
		CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = True
		CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = True
		If g_nPleats(3) > 0 Then CType(MainForm.Controls("txtShoulderPleat2"), Object).Text = CStr(g_nPleats(3))
		If g_nPleats(4) > 0 Then CType(MainForm.Controls("txtShoulderPleat1"), Object).Text = CStr(g_nPleats(4))
		
		CType(MainForm.Controls("frmCalculate"), Object).Enabled = True
		CType(MainForm.Controls("cboPressure"), Object).Enabled = True
		CType(MainForm.Controls("cboPressure"), Object).SelectedIndex = g_iPressure
		CType(MainForm.Controls("cmdCalculate"), Object).Enabled = True
		
	End Sub
	
	Sub PR_ExtendTo_Click(ByRef Index As Short)
		Static iStart, iFirst, ii, iLast, iElbow As Short
		
		iLast = 23
		iFirst = 8
		iElbow = 16 'Elbow w.r.t. txtExtCir()
		iStart = 8
		
		Select Case Index
			Case 0 'Normal Glove
				'Disable tapes above those required
				For ii = iStart To iLast
					CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = False
					CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = False
					PR_GrdInchesDisplay(ii - 8, 0)
				Next ii
				
				'Disable all MMs fields etc
				For ii = iFirst To iLast
					CType(MainForm.Controls("mms"), Object)(ii).Enabled = False
					CType(MainForm.Controls("mms"), Object)(ii).Text = ""
					PR_GramRedDisplay(ii - 8, 0, 0)
				Next ii
				
				PR_DisableFigureArm()
				PR_DisableGloveToAxilla()
				
				'set disable and disable
				CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = -1
				CType(MainForm.Controls("cboDistalTape"), Object).Enabled = False
				CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = -1
				CType(MainForm.Controls("cboProximalTape"), Object).Enabled = False
				
			Case 1 'Glove to Elbow
				'Enable MMs fields etc
				For ii = iFirst To iLast
					If ii <= iElbow Then
						CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = True
						CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
						CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = True
					Else
						CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = False
						CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = False
						CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
						CType(MainForm.Controls("mms"), Object)(ii).Text = ""
						PR_GramRedDisplay(ii - 8, 0, 0)
						PR_GrdInchesDisplay(ii - 8, 0)
					End If
				Next ii
				
				'Enable other bits
				PR_EnableFigureArm()
				PR_DisableGloveToAxilla()
				
				'Set Wrist and EOS pointers
				If g_iNumTapesWristToEOS > 0 Then
					'Set wrist to given tape
					If g_iWristPointer > g_iFirstTape Then
						CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = g_iWristPointer
					Else
						'set to first tape
						CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
					End If
					
					'Set EOS to given tape or to Elbow if it extends
					'past the elbow
					'NB. The order of tapes is reversed in this list
					'starting at 19-1/2 and finishing at 0
					If g_iEOSPointer < ELBOW_TAPE Then
						If g_iEOSPointer >= g_iLastTape Then
							CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
						Else
							CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - g_iEOSPointer
						End If
					ElseIf g_iEOSPointer > ELBOW_TAPE Or g_iLastTape > ELBOW_TAPE Then 
						CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - ELBOW_TAPE
					Else
						'set to last tape
						CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
					End If
				Else
					CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
					CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
				End If
				
				CType(MainForm.Controls("cboDistalTape"), Object).Enabled = True
				CType(MainForm.Controls("cboProximalTape"), Object).Enabled = True
				
			Case 2 'Glove to Axilla
				For ii = iFirst To iLast
					CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
					CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = True
					CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = True
					'                MainForm!mms(ii) = ""
					'                PR_GramRedDisplay ii - 8, 0, 0
				Next ii
				
				'Enable other bits
				PR_EnableFigureArm()
				
				CType(MainForm.Controls("frmGloveToAxilla"), Object).Enabled = True
				If g_EOSType = ARM_FLAP Then CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Else CType(MainForm.Controls("optProximalTape"), Object)(0).Checked = True
				
				
				CType(MainForm.Controls("optProximalTape"), Object)(0).Enabled = True
				CType(MainForm.Controls("optProximalTape"), Object)(1).Enabled = True
				
				If g_iNumTapesWristToEOS > 0 Then
					'Set wrist to given tape
					If g_iWristPointer > g_iFirstTape Then
						CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = g_iWristPointer
					Else
						'set to first tape
						CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
					End If
					
					'Set EOS to given tape
					'NB. The order of tapes is reversed in this list
					'starting at 19-1/2 and finishing at 0
					If g_iEOSPointer < g_iLastTape Then
						CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - g_iEOSPointer
					Else
						'set to first tape
						CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
					End If
				Else
					For ii = iFirst To iLast
						CType(MainForm.Controls("mms"), Object)(ii).Text = ""
						PR_GramRedDisplay(ii - 8, 0, 0)
					Next ii
					CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
					CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
				End If
				
				CType(MainForm.Controls("cboDistalTape"), Object).Enabled = True
				CType(MainForm.Controls("cboProximalTape"), Object).Enabled = True
				
		End Select
		
	End Sub
	Sub PR_PrintDlgAboveWrist()
		'Debug routine only
		Dim sMessage As String
		Dim ii As Short
		'Prints
		'    g_nCir(1, NOFF_ARMTAPES) As Double
		'    g_iFirstTape            As Integer
		'    g_iLastTape             As Integer
		'    g_iWristPointer         As Integer
		'    g_iEOSPointer           As Integer
		'    g_iNumTotalTapes        As Integer
		'    g_iNumTapesWristToEOS   As Integer
		'    g_EOSType               As Integer
		'    g_OnFold                As Integer
		'    g_ExtendTo              As Integer
		'    g_nPleats(1 To 4)       As Double
		'    g_iPressure             As Integer
		For ii = 1 To NOFF_ARMTAPES
			sMessage = sMessage & "g_nCir(" & Str(ii) & ")=" & Str(g_nCir(ii)) & Chr(13)
		Next ii
		MsgBox(sMessage)
		sMessage = ""
		sMessage = "g_iFirstTape=" & Str(g_iFirstTape) & Chr(13)
		sMessage = sMessage & "g_iLastTape=" & Str(g_iLastTape) & Chr(13)
		sMessage = sMessage & "g_iWristPointer=" & Str(g_iWristPointer) & Chr(13)
		MsgBox(sMessage)
		
	End Sub
	Sub PR_GetDlgAboveWrist()
		'MsgBox "PR_GetDlgAboveWrist"
		'General procedure to read the data given in the
		'dialogue controls and copy to module level variables.
		'Saves having to do this more than once.
		'
		'Updates:-
		'
		'    g_nCir(1, NOFF_ARMTAPES) As Double
		'    g_iFirstTape            As Integer
		'    g_iLastTape             As Integer
		'    g_iWristPointer         As Integer
		'    g_iEOSPointer           As Integer
		'    g_iNumTotalTapes        As Integer
		'    g_iNumTapesWristToEOS   As Integer
		'    g_EOSType               As Integer
		'    g_OnFold                As Integer
		'    g_ExtendTo              As Integer
		'    g_nPleats(1 To 4)       As Double
		'    g_iPressure             As Integer
		'
		'NOTE:
		'    We ignore the fact that there may be missing tapes.
		'
		'    g_iWristPointer and g_iEOSPointer are used to indicate the
		'    start and finish tapes in the arrays and not in the
		'    txtExtCir() array
		
		Dim nCir As Double
		Dim ii As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("SSTab1"), Object).Tab <> 1 Then CType(MainForm.Controls("SSTab1"), Object).Tab = 1
		
		g_iFirstTape = -1
		g_iLastTape = -1
		g_iWristPointer = -1
		g_iEOSPointer = -1
		g_iNumTotalTapes = 0
		g_iNumTapesWristToEOS = 0
		
		'Glove type
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then
			g_ExtendTo = GLOVE_NORMAL
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CType(MainForm.Controls("optExtendTo"), Object)(1).Value = True Then 
			g_ExtendTo = GLOVE_ELBOW
		Else
			g_ExtendTo = GLOVE_AXILLA
		End If
		
		
		
		Dim iIndex As Short
		If g_ExtendTo = GLOVE_NORMAL Then
			'Get values from Hand Data TAB
			'We will need to update this w.r.t CAD Glove later
			'Assummes no holes in data
			If g_nCir(2) <= 0 Then
				g_iNumTotalTapes = 1
				g_iNumTapesWristToEOS = 1
			ElseIf g_nCir(3) > 0 Then 
				g_iNumTotalTapes = 3
				g_iNumTapesWristToEOS = 3
			Else
				g_iNumTotalTapes = 2
				g_iNumTapesWristToEOS = 2
			End If
			'        g_iWristPointer = 1
			g_iWristPointer = 0
			g_iEOSPointer = 1000 'Stupid value to fool stupid code
			'in PR_ExtendTo_Click
		Else
			'Get First and last tape (if any)
			For ii = 8 To 23
				nCir = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtExtCir"), Object)(ii))
				If nCir > 0 And g_iFirstTape = -1 Then g_iFirstTape = ii - 7
				g_nCir(ii - 7) = nCir
			Next ii
			
			For ii = 23 To 8 Step -1
				If Val(CType(MainForm.Controls("txtExtCir"), Object)(ii).ToString()) > 0 Then
					g_iLastTape = ii - 7
					Exit For
				End If
			Next ii
			
			'Total number of tapes
			If (g_iLastTape = g_iFirstTape) Then
				If g_iLastTape <> -1 Then g_iNumTotalTapes = 1
			Else
				g_iNumTotalTapes = (g_iLastTape - g_iFirstTape) + 1
			End If
			'Get values from "Above Wrist" TAB
			'Wrist tape
			'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
			iIndex = CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex
			If iIndex = 0 Or iIndex = -1 Then
				g_iWristPointer = g_iFirstTape
			Else
				g_iWristPointer = iIndex
			End If
			
			'EOS tape
			'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
			iIndex = CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex
			If iIndex = 0 Or iIndex = -1 Then
				g_iEOSPointer = g_iLastTape
			Else
				'NB. The order of tapes is reversed in this list
				'starting at 19-1/2 and finishing at 0
				g_iEOSPointer = 17 - iIndex
			End If
			
			
			'Number of tapes between
			If (g_iWristPointer = g_iEOSPointer) Then
				'Just in case there are no tapes
				'And "first" and "last" are used for Wrist and EOS
				If g_iLastTape <> -1 Then g_iNumTapesWristToEOS = 1
			Else
				g_iNumTapesWristToEOS = (g_iEOSPointer - g_iWristPointer) + 1
			End If
			
			'Pleats
			If CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = True Then g_nPleats(1) = Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) Else g_nPleats(1) = 0
			If CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = True Then g_nPleats(2) = Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) Else g_nPleats(2) = 0
			If CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = True Then g_nPleats(3) = Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) Else g_nPleats(3) = 0
			If CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = True Then g_nPleats(4) = Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) Else g_nPleats(4) = 0
			
			g_iPressure = CType(MainForm.Controls("cboPressure"), Object).SelectedIndex
			
		End If
		
		
		'End of support type
		g_EOSType = ARM_PLAIN
		If g_ExtendTo = GLOVE_AXILLA And CType(MainForm.Controls("optProximalTape"), Object)(1).Checked Then
			g_EOSType = ARM_FLAP
			g_sFlapType = CType(MainForm.Controls("cboFlaps"), Object).Text
			g_iFlapType = CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex
			If CType(MainForm.Controls("txtStrapLength"), Object).Text <> "" Then g_nStrapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtStrapLength"), Object))
			If CType(MainForm.Controls("txtFrontStrapLength"), Object).Text <> "" Then g_nFrontStrapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtFrontStrapLength"), Object))
			If CType(MainForm.Controls("txtCustFlapLength"), Object).Text <> "" Then g_nCustFlapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCustFlapLength"), Object))
			If CType(MainForm.Controls("txtWaistCir"), Object).Text <> "" Then g_nWaistCir = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtWaistCir"), Object))
		End If
		
		'On or off fold
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optFold"), Object)(1).Value = True Then g_OnFold = True Else g_OnFold = False
		
		
	End Sub
	
	Sub PR_GetExtensionDDE_Data()
		'Procedure to use the data given in the Form for the
		'glove extension (if any) along the arm
		'
		Dim nAge, Flap As Short
		Dim jj, ii, nn As Short
		Dim nValue As Double
		Dim iGms, iMMs, iRed As Short
		Dim sDiag As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 1
		
		'Defaults
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("optFold"), Object)(1).Value = True 'On fold
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True 'Normal Glove
		CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0 '1st Tape
		CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0 'Last Tape
		Flap = False
		
		
		'Set default pressure w.r.t diagnosis
		nAge = Val(CType(MainForm.Controls("txtAge"), Object).Text)
		
		' if pressure empty set from diagnosis
		sDiag = UCase(Mid(CType(MainForm.Controls("txtDiagnosis"), Object).Text, 1, 6))
		'Arterial Insufficiency
		If sDiag = "ARTERI" Then
			g_iPressure = 0
			'Assist fluid dynamics
		ElseIf sDiag = "ASSIST" Then 
			g_iPressure = 1
			'Blood Clots
		ElseIf sDiag = "BLOOD " Then 
			g_iPressure = 1
			'Burns
		ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then 
			g_iPressure = 0
			'Cancer
		ElseIf sDiag = "CANCER" Then 
			g_iPressure = 1
			'Cardial Vascular Arrest
		ElseIf sDiag = "CARDIA" Then 
			g_iPressure = 0
			'Carpal Tunnel Syndrome
		ElseIf sDiag = "CARPAL" Then 
			g_iPressure = 1
			'Cellulitis
		ElseIf sDiag = "CELLUL" Then 
			g_iPressure = 1
			'Chronic Venous Insufficiency
		ElseIf sDiag = "CHRONI" Then 
			g_iPressure = 1
			'Heart Condition
		ElseIf sDiag = "HEART " Then 
			g_iPressure = 0
			'Hemangioma
		ElseIf sDiag = "HEMANG" Then 
			g_iPressure = 0
			'Lymphedema, Lymphedema 1+, Lymphedema 2+
			'N.B. Lymphedema <=> Lymphedema 1+
		ElseIf sDiag = "LYMPHE" Then 
			If InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "2") > 0 Then
				If nAge <= 70 Then
					g_iPressure = 2
				Else
					g_iPressure = 1
				End If
			Else
				If nAge <= 70 Then
					g_iPressure = 1
				Else
					g_iPressure = 0
				End If
			End If
			'Night Wear
		ElseIf sDiag = "NIGHT " Then 
			g_iPressure = 0
			'Post Fracture
		ElseIf sDiag = "POST F" Then 
			g_iPressure = 1
			'Post Mastectomy
		ElseIf sDiag = "POST M" Then 
			g_iPressure = 1
			'Postphlebetic Syndrome
		ElseIf sDiag = "POSTPH" Then 
			g_iPressure = 1
			'Renal Disease (Kidney)
		ElseIf sDiag = "RENAL " Then 
			g_iPressure = 1
			'Skin Graft
		ElseIf sDiag = "SKIN G" Then 
			g_iPressure = 0
			'Stroke
		ElseIf sDiag = "STROKE" Then 
			g_iPressure = 0
			'Tendonitis
		ElseIf sDiag = "TENDON" Then 
			g_iPressure = 0
			'Thrombophlebitis
		ElseIf sDiag = "THROMB" Then 
			g_iPressure = 1
			'Trauma
		ElseIf sDiag = "TRAUMA" Then 
			g_iPressure = 1
			'Varicose Veins
		ElseIf sDiag = "VARICO" Then 
			g_iPressure = 1
			'Request for 30 m/m, 40 m/m and 50 m/m
		ElseIf sDiag = "REQUES" Then 
			If InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "30") > 0 Then
				g_iPressure = 0
			ElseIf InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "40") > 0 Then 
				g_iPressure = 1
			ElseIf InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "50") > 0 Then 
				g_iPressure = 2
			End If
			'Vein Removal
		ElseIf sDiag = "VEIN R" Then 
			g_iPressure = 1
		End If
		
		'Glove extensions
		'Code for extended glove
		'Glove Option Buttons
		Dim nLen As Double
		
		If CType(MainForm.Controls("txtDataGlove"), Object).Text <> "" Then
			For ii = 0 To 5
				nValue = Val(Mid(CType(MainForm.Controls("txtDataGlove"), Object).Text, (ii * 2) + 1, 2))
				If ii = 0 Then
					'Fold options
					'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CType(MainForm.Controls("optFold"), Object)(nValue).Value = True
				ElseIf ii = 1 And nValue >= 0 Then 
					'Pressure w.r.t Figuring and MMs
					g_iPressure = nValue
				ElseIf ii = 2 Then 
					'Wrist Tape
					CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = nValue
				ElseIf ii = 3 Then 
					'Proximal Tape
					CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = nValue
				ElseIf ii = 4 Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CType(MainForm.Controls("optExtendTo"), Object)(nValue).Value = True
					If nValue = 2 Then Flap = True
				ElseIf ii = 5 Then 
					'Only if glove to axilla is set do we
					'do anything
					If Flap And nValue = 1 Then
						'Flap, break up flap multiple field
						CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True
						PR_ProximalTape_Click((1))
						For jj = 0 To 4
							nValue = Val(Mid(CType(MainForm.Controls("txtFlap"), Object).Text, (jj * 3) + 1, 3))
							If jj = 0 Then
								CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = nValue
							ElseIf jj = 1 And nValue > 0 Then 
								g_nStrapLength = nValue / 10
								CType(MainForm.Controls("txtStrapLength"), Object).Text = CStr(g_nStrapLength)
								PR_DisplayTextInches(CType(MainForm.Controls("txtStrapLength"), Object), CType(MainForm.Controls("labStrap"), Object))
							ElseIf jj = 2 And nValue > 0 Then 
								g_nFrontStrapLength = nValue / 10
								CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = CStr(g_nFrontStrapLength)
								PR_DisplayTextInches(CType(MainForm.Controls("txtFrontStrapLength"), Object), CType(MainForm.Controls("labFrontStrapLength"), Object))
							ElseIf jj = 3 And nValue > 0 Then 
								g_nCustFlapLength = nValue / 10
								CType(MainForm.Controls("txtCustFlapLength"), Object).Text = CStr(g_nCustFlapLength)
								PR_DisplayTextInches(CType(MainForm.Controls("txtCustFlapLength"), Object), CType(MainForm.Controls("labCustFlapLength"), Object))
							ElseIf jj = 4 And nValue > 0 Then 
								g_nWaistCir = nValue / 10
								CType(MainForm.Controls("txtWaistCir"), Object).Text = CStr(g_nWaistCir)
								PR_DisplayTextInches(CType(MainForm.Controls("txtWaistCir"), Object), CType(MainForm.Controls("labWaistCir"), Object))
							End If
						Next jj
					Else
						'Disable flaps
						PR_ProximalTape_Click((0))
					End If
				End If
			Next ii
		End If
		
		'Set value for pressure
		CType(MainForm.Controls("cboPressure"), Object).SelectedIndex = g_iPressure
		
		'Glove to elbow and Glove to axilla
		'These can start at -3 (However to allow for possible
		'changes later we save up to -4-1/2 but ignore -4-1/2 for the
		'mean time)
		If CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text <> "" Then
			ii = 1
			For nn = 8 To 23
				nValue = Val(Mid(CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text, (ii * 3) + 1, 3))
				If nValue > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!txtExtCir(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CType(MainForm.Controls("txtExtCir"), Object)(nn) = nValue / 10
					nLen = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtExtCir"), Object)(nn))
					If nLen <> -1 Then PR_GrdInchesDisplay(nn - 8, nLen)
					g_nCir(ii) = nLen
				Else
					PR_GrdInchesDisplay(nn - 8, 0)
				End If
				iMMs = Val(Mid(CType(MainForm.Controls("txtTapeMMs"), Object).Text, (ii * 3) + 1, 3))
				If iMMs > 0 Then
					g_CalculatedExtension = True
					iGms = Val(Mid(CType(MainForm.Controls("txtGrams"), Object).Text, (ii * 3) + 1, 3))
					iRed = Val(Mid(CType(MainForm.Controls("txtReduction"), Object).Text, (ii * 3) + 1, 3))
					CType(MainForm.Controls("mms"), Object)(nn).Text = CStr(iMMs)
					PR_GramRedDisplay(nn - 8, iGms, iRed)
					g_iMMs(ii) = iMMs
					g_iRed(ii) = iRed
					g_iGms(ii) = iGms
				End If
				ii = ii + 1
			Next nn
		End If
		
		'Pleats
		For ii = 0 To 1
			nValue = Val(Mid(CType(MainForm.Controls("txtWristPleat"), Object).Text, (ii * 3) + 1, 3))
			If ii = 0 And nValue > 0 Then
				CType(MainForm.Controls("txtWristPleat1"), Object).Text = CStr(nValue / 10)
			ElseIf ii = 1 And nValue > 0 Then 
				CType(MainForm.Controls("txtWristPleat2"), Object).Text = CStr(nValue / 10)
			End If
			nValue = Val(Mid(CType(MainForm.Controls("txtShoulderPleat"), Object).Text, (ii * 3) + 1, 3))
			If ii = 0 And nValue > 0 Then
				CType(MainForm.Controls("txtShoulderPleat1"), Object).Text = CStr(nValue / 10)
			ElseIf ii = 1 And nValue > 0 Then 
				CType(MainForm.Controls("txtShoulderPleat2"), Object).Text = CStr(nValue / 10)
			End If
		Next ii
		
		
	End Sub
	
	
	Sub PR_GramRedDisplay(ByRef iIndex As Short, ByRef Gram As Short, ByRef Reduction As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdDisplay"), Object).Row = iIndex
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdDisplay"), Object).Col = 0
		If Gram = 0 Then CType(MainForm.Controls("grdDisplay"), Object).Text = "" Else CType(MainForm.Controls("grdDisplay"), Object).Text = Str(Gram)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdDisplay"), Object).Col = 1
		If Reduction = 0 Then CType(MainForm.Controls("grdDisplay"), Object).Text = "" Else CType(MainForm.Controls("grdDisplay"), Object).Text = Str(Reduction)
		
	End Sub
	
	Sub PR_GrdInchesDisplay(ByRef iIndex As Short, ByRef nLen As Double)
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdInches"), Object).Row = iIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdInches"), Object).Col = 0
		CType(MainForm.Controls("grdInches"), Object).Text = ARMEDDIA1.fnInchestoText(nLen)
	End Sub
	
	Sub PR_LoadPowernetChart()
		'Procedure to load the reduction charts for the arms from
		'disk
		'
		Static sModulus, sFile, sLine As String
		Static jj, fChart, ii, nn As Short
		
		sFile = g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(sFile) = "" Then
			MsgBox("Template file not found " & sFile, 48, g_sDialogID)
			Exit Sub
		End If
		
		'Open file
		fChart = FreeFile
		FileOpen(fChart, sFile, OpenMode.Input)
		
		'Get ARM chart reductions
		'NB these are fixed format
		ii = 1
		While Not EOF(fChart) And ii <= NOFF_MODULUS
			Input(fChart, sModulus)
			Input(fChart, sLine)
			
			For jj = 0 To 22
				g_iPowernet(ii, jj + 1) = Val(Mid(sLine, (jj * 4) + 1, 4))
			Next jj
			
			ii = ii + 1
			
		End While
		
		'Close file
		FileClose(fChart)
		sLine = ""
		sModulus = ""
		
	End Sub
	
	Sub PR_LoadReductionCharts(ByRef sFabric As String)
		'Procedure to load the reduction charts from
		'disk
		'Two charts are loaded
		'    1. Length chart
		'    2. Circumferences based on the Fabric
		'
		
		Static sFile, sLine As String
		Static jj, fChart, iModulus, ii, nn As Short
		
		If sFabric = "" Then Exit Sub
		If UCase(Mid(sFabric, 1, 3)) <> "POW" Then
			MsgBox("Fabric chosen is not Powernet", 48, g_sDialogID)
			Exit Sub
		End If
		
		'Establish chart to be loaded
		'    Fabric Format  Pow MMM-XX Comment
		'    if mm < 230 use 230 chart
		'    if MM > 280 use 280 chart
		'
		iModulus = Val(Mid(sFabric, 5, 3))
		If iModulus < 230 Then iModulus = 230
		If iModulus > 280 Then iModulus = 280
		sFile = g_sPathJOBST & "\TEMPLTS\GLV_" & Trim(Str(iModulus)) & ".DAT"
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(sFile) = "" Then
			MsgBox("Fabric Chart, Template file not found " & sFile, 48, g_sDialogID)
			Exit Sub
		End If
		
		'Open file
		fChart = FreeFile
		FileOpen(fChart, sFile, OpenMode.Input)
		
		'Get finger and thumb reducations
		'NB these are fixed format
		ii = 1
		While Not EOF(fChart) And ii <= NOFF_FINGER_RED
			sLine = LineInput(fChart)
			'Ignore comments and blank lines
			'NB use of "." to repeat previous number
			If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
				
				iFingerRed(1, ii) = Val(Mid(sLine, 5, 2))
				
				If Mid(sLine, 8, 1) = "." Then
					iFingerRed(2, ii) = iFingerRed(1, ii)
				Else
					iFingerRed(2, ii) = Val(Mid(sLine, 8, 2))
				End If
				
				If Mid(sLine, 11, 1) = "." Then
					iFingerRed(3, ii) = iFingerRed(2, ii)
				Else
					iFingerRed(3, ii) = Val(Mid(sLine, 11, 2))
				End If
				
				ii = ii + 1
			End If
		End While
		
		'Wrist and Palm to 9 tape reductions
		ii = 1
		While Not EOF(fChart) And ii <= NOFF_ARM_RED
			sLine = LineInput(fChart)
			'Ignore comments and blank lines
			'NB: use of "." to repeat previous number
			'    also translation from inches and eights to decimal inches)
			If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
				nArmRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
				jj = 1
				For nn = 9 To 34 Step 4
					jj = jj + 1
					If Mid(sLine, nn, 1) = "." Then
						nArmRed(jj, ii) = nArmRed(jj - 1, ii)
					Else
						nArmRed(jj, ii) = Val(Mid(sLine, nn, 2)) + (Val(Mid(sLine, nn + 2, 1)) * 0.125)
					End If
				Next nn
				ii = ii + 1
			End If
		End While
		
		'Close file
		FileClose(fChart)
		
		
		'Length Reduction chart
		sFile = g_sPathJOBST & "\TEMPLTS\GLV_LEN.DAT"
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(sFile) = "" Then
			MsgBox("Template file not found " & sFile, 48, g_sDialogID)
			Exit Sub
		End If
		
		'Open file
		fChart = FreeFile
		FileOpen(fChart, sFile, OpenMode.Input)
		ii = 1
		While Not EOF(fChart) And ii <= NOFF_LENGTH_RED
			sLine = LineInput(fChart)
			'Ignore comments and blank lines
			'NB: use of "." to repeat previous number
			'    also translation from inches and eights to decimal inches)
			If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
				
				nLengthRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
				
				If Mid(sLine, 9, 1) = "." Then
					nLengthRed(2, ii) = nLengthRed(1, ii)
				Else
					nLengthRed(2, ii) = Val(Mid(sLine, 9, 2)) + (Val(Mid(sLine, 11, 1)) * 0.125)
				End If
				
				ii = ii + 1
			End If
		End While
		
		'Close file
		FileClose(fChart)
		
		sLine = ""
		
	End Sub
	
	Sub PR_ProximalTape_Click(ByRef Index As Short)
		Dim ii As Short
		If Index = 0 Then
			'Disable flaps
			g_iFlapType = CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex
			CType(MainForm.Controls("cboFlaps"), Object).Enabled = False
			CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = -1
			For ii = 0 To 4
				CType(MainForm.Controls("lblFlap"), Object)(ii).Enabled = False
			Next ii
			CType(MainForm.Controls("txtStrapLength"), Object).Enabled = False
			g_nStrapLength = Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)
			CType(MainForm.Controls("txtStrapLength"), Object).Text = ""
			CType(MainForm.Controls("labStrap"), Object).Text = ""
			
			CType(MainForm.Controls("txtFrontStrapLength"), Object).Enabled = False
			g_nFrontStrapLength = Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text)
			CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = ""
			CType(MainForm.Controls("labFrontStrapLength"), Object).Text = ""
			
			CType(MainForm.Controls("txtCustFlapLength"), Object).Enabled = False
			g_nCustFlapLength = Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text)
			CType(MainForm.Controls("txtCustFlapLength"), Object).Text = ""
			CType(MainForm.Controls("labCustFlapLength"), Object).Text = ""
			
			CType(MainForm.Controls("txtWaistCir"), Object).Enabled = False
			g_nWaistCir = Val(CType(MainForm.Controls("txtWaistCir"), Object).Text)
			CType(MainForm.Controls("txtWaistCir"), Object).Text = ""
			CType(MainForm.Controls("labWaistCir"), Object).Text = ""
			
		Else
			'Enable flaps
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CType(MainForm.Controls("optFold"), Object)(0) = True 'Ensure off the fold only
			CType(MainForm.Controls("cboFlaps"), Object).Enabled = True
			CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = g_iFlapType
			
			For ii = 0 To 3
				CType(MainForm.Controls("lblFlap"), Object)(ii).Enabled = True
			Next ii
			CType(MainForm.Controls("txtStrapLength"), Object).Enabled = True
			If g_nStrapLength > 0 Then
				CType(MainForm.Controls("txtStrapLength"), Object).Text = CStr(g_nStrapLength)
				PR_DisplayTextInches(CType(MainForm.Controls("txtStrapLength"), Object), CType(MainForm.Controls("labStrap"), Object))
			End If
			
			CType(MainForm.Controls("txtFrontStrapLength"), Object).Enabled = True
			If g_nFrontStrapLength > 0 Then
				CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = CStr(g_nFrontStrapLength)
				PR_DisplayTextInches(CType(MainForm.Controls("txtFrontStrapLength"), Object), CType(MainForm.Controls("labFrontStrapLength"), Object))
			End If
			
			CType(MainForm.Controls("txtCustFlapLength"), Object).Enabled = True
			If g_nCustFlapLength > 0 Then
				CType(MainForm.Controls("txtCustFlapLength"), Object).Text = CStr(g_nCustFlapLength)
				PR_DisplayTextInches(CType(MainForm.Controls("txtCustFlapLength"), Object), CType(MainForm.Controls("labCustFlapLength"), Object))
			End If
			
			If InStr(1, CType(MainForm.Controls("cboFlaps"), Object).Text, "D") > 0 Then
				CType(MainForm.Controls("lblFlap"), Object)(4).Enabled = True
				CType(MainForm.Controls("txtWaistCir"), Object).Enabled = True
				If g_nWaistCir > 0 Then
					CType(MainForm.Controls("txtWaistCir"), Object).Text = CStr(g_nWaistCir)
					PR_DisplayTextInches(CType(MainForm.Controls("txtWaistCir"), Object), CType(MainForm.Controls("labWaistCir"), Object))
				End If
			End If
			
		End If
		
	End Sub
	
	Sub PR_SetMMs(ByRef sPressure As String)
		'Set the mms based on the pressure and the wrist and EOS
		'tapes given
		'REF:    GOP 01-02/16, Section 1.4
		'
		Dim ii As Short
		
		For ii = g_iWristPointer To g_iEOSPointer
			Select Case sPressure
				Case "15"
					If ii = g_iWristPointer Then
						g_iMMs(ii) = 12
					ElseIf ii < ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 15
					ElseIf ii = ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 12
					ElseIf ii = ELBOW_TAPE Then 
						g_iMMs(ii) = 8
					ElseIf ii = ELBOW_TAPE + 1 Then 
						g_iMMs(ii) = 10
					ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then 
						g_iMMs(ii) = 12
					ElseIf ii = g_iEOSPointer - 1 Then 
						g_iMMs(ii) = 10
					ElseIf ii = g_iEOSPointer Then 
						g_iMMs(ii) = 8
					End If
				Case "20"
					If ii = g_iWristPointer Then
						g_iMMs(ii) = 16
					ElseIf ii < ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 20
					ElseIf ii = ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 16
					ElseIf ii = ELBOW_TAPE Then 
						g_iMMs(ii) = 10
					ElseIf ii = ELBOW_TAPE + 1 Then 
						g_iMMs(ii) = 13
					ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then 
						g_iMMs(ii) = 16
					ElseIf ii = g_iEOSPointer - 1 Then 
						g_iMMs(ii) = 13
					ElseIf ii = g_iEOSPointer Then 
						g_iMMs(ii) = 10
					End If
				Case "25"
					If ii = g_iWristPointer Then
						g_iMMs(ii) = 20
					ElseIf ii < ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 25
					ElseIf ii = ELBOW_TAPE - 1 Then 
						g_iMMs(ii) = 20
					ElseIf ii = ELBOW_TAPE Then 
						g_iMMs(ii) = 13
					ElseIf ii = ELBOW_TAPE + 1 Then 
						g_iMMs(ii) = 17
					ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then 
						g_iMMs(ii) = 20
					ElseIf ii = g_iEOSPointer - 1 Then 
						g_iMMs(ii) = 17
					ElseIf ii = g_iEOSPointer Then 
						g_iMMs(ii) = 13
					End If
			End Select
		Next ii
		
	End Sub
	
	Sub PR_SetModulusIndex(ByRef sFabric As String)
		'Set the module level variable iModulus Index based on
		'the chosen fabric
		'
		' Where sFabric = Pow 230-2B
		'
		'    g_iModulusIndex Start = LOW_MODULUS  = 160
		'    g_iModulusIndex End   = HIGH_MODULUS = 340
		'
		' hence
		'    g_iModulusIndex = ((230 - 160) + 10) / 10
		'
		'
		Dim iMod As Short
		
		iMod = Val(Mid(sFabric, 5, 3))
		If iMod < LOW_MODULUS Then
			g_iModulusIndex = LOW_MODULUS
		ElseIf iMod > HIGH_MODULUS Then 
			g_iModulusIndex = HIGH_MODULUS
		Else
			g_iModulusIndex = ((iMod - LOW_MODULUS) + 10) / 10
		End If
		
	End Sub
	
	Sub PR_UpdateDDE_Extension()
		'Procedure to update the fields used when data is transfered
		'from DRAFIX using DDE.
		'Although the transfer back to DRAFIX is not via DDE we use the same controls
		'simply to illustrate the method by which the data is packed into the
		'fields
		
		'This routine is similar to PR_UpdateDDE in MANGLOVE.FRM#
		'it is split so that we can acccess the Module level variables
		Dim iLen, ii As Short
		Dim sLen, sPacked As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 1
		
		
		'Pack Glove extensions
		sPacked = "   " 'Allow for possible extension to -4 tape
		For ii = 8 To 23
			iLen = Val(CType(MainForm.Controls("txtExtCir"), Object)(ii).Text) * 10 'Shift decimal place
			
			If iLen <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(iLen)), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			
			sPacked = sPacked & sLen
			
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!txtTapeLengthPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("txtTapeLengthPt1"), Object) = sPacked
		
		'Pack Grams, MMs and Reductions
		Dim sTapeMMs, sGrams, sReduction As String
		
		sGrams = "   " 'Allow for possible extension to -4 tape
		sTapeMMs = "   " 'Allow for possible extension to -4 tape
		sReduction = "   " 'Allow for possible extension to -4 tape
		
		For ii = 1 To 16
			'Grams
			If g_iGms(ii) <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(g_iGms(ii))), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			sGrams = sGrams & sLen
			
			'Tape pressures
			If g_iMMs(ii) <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(g_iMMs(ii))), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			sTapeMMs = sTapeMMs & sLen
			
			'Reductions
			If g_iRed(ii) <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(g_iRed(ii))), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			sReduction = sReduction & sLen
			
		Next ii
		
		CType(MainForm.Controls("txtReduction"), Object).Text = sReduction
		CType(MainForm.Controls("txtTapeMMs"), Object).Text = sTapeMMs
		CType(MainForm.Controls("txtGrams"), Object).Text = sGrams
		
		'Fold, variations etc etc.
		sPacked = ""
		
		'Fold
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optFold"), Object)(0).Value = True Then sPacked = " 0" Else sPacked = " 1"
		
		'Pressure
		sLen = New String(" ", 2)
		sLen = RSet(Trim(Str(CType(MainForm.Controls("cboPressure"), Object).SelectedIndex)), Len(sLen))
		sPacked = sPacked & sLen
		
		'Wrist Tape
		sLen = RSet(Trim(Str(CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex)), Len(sLen))
		sPacked = sPacked & sLen
		
		'EOS Tape
		sLen = RSet(Trim(Str(CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex)), Len(sLen))
		sPacked = sPacked & sLen
		
		'Variations
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then
			sPacked = sPacked & " 0 0" 'Second Zero is w.r.t Flap
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CType(MainForm.Controls("optExtendTo"), Object)(1).Value = True Then 
			sPacked = sPacked & " 1 0" 'Second Zero is w.r.t Flap
		Else
			sPacked = sPacked & " 2"
			'Check for flaps
			If CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Then sPacked = sPacked & " 1" Else sPacked = sPacked & " 0"
		End If
		
		'Thumb Web Drop in 1/16ths for use only with web spacers
		'Ignored by every thing else
		sLen = RSet(Trim(Str(g_iThumbWebDrop)), Len(sLen))
		sPacked = sPacked & sLen
		
		
		CType(MainForm.Controls("txtDataGlove"), Object).Text = sPacked
		
		'Flap
		sPacked = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(2).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optExtendTo"), Object)(2).Value = True And CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Then
			'Flap Type
			sLen = New String(" ", 3)
			sLen = RSet(Trim(Str(CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex)), Len(sLen))
			sPacked = sPacked & sLen
			
			sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
			sPacked = sPacked & sLen
			
			sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
			sPacked = sPacked & sLen
			
			sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
			sPacked = sPacked & sLen
			
			sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWaistCir"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
			sPacked = sPacked & sLen
			
		End If
		
		CType(MainForm.Controls("txtFlap"), Object).Text = sPacked
		
		'Pleats
		sLen = New String(" ", 3)
		If Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) > 0 Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
		sPacked = sLen
		
		sLen = New String(" ", 3)
		If Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
		CType(MainForm.Controls("txtWristPleat"), Object).Text = sPacked & sLen
		
		sLen = New String(" ", 3)
		If Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) > 0 Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
		sPacked = sLen
		
		sLen = New String(" ", 3)
		If Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
		CType(MainForm.Controls("txtShoulderPleat"), Object).Text = sPacked & sLen
		
	End Sub
End Module
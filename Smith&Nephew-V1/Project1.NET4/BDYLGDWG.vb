Option Strict Off
Option Explicit On
Module BDYLGDWG
	'Project:   BODYDRAW.MAK
	'Module:    BDYLGDWG.BAS
	'
	'Purpose:   To collect all leg functions for the
	'           body suit into a single module so that they can
	'           share the module level variables
	'           as declared below.
	'
	'           In the longer term this same module could be
	'           generalised to the other lower extremity supports (and then again maybe not!)
	'
	'
	'Version:   1.00
	'Date:      9 October 1997
	'Author:    Gary George
	'
	'           © C-Gem Ltd.
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	'
	'
	'
	'
	
	'50MMHG Chart values
	Dim m_nNo(30) As Double
	Dim m_sScale(30) As String
	Dim m_nSpace(30) As Double
	Dim m_n20Len(30) As Double
	Dim m_nReduction(30) As Double
	
	'Left and right caculated values
	Dim m_nLeftOriginalTapeValues(30) As Double
	Dim m_nRightOriginalTapeValues(30) As Double
	
	Dim m_iLeftFirstTape As Short
	Dim m_iLeftLastTape As Short
	Dim m_iRightFirstTape As Short
	Dim m_iRightLastTape As Short
	
	'Constants
	Const SEAM As Double = 0.25

    'Flags
    Dim m_bFileLoaded As Short

    Public Structure XY
        Public x As Double
        Public y As Double
    End Structure

    Function FN_AverageLeftAndRightLegs() As Short
		Dim iLeftLegProfileCount, ii, bOutwithTolerance, iRightLegProfileCount As Short
        Dim xyTmp, xyTmp1 As XY
        Dim nAverage, nScale, nDistanceApart As Double
		
		FN_AverageLeftAndRightLegs = True
		nDistanceApart = 0
		
		'Compare first tapes
		If m_iLeftFirstTape <> m_iRightFirstTape Then
			FN_AverageLeftAndRightLegs = False
			nDistanceApart = 1000 'Dummy to fool averaging below
		End If
		
		'Compare lengths (takes care of pleats)
		If g_nLeftLegLength <> g_nRightLegLength Then
			FN_AverageLeftAndRightLegs = False
			nDistanceApart = 1000 'Dummy to fool averaging below
		End If
		
		
		'Average legs
		iLeftLegProfileCount = 0
		iRightLegProfileCount = 0
		If ii >= m_iRightFirstTape Then g_RightLegProfile.n = g_RightLegProfile.n + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object max(m_iLeftLastTape, m_iRightLastTape). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object min(m_iLeftFirstTape, m_iRightFirstTape). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For ii = BDYUTILS.min(m_iLeftFirstTape, m_iRightFirstTape) To ARMDIA1.max(m_iLeftLastTape, m_iRightLastTape)
			If ii >= m_iLeftFirstTape Then iLeftLegProfileCount = iLeftLegProfileCount + 1
			If ii >= m_iRightFirstTape Then iRightLegProfileCount = iRightLegProfileCount + 1
			
			'Y scaling (m_n20Len(ii) / 20) from 50MMHG Chart
			If m_nRightOriginalTapeValues(ii) > 0 And m_nLeftOriginalTapeValues(ii) > 0 Then
				If System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii)) <= INCH3_8 Then
					'Average leg
					If ii = m_iLeftFirstTape And ii = m_iRightFirstTape Then
						If g_bLeftAboveKnee Then
							nScale = 0.95 / 2 '5 reduction
						Else
							nScale = 0.92 / 2 '8 reduction
						End If
					ElseIf ii = m_iLeftFirstTape Or ii = m_iRightFirstTape Then 
						'Do nothing
						'keep original legs tape values
						GoTo Continue_Loop
					Else
						nScale = m_n20Len(ii) / 20
					End If
					nAverage = (m_nRightOriginalTapeValues(ii) + m_nLeftOriginalTapeValues(ii)) / 2
					m_nLeftOriginalTapeValues(ii) = nAverage
					m_nRightOriginalTapeValues(ii) = nAverage
					g_LeftLegProfile.y(iLeftLegProfileCount) = (nAverage * nScale) + SEAM
					g_RightLegProfile.y(iRightLegProfileCount) = (nAverage * nScale) + SEAM
					
				Else
					'Do nothing
					'keep original legs tape values
					FN_AverageLeftAndRightLegs = False
					If System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii)) > nDistanceApart Then
						nDistanceApart = System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii))
						If iLeftLegProfileCount >= 1 Then
							ARMDIA1.PR_MakeXY(xyTmp, g_LeftLegProfile.x(iLeftLegProfileCount), g_LeftLegProfile.y(iLeftLegProfileCount))
							ARMDIA1.PR_MakeXY(xyTmp1, g_LeftLegProfile.x(iLeftLegProfileCount + 1), g_LeftLegProfile.y(iLeftLegProfileCount + 1))
						ElseIf iLeftLegProfileCount = g_LeftLegProfile.n Then 
							ARMDIA1.PR_MakeXY(xyTmp, g_LeftLegProfile.x(iLeftLegProfileCount), g_LeftLegProfile.y(iLeftLegProfileCount))
							ARMDIA1.PR_MakeXY(xyTmp1, g_LeftLegProfile.x(iLeftLegProfileCount - 1), g_LeftLegProfile.y(iLeftLegProfileCount - 1))
						End If
						ARMDIA1.PR_CalcPolar(xyTmp1, ARMDIA1.FN_CalcAngle(xyTmp1, xyTmp), ARMDIA1.FN_CalcLength(xyTmp1, xyTmp) * 2 / 3, xyLeftLegLabelPoint)
						
						If iRightLegProfileCount >= 1 Then
							ARMDIA1.PR_MakeXY(xyTmp, g_RightLegProfile.x(iRightLegProfileCount), g_RightLegProfile.y(iRightLegProfileCount))
							ARMDIA1.PR_MakeXY(xyTmp1, g_RightLegProfile.x(iRightLegProfileCount + 1), g_RightLegProfile.y(iRightLegProfileCount + 1))
						ElseIf iRightLegProfileCount = g_RightLegProfile.n Then 
							ARMDIA1.PR_MakeXY(xyTmp, g_RightLegProfile.x(iRightLegProfileCount), g_RightLegProfile.y(iRightLegProfileCount))
							ARMDIA1.PR_MakeXY(xyTmp1, g_RightLegProfile.x(iRightLegProfileCount - 1), g_RightLegProfile.y(iRightLegProfileCount - 1))
						End If
						ARMDIA1.PR_CalcPolar(xyTmp1, ARMDIA1.FN_CalcAngle(xyTmp1, xyTmp), ARMDIA1.FN_CalcLength(xyTmp1, xyTmp) * 1 / 3, xyRightLegLabelPoint)
						
					End If
				End If
			Else
				'Do nothing
				'keep original legs tape values
				FN_AverageLeftAndRightLegs = False
				
			End If
Continue_Loop: 
		Next ii
		'Revise label points so that they sit inside the lisne
		ARMDIA1.PR_MakeXY(xyRightLegLabelPoint, xyRightLegLabelPoint.x - g_nRightLegLength, xyRightLegLabelPoint.y - INCH1_16)
		ARMDIA1.PR_MakeXY(xyLeftLegLabelPoint, xyLeftLegLabelPoint.x - g_nLeftLegLength, xyLeftLegLabelPoint.y - INCH1_16)
		
	End Function
	
	Function FN_InitiliseLegRoutines() As Short
		'
		Dim ii, fFileNum As Short
		Dim sFileName As String
		
		'Open leg template data file
		'Check to see if has been loaded from file before,
		'so we don't load it twice
		If Not m_bFileLoaded Then
			fFileNum = FreeFile
			sFileName = g_sPathJOBST & "\TEMPLTS\Wh50mmhg.DAT"
			If FileLen(sFileName) = 0 Then
				MsgBox("Can't open the template file " & sFileName & ".  Unable to draw the leg curve.", 48)
				FN_InitiliseLegRoutines = False
				Exit Function
			End If
			ii = 0
			FileOpen(fFileNum, sFileName, OpenMode.Input)
			Do While Not EOF(fFileNum) And ii < 30
				ii = ii + 1
				Input(fFileNum, m_nNo(ii))
				Input(fFileNum, m_sScale(ii))
				Input(fFileNum, m_nSpace(ii))
				Input(fFileNum, m_n20Len(ii))
				Input(fFileNum, m_nReduction(ii))
			Loop 
			FileClose(fFileNum)
			m_bFileLoaded = True
		End If
		
		FN_InitiliseLegRoutines = True
		
	End Function
	
	Function FN_ValidateLegData() As Short
		Dim sError As String
		sError = ""
		If g_sLeftLeg = "Panty" And Val(bodydia.txtLeftUidLeg.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Panty requested for Left leg but no data found!" & NL
		End If
		If g_sRightLeg = "Panty" And Val(bodydia.txtRightUidLeg.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Panty requested for Right leg but no data found!" & NL
		End If
		
		'Display errors if found
		If Len(sError) > 0 Then
			MsgBox(sError, 52, "Severe Problems with data")
			FN_ValidateLegData = False
		Else
			FN_ValidateLegData = True
		End If
	End Function
	
	Sub PR_CalculateLeg(ByRef sLeg As String, ByRef nLength As Double)
        'Routine to calculate the leg
        'Returns:   nLength = the length the of the leg
        'Updates:   m_*  (Module level variables)
        '
        'Assumes:   1. That data exists.
        '           2. That FN_InitiliseLegRoutines has been called
        '

        'UPGRADE_WARNING: Arrays in structure LegProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim LegProfile As ARMDIA1.curve
        Dim sTapes As String
		Dim sPleats As String
		Dim nValue, nSpace, nScale As Double
		Dim nTopLegPleat1, nFootPleat1, nFootPleat2, nTopLegPleat2 As Double
		Dim iLastTape, iFirstTape, bAboveKnee As Short
		Dim nn, ii, iValue As Short
		Dim iLowest As Short
		Dim nValueLowest As Double
		Dim xyTmp, xyTmp1 As xy
		
		If sLeg = "Left" Then
			sTapes = CType(bodydia.Controls("txtLeftLengths"), Object).Text
			sPleats = CType(bodydia.Controls("txtLeftLegPleats"), Object).Text
			iValue = Val(CType(bodydia.Controls("txtLeftKneePosition"), Object).Text)
			If iValue = 1 Then g_bLeftAboveKnee = True Else g_bLeftAboveKnee = False
			bAboveKnee = g_bLeftAboveKnee
		Else
			sTapes = CType(bodydia.Controls("txtRightLengths"), Object).Text
			sPleats = CType(bodydia.Controls("txtRightLegPleats"), Object).Text
			iValue = Val(CType(bodydia.Controls("txtRightKneePosition"), Object).Text)
			If iValue = 1 Then g_bRightAboveKnee = True Else g_bRightAboveKnee = False
			bAboveKnee = g_bRightAboveKnee
		End If
		
		If sTapes = "" Then Exit Sub
		
		'Split leg tape circumferances and establish start and finish tapes
		LegProfile.n = 0
		iFirstTape = -1
		iLastTape = 30
		For ii = 0 To 29
			nValue = Val(Mid(sTapes, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				LegProfile.n = LegProfile.n + 1
				nValue = ARMEDDIA1.fnDisplayToInches(nValue)
				LegProfile.y(LegProfile.n) = nValue
				If sLeg = "Left" Then m_nLeftOriginalTapeValues(ii) = nValue Else m_nRightOriginalTapeValues(ii) = nValue
			End If
			
			'Set first and last tape (assumes no holes in data)
			If iFirstTape < 0 And nValue > 0 Then iFirstTape = ii
			If iLastTape = 30 And iFirstTape > 0 And nValue = 0 Then iLastTape = ii - 1
		Next ii
		
		'Using the 50MMHG chart loaded on initilisation and the pleats (if any) we need to
		'calculate the x values of the profile and the overall length of the leg.
		'NB
		'    These are relative X values only as the absolute values will be calculated
		'    only when the leg is to be drawn.
		nFootPleat1 = Val(ARMDIA1.fnGetString(sPleats, 1, ","))
		nFootPleat2 = Val(ARMDIA1.fnGetString(sPleats, 2, ","))
		nTopLegPleat2 = Val(ARMDIA1.fnGetString(sPleats, 3, ","))
		nTopLegPleat1 = Val(ARMDIA1.fnGetString(sPleats, 4, ","))
		
		nn = 1 'Counter for LegProfile
		nLength = 0 'Note - length to be returned above
		'    nValueLowest = 10000
		For ii = iFirstTape To iLastTape
			If (ii = iFirstTape) Then
				nSpace = 0
			ElseIf (ii = iFirstTape + 1) And (nFootPleat1 <> 0) Then 
				nSpace = nFootPleat1
			ElseIf (ii = iFirstTape + 2) And (nFootPleat2 <> 0) Then 
				nSpace = nFootPleat2
			ElseIf (ii = iLastTape - 1) And (nTopLegPleat2 <> 0) Then 
				nSpace = nTopLegPleat2
			ElseIf (ii = iLastTape) And (nTopLegPleat1 <> 0) Then 
				nSpace = nTopLegPleat1
			Else
				nSpace = m_nSpace(ii - 1) 'Spacing from 50MMHG Chart
			End If
			nLength = nLength + nSpace
			LegProfile.x(nn) = nLength
			
			'Y scaling (m_n20Len(ii) / 20) from 50MMHG Chart
			If ii = iFirstTape Then
				If bAboveKnee Then
					nScale = 0.95 / 2 '5 reduction
				Else
					nScale = 0.92 / 2 '8 reduction
				End If
			Else
				nScale = m_n20Len(ii) / 20
			End If
			LegProfile.y(nn) = (LegProfile.y(nn) * nScale) + SEAM
			
			nn = nn + 1
		Next ii
		ARMDIA1.PR_MakeXY(xyTmp, LegProfile.x(1), LegProfile.y(1))
		ARMDIA1.PR_MakeXY(xyTmp1, LegProfile.x(2), LegProfile.y(2))
		
		'Assign calculated values to left and right as required
		If sLeg = "Left" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_LeftLegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_LeftLegProfile = LegProfile
			m_iLeftFirstTape = iFirstTape
			m_iLeftLastTape = iLastTape
			BDYUTILS.PR_CalcMidPoint(xyTmp1, xyTmp, xyLeftLegLabelPoint)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object g_RightLegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_RightLegProfile = LegProfile
			m_iRightFirstTape = iFirstTape
			m_iRightLastTape = iLastTape
			BDYUTILS.PR_CalcMidPoint(xyTmp1, xyTmp, xyRightLegLabelPoint)
		End If
		
	End Sub
	
	
	Sub PR_CopyRightToLeftLeg()
		Dim ii As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object g_LeftLegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_LeftLegProfile = g_RightLegProfile
		g_nLeftLegLength = g_nRightLegLength
		m_iLeftFirstTape = m_iRightFirstTape
		m_iLeftLastTape = m_iRightLastTape
		'UPGRADE_WARNING: Couldn't resolve default property of object xyLeftLegLabelPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyLeftLegLabelPoint = xyRightLegLabelPoint
		
		For ii = 0 To 29
			m_nLeftOriginalTapeValues(ii) = m_nRightOriginalTapeValues(ii)
		Next ii
		
	End Sub
	
	Sub PR_DrawLegTemplate(ByRef sLeg As String, ByRef xyStart As xy)
		'Draw the leg template based on the calculated values.
		'We are actually drawing the tick marks and the labels at
		'each tape.
		'  sLeg    "Left"
		'          "Right"
		'          "Both"  combine the two sets of templates
		'                  this becomes a bit tricky when we have pleats
		'Note the "profile" for the leg will be drawn elsewhere as the leg is a continuation
		'of the back body and is not drawn seperatly.
		'
		Dim iLastTape, iFirstTape, ii As Short
		Dim nMaxHt, nX, nMinHt As Double
		'UPGRADE_WARNING: Arrays in structure LegProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim LegProfile As curve
		
		'get pre-calculated module level variables
		If sLeg = "Left" Then
			iFirstTape = m_iLeftFirstTape
			iLastTape = m_iLeftLastTape
			'UPGRADE_WARNING: Couldn't resolve default property of object LegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LegProfile = g_LeftLegProfile
		ElseIf sLeg = "Right" Then 
			iFirstTape = m_iRightFirstTape
			iLastTape = m_iRightLastTape
			'UPGRADE_WARNING: Couldn't resolve default property of object LegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LegProfile = g_RightLegProfile
		End If
		
		If sLeg = "Both" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iFirstTape = BDYUTILS.min(m_iRightFirstTape, m_iLeftFirstTape)
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iLastTape = ARMDIA1.max(m_iRightLastTape, m_iLeftLastTape)
			
			BDYUTILS.PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST))
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\BODY\BD_BOTH.D;")
			
			nX = xyStart.x
			For ii = iFirstTape To iLastTape
				'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMaxHt = ARMDIA1.max(m_nLeftOriginalTapeValues(ii), m_nRightOriginalTapeValues(ii))
				'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMinHt = BDYUTILS.min(m_nLeftOriginalTapeValues(ii), m_nRightOriginalTapeValues(ii))
				If m_nLeftOriginalTapeValues(ii) < m_nRightOriginalTapeValues(ii) Then
					BDYUTILS.PR_PutStringAssign("sMinHtColour", "blue")
					BDYUTILS.PR_PutStringAssign("sMaxHtColour", "red")
				ElseIf m_nLeftOriginalTapeValues(ii) = m_nRightOriginalTapeValues(ii) Then 
					BDYUTILS.PR_PutStringAssign("sMinHtColour", "black")
					BDYUTILS.PR_PutStringAssign("sMaxHtColour", "black")
				Else
					BDYUTILS.PR_PutStringAssign("sMinHtColour", "red")
					BDYUTILS.PR_PutStringAssign("sMaxHtColour", "blue")
				End If
				BDYUTILS.PR_PutStringAssign("sSymbol", Trim(Str(ii)) & "tape")
				BDYUTILS.PR_PutNumberAssign("n20Len", m_n20Len(ii))
				BDYUTILS.PR_PutNumberAssign("nReduction", m_nReduction(ii))
				BDYUTILS.PR_PutNumberAssign("nMaxHt", nMaxHt)
				BDYUTILS.PR_PutNumberAssign("nMinHt", nMinHt)
				BDYUTILS.PR_PutNumberAssign("nX", nX)
				nX = nX + m_nSpace(ii)
				
				BDYUTILS.PR_PutLine("PR_DrawTapeScale();")
				
			Next ii
			
		Else
			'sLeg = Left or sLeg = Right
			BDYUTILS.PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST))
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\BODY\BD_SNGLE.D;")
			nX = xyStart.x
			For ii = iFirstTape To iLastTape
				If sLeg = "Left" Then
					nMaxHt = m_nLeftOriginalTapeValues(ii)
				Else
					nMaxHt = m_nRightOriginalTapeValues(ii)
				End If
				BDYUTILS.PR_PutStringAssign("sSymbol", Trim(Str(ii)) & "tape")
				BDYUTILS.PR_PutNumberAssign("n20Len", m_n20Len(ii))
				BDYUTILS.PR_PutNumberAssign("nReduction", m_nReduction(ii))
				BDYUTILS.PR_PutNumberAssign("nHt", nMaxHt)
				BDYUTILS.PR_PutNumberAssign("nX", nX)
				nX = nX + m_nSpace(ii)
				
				BDYUTILS.PR_PutLine("PR_DrawTapeScale();")
				
			Next ii
		End If
		
		
	End Sub
End Module
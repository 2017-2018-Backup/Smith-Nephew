Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class manglved
	Inherits System.Windows.Forms.Form
	'Project:   MANGLVED.MAK
	'Purpose:   Editor for Manual glove that extends to the
	'           elbow or axilla.
	'
	'Version:   1.01
	'Date:      16.May.96
	'Author:    Gary George
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Dec 98     GG      Ported to VB5
	'
	'Notes:-
	'
	' Known Bugs
	'
	' 1 Fails if DRAFIX Help is in use.
	'
	'
	
	Private Function Arccos(ByRef X As Double) As Double
		Arccos = System.Math.Atan(-X / System.Math.Sqrt(-X * X + 1)) + 1.5708
	End Function
	
	'UPGRADE_WARNING: Event cboContracture.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboContracture_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboContracture.SelectedIndexChanged
		Dim sType As String
		
		If g_NoElbowTape Then Exit Sub
		
		If cboContracture.Text <> txtContracture.Text Then
			txtContracture.Text = cboContracture.Text
			sType = VB.Left(txtContracture.Text, 2)
			Select Case sType
				Case "No"
					'Reset mms at elbow to standard and recalculate
					Select Case txtMM.Text
						Case "15mm"
							EditMMs(g_iElbowTape).Text = CStr(12)
						Case "20mm"
							EditMMs(g_iElbowTape).Text = CStr(16)
						Case "25mm"
							EditMMs(g_iElbowTape).Text = CStr(20)
					End Select
					PR_CalculateFromMMs(g_iElbowTape)
					cboContracture.Text = ""
					txtContracture.Text = ""
				Case "10"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 12)
				Case "36"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 15)
				Case "71"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 17)
			End Select
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboLining.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLining_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLining.SelectedIndexChanged
		If cboLining.Text <> txtLining.Text Then
			If cboLining.Text = "None" Then cboLining.Text = ""
			txtLining.Text = cboLining.Text
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		'This cancel will delete the tempory curve then exit.
		'Checks to see if there is an Available drafix instance to
		'sendkeys to.
		Dim sTask As String
		If g_ReDrawn = True Then
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				PR_CancelProfileEdits()
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
			End If
		End If
		End
	End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		'Check that data is all present and commit into drafix
		If FN_Validate_Data() Then
			PR_DrawAndCommitProfileEdits()
			AppActivate(fnGetDrafixWindowTitleText())
			System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
			End
		End If
	End Sub
	
	Private Sub cmdDraw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDraw.Click
		'Check that data is all present and insert into drafix
		If FN_Validate_Data() Then
			PR_DrawProfileEdits()
			AppActivate(fnGetDrafixWindowTitleText())
			g_ReDrawn = True
			System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
		End If
	End Sub
	
	Private Sub EditMMs_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditMMs.Enter
		Dim Index As Short = EditMMs.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(EditMMs(Index))
	End Sub
	
	Private Sub EditMMs_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditMMs.Leave
		Dim Index As Short = EditMMs.GetIndex(eventSender)
		If IsNumeric(EditMMs(Index).Text) Then PR_CalculateFromMMs(Index)
	End Sub
	
	Private Sub EditTapes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditTapes.Enter
		Dim Index As Short = EditTapes.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(EditTapes(Index))
	End Sub
	
	Private Sub EditTapes_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditTapes.Leave
		Dim Index As Short = EditTapes.GetIndex(eventSender)
		'Check length given is valid
		'Display value in inches
		If IsNumeric(EditTapes(Index).Text) Then
			InchText(Index).Text = ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(EditTapes(Index).Text)))
			PR_CalculateFromMMs(Index)
		End If
	End Sub
	
	Private Function FN_CirLinInt(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyCen As XY, ByRef nRad As Double, ByRef xyInt As XY) As Short
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
		nSlope = ARMDIA1.FN_CalcAngle(xyStart, xyEnd)
		
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
				If nRoot >= BDYUTILS.min(xyStart.X, xyEnd.X) And nRoot <= ARMDIA1.max(xyStart.X, xyEnd.X) Then
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
				If nRoot >= BDYUTILS.min(xyStart.y, xyEnd.y) And nRoot <= ARMDIA1.max(xyStart.y, xyEnd.y) Then
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
				If nRoot >= BDYUTILS.min(xyStart.X, xyEnd.X) And nRoot <= ARMDIA1.max(xyStart.X, xyEnd.X) Then
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
	
	Private Function FN_DrawOpen(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_DrawOpen = fNum
		
		'Initialise String globals
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma ( , )
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(10) 'The new line character
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes ( " )
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCQ = QQ & CC & QQ 'Quote Comma Quote ( "," )
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QC = QQ & CC 'Quote Comma ( ", )
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CQ = CC & QQ 'Comma Quote ( ," )
		
		'Globals to reduced drafix code written to file
		g_sCurrentLayer = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAspect = 0.6
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Arm Editing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		
	End Function
	
	Private Function FN_FiguredValue(ByRef nLength As Double, ByRef iRed As Short) As Double
		'Fuction to reduce code only
		Dim nValue As Double
		nValue = ((ARMEDDIA1.fnDisplayToInches(nLength) * (System.Math.Abs(iRed - 100) / 100)))
		Select Case g_iInsertStyle
			Case 0
				If Not g_OffFold Then
					nValue = ((nValue + EIGHTH) - g_nInsertSize)
				Else
					nValue = ((nValue + (3 * EIGHTH)) - (2 * g_nInsertSize))
				End If
			Case 1
				nValue = ((nValue + (3 * EIGHTH)) - g_nInsertSize)
			Case 2, 3
				nValue = (nValue + (2 * EIGHTH))
		End Select
		
		FN_FiguredValue = nValue / 2
		
	End Function
	
	Private Function FN_Validate_Data() As Short
		
		Dim sError, NL As String
		Dim nFirstTape, ii, nLastTape As Short
		Dim iError As Short
		Dim nValue As Double
		
		NL = Chr(10) 'new line
		
		
		'Display Error message (if required) and return
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			FN_Validate_Data = False
			Exit Function
		Else
			FN_Validate_Data = True
		End If
		
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		Dim iValue, ii, nn, filenum As Short
		Dim nValue As Double
		Dim iFirstTape, iLastTape As Object
		Dim textline, sString As String
		
		Timer1.Enabled = False
		'Test that this is an editable glove (Warn and exit if Not)
		g_GloveType = Val(Mid(txtData.Text, (4 * 2) + 1, 2))
		g_EOSType = Val(Mid(txtData.Text, (5 * 2) + 1, 2))
		
		If g_GloveType = GLOVE_NORMAL Then
			MsgBox("This is a NORMAL Glove, Can't EDIT the arm of this type of glove.", 16, "Glove Edit Warning")
			End
		End If
		
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Other globals
		g_sID = txtID.Text
		g_sFileNo = Mid(txtID.Text, 5)
		g_sSide = txtSide.Text
		
		
		'Title bar display
		If g_sSide = "Left" Then
			g_sFileNo = Mid(g_sFileNo, Len(g_sFileNo) - 4)
			Me.Text = "GLOVE Edit - Left [" & g_sFileNo & "]"
			g_Direction = -1
		End If
		If g_sSide = "Right" Then
			g_sFileNo = Mid(g_sFileNo, Len(g_sFileNo) - 5)
			Me.Text = "GLOVE Edit - Right [" & g_sFileNo & "]"
			g_Direction = 1
		End If
		
		'Display - Given Length, Length in inches, MMs, Grams and Reduction at each tape
		'Store values into working arrays. Store also the initial values for change checking
		txtTapeLengthsPt1.Text = "   " & txtTapeLengthsPt1.Text
		txtTapeMMs.Text = "   " & txtTapeMMs.Text
		txtGrams.Text = "   " & txtGrams.Text
		txtReduction.Text = "   " & txtReduction.Text
		
		For ii = 0 To 17
			nValue = Val(Mid(txtTapeLengthsPt1.Text, (ii * 3) + 1, 3)) / 10
			If nValue > 0 Then
				
				EditTapes(ii).Text = CStr(nValue)
				g_nLengths(ii) = nValue
				g_nLengthsInit(ii) = nValue
				InchText(ii).Text = ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(nValue))
				
				iValue = Val(Mid(txtTapeMMs.Text, (ii * 3) + 1, 3))
				EditMMs(ii).Text = CStr(iValue)
				g_iMMs(ii) = iValue
				g_iMMsInit(ii) = iValue
				
				iValue = Val(Mid(txtGrams.Text, (ii * 3) + 1, 3))
				EditGrams(ii).Text = Str(iValue)
				g_iGms(ii) = iValue
				g_iGmsInit(ii) = iValue
				
				iValue = Val(Mid(txtReduction.Text, (ii * 3) + 1, 3))
				EditReductions(ii).Text = Str(iValue)
				g_iRed(ii) = iValue
				g_iRedInit(ii) = iValue
				
			End If
		Next ii
		
		'Establish First tape
		'Disable unused tapes
		g_iFirstTape = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object iFirstTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iFirstTape = Val(Mid(txtData.Text, (2 * 2) + 1, 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object iFirstTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If iFirstTape > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object iFirstTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_iFirstTape = iFirstTape + 1
			For ii = 0 To g_iFirstTape - 1
				EditTapes(ii).Enabled = False
				EditMMs(ii).Enabled = False
				EditMMs(ii).Text = ""
				EditGrams(ii).Text = ""
				EditReductions(ii).Text = ""
				InchText(ii).Text = ""
				EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next ii
		Else
			nValue = 0
			For ii = 0 To 17
				nValue = Val(EditTapes(ii).Text)
				If nValue > 0 Then Exit For
				EditTapes(ii).Enabled = False
				EditMMs(ii).Enabled = False
				EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next ii
			g_iFirstTape = ii
		End If
		
		
		'Establish Last tape (DISABLE down to Last Tape)
		g_iLastTape = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object iLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iLastTape = Val(Mid(txtData.Text, (3 * 2) + 1, 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object iLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If iLastTape > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object iLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_iLastTape = (17 - iLastTape) + 1
			For ii = 17 To g_iLastTape + 1 Step -1
				EditTapes(ii).Enabled = False
				EditMMs(ii).Enabled = False
				EditMMs(ii).Text = ""
				EditGrams(ii).Text = ""
				EditReductions(ii).Text = ""
				InchText(ii).Text = ""
				EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next ii
		Else
			nValue = 0
			For ii = 17 To 0 Step -1
				nValue = Val(EditTapes(ii).Text)
				If nValue > 0 Then Exit For
				EditTapes(ii).Enabled = False
				EditMMs(ii).Enabled = False
				EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next ii
			g_iLastTape = ii
		End If
		
		'If the Proximal style is Flap
		'then disable last tape editing
		If Val(Mid(txtData.Text, (5 * 2) + 1, 2)) = 1 Then
			EditTapes(g_iLastTape).Enabled = False
			EditMMs(g_iLastTape).Enabled = False
			EditMMs(g_iLastTape).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			EditGrams(ii).Enabled = False
			EditReductions(ii).Enabled = False
			InchText(ii).Enabled = False
		End If
		
		'Establish if drawn Off fold
		g_OffFold = False
		If Val(Mid(txtData.Text, (0 * 2) + 1, 2)) = 0 Then g_OffFold = True
		
		'Establish insert size
		'Insert size is stored in Eighths
		sString = Mid(txtTapeLengths2.Text, 15, 2)
		If Val(sString) > 0 Then g_nInsertSize = Val(sString) * 0.125
		
		'Insert Options
		g_iInsertStyle = Val(Mid(txtTapeLengths2.Text, 19, 1))
		
		'Load the original profile from file
		'Setup profile vertex to tape mappings
		PR_GetProfileFromFile("C:\JOBST\ARMCURVE.DAT")
		
		
		'Load fabric conversion charts from file
		'Establish Fabric
		If Mid(txtFabric.Text, 1, 3) = "Pow" Then
			ARMEDDIA1.PR_LoadFabricFromFile(g_MATERIAL, g_sPathJOBST & "\TEMPLTS\POWERNET.DAT")
		Else
			MsgBox("Unsupported fabric type, Can't EDIT using this type of Fabric.", 16, "Glove Edit Warning")
			End
		End If
		
		'Establish Fabric modulus
		g_iModulus = Val(Mid(txtFabric.Text, 5, 3))
		
		'If there is no elbow tape or elbow tape is the last tape
		'then disable linings and contractures
		g_iElbowTape = 10
		g_sOriginalContracture = txtContracture.Text
		g_sOriginalLining = txtLining.Text
		If Val(EditTapes(g_iElbowTape).Text) = 0 Or g_iElbowTape = g_iFirstTape Or g_iElbowTape = g_iLastTape Then
			cboContracture.Enabled = False
			cboLining.Enabled = False
			g_NoElbowTape = True
		Else
			cboContracture.Text = txtContracture.Text
			cboLining.Text = txtLining.Text
		End If
		
		'Set redraw flag to False
		g_ReDrawn = False
		g_BlendingCurveChanged = False
		
		Show()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
		cmdCancel.Focus()
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub manglved_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim filenum As Short
		Dim textline As String
		'Hide form while loading
		Hide()
		
		'Start Time out timer
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("The Glove Edit Module is already running!" & Chr(13) & "Use ALT-TAB to access it .", 16, "Leg Edit Warning")
			End
		End If
		
		'Position to "off" center
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 10)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
		
		MainForm = Me
		
		'Set Tape Numbers
		num(0).Text = "-6"
		num(1).Text = "-4" & Chr(189)
		num(2).Text = "-3"
		num(3).Text = "-1" & Chr(189)
		num(4).Text = "0"
		num(5).Text = "1" & Chr(189)
		num(6).Text = "3"
		num(7).Text = "4" & Chr(189)
		num(8).Text = "6"
		num(9).Text = "7" & Chr(189)
		num(10).Text = "9"
		num(11).Text = "10" & Chr(189)
		num(12).Text = "12"
		num(13).Text = "13" & Chr(189)
		num(14).Text = "15"
		num(15).Text = "16" & Chr(189)
		num(16).Text = "18"
		num(17).Text = "19" & Chr(189)
		
		txtUnits.Text = ""
		txtUIDETS.Text = ""
		txtUIDLFS.Text = ""
		txtUIDEOS.Text = ""
		txtUIDTempETS.Text = ""
		txtUIDTempLFS.Text = ""
		txtUIDPALMER.Text = ""
		txtSide.Text = ""
		txtID.Text = ""
		txtData.Text = ""
		txtTapeLengthsPt1.Text = ""
		txtTapeLengths2.Text = ""
		txtTapeMMs.Text = ""
		txtMM.Text = ""
		txtGrams.Text = ""
		txtReduction.Text = ""
		txtContracture.Text = ""
		txtLining.Text = ""
		txtFabric.Text = ""
		
		'Linings Combo
		cboLining.Items.Add("Inside Lining")
		cboLining.Items.Add("Outside Lining")
		cboLining.Items.Add("Lining")
		cboLining.Items.Add("Full Lining")
		cboLining.Items.Add("Reinforced Elbow")
		cboLining.Items.Add("None")
		
		'Contractures Combo
		cboContracture.Items.Add("10-35 Degrees")
		cboContracture.Items.Add("36-70 Degrees")
		cboContracture.Items.Add("71 Degrees and Over")
		cboContracture.Items.Add("None")
		
		g_nUnitsFac = 1 'Default to inches
		g_NoElbowTape = False
		
		
		g_sPathJOBST = fnPathJOBST()
		
	End Sub
	
	Private Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
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
	
	Private Function min(ByRef nFirst As Object, ByRef nSecond As Object) As Object
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
	
	Private Sub PR_CalculateFromMMs(ByRef Index As Short)
		'Calculate the reduction based on the given lengths
		'and the given MMs
		
		'If no change has been made to either length or MMs then exit
		If Val(EditTapes(Index).Text) = g_nLengths(Index) And Val(EditMMs(Index).Text) = g_iMMs(Index) Then Exit Sub
		
		Dim sConversion As String
		Dim nLength As Double
		Dim iPrevVal, ii, iVal As Short
		Dim iGrams, iReduction, iMMs As Short
		
		'Calculate grams
		'NB allows use of decimal MMs to fudge value
		'
		nLength = ARMEDDIA1.fnDisplayToInches(CDbl(EditTapes(Index).Text))
		iGrams = Int(nLength * Val(EditMMs(Index).Text))
		iMMs = Int(Val(EditMMs(Index).Text))
		
		'Get conversion string based on modulus
		sConversion = ""
		For ii = 0 To 17
			If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
				sConversion = g_MATERIAL.Conversion_Renamed(ii)
				Exit For
			End If
		Next ii
		
		If sConversion = "" Then
			MsgBox("Fabric Modulus not found in conversion chart", 16, "ARM Edit")
			cmdCancel_Click(cmdCancel, New System.EventArgs())
			End
		End If
		
		iPrevVal = 0
		For ii = 0 To 22
			iVal = Val(Mid(sConversion, (ii * 4) + 1, 4))
			If iVal >= iGrams Then Exit For
			iPrevVal = iVal
		Next ii
		
		Select Case ii
			Case 0
				'Minimum reduction
				iReduction = 10
			Case 23
				'Maximum reduction
				iReduction = 32
			Case Else
				'Get reduction closest to given grams
				If (iGrams - iPrevVal) < (iVal - iGrams) Then
					iReduction = ii + 9
				Else
					iReduction = ii + 10
				End If
		End Select
		
		'Modify stored values
		g_iMMs(Index) = iMMs
		g_nLengths(Index) = Val(EditTapes(Index).Text)
		g_iRed(Index) = iReduction
		g_iGms(Index) = iGrams
		
		'Change display
		EditMMs(Index).Text = CStr(iMMs) 'Reset MMs to integer value
		EditGrams(Index).Text = Str(iGrams)
		EditReductions(Index).Text = Str(iReduction)
		
		'Show that this tape has been modified
		g_iChanged(Index) = 1
		
	End Sub
	
	Private Sub PR_CalculateFromReduction(ByRef Index As Short, ByRef iReduction As Short)
		
		'Back calculate from the given reduction revised grams and mms
		
		Dim sConversion As String
		Dim nLength As Double
		Dim iPrevVal, ii, iVal As Short
		Dim iGrams, iMMs As Short
		
		'Quiet exit for bad data
		nLength = ARMEDDIA1.fnDisplayToInches(CDbl(EditTapes(Index).Text))
		If nLength <= 0 Then Exit Sub
		
		'Get conversion string based on modulus
		sConversion = ""
		For ii = 0 To 17
			If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
				sConversion = g_MATERIAL.Conversion_Renamed(ii)
				Exit For
			End If
		Next ii
		
		If sConversion = "" Then
			MsgBox("Fabric Modulus not found in conversion chart", 16, "ARM Edit")
			cmdCancel_Click(cmdCancel, New System.EventArgs())
			End
		End If
		
		
		'Get grams from conversion string based on given reduction
		'N.B. Reductions only run from 10 to 32
		'
		Select Case iReduction
			Case Is < 10
				'Use lowest available grams for those less than 10
				iGrams = Val(Mid(sConversion, 1, 4))
			Case Is > 32
				'Use highest available grams for those greater than 32
				iGrams = Val(Mid(sConversion, (22 * 4) + 1, 4))
			Case Else
				iGrams = Val(Mid(sConversion, ((iReduction - 10) * 4) + 1, 4))
		End Select
		
		'Back calculate MMs from grams and length
		iMMs = iGrams / nLength
		
		'Modify stored values
		g_iMMs(Index) = iMMs
		g_nLengths(Index) = Val(EditTapes(Index).Text)
		g_iRed(Index) = iReduction
		g_iGms(Index) = iGrams
		
		'Change display
		EditMMs(Index).Text = CStr(iMMs) 'Reset MMs to integer value
		EditGrams(Index).Text = Str(iGrams)
		EditReductions(Index).Text = Str(iReduction)
		
		'Show that this tape has been modified
		g_iChanged(Index) = 1
		
	End Sub
	
	Private Sub PR_CalcWristBlendingProfile(ByRef xyTSt As XY, ByRef xyTEnd As XY, ByRef xyBSt As XY, ByRef xyBEnd As XY, ByRef ReturnProfile As Curve, ByRef xyThumbPalm As XY)
		'Procedure to return a smooth inflecting curve between two points
		'
		'N.B. Parameters are given in the order of decreasing Y
		'
		Dim nR2, aInc, aA3, aA1, nL2, nL1, nL3, aA2, rAngle, nR1, nTol As Double
		Dim xyBotSt, xyTopSt, xyCenter, xyMidPoint, xyTopEnd, xyBotEnd As XY
		Dim xyR1, xyR2 As XY
		Dim nA, nThirdOfL2, aAngle As Double
		Dim xyPt2, xyPt1, xyInt As XY
		Dim BottomIsArc, MirrorResult, ii, Direction, TopIsArc As Short
		Dim Intersection As Double
		'UPGRADE_WARNING: Lower bound of array xyTmp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim xyTmp(10) As XY
		
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
			ReturnProfile.y(1) = xyTSt.y
			ReturnProfile.X(2) = xyBEnd.X
			ReturnProfile.y(2) = xyBEnd.y
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
			If nA - nR2 > nTol Then
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
		
		
		ReturnProfile.n = 10
		ReturnProfile.X(1) = xyTopSt.X
		ReturnProfile.y(1) = xyTopSt.y
		ReturnProfile.X(10) = xyBotEnd.X
		ReturnProfile.y(10) = xyBotEnd.y
		
		For ii = 2 To 9
			ReturnProfile.X(ii) = xyTmp(ii).X
			ReturnProfile.y(ii) = xyTmp(ii).y
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
	
	Private Sub PR_CancelProfileEdits()
		Dim ii As Short
		
		If g_ReDrawn = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Cancel", g_sSide)
			'Delete Curve Copy
			BDYUTILS.PR_PutLine("HANDLE hTempCurv;")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempETS.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "DeleteEntity(hTempCurv);")
			If g_OffFold Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempLFS.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "DeleteEntity(hTempCurv);")
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "ViewRedraw" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
		End If
	End Sub
	
	Private Sub PR_ChangeEOS(ByRef xyPt As XY, ByRef xyPt1 As XY)
		'The EOS is given by the UID txtUIDEOS
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEOS = UID (" & QQ & "find" & QC & Val(txtUIDEOS.Text) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEOS) {")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "          SetGeometry (hEOS," & xyPt.X & "," & xyPt.y & "," & xyPt1.X & "," & xyPt1.y & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "          }")
		
		
	End Sub
	
	Private Sub PR_DrawAndCommitProfileEdits()
		'Create a macro to commit the changes to the Original profile
		
		Static xyPt, xyPt1 As XY
		Static ii, iProfileVertex As Short
		Static ProfileChanged As Short
		Static nFiguredValue As Double
		Static nOffset, aAngle, nLength As Double
		Static sGms, sLen, sMM, sRed As String
		Static sPackedGms, sPackedLengths, sPackedMMs, sPackedRed As String
		Static iGms, iLen, iMM, iRed As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Manual Glove", g_sSide)
		BDYUTILS.PR_PutLine("HANDLE  hCurvETS, hCurvLFS, hTempCurv, hArm, hChan, hEnt, hEOS, hPALMER, hPALM6;")
		
		ARMDIA1.PR_SetLayer("Construct")
		
		'Delete Curve Copy (Only if redraw has been used)
		If g_ReDrawn = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempETS.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "DeleteEntity(hTempCurv);")
			If g_OffFold Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempLFS.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "DeleteEntity(hTempCurv);")
			End If
		End If
		
		'Modify Original curve  (need not have been redrawn first)
		ProfileChanged = False
		
		If g_OffFold Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hCurvETS = UID (" & QQ & "find" & QC & Val(txtUIDETS.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hCurvLFS = UID (" & QQ & "find" & QC & Val(txtUIDLFS.Text) & ");")
			If g_BlendingCurveChanged Then
				For ii = 1 To 8
					iProfileVertex = g_iVertexETSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & g_ETSBlendProfile.X(ii + 1) & "," & g_ETSBlendProfile.y(ii + 1) & ");")
				Next ii
				For ii = 1 To 8
					iProfileVertex = g_iVertexLFSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvLFS," & iProfileVertex & "," & g_LFSBlendProfile.X(ii + 1) & "," & g_LFSBlendProfile.y(ii + 1) & ");")
				Next ii
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hPALM6 = UID (" & QQ & "find" & QC & Val(txtUIDPALM6.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "SetGeometry(hPALM6," & QQ & "xmarker" & QC & xyPALM6.X & CC & xyPALM6.y & ", 0.0625, 0.0625, 0);")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hPALMER = UID (" & QQ & "find" & QC & Val(txtUIDPALMER.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "SetGeometry(hPALMER," & QQ & "xmarker" & QC & xyPALMER.X & CC & xyPALMER.y & ", 0.0625,0.0625, 0);")
				ARMDIA1.PR_PutTapeLabel(g_iFirstTape, ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iFirstTape)), g_iMMs(g_iFirstTape), g_iGms(g_iFirstTape), g_iRed(g_iFirstTape))
				ProfileChanged = True
			End If
			
			For ii = 0 To 17
				If g_iChanged(ii) < 0 Or g_iChanged(ii) > 0 Then
					iProfileVertex = g_iVertexETSMap(ii)
					'nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
					nFiguredValue = FN_FiguredValue(g_nLengths(ii), g_iRed(ii))
					xyPt.y = xyProfileETS(iProfileVertex - 1).y
					xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
					iProfileVertex = g_iVertexLFSMap(ii)
					xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
					xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvLFS," & iProfileVertex & "," & xyPt1.X & "," & xyPt1.y & ");")
					ARMDIA1.PR_PutTapeLabel(ii, ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)), g_iMMs(ii), g_iGms(ii), g_iRed(ii))
					
					ProfileChanged = True
				End If
			Next ii
			
			
			'Modify EOS (Keeps it tidy for Zipper module)
			If g_GloveType = GLOVE_ELBOW And g_iLastTape = ELBOW_TAPE And ((g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0) Or (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0)) Then
				'Glove to elbow, if either the last tape or tape befor the last has changed
				iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape - 1), g_iRed(g_iLastTape - 1))
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
				
				iProfileVertex = g_iVertexETSMap(g_iLastTape)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape)) * (Abs(g_iRed(g_iLastTape) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape), g_iRed(g_iLastTape))
				xyPt1.y = xyProfileETS(iProfileVertex - 1).y
				xyPt1.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
				
				aAngle = System.Math.Abs(90 - ARMDIA1.FN_CalcAngle(xyPt1, xyPt))
				If Val(txtAge.Text) > 10 Then nOffset = 0.75 Else nOffset = 0.375
				nLength = nOffset / System.Math.Cos(aAngle * (PI / 180))
				ARMDIA1.PR_CalcPolar(xyPt1, ARMDIA1.FN_CalcAngle(xyPt1, xyPt), nLength, xyPt)
				
				'Reflect point in X= xyTemplate.
				xyPt1.X = xyOtemplate.X
				xyPt1.y = xyPt.y
				xyPt1.X = xyOtemplate.X + (ARMDIA1.FN_CalcLength(xyPt1, xyPt) * g_Direction)
				
				PR_ChangeEOS(xyPt, xyPt1)
				
			ElseIf g_GloveType = GLOVE_AXILLA And g_EOSType = ARM_FLAP And (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0) Then 
				'Glove to axilla with a flap, if the tape befor last has change.
				iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape - 1), g_iRed(g_iLastTape - 1))
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
				iProfileVertex = g_iVertexLFSMap(g_iLastTape - 1)
				xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
				xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
				
				PR_ChangeEOS(xyPt, xyPt1)
				
			ElseIf g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0 Then 
				'All others, if last tape has changed.
				'in this case we can reuse xyPt & xyPt1 set above
				PR_ChangeEOS(xyPt, xyPt1)
				
			End If
			
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hCurvETS = UID (" & QQ & "find" & QC & Val(txtUIDETS.Text) & ");")
			If g_BlendingCurveChanged Then
				For ii = 1 To 8
					iProfileVertex = g_iVertexETSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & g_ETSBlendProfile.X(ii + 1) & "," & g_ETSBlendProfile.y(ii + 1) & ");")
				Next ii
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hPALM6 = UID (" & QQ & "find" & QC & Val(txtUIDPALM6.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "SetGeometry(hPALM6," & QQ & "xmarker" & QC & xyPALM6.X & CC & xyPALM6.y & ", 0.0625,0.0625, 0);")
				ARMDIA1.PR_PutTapeLabel(g_iFirstTape, ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iFirstTape)), g_iMMs(g_iFirstTape), g_iGms(g_iFirstTape), g_iRed(g_iFirstTape))
				ProfileChanged = True
			End If
			
			For ii = 0 To 17
				If g_iChanged(ii) < 0 Or g_iChanged(ii) > 0 Then
					'nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + .125) - g_nInsertSize) / 2)
					nFiguredValue = FN_FiguredValue(g_nLengths(ii), g_iRed(ii))
					iProfileVertex = g_iVertexETSMap(ii)
					xyPt.y = xyProfileETS(iProfileVertex - 1).y
					xyPt.X = xyOtemplate.X - (nFiguredValue) * g_Direction
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
					ARMDIA1.PR_PutTapeLabel(ii, ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)), g_iMMs(ii), g_iGms(ii), g_iRed(ii))
					ProfileChanged = True
				End If
			Next ii
			
			
			'Modify EOS (Keeps it tidy for Zipper module)
			If g_GloveType = GLOVE_ELBOW And g_iLastTape = ELBOW_TAPE And ((g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0) Or (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0)) Then
				'Glove to elbow, if either the last tape or tape befor the last has changed
				iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + .125) - (g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape - 1), g_iRed(g_iLastTape - 1))
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - ((nFiguredValue) * g_Direction)
				
				iProfileVertex = g_iVertexETSMap(g_iLastTape)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape)) * (Abs(g_iRed(g_iLastTape) - 100) / 100)) + .125) - (g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape), g_iRed(g_iLastTape))
				xyPt1.y = xyProfileETS(iProfileVertex - 1).y
				xyPt1.X = xyOtemplate.X - ((nFiguredValue) * g_Direction)
				
				aAngle = System.Math.Abs(90 - ARMDIA1.FN_CalcAngle(xyPt1, xyPt))
				If Val(txtAge.Text) > 10 Then nOffset = 0.75 Else nOffset = 0.375
				nLength = nOffset / System.Math.Cos(aAngle * (PI / 180))
				ARMDIA1.PR_CalcPolar(xyPt1, ARMDIA1.FN_CalcAngle(xyPt1, xyPt), nLength, xyPt)
				
				'Set other point to lie on the fold
				xyPt1.X = xyOtemplate.X
				xyPt1.y = xyPt.y
				
				PR_ChangeEOS(xyPt, xyPt1)
				
			ElseIf g_GloveType = GLOVE_AXILLA And g_EOSType = ARM_FLAP And (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0) Then 
				'Glove to axilla with a flap, if the tape befor last has change.
				iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + .125) - (g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iLastTape - 1), g_iRed(g_iLastTape - 1))
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - (nFiguredValue * g_Direction)
				xyPt1.y = xyPt.y
				xyPt1.X = xyOtemplate.X
				
				PR_ChangeEOS(xyPt, xyPt1)
				
			ElseIf g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0 Then 
				'All others, if last tape has changed.
				'in this case we can reuse xyPt  set above
				xyPt1.X = xyOtemplate.X
				xyPt1.y = xyPt.y
				PR_ChangeEOS(xyPt, xyPt1)
				
			End If
			
		End If
		
		
		'Update contracture if it has been changed
		If txtContracture.Text <> g_sOriginalContracture And g_NoElbowTape = False Then
			ARMEUTIL.PR_DeleteByID(g_sID & "Contracture")
			PR_DrawContracture()
		End If
		
		'Add revised lining if it has been changed
		If txtLining.Text <> g_sOriginalLining And g_NoElbowTape = False Then
			PR_DrawLining()
		End If
		
		'Update Origin Marker (only if Profile has been changed)
		If ProfileChanged = True Then
			
			For ii = 1 To 17
				iLen = g_nLengths(ii) * 10 'Shift decimal place
				iMM = g_iMMs(ii)
				iRed = g_iRed(ii)
				iGms = g_iGms(ii)
				
				If iLen <> 0 Then
					sLen = New String(" ", 3)
					sLen = RSet(Trim(Str(iLen)), Len(sLen))
					sMM = New String(" ", 3)
					sMM = RSet(Trim(Str(iMM)), Len(sMM))
					sRed = New String(" ", 3)
					sRed = RSet(Trim(Str(iRed)), Len(sRed))
					sGms = New String(" ", 3)
					sGms = RSet(Trim(Str(iGms)), Len(sGms))
				Else
					sLen = New String(" ", 3)
					sMM = New String(" ", 3)
					sRed = New String(" ", 3)
					sGms = New String(" ", 3)
				End If
				
				sPackedLengths = sPackedLengths & sLen
				sPackedMMs = sPackedMMs & sMM
				sPackedRed = sPackedRed & sRed
				sPackedGms = sPackedGms & sGms
				
			Next ii
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hArm = UID (" & QQ & "find" & QC & Val(txtUIDPALMER.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (!hArm)Exit(%cancel," & QQ & "Can't find PALMER marker to Update" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "TapeLengthsPt1" & QCQ & sPackedLengths & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "TapeMMs" & QCQ & sPackedMMs & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "Grams" & QCQ & sPackedGms & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "Reduction" & QCQ & sPackedRed & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "Contracture" & QCQ & txtContracture.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hArm, " & QQ & "Lining" & QCQ & txtLining.Text & QQ & ");")
			
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "ViewRedraw" & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_DrawContracture()
		Dim xyContractETS, xyContractLFS, xyText As XY
		Dim iElbowVertex As Short
		Dim nFiguredValue, nContractureWidth, nProfileOffset As Double
		Dim sContracture As String
		'UPGRADE_WARNING: Arrays in structure Contracture may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Contracture As Curve
		
		If g_NoElbowTape Or txtContracture.Text = "" Then Exit Sub
		
		'Get elbow position on profile
		iElbowVertex = g_iVertexETSMap(g_iElbowTape)
		xyText.y = xyProfileETS(iElbowVertex - 1).y
		
		If g_OffFold Then
			nFiguredValue = (((ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iElbowTape)) * (System.Math.Abs(g_iRed(g_iElbowTape) - 100) / 100)) - (2 * g_nInsertSize)) / 2)
			xyContractETS.y = xyProfileETS(iElbowVertex - 1).y
			xyContractETS.X = xyOtemplate.X - (nFiguredValue / 2) * g_Direction
			xyContractLFS.y = xyProfileLFS(iElbowVertex - 1).y
			xyContractLFS.X = xyOtemplate.X + (nFiguredValue / 2) * g_Direction
			xyText.X = xyOtemplate.X
		Else
			nFiguredValue = (((ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iElbowTape)) * (System.Math.Abs(g_iRed(g_iElbowTape) - 100) / 100)) - g_nInsertSize) / 2)
			xyContractETS.y = xyProfileETS(iElbowVertex - 1).y
			xyContractETS.X = xyOtemplate.X - nFiguredValue * g_Direction
			xyContractLFS.y = xyProfileLFS(iElbowVertex - 1).y
			xyContractLFS.X = xyOtemplate.X
			xyText.X = xyOtemplate.X - (nFiguredValue / 2)
		End If
		
		'Contracture width
		sContracture = VB.Left(txtContracture.Text, 2)
		Select Case sContracture
			Case "No"
				'Just in case
				Exit Sub
			Case "10"
				nContractureWidth = 0.5
			Case "36"
				nContractureWidth = 1
			Case "71"
				nContractureWidth = 1.5
		End Select
		
		'Create contracture polyline
		Contracture.n = 5
		
		'Start at Thumbside
		Contracture.X(1) = xyContractETS.X
		Contracture.y(1) = xyContractETS.y
		
		'To the top
		Contracture.y(2) = xyContractETS.y + (nContractureWidth / 2)
		Contracture.X(2) = xyContractETS.X + ((xyContractLFS.X - xyContractETS.X) / 2)
		
		'To Little finger side
		Contracture.X(3) = xyContractLFS.X
		Contracture.y(3) = xyContractLFS.y
		
		'To the bottom
		Contracture.y(4) = Contracture.y(2) - nContractureWidth
		Contracture.X(4) = Contracture.X(2)
		
		'Close to the thumbside
		Contracture.X(5) = xyContractETS.X
		Contracture.y(5) = xyContractETS.y
		
		'Draw contracture
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_DrawPoly(Contracture)
		ARMDIA1.PR_AddEntityID(g_sID, "", "Contracture")
		ARMDIA1.PR_SetTextData(2, 32, -1, -1, -1)
		ARMDIA1.PR_DrawText("Remove For Contracture", xyText, 0.125)
		
	End Sub
	
	Private Sub PR_DrawLining()
		Dim xyElbow As XY
		Dim iElbowVertex As Short
		Dim nFiguredValue As Double
		
		If g_NoElbowTape Or txtLining.Text = "" Then Exit Sub
		
		'Get elbow position on profile
		iElbowVertex = g_iVertexETSMap(g_iElbowTape)
		xyElbow.y = xyProfileETS(iElbowVertex - 1).y
		
		If g_OffFold Then
			xyElbow.X = xyOtemplate.X
		Else
			nFiguredValue = (((ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iElbowTape)) * (System.Math.Abs(g_iRed(g_iElbowTape) - 100) / 100)) - g_nInsertSize) / 2)
			xyElbow.X = xyOtemplate.X - (nFiguredValue / 2) * g_Direction
		End If
		
		'Draw Lining Text
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_SetTextData(2, 8, -1, -1, -1)
		ARMDIA1.PR_DrawText(txtLining, xyElbow, 0.125)
		
	End Sub
	
	Private Sub PR_DrawProfileEdits()
		'Create a macro to copy part of the arm profile and
		'modify this copy
		'Modifications will only be committed to the original profile on "Finish"
		'The exception to this is the ShortArm where due to DRAFIX bugs the edits
		'have to be done directly on the profile.
		
		Dim xyNull, xyPt, xyPt1 As XY
		Dim ii, iProfileVertex As Short
		Dim nFiguredValue As Double
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Profile Edits", g_sSide)
		BDYUTILS.PR_PutLine("HANDLE  hCurvETS, hCurvLFS;")
		
		
		If g_ReDrawn <> True Then
			'Make Copy of curve (First time through only)
			
			ARMDIA1.PR_SetLayer("Construct")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hCurvETS = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
			
			For ii = 0 To g_iProfileETS
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Print(fNum, CC & xyProfileETS(ii).X & CC & xyProfileETS(ii).y)
			Next ii
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, ");")
			
			If g_OffFold Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hCurvLFS = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
				
				For ii = 0 To g_iProfileLFS
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Print(fNum, CC & xyProfileLFS(ii).X & CC & xyProfileLFS(ii).y)
				Next ii
				
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, ");")
			End If
			
		Else
			'Modify Curve Copy
			If txtUIDTempETS.Text <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hCurvETS = UID (" & QQ & "find" & QC & Val(txtUIDTempETS.Text) & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If g_OffFold Then PrintLine(fNum, "hCurvLFS = UID (" & QQ & "find" & QC & Val(txtUIDTempLFS.Text) & ");")
			End If
		End If
		
		'Loop through all tape values.  If changed then modify profile copy
		If g_OffFold Then
			If g_iChanged(g_iFirstTape) > 0 Or g_iChanged(g_iFirstTape + 1) > 0 Then
				g_iChanged(g_iFirstTape) = 0
				
				g_BlendingCurveChanged = True
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape)) * (Abs(g_iRed(g_iFirstTape) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iFirstTape), g_iRed(g_iFirstTape))
				xyPALM6.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
				xyPALMER.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
				
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape + 1)) * (Abs(g_iRed(g_iFirstTape + 1) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iFirstTape + 1), g_iRed(g_iFirstTape + 1))
				iProfileVertex = g_iVertexETSMap(g_iFirstTape + 1)
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
				xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
				xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
				
				PR_CalcWristBlendingProfile(xyPALM4, xyPALM5, xyPALM6, xyPt, g_ETSBlendProfile, xyNull)
				For ii = 1 To 8
					iProfileVertex = g_iVertexETSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & g_ETSBlendProfile.X(ii + 1) & "," & g_ETSBlendProfile.y(ii + 1) & ");")
				Next ii
				
				PR_CalcWristBlendingProfile(xyPALM3, xyPALM2, xyPALMER, xyPt1, g_LFSBlendProfile, xyNull)
				For ii = 1 To 8
					iProfileVertex = g_iVertexLFSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvLFS," & iProfileVertex & "," & g_LFSBlendProfile.X(ii + 1) & "," & g_LFSBlendProfile.y(ii + 1) & ");")
				Next ii
				
			End If
			For ii = 0 To 17
				If g_iChanged(ii) > 0 Then
					g_iChanged(ii) = -1
					iProfileVertex = g_iVertexETSMap(ii)
					'nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + .375) - (2 * g_nInsertSize)) / 2)
					nFiguredValue = FN_FiguredValue(g_nLengths(ii), g_iRed(ii))
					xyPt.y = xyProfileETS(iProfileVertex - 1).y
					xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
					iProfileVertex = g_iVertexLFSMap(ii)
					xyPt.y = xyProfileLFS(iProfileVertex - 1).y
					xyPt.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvLFS," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
				End If
			Next ii
		Else
			If g_iChanged(g_iFirstTape) > 0 Or g_iChanged(g_iFirstTape + 1) > 0 Then
				' g_iChanged(g_iFirstTape) = -1
				g_iChanged(g_iFirstTape) = 0
				g_BlendingCurveChanged = True
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape)) * (Abs(g_iRed(g_iFirstTape) - 100) / 100)) + .125) - g_nInsertSize) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iFirstTape), g_iRed(g_iFirstTape))
				xyPALM6.X = xyOtemplate.X - (nFiguredValue * g_Direction)
				
				'nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape + 1)) * (Abs(g_iRed(g_iFirstTape + 1) - 100) / 100)) + .125) - g_nInsertSize) / 2)
				nFiguredValue = FN_FiguredValue(g_nLengths(g_iFirstTape + 1), g_iRed(g_iFirstTape + 1))
				iProfileVertex = g_iVertexETSMap(g_iFirstTape + 1)
				xyPt.y = xyProfileETS(iProfileVertex - 1).y
				xyPt.X = xyOtemplate.X - (nFiguredValue * g_Direction)
				PR_CalcWristBlendingProfile(xyPALM4, xyPALM5, xyPALM6, xyPt, g_ETSBlendProfile, xyNull)
				
				For ii = 1 To 8
					iProfileVertex = g_iVertexETSBlendMap(ii)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & g_ETSBlendProfile.X(ii + 1) & "," & g_ETSBlendProfile.y(ii + 1) & ");")
				Next ii
			End If
			For ii = 0 To 17
				If g_iChanged(ii) > 0 Then
					g_iChanged(ii) = -1
					'nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + .125) - g_nInsertSize) / 2)
					nFiguredValue = FN_FiguredValue(g_nLengths(ii), g_iRed(ii))
					iProfileVertex = g_iVertexETSMap(ii)
					xyPt.y = xyProfileETS(iProfileVertex - 1).y
					xyPt.X = xyOtemplate.X - (nFiguredValue * g_Direction)
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "SetVertex(hCurvETS," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
				End If
			Next ii
		End If
		If g_ReDrawn <> True Then
			BDYUTILS.PR_PutLine("HANDLE  hDDE;")
			
			'Poke Curve UID back to the VB program
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hDDE = Open (" & QQ & "dde" & QCQ & "manglved" & QCQ & "manglved" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Poke ( hDDE, " & QQ & "txtUIDTempETS" & QC & "MakeString(" & QQ & "long" & QC & "UID(" & QQ & "get" & QQ & ",hCurvETS)));")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If g_OffFold Then PrintLine(fNum, "Poke ( hDDE, " & QQ & "txtUIDTempLFS" & QC & "MakeString(" & QQ & "long" & QC & "UID(" & QQ & "get" & QQ & ",hCurvLFS)));")
			
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "ViewRedraw" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_GetProfileFromFile(ByRef sFileName As String)
		'Procedure to read curve data from file
		'Data is a series of x, y values.
		'In the following format
		'
		'    Line    Type
		'-----------------------------------------------
		'    1       PALMER Marker xyOTemplate
		'    2       PALM2 Marker
		'    .       .       .
		'    .       .       .
		'    5       PALM5 Marker
		'    6       PALM6 Marker
		'    5       n_LFS = number of vertices Little Finger side
		'    8  }
		'    .  }}
		'    .  }}}  Vertices of profile (Profile)
		'    .  }}
		'    n  }
		'    n_LFS+2 N_ETS = number of vertices Thumb side
		'    n+2}
		'    .  }}
		'    .  }}}  Vertices of profile (Profile)
		'    .  }}
		'    n_ETS  }
		'
		'The editable vertices are setup by use of an array that maps the
		'vertex of the profile to the tapes.
		
		Dim iPalm, ii, iVertex As Short
		Dim fFileNum As Short
		Dim nTolerance, nLength As Double
		
		fFileNum = FreeFile
		
		If FileLen(sFileName) = 0 Then
			MsgBox(sFileName & "Not found", 48)
			Exit Sub
		End If
		
		FileOpen(fFileNum, sFileName, OpenMode.Input)
		
		'Get control points
		Input(fFileNum, xyOtemplate.X)
		Input(fFileNum, xyOtemplate.y)
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPALMER. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPALMER = xyOtemplate
		
		Input(fFileNum, xyPALM2.X)
		Input(fFileNum, xyPALM2.y)
		Input(fFileNum, xyPALM3.X)
		Input(fFileNum, xyPALM3.y)
		Input(fFileNum, xyPALM4.X)
		Input(fFileNum, xyPALM4.y)
		Input(fFileNum, xyPALM5.X)
		Input(fFileNum, xyPALM5.y)
		Input(fFileNum, xyPALM6.X)
		Input(fFileNum, xyPALM6.y)
		
		'Get profile points
		g_iProfileLFS = 0
		Do While Not EOF(fFileNum)
			Input(fFileNum, g_iProfileLFS)
			g_iProfileLFS = g_iProfileLFS - 1
			For ii = 0 To g_iProfileLFS
				Input(fFileNum, xyProfileLFS(ii).X)
				Input(fFileNum, xyProfileLFS(ii).y)
			Next ii
			Input(fFileNum, g_iProfileETS)
			g_iProfileETS = g_iProfileETS - 1
			For ii = 0 To g_iProfileETS
				Input(fFileNum, xyProfileETS(ii).X)
				Input(fFileNum, xyProfileETS(ii).y)
			Next ii
		Loop 
		FileClose(fFileNum)
		
		
		'For editing purposes we must find the profile vertex that maps to
		'a tape. We therefor create an array that contains the profile vertex
		'number for each tape
		'
		
		
		'by starting at the last tape and working forward
		'we can map vertex to editable profile
		iVertex = g_iProfileETS + 1
		For ii = g_iLastTape To g_iFirstTape - 1 Step -1
			g_iVertexETSMap(ii) = iVertex
			iVertex = iVertex - 1
		Next ii
		
		'The next stage is map the blending curve vertex
		'
		iVertex = iVertex + 2
		For ii = 8 To 1 Step -1
			g_iVertexETSBlendMap(ii) = iVertex
			iVertex = iVertex - 1
		Next ii
		
		If g_OffFold Then
			
			'Map editable profile
			iVertex = g_iProfileLFS + 1
			For ii = g_iLastTape To g_iFirstTape - 1 Step -1
				g_iVertexLFSMap(ii) = iVertex
				iVertex = iVertex - 1
			Next ii
			
			'The next stage is map the blending curve vertex
			'
			iVertex = iVertex + 2
			For ii = 8 To 1 Step -1
				g_iVertexLFSBlendMap(ii) = iVertex
				iVertex = iVertex - 1
			Next ii
			
			'Reset the origin to lie on the midpoint of the wrist line
			nLength = System.Math.Abs(xyPALM6.X - xyOtemplate.X)
			If g_sSide = "Left" Then
				xyOtemplate.X = xyOtemplate.X + nLength / 2
			Else
				xyOtemplate.X = xyOtemplate.X - nLength / 2
			End If
		End If
		
		
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
	End Sub
	
	Private Sub PR_SetLayer(ByRef sNewLayer As String)
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
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		End
	End Sub
End Class
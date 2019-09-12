Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class manglove
	Inherits System.Windows.Forms.Form
	'Project:   MANGLOVE.MAK
	'Purpose:   Manual glove
	'
	'
	'Version:   3.01
	'Date:      17.Jan.96
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Dec 98     GG      Ported to VB5
	'
	'Notes:-
	'
	'
	'
	'
	'Windows API Functions Declarations
	'    Private Declare Function GetActiveWindow Lib "User" () As Integer
	'    Private Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer
	'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Private Declare Function GetNumTasks Lib "Kernel" () As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	
	
	'   'Constanst used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	'MsgBox constant
	'    Const IDCANCEL = 2
	'    Const IDYES = 6
	'    Const IDNO = 7
	
	'Open tip constants
	Const AGE_CUTOFF As Short = 10 'Years
	Const ADULT_STD_OPEN_TIP As Double = 0.5 'Inches
	Const CHILD_STD_OPEN_TIP As Double = 0.375 'Inches
    Public g_iInsertSize As Short
    Public Const g_sTapeText As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public g_OnFold As Short
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote

    Public Const g_sDialogID As String = "MANUAL Glove Dialogue"
    Public g_sFileNo As String 'The patients file no

    'XY data type to represent points
    Structure XY
        Dim X As Double
        Dim y As Double
    End Structure

    Public Structure curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim y() As Double

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim X(100)
            'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim y(100)
        End Sub
    End Structure

    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask As String
		
		'Check if data has been modified
		If FN_GloveDataChanged() Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, g_sDialogID)
			Select Case Response
				Case IDYES
					PR_UpdateDDE()
					CADGLEXT.PR_UpdateDDE_Extension()
					PR_CreateSaveMacro("c:\jobst\draw.d")
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(fnGetDrafixWindowTitleText())
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
                        Return
                    Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, g_sDialogID)
					End If
				Case IDNO
                    Return
                Case IDCANCEL
					Exit Sub
			End Select
		Else
            Return
        End If
		
	End Sub
	
	'UPGRADE_WARNING: Event cboFlaps.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFlaps_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFlaps.SelectedIndexChanged
		
		If InStr(1, CType(MainForm.Controls("cboFlaps"), Object).Text, "D") > 0 Then
			CType(MainForm.Controls("lblFlap"), Object)(4).Enabled = True
			CType(MainForm.Controls("txtWaistCir"), Object).Enabled = True
			CType(MainForm.Controls("labWaistCir"), Object).Enabled = True
		ElseIf CType(MainForm.Controls("lblFlap"), Object)(4).Enabled = True Then 
			CType(MainForm.Controls("txtWaistCir"), Object).Text = ""
			CType(MainForm.Controls("labWaistCir"), Object).Text = ""
			CType(MainForm.Controls("lblFlap"), Object)(4).Enabled = False
			CType(MainForm.Controls("txtWaistCir"), Object).Enabled = False
			CType(MainForm.Controls("labWaistCir"), Object).Enabled = False
		End If
		
		If MainForm.Visible Then CADGLEXT.PR_CalculateArmTapeReductions()
		
	End Sub
	
	Private Sub cmdCalculate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalculate.Click
		CADGLEXT.PR_CalculateArmTapeReductions()
	End Sub
	
	Private Sub cmdCalculateInsert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalculateInsert.Click
		Static iError, iIndex As Short
		If FN_ValidateData("Ignore Webs") Then
			'Common Dialogue elements
			If optFold(1).Checked = True Then g_OnFold = True Else g_OnFold = False
			PR_GetDlgHandData()
			CADGLEXT.PR_LoadReductionCharts((cboFabric.Text))
			'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert(g_iInsertSize, CALC). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If FN_CalculateInsert(g_iInsertSize, "CALC") Then
				'Translate and display
				Select Case g_iInsertSize
					Case 1 To 5
						iIndex = g_iInsertSize + 3
					Case 6
						iIndex = 3
				End Select
				
				'We need to force a calculation if the user changes
				'fold options but if this is done before we have enough
				'data we will confuse the user as they have not explititly
				'calculated.
				
				g_DataIsCalcuable = True
				
			Else
				'Set Drop down to calculate on fail
				iIndex = 0
			End If
			
			cboInsert.SelectedIndex = iIndex
			
		End If
		
	End Sub
	
	Private Function FN_CalcGlovePoints(ByRef iInsert As Short) As Short
		'Calculate the points on the glove based on the insert given
		'and the data in the dialogue
		'Note:-
		' Figure both Left and Right the same.
		' We then mirror the result in the vertical line
		' given by X = xyDatum.x
		FN_CalcGlovePoints = True
		
		Dim nFiguredPalm, nFiguredWrist As Double
		
		Dim nPIPs, nMaxWeb As Double
		Dim nFiguredPIPs, nMissingPIPs As Double
		Dim ii, nn, Response As Short
		Dim nMinDist, nThumbSlantLength, nInsert, nSlantInsert, aPalmThumbEnd, nInc As Double
		Dim nMinRadius, nThumbLineToWeb, nLength As Double
		Dim xyThumbOriginal, xyMinRadius, xyMinRadiusC1, xyThumbPalm5 As XY
		Dim iMaxLoopCount, Success, ForcedFail, iLoopCount As Short
		Dim nCheckDistance As Double
		
		nInsert = g_iInsertValue(iInsert) * EIGHTH
		nSlantInsert = 0
		nInc = 0.03125 '1/32nd of an inch
		
		'Calculate Figured DIPs and PIPs for the fingers
		For nn = 1 To 4
			If g_iFINGER_CHART(nn) = THUMB Then
				g_Finger(nn).PIP = FN_FigureFinger(THUMB, g_nFingerPIPCir(nn), iInsert)
				g_Finger(nn).DIP = FN_FigureFinger(THUMB, g_nFingerDIPCir(nn), iInsert)
			Else
				g_Finger(nn).PIP = FN_FigureFinger(PIP, g_nFingerPIPCir(nn), iInsert)
				g_Finger(nn).DIP = FN_FigureFinger(DIP, g_nFingerDIPCir(nn), iInsert)
			End If
			'If not missing then minimun value is 3/8th"
			If g_Finger(nn).PIP > 0 Then
				g_Finger(nn).PIP = g_Finger(nn).PIP + g_nFINGER_FIGURE(nn)
				If g_Finger(nn).PIP < 0.375 And g_Finger(nn).PIP > 0 Then g_Finger(nn).PIP = 0.375
				g_Finger(nn).DIP = g_Finger(nn).DIP + g_nFINGER_FIGURE(nn)
				If g_Finger(nn).DIP < 0.375 And g_Finger(nn).DIP > 0 Then g_Finger(nn).DIP = 0.375
			End If
			nFiguredPIPs = nFiguredPIPs + g_Finger(nn).PIP
		Next nn
		
		'Figure palm and wrist
		nFiguredPalm = (FN_FigureTape(PALM, g_nPalm) + g_nPALM_FIGURE) / 2
		nFiguredWrist = (FN_FigureTape(WRIST, g_nWrist) + g_nPALM_FIGURE) / 2
		
		'Thumb
		If g_iThumbStyle = 2 Then
			'Setin thumb
			g_Finger(5).PIP = FN_FigureFinger(PIP, g_nFingerPIPCir(5), iInsert)
		ElseIf g_iThumbStyle = 0 Then 
			'Normal thumbs
			g_Finger(5).PIP = FN_FigureFinger(THUMB, g_nFingerPIPCir(5), iInsert)
		Else
			'Edged thumbs
			'Do not figure
			g_Finger(5).PIP = g_nFingerPIPCir(5)
		End If
		g_Finger(5).DIP = 0
		
		'Check for missing fingers
		'If there are missing fingers then we need establish the width of the
		'PIP for the missing fingers, we use the Figured palm as the total
		'width of the PIPs, subtract the actual total PIPs and spread the rest
		'over the missing fingers
		'Note the exception for Fused Fingers
		If g_MissingFingers > 0 And Not g_FusedFingers Then
			nMissingPIPs = System.Math.Abs(nFiguredPalm - nFiguredPIPs) / g_MissingFingers
			For ii = 1 To 4
				If g_Missing(ii) = True Then g_Finger(ii).PIP = nMissingPIPs
			Next ii
		End If
		
		
		'Establish Datum point
		If g_OnFold Then
			ARMDIA1.PR_MakeXY(xyDatum, 0, CADGLEXT.FN_LengthWristToEOS())
		Else
			ARMDIA1.PR_MakeXY(xyDatum, 1, CADGLEXT.FN_LengthWristToEOS())
		End If
		
		'Missing web heights can represent amputated fingers etc
		'we need to account for these
		
		'Missing thumb web (use lowest - 0.375)
		If g_nWeb(4) = 0 Then
			g_nWeb(4) = 1000000 'Dummy high value
			For ii = 3 To 1 Step -1
				If g_nWeb(ii) < g_nWeb(4) Then g_nWeb(4) = g_nWeb(ii)
			Next ii
			g_nWeb(4) = g_nWeb(4) - 0.375
		End If
		
		'Missing middle
		If g_nWeb(3) = 0 Then g_nWeb(3) = g_nWeb(4) + 0.25
		
		'Missing Ring
		If g_nWeb(2) = 0 And g_nWeb(1) <> 0 Then g_nWeb(2) = g_nWeb(1) 'Use Little
		If g_nWeb(2) = 0 And g_nWeb(1) = 0 Then g_nWeb(2) = g_nWeb(3) 'Use Middle if no little
		
		'Missing Little
		If g_nWeb(1) = 0 Then g_nWeb(1) = g_nWeb(2) 'Use Ring
		
		'Little Finger (at web)
		nPIPs = g_Finger(1).PIP
		ARMDIA1.PR_MakeXY(xyLittle, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(1))
		
		'Little Finger at side of hand
		ARMDIA1.PR_MakeXY(xyLFS, xyDatum.X, xyLittle.Y)
		
		'Ring finger
		nPIPs = g_Finger(1).PIP + g_Finger(2).PIP
		ARMDIA1.PR_MakeXY(xyRing, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(2))
		
		'Middle finger
		nPIPs = g_Finger(1).PIP + g_Finger(2).PIP + g_Finger(3).PIP
		ARMDIA1.PR_MakeXY(xyMiddle, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(3))
		
		'Index finger
		nPIPs = g_Finger(1).PIP + g_Finger(2).PIP + g_Finger(3).PIP + g_Finger(4).PIP
		ARMDIA1.PR_MakeXY(xyIndex, xyDatum.X + nPIPs, xyMiddle.Y)
		
		'Thumb
		If g_OnFold Then
			ARMDIA1.PR_MakeXY(xyThumb, xyDatum.X + nFiguredPalm, xyDatum.Y + g_nWeb(4))
		Else
			If nPIPs <> nFiguredPalm Then
				'Avoid overflow
				ARMDIA1.PR_MakeXY(xyThumb, xyDatum.X + (((nPIPs - nFiguredPalm) / 2) + nFiguredPalm), xyDatum.Y + g_nWeb(4))
			Else
				ARMDIA1.PR_MakeXY(xyThumb, xyDatum.X + nFiguredPalm, xyDatum.Y + g_nWeb(4))
			End If
		End If
		
		'Set finger points
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyUlnar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(1).xyUlnar = xyLFS
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyRadial. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(1).xyRadial = xyLittle
		g_Finger(1).WebHt = xyLittle.Y
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyUlnar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(2).xyUlnar = xyLittle
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyRadial. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(2).xyRadial = xyRing
		g_Finger(2).WebHt = xyRing.Y
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyUlnar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(3).xyUlnar = xyRing
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyRadial. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(3).xyRadial = xyMiddle
		g_Finger(3).WebHt = xyMiddle.Y
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyUlnar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(4).xyUlnar = xyMiddle
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Finger().xyRadial. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_Finger(4).xyRadial = xyIndex
		g_Finger(4).WebHt = xyMiddle.Y
		
		
		'Palm points
		'Wrist line
		If g_OnFold Then
			ARMDIA1.PR_MakeXY(xyPalm(1), xyLFS.X, xyDatum.Y)
		Else
			If nPIPs <> nFiguredWrist Then
				'Avoids overflow
				ARMDIA1.PR_MakeXY(xyPalm(1), xyDatum.X + ((nPIPs - nFiguredWrist) / 2), xyDatum.Y)
			Else
				ARMDIA1.PR_MakeXY(xyPalm(1), xyDatum.X, xyDatum.Y)
			End If
		End If
		
		'Offset from wrist line
		If g_OnFold Then
			xyPalm(2).X = xyLFS.X
		Else
			If nPIPs <> nFiguredPalm Then
				'Avoids overflow
				xyPalm(2).X = xyLFS.X + (nPIPs - nFiguredPalm) / 2
			Else
				xyPalm(2).X = xyLFS.X
			End If
		End If
		
		'Establish maximum Web Height to get offset
		nMaxWeb = 0
		For ii = 1 To 3
			If g_nWeb(ii) > nMaxWeb Then nMaxWeb = g_nWeb(ii)
		Next ii
		
		If nMaxWeb > 3.5 Then
			xyPalm(2).Y = xyDatum.Y + 1
		Else
			xyPalm(2).Y = xyDatum.Y + 0.5
		End If
		
		'Offset from wrist at Thumb Side
		xyPalm(5).X = xyPalm(2).X + nFiguredPalm
		xyPalm(5).Y = xyPalm(2).Y
		
		'Wrist at Thumb Side
		xyPalm(6).X = xyPalm(1).X + nFiguredWrist
		xyPalm(6).Y = xyPalm(1).Y
		
		
		'Figure Lengths for fingers and thumbs
		Dim nTipCutBack As Double
		Dim iDigitType As Short
		
		If Val(txtAge.Text) <= AGE_CUTOFF Then
			nTipCutBack = CHILD_STD_OPEN_TIP
		Else
			nTipCutBack = ADULT_STD_OPEN_TIP
		End If
		
		iDigitType = FINGER_LEN
		For ii = 1 To 5
			If ii = 5 Then iDigitType = THUMB_LEN 'Swap for thumbs
			Select Case g_Finger(ii).Tip
				Case 0, 10
					'Closed tips (Figured Length)
					g_Finger(ii).Len_Renamed = FN_FigureLength(iDigitType, g_nFingerLen(ii))
				Case 1, 11
					'Standard Cut back for open tips (GivenLength - StdCutBack)
					g_Finger(ii).Len_Renamed = g_nFingerLen(ii) - nTipCutBack
				Case 2, 12
					'Open at a desired given length
					g_Finger(ii).Len_Renamed = g_nFingerLen(ii)
			End Select
		Next ii
		
		'Thumbpoints on palm
		'Angles
		aPalmThumbEnd = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbOriginal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyThumbOriginal = xyThumb 'Store original to give a message later
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPalmThumb(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPalmThumb(1) = xyThumb
		
		Dim nDrop As Double
		If g_iThumbStyle <> 1 Then
			'Normal or SetIn thumbs
			xyPalmThumb(1).Y = xyThumb.Y + 0.75
			
			'Ensure mininum distance of 3/8" from Index web
			If (xyIndex.Y - xyPalmThumb(1).Y) < 0.375 Then
				xyPalmThumb(1).Y = xyIndex.Y - 0.375
			End If
			'Ensure mininum distance of 1/8" from top of thumb curve to thumb web
			If xyPalmThumb(1).Y - xyThumb.Y < 0.125 Then
				xyThumb.Y = xyPalmThumb(1).Y - 0.125
			End If
			nThumbSlantLength = g_Finger(5).PIP / (1 / System.Math.Sqrt(2))
			
			'This loop counter ensures that the bottom of thumb curve does not go closer
			'than 0.25" to the EOS
			
			'        iMaxLoopCount = Int((((xyThumb.Y - xyPalm(6).Y) - nThumbSlantLength)) / .0625)
			'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iMaxLoopCount = BDYUTILS.min(25, Int((((xyThumb.Y - 0.25) - nThumbSlantLength)) / 0.0625))
			iLoopCount = 0
			Do  'at least once
				'Get blending curve and revise xyPalmThumb(1) X value
				PR_CalcFingerBlendingProfile(xyIndex, xyThumb, xyMiddle, g_Finger(4), g_FingerThumbBlend, xyPalmThumb(1))
				
				xyPalmThumb(2).Y = xyPalmThumb(1).Y
				xyPalmThumb(2).X = xyPalmThumb(1).X - nInsert
				
				xyPalmThumb(5).Y = xyPalmThumb(2).Y
				'           xyPalmThumb(5).X = xyPalmThumb(2).X - .125
				xyPalmThumb(5).X = xyPalmThumb(2).X - 0.25
				
				xyPalmThumb(3).Y = xyThumb.Y
				xyPalmThumb(3).X = xyPalmThumb(1).X - ((nFiguredPalm / 2) - 0.125)
				
				'Check that Point 3 is always less than point 5, if it is not
				'then make it so by an eighth
				If xyPalmThumb(5).X - xyPalmThumb(3).X < 0.125 Then xyPalmThumb(3).X = xyPalmThumb(5).X - 0.125
				
				
				'Get top thumb curve
				'We find the top curve by iteration, accepting the first
				'by changing the X value of xyPalmThumb(5)
				'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMinDist = BDYUTILS.min(System.Math.Abs(xyPalmThumb(3).Y - xyPalmThumb(5).Y), System.Math.Abs(xyPalmThumb(3).X - xyPalmThumb(5).X))
				'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyThumbPalm5 = xyPalmThumb(5) 'retain for later use
				Success = False
				ForcedFail = False
				nCheckDistance = System.Math.Abs(xyPalmThumb(3).X - xyPalmThumb(5).X)
				'Try first moving towards the center of the palm away from xyPalmThumb(2)
				While Not Success And nCheckDistance > nMinDist And Not ForcedFail
					ii = BDYUTILS.FN_BiArcCurve(xyPalmThumb(3), 90, xyPalmThumb(5), aPalmThumbEnd, g_ThumbTopCurve)
					
					If ii Then
						'Check that the distance from the thumbline is more than 3/8th from xyMiddle
						'allow adjustement to continue the other way if it is
						nLength = ARMDIA1.FN_CalcLength(g_ThumbTopCurve.xyR1, xyRing) - g_ThumbTopCurve.nR1
						nLength = ARMDIA1.round(nLength / 0.0625)
						'                    If nLength < .375 Then
						If nLength < 6 Then
							ForcedFail = True
						Else
							Success = True
						End If
					Else
						'Adjust and try again BUT!
						'Check that xyPalmThumb(5).x - xyPalmThumb(3) is not silly
						'IE stop any infinite loops.
						'                    xyPalmThumb(5).X = xyPalmThumb(5).X + nInc
						xyPalmThumb(5).X = xyPalmThumb(5).X - nInc
						If System.Math.Abs(xyPalmThumb(3).X - xyPalmThumb(5).X) > nCheckDistance Then ForcedFail = True
						nCheckDistance = System.Math.Abs(xyPalmThumb(3).X - xyPalmThumb(5).X)
					End If
				End While
				
				'If the above does not work then try moving away from the center
				'of the palm towards (but not past) xyPalmThumb(2)
				If Not Success Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyPalmThumb(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyPalmThumb(5) = xyThumbPalm5
					For nLength = xyPalmThumb(5).X To xyPalmThumb(2).X Step (nInc / 2)
						ii = BDYUTILS.FN_BiArcCurve(xyPalmThumb(3), 90, xyPalmThumb(5), aPalmThumbEnd, g_ThumbTopCurve)
						If ii Then
							'Check that the distance from the thumbline is more than 3/8th from xyRing
							'allow adjustement to continue until fail
							If (ARMDIA1.FN_CalcLength(g_ThumbTopCurve.xyR1, xyRing) - g_ThumbTopCurve.nR1) >= 0.375 Then
								Success = True
								Exit For
							Else
								'Adjust and try again
								xyPalmThumb(5).X = nLength
							End If
						Else
							'Adjust and try again
							xyPalmThumb(5).X = nLength
						End If
					Next nLength
				End If
				
				iLoopCount = iLoopCount + 1
				
				If Success Or iLoopCount >= iMaxLoopCount Then Exit Do
				
				
				'If we don't have a sucessfull top thumb line then lower
				'the top line or if this is not possible then drop xyThumb
				'Do this in increments of nDrop
				nDrop = 0.0625
				'Drop top thumb curve (Ensure a minimum of 1/8th inch)
				If (xyPalmThumb(1).Y - xyThumb.Y) >= 0.25 Then
					'Drop top curve
					xyPalmThumb(1).Y = xyPalmThumb(1).Y - 0.125
				Else
					'Drop thumb
					xyThumb.Y = xyThumb.Y - nDrop
				End If
				
				
			Loop 
			
			If Not Success Then
				MsgBox("Can't form TOP part of thumb curve on palm!. Adjust the palm curve and use the Transfer tool to update the thumb", 48, g_sDialogID)
				'Degenerate "ThumbTopCurve" to a straight line
				'UPGRADE_WARNING: Couldn't resolve default property of object xyPalmThumb(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyPalmThumb(5) = xyThumbPalm5
				'UPGRADE_WARNING: Couldn't resolve default property of object g_ThumbTopCurve.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_ThumbTopCurve.xyStart = xyPalmThumb(3)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_ThumbTopCurve.xyEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_ThumbTopCurve.xyEnd = xyThumbPalm5
				g_ThumbTopCurve.nR1 = 0
				g_ThumbTopCurve.nR2 = 0
			Else
				'Check and warn
				If ARMDIA1.FN_CalcLength(g_ThumbTopCurve.xyR1, xyLittle) - g_ThumbTopCurve.nR1 < 0.375 Then
					Response = MsgBox("The insert used creates a Top thumb curve on the palm that is closer than 3/8ths to the little finger web!" & Chr(13) & "Use YES to continue or NO to return to the dialogue.", 52, g_sDialogID)
				End If
				If xyThumbOriginal.Y <> xyThumb.Y Then
					nLength = System.Math.Abs(xyThumbOriginal.Y - xyThumb.Y)
					Response = MsgBox("Warnimg the thumb web position has been lowered by" & Str(nLength) & " inches, so that the thumb curve is 3/8ths away from closest web!" & Chr(13) & "Use YES to continue or NO to return to the dialogue.", 52, g_sDialogID)
					'Save the drop for later use with Web spacers
					g_iThumbWebDrop = Int(nLength / 0.0625) 'this will give the value in 1/16 ths
				End If
				Select Case Response
					Case IDYES
						FN_CalcGlovePoints = True
					Case IDNO, IDCANCEL
						FN_CalcGlovePoints = False
						Exit Function
				End Select
				
			End If
			
		Else
			
			'Edged thumb
			nThumbSlantLength = g_Finger(5).PIP / 2
			PR_CalcFingerBlendingProfile(xyIndex, xyThumb, xyMiddle, g_Finger(4), g_FingerThumbBlend, xyPalmThumb(1))
			
		End If
		
		xyPalmThumb(4).X = xyThumb.X
		xyPalmThumb(4).Y = xyThumb.Y - nThumbSlantLength
		
		'Palm points (Points dependant on thumb y)
		'LFS at Thumb web height
		If g_OnFold Then
			xyPalm(3).X = xyLFS.X
		Else
			xyPalm(3).X = xyPalm(2).X
		End If
		
		xyPalm(3).Y = xyThumb.Y
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPalm(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPalm(4) = xyThumb
		
	End Function
	
	Private Function FN_CalculateInsert(ByRef iInsert As Short, ByRef sType As String) As Object
		'This procedure uses the data for the finger PIP and the
		'Palm circumference to calculate the insert sizes
		'based on the procedure
		'GOP 01/02/26, Effective date 2nd June 1992
		'
		'
		'Notes:-
		'
		Dim ii, nn, Response As Short
		Dim nAge As Object
		Dim nFiguredPalm, nFiguredPIPs, nPermittedDiff As Double
		Dim sText As String
		'UPGRADE_WARNING: Couldn't resolve default property of object nAge. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nAge = Val(CType(MainForm.Controls("txtAge"), Object).Text)
		
		'Dim ifile%
		'Initially true
		'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_CalculateInsert = True
		
		'CALCULATE an insert
		'setup the parameters based on radio buttons
		'ifile = FreeFile
		'Open "C:\TMP\ERR.DAT" For Output As ifile
		'If g_FusedFingers = True Then Print #ifile, "Fused Fingers"
		
		'Loop through the inserts starting at
		'1/2" (ii=1) to 1" (ii=5) then try as a last resort 3/8" (ii=6)
		'If we find an insert that satisfies our test criteria then break the
		'for next loop
		'NB addition wrt Fused Fingers
		'
		For ii = 1 To 4
			g_nFINGER_FIGURE(ii) = 0
		Next ii
		
		Dim iStart, iEnd As Short
		If sType = "CALC" Then
			iStart = 1
			iEnd = 6
		Else
			iStart = iInsert
			iEnd = iInsert
		End If
		
		For iInsert = iStart To iEnd
			Select Case g_iInsertStyle
				Case 0
					'Normal
					If g_OnFold Then
						'Use thumb for little finger
						g_iFINGER_CHART(1) = THUMB
						g_nFINGER_FIGURE(1) = -(2 * SIXTEENTH)
						For ii = 2 To 4
							g_iFINGER_CHART(ii) = PIP
						Next ii
						g_nPALM_FIGURE = EIGHTH - (g_iInsertValue(iInsert) * EIGHTH)
					Else
						For ii = 1 To 4
							g_iFINGER_CHART(ii) = PIP
						Next ii
						g_nPALM_FIGURE = (3 * EIGHTH) - ((g_iInsertValue(iInsert) * EIGHTH) * 2)
					End If
					
				Case 1
					g_iFINGER_CHART(1) = THUMB
					For ii = 2 To 4
						g_iFINGER_CHART(ii) = PIP
					Next ii
					g_nPALM_FIGURE = (3 * EIGHTH) - (g_iInsertValue(iInsert) * EIGHTH)
					
				Case 2
					g_iFINGER_CHART(1) = THUMB
					g_iFINGER_CHART(2) = PIP
					g_iFINGER_CHART(3) = PIP
					g_iFINGER_CHART(4) = THUMB
					g_nPALM_FIGURE = (2 * EIGHTH)
					
				Case 3
					For ii = 1 To 4
						g_iFINGER_CHART(ii) = THUMB
					Next ii
					g_nPALM_FIGURE = (2 * EIGHTH)
					
			End Select
			
			'set permitted differencses
			If g_OnFold Then
				nPermittedDiff = EIGHTH
			Else
				nPermittedDiff = 2 * EIGHTH
			End If
			
			'Common figuring used latter
			g_iFINGER_CHART(5) = THUMB
			g_nWRIST_FIGURE = g_nPALM_FIGURE
			
			'Figure fingers
			nFiguredPIPs = 0
			For ii = 1 To 4
				nFiguredPIPs = nFiguredPIPs + FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert) + g_nFINGER_FIGURE(ii)
			Next ii
			
			'Figure palm
			nFiguredPalm = (FN_FigureTape(PALM, g_nPalm) + g_nPALM_FIGURE) / 2
			'Dim nValue#
			'Print #ifile, "iInsert= "; iInsert
			'Print #ifile, "nFiguredPalm = "; nFiguredPalm
			'Print #ifile, "nFiguredPIPs = "; nFiguredPIPs
			'For ii = 1 To 4
			'nValue = FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert) + g_nFINGER_FIGURE(ii)
			'Print #ifile, "Finger"; ii
			'Print #ifile, "Figured"; nValue; "Cir="; g_nFingerPIPCir(ii); "Figured= "; FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert); "Figure all="; g_nFINGER_FIGURE(ii)
			'Next ii
			'Exit "FOR" if conditions met
			If System.Math.Abs(nFiguredPalm - nFiguredPIPs) <= nPermittedDiff Or (g_MissingFingers <> 0 And Not g_FusedFingers) Then Exit For
			If sType <> "CALC" Then Exit For 'Exit if we have been given an Insert size
			
		Next iInsert
		
		'Close #ifile
		'Check to see if all of the above have been tried
		'issue a warning message and exit this sub
		If iInsert >= 7 Then
			If g_OnFold Then
				sText = "Unable to calculate an insert. Try OFF the fold"
			Else
				sText = "Unable to calculate an insert. Try again using one of the other options."
			End If
			MsgBox(sText, 16, "MANUAL Glove - Insert Calculation")
			'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FN_CalculateInsert = False
			Exit Function
		End If
		
		'Check and warn with respect to age and insert size
		'UPGRADE_WARNING: Couldn't resolve default property of object nAge. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If ((iInsert = 4 Or iInsert = 5) And nAge < 6) Then
			sText = "It is not recommended that an insert of 7/8"" or greater is given for children under 6 years old. Try using a ""one insert glove - on the fold""." & Chr(13) & "Do you wish to continue?"
			Response = MsgBox(sText, 36, "MANUAL Glove - Insert Calculation")
			Select Case Response
				Case IDYES
					'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FN_CalculateInsert = True
				Case IDNO, IDCANCEL
					'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FN_CalculateInsert = False
			End Select
		End If
		
		If (sType <> "CALC" And System.Math.Abs(nFiguredPalm - nFiguredPIPs) > nPermittedDiff) And g_MissingFingers = 0 And Not g_FusedFingers Then
			sText = "Incorrect figuring with the insert specified." & Chr(13) & "Do you wish to continue?"
			Response = MsgBox(sText, 36, "MANUAL Glove - Insert Calculation")
			Select Case Response
				Case IDYES
					'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FN_CalculateInsert = True
				Case IDNO, IDCANCEL
					'UPGRADE_WARNING: Couldn't resolve default property of object FN_CalculateInsert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FN_CalculateInsert = False
			End Select
		End If
		
		
	End Function
	
	Private Function FN_CompareWithOtherGlove(ByRef sFabric As String, ByRef Onfold As Short, ByRef iInsertSize As Short) As Short
		'Compares the current glove with the other glove
		'
		'Globals
		'     g_iInsertOtherGlv
		'     g_OnFoldOtherGlv
		'     txtFabric.Text
		'     g_sSide
		'
		Static sThisInsert, sText, sOtherInsert As String
		Static Response As Short
		
		FN_CompareWithOtherGlove = True
		sText = ""
		
		If g_iInsertOtherGlv = -1 Then Exit Function 'Other glove does not exist
		
		If sFabric <> txtFabric.Text Then
			sText = g_sSide & " glove fabric is " & sFabric & "." & Chr(13)
			sText = sText & "Other glove fabric is " & txtFabric.Text & "." & Chr(13)
		End If
		
		If g_iInsertOtherGlv <> iInsertSize Then
			sText = sText & "Insert sizes are different." & Chr(13)
			sThisInsert = Trim(ARMEDDIA1.fnInchestoText(g_iInsertValue(iInsertSize) * EIGHTH))
			If Mid(sThisInsert, 1, 1) = "-" Then sThisInsert = Mid(sThisInsert, 2)
			sOtherInsert = Trim(ARMEDDIA1.fnInchestoText(g_iInsertValue(g_iInsertOtherGlv) * EIGHTH))
			If Mid(sOtherInsert, 1, 1) = "-" Then sOtherInsert = Mid(sOtherInsert, 2)
			sText = sText & "This glove uses a " & sThisInsert & " insert, The other glove uses a " & sOtherInsert & " insert." & Chr(13)
		End If
		
		If Onfold <> g_OnFoldOtherGlv Then
			If g_sSide = "Left" And g_OnFoldOtherGlv Then sText = sText & "Right glove is Drawn on the fold" & Chr(13)
			If g_sSide = "Left" And Not g_OnFoldOtherGlv Then sText = sText & "Right glove is Drawn off the fold" & Chr(13)
			If g_sSide = "Right" And g_OnFoldOtherGlv Then sText = sText & "Left glove is Drawn on the fold" & Chr(13)
			If g_sSide = "Right" And Not g_OnFoldOtherGlv Then sText = sText & "Left glove is Drawn off the fold" & Chr(13)
		End If
		
		'Check if error has occured
		If sText <> "" Then
			If g_sSide = "Left" Then
				sText = "This glove differs from the RIGHT glove!" & Chr(13) & sText & Chr(13) & "Do you wish to continue and draw the glove anyway?"
			Else
				sText = "This glove differs from the LEFT glove!" & Chr(13) & sText & Chr(13) & "Do you wish to continue and draw the glove anyway?"
			End If
			Response = MsgBox(sText, 36, g_sDialogID)
			Select Case Response
				Case IDYES
					FN_CompareWithOtherGlove = True
				Case IDNO, IDCANCEL
					FN_CompareWithOtherGlove = False
			End Select
		End If
		
	End Function
	
	Private Function FN_GloveDataChanged() As Short
		'Function that will be used to check if data has changed
		'The first Time that it is called the function will always
		'return False
		'
		'NOTE:-
		'    The change checked for is always the state of the data the
		'    first time the function is called
		'
		'    It is intendee to be run After Link Close to save the state
		'    at that point
		'    If the user uses Close then we can check if any changes have been
		'    made
		
		
		Static sCheck, sTmp As String
		Static ii As Short
		
		'Get data from the controls
		sTmp = ""
		'    MainForm!SSTab1.Tab = 0
		
		For ii = 0 To 12
			sTmp = sTmp & txtCir(ii).Text
		Next ii
		
		For ii = 0 To 8
			sTmp = sTmp & txtLen(ii).Text
		Next ii
		
		'    MainForm!SSTab1.Tab = 1
		
		For ii = 8 To 23
			sTmp = sTmp & txtExtCir(ii).Text & mms(ii).Text
		Next ii
		
		sTmp = sTmp & txtWristPleat1.Text & txtWristPleat2.Text & txtShoulderPleat1.Text & txtShoulderPleat2.Text
		
		sTmp = sTmp & cboFlaps.Text & txtStrapLength.Text & txtFrontStrapLength.Text & txtCustFlapLength.Text & txtWaistCir.Text
		
		sTmp = sTmp & Str(optExtendTo(0).Checked) & Str(optExtendTo(1).Checked) & Str(optExtendTo(2).Checked)
		
		sTmp = sTmp & Str(optThumbStyle(0).Checked) & Str(optThumbStyle(1).Checked) & Str(optThumbStyle(2).Checked)
		
		sTmp = sTmp & Str(optProximalTape(0).Checked) & Str(optProximalTape(1).Checked)
		
		sTmp = sTmp & Str(optFold(0).Checked) & Str(optFold(1).Checked)
		
		sTmp = sTmp & cboFabric.Text & cboDistalTape.Text & cboProximalTape.Text & cboPressure.Text & cboInsert.Text
		
		sTmp = sTmp & Str(chkSlantedInserts.CheckState) & Str(chkPalm.CheckState) & Str(chkDorsal.CheckState)
		
		'Check for changes
		If sCheck = "" Then
			FN_GloveDataChanged = False
			sCheck = sTmp
		Else
			If sTmp <> sCheck Then FN_GloveDataChanged = True
		End If
		
	End Function
	
	Private Function FN_Open(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
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
		NL = Chr(10) 'The new line character
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
		
		'Initialise patient globals
		'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sFileNo = sPatientFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sLeftorRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sSide = sLeftorRight
		'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sPatient = sName
		
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
		PrintLine(fNum, "//Patient - " & g_sPatient & ", " & g_sFileNo & ", Hand - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, GLOVES - Drawing")
		
		'Define DRAFIX variables
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY     xyStart, xyOrigin, xyScale, xyO;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sFileNo, sSide, sID, sName;")
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
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Data" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "ZipperLength" & QCQ & "length" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Zipper" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "TapeLengths" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "TapeLengths2" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "TapeLengthPt1" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "curvetype" & QCQ & "string" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Leg" & QCQ & "string" & QQ & ");")
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
		
		'Set values for use futher on by other macros
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sSide = " & QQ & g_sSide & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sFileNo = " & QQ & g_sFileNo & QQ & ";")
		
		'Get Start point
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "GetUser (" & QQ & "xy" & QCQ & "Indicate Start Point" & QC & "&xyStart);")
		
		'Place a marker at the start point for later use.
		'Get a UID and create the unique 4 character start to the ID code
		'Note this is a bit dogey if the drawing contains more than 9999 entities
		ARMDIA1.PR_Setlayer("Construct")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hOrigin = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart" & CC & "0.125);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hOrigin) {")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  sID=StringMiddle(MakeString(" & QQ & "long" & QQ & ",UID(" & QQ & "get" & QQ & ",hOrigin)), 1, 4) ; ")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  while (StringLength(sID) < 4) sID = sID + " & QQ & " " & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  sID = sID + sFileNo + sSide ;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  SetDBData(hOrigin," & QQ & "ID" & QQ & ",sID);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  }")
		
		'Display Hour Glass Symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Drawing" & QQ & ");")
		
		'Set values for use futher on by other macros
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyOrigin = xyStart" & ";")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hMPD)")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "else")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")
		
		'Start drawing on correct side
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
	End Function
	Private Function FN_TranslateInsert(ByRef iListIndex As Short, ByRef iInsert As Short) As Short
		'Function to translate the cboInsert value to one that can be
		'used in a manual glove
		'It will also attempt reconcile the two
		'
		'   cboInsert.AddItem "Calculate"   0
		'   cboInsert.AddItem "1/8"         1
		'   cboInsert.AddItem "1/4"         2
		'   cboInsert.AddItem "3/8"         3      Valid Manual
		'   cboInsert.AddItem "1/2"         4      Valid Manual
		'   cboInsert.AddItem "5/8"         5      Valid Manual
		'   cboInsert.AddItem "3/4"         6      Valid Manual
		'   cboInsert.AddItem "7/8"         7      Valid Manual
		'   cboInsert.AddItem "1"           8      Valid Manual
		'   cboInsert.AddItem "1-1/8"       9
		'   cboInsert.AddItem "1-1/4"       10
		'   cboInsert.AddItem "1-3/8"       11
		'   cboInsert.AddItem "1-1/2"       12
		'   cboInsert.AddItem "1-5/8"       13
		'   cboInsert.AddItem "1-3/4"       14
		'   cboInsert.AddItem "1-7/8"       15
		'
		FN_TranslateInsert = True
		Select Case iListIndex
			Case 0 To 2, 9 To 15
				MsgBox("Not valid insert size for MANUAL-Gloves. Use a value in the range 3/8th to 1 inch.  Or use Calculate to force a calculation ", 16, g_sDialogID)
				FN_TranslateInsert = False
				iInsert = 0
			Case 4 To 8
				iInsert = iListIndex - 3
			Case 3
				iInsert = 6
		End Select
		
		
	End Function
	
	Private Function FN_ValidateData(ByRef sOption As String) As Short
		'This function is used only to make gross checks
		'for missing data.
		'It does not perform any sensibility checks on the
		'data
		'The sOption allows for a less rigourous check
		'    "Check All"
		'    "Ignore Webs"
		'
		Dim NL, sError, sWebError As String
		Dim ii, nn As Short
		Dim nThumbWebHt As Double
		
		'Initialise
		FN_ValidateData = False
		sError = ""
		NL = Chr(13)
		
		Dim sCircum(12) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(0) = "Little Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(1) = "Little Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(2) = "Ring Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(3) = "Ring Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(4) = "Middle Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(5) = "Middle Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(6) = "Index Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(7) = "Index Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(8) = "Thumb"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(9) = "Palm"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(10) = "Wrist"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(11) = "Tape 1½"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(12) = "Tape 3"
		
		Dim nTotalCir(4) As Object
		
		Dim sLengths(8) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(0) = "Little Finger tip  to web"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(1) = "Ring Finger tip to web"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(2) = "Middle Finger tip to web"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(3) = "Index Finger tip to web"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(4) = "Thumb to tip web"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(5) = "Wrist to web at Little Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(6) = "Wrist to web at Middle Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(7) = "Wrist to web at Index Finger"
		'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLengths(8) = "Wrist to web at Thumb"
		
		
		'Check circumferences
		nn = 0
		For ii = 0 To 7 Step 2
			If Val(txtCir(ii).Text) <> 0 And Val(txtCir(ii + 1).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing PIP for " & sCircum(ii) & "!" & NL
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(nn). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nTotalCir(nn) = Val(txtCir(ii).Text) + Val(txtCir(ii + 1).Text)
			nn = nn + 1
		Next ii
		
		'Settings for thumb (as not set above)
		'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nTotalCir(4) = Val(txtCir(8).Text)
		
		For ii = 8 To 10
			If Val(txtCir(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing circumference for " & sCircum(ii) & "!" & NL
			End If
		Next ii
		
		
		'Check on lengths
		For ii = 0 To 3
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(ii).Text) = 0 And nTotalCir(ii) <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing length " & sLengths(ii) & "!" & NL
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(ii).Text) <> 0 And nTotalCir(ii) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Length given but missing circumference " & sCircum(ii * 2) & "!" & NL
			End If
		Next ii
		
		'Length not required for edged thumbs
		If optThumbStyle(1).Checked <> True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(4).Text) = 0 And nTotalCir(4) <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing length " & sLengths(4) & "!" & NL
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(4).Text) <> 0 And nTotalCir(4) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Length given but missing circumference " & sCircum(4 * 2) & "!" & NL
			End If
			
		End If
		
		If sOption <> "Ignore Webs" Then
			sWebError = ""
			For ii = 5 To 8
				If Val(txtLen(ii).Text) = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sWebError = sWebError & "Missing length for " & sLengths(ii) & "!" & NL
				End If
			Next ii
			If sWebError <> "" Then
				If MsgBox(sWebError & NL & "If the missing webs represent amputated fingers then, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
					FN_ValidateData = False
					Exit Function
				End If
			End If
			
			'Check that the thumb web is lower that all the other webs
			sWebError = ""
			nThumbWebHt = Val(txtLen(8).Text)
			For ii = 5 To 7
				If ((Val(txtLen(ii).Text)) <= nThumbWebHt) And nThumbWebHt <> 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sWebError = sWebError & "Thumb web is higher than " & sLengths(ii) & "!" & NL
				End If
			Next ii
			If sWebError <> "" Then
				If MsgBox(sWebError & NL & "Use, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
					FN_ValidateData = False
					Exit Function
				End If
			End If
			
		End If
		
		If cboFabric.Text = "" Then
			sError = sError & "Fabric not given! " & NL
		Else
			If UCase(Mid(cboFabric.Text, 1, 3)) <> "POW" Then
				sError = sError & "Fabric chosen is not Powernet"
			End If
			If Val(Mid(cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid(cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
				sError = sError & "Modulus of fabric chosen is not available on Gram / Tension reduction chart"
			End If
			
		End If
		
		'Validate extension data
		If sOption = "Check All" Then
			sError = sError & CADGLEXT.FN_ValidateExtensionData()
		End If
		
		If sError <> "" Then
			MsgBox(sError, 16, g_sDialogID)
			FN_ValidateData = False
		Else
			FN_ValidateData = True
		End If
		
	End Function
	
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Dim nAge, ii, nn, jj As Short
		Dim nValue As Double
		Dim sTapesBeyondWrist, sValue As String
		
		'Disable timeout timer as a link close event has happened
		Timer1.Enabled = False
		
		'Check if MainPatientDetails exists
		'If it dosn't then exit
		If txtUidMPD.Text = "" Then
			MsgBox("No Patient details found!", 16, g_sDialogID)
			End
		End If
		
		'Scaling factor
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Hand Option buttons
		If txtSide.Text = "Left" Then
			optHand(0).Checked = True
			optHand(1).Enabled = False
			g_sSide = "Left"
		End If
		If txtSide.Text = "Right" Then
			optHand(1).Checked = True
			optHand(0).Enabled = False
			g_sSide = "Right"
		End If
		
		'Circumferences
		For ii = 0 To 10
			nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3))
			If nValue > 0 Then
				txtCir(ii).Text = CStr(nValue / 10)
				txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
			End If
		Next ii
		
		'Lengths
		For ii = 0 To 8
			nn = ii + 11
			nValue = Val(Mid(txtTapeLengths.Text, (nn * 3) + 1, 3))
			If nValue > 0 Then
				txtLen(ii).Text = CStr(nValue / 10)
				txtLen_Leave(txtLen.Item(ii), New System.EventArgs())
			End If
		Next ii
		
		'Tip Option Buttons
		'     0 => Closed tip
		'     1 => Standard Open Tip
		'     2 => Disired Open Tip
		If txtTapeLengths2.Text <> "" Then
			For ii = 0 To 6
				nValue = Val(Mid(txtTapeLengths2.Text, ii + 1, 1))
				If ii = 0 Then
					optLittleTip(nValue).Checked = True
				ElseIf ii = 1 Then 
					optRingTip(nValue).Checked = True
				ElseIf ii = 2 Then 
					optMiddleTip(nValue).Checked = True
				ElseIf ii = 3 Then 
					optIndexTip(nValue).Checked = True
				ElseIf ii = 4 Then 
					optThumbTip(nValue).Checked = True
				ElseIf ii = 5 Then 
					chkSlantedInserts.CheckState = nValue
				ElseIf ii = 6 Then 
					chkPalm.CheckState = nValue
				End If
			Next ii
			'Tapes beyond wrist
			'normal glove
			
			ii = 0
			sTapesBeyondWrist = Mid(txtTapeLengths2.Text, 8)
			For nn = 11 To 12
				nValue = Val(Mid(sTapesBeyondWrist, (ii * 3) + 1, 3))
				If nValue > 0 Then
					txtCir(nn).Text = CStr(nValue / 10)
					txtCir_Leave(txtCir.Item(nn), New System.EventArgs())
				End If
				ii = ii + 1
			Next nn
		End If
		
		
		'Setup for specific insert size
		cboInsert.Items.Add("Calculate")
		cboInsert.Items.Add("1/8")
		cboInsert.Items.Add("1/4")
		cboInsert.Items.Add("3/8")
		cboInsert.Items.Add("1/2")
		cboInsert.Items.Add("5/8")
		cboInsert.Items.Add("3/4")
		cboInsert.Items.Add("7/8")
		cboInsert.Items.Add("1")
		cboInsert.Items.Add("1-1/8")
		cboInsert.Items.Add("1-1/4")
		cboInsert.Items.Add("1-3/8")
		cboInsert.Items.Add("1-1/2")
		cboInsert.Items.Add("1-5/8")
		cboInsert.Items.Add("1-3/4")
		cboInsert.Items.Add("1-7/8")
		
		'Stored insert value (This is stored in 8ths)
		Dim sInsert As String
		Dim nInsert As Short
		
		'This is a flag used to allow for automatic recalculation
		'if the user switches between on and off the fold
		'This stops confusing the user with warning messages about
		'missing data when they have not requested anything to happen
		'but have simply used the Fold option buttons
		g_DataIsCalcuable = False
		
		sInsert = Mid(txtTapeLengths2.Text, 15, 2)
		If sInsert <> "" Then
			nInsert = Val(sInsert)
			If nInsert <= cboInsert.Items.Count Then
				cboInsert.SelectedIndex = nInsert
			Else
				cboInsert.SelectedIndex = 0
			End If
		Else
			cboInsert.SelectedIndex = 0
		End If
		
		If cboInsert.SelectedIndex > 0 Then
			g_DataIsCalcuable = True
			g_iInsertSize = nInsert
		End If
		
		'Reinforced DORSAL
		nValue = Val(Mid(txtTapeLengths2.Text, 17, 1))
		chkDorsal.CheckState = nValue
		
		'Thumb Options
		nValue = Val(Mid(txtTapeLengths2.Text, 18, 1))
		optThumbStyle(nValue).Checked = True
		
		'Inset Options
		nValue = Val(Mid(txtTapeLengths2.Text, 19, 1))
		optInsertStyle(nValue).Checked = True
		
		'Print Fold Options
		sValue = Mid(txtTapeLengths2.Text, 20, 1)
		chkPrintFold.CheckState = System.Windows.Forms.CheckState.Checked
		If sValue = "0" Then chkPrintFold.CheckState = System.Windows.Forms.CheckState.Unchecked
		
		'Fused fingers options
		nValue = Val(Mid(txtTapeLengths2.Text, 21, 1))
		chkFusedFingers.CheckState = nValue
		
		'Fabric
		If txtFabric.Text <> "" Then cboFabric.Text = txtFabric.Text
		
		'Get DDE data for extension
		'This procedure is in the module GLVEXTEN.BAS
		g_CalculatedExtension = False
		
		CADGLEXT.PR_GetExtensionDDE_Data()
		
		'Update the values in the glove extension module
		'This takes the values from the dialogue and a sets it up for subsequent
		'procedures in this module
		CADGLEXT.PR_GetDlgAboveWrist()
		
		'Having set all the fields with raw data we NOW!
		'refine the display by executing the relevent procedure
		For ii = 0 To 2
			If optExtendTo(ii).Checked = True Then
				CADGLEXT.PR_ExtendTo_Click((ii))
				Exit For
			End If
		Next ii
		
		'Get data for other hand to ensure that the two gloves are
		'similar
		g_iInsertOtherGlv = -1
		g_OnFoldOtherGlv = False
		If g_sSide = "Right" Then
			If Val(Mid(txtDataGC.Text, 1, 2)) > 0 Then
				g_iInsertOtherGlv = Val(Mid(txtDataGC.Text, 1, 2))
				If Val(Mid(txtDataGC.Text, 3, 2)) > 0 Then g_OnFoldOtherGlv = True
			End If
		Else
			If Val(Mid(txtDataGC.Text, 5, 2)) <> 0 Then
				g_iInsertOtherGlv = Val(Mid(txtDataGC.Text, 5, 2))
				If Val(Mid(txtDataGC.Text, 7, 2)) > 0 Then g_OnFoldOtherGlv = True
			End If
		End If
		
		
		'Setup to test for changes
		ii = FN_GloveDataChanged()
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
		'Show after link is closed
		Show()
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub manglove_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		'Hide during DDE transfer
		Hide()
		
		'Enable Timeout timer
		'Use 6 seconds (a 1/10th of a minute).
		'Timeout is disabled in Link close
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
		MainForm = Me
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MainForm.Width)) / 2)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MainForm.Height)) / 2)
		
		'Get the path to the JOBST system installation directory
		g_sPathJOBST = fnPathJOBST()
		
		'Set up fabric list
		LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\FABRIC.DAT")
		
		'Clear DDE transfer boxes
		cboFabric.Text = ""
		txtUidGlove.Text = ""
		txtTapeLengths.Text = ""
		txtTapeLengths2.Text = ""
		txtTapeLengthPt1.Text = ""
		txtSide.Text = ""
		txtWorkOrder.Text = ""
		txtUidGC.Text = ""
		txtFabric.Text = ""
		txtUidMPD.Text = ""
		txtGrams.Text = ""
		txtReduction.Text = ""
		txtDataGlove.Text = ""
		txtDataGC.Text = ""
		txtFlap.Text = ""
		txtTapeMMs.Text = ""
		txtWristPleat.Text = ""
		txtShoulderPleat.Text = ""
		
		'Default units
		'g_nUnitsFac = 1             'Metric
		g_nUnitsFac = 10 / 25.4 'Inches
		
		'Set Default Tips to Closed
		optLittleTip(0).Checked = 1
		optRingTip(0).Checked = 1
		optMiddleTip(0).Checked = 1
		optIndexTip(0).Checked = 1
		optThumbTip(0).Checked = 1
		
		
		'Setup display inches grid
		grdInches.set_ColWidth(0, 615)
		grdInches.set_ColAlignment(0, 2)
		For ii = 0 To 15
			'        grdInches.RowHeight(ii) = 266
			grdInches.set_RowHeight(ii, 270)
		Next ii
		
		'Setup display of results grid
		For ii = 0 To 1
			grdDisplay.set_ColWidth(ii, 488)
			grdDisplay.set_ColAlignment(ii, 2)
		Next ii
		
		For ii = 0 To 15
			'        grdDisplay.RowHeight(ii) = 266
			grdDisplay.set_RowHeight(ii, 270)
		Next ii
		
		'ListBoxes
		'g_sTapeText, From GLVEXTEN.BAS
		
		cboDistalTape.Items.Add("1st")
		For ii = 2 To 17
			cboDistalTape.Items.Add(LTrim(Mid(g_sTapeText, (ii * 3) + 1, 3)))
		Next ii
		cboDistalTape.SelectedIndex = 0
		
		cboProximalTape.Items.Add("Last")
		For ii = 17 To 2 Step -1
			cboProximalTape.Items.Add(LTrim(Mid(g_sTapeText, (ii * 3) + 1, 3)))
		Next ii
		cboProximalTape.SelectedIndex = 0
		
		cboFlaps.Items.Add("Style A ")
		cboFlaps.Items.Add("Style B ")
		cboFlaps.Items.Add("Style C ")
		cboFlaps.Items.Add("Style D ")
		cboFlaps.Items.Add("Style E ")
		cboFlaps.Items.Add("Style F ")
		cboFlaps.Items.Add("Raglan A")
		cboFlaps.Items.Add("Raglan B")
		cboFlaps.Items.Add("Raglan C")
		cboFlaps.Items.Add("Raglan D")
		cboFlaps.Items.Add("Raglan E")
		cboFlaps.Items.Add("Raglan F")
		
		cboPressure.Items.Add("15")
		cboPressure.Items.Add("20")
		cboPressure.Items.Add("25")
		cboPressure.SelectedIndex = 0
		
		'Default side
		g_sSide = "Left"
		
		'load
		CADGLEXT.PR_LoadPowernetChart()
		
	End Sub
	
	Private Sub labCir_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labCir.DoubleClick
		Dim Index As Short = labCir.GetIndex(eventSender)
		'Procedure to disable the circumference fields
		'to enable the use to select only the relevent
		'data to be drawn
		'Saves having to delele and re-enter to get variations
		'on a glove
		Dim iDIP, iPIP As Short
		If System.Drawing.ColorTranslator.ToOle(labCir(Index).ForeColor) = &HC0C0C0 Then
			Select Case Index
				Case 0 To 7, 11 To 12
					PR_EnableCir(Index)
				Case 8
					PR_EnableCir(Index)
					'Disable thumb length
					labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					txtLen(4).Enabled = True
					lblLen(4).Enabled = True
				Case 20 To 23
					iDIP = (Index - 20) * 2
					iPIP = iDIP + 1
					labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					PR_EnableCir(iDIP)
					PR_EnableCir(iPIP)
					'Enable finger length
					labLen(Index - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					txtLen(Index - 20).Enabled = True
					lblLen(Index - 20).Enabled = True
			End Select
		Else
			Select Case Index
				Case 0 To 7, 12
					PR_DisableCir(Index)
				Case 8
					PR_DisableCir(Index)
					'Disable thumb length
					labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					txtLen(4).Enabled = False
					lblLen(4).Enabled = False
				Case 11
					PR_DisableCir(Index)
					PR_DisableCir(Index + 1)
				Case 20 To 23
					iDIP = (Index - 20) * 2
					iPIP = iDIP + 1
					labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					PR_DisableCir(iDIP)
					PR_DisableCir(iPIP)
					'Disable finger length
					labLen(Index - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					txtLen(Index - 20).Enabled = False
					lblLen(Index - 20).Enabled = False
			End Select
		End If
		
	End Sub
	
	Private Sub labLen_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labLen.DoubleClick
		Dim Index As Short = labLen.GetIndex(eventSender)
		If Index = 5 Then Exit Sub
		If System.Drawing.ColorTranslator.ToOle(labLen(Index).ForeColor) = &HC0C0C0 Then
			labLen(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
			txtLen(Index).Enabled = True
			lblLen(Index).Enabled = True
		Else
			labLen(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			txtLen(Index).Enabled = False
			lblLen(Index).Enabled = False
		End If
	End Sub
	
	Private Sub lblPleat_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPleat.DoubleClick
		Dim Index As Short = lblPleat.GetIndex(eventSender)
		'Procedure to disable the pleat fields
		'to enable the user to select only the relevent
		'data to be drawn
		'Saves having to delele and re-enter to get variations
		'on a glove
		Dim iDIP, iPIP As Short
		If System.Drawing.ColorTranslator.ToOle(lblPleat(Index).ForeColor) = &HC0C0C0 Then
			lblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
			Select Case Index
				Case 0
					txtWristPleat1.Enabled = True
				Case 1
					txtWristPleat2.Enabled = True
				Case 2
					txtShoulderPleat2.Enabled = True
				Case 3
					txtShoulderPleat1.Enabled = True
			End Select
		Else
			lblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			Select Case Index
				Case 0
					txtWristPleat1.Enabled = False
				Case 1
					txtWristPleat2.Enabled = False
				Case 2
					txtShoulderPleat2.Enabled = False
				Case 3
					txtShoulderPleat1.Enabled = False
			End Select
		End If
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim sTask As String
		Dim iInsert, iError As Short
		'Don't allow multiple clicking
		'
		OK.Enabled = False
		
		'Get data from the Dialogue
		PR_GetDlgData()
		
		If FN_ValidateData("Check All") Then
			'Display and hourglass
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			CADGLEXT.PR_LoadReductionCharts((cboFabric.Text))
			
			'These functions, FN_CalculateInsert() and FN_TranslateInsert()
			'interact with the user.
			'If they returns false then either an error has occoured
			'or the user has chosen to make changes.
			'In either case the user is returned to the dialogue
			
			If cboInsert.Text = "Calculate" Then
				If Not FN_CalculateInsert(g_iInsertSize, "CALC") Then GoTo EXIT_OK
			Else
				'Translate cboInsert.Text to correct format
				'for use below
				If Not FN_TranslateInsert((cboInsert.SelectedIndex), g_iInsertSize) Then
					GoTo EXIT_OK
				Else
					'Use the given insertsize
					If Not FN_CalculateInsert(g_iInsertSize, "USE") Then GoTo EXIT_OK
				End If
			End If
			
			'Check with the glove common
			If Not FN_CompareWithOtherGlove((cboFabric.Text), g_OnFold, g_iInsertSize) Then
				GoTo EXIT_OK
			End If
			
			'Calculate glove points based on insert given above and data.
			If Not FN_CalcGlovePoints(g_iInsertSize) Then GoTo EXIT_OK
			
			'Calculate glove extensions (if any)
			CADGLEXT.PR_CalculateExtension(xyPalm(1), xyPalm(6), g_iInsertValue(g_iInsertSize) * EIGHTH)
			
			PR_UpdateDDE()
			
			CADGLEXT.PR_UpdateDDE_Extension()
			
			'N.B. Use of local JOBST Directory C:\JOBST
			PR_CreateDrawMacro("C:\JOBST\DRAW.D", g_iInsertSize)
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				End
			Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, g_sDialogID)
			End If
		End If
		
EXIT_OK: 
		
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Show()
		Exit Sub
		
		
	End Sub
	
	'UPGRADE_WARNING: Event optExtendTo.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optExtendTo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExtendTo.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optExtendTo.GetIndex(eventSender)
			'Use Procedure in module GLVEXTEN.BAS
			If MainForm.Visible Then CADGLEXT.PR_ExtendTo_Click(Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optFold.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optFold_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFold.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optFold.GetIndex(eventSender)
			If g_DataIsCalcuable Then cmdCalculateInsert_Click(cmdCalculateInsert, New System.EventArgs())
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optHand.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optHand_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHand.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optHand.GetIndex(eventSender)
			If Index = 0 Then
				g_sSide = "Left"
			Else
				g_sSide = "Right"
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optIndexTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optIndexTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optIndexTip.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optIndexTip.GetIndex(eventSender)
			PR_CutBackTip(3, Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optInsertStyle.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optInsertStyle_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optInsertStyle.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optInsertStyle.GetIndex(eventSender)
			If Index = 1 Then
				optFold(0).Checked = True
				optFold(1).Enabled = False
			Else
				optFold(1).Enabled = True
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optLittleTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optLittleTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optLittleTip.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optLittleTip.GetIndex(eventSender)
			PR_CutBackTip(0, Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMiddleTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMiddleTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMiddleTip.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optMiddleTip.GetIndex(eventSender)
			PR_CutBackTip(2, Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optProximalTape.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optProximalTape_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optProximalTape.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optProximalTape.GetIndex(eventSender)
			'Use Procedure in module GLVEXTEN.BAS
			CADGLEXT.PR_ProximalTape_Click(Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRingTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRingTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRingTip.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optRingTip.GetIndex(eventSender)
			PR_CutBackTip(1, Index)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optThumbTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optThumbTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optThumbTip.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optThumbTip.GetIndex(eventSender)
			PR_CutBackTip(4, Index)
		End If
	End Sub
	
	Private Sub PR_CreateDrawMacro(ByRef sFileName As String, ByRef iInsertSize As Short)
		'
		' GLOBALS
		'
		'    g_sSide
		'    g_FingerThumbBlend
		'
		Dim ii As Short
		Dim xyTextAlign, xyPt2, xyPt1, xyPt3, xyTmp As XY
		Dim nAge, iiStart As Short
		Dim TmpY, nInsert, TmpX, nOffSet As Double
		'UPGRADE_WARNING: Arrays in structure WristLFSBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure FingerLFSBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim FingerLFSBlend, WristLFSBlend As Curve
		'UPGRADE_WARNING: Arrays in structure WristThumbBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim WristThumbBlend As Curve
		'UPGRADE_WARNING: Arrays in structure Reinforced may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure LFS may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure ThumbSide may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ThumbSide, LFS, Reinforced As Curve
		Dim nLength, aAngle, nTranslate As Double
		Dim sWorkOrder, sText As String
		
		
		'Initialise
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)
		
		'Draw on correct layer
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		'Get Blending profiles on thumbside
		'N.B. Modifications to xyPalmThumb(1) and xyPalmThumb(4)
		If RadialProfile.n = 1 Then iiStart = 1 Else iiStart = 2
		xyPt1.X = RadialProfile.X(iiStart)
		xyPt1.Y = RadialProfile.Y(iiStart)
		
		'    PR_CalcFingerBlendingProfile xyIndex, xyPalm(4), xyMiddle, g_Finger(4), FingerThumbBlend, xyPalmThumb(1)
		PR_CalcWristBlendingProfile(xyPalm(4), xyPalm(5), xyPalm(6), xyPt1, WristThumbBlend, xyPalmThumb(4))
		
		'Get Blending profiles on Little Finger Side
		If Not g_OnFold Then
			xyPt1.X = UlnarProfile.X(iiStart)
			xyPt1.Y = UlnarProfile.Y(iiStart)
			PR_CalcWristBlendingProfile(xyPalm(3), xyPalm(2), xyPalm(1), xyPt1, WristLFSBlend, xyTmp)
		End If
		ARMDIA1.PR_MakeXY(xyPt3, xyPalm(3).X, xyPalm(4).Y)
		PR_CalcFingerBlendingProfile(xyLFS, xyPt3, xyLittle, g_Finger(1), FingerLFSBlend, xyTmp)
		
		'Calculate other PalmThumb points based on point xyPalmThumb(1)
		'Done here so it can be handed below
		'
		nInsert = g_iInsertValue(iInsertSize) * EIGHTH
		
		'Calculate a point that will be used to align text
		ARMDIA1.PR_MakeXY(xyTextAlign, xyPalm(1).X + 0.25, xyPalm(1).Y + 1.625)
		
		'Mirror and translate for right hand side
		'To now all the data has been caculated for the LEFT
		'side
		If g_sSide = "Right" Then
			'Find farthest point from datum
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nTranslate = ARMDIA1.max(g_Finger(4).xyRadial.X, RadialProfile.X(RadialProfile.n))
			nTranslate = nTranslate - xyDatum.X
			'Fingers
			For ii = 1 To 4
				CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_Finger(ii).xyRadial)
				CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_Finger(ii).xyUlnar)
			Next ii
			'Curves
			CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, g_FingerThumbBlend)
			CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, WristThumbBlend)
			If Not g_OnFold Then
				CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, WristLFSBlend)
			End If
			CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, FingerLFSBlend)
			CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, RadialProfile)
			CADUTILS.PR_MirrorCurveInYaxis(xyDatum.X, nTranslate, UlnarProfile)
			'Points
			For ii = 1 To 6
				CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyPalm(ii))
			Next ii
			For ii = 1 To 5
				CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyPalmThumb(ii))
			Next ii
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyLFS)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyLittle)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyRing)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyMiddle)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyIndex)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, xyThumb)
			
			'Top of thumb curve
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_ThumbTopCurve.xyStart)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_ThumbTopCurve.xyTangent)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_ThumbTopCurve.xyEnd)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_ThumbTopCurve.xyR1)
			CADUTILS.PR_MirrorPointInYaxis(xyDatum.X, nTranslate, g_ThumbTopCurve.xyR2)
			
			'Text alignment
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nLength = ARMDIA1.max((Len(txtPatientName.Text) * 0.075) + 0.125, 1.5)
			ARMDIA1.PR_MakeXY(xyTextAlign, xyPalm(1).X - nLength, xyPalm(1).Y + 1.625)
		End If
		
		
		'Draw Fingers (For missing see later)
		For ii = 1 To 4
			PR_DrawFingers(g_Finger(ii))
		Next ii
		'For ii = 1 To 5
		'    PR_DrawMarkerNamed "detail", xyPalmThumb(ii), .1, .1, 0
		'    PR_DrawText Trim$(Str$(ii)), xyPalmThumb(ii), .075
		'Next ii
		'Draw Thumb
		If g_nFingerLen(5) <> 0 Then
			Select Case g_iThumbStyle
				Case 0
					PR_DrawNormalThumb(nInsert)
				Case 1
					'do nothing as this is an edged thumb which does not require
					'a length, drawn below
				Case 2
					PR_DrawSetInThumb(nInsert)
			End Select
		Else
			If g_iThumbStyle = 1 Then PR_DrawEdgedThumb(nInsert)
		End If
		
		'Draw arm extension
		'Create profile for LitteFingerSide and ThumbSide
		'Little Finger Side
		If Not g_OnFold Then
			'Join blending profiles to extension
			LFS.n = 0
			For ii = 1 To FingerLFSBlend.n
				LFS.n = LFS.n + 1
				LFS.X(LFS.n) = FingerLFSBlend.X(ii)
				LFS.Y(LFS.n) = FingerLFSBlend.Y(ii)
			Next ii
			For ii = 2 To WristLFSBlend.n - 1
				LFS.n = LFS.n + 1
				LFS.X(LFS.n) = WristLFSBlend.X(ii)
				LFS.Y(LFS.n) = WristLFSBlend.Y(ii)
			Next ii
			'Add wrist point, only if RadialProfile.n = 1
			If RadialProfile.n = 1 Then
				LFS.n = LFS.n + 1
				LFS.X(LFS.n) = WristLFSBlend.X(WristLFSBlend.n)
				LFS.Y(LFS.n) = WristLFSBlend.Y(WristLFSBlend.n)
			End If
			For ii = 2 To UlnarProfile.n
				LFS.n = LFS.n + 1
				LFS.X(LFS.n) = UlnarProfile.X(ii)
				LFS.Y(LFS.n) = UlnarProfile.Y(ii)
			Next ii
		Else
			'To reduce code in the Zipper drawing module we
			'artificially create a polyline of three points
			'We know that the macro language always creates a polyline
			'if there are only 3 points in a fitted curve.
			'This is a bug in Drafix 3.01, which we take care of here
			'however this may be fixed for a subsequent release so beware!
			
			LFS.n = 0
			For ii = 1 To FingerLFSBlend.n
				LFS.n = LFS.n + 1
				LFS.X(LFS.n) = FingerLFSBlend.X(ii)
				LFS.Y(LFS.n) = FingerLFSBlend.Y(ii)
			Next ii
			LFS.X(LFS.n) = xyLFS.X
			LFS.Y(LFS.n) = UlnarProfile.Y(UlnarProfile.n) + ((xyLFS.Y - UlnarProfile.Y(UlnarProfile.n)) / 2)
			
			LFS.n = LFS.n + 1
			LFS.X(LFS.n) = UlnarProfile.X(UlnarProfile.n)
			LFS.Y(LFS.n) = UlnarProfile.Y(UlnarProfile.n)
		End If
		
		'Draw
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		ARMDIA1.PR_DrawFitted(LFS)
        MANUTILS.PR_AddDBValueToLast("Zipper", "LFS")
        MANUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
        'Only add this data base field if this is an editable curve
        If g_ExtendTo <> GLOVE_NORMAL Then BDYUTILS.PR_AddDBValueToLast("curvetype", "LFS")
		
		'Thumb side
		'Join blending profiles to extension
		ThumbSide.n = 0
		For ii = 1 To g_FingerThumbBlend.n
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = g_FingerThumbBlend.X(ii)
			ThumbSide.Y(ThumbSide.n) = g_FingerThumbBlend.Y(ii)
		Next ii
		For ii = 2 To WristThumbBlend.n - 1
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(ii)
			ThumbSide.Y(ThumbSide.n) = WristThumbBlend.Y(ii)
		Next ii
		'Add wrist point, only if RadialProfile.n = 1
		If RadialProfile.n = 1 Then
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(WristThumbBlend.n)
			ThumbSide.Y(ThumbSide.n) = WristThumbBlend.Y(WristThumbBlend.n)
		End If
		For ii = 2 To RadialProfile.n
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = RadialProfile.X(ii)
			ThumbSide.Y(ThumbSide.n) = RadialProfile.Y(ii)
		Next ii
		'Draw
		ARMDIA1.PR_DrawFitted(ThumbSide)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "ETS")
		BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
		'Only add this data base field if this is an editable curve
		If g_ExtendTo <> GLOVE_NORMAL Then BDYUTILS.PR_AddDBValueToLast("curvetype", "ETS")
		
		'Add labels and construction lines
		ARMDIA1.PR_Setlayer("Construct")
		
		
		'Draw notes for each tape
		'If on the fold draw notes at a fixed distance
		If g_OnFold Then
			nLength = System.Math.Abs(RadialProfile.X(1) - UlnarProfile.X(1)) / 2
		Else
			nLength = -1
		End If
		For ii = 1 To RadialProfile.n
			ARMDIA1.PR_MakeXY(xyPt1, UlnarProfile.X(ii), UlnarProfile.Y(ii))
			ARMDIA1.PR_MakeXY(xyPt2, RadialProfile.X(ii), RadialProfile.Y(ii))
			If ii = RadialProfile.n - 1 And g_EOSType = ARM_FLAP Then
				'Provide a dummy EOS for the zipper programme
				PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "EOS")
			Else
				PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "")
			End If
			If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
				'Draw elbow line
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_Setlayer("Notes")
				ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.Y = xyPt3.Y + 0.0625
				ARMDIA1.PR_DrawText("E", xyPt3, 0.1)
			End If
		Next ii
		
		
		If g_EOSType = ARM_FLAP Then
			CADGLEXT.PR_DrawShoulderFlaps(xyPt1, xyPt2)
		Else
			'Draw EOS, account for elastic if last tape is at elbow
			'N.B. our elbow is fixed at tape no 9
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
			ARMDIA1.PR_Setlayer("Notes")
			If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
				ARMDIA1.PR_MakeXY(xyPt3, RadialProfile.X(RadialProfile.n - 1), RadialProfile.Y(RadialProfile.n - 1))
				aAngle = System.Math.Abs(90 - ARMDIA1.FN_CalcAngle(xyPt2, xyPt3))
				If Val(txtAge.Text) > 10 Then nOffSet = 0.75 Else nOffSet = 0.375
				nLength = nOffSet / System.Math.Cos(aAngle * (PI / 180))
				ARMDIA1.PR_CalcPolar(xyPt2, ARMDIA1.FN_CalcAngle(xyPt2, xyPt3), nLength, xyPt2)
				If Not g_OnFold Then
					ARMDIA1.PR_MakeXY(xyPt3, UlnarProfile.X(UlnarProfile.n - 1), UlnarProfile.Y(UlnarProfile.n - 1))
					ARMDIA1.PR_CalcPolar(xyPt1, ARMDIA1.FN_CalcAngle(xyPt1, xyPt3), nLength, xyPt1)
				Else
					xyPt1.Y = xyPt1.Y + nOffSet
				End If
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.Y = xyPt3.Y + 0.0625
				ARMDIA1.PR_DrawText("E", xyPt3, 0.1)
			End If
			
			'Elastic for children
			If Val(txtAge.Text) <= 10 Then
				ARMDIA1.PR_Setlayer("Notes")
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.Y = xyPt3.Y + 0.25
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_DrawText("1/2\"" Elastic", xyPt3, 0.1)
			End If
			
			'Draw EOS line
			ARMDIA1.PR_Setlayer("Template" & g_sSide)
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "EOS")
			BDYUTILS.PR_AddDBValueToLast("curvetype", "EOS")
			
		End If
		
		'Draw Centre line
		If Not g_OnFold Then
			ARMDIA1.PR_Setlayer("Construct")
			aAngle = ARMDIA1.FN_CalcAngle(xyPt1, xyPt2)
			nLength = ARMDIA1.FN_CalcLength(xyPt1, xyPt2)
			ARMDIA1.PR_CalcPolar(xyPt1, aAngle, nLength / 2, xyPt1)
			ARMDIA1.PR_MakeXY(xyPt1, xyPt1.X, UlnarProfile.Y(1))
			ARMDIA1.PR_MakeXY(xyPt2, xyPt1.X, UlnarProfile.Y(UlnarProfile.n))
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
		End If
		
		
		'Insert Size (Text)
		'Position Insert text Size w.r.t Little Finger web
		ARMDIA1.PR_Setlayer("Notes")
		If nInsert = 0.375 Then
			'Substitute 1/2" for 3/8" inch for printing only
			sText = Trim(ARMEDDIA1.fnInchestoText(0.5))
		Else
			sText = Trim(ARMEDDIA1.fnInchestoText(nInsert))
		End If
		If Mid(sText, 1, 1) = "-" Then sText = Mid(sText, 2) 'Strip leading "-" sign
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = sText & "\" & QQ
		sText = "INSERT " & sText
		ARMDIA1.PR_MakeXY(xyPt3, xyLittle.X, xyLittle.Y - (nInsert + 0.25))
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		'Inserts
		Select Case g_iInsertStyle
			Case 1 'One insert, Little Finger Only, Off the Fold
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyLittle, 1, 0)
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyRing, 1, 0)
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyMiddle, 1, 0)
				If g_sSide = "Left" Then
					ARMDIA1.PR_MakeXY(xyPt1, xyPalmThumb(1).X - 0.0625, xyPalmThumb(1).Y + 0.03125)
				Else
					ARMDIA1.PR_MakeXY(xyPt1, xyPalmThumb(1).X + 0.0625, xyPalmThumb(1).Y + 0.03125)
				End If
				ARMDIA1.PR_MakeXY(xyPt2, xyPt1.X, xyPalmThumb(1).Y + 0.25)
				ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
				BDYUTILS.PR_DrawMarkerNamed("Open Short Arrow", xyPt1, 0.1, 0.125, 90)
				BDYUTILS.PR_DrawMarkerNamed("Open Short Arrow", xyPt2, 0.1, 0.125, 270)
			Case 2 'One insert, Little and Index Finger, On or Off the Fold
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyLittle, 1, 0)
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyRing, 1, 0)
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyMiddle, 1, 0)
			Case 3 'One insert Glove, On or Off the Fold
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyLittle, 1, 0)
				ARMDIA1.PR_InsertSymbol("GloveInsertU", xyMiddle, 1, 0)
		End Select
		
		'Add wrist line on notes
		ARMDIA1.PR_DrawLine(xyPalm(1), xyPalm(6))
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "WRIST")
		BDYUTILS.PR_CalcMidPoint(xyPalm(1), xyPalm(6), xyPt3)
		xyPt3.Y = xyPt3.Y + 0.0625
		ARMDIA1.PR_DrawText("W", xyPt3, 0.1)
		
		'print fold line if drawn on the fold and print requested
		If g_PrintFold And g_OnFold Then
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt3 = xyPalm(1)
			xyPt3.Y = xyPalm(1).Y - 0.25
			If g_sSide = "Right" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_nCurrTextAngle = 90
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_nCurrTextAngle = -90
			End If
			ARMDIA1.PR_DrawText("Fold", xyPt3, 0.1)
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
		End If
		
		'Slanted inserts
		If chkSlantedInserts.CheckState = 1 Then
			nInsert = nInsert - 0.125
			For ii = 2 To 4
				If g_Missing(ii) <> True Then
					If ii = 2 Then
						TmpX = xyLittle.X
						TmpY = xyLittle.Y
					ElseIf ii = 3 Then 
						TmpX = xyRing.X
						TmpY = xyRing.Y
					Else
						TmpX = xyMiddle.X
						TmpY = xyMiddle.Y
					End If
					'Vertical line
					ARMDIA1.PR_MakeXY(xyPt1, TmpX, TmpY)
					ARMDIA1.PR_MakeXY(xyPt2, TmpX, TmpY - nInsert)
					ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
					
					'Horizontal line
					ARMDIA1.PR_MakeXY(xyPt1, TmpX + EIGHTH, TmpY - nInsert)
					ARMDIA1.PR_MakeXY(xyPt2, TmpX - EIGHTH, TmpY - nInsert)
					ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
					
				End If
			Next ii
		End If
		
		
		'Reinforced Palms
		If chkPalm.CheckState = 1 Or chkDorsal.CheckState = 1 Then
			Reinforced.n = 5
			Reinforced.X(1) = xyLFS.X
			Reinforced.Y(1) = xyLFS.Y - 0.25
			Reinforced.X(2) = xyLittle.X
			Reinforced.Y(2) = xyLittle.Y - 0.25
			Reinforced.X(3) = xyRing.X
			Reinforced.Y(3) = xyRing.Y - 0.25
			Reinforced.X(4) = xyMiddle.X
			Reinforced.Y(4) = xyMiddle.Y - 0.25
			Reinforced.X(5) = xyIndex.X
			Reinforced.Y(5) = xyMiddle.Y - 0.25
			
			'Draw reinforcing lines using Long Dash line type
			BDYUTILS.PR_SetLineStyle("Long Dash")
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
			
			'Establish a text insert point
			BDYUTILS.PR_CalcMidPoint(xyPalm(3), xyPalm(4), xyPt3)
			xyPt3.Y = xyPt3.Y - 0.75
			
			If chkPalm.CheckState = 1 Then
				ARMDIA1.PR_DrawFitted(Reinforced)
				'Add REINFORCED text
				ARMDIA1.PR_DrawText("REINFORCED\nPALM", xyPt3, 0.1)
			End If
			
			If chkDorsal.CheckState = 1 Then
				If chkSlantedInserts.CheckState = 1 Then
					For ii = 1 To 5
						Reinforced.Y(ii) = Reinforced.Y(ii) - nInsert
					Next ii
				End If
				ARMDIA1.PR_DrawFitted(Reinforced)
				'Add REINFORCED text
				If chkPalm.CheckState = 1 Then
					'Shift text so that it misses PALM Text
					If txtSide.Text = "Right" Then
						xyPt3.Y = xyPt3.Y - 0.4
					Else
						xyPt3.Y = xyPt3.Y + 0.4
					End If
				End If
				ARMDIA1.PR_DrawText("REINFORCED\nDORSAL", xyPt3, 0.1)
			End If
			
			'Close at wrist
			ARMDIA1.PR_DrawLine(xyPalm(1), xyPalm(6))
			BDYUTILS.PR_SetLineStyle("bylayer")
		End If
		
		
		'Missing Fingers
		PR_DrawMissingFingers()
		
		
		'Patient Details
		ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPt3 = xyTextAlign
		If txtWorkOrder.Text = "" Then
			sWorkOrder = "-"
		Else
			sWorkOrder = txtWorkOrder.Text
		End If
		sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim(Mid(cboFabric.Text, 4))
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		'Other patient details in black on layer construct
		ARMDIA1.PR_Setlayer("Construct")
		ARMDIA1.PR_MakeXY(xyPt3, xyPt3.X, xyPt3.Y - 0.8)
		sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		
		'Markers for zippers
		'Place markers at lowest Web positions
		'For latter use with Zips, if a slant insert given then
		'supply the length of the slant insert
		'Added here capture mods to nInsert and drawn on layer construct
		
		BDYUTILS.PR_DrawMarker(xyPalm(1))
		BDYUTILS.PR_AddDBValueToLast("Zipper", "PALMER")
		ARMDIA1.PR_NamedHandle("hPalmer") 'Save a HANDLE for later use
		
		BDYUTILS.PR_DrawMarker(xyLittle)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "PALMER-WEB")
		If chkSlantedInserts.CheckState = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		BDYUTILS.PR_DrawMarker(xyRing)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "DORSAL-WEB")
		If chkSlantedInserts.CheckState = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		BDYUTILS.PR_DrawMarker(xyPalmThumb(3))
		BDYUTILS.PR_AddDBValueToLast("Zipper", "DORSAL")
		
		BDYUTILS.PR_DrawMarker(xyMiddle)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "OUTSIDE-WEB")
		If chkSlantedInserts.CheckState = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		'Markers for editors
		For ii = 2 To 6
			BDYUTILS.PR_DrawMarker(xyPalm(ii))
			BDYUTILS.PR_AddDBValueToLast("Data", "PALM" & Trim(Str(ii)))
		Next ii
		
		
		'Update symbol gloveglove, glovecommon and marker palmer with data
		PR_UpdateDBFields("Draw")
		
		ARMDIA1.PR_Setlayer("1")
		
		
debugclose: 
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CreateSaveMacro(ByRef sDrafixFile As String)
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		
		'Initialise String globals
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(10) 'The new line character
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
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient " & txtPatientName.Text & ", " & txtFileNo.Text & ", Hand - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, GLOVES - Save only")
		
		'Define DRAFIX variables
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hLayer, hChan, hSym, hEnt, hOrigin, hMPD;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY     xyO, xyScale;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sFileNo, sSide, sID, sName;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "ANGLE  aAngle;")
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		
		'Display Hour Glass symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Saving" & QQ & ");")
		
		'Get Handle and Origin of "mainpatientdetails symbol"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hMPD)")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "else")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")
		
		'Save DB fileds
		PR_UpdateDBFields("Save")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CutBackTip(ByRef iFinger As Short, ByRef iOption As Object)
		'Function to cut down code w.r.t.
		' optLittleTip_Click
		' optMiddleTip_Click
		' optRingTip_Click
		' optIndexTip_Click
		' optThumbTip_Click
		
		Dim nLen As Double
		Dim nTipCutBack As Double
		
		If txtLen(iFinger).Enabled = False Then Exit Sub
		
		nLen = CADGLOVE1.FN_InchesValue(txtLen(iFinger))
		If nLen <= 0 Then Exit Sub
		
		If Val(txtAge.Text) <= AGE_CUTOFF Then
			nTipCutBack = CHILD_STD_OPEN_TIP
		Else
			nTipCutBack = ADULT_STD_OPEN_TIP
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object iOption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If iOption = 1 Then nLen = nLen - nTipCutBack
		
		If nLen > 0 Then lblLen(iFinger).Text = ARMEDDIA1.fnInchestoText(nLen)
		
		
	End Sub
	
	Private Sub PR_DisableCir(ByRef Index As Short)
		'Read as part of labCir_DblClick
		labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		txtCir(Index).Enabled = False
		lblCir(Index).Enabled = False
	End Sub
	
	Private Sub PR_DrawEdgedThumb(ByRef nInsert As Object)
		'Procedure to draw an edged thumb
		Static aAngle As Short
		Static nFillet As Single
		Static xyTmp3, xyTmp1, xyTmp2, xyKeep As XY
		
		'Set angle depending on side
		If g_sSide = "Right" Then aAngle = 360 Else aAngle = 180
		
		nFillet = 0.125
		
		'Thumb on layer Notes
		ARMDIA1.PR_Setlayer("Notes")
		
		ARMDIA1.PR_CalcPolar(xyPalmThumb(1), aAngle, 0.25 - nFillet, xyTmp1)
		ARMDIA1.PR_DrawLine(xyPalmThumb(1), xyTmp1)
		ARMDIA1.PR_CalcPolar(xyTmp1, 270, nFillet, xyTmp2)
		ARMDIA1.PR_CalcPolar(xyTmp2, aAngle, nFillet, xyTmp3)
		ARMDIA1.PR_DrawArc(xyTmp2, xyTmp1, xyTmp3)
		'UPGRADE_WARNING: Couldn't resolve default property of object xyKeep. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyKeep = xyTmp3
		
		ARMDIA1.PR_CalcPolar(xyPalmThumb(4), aAngle, 0.25 - nFillet, xyTmp1)
		ARMDIA1.PR_DrawLine(xyPalmThumb(4), xyTmp1)
		ARMDIA1.PR_CalcPolar(xyTmp1, 90, nFillet, xyTmp2)
		ARMDIA1.PR_CalcPolar(xyTmp2, aAngle, nFillet, xyTmp3)
		If g_sSide = "Right" Then ARMDIA1.PR_DrawArc(xyTmp2, xyTmp3, xyTmp1) Else ARMDIA1.PR_DrawArc(xyTmp2, xyTmp1, xyTmp3)
		ARMDIA1.PR_DrawLine(xyKeep, xyTmp3)
		
		BDYUTILS.PR_CalcMidPoint(xyKeep, xyTmp3, xyTmp1)
		ARMDIA1.PR_CalcPolar(xyTmp1, aAngle, 0.05, xyTmp1)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = aAngle - 90
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
		ARMDIA1.PR_DrawText("EDGED", xyTmp1, 0.1)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		
		
		
	End Sub
	
	Private Sub PR_DrawFingers(ByRef Digit As Finger)
		'Procedure to draw a Finger based on the data given in the structure
		'Finger
		'NB
		'    NOT USED TO DRAW A THUMB
		'
		' Fingers are drawn vertically
		' All dimensions assummed to be figured
		'
		'
		Static xyPt3, xyPt1, xyPt2, xyPt4 As XY
		Static xyPt7, xyPt5, xyPt6, xyPt8 As XY
		Static xyPt9, xyMid As XY
		Static nX, nY As Double
		Static nA, aTheta, nRadius As Double
		Static rTheta As Double
		Static nHt, nWidth As Double
		Static nTipWidth, nBulge As Double
		Static Sign As Short
		
		'Simple case for missing fingers
		'Don't draw anything
		If Digit.Len_Renamed <= 0 Then Exit Sub
		
		'Establish if handing is required
		Sign = 1
		If Digit.xyUlnar.X > Digit.xyRadial.X Then Sign = -1
		
		'Case for missing DIP
		If Digit.DIP <= 0 Then
			ARMDIA1.PR_MakeXY(xyPt1, Digit.xyRadial.X, Digit.WebHt + Digit.Len_Renamed)
			If Sign > 0 Then
				ARMDIA1.PR_CalcPolar(xyPt1, 180, Digit.PIP, xyPt2)
			Else
				ARMDIA1.PR_CalcPolar(xyPt1, 0, Digit.PIP, xyPt2)
			End If
			Select Case Digit.Tip
				Case 0, 10 'Closed tip (ON and OFF Fold)
					nHt = ARMDIA1.FN_CalcLength(xyPt1, xyPt2) / 2 'Bug fix 04/02/97
					If xyPt1.Y - nHt <= Digit.WebHt Then '     "    "
						xyPt1.Y = Digit.WebHt '     "    "
						xyPt2.Y = xyPt1.Y '     "    "
					Else '     "    "
						xyPt1.Y = xyPt1.Y - nHt '     "    "
						xyPt2.Y = xyPt1.Y '     "    "
					End If '     "    "
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_AddVertex(xyPt1, (Sign))
					BDYUTILS.PR_AddVertex(xyPt2, 0)
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_EndPoly()
					
					
				Case 1, 11, 2, 12 'Standard open tip & Desired length open tip (ON Fold)
					'The cutback for standard open tips has been done in
					'FN_CalculateGlove
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_AddVertex(xyPt1, 0)
					BDYUTILS.PR_AddVertex(xyPt2, 0)
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_EndPoly()
					
			End Select
			
		Else
			
			
			Select Case Digit.Tip
				Case 0 'Closed tip (OFF Fold)
					nHt = Digit.Len_Renamed / 3
					nTipWidth = Digit.DIP * 0.8
					If nTipWidth < 0.375 Then nTipWidth = 0.375
					
					'PIP
					ARMDIA1.PR_MakeXY(xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt)
					
					'DIP
					xyPt2.Y = xyPt1.Y + nHt
					xyPt2.X = xyPt1.X + (Sign * (System.Math.Abs(Digit.PIP - Digit.DIP) / 2))
					
					'At Tip
					xyPt3.Y = xyPt2.Y + nHt
					xyPt3.X = xyPt2.X + (Sign * (System.Math.Abs(nTipWidth - Digit.DIP) / 2))
					
					'Find center of arc
					rTheta = System.Math.Atan((System.Math.Abs(nTipWidth - Digit.DIP) / 2) / nHt)
					
					aTheta = rTheta * (180 / PI)
					rTheta = (rTheta + (PI / 2)) / 2
					nA = (nTipWidth / 2) / System.Math.Cos(rTheta)
					nRadius = nA * System.Math.Sin(rTheta)
					
					If Sign > 0 Then
						ARMDIA1.PR_CalcPolar(xyPt3, 315 - (aTheta / 2), nA, xyPt4)
						ARMDIA1.PR_CalcPolar(xyPt4, 180 - aTheta, nRadius, xyPt5)
						ARMDIA1.PR_CalcPolar(xyPt4, aTheta, nRadius, xyPt6)
					Else
						ARMDIA1.PR_CalcPolar(xyPt3, 225 + (aTheta / 2), nA, xyPt4)
						ARMDIA1.PR_CalcPolar(xyPt4, aTheta, nRadius, xyPt5)
						ARMDIA1.PR_CalcPolar(xyPt4, 180 - aTheta, nRadius, xyPt6)
					End If
					
					ARMDIA1.PR_MakeXY(xyPt7, xyPt3.X + (Sign * nTipWidth), xyPt3.Y)
					ARMDIA1.PR_MakeXY(xyPt8, xyPt2.X + (Sign * Digit.DIP), xyPt2.Y)
					ARMDIA1.PR_MakeXY(xyPt9, xyPt1.X + (Sign * Digit.PIP), xyPt1.Y)
					
					'Draw Finger
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_AddVertex(xyPt1, 0)
					BDYUTILS.PR_AddVertex(xyPt2, 0)
					'Calculate bulge
					nBulge = ((nRadius - System.Math.Abs(xyPt5.Y - xyPt4.Y)) / (ARMDIA1.FN_CalcLength(xyPt5, xyPt6) / 2)) * (Sign * -1)
					BDYUTILS.PR_AddVertex(xyPt5, nBulge)
					BDYUTILS.PR_AddVertex(xyPt6, 0)
					BDYUTILS.PR_AddVertex(xyPt8, 0)
					BDYUTILS.PR_AddVertex(xyPt9, 0)
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_EndPoly()
					
				Case 1, 2 'Standard open tip & Desired length open tip (OFF Fold
					'The cutback for standard open tips has been done in
					'FN_CalculateGlove
					nHt = Digit.Len_Renamed / 2
					
					'PIP
					' xyPt1.x = Digit.xyUlnar.x
					' xyPt1.y = Digit.xyUlnar.y + nHt
					ARMDIA1.PR_MakeXY(xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object xyPt4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyPt4 = xyPt1
					xyPt4.X = xyPt1.X + (Sign * Digit.PIP)
					
					'DIP
					xyPt2.Y = xyPt1.Y + nHt
					xyPt2.X = xyPt1.X + (Sign * (System.Math.Abs(Digit.PIP - Digit.DIP) / 2))
					
					'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyPt3 = xyPt2
					xyPt3.X = xyPt2.X + (Sign * Digit.DIP)
					
					'Draw Finger
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_AddVertex(xyPt1, 0)
					BDYUTILS.PR_AddVertex(xyPt2, 0)
					BDYUTILS.PR_AddVertex(xyPt3, 0)
					BDYUTILS.PR_AddVertex(xyPt4, 0)
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_EndPoly()
					
				Case 10 'Closed tip (ON Fold)
					nHt = Digit.Len_Renamed / 3
					nTipWidth = Digit.DIP * 0.8
					If nTipWidth < 0.375 Then nTipWidth = 0.375
					
					'PIP
					ARMDIA1.PR_MakeXY(xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt)
					
					'DIP
					xyPt2.Y = xyPt1.Y + nHt
					xyPt2.X = Digit.xyUlnar.X
					
					'At Tip
					xyPt3.Y = xyPt2.Y + nHt
					xyPt3.X = Digit.xyUlnar.X
					
					'Center of arc
					nRadius = nTipWidth / 2
					ARMDIA1.PR_MakeXY(xyPt4, xyPt3.X + (Sign * nRadius), xyPt3.Y - nRadius)
					
					'Find Ends of arc
					rTheta = System.Math.Atan(System.Math.Abs(nTipWidth - Digit.DIP) / nHt)
					aTheta = rTheta * (180 / PI)
					
					ARMDIA1.PR_MakeXY(xyPt5, xyPt3.X, xyPt3.Y - nRadius)
					
					If Sign > 0 Then
						ARMDIA1.PR_CalcPolar(xyPt4, aTheta, nRadius, xyPt6)
					Else
						ARMDIA1.PR_CalcPolar(xyPt4, 180 - aTheta, nRadius, xyPt6)
					End If
					
					ARMDIA1.PR_MakeXY(xyPt7, xyPt3.X + (Sign * nTipWidth), xyPt3.Y)
					ARMDIA1.PR_MakeXY(xyPt8, xyPt2.X + (Sign * Digit.DIP), xyPt2.Y)
					ARMDIA1.PR_MakeXY(xyPt9, xyPt1.X + (Sign * Digit.PIP), xyPt1.Y)
					
					'Draw finger
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_AddVertex(xyPt5, 0)
					'Calculate bulge
					BDYUTILS.PR_CalcMidPoint(xyPt5, xyPt6, xyMid)
					nBulge = ((nRadius - ARMDIA1.FN_CalcLength(xyMid, xyPt4)) / (ARMDIA1.FN_CalcLength(xyPt5, xyPt6) / 2)) * (Sign * -1)
					BDYUTILS.PR_AddVertex(xyPt5, nBulge)
					BDYUTILS.PR_AddVertex(xyPt6, 0)
					BDYUTILS.PR_AddVertex(xyPt8, 0)
					BDYUTILS.PR_AddVertex(xyPt9, 0)
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_EndPoly()
					
					
				Case 11, 12 'Standard open tip & Desired length open tip (ON Fold)
					'The cutback for standard open tips has been done in
					'FN_CalculateGlove
					nHt = Digit.Len_Renamed / 2
					
					'PIP
					ARMDIA1.PR_MakeXY(xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object xyPt4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyPt4 = xyPt1
					xyPt4.X = xyPt1.X + (Sign * Digit.PIP)
					
					'DIP
					xyPt2.Y = xyPt1.Y + nHt
					xyPt2.X = xyPt1.X
					
					'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyPt3 = xyPt2
					xyPt3.X = xyPt2.X + (Sign * Digit.DIP)
					
					'Draw Finger
					BDYUTILS.PR_StartPoly()
					BDYUTILS.PR_AddVertex(Digit.xyUlnar, 0)
					BDYUTILS.PR_AddVertex(xyPt2, 0)
					BDYUTILS.PR_AddVertex(xyPt3, 0)
					BDYUTILS.PR_AddVertex(xyPt4, 0)
					BDYUTILS.PR_AddVertex(Digit.xyRadial, 0)
					BDYUTILS.PR_EndPoly()
					
			End Select
		End If
		
	End Sub
	
	Private Sub PR_DrawMissing(ByRef iFingers As Short, ByRef iTip As Short, ByRef xy1 As XY, ByRef xy2 As XY, ByRef xy3 As XY, ByRef xy4 As XY)
		
		Dim xyMidPoint As XY
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		Select Case iFingers
			Case 1
				BDYUTILS.PR_CalcMidPoint(xy1, xy2, xyMidPoint)
				BDYUTILS.PR_StartPoly()
				BDYUTILS.PR_AddVertex(xy1, 0)
				BDYUTILS.PR_AddVertex(xyMidPoint, 0)
				BDYUTILS.PR_AddVertex(xy2, 0)
				BDYUTILS.PR_EndPoly()
			Case 2
				BDYUTILS.PR_StartPoly()
				BDYUTILS.PR_AddVertex(xy1, 0)
				BDYUTILS.PR_AddVertex(xy2, 0)
				BDYUTILS.PR_AddVertex(xy3, 0)
				BDYUTILS.PR_EndPoly()
				'UPGRADE_WARNING: Couldn't resolve default property of object xyMidPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyMidPoint = xy2
			Case 3
				BDYUTILS.PR_StartPoly()
				BDYUTILS.PR_AddVertex(xy1, 0)
				BDYUTILS.PR_AddVertex(xy2, 0)
				BDYUTILS.PR_AddVertex(xy3, 0)
				BDYUTILS.PR_AddVertex(xy4, 0)
				BDYUTILS.PR_EndPoly()
				BDYUTILS.PR_CalcMidPoint(xy2, xy3, xyMidPoint)
			Case Is < 1
				'Add note only
				BDYUTILS.PR_CalcMidPoint(xy1, xy2, xyMidPoint)
				xyMidPoint.Y = xyMidPoint.Y - (EIGHTH * 1)
			Case Else
				Exit Sub
		End Select
		
		ARMDIA1.PR_Setlayer("Notes")
		xyMidPoint.Y = xyMidPoint.Y - (EIGHTH * 2)
		If iTip = 0 Or iTip = 10 Then
			ARMDIA1.PR_DrawText("CLOSED", xyMidPoint, 0.1)
		Else
			ARMDIA1.PR_DrawText("OPEN", xyMidPoint, 0.1)
		End If
		
	End Sub
	
	Private Sub PR_DrawMissingFingers()
		Dim ClosedLittle, iMissingFingers, ClosedIndex As Short
		Dim ii As Short
		Dim xyNull As XY
		
		If g_FusedFingers Then Exit Sub
		
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
		
		'Check for missing fingers
		iMissingFingers = 0
		For ii = 0 To 3
			If g_Finger(ii + 1).Len_Renamed <= 0 Then iMissingFingers = iMissingFingers + (2 ^ ii)
		Next ii
		
		ClosedLittle = False
		ClosedIndex = False
		If g_Finger(1).Tip = 0 Or g_Finger(1).Tip = 10 Then ClosedLittle = True
		If g_Finger(4).Tip = 0 Or g_Finger(4).Tip = 10 Then ClosedIndex = True
		
		Select Case iMissingFingers
			'NB -1 => text only
			Case 1
				If ClosedLittle Then
					PR_DrawMissing(-1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				End If
			Case 2
				PR_DrawMissing(1, g_Finger(2).Tip, xyLittle, xyRing, xyNull, xyNull)
			Case 3
				If ClosedLittle Then
					PR_DrawMissing(1, g_Finger(1).Tip, xyLittle, xyRing, xyNull, xyNull)
				Else
					PR_DrawMissing(2, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyNull)
				End If
			Case 4
				PR_DrawMissing(1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull)
			Case 5
				If ClosedLittle Then
					PR_DrawMissing(-1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				End If
				PR_DrawMissing(1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull)
			Case 6
				PR_DrawMissing(2, g_Finger(2).Tip, xyLittle, xyRing, xyMiddle, xyNull)
			Case 7
				If ClosedLittle Then
					PR_DrawMissing(2, g_Finger(1).Tip, xyLittle, xyRing, xyMiddle, xyNull)
				Else
					PR_DrawMissing(3, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyMiddle)
				End If
			Case 8
				If ClosedIndex Then
					PR_DrawMissing(-1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				End If
			Case 9
				If ClosedLittle Then
					PR_DrawMissing(-1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				End If
				If ClosedIndex Then
					PR_DrawMissing(-1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				End If
			Case 10
				PR_DrawMissing(1, g_Finger(2).Tip, xyLittle, xyRing, xyNull, xyNull)
				If ClosedIndex Then
					PR_DrawMissing(-1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				End If
			Case 11
				If ClosedLittle Then
					PR_DrawMissing(1, g_Finger(1).Tip, xyLittle, xyRing, xyNull, xyNull)
				Else
					PR_DrawMissing(2, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyNull)
				End If
				If ClosedIndex Then
					PR_DrawMissing(-1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull)
				End If
			Case 12
				If ClosedIndex Then
					PR_DrawMissing(1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull)
				Else
					PR_DrawMissing(2, g_Finger(3).Tip, xyRing, xyMiddle, xyIndex, xyNull)
				End If
			Case 13
				If ClosedLittle Then
					PR_DrawMissing(-1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				Else
					PR_DrawMissing(1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
				End If
				If ClosedIndex Then
					PR_DrawMissing(1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull)
				Else
					PR_DrawMissing(2, g_Finger(3).Tip, xyRing, xyMiddle, xyIndex, xyNull)
				End If
			Case 14
				If ClosedIndex Then
					PR_DrawMissing(2, g_Finger(2).Tip, xyLittle, xyRing, xyMiddle, xyNull)
				Else
					PR_DrawMissing(3, g_Finger(3).Tip, xyLittle, xyRing, xyMiddle, xyIndex)
				End If
			Case 15
				If ClosedIndex And ClosedLittle Then
					PR_DrawMissing(1, g_Finger(1).Tip, xyLittle, xyMiddle, xyNull, xyNull)
				ElseIf ClosedLittle Then 
					PR_DrawMissing(-1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull)
					PR_DrawMissing(3, g_Finger(1).Tip, xyLittle, xyRing, xyMiddle, xyIndex)
				ElseIf ClosedIndex Then 
					PR_DrawMissing(-1, g_Finger(1).Tip, xyMiddle, xyIndex, xyNull, xyNull)
					PR_DrawMissing(3, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyMiddle)
				Else
					PR_DrawMissing(1, g_Finger(1).Tip, xyLFS, xyIndex, xyNull, xyNull)
				End If
			Case Else
				'=> No missing fingers so "DO NOTHING"
		End Select
		
	End Sub
	
	Private Sub PR_DrawNormalThumb(ByRef nInsert As Double)
		'UPGRADE_WARNING: Lower bound of array xyT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim xyT(12)
		Static xyTmpStart, xyTmp, xyTmpEnd, xyTranslate As XY
		Static xyThumbPalm5 As XY
		Static aAngle, nLength, aThumbAngle As Double
		Static nJ_Radius, nSeam, nInc As Double
		Static nSlantInsert, nThumbSlantLength As Double
		Static aPalmThumbStart, aPalmThumbEnd As Double
		Static ThumbTopCurve, ThumbBottCurve As BiArc
		Static iThumbDirection, ii, iAngleDirection, Success As Short
		Static nFiguredPalm, nMinDist, nOffSet, nBulge As Double
		Static sText As String
		
		'Globals
		'     g_ThumbTopCurve, from FN_CalculateGlove
		'
		
		'Angles
		aPalmThumbStart = ARMDIA1.FN_CalcAngle(xyPalmThumb(4), xyPalmThumb(3)) + 20
		aPalmThumbEnd = 0
		iAngleDirection = 1
		iThumbDirection = -1
		nOffSet = 3
		
		'Mirror Thumb angles for right side
		If g_sSide = "Right" Then
			'Angles
			aPalmThumbStart = aPalmThumbStart - 40
			aPalmThumbEnd = 180
			iAngleDirection = -1
			iThumbDirection = 1
			nOffSet = 6
		End If
		
		nInc = 0.03125 '1/32nd of an inch
		
		'Thumb curve on palm
		ARMDIA1.PR_Setlayer("Notes")
		
		'BOTTOM Curve
		'We find by iteration the first curve that fits the data.  We change the
		'start angle as part of the iteration
		Success = False
		While Not Success And (aPalmThumbStart < 180 Or aPalmThumbStart > 0)
			ii = BDYUTILS.FN_BiArcCurve(xyPalmThumb(4), aPalmThumbStart, xyPalmThumb(3), 90, ThumbBottCurve)
			If ii Then
				Success = True
			Else
				'Adjust and try again
				aPalmThumbStart = aPalmThumbStart + (iAngleDirection * 1)
			End If
		End While
		
		If Not Success Then
			MsgBox("Can't form BOTTOM part of thumb curve on palm! Adjust the palm curve and use the Transfer tool to update the thumb", 48, g_sDialogID)
			'Degenerate "ThumbBottCurve" to a straight line
			'UPGRADE_WARNING: Couldn't resolve default property of object ThumbBottCurve.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ThumbBottCurve.xyStart = xyPalmThumb(4)
			'UPGRADE_WARNING: Couldn't resolve default property of object ThumbBottCurve.xyEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ThumbBottCurve.xyEnd = xyPalmThumb(3)
			ThumbBottCurve.nR1 = 0
			ThumbBottCurve.nR2 = 0
		End If
		
		'Draw thumb curve on palm
		PR_DrawThumbCurve(ThumbBottCurve, g_ThumbTopCurve)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "ThumbPalmCurve")
		
		'    PR_DrawLine xyPalmThumb(5), xyPalmThumb(1)
		ARMDIA1.PR_DrawLine(g_ThumbTopCurve.xyEnd, xyPalmThumb(1))
		
		'Draw the actual thumb
		'Translate thumb curves
		nSeam = 0.125
		nJ_Radius = 0.125
		
		ThumbBottCurve.xyR1.X = ThumbBottCurve.xyR1.X + nOffSet
		ThumbBottCurve.xyR2.X = ThumbBottCurve.xyR2.X + nOffSet
		ThumbBottCurve.xyTangent.X = ThumbBottCurve.xyTangent.X + nOffSet
		ThumbBottCurve.xyEnd.X = ThumbBottCurve.xyEnd.X + nOffSet
		ThumbBottCurve.xyStart.X = ThumbBottCurve.xyStart.X + nOffSet
		g_ThumbTopCurve.xyR1.X = g_ThumbTopCurve.xyR1.X + nOffSet
		g_ThumbTopCurve.xyR2.X = g_ThumbTopCurve.xyR2.X + nOffSet
		g_ThumbTopCurve.xyTangent.X = g_ThumbTopCurve.xyTangent.X + nOffSet
		g_ThumbTopCurve.xyEnd.X = g_ThumbTopCurve.xyEnd.X + nOffSet
		g_ThumbTopCurve.xyStart.X = g_ThumbTopCurve.xyStart.X + nOffSet
		
		'Draw thumb curve (Seam line)
		PR_DrawThumbCurve(ThumbBottCurve, g_ThumbTopCurve)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "ThumbPieceSeam")
		
		aAngle = ARMDIA1.FN_CalcAngle(xyPalmThumb(4), xyPalmThumb(2))
		aThumbAngle = aAngle + (iThumbDirection * 45)
		'UPGRADE_WARNING: Couldn't resolve default property of object xyT(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyT(1) = ThumbBottCurve.xyStart
		
		nLength = g_Finger(5).PIP / (1 / System.Math.Sqrt(2))
		ARMDIA1.PR_CalcPolar(xyT(1), aAngle, nLength, xyT(4))
		ARMDIA1.PR_CalcPolar(xyT(4), aThumbAngle, g_Finger(5).Len_Renamed, xyT(3))
		ARMDIA1.PR_CalcPolar(xyT(3), aThumbAngle + (iThumbDirection * 90), g_Finger(5).PIP, xyT(2))
		
		nLength = nJ_Radius / System.Math.Tan((45 / 2) * (PI / 180))
		ARMDIA1.PR_CalcPolar(xyT(4), aThumbAngle, nLength, xyT(5))
		ARMDIA1.PR_CalcPolar(xyT(4), aAngle, nLength, xyT(6))
		ARMDIA1.PR_CalcPolar(xyT(6), aAngle + (iThumbDirection * 90), nJ_Radius, xyT(7))
		
		nLength = ARMDIA1.FN_CalcLength(xyPalmThumb(4), xyPalmThumb(2))
		ARMDIA1.PR_CalcPolar(xyT(1), aAngle, nLength, xyT(10))
		ARMDIA1.PR_DrawLine(g_ThumbTopCurve.xyEnd, xyT(10))
		
		nLength = nSeam / System.Math.Abs(System.Math.Sin(aAngle * (PI / 180)))
		nLength = ARMDIA1.FN_CalcLength(xyPalmThumb(4), xyPalmThumb(2)) + nLength
		ARMDIA1.PR_CalcPolar(xyT(1), aAngle, nLength, xyT(8))
		
		ARMDIA1.PR_CalcPolar(xyT(1), aThumbAngle + 180, 1, xyTmp)
		ii = BDYUTILS.FN_CirLinInt(xyT(1), xyTmp, ThumbBottCurve.xyR1, ThumbBottCurve.nR1 + nSeam, xyT(9))
		
		ARMDIA1.PR_DrawLine(xyT(6), xyT(1))
		
		'Offset line
		ARMDIA1.PR_CalcPolar(xyT(6), aAngle + (iAngleDirection * 90), 0.25, xyTmp)
		ARMDIA1.PR_CalcPolar(xyTmp, aAngle, 10, xyT(11))
		If Not BDYUTILS.FN_CirLinInt(xyT(11), xyTmp, g_ThumbTopCurve.xyR2, g_ThumbTopCurve.nR2 + nSeam, xyT(11)) Then
			'make a best guess
			ARMDIA1.PR_CalcPolar(xyT(8), aAngle + (iAngleDirection * 90), 0.25, xyT(11))
		End If
		
		ARMDIA1.PR_CalcPolar(xyTmp, aAngle + 180, 10, xyT(12))
		If Not BDYUTILS.FN_CirLinInt(xyT(12), xyTmp, ThumbBottCurve.xyR1, ThumbBottCurve.nR1 + nSeam, xyT(12)) Then
			'make a best guess
			ARMDIA1.PR_CalcPolar(xyT(1), aAngle + (iAngleDirection * 90), 0.25, xyT(12))
		End If
		
		ARMDIA1.PR_DrawLine(xyT(11), xyT(12))
		
		'Draw Text
		BDYUTILS.PR_CalcMidPoint(xyT(1), xyT(2), xyTmp)
		
		ARMDIA1.PR_CalcPolar(xyTmp, aThumbAngle + (iAngleDirection * 90), g_Finger(5).PIP * 2 / 3, xyTmp)
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, EIGHTH, CURRENT, CURRENT)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If g_sSide = "Left" Then
			g_nCurrTextAngle = aThumbAngle
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 180 + aThumbAngle
		End If
		sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & txtWorkOrder.Text
		ARMDIA1.PR_DrawText(sText, xyTmp, 0.1)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		'Draw Outside profile
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyT(9), 0)
		
		'Draw tip
		If g_Finger(5).Tip = 0 Then
			'Closed tip
			ARMDIA1.PR_CalcPolar(xyT(2), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(2))
			ARMDIA1.PR_CalcPolar(xyT(3), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(3))
			BDYUTILS.PR_CalcMidPoint(xyT(2), xyT(3), xyTmp)
			nBulge = 1 * (iThumbDirection * -1)
		Else
			'Open tip
			nBulge = 0
		End If
		
		BDYUTILS.PR_AddVertex(xyT(2), nBulge)
		BDYUTILS.PR_AddVertex(xyT(3), 0)
		
		'Calculate bulge
		BDYUTILS.PR_CalcMidPoint(xyT(5), xyT(6), xyTmp)
		nBulge = ((nJ_Radius - ARMDIA1.FN_CalcLength(xyTmp, xyT(7))) / (ARMDIA1.FN_CalcLength(xyT(5), xyT(6)) / 2)) * iThumbDirection
		BDYUTILS.PR_AddVertex(xyT(5), nBulge)
		
		BDYUTILS.PR_AddVertex(xyT(6), 0)
		BDYUTILS.PR_AddVertex(xyT(8), 0)
		
		BDYUTILS.PR_EndPoly()
		
		If ThumbBottCurve.nR1 > 0 And ThumbBottCurve.nR1 > 0 Then
			'Offset curves
			ThumbBottCurve.nR1 = ThumbBottCurve.nR1 + nSeam
			ThumbBottCurve.nR2 = ThumbBottCurve.nR2 + nSeam
			aAngle = ARMDIA1.FN_CalcAngle(ThumbBottCurve.xyR1, ThumbBottCurve.xyStart)
			ARMDIA1.PR_CalcPolar(ThumbBottCurve.xyR1, aAngle, ThumbBottCurve.nR1, ThumbBottCurve.xyStart)
			aAngle = ARMDIA1.FN_CalcAngle(ThumbBottCurve.xyR1, ThumbBottCurve.xyTangent)
			ARMDIA1.PR_CalcPolar(ThumbBottCurve.xyR1, aAngle, ThumbBottCurve.nR1, ThumbBottCurve.xyTangent)
			aAngle = ARMDIA1.FN_CalcAngle(ThumbBottCurve.xyR2, ThumbBottCurve.xyEnd)
			ARMDIA1.PR_CalcPolar(ThumbBottCurve.xyR2, aAngle, ThumbBottCurve.nR2, ThumbBottCurve.xyEnd)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object ThumbBottCurve.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ThumbBottCurve.xyStart = xyT(9)
			ThumbBottCurve.xyEnd.X = ThumbBottCurve.xyEnd.X + (iThumbDirection * nSeam)
		End If
		
		If g_ThumbTopCurve.nR1 > 0 And g_ThumbTopCurve.nR1 > 0 Then
			g_ThumbTopCurve.nR1 = g_ThumbTopCurve.nR1 + nSeam
			g_ThumbTopCurve.nR2 = g_ThumbTopCurve.nR2 + nSeam
			aAngle = ARMDIA1.FN_CalcAngle(g_ThumbTopCurve.xyR1, g_ThumbTopCurve.xyStart)
			ARMDIA1.PR_CalcPolar(g_ThumbTopCurve.xyR1, aAngle, g_ThumbTopCurve.nR1, g_ThumbTopCurve.xyStart)
			aAngle = ARMDIA1.FN_CalcAngle(g_ThumbTopCurve.xyR1, g_ThumbTopCurve.xyTangent)
			ARMDIA1.PR_CalcPolar(g_ThumbTopCurve.xyR1, aAngle, g_ThumbTopCurve.nR1, g_ThumbTopCurve.xyTangent)
			aAngle = ARMDIA1.FN_CalcAngle(g_ThumbTopCurve.xyR2, g_ThumbTopCurve.xyEnd)
			ARMDIA1.PR_CalcPolar(g_ThumbTopCurve.xyR2, aAngle, g_ThumbTopCurve.nR2, g_ThumbTopCurve.xyEnd)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object g_ThumbTopCurve.xyEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_ThumbTopCurve.xyEnd = xyT(8)
			'UPGRADE_WARNING: Couldn't resolve default property of object g_ThumbTopCurve.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_ThumbTopCurve.xyStart = ThumbBottCurve.xyEnd
		End If
		
		'Draw thumb curve
		PR_DrawThumbCurve(ThumbBottCurve, g_ThumbTopCurve)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "ThumbPieceOutsideEdge")
		
		ARMDIA1.PR_DrawLine(g_ThumbTopCurve.xyEnd, xyT(8))
		
		
	End Sub
	
	Private Sub PR_DrawSetInThumb(ByRef nInsert As Object)
		
		'Procedure to draw a Radial Set In thumb
		
		'UPGRADE_WARNING: Lower bound of array xyT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim xyT(12)
		Dim xyTmpStart, xyTmp, xyTmpEnd, xyTranslate As XY
		Dim aAngle, nLength, aThumbAngle As Double
		Dim nSeam, nJ_Radius As Double
		Dim nSlantInsert, nThumbSlantLength As Double
		Dim aPalmThumbStart, aPalmThumbEnd As Double
		Dim ThumbTopCurve, ThumbBottCurve As BiArc
		Dim iThumbDirection, ii, iAngleDirection, Success As Short
		Dim nMinDist, nOffSet As Double
		
		
		'Angles
		iAngleDirection = 1
		iThumbDirection = -1
		aThumbAngle = 45
		
		'Mirror Thumb angles for right side
		If g_sSide = "Right" Then
			'Angles
			iAngleDirection = -1
			iThumbDirection = 1
			aThumbAngle = 135
		End If
		
		'Calculate thumb points
		'Draw profile
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xyT(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyT(4) = xyThumb
		'UPGRADE_WARNING: Couldn't resolve default property of object xyT(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyT(1) = xyThumb
		xyT(1).Y = xyThumb.Y - (System.Math.Sqrt(2) * g_Finger(5).PIP)
		
		ARMDIA1.PR_CalcPolar(xyT(4), aThumbAngle, g_Finger(5).Len_Renamed, xyT(3))
		ARMDIA1.PR_CalcPolar(xyT(3), aThumbAngle + (iThumbDirection * 90), g_Finger(5).PIP, xyT(2))
		
		nJ_Radius = 0.125
		nLength = nJ_Radius * System.Math.Tan((22.5) * (PI / 180))
		ARMDIA1.PR_CalcPolar(xyT(4), aThumbAngle, nJ_Radius, xyT(5))
		ARMDIA1.PR_CalcPolar(xyT(4), 90, nJ_Radius, xyT(6))
		ARMDIA1.PR_CalcPolar(xyT(6), 90 + (iThumbDirection * 90), nLength, xyT(7))
		
		nJ_Radius = 0.25
		nLength = nJ_Radius * System.Math.Cos((67.5) * (PI / 180))
		ARMDIA1.PR_CalcPolar(xyT(1), aThumbAngle, nLength, xyT(8))
		ARMDIA1.PR_CalcPolar(xyT(1), 270, nLength, xyT(9))
		ARMDIA1.PR_CalcPolar(xyT(9), 90 + (iThumbDirection * 90), nJ_Radius, xyT(10))
		
		'Draw tip
		If g_Finger(5).Tip = 0 Then
			'Closed tip
			ARMDIA1.PR_CalcPolar(xyT(2), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(2))
			ARMDIA1.PR_CalcPolar(xyT(3), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(3))
			BDYUTILS.PR_CalcMidPoint(xyT(2), xyT(3), xyTmp)
			ARMDIA1.PR_DrawArc(xyTmp, xyT(2), xyT(3))
		Else
			'Open tip
			ARMDIA1.PR_DrawLine(xyT(2), xyT(3))
		End If
		
		ARMDIA1.PR_DrawLine(xyT(8), xyT(2))
		ARMDIA1.PR_DrawLine(xyT(3), xyT(5))
		ARMDIA1.PR_DrawArc(xyT(7), xyT(6), xyT(5))
		ARMDIA1.PR_DrawArc(xyT(10), xyT(8), xyT(9))
		
		
	End Sub
	
	Private Sub PR_DrawTapeNotes(ByRef iVertex As Short, ByRef xyStart As XY, ByRef xyEnd As XY, ByRef nOffSet As Double, ByRef sZipperID As String)
		'Draws the the notes at ecah tape
		Dim xyDatum, xyInsert As XY
		Dim nLength, aAngle, nDec As Double
		Dim sText As String
		Dim iInt As Short
		
		'Exit from sub if no Tape position Given as
		'this implies that the TapeNote is Empty
		'(EG with GLOVE_NORMAL)
		If TapeNote(iVertex).iTapePos = 0 Then Exit Sub
		
		'Draw Construction line
		ARMDIA1.PR_Setlayer("Construct")
		
		ARMDIA1.PR_DrawLine(xyStart, xyEnd)
		If sZipperID <> "" Then
			If sZipperID <> "EOS" Then BDYUTILS.PR_AddDBValueToLast("curvetype", "EOS")
			BDYUTILS.PR_AddDBValueToLast("Zipper", sZipperID)
		End If
		
		'Get center point on which text positions are based
		'Doing it this way accounts for case when xyStart.x > xyEnd.x
		If nOffSet > 0 Then
			aAngle = ARMDIA1.FN_CalcAngle(xyStart, xyEnd)
			ARMDIA1.PR_CalcPolar(xyStart, aAngle, nOffSet, xyDatum)
		Else
			BDYUTILS.PR_CalcMidPoint(xyStart, xyEnd, xyDatum)
		End If
		
		'Insert Symbol
		ARMDIA1.PR_InsertSymbol("glvTapeNotes", xyDatum, 1, 0)
		
		'Setup for Tape Lable
		BDYUTILS.PR_AddDBValueToLast("Data", TapeNote(iVertex).sTapeText)
		
		'MMs
		BDYUTILS.PR_AddDBValueToLast("MM", Trim(Str(TapeNote(iVertex).iMMs)) & "mm")
		
		'Grams
		BDYUTILS.PR_AddDBValueToLast("Grams", Trim(Str(TapeNote(iVertex).iGms)) & "gm")
		
		'Reduction
		BDYUTILS.PR_AddDBValueToLast("Reduction", Trim(Str(TapeNote(iVertex).iRed)))
		
		'Circumference  in Inches and Eighths format
		iInt = Int(TapeNote(iVertex).nCir)
		BDYUTILS.PR_AddDBValueToLast("TapeLengths", Trim(Str(iInt)))
		
		nDec = (TapeNote(iVertex).nCir - iInt)
		If nDec > 0 Then
			iInt = ARMDIA1.round((nDec / EIGHTH) * 10) / 10
			If iInt <> 0 Then BDYUTILS.PR_AddDBValueToLast("TapeLengths2", Trim(Str(iInt)))
		End If
		
		'Check for Notes
		If TapeNote(iVertex).sNote <> "" Then
			'set text justification etc.
			ARMDIA1.PR_SetTextData(RIGHT_, VERT_CENTER, EIGHTH, CURRENT, CURRENT)
			ARMDIA1.PR_MakeXY(xyInsert, xyDatum.X - 0.1, xyDatum.Y + 0.35)
			If TapeNote(iVertex).sNote <> "PLEAT" Then ARMDIA1.PR_Setlayer("Notes")
			ARMDIA1.PR_DrawText(TapeNote(iVertex).sNote, xyInsert, 0.1)
		End If
		
	End Sub
	
	Private Sub PR_DrawThumbCurve(ByRef BottCurve As BiArc, ByRef TopCurve As BiArc)
		'Procedure to draw the specific thumb curve as a fitted polyline.
		'Created to reduce code in PR_DrawThumb.
		'
		'N.B. (Very Carefully)
		'
		'    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'    This is not a general procedure it is specific to the
		'    thumb case defined by the Two bi-arcs.  It is order
		'    and data dependant.
		'    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'
		'
		'UPGRADE_WARNING: Arrays in structure Profile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Profile As Curve
		Dim ii, iVertex As Short
		Dim xyVertex As XY
		Dim aStep, aEnd, aFirst, aTangent, aDiff As Double
		
		iVertex = 1
		Profile.X(iVertex) = BottCurve.xyStart.X
		Profile.Y(iVertex) = BottCurve.xyStart.Y
		
		If BottCurve.nR1 > 0 And BottCurve.nR2 > 0 Then
			
			aFirst = ARMDIA1.FN_CalcAngle(BottCurve.xyR1, BottCurve.xyStart)
			aTangent = ARMDIA1.FN_CalcAngle(BottCurve.xyR1, BottCurve.xyTangent)
			aEnd = ARMDIA1.FN_CalcAngle(BottCurve.xyR2, BottCurve.xyEnd)
			aStep = (aTangent - aFirst) / 4
			For ii = 1 To 3
				iVertex = iVertex + 1
				ARMDIA1.PR_CalcPolar(BottCurve.xyR1, aFirst + (ii * aStep), BottCurve.nR1, xyVertex)
				Profile.X(iVertex) = xyVertex.X
				Profile.Y(iVertex) = xyVertex.Y
			Next ii
			
			iVertex = iVertex + 1
			Profile.X(iVertex) = BottCurve.xyTangent.X
			Profile.Y(iVertex) = BottCurve.xyTangent.Y
			
			aDiff = aEnd - aTangent
			If System.Math.Abs(aDiff) >= 180 Then aDiff = 360 - System.Math.Abs(aDiff)
			aStep = aDiff / 4
			For ii = 1 To 3
				iVertex = iVertex + 1
				ARMDIA1.PR_CalcPolar(BottCurve.xyR2, aTangent + (ii * aStep), BottCurve.nR2, xyVertex)
				Profile.X(iVertex) = xyVertex.X
				Profile.Y(iVertex) = xyVertex.Y
			Next ii
			
		End If
		
		iVertex = iVertex + 1
		Profile.X(iVertex) = BottCurve.xyEnd.X
		Profile.Y(iVertex) = BottCurve.xyEnd.Y
		
		If TopCurve.nR1 > 0 And TopCurve.nR2 > 0 Then
			aFirst = ARMDIA1.FN_CalcAngle(TopCurve.xyR1, TopCurve.xyStart)
			aTangent = ARMDIA1.FN_CalcAngle(TopCurve.xyR1, TopCurve.xyTangent)
			aEnd = ARMDIA1.FN_CalcAngle(TopCurve.xyR2, TopCurve.xyEnd)
			
			If aFirst > 270 Then
				aStep = (aTangent + (360 - aFirst)) / 4
			Else
				aStep = (aTangent - aFirst) / 4
			End If
			For ii = 1 To 3
				iVertex = iVertex + 1
				ARMDIA1.PR_CalcPolar(TopCurve.xyR1, aFirst + (ii * aStep), TopCurve.nR1, xyVertex)
				Profile.X(iVertex) = xyVertex.X
				Profile.Y(iVertex) = xyVertex.Y
			Next ii
			
			iVertex = iVertex + 1
			Profile.X(iVertex) = TopCurve.xyTangent.X
			Profile.Y(iVertex) = TopCurve.xyTangent.Y
			
			aStep = (aEnd - aTangent) / 4
			For ii = 1 To 3
				iVertex = iVertex + 1
				ARMDIA1.PR_CalcPolar(TopCurve.xyR2, aTangent + (ii * aStep), TopCurve.nR2, xyVertex)
				Profile.X(iVertex) = xyVertex.X
				Profile.Y(iVertex) = xyVertex.Y
			Next ii
			
		End If
		
		iVertex = iVertex + 1
		Profile.X(iVertex) = TopCurve.xyEnd.X
		Profile.Y(iVertex) = TopCurve.xyEnd.Y
		
		Profile.n = iVertex
		ARMDIA1.PR_DrawFitted(Profile)
		
	End Sub
	
	Private Sub PR_EnableCir(ByRef Index As Short)
		'Read as part of labCir_DblClick
		labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
		txtCir(Index).Enabled = True
		lblCir(Index).Enabled = True
	End Sub
	
	Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.Control, ByRef sFileName As String)
		'General procedure to create the list section of
		'a combo box reading the data from a file
		
		Dim sLine As String
		Dim fFileNum As Short
		
		fFileNum = FreeFile
		
		If FileLen(sFileName) = 0 Then
			MsgBox(sFileName & "Not found", 48, g_sDialogID)
			Exit Sub
		End If
		
		FileOpen(fFileNum, sFileName, OpenMode.Input)
		Do While Not EOF(fFileNum)
			sLine = LineInput(fFileNum)
			'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Combo_Name.AddItem(sLine)
		Loop 
		FileClose(fFileNum)
		
	End Sub
	
	Private Sub PR_GetDlgData()
		'Subroutine to extract from the Dialogue the data
		'putting it into global variables
		'
		'Common Dialogue elements
		If optFold(1).Checked = True Then g_OnFold = True Else g_OnFold = False
		
		'Get data from tabbed
		PR_GetDlgHandData()
		CADGLEXT.PR_GetDlgAboveWrist()
		
	End Sub
	
	Private Sub PR_GetDlgHandData()
		'Procedure to setup Module level variables based on the data in the
		'dialogue
		'Save having to do this more than once
		
		Dim ii As Short
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("SSTab1"), Object).Tab <> 0 Then CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
		'Insert values in Eighths
		g_iInsertValue(1) = 4
		g_iInsertValue(2) = 5
		g_iInsertValue(3) = 6
		g_iInsertValue(4) = 7
		g_iInsertValue(5) = 8
		g_iInsertValue(6) = 3
		
		'Missing Finger
		For ii = 1 To 5
			g_Missing(ii) = False
		Next ii
		
		'Given Dimensions and missing fingers
		g_MissingFingers = 0
		
		If txtCir(0).Enabled Then g_nFingerDIPCir(1) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(0).Text)) 'Little finger DIP
		If txtCir(1).Enabled Then g_nFingerPIPCir(1) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(1).Text)) 'Little finger PIP
		If g_nFingerPIPCir(1) = 0 Then
			g_MissingFingers = g_MissingFingers + 1
			g_Missing(1) = True
		End If
		
		
		If txtCir(2).Enabled Then g_nFingerDIPCir(2) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(2).Text)) 'Ring finger DIP
		If txtCir(3).Enabled Then g_nFingerPIPCir(2) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(3).Text)) 'Ring finger PIP
		If g_nFingerPIPCir(2) = 0 Then
			g_MissingFingers = g_MissingFingers + 1
			g_Missing(2) = True
		End If
		
		If txtCir(4).Enabled Then g_nFingerDIPCir(3) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(4).Text)) 'Middle Finger DIP
		If txtCir(5).Enabled Then g_nFingerPIPCir(3) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(5).Text)) 'Middle Finger PIP
		If g_nFingerPIPCir(3) = 0 Then
			g_MissingFingers = g_MissingFingers + 1
			g_Missing(3) = True
		End If
		
		If txtCir(6).Enabled Then g_nFingerDIPCir(4) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(6).Text)) 'Index Finger DIP
		If txtCir(7).Enabled Then g_nFingerPIPCir(4) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(7).Text)) 'Index Finger PIP
		If g_nFingerPIPCir(4) = 0 Then
			g_MissingFingers = g_MissingFingers + 1
			g_Missing(4) = True
		End If
		
		If txtCir(8).Enabled Then g_nFingerPIPCir(5) = ARMEDDIA1.fnDisplayToInches(Val(txtCir(8).Text)) 'Thumb
		If g_nFingerPIPCir(5) = 0 Then
			g_Missing(5) = True
		End If
		
		g_nPalm = ARMEDDIA1.fnDisplayToInches(Val(txtCir(9).Text)) 'Palm
		g_nWrist = ARMEDDIA1.fnDisplayToInches(Val(txtCir(10).Text)) 'Wrist
		
		'Finger Lengths
		For ii = 0 To 4
			If txtLen(ii).Enabled Then g_nFingerLen(ii + 1) = ARMEDDIA1.fnDisplayToInches(Val(txtLen(ii).Text))
		Next ii
		
		'Webs
		For ii = 1 To 4
			g_nWeb(ii) = ARMEDDIA1.fnDisplayToInches(Val(txtLen(ii + 4).Text))
		Next ii
		
		'Set finger tips
		For ii = 0 To 2
			If optLittleTip(ii).Checked = True Then
				If g_OnFold Then
					g_Finger(1).Tip = ii + 10
				Else
					g_Finger(1).Tip = ii
				End If
			End If
			If optRingTip(ii).Checked = True Then g_Finger(2).Tip = ii
			If optMiddleTip(ii).Checked = True Then g_Finger(3).Tip = ii
			If optIndexTip(ii).Checked = True Then g_Finger(4).Tip = ii
			If optThumbTip(ii).Checked = True Then g_Finger(5).Tip = ii
		Next ii
		
		g_nCir(1) = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCir"), Object)(10))
		If CType(MainForm.Controls("txtCir"), Object)(11).Enabled Then g_nCir(2) = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCir"), Object)(11))
		If CType(MainForm.Controls("txtCir"), Object)(12).Enabled Then g_nCir(3) = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCir"), Object)(12))
		
		'Thumb styles
		For ii = 0 To 2
			If optThumbStyle(ii).Checked = True Then g_iThumbStyle = ii
		Next ii
		
		'Insert styles
		For ii = 0 To 3
			If optInsertStyle(ii).Checked = True Then g_iInsertStyle = ii
		Next ii
		
		'print fold
		g_PrintFold = False
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkPrintFold.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkPrintFold"), Object).Value <> 0 Then g_PrintFold = True
		
		'Fused fingers
		g_FusedFingers = False
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkFusedFingers.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkFusedFingers"), Object).Value <> 0 Then g_FusedFingers = True
		
	End Sub
	
	Private Sub PR_UpdateDBFields(ByRef sType As String)
		'sType = "Save" for use with save only
		'sType = "Draw" for use with a drawing macro
		
		Dim sSymbol As String
		
		'Glove common
		sSymbol = "glovecommon"
		If txtUidGC.Text = "" Then
			'Insert a new symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x, xyO.y);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  }")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "else")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
		Else
			'Use existing symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hSym = UID (" & QQ & "find" & QC & Val(txtUidGC.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (!hSym) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
		End If
		
		'Update DB fields
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Data" & QCQ & txtDataGC.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		
		'Glove
		sSymbol = "gloveglove"
		If txtUidGlove.Text = "" Then
			'Insert a new symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			If optHand(0).Checked = True Then
				'Left Hand
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x+1.5 , xyO.y);")
			Else
				'Right Hand
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x+3 , xyO.y);")
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  }")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "else")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
		Else
			'Use existing symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hSym = UID (" & QQ & "find" & QC & Val(txtUidGlove.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (!hSym) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
		End If
		
		'Update DB fields
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeLengths" & QCQ & txtTapeLengths.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeLengths2" & QCQ & txtTapeLengths2.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeLengthPt1" & QCQ & txtTapeLengthPt1.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Sleeve" & QCQ & txtSide.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Grams" & QCQ & txtGrams.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Reduction" & QCQ & txtReduction.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeMMs" & QCQ & txtTapeMMs.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Data" & QCQ & txtDataGlove.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "WristPleat" & QCQ & txtWristPleat.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "ShoulderPleat" & QCQ & txtShoulderPleat.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Flap" & QCQ & txtFlap.Text & QQ & ");")
		
		If sType = "Draw" Then
			'Update the fields associated with the PALMER marker
			'This was given an explicit handle in the procedure PR_CreateDrawMacro
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "TapeLengths" & QCQ & txtTapeLengths.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "TapeLengths2" & QCQ & txtTapeLengths2.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "TapeLengthPt1" & QCQ & txtTapeLengthPt1.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "TapeMMs" & QCQ & txtTapeMMs.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Sleeve" & QCQ & txtSide.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Data" & QCQ & txtDataGlove.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "age" & QCQ & txtAge.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "units" & QCQ & txtUnits.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Grams" & QCQ & txtGrams.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Reduction" & QCQ & txtReduction.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "TapeMMs" & QCQ & txtTapeMMs.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Data" & QCQ & txtDataGlove.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "WristPleat" & QCQ & txtWristPleat.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "ShoulderPleat" & QCQ & txtShoulderPleat.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Flap" & QCQ & txtFlap.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hPalmer" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
		End If
	End Sub
	
	Private Sub PR_UpdateDDE()
		
		'Procedure to update the fields used when data is transfered
		'from DRAFIX using DDE.
		'Although the transfer back to DRAFIX is not via DDE we use the same controls
		'simply to illustrate the method by which the data is packed into the
		'fields
		
		'The decisons on format for each is not very neat!
		'The reason is the iteritive development process and the 63 char length
		'limitation on DRAFIX DB Fields.
		
		'Backward compatibility with the CAD-Glove is required in this
		'procedure
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
		Dim iLen, ii As Short
		Dim sLen, sPacked As String
		
		'Initialise
		sPacked = ""
		
		'Pack Circumferences
		For ii = 0 To 10
			iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
			
			If iLen <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(iLen)), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			
			sPacked = sPacked & sLen
			
		Next ii
		
		'Pack Lengths
		For ii = 0 To 8
			iLen = Val(txtLen(ii).Text) * 10 'Shift decimal place
			
			If iLen <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(iLen)), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			
			sPacked = sPacked & sLen
			
		Next ii
		
		'Store to DDE text boxes
		txtTapeLengths.Text = sPacked
		
		'Tip options
		'UPGRADE_NOTE: Closed was upgraded to Closed_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Closed_Renamed(4) As Object
		Dim StdOpen(4) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed_Renamed(0) = optLittleTip(0).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed_Renamed(1) = optRingTip(0).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed_Renamed(2) = optMiddleTip(0).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed_Renamed(3) = optIndexTip(0).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed_Renamed(4) = optThumbTip(0).Checked
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StdOpen(0) = optLittleTip(1).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StdOpen(1) = optRingTip(1).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StdOpen(2) = optMiddleTip(1).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StdOpen(3) = optIndexTip(1).Checked
		'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StdOpen(4) = optThumbTip(1).Checked
		
		sPacked = ""
		For ii = 0 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object Closed_Renamed(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Closed_Renamed(ii) = True Then
				sPacked = sPacked & "0"
				'UPGRADE_WARNING: Couldn't resolve default property of object StdOpen(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf StdOpen(ii) = True Then 
				sPacked = sPacked & "1"
			Else
				sPacked = sPacked & "2"
			End If
		Next ii
		
		'CheckBoxes
		sPacked = sPacked & Trim(Str(chkSlantedInserts.CheckState))
		sPacked = sPacked & Trim(Str(chkPalm.CheckState))
		
		'Tapes past wrist (Store here due to limitations on DB field length
		'Pack Circumferences
		For ii = 11 To 12
			iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
			
			If iLen <> 0 Then
				sLen = New String(" ", 3)
				sLen = RSet(Trim(Str(iLen)), Len(sLen))
			Else
				sLen = New String(" ", 3)
			End If
			
			sPacked = sPacked & sLen
			
		Next ii
		
		'Insert size
		'Translate from Manual glove format to Eighths
		'NOTE:-
		'    This is rather stupid but it made life easier earlier
		If g_iInsertSize = 6 Then
			iLen = 3
		ElseIf g_iInsertSize > 0 Then 
			iLen = 3 + g_iInsertSize
		End If
		sLen = New String(" ", 3)
		sLen = RSet(Trim(Str(iLen)), Len(sLen))
		sPacked = sPacked & sLen
		
		'Reinforced dorsal
		sPacked = sPacked & Trim(Str(chkDorsal.CheckState))
		
		'Thumb Options
		For ii = 0 To 2
			If optThumbStyle(ii).Checked = True Then sPacked = sPacked & Trim(Str(ii))
		Next ii
		'Insert Options
		For ii = 0 To 3
			If optInsertStyle(ii).Checked = True Then sPacked = sPacked & Trim(Str(ii))
		Next ii
		
		'Print Fold
		sPacked = sPacked & Trim(Str(chkPrintFold.CheckState))
		
		'Fused Fingers
		sPacked = sPacked & Trim(Str(chkFusedFingers.CheckState))
		
		'Store to DDE text boxes
		txtTapeLengths2.Text = sPacked
		
		'Hand & Glove common
		'Glove common is used to carry the fabric and insert size and
		'on\off fold data, so that both gloves can be harmonized
		sLen = New String(" ", 2)
		sLen = RSet(Trim(Str(g_iInsertSize)), Len(sLen))
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optFold"), Object)(0).Value = True Then sPacked = sLen & " 0" Else sPacked = sLen & " 1"
		
		If optHand(0).Checked = True Then
			txtSide.Text = "Left"
			sPacked = sPacked & Mid(txtDataGC.Text, 5, 4)
		Else
			txtSide.Text = "Right"
			If Mid(txtDataGC.Text, 1, 4) <> "" Then
				sPacked = Mid(txtDataGC.Text, 1, 4) & sPacked
			Else
				sPacked = "    " & sPacked
			End If
		End If
		txtDataGC.Text = sPacked
		
		'Fabric
		If cboFabric.Text <> "" Then txtFabric.Text = cboFabric.Text
		
		
	End Sub
	
	Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
		'Allows the user to use enter as a tab
		
		System.Windows.Forms.SendKeys.Send("{TAB}")
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'End as the link close event has not
		'occured.
		End
	End Sub
	
	Private Sub txtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Enter
		Dim Index As Short = txtCir.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtCir(Index))
	End Sub
	
	Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
		Dim Index As Short = txtCir.GetIndex(eventSender)
		
		Dim nLen As Double
		Dim sString As String
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtCir(Index))
		'Display value in inches if Valid
		If nLen <> -1 Then
			sString = Trim(ARMEDDIA1.fnInchestoText(nLen))
			If VB.Left(sString, 1) = "-" Then sString = Mid(sString, 2)
			lblCir(Index).Text = sString
		End If
		
	End Sub
	
	Private Sub txtCustFlapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustFlapLength.Enter
		
		CADGLOVE1.PR_Select_Text(txtCustFlapLength)
		
	End Sub
	
	Private Sub txtCustFlapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustFlapLength.Leave
		
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtCustFlapLength)
		
		'Display value in inches if Valid
		If nLen <> -1 Then labCustFlapLength.Text = ARMEDDIA1.fnInchestoText(nLen)
		
		
	End Sub
	
	Private Sub txtExtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExtCir.Enter
		Dim Index As Short = txtExtCir.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtExtCir(Index))
	End Sub
	
	Private Sub txtExtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExtCir.Leave
		Dim Index As Short = txtExtCir.GetIndex(eventSender)
		
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtExtCir(Index))
		
		'Display value in inches if Valid
		If nLen <> -1 Then
			'Note we use a grid here that starts at 0
			CADGLEXT.PR_GrdInchesDisplay(Index - 8, nLen)
		End If
		
	End Sub
	
	Private Sub txtFrontStrapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontStrapLength.Enter
		
		CADGLOVE1.PR_Select_Text(txtFrontStrapLength)
		
	End Sub
	
	Private Sub txtFrontStrapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontStrapLength.Leave
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtFrontStrapLength)
		
		'Display value in inches if Valid
		If nLen <> -1 Then labFrontStrapLength.Text = ARMEDDIA1.fnInchestoText(nLen)
		
	End Sub
	
	Private Sub txtLen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Enter
		Dim Index As Short = txtLen.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtLen(Index))
	End Sub
	
	Private Sub txtLen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Leave
		Dim Index As Short = txtLen.GetIndex(eventSender)
		Dim nLen As Double
		Dim nTipCutBack As Double
		Dim sString As String
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtLen(Index))
		If nLen < 0 Then
			lblLen(Index).Text = ""
			Exit Sub
		End If
		
		If Index <= 4 Then
			'Special case for finger lengths w.r.t. tips
			If Val(txtAge.Text) <= AGE_CUTOFF Then
				nTipCutBack = CHILD_STD_OPEN_TIP
			Else
				nTipCutBack = ADULT_STD_OPEN_TIP
			End If
			
			Select Case Index
				Case 0
					'Little finger
					If optLittleTip(1).Checked Then nLen = nLen - nTipCutBack
				Case 1
					'Ring finger
					If optRingTip(1).Checked Then nLen = nLen - nTipCutBack
				Case 2
					'Middle finger
					If optMiddleTip(1).Checked Then nLen = nLen - nTipCutBack
				Case 3
					'Index finger
					If optIndexTip(1).Checked Then nLen = nLen - nTipCutBack
				Case 4
					'Thumb finger
					If optThumbTip(1).Checked Then nLen = nLen - nTipCutBack
			End Select
		End If
		
		'Display length text
		If nLen > 0 Then
			sString = Trim(ARMEDDIA1.fnInchestoText(nLen))
			If VB.Left(sString, 1) = "-" Then sString = Mid(sString, 2)
			lblLen(Index).Text = sString
		Else
			lblLen(Index).Text = ""
		End If
		
	End Sub
	
	Private Sub txtStrapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStrapLength.Enter
		
		CADGLOVE1.PR_Select_Text(txtStrapLength)
		
	End Sub
	
	Private Sub txtStrapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStrapLength.Leave
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtStrapLength)
		
		'Display value in inches if Valid
		If nLen <> -1 Then labStrap.Text = ARMEDDIA1.fnInchestoText(nLen)
		
	End Sub
	
	Private Sub txtWaistCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Enter
		
		CADGLOVE1.PR_Select_Text(txtWaistCir)
		
	End Sub
	
	Private Sub txtWaistCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Leave
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtWaistCir)
		
		'Display value in inches if Valid
		If nLen <> -1 Then labWaistCir.Text = ARMEDDIA1.fnInchestoText(nLen)
		
		
	End Sub
End Class
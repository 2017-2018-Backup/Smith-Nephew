Option Strict Off
Option Explicit On
Friend Class webspacr
	Inherits System.Windows.Forms.Form
	'Project:   WEBSPACR.MAK
	'Purpose:   Draw web spacers
	'
	'
	'Version:   3.02
	'Date:      20.Sep.96
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Dec 98     GG      Port to VB5
	'
	'Notes:-
	'   This has been quickly cobbled from the MANUAL and
	'   CAD GLOVE code so it's a bit quirky.
	'   but it works and was cheap to do. (I didn't write this bit !Honest Guv!)
	'
	'
	'   'Windows API Functions Declarations
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
	
	'   'MsgBox constant
	'    Const IDCANCEL = 2
	'    Const IDYES = 6
	'    Const IDNO = 7
	
	'Open tip constants
	Const AGE_CUTOFF As Short = 10 'Years
	Const ADULT_STD_OPEN_TIP As Double = 0.5 'Inches
	Const CHILD_STD_OPEN_TIP As Double = 0.375 'Inches
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		End
	End Sub
	
	Private Function FN_CalculateWebSpacer() As Short
		'Procedure to setup Module level variables based on the data in the
		'dialogue
		'Returns true if there is enough data to work with
		'else returns false if there isn't
		
		Dim nWristCir, nThumbWeb, nWeb, nHighWeb, nPalmCir, nThumbCir As Double
		Dim nStrapLengthLittle, nWebToWeb, nStrapLengthMiddleAndIndex, nStrapWidth As Double
		Dim nFigure, nStrapConstruct, nGaunletCrook, nExtendThumb As Double
		Dim nSpace, nSeam, nInsert, nLength As Double
		Dim nBotRadFac, nWristNotch As Double
		Dim nThumbOffset, aThumb, aAngle, nTaperNotchRad, nWristOffset As Double
		Dim xyTmp, xyTmp1 As XY
		Dim iAge, ii, iError As Short
		Dim sError, NL As String
		'UPGRADE_WARNING: Lower bound of array nWebHt was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim nWebHt(4) As Double
		Dim nThumbHt1 As Double
		'Initialise
		NL = Chr(13) 'New Line character
		FN_CalculateWebSpacer = False 'Default to false
		
		'Patients age
		iAge = Val(txtAge.Text)
		
		'Circumferences
		nThumbCir = ARMEDDIA1.fnDisplayToInches(Val(txtCir(8).Text)) 'Thumb
		nPalmCir = ARMEDDIA1.fnDisplayToInches(Val(txtCir(9).Text)) 'Palm
		nWristCir = ARMEDDIA1.fnDisplayToInches(Val(txtCir(10).Text)) 'Wrist
		
		'Webs
		'The thumb web can be lowered when the thumb
		'is being calculated in the Manual glove drawing
		'programme
		'Therefor we use the value that has been calculated and
		'stored
		
		nHighWeb = 0
		For ii = 1 To 4
			nWeb = ARMEDDIA1.fnDisplayToInches(Val(txtLen(ii + 4).Text))
			If ii = 4 Then nWeb = nWeb - g_nThumbWebDrop 'Calculated thumb drop
			nWebHt(ii) = nWeb
			If nWeb > nHighWeb Then nHighWeb = nWeb
		Next ii
		'Last web is thumb web
		nThumbWeb = nWeb
		
		
		'Check on available data
		sError = ""
		If nThumbCir = 0 Then
			sError = sError & "Missing THUMB circumference!" & NL
		End If
		If nPalmCir = 0 Then
			sError = sError & "Missing PALM circumference!" & NL
		End If
		If nWristCir = 0 Then
			sError = sError & "Missing WRIST circumference!" & NL
		End If
		If nHighWeb = 0 Then
			sError = sError & "Missing WEB lengths!" & NL
		End If
		If nThumbWeb = 0 Then
			sError = sError & "Missing THUMB WEB length!" & NL
		End If
		
		If g_iInsertSize = 0 Then
			sError = sError & "Insert not specified!" & NL
		End If
		
		If g_iInsertSize < 3 Or g_iInsertSize > 8 Then
			sError = sError & "Specified insert is not found on table ""ONE INSERT WITH LARGE SEAMS"", ref GOP 01-02/28, page 4!" & NL
		End If
		
		If sError <> "" Then
			sError = "Unable to draw Web Spacer because of missing or incomplete data." & NL & sError
			MsgBox(sError, 16, g_sDialogID)
			FN_CalculateWebSpacer = False
			Exit Function
		End If
		
		'Setup figuring and construction factors
		'NB age related construction factors
		nGaunletCrook = 3 * SIXTEENTH
		nFigure = 0.86
		nSeam = 0.25
		nExtendThumb = 0.5
		nInsert = g_iInsertSize * EIGHTH
		
		If iAge > 10 Then
			'Adults
			nStrapLengthMiddleAndIndex = 1.25
			nStrapLengthLittle = 1
			nStrapWidth = 0.5
			nStrapConstruct = 0.25
		Else
			'Children
			nStrapLengthMiddleAndIndex = 1
			nStrapLengthLittle = 0.875
			nStrapWidth = 0.375
			nStrapConstruct = 0.1875
		End If
		
		'Figure dimensions
		nPalmCir = ((nPalmCir * nFigure) - nInsert) + QUARTER
		nWristCir = nWristCir * nFigure
		nWebToWeb = (nHighWeb - nThumbWeb) * nFigure
		nHighWeb = nHighWeb * nFigure
		nThumbCir = FN_FigureThumb(nThumbCir)
		If nThumbCir <= 0 Then
			sError = "Unable to figure the thumb with the given insert and Thumb circumference." & NL
			MsgBox(sError, 16, g_sDialogID)
			FN_CalculateWebSpacer = False
			Exit Function
		End If
		
		'Calculate points
		'Work w.r.t the right hand first
		'when we have caculated the thumb and we know the
		'full extents of the drawing then we can mirror and
		'translate the final image etc.
		
		'NOTE
		'The construction points xyW(n) are completly
		'arbitary and given im no particular order
		'refer to drawing
		
		ARMDIA1.PR_MakeXY(xyFold, 10, 0)
		
		ARMDIA1.PR_MakeXY(xyW(1), xyFold.X - (nWristCir / 2), 0)
		ARMDIA1.PR_MakeXY(xyW(4), xyFold.X + (nWristCir / 2), xyW(1).y)
		
		ARMDIA1.PR_MakeXY(xyW(2), xyFold.X - (nPalmCir / 2), nHighWeb)
		ARMDIA1.PR_MakeXY(xyW(3), xyFold.X + (nPalmCir / 2), xyW(2).y)
		
		ARMDIA1.PR_MakeXY(xyW(8), xyW(3).X - 0.5, xyW(3).y)
		ARMDIA1.PR_MakeXY(xyW(9), xyFold.X + nSeam, xyW(3).y - nSeam)
		ARMDIA1.PR_MakeXY(xyW(5), xyW(2).X + nSeam, xyW(2).y)
		
		ARMDIA1.PR_MakeXY(xyW(6), xyW(3).X, xyW(3).y - nWebToWeb)
		
		'Finger strap points
		'Assume for the mean time no missing fingers
		nLength = ARMDIA1.FN_CalcLength(xyW(5), xyW(9))
		nSpace = (nLength - (3 * nStrapWidth)) / 4
		
		ARMDIA1.PR_MakeXY(xyW(10), xyW(5).X + nSpace, xyW(2).y)
		
		ARMDIA1.PR_CalcPolar(xyW(10), 90, nStrapLengthMiddleAndIndex, xyW(20))
		ARMDIA1.PR_CalcPolar(xyW(20), 0, nStrapWidth, xyW(21))
		
		nLength = (nStrapLengthMiddleAndIndex + nSeam) - (nSpace / 2)
		ARMDIA1.PR_CalcPolar(xyW(21), 270, nLength, xyW(11))
		ARMDIA1.PR_CalcPolar(xyW(11), 0, nSpace, xyW(12))
		ARMDIA1.PR_CalcPolar(xyW(12), 90, nLength, xyW(22))
		ARMDIA1.PR_CalcPolar(xyW(22), 0, nStrapWidth, xyW(23))
		ARMDIA1.PR_CalcPolar(xyW(23), 270, nLength, xyW(13))
		ARMDIA1.PR_CalcPolar(xyW(13), 0, nSpace, xyW(14))
		
		nLength = (nStrapLengthLittle + nSeam) - (nSpace / 2)
		ARMDIA1.PR_CalcPolar(xyW(14), 90, nLength, xyW(24))
		ARMDIA1.PR_CalcPolar(xyW(24), 0, nStrapWidth, xyW(25))
		ARMDIA1.PR_CalcPolar(xyW(25), 270, nLength, xyW(15))
		
		ARMDIA1.PR_MakeXY(xyW(26), xyW(15).X + (nSpace / 2), xyW(9).y)
		
		
		'Thumb points, construction circles etc.
		'    nNotchRad = .09375
		'    nBotRadFac = 1.3
		'    aThumb = 53
		nNotchRad = 0.05
		nBotRadFac = 1.4
		aThumb = 70
		nThumbOffset = 0.5
		nWristOffset = 0.5
		nWristNotch = 0.125
		
		nTopThumbRad = nThumbCir + nNotchRad
		nBotThumbRad = ((xyW(6).y - xyW(4).y) - nThumbCir)
		
		If ((xyW(6).y - xyW(4).y) - nThumbCir) <= 0 Then
			sError = "Query thumb circumference as it is impossible to draw the thumb with the given dimension!." & NL
			MsgBox(sError, 16, g_sDialogID)
			FN_CalculateWebSpacer = False
			Exit Function
		Else
			nBotThumbRad = nBotThumbRad * nBotRadFac
		End If
		
		xyTopThumbRad.X = xyW(6).X + nNotchRad
		xyTopThumbRad.y = xyW(6).y + nNotchRad
		
		'Calculate notch points
		'Note use of factor to create a taper of the thumb and the
		'seam into the notch
		nTaperNotchRad = nNotchRad * 0.95
		xyT(1).X = xyTopThumbRad.X - nTaperNotchRad
		xyT(1).y = xyTopThumbRad.y
		ARMDIA1.PR_CalcPolar(xyTopThumbRad, 270 + aThumb, nTopThumbRad + 1, xyTmp)
		
		iError = BDYUTILS.FN_CirLinInt(xyTopThumbRad, xyTmp, xyTopThumbRad, nTaperNotchRad, xyT(2))
		
		'Other thumb points at top
		ARMDIA1.PR_CalcPolar(xyTopThumbRad, 270 + aThumb, nTopThumbRad, xyT(5))
		ARMDIA1.PR_CalcPolar(xyT(5), aThumb, nThumbOffset, xyT(4))
		ARMDIA1.PR_CalcPolar(xyT(4), aThumb + 90, nThumbCir, xyT(3))
		
		'Calculate bottom circle
		'Note use of cosine rule
		Dim nC, nA, nB, nCosA As Double
		nA = nTopThumbRad + nBotThumbRad
		nB = ARMDIA1.FN_CalcLength(xyTopThumbRad, xyW(4))
		nC = nBotThumbRad
		nCosA = ((nB ^ 2 + nC ^ 2) - nA ^ 2) / (2 * nB * nC)
		aAngle = ARMDIA1.FN_CalcAngle(xyW(4), xyTopThumbRad) - (BDYUTILS.Arccos(nCosA) * (180 / PI))
		ARMDIA1.PR_CalcPolar(xyW(4), aAngle, nBotThumbRad, xyBotThumbRad)
		
		'Notch at EOS
		CADGLEXT.PR_GetDlgAboveWrist()
		If g_ExtendTo = GLOVE_NORMAL Then
			'Calculate notch
			xyTmp.X = xyBotThumbRad.X
			xyTmp.y = xyFold.y + nWristOffset
			xyTmp1.X = xyFold.X
			xyTmp1.y = xyTmp.y
			
			If BDYUTILS.FN_CirLinInt(xyTmp, xyTmp1, xyBotThumbRad, nBotThumbRad, xyN(3)) Then
				'Result in xyN(3)
			Else
				'No intesection on bottom circle
				aAngle = ARMDIA1.FN_CalcAngle(xyW(4), xyT(4))
				nLength = nWristOffset / System.Math.Sin(aAngle * (PI / 180))
				ARMDIA1.PR_CalcPolar(xyW(4), aAngle, nLength, xyN(3))
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xyN(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyN(2) = xyN(3)
			xyN(2).X = xyN(2).X - nWristNotch
			'UPGRADE_WARNING: Couldn't resolve default property of object xyN(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyN(1) = xyW(4)
		Else
			'If there is data for an extension then we calculate it here
			CADGLEXT.PR_CalculateExtension()
		End If
		
		'Exit with true if we get this far
		FN_CalculateWebSpacer = True
		
	End Function
	
	Private Function FN_FigureThumb(ByRef nThumbCir As Object) As Double
		'Figure the thumb
		'Based on the Table
		'    "ONE INSERT WITH LARGE SEAMS"
		'GOP 01-02/28, page 4
		'Date May 1991
		'
		' Globals
		'    g_iInsertSize   -   Given insert size on EIGHTHs
		'
		' Returns
		'    Figured dimension in inches
		'    0 if not able to calculate
		
		Dim iThumb As Short
		
		FN_FigureThumb = 0
		
		If g_iInsertSize < 3 Or g_iInsertSize > 8 Then Exit Function
		
		'Translate thumb circumference to EIGHTHs
		'UPGRADE_WARNING: Couldn't resolve default property of object nThumbCir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iThumb = ARMDIA1.round(nThumbCir / EIGHTH)
		
		'Check that thumb lies within range
		If iThumb < 8 Then Exit Function
		
		'Calculate Table entry for given thumb and insert
		iThumb = (3 + (7 - g_iInsertSize)) + (iThumb - 8)
		
		If iThumb < 0 Then Exit Function
		
		FN_FigureThumb = iThumb * SIXTEENTH
		
	End Function
	
	Private Function FN_LengthWristToEOS() As Double
		'Calculates the distance from the EOS to the
		'wrist based on values given in the dialogue
		'
		Static ii As Short
		Static nLen, nSpace As Double
		
		'Get the data from the dialogue
		nLen = 0
		
		Select Case g_ExtendTo
			Case GLOVE_NORMAL
				nLen = 0
			Case GLOVE_ELBOW, GLOVE_PASTWRIST
				For ii = 1 To g_iNumTapesWristToEOS - 1
					nSpace = 1.375
					If ii = 1 And g_nPleats(1) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
					If ii = 2 And g_nPleats(2) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
					If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
					If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
					nLen = nLen + nSpace
				Next ii
				
		End Select
		
		FN_LengthWristToEOS = nLen
		
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
		'Note this is a bit dodgey if the drawing contains more than 9999 entities
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
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Dim nAge, ii, nn, jj As Short
		Dim nValue As Double
		Dim sTapesBeyondWrist As String
		
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
		
		If txtUidGlove.Text = "" Then
			'Disable text boxes that will not recieve data
			
			For ii = 20 To 23
				labCir_DoubleClick(labCir.Item(ii), New System.EventArgs())
			Next ii
			
			labCir_DoubleClick(labCir.Item(11), New System.EventArgs())
			txtLen(4).Text = CStr(0.5 / g_nUnitsFac)
			txtLen_Leave(txtLen.Item(4), New System.EventArgs())
			
			optExtendTo(1).Enabled = False
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
		
		sInsert = Mid(txtTapeLengths2.Text, 15, 2)
		If sInsert <> "" Then
			g_iInsertSize = Val(sInsert)
		Else
			g_iInsertSize = 4
		End If
		If g_iInsertSize <= cboInsert.Items.Count Then
			cboInsert.SelectedIndex = g_iInsertSize
		Else
			cboInsert.SelectedIndex = 4
		End If
		
		
		cboFabric.Text = txtFabric.Text
		
		'Default always to normal webspacer
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True
		
		If Val(txtTapeLengthPt1.Text) > 0 Then
			CADGLEXT.PR_GetExtensionDDE_Data()
			If Not g_TapesFound Then
				'disable access to the extension
				'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.TabEnabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(MainForm.Controls("SSTab1"), Object).TabEnabled(1) = False
				CType(MainForm.Controls("optExtendTo"), Object)(1).Enabled = False
			End If
		Else
			'disable access to the extension
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.TabEnabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CType(MainForm.Controls("SSTab1"), Object).TabEnabled(1) = False
			CType(MainForm.Controls("optExtendTo"), Object)(1).Enabled = False
		End If
		
		
		'Revised Web hts
		If CType(MainForm.Controls("txtDataGlove"), Object).Text <> "" Then
			ii = 6
			nValue = Val(Mid(CType(MainForm.Controls("txtDataGlove"), Object).Text, (ii * 2) + 1, 2))
			'Amount thumbweb was dropped in 1/16ths
			g_nThumbWebDrop = nValue * 0.0625
		End If
		
		
		'Show after link is closed
		Show()
		'Use ok click to draw the web spacer automatically
		' OK_Click
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub webspacr_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		'Hide during DDE transfer
		Hide()
		
		'Enable Timeout timer
		'Use 6 seconds (a 1/10th of a minute).
		'Timeout is disabled in Link close
		Timer1.Interval = 10000
		Timer1.Enabled = True
		
		MainForm = Me
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(MainForm.Width)) / 2)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(MainForm.Height)) / 2)
		
		'Get the path to the JOBST system installation directory
		g_sPathJOBST = fnPathJOBST()
		
		'Clear DDE transfer boxes
		txtUidGlove.Text = ""
		txtTapeLengths.Text = ""
		txtTapeLengths2.Text = ""
		txtTapeLengthPt1.Text = ""
		txtWristPleat.Text = ""
		txtShoulderPleat.Text = ""
		txtSide.Text = ""
		txtWorkOrder.Text = ""
		txtUidGC.Text = ""
		txtFabric.Text = ""
		txtUidMPD.Text = ""
		txtDataGlove.Text = ""
		
		'Default units
		'g_nUnitsFac = 1             'Metric
		g_nUnitsFac = 10 / 25.4 'Inches
		
		'Set up fabric list
		LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\FABRIC.DAT")
		
		'Setup display inches grid
		grdInches.set_ColWidth(0, 615)
		grdInches.set_ColAlignment(0, 2)
		For ii = 0 To 15
			grdInches.set_RowHeight(ii, 270)
		Next ii
		
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
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim sTask As String
		
		'Don't allow multiple clicking
		'
		OK.Enabled = False
		
		If FN_CalculateWebSpacer() Then
			'Display and hourglass
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			
			'N.B. Use of local JOBST Directory C:\JOBST
			PR_CreateDrawMacro("C:\JOBST\DRAW.D")
			
			'Find Drafix and start macro
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
		
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Show()
		Exit Sub
		
		
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
	
	Private Sub PR_CalculateExtension()
		'Procedure to calculate the POINTS
		'used to draw the extension of the glove above the
		'wrist
		'The Procedure also supplies the data that is to be
		'printed at each tape
		'N.B.
		
		Dim nValue, nMidWristX, nSpacing As Double
		Dim iLastTape, ii, iStartTape, iError As Short
		Dim nFiguredValue As Double
		Dim iVertex As Short
		Dim nWristOffset, nHt, nWristNotch As Double
		Dim xyTmp As XY
		
		
		
		If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_PASTWRIST Then
			'As we have recalculated the wrist we need to start at the
			'wrist
			nWristOffset = 0.5
			nWristNotch = 0.125
			
			nHt = CADGLEXT.FN_LengthWristToEOS()
			UlnarProfile.n = g_iNumTapesWristToEOS
			UlnarProfile.y(1) = nHt
			
			RadialProfile.n = g_iNumTapesWristToEOS
			RadialProfile.y(1) = nHt
			
			iStartTape = g_iWristPointer
			iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)
			
			'Note: As we can allow the wrist to be any tape we need a
			'vertex counter
			iVertex = 1
			
			For ii = iStartTape To iLastTape
				'NB use of 1/2 scale
				nValue = g_nCir(ii) * 0.86
				UlnarProfile.X(iVertex) = xyFold.X - (nValue / 2)
				RadialProfile.X(iVertex) = xyFold.X + (nValue / 2)
				
				'Setup the notes for the vertex
				'We do this here as we have all the data and it simplifies
				'the drawing side
				TapeNote(iVertex).sTapeText = LTrim(Mid(g_sTapeText, ((ii + 1) * 3) + 1, 3))
				TapeNote(iVertex).iTapePos = ii
				TapeNote(iVertex).nCir = g_nCir(ii)
				
				'Standard spacing
				nSpacing = 1.375
				
				'Account all for pleats
				'Wrist (as we have started at the wrist we use, iStartTape + 1
				'and iStartTape + 2 in this case)
				If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
				If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
				
				'Axilla
				If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
				If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
				
				'We are working from the top (wrist) down
				'iVertex = 1 is set before the For Loop (Values from wrist)
				If iVertex <> 1 Then
					UlnarProfile.y(iVertex) = UlnarProfile.y(iVertex - 1) - nSpacing
					RadialProfile.y(iVertex) = RadialProfile.y(iVertex - 1) - nSpacing
					If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
				End If
				
				'Increment vertex count
				iVertex = iVertex + 1
				
			Next ii
			
			'
			'Calculate Notch
			ARMDIA1.PR_MakeXY(xyN(1), RadialProfile.X(RadialProfile.n), RadialProfile.y(RadialProfile.n))
			ARMDIA1.PR_MakeXY(xyTmp, RadialProfile.X(RadialProfile.n - 1), RadialProfile.y(RadialProfile.n - 1))
			If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
				'Cutback by 3/4 of an inch
				iError = BDYUTILS.FN_CirLinInt(xyN(1), xyTmp, xyN(1), 0.75, xyN(1))
			End If
			
			iError = BDYUTILS.FN_CirLinInt(xyN(1), xyTmp, xyN(1), nWristOffset, xyN(3))
			'UPGRADE_WARNING: Couldn't resolve default property of object xyN(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyN(2) = xyN(3)
			xyN(2).X = xyN(2).X - nWristNotch
			
			'Reset end of profile
			RadialProfile.X(RadialProfile.n) = xyN(3).X
			RadialProfile.y(RadialProfile.n) = xyN(3).y
			UlnarProfile.X(UlnarProfile.n) = xyFold.X - (xyN(3).X - xyFold.X)
			UlnarProfile.y(UlnarProfile.n) = xyN(3).y
			
		End If
		
	End Sub
	
	Private Sub PR_CreateDrawMacro(ByRef sFileName As String)
		'
		' GLOBALS
		'            g_sSide
		'
		
		Dim nInsert, nFingerSpace As Double
		Dim ii As Short
		Dim xyPt3, xyPt1, xyPt2, xyTextAlign As XY
		Dim xyWristPt1, xyWristPt2 As XY
		Dim xyTmp As XY
		Dim aInc, aAngle, nTranslate As Double
		'UPGRADE_WARNING: Arrays in structure Thumb may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Thumb As Curve
		Dim sText, sWorkOrder As String
		
		'Initialise
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)
		
		'Draw on correct layer
		ARMDIA1.PR_Setlayer("Template" & g_sSide)
		
		'    PR_DebugPoint "xyFold", xyFold
		'    For ii = 1 To 20
		'        PR_DebugPoint "W" & Trim$(Str$(ii)), xyW(ii)
		'    Next ii
		'    For ii = 1 To 8
		'        PR_DebugPoint "T" & Trim$(Str$(ii)), xyT(ii)
		'    Next ii
		
		'PR_DebugPoint "xyTopThumbRad", xyTopThumbRad
		'PR_DebugPoint "xyBotThumbRad", xyBotThumbRad
		
		'Having established the points we now draw them
		'We must mirror for the right hand side and
		'translate all points relative to our temporary datum xyFold
		
		'Translate points relative to xyFold
		Dim nYTranslate As Double
		nYTranslate = CADGLEXT.FN_LengthWristToEOS()
		nTranslate = -(xyFold.X + (xyFold.X - xyT(4).X))
		xyFold.X = xyFold.X + nTranslate
		xyFold.y = xyFold.y + nYTranslate
		
		For ii = 1 To 5
			xyT(ii).X = xyT(ii).X + nTranslate
			xyT(ii).y = xyT(ii).y + nYTranslate
		Next ii
		
		For ii = 1 To 3
			xyN(ii).X = xyN(ii).X + nTranslate
		Next ii
		
		For ii = 1 To 26
			xyW(ii).X = xyW(ii).X + nTranslate
			xyW(ii).y = xyW(ii).y + nYTranslate
			'UPGRADE_WARNING: Couldn't resolve default property of object xyW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If g_sSide = "Left" Then CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyW(ii))
		Next ii
		xyTopThumbRad.X = xyTopThumbRad.X + nTranslate
		xyTopThumbRad.y = xyTopThumbRad.y + nYTranslate
		xyBotThumbRad.X = xyBotThumbRad.X + nTranslate
		xyBotThumbRad.y = xyBotThumbRad.y + nYTranslate
		
		
		'Drawfingerstraps
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyW(2), 0)
		BDYUTILS.PR_AddVertex(xyW(10), 0)
		BDYUTILS.PR_AddVertex(xyW(20), 0)
		BDYUTILS.PR_AddVertex(xyW(21), 0)
		If g_sSide = "Left" Then BDYUTILS.PR_AddVertex(xyW(11), -1) Else BDYUTILS.PR_AddVertex(xyW(11), 1)
		BDYUTILS.PR_AddVertex(xyW(12), 0)
		BDYUTILS.PR_AddVertex(xyW(22), 0)
		BDYUTILS.PR_AddVertex(xyW(23), 0)
		If g_sSide = "Left" Then BDYUTILS.PR_AddVertex(xyW(13), -1) Else BDYUTILS.PR_AddVertex(xyW(13), 1)
		BDYUTILS.PR_AddVertex(xyW(14), 0)
		BDYUTILS.PR_AddVertex(xyW(24), 0)
		BDYUTILS.PR_AddVertex(xyW(25), 0)
		If g_sSide = "Left" Then BDYUTILS.PR_AddVertex(xyW(15), -0.414) Else BDYUTILS.PR_AddVertex(xyW(15), 0.414)
		BDYUTILS.PR_AddVertex(xyW(26), 0)
		BDYUTILS.PR_AddVertex(xyW(9), 0)
		BDYUTILS.PR_AddVertex(xyW(8), 0)
		BDYUTILS.PR_AddVertex(xyW(3), 0)
		BDYUTILS.PR_EndPoly()
		
		Dim nLength As Double
		
		If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_PASTWRIST Then
			
			For ii = 1 To RadialProfile.n
				UlnarProfile.X(ii) = UlnarProfile.X(ii) + nTranslate
				RadialProfile.X(ii) = RadialProfile.X(ii) + nTranslate
			Next ii
			
			ARMDIA1.PR_DrawPoly(UlnarProfile)
			ARMDIA1.PR_DrawPoly(RadialProfile)
			
			'Draw notes for each tape
			If g_ExtendTo = GLOVE_ELBOW Then
				nLength = -1
				For ii = 1 To RadialProfile.n
					ARMDIA1.PR_MakeXY(xyPt1, UlnarProfile.X(ii), UlnarProfile.y(ii))
					ARMDIA1.PR_MakeXY(xyPt2, RadialProfile.X(ii), RadialProfile.y(ii))
					If ii = 1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyWristPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyWristPt1 = xyPt1
						'UPGRADE_WARNING: Couldn't resolve default property of object xyWristPt2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyWristPt2 = xyPt2
					End If
					If ii <> RadialProfile.n Then
						PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "")
					Else
						If TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
							'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							xyTmp = xyN(1)
							'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							xyPt3 = xyN(1)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							xyTmp = xyN(1)
							xyTmp.y = 0
							'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							xyPt3 = xyTmp
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyPt3)
						PR_DrawTapeNotes(ii, xyTmp, xyPt3, nLength, "")
					End If
					
					If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
						'Draw elbow line
						ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
						ARMDIA1.PR_Setlayer("Notes")
						ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
						BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
						xyPt3.y = xyPt3.y + 0.0625
						ARMDIA1.PR_DrawText("E", xyPt3, 0.1)
					End If
				Next ii
				
				'Center line
				ARMDIA1.PR_Setlayer("Construct")
				BDYUTILS.PR_CalcMidPoint(xyWristPt1, xyWristPt2, xyTmp)
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				ARMDIA1.PR_DrawLine(xyPt3, xyTmp)
			End If
			
			'Wrist line
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
			
			ARMDIA1.PR_Setlayer("Notes")
			ARMDIA1.PR_DrawLine(xyW(1), xyW(4))
			BDYUTILS.PR_CalcMidPoint(xyW(1), xyW(4), xyPt3)
			xyPt3.y = xyPt3.y + 0.0625
			ARMDIA1.PR_DrawText("W", xyPt3, 0.1)
			
			'Elastic Text
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt1 = xyN(1)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt2 = xyN(1)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyPt2)
			
			sText = Trim(ARMEDDIA1.fnInchestoText(ARMDIA1.FN_CalcLength(xyPt1, xyPt2) + 1.5))
			BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
			xyPt3.y = xyPt3.y + 0.375
			sText = sText & "\"" ELASTIC"
			ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
			
			'Elastic for children
			If Val(txtAge.Text) <= 10 Then
				'           PR_Setlayer "Notes"
				'           PR_CalcMidPoint xyPt1, xyPt2, xyPt3
				xyPt3.y = xyPt3.y - 0.25
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_DrawText("1/2\"" ELASTIC", xyPt3, 0.1)
			End If
			
			
			'EOS
			ARMDIA1.PR_Setlayer("Template" & g_sSide)
			
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
			
		Else
			
			'Elastic text at EOS
			ARMDIA1.PR_Setlayer("Notes")
			sText = Trim(ARMEDDIA1.fnInchestoText(ARMDIA1.FN_CalcLength(xyW(1), xyW(4)) + 1.5))
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
			BDYUTILS.PR_CalcMidPoint(xyW(1), xyW(4), xyPt3)
			xyPt3.y = xyPt3.y + 0.25
			sText = sText & "\"" ELASTIC"
			ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
			
			'EOS at wrist
			ARMDIA1.PR_Setlayer("Template" & g_sSide)
			ARMDIA1.PR_DrawLine(xyW(1), xyW(4))
			
		End If
		
		'Drawthumb (LEFT)
		If g_sSide = "Left" Then ARMDIA1.PR_DrawLine(xyT(1), xyW(2)) Else ARMDIA1.PR_DrawLine(xyT(1), xyW(3))
		
		ARMDIA1.PR_DrawArc(xyTopThumbRad, xyT(1), xyT(2))
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyT(2), 0)
		BDYUTILS.PR_AddVertex(xyT(3), 0)
		BDYUTILS.PR_AddVertex(xyT(4), 0)
		BDYUTILS.PR_AddVertex(xyT(5), 0)
		BDYUTILS.PR_EndPoly()
		
		'Notch
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyN(1), 0)
		BDYUTILS.PR_AddVertex(xyN(2), 0)
		BDYUTILS.PR_AddVertex(xyN(3), 0)
		BDYUTILS.PR_EndPoly()
		
		'Calculate thumb curve
		'For the left then mirror for right
		Thumb.n = 8
		
		If g_ExtendTo = GLOVE_NORMAL Then
			'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyTmp = xyN(3)
		Else
			If g_sSide = "Left" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyTmp = xyW(1)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyTmp = xyW(4)
			End If
		End If
		Thumb.n = 1
		Thumb.X(Thumb.n) = xyTmp.X
		Thumb.y(Thumb.n) = xyTmp.y
		
		aAngle = ARMDIA1.FN_CalcAngle(xyBotThumbRad, xyTmp)
		aInc = (aAngle - ARMDIA1.FN_CalcAngle(xyBotThumbRad, xyTopThumbRad)) / 5
		For ii = 2 To 4
			aAngle = aAngle - aInc
			ARMDIA1.PR_CalcPolar(xyBotThumbRad, aAngle, nBotThumbRad, xyTmp)
			If xyTmp.y > xyN(3).y Then
				Thumb.n = Thumb.n + 1
				Thumb.X(Thumb.n) = xyTmp.X
				Thumb.y(Thumb.n) = xyTmp.y
			End If
		Next ii
		
		aAngle = ARMDIA1.FN_CalcAngle(xyTopThumbRad, xyBotThumbRad)
		aInc = (ARMDIA1.FN_CalcAngle(xyTopThumbRad, xyT(5)) - aAngle) / 4
		For ii = 5 To 8
			aAngle = aAngle + aInc
			ARMDIA1.PR_CalcPolar(xyTopThumbRad, aAngle, nTopThumbRad, xyTmp)
			If xyTmp.y > xyN(3).y Then
				Thumb.n = Thumb.n + 1
				Thumb.X(Thumb.n) = xyTmp.X
				Thumb.y(Thumb.n) = xyTmp.y
			End If
		Next ii
		
		ARMDIA1.PR_DrawFitted(Thumb)
		
		'Draw thumb (Right)
		For ii = 1 To 5
			'UPGRADE_WARNING: Couldn't resolve default property of object xyT(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyT(ii))
		Next ii
		For ii = 1 To 3
			'UPGRADE_WARNING: Couldn't resolve default property of object xyN(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyN(ii))
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object xyTopThumbRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CADUTILS.PR_MirrorPointInYaxis(xyFold.X, xyTopThumbRad)
		'UPGRADE_WARNING: Couldn't resolve default property of object Thumb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CADUTILS.PR_MirrorCurveInYaxis(xyFold.X, Thumb)
		If g_sSide = "Left" Then ARMDIA1.PR_DrawLine(xyT(1), xyW(3)) Else ARMDIA1.PR_DrawLine(xyT(1), xyW(2))
		
		ARMDIA1.PR_DrawArc(xyTopThumbRad, xyT(1), xyT(2))
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyT(2), 0)
		BDYUTILS.PR_AddVertex(xyT(3), 0)
		BDYUTILS.PR_AddVertex(xyT(4), 0)
		BDYUTILS.PR_AddVertex(xyT(5), 0)
		BDYUTILS.PR_EndPoly()
		
		'notch
		BDYUTILS.PR_StartPoly()
		BDYUTILS.PR_AddVertex(xyN(1), 0)
		BDYUTILS.PR_AddVertex(xyN(2), 0)
		BDYUTILS.PR_AddVertex(xyN(3), 0)
		BDYUTILS.PR_EndPoly()
		ARMDIA1.PR_DrawFitted(Thumb)
		
		'PR_DrawCircle xyBotThumbRad, nBotThumbRad
		'PR_DrawCircle xyTopThumbRad, nTopThumbRad
		'PR_DrawCircle xyTopThumbRad, nNotchRad
		
		'Add notes etc.
		ARMDIA1.PR_Setlayer("Notes")
		
		
		'Insert Size (Text)
		'Position Insert text Size w.r.t Little Finger web
		nInsert = g_iInsertSize * EIGHTH
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
		
		If g_sSide = "Left" Then
			BDYUTILS.PR_CalcMidPoint(xyW(15), xyW(14), xyTextAlign)
		Else
			BDYUTILS.PR_CalcMidPoint(xyW(10), xyW(11), xyTextAlign)
		End If
		xyTextAlign.y = xyTextAlign.y - 0.5
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPt3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPt3 = xyTextAlign
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		'Arrow
		BDYUTILS.PR_DrawMarkerNamed("closed arrow", xyW(9), 0.125, 0.05, 270)
		
		'Patient Details
		ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
		ARMDIA1.PR_MakeXY(xyPt3, xyPt3.X, xyPt3.y - 0.3)
		If txtWorkOrder.Text = "" Then
			sWorkOrder = "-"
		Else
			sWorkOrder = txtWorkOrder.Text
		End If
		sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim(Mid(txtFabric.Text, 4))
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		'Other patient details in black on layer construct
		ARMDIA1.PR_Setlayer("Construct")
		ARMDIA1.PR_MakeXY(xyPt3, xyPt3.X, xyPt3.y - 0.8)
		sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
		ARMDIA1.PR_DrawText(sText, xyPt3, 0.1)
		
		'This will only save if there is no existing glove data
		PR_UpdateDDE()
		PR_UpdateDBFields()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_DisableCir(ByRef Index As Short)
		'Read as part of labCir_DblClick
		labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		txtCir(Index).Enabled = False
		lblCir(Index).Enabled = False
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
		
		
		'Circumference  in Inches and Eighths format
		iInt = Int(TapeNote(iVertex).nCir)
		BDYUTILS.PR_AddDBValueToLast("TapeLengths", Trim(Str(iInt)))
		
		nDec = (TapeNote(iVertex).nCir - iInt)
		If nDec > 0 Then
			iInt = ARMDIA1.round((nDec / EIGHTH) * 10) / 10
			If iInt <> 0 Then BDYUTILS.PR_AddDBValueToLast("TapeLengths2", Trim(Str(iInt)))
		End If
		
		BDYUTILS.PR_AddDBValueToLast("Reduction", "14")
		
		'Check for Notes
		If TapeNote(iVertex).sNote <> "" Then
			'set text justification etc.
			ARMDIA1.PR_SetTextData(RIGHT_, VERT_CENTER, EIGHTH, CURRENT, CURRENT)
			ARMDIA1.PR_MakeXY(xyInsert, xyDatum.X - 0.1, xyDatum.y + 0.35)
			If TapeNote(iVertex).sNote <> "PLEAT" Then ARMDIA1.PR_Setlayer("Notes")
			ARMDIA1.PR_DrawText(TapeNote(iVertex).sNote, xyInsert, 0.1)
		End If
		
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
	
	Private Sub PR_GetDlgAboveWrist()
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
		End If
		
		
		Dim iIndex As Short
		If g_ExtendTo = GLOVE_NORMAL Then
			'Get values from Hand Data TAB
			For ii = 10 To 12
				nCir = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCir"), Object)(ii))
				g_nCir(ii - 9) = nCir
			Next ii
			g_iFirstTape = 1
			
			'Assumes no holes in data
			If g_nCir(2) <= 0 Then
				g_iNumTotalTapes = 1
				g_iNumTapesWristToEOS = 1
				g_iLastTape = 1
			ElseIf g_nCir(3) > 0 Then 
				g_iNumTotalTapes = 3
				g_iNumTapesWristToEOS = 3
				g_iLastTape = 3
				g_ExtendTo = GLOVE_PASTWRIST
			Else
				g_iNumTotalTapes = 2
				g_iNumTapesWristToEOS = 2
				g_iLastTape = 2
				g_ExtendTo = GLOVE_PASTWRIST
			End If
			'        g_iWristPointer = 1
			g_iWristPointer = 1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("SSTab1"), Object).Tab <> 1 Then CType(MainForm.Controls("SSTab1"), Object).Tab = 1
			
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
			
			
		End If
		
		
		
	End Sub
	
	Private Sub PR_GetExtensionDDE_Data()
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
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True 'Normal Glove
		CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0 '1st Tape
		CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0 'Last Tape
		Flap = False
		
		
		'Set default pressure w.r.t diagnosis
		nAge = Val(CType(MainForm.Controls("txtAge"), Object).Text)
		
		
		'Glove extensions
		'Code for extended glove
		'Glove Option Buttons
		Dim nLen As Double
		
		If CType(MainForm.Controls("txtDataGlove"), Object).Text <> "" Then
			For ii = 0 To 5
				nValue = Val(Mid(CType(MainForm.Controls("txtDataGlove"), Object).Text, (ii * 2) + 1, 2))
				If ii = 0 Then
					'Fold options
					'Do nothing Off fold is default
				ElseIf ii = 1 And nValue >= 0 Then 
					'Pressure w.r.t Figuring and MMs
					'ignore
				ElseIf ii = 2 Then 
					'Wrist Tape
					CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = nValue
				ElseIf ii = 3 Then 
					'Proximal Tape
					CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = nValue
				ElseIf ii = 4 Then 
					'Ignore
				ElseIf ii = 5 Then 
					'Ignore flaps
				End If
			Next ii
		End If
		
		'Glove to elbow and Glove to axilla
		'These can start at -3 (However to allow for possible
		'changes later we save up to -4-1/2 but ignore -4-1/2 for the
		'mean time)
		g_TapesFound = False
		If CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text <> "" Then
			ii = 1
			For nn = 8 To 23
				nValue = Val(Mid(CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text, (ii * 3) + 1, 3))
				If nValue > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!txtExtCir(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CType(MainForm.Controls("txtExtCir"), Object)(nn) = nValue / 10
					nLen = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtExtCir"), Object)(nn))
					If nLen <> -1 Then
						CADGLEXT.PR_GrdInchesDisplay(nn - 8, nLen)
						g_TapesFound = True
					End If
					g_nCir(ii) = nLen
				Else
					CADGLEXT.PR_GrdInchesDisplay(nn - 8, 0)
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
	End Sub
	
	Private Sub PR_GrdInchesDisplay(ByRef iIndex As Short, ByRef nLen As Double)
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdInches"), Object).Row = iIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("grdInches"), Object).Col = 0
		CType(MainForm.Controls("grdInches"), Object).Text = ARMEDDIA1.fnInchestoText(nLen)
	End Sub
	
	Private Sub PR_UpdateDBFields()
		
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
			'Update DB fields
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		End If
		
		
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
			'Update DB fields
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeLengths" & QCQ & txtTapeLengths.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "TapeLengths2" & QCQ & txtTapeLengths2.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "Sleeve" & QCQ & txtSide.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hSym" & CQ & "Data" & QCQ & txtDataGlove.Text & QQ & ");")
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
		sPacked = ""
		For ii = 0 To 4
			sPacked = sPacked & "0"
		Next ii
		
		'CheckBoxes
		sPacked = sPacked & "0"
		sPacked = sPacked & "0"
		
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
		sLen = New String(" ", 3)
		sLen = RSet(Trim(Str(g_iInsertSize)), Len(sLen))
		sPacked = sPacked & sLen
		
		'Reinforced dorsal
		sPacked = sPacked & "0"
		
		'Thumb Options
		sPacked = sPacked & "0"
		
		'Insert Options
		sPacked = sPacked & "0"
		
		'Store to DDE text boxes
		txtTapeLengths2.Text = sPacked
		
		
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
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtCir(Index))
		
		'Display value in inches if Valid
		If nLen <> -1 Then lblCir(Index).Text = ARMEDDIA1.fnInchestoText(nLen)
		
	End Sub
	
	Private Sub txtLen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Enter
		Dim Index As Short = txtLen.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtLen(Index))
	End Sub
	
	Private Sub txtLen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Leave
		Dim Index As Short = txtLen.GetIndex(eventSender)
		Dim nLen As Double
		Dim nTipCutBack As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtLen(Index))
		If nLen < 0 Then
			lblLen(Index).Text = ""
			Exit Sub
		End If
		
		
		'Display length text
		If nLen > 0 Then
			lblLen(Index).Text = ARMEDDIA1.fnInchestoText(nLen)
		Else
			lblLen(Index).Text = ""
		End If
		
	End Sub
End Class
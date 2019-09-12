Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class lgedtdia
	Inherits System.Windows.Forms.Form
    'Project:   LGEDTDIA.D
    'Purpose:   Editor for Waist Height with two legs.
    '           For legs drawn with JOBSTEX material and
    '           gradient pressure
    '
    'Version:   1.01
    'Date:      13.Jan.94
    'Author:    Gary George
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    'Jan 99     GG      Ported to VB5
    '
    'Note:-
    '
    '    Much of the code and form has been hacked from
    '    WHFIGURE.FRM and WHLEGDIA.FRM
    '    As both of these are proven.
    '
    '    Reference to the left leg should be taken to mean
    '    the current leg
    '

    '    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
    '    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)

    Dim g_sTextList As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
    'Globals
    Public g_nUnitsFac As Double
    Public g_iLtAnkle As Short
    Public g_JOBSTEX_FL As Short

    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote

    Public g_ReDrawn As Short
    Public g_sCurrentLayer As String
    Public g_sPathJOBST As String
    Public g_iStyleFirstTape As Short
    Public g_iStyleLastTape As Short
    Public g_iFirstTape As Short
    Public g_iLastTape As Short
    Public g_nFrontStrapLength As Double
    Public g_nGauntletExtension As Double
    Public g_nLtLastHeel As Double
    Public g_nLtLastAnkle As Double
    Public g_iLtLastMM As Double
    Public g_iLtLastStretch As Short
    Public g_iLtLastZipper As Short
    Public g_iLtLastFabric As Short

    Public g_iLtStretch(29) As Short
    Public g_iLtRed(29) As Short
    Public g_iLtMM(29) As Short
    Public g_nLtLengths(29) As Double

    Public g_iLtStretchInit(29) As Short
    Public g_iLtRedInit(29) As Short
    Public g_iLtMMInit(29) As Short
    Public g_nLtLengthsInit(29) As Double
    Public g_iLtChanged(29) As Short
    'Profile Globals
    ' Public xyProfile(200) As xy
    'Public xyOriginal(200) As xy
    Public g_iProfile As Short
    'Public xyOtemplate As xy
    Public g_iFirstEditableVertex As Short
    Public g_iFirstEditTape As Short
    Public g_iLastEditTape As Short
    Public g_ShortArm As Short
    Public g_iElbowTape As Short
    Public g_NoElbowTape As Short
    Public g_sOriginalContracture As String
    Public g_sOriginalLining As String

    'XY data type to represent points
    Structure XY
        Dim X As Double
        Dim y As Double
    End Structure

    'Public Structure curve
    '    Dim n As Short
    '    <VBFixedArray(100)> Dim X() As Double
    '    <VBFixedArray(100)> Dim y() As Double

    '    'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
    '    Public Sub Initialize()
    '        'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    '        ReDim X(100)
    '        'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    '        ReDim y(100)
    '    End Sub
    'End Structure

    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
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
        Return
    End Sub
	
	Private Sub Draw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Draw.Click
		'Check that data is all present and insert into drafix
		If Validate_Data() Then
			PR_DrawProfileEdits()
			AppActivate(fnGetDrafixWindowTitleText())
			g_ReDrawn = True
			System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
		End If
	End Sub
	
	Private Sub Finished_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Finished.Click
		'Check that data is all present and commit into drafix
		If Validate_Data() Then
			PR_DrawAndCommitProfileEdits()
			AppActivate(fnGetDrafixWindowTitleText())
			System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
            Return
        End If
		
	End Sub
	
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
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Leg Editing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		
	End Function
	
	Private Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Dim ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As String
		Dim sEscapedString As String
		
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
		
	End Function
	
	
	Private Function fnGetNumber(ByVal sString As String, ByRef iIndex As Short) As Double
		'Function to return as a numerical value the iIndexth item in a string
		'that uses blanks (spaces) as delimiters.
		'EG
		'    sString = "12.3 65.1 45"
		'    fnGetNumber( sString, 2) = 65.1
		'
		'If the iIndexth item is not found then return -1 to indicate an error.
		'This assumes that the string will not be used to store -ve numbers.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
		If Len(sString) = 0 Then
			fnGetNumber = -1
			Exit Function
		End If
		
		'Prepare string
		sString = sString & " " 'Trailing blank as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iIndex
			iPos = InStr(sString, " ")
			If ii = iIndex Then
				sString = VB.Left(sString, iPos - 1)
				fnGetNumber = Val(sString)
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos))
				If Len(sString) = 0 Then
					fnGetNumber = -1
					Exit Function
				End If
			End If
		Next ii
		
		'The function should have exited befor this, however just in case
		'(iIndex = 0) we indicate an error,
		fnGetNumber = -1
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Dim nAge, ii, nn, iFabricClass As Short
		Dim iZipper, iMMHg, iValue As Short
		Dim iStrValue, iMMValue, iRedValue As Short
		Dim nValue, nStretch As Double
		Dim sMessage As String
		Dim iLastStyleFound As Short
		Dim sSpacer As Object
		
		'Disable TIMEOUT
		Timer1.Enabled = False
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		
		'Leg Option buttons
		If txtLeg.Text = "Left" Then
			Me.Text = "LEG Edit - Left [" & txtFileNo.Text & "]"
		End If
		If txtLeg.Text = "Right" Then
			Me.Text = "LEG Edit - Right [" & txtFileNo.Text & "]"
		End If
		
		'From the AnkleString extract the saved ankle figuring
		g_iLtAnkle = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1) - 1
		g_iLtLastMM = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 2)
		g_iLtLastStretch = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 3)
		g_iLtLastZipper = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 7)
		
		'Fabric
		g_iLtLastFabric = Val(VB.Left(txtFabric.Text, 2))
		
		g_iLtMM(g_iLtAnkle) = g_iLtLastMM
		g_iLtRed(g_iLtAnkle) = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4)
		g_iLtStretch(g_iLtAnkle) = g_iLtLastStretch
		
		'Enable Display of the extracted values
		'Set display value depending on Ankle position
		If g_iLtAnkle = 7 Then
			txtLeftMM(6).Visible = False
			txtLeftStr(6).Visible = False
		End If
		
		
		'Display ankle values
		txtLeftMM(g_iLtAnkle).Text = CStr(g_iLtLastMM)
		txtLeftMM(g_iLtAnkle).Enabled = False
		txtLeftMM(g_iLtAnkle).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
		
		txtLeftStr(g_iLtAnkle).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
		txtLeftStr(g_iLtAnkle).Text = CStr(g_iLtStretch(g_iLtAnkle))
		txtLeftStr(g_iLtAnkle).Enabled = False
		grdLeftDisplay.Col = 0
		grdLeftDisplay.Row = g_iLtAnkle - 6 'Display from ankle
		grdLeftDisplay.Text = Str(g_iLtRed(g_iLtAnkle))
		
		'Update tape boxes
		'Get first and last tapes in the text boxes
		g_iFirstTape = -1
		g_iLastTape = -1
		
		nValue = 0
		grdLeftInches.Col = 0
		For ii = 5 To 29
			nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				txtLeft(ii).Text = CStr(nValue)
				nValue = ARMEDDIA1.fnDisplaytoInches(nValue)
				grdLeftInches.Row = ii
				grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
			End If
			g_nLtLengths(ii) = nValue
			If ii <= g_iLtAnkle Then txtLeft(ii).Enabled = False
			If ii = 5 Then g_nLtLastHeel = nValue
			If ii = g_iLtAnkle Then g_nLtLastAnkle = nValue
			If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii
			If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii - 1
		Next ii
		If nValue > 0 Then g_iLastTape = 29
		
		g_JOBSTEX_FL = True
		
		'Initialise display from stored data
		For ii = g_iLtAnkle + 1 To 29
			iMMValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
			iStrValue = Val(Mid(txtLeftStretch.Text, (ii * 3) + 1, 3))
			iRedValue = Val(Mid(txtLeftRed.Text, (ii * 3) + 1, 3))
			If iMMValue > 0 Then
				txtLeftMM(ii).Text = CStr(iMMValue)
				g_iLtMM(ii) = iMMValue
				txtLeftStr(ii).Text = CStr(iStrValue)
				g_iLtStretch(ii) = iStrValue
				grdLeftDisplay.Row = ii - 6
				grdLeftDisplay.Text = Str(iRedValue)
				g_iLtRed(ii) = iRedValue
			End If
		Next ii
		
		'Store inital values for comparision and for use in cancel command
		For ii = 0 To 29
			g_iLtStretchInit(ii) = g_iLtStretch(ii)
			g_iLtRedInit(ii) = g_iLtRed(ii)
			g_iLtMMInit(ii) = g_iLtMM(ii)
			g_nLtLengthsInit(ii) = g_nLtLengths(ii)
			g_iLtChanged(ii) = 0
		Next ii
		
		g_ReDrawn = False
		
		PR_GetProfileFromFile("C:\JOBST\LEGCURVE.DAT")
		
		'Disable Tape above fold Ht
		For ii = g_iLtAnkle + (g_iLastEditTape - g_iFirstEditTape) + 2 To 29
			txtLeft(ii).Enabled = False
			txtLeftMM(ii).Enabled = False
			txtLeftMM(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			txtLeftStr(ii).Enabled = False
			txtLeftStr(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
		Next ii
		
		Show()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
		Cancel.Focus()
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
            Return
        End If
	End Sub
	
	Private Sub lgedtdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		Hide()
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left_Renamed.Text = CStr((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		MainForm = Me
		
		g_nUnitsFac = 1 'Default to inches
		g_JOBSTEX_FL = True
		g_sTextList = "-7½ -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
		
		'Setup display inches grid
		grdLeftInches.set_ColWidth(0, 700)
		grdLeftInches.set_ColAlignment(0, 2)
		For ii = 0 To 29
			grdLeftInches.set_RowHeight(ii, 266)
		Next ii
		
		'Setup display of results grid
		grdLeftDisplay.set_ColWidth(0, 488)
		grdLeftDisplay.set_ColAlignment(0, 2)
		For ii = 0 To 23
			grdLeftDisplay.set_RowHeight(ii, 266)
		Next ii
		
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("The Edit Module is already running!" & Chr(13) & "Use ALT-TAB to access it .", 16, "Leg Edit Warning")
            Return
        End If
		
		g_sPathJOBST = fnPathJOBST()
		
		'Set up to \timeout after 6 seconds
		'This is disabled in link close
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
	End Sub
	
	Private Sub PR_CancelProfileEdits()
		
		If g_ReDrawn = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
			BDYUTILS.PR_PutLine("HANDLE hTempCurv;")
			
			'Delete Curve Copy
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "DeleteEntity(hTempCurv);")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
		End If
		
	End Sub
	
	Private Sub PR_DeleteTapeLabel(ByRef xyPt As xy)
		'Deletes the tape label and the text at the given point
		'Uses a marquee to do so
		Dim y1, x1, x2, y2 As Double
		
		x1 = xyPt.X - 0.5
		x2 = xyPt.X + 0.5
		y1 = xyPt.y - 1.5
		y2 = xyPt.y + 0.5
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hChan=Open(" & QQ & "selection" & QCQ & " ( type = 'Text' OR type = 'Symbol' ) AND layer = 'Construct' AND TOTally INside " & x1 & y1 & x2 & y2 & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if(hChan)")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
		
	End Sub
	
	Private Sub PR_DrawAndCommitProfileEdits()
		
		'Create a macro to commit the changes to the Original profile
		
		Dim xyPt As xy
		Dim iTempLegVertex, ii, iProfileVertex As Short
		Dim iStart, iEnd As Short
		Dim ProfileChanged As Short
		Dim nSeam As Double
		Dim sStr, sLen, sMM, sRed As String
		Dim sLeftStr, sLeftLengths, sLeftMMs, sLeftRed As String
		Dim nLen As Double
		Dim iStr, iMM, iRed As Short
		Dim sLine As String
		
		nSeam = 0.1875
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
		BDYUTILS.PR_PutLine("HANDLE  hCurv, hTempCurv, hLeg, hChan, hEnt;")
		
		sLine = ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST) & "\\JOBST.SLB"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetSymbolLibrary(" & QQ & sLine & QQ & ");")
		ARMDIA1.PR_SetLayer("Construct")
		
		iStart = g_iFirstEditTape - 3
		If iStart > g_iFirstEditTape Then iStart = g_iFirstEditTape
		
		iEnd = g_iLastEditTape + 3
		If iEnd > g_iLegProfile Then iEnd = g_iLegProfile
		
		'Delete Curve Copy (Only if redraw has been used)
		If g_ReDrawn = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hTempCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "DeleteEntity(hTempCurv);")
		End If
		
		'Modify Original curve  (need not have been redrawn first)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hCurv = UID (" & QQ & "find" & QC & Val(txtUIDCurv.Text) & ");")
		ProfileChanged = False
		For ii = g_iLtAnkle + 1 To g_iLastTape
			If g_iLtChanged(ii) < 0 Or g_iLtChanged(ii) > 0 Then
				iProfileVertex = ((ii - g_iLtAnkle) + g_iFirstEditTape)
				xyPt.X = xyLegProfile(iProfileVertex - 1).X
				xyPt.y = ((g_nLtLengths(ii) * (System.Math.Abs(g_iLtRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "SetVertex(hCurv," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
				PR_DeleteTapeLabel(xyLegProfile(iProfileVertex - 1))
				PR_DrawTapeLabel(xyPt, ii)
				ProfileChanged = True
			End If
		Next ii
		
		'Update Leg Box (only if Profile has been changed)
		If ProfileChanged = True Then
			sLeftLengths = VB.Left(txtLeftLengths.Text, 20)
			For ii = 5 To 29
				nLen = Val(txtLeft(ii).Text)
				
				If nLen <> 0 Then
					nLen = nLen * 10 'Shift decimal place to right
					sLen = New String(" ", 4)
					sLen = RSet(Trim(Str(nLen)), Len(sLen))
				Else
					sLen = New String(" ", 4)
				End If
				sLeftLengths = sLeftLengths & sLen
				
			Next ii
			
			For ii = 0 To 29
				iMM = g_iLtMM(ii)
				iRed = g_iLtRed(ii)
				iStr = g_iLtStretch(ii)
				
				If iMM <> 0 Then
					sMM = New String(" ", 3)
					sMM = RSet(Trim(Str(iMM)), Len(sMM))
					sRed = New String(" ", 3)
					sRed = RSet(Trim(Str(iRed)), Len(sRed))
					sStr = New String(" ", 3)
					sStr = RSet(Trim(Str(iStr)), Len(sStr))
				Else
					sMM = New String(" ", 3)
					sRed = New String(" ", 3)
					sStr = New String(" ", 3)
				End If
				
				sLeftMMs = sLeftMMs & sMM
				sLeftRed = sLeftRed & sRed
				sLeftStr = sLeftStr & sStr
				
			Next ii
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hLeg = UID (" & QQ & "find" & QC & Val(txtUidLeftLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (!hLeg)Exit(%cancel," & QQ & "Can't find LEGBOX to Update" & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "TapeLengthsPt1" & QCQ & Mid(sLeftLengths, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "TapeLengthsPt2" & QCQ & Mid(sLeftLengths, 61, 60) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "TapeMMs" & QCQ & Mid(sLeftMMs, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "TapeMMs2" & QCQ & Mid(sLeftMMs, 61, 60) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "Grams" & QCQ & Mid(sLeftStr, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "Grams2" & QCQ & Mid(sLeftStr, 61, 60) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "Reduction" & QCQ & Mid(sLeftRed, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData( hLeg, " & QQ & "Reduction2" & QCQ & Mid(sLeftRed, 61, 60) & QQ & ");")
			
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "ViewRedraw" & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_DrawProfileEdits()
		'Create a macro to copy part of the leg profile and
		'modify this copy
		'Modifications will only be committed to the original profile on "Finish"
		
		Dim xyPt As xy
		Dim iTempLegVertex, ii, iProfileVertex As Short
		Dim iStart, iEnd As Short
		Dim nSeam As Double
		
		nSeam = 0.1875
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
		BDYUTILS.PR_PutLine("HANDLE  hCurv;")
		
		iStart = g_iFirstEditTape - 3
		If iStart > g_iFirstEditTape Then iStart = g_iFirstEditTape
		
		iEnd = g_iLastEditTape + 3
		If iEnd > g_iLegProfile Then iEnd = g_iLegProfile
		
		If g_ReDrawn <> True Then
			'Make Copy of curve
			ARMDIA1.PR_SetLayer("Construct")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hCurv = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
			
			For ii = iStart To iEnd
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Print(fNum, CC & xyLegProfile(ii).X & CC & xyLegProfile(ii).y)
			Next ii
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, ");")
			
		Else
			'Modify Curve Copy
			If txtUIDTempLeg.Text <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hCurv = UID (" & QQ & "find" & QC & Val(txtUIDTempLeg.Text) & ");")
			Else
				
			End If
		End If
		
		For ii = g_iLtAnkle + 1 To g_iLastTape
			If g_iLtChanged(ii) > 0 Then
				g_iLtChanged(ii) = -1
				iTempLegVertex = (ii - g_iLtAnkle) + (g_iFirstEditTape - iStart)
				iProfileVertex = ((ii - g_iLtAnkle) + g_iFirstEditTape) - 1
				If iTempLegVertex > iEnd Then Exit For
				xyPt.X = xyLegProfile(iProfileVertex).X
				xyPt.y = ((g_nLtLengths(ii) * (System.Math.Abs(g_iLtRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "SetVertex(hCurv," & iTempLegVertex & "," & xyPt.X & "," & xyPt.y & ");")
			End If
		Next ii
		
		If g_ReDrawn <> True Then
			BDYUTILS.PR_PutLine("HANDLE  hDDE;")
			
			'Poke Curve UID back to the VB program
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hDDE = Open (" & QQ & "dde" & QCQ & "lgedtdia" & QCQ & "lgedtdia" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Poke ( hDDE, " & QQ & "txtUIDTempLeg" & QC & "MakeString(" & QQ & "long" & QC & "UID(" & QQ & "get" & QQ & ",hCurv)));")
			
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "ViewRedraw" & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_DrawTapeLabel(ByRef xyPt As xy, ByRef Index As Short)
		Dim sSymbol As String
		Dim sText As Object
		
		sSymbol = Trim(Str(Index)) & "tape"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & xyPt.X & CC & xyPt.y & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & "2);")
		'UPGRADE_WARNING: Couldn't resolve default property of object sText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = ARMEDDIA1.fnInchesToText(g_nLtLengths(Index))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & xyPt.X & CC & xyPt.y - 0.5 & CC & "0.075" & CC & "0.125" & ",0);")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = Trim(Str(g_iLtRed(Index)))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & xyPt.X & CC & xyPt.y - 0.7 & CC & "0.075" & CC & "0.125" & ",0);")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = Trim(Str(g_iLtStretch(Index)))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & xyPt.X & CC & xyPt.y - 0.9 & CC & "0.075" & CC & "0.125" & ",0);")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = Trim(Str(g_iLtMM(Index)))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & xyPt.X & CC & xyPt.y - 1.1 & CC & "0.075" & CC & "0.125" & ",0);")
		
		
	End Sub
	
	Private Sub PR_FigureLeftTape(ByRef Index As Short, ByRef Force As Short)
		'Figures the Pressure/Stretch at a LEFT LegTape
		
		Dim nMMHg, nStretch As Double
		Dim nTapeCir As Double
		Dim UseMMHg As Short
		
		grdLeftDisplay.Col = 0
		grdLeftDisplay.Row = Index - 6 'Display from ankle
		
		'Establish if enough data exists to calculate.
		'If not then EXIT quietly
		If Val(txtLeftMM(Index).Text) <= 0 And Val(txtLeftStr(Index).Text) <= 0 Then
			Beep()
			grdLeftDisplay.Text = ""
			'Select_Text_in_Box txtLeftMM(Index)
			Exit Sub
		End If
		
		'Establish if pressure or stretch has changed from last time
		'If it hasn't then exit
		'
		If Val(txtLeftMM(Index).Text) = g_iLtMM(Index) And Val(txtLeftStr(Index).Text) = g_iLtStretch(Index) And Force = False Then Exit Sub
		
		'MMHg and Stretch for JOBSTEX fabric
		'Don't Calculate if no values for Ankle tape
		If Val(txtLeftMM(g_iLtAnkle).Text) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
		
		nTapeCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(Index).Text))
		nMMHg = Val(txtLeftMM(Index).Text)
		nStretch = Val(txtLeftStr(Index).Text)
		
		'Establish if caclculation is based on given Stretch or Pressure (MMhg)
		UseMMHg = -1
		
		'Simple case where one of the values is missing
		If nMMHg <= 0 And nStretch > 0 Then UseMMHg = 0
		If nStretch <= 0 And nMMHg > 0 Then UseMMHg = 1
		
		If nStretch > 0 And nMMHg > 0 Then
			If nMMHg <> g_iLtMM(Index) And nStretch <> g_iLtStretch(Index) Then
				'Changes in Pressure take precedence
				UseMMHg = 1
			Else
				If nMMHg <> g_iLtMM(Index) Then
					UseMMHg = 1
				Else
					UseMMHg = 0
				End If
			End If
		End If
		
		'Exit if UseMMHg flag not set
		If UseMMHg < 0 Then Exit Sub
		
		'Calculate using Pressure solving for stretch
		If UseMMHg = 1 Then
			nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, g_iLtLastZipper, g_iLtLastFabric)
			If nStretch <= 0 Then
				Beep()
				grdLeftDisplay.Text = ""
				txtLeftStr(Index).Text = ""
				ARMEDDIA1.Select_Text_in_Box(txtLeftMM(Index))
				Exit Sub
			End If
		Else
			nMMHg = FN_JOBSTEX_Pressure(nTapeCir, g_nLtLastHeel, nStretch, g_iLtLastZipper, g_iLtLastFabric)
			If nMMHg <= 0 Then
				Beep()
				grdLeftDisplay.Text = ""
				txtLeftMM(Index).Text = ""
				ARMEDDIA1.Select_Text_in_Box(txtLeftStr(Index))
				Exit Sub
			End If
		End If
		
		'Store values
		g_iLtStretch(Index) = ARMDIA1.round(nStretch)
		g_iLtRed(Index) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
		g_iLtMM(Index) = nMMHg
		
		'Display values
		txtLeftMM(Index).Text = CStr(g_iLtMM(Index))
		txtLeftStr(Index).Text = CStr(g_iLtStretch(Index))
		grdLeftDisplay.Text = Str(g_iLtRed(Index))
		
		'Flag Change position
		g_iLtChanged(Index) = 1
		
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
		PrintLine(fNum, "if ( hLayer > %zero && hLayer != 32768)" & "Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "hLayer);")
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'This event means:-
        'That the DDE link to Drafix has not been established
        'therefore we End.
        Return
    End Sub
	
	Private Sub txtLeft_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Enter
		Dim Index As Short = txtLeft.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_in_Box(txtLeft(Index))
	End Sub
	
	Private Sub txtLeft_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Leave
		Dim Index As Short = txtLeft.GetIndex(eventSender)
		BDLEGDIA1.Validate_And_Display_Text_In_Box(txtLeft(Index), grdLeftInches, Index)
	End Sub
	
	Private Sub txtLeftMM_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftMM.Enter
		Dim Index As Short = txtLeftMM.GetIndex(eventSender)
		'Select_Text_in_Box txtLeftMM(Index)
	End Sub
	
	Private Sub txtLeftMM_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftMM.Leave
		Dim Index As Short = txtLeftMM.GetIndex(eventSender)
		If Index <> g_iLtAnkle Then
			PR_FigureLeftTape(Index, 0)
		End If
	End Sub
	
	Private Sub txtLeftStr_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftStr.Enter
		Dim Index As Short = txtLeftStr.GetIndex(eventSender)
		'Select_Text_in_Box txtLeftStr(Index)
	End Sub
	
	Private Sub txtLeftStr_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftStr.Leave
		Dim Index As Short = txtLeftStr.GetIndex(eventSender)
		If Index <> g_iLtAnkle Then
			PR_FigureLeftTape(Index, 0)
		End If
	End Sub
	
	Private Sub Validate_And_Display_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control, ByRef Display_Control As System.Windows.Forms.Control, ByRef Index As Short)
		'Subroutine that is activated when the focus is lost.
		'Displays value in inches in the grid given
		'
		'If Value is changed then
		'    Modifies Stored Value
		'    Recalculates tape reduction and stretch
		'
		Dim rTextBoxValue As Double
		Dim iFirstTape, iLastTape As Short
		Dim ii, iTest As Short
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = ARMEDDIA1.fnDisplaytoInches(Val(Text_Box_Name.Text))
		
		'Only relevant if units are Inches
		If rTextBoxValue < 0 Then
			MsgBox("Invalid Format for inches", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		'Check that each character is numeric or a decimal point
		'N.B.
		'    Asc("0") = 48
		'    Asc("9") = 57
		'    Asc(".") = 46
		'
		For ii = 1 To Len(Text_Box_Name)
			iTest = Asc(Mid(Text_Box_Name.Text, ii, 1))
			If (iTest < 48 Or iTest > 57) And iTest <> 46 Then
				rTextBoxValue = -1
			End If
		Next ii
		
		If rTextBoxValue < 0 Then
			MsgBox("Invalid or Negative value given", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
			MsgBox("Zero value given", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		If rTextBoxValue > 999.9 Then
			MsgBox("Given value too Large", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		If Index >= 0 And g_nLtLengths(Index) <> rTextBoxValue Then
			'Treat Display as Grid control
			'UPGRADE_WARNING: Couldn't resolve default property of object Display_Control.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Display_Control.Row = Index
			Display_Control.Text = ARMEDDIA1.fnInchesToText(rTextBoxValue)
			g_nLtLengths(Index) = rTextBoxValue
			PR_FigureLeftTape(Index, True)
		End If
		
		
	End Sub
	
	Private Function Validate_Data() As Short
		
		Dim sError, NL As String
		Dim nFirstTape, ii, nLastTape As Short
		Dim iError As Short
		Dim nValue As Double
		
		NL = Chr(10) 'new line
		
		'Check Tape length data text boxes for holes (ie missing values)
		'Establish First and last tape
		
		sError = ""
		
		For ii = 5 To 29 Step 1
			If Val(txtLeft(ii).Text) > 0 Then Exit For
		Next ii
		nFirstTape = ii
		
		For ii = 29 To 5 Step -1
			If Val(txtLeft(ii).Text) > 0 Then Exit For
		Next ii
		nLastTape = ii
		
		For ii = nFirstTape To nLastTape
			If Val(txtLeft(ii).Text) = 0 Then
				sError = sError & "Missing Tape length - " & LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)) & NL
			End If
		Next ii
		
		'Check that a minimum of 2 tapes are given
		If (nLastTape - nFirstTape) = 0 Then
			sError = sError & "More than 1 tape must be given." & NL
		End If
		
		If (nLastTape - nFirstTape) < 0 Then
			sError = sError & "No Tape values given." & NL
		End If
		
		'Display Error message (if required) and return
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			Validate_Data = False
			Exit Function
		Else
			Validate_Data = True
		End If
		
		
	End Function
End Class
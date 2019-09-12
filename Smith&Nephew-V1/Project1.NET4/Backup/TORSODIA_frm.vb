Option Strict Off
Option Explicit On
Friend Class torsodia
	Inherits System.Windows.Forms.Form
	'Project:   TORSODIA.MAK
	'Purpose:   TORSO Band Dialogue
	'
	'
	'Version:   3.01
	'Date:      13th Jan 1998
	'Author:    Gary George
	'Copyright  C-Gem Ltd
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
	
	'MsgBox constant
	'    Const IDCANCEL = 2
	'    Const IDYES = 6
	'    Const IDNO = 7
	
	
	
	
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()
		
		If sCurrentValues <> g_sChangeChecker Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
			Select Case Response
				Case IDYES
					PR_CreateMacro_Save("c:\jobst\draw.d")
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(sTask)
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue")
					End If
				Case IDNO
					End
				Case IDCANCEL
					Exit Sub
			End Select
		Else
			End
		End If
	End Sub
	
	Private Function FN_EscapeQuotesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Dim ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As String
		Dim sEscapedString As String
		
		FN_EscapeQuotesInString = ""
		
		For ii = 1 To Len(sAssignedString)
			'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Char_Renamed = Mid(sAssignedString, ii, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Char_Renamed = QQ Then
				sEscapedString = sEscapedString & "\" & Char_Renamed
			Else
				sEscapedString = sEscapedString & Char_Renamed
			End If
		Next ii
		
		FN_EscapeQuotesInString = sEscapedString
		
	End Function
	
	Private Function FN_OpenSave(ByRef sDrafixFile As String, ByRef sType As String, ByRef sName As Object, ByRef sFileNo As Object) As Short
		'Open the DRAFIX macro file
		'Return the file number
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_OpenSave = fNum
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & sName & ", " & sFileNo & "")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, TORSO Band")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//type - " & sType & "")
		
		
	End Function
	
	
	
	Private Function FN_ValidateData() As Short
		'This function is used only to make gross checks
		'for missing data.
		'It does not perform any sensibility checks on the
		'data
		Dim sError As String
		Dim ii, nn As Short
		
		'Initialise
		FN_ValidateData = False
		sError = ""
		
		Dim sCircum(8) As Object
		'    sCircum(0) = "Left shoulder circ."
		'    sCircum(1) = "Right shoulder circ."
		'    sCircum(2) = "Neck circ."
		'    sCircum(3) = "Shoulder width"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(4) = "Chest to waist"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(5) = "Chest circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(6) = "Waist circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(7) = "Chest to EOS"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(8) = "Circ. at EOS"
		
		'Vest measurements
		'    For ii = 0 To 1
		'        If Val(txtCir(ii).Text) = 0 Then
		'            sError = sError & "Missing dimension for " & sCircum(ii) & "!" & NL
		'        End If
		'    Next ii
		
		For ii = 4 To 6
			If Val(txtCir(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing dimension for " & sCircum(ii) & "!" & NL
			End If
		Next ii
		
		'EOS Measurements (if one given both must be given)
		If Val(txtCir(7).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing dimension for " & sCircum(7) & "!" & NL
		End If
		If Val(txtCir(8).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing dimension for " & sCircum(8) & "!" & NL
		End If
		
		If cboClosure.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Closure not given! " & NL
		End If
		
		If cboFabric.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Fabric not given! " & NL
		End If
		
		If sError <> "" Then
			MsgBox(sError, 16, "TORSO Band - Dialogue")
			FN_ValidateData = False
		Else
			FN_ValidateData = True
		End If
		
	End Function
	
	Private Function FN_ValuesString() As String
		'Create a string of all the values given in the
		'Text and Combo boxes.
		'
		'Ignore patient details
		'
		Dim ii As Short
		Dim sString As String
		
		FN_ValuesString = ""
		sString = ""
		
		For ii = 4 To 8
			sString = sString & txtCir(ii).Text
		Next ii
		
		sString = sString & cboClosure.Text
		
		sString = sString & cboFabric.Text
		
		FN_ValuesString = sString
		
	End Function
	
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		Dim ii As Short
		
		'Stop the timer used to ensure that the Dialogue dies
		'if the DRAFIX macro fails to establish a DDE Link
		Timer1.Enabled = False
		
		'Check that a "MainPatientDetails" Symbol has been
		'found
		If txtUidMPD.Text = "" Then
			MsgBox("No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue")
			End
		End If
		
		If txtCombo(9).Text <> "" Then cboClosure.Text = txtCombo(9).Text Else cboClosure.SelectedIndex = 0
		cboFabric.Text = txtCombo(10).Text
		
		'Set up units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Display dimesions sizes in inches
		For ii = 4 To 8
			txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
		Next ii
		
		
		'Save the values in the text to a string
		'this can then be used to check if they have changed
		'on use of the close button
		g_sChangeChecker = FN_ValuesString()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to hourglass.
		
		Show()
	End Sub
	
	Private Sub torsodia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		
		Hide()
		
		'Start a timer
		'The Timer is disabled in LinkClose
		'If after 6 seconds the timer event will "End" the programme
		'This ensures that the dialogue dies in event of a failure
		'on the drafix macro side
		Timer1.Interval = 6000 'Approx 6 Seconds
		Timer1.Enabled = True
		
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("TORSO Band Dialogue is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting VEST Body - Dialogue")
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		MainForm = Me
		
		'Initialize globals
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes (")
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(13) 'New Line
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
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
		
		g_nUnitsFac = 1
		g_PathJOBST = fnPathJOBST()
		
		'Clear fields
		'Circumferences and lengths
		For ii = 4 To 8
			txtCir(ii).Text = ""
		Next ii
		
		'The data from these DDE text boxes is copied
		'to the combo boxes on Link close
		'
		txtCombo(0).Text = ""
		txtCombo(1).Text = ""
		txtCombo(9).Text = ""
		txtCombo(10).Text = ""
		
		'closure
		'    cboClosure.AddItem "Velcro"
		'    cboClosure.AddItem "Zip"
		cboClosure.Items.Add("Front Velcro")
		cboClosure.Items.Add("Front Velcro (Reversed)")
		cboClosure.Items.Add("Back Velcro")
		cboClosure.Items.Add("Back Velcro (Reversed)")
		cboClosure.Items.Add("Front Zip")
		cboClosure.Items.Add("Back Zip")
		
		
		'Patient details
		txtFileNo.Text = ""
		txtUnits.Text = ""
		txtPatientName.Text = ""
		txtDiagnosis.Text = ""
		txtAge.Text = ""
		txtSex.Text = ""
		txtWorkOrder.Text = ""
		
		
		'UID of symbols
		txtUidMPD.Text = ""
		txtUidVB.Text = ""
		
		LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_PathJOBST & "\FABRIC.DAT")
		
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim sTask As String
		'Don't allow multiple clicking
		'
		OK.Enabled = False
		If FN_ValidateData() Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			Hide()
			
			PR_CreateMacro_Draw("c:\jobst\draw.d")
			
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				End
			Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue")
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub PR_CreateMacro_Draw(ByRef sFile As String)
		'Since this is a fairly straight forward module we
		'will keep calculation and drawing in the same procedure
		'
		'The assumption is that all data has been validated before
		'this procedure is called
		
		'Figured dimemsions
		Dim nEOSCir As Double
		Dim nEOStoShoulder As Double 'Misnomer due to mods 04.Jun.98
		
		Dim nWaistCir As Double
		Dim nWaisttoShoulder As Double 'Misnomer due to mods 04.Jun.98
		
		Dim nChestCir As Double
		'    Dim nRightShoulderCir   As Double
		'    Dim nLeftShoulderCir    As Double
		'    Dim nChesttoShoulder    As Double
		
		Dim nLowShoulderLine As Double
		
		'Points
		Dim xyO As xy
		Dim xyEOS As xy
		Dim xyEOSatCL As xy
		Dim xyChest As xy
		Dim xyChestatCL As xy
		Dim xyWaist As xy
		Dim xyWaistatCL As xy
		
		'Figuring factors
		Dim nCirFactor As Double
		'    Dim nShoulderCirFactor  As Double
		Dim nSeamAllowance As Double
		
		'Flags
		Dim bEOSGiven As Short
		
		'other
		Dim sText As String
		Dim sWorkOrder As String
		
		Dim nZiplength As Double
		
		Dim xyText As xy
		
		'Set figuring factors, flags, defaults etc...
		bEOSGiven = False
		nCirFactor = 0.9
		'    nShoulderCirFactor = 2.5
		nSeamAllowance = 0.1875
		ARMDIA1.PR_MakeXY(xyO, 0, 0)
		
		'Figure given dimensions
		'N.B.
		'txtCir(0)  = Left Shoulder cir.     (not required)
		'txtCir(1)  = Right Shoulder cir.    (not required)
		'txtCir(2)  = Neck cir.              (not required)
		'txtCir(3)  = Shoulder width         (not required)
		'txtCir(4)  = Chest to waist
		'txtCir(5)  = Chest cir.
		'txtCir(6)  = Waist cir.
		'txtCir(7)  = Chest to EOS           (optional, but must be given if EOS Cir. below given
		'txtCir(8)  = EOS cir.               (optional)
		
		'EOS (This is optional)
		If Val(txtCir(8).Text) > 0 Then
			bEOSGiven = True
			nEOSCir = ARMEDDIA1.fnDisplayToInches(CDbl(txtCir(8).Text)) * nCirFactor / 4
			nEOStoShoulder = ARMEDDIA1.fnDisplayToInches(CDbl(txtCir(7).Text))
		End If
		
		'Waist
		nWaistCir = ARMEDDIA1.fnDisplayToInches(CDbl(txtCir(6).Text)) * nCirFactor / 4
		nWaisttoShoulder = ARMEDDIA1.fnDisplayToInches(CDbl(txtCir(4).Text))
		
		'Chest
		nChestCir = ARMEDDIA1.fnDisplayToInches(CDbl(txtCir(5).Text)) * nCirFactor / 4
		
		'    nRightShoulderCir = fnDisplayToInches(txtCir(1))
		'    nLeftShoulderCir = fnDisplayToInches(txtCir(0))
		'
		'    If Abs(nRightShoulderCir - nLeftShoulderCir) > 1 Then
		'       'use the highest
		'        nRightShoulderCir = round(nRightShoulderCir / nShoulderCirFactor)
		'        nLeftShoulderCir = round(nLeftShoulderCir / nShoulderCirFactor)
		'        nChesttoShoulder = min(nRightShoulderCir, nLeftShoulderCir)
		'    Else
		'       'use the average
		'        nChesttoShoulder = round(((nRightShoulderCir + nLeftShoulderCir) / 2) / nShoulderCirFactor)
		'    End If
		'   'Drop "nChestToShoulder" by 1/2" to match changes requested to the vest
		'   'on 16.Oct.97
		'    nChesttoShoulder = nChesttoShoulder - .5
		
		'High shoulder line
		If bEOSGiven Then
			nLowShoulderLine = nEOStoShoulder
		Else
			nLowShoulderLine = nWaisttoShoulder
		End If
		
		'Key Points
		'End of Support (If given) and Waist points.
		If bEOSGiven Then
			xyEOSatCL.X = xyO.X
			xyEOSatCL.Y = xyO.Y
			xyEOS.X = xyO.X
			xyEOS.Y = nEOSCir + nSeamAllowance + xyO.Y
			xyWaist.X = nLowShoulderLine - nWaisttoShoulder + xyO.X
			xyWaist.Y = nWaistCir + nSeamAllowance + xyO.Y
			xyWaistatCL.X = xyWaist.X
			xyWaistatCL.Y = xyO.Y
		Else
			xyWaistatCL.X = xyO.X
			xyWaistatCL.Y = xyO.Y
			xyWaist.X = xyO.X
			xyWaist.Y = nWaistCir + nSeamAllowance + xyO.Y
		End If
		
		'Chest points
		'    xyChestatCL.X = nLowShoulderLine - nChesttoShoulder + xyO.X
		xyChestatCL.X = nLowShoulderLine + xyO.X
		xyChestatCL.Y = xyO.Y
		xyChest.X = xyChestatCL.X
		xyChest.Y = nChestCir + nSeamAllowance + xyO.Y
		
		'Zippers
		If bEOSGiven Then
			nZiplength = ARMDIA1.FN_CalcLength(xyEOSatCL, xyChestatCL)
		Else
			nZiplength = ARMDIA1.FN_CalcLength(xyChestatCL, xyWaistatCL)
		End If
		nZiplength = (nZiplength - 0.125) / 0.95
		
		'DRAW DRAW DRAW DRAW DRAW DRAW DRAW DRAW
		'(Join the dots etc..)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sFile)
		
		'Draw Text items
		ARMDIA1.PR_Setlayer("Notes")
		
		'Zippers
		ARMDIA1.PR_SetTextData(2, 32, -1, -1, -1) 'Horiz-Cen, Vertical-Bottom
		sText = ARMEDDIA1.fnInchesToText(nZiplength) & "\"" " & cboClosure.Text
		BDYUTILS.PR_CalcMidPoint(xyChestatCL, xyWaistatCL, xyText)
		ARMDIA1.PR_MakeXY(xyText, xyText.X, xyText.Y + 0.25)
		ARMDIA1.PR_DrawText(sText, xyText, 0.1)
		
		'Elastic Top
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 270
		BDYUTILS.PR_CalcMidPoint(xyChestatCL, xyChest, xyText)
		ARMDIA1.PR_MakeXY(xyText, xyText.X - 0.25, xyText.Y)
		ARMDIA1.PR_DrawText("TOP ELASTIC", xyText, 0.1)
		
		'Elastic Bottom
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 90
		If bEOSGiven Then
			BDYUTILS.PR_CalcMidPoint(xyEOSatCL, xyEOS, xyText)
		Else
			BDYUTILS.PR_CalcMidPoint(xyWaistatCL, xyWaist, xyText)
		End If
		ARMDIA1.PR_MakeXY(xyText, xyText.X + 0.25, xyText.Y)
		ARMDIA1.PR_DrawText("BOTTOM ELASTIC", xyText, 0.1)
		
		'Patient details
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
		BDYUTILS.PR_CalcMidPoint(xyChest, xyWaistatCL, xyText)
		If txtWorkOrder.Text = "" Then
			sWorkOrder = "-"
		Else
			sWorkOrder = txtWorkOrder.Text
		End If
		sText = txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim(Mid(cboFabric.Text, 4))
		ARMDIA1.PR_DrawText(sText, xyText, 0.1)
		
		'Remaining patient details in black on layer construct
		ARMDIA1.PR_Setlayer("Construct")
		ARMDIA1.PR_MakeXY(xyText, xyText.X, xyText.Y - 0.8)
		sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
		ARMDIA1.PR_DrawText(sText, xyText, 0.1)
		
		'Draw construction lines
		If bEOSGiven Then ARMDIA1.PR_DrawLine(xyWaistatCL, xyWaist)
		
		'Draw profile
		ARMDIA1.PR_Setlayer("TemplateLeft")
		BDYUTILS.PR_StartPoly()
		
		If bEOSGiven Then
			BDYUTILS.PR_AddVertex(xyEOSatCL, 0)
			BDYUTILS.PR_AddVertex(xyEOS, 0)
			BDYUTILS.PR_AddVertex(xyWaist, 0)
			BDYUTILS.PR_AddVertex(xyChest, 0)
			BDYUTILS.PR_AddVertex(xyChestatCL, 0)
			BDYUTILS.PR_AddVertex(xyEOSatCL, 0)
		Else
			BDYUTILS.PR_AddVertex(xyWaistatCL, 0)
			BDYUTILS.PR_AddVertex(xyWaist, 0)
			BDYUTILS.PR_AddVertex(xyChest, 0)
			BDYUTILS.PR_AddVertex(xyChestatCL, 0)
			BDYUTILS.PR_AddVertex(xyWaistatCL, 0)
		End If
		BDYUTILS.PR_EndPoly()
		
		PR_UpdateDB()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_OpenSave(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hBody;")
		
		PR_UpdateDB()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	
	Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.Control, ByRef sFileName As String)
		'General procedure to create the list section of
		'a combo box reading the data from a file
		
		Dim sLine As String
		Dim fFileNum As Short
		
		fFileNum = FreeFile
		
		If FileLen(sFileName) = 0 Then
			MsgBox(sFileName & "Not found", 48, "CAD - Glove Dialogue")
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
	
	Private Sub PR_UpdateDB()
		'Procedure called from
		'    PR_CreateMacro_Save
		'and
		'    PR_CreateMacro_Draw
		'
		'Used to stop duplication on code
		
		Dim sSymbol As String
		
		sSymbol = "vestbody"
		
		If txtUidVB.Text = "" Then
			'Define DB Fields
			BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\VFIELDS.D;")
			
			'Find "mainpatientdetails" and get position
			BDYUTILS.PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
			BDYUTILS.PR_PutLine("STRING sMPD_Name;")
			BDYUTILS.PR_PutLine("ANGLE  aMPD_Angle;")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");")
			BDYUTILS.PR_PutLine("if (hMPD)")
			BDYUTILS.PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
			BDYUTILS.PR_PutLine("else")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")
			
			'Insert bodybox
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  hBody = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);")
			BDYUTILS.PR_PutLine("  }")
			BDYUTILS.PR_PutLine("else")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
		Else
			'Use existing symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hBody = UID (" & QQ & "find" & QC & Val(txtUidVB.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			
		End If
		
		'Update the BODY Box symbol
		'    PR_PutLine "SetDBData( hBody" & CQ & "LtSCir" & QCQ & txtCir(0).Text & QQ & ");"
		'    PR_PutLine "SetDBData( hBody" & CQ & "RtSCir" & QCQ & txtCir(1).Text & QQ & ");"
		'    PR_PutLine "SetDBData( hBody" & CQ & "NeckCir" & QCQ & txtCir(2).Text & QQ & ");"
		'    PR_PutLine "SetDBData( hBody" & CQ & "SWidth" & QCQ & txtCir(3).Text & QQ & ");"
		'    PR_PutLine "SetDBData( hBody" & CQ & "S_Waist" & QCQ & txtCir(4).Text & QQ & ");"
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "SLgButt" & QCQ & txtCir(4).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ChestCir" & QCQ & txtCir(5).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtCir(6).Text & QQ & ");")
		
		'    PR_PutLine "SetDBData( hBody" & CQ & "S_EOS" & QCQ & txtCir(7).Text & QQ & ");"
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "SFButt" & QCQ & txtCir(7).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "EOSCir" & QCQ & txtCir(8).Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Closure" & QCQ & cboClosure.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Fabric" & QCQ & cboFabric.Text & QQ & ");")
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		
	End Sub
	
	Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
		'Allows the user to use enter as a tab
		
		System.Windows.Forms.SendKeys.Send("{TAB}")
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'It is assumed that the link open from Drafix has failed
		'Therefor we "End" here
		End
	End Sub
	
	Private Sub txtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Enter
		Dim Index As Short = txtCir.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtCir(Index))
	End Sub
	
	Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
		Dim Index As Short = txtCir.GetIndex(eventSender)
		Dim nLen As Double
		
		nLen = CADGLOVE1.FN_InchesValue(txtCir(Index))
		If nLen > 0 Then
			lblCir(Index).Text = ARMEDDIA1.fnInchesToText(nLen)
		Else
			lblCir(Index).Text = ""
		End If
		
	End Sub
End Class
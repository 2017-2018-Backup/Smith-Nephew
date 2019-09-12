Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class whboddia
	Inherits System.Windows.Forms.Form
	'   '* Windows API Functions Declarations
	'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	'
	'   'Constants used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4''
	'
	'   'MsgBox constant
	'    Const IDYES = 6
	'    Const IDNO = 7
	'    Const IDCANCEL = 2
	
	
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()
		
		If sCurrentValues <> g_sChangeChecker Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, "WAIST Height Body - Dialogue")
			Select Case Response
				Case IDYES
					PR_CreateMacro_Save("c:\jobst\draw.d")
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(sTask)
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Dialogue")
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
	
	'UPGRADE_WARNING: Event cboCrotchStyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCrotchStyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCrotchStyle.SelectedIndexChanged
		Dim sString, sError As String
		Dim iIndex As Short
		
		sError = ""
		
		'Check Crotch style against sex
		sString = VB6.GetItemString(cboCrotchStyle, cboCrotchStyle.SelectedIndex)
		
		'Check for male crotch and female patient
		iIndex = 0
		iIndex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
		If iIndex <> 0 And txtSex.Text = "Female" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male Crotch type specified for Female patient" + NL
		End If
		
		'Check for female crotch and male patient
		iIndex = 0
		iIndex = InStr(1, sString, "Female", 0)
		If iIndex <> 0 And txtSex.Text = "Male" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Female Crotch type specified for Male patient" + NL
		End If
		
		
		'Display Error message (if required) and return
		'These are fatal errors
		If Len(sError) > 0 Then
			MsgBox(sError, 0, "Crotch Style Error")
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event cboLegStyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLegStyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLegStyle.SelectedIndexChanged
		'
		Dim sLegStyle, sReductionFactor As String
		Dim iAge As Short
		
		If VB.Left(g_sPreviousLegStyle, 4) = "Chap" Then prEnable_WRT_Chap()
		If VB.Left(g_sPreviousLegStyle, 4) = "W4OC" Then prEnable_WRT_W4OC()
		If VB.Left(g_sPreviousLegStyle, 5) = "Panty" Then prRestore_Reductions()
		If VB.Left(g_sPreviousLegStyle, 5) = "Brief" Then prRestore_Reductions()
		
		'Get the currently selected leg stlye
		Select Case VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
			
			Case "Full-Leg"
				prOptLeftLeg(1, 0, 0, 1, 1, 0)
				prOptRightLeg(1, 0, 0, 1, 1, 0)
				'Set dropdown combo boxes for reductions
				'Only set if Leg Style had been changed
				If g_sPreviousLegStyle <> "" Then
					prStore_Reductions()
					'Reductions, set defaults
					sReductionFactor = "0.83"
					iAge = Val(txtAge.Text)
					If iAge <= 3 Then sReductionFactor = "0.88"
					If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
					
					'NB. Special case for TOS as TOS need not be given.
					If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
						prSet_Red_Combo(sReductionFactor, cboTOSRed)
					End If
					prSet_Red_Combo(sReductionFactor, cboWaistRed)
					prSet_Red_Combo(sReductionFactor, cboMidPointRed)
					prSet_Red_Combo(sReductionFactor, cboLargestRed)
					prSet_Red_Combo(sReductionFactor, cboThighRed)
				End If
				
				
			Case "Panty"
				prOptLeftLeg(0, 1, 0, 0, 1, 1)
				prOptRightLeg(0, 1, 0, 0, 1, 1)
				'Set dropdown combo boxes for reductions
				'Only set if Leg Style had been changed
				If g_sPreviousLegStyle <> "" Then
					prStore_Reductions()
					sReductionFactor = "0.85"
					'NB. Special case for TOS as TOS need not be given.
					If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
						prSet_Red_Combo(sReductionFactor, cboTOSRed)
					End If
					prSet_Red_Combo(sReductionFactor, cboWaistRed)
					prSet_Red_Combo(sReductionFactor, cboMidPointRed)
					prSet_Red_Combo(sReductionFactor, cboLargestRed)
					prSet_Red_Combo(sReductionFactor, cboThighRed)
				End If
				
			Case "Chap-Left"
				prDisable_WRT_Chap()
				prOptLeftLeg(1, 0, 0, 1, 1, 0)
				prOptRightLeg(0, 0, 0, 0, 0, 0)
				If g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)
				
			Case "Chap-Right"
				prDisable_WRT_Chap()
				prOptLeftLeg(0, 0, 0, 0, 0, 0)
				prOptRightLeg(1, 0, 0, 1, 1, 0)
				If g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)
				
			Case "Chap-Both Legs"
				prDisable_WRT_Chap()
				prOptLeftLeg(1, 0, 0, 1, 1, 0)
				prOptRightLeg(1, 0, 0, 1, 1, 0)
				If g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)
				
			Case "Brief"
				prOptLeftLeg(0, 0, 1, 0, 0, 1)
				prOptRightLeg(0, 0, 1, 0, 0, 1)
				'Set dropdown combo boxes for reductions
				'Only set if Leg Style had been changed
				If g_sPreviousLegStyle <> "" Then
					prStore_Reductions()
					sReductionFactor = "0.85"
					'NB. Special case for TOS as TOS need not be given.
					If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
						prSet_Red_Combo(sReductionFactor, cboTOSRed)
					End If
					prSet_Red_Combo(sReductionFactor, cboWaistRed)
					prSet_Red_Combo(sReductionFactor, cboMidPointRed)
					prSet_Red_Combo(sReductionFactor, cboLargestRed)
					prSet_Red_Combo(sReductionFactor, cboThighRed)
				End If
				
				
			Case "W4OC-Left"
				prDisable_WRT_W4OC()
				'Enable Left thigh, restore if value available
				prOptLeftLeg(1, 0, 0, 1, 1, 0)
				txtLeftThighCir.Enabled = True
				labLeftThighCir.Enabled = True
				labLeftThigh.Enabled = True
				If Val(txtLeftThighCir.Text) = 0 And g_nLeftThighCir Then txtLeftThighCir.Text = CStr(g_nLeftThighCir)
				txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())
				
				'Disable Right thigh data, save if required for a restore
				If Val(txtRightThighCir.Text) > 0 Then g_nRightThighCir = Val(txtRightThighCir.Text)
				prOptRightLeg(0, 0, 0, 0, 0, 0)
				'        txtRightThighCir = ""
				'        txtRightThighCir_LostFocus
				labRightThigh.Enabled = False
				txtRightThighCir.Enabled = False
				labRightThighCir.Enabled = False
				
			Case "W4OC-Right"
				prDisable_WRT_W4OC()
				'Enable Right thigh, restore if value available
				prOptRightLeg(1, 0, 0, 1, 1, 0)
				txtRightThighCir.Enabled = True
				labRightThighCir.Enabled = True
				labRightThigh.Enabled = True
				If Val(txtRightThighCir.Text) = 0 And g_nRightThighCir Then txtRightThighCir.Text = CStr(g_nRightThighCir)
				txtRightThighCir_Leave(txtRightThighCir, New System.EventArgs())
				
				'Disable Left thigh data, save if required for a restore
				If Val(txtLeftThighCir.Text) > 0 Then g_nLeftThighCir = Val(txtLeftThighCir.Text)
				prOptLeftLeg(0, 0, 0, 0, 0, 0)
				'        txtLeftThighCir = ""
				'        txtLeftThighCir_LostFocus
				labLeftThigh.Enabled = False
				txtLeftThighCir.Enabled = False
				labLeftThighCir.Enabled = False
				
		End Select
		
		'Store leg style for use when it changes
		g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		
	End Sub
	
	'UPGRADE_WARNING: Event chkDrawRevisedHts.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDrawRevisedHts_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDrawRevisedHts.CheckStateChanged
		If chkDrawRevisedHts.CheckState = 1 Then
			If Val(txtTOSHtRevised.Text) > 0 Then
				txtTOSHt.Enabled = False
				labTOSHt.Enabled = False
			End If
			If Val(txtWaistHtRevised.Text) > 0 Then
				txtWaistHt.Enabled = False
				labWaistHt.Enabled = False
			End If
			If Val(txtFoldHtRevised.Text) > 0 Then
				txtFoldHt.Enabled = False
				labFoldHt.Enabled = False
			End If
			txtWaistHtRevised.Enabled = True
			txtTOSHtRevised.Enabled = True
			txtFoldHtRevised.Enabled = True
			labWaistHtRevised.Enabled = True
			labTOSHtRevised.Enabled = True
			labFoldHtRevised.Enabled = True
			labRevisedHts.Enabled = True
		Else
			txtTOSHtRevised.Enabled = False
			txtWaistHtRevised.Enabled = False
			txtFoldHtRevised.Enabled = False
			labTOSHtRevised.Enabled = False
			labWaistHtRevised.Enabled = False
			labFoldHtRevised.Enabled = False
			txtTOSHt.Enabled = True
			txtWaistHt.Enabled = True
			txtFoldHt.Enabled = True
			labTOSHt.Enabled = True
			labWaistHt.Enabled = True
			labFoldHt.Enabled = True
			labRevisedHts.Enabled = False
		End If
	End Sub
	
	Private Sub cmdTab_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTab.Click
		'Allows the user to use enter as a tab
		System.Windows.Forms.SendKeys.Send("{TAB}")
		
	End Sub
	
	Private Function FN_Open(ByRef sDrafixFile As String, ByRef sType As String, ByRef sName As Object, ByRef sFileNo As Object) As Short
		'Open the DRAFIX macro file
		'Return the file number
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_Open = fNum
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & sName & ", " & sFileNo & "")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, Waist Height Body")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//type - " & sType & "")
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		
		'Display Hour Glass symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & sType & QQ & ");")
		
		
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
		
		'Dimensions
		sString = sString & txtTOSCir.Text & txtTOSHt.Text & txtWaistCir.Text & txtWaistHt.Text & txtMidPointCir.Text
		sString = sString & txtMidPointHt.Text & txtLargestCir.Text & txtLargestHt.Text & txtLeftThighCir.Text
		sString = sString & txtRightThighCir.Text & txtFoldHt.Text
		
		sString = sString & cboLegStyle.Text & cboCrotchStyle.Text
		
		sString = sString & chkDrawRevisedHts.CheckState
		
		sString = sString & optLeftLeg(0).Checked & optLeftLeg(1).Checked & optLeftLeg(2).Checked
		
		sString = sString & optRightLeg(0).Checked & optRightLeg(1).Checked & optRightLeg(2).Checked
		
		FN_ValuesString = sString
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		'On link close.
		Dim sReductionFactor As String
		Dim nAge, ii As Short
		Dim nValue As Double
		
		'Disable Time out
		Timer1.Enabled = False
		
		
		
		'Set global Units Factor
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Waist Height details
		'Note
		' The procedure Validate_Text_In_Box also displays the size in
		' inches.
		'
		BDLEGDIA1.Validate_Text_In_Box(txtTOSCir, labTOSCir)
		BDLEGDIA1.Validate_Text_In_Box(txtTOSHt, labTOSHt)
		BDLEGDIA1.Validate_Text_In_Box(txtWaistCir, labWaistCir)
		BDLEGDIA1.Validate_Text_In_Box(txtWaistHt, labWaistHt)
		BDLEGDIA1.Validate_Text_In_Box(txtMidPointCir, labMidPointCir)
		BDLEGDIA1.Validate_Text_In_Box(txtMidPointHt, labMidPointHt)
		BDLEGDIA1.Validate_Text_In_Box(txtLargestCir, labLargestCir)
		BDLEGDIA1.Validate_Text_In_Box(txtLargestHt, labLargestHt)
		BDLEGDIA1.Validate_Text_In_Box(txtLeftThighCir, labLeftThighCir)
		BDLEGDIA1.Validate_Text_In_Box(txtRightThighCir, labRightThighCir)
		BDLEGDIA1.Validate_Text_In_Box(txtFoldHt, labFoldHt)
		
		'Set dropdown combo boxes
		'Crotch Style
		cboCrotchStyle.Items.Add("Open Crotch") 'Both
		cboCrotchStyle.Items.Add("Horizontal Fly") 'Male only
		cboCrotchStyle.Items.Add("Diagonal Fly") 'Male only
		cboCrotchStyle.Items.Add("Gusset") 'Both
		cboCrotchStyle.Items.Add("Mesh Gusset") 'Both
		cboCrotchStyle.Items.Add("Female Mesh Gusset") 'Female only
		cboCrotchStyle.Items.Add("Male Mesh Gusset") 'Male only
		
		
		'Set value
		For ii = 0 To (cboCrotchStyle.Items.Count - 1)
			If VB6.GetItemString(cboCrotchStyle, ii) = txtCrotchStyle.Text Then
				cboCrotchStyle.SelectedIndex = ii
			End If
		Next ii
		
		'Leg Style
		cboLegStyle.Items.Add("Full-Leg")
		cboLegStyle.Items.Add("Panty")
		cboLegStyle.Items.Add("Brief")
		cboLegStyle.Items.Add("W4OC-Left")
		cboLegStyle.Items.Add("W4OC-Right")
		cboLegStyle.Items.Add("Chap-Left")
		cboLegStyle.Items.Add("Chap-Right")
		cboLegStyle.Items.Add("Chap-Both Legs")
		
		
		'Set values of combo and option buttons
		g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
		cboLegStyle.SelectedIndex = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		
		'Set this now to avoid problems with cboLegStyle_Click
		If cboLegStyle.SelectedIndex = -1 Then
			g_sPreviousLegStyle = "None"
		Else
			g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		End If
		
		'   ii = fnGetNumber(txtLegStyle.Text, 2)
		'   If ii >= 0 Then optLeftLeg(ii).Value = True
		'   ii = fnGetNumber(txtLegStyle.Text, 3)
		'   If ii >= 0 Then optRightLeg(ii).Value = True
		
		'Reductions (set defaults if no handle to the body)
		sReductionFactor = ""
		nAge = Val(txtAge.Text)
		If Val(txtUidWHBody.Text) = 0 Then
			sReductionFactor = "0.83"
			If nAge <= 3 Then
				sReductionFactor = "0.88"
			End If
			If nAge <= 10 And nAge > 3 Then
				sReductionFactor = "0.86"
			End If
		End If
		
		'Set dropdown combo boxes for reductions
		'NB. Special case for TOS as TOS need not be given.
		'    Also if this is the first time through then display reduction.
		prInitialise_Red_Combo(sReductionFactor, txtTOSRed, cboTOSRed)
		If (Len(txtTOSCir.Text) = 0 And Len(txtTOSHt.Text) = 0) Or Len(txtUidWHBody.Text) = 0 Then
			cboTOSRed.SelectedIndex = -1
			cboTOSRed.Enabled = False
		End If
		prInitialise_Red_Combo(sReductionFactor, txtWaistRed, cboWaistRed)
		prInitialise_Red_Combo(sReductionFactor, txtMidPointRed, cboMidPointRed)
		prInitialise_Red_Combo(sReductionFactor, txtLargestRed, cboLargestRed)
		prInitialise_Red_Combo(sReductionFactor, txtThighRed, cboThighRed)
		
		ii = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 2)
		If ii >= 0 Then optLeftLeg(ii).Checked = True
		ii = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 3)
		If ii >= 0 Then optRightLeg(ii).Checked = True
		
		
		'Revised Heights
		'From Body DB field
		nValue = ARMEDDIA1.fnGetNumber(txtBody.Text, 5)
		If nValue > 0 Then
			txtTOSHtRevised.Text = CStr(nValue)
			BDLEGDIA1.Validate_Text_In_Box(txtTOSHtRevised, labTOSHtRevised)
		End If
		
		nValue = ARMEDDIA1.fnGetNumber(txtBody.Text, 6)
		If nValue > 0 Then
			txtWaistHtRevised.Text = CStr(nValue)
			BDLEGDIA1.Validate_Text_In_Box(txtWaistHtRevised, labWaistHtRevised)
		End If
		
		nValue = ARMEDDIA1.fnGetNumber(txtBody.Text, 7)
		If nValue > 0 Then
			txtFoldHtRevised.Text = CStr(nValue)
			BDLEGDIA1.Validate_Text_In_Box(txtFoldHtRevised, labFoldHtRevised)
		End If
		
		If ARMEDDIA1.fnGetNumber(txtBody.Text, 4) > 0 Then
			chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		
		'Store loded values w.r.t cancel button
		'
		g_sChangeChecker = FN_ValuesString()
		
		Show()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to default.
		Me.Activate()
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub whboddia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Hide()
		'Check if this is already running
		'If it is then warn the user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			End
		End If
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to Hourglass.
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
		
		g_sPathJOBST = fnPathJOBST()
		
		'Set Units global to default to inches
		g_nUnitsFac = 1
		
		'Disable revised Hts (default), set from Link Close event
		chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
		chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked
		
		'Start a timer that will ensure that the
		'the Vb programme ends if the DRAFIX DDE
		'link fails.
		Timer1.Interval = 6000
		Timer1.Enabled = True
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		'Check that data is all present and insert into drafix
		
		Dim sTask As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Validate_Data() Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			Hide()
			
			PR_CreateMacro_Save("c:\jobst\draw.d")
			
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				End
			Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Data")
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	
	'UPGRADE_WARNING: Event optLeftLeg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optLeftLeg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optLeftLeg.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optLeftLeg.GetIndex(eventSender)
			'0 = Full-Leg
			'1 = Panty
			'2 = Brief
			
			Dim sLegStyle, sReductionFactor As String
			Dim iAge As Short
			
			sLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
			
			If sLegStyle = "Full-Leg" Then
				If Index = 0 Then
					optRightLeg(1).Enabled = True
					optRightLeg(2).Enabled = True
				Else
					optRightLeg(1).Enabled = False
					optRightLeg(2).Enabled = False
				End If
			End If
			
			If sLegStyle = "Panty" Then
				If Index = 1 Then
					optRightLeg(2).Enabled = True
				Else
					optRightLeg(2).Enabled = False
				End If
			End If
			
			If sLegStyle = "W4OC-Left" And g_sPreviousLegStyle <> "" Then
				'Reductions, set defaults
				prStore_Reductions()
				If Index = 0 Then
					'Reductions, for full leg
					sReductionFactor = "0.83"
					iAge = Val(txtAge.Text)
					If iAge <= 3 Then sReductionFactor = "0.88"
					If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
				Else
					'Reductions panty leg
					sReductionFactor = "0.85"
				End If
				
				'NB. Special case for TOS as TOS need not be given.
				If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
					prSet_Red_Combo(sReductionFactor, cboTOSRed)
				End If
				prSet_Red_Combo(sReductionFactor, cboWaistRed)
				prSet_Red_Combo(sReductionFactor, cboMidPointRed)
				prSet_Red_Combo(sReductionFactor, cboLargestRed)
				prSet_Red_Combo(sReductionFactor, cboThighRed)
			End If
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRightLeg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRightLeg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRightLeg.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optRightLeg.GetIndex(eventSender)
			'0 = Full-Leg
			'1 = Panty
			'2 = Brief
			
			Dim sLegStyle, sReductionFactor As String
			Dim iAge As Short
			
			sLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
			
			If sLegStyle = "Full-Leg" Then
				If Index = 0 Then
					optLeftLeg(1).Enabled = True
					optLeftLeg(2).Enabled = True
				Else
					optLeftLeg(1).Enabled = False
					optLeftLeg(2).Enabled = False
				End If
			End If
			
			If sLegStyle = "Panty" Then
				If Index = 1 Then
					optLeftLeg(2).Enabled = True
				Else
					optLeftLeg(2).Enabled = False
				End If
			End If
			
			If sLegStyle = "W4OC-Right" And g_sPreviousLegStyle <> "" Then
				'Reductions, set defaults
				prStore_Reductions()
				If Index = 0 Then
					'Reductions, for full leg
					sReductionFactor = "0.83"
					iAge = Val(txtAge.Text)
					If iAge <= 3 Then sReductionFactor = "0.88"
					If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
				Else
					'Reductions panty leg
					sReductionFactor = "0.85"
				End If
				
				'NB. Special case for TOS as TOS need not be given.
				If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
					prSet_Red_Combo(sReductionFactor, cboTOSRed)
				End If
				prSet_Red_Combo(sReductionFactor, cboWaistRed)
				prSet_Red_Combo(sReductionFactor, cboMidPointRed)
				prSet_Red_Combo(sReductionFactor, cboLargestRed)
				prSet_Red_Combo(sReductionFactor, cboThighRed)
			End If
			
		End If
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hBody;")
		
		Update_DDE_Text_Boxes()
		
		Dim sSymbol As String
		
		sSymbol = "waistbody"
		
		If txtUidWHBody.Text = "" Then
			'Define DB Fields
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFIELDS.D;")
			
			'Find "mainpatientdetails" and get position
			BDYUTILS.PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
			BDYUTILS.PR_PutLine("STRING sMPD_Name;")
			BDYUTILS.PR_PutLine("ANGLE  aMPD_Angle;")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hMPD = UID (" & QQ & "find" & QC & Val(txtUidTitle.Text) & ");")
			BDYUTILS.PR_PutLine("if (hMPD)")
			BDYUTILS.PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
			BDYUTILS.PR_PutLine("else")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")
			
			'Insert body symbol
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
			BDYUTILS.PR_PutLine("hBody = UID (" & QQ & "find" & QC & Val(txtUidWHBody.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			
		End If
		
		'Update the BODY Box symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "TOSCir" & QCQ & txtTOSCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "TOSGivenRed" & QCQ & txtTOSRed.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "TOSHt" & QCQ & txtTOSHt.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtWaistCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistGivenRed" & QCQ & txtWaistRed.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistHt" & QCQ & txtWaistHt.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "MidPointCir" & QCQ & txtMidPointCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "MidPointGivenRed" & QCQ & txtMidPointRed.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "MidPointHt" & QCQ & txtMidPointHt.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LargestCir" & QCQ & txtLargestCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LargestGivenRed" & QCQ & txtLargestRed.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LargestHt" & QCQ & txtLargestHt.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LeftThighCir" & QCQ & txtLeftThighCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "RightThighCir" & QCQ & txtRightThighCir.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ThighGivenRed" & QCQ & txtThighRed.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "FoldHt" & QCQ & txtFoldHt.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "CrotchStyle" & QCQ & txtCrotchStyle.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ID" & QCQ & txtLegStyle.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Body" & QCQ & txtBody.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
		
	End Sub
	
	Private Sub prDisable_WRT_Chap()
		'Disables the text boxes with respect to the Chap Style
		'Called from cboLegStyle_Click
		
		'Store existing values (nb 0 values are not stored)
		prStoreValues()
		
		'    txtTOSCir = ""
		'    txtTOSCir_LostFocus
		'    txtTOSHt = ""
		'    txtTOSHt_LostFocus
		'    txtTOSHtRevised = ""
		'    txtTOSHtRevised_LostFocus
		
		'    txtTOSCir.Enabled = False
		'    labTOSCir.Enabled = False
		'    txtTOSHt.Enabled = False
		'    labTOSHt.Enabled = False
		'    txtTOSHtRevised.Enabled = False
		'    labTOSHtRevised.Enabled = False
		'    labTOS.Enabled = False
		'    cboTOSRed.ListIndex = -1
		'    cboTOSRed.Enabled = False
		
		'    txtMidPointCir = ""
		'    txtMidPointCir_LostFocus
		'    txtMidPointHt = ""
		'    txtMidPointHt_LostFocus
		txtMidPointCir.Enabled = False
		labMidPointCir.Enabled = False
		txtMidPointHt.Enabled = False
		labMidPointHt.Enabled = False
		labMidPoint.Enabled = False
		cboMidPointRed.SelectedIndex = -1
		cboMidPointRed.Enabled = False
		
		'    txtLargestCir = ""
		'    txtLargestCir_LostFocus
		'    txtLargestHt = ""
		'    txtLargestHt_LostFocus
		txtLargestCir.Enabled = False
		labLargestCir.Enabled = False
		txtLargestHt.Enabled = False
		labLargestHt.Enabled = False
		labLargest.Enabled = False
		cboLargestRed.SelectedIndex = -1
		cboLargestRed.Enabled = False
		
		'    txtLeftThighCir = ""
		'    txtLeftThighCir_LostFocus
		txtLeftThighCir.Enabled = False
		labLeftThighCir.Enabled = False
		labLeftThigh.Enabled = False
		
		'    txtRightThighCir = ""
		'    txtRightThighCir_LostFocus
		txtRightThighCir.Enabled = False
		labRightThighCir.Enabled = False
		labRightThigh.Enabled = False
		cboThighRed.SelectedIndex = -1
		cboThighRed.Enabled = False
		
		cboCrotchStyle.SelectedIndex = -1 'No crotch for a chap style
		cboCrotchStyle.Enabled = False
		
	End Sub
	
	Private Sub prDisable_WRT_W4OC()
		
		cboCrotchStyle.SelectedIndex = 0 'Open Crotch
		cboCrotchStyle.Enabled = False
		
	End Sub
	
	Private Sub prEnable_WRT_Chap()
		'Enable all the Text boxes that where disabled when the
		'Chap style was chosen before
		'
		'Restore the original values
		Dim ii As Short
		
		If g_nTOSCir > 0 Then txtTOSCir.Text = CStr(g_nTOSCir)
		txtTOSCir_Leave(txtTOSCir, New System.EventArgs())
		If g_nTOSHt > 0 Then txtTOSHt.Text = CStr(g_nTOSHt)
		txtTOSHt_Leave(txtTOSHt, New System.EventArgs())
		If g_nTOSHtRevised > 0 Then txtTOSHtRevised.Text = CStr(g_nTOSHtRevised)
		txtTOSHtRevised_Leave(txtTOSHtRevised, New System.EventArgs())
		txtTOSCir.Enabled = True
		labTOSCir.Enabled = True
		txtTOSHt.Enabled = True
		labTOSHt.Enabled = True
		txtTOSHtRevised.Enabled = True
		labTOSHtRevised.Enabled = True
		labTOS.Enabled = True
		
		If g_nMidPointCir > 0 Then txtMidPointCir.Text = CStr(g_nMidPointCir)
		txtMidPointCir_Leave(txtMidPointCir, New System.EventArgs())
		If g_nMidPointHt > 0 Then txtMidPointHt.Text = CStr(g_nMidPointHt)
		txtMidPointHt_Leave(txtMidPointHt, New System.EventArgs())
		txtMidPointCir.Enabled = True
		labMidPointCir.Enabled = True
		txtMidPointHt.Enabled = True
		labMidPointHt.Enabled = True
		labMidPoint.Enabled = True
		
		If g_nLargestCir > 0 Then txtLargestCir.Text = CStr(g_nLargestCir)
		txtLargestCir_Leave(txtLargestCir, New System.EventArgs())
		If g_nLargestHt > 0 Then txtLargestHt.Text = CStr(g_nLargestHt)
		txtLargestHt_Leave(txtLargestHt, New System.EventArgs())
		txtLargestCir.Enabled = True
		txtLargestHt.Enabled = True
		labLargestCir.Enabled = True
		labLargestHt.Enabled = True
		labLargest.Enabled = True
		
		If g_nLeftThighCir > 0 Then txtLeftThighCir.Text = CStr(g_nLeftThighCir)
		txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())
		txtLeftThighCir.Enabled = True
		labLeftThighCir.Enabled = True
		labLeftThigh.Enabled = True
		
		If g_nRightThighCir > 0 Then txtRightThighCir.Text = CStr(g_nRightThighCir)
		txtRightThighCir_Leave(txtRightThighCir, New System.EventArgs())
		txtRightThighCir.Enabled = True
		labRightThighCir.Enabled = True
		labRightThigh.Enabled = True
		
		cboTOSRed.Enabled = True
		cboMidPointRed.Enabled = True
		cboLargestRed.Enabled = True
		cboThighRed.Enabled = True
		
		prRestore_Reductions()
		
		'Set value
		For ii = 0 To (cboCrotchStyle.Items.Count - 1)
			If VB6.GetItemString(cboCrotchStyle, ii) = txtCrotchStyle.Text Then
				cboCrotchStyle.SelectedIndex = ii
			End If
		Next ii
		cboCrotchStyle.Enabled = True
		
		
	End Sub
	
	Private Sub prEnable_WRT_W4OC()
		Dim ii As Short
		
		'Enable Left thigh, restore if value available
		prOptLeftLeg(0, 1, 0, 1, 1, 0)
		txtLeftThighCir.Enabled = True
		labLeftThigh.Enabled = True
		If Val(txtLeftThighCir.Text) = 0 And g_nLeftThighCir Then txtLeftThighCir.Text = CStr(g_nLeftThighCir)
		txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())
		
		'Enable Right thigh, restore if value available
		prOptRightLeg(0, 1, 0, 1, 1, 0)
		txtRightThighCir.Enabled = True
		labRightThigh.Enabled = True
		If Val(txtRightThighCir.Text) = 0 And g_nRightThighCir Then txtRightThighCir.Text = CStr(g_nRightThighCir)
		txtRightThighCir_Leave(txtRightThighCir, New System.EventArgs())
		
		'Set value for crotch style
		For ii = 0 To (cboCrotchStyle.Items.Count - 1)
			If VB6.GetItemString(cboCrotchStyle, ii) = txtCrotchStyle.Text Then
				cboCrotchStyle.SelectedIndex = ii
			End If
		Next ii
		cboCrotchStyle.Enabled = True
		
	End Sub
	
	Private Sub prInitialise_Red_Combo(ByRef sRed As String, ByRef Text_Box_Name As System.Windows.Forms.Control, ByRef Combo_Box_Name As System.Windows.Forms.Control)
		Dim nIndex As Short
		
		'Set of combo box
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.78") '0
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.79") '1
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.80") '2
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.81") '3
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.82") '4
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.83") '5
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.84") '6
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.85") '7
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.86") '8
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.87") '9
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.88") '10
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.89") '11
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.90") '12
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.91") '13
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.92") '15
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.AddItem("0.93") '15
		
		'If no values given exit
		If Val(sRed) = 0 And Val(Text_Box_Name.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Combo_Box_Name.ListIndex = -1
			Exit Sub
		End If
		
		'Set the value of the combo box as given
		If sRed = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Combo_Box_Name.ListIndex = (Val(Text_Box_Name.Text) * 100) - 78
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Combo_Box_Name.ListIndex = (Val(sRed) * 100) - 78
		End If
		
	End Sub
	
	Private Sub prOptLeftLeg(ByRef i0 As Short, ByRef i1 As Short, ByRef i2 As Short, ByRef j0 As Short, ByRef j1 As Short, ByRef j2 As Short)
		'Procedure to set the option buttons
		'sets them according to the pattern i0-i2
		'enables or disables according to the pattern j0-j2
		
		optLeftLeg(0).Checked = i0
		optLeftLeg(1).Checked = i1
		optLeftLeg(2).Checked = i2
		optLeftLeg(0).Enabled = j0
		optLeftLeg(1).Enabled = j1
		optLeftLeg(2).Enabled = j2
		
	End Sub
	
	Private Sub prOptRightLeg(ByRef i0 As Short, ByRef i1 As Short, ByRef i2 As Short, ByRef j0 As Short, ByRef j1 As Short, ByRef j2 As Short)
		'Procedure to set the option buttons
		'sets them according to the pattern i0-i2
		'enables or disables according to the pattern j0-j2
		
		optRightLeg(0).Checked = i0
		optRightLeg(1).Checked = i1
		optRightLeg(2).Checked = i2
		optRightLeg(0).Enabled = j0
		optRightLeg(1).Enabled = j1
		optRightLeg(2).Enabled = j2
		
	End Sub
	
	Private Sub prRestore_Reductions()
		
		'Restore reduction values from stored ones
		cboThighRed.SelectedIndex = g_iThighRed
		cboLargestRed.SelectedIndex = g_iLargestRed
		cboMidPointRed.SelectedIndex = g_iMidPointRed
		cboWaistRed.SelectedIndex = g_iWaistRed
		cboTOSRed.SelectedIndex = g_iTOSRed
		
	End Sub
	
	Private Sub prSet_Red_Combo(ByRef sRed As String, ByRef Combo_Box_Name As System.Windows.Forms.Control)
		'Set the value of the combo box as given
		'Find index into combo box list
		
		If Val(sRed) = 0 Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo_Box_Name.ListIndex = (Val(sRed) * 100) - 78
		
	End Sub
	
	Private Sub prStore_Reductions()
		'Store Reduction settings for later restore
		g_iThighRed = cboThighRed.SelectedIndex
		g_iLargestRed = cboLargestRed.SelectedIndex
		g_iMidPointRed = cboMidPointRed.SelectedIndex
		g_iWaistRed = cboWaistRed.SelectedIndex
		g_iTOSRed = cboTOSRed.SelectedIndex
		
	End Sub
	
	Private Sub prStoreValues()
		'Stores the current values of the text boxes.
		'
		'Called from cboLegStyle_Click
		'
		'Used when the leg style changes so that the text box
		'values can be restored if the operator changes the leg style
		'again
		
		If Val(txtTOSCir.Text) > 0 Then g_nTOSCir = Val(txtTOSCir.Text)
		If Val(txtTOSHt.Text) > 0 Then g_nTOSHt = Val(txtTOSHt.Text)
		If Val(txtTOSHtRevised.Text) > 0 Then g_nTOSHtRevised = Val(txtTOSHtRevised.Text)
		
		If Val(txtWaistCir.Text) > 0 Then g_nWaistCir = Val(txtWaistCir.Text)
		If Val(txtWaistHt.Text) > 0 Then g_nWaistHt = Val(txtWaistHt.Text)
		If Val(txtWaistHtRevised.Text) > 0 Then g_nWaistHtRevised = Val(txtWaistHtRevised.Text)
		
		If Val(txtMidPointCir.Text) > 0 Then g_nMidPointCir = Val(txtMidPointCir.Text)
		If Val(txtMidPointHt.Text) > 0 Then g_nMidPointHt = Val(txtMidPointHt.Text)
		
		If Val(txtLargestCir.Text) > 0 Then g_nLargestCir = Val(txtLargestCir.Text)
		If Val(txtLargestHt.Text) > 0 Then g_nLargestHt = Val(txtLargestHt.Text)
		
		If Val(txtLeftThighCir.Text) > 0 Then g_nLeftThighCir = Val(txtLeftThighCir.Text)
		If Val(txtRightThighCir.Text) > 0 Then g_nRightThighCir = Val(txtRightThighCir.Text)
		
		If Val(txtFoldHt.Text) > 0 Then g_nFoldHt = Val(txtFoldHt.Text)
		If Val(txtFoldHtRevised.Text) > 0 Then g_nFoldHtRevised = Val(txtFoldHtRevised.Text)
		
		prStore_Reductions()
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'If this event is called then the
		'DDE link with DRAFIX has not closed
		'Therefor we time out after 6 seconds
		End
	End Sub
	
	Private Sub txtFoldHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHt.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFoldHt)
	End Sub
	
	Private Sub txtFoldHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHt.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtFoldHt, labFoldHt)
	End Sub
	
	Private Sub txtFoldHtRevised_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHtRevised.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFoldHtRevised)
	End Sub
	
	Private Sub txtFoldHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHtRevised.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtFoldHtRevised, labFoldHtRevised)
		If Val(txtFoldHtRevised.Text) > 0 Then
			txtFoldHt.Enabled = False
			labFoldHt.Enabled = False
		Else
			txtFoldHt.Enabled = True
			labFoldHt.Enabled = True
		End If
	End Sub
	
	Private Sub txtLargestCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtLargestCir)
	End Sub
	
	Private Sub txtLargestCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtLargestCir, labLargestCir)
	End Sub
	
	Private Sub txtLargestHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestHt.Enter
		ARMEDDIA1.Select_Text_In_Box(txtLargestHt)
	End Sub
	
	Private Sub txtLargestHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestHt.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtLargestHt, labLargestHt)
	End Sub
	
	Private Sub txtLeftThighCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftThighCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtLeftThighCir)
	End Sub
	
	Private Sub txtLeftThighCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftThighCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtLeftThighCir, labLeftThighCir)
	End Sub
	
	Private Sub txtMidPointCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtMidPointCir)
	End Sub
	
	Private Sub txtMidPointCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtMidPointCir, labMidPointCir)
	End Sub
	
	Private Sub txtMidPointHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointHt.Enter
		ARMEDDIA1.Select_Text_In_Box(txtMidPointHt)
	End Sub
	
	Private Sub txtMidPointHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointHt.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtMidPointHt, labMidPointHt)
	End Sub
	
	Private Sub txtRightThighCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightThighCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtRightThighCir)
	End Sub
	
	Private Sub txtRightThighCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightThighCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtRightThighCir, labRightThighCir)
	End Sub
	
	Private Sub txtTOSCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtTOSCir)
	End Sub
	
	Private Sub txtTOSCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtTOSCir, labTOSCir)
		
		'Special case where we are adding a TOS to existing
		Dim iAge As Short
		Dim sReductionFactor As String
		
		If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Then
			cboTOSRed.Enabled = True
			'Reductions, set defaults
			sReductionFactor = "0.83"
			iAge = Val(txtAge.Text)
			If iAge <= 3 Then sReductionFactor = "0.88"
			If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
			If VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex) = "Panty" Then sReductionFactor = "0.85"
			If cboTOSRed.SelectedIndex = -1 Then
				prSet_Red_Combo(sReductionFactor, cboTOSRed)
			End If
		End If
		
	End Sub
	
	Private Sub txtTOSHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHt.Enter
		ARMEDDIA1.Select_Text_In_Box(txtTOSHt)
	End Sub
	
	Private Sub txtTOSHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHt.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtTOSHt, labTOSHt)
		
		'Special case where we are adding a TOS to existing
		Dim iAge As Short
		Dim sReductionFactor As String
		
		If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Then
			cboTOSRed.Enabled = True
			'Reductions, set defaults
			sReductionFactor = "0.83"
			iAge = Val(txtAge.Text)
			If iAge <= 3 Then sReductionFactor = "0.88"
			If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
			If VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex) = "Panty" Then sReductionFactor = "0.85"
			If cboTOSRed.SelectedIndex = -1 Then
				prSet_Red_Combo(sReductionFactor, cboTOSRed)
			End If
		End If
		
	End Sub
	
	Private Sub txtTOSHtRevised_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHtRevised.Enter
		ARMEDDIA1.Select_Text_In_Box(txtTOSHtRevised)
	End Sub
	
	Private Sub txtTOSHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHtRevised.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtTOSHtRevised, labTOSHtRevised)
		If Val(txtTOSHtRevised.Text) > 0 Then
			txtTOSHt.Enabled = False
			labTOSHt.Enabled = False
		Else
			txtTOSHt.Enabled = True
			labTOSHt.Enabled = True
		End If
	End Sub
	
	Private Sub txtWaistCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Enter
		ARMEDDIA1.Select_Text_In_Box(txtWaistCir)
	End Sub
	
	Private Sub txtWaistCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtWaistCir, labWaistCir)
	End Sub
	
	Private Sub txtWaistHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHt.Enter
		ARMEDDIA1.Select_Text_In_Box(txtWaistHt)
	End Sub
	
	Private Sub txtWaistHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHt.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtWaistHt, labWaistHt)
	End Sub
	
	Private Sub txtWaistHtRevised_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHtRevised.Enter
		ARMEDDIA1.Select_Text_In_Box(txtWaistHtRevised)
	End Sub
	
	Private Sub txtWaistHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHtRevised.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtWaistHtRevised, labWaistHtRevised)
		If Val(txtWaistHtRevised.Text) > 0 Then
			txtWaistHt.Enabled = False
			labWaistHt.Enabled = False
		Else
			txtWaistHt.Enabled = True
			labWaistHt.Enabled = True
		End If
	End Sub
	
	Private Sub Update_DDE_Text_Boxes()
		'Called from OK_Click, assumes data has been validated
		'Updates the text boxes used to transfer DDE data
		
		Dim nCir, nLargestCir, nCrotchFrontFactor As Double
		Dim nOpenBack, nOpenFront As Double
		Dim sString As String
		Dim iRightleg, iLeftleg, ii As Short
		Dim iLegStyle As Short
		
		'Update from COMBO boxes
		'NB. Special case for TOS
		If txtTOSCir.Text <> "" Then
			txtTOSRed.Text = VB6.GetItemString(cboTOSRed, cboTOSRed.SelectedIndex)
		End If
		txtWaistRed.Text = VB6.GetItemString(cboWaistRed, cboWaistRed.SelectedIndex)
		txtMidPointRed.Text = VB6.GetItemString(cboMidPointRed, cboMidPointRed.SelectedIndex)
		txtLargestRed.Text = VB6.GetItemString(cboLargestRed, cboLargestRed.SelectedIndex)
		txtThighRed.Text = VB6.GetItemString(cboThighRed, cboThighRed.SelectedIndex)
		txtLegStyle.Text = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		txtCrotchStyle.Text = VB6.GetItemString(cboCrotchStyle, cboCrotchStyle.SelectedIndex)
		
		'Calculate:-
		'    nCrotchFrontFactor
		'    nOpenBack
		'    nOpenFront
		'These are used in the DRAFIX macro WHCUTDBD.D, transfered via the
		'DB Field "Body".
		'The point of this is to Offload conditional decisions to VB where
		'it is faster, also to reduce the DRAFIX macro code.
		'
		'Establish Largest circumferance
		'NOTE:- this value is also used by the macro WH_LABL.D, passed through "ID"
		'
		nLargestCir = ARMEDDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
		iLegStyle = cboLegStyle.SelectedIndex
		
		'nOpenBack, nOpenFront
		If txtCrotchStyle.Text = "Open Crotch" Then
			If iLegStyle = 3 Or iLegStyle = 4 Then
				'W4OC-Left and W4OC-Right
				If CDbl(txtAge.Text) <= 10 Then
					nOpenBack = 2#
					nOpenFront = 2#
				Else
					If nLargestCir >= 45 Then
						nOpenBack = 4#
						nOpenFront = 4#
					Else
						nOpenBack = 3.5
						nOpenFront = 3.5
					End If
				End If
			Else
				'All other leg styles
				If txtSex.Text = "Male" Then
					If CDbl(txtAge.Text) <= 10 Then
						nOpenBack = 1.5
						nOpenFront = 2.25
					Else
						If nLargestCir >= 45 Then
							nOpenBack = 3.5
							nOpenFront = 4.5
						Else
							nOpenBack = 3#
							nOpenFront = 4#
						End If
					End If
				Else
					If CDbl(txtAge.Text) <= 10 Then
						nOpenBack = 2.25
						nOpenFront = 1.5
					Else
						If nLargestCir >= 45 Then
							nOpenBack = 4.5
							nOpenFront = 3.5
						Else
							nOpenBack = 4#
							nOpenFront = 3#
						End If
					End If
				End If
			End If
		End If
		
		'nCrotchFrontFactor
		If txtSex.Text = "Male" And CDbl(txtAge.Text) >= 3 Then
			If txtCrotchStyle.Text = "Open Crotch" Then
				nCrotchFrontFactor = 0.5
			Else
				nCrotchFrontFactor = 0.4
			End If
		End If
		
		If txtSex.Text = "Female" Or CDbl(txtAge.Text) < 3 Then
			nCrotchFrontFactor = 0.5
		End If
		
		'Update DDE data field "Body". Overwrite previous values
		'As this field will be read using ScanLine(sBody, "blank", ...
		'Accept the leading blank given by Str$, as txtBody will be
		'blank delimited as required by ScanLine.
		'
		'Add the data for revised Waist and TOS heights
		sString = Str(nCrotchFrontFactor) & Str(nOpenFront) & Str(nOpenBack)
		
		If chkDrawRevisedHts.CheckState = 1 Then
			sString = sString & " 1" 'Draw with revised hts
		Else
			sString = sString & " 0" 'Don't draw with revised hts
		End If
		
		sString = sString & Str(Val(txtTOSHtRevised.Text)) & Str(Val(txtWaistHtRevised.Text)) & Str(Val(txtFoldHtRevised.Text))
		txtBody.Text = sString
		
		'Leg Style data field
		'Update DDE data field "ID". Overwrite previous values
		'NOTE:- Includes largest circumferance data, used by WH_LABL.D
		'       GG 8.Nov.94 - Production modifications
		'This is now revised to incorporate data about the Left and Right Legs
		'                CODE       LEFT        RIGHT
		'Full-Leg        0       0 | 1 | 2       0 | 1 | 2
		'Panty           1       1 | 2           1 | 2
		'Brief           2       2               2
		'W4OC-Left       3       0 | 1           -1
		'W4OC-Right      4       -1              0 | 1
		'Chap-Left       5       0 | 1           -1
		'Chap-Right      6       -1              0 | 1
		'Chap Both Legs  7       0 | 1           0 | 1
		'Where:-
		'           -1   =>  No leg style
		'            0   =>  Full-Leg
		'            1   =>  Panty
		'            2   =>  Brief
		
		iRightleg = -1
		iLeftleg = -1
		For ii = 0 To 2
			If optLeftLeg(ii).Checked = True Then iLeftleg = ii
			If optRightLeg(ii).Checked = True Then iRightleg = ii
		Next ii
		sString = Str(cboLegStyle.SelectedIndex) & " " & Str(iLeftleg) & " " & Str(iRightleg) & " " & Str(nLargestCir)
		txtLegStyle.Text = sString
		
	End Sub
	
	Private Function Validate_Data() As Object
		'Checks data for holes etc.
		'Called from OK_Click
		
		'This breaks into two section the first is for fatal errors
		'that must be corrected
		
		'The second is for non-fatal error that require warning to the user
		'who can then fix them or not depending on requirements.
		
		
		Dim sString, sError, NL, sLegStyle As String
		Dim sHt As String
		
		Dim iError, iIndex, ii As Short
		
		Dim nTOSHt, nWaistHt, nMaxCir As Double
		Dim nCir, nFoldHt, nDistance As Double
		
		Dim nLargestHt, nLargestCir, nLargestRed As Double
		Dim nCO_LargestBottOff, nLargestOff, nCO_LargestTopOff As Double
		Dim nButtockArcRad, nCO_Diameter As Double
		
		Dim nDiff, nCrotchFrontFactor As Double
		
		Dim nSeam, nYscale, nXscale, nLength As Double
		Dim nDatum As Double
		
		Dim nLeftThighCir, nThighCir, nRightThighCir As Double
		Dim nThighRed, nThighCirOriginal As Double
		
		Dim xyFold, xyButtockArcCen As XY
		
		Dim nBodyBackCutOutMinTol, nCutOutDiaMaxTol As Double
		
		
		NL = Chr(10) 'new line
		sLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		
		'FATAL Errors
		'~~~~~~~~~~~~
		If (txtTOSCir.Text = "" And txtTOSHt.Text <> "") Or (txtTOSCir.Text <> "" And txtTOSHt.Text = "") Then
			sError = sError & "Missing Top of support value/s." & NL
		End If
		
		If txtWaistCir.Text = "" Or txtWaistHt.Text = "" Then
			sError = sError & "Missing Waist value/s." & NL
		End If
		
		If txtMidPointCir.Text = "" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
			sError = sError & "Missing Mid point Circumferance." & NL
		End If
		
		If txtLargestCir.Text = "" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
			sError = sError & "Missing Largest part of buttocks Circumferance." & NL
		End If
		
		If txtLeftThighCir.Text = "" And sLegStyle <> "W4OC-Right" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
			sError = sError & "Missing Left Thigh value." & NL
		End If
		
		If txtRightThighCir.Text = "" And sLegStyle <> "W4OC-Left" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
			sError = sError & "Missing Right Thigh value." & NL
		End If
		
		If txtFoldHt.Text = "" Then
			sError = sError & "Missing Fold Height value." & NL
		End If
		
		'Except for the Chap style
		If cboCrotchStyle.SelectedIndex < 0 And cboLegStyle.SelectedIndex < 5 Then
			sError = sError & "Missing Crotch Style." & NL
		End If
		
		If cboLegStyle.SelectedIndex < 0 Then
			sError = sError & "Missing Leg Style." & NL
		End If
		
		'Check Crotch style against sex
		sString = VB6.GetItemString(cboCrotchStyle, cboCrotchStyle.SelectedIndex)
		
		'Display Error message (if required) and return
		'These are fatal errors
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Validate_Data = False
			Exit Function
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Validate_Data = True
		End If
		
		'NON-FATAL Errors (warnings)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'Establish Largest circumferance
		nMaxCir = ARMEDDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
		
		'Establish hts (original or revised)
		'If no revised ht given then use the original
		If chkDrawRevisedHts.CheckState = 1 Then
			nFoldHt = ARMEDDIA1.fnDisplayToInches(Val(txtFoldHtRevised.Text))
			nWaistHt = ARMEDDIA1.fnDisplayToInches(Val(txtWaistHtRevised.Text))
			nTOSHt = ARMEDDIA1.fnDisplayToInches(Val(txtTOSHtRevised.Text))
			'If no revised ht given then use the original
			If nFoldHt = 0 Then nFoldHt = ARMEDDIA1.fnDisplayToInches(Val(txtFoldHt.Text))
			If nWaistHt = 0 Then nWaistHt = ARMEDDIA1.fnDisplayToInches(Val(txtWaistHt.Text))
			If nTOSHt = 0 Then nTOSHt = ARMEDDIA1.fnDisplayToInches(Val(txtTOSHt.Text))
		Else
			nFoldHt = ARMEDDIA1.fnDisplayToInches(Val(txtFoldHt.Text))
			nWaistHt = ARMEDDIA1.fnDisplayToInches(Val(txtWaistHt.Text))
			nTOSHt = ARMEDDIA1.fnDisplayToInches(Val(txtTOSHt.Text))
		End If
		
		'Get difference between TOS and fold (note case for explicit TOS)
		If nTOSHt > 0 Then
			nDiff = nTOSHt - nFoldHt
			sHt = "TOS"
		Else
			nDiff = nWaistHt - nFoldHt
			sHt = "Waist"
		End If
		
		'Check Fold Height to TOS/Waist Height
		If nMaxCir <= 39 And nDiff < 5 And nDiff > 0 Then
			sError = sError & "Largest Part of Buttocks Circumference is" & Str(nMaxCir) & " inches" & NL
			sError = sError & "Fold to " & sHt & " distance is" & Str(nDiff) & " inches" & NL
			sError = sError & NL
			sError = sError & "If the buttocks circumferance is :-" & NL
			sError = sError & "less than 39 inches, there must be at least 5 inches" & NL
			sError = sError & "from fold to the " & sHt & "." & NL
			sError = sError & "ref:- 8.1.1.1, GOP 01-02/12, page 5 of 24" & NL & NL
			iError = True
		End If
		
		'Check Fold Height to TOS/Waist Height
		If nMaxCir > 39 And nMaxCir <= 44.875 And ((nDiff < 10 And txtSex.Text = "Male") Or (nDiff < 12 And txtSex.Text = "Female")) And nDiff > 0 Then
			sError = sError & "Largest Part of Buttocks Circumference is" & Str(nMaxCir) & " inches" & NL
			sError = sError & "Fold to " & sHt & " distance is" & Str(nDiff) & " inches" & NL
			sError = sError & NL
			sError = sError & "If the buttocks circumferance is :-" & NL
			sError = sError & "between 39 and 44-7/8 inches, there must be at least" & NL
			sError = sError & "10 inches for Males and 12 inches for Females from the" & NL
			sError = sError & "fold to the " & sHt & "." & NL
			sError = sError & "ref:- 8.1.1.2, GOP 01-02/12, page 5 of 24" & NL & NL
			iError = True
		End If
		
		'Check Fold Height to TOS/Waist Height
		If nMaxCir > 44.875 And nDiff < 15 And nDiff > 0 Then
			sError = sError & "Largest Part of Buttocks Circumference is" & Str(nMaxCir) & " inches" & NL
			sError = sError & "Fold to " & sHt & " distance is" & Str(nDiff) & " inches" & NL
			sError = sError & NL
			sError = sError & "If the buttocks circumferance is :-" & NL
			sError = sError & "greater than 45 inches, there must be at least 15 inches" & NL
			sError = sError & "from the fold to the " & sHt & "." & NL
			sError = sError & "ref:- 8.1.1.3, GOP 01-02/12, page 5 of 24" & NL & NL
			iError = True
		End If
		
		'Check distance from waist to Explicity given TOS
		nDiff = nTOSHt - nWaistHt
		
		If (nDiff < 1.5 Or nDiff > 5) And nTOSHt > 0 Then
			sError = sError & "TOS to Waist distance is" & Str(nDiff) & " inches" & NL
			sError = sError & NL
			sError = sError & "The distance between the Waist and the TOS should be no" & NL
			sError = sError & "less than 1-1/2 inches or more than 5 inches" & NL
			sError = sError & "ref:- 8.2 GOP 01-02/12, page 5 of 24" & NL & NL
			iError = True
		End If
		
		'Thigh figuring
		'Take account of W4OC and CHAP w.r.t. thigh circumferences
		Select Case cboLegStyle.SelectedIndex
			Case 0, 1, 2, 7
				nRightThighCir = ARMEDDIA1.fnDisplayToInches(Val(txtRightThighCir.Text))
				nLeftThighCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeftThighCir.Text))
				
			Case 3, 5 'W4OC-Left and CHAP-Left
				nLeftThighCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeftThighCir.Text))
				nRightThighCir = nLeftThighCir
				
			Case 4, 6 'W4OC-Right and CHAP-Right
				nRightThighCir = ARMEDDIA1.fnDisplayToInches(Val(txtRightThighCir.Text))
				nLeftThighCir = nRightThighCir
				
		End Select
		
		nThighRed = Val(VB6.GetItemString(cboThighRed, cboThighRed.SelectedIndex))
		'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nThighCirOriginal = BDYUTILS.min(nRightThighCir, nLeftThighCir)
		nThighCir = ARMEDDIA1.fnRoundInches(nThighCirOriginal * nThighRed)
		
		'Check for male crotch and female patient
		iIndex = 0
		iIndex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
		If iIndex <> 0 And txtSex.Text = "Female" Then
			sError = sError & "Male Crotch type specified for Female patient" & NL
		End If
		
		'Check for female crotch and male patient
		iIndex = 0
		iIndex = InStr(1, sString, "Female", 0)
		If iIndex <> 0 And txtSex.Text = "Male" Then
			sError = sError & "Female Crotch type specified for Male patient" & NL
		End If
		
		
		'Display Error message (if required) and return
		'These are non-fatal errors
		sError = sError & NL & "The above problems have been found in the data do you" & NL
		sError = sError & "wish to continue?"
		
		If iError = True Then
			iError = MsgBox(sError, 36, "Problems with data")
			If iError = IDYES Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Validate_Data = True
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Validate_Data = False
				Exit Function
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Validate_Data = True
		End If
		
		'nCrotchFrontFactor
		If txtSex.Text = "Male" And Val(txtAge.Text) >= 3 Then
			If VB6.GetItemString(cboCrotchStyle, cboCrotchStyle.SelectedIndex) = "Open Crotch" Then
				nCrotchFrontFactor = 0.5
			Else
				nCrotchFrontFactor = 0.4
			End If
		End If
		
		If txtSex.Text = "Female" Or Val(txtAge.Text) < 3 Then
			nCrotchFrontFactor = 0.5
		End If
		
		'Check Cut-Out values (excluding CHAP styles)
		
		Dim nMinRed, nMaxRed As Short
		If cboLegStyle.SelectedIndex < 5 Then
			'Heights are figured with the fold height as the datum
			nSeam = 0.1875
			nYscale = 0.5
			nXscale = 1.25 / 1.5
			iError = 0
			sError = ""
			iIndex = -1 'Also acts as a flag
			
			nDatum = Int(nFoldHt / 1.5) * 1.5
			
			'Largest part of buttocks figuring
			nLargestHt = ((nWaistHt - nFoldHt) / 3) + nFoldHt
			
			nLargestHt = ARMEDDIA1.fnRoundInches((nLargestHt - nDatum) * nXscale)
			
			nFoldHt = ARMEDDIA1.fnRoundInches((nFoldHt - nDatum) * nXscale)
			
			xyFold.X = nFoldHt
			xyFold.Y = nThighCir * nYscale + nSeam
			
			xyButtockArcCen.X = nLargestHt
			xyButtockArcCen.Y = ((nThighCir / 2) * nYscale) + nSeam
			
			nButtockArcRad = ARMDIA1.FN_CalcLength(xyFold, xyButtockArcCen)
			
			nLargestOff = xyButtockArcCen.Y + nButtockArcRad
			
			'NOTE:-
			'    In both of the following tests we loop only twice.
			'    There can be no more that 5% difference between the reductions
			'    at each circumference.
			
			'Loop to test for 6" cut out
			'We also test to see if the cut-out is -ve, this will only occur
			'with v.small body
			'Test and fix the first time - give information message
			'The second time - give error message
			For ii = 1 To 2
				nLargestRed = Val(VB6.GetItemString(cboLargestRed, cboLargestRed.SelectedIndex))
				nLargestCir = ARMEDDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
				'nLargestCir = fnRoundInches(nLargestCir * nLargestRed)
				nLargestCir = nLargestCir * nLargestRed
				
				nLength = nLargestOff - nSeam
				
				nCO_LargestBottOff = (((nLargestCir * nYscale) - nLength) * nCrotchFrontFactor) + nSeam
				nCO_LargestTopOff = nLargestOff - (((nLargestCir * nYscale) - nLength) * (1 - nCrotchFrontFactor))
				'nCO_Diameter = fnRoundInches(nCO_LargestTopOff - nCO_LargestBottOff)
				nCO_Diameter = nCO_LargestTopOff - nCO_LargestBottOff
				
				nCutOutDiaMaxTol = 6
				If nCO_Diameter < nCutOutDiaMaxTol And nCO_Diameter > 0 Then
					Exit For
				End If
				If nCO_Diameter <= 0 Then
					'Negative or Zero Cut-Out
					If ii = 1 Then
						iIndex = cboLargestRed.SelectedIndex - 5
						If iIndex < 0 Then
							cboLargestRed.SelectedIndex = 0
						Else
							cboLargestRed.SelectedIndex = iIndex
						End If
						sError = sError & "With the given Buttock Reduction, Cut-Out diameter" & NL
						sError = sError & "is negative or zero!" & NL
						sError = sError & "Reduction decreased by 5% to rectify this." & NL
						iError = iError + 1
					Else
						sError = sError & NL & "Severe - Warning!" & NL
						sError = sError & "Cut-Out is still negative or zero!" & NL
						sError = sError & "Even with the reduction decreased by 5%." & NL
						iError = iError + 4
					End If
					
				Else
					'Cut-Out greater than 6"
					If ii = 1 Then
						iIndex = cboLargestRed.SelectedIndex + 5
						If iIndex >= cboLargestRed.Items.Count Then
							cboLargestRed.SelectedIndex = cboLargestRed.Items.Count - 1
						Else
							cboLargestRed.SelectedIndex = iIndex
						End If
						sError = sError & "With given Buttock Reduction, Cut-Out is larger than 6 inches" & NL
						sError = sError & "Reduction increased by 5% to rectify this." & NL
						sError = sError & "ref:- 13.4.3.1 GOP 01-02/12, page 7 of 24" & NL
						iError = iError + 1
					Else
						sError = sError & NL & "Severe - Warning!" & NL
						sError = sError & "Cut-Out is still larger than 6 inches." & NL
						sError = sError & "Even with reduction already increased by 5%." & NL
						iError = iError + 4
					End If
				End If
			Next ii
			
			'Check that minimum distance of 3" from Back Cut-Out to Largest part of
			'buttocks is met
			'Loop to test for 3" minimum
			'Test and fix the first time - give information message.
			'    Except if the reduction has been already modified above wrt
			'    6" Cut-Out maximum.
			'    Uses iIndex as the flag for this.
			'The second time - give error message.
			For ii = 1 To 2
				nLargestRed = Val(VB6.GetItemString(cboLargestRed, cboLargestRed.SelectedIndex))
				nLargestCir = ARMEDDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
				nLargestCir = ARMEDDIA1.fnRoundInches(nLargestCir * nLargestRed)
				
				nCO_LargestTopOff = nLargestOff - (((nLargestCir * nYscale) - nLength) * (1 - nCrotchFrontFactor))
				nDistance = ARMEDDIA1.fnRoundInches(nLargestOff - nCO_LargestTopOff)
				
				nBodyBackCutOutMinTol = 3
				If nDistance >= (nBodyBackCutOutMinTol * nYscale) Then
					Exit For
				Else
					If ii = 1 Then
						If iIndex < 0 Then
							iIndex = cboLargestRed.SelectedIndex + 5
							sError = sError & "With given Buttock Reduction, distance from back of Cut-Out" & NL
							sError = sError & "to largest part of buttocks is less than 3 inches!" & NL
							sError = sError & "Reduction increased by 5% to rectify this." & NL
							sError = sError & "ref:- 16.1 GOP 01-02/12, page 8 of 24" & NL
						End If
						If iIndex >= cboLargestRed.Items.Count Then
							cboLargestRed.SelectedIndex = cboLargestRed.Items.Count - 1
						Else
							cboLargestRed.SelectedIndex = iIndex
						End If
						iError = iError + 2
					Else
						sError = sError & NL & "Severe - Warning!" & NL
						sError = sError & "Distance from back of Cut-Out to largest part of buttocks" & NL
						sError = sError & "is still less than 3 inches!" & NL
						sError = sError & "Even with reduction already increased by 5%." & NL
						iError = iError + 8
					End If
				End If
			Next ii
			
			'Check that all of the reductions are within 5% of each other
			nMinRed = 15
			nMaxRed = 0
			
			If cboTOSRed.SelectedIndex <> -1 And cboTOSRed.SelectedIndex < nMinRed And Val(txtTOSCir.Text) <> 0 Then nMinRed = cboTOSRed.SelectedIndex
			If cboTOSRed.SelectedIndex <> -1 And cboTOSRed.SelectedIndex > nMaxRed And Val(txtTOSCir.Text) <> 0 Then nMaxRed = cboTOSRed.SelectedIndex
			
			If cboWaistRed.SelectedIndex <> -1 And cboWaistRed.SelectedIndex < nMinRed Then nMinRed = cboWaistRed.SelectedIndex
			If cboWaistRed.SelectedIndex <> -1 And cboWaistRed.SelectedIndex > nMaxRed Then nMaxRed = cboWaistRed.SelectedIndex
			
			If cboMidPointRed.SelectedIndex <> -1 And cboMidPointRed.SelectedIndex < nMinRed Then nMinRed = cboMidPointRed.SelectedIndex
			If cboMidPointRed.SelectedIndex <> -1 And cboMidPointRed.SelectedIndex > nMaxRed Then nMaxRed = cboMidPointRed.SelectedIndex
			
			If cboLargestRed.SelectedIndex <> -1 And cboLargestRed.SelectedIndex < nMinRed Then nMinRed = cboLargestRed.SelectedIndex
			If cboLargestRed.SelectedIndex <> -1 And cboLargestRed.SelectedIndex > nMaxRed Then nMaxRed = cboLargestRed.SelectedIndex
			
			If cboThighRed.SelectedIndex <> -1 And cboThighRed.SelectedIndex < nMinRed Then nMinRed = cboThighRed.SelectedIndex
			If cboThighRed.SelectedIndex <> -1 And cboThighRed.SelectedIndex > nMaxRed Then nMaxRed = cboThighRed.SelectedIndex
			
			If (nMaxRed - nMinRed) > 5 Then
				sError = sError & NL & "Severe - Warning!" & NL
				sError = sError & "All Reductions should be within 5% of each other" & NL
				iError = iError + 8
			End If
			
			'Display Error message (if required) and return
			'These are non-fatal errors
			Select Case iError
				Case 0
					'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Validate_Data = True
				Case 1 To 3
					MsgBox(sError, 64, "Warning - Problems with data")
					'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Validate_Data = False
				Case 4 To 100
					sError = sError & NL & "The above problems have been found in the data do you" & NL
					sError = sError & "wish to continue ?"
					iError = MsgBox(sError, 52, "Severe Problems with data")
					If iError = IDYES Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Validate_Data = True
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Validate_Data = False
					End If
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Validate_Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Validate_Data = True
			End Select
			
		End If 'End of If to exclude CHAP leg styles
		
	End Function
End Class
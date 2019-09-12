Option Strict Off
Option Explicit On
Friend Class cadglove
	Inherits System.Windows.Forms.Form
	'Project:   CADGLOVE.MAK
	'Purpose:   To draw the CAD glove with the addition
	'           of glove to elbow / axilla
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
	'15.Aug.96  GG      F.O.C. (Because I'm a Stupid bastard!)
	'                   Upgrade to the Original CAD Glove
	'                   program, to extend above the wrist in
	'                   the same way as the Manual Glove
	'                   This has cost ME! the proverbial arm
	'                   if not a leg as well.
	'
	'04.Jun.98  GG      Bug fix in respect of insert size
	'                   Warning given if insert size below 1/2"
	'
	'06.Dec.98  GG      Upgrade to VB5.
	'                   The code for KEVIN.EXE (the original CAD glove) is now
	'                   contained in a Module KEVIN
	'
	'
	'Notes:-
	'
	'
	'
	'
	'These are now defined in the common module Public.bas
	'Windows API Functions Declarations
	'Private Declare Function GetActiveWindow Lib "User" () As Integer
	'Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	'Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
	'Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
	
	
	
	
	
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
						AppActivate(sTask)
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, g_sDialogID)
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
	
	
	Private Function FN_ValidateData(ByRef sOption As String) As Short
		'This function is used only to make gross checks
		'for missing data.
		'It does not perform any sensibility checks on the
		'data
		'The sOption allows for a less rigourous check
		'    "Check All"
		'    "Ignore Webs"
		'
		Dim nL, sError, sWebError As String
		Dim ii, nn As Short
		
		'Initialise
		FN_ValidateData = False
		sError = ""
		nL = Chr(13)
		
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
				sError = sError & "Missing PIP for " & sCircum(ii) & "!" & nL
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
				sError = sError & "Missing circumference for " & sCircum(ii) & "!" & nL
			End If
		Next ii
		
		
		'Check on lengths
		For ii = 0 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(ii).Text) = 0 And nTotalCir(ii) <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing length " & sLengths(ii) & "!" & nL
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object nTotalCir(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Val(txtLen(ii).Text) <> 0 And nTotalCir(ii) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Length given but missing circumference " & sCircum(ii * 2) & "!" & nL
			End If
		Next ii
		
		If sOption <> "Ignore Webs" Then
			sWebError = ""
			For ii = 5 To 8
				If Val(txtLen(ii).Text) = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sLengths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sWebError = sWebError & "Missing length for " & sLengths(ii) & "!" & nL
				End If
			Next ii
			If sWebError <> "" Then
				If MsgBox(sWebError & nL & "If the missing webs represent amputated fingers then, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
					FN_ValidateData = False
					Exit Function
				End If
			End If
		End If
		
		If cboFabric.Text = "" Then
			sError = sError & "Fabric not given! " & nL
		Else
			If UCase(Mid(cboFabric.Text, 1, 3)) <> "POW" Then
				sError = sError & "Fabric chosen is not Powernet"
			End If
			If Val(Mid(cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid(cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
				sError = sError & "Modulus of fabric chosen is not available on Gram / Tension reduction chart"
			End If
			
		End If
		
		'Validate extension data
		If sOption = "Check All" And g_ExtendTo <> GLOVE_NORMAL Then
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
		
		Static nAge, ii, nn, jj As Short
		Static nValue As Double
		Static sTapesBeyondWrist As String
		
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
		Static sINSERT As String
		Static nInsert As Short
		
		g_iInsertSize = 0
		cboInsert.SelectedIndex = 0
		
		nInsert = Val(Mid(txtTapeLengths2.Text, 15, 2))
		If nInsert > 0 And nInsert <= cboInsert.Items.Count Then
			cboInsert.SelectedIndex = nInsert
			g_iInsertSize = nInsert
		End If
		
		
		'Reinforced DORSAL
		nValue = Val(Mid(txtTapeLengths2.Text, 17, 1))
		chkDorsal.CheckState = nValue
		
		'Fabric
		If txtFabric.Text <> "" Then cboFabric.Text = txtFabric.Text
		
		'Get DDE data for extension
		'This procedure is in the module GLVEXTEN.BAS
		CADGLEXT.PR_GetExtensionDDE_Data()
		
		'Update the values in the glove extension module
		'This takes the values from the dialogue and a sets it up for subsequent
		'procedures in this module
		CADGLEXT.PR_GetDlgAboveWrist()
		
		
		'Having set all the fields with raw data we NOW!
		'refine the display by executing the relevent procedure
		For ii = 0 To 2
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(ii).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("optExtendTo"), Object)(ii).Value = True Then
				CADGLEXT.PR_ExtendTo_Click((ii))
				Exit For
			End If
		Next ii
		
		
		'Setup to test for changes
		ii = FN_GloveDataChanged()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
		g_OnFold = False
		
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
	
	Private Sub cadglove_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
			grdInches.set_RowHeight(ii, 270)
		Next ii
		
		'Setup display of results grid
		For ii = 0 To 1
			grdDisplay.set_ColWidth(ii, 488)
			grdDisplay.set_ColAlignment(ii, 2)
		Next ii
		
		For ii = 0 To 15
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
		Dim INDEX As Short = labCir.GetIndex(eventSender)
		'Procedure to disable the circumference fields
		'to enable the use to select only the relevent
		'data to be drawn
		'Saves having to delele and re-enter to get variations
		'on a glove
		Dim iDIP, iPIP As Short
		If System.Drawing.ColorTranslator.ToOle(labCir(INDEX).ForeColor) = &HC0C0C0 Then
			Select Case INDEX
				Case 0 To 7, 11 To 12
					PR_EnableCir(INDEX)
				Case 8
					PR_EnableCir(INDEX)
					'Disable thumb length
					labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					txtLen(4).Enabled = True
					lblLen(4).Enabled = True
				Case 20 To 23
					iDIP = (INDEX - 20) * 2
					iPIP = iDIP + 1
					labCir(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					PR_EnableCir(iDIP)
					PR_EnableCir(iPIP)
					'Enable finger length
					labLen(INDEX - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
					txtLen(INDEX - 20).Enabled = True
					lblLen(INDEX - 20).Enabled = True
			End Select
		Else
			Select Case INDEX
				Case 0 To 7, 12
					PR_DisableCir(INDEX)
				Case 8
					PR_DisableCir(INDEX)
					'Disable thumb length
					labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					txtLen(4).Enabled = False
					lblLen(4).Enabled = False
				Case 11
					PR_DisableCir(INDEX)
					PR_DisableCir(INDEX + 1)
				Case 20 To 23
					iDIP = (INDEX - 20) * 2
					iPIP = iDIP + 1
					labCir(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					PR_DisableCir(iDIP)
					PR_DisableCir(iPIP)
					'Disable finger length
					labLen(INDEX - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					txtLen(INDEX - 20).Enabled = False
					lblLen(INDEX - 20).Enabled = False
			End Select
		End If
		
	End Sub
	
	Private Sub labLen_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labLen.DoubleClick
		Dim INDEX As Short = labLen.GetIndex(eventSender)
		If INDEX = 5 Then Exit Sub
		If System.Drawing.ColorTranslator.ToOle(labLen(INDEX).ForeColor) = &HC0C0C0 Then
			labLen(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
			txtLen(INDEX).Enabled = True
			lblLen(INDEX).Enabled = True
		Else
			labLen(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			txtLen(INDEX).Enabled = False
			lblLen(INDEX).Enabled = False
		End If
	End Sub
	
	Private Sub lblPleat_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPleat.DoubleClick
		Dim INDEX As Short = lblPleat.GetIndex(eventSender)
		'Procedure to disable the pleat fields
		'to enable the user to select only the relevent
		'data to be drawn
		'Saves having to delele and re-enter to get variations
		'on a glove
		Dim iDIP, iPIP As Short
		If System.Drawing.ColorTranslator.ToOle(lblPleat(INDEX).ForeColor) = &HC0C0C0 Then
			lblPleat(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
			Select Case INDEX
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
			lblPleat(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			Select Case INDEX
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("SSTab1"), Object).Tab <> 0 Then CType(MainForm.Controls("SSTab1"), Object).Tab = 0
		
		If FN_ValidateData("Check All") Then
			'    If True Then
			'Display and hourglass
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			
			'Create the data for the EXTERNAL CAD Glove
			'programme
			PR_CreateCADGloveData()
			
			'Run the CAD Glove
			'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
			PR_RunCADGlove()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fNum = ARMDIA1.FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)
			
			'N.B. Use of local JOBST Dirctory C:\JOBST
			'Set translations and rotations
			XTRANSLATE = -12.937
			YTRANSLATE = CADGLEXT.FN_LengthWristToEOS()
			ROTATION = -90
			
			PR_DrawGloveHPGL_File(("C:\JOBST\GLOVDATA.PLT"))
			
			If g_ExtendTo <> GLOVE_NORMAL Then
				'Calculate glove extensions (if any)
				CADGLEXT.PR_CalculateExtension()
				PR_DrawExtension()
				CADGLEXT.PR_UpdateDDE_Extension()
			End If
			
			PR_UpdateDDE()
			PR_UpdateDBFields("Draw")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
			
			'
			'      'Check that given or returned insert is above a minimum of 3/8"
			'       If g_iInsertSize < 3 Then
			'            If IDNO = MsgBox("The insert size is less than 3/8""! Do you want to continue.", 36, g_sDialogID) Then GoTo EXIT_OK
			'       End If
			'
			' The above was replaces at the request of G.Dunne 04.Jun.98 so that the minimum
			' is now 1/2".
			
			'Check that given or returned insert is above a minimum of 3/8"
			If g_iInsertSize <= 3 Then
				If IDNO = MsgBox("The insert size is less than 1/2""! Do you want to continue.", 36, g_sDialogID) Then GoTo EXIT_OK
			End If
			
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
			Dim INDEX As Short = optExtendTo.GetIndex(eventSender)
			'Use Procedure in module GLVEXTEN.BAS
			If MainForm.Visible Then CADGLEXT.PR_ExtendTo_Click(INDEX)
		End If
	End Sub
	
	Sub PR_CreateCADGloveData()
		Dim fGloveData As Short
		Dim X As Short
		
		Const GLOVEDATA As String = "C:\JOBST\GLOVDATA.DAT"
		
		'Create Glove data file that is input to
		'the CAD glove programme
		
		fGloveData = FreeFile
		FileOpen(fGloveData, GLOVEDATA, OpenMode.Output)
		'Path to executables
		PrintLine(fGloveData, g_sPathJOBST & "\GLOVECAD")
		
		'Standard header
		PrintLine(fGloveData, "VBDATA") 'Name
		PrintLine(fGloveData, "1234") 'File Number
		PrintLine(fGloveData, "1") 'Product types 1 only
		PrintLine(fGloveData, "2") 'Burns
		PrintLine(fGloveData, "N") 'One Glove only
		
		'Which glove Left or Right
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optHand(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optHand"), Object)(0).Value = True Then
			PrintLine(fGloveData, "L") 'Left Hand
		Else
			PrintLine(fGloveData, "R") 'Right Hand
		End If
		
		PrintLine(fGloveData, "N") 'No Zipper
		PrintLine(fGloveData, "N") 'No Slant Inserts
		PrintLine(fGloveData, "N") 'No Reinforcements
		
		'Number of Fingers (Excluding thumb)
		Dim ii, iFingers As Short
		iFingers = 0
		For ii = 0 To 3
			If Val(CType(MainForm.Controls("txtLen"), Object)(ii).Text) <> 0 Then iFingers = iFingers + 1
		Next ii
		PrintLine(fGloveData, Trim(Str(iFingers))) 'Number of fingers
		
		PrintLine(fGloveData, "Y") 'Name is correct
		PrintLine(fGloveData, "Y") 'File# is correct
		PrintLine(fGloveData, "Y") 'Hand is correct
		PrintLine(fGloveData, "1") '1 glove only
		PrintLine(fGloveData, "N") 'N to Lymphedema
		
		'Fabric range
		Dim sFabric As String
		Dim nFabric As Short
		'        sFabric = MainForm!cboFabric.List(MainForm!cboFabric.ListIndex)
		sFabric = CType(MainForm.Controls("cboFabric"), Object).Text
		nFabric = Val(Mid(sFabric, 5, 3))
		
		If nFabric <= 190 Then
			PrintLine(fGloveData, "A") 'A Range Fabric
		Else
			PrintLine(fGloveData, "B") 'B Range Fabric
		End If
		
		PrintLine(fGloveData, "E") 'English measurments
		
		'Circumferences and Lengths
		For ii = 0 To 10
			'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
			PrintLine(fGloveData, FN_CADGloveFormat(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(ii).Text))))
		Next ii
		For ii = 0 To 8
			'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
			PrintLine(fGloveData, FN_CADGloveFormat(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtLen"), Object)(ii).Text))))
		Next ii
		
		'Tapes extending past wrist
		If Val(CType(MainForm.Controls("txtCir"), Object)(11).Text) <> 0 Then
			PrintLine(fGloveData, "Y") 'Does extend beyond wrist
			If Val(CType(MainForm.Controls("txtCir"), Object)(12).Text) = 0 Then
				PrintLine(fGloveData, "1") 'Only one tape past wrist
				'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
				PrintLine(fGloveData, FN_CADGloveFormat(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(11).Text))))
			Else
				PrintLine(fGloveData, "2") 'Two tapes past wrist
				'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
				PrintLine(fGloveData, FN_CADGloveFormat(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(11).Text))))
				'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
				PrintLine(fGloveData, FN_CADGloveFormat(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(12).Text))))
			End If
		Else
			PrintLine(fGloveData, "N") 'Doesn't extend beyond wrist
		End If
		
		PrintLine(fGloveData, "Y") 'Everything correct
		PrintLine(fGloveData, "VB") 'Designers Name
		
		'Tips
		Dim nClosed, nStdOpen As Short
		'UPGRADE_NOTE: Closed was upgraded to Closed_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Closed_Renamed(4) As Short
		Dim StdOpen(4) As Short
		
		nClosed = 0
		nStdOpen = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optThumbTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optIndexTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optMiddleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRingTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLittleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nClosed = System.Math.Abs(CType(MainForm.Controls("optLittleTip"), Object)(0).Value) + System.Math.Abs(CType(MainForm.Controls("optRingTip"), Object)(0).Value) + System.Math.Abs(CType(MainForm.Controls("optMiddleTip"), Object)(0).Value) + System.Math.Abs(CType(MainForm.Controls("optIndexTip"), Object)(0).Value) + System.Math.Abs(CType(MainForm.Controls("optThumbTip"), Object)(0).Value)
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optThumbTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optIndexTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optMiddleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRingTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLittleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nStdOpen = System.Math.Abs(CType(MainForm.Controls("optLittleTip"), Object)(1).Value) + System.Math.Abs(CType(MainForm.Controls("optRingTip"), Object)(1).Value) + System.Math.Abs(CType(MainForm.Controls("optMiddleTip"), Object)(1).Value) + System.Math.Abs(CType(MainForm.Controls("optIndexTip"), Object)(1).Value) + System.Math.Abs(CType(MainForm.Controls("optThumbTip"), Object)(1).Value)
		
		'The simple case is either all open or all closed
		Dim nLen As Double
		Dim nTipCutBack As Double
		If nClosed = 5 Then
			'All tips closed
			PrintLine(fGloveData, "N") 'Open tips Not required
			' The following commented out in response to the fact that
			' the cutback for standard tips when performed by
			' KEVIN.EXE produces strange and varing results.
			' Now force all lengths to desired
			'
			'        ElseIf nStdOpen = 5 Then
			'           'All tips open
			'            Print #fGloveData, "Y"  'Open tips required
			'            Print #fGloveData, "Y"  'On all 4 fingers and thumb
		Else
			'Some open, some closed, some desired
			'If a std tip is required then the finger length is
			'the given - std length
			PrintLine(fGloveData, "Y") 'Open tips required
			PrintLine(fGloveData, "N") 'On all fingers and thumb
			
			'Setup
			If Val(CType(MainForm.Controls("txtAge"), Object).Text) <= 10 Then
				nTipCutBack = CHILD_STD_OPEN_TIP
			Else
				nTipCutBack = ADULT_STD_OPEN_TIP
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLittleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Closed_Renamed(0) = CType(MainForm.Controls("optLittleTip"), Object)(0).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRingTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Closed_Renamed(1) = CType(MainForm.Controls("optRingTip"), Object)(0).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optMiddleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Closed_Renamed(2) = CType(MainForm.Controls("optMiddleTip"), Object)(0).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optIndexTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Closed_Renamed(3) = CType(MainForm.Controls("optIndexTip"), Object)(0).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optThumbTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Closed_Renamed(4) = CType(MainForm.Controls("optThumbTip"), Object)(0).Value
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLittleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StdOpen(0) = CType(MainForm.Controls("optLittleTip"), Object)(1).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRingTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StdOpen(1) = CType(MainForm.Controls("optRingTip"), Object)(1).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optMiddleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StdOpen(2) = CType(MainForm.Controls("optMiddleTip"), Object)(1).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optIndexTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StdOpen(3) = CType(MainForm.Controls("optIndexTip"), Object)(1).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optThumbTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StdOpen(4) = CType(MainForm.Controls("optThumbTip"), Object)(1).Value
			
			For ii = 0 To 4
				nLen = ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtLen"), Object)(ii).Text))
				
				If nLen > 0 And Closed_Renamed(ii) <> True Then
					PrintLine(fGloveData, "Y") 'Open tip required
					If StdOpen(ii) = True Then
						'Cut back for a standard tip
						PrintLine(fGloveData, "N") 'No length is wrong
						'UPGRADE_ISSUE: COM expression not supported: Module methods of COM objects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="5D48BAC6-2CD4-45AD-B1CC-8E4A241CDB58"'
						PrintLine(fGloveData, FN_CADGloveFormat(nLen - nTipCutBack))
					Else
						'Yes this is the desired length
						PrintLine(fGloveData, "Y")
					End If
				Else
					PrintLine(fGloveData, "N") 'Open tip not required
				End If
			Next ii
		End If
		
		PrintLine(fGloveData, "Y") 'Force plot
		
		FileClose(fGloveData)
		
	End Sub
	
	
	
	
	'UPGRADE_WARNING: Event optHand.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optHand_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHand.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optHand.GetIndex(eventSender)
			If INDEX = 0 Then
				g_sSide = "Left"
			Else
				g_sSide = "Right"
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optIndexTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optIndexTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optIndexTip.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optIndexTip.GetIndex(eventSender)
			PR_CutBackTip(3, INDEX)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optLittleTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optLittleTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optLittleTip.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optLittleTip.GetIndex(eventSender)
			PR_CutBackTip(0, INDEX)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMiddleTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMiddleTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMiddleTip.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optMiddleTip.GetIndex(eventSender)
			PR_CutBackTip(2, INDEX)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optProximalTape.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optProximalTape_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optProximalTape.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optProximalTape.GetIndex(eventSender)
			'Use Procedure in module GLVEXTEN.BAS
			CADGLEXT.PR_ProximalTape_Click(INDEX)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRingTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRingTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRingTip.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optRingTip.GetIndex(eventSender)
			PR_CutBackTip(1, INDEX)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optThumbTip.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optThumbTip_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optThumbTip.CheckedChanged
		If eventSender.Checked Then
			Dim INDEX As Short = optThumbTip.GetIndex(eventSender)
			PR_CutBackTip(4, INDEX)
		End If
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
		'UPGRADE_WARNING: Couldn't resolve default property of object nL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nL = Chr(10) 'The new line character
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
	
	Private Sub PR_DisableCir(ByRef INDEX As Short)
		'Read as part of labCir_DblClick
		labCir(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		txtCir(INDEX).Enabled = False
		lblCir(INDEX).Enabled = False
	End Sub
	
	Private Sub PR_DrawExtension()
		'
		' GLOBALS
		'
		'    g_sSide
		'
		'
		Dim ii As Short
		Dim xyTextAlign, xyPt2, xyPt1, xyPt3, xyTmp As XY
		Dim nAge, iiStart As Short
		Dim TmpY, nInsert, TmpX, nOffSet As Double
		'UPGRADE_WARNING: Arrays in structure WristThumbBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure WristLFSBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim WristLFSBlend, WristThumbBlend As Curve
		'UPGRADE_WARNING: Arrays in structure Reinforced may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure LFS may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure ThumbSide may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ThumbSide, LFS, Reinforced As Curve
		Dim nLength, aAngle, nTranslate As Double
		Dim sWorkOrder, sText As String
		Dim xyWrist1, xyNull, xyWrist2 As XY
		
		'Initialise
		nInsert = g_iInsertSize * EIGHTH
		xyDatum.X = 0.80409 'I don't know why either
		xyDatum.y = YTRANSLATE 'I only wrote the %$$&^&*() thing
		
		'Draw on correct layer
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		
		'Get Blending profiles on thumbside
		xyWrist1.X = RadialProfile.X(2)
		xyWrist1.y = RadialProfile.y(2)
		xyWrist2.X = UlnarProfile.X(2)
		xyWrist2.y = UlnarProfile.y(2)
		
		If g_sSide = "Right" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyTmp = xyWrist2
			'UPGRADE_WARNING: Couldn't resolve default property of object xyWrist2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyWrist2 = xyWrist1
			'UPGRADE_WARNING: Couldn't resolve default property of object xyWrist1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyWrist1 = xyTmp
			'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyTmp = xyPalm(6)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPalm(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPalm(6) = xyPalm(1)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPalm(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPalm(1) = xyTmp
		End If
		
		PR_CalcWristBlendThumbSide(xyPalm(4), xyPalm(5), xyPalm(6), xyWrist1, WristThumbBlend)
		
		'Get Blending profiles on Little Finger Side
		CADGLEXT.PR_CalcWristBlendLFS(xyLFAcen, aLFAsweep, xyPalm(3), xyPalm(1), xyWrist2, WristLFSBlend)
		
		'Calculate a point that will be used to align text
		ARMDIA1.PR_MakeXY(xyTextAlign, xyPalm(1).X + 0.25, xyPalm(1).y + 1.625)
		
		
		'Mirror and translate for right hand side
		'To now all the data has been caculated for the LEFT
		'side
		If g_sSide = "Right" Then
			'Find farthest point from datum
			nTranslate = 0
			'Curves
			CADUTILS.PR_MirrorCurveInYaxis(xyPalm(7).X, nTranslate, RadialProfile)
			CADUTILS.PR_MirrorCurveInYaxis(xyPalm(7).X, nTranslate, UlnarProfile)
			'Points
			For ii = 1 To 6
				CADUTILS.PR_MirrorPointInYaxis(xyPalm(7).X, nTranslate, xyPalm(ii))
			Next ii
			'Text alignment
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nLength = ARMDIA1.max((Len(txtPatientName.Text) * 0.075) + 0.125, 1.5)
			ARMDIA1.PR_MakeXY(xyTextAlign, xyPalm(1).X - nLength, xyPalm(1).y + 1.625)
		End If
		
		
		'Draw arm extension
		'Create profile for LitteFingerSide and ThumbSide
		'Little Finger Side
		'Join blending profiles to extension
		LFS.n = 0
		For ii = 1 To WristLFSBlend.n
			LFS.n = LFS.n + 1
			LFS.X(LFS.n) = WristLFSBlend.X(ii)
			LFS.y(LFS.n) = WristLFSBlend.y(ii)
		Next ii
		For ii = 2 To UlnarProfile.n
			LFS.n = LFS.n + 1
			LFS.X(LFS.n) = UlnarProfile.X(ii)
			LFS.y(LFS.n) = UlnarProfile.y(ii)
		Next ii
		
		'Draw
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		ARMDIA1.PR_DrawFitted(LFS)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "LFS")
		BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
		BDYUTILS.PR_AddDBValueToLast("curvetype", "LFS")
		
		'Thumb side
		'Join blending profiles to extension
		ThumbSide.n = 0
		For ii = 1 To WristThumbBlend.n
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(ii)
			ThumbSide.y(ThumbSide.n) = WristThumbBlend.y(ii)
		Next ii
		For ii = 2 To RadialProfile.n
			ThumbSide.n = ThumbSide.n + 1
			ThumbSide.X(ThumbSide.n) = RadialProfile.X(ii)
			ThumbSide.y(ThumbSide.n) = RadialProfile.y(ii)
		Next ii
		'Draw
		ARMDIA1.PR_DrawFitted(ThumbSide)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "ETS")
		BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
		BDYUTILS.PR_AddDBValueToLast("curvetype", "ETS")
		
		'Add labels and construction lines
		ARMDIA1.PR_SetLayer("Construct")
		
		
		'Draw notes for each tape
		nLength = -1
		For ii = 1 To RadialProfile.n
			ARMDIA1.PR_MakeXY(xyPt1, UlnarProfile.X(ii), UlnarProfile.y(ii))
			ARMDIA1.PR_MakeXY(xyPt2, RadialProfile.X(ii), RadialProfile.y(ii))
			If ii = RadialProfile.n - 1 And g_EOSType = ARM_FLAP Then
				'Provide a dummy EOS for the zipper programme
				PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "EOS")
			Else
				PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "")
			End If
			If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
				'Draw elbow line
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_SetLayer("Notes")
				ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.y = xyPt3.y + 0.0625
				ARMDIA1.PR_DrawText("E", xyPt3, 0.1)
			End If
		Next ii
		
		
		If g_EOSType = ARM_FLAP Then
			CADGLEXT.PR_DrawShoulderFlaps(xyPt1, xyPt2)
		Else
			'Draw EOS, account for elastic if last tape is at elbow
			'N.B. our elbow is fixed at tape no 9
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
			ARMDIA1.PR_SetLayer("Notes")
			If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
				ARMDIA1.PR_MakeXY(xyPt3, RadialProfile.X(RadialProfile.n - 1), RadialProfile.y(RadialProfile.n - 1))
				aAngle = System.Math.Abs(90 - ARMDIA1.FN_CalcAngle(xyPt2, xyPt3))
				If Val(txtAge.Text) > 10 Then nOffSet = 0.75 Else nOffSet = 0.375
				nLength = nOffSet / System.Math.Cos(aAngle * (PI / 180))
				ARMDIA1.PR_CalcPolar(xyPt2, ARMDIA1.FN_CalcAngle(xyPt2, xyPt3), nLength, xyPt2)
				If Not g_OnFold Then
					ARMDIA1.PR_MakeXY(xyPt3, UlnarProfile.X(UlnarProfile.n - 1), UlnarProfile.y(UlnarProfile.n - 1))
					ARMDIA1.PR_CalcPolar(xyPt1, ARMDIA1.FN_CalcAngle(xyPt1, xyPt3), nLength, xyPt1)
				Else
					xyPt1.y = xyPt1.y + nOffSet
				End If
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.y = xyPt3.y + 0.0625
				ARMDIA1.PR_DrawText("E", xyPt3, 0.1)
			End If
			
			'Elastic for children
			If Val(txtAge.Text) <= 10 Then
				ARMDIA1.PR_SetLayer("Notes")
				BDYUTILS.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
				xyPt3.y = xyPt3.y + 0.25
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_DrawText("1/2\"" Elastic", xyPt3, 0.1)
			End If
			
			'Draw EOS line
			ARMDIA1.PR_SetLayer("Template" & g_sSide)
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "EOS")
			BDYUTILS.PR_AddDBValueToLast("curvetype", "EOS")
			
		End If
		
		'Draw Centre line
		ARMDIA1.PR_SetLayer("Construct")
		aAngle = ARMDIA1.FN_CalcAngle(xyPt1, xyPt2)
		nLength = ARMDIA1.FN_CalcLength(xyPt1, xyPt2)
		ARMDIA1.PR_CalcPolar(xyPt1, aAngle, nLength / 2, xyPt1)
		ARMDIA1.PR_MakeXY(xyPt1, xyPt1.X, UlnarProfile.y(1))
		ARMDIA1.PR_MakeXY(xyPt2, xyPt1.X, UlnarProfile.y(UlnarProfile.n))
		ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "CL")
		
		
		'Add wrist line on notes
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_DrawLine(xyPalm(1), xyPalm(6))
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "WRIST")
		BDYUTILS.PR_CalcMidPoint(xyPalm(1), xyPalm(6), xyPt3)
		xyPt3.y = xyPt3.y + 0.0625
		ARMDIA1.PR_DrawText("W", xyPt3, 0.1)
		
		'Update symbol gloveglove, glovecommon and marker palmer with data
		'    PR_UpdateDBFields "Draw"
		If g_sSide = "Right" Then
			PR_RepositionPalmerMarker(xyPalm(6))
		Else
			PR_RepositionPalmerMarker(xyPalm(1))
		End If
		ARMDIA1.PR_SetLayer("1")
		
		
		
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
		ARMDIA1.PR_SetLayer("Construct")
		
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
			ARMDIA1.PR_MakeXY(xyInsert, xyDatum.X - 0.1, xyDatum.y + 0.35)
			If TapeNote(iVertex).sNote <> "PLEAT" Then ARMDIA1.PR_SetLayer("Notes")
			ARMDIA1.PR_DrawText(TapeNote(iVertex).sNote, xyInsert, 0.1)
		End If
		
	End Sub
	
	Private Sub PR_EnableCir(ByRef INDEX As Short)
		'Read as part of labCir_DblClick
		labCir(INDEX).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
		txtCir(INDEX).Enabled = True
		lblCir(INDEX).Enabled = True
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
		Dim ii As Object
		'Common Dialogue elements
		g_OnFold = False
		
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
		
		g_nWrist = ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(10).Text))
		If g_ExtendTo <> GLOVE_NORMAL Then
			CADGLEXT.PR_GetDlgAboveWrist()
		Else
			g_iNumTapesWristToEOS = 0
			For ii = 11 To 12
				If Val(CType(MainForm.Controls("txtCir"), Object)(ii).Text) > 0 And CType(MainForm.Controls("txtCir"), Object)(ii).Enabled Then g_iNumTapesWristToEOS = g_iNumTapesWristToEOS + 1
			Next ii
		End If
		
	End Sub
	
	Private Sub PR_RepositionPalmerMarker(ByRef xyPalm As XY)
		'As the extenstion has  changed the position of
		'the wrist we need to adjust the palmer marker
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetGeometry(hPalmer," & QQ & "xmarker" & QQ & CC & "xyStart.x+" & xyPalm.X & CC & "xyStart.y+" & xyPalm.y & ",0.0625);")
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
		Dim INDEX As Short = txtCir.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtCir(INDEX))
	End Sub
	
	Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
		Dim INDEX As Short = txtCir.GetIndex(eventSender)
		
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtCir(INDEX))
		
		'Display value in inches if Valid
		If nLen <> -1 Then lblCir(INDEX).Text = ARMEDDIA1.fnInchestoText(nLen)
		
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
		Dim INDEX As Short = txtExtCir.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtExtCir(INDEX))
	End Sub
	
	Private Sub txtExtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExtCir.Leave
		Dim INDEX As Short = txtExtCir.GetIndex(eventSender)
		
		Dim nLen As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtExtCir(INDEX))
		
		'Display value in inches if Valid
		If nLen <> -1 Then
			'Note we use a grid here that starts at 0
			CADGLEXT.PR_GrdInchesDisplay(INDEX - 8, nLen)
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
		Dim INDEX As Short = txtLen.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtLen(INDEX))
	End Sub
	
	Private Sub txtLen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Leave
		Dim INDEX As Short = txtLen.GetIndex(eventSender)
		Dim nLen As Double
		Dim nTipCutBack As Double
		
		'Check that number is valid (-ve value => invalid)
		nLen = CADGLOVE1.FN_InchesValue(txtLen(INDEX))
		If nLen < 0 Then
			lblLen(INDEX).Text = ""
			Exit Sub
		End If
		
		If INDEX <= 4 Then
			'Special case for finger lengths w.r.t. tips
			If Val(txtAge.Text) <= AGE_CUTOFF Then
				nTipCutBack = CHILD_STD_OPEN_TIP
			Else
				nTipCutBack = ADULT_STD_OPEN_TIP
			End If
			
			Select Case INDEX
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
			lblLen(INDEX).Text = ARMEDDIA1.fnInchestoText(nLen)
		Else
			lblLen(INDEX).Text = ""
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
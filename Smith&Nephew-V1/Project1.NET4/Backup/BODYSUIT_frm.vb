Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class bodysuit
	Inherits System.Windows.Forms.Form
	'Project:   BODYSUIT.MAK
	'Purpose:   Body Dialogue
	'
	'
	'Version:   3.00
	'Date:      3.May.96
	'Author:    Paul O'Rawe
	'           Gary George
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
	
	
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()
		
		If sCurrentValues <> g_sChangeChecker Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, "Bodysuit - Glove Dialogue")
			Select Case Response
				Case IDYES
					PR_CreateMacro_Save("c:\jobst\draw.d")
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(sTask)
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, "Bodysuit - Dialogue")
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
	
	Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click
		
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()
		If sCurrentValues <> g_sChangeChecker Then
			'Check all values are present
			PR_GetValuesFromDialogue()
			If FN_ValidateAndCalculateData(True) Then 'else stay in dialog
				
				PR_CreateMacro_Save("c:\jobst\draw.d")
				sTask = fnGetDrafixWindowTitleText()
				If sTask <> "" Then
					AppActivate(sTask)
					System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
					End
				Else
					MsgBox("Can't find a Drafix Drawing to update!", 16, "BodySuit - Dialogue")
				End If
			End If
		Else
			End
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event cboBackNeck.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboBackNeck_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboBackNeck.SelectedIndexChanged
		Select Case VB.Left(cboBackNeck.Text, 1)
			Case "R", "S" 'Regular or Scoop
				txtBackNeck.Text = ""
				lblBackNeck.Text = ""
				txtBackNeck.Enabled = False
				labBackNeck.Enabled = False
			Case Else 'Measured Scoop
				If g_sBackNeck <> "" Then
					txtBackNeck.Text = g_sFrontNeck
					txtBackNeck_Leave(txtBackNeck, New System.EventArgs()) 'Display inches
				End If
				txtBackNeck.Enabled = True
				labBackNeck.Enabled = True
		End Select
	End Sub
	
	'UPGRADE_WARNING: Event cboCrotchStyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCrotchStyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCrotchStyle.SelectedIndexChanged
		Dim sString, sError As String
		Dim iindex As Short
		
		sError = ""
		
		'Check Crotch style against sex
		sString = VB6.GetItemString(cboCrotchStyle, cboCrotchStyle.SelectedIndex)
		
		'Check for male crotch and female patient
		iindex = 0
		iindex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
		If iindex <> 0 And txtSex.Text = "Female" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male Crotch type specified for Female patient" + NL
		End If
		
		'Check for female crotch and male patient
		iindex = 0
		iindex = InStr(1, sString, "Female", 0)
		If iindex <> 0 And txtSex.Text = "Male" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Female Crotch type specified for Male patient" + NL
		End If
		
		
		'Display Error message (if required) and return
		'These are fatal errors
		If Len(sError) > 0 Then
			MsgBox(sError, 0, "Crotch Style Error")
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event cboFrontNeck.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFrontNeck_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFrontNeck.SelectedIndexChanged
		Select Case VB.Left(cboFrontNeck.Text, 1)
			Case "R", "S", "V" 'Regular or Scoop or V neck
				txtFrontNeck.Text = ""
				lblFrontNeck.Text = ""
				txtFrontNeck.Enabled = False
				labFrontNeck.Enabled = False
			Case Else 'Measured Scoop, Measured V or Turtle neck
				If g_sFrontNeck <> "" Then
					txtFrontNeck.Text = g_sFrontNeck
					txtFrontNeck_Leave(txtFrontNeck, New System.EventArgs()) 'Display inches
				End If
				txtFrontNeck.Enabled = True
				labFrontNeck.Enabled = True
		End Select
	End Sub
	
	'UPGRADE_WARNING: Event cboLeftAxilla.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLeftAxilla_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLeftAxilla.SelectedIndexChanged
		If cboLeftAxilla.Text = "Sleeveless" Then
			cboRightAxilla.Text = "Sleeveless"
		Else
			If cboRightAxilla.Text = "Sleeveless" Then
				cboRightAxilla.Text = cboLeftAxilla.Text
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboLeftCup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLeftCup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLeftCup.SelectedIndexChanged
		If Me.Visible And (cboLeftCup.Text = "" Or cboLeftCup.Text = "None") Then
			txtLeftDisk.Text = ""
			g_sLeftCup = ""
			g_sLeftDisk = ""
		Else
			PR_EnableCalculateDiskButton()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboLegStyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLegStyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLegStyle.SelectedIndexChanged
		'
		Dim sLegStyle As String
		
		
		'Get the currently selected leg style
		sLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		If sLegStyle = g_sPreviousLegStyle Then Exit Sub
		
		Select Case sLegStyle
			Case "Panty"
				optLeftLeg(0).Enabled = True
				optRightLeg(0).Enabled = True
				
				optLeftLeg(0).Checked = 1 'Left leg Panty
				optLeftLeg(1).Checked = 0 'Left leg Brief
				
				optRightLeg(0).Checked = 1 'Right leg Panty
				optRightLeg(1).Checked = 0 'Right leg Brief
				
				
			Case "Brief", "Brief-French"
				optLeftLeg(0).Checked = 0 'Left leg Panty
				optLeftLeg(0).Enabled = False
				optLeftLeg(1).Checked = 1 'Left leg Brief
				
				optRightLeg(0).Checked = 0 'Right leg Panty
				optRightLeg(0).Enabled = False
				optRightLeg(1).Checked = 1 'Right leg Brief
				
				optLeftLeg(1).Enabled = True 'We enable just in case as
				optRightLeg(1).Enabled = True '
				
		End Select
		
		'Store leg style for use when it changes
		g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		
	End Sub
	
	'UPGRADE_WARNING: Event cboRightAxilla.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboRightAxilla_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboRightAxilla.SelectedIndexChanged
		If cboRightAxilla.Text = "Sleeveless" Then
			cboLeftAxilla.Text = "Sleeveless"
		Else
			If cboLeftAxilla.Text = "Sleeveless" Then
				cboLeftAxilla.Text = cboRightAxilla.Text
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboRightCup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboRightCup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboRightCup.SelectedIndexChanged
		If Me.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
			txtRightDisk.Text = ""
			g_sRightCup = ""
			g_sRightDisk = ""
		Else
			PR_EnableCalculateDiskButton()
		End If
	End Sub
	
	Private Sub cmdCalculateBraDisks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalculateBraDisks.Click
		PR_DoBraCupsAndDisks()
	End Sub
	
	Private Sub cmdTab_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTab.Click
		'Allows the user to use enter as a tab
		System.Windows.Forms.SendKeys.Send("{TAB}")
	End Sub
	
	Private Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Fuction to return the length between two points
		'Greatfull thanks to Pythagorus
		
		FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)
		
	End Function
	
	Private Function FN_CalculateBra(ByRef nChest As Double, ByRef nUnderBreast As Double, ByRef nNipple As Double, ByRef sCup As String, ByRef iDisk As Short, ByRef sSide As String) As Short
		'Procedure to calcluate the size of a BRA
		'Returning the CUP size and the Bra disk size.
		'
		'INPUT
		'    nChest   Chest Circumference (Inches)
		'
		'    nNipple  Circumference over nipple line (Inches)
		'
		'    nUnderBreast
		'             Circumference just under the breast (Inches)
		'
		'
		'    sCup        Cup size
		'                If a cup size is given then disk size
		'                is based on the given cup.
		'
		'    sSide       Used when displaying error & warning messages
		'
		'OUTPUT
		'    sCup        If no cup is given then a cup is calculated
		'                and the disk size calculated from this cup.
		'
		'    iDisk       Size of disk to be used on template
		'
		'    FN_CalculateBra
		'                True if no errors
		'                False if errors, Message returned in sError
		'
		'    If sCup [and | or] iDisk are given then these are not changed
		'    the values are calculated and a warning given if the calculated
		'    values are different from the given
		'
		'NOTES
		'    Cup sizes   TRAINING, A, B, C, D, E, NONE
		'    Cup sizes   -1, 0 and 1 to 9
		'                Where 0 => a specified missing cup.
		'                and  -1 => failure to calculate.
		'
		'SPECIFICATIONS
		'    GOP 01-02/17    VEST WITHOUT SLEEVES
		'                    Pages 11,12 29.October.1991
		'    GOP 01-02/17    VEST WITH SLEEVES
		'                    Pages 17,18 29.October.1991
		'
		'    BODYBRA.D       DRAFIX macro version 1.04
		'
		'    BRACHART.DAT    Data used in conjunction with
		'                    the macro BODYBRA.D
		'
		
		'Variables
		Dim nDiff As Double
		Dim nSelctedDisk As Short
		Dim sError As String
		
		'Initially set to false
		FN_CalculateBra = False
		sError = ""
		
		'Simple case for Cup type = "None"
		If sCup = "None" Then
			FN_CalculateBra = True
			iDisk = 0
			Exit Function
		End If
		
		
		'Calulate bracup
		If sCup = "" Then
			If nNipple = 0 Then
				FN_CalculateBra = False
				sError = "Can't calculate a cup as Circumference over Nipple line is missing"
				MsgBox(sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")")
				Exit Function
			End If
			
			nDiff = nNipple - nChest
			If nDiff < 0 Then
				FN_CalculateBra = False
				sError = "Circumference over Nipple line is smaller than Chest circumference"
				MsgBox(sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")")
				Exit Function
			End If
			
			'Calculate bra cup
			If nDiff <= 1.375 Then
				sCup = "A"
			ElseIf nDiff <= 2.375 Then 
				sCup = "B"
			ElseIf nDiff <= 3.375 Then 
				sCup = "C"
			ElseIf nDiff <= 4.375 Then 
				sCup = "D"
			Else
				sCup = "E"
			End If
		End If
		
		
		'Bra cup to disk mappings
		nSelctedDisk = -1
		Select Case nUnderBreast
			Case 26 To 28.99
				If sCup = "Training" Then
					nSelctedDisk = 1
				ElseIf sCup = "A" Then 
					nSelctedDisk = 1
				ElseIf sCup = "B" Then 
					nSelctedDisk = 2
				ElseIf sCup = "C" Then 
					nSelctedDisk = 3
				ElseIf sCup = "D" Then 
					nSelctedDisk = 4
				ElseIf sCup = "E" Then 
					nSelctedDisk = 5
				Else : nSelctedDisk = -1
				End If
			Case 29 To 31.99
				If sCup = "Training" Then
					nSelctedDisk = 1
				ElseIf sCup = "A" Then 
					nSelctedDisk = 2
				ElseIf sCup = "B" Then 
					nSelctedDisk = 3
				ElseIf sCup = "C" Then 
					nSelctedDisk = 4
				ElseIf sCup = "D" Then 
					nSelctedDisk = 5
				ElseIf sCup = "E" Then 
					nSelctedDisk = 6
				Else : nSelctedDisk = -1
				End If
			Case 32 To 34.99
				If sCup = "Training" Then
					nSelctedDisk = 2
				ElseIf sCup = "A" Then 
					nSelctedDisk = 3
				ElseIf sCup = "B" Then 
					nSelctedDisk = 4
				ElseIf sCup = "C" Then 
					nSelctedDisk = 5
				ElseIf sCup = "D" Then 
					nSelctedDisk = 6
				ElseIf sCup = "E" Then 
					nSelctedDisk = 7
				Else : nSelctedDisk = -1
				End If
			Case 35 To 37.99
				If sCup = "Training" Then
					nSelctedDisk = 3
				ElseIf sCup = "A" Then 
					nSelctedDisk = 4
				ElseIf sCup = "B" Then 
					nSelctedDisk = 5
				ElseIf sCup = "C" Then 
					nSelctedDisk = 6
				ElseIf sCup = "D" Then 
					nSelctedDisk = 7
				ElseIf sCup = "E" Then 
					nSelctedDisk = 8
				Else : nSelctedDisk = -1
				End If
			Case 38 To 40.99
				If sCup = "A" Then
					nSelctedDisk = 5
				ElseIf sCup = "B" Then 
					nSelctedDisk = 6
				ElseIf sCup = "C" Then 
					nSelctedDisk = 7
				ElseIf sCup = "D" Then 
					nSelctedDisk = 8
				ElseIf sCup = "E" Then 
					nSelctedDisk = 9
				Else : nSelctedDisk = -1
				End If
			Case 41 To 44
				If sCup = "A" Then
					nSelctedDisk = 6
				ElseIf sCup = "B" Then 
					nSelctedDisk = 7
				ElseIf sCup = "C" Then 
					nSelctedDisk = 8
				ElseIf sCup = "D" Then 
					nSelctedDisk = 9
				Else : nSelctedDisk = -1
				End If
		End Select
		
		If nSelctedDisk = -1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = "No Bra disk is availabe for a Cup Size " & sCup & NL & "For an under breast circumference of " & ARMEDDIA1.fnInchestoText(nUnderBreast)
			MsgBox(sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")")
			FN_CalculateBra = False
			Exit Function
		End If
		
		'Cut back disk by 1 if it is over 5
		If nSelctedDisk > 5 Then nSelctedDisk = nSelctedDisk - 1
		
		
		'Set disk size.
		'If disk is given then only check that size is the same, warn if not!
		If iDisk > 0 Then
			If iDisk <> nSelctedDisk Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = "Warning" & NL & "The Given disk is different in size to the disk that would have been calculated!"
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & NL & "Given Disk      : " & Str(iDisk)
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & NL & "Calculated Disk : " & Str(nSelctedDisk)
				MsgBox(sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")")
			End If
		Else
			iDisk = nSelctedDisk
		End If
		
		FN_CalculateBra = True
		
		
	End Function
	
	Private Function FN_Decimalise(ByRef nDisplay As Double) As Double
		'This function takes the value given and converts it to its
		'decimal equivelent.
		'
		'
		'Input:-
		'        nDisplay is the value as input by the operator in the
		'        dialog.
		'        The convention is that, Metric dimensions use the decimal
		'        point to indicate the division between CMs and mms
		'        ie 7.6 = 7 cm and 6 mm.
		'        Whereas the decimal point for imperial measurements indicates
		'        the division between inches and eighths
		'        ie 7.6 = 7 inches and 6 eighths
		'
		'Globals:-
		'        g_nUnitsFac = 1       => nDisplay in Inches
		'        g_nUnitsFac = 10/25.5 => nDisplay in CMs
		'
		'Returns:-
		'        Single,
		'        For CMs value is returned unaltered.
		'        For Inches value is returned in decimal inches
		'        -1,     on conversion error.
		'
		'Errors:-
		'        The returned value is usually +ve. Unless it can't
		'        be sucessfully converted to inches.
		'        Eg 7.8 is an invalid number if g_nUnitsFac = 1
		'
		'                            WARNING
		'                            ~~~~~~~
		'In most cases the input is a +ve number.  This function will handle a
		'-ve number but in this case the error checking is invalid.  This
		'is done to provide a general conversion tool.  Where the input is
		'likley to be -ve then the calling subroutine or function should check
		'the sensibility of the returned value for that specific case.
		'
		
		Dim iInt, iSign As Short
		Dim nDec As Double
		'retain sign
		iSign = System.Math.Sign(nDisplay)
		nDisplay = System.Math.Abs(nDisplay)
		
		'Simple case where Units are CM
		If g_nUnitsFac <> 1 Then
			FN_Decimalise = nDisplay * iSign
			Exit Function
		End If
		
		'Imperial units
		iInt = Int(nDisplay)
		nDec = nDisplay - iInt
		'Check that conversion is possible (return -1 if not)
		If nDec > 0.8 Then
			FN_Decimalise = -1
		Else
			FN_Decimalise = (iInt + (nDec * 0.125 * 10)) * iSign
		End If
		
	End Function
	
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
	
	Private Function FN_InchesValue(ByRef TextBox As System.Windows.Forms.Control) As Double
		' Check for numeric values
		Dim sChar, sText As String
		Dim iLen, nn As Short
		Dim nLen As Double
		
		sText = TextBox.Text
		iLen = Len(sText)
		
		'Check the actual structure of the input
		FN_InchesValue = -1
		For nn = 1 To iLen
			sChar = Mid(sText, nn, 1)
			If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
				MsgBox("Invalid - Dimension has been entered", 48, "BodySuit - Dialogue")
				TextBox.Focus()
				FN_InchesValue = -1
				Exit For
			End If
		Next nn
		
		'Convert to inches
		nLen = ARMEDDIA1.fnDisplayToInches(Val(TextBox.Text))
		If nLen = -1 Then
			MsgBox("Invalid - Length has been entered", 48, "BodySuit Details")
			TextBox.Focus()
			FN_InchesValue = -1
		Else
			FN_InchesValue = nLen
		End If
		
	End Function
	
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
		PrintLine(fNum, "//by Visual Basic, BodySuit")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//type - " & sType & "")
		
		
	End Function
	
	
	
	Private Function Fn_ValidateData() As Short
		Dim iError As Object
		Dim ii As Object
		Dim nShldToWaistRed As Object
		Dim nShldToFoldRed As Object
		Dim nRightThighCirGiven As Object
		Dim nLeftThighCirGiven As Object
		Dim nShldToWaistGiven As Object
		Dim nShldWidthGiven As Object
		'This function is used only to make gross checks
		'for missing data.
		'It does not perform any sensibility checks on the
		'data
		'Check CutOut, Back Seams and Shoulder width to see
		'if any changes are required - if there is the dialog
		'is updated.
		
		Dim sError As String
		Dim iFatalError As Short
		Dim nShoulderWidthGiven As Double
		Dim nChestCirGiven As Double
		Dim nWaistCirGiven As Double
		Dim nShldToFoldGiven As Double
		Dim nButtocksCirGiven As Double
		Dim nWaistCirRed As Double
		Dim nChestCirRed As Double
		Dim nButtocksCirRed As Double
		Dim nThighCirRed As Double
		Dim sCrotchStyle As String
		Dim nSmallestThighGiven As Double
		Dim nWaistRem As Double
		Dim nChestRem As Double
		Dim nChestCir As Double
		Dim nWaistCir As Double
		Dim nButtocksCir As Double
		Dim nGroinHeight As Double
		Dim nShldToFold As Double
		Dim nShldToWaist As Double
		Dim nButtocksLength As Double
		Dim nButtBackSeamRatio As Double
		Dim nButtFrontSeamRatio As Double
		Dim nButtRedAtLimit As Short
		Dim nThighRedAtLimit As Short
		Dim nButtRedIncreased As Short
		Dim nThighRedDecreased As Short
		Dim nHalfGroinHeight As Double
		Dim nButtRadius As Double
		Dim nButtCirCalc As Double
		Dim nButtBackSeam As Double
		Dim nButtFrontSeam As Double
		Dim nButtCutOut As Double
		Dim nWaistFrontSeam As Double
		Dim nChestFrontSeam As Double
		Dim nWaistCutOut As Double
		Dim nWaistCutOutModified As Short
		
		'Initialise
		Fn_ValidateData = False
		sError = ""
		iFatalError = False
		
		'Get values from dialog
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldWidthGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldWidthGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(3).Text))
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToWaistGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldToWaistGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(4).Text))
		nChestCirGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(5).Text))
		nWaistCirGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(6).Text))
		nShldToFoldGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(12).Text))
		nButtocksCirGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(14).Text))
		'UPGRADE_WARNING: Couldn't resolve default property of object nLeftThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nLeftThighCirGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(15).Text))
		'UPGRADE_WARNING: Couldn't resolve default property of object nRightThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nRightThighCirGiven = ARMEDDIA1.fnDisplayToInches(Val(txtCir(16).Text))
		sCrotchStyle = cboCrotchStyle.Text
		
		'Reductions
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToFoldRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldToFoldRed = 0.95 '95%
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToWaistRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldToWaistRed = 0.95 '95%
		If cboRed(1).Text <> "" Then
			nChestCirRed = Val(cboRed(1).Text)
		Else
			nChestCirRed = 0.85 '85% default
		End If
		If cboRed(0).Text <> "" Then
			nWaistCirRed = Val(cboRed(0).Text)
		Else
			nWaistCirRed = 0.85 '85% default
		End If
		If cboRed(3).Text <> "" Then
			nButtocksCirRed = Val(cboRed(3).Text)
		Else
			nButtocksCirRed = 0.85 '85% default
		End If
		If cboRed(4).Text <> "" Then
			nThighCirRed = Val(cboRed(4).Text)
		Else
			nThighCirRed = 0.85 '85% default
		End If
		
		'Test minimum value for Shoulder width
		If nShldWidthGiven < 0.5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Shoulder Width is less than 1/2""!" & NL
		End If
		
		Dim sCircum(16) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(0) = "Left shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(1) = "Right shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(2) = "Neck circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(3) = "Shoulder width."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(4) = "Shoulder to waist."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(5) = "Chest circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(6) = "Waist circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(9) = "Shoulder to under breast."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(10) = "Circ. under breast."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(11) = "Circ. over nipple."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(12) = "Shoulder to Fold of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(13) = "Shoulder to Large Part of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(14) = "Circ. of Large Part of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(15). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(15) = "Left Thigh Circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(16). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(16) = "Right Thigh Circ."
		
		'FATAL Errors
		'~~~~~~~~~~~~
		'Body measurements (all must be present)
		For ii = 0 To 6
			If Val(txtCir(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing " & sCircum(ii) & NL
			End If
		Next ii
		For ii = 12 To 16
			If Val(txtCir(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing " & sCircum(ii) & NL
			End If
		Next ii
		
		'Bra Cups
		'Note:
		'    The Circumference over nipple is optional unless a cup has been
		'    specified
		'
		If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing " & sCircum(9) & NL
		End If
		If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing " & sCircum(10) & NL
		End If
		
		If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & NL
		End If
		
		If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra disks given! But Missing " & sCircum(9) & NL
		End If
		
		'If cups or dimensions given then a disk must be present
		'NB
		'    cboXXXXCup.ListIndex = 6 = "None"
		'    cboXXXXCup.ListIndex = 7 = ""
		
		If cboLeftCup.SelectedIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Left BRA Cup!" & NL
		End If
		If cboRightCup.SelectedIndex < 6 And Val(txtRightDisk.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Right BRA Cup!" & NL
		End If
		
		'Sex error
		If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.SelectedIndex < 0 Or cboRightCup.SelectedIndex < 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & NL
		End If
		
		'Neck at back and front
		Dim sChar As New VB6.FixedLengthString(1)
		sChar.Value = VB.Left(cboBackNeck.Text, 1)
		If sChar.Value = "M" And txtBackNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Back Meck Measured Scoop!" & NL
		End If
		
		sChar.Value = VB.Left(cboFrontNeck.Text, 1)
		If sChar.Value = "M" And txtFrontNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Front Neck Measured Scoop!" & NL
		End If
		
		'
		If cboLeftAxilla.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Left Axilla not given!" & NL
		End If
		
		If cboRightAxilla.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Right Axilla not given!" & NL
		End If
		
		If cboFrontNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Neck not given!" & NL
		End If
		
		If cboBackNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Back neck not given!" & NL
		End If
		
		If cboClosure.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Closure not given!" & NL
		End If
		
		If cboFabric.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Fabric not given!" & NL
		End If
		
		'Extra Values
		If cboLegStyle.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Leg Style not given!" & NL
		End If
		If cboCrotchStyle.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Crotch Style not given!" & NL
		End If
		
		'Display Error message (if required) and return
		'These are fatal errors
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			Fn_ValidateData = False
			Exit Function
		Else
			Fn_ValidateData = True
		End If
		
		'Possible FATAL Errors and WARNINGS
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'Figured Variables for tests
		'Note: Shoulder Width & Nipple Circumference not figured
		nChestCir = ARMEDDIA1.fnRoundInches(nChestCirGiven * nChestCirRed) / 2 'half scale
		nWaistCir = ARMEDDIA1.fnRoundInches(nWaistCirGiven * nWaistCirRed) / 2 'half-scale
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToFoldRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldToFold = ARMEDDIA1.fnRoundInches(nShldToFoldGiven * nShldToFoldRed) + INCH1_2
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToWaistRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nShldToWaistGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nShldToWaist = ARMEDDIA1.fnRoundInches(nShldToWaistGiven * nShldToWaistRed) + INCH1_2
		nButtocksLength = ARMEDDIA1.fnRoundInches((nShldToFold - nShldToWaist) / 3)
		
		'Test Measurements
		
		'NOTE:-
		'    There can be no more that 5% difference between the reductions
		'    at each circumference.
		
		'Use smallest thigh size.
		'UPGRADE_WARNING: Couldn't resolve default property of object nRightThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nLeftThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (nLeftThighCirGiven < nRightThighCirGiven) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nLeftThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nSmallestThighGiven = nLeftThighCirGiven
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nRightThighCirGiven. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nSmallestThighGiven = nRightThighCirGiven
		End If
		
		If InStr(sCrotchStyle, "Fly") <> 0 Then
			nButtBackSeamRatio = 0.6 '60%
			nButtFrontSeamRatio = 0.4 '40%
		Else
			nButtBackSeamRatio = 0.5 '50%
			nButtFrontSeamRatio = 0.5 '50%
		End If
		
		'CutOut and Back Seam Test
		nButtRedAtLimit = False
		nThighRedAtLimit = False
		nButtRedIncreased = False
		nThighRedDecreased = False
		Do 
			If nButtocksCirRed = 0.9 Then 'at Reduction limit of 90%
				nButtRedAtLimit = True
			End If
			'I had to use Val and Str$ functions because VB3 has a problem
			'when returning values from the minus operation,
			'i.e. nThighCirRed was not equal to .8 exactly (but it should be!!!)
			If Val(Str(nThighCirRed)) = 0.8 Then 'at Reduction limit of 80%
				nThighRedAtLimit = True
			End If
			If nThighRedAtLimit And nButtRedAtLimit Then
				Exit Do
			End If
			nButtocksCir = ARMEDDIA1.fnRoundInches(nButtocksCirGiven * nButtocksCirRed) / 2 'half scale
			nGroinHeight = ARMEDDIA1.fnRoundInches(nSmallestThighGiven * nThighCirRed) / 2 'half scale
			nHalfGroinHeight = (nGroinHeight / 2)
			nButtRadius = System.Math.Sqrt((nButtocksLength ^ 2) + (nHalfGroinHeight ^ 2))
			nButtCirCalc = nButtRadius + nHalfGroinHeight
			nButtBackSeam = (nButtocksCir - nButtCirCalc) * nButtBackSeamRatio
			If nButtBackSeam < 1.5 Then 'less than 3" full scale
				If nButtRedAtLimit Then
					nThighCirRed = nThighCirRed - 0.05 'decrease the reduction by 5%
					nThighRedDecreased = True
				Else
					nButtocksCirRed = nButtocksCirRed + 0.05 'increase the reduction by 5%
					nButtRedIncreased = True
				End If
			Else
				Exit Do 'greater than 3" full scale - acceptable
			End If
		Loop 
		If nButtRedIncreased Then
			cboRed(3).Text = CStr(nButtocksCirRed)
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With given Buttock Reduction, distance from back of Cut-Out" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "to largest part of buttocks is less than 3 inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Buttocks Reduction increased to rectify this." + NL
			'sError = sError + "ref:- " + NL
		End If
		If nThighRedDecreased Then
			cboRed(4).Text = CStr(nThighCirRed)
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With given Thigh Reduction, distance from back of Cut-Out" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "to largest part of buttocks is less than 3 inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Thigh Reduction decreased to rectify this." + NL
			'sError = sError + "ref:- " + NL
		End If
		'One last time with Reductions
		nButtocksCir = ARMEDDIA1.fnRoundInches(nButtocksCirGiven * nButtocksCirRed) / 2 'half scale
		nGroinHeight = ARMEDDIA1.fnRoundInches(nSmallestThighGiven * nThighCirRed) / 2 'half scale
		nHalfGroinHeight = (nGroinHeight / 2)
		nButtRadius = System.Math.Sqrt((nButtocksLength ^ 2) + (nHalfGroinHeight ^ 2))
		nButtCirCalc = nButtRadius + nHalfGroinHeight
		nButtBackSeam = (nButtocksCir - nButtCirCalc) * nButtBackSeamRatio
		If nButtBackSeam <= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "is still negative or zero!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Even with the modified reductions." + NL
			iFatalError = True
		ElseIf nButtBackSeam < 1.5 Then  'less than 3" full scale
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "is still less than 3 inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Even with the modified reductions." + NL
		End If
		
		'All other seams test
		nButtFrontSeam = (nButtocksCir - nButtCirCalc) * nButtFrontSeamRatio
		nButtCutOut = nButtCirCalc - (nButtBackSeam + nButtFrontSeam)
		If (nWaistCirGiven < nButtocksCirGiven) Then
			nWaistFrontSeam = System.Math.Abs(((nButtocksCirGiven - nWaistCirGiven) / 2) - nButtFrontSeam)
		Else
			nWaistFrontSeam = nButtFrontSeam + (Int(nWaistCirGiven - nButtocksCirGiven) * INCH1_8)
		End If
		If (nChestCirGiven < nButtocksCirGiven) Then
			nChestFrontSeam = System.Math.Abs(((nButtocksCirGiven - nChestCirGiven) / 2) - nButtFrontSeam)
		Else
			nChestFrontSeam = nButtFrontSeam + (Int(nChestCirGiven - nButtocksCirGiven) * INCH1_8)
		End If
		nWaistCutOut = (nButtCutOut * 0.87)
		nWaistCutOutModified = 0
		Do 
			nWaistRem = nWaistCir - (nWaistCutOut + (nWaistFrontSeam * 2))
			nChestRem = nChestCir - (nWaistCutOut + nWaistFrontSeam + nChestFrontSeam)
			If nWaistRem < 1.5 Or nChestRem < 1.5 Then 'less than 3" @ full scale
				If nWaistRem <= 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError + NL + "Severe - Warning!" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to Waist" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "is negative or zero!" + NL
					If nWaistCutOutModified > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sError = sError & "Even with the Waist CutOut decreased." + NL
					End If
					iFatalError = True
					Exit Do
				End If
				If nChestRem <= 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError + NL + "Severe - Warning!" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to Chest" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "is negative or zero!" + NL
					If nWaistCutOutModified > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sError = sError & "Even with the Waist CutOut decreased." + NL
					End If
					iFatalError = True
					Exit Do
				End If
				'Decreae Waist CutOut
				nWaistCutOut = nWaistCutOut - (2 * (1.5 - nWaistRem))
				nWaistCutOutModified = nWaistCutOutModified + 1
				If nWaistCutOut <= 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError + NL + "Severe - Warning!" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to Waist or Chest" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "is less than 3 inches!" + NL
					Select Case nWaistCutOutModified
						Case 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " time" & NL
							'sError = sError + "ref:- " + NL
						Case Is > 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " times" & NL
							'sError = sError + "ref:- " + NL
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "but is still negative or zero!" + NL
					iFatalError = True
					Exit Do
				End If
			Else
				If nWaistCutOutModified > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to Waist or Chest" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "was less than 3 inches!" + NL
					Select Case nWaistCutOutModified
						Case 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " time." & NL
							'sError = sError + "ref:- " + NL
						Case Is > 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " times." & NL
							'sError = sError + "ref:- " + NL
					End Select
				End If
				Exit Do 'greater than 3" - acceptable
			End If
		Loop 
		
		'Check that all of the reductions are within 5% of each other
		Dim nMaxRed, nMinRed, nCurrentValue As Short
		nMinRed = 0
		nMaxRed = 0
		
		'turn all reductions into integers for easy comparing
		'Waist
		nMinRed = (nWaistCirRed * 100)
		nMaxRed = nMinRed
		'Chest
		nCurrentValue = (nChestCirRed * 100)
		If nMinRed > nCurrentValue Then
			nMinRed = nCurrentValue
		Else
			If nMaxRed < nCurrentValue Then
				nMaxRed = nCurrentValue
			End If
		End If
		'Buttocks
		nCurrentValue = (nButtocksCirRed * 100)
		If nMinRed > nCurrentValue Then
			nMinRed = nCurrentValue
		Else
			If nMaxRed < nCurrentValue Then
				nMaxRed = nCurrentValue
			End If
		End If
		'Thigh
		nCurrentValue = (nThighCirRed * 100)
		If nMinRed > nCurrentValue Then
			nMinRed = nCurrentValue
		Else
			If nMaxRed < nCurrentValue Then
				nMaxRed = nCurrentValue
			End If
		End If
		
		If (nMaxRed - nMinRed) > 5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "All Reductions should be within 5% of each other" + NL
			'       iFatalError = True   'gg
		End If
		
		'Display Error message (if required) and return
		'There could be fatal errors found
		If Len(sError) > 0 Then
			If iFatalError = True Then
				MsgBox(sError, 64, "Warning - Problems with data")
				Fn_ValidateData = False
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError + NL + "The above problems have been found in the data do you" + NL
				sError = sError & "wish to continue ?"
				'UPGRADE_WARNING: Couldn't resolve default property of object iError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iError = MsgBox(sError, 52, "Severe Problems with data")
				'UPGRADE_WARNING: Couldn't resolve default property of object iError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If iError = IDYES Then
					Fn_ValidateData = True
				Else
					Fn_ValidateData = False
				End If
			End If
		Else
			Fn_ValidateData = True
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
		
		For ii = 0 To 4
			sString = sString & cboRed(ii).Text
		Next ii
		
		For ii = 0 To 6
			sString = sString & txtCir(ii).Text
		Next ii
		'Reason for skip is that 7 & 8 were taken out during development
		For ii = 9 To 16
			sString = sString & txtCir(ii).Text
		Next ii
		
		
		sString = sString & cboLeftCup.Text & txtLeftDisk.Text
		
		sString = sString & cboRightCup.Text & txtRightDisk.Text
		
		sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text
		
		sString = sString & cboFrontNeck.Text & txtFrontNeck.Text
		
		sString = sString & cboBackNeck.Text & txtBackNeck.Text
		
		sString = sString & cboClosure.Text
		
		sString = sString & cboFabric.Text
		
		'Extra Values
		
		sString = sString & cboLegStyle.Text
		
		'Index 0   =>  Panty
		'Index 1   =>  Brief
		If optLeftLeg(0).Checked = True Then
			sString = sString & "Panty"
		Else
			sString = sString & "Brief"
		End If
		If optRightLeg(0).Checked = True Then
			sString = sString & "Panty"
		Else
			sString = sString & "Brief"
		End If
		
		sString = sString & cboCrotchStyle.Text
		
		
		FN_ValuesString = sString
		
	End Function
	
	Private Function fnDisplayToInches(ByVal nDisplay As Double) As Double
		'This function takes the value given and converts it
		'into a decimal version in inches, rounded to the nearest eighth
		'of an inch.
		'
		'Input:-
		'        nDisplay is the value as input by the operator in the
		'        dialog.
		'        The convention is that, Metric dimensions use the decimal
		'        point to indicate the division between CMs and mms
		'        ie 7.6 = 7 cm and 6 mm.
		'        Whereas the decimal point for imperial measurements indicates
		'        the division between inches and eighths
		'        ie 7.6 = 7 inches and 6 eighths
		'Globals:-
		'        g_nUnitsFac = 1       => nDisplay in Inches
		'        g_nUnitsFac = 10/25.5 => nDisplay in CMs
		'Returns:-
		'        Single, Inches rounded to the nearest eighth (0.125)
		'        -1,     on conversion error.
		'
		'Errors:-
		'        The returned value is usually +ve. Unless it can't
		'        be sucessfully converted to inches.
		'        Eg 7.8 is an invalid number if g_nUnitsFac = 1
		'
		'                            WARNING
		'                            ~~~~~~~
		'In most cases the input is a +ve number.  This function will handle a
		'-ve number but in this case the error checking is invalid.  This
		'is done to provide a general conversion tool.  Where the input is
		'likley to be -ve then the calling subroutine or function should check
		'the sensibility of the returned value for that specific case.
		'
		
		Dim iInt, iSign As Short
		Dim nDec As Double
		'retain sign
		iSign = System.Math.Sign(nDisplay)
		nDisplay = System.Math.Abs(nDisplay)
		
		'Simple case where Units are CM
		If g_nUnitsFac <> 1 Then
			fnDisplayToInches = ARMEDDIA1.fnRoundInches(nDisplay * g_nUnitsFac) * iSign
			Exit Function
		End If
		
		'Imperial units
		iInt = Int(nDisplay)
		nDec = nDisplay - iInt
		
		'Check that conversion is possible (return -1 if not)
		If nDec > 0.8 Then
			fnDisplayToInches = -1
		Else
			fnDisplayToInches = ARMEDDIA1.fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
		End If
	End Function
	
	
	
	Private Function fnGetString(ByVal sString As String, ByRef iindex As Short) As String
		'Function to return as a string the iIndexth item in a string
		'that uses blanks (spaces) as delimiters.
		'EG
		'    sString = "Sam Spade Hello"
		'    fnGetNumber( sString, 2) = "Spade"
		'
		'If the iIndexth item is not found then return -1 to indicate an error.
		'This assumes that the string will not be used to store -ve numbers.
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
		sString = sString & " " 'Trailing blank as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iindex
			iPos = InStr(sString, " ")
			If ii = iindex Then
				sString = VB.Left(sString, iPos - 1)
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
	
	Private Function fnInchestoText(ByRef nInches As Double) As String
		'Function returns a decimal value in inches as a string
		'
		Dim nPrecision, nDec As Double
		Dim iInt, iEighths As Short
		Dim sString As String
		nPrecision = 0.125
		
		'Split into decimal parts
		iInt = Int(nInches)
		nDec = nInches - iInt
		If nDec <> 0 Then 'Avoid overflow
			iEighths = Int(nDec / nPrecision)
		Else
			iEighths = 0
		End If
		
		'Format string
		If iInt <> 0 Then
			sString = LTrim(Str(iInt))
		Else
			sString = "  "
		End If
		If iEighths <> 0 Then
			Select Case iEighths
				Case 2, 6
					sString = sString & "-" & LTrim(Str(iEighths / 2)) & "/4"
				Case 4
					sString = sString & "-" & "1/2"
				Case Else
					sString = sString & "-" & LTrim(Str(iEighths)) & "/8"
			End Select
		Else
			sString = sString & "   "
		End If
		
		'Return formatted string
		fnInchestoText = sString
		
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
		
		'Assign combo boxes (set the defaults if empty)
		'Reductions
		If txtCombo(0).Text <> "" Then cboRed(0).Text = txtCombo(0).Text Else cboRed(0).Text = ""
		If txtCombo(1).Text <> "" Then cboRed(1).Text = txtCombo(1).Text Else cboRed(1).Text = ""
		If txtCombo(2).Text <> "" Then cboRed(2).Text = txtCombo(2).Text Else cboRed(2).Text = ""
		If txtCombo(11).Text <> "" Then cboRed(3).Text = txtCombo(11).Text Else cboRed(3).Text = ""
		If txtCombo(12).Text <> "" Then cboRed(4).Text = txtCombo(12).Text Else cboRed(4).Text = ""
		
		'Bra cups and disk
		For ii = 0 To (cboLeftCup.Items.Count - 1)
			If txtCombo(3).Text = VB6.GetItemString(cboLeftCup, ii) Then
				g_sLeftCup = VB6.GetItemString(cboLeftCup, ii)
				cboLeftCup.SelectedIndex = ii
			End If
		Next ii
		For ii = 0 To (cboRightCup.Items.Count - 1)
			If txtCombo(4).Text = VB6.GetItemString(cboRightCup, ii) Then
				g_sRightCup = VB6.GetItemString(cboRightCup, ii)
				cboRightCup.SelectedIndex = ii
			End If
		Next ii
		
		g_sLeftDisk = txtLeftDisk.Text
		g_sRightDisk = txtRightDisk.Text
		
		If txtCombo(5).Text <> "" Then cboLeftAxilla.Text = txtCombo(5).Text Else cboLeftAxilla.SelectedIndex = 0
		If txtCombo(6).Text <> "" Then cboRightAxilla.Text = txtCombo(6).Text Else cboRightAxilla.SelectedIndex = 0
		If txtCombo(7).Text <> "" Then cboFrontNeck.Text = txtCombo(7).Text Else cboFrontNeck.SelectedIndex = 0
		If txtCombo(8).Text <> "" Then cboBackNeck.Text = txtCombo(8).Text Else cboBackNeck.SelectedIndex = 0
		If txtCombo(9).Text <> "" Then cboClosure.Text = txtCombo(9).Text Else cboClosure.SelectedIndex = 0
		cboFabric.Text = txtCombo(10).Text
		
		'Set up units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Display dimesions sizes in inches
		For ii = 0 To 6
			txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
		Next ii
		For ii = 9 To 16
			txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
		Next ii
		txtFrontNeck_Leave(txtFrontNeck, New System.EventArgs())
		txtBackNeck_Leave(txtBackNeck, New System.EventArgs())
		
		'Store values used to check for mods to bra cups
		g_nChest = Val(txtCir(5).Text)
		g_nUnderBreast = Val(txtCir(10).Text)
		g_nNipple = Val(txtCir(11).Text)
		
		'Store values to use on change etc
		Dim sChar As New VB6.FixedLengthString(1)
		g_sBackNeck = txtBackNeck.Text
		sChar.Value = VB.Left(cboBackNeck.Text, 1)
		If sChar.Value = "R" Or sChar.Value = "S" Or sChar.Value = "V" Then
			'Disable length box for Regular, Scoop, V neck
			txtBackNeck.Enabled = False
			labBackNeck.Enabled = False
		Else
			txtBackNeck.Enabled = True
			labBackNeck.Enabled = True
		End If
		
		g_sFrontNeck = txtFrontNeck.Text
		sChar.Value = VB.Left(cboFrontNeck.Text, 1)
		If sChar.Value = "R" Or sChar.Value = "S" Then
			txtFrontNeck.Enabled = False
			labFrontNeck.Enabled = False
		Else
			labFrontNeck.Enabled = True
			txtFrontNeck.Enabled = True
		End If
		
		'Set dropdown combo boxes
		'Crotch Style
		cboCrotchStyle.Items.Add("Open Crotch") 'Both
		cboCrotchStyle.Items.Add("Horizontal Fly") 'Male only
		cboCrotchStyle.Items.Add("Diagonal Fly") 'Male only
		cboCrotchStyle.Items.Add("Snap Crotch") 'Both
		cboCrotchStyle.Items.Add("Gusset") 'Both
		
		'Set value
		For ii = 0 To (cboCrotchStyle.Items.Count - 1)
			If VB6.GetItemString(cboCrotchStyle, ii) = txtCrotchStyle.Text Then
				cboCrotchStyle.SelectedIndex = ii
			End If
		Next ii
		
		'Leg Style
		cboLegStyle.Items.Add("Panty")
		cboLegStyle.Items.Add("Brief")
		cboLegStyle.Items.Add("Brief-French")
		
		'Set values of combo and option buttons
		g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
		For ii = 0 To (cboLegStyle.Items.Count - 1)
			If VB6.GetItemString(cboLegStyle, ii) = ARMDIA1.fnGetString(txtLegStyle.Text, 1) Then
				cboLegStyle.SelectedIndex = ii
			End If
		Next ii
		
		'Set this now to avoid problems with cboLegStyle_Click
		If cboLegStyle.SelectedIndex = -1 Then
			g_sPreviousLegStyle = "None"
		Else
			g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		End If
		
		If ARMDIA1.fnGetString(txtLegStyle.Text, 2) <> "" Then
			If ARMDIA1.fnGetString(txtLegStyle.Text, 2) = "Panty" Then
				optLeftLeg(0).Checked = True
			Else
				optLeftLeg(1).Checked = True 'Brief
			End If
		End If
		If ARMDIA1.fnGetString(txtLegStyle.Text, 3) <> "" Then
			If ARMDIA1.fnGetString(txtLegStyle.Text, 3) = "Panty" Then
				optRightLeg(0).Checked = True
			Else
				optRightLeg(1).Checked = True 'Brief
			End If
		End If
		
		'Save the values in the text to a string
		'this can then be used to check if they have changed
		'on use of the close button
		'    g_sChangeChecker = FN_ValuesString()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer back to normal.
		
		'Disable bras for Males
		If txtSex.Text = "Male" Then
			frmBra.Enabled = False
			For ii = 0 To 11
				LabBra(ii).Enabled = False
			Next ii
			cboLeftCup.Enabled = False
			txtLeftDisk.Enabled = False
			cboRightCup.Enabled = False
			txtRightDisk.Enabled = False
			cboRed(2).Enabled = False
			cmdCalculateBraDisks.Enabled = False
		Else
			PR_EnableCalculateDiskButton()
		End If
		Show()
		g_sChangeChecker = FN_ValuesString()
		
	End Sub
	
	Private Sub bodysuit_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		
		Hide()
		
		MainForm = Me
		
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
			MsgBox("BodySuit Dialogue is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting BodySuit - Dialogue")
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		
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
		g_sPathJOBST = fnPathJOBST()
		
		'Clear fields
		'Circumferences and lengths
		For ii = 0 To 6
			txtCir(ii).Text = ""
		Next ii
		For ii = 9 To 16
			txtCir(ii).Text = ""
		Next ii
		
		'The data from these DDE text boxes is copied
		'to the combo boxes on Link close
		'
		For ii = 0 To 12
			txtCombo(ii).Text = ""
		Next ii
		
		'Bra cups
		txtLeftDisk.Text = ""
		txtRightDisk.Text = ""
		
		'Patient details
		txtFileNo.Text = ""
		txtUnits.Text = ""
		txtPatientName.Text = ""
		txtDiagnosis.Text = ""
		txtAge.Text = ""
		txtSex.Text = ""
		txtWorkOrder.Text = ""
		
		'Design choices
		txtFrontNeck.Text = ""
		txtBackNeck.Text = ""
		
		'UID of symbols
		txtUidMPD.Text = ""
		txtUidVB.Text = ""
		
		'Setup combo box fields
		'Waist Default 85% +/- 5%
		cboRed(0).Items.Add("0.90")
		cboRed(0).Items.Add("0.85") 'gg
		cboRed(0).Items.Add("0.80")
		cboRed(0).Items.Add("")
		
		'Chest Default 90% +/- 5%
		cboRed(1).Items.Add("0.95")
		cboRed(1).Items.Add("0.90") 'gg
		cboRed(1).Items.Add("0.85")
		cboRed(1).Items.Add("")
		
		'Under Breast 90% +/- 5%
		cboRed(2).Items.Add("0.95")
		cboRed(2).Items.Add("0.90") 'gg
		cboRed(2).Items.Add("0.85")
		cboRed(2).Items.Add("")
		
		'Large part of Buttocks 85% +/- 5%
		cboRed(3).Items.Add("0.90")
		cboRed(3).Items.Add("0.85") 'gg
		cboRed(3).Items.Add("0.80")
		cboRed(3).Items.Add("")
		
		'Left Thigh 85% +/- 5%
		cboRed(4).Items.Add("0.90")
		cboRed(4).Items.Add("0.85") 'gg
		cboRed(4).Items.Add("0.80")
		cboRed(4).Items.Add("")
		'Right Thigh uses Left Thigh cboRed value
		
		cboLeftCup.Items.Add("Training")
		cboLeftCup.Items.Add("A")
		cboLeftCup.Items.Add("B")
		cboLeftCup.Items.Add("C")
		cboLeftCup.Items.Add("D")
		cboLeftCup.Items.Add("E")
		cboLeftCup.Items.Add("None")
		cboLeftCup.Items.Add("")
		
		cboRightCup.Items.Add("Training")
		cboRightCup.Items.Add("A")
		cboRightCup.Items.Add("B")
		cboRightCup.Items.Add("C")
		cboRightCup.Items.Add("D")
		cboRightCup.Items.Add("E")
		cboRightCup.Items.Add("None")
		cboRightCup.Items.Add("")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboLeftAxilla.Items.Add("Regular 2" & QQ)
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboLeftAxilla.Items.Add("Regular 1½" & QQ)
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboLeftAxilla.Items.Add("Regular 2½" & QQ)
		cboLeftAxilla.Items.Add("Open")
		cboLeftAxilla.Items.Add("Mesh")
		cboLeftAxilla.Items.Add("Lining")
		cboLeftAxilla.Items.Add("Sleeveless")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboRightAxilla.Items.Add("Regular 2" & QQ)
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboRightAxilla.Items.Add("Regular 1½" & QQ)
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboRightAxilla.Items.Add("Regular 2½" & QQ)
		cboRightAxilla.Items.Add("Open")
		cboRightAxilla.Items.Add("Mesh")
		cboRightAxilla.Items.Add("Lining")
		cboRightAxilla.Items.Add("Sleeveless")
		
		cboFrontNeck.Items.Add("Regular")
		cboFrontNeck.Items.Add("Scoop")
		cboFrontNeck.Items.Add("Measured Scoop")
		cboFrontNeck.Items.Add("V neck")
		cboFrontNeck.Items.Add("Measured V neck")
		cboFrontNeck.Items.Add("Turtle")
		cboFrontNeck.Items.Add("Turtle - Fabric same as Vest")
		cboFrontNeck.Items.Add("Turtle Detachable")
		cboFrontNeck.Items.Add("Turtle Detach. Fabric")
		
		cboBackNeck.Items.Add("Regular")
		cboBackNeck.Items.Add("Scoop")
		cboBackNeck.Items.Add("Measured Scoop")
		
		cboClosure.Items.Add("Front Zip")
		cboClosure.Items.Add("Back Zip")
		
		LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\FABRIC.DAT")
		
		
	End Sub
	
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
	
	'UPGRADE_WARNING: Event optLeftLeg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optLeftLeg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optLeftLeg.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optLeftLeg.GetIndex(eventSender)
			
			If VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex) = "Panty" Then
				Select Case Index
					Case 0 'Panty
						'Enable Right side
						optRightLeg(1).Enabled = True
						
					Case 1 'Brief
						'Disable Right side
						optRightLeg(1).Enabled = False
				End Select
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRightLeg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRightLeg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRightLeg.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optRightLeg.GetIndex(eventSender)
			If VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex) = "Panty" Then
				Select Case Index
					Case 0 'Panty
						'Enable left side
						optLeftLeg(1).Enabled = True
						
					Case 1 'Brief
						'Disable Left side
						optLeftLeg(1).Enabled = False
				End Select
			End If
			
		End If
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a body then Define the DATA Base
		'fields for the BodySuit and insert the BODYSUIT symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hBody;")
		
		PR_UpdateDB()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	
	Private Sub PR_DoBraCupsAndDisks()
		'Procedure to either
		'    1. Check if there is enough data to enable a
		'       calculation.
		' or
		'    2. Check if a change has been made to the data
		'       requiring a recalculation.
		'
		'Then calculate the bra cups and disks.
		'
		'NOTE
		'    txtCir(5) == Chest Circumference
		'    txtCir(10) == Under Breast Circumference
		'    txtCir(11) == Circumference over nipple line
		'
		'GLOBALS
		'    g_nUnitsFac As Double
		'    g_nUnderBreast As Double
		'    g_nChest As Double
		'    g_nNipple As Double
		'    g_sLeftCup As String
		'    g_sRightCup As String
		'    g_sLeftDisk As String
		'    g_sRightDisk As String
		
		Dim sLeftCup, sRightCup As String
		Dim sSide As Short
		Dim ii, iError As Short
		Dim nUnderBreast, nNipple, nChest As Double
		Dim sError As String
		Dim iLeftDisk, iRightDisk As Short
		
		sLeftCup = cboLeftCup.Text
		sRightCup = cboRightCup.Text
		
		'Check if enough data to caclulate
		'Exit sub if not
		If Val(txtCir(5).Text) = 0 Then
			'No chest cir.
			Exit Sub
		ElseIf Val(txtCir(10).Text) = 0 Then 
			'No under breast cir.
			Exit Sub
		ElseIf Val(txtCir(11).Text) = 0 And ((sLeftCup = "" Or sLeftCup = "None") And (sRightCup = "" Or sRightCup = "None")) Then 
			'No over nipple circumference and no bra cups given for either side
			Exit Sub
		End If
		
		'Check if changed
		Dim Changed As Short
		
		Changed = False
		
		If sLeftCup <> g_sLeftCup Then
			Changed = True
			g_sLeftCup = sLeftCup
		End If
		
		If sRightCup <> g_sRightCup Then
			Changed = True
			g_sRightCup = sRightCup
		End If
		
		If Val(txtCir(10).Text) <> g_nUnderBreast Then
			Changed = True
			g_nUnderBreast = Val(txtCir(10).Text)
		End If
		
		If Val(txtCir(11).Text) <> g_nNipple Then
			Changed = True
			g_nNipple = Val(txtCir(11).Text)
			'Force a recalculation of cup
			If sLeftCup <> "None" Then sLeftCup = ""
			If sRightCup <> "None" Then sRightCup = ""
		End If
		
		If Val(txtCir(5).Text) <> g_nChest Then
			g_nChest = Val(txtCir(5).Text)
			'Force a recalculation of cup if a nipple circum. is given
			'Changed chest is only significant if a nipple line measurement
			'is given from which the Cup size is calculated
			If g_nNipple <> 0 Then
				Changed = True
				If sLeftCup <> "None" Then sLeftCup = ""
				If sRightCup <> "None" Then sRightCup = ""
			End If
		End If
		
		'Force a calculation if either disk is empty
		If txtLeftDisk.Text = "" Or txtRightDisk.Text = "" Then Changed = True
		
		'Force a check if disk has been changed
		If txtLeftDisk.Text <> g_sLeftDisk Or txtRightDisk.Text <> g_sRightDisk Then
			Changed = True
			iRightDisk = Val(txtRightDisk.Text)
			iLeftDisk = Val(txtLeftDisk.Text)
		End If
		
		'If no modifications then exit
		If Changed <> True Then Exit Sub
		
		
		'Recalculate if Changed
		'Convert to inches
		nNipple = ARMEDDIA1.fnDisplayToInches(g_nNipple)
		nUnderBreast = ARMEDDIA1.fnDisplayToInches(g_nUnderBreast)
		nChest = ARMEDDIA1.fnDisplayToInches(g_nChest)
		
		'Right cup
		If sRightCup <> "" Or nNipple <> 0 Then
			iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sRightCup, iRightDisk, "Right")
			If iError Then
				For ii = 0 To (cboRightCup.Items.Count - 1)
					If VB6.GetItemString(cboRightCup, ii) = sRightCup Then cboRightCup.SelectedIndex = ii
				Next ii
				If iRightDisk = 0 Then txtRightDisk.Text = "" Else txtRightDisk.Text = CStr(iRightDisk)
				g_sRightDisk = Trim(Str(iRightDisk))
				g_sRightCup = sRightCup
			Else
				txtRightDisk.Text = ""
				cboRightCup.SelectedIndex = -1
				g_sRightDisk = ""
				g_sRightCup = ""
			End If
		End If
		
		If sLeftCup <> "" Or nNipple <> 0 Then
			iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sLeftCup, iLeftDisk, "Left")
			If iError Then
				For ii = 0 To (cboLeftCup.Items.Count - 1)
					If VB6.GetItemString(cboLeftCup, ii) = sLeftCup Then cboLeftCup.SelectedIndex = ii
				Next ii
				If iLeftDisk <= 0 Then txtLeftDisk.Text = "" Else txtLeftDisk.Text = CStr(iLeftDisk)
				g_sLeftDisk = Trim(Str(iLeftDisk))
				g_sLeftCup = sLeftCup
			Else
				txtLeftDisk.Text = ""
				cboLeftCup.SelectedIndex = -1
				g_sLeftDisk = ""
				g_sLeftCup = ""
			End If
		End If
		
		PR_EnableCalculateDiskButton()
		
		
	End Sub
	
	Private Sub PR_EnableCalculateDiskButton()
		'Procedure to check if there is enough data to
		'enable the caclculate disks command button
		'
		Dim sLeftCup, sRightCup As String
		
		sLeftCup = cboLeftCup.Text
		sRightCup = cboRightCup.Text
		
		'Check if enough data to caclulate
		'Exit sub if not
		If Val(txtCir(5).Text) = 0 Then
			'No chest cir.
			cmdCalculateBraDisks.Enabled = False
			Exit Sub
		ElseIf Val(txtCir(10).Text) = 0 Then 
			'No under breast cir.
			cmdCalculateBraDisks.Enabled = False
			Exit Sub
		ElseIf Val(txtCir(11).Text) = 0 And (sLeftCup = "" Or sLeftCup = "None") And (sRightCup = "" Or sRightCup = "None") Then 
			'No over nipple circumference and no bra cups given for either side
			cmdCalculateBraDisks.Enabled = False
			Exit Sub
		End If
		
		cmdCalculateBraDisks.Enabled = True
		
	End Sub
	
	Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.Control, ByRef sFileName As String)
		'General procedure to create the list section of
		'a combo box reading the data from a file
		
		Dim sLine As String
		Dim fFileNum As Short
		
		fFileNum = FreeFile
		
		If FileLen(sFileName) = 0 Then
			MsgBox(sFileName & "Not found", 48, "BodySuit Dialogue")
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
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
		
	End Sub
	
	Private Sub PR_PutNumberAssign(ByRef sVariableName As String, ByRef nAssignedNumber As Object)
		'Procedure to put a number assignment
		'Adds a semi-colon
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAssignedNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & Str(nAssignedNumber) & ";")
		
	End Sub
	
	Private Sub PR_PutStringAssign(ByRef sVariableName As String, ByRef sAssignedString As Object)
		'Procedure to put a string assignment
		'Encloses String in quotes and adds a semi-colon
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & QQ & sAssignedString & QQ & ";")
		
	End Sub
	
	Private Sub PR_Select_Text(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelStart = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
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
	
	Private Sub PR_UpdateDB()
		'Procedure called from
		'    PR_CreateMacro_Save
		'and
		'    PR_CreateMacro_Data
		'
		'Used to stop duplication on code
		
		Dim sSymbol As String
		Dim sLeftleg, sRightleg As String
		
		sSymbol = "suitbody"
		
		If txtUidVB.Text = "" Then
			'Define DB Fields
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\BODY\BFIELDS.D;")
			
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
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LtSCir" & QCQ & txtCir(0).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "RtSCir" & QCQ & txtCir(1).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "NeckCir" & QCQ & txtCir(2).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "SWidth" & QCQ & txtCir(3).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "S_Waist" & QCQ & txtCir(4).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ChestCir" & QCQ & txtCir(5).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtCir(6).Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "S_Breast" & QCQ & txtCir(9).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BreastCir" & QCQ & txtCir(10).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "NippleCir" & QCQ & txtCir(11).Text & QQ & ");")
		
		'Extra Values
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "SFButt" & QCQ & txtCir(12).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "SLgButt" & QCQ & txtCir(13).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LgButtCir" & QCQ & txtCir(14).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LtThCir" & QCQ & txtCir(15).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "RtThCir" & QCQ & txtCir(16).Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BraLtCup" & QCQ & cboLeftCup.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BraRtCup" & QCQ & cboRightCup.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BraLtDisk" & QCQ & txtLeftDisk.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BraRtDisk" & QCQ & txtRightDisk.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LtAxillaType" & QCQ & FN_EscapeQuotesInString(cboLeftAxilla) & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "RtAxillaType" & QCQ & FN_EscapeQuotesInString(cboRightAxilla) & QQ & ");")
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "NeckType" & QCQ & cboFrontNeck.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "NeckDimension" & QCQ & txtFrontNeck.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BackNeckType" & QCQ & cboBackNeck.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BackNeckDim" & QCQ & txtBackNeck.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Closure" & QCQ & cboClosure.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Fabric" & QCQ & cboFabric.Text & QQ & ");")
		
		'Extra Values
		'Index 0   =>  Panty
		'Index 1   =>  Brief
		If optLeftLeg(0).Checked = True Then
			sLeftleg = "Panty"
		Else
			sLeftleg = "Brief"
		End If
		If optRightLeg(0).Checked = True Then
			sRightleg = "Panty"
		Else
			sRightleg = "Brief"
		End If
		'txtLegStyle = cboLegStyle.Text & " " & sLeftleg & " " & sRightleg & " " & Str$(nLargestCir)
		txtLegStyle.Text = cboLegStyle.Text & " " & sLeftleg & " " & sRightleg
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LegStyle" & QCQ & txtLegStyle.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "CrotchStyle" & QCQ & cboCrotchStyle.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistCirUserFac" & QCQ & cboRed(0).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ChestCirUserFac" & QCQ & cboRed(1).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BreastCirUserFac" & QCQ & cboRed(2).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "LgButtCirUserFac" & QCQ & cboRed(3).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ThCirUserFac" & QCQ & cboRed(4).Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		
	End Sub
	
	Private Function round(ByVal nNumber As Double) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35)  = 1
		'    round(1.55)  = 2
		'    round(2.50)  = 3
		'    round(-2.50) = -3
		'    round(0)     = 0
		'
		
		Dim iInt, iSign As Short
		
		'Avoid extra work. Return 0 if input is 0
		If nNumber = 0 Then
			round = 0
			Exit Function
		End If
		
		'Split input
		iSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		iInt = Int(nNumber)
		
		'Effect rounding
		If (nNumber - iInt) >= 0.5 Then
			round = (iInt + 1) * iSign
		Else
			round = iInt * iSign
		End If
		
	End Function
	
	Private Sub Select_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelStart = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'It is assumed that the link open from Drafix has failed
		'Therefore we "End" here
		End
	End Sub
	
	Private Sub txtBackNeck_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBackNeck.Enter
		CADGLOVE1.PR_Select_Text(txtBackNeck)
	End Sub
	
	Private Sub txtBackNeck_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBackNeck.Leave
		Dim nLen As Double
		nLen = CADGLOVE1.FN_InchesValue(txtBackNeck)
		If nLen >= 0 Then
			lblBackNeck.Text = ARMEDDIA1.fnInchestoText(nLen)
			g_sBackNeck = txtBackNeck.Text
		Else
			lblBackNeck.Text = ""
		End If
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
			lblCir(Index).Text = ARMEDDIA1.fnInchestoText(nLen)
		Else
			lblCir(Index).Text = ""
		End If
		
		If Index > 10 Or Index = 5 Then PR_EnableCalculateDiskButton()
		
		
	End Sub
	
	Private Sub txtFrontNeck_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontNeck.Enter
		CADGLOVE1.PR_Select_Text(txtFrontNeck)
	End Sub
	
	Private Sub txtFrontNeck_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontNeck.Leave
		Dim nLen As Double
		nLen = CADGLOVE1.FN_InchesValue(txtFrontNeck)
		If nLen >= 0 Then
			lblFrontNeck.Text = ARMEDDIA1.fnInchestoText(nLen)
			g_sFrontNeck = txtFrontNeck.Text 'Keep this for restore
		Else
			lblFrontNeck.Text = ""
		End If
	End Sub
	
	Private Sub txtLeftDisk_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftDisk.Enter
		CADGLOVE1.PR_Select_Text(txtLeftDisk)
	End Sub
	
	Private Sub txtLeftDisk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLeftDisk.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'Only Disks sizes 1 to 9 allowed
		Select Case KeyAscii
			Case 32
				KeyAscii = 0
			Case 8, 9, 10, 13, 49 To 57
				'Do nothing
			Case Else
				KeyAscii = 0
				MsgBox("Bra disk sizes 1 to 9 only.", 48, "BodySuit - Dialogue")
		End Select
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtRightDisk_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightDisk.Enter
		CADGLOVE1.PR_Select_Text(txtRightDisk)
	End Sub
	
	Private Sub txtRightDisk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRightDisk.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'Only Disks sizes 1 to 9 allowed
		Select Case KeyAscii
			Case 32
				KeyAscii = 0
			Case 8, 9, 10, 13, 49 To 57
				'Do nothing
			Case Else
				KeyAscii = 0
				MsgBox("Bra disk sizes 1 to 9 only.", 48, "BodySuit - Dialogue")
		End Select
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub Validate_Text(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'Subroutine that is activated when the focus is lost.
		Dim rTextBoxValue As Double
		Dim ii, iTest As Short
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = ARMEDDIA1.fnDisplayToInches(Val(Text_Box_Name.Text))
		
		'Only if relevant units are Inches
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
		End If
		
		If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
			MsgBox("Zero value given", 48, "Data input Error")
			Text_Box_Name.Focus()
		End If
		
		If rTextBoxValue > 999.9 Then
			MsgBox("Given value too Large", 48, "Data input Error")
			Text_Box_Name.Focus()
		End If
		
	End Sub
	
	Private Sub Validate_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control, ByRef Label_Name As System.Windows.Forms.Control)
		'Subroutine that is activated when the focus is lost.
		Dim rTextBoxValue, rDec As Double
		Dim ii, iTest As Short
		Dim iEighths, iInt As Short
		Dim sString As String
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = ARMEDDIA1.fnDisplayToInches(Val(Text_Box_Name.Text))
		
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
		
		Label_Name.Text = ARMEDDIA1.fnInchestoText(rTextBoxValue)
		
	End Sub
End Class
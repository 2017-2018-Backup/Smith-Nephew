Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class vestdia
	Inherits System.Windows.Forms.Form
	'Project:   VESTDIA
	'Purpose:   Vest Dialogue
	'
	'
	'Version:   3.01
	'Date:      4.Oct.95
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
						MsgBox("Can't find a Drafix Drawing to update!", 16, "VEST Body - Dialogue")
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
	
	'UPGRADE_WARNING: Event cboFrontNeck.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFrontNeck_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFrontNeck.SelectedIndexChanged
		Select Case VB.Left(cboFrontNeck.Text, 1)
			Case "R", "S" 'Regular or Scoop
				txtFrontNeck.Text = ""
				lblFrontNeck.Text = ""
				txtFrontNeck.Enabled = False
				labFrontNeck.Enabled = False
			Case Else 'Measured Scoop or Turtle neck
				If g_sFrontNeck <> "" Then
					txtFrontNeck.Text = g_sFrontNeck
					txtFrontNeck_Leave(txtFrontNeck, New System.EventArgs()) 'Display inches
				End If
				txtFrontNeck.Enabled = True
				labFrontNeck.Enabled = True
		End Select
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
	
	'UPGRADE_WARNING: Event cboRightCup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboRightCup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboRightCup.SelectedIndexChanged
		If Me.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
			txtRightDisk.Text = ""
			g_sRightCup = ""
			g_sRightDisk = ""
		Else
			PR_EnableCalculateDiskButton()
			'PR_DoBraCupsAndDisks
		End If
	End Sub
	
	Private Sub cmdCalculateBraDisks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalculateBraDisks.Click
		PR_DoBraCupsAndDisks()
	End Sub
	
	Private Sub cmdTab_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTab.Click
		'Allows the user to use enter as a tab
		System.Windows.Forms.SendKeys.Send("{TAB}")
	End Sub
	
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
				MsgBox(sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")")
				Exit Function
			End If
			
			nDiff = nNipple - nChest
			If nDiff < 0 Then
				FN_CalculateBra = False
				sError = "Circumference over Nipple line is smaller than Chest circumference"
				MsgBox(sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")")
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
			MsgBox(sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")")
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
				MsgBox(sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")")
			End If
		Else
			iDisk = nSelctedDisk
		End If
		
		FN_CalculateBra = True
		
		
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
		PrintLine(fNum, "//by Visual Basic, VEST Body")
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
		
		Dim sCircum(11) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(0) = "Left shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(1) = "Right shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(2) = "Neck circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(3) = "Shoulder width"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(4) = "Shoulder to waist"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(5) = "Chest circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(6) = "Waist circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(7) = "Shoulder to EOS"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(8) = "Circ. at EOS"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(9) = "Shoulder to under breast"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(10) = "Circ. under breast"
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(11) = "Circ. over nipple"
		
		'Vest measurements (all must be present)
		For ii = 0 To 6
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
		
		'Bra Cups
		'Note:
		'    The Circumference over nipple is optional unless a cup has been
		'    specified
		'
		If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing dimension for " & sCircum(9) & "!" & NL
		End If
		If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing dimension for " & sCircum(10) & "!" & NL
		End If
		
		If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated" & "!" & NL
		End If
		
		If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra disks given! But missing dimension for " & sCircum(9) & "!" & NL
		End If
		
		'If cups or dimensions given then a disk must be present
		'NB
		'    cboXXXXCup.ListIndex = 6 = "None"
		'    cboXXXXCup.ListIndex = 7 = ""
		
		If cboLeftCup.SelectedIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Left BRA Cup" & "!" & NL
		End If
		If cboRightCup.SelectedIndex < 6 And Val(txtRightDisk.Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Right BRA Cup" & "!" & NL
		End If
		
		'Sex error
		If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.SelectedIndex < 0 Or cboRightCup.SelectedIndex < 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male patient but Bra Measurements or Bra Cups requested " & "!" & NL
		End If
		
		'Neck at back and front
		Dim sChar As New VB6.FixedLengthString(1)
		sChar.Value = VB.Left(cboBackNeck.Text, 1)
		If sChar.Value = "M" And txtBackNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Back Meck Measured Scoop! " & NL
		End If
		
		sChar.Value = VB.Left(cboFrontNeck.Text, 1)
		If sChar.Value = "M" And txtFrontNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Front Neck Measured Scoop! " & NL
		End If
		
		'
		If cboLeftAxilla.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Left Axilla not given! " & NL
		End If
		
		If cboRightAxilla.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Right Axilla not given! " & NL
		End If
		
		If cboFrontNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Neck not given! " & NL
		End If
		
		If cboBackNeck.Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Back neck not given! " & NL
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
			MsgBox(sError, 16, "VEST Body - Dialogue")
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
		
		For ii = 0 To 11
			sString = sString & txtCir(ii).Text
		Next ii
		
		For ii = 0 To 2
			sString = sString & cboRed(ii).Text
		Next ii
		
		sString = sString & cboLeftCup.Text & txtLeftDisk.Text
		
		sString = sString & cboRightCup.Text & txtRightDisk.Text
		
		sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text
		
		sString = sString & cboFrontNeck.Text & txtFrontNeck.Text
		
		sString = sString & cboBackNeck.Text & txtBackNeck.Text
		
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
		
		'Assign combo boxes (set the defaults if empty)
		'Reductions
		If txtCombo(0).Text <> "" Then cboRed(0).Text = txtCombo(0).Text Else cboRed(0).Text = ""
		If txtCombo(1).Text <> "" Then cboRed(1).Text = txtCombo(1).Text Else cboRed(1).Text = ""
		If txtCombo(2).Text <> "" Then cboRed(2).Text = txtCombo(2).Text Else cboRed(2).Text = ""
		
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
		For ii = 0 To 11
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
		If sChar.Value = "R" Or sChar.Value = "S" Then
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
		
		'Save the values in the text to a string
		'this can then be used to check if they have changed
		'on use of the close button
		g_sChangeChecker = FN_ValuesString()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to hourglass.
		
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
	End Sub
	
	Private Sub vestdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
			MsgBox("VEST Body Dialogue is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting VEST Body - Dialogue")
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
		For ii = 0 To 11
			txtCir(ii).Text = ""
		Next ii
		
		'The data from these DDE text boxes is copied
		'to the combo boxes on Link close
		'
		For ii = 0 To 10
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
		cboRed(0).Items.Add("0.95")
		cboRed(0).Items.Add("0.85")
		cboRed(0).Items.Add("")
		
		cboRed(1).Items.Add("0.95")
		cboRed(1).Items.Add("0.85")
		cboRed(1).Items.Add("")
		
		cboRed(2).Items.Add("0.95")
		cboRed(2).Items.Add("0.85")
		cboRed(2).Items.Add("")
		
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
		cboFrontNeck.Items.Add("Turtle")
		cboFrontNeck.Items.Add("Turtle - Fabric same as Vest")
		cboFrontNeck.Items.Add("Turtle Detachable")
		cboFrontNeck.Items.Add("Turtle Detach. Fabric")
		
		cboBackNeck.Items.Add("Regular")
		cboBackNeck.Items.Add("Scoop")
		cboBackNeck.Items.Add("Measured Scoop")
		
		cboClosure.Items.Add("Velcro")
		cboClosure.Items.Add("Zip")
		cboClosure.Items.Add("Front Velcro")
		cboClosure.Items.Add("Front Velcro (Reversed)")
		cboClosure.Items.Add("Back Velcro")
		cboClosure.Items.Add("Back Velcro (Reversed)")
		cboClosure.Items.Add("Front Zip")
		cboClosure.Items.Add("Back Zip")
		
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
			
			PR_CreateMacro_Data("c:\jobst\draw.d")
			PR_CreateMacro_Axilla("c:\jobst\draw_1.d", "c:\jobst\draw_3.d")
			PR_CreateMacro_Bra("c:\jobst\draw_2.d")
			
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@" & g_PathJOBST & "\VEST\BODY.D{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				End
			Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, "CAD Glove Dialogue")
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub PR_CreateMacro_Axilla(ByRef sfile As String, ByRef sFile2 As String)
		'Assumes that data has been validated befor this
		'procedure is called.
		
		g_bMeshAxilla = False
		
		Dim sRightAxilla, sLeftAxilla, sAxilla As String
		Dim iLoop, ii As Short
		
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sfile, "Axilla Drawing MACRO/S", (txtPatientName.Text), (txtFileNo.Text))
		
		sLeftAxilla = VB.Left(cboLeftAxilla.Text, 1)
		sRightAxilla = VB.Left(cboRightAxilla.Text, 1)
		
		If sLeftAxilla <> sRightAxilla Then iLoop = 2 Else iLoop = 1
		
		For ii = 1 To iLoop
			If ii = 1 Then
				sAxilla = sLeftAxilla
			Else
				sAxilla = sRightAxilla
			End If
			Select Case sAxilla
				Case "R"
					'Regular Axilla
					BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\BODREGUL.D;")
				Case "O", "L"
					'Open or Lining Axilla
					BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\BODOTHRS.D;")
				Case "M"
					'Mesh Axilla
					g_bMeshAxilla = True
					BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\BODMESH.D;")
				Case "S"
					'Sleeveless Axilla
					BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\BODSLESS.D;")
			End Select
		Next ii
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sFile2, "Draw mesh", (txtPatientName.Text), (txtFileNo.Text))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If g_bMeshAxilla Then PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshvest" & QCQ & "normal" & QQ & " );")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CreateMacro_Bra(ByRef sfile As String)
		'Assumes that data has been validated before this
		'procedure is called.
		Dim nDiskXoff, nBraCLOffset, nDiskYoff As Double
		Dim iLeftDisk, ii, iDisk, iRightDisk As Short
		
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sfile, "Load Bradisks", (txtPatientName.Text), (txtFileNo.Text))
		
		If cboLeftCup.SelectedIndex < 7 Or cboLeftCup.SelectedIndex < 7 Or txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "" Then
			'A cup or a disk of some sort has been given
			If cboLeftCup.Text = "None" Then iLeftDisk = -1 Else iLeftDisk = Val(txtLeftDisk.Text)
			If cboRightCup.Text = "None" Then iRightDisk = -1 Else iRightDisk = Val(txtRightDisk.Text)
			
			'Loop through both cups and give positioning information
			For ii = 1 To 2
				If ii = 1 Then iDisk = iLeftDisk Else iDisk = iRightDisk
				Select Case iDisk
					Case -1, 0
						nBraCLOffset = 0
						nDiskXoff = 0
						nDiskYoff = 0
					Case 1
						nBraCLOffset = 1.25
						nDiskXoff = 1.45
						nDiskYoff = 1.797
					Case 2
						nBraCLOffset = 1.125
						nDiskXoff = 1.652
						nDiskYoff = 2.029
					Case 3
						nBraCLOffset = 1#
						nDiskXoff = 1.893
						nDiskYoff = 2.288
					Case 4
						nBraCLOffset = 0.875
						nDiskXoff = 2.175
						nDiskYoff = 2.572
					Case 5
						nBraCLOffset = 0.75
						nDiskXoff = 2.293
						nDiskYoff = 2.882
					Case 6
						nBraCLOffset = 0.625
						nDiskXoff = 2.625
						nDiskYoff = 3.025
					Case 7
						nBraCLOffset = 0.5
						nDiskXoff = 2.94
						nDiskYoff = 3.335
					Case 8
						nBraCLOffset = 0.5
						nDiskXoff = 3.139
						nDiskYoff = 3.518
					Case 9
						nBraCLOffset = 0.5
						nDiskXoff = 3.393
						nDiskYoff = 3.938
				End Select
				
				If ii = 1 Then
					'Left Disks
					BDYUTILS.PR_PutNumberAssign("nDiskLt", iLeftDisk)
					BDYUTILS.PR_PutNumberAssign("nBraCLOffsetLt", nBraCLOffset)
					BDYUTILS.PR_PutNumberAssign("nDiskXoffLt", nDiskXoff)
					BDYUTILS.PR_PutNumberAssign("nDiskYoffLt", nDiskYoff)
				Else
					'Right Disks
					BDYUTILS.PR_PutNumberAssign("nDiskRt", iRightDisk)
					BDYUTILS.PR_PutNumberAssign("nBraCLOffsetRt", nBraCLOffset)
					BDYUTILS.PR_PutNumberAssign("nDiskXoffRt", nDiskXoff)
					BDYUTILS.PR_PutNumberAssign("nDiskYoffRt", nDiskYoff)
					
				End If
			Next ii
			
			'Insert disks and text using the Assignments from above
			BDYUTILS.PR_PutLine("@" & g_PathJOBST & "\VEST\BODYBRA.D;")
			
		Else
			'No bra cups are required
			BDYUTILS.PR_PutLine("// --------------------")
			BDYUTILS.PR_PutLine("// NO BRA CUPS REQUIRED")
			BDYUTILS.PR_PutLine("// --------------------")
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CreateMacro_Data(ByRef sfile As String)
		
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sfile, "Save Data and Set Variables for BODY.D", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hBody;")
		
		PR_UpdateDB()
		
		'Variable data
		'If display is in inches then convert to decimal inches
		'The conversion of CMs to Inches is done in BODY.D
		'in this way we do not add any extra problems due to rounding errors
		'
		'Patient details
		BDYUTILS.PR_PutStringAssign("sFileNo", (txtFileNo.Text))
		BDYUTILS.PR_PutStringAssign("sPatient", (txtPatientName.Text))
		BDYUTILS.PR_PutStringAssign("sAge", (txtAge.Text))
		BDYUTILS.PR_PutStringAssign("sSEX", (txtSex.Text))
		BDYUTILS.PR_PutStringAssign("sDiagnosis", (txtDiagnosis.Text))
		BDYUTILS.PR_PutNumberAssign("nUnitsFac", g_nUnitsFac)
		BDYUTILS.PR_PutNumberAssign("nAge", Val(txtAge.Text))
		'PR_PutStringAssign "sUnitType", txtUnits.Text
		
		'VEST Body Details
		BDYUTILS.PR_PutNumberAssign("nLtSCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(0).Text)))
		BDYUTILS.PR_PutNumberAssign("nRtSCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(1).Text)))
		BDYUTILS.PR_PutNumberAssign("nNeckCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(2).Text)))
		BDYUTILS.PR_PutNumberAssign("nSWidth", SCOOPDIA1.FN_Decimalise(Val(txtCir(3).Text)))
		BDYUTILS.PR_PutNumberAssign("nS_Waist", SCOOPDIA1.FN_Decimalise(Val(txtCir(4).Text)))
		BDYUTILS.PR_PutNumberAssign("nChestCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(5).Text)))
		BDYUTILS.PR_PutNumberAssign("nChestCirActual", SCOOPDIA1.FN_Decimalise(Val(txtCir(5).Text)))
		BDYUTILS.PR_PutNumberAssign("nWaistCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(6).Text)))
		BDYUTILS.PR_PutNumberAssign("nS_EOS", SCOOPDIA1.FN_Decimalise(Val(txtCir(7).Text)))
		BDYUTILS.PR_PutNumberAssign("nEOSCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(8).Text)))
		BDYUTILS.PR_PutNumberAssign("nS_Breast", SCOOPDIA1.FN_Decimalise(Val(txtCir(9).Text)))
		BDYUTILS.PR_PutNumberAssign("nBreastCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(10).Text)))
		BDYUTILS.PR_PutNumberAssign("nBreastCirActual", SCOOPDIA1.FN_Decimalise(Val(txtCir(10).Text)))
		BDYUTILS.PR_PutNumberAssign("nNippleCir", SCOOPDIA1.FN_Decimalise(Val(txtCir(11).Text)))
		
		BDYUTILS.PR_PutStringAssign("sBraLtCup", (cboLeftCup.Text))
		BDYUTILS.PR_PutStringAssign("sBraRtCup", (cboRightCup.Text))
		BDYUTILS.PR_PutStringAssign("sBraLtDisk", (txtLeftDisk.Text))
		BDYUTILS.PR_PutStringAssign("sBraRtDisk", (txtRightDisk.Text))
		
		BDYUTILS.PR_PutStringAssign("sLtAxillaType", FN_EscapeQuotesInString(cboLeftAxilla))
		BDYUTILS.PR_PutStringAssign("sRtAxillaType", FN_EscapeQuotesInString(cboRightAxilla))
		
		BDYUTILS.PR_PutStringAssign("sNeckType", (cboFrontNeck.Text))
		BDYUTILS.PR_PutNumberAssign("nNeckDimension", SCOOPDIA1.FN_Decimalise(Val(txtFrontNeck.Text)))
		BDYUTILS.PR_PutStringAssign("sBackNeckType", (cboBackNeck.Text))
		BDYUTILS.PR_PutNumberAssign("nBackNeckDim", SCOOPDIA1.FN_Decimalise(Val(txtBackNeck.Text)))
		
		BDYUTILS.PR_PutStringAssign("sClosure", (cboClosure.Text))
		BDYUTILS.PR_PutStringAssign("sFabric", (cboFabric.Text))
		
		BDYUTILS.PR_PutNumberAssign("nWaistCirUserFac", Val(cboRed(0).Text))
		BDYUTILS.PR_PutNumberAssign("nEOSCirUserFac", Val(cboRed(1).Text))
		BDYUTILS.PR_PutNumberAssign("nBreastCirUserFac", Val(cboRed(0).Text))
		
		If txtWorkOrder.Text = "" Then
			BDYUTILS.PR_PutStringAssign("sWorkOrder", "-")
		Else
			BDYUTILS.PR_PutStringAssign("sWorkOrder", (txtWorkOrder.Text))
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
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
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "S_EOS" & QCQ & txtCir(7).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "EOSCir" & QCQ & txtCir(8).Text & QQ & ");")
		
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "WaistCirUserFac" & QCQ & cboRed(0).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "EOSCirUserFac" & QCQ & cboRed(1).Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "BreastCirUserFac" & QCQ & cboRed(2).Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'It is assumed that the link open from Drafix has failed
		'Therefor we "End" here
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
				MsgBox("Bra disk sizes 1 to 9 only.", 48, "VEST Body - Dialogue")
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
				MsgBox("Bra disk sizes 1 to 9 only.", 48, "VEST Body - Dialogue")
		End Select
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class
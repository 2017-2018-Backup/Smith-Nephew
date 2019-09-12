Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class bodydia
	Inherits System.Windows.Forms.Form
	'Project:   bodydraw
	'Purpose:   Body Drawing function
	'
	'
	'Version:   3.00
	'Date:      3.May.96
	'Author:    Paul O'Rawe
	'           Gary George
	'           © C-Gem Ltd.
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Dec 98     GG      Port to VB5
	'
	'Notes:-
	'
	'
	'
	'
	
	'Windows API Functions Declarations
	'    Private Declare Function GetActiveWindow Lib "User" () As Integer
	'    Private Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
	'    Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
	'    Private Declare Function GetNumTasks Lib "Kernel" () As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
	
	
	'Constanst used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
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
		If iIndex <> 0 And CType(Me.Controls("txtSex"), Object).Text = "Female" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male Crotch type specified for Female patient" + NL
		End If
		
		'Check for female crotch and male patient
		iIndex = 0
		iIndex = InStr(1, sString, "Female", 0)
		If iIndex <> 0 And CType(Me.Controls("txtSex"), Object).Text = "Male" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Female Crotch type specified for Male patient" + NL
		End If
		
		
		'Display Error message (if required) and return
		'These are fatal errors
		If Len(sError) > 0 Then
			MsgBox(sError, 0, "Crotch Style Error")
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
		'Dim sReductionFactor  As String
		'Dim iAge As Integer
		
		'Get the currently selected leg style
		Select Case VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
			
			Case "Panty"
				prOptLeftLeg(1, 0)
				prOptRightLeg(1, 0)
				'Set dropdown combo boxes for reductions
				'Only set if Leg Style had been changed
				'If g_sPreviousLegStyle <> "" Then
				'prStore_Reductions
				'sReductionFactor = "0.85"
				'NB. Special case for TOS as TOS need not be given.
				'If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
				'    prSet_Red_Combo sReductionFactor, cboTOSRed
				'End If
				'prSet_Red_Combo sReductionFactor, cboWaistRed
				'prSet_Red_Combo sReductionFactor, cboMidPointRed
				'prSet_Red_Combo sReductionFactor, cboLargestRed
				'prSet_Red_Combo sReductionFactor, cboThighRed
				'End If
				
			Case "Brief", "Brief-French"
				prOptLeftLeg(0, 1)
				prOptRightLeg(0, 1)
				'Set dropdown combo boxes for reductions
				'Only set if Leg Style had been changed
				'If g_sPreviousLegStyle <> "" Then
				'prStore_Reductions
				'sReductionFactor = "0.85"
				'NB. Special case for TOS as TOS need not be given.
				'If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
				'    prSet_Red_Combo sReductionFactor, cboTOSRed
				'End If
				'prSet_Red_Combo sReductionFactor, cboWaistRed
				'prSet_Red_Combo sReductionFactor, cboMidPointRed
				'prSet_Red_Combo sReductionFactor, cboLargestRed
				'prSet_Red_Combo sReductionFactor, cboThighRed
				'End If
				
		End Select
		
		'Store leg style for use when it changes
		g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		
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
	
	Private Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Fuction to return the length between two points
		'Greatfull thanks to Pythagorus
		
		FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)
		
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
			If Char_Renamed = """" Then
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
	
	Private Function FN_Open(ByRef sDrafixFile As String) As Short
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
		NL = Chr(13) 'The new line character
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
		PrintLine(fNum, "//by Visual Basic, BODYSUIT - Drawing")
		
		'Define DRAFIX variables
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY     xyStart, xyOrigin, xyScale, xyO;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sFileNo, sSide, sID, sPathJOBST;")
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
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "ZipperLength" & QCQ & "length" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Zipper" & QCQ & "string" & QQ & ");")
		
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
		ARMDIA1.PR_SetLayer("Construct")
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
		
		'Start drawing on correct side
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		
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
		
		For ii = 0 To 6
			sString = sString & txtCir(ii).Text
		Next ii
		'Reason for skip is that 7 & 8 were taken out during development
		For ii = 9 To 16
			sString = sString & txtCir(ii).Text
		Next ii
		
		For ii = 0 To 4
			sString = sString & cboRed(ii).Text
		Next ii
		
		sString = sString & cboLeftCup.Text & txtLeftDisk.Text
		
		sString = sString & cboRightCup.Text & txtRightDisk.Text
		
		sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text
		
		sString = sString & cboFrontNeck.Text & txtFrontNeck.Text
		
		sString = sString & cboBackNeck.Text & txtBackNeck.Text
		
		sString = sString & cboClosure.Text
		
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
	
	Private Function fnRoundInches(ByVal nNumber As Double) As Double
		'Function to return the rounded value in decimal inches
		'returns to the nearest eighth (0.125)
		'E.G.
		'    5.67         = 5 inches and 0.67 inches
		'                   0.67 / 0.125 = 5.36 eighths
		'                   5.36 eighths = 5 eighths (rounded to nearest eighth)
		'    5.67         = 5 inches and 5 eighths
		'    5.67         = 5 + ( 5 * 0.125)
		'    5.67         = 6.625 inches
		'
		
		Dim iInt, iSign As Short
		Dim nPrecision, nDec As Double
		
		
		'Return 0 if input is Zero
		If nNumber = 0 Then
			fnRoundInches = 0
			Exit Function
		End If
		
		'Set precision
		nPrecision = 0.125
		
		'Break input into components
		iSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		iInt = Int(nNumber)
		nDec = nNumber - iInt
		
		'Get decimal part in precision units
		If nDec <> 0 Then
			nDec = nDec / nPrecision 'Avoid overflow
		End If
		nDec = ARMDIA1.round(nDec)
		
		'Return value
		fnRoundInches = (iInt + (nDec * nPrecision)) * iSign
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		Dim ii, iError As Short
		
		'Stop the timer used to ensure that the Dialogue dies
		'if the DRAFIX macro fails to establish a DDE Link
		Timer1.Enabled = False
		
		'Check that a "MainPatientDetails" Symbol has been
		'found
		If txtUidMPD.Text = "" Then
			MsgBox("No Patient Details have been found in drawing!", 16, "Error, Bodysuit Drawing")
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
		
		g_sFrontNeck = txtFrontNeck.Text
		sChar.Value = VB.Left(cboFrontNeck.Text, 1)
		
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
			If VB6.GetItemString(cboLegStyle, ii) = ARMDIA1.fnGetString(txtLegStyle.Text, 1, " ") Then
				cboLegStyle.SelectedIndex = ii
			End If
		Next ii
		
		'Set this now to avoid problems with cboLegStyle_Click
		If cboLegStyle.SelectedIndex = -1 Then
			g_sPreviousLegStyle = "None"
		Else
			g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
		End If
		
		If ARMDIA1.fnGetString(txtLegStyle.Text, 2, " ") <> "" Then
			If ARMDIA1.fnGetString(txtLegStyle.Text, 2, " ") = "Panty" Then
				optLeftLeg(0).Checked = True
			Else
				optLeftLeg(1).Checked = True 'Brief
			End If
		End If
		If ARMDIA1.fnGetString(txtLegStyle.Text, 3, " ") <> "" Then
			If ARMDIA1.fnGetString(txtLegStyle.Text, 3, " ") = "Panty" Then
				optRightLeg(0).Checked = True
			Else
				optRightLeg(1).Checked = True 'Brief
			End If
		End If
		
		'Create and check all values are present to allow creation of drawing.
		iError = FN_ValidateAndCalculateData(False) 'Sets variables from dialog box
		'N.B. Use of local JOBST Directory C:\JOBST within this Routine
		PR_CalculateBodySuit()
		PR_CreateDrawMacro()
		sTask = fnGetDrafixWindowTitleText()
		If sTask <> "" Then
			AppActivate(sTask)
			System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
		Else
			MsgBox("Can't find a Drafix Drawing to update! - Terminating.")
		End If
		
		'Missing values - User is informed of which in FnAllValuesPresent
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer back to normal.
		End
		
	End Sub
	
	Private Sub bodydia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		
		MainForm = Me 'load
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(13) 'The new line character
		g_sSide = VB.Command()
		
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
			MsgBox("Body Draw Dialogue is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting BodySuit - Dialogue")
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
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
		cboRed(0).Items.Add("0.80")
		cboRed(0).Items.Add("")
		
		'Chest Default 90% +/- 5%
		cboRed(1).Items.Add("0.95")
		cboRed(1).Items.Add("0.85")
		cboRed(1).Items.Add("")
		
		'Under Breast 90% +/- 5%
		cboRed(2).Items.Add("0.95")
		cboRed(2).Items.Add("0.85")
		cboRed(2).Items.Add("")
		
		'Large part of Buttocks 85% +/- 5%
		cboRed(3).Items.Add("0.90")
		cboRed(3).Items.Add("0.80")
		cboRed(3).Items.Add("")
		
		'Left Thigh 85% +/- 5%
		cboRed(4).Items.Add("0.90")
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
	
	Private Sub PR_CalculateBodySuit()
		'This function caclulates the body suit
		'
		'This function is used by the project
		'   BODYDRAW.MAK
		'
		
		Static nBackNeckOffset As Double
		Static nFrontAngle As Double
		Static nRaglanOffset As Double
		Static nShldWidthOffset As Double
		Static nFrontScoopDistance As Double
		Static nBackScoopDistance As Double
		Static nScoopAngle As Double
		Static nGussetSize As Double
		Static nGussetLength As Double
		Static nGussetRadius As Double
		Static nCutSegmentAngle As Double
		Static nCutArcLength As Double
		Static nStrapAngle As Double
		Static nStrapLength As Double
		Static nGussetAngle As Double
		Static n8_7Length As Double
		Static iError As Short
		Static nCutOut7Angle As Double
		Static nTmpRadius As Double
		Static nCrotchFilletRadius As Double
		Static nCutExtendedLength As Double
		Static xyTmpCutOut7 As XY
		Static xyTmp1 As XY
		Static xyTmp2 As XY
		Static xyTmp3 As XY
		Static bShldWidthModified As Short
		Static nTangentAngle As Double
		Static nGroinAngle As Double
		Static nThighAngle As Double
		Static nGroinToThighAngle As Double
		Static nGroinToThighMidPoint As Double
		Static sLeg As String
		Static sSleeve As String
		Static xyRegularTmpCutOut8 As XY
		Static nRegularRadiusFrontNeck As Double
		Static aAngle As Double
		Static nLength As Double
		Static ii As Short
		
		'Drawing variables
		g_bMissCutOut7 = False
		nScoopAngle = 0
		nFrontScoopDistance = 0
		nBackScoopDistance = 0
		
		'Body Co-ordinate points and values
		'Setup which side to be drawn
		If g_bDiffAxillaHeight Then
			'Draw the body using the highest axilla
			If g_nLeftShldToAxilla < g_nRightShldToAxilla Then
				g_nShldToAxilla = g_nLeftShldToAxilla
				g_nShldToAxillaLow = g_nRightShldToAxilla
				g_nAxillaCir = g_nLeftAxillaCir
				g_nAxillaCirLow = g_nRightAxillaCir
				g_sAxillaSide = "Left"
				g_sAxillaSideLow = "Right"
				g_sAxillaType = g_sLeftSleeve
				g_sAxillaTypeLow = g_sRightSleeve
				
			Else
				g_nShldToAxilla = g_nRightShldToAxilla
				g_nShldToAxillaLow = g_nLeftShldToAxilla
				g_nAxillaCir = g_nRightAxillaCir
				g_nAxillaCirLow = g_nLeftAxillaCir
				g_sAxillaSide = "Right"
				g_sAxillaSideLow = "Left"
				g_sAxillaType = g_sRightSleeve
				g_sAxillaTypeLow = g_sLeftSleeve
			End If
		Else
			'Default to left side
			g_sAxillaSide = "Left" 'Draw axilla by default on left side
			g_sAxillaType = g_sLeftSleeve
			g_nShldToAxilla = g_nLeftShldToAxilla
			g_nAxillaCir = g_nLeftAxillaCir
		End If
		
		ARMDIA1.PR_MakeXY(xySeamFold, 0, INCH1_4)
		ARMDIA1.PR_MakeXY(xySeamButt, xySeamFold.X + g_nButtocksLength, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xySeamHighShld, xySeamFold.X + g_nShldToFold, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xySeamLowShld, xySeamHighShld.X - INCH1_2, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xySeamChest, xySeamHighShld.X - g_nShldToAxilla, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xySeamChestAxillaLow, xySeamHighShld.X - g_nShldToAxillaLow, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xySeamWaist, xySeamHighShld.X - g_nShldToWaist, xySeamFold.y)
		ARMDIA1.PR_MakeXY(xyProfileButt, xySeamButt.X, xySeamButt.y + g_nButtCirCalc)
		ARMDIA1.PR_MakeXY(xyCutOut3, xyProfileButt.X, xyProfileButt.y - g_nButtBackSeam)
		ARMDIA1.PR_MakeXY(xyCutOut4, xySeamFold.X + INCH3_4, xyCutOut3.y - (g_nButtCutOut / 2))
		ARMDIA1.PR_MakeXY(xyCutOut5, xySeamButt.X, xySeamButt.y + g_nButtFrontSeam)
		ARMDIA1.PR_MakeXY(xyCutOut6, xySeamWaist.X, xySeamWaist.y + g_nWaistFrontSeam)
		ARMDIA1.PR_MakeXY(xyCutOut2, xyCutOut6.X, xyCutOut6.y + g_nCutOut)
		ARMDIA1.PR_MakeXY(xyProfileWaist, xyCutOut2.X, xyCutOut2.y + g_nWaistBackSeam)
		ARMDIA1.PR_MakeXY(xyCutOut7, xySeamChest.X, xySeamChest.y + g_nChestFrontSeam)
		ARMDIA1.PR_MakeXY(xyProfileChest, xySeamChest.X, xyCutOut2.y + g_nChestBackSeam)
		
		'    g_nCutOutRadius = (FN_CalcLength(xyCutOut4, xyCutOut3) / 2) / Cos(FN_CalcAngle(xyCutOut4, xyCutOut3) * (PI / 180))
		'Original 85% mark
		
		If xyCutOut5.X - xyCutOut4.X > (xyCutOut3.y - xyCutOut5.y) / 2 Then
			'Extreme measurement case
			g_bExtremeCrotch = True
			g_nCutOutRadius = (ARMDIA1.FN_CalcLength(xyCutOut5, xyCutOut3) / 2)
			ARMDIA1.PR_MakeXY(xyCutOut9, xyCutOut4.X + g_nCutOutRadius, xyCutOut3.y)
			ARMDIA1.PR_MakeXY(xyCutOut10, xyCutOut9.X, xyCutOut5.y)
		Else
			'Normal case
			aAngle = ARMDIA1.FN_CalcAngle(xyCutOut4, xyCutOut3)
			If aAngle <> 90 Then
				g_nCutOutRadius = (ARMDIA1.FN_CalcLength(xyCutOut4, xyCutOut3) / 2) / System.Math.Cos(aAngle * (PI / 180))
			Else
				g_nCutOutRadius = ARMDIA1.FN_CalcLength(xyCutOut4, xyCutOut3) / 2
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xyCutOut9. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyCutOut9 = xyCutOut3
			'UPGRADE_WARNING: Couldn't resolve default property of object xyCutOut10. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyCutOut10 = xyCutOut5
		End If
		
		ARMDIA1.PR_MakeXY(xyCutOutArcCentre, xyCutOut4.X + g_nCutOutRadius, xyCutOut4.y)
		
		'Calculate Centre Point of Buttocks Arc for the averaged/smallest thigh
		ARMDIA1.PR_MakeXY(xyButtocksArcCentre, xySeamButt.X, xySeamButt.y + g_nHalfGroinHeight)
		ARMDIA1.PR_MakeXY(xyProfileGroin, xySeamFold.X + INCH3_4, xySeamFold.y + g_nGroinHeight)
		
		'Calculate above for largest thigh (if given)
		'    If g_bDiffThigh Then
		ARMDIA1.PR_MakeXY(xyLT_ButtocksArcCentre, xySeamButt.X, xySeamButt.y + g_nLT_HalfGroinHeight)
		ARMDIA1.PR_MakeXY(xyLT_ProfileGroin, xySeamFold.X + INCH3_4, xySeamFold.y + g_nLT_GroinHeight)
		'    End If
		
		
		'FRONT NECK
		'Needed variables for neck
		If g_bSleeveless Then
			nBackNeckOffset = ARMEDDIA1.fnRoundInches(g_nNeckCir / 6)
			g_nRadiusFrontNeck = ARMEDDIA1.fnRoundInches(((g_nNeckCir - (nBackNeckOffset * 2)) / 3.14) - INCH1_8) 'Full scale
		Else
			nBackNeckOffset = ARMEDDIA1.fnRoundInches(g_nNeckCir / 9)
			g_nRadiusFrontNeck = ARMEDDIA1.fnRoundInches(((g_nNeckCir - ((nBackNeckOffset * 2) + 1.5)) / 3.14) - INCH1_8) 'Full scale
		End If
		If InStr(1, g_sFrontNeckStyle, "Scoop") > 0 Or g_sFrontNeckStyle = "V neck" Then
			'Scoop, Measured scoop, V neck and Measured V neck
			ARMDIA1.PR_MakeXY(xyRegularTmpCutOut8, xySeamHighShld.X - g_nRadiusFrontNeck, xySeamHighShld.y + g_nRadiusFrontNeck)
			g_nRadiusFrontNeck = g_nRadiusFrontNeck * 1.2
			g_nShldWidth = g_nShldWidth - 1
			bShldWidthModified = True
		End If
		ARMDIA1.PR_MakeXY(xyTmpCutOut8, xySeamHighShld.X - g_nRadiusFrontNeck, xySeamHighShld.y + g_nRadiusFrontNeck)
		
		ARMDIA1.PR_MakeXY(xyFrontNeckArcCentre, xySeamHighShld.X, xySeamHighShld.y + g_nRadiusFrontNeck)
		
		Select Case g_sFrontNeckStyle
			Case "Measured Scoop", "Measured V neck"
				n8_7Length = ARMDIA1.FN_CalcLength(xyRegularTmpCutOut8, xyCutOut7)
				If (n8_7Length < g_nFrontNeckSize) Then
					nScoopAngle = ARMDIA1.FN_CalcAngle(xyCutOut7, xyCutOut6)
					ARMDIA1.PR_CalcPolar(xyCutOut7, nScoopAngle, g_nFrontNeckSize - n8_7Length, xyCutOut8)
					g_bMissCutOut7 = True
					'Get intersection on ordinary scoop neck line
					'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyTmp1 = xyCutOut8
					xyTmp1.y = xyCutOut8.y + 10 'Extend line
					ii = BDYUTILS.FN_LinLinInt(xyCutOut6, xyTmpCutOut8, xyCutOut8, xyTmp1, xyCutOut8)
				Else
					nScoopAngle = ARMDIA1.FN_CalcAngle(xyRegularTmpCutOut8, xyCutOut7)
					ARMDIA1.PR_CalcPolar(xyRegularTmpCutOut8, nScoopAngle, g_nFrontNeckSize, xyCutOut8)
					'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyTmp1 = xyCutOut8
					xyTmp1.y = xyCutOut8.y + 10 'Extend line
					ii = BDYUTILS.FN_LinLinInt(xyCutOut7, xyTmpCutOut8, xyCutOut8, xyTmp1, xyCutOut8)
				End If
				''revise y of xyCutOut8
				'PR_MakeXY xyCutOut8, xyCutOut8.x, xyTmpCutOut8.y
			Case Else
				'Regular, V neck and Scoop
				'UPGRADE_WARNING: Couldn't resolve default property of object xyCutOut8. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyCutOut8 = xyTmpCutOut8
		End Select
		
		
		'BACK NECK Check for scoop
		If InStr(g_sBackNeckStyle, "Scoop") <> 0 Then
			If Not bShldWidthModified Then
				g_nShldWidth = g_nShldWidth - 1
			End If
			'Check for scoops here
			If g_sBackNeckStyle = "Scoop" Then
				If g_nAdult Then
					nBackScoopDistance = INCH3_4
				Else
					nBackScoopDistance = INCH3_8
				End If
			Else
				nBackScoopDistance = g_nBackNeckSize
			End If
		End If
		nFrontAngle = -((360 / (2 * PI * g_nRadiusFrontNeck)) + 90)
		If (g_sSide = "Left") Then
			sSleeve = g_sLeftSleeve
		Else
			sSleeve = g_sRightSleeve
		End If
		
		
		If sSleeve = "Sleeveless" Then
			'Neck Co-ordinate points and values
			ARMDIA1.PR_MakeXY(xyCutOutBackNeck, xySeamHighShld.X - (INCH1_4 + nBackScoopDistance), xyCutOut2.y)
			ARMDIA1.PR_MakeXY(xyProfileNeck, xySeamHighShld.X, xyCutOut2.y + nBackNeckOffset)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyCutOutFrontNeck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyCutOutFrontNeck = xySeamHighShld
			ARMDIA1.PR_MakeXY(xyProfileNeckMirror, xyProfileNeck.X, -xyProfileNeck.y)
			
			'Axilla Co-ordinate points and values
			'Because it's sleeveless the axilla length is figured
			'Needed variables for Axilla
			nRaglanOffset = ((xyCutOutFrontNeck.y - xyProfileNeckMirror.y) - INCH1_2) / 2
			'This variable is only used for Sleeveless
			nShldWidthOffset = System.Math.Sqrt((g_nShldWidth ^ 2) - (INCH1_2 ^ 2))
			If (nRaglanOffset < nShldWidthOffset) Then
				nShldWidthOffset = nRaglanOffset
			End If
			ARMDIA1.PR_MakeXY(xyRaglan1, xySeamHighShld.X, xySeamHighShld.y - INCH1_4)
			ARMDIA1.PR_MakeXY(xyRaglan2, xySeamLowShld.X, xyRaglan1.y - nShldWidthOffset)
			ARMDIA1.PR_MakeXY(xyRaglan3, xySeamLowShld.X, xyRaglan1.y - nRaglanOffset)
			ARMDIA1.PR_MakeXY(xyRaglan4, xyRaglan3.X - (g_nAxillaCir / 2), xyRaglan3.y)
			If g_bDiffAxillaHeight Then
				ARMDIA1.PR_MakeXY(xyRaglan4LowAxilla, xyRaglan3.X - (g_nAxillaCirLow / 2), xyRaglan3.y)
			End If
			ARMDIA1.PR_MakeXY(xyRaglan6, xyProfileNeckMirror.X, xyProfileNeckMirror.y + INCH1_4)
			ARMDIA1.PR_MakeXY(xyRaglan5, xySeamLowShld.X, xyRaglan6.y + nShldWidthOffset)
		Else
			'Neck Co-ordinate points and values
			ARMDIA1.PR_MakeXY(xyCutOutBackNeck, xySeamLowShld.X - (INCH1_4 + nBackScoopDistance), xyCutOut2.y)
			ARMDIA1.PR_MakeXY(xyProfileNeck, xySeamLowShld.X, xyCutOut2.y + nBackNeckOffset)
			ARMDIA1.PR_CalcPolar(xyFrontNeckArcCentre, nFrontAngle, g_nRadiusFrontNeck, xyCutOutFrontNeck)
			ARMDIA1.PR_MakeXY(xyProfileNeckMirror, xyProfileNeck.X, -xyProfileNeck.y)
			
			'Axilla Co-ordinate points and values
			'Needed variables for Axilla
			nRaglanOffset = ((xyCutOutFrontNeck.y - xyProfileNeckMirror.y) - INCH1_2) / 2
			ARMDIA1.PR_MakeXY(xyRaglan1, xyCutOutFrontNeck.X, xyCutOutFrontNeck.y - INCH1_4)
			ARMDIA1.PR_MakeXY(xyRaglan2, xySeamChest.X, xySeamChest.y - nRaglanOffset)
			ARMDIA1.PR_MakeXY(xyRaglan3, xyProfileNeckMirror.X, xyProfileNeckMirror.y + INCH1_4)
			If g_bDiffAxillaHeight Then
				ARMDIA1.PR_MakeXY(xyRaglan2LowAxilla, xySeamHighShld.X - g_nShldToAxillaLow, xySeamChest.y - nRaglanOffset)
				If g_sAxillaSide = "Left" Then
					g_nAxillaBackNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan3)
					g_nAxillaFrontNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
					g_nAFNRadRight = ARMDIA1.FN_CalcLength(xyRaglan2LowAxilla, xyRaglan1)
					g_nABNRadRight = ARMDIA1.FN_CalcLength(xyRaglan2LowAxilla, xyRaglan3)
				Else
					g_nAxillaBackNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2LowAxilla, xyRaglan3)
					g_nAxillaFrontNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2LowAxilla, xyRaglan1)
					g_nAFNRadRight = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
					g_nABNRadRight = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan3)
				End If
			Else
				g_nAxillaBackNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan3)
				g_nAxillaFrontNeckRad = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
				g_nABNRadRight = 0
				g_nAFNRadRight = 0
			End If
		End If
		
		
		If g_bDiffThigh Then 'One panty & brief
			'Release extra 5% reduction
			g_nThighCir = ARMEDDIA1.fnRoundInches(g_nSmallestThighGiven * (g_nThighCirRed + 0.05)) / 2 'half scale
		Else
			g_nThighCir = ARMEDDIA1.fnRoundInches(g_nSmallestThighGiven * g_nThighCirRed) / 2 'half scale
		End If
		
		'Calculate Crotch Co-ordinate points and values
		Select Case g_sCrotchStyle
			Case "Open Crotch"
				'Set sizes
				PR_SetCrotchSizes()
				If g_sSex = "Male" Then
					nCrotchFilletRadius = INCH3_8
					If Not g_nAdult Then
						nCrotchFilletRadius = INCH3_16
					End If
				Else
					nCrotchFilletRadius = INCH3_16
				End If
				nTmpRadius = g_nCutOutRadius + nCrotchFilletRadius
				
				If g_bExtremeCrotch Then
					'Extreme crotch
					nCutArcLength = (2 * PI * g_nCutOutRadius) / 4
					nCutExtendedLength = ARMDIA1.FN_CalcLength(xyCutOut10, xyCutOut5)
					'Back crotch
					If nCutArcLength > g_nBackCrotchSize Then
						'Open crotch ends on arc
						'UPGRADE_WARNING: Couldn't resolve default property of object xyBackCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyBackCrotch1 = xyCutOut9
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBackCrotch1)
						xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
						xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
						xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
						xyTmp2.y = xyBackCrotch1.y + 1
						iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
						ARMDIA1.PR_MakeXY(xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y)
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
						iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
					ElseIf (nCutArcLength + nCutExtendedLength) < g_nBackCrotchSize Then 
						'Open crotch ends at xyCutOut3
						'IE we can't go past the Buttock height
						'UPGRADE_WARNING: Couldn't resolve default property of object xyBackCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyBackCrotch1 = xyCutOut3
						ARMDIA1.PR_MakeXY(xyBackCrotch2, xyBackCrotch1.X, xyBackCrotch1.y + nCrotchFilletRadius)
						ARMDIA1.PR_MakeXY(xyBackCrotchFilletCentre, xyBackCrotch1.X - nCrotchFilletRadius, xyBackCrotch2.y)
						ARMDIA1.PR_MakeXY(xyBackCrotch3, xyBackCrotchFilletCentre.X, xyBackCrotchFilletCentre.y + nCrotchFilletRadius)
					Else
						'Open crotch ends between xyCutOut9 and xyCutout3
						'UPGRADE_WARNING: Couldn't resolve default property of object xyBackCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyBackCrotch1 = xyCutOut3
						xyBackCrotch1.X = xyCutOut3.X - ((nCutArcLength + nCutExtendedLength) - g_nBackCrotchSize)
						'Get center of fillet radius(initially relative to the arc)
						If (xyBackCrotch1.X - nCrotchFilletRadius) < xyCutOut9.X Then
							'Fillet intersects arc
							xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
							xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
							xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
							xyTmp2.y = xyBackCrotch1.y + 1
							iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
							ARMDIA1.PR_MakeXY(xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y)
							ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
							iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
						Else
							'Fillet is on straight bit
							ARMDIA1.PR_MakeXY(xyBackCrotch2, xyBackCrotch1.X, xyBackCrotch1.y + nCrotchFilletRadius)
							ARMDIA1.PR_MakeXY(xyBackCrotchFilletCentre, xyBackCrotch1.X - nCrotchFilletRadius, xyBackCrotch2.y)
							ARMDIA1.PR_MakeXY(xyBackCrotch3, xyBackCrotchFilletCentre.X, xyBackCrotchFilletCentre.y + nCrotchFilletRadius)
						End If
					End If
					
					'Front crotch
					If nCutArcLength > g_nFrontCrotchSize Then
						'Open crotch ends on arc
						'UPGRADE_WARNING: Couldn't resolve default property of object xyFrontCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyFrontCrotch1 = xyCutOut10
						xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
						xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
						xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
						xyTmp2.y = xyFrontCrotch1.y - 1
						iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
						ARMDIA1.PR_MakeXY(xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y)
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
						iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
					ElseIf (nCutArcLength + nCutExtendedLength) < g_nFrontCrotchSize Then 
						'Open crotch ends at xyCutOut5
						'IE we can't go past the Buttock height
						'UPGRADE_WARNING: Couldn't resolve default property of object xyFrontCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyFrontCrotch1 = xyCutOut5
						ARMDIA1.PR_MakeXY(xyFrontCrotch2, xyFrontCrotch1.X, xyFrontCrotch1.y - nCrotchFilletRadius)
						ARMDIA1.PR_MakeXY(xyFrontCrotchFilletCentre, xyFrontCrotch1.X - nCrotchFilletRadius, xyFrontCrotch2.y)
						ARMDIA1.PR_MakeXY(xyFrontCrotch3, xyFrontCrotchFilletCentre.X, xyFrontCrotchFilletCentre.y - nCrotchFilletRadius)
					Else
						'Open crotch ends between xyCutOut10 and xyCutout5
						'UPGRADE_WARNING: Couldn't resolve default property of object xyFrontCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyFrontCrotch1 = xyCutOut5
						xyFrontCrotch1.X = xyCutOut5.X - ((nCutArcLength + nCutExtendedLength) - g_nFrontCrotchSize)
						'Get center of fillet radius(initially relative to the arc)
						If (xyFrontCrotch1.X - nCrotchFilletRadius) < xyCutOut10.X Then
							'Fillet intersects arc
							xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
							xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
							xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
							xyTmp2.y = xyFrontCrotch1.y - 1
							iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
							ARMDIA1.PR_MakeXY(xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y)
							ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
							iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
						Else
							'Fillet is on straight bit
							ARMDIA1.PR_MakeXY(xyFrontCrotch2, xyFrontCrotch1.X, xyFrontCrotch1.y - nCrotchFilletRadius)
							ARMDIA1.PR_MakeXY(xyFrontCrotchFilletCentre, xyFrontCrotch1.X - nCrotchFilletRadius, xyFrontCrotch2.y)
							ARMDIA1.PR_MakeXY(xyFrontCrotch3, xyFrontCrotchFilletCentre.X, xyFrontCrotchFilletCentre.y - nCrotchFilletRadius)
						End If
					End If
					
				Else
					'Normal crotch
					nCutSegmentAngle = ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut5) - ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
					nCutArcLength = ARMEDDIA1.fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
					If (nCutArcLength <= g_nBackCrotchSize) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyBackCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyBackCrotch1 = xyCutOut3
					Else
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBackCrotch1)
					End If
					'Back of crotch
					xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
					xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
					xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
					xyTmp2.y = xyBackCrotch1.y + 1
					iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
					ARMDIA1.PR_MakeXY(xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y)
					ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
					iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
					'Front of crotch
					If (nCutArcLength <= g_nFrontCrotchSize) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyFrontCrotch1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyFrontCrotch1 = xyCutOut5
					Else
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, 180 + ((g_nFrontCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyFrontCrotch1)
					End If
					xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
					xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
					xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
					xyTmp2.y = xyFrontCrotch1.y - 1
					nTmpRadius = g_nCutOutRadius + nCrotchFilletRadius
					iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
					ARMDIA1.PR_MakeXY(xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y)
					ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
					iError = BDYUTILS.FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
				End If
				'Common
				ARMDIA1.PR_MakeXY(xyLegPoint, INCH3_4, 0)
				ARMDIA1.PR_MakeXY(xyCrotchMarker, xyCutOut4.X - (2 * nCrotchFilletRadius), xyCutOut4.y)
				If g_sSex = "Male" Then
					'1 1/4" + 3/4" to xyCutOut4 = 2"
					ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - (1 + INCH1_4), xySeamFold.y)
					If Not g_nAdult Then
						'7/8" + 3/4" to xyCutOut4 = 1 5/8"
						ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y)
					End If
				Else
					'7/8" + 3/4" to xyCutOut4 = 1 5/8"
					ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y)
				End If
				'Test that at least 1-1/4" between cutout edge and bottom of brief.
				If (xyCrotchMarker.X - xySeamThigh.X) < 1.25 Then 'less than 1-1/4"
					ARMDIA1.PR_MakeXY(xySeamThigh, xyCrotchMarker.X - 1.25, xySeamFold.y)
				End If
				ARMDIA1.PR_MakeXY(xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2))
				
			Case "Horizontal Fly"
				'7/8" + 3/4" to xyCutOut4 = 1 5/8"
				ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y)
				ARMDIA1.PR_MakeXY(xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2))
				PR_SetFlySize("Hor")
				ARMDIA1.PR_MakeXY(xyLegPoint, 0, 0)
				ARMDIA1.PR_MakeXY(xyCrotchMarker, xyCutOut4.X, xyCutOut4.y)
			Case "Diagonal Fly"
				'7/8" + 3/4" to xyCutOut4 = 1 5/8"
				ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y)
				ARMDIA1.PR_MakeXY(xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2))
				PR_SetFlySize("Dia")
				ARMDIA1.PR_MakeXY(xyLegPoint, 0, 0)
				ARMDIA1.PR_MakeXY(xyCrotchMarker, xyCutOut4.X, xyCutOut4.y)
				
			Case "Gusset"
				'7/8" + 3/4" to xyCutOut4 = 1 5/8"
				ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y)
				ARMDIA1.PR_MakeXY(xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2))
				PR_SetGussetSize()
				ARMDIA1.PR_MakeXY(xyLegPoint, 0, 0)
				ARMDIA1.PR_MakeXY(xyCrotchMarker, xyCutOut4.X, xyCutOut4.y)
				
			Case "Snap Crotch"
				'1/2" + 3/4" to xyCutOut4 = 1 1/4"
				ARMDIA1.PR_MakeXY(xySeamThigh, xySeamFold.X - INCH1_2, xySeamFold.y)
				ARMDIA1.PR_MakeXY(xyProfileBrief, xySeamThigh.X, xyCutOut4.y)
				PR_SetGussetSize()
				nGussetLength = g_nGussetLength + 1.5 '1-1/2"
				If nGussetLength < 2.5 Then '2-1/2" Minimum
					MsgBox(">>>>>>>>>>>>>>>>>WARNING <<<<<<<<<<<<" & Chr(13) & "Gusset Strap shorter than 2-1/2"" which is unacceptable.")
					' End
				End If
				g_nLengthStrap1ToCutOut5 = 0
				If g_bExtremeCrotch Then
					
					'Special case for extreme crotch
					nCutSegmentAngle = ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut10) - ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
					nCutArcLength = ARMEDDIA1.fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
					nCutExtendedLength = ARMDIA1.FN_CalcLength(xyCutOut10, xyCutOut5)
					
					If nCutArcLength > nGussetLength Then
						'End of strap ends on arc
						nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / PI))
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyStrap1)
						g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength
						
					ElseIf (nCutArcLength + nCutExtendedLength) > nGussetLength Then 
						'End of strap ends between xyCutOut10 and xyCutout5
						nStrapAngle = ARMDIA1.FN_CalcAngle(xyCutOut10, xyCutOut5)
						nStrapLength = nGussetLength - nCutArcLength
						ARMDIA1.PR_CalcPolar(xyCutOut10, nStrapAngle, nStrapLength, xyStrap1)
						g_nLengthStrap1ToCutOut5 = ARMDIA1.FN_CalcLength(xyStrap1, xyCutOut5)
					Else
						'End of strap ends between xyCutOut5 and xyCutout6
						nStrapAngle = ARMDIA1.FN_CalcAngle(xyCutOut5, xyCutOut6)
						nStrapLength = nGussetLength - (nCutArcLength + nCutExtendedLength)
						ARMDIA1.PR_CalcPolar(xyCutOut5, nStrapAngle, nStrapLength, xyStrap1)
						g_nLengthStrap1ToCutOut5 = 0
					End If
				Else
					'Normal crotch
					nCutSegmentAngle = ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut5) - ARMDIA1.FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
					nCutArcLength = ARMEDDIA1.fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
					If (nCutArcLength < nGussetLength) Then
						nStrapAngle = ARMDIA1.FN_CalcAngle(xyCutOut5, xyCutOut6)
						nStrapLength = nGussetLength - nCutArcLength
						ARMDIA1.PR_CalcPolar(xyCutOut5, nStrapAngle, nStrapLength, xyStrap1)
						g_nLengthStrap1ToCutOut5 = 0
					Else
						nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / PI))
						ARMDIA1.PR_CalcPolar(xyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyStrap1)
						g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength
					End If
				End If
				
				ARMDIA1.PR_MakeXY(xyStrap2, xyStrap1.X, xyStrap1.y - 1.75) '1-3/4"
				nGussetRadius = xyCutOutArcCentre.y - xyStrap2.y
				ARMDIA1.PR_MakeXY(xyGussetArcCentre, xyProfileBrief.X + nGussetRadius, xyProfileBrief.y)
				
				If xyStrap1.X > xyGussetArcCentre.X Then
					'Normal end of Snap crotch arc
					ARMDIA1.PR_MakeXY(xyStrap3, xyGussetArcCentre.X, xyStrap2.y) 'Roughly!!!
				Else
					'End of arc
					ARMDIA1.PR_MakeXY(xyStrap3, xyStrap1.X - INCH3_4, xyStrap2.y)
					'Revise xyGussetArcCentre
					'                aAngle = Fn_CalcAngle(xyStrap3, xyProfileBrief) - 90
					'                nLength = FN_CalcLength(xyStrap3, xyProfileBrief) / 2
					'                Pr_CalcMidPoint xyStrap3, xyProfileBrief, xyTmp1
					'                nLength = Sqr(nGussetRadius ^ 2 - nLength ^ 2)
					'                PR_CalcPolar xyTmp1, aAngle, nLength, xyGussetArcCentre
					aAngle = 180 - ARMDIA1.FN_CalcAngle(xyStrap3, xyProfileBrief)
					nLength = ARMDIA1.FN_CalcLength(xyStrap3, xyProfileBrief) / 2
					nGussetRadius = nLength / System.Math.Cos(aAngle * (PI / 180))
					ARMDIA1.PR_MakeXY(xyGussetArcCentre, xyProfileBrief.X + nGussetRadius, xyProfileBrief.y)
				End If
				ARMDIA1.PR_MakeXY(xyStrap4, INCH3_4, xyStrap2.y)
				ARMDIA1.PR_MakeXY(xyLegPoint, INCH3_4, 0)
				ARMDIA1.PR_MakeXY(xyCrotchMarker, xyCutOut4.X, xyCutOut4.y)
				
		End Select
		
		'Panty
		'NB We check at this point for complete leg data we can't do this in the
		'in the main validation routine as we do not have leg data
		'at that point
		If g_sLeftLeg = "Panty" Or g_sRightLeg = "Panty" Then
			
			If Not FN_ValidateLegData() Then End
			If Not FN_InitiliseLegRoutines() Then End
			
			If g_sLeftLeg = "Panty" Then PR_CalculateLeg("Left", g_nLeftLegLength)
			If g_sRightLeg = "Panty" Then PR_CalculateLeg("Right", g_nRightLegLength)
			If g_sLeftLeg = "Panty" And g_sRightLeg = "Panty" Then g_bDrawSingleLeg = FN_AverageLeftAndRightLegs()
			If g_sLeftLeg <> "Panty" And g_sRightLeg = "Panty" Then
				g_bDrawSingleLeg = True
				PR_CopyRightToLeftLeg()
				xyLeftLegLabelPoint.X = xyLeftLegLabelPoint.X - g_nLeftLegLength 'This normally a side efect of averaging the legs'Bad practice but what the hell
			End If
			If g_sLeftLeg = "Panty" And g_sRightLeg <> "Panty" Then
				g_bDrawSingleLeg = True
				xyLeftLegLabelPoint.X = xyLeftLegLabelPoint.X - g_nLeftLegLength 'As above
			End If
			
			ARMDIA1.PR_MakeXY(xyLT_Fold, xySeamFold.X, xySeamFold.y + g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
			ARMDIA1.PR_MakeXY(xyFold, xySeamFold.X, xySeamFold.y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
			
		End If
		
		If g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief" Then
			'Check to see if it can reach, if not use the tangent from the groin
			'point and join them with a straight line.
			'Calculate points for Average / Smallest thigh
			nGroinAngle = ARMDIA1.FN_CalcAngle(xyButtocksArcCentre, xyProfileGroin)
			If (xyButtocksArcCentre.X - xySeamThigh.X) < g_nButtRadius Then
				'Curve reaches bottom of Brief
				g_bDrawBriefCurve = True
				ARMDIA1.PR_MakeXY(xyTmp1, xySeamThigh.X, xyButtocksArcCentre.y)
				ARMDIA1.PR_MakeXY(xyTmp2, xyTmp1.X, xyProfileButt.y)
				iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyButtocksArcCentre, g_nButtRadius, xyProfileThigh)
				nThighAngle = ARMDIA1.FN_CalcAngle(xyButtocksArcCentre, xyProfileThigh)
				ARMDIA1.PR_CalcPolar(xyButtocksArcCentre, nThighAngle - ((nThighAngle - nGroinAngle) / 3), g_nButtRadius, xyProfileThighExtraPT)
			Else
				'Curve doesn't reaches bottom of Brief - draw tangential line
				g_bDrawBriefCurve = False
				nTangentAngle = nGroinAngle + 90
				ARMDIA1.PR_MakeXY(xyTmp1, xyProfileBrief.X, xyProfileButt.y)
				ARMDIA1.PR_CalcPolar(xyProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2)
				iError = BDYUTILS.FN_LinLinInt(xyProfileBrief, xyTmp1, xyProfileGroin, xyTmp2, xyProfileThigh)
				nGroinToThighAngle = ARMDIA1.FN_CalcAngle(xyProfileGroin, xyProfileThigh)
				nGroinToThighMidPoint = (ARMDIA1.FN_CalcLength(xyProfileGroin, xyProfileThigh) / 2)
				ARMDIA1.PR_CalcPolar(xyProfileGroin, nGroinToThighAngle, nGroinToThighMidPoint, xyProfileThighExtraPT)
			End If
			
			'Calculate points for largest thigh (if required)
			'I know parts of this is identical to above except for a few variable changes but I
			'am trying to retrofit this code without making major changes to code that has
			'already been validated. (Yeah! I know it sound feeble, but this is not paying me enough to
			'really care) GG
			If g_bDiffThigh Then
				nGroinAngle = ARMDIA1.FN_CalcAngle(xyLT_ButtocksArcCentre, xyLT_ProfileGroin)
				If (xyLT_ButtocksArcCentre.X - xySeamThigh.X) < g_nLT_ButtRadius Then
					'Curve reaches bottom of Brief
					ARMDIA1.PR_MakeXY(xyTmp1, xySeamThigh.X, xyLT_ButtocksArcCentre.y)
					ARMDIA1.PR_MakeXY(xyTmp2, xyTmp1.X, xyProfileButt.y + 2) 'Ensure it clears
					iError = BDYUTILS.FN_CirLinInt(xyTmp1, xyTmp2, xyLT_ButtocksArcCentre, g_nLT_ButtRadius, xyLT_ProfileThigh)
				Else
					'Curve doesn't reaches bottom of Brief - draw tangential line
					nTangentAngle = nGroinAngle + 90
					ARMDIA1.PR_MakeXY(xyTmp1, xyProfileBrief.X, xyProfileButt.y + 2)
					ARMDIA1.PR_CalcPolar(xyLT_ProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2)
					iError = BDYUTILS.FN_LinLinInt(xyProfileBrief, xyTmp1, xyLT_ProfileGroin, xyTmp2, xyLT_ProfileThigh)
				End If
				ARMDIA1.PR_MakeXY(xyLT_Fold, xySeamFold.X, xySeamFold.y + g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
				'Revise xyLT_ButtocksArcCentre so the point can be used construct the back arc
				'Note that we are trying to blend from the largest thigh into the body based on
				'the smallest thigh
				aAngle = ARMDIA1.FN_CalcAngle(xyLT_Fold, xyProfileButt) - 90 'perpendicular angle to the line xyLT_Fold, xyProfileButt
				nLength = System.Math.Sqrt(g_nButtRadius ^ 2 - (ARMDIA1.FN_CalcLength(xyLT_Fold, xyProfileButt) / 2) ^ 2)
				BDYUTILS.PR_CalcMidPoint(xyLT_Fold, xyProfileButt, xyTmp1)
				ARMDIA1.PR_CalcPolar(xyTmp1, aAngle, nLength, xyLT_ButtocksArcCentre)
			End If
		End If
		
		'Average/Smallest thigh
		ARMDIA1.PR_MakeXY(xyFold, xySeamFold.X, xySeamFold.y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
		ARMDIA1.PR_MakeXY(xy85Fold, xySeamFold.X, xySeamFold.y + g_n85FoldHeight)
		ARMDIA1.PR_MakeXY(xyProfileThighMirror, xyProfileThigh.X, -xyProfileThigh.y)
		' If g_bDiffThigh Then PR_MakeXY xyLT_ProfileThighMirror, xyLT_ProfileThigh.x, -xyLT_ProfileThigh.y
		ARMDIA1.PR_MakeXY(xyLT_ProfileThighMirror, xyLT_ProfileThigh.X, -xyLT_ProfileThigh.y)
		ARMDIA1.PR_MakeXY(xyProfileThighMirror1, xyProfileThighMirror.X, xyProfileThighMirror.y + INCH1_4)
		
		'largest thigh (if required)
		
	End Sub
	
	Private Sub PR_CreateDrawMacro()
		Static sText As String
		Static xyText As XY
		'UPGRADE_WARNING: Lower bound of array xyZipMark was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static xyZipMark(4) As XY
		Static xyTmp As XY
		Static xyTmp1 As XY
		Static aAngle As Double
		Static aInc As Double
		Static ii As Short
		Static nLength As Double
		Static nOffset As Double
		'UPGRADE_WARNING: Lower bound of array xyOpen was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static xyOpen(4) As XY
		'UPGRADE_WARNING: Lower bound of array aOpen was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static aOpen(4) As Double
		
		Static sAxillaText As String
		Static sAxillaTextLow As String
		
		Static sAxillaTextThis As String
		Static sThis As String
		Static sOther As String
		
		Static nVertLinesOffset As Double
		Static nHorLinesOffset As Short
		Static nSmallestLegXOffset As Double
		Static nLargestLegXOffset As Double
		Static xySeamLineStart As XY
		Static xySeamLineEnd As XY
		Static xyCentreLineStart As XY
		Static xyCentreLineEnd As XY
		Static xyBackNeckMid1 As XY
		Static xyBackNeckMid2 As XY
		Static xyFrontNeckMid1 As XY
		Static xyFrontNeckMid2 As XY
		Static xyGroinExtraPT As XY
		Static xyGroinExtraPT1 As XY
		Static xyTmpRaglanCentre As XY
		Static nBackNeckMidOffsetX As Double
		Static nBackNeckMidOffsetY As Double
		Static nFlip As Short
		Static xyCrotchText As XY
		Static sCrotchText As String
		Static xyGussetOffCut As XY
		Static xyLLGussetRectangle As XY
		Static xyURGussetRectangle As XY
		Static xyNeckText As XY
		Static xyCutOutArcStart As XY
		Static xyCutOutArcEnd As XY
		Static xyProfileThighRelease As XY
		Static xyLegStart As XY
		Static xyMesh As XY
		
		
		
		'UPGRADE_WARNING: Arrays in structure cProfilePoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cProfilePoly As curve
		'UPGRADE_WARNING: Arrays in structure cRaglanPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cRaglanPoly As curve
		'UPGRADE_WARNING: Arrays in structure cRaglanCurve may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cRaglanCurve As curve
		'UPGRADE_WARNING: Arrays in structure cFrontCutOutPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cFrontCutOutPoly As curve
		'UPGRADE_WARNING: Arrays in structure cBackCutOutPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cBackCutOutPoly As curve
		'UPGRADE_WARNING: Arrays in structure cFrontNeckArc may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cFrontNeckArc As curve
		'UPGRADE_WARNING: Arrays in structure cBackNeckArc may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static cBackNeckArc As curve
		Static sSmallestElasticText As String
		Static sLargestElasticText As String
		
		Static MeshProfile As BiArc
		Static MeshSeam As BiArc
		Static nMeshSize As Double
		Static nMeshDistanceAlongRaglan As Double
		
		'UPGRADE_WARNING: Arrays in structure SmallestLegCurve may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static SmallestLegCurve As curve
		'UPGRADE_WARNING: Arrays in structure LargestLegCurve may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static LargestLegCurve As curve
		
		
		Static sLeg As String
		
		'Initialise
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open("C:\JOBST\DRAW.D")
		
		'Add labels and construction lines
		nVertLinesOffset = xyProfileButt.y + 1
		nHorLinesOffset = g_nShldToFold * 0.05
		ARMDIA1.PR_SetLayer("Construct")
		If Not g_sLegStyle = "Panty" Then PR_DrawVerticalLineThruPoint(xySeamThigh, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamFold, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamButt, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamWaist, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamLowShld, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamHighShld, nVertLinesOffset)
		PR_DrawVerticalLineThruPoint(xySeamChest, nVertLinesOffset)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xySeamLineStart, 0 - (ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), INCH1_4)
		ARMDIA1.PR_MakeXY(xySeamLineEnd, xySeamHighShld.X + nHorLinesOffset, INCH1_4)
		'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ARMDIA1.PR_MakeXY(xyCentreLineStart, 0 - (ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), 0)
		ARMDIA1.PR_MakeXY(xyCentreLineEnd, xySeamHighShld.X + nHorLinesOffset, 0)
		
		
		'Useful Points
		ARMDIA1.PR_MakeXY(xyBackNeckMid1, xyCutOutBackNeck.X + ((xyProfileNeck.X - xyCutOutBackNeck.X) / 6), xyCutOutBackNeck.y + ((xyProfileNeck.y - xyCutOutBackNeck.y) / 3))
		ARMDIA1.PR_MakeXY(xyBackNeckMid2, xyProfileNeck.X - ((xyProfileNeck.X - xyCutOutBackNeck.X) / 2), xyProfileNeck.y - ((xyProfileNeck.y - xyCutOutBackNeck.y) / 3))
		ARMDIA1.PR_MakeXY(xyFrontNeckMid1, xyCutOutFrontNeck.X - ((xyCutOutFrontNeck.X - xyCutOut8.X) / 3), xyCutOutFrontNeck.y + ((xyCutOut8.y - xyCutOutFrontNeck.y) / 6))
		ARMDIA1.PR_MakeXY(xyFrontNeckMid2, xyCutOut8.X + ((xyCutOutFrontNeck.X - xyCutOut8.X) / 3), xyCutOut8.y - ((xyCutOut8.y - xyCutOutFrontNeck.y) / 2))
		ARMDIA1.PR_CalcPolar(xyButtocksArcCentre, 45, g_nButtRadius, xyCutOutArcStart)
		ARMDIA1.PR_CalcPolar(xyButtocksArcCentre, 135, g_nButtRadius, xyCutOutArcEnd)
		ARMDIA1.PR_MakeXY(xyGroinExtraPT, xySeamFold.X + (g_nButtocksLength / 2), xySeamFold.y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength / 2) ^ 2)))
		ARMDIA1.PR_MakeXY(xyGroinExtraPT1, xySeamFold.X + ((g_nButtocksLength / 4) * 3), xySeamFold.y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength / 4) ^ 2)))
		
		'Calculated for manual editing
		ARMDIA1.PR_MakeXY(xyProfileThighRelease, xySeamThigh.X, xySeamThigh.y + g_nThighCir) 'this is the only place that g_nThighCir is used
		BDYUTILS.PR_DrawMarker(xyProfileThighRelease)
		BDYUTILS.PR_AddDBValueToLast("Data", "Thigh")
		BDYUTILS.PR_DrawMarker(xy85Fold)
		BDYUTILS.PR_AddDBValueToLast("Data", "At original 85% reduction")
		'Draw Construction Lines
		ARMDIA1.PR_DrawLine(xySeamLineStart, xySeamLineEnd)
		ARMDIA1.PR_DrawLine(xyCentreLineStart, xyCentreLineEnd)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "CentreLine")
		'Draw Buttocks Arc for reference
		ARMDIA1.PR_DrawArc(xyButtocksArcCentre, xyCutOutArcStart, xyCutOutArcEnd)
		'Insert Key Markers for reference
		BDYUTILS.PR_DrawMarker(xyProfileThigh)
		BDYUTILS.PR_DrawMarker(xyProfileGroin)
		BDYUTILS.PR_DrawMarker(xyProfileButt)
		BDYUTILS.PR_DrawMarker(xyProfileWaist)
		'YES - 2 at same point (one with DB value)!!!!
		BDYUTILS.PR_DrawMarker(xyProfileChest)
		BDYUTILS.PR_DrawMarker(xyProfileChest)
		BDYUTILS.PR_AddDBValueToLast("curvetype", "BackCurveMarker")
		BDYUTILS.PR_DrawMarker(xyProfileNeck)
		BDYUTILS.PR_DrawMarker(xyCutOutArcCentre)
		BDYUTILS.PR_DrawMarker(xyFrontNeckArcCentre)
		BDYUTILS.PR_DrawMarker(xyFold)
		BDYUTILS.PR_DrawMarker(xyButtocksArcCentre)
		
		'Draw DB markers for Zipper macros
		BDYUTILS.PR_DrawMarker(xyCutOutBackNeck)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "BackN")
		BDYUTILS.PR_DrawMarker(xyCutOut2)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "BackW")
		BDYUTILS.PR_DrawMarker(xyCutOut3)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "BackB")
		BDYUTILS.PR_DrawMarker(xyCutOut5)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "FrontB")
		BDYUTILS.PR_AddDBValueToLast("Data", Str(g_nLengthStrap1ToCutOut5))
		
		BDYUTILS.PR_DrawMarker(xyCutOut6)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "FrontW")
		BDYUTILS.PR_DrawMarker(xyCutOut8)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "FrontN")
		If g_sFrontNeckStyle = "Turtle" Then
			BDYUTILS.PR_AddDBValueToLast("Data", Str(g_nFrontNeckSize))
		End If
		
		
		'Other patient details in black on layer construct
		ARMDIA1.PR_MakeXY(xyText, xySeamButt.X + 1, xySeamButt.y - 1.8)
		sText = g_sFileNo & "\n" & g_sDiagnosis & "\n" & Trim(Str(g_nAge)) & "\n" & g_sSex
		ARMDIA1.Pr_DrawText(sText, xyText, 0.125)
		
		'Notes Text - e.g. Patient Details, Sewing instructions
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_MakeXY(xyText, xySeamButt.X + 1, xySeamButt.y - 1)
		sText = g_sPatient & "\n" & g_sWorkOrder & "\n" & Trim(Mid(g_sFabric, 4))
		ARMDIA1.Pr_DrawText(sText, xyText, 0.125)
		
		
		
		'Draw Measured Front Neck Text
		If InStr(g_sFrontNeckStyle, "Scoop") > 0 Then
			sText = "SCOOP NECK"
			ARMDIA1.PR_MakeXY(xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25)
		ElseIf InStr(g_sFrontNeckStyle, "V neck") > 0 Then 
			sText = "V NECK"
			ARMDIA1.PR_MakeXY(xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25)
		ElseIf InStr(g_sFrontNeckStyle, "Turtle") > 0 Then 
			sText = "NECK CIRC. " & ARMEDDIA1.fnInchestoText(g_nNeckCir) & "\""" & "\nTURTLE NECK WIDTH " & ARMEDDIA1.fnInchestoText(g_nFrontNeckSize) & "\"""
			ARMDIA1.PR_MakeXY(xyNeckText, xyCutOutFrontNeck.X - 3.5, xyCutOutFrontNeck.y + 0.5)
			If InStr(g_sFrontNeckStyle, "same") Then
				If g_sLegStyle = "Brief" Then
					sText = sText & "\nIN SAME FABRIC AS BRIEF"
				Else
					sText = sText & "\nIN SAME FABRIC AS SUIT"
				End If
			ElseIf InStr(g_sFrontNeckStyle, "Detach. Fabric") Then 
				sText = sText & "\nDETACHABLE FABRIC"
			Else
				sText = sText & "\nDETACHABLE"
			End If
			
		Else 'Regular
			sText = "NECK CIRC. " & ARMEDDIA1.fnInchestoText(g_nNeckCir) & "\"""
			ARMDIA1.PR_MakeXY(xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25)
		End If
		If InStr(g_sFrontNeckStyle, "Measured") > 0 Then sText = "MEASURED " & sText
		ARMDIA1.Pr_DrawText(sText, xyNeckText, 0.125)
		
		'Draw Measured Back Neck Text
		If InStr(g_sBackNeckStyle, "Scoop") > 0 Then
			If g_sBackNeckStyle = "Measured Scoop" Then
				sText = "MEASURED SCOOP NECK"
			Else
				sText = "SCOOP NECK"
			End If
			ARMDIA1.PR_MakeXY(xyNeckText, xyCutOutBackNeck.X - 3, xyCutOutBackNeck.y + 0.25)
			ARMDIA1.Pr_DrawText(sText, xyNeckText, 0.125)
		End If
		
		'Draw Zipper
		If g_sClosure = "Back Zip" Then
			'Back zipper
			BDYUTILS.PR_CalcMidPoint(xyCutOut3, xyCutOut2, xyZipMark(1))
			nLength = ARMDIA1.FN_CalcLength(xyCutOut2, xyCutOutBackNeck) + ARMDIA1.FN_CalcLength(xyCutOut3, xyZipMark(1))
			aAngle = ARMDIA1.FN_CalcAngle(xyCutOut3, xyZipMark(1))
			ARMDIA1.PR_MakeXY(xyZipMark(2), xyZipMark(1).X, xyZipMark(1).y + 0.125)
			ARMDIA1.PR_CalcPolar(xyZipMark(1), aAngle, 0.5, xyZipMark(3))
			BDYUTILS.PR_CalcMidPoint(xyCutOut2, xyCutOutBackNeck, xyZipMark(4)) 'Text insertion point
			xyZipMark(4).y = xyZipMark(4).y + 0.125
		End If
		
		Static nLen As Double
		If g_sClosure = "Front Zip" Then
			'Front Zipper
			If g_bMissCutOut7 Then
				nLength = ARMDIA1.FN_CalcLength(xyCutOut8, xyCutOut6)
				BDYUTILS.PR_CalcMidPoint(xyCutOut6, xyCutOut8, xyZipMark(4)) 'Text insertion point
			Else
				nLength = ARMDIA1.FN_CalcLength(xyCutOut8, xyCutOut7) + ARMDIA1.FN_CalcLength(xyCutOut7, xyCutOut6)
				BDYUTILS.PR_CalcMidPoint(xyCutOut6, xyCutOut7, xyZipMark(4)) 'Text insertion point
			End If
			nLength = nLength + ARMDIA1.FN_CalcLength(xyCutOut5, xyCutOut6)
			
			xyZipMark(4).y = xyZipMark(4).y - 0.25
			
			aAngle = ARMDIA1.FN_CalcAngle(xyCutOut5, xyCutOut6)
			If g_sCrotchStyle = "Snap Crotch" Then
				If g_nLengthStrap1ToCutOut5 > 0 Then
					'This takes care of the special case where the snap crotch lands
					'before xyCutOut5
					ARMDIA1.PR_CalcPolar(xyCutOut5, aAngle, 1.125 - g_nLengthStrap1ToCutOut5, xyZipMark(1))
					nLength = nLength - ARMDIA1.FN_CalcLength(xyCutOut5, xyZipMark(1))
				Else
					'We also need to take care of another special case
					'where the 1.25" along the cut-out to the start of the
					'zip mark takes it around the corner (so to speak)
					nLen = ARMDIA1.FN_CalcLength(xyStrap1, xyCutOut6)
					If nLen >= 1.625 Then
						'Simple case Zip mark is on the line (xyCutOut5, xyCutOut6)
						ARMDIA1.PR_CalcPolar(xyStrap1, aAngle, 1.125, xyZipMark(1))
						nLength = nLength - ARMDIA1.FN_CalcLength(xyCutOut5, xyZipMark(1))
					ElseIf nLen > 1.125 And nLen < 1.625 Then 
						'Zip mark is round the corner
						ARMDIA1.PR_CalcPolar(xyStrap1, aAngle, 1.125, xyZipMark(1))
						nLength = nLength - ARMDIA1.FN_CalcLength(xyCutOut5, xyZipMark(1))
						aAngle = ARMDIA1.FN_CalcAngle(xyCutOut6, xyCutOut7)
					Else
						'Zip mark is on the line (xyCutOut6, xyCutOut7)
						nLen = 1.125 - nLen
						aAngle = ARMDIA1.FN_CalcAngle(xyCutOut6, xyCutOut7)
						ARMDIA1.PR_CalcPolar(xyCutOut6, aAngle, nLen, xyZipMark(1))
						nLength = nLength - ARMDIA1.FN_CalcLength(xyCutOut5, xyCutOut6) - nLen
					End If
				End If
			Else
				BDYUTILS.PR_CalcMidPoint(xyCutOut5, xyCutOut6, xyZipMark(1))
				nLength = nLength - (ARMDIA1.FN_CalcLength(xyCutOut5, xyCutOut6) / 2)
			End If
			ARMDIA1.PR_MakeXY(xyZipMark(2), xyZipMark(1).X, xyZipMark(1).y - 0.125)
			ARMDIA1.PR_CalcPolar(xyZipMark(1), aAngle, 0.5, xyZipMark(3))
			
		End If
		
		'Draw the zip and text
		If g_sClosure = "Back Zip" Or g_sClosure = "Front Zip" Then
			ARMDIA1.PR_DrawLine(xyZipMark(1), xyZipMark(2))
			BDYUTILS.PR_AddDBValueToLast("Zipper", "1")
			ARMDIA1.PR_DrawLine(xyZipMark(2), xyZipMark(3))
			BDYUTILS.PR_AddDBValueToLast("Zipper", "1")
			ARMDIA1.PR_InsertSymbol("TextAsSymbol", xyZipMark(4), 1, 0)
			'Adjust length for turtle necks
			If g_sFrontNeckStyle = "Turtle" Then
				nLength = nLength + g_nFrontNeckSize
			End If
			BDYUTILS.PR_AddDBValueToLast("Zipper", "1")
			nLength = (nLength - 0.125) / 0.95
			sText = ARMEDDIA1.fnInchestoText(nLength) & "\"" " & g_sClosure
			BDYUTILS.PR_AddDBValueToLast("Data", sText)
			BDYUTILS.PR_AddFormatedDBValueToLast("Data", "", nLength, " " & g_sClosure)
		End If
		
		'Bracups
		If g_sSex = "Female" Then PR_DrawBra()
		
		'Draw Body Suit/Brief
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		
		'Draw Front Neck
		'Temporary 3 point polyline - will use arc for all options later
		If g_sFrontNeckStyle = "Measured Scoop" Or g_sFrontNeckStyle = "Measured V neck" Then
			cFrontNeckArc.X(1) = xyCutOutFrontNeck.X
			cFrontNeckArc.y(1) = xyCutOutFrontNeck.y
			cFrontNeckArc.X(2) = xyFrontNeckMid1.X
			cFrontNeckArc.y(2) = xyFrontNeckMid1.y
			cFrontNeckArc.X(3) = xyFrontNeckMid2.X
			cFrontNeckArc.y(3) = xyFrontNeckMid2.y
			cFrontNeckArc.X(4) = xyCutOut8.X
			cFrontNeckArc.y(4) = xyCutOut8.y
			cFrontNeckArc.n = 4
			ARMDIA1.PR_DrawFitted(cFrontNeckArc)
		Else
			ARMDIA1.PR_DrawArc(xyFrontNeckArcCentre, xyCutOutFrontNeck, xyCutOut8)
		End If
		
		'Draw Cutout
		If g_bMissCutOut7 Then
			cFrontCutOutPoly.X(1) = xyCutOut5.X
			cFrontCutOutPoly.y(1) = xyCutOut5.y
			cFrontCutOutPoly.X(2) = xyCutOut6.X
			cFrontCutOutPoly.y(2) = xyCutOut6.y
			cFrontCutOutPoly.X(3) = xyCutOut8.X
			cFrontCutOutPoly.y(3) = xyCutOut8.y
			cFrontCutOutPoly.n = 3
			ARMDIA1.PR_DrawPoly(cFrontCutOutPoly)
			BDYUTILS.PR_AddDBValueToLast("curvetype", "CutOutCurve")
			BDYUTILS.PR_AddDBValueToLast("Zipper", "Front")
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
			'Draw DB markers for Zipper macros
			BDYUTILS.PR_DrawMarker(xyCutOut8) 'Position this Marker at the neck for simplicity
			BDYUTILS.PR_AddDBValueToLast("Zipper", "FrontC")
		Else
			cFrontCutOutPoly.X(1) = xyCutOut5.X
			cFrontCutOutPoly.y(1) = xyCutOut5.y
			cFrontCutOutPoly.X(2) = xyCutOut6.X
			cFrontCutOutPoly.y(2) = xyCutOut6.y
			cFrontCutOutPoly.X(3) = xyCutOut7.X
			cFrontCutOutPoly.y(3) = xyCutOut7.y
			cFrontCutOutPoly.X(4) = xyCutOut8.X
			cFrontCutOutPoly.y(4) = xyCutOut8.y
			cFrontCutOutPoly.n = 4
			ARMDIA1.PR_DrawPoly(cFrontCutOutPoly)
			BDYUTILS.PR_AddDBValueToLast("curvetype", "CutOutCurve")
			BDYUTILS.PR_AddDBValueToLast("Zipper", "Front")
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
			BDYUTILS.PR_DrawMarker(xyCutOut7)
			BDYUTILS.PR_DrawMarker(xyCutOut7)
			BDYUTILS.PR_AddDBValueToLast("curvetype", "CutOutMarker")
			'Draw DB markers for Zipper macros
			BDYUTILS.PR_DrawMarker(xyCutOut7)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "FrontC")
		End If
		
		cBackCutOutPoly.X(1) = xyCutOut3.X
		cBackCutOutPoly.y(1) = xyCutOut3.y
		cBackCutOutPoly.X(2) = xyCutOut2.X
		cBackCutOutPoly.y(2) = xyCutOut2.y
		cBackCutOutPoly.X(3) = xyCutOutBackNeck.X
		cBackCutOutPoly.y(3) = xyCutOutBackNeck.y
		cBackCutOutPoly.n = 3
		ARMDIA1.PR_DrawFitted(cBackCutOutPoly)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "Back")
		BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
		
		'Draw Back Neck
		cBackNeckArc.X(1) = xyCutOutBackNeck.X
		cBackNeckArc.y(1) = xyCutOutBackNeck.y
		cBackNeckArc.X(2) = xyBackNeckMid1.X
		cBackNeckArc.y(2) = xyBackNeckMid1.y
		cBackNeckArc.X(3) = xyBackNeckMid2.X
		cBackNeckArc.y(3) = xyBackNeckMid2.y
		cBackNeckArc.X(4) = xyProfileNeck.X
		cBackNeckArc.y(4) = xyProfileNeck.y
		cBackNeckArc.n = 4
		ARMDIA1.PR_DrawFitted(cBackNeckArc)
		
		If InStr(g_sLegStyle, "Brief") > 0 Then
			'Draw back profile
			cProfilePoly.X(1) = xyProfileNeck.X
			cProfilePoly.y(1) = xyProfileNeck.y
			cProfilePoly.X(2) = xyProfileChest.X
			cProfilePoly.y(2) = xyProfileChest.y
			cProfilePoly.X(3) = xyProfileWaist.X
			cProfilePoly.y(3) = xyProfileWaist.y
			cProfilePoly.X(4) = xyProfileButt.X
			cProfilePoly.y(4) = xyProfileButt.y
			cProfilePoly.X(5) = xyGroinExtraPT1.X
			cProfilePoly.y(5) = xyGroinExtraPT1.y
			cProfilePoly.X(6) = xyGroinExtraPT.X
			cProfilePoly.y(6) = xyGroinExtraPT.y
			cProfilePoly.X(7) = xyProfileGroin.X
			cProfilePoly.y(7) = xyProfileGroin.y
			
			If g_bDrawBriefCurve Then
				cProfilePoly.X(8) = xyFold.X
				cProfilePoly.y(8) = xyFold.y
				cProfilePoly.X(9) = xyProfileThighExtraPT.X
				cProfilePoly.y(9) = xyProfileThighExtraPT.y
				cProfilePoly.X(10) = xyProfileThigh.X
				cProfilePoly.y(10) = xyProfileThigh.y
				cProfilePoly.n = 10
			Else
				cProfilePoly.X(8) = xyProfileThighExtraPT.X
				cProfilePoly.y(8) = xyProfileThighExtraPT.y
				cProfilePoly.X(9) = xyProfileThigh.X
				cProfilePoly.y(9) = xyProfileThigh.y
				cProfilePoly.n = 9
			End If
			ARMDIA1.PR_DrawFitted(cProfilePoly)
			BDYUTILS.PR_AddDBValueToLast("curvetype", "BackCurve")
			BDYUTILS.PR_AddDBValueToLast("Leg", "Left&Right")
			
			If g_bDiffThigh Then
				'Revise data above
				BDYUTILS.PR_AddDBValueToLast("Leg", g_sSmallestThighGiven)
				'Draw the largest thigh profile
				
				aAngle = ARMDIA1.FN_CalcAngle(xyLT_ButtocksArcCentre, xyProfileButt)
				aInc = (aAngle - ARMDIA1.FN_CalcAngle(xyLT_ButtocksArcCentre, xyLT_Fold)) / 4
				cProfilePoly.n = 0
				For ii = 1 To 5
					ARMDIA1.PR_CalcPolar(xyLT_ButtocksArcCentre, aAngle, g_nButtRadius, xyTmp)
					cProfilePoly.n = cProfilePoly.n + 1
					cProfilePoly.X(cProfilePoly.n) = xyTmp.X
					cProfilePoly.y(cProfilePoly.n) = xyTmp.y
					aAngle = aAngle - aInc
				Next ii
				cProfilePoly.n = cProfilePoly.n + 1
				cProfilePoly.X(cProfilePoly.n) = xyLT_ProfileThigh.X
				cProfilePoly.y(cProfilePoly.n) = xyLT_ProfileThigh.y
				
				ARMDIA1.PR_SetLayer("Template" & g_sLargestThighGiven)
				ARMDIA1.PR_DrawFitted(cProfilePoly)
				BDYUTILS.PR_AddDBValueToLast("curvetype", "BackCurveLargest")
				BDYUTILS.PR_AddDBValueToLast("Leg", g_sLargestThighGiven)
				
				'Label Left and Right
				ARMDIA1.PR_SetLayer("Notes")
				'Insert ArrowHead symbol
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyLT_Fold, INCH1_4, INCH1_8, -90)
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyProfileThighExtraPT, INCH1_4, INCH1_8, -90)
				ARMDIA1.PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
				ARMDIA1.PR_MakeXY(xyText, xyProfileThighExtraPT.X, xyProfileThighExtraPT.y - INCH1_8)
				ARMDIA1.Pr_DrawText(g_sSmallestThighGiven, xyText, 0.1)
				ARMDIA1.PR_MakeXY(xyText, xyLT_Fold.X, xyLT_Fold.y - INCH1_8)
				ARMDIA1.Pr_DrawText(g_sLargestThighGiven, xyText, 0.1)
				
				'Reset back to original
				ARMDIA1.PR_SetLayer("Template" & g_sSmallestThighGiven)
				
			End If
		End If
		
		If g_sLegStyle = "Panty" Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ARMDIA1.PR_MakeXY(xyLegStart, -ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength), 0)
			ARMDIA1.PR_SetLayer("Construct")
			If g_bDrawSingleLeg Then
				PR_DrawLegTemplate("Left", xyLegStart)
				'UPGRADE_WARNING: Couldn't resolve default property of object SmallestLegCurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SmallestLegCurve = g_LeftLegProfile
				nSmallestLegXOffset = g_nLeftLegLength
				If g_bLeftAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
			Else
				PR_DrawLegTemplate("Both", xyLegStart)
				If g_sSmallestThighGiven = "Left" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object SmallestLegCurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SmallestLegCurve = g_LeftLegProfile
					'UPGRADE_WARNING: Couldn't resolve default property of object LargestLegCurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LargestLegCurve = g_RightLegProfile
					nLargestLegXOffset = g_nRightLegLength
					nSmallestLegXOffset = g_nLeftLegLength
					If g_bLeftAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
					If g_bRightAboveKnee Then sLargestElasticText = "ELASTIC" Else sLargestElasticText = "NO ELASTIC"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object SmallestLegCurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SmallestLegCurve = g_RightLegProfile
					'UPGRADE_WARNING: Couldn't resolve default property of object LargestLegCurve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LargestLegCurve = g_LeftLegProfile
					nLargestLegXOffset = g_nLeftLegLength
					nSmallestLegXOffset = g_nRightLegLength
					If g_bLeftAboveKnee Then sLargestElasticText = "ELASTIC" Else sLargestElasticText = "NO ELASTIC"
					If g_bRightAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
				End If
				
			End If
			cProfilePoly.X(1) = xyProfileNeck.X
			cProfilePoly.y(1) = xyProfileNeck.y
			cProfilePoly.X(2) = xyProfileChest.X
			cProfilePoly.y(2) = xyProfileChest.y
			cProfilePoly.X(3) = xyProfileWaist.X
			cProfilePoly.y(3) = xyProfileWaist.y
			cProfilePoly.X(4) = xyProfileButt.X
			cProfilePoly.y(4) = xyProfileButt.y
			cProfilePoly.X(5) = xyGroinExtraPT1.X
			cProfilePoly.y(5) = xyGroinExtraPT1.y
			cProfilePoly.X(6) = xyGroinExtraPT.X
			cProfilePoly.y(6) = xyGroinExtraPT.y
			' cProfilePoly.x(7) = xyProfileGroin.x
			' cProfilePoly.y(7) = xyProfileGroin.y
			cProfilePoly.n = 6
			
			'Join leg
			For ii = SmallestLegCurve.n To 1 Step -1
				cProfilePoly.n = cProfilePoly.n + 1
				cProfilePoly.X(cProfilePoly.n) = SmallestLegCurve.X(ii) - nSmallestLegXOffset
				cProfilePoly.y(cProfilePoly.n) = SmallestLegCurve.y(ii)
			Next ii
			If g_bDrawSingleLeg Then ARMDIA1.PR_SetLayer("TemplateLeft") Else ARMDIA1.PR_SetLayer("Template" & g_sSmallestThighGiven)
			ARMDIA1.PR_DrawFitted(cProfilePoly)
			
			BDYUTILS.PR_AddDBValueToLast("curvetype", "BackCurve")
			
			'Label legs
			ARMDIA1.PR_SetLayer("Notes")
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
			If Not g_bDrawSingleLeg Then
				BDYUTILS.PR_AddDBValueToLast("Leg", g_sSmallestThighGiven)
			Else
				BDYUTILS.PR_AddDBValueToLast("Leg", "Left&Right")
			End If
			
			'Draw closure
			If g_bDrawSingleLeg Then ARMDIA1.PR_SetLayer("TemplateLeft") Else ARMDIA1.PR_SetLayer("Template" & g_sSmallestThighGiven)
			ARMDIA1.PR_MakeXY(xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n))
			ARMDIA1.PR_MakeXY(xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n))
			ARMDIA1.PR_DrawLine(xyTmp, xyTmp1)
			
			'Add elastic text for smallest
			ARMDIA1.PR_SetLayer("Notes")
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 270
			ARMDIA1.PR_MakeXY(xyText, xyTmp1.X + 0.3, 0)
			ARMDIA1.Pr_DrawText(sSmallestElasticText, xyText, 0.125)
			
			'Draw other leg
			If Not g_bDrawSingleLeg Then
				
				'Draw the largest thigh profile
				aAngle = ARMDIA1.FN_CalcAngle(xyButtocksArcCentre, xyProfileButt)
				aInc = (aAngle - ARMDIA1.FN_CalcAngle(xyButtocksArcCentre, xyGroinExtraPT)) / 4
				cProfilePoly.n = 0
				For ii = 1 To 5
					ARMDIA1.PR_CalcPolar(xyButtocksArcCentre, aAngle, g_nButtRadius, xyTmp)
					cProfilePoly.n = cProfilePoly.n + 1
					cProfilePoly.X(cProfilePoly.n) = xyTmp.X
					cProfilePoly.y(cProfilePoly.n) = xyTmp.y
					aAngle = aAngle - aInc
				Next ii
				
				'Join leg
				For ii = LargestLegCurve.n To 1 Step -1
					cProfilePoly.n = cProfilePoly.n + 1
					cProfilePoly.X(cProfilePoly.n) = LargestLegCurve.X(ii) - nLargestLegXOffset
					cProfilePoly.y(cProfilePoly.n) = LargestLegCurve.y(ii)
				Next ii
				
				ARMDIA1.PR_SetLayer("Template" & g_sLargestThighGiven)
				ARMDIA1.PR_DrawFitted(cProfilePoly)
				BDYUTILS.PR_AddDBValueToLast("curvetype", "BackCurveLargest")
				BDYUTILS.PR_AddDBValueToLast("Leg", g_sLargestThighGiven)
				ARMDIA1.PR_MakeXY(xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n))
				ARMDIA1.PR_MakeXY(xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n))
				ARMDIA1.PR_DrawLine(xyTmp, xyTmp1)
				
				'Add elastic text
				ARMDIA1.PR_SetLayer("Notes")
				ARMDIA1.PR_MakeXY(xyText, xyTmp1.X + 0.3, 0)
				ARMDIA1.Pr_DrawText(sLargestElasticText, xyText, 0.125)
			End If
		End If
		
		
		'Setup which side is being drawn
		If g_sSide = "Left" Then
			g_sSleeveType = g_sLeftSleeve
		Else
			g_sSleeveType = g_sRightSleeve
		End If
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
		
		'Draw Axilla
		If g_sSleeveType = "Sleeveless" Then
			'Draw Sleeveless Axilla
			cRaglanPoly.X(1) = xyCutOutFrontNeck.X
			cRaglanPoly.y(1) = xyCutOutFrontNeck.y
			cRaglanPoly.X(2) = xyRaglan1.X
			cRaglanPoly.y(2) = xyRaglan1.y
			cRaglanPoly.X(3) = xyRaglan2.X
			cRaglanPoly.y(3) = xyRaglan2.y
			cRaglanPoly.n = 3
			ARMDIA1.PR_DrawPoly(cRaglanPoly)
			cRaglanPoly.X(1) = xyRaglan5.X
			cRaglanPoly.y(1) = xyRaglan5.y
			cRaglanPoly.X(2) = xyRaglan6.X
			cRaglanPoly.y(2) = xyRaglan6.y
			cRaglanPoly.X(3) = xyProfileNeckMirror.X
			cRaglanPoly.y(3) = xyProfileNeckMirror.y
			cRaglanPoly.n = 3
			ARMDIA1.PR_DrawPoly(cRaglanPoly)
			
			'Draw Raglan Curves
			ARMDIA1.PR_SetLayer("Template" & g_sAxillaSide)
			nFlip = True
			MESHCALC.PR_CalcRaglan(xyRaglan4, xyRaglan2, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), -1)
			ARMDIA1.PR_DrawFitted(cRaglanCurve)
			nFlip = False
			MESHCALC.PR_CalcRaglan(xyRaglan4, xyRaglan5, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1)
			ARMDIA1.PR_DrawFitted(cRaglanCurve)
			If g_bDiffAxillaHeight Then
				ARMDIA1.PR_SetLayer("Template" & g_sAxillaSideLow)
				nFlip = False
				MESHCALC.PR_CalcRaglan(xyRaglan4LowAxilla, xyRaglan5, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1)
				ARMDIA1.PR_DrawFitted(cRaglanCurve)
				nFlip = True
				MESHCALC.PR_CalcRaglan(xyRaglan4LowAxilla, xyRaglan2, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), -1)
				ARMDIA1.PR_DrawFitted(cRaglanCurve)
				
				ARMDIA1.PR_SetLayer("Notes")
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyRaglan4LowAxilla, INCH1_4, INCH1_8, 180)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_nCurrTextAngle = 270
				ARMDIA1.PR_MakeXY(xyText, xyRaglan4LowAxilla.X - 0.3, xyRaglan4LowAxilla.y)
				ARMDIA1.Pr_DrawText(g_sAxillaSideLow, xyText, 0.125)
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyRaglan4, INCH1_4, INCH1_8, 180)
				ARMDIA1.PR_MakeXY(xyText, xyRaglan4.X - 0.3, xyRaglan4.y)
				ARMDIA1.Pr_DrawText(g_sAxillaSide, xyText, 0.125)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_nCurrTextAngle = 0
			End If
		Else
			'Draw Sleeved Axilla
			ARMDIA1.PR_DrawLine(xyCutOutFrontNeck, xyRaglan1)
			ARMDIA1.PR_DrawLine(xyRaglan3, xyProfileNeckMirror)
			
			'Draw Raglan Curves
			ARMDIA1.PR_SetLayer("Template" & g_sAxillaSide)
			nFlip = False
			MESHCALC.PR_CalcRaglan(xyRaglan2, xyRaglan3, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1)
			ARMDIA1.PR_DrawFitted(cRaglanCurve)
			nFlip = True
			nLength = ARMDIA1.FN_CalcLength(xyRaglan2, xyOpen(2))
			MESHCALC.PR_CalcRaglan(xyRaglan2, xyRaglan1, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), nLength)
			ARMDIA1.PR_DrawFitted(cRaglanCurve)
			
			Select Case g_sAxillaType
				Case "Open"
					ARMDIA1.PR_SetLayer("Notes")
					BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(1), INCH1_4, INCH1_8, aOpen(1))
					BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(2), INCH1_4, INCH1_8, aOpen(2))
					sAxillaText = "Open"
					
				Case "Mesh"
					PR_DrawMeshConstruction(xyRaglan2, xyMesh, sAxillaText, nMeshSize)
					nFlip = False
					nMeshDistanceAlongRaglan = 0
					If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
						MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
					Else
						PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSide)
					End If
					'Calculate the mesh seam line along back neck raglan
					'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
					nFlip = True
					If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
						MsgBox("Unable to calculate mesh along front neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
					Else
						PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSide)
					End If
					sText = Trim(Str(nMeshDistanceAlongRaglan)) & "," & Trim(Str(nMeshSize))
					g_sMeshLeft = sText
					g_sMeshRight = sText
					
					'regular axilla (because CAD Block can't handle 1/2's)
				Case "Regular 1½"""
					sAxillaText = "Regular 1-1/2\"""
				Case "Regular 2"""
					sAxillaText = "Regular 2\"""
				Case "Regular 2½"""
					sAxillaText = "Regular 2-1/2\"""
					
				Case Else
					sAxillaText = g_sAxillaType
			End Select
			
			'Draw Lowest Axilla and Label
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 270
			If g_bDiffAxillaHeight Then
				ARMDIA1.PR_SetLayer("Template" & g_sAxillaSideLow)
				nFlip = False
				MESHCALC.PR_CalcRaglan(xyRaglan2LowAxilla, xyRaglan3, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1)
				ARMDIA1.PR_DrawFitted(cRaglanCurve)
				nFlip = True
				nLength = ARMDIA1.FN_CalcLength(xyRaglan2LowAxilla, xyOpen(4))
				MESHCALC.PR_CalcRaglan(xyRaglan2LowAxilla, xyRaglan1, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), nLength)
				ARMDIA1.PR_DrawFitted(cRaglanCurve)
				
				Select Case g_sAxillaTypeLow
					Case "Open"
						ARMDIA1.PR_SetLayer("Notes")
						BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(3), INCH1_4, INCH1_8, aOpen(3))
						BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(4), INCH1_4, INCH1_8, aOpen(4))
						sAxillaTextLow = "Open"
					Case "Mesh"
						PR_DrawMeshConstruction(xyRaglan2LowAxilla, xyMesh, sAxillaTextLow, nMeshSize)
						nFlip = False
						nMeshDistanceAlongRaglan = 0
						If Not FN_CalcAxillaMesh(xyRaglan2LowAxilla, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
							MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
						Else
							PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSideLow)
						End If
						'Calculate the mesh seam line along back neck raglan
						'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
						nFlip = True
						If Not FN_CalcAxillaMesh(xyRaglan2LowAxilla, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
							MsgBox("Unable to calculate mesh along front neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
						Else
							PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSideLow)
						End If
						sText = Trim(Str(nMeshDistanceAlongRaglan)) & "," & Trim(Str(nMeshSize))
						If g_sAxillaSideLow = "Left" Then g_sMeshLeft = sText Else g_sMeshRight = sText
						
						
						'regular axilla
					Case "Regular 1½"""
						sAxillaTextLow = "Regular 1-1/2\"""
					Case "Regular 2"""
						sAxillaTextLow = "Regular 2\"""
					Case "Regular 2½"""
						sAxillaTextLow = "Regular 2-1/2\"""
					Case Else
						sAxillaTextLow = g_sAxillaTypeLow
						
				End Select
				
				ARMDIA1.PR_SetLayer("Notes")
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyRaglan2LowAxilla, INCH1_4, INCH1_8, 180)
				ARMDIA1.PR_MakeXY(xyText, xyRaglan2LowAxilla.X - 0.3, xyRaglan2LowAxilla.y)
				ARMDIA1.Pr_DrawText(g_sAxillaSideLow & " " & sAxillaTextLow, xyText, 0.125)
				
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyRaglan2, INCH1_4, INCH1_8, 180)
				ARMDIA1.PR_MakeXY(xyText, xyRaglan2.X - 0.3, xyRaglan2.y)
				ARMDIA1.Pr_DrawText(g_sAxillaSide & " " & sAxillaText, xyText, 0.125)
				
			Else
				'Take care of different axilla types at same axilla height
				If g_sLeftSleeve <> g_sRightSleeve Then
					'Do the opposite axilla to that done above
					If g_sLeftSleeve = g_sAxillaType Then
						g_sAxillaType = g_sRightSleeve
						sThis = "Right"
						sOther = "Left"
					Else
						g_sAxillaType = g_sLeftSleeve
						sThis = "Left"
						sOther = "Right"
					End If
					Select Case g_sAxillaType
						Case "Open"
							ARMDIA1.PR_SetLayer("Notes")
							BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(1), INCH1_4, INCH1_8, aOpen(1))
							BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyOpen(2), INCH1_4, INCH1_8, aOpen(2))
							sAxillaTextThis = "Open"
						Case "Mesh"
							PR_DrawMeshConstruction(xyRaglan2, xyMesh, sAxillaText, nMeshSize)
							nFlip = False
							nMeshDistanceAlongRaglan = 0
							If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
								MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
							Else
								PR_DrawMesh(MeshProfile, MeshSeam, sThis)
							End If
							'Calculate the mesh seam line along back neck raglan
							'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
							nFlip = True
							If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
								MsgBox("Unable to calculate mesh along front neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
							Else
								PR_DrawMesh(MeshProfile, MeshSeam, sThis)
							End If
							
							sText = Trim(Str(nMeshDistanceAlongRaglan)) & "," & Trim(Str(nMeshSize))
							If sThis = "Left" Then g_sMeshLeft = sText Else g_sMeshRight = sText
							
							'regular axilla
						Case "Regular 1½"""
							sAxillaTextThis = "Regular 1-1/2\"""
						Case "Regular 2"""
							sAxillaTextThis = "Regular 2\"""
						Case "Regular 2½"""
							sAxillaTextThis = "Regular 2-1/2\"""
						Case Else
							sAxillaTextThis = g_sAxillaTypeLow
							
					End Select
					
					ARMDIA1.PR_SetLayer("Notes")
					ARMDIA1.PR_MakeXY(xyText, xyRaglan2.X - 0.3, xyRaglan2.y)
					ARMDIA1.Pr_DrawText(sOther & " " & sAxillaText & "\n" & sThis & " " & sAxillaTextThis, xyText, 0.125)
					
				Else
					ARMDIA1.PR_SetLayer("Notes")
					ARMDIA1.PR_MakeXY(xyText, xyRaglan2.X - 0.3, xyRaglan2.y)
					ARMDIA1.Pr_DrawText(sAxillaText, xyText, 0.125)
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
			
			'Update the database for use by sleeve drawing module
			PR_UpdateDB()
			
		End If
		
		'Draw Crotch (based on some fixed rules checked in BodySuit dialogue box)
		ARMDIA1.PR_MakeXY(xyCrotchText, xyCrotchMarker.X - 0.375, xyCrotchMarker.y)
		BDYUTILS.PR_DrawMarker(xyCrotchMarker)
		
		'Draw Crotch styles
		ARMDIA1.PR_SetLayer("Template" & "Left")
		
		Select Case g_sCrotchStyle
			Case "Open Crotch"
				'Draw Cutout Arc
				ARMDIA1.PR_DrawLine(xyBackCrotch1, xyBackCrotch2)
				ARMDIA1.PR_DrawLine(xyFrontCrotch1, xyFrontCrotch2)
				ARMDIA1.PR_DrawArc(xyBackCrotchFilletCentre, xyBackCrotch2, xyBackCrotch3)
				ARMDIA1.PR_DrawArc(xyFrontCrotchFilletCentre, xyFrontCrotch3, xyFrontCrotch2)
				
				If g_bExtremeCrotch Then
					'Extreme crotch
					'Back crotch
					If xyBackCrotch1.X < xyCutOut9.X Then
						ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyCutOut9, xyBackCrotch1)
						ARMDIA1.PR_DrawLine(xyCutOut9, xyCutOut3)
					ElseIf xyBackCrotch1.X < xyCutOut3.X Then 
						ARMDIA1.PR_DrawLine(xyBackCrotch1, xyCutOut3)
					End If
					'Reuse xyBackCrotch3 here as the start of the arc
					If xyBackCrotch3.X > xyCutOut9.X Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyTmp = xyCutOut9 : xyTmp.y = xyBackCrotch3.y
						ARMDIA1.PR_DrawLine(xyBackCrotch3, xyTmp)
						xyBackCrotch3.X = xyCutOut9.X
					Else
						'Do nothing
					End If
					'Front crotch
					If xyFrontCrotch1.X < xyCutOut10.X Then
						ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyCutOut10, xyFrontCrotch1)
						ARMDIA1.PR_DrawLine(xyCutOut10, xyCutOut5)
					ElseIf xyFrontCrotch1.X < xyCutOut5.X Then 
						ARMDIA1.PR_DrawLine(xyFrontCrotch1, xyCutOut5)
					End If
					'Reuse xyFrontCrotch3 here as the start of the arc
					If xyFrontCrotch3.X > xyCutOut10.X Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyTmp = xyCutOut10 : xyTmp.y = xyFrontCrotch3.y
						ARMDIA1.PR_DrawLine(xyFrontCrotch3, xyTmp)
						xyFrontCrotch3.X = xyCutOut10.X
					Else
						'Do nothing
					End If
					'Note:- Reuse of xyFrontCrotch3and xyBackCrotch3 here as the start and end of the arc
					ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyBackCrotch3, xyFrontCrotch3)
					
					
				Else
					'Normal crotch
					If xyBackCrotch1.X <> xyCutOut3.X And xyBackCrotch1.y <> xyCutOut3.y Then
						ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyCutOut3, xyBackCrotch1)
					End If
					If xyFrontCrotch1.X <> xyCutOut5.X And xyFrontCrotch1.y <> xyCutOut5.y Then
						ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyFrontCrotch1, xyCutOut5)
					End If
					ARMDIA1.PR_DrawArc(xyCutOutArcCentre, xyBackCrotch3, xyFrontCrotch3)
				End If
				
				If g_sSex = "Male" Then
					sCrotchText = ""
					If Not g_nAdult Then
						sCrotchText = "1/2\"" ELASTIC"
					End If
				Else
					sCrotchText = "1/2\"" ELASTIC"
				End If
				
			Case "Horizontal Fly"
				sCrotchText = "Horizontal Fly," & g_sFlySize
				
			Case "Diagonal Fly"
				sCrotchText = "Diagonal Fly," & g_sFlySize
				ARMDIA1.PR_SetLayer("Notes")
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyCutOut3, INCH1_4, INCH1_8, 90)
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyCutOut5, INCH1_4, INCH1_8, 270)
				
			Case "Snap Crotch"
				
				sCrotchText = "GUSSET, MESH " & g_sGussetSize
				
				'DRAW Gusset rectangle
				ARMDIA1.PR_SetLayer("Template" & g_sSide)
				ARMDIA1.PR_MakeXY(xyLLGussetRectangle, xyProfileButt.X - 2, xyProfileButt.y + 2)
				ARMDIA1.PR_MakeXY(xyURGussetRectangle, xyLLGussetRectangle.X + (xyStrap2.X - xyStrap4.X), xyLLGussetRectangle.y + (xyStrap1.y - xyStrap2.y))
				BDYUTILS.PR_DrawRectangle(xyLLGussetRectangle, xyURGussetRectangle)
				
				ARMDIA1.PR_SetTextData(LEFT_, TOP_, CURRENT, CURRENT, CURRENT)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_nCurrTextAngle = 0
				ARMDIA1.PR_SetLayer("Notes")
				ARMDIA1.PR_MakeXY(xyText, xyLLGussetRectangle.X + 0.25, xyURGussetRectangle.y - 0.25)
				sText = g_sPatient & "\n" & g_sWorkOrder
				ARMDIA1.Pr_DrawText(sText, xyText, 0.125)
				
				'Draw DB markers for Zipper macros
				ARMDIA1.PR_SetLayer("Construct")
				BDYUTILS.PR_DrawMarker(xyStrap1)
				BDYUTILS.PR_AddDBValueToLast("Zipper", "Snap")
				
			Case "Gusset"
				sCrotchText = "Gusset, " & g_sGussetSize
				
		End Select
		
		'Draw Crotch Text and marker
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 270
		If sCrotchText <> "" Then ARMDIA1.Pr_DrawText(sCrotchText, xyCrotchText, 0.125)
		
		'Insert ArrowHead symbol
		BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyCrotchMarker, INCH1_4, INCH1_8, 180)
		
		'Draw Leg style
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		
		'Draw CutOut Arc
		If Not g_sCrotchStyle = "Open Crotch" Then
			If g_bExtremeCrotch Then
				ARMDIA1.PR_DrawLine(xyCutOut5, xyCutOut10)
				ARMDIA1.PR_DrawLine(xyCutOut9, xyCutOut3)
				BDYUTILS.PR_AddEntityArc(xyCutOutArcCentre, xyCutOut9, xyCutOut10)
			Else
				BDYUTILS.PR_AddEntityArc(xyCutOutArcCentre, xyCutOut3, xyCutOut5)
			End If
		End If
		
		If g_sLegStyle = "Panty" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
			ARMDIA1.PR_SetLayer("Notes")
			If (g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief") Then
				'There is only a single leg draw which is given by the left leg
				ARMDIA1.PR_MakeXY(xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8)
				BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90)
				If g_sLeftLeg = "Brief" Then
					ARMDIA1.Pr_DrawText("Right", xyText, 0.125)
					ARMDIA1.PR_SetLayer("TemplateLeft")
				Else
					ARMDIA1.Pr_DrawText("Left", xyText, 0.125)
					ARMDIA1.PR_SetLayer("TemplateRight")
				End If
				PR_DrawBriefStandard(0.875)
			Else
				If g_bDrawSingleLeg Then
					ARMDIA1.PR_MakeXY(xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8)
					BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90)
					ARMDIA1.Pr_DrawText("Left & Right", xyText, 0.125)
				Else
					ARMDIA1.PR_MakeXY(xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8)
					BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90)
					ARMDIA1.Pr_DrawText("Left", xyText, 0.125)
					ARMDIA1.PR_MakeXY(xyText, xyRightLegLabelPoint.X, xyRightLegLabelPoint.y - INCH1_8)
					BDYUTILS.PR_DrawMarkerNamed("Closed Arrow", xyRightLegLabelPoint, INCH1_4, INCH1_8, -90)
					ARMDIA1.Pr_DrawText("Right", xyText, 0.125)
				End If
			End If
		End If
		
		'Draw on body layer
		ARMDIA1.PR_SetLayer("Template" & g_sSide)
		
		If g_sLegStyle = "Brief" Then
			PR_DrawBriefStandard(0)
		End If
		
		If g_sLegStyle = "Brief-French" Then
			'calculate additional French-cut points
			If g_sSleeveType = "Sleeveless" Then xyFrenchCut.y = xyRaglan4.y Else xyFrenchCut.y = xyRaglan2.y
			If g_nAdult Then xyFrenchCut.X = xySeamFold.X + 2.5 Else xyFrenchCut.X = xySeamFold.X + 2
			PR_DrawBriefFrenchCut()
		End If
		
		
		'Draw Lower axilla construction line
		If g_bDiffAxillaHeight Then
			ARMDIA1.PR_SetLayer("Construct")
			PR_DrawVerticalLineThruPoint(xySeamChestAxillaLow, nVertLinesOffset)
		End If
		
		'Set Neutral Layer
		ARMDIA1.PR_SetLayer("1")
		
		'PR_DebugPointText
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		
		
		
		
	End Sub
	
	Private Sub PR_DebugPointText()
		Dim nn As Double
		nn = 0.1
		ARMDIA1.PR_SetTextData(RIGHT_, BOTTOM_, CURRENT, CURRENT, CURRENT)
		'    Pr_DrawText "xyCrotchMarker", xyCrotchMarker, nn
		'    PR_DrawText "xyProfileThigh", xyProfileThigh, nn
		'    PR_DrawText "xyProfileGroin", xyProfileGroin, nn
		'    PR_DrawText "xyFold", xyFold, nn
		'    PR_DrawText "xyLT_ProfileThigh", xyLT_ProfileThigh, nn
		'    PR_DrawText "xyLT_ProfileGroin", xyLT_ProfileGroin, nn
		'    PR_DrawText "xyLT_Fold", xyLT_Fold, nn
		'    PR_DrawText "xyLT_ButtocksArcCentre", xyLT_ButtocksArcCentre, nn
		'    PR_DrawText "xyProfileButt", xyProfileButt, nn
		'    PR_DrawText "xyProfileBrief", xyProfileBrief, nn
		
		'    Pr_DrawText "xyProfileWaist", xyProfileWaist, nn
		'    Pr_DrawText "xyProfileChest", xyProfileChest, nn
		'    Pr_DrawText "xyProfileNeck", xyProfileNeck, nn
		'    PR_DrawText "xyCutOutArcCentre", xyCutOutArcCentre, nn
		'    Pr_DrawText "xyFold", xyFold, nn
		'    PR_DrawText "xyButtocksArcCentre", xyButtocksArcCentre, nn
		'    Pr_DrawText "xyCutOutBackNeck", xyCutOutBackNeck, nn
		'    Pr_DrawText "xyCutOut2", xyCutOut2, nn
		ARMDIA1.Pr_DrawText("xyCutOut3", xyCutOut3, nn)
		ARMDIA1.Pr_DrawText("xyCutOut4", xyCutOut4, nn)
		'    Pr_DrawText "xyCutOut5", xyCutOut5, nn
		'    Pr_DrawText "xyCutOut6", xyCutOut6, nn
		'    Pr_DrawText "xyCutOut7", xyCutOut7, nn
		'   Pr_DrawText "xyCutOut8", xyCutOut8, nn
		'
		'    Pr_DrawText "xyBackCrotch1", xyBackCrotch1, nn
		'    Pr_DrawText "xyBackCrotch2", xyBackCrotch2, nn
		'    Pr_DrawText "xyBackCrotch3", xyBackCrotch3, nn
		'    Pr_DrawText "xyBackCrotchFilletCentre", xyBackCrotchFilletCentre, nn
		'   Pr_DrawText "xyFrontCrotch1", xyFrontCrotch1, nn
		'    Pr_DrawText "xyFrontCrotch2", xyFrontCrotch2, nn
		'    Pr_DrawText "xyFrontCrotch3", xyFrontCrotch3, nn
		'    Pr_DrawText "xyFrontCrotchFilletCentre", xyFrontCrotchFilletCentre, nn
		
		'    Pr_DrawText "xyRaglan1", xyRaglan1, nn
		'    Pr_DrawText "xyRaglan2", xyRaglan2, nn
		'    Pr_DrawText "xyRaglan3", xyRaglan3, nn
		'    Pr_DrawText "xyRaglan4", xyRaglan4, nn
		'    Pr_DrawText "xyRaglan5", xyRaglan5, nn
		'    Pr_DrawText "xyFrenchCut", xyFrenchCut, nn
		
		
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
	
	Private Sub PR_DrawBra()
		'Assumes that data has been validated before this
		'procedure is called.
		Dim nDiskXoff, nBraCLOffset, nDiskYoff As Double
		Dim iRightDisk, iDisk, ii, iLeftDisk, iNextDisk As Short
		Dim sBraText, sDisk As String
		Dim xyDisk As XY
		'UPGRADE_WARNING: Lower bound of array xyPt was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim xyPt(3) As XY
		
		If cboLeftCup.SelectedIndex < 7 Or cboLeftCup.SelectedIndex < 7 Or txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "" Then
			'A cup or a disk of some sort has been given
			If cboLeftCup.Text = "None" Then iLeftDisk = -1 Else iLeftDisk = Val(txtLeftDisk.Text)
			If cboRightCup.Text = "None" Then iRightDisk = -1 Else iRightDisk = Val(txtRightDisk.Text)
			
			If iLeftDisk = iRightDisk Then
				iNextDisk = 1
				If iLeftDisk = -1 Then
					sBraText = "NO CUP LEFT & RIGHT"
				Else
					sBraText = "No." & Str(iLeftDisk) & " LEFT & RIGHT"
				End If
			Else
				iNextDisk = 2
				If iLeftDisk = -1 Then
					sBraText = "NO CUP LEFT"
				Else
					sBraText = "No." & Str(iLeftDisk) & " LEFT"
				End If
				
				If iRightDisk = -1 Then
					sBraText = sBraText & "\nNO CUP RIGHT"
				Else
					sBraText = sBraText & "\nNo." & Str(iRightDisk) & " RIGHT"
				End If
			End If
			
			'Loop through both cups and give positioning information
			For ii = 1 To iNextDisk
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
						nDiskXoff = 2.293 + 0.104
						nDiskYoff = 2.882
					Case 6
						nBraCLOffset = 0.625
						nDiskXoff = 2.625 + 0.0272
						nDiskYoff = 3.025
					Case 7
						nBraCLOffset = 0.5
						nDiskXoff = 2.94 + 0.02
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
				
				If iDisk > 0 Then
					'Insert disk
					ARMDIA1.PR_SetLayer("Notes")
					xyDisk.X = (xySeamChest.X - 0.75) - nDiskXoff
					xyDisk.y = (xyCutOut7.y - nBraCLOffset) '- (nDiskYoff * 2)
					
					sDisk = "bradisk" & Trim(Str(iDisk)) & "mirrored"
					ARMDIA1.PR_InsertSymbol(sDisk, xyDisk, 1, 0)
					BDYUTILS.PR_AddDBValueToLast("curvetype", "Bra")
					
					'Draw Construction lines
					ARMDIA1.PR_SetLayer("Construct")
					ARMDIA1.PR_MakeXY(xyPt(1), xySeamChest.X - 0.75, xyCutOut7.y - nBraCLOffset)
					ARMDIA1.PR_MakeXY(xyPt(2), xyPt(1).X, (xyCutOut7.y - nBraCLOffset) - (nDiskYoff * 2))
					ARMDIA1.PR_MakeXY(xyPt(3), xyPt(1).X - (nDiskXoff * 2), xyPt(1).y)
					ARMDIA1.PR_DrawLine(xyPt(1), xyPt(2))
					BDYUTILS.PR_AddDBValueToLast("curvetype", "Bra")
					ARMDIA1.PR_DrawLine(xyPt(1), xyPt(3))
					BDYUTILS.PR_AddDBValueToLast("curvetype", "Bra")
					If ii = iNextDisk Then
						'Insert bra label text
						ARMDIA1.PR_SetLayer("Notes")
						ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
						BDYUTILS.PR_CalcMidPoint(xyPt(2), xyPt(3), xyPt(1))
						ARMDIA1.Pr_DrawText(sBraText, xyPt(1), 0.125)
					End If
				Else
					'No Bra cup.
					'Do nothing as this is dealt with as text.
					'Except for this very special case
					If iNextDisk = 1 Then
						'Insert Text (Special case No Cups Left and Right ????)
						ARMDIA1.PR_SetLayer("Notes")
						ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
						ARMDIA1.PR_MakeXY(xyPt(1), xySeamChest.X - 2.75, xySeamChest.y - 2.5)
						ARMDIA1.Pr_DrawText(sBraText, xyPt(1), 0.125)
					End If
				End If
			Next ii
		End If
	End Sub
	
	Private Sub PR_DrawBriefFrenchCut()
		'UPGRADE_WARNING: Arrays in structure cBriefPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cBriefPoly As curve
		'UPGRADE_WARNING: Arrays in structure cCrotchPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cCrotchPoly As curve
		Dim nArc As Double
		Dim xyArc As XY
		Dim aAngle, aArcAngle, aConstruct As Double
		Dim ii As Short
		Dim xyTmp As XY
		
		If g_bDiffThigh Then
			ARMDIA1.PR_DrawLine(xyLT_ProfileThigh, xyProfileBrief)
			ARMDIA1.PR_DrawLine(xyProfileThighMirror1, xyLT_ProfileThighMirror)
		Else
			ARMDIA1.PR_DrawLine(xyProfileThigh, xyProfileBrief)
			ARMDIA1.PR_DrawLine(xyProfileThighMirror1, xyProfileThighMirror)
		End If
		cBriefPoly.n = 0
		
		If g_sCrotchStyle <> "Snap Crotch" Then
			'All other crotch styles
			cBriefPoly.n = 0
			aConstruct = ARMDIA1.FN_CalcAngle(xyLegPoint, xyProfileBrief)
			aConstruct = aConstruct - ((aConstruct - 90) / 2)
			nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
			' PR_MakeXY xyArc, xyProfileBrief.X + nArc, xyProfileBrief.Y
			ARMDIA1.PR_CalcPolar(xyProfileBrief, aConstruct - 90, nArc, xyArc)
			aAngle = ARMDIA1.FN_CalcAngle(xyArc, xyProfileBrief)
			aArcAngle = 20
			For ii = 1 To 3
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle + (aArcAngle / 2)
			Next ii
			
			cBriefPoly.n = cBriefPoly.n + 1
			cBriefPoly.X(cBriefPoly.n) = xyLegPoint.X
			cBriefPoly.y(cBriefPoly.n) = xyLegPoint.y
			
			nArc = (xyFrenchCut.X - xySeamFold.X)
			If Not g_nAdult Then
				nArc = nArc * 0.5
				aAngle = 40
				aArcAngle = 80
			Else
				aAngle = 45
				aArcAngle = 90
			End If
			ARMDIA1.PR_MakeXY(xyArc, xyFrenchCut.X - nArc, xyFrenchCut.y)
			For ii = 1 To 5
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle - (aArcAngle / 4)
			Next ii
			
			nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
			If Not g_nAdult Then
				nArc = nArc * 0.25
				ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror.X + nArc, xyProfileThighMirror.y)
				nArc = ARMDIA1.FN_CalcLength(xyProfileThighMirror1, xyArc)
				aArcAngle = 40
				aAngle = ARMDIA1.FN_CalcAngle(xyProfileThighMirror1, xyArc) - aArcAngle
			Else
				ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y)
				aArcAngle = 40
				aAngle = 180 - aArcAngle
			End If
			ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y)
			aArcAngle = 40
			aAngle = 180 - aArcAngle
			For ii = 1 To 5
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle + (aArcAngle / 4)
			Next ii
			
			ARMDIA1.PR_DrawFitted(cBriefPoly)
			
		Else
			'Snap Crotch
			PR_DrawSnapCrotchArc()
			cCrotchPoly.X(1) = xyStrap1.X
			cCrotchPoly.y(1) = xyStrap1.y
			cCrotchPoly.X(2) = xyStrap2.X
			cCrotchPoly.y(2) = xyStrap2.y
			cCrotchPoly.X(3) = xyStrap4.X
			cCrotchPoly.y(3) = xyStrap4.y
			cCrotchPoly.n = 3
			ARMDIA1.PR_DrawPoly(cCrotchPoly)
			
			ARMDIA1.PR_DrawLine(xyStrap4, xyLegPoint)
			
			cBriefPoly.n = cBriefPoly.n + 1
			cBriefPoly.X(cBriefPoly.n) = xyLegPoint.X
			cBriefPoly.y(cBriefPoly.n) = xyLegPoint.y
			
			nArc = (xyFrenchCut.X - xySeamFold.X)
			If Not g_nAdult Then
				nArc = nArc * 0.5
				aAngle = 40
				aArcAngle = 80
			Else
				aAngle = 45
				aArcAngle = 90
			End If
			ARMDIA1.PR_MakeXY(xyArc, xyFrenchCut.X - nArc, xyFrenchCut.y)
			aAngle = 45
			aArcAngle = 90
			For ii = 1 To 5
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle - (aArcAngle / 4)
			Next ii
			
			nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
			If Not g_nAdult Then
				nArc = nArc * 0.25
				ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror.X + nArc, xyProfileThighMirror.y)
				nArc = ARMDIA1.FN_CalcLength(xyProfileThighMirror1, xyArc)
				aArcAngle = 40
				aAngle = ARMDIA1.FN_CalcAngle(xyProfileThighMirror1, xyArc) - aArcAngle
			Else
				ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y)
				aArcAngle = 40
				aAngle = 180 - aArcAngle
			End If
			
			ARMDIA1.PR_MakeXY(xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y)
			aArcAngle = 40
			aAngle = 180 - aArcAngle
			For ii = 1 To 5
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle + (aArcAngle / 4)
			Next ii
			
			ARMDIA1.PR_DrawFitted(cBriefPoly)
		End If
	End Sub
	
	Private Sub PR_DrawBriefStandard(ByRef nXOffset As Double)
		'UPGRADE_WARNING: Arrays in structure cBriefPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cBriefPoly As curve
		'UPGRADE_WARNING: Arrays in structure cCrotchPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cCrotchPoly As curve
		Dim nArc As Double
		Dim xyArc As XY
		Dim aAngle, aArcAngle, aConstruct As Double
		Dim ii As Short
		Dim xyTmp, xyTmp1 As XY
		Dim nRadius As Double
		'UPGRADE_WARNING: Arrays in structure cSnapCrotchPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cSnapCrotchPoly As curve
		
		'To take account of the One leg Panty and One leg Brief
		'we use the nXOffset to move the profile
		'This is a fudge but one that works.
		'Normal brief only not applicable to the Brief-French
		'NB we can reset the Global xy...  points as they are not used again.
		
		If nXOffset > 0 Then '==>Mixed Panty & Brief
			'UPGRADE_WARNING: Couldn't resolve default property of object xyProfileThigh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyProfileThigh = xyFold
			ARMDIA1.PR_MakeXY(xyProfileThighMirror, xyProfileThigh.X, -xyProfileThigh.y)
			ARMDIA1.PR_MakeXY(xyProfileThighMirror1, xyProfileThighMirror.X, xyProfileThighMirror.y + INCH1_4)
			xyProfileBrief.X = xyProfileBrief.X + nXOffset
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object xyLT_ProfileThigh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyLT_ProfileThigh = xyLT_Fold
			ARMDIA1.PR_MakeXY(xyLT_ProfileThighMirror, xyLT_ProfileThigh.X, -xyLT_ProfileThigh.y)
			'        PR_MakeXY xyLT_ProfileThighMirror1, xyLT_ProfileThighMirror.x, xyLT_ProfileThighMirror.y + inch1_4
			
			xyLegPoint.X = xyLegPoint.X + nXOffset
			xyCutOutArcCentre.X = xyCutOutArcCentre.X + nXOffset
			xySeamThigh.X = xySeamThigh.X + nXOffset
			
			xyLT_ProfileThighMirror.X = xyLT_ProfileThighMirror.X + nXOffset
			xyStrap4.X = xyStrap4.X + nXOffset
			xyGussetArcCentre.X = xyGussetArcCentre.X + nXOffset
			
		End If
		
		If g_bDiffThigh Then
			ARMDIA1.PR_DrawLine(xyLT_ProfileThigh, xyProfileBrief)
		Else
			ARMDIA1.PR_DrawLine(xyProfileThigh, xyProfileBrief)
		End If
		
		If g_sCrotchStyle <> "Snap Crotch" Then
			'All other crotch styles
			cBriefPoly.n = 0
			aConstruct = ARMDIA1.FN_CalcAngle(xyLegPoint, xyProfileBrief)
			aConstruct = aConstruct - ((aConstruct - 90) / 2)
			nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
			' PR_MakeXY xyArc, xyProfileBrief.X + nArc, xyProfileBrief.Y
			ARMDIA1.PR_CalcPolar(xyProfileBrief, aConstruct - 90, nArc, xyArc)
			aAngle = ARMDIA1.FN_CalcAngle(xyArc, xyProfileBrief)
			aArcAngle = 20
			For ii = 1 To 3
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle + (aArcAngle / 2)
			Next ii
			
			nArc = (xyLegPoint.X - xySeamThigh.X) * (3 / 2)
			'PR_MakeXY xyArc, xyLegPoint.X - nArc, xyLegPoint.Y
			ARMDIA1.PR_CalcPolar(xyLegPoint, aConstruct + 90, nArc, xyArc)
			aArcAngle = 15
			aAngle = ARMDIA1.FN_CalcAngle(xyArc, xyLegPoint) + aArcAngle
			For ii = 1 To 3
				ARMDIA1.PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
				cBriefPoly.n = cBriefPoly.n + 1
				cBriefPoly.X(cBriefPoly.n) = xyTmp.X
				cBriefPoly.y(cBriefPoly.n) = xyTmp.y
				aAngle = aAngle - (aArcAngle / 2)
			Next ii
			
			ARMDIA1.PR_DrawFitted(cBriefPoly)
			
			cBriefPoly.X(1) = xyLegPoint.X
			cBriefPoly.y(1) = xyLegPoint.y
			cBriefPoly.X(2) = xyProfileThighMirror1.X
			cBriefPoly.y(2) = xyProfileThighMirror1.y
			If g_bDiffThigh Then
				cBriefPoly.X(3) = xyLT_ProfileThighMirror.X
				cBriefPoly.y(3) = xyLT_ProfileThighMirror.y
			Else
				cBriefPoly.X(3) = xyProfileThighMirror.X
				cBriefPoly.y(3) = xyProfileThighMirror.y
			End If
			cBriefPoly.n = 3
			ARMDIA1.PR_DrawPoly(cBriefPoly)
		Else
			'Snap Crotch
			PR_DrawSnapCrotchArc()
			
			cCrotchPoly.X(1) = xyStrap1.X
			cCrotchPoly.y(1) = xyStrap1.y
			cCrotchPoly.X(2) = xyStrap2.X
			cCrotchPoly.y(2) = xyStrap2.y
			cCrotchPoly.X(3) = xyStrap4.X
			cCrotchPoly.y(3) = xyStrap4.y
			cCrotchPoly.n = 3
			ARMDIA1.PR_DrawPoly(cCrotchPoly)
			cBriefPoly.X(1) = xyStrap4.X
			cBriefPoly.y(1) = xyStrap4.y
			cBriefPoly.X(2) = xyLegPoint.X
			cBriefPoly.y(2) = xyLegPoint.y
			cBriefPoly.X(3) = xyProfileThighMirror1.X
			cBriefPoly.y(3) = xyProfileThighMirror1.y
			If g_bDiffThigh Then
				cBriefPoly.X(4) = xyLT_ProfileThighMirror.X
				cBriefPoly.y(4) = xyLT_ProfileThighMirror.y
			Else
				cBriefPoly.X(4) = xyProfileThighMirror.X
				cBriefPoly.y(4) = xyProfileThighMirror.y
			End If
			cBriefPoly.n = 4
			ARMDIA1.PR_DrawPoly(cBriefPoly)
		End If
	End Sub
	
	Private Sub PR_DrawMesh(ByRef MeshProfile As BiArc, ByRef MeshSeam As BiArc, ByRef sSide As String)
		'Subroutime to draw the given mesh
		'used to save code
		'NB the misnomer. Too complicated to change in the DrawMacro
		'procedure so do it here
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_DrawArc(MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent)
		ARMDIA1.PR_DrawArc(MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd)
		ARMDIA1.PR_SetLayer("Template" & sSide)
		ARMDIA1.PR_DrawArc(MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent)
		ARMDIA1.PR_DrawArc(MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd)
		
	End Sub
	
	Private Sub PR_DrawMeshConstruction(ByRef xyRaglanStart As XY, ByRef xyMeshStart As XY, ByRef sAxillaText As String, ByRef nMeshLength As Double)
		'Draws the custruction lines for a mesh
		'Save code
		'Returns
		'    xyMeshStart point
		'    AxillaText
		'    Length of Mesh
		Dim nOffset As Double
		Dim xyTmp, xyTmp1 As XY
		ARMDIA1.PR_SetLayer("1")
		If g_nAge < 14 Then
			nOffset = 0.75
			nMeshLength = ONE_and_THREE_QUARTER_GUSSET '2.58
			sAxillaText = "GUSSET 1-3/4\"""
		Else
			nOffset = 0.8125
			nMeshLength = BOYS_GUSSET '2.8
			sAxillaText = "BOY'S GUSSET"
		End If
		
		ARMDIA1.PR_MakeXY(xyTmp, xyRaglanStart.X - nOffset, xyRaglanStart.y)
		ARMDIA1.PR_DrawLine(xyRaglanStart, xyTmp)
		ARMDIA1.PR_MakeXY(xyTmp1, xyTmp.X, xyTmp.y + 0.25)
		ARMDIA1.PR_MakeXY(xyTmp, xyTmp1.X, xyTmp1.y - 0.5)
		ARMDIA1.PR_DrawLine(xyTmp1, xyTmp)
		
		ARMDIA1.PR_MakeXY(xyMeshStart, xyRaglanStart.X - nOffset, xyRaglanStart.y)
		
	End Sub
	
	Private Sub PR_DrawSnapCrotchArc()
		Dim aAngle As Double
		Dim xyTmp, xyTmp1 As XY
		Dim nRadius As Double
		'UPGRADE_WARNING: Arrays in structure cSnapCrotchPoly may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cSnapCrotchPoly As curve
		Dim sLayer As String
		
		'Check if the arc is too close to the cut-out
		'it should be no closer than 1.25 inches
		nRadius = ARMDIA1.FN_CalcLength(xyGussetArcCentre, xyProfileBrief)
		
		If FN_CalcCirCirInt(xyGussetArcCentre, nRadius, xyCutOut5, 1.25, xyTmp, xyTmp1) Then
			'The arc is too close
			'draw as a fitted curve to allow for user adjustment
			MsgBox("The calculated arc for the Snap Crotch was too close to the cut-out. It will be drawn as fitted curve.  Adjust this as required.")
			cSnapCrotchPoly.n = 3
			cSnapCrotchPoly.X(1) = xyProfileBrief.X
			cSnapCrotchPoly.y(1) = xyProfileBrief.y
			aAngle = ARMDIA1.FN_CalcAngle(xyGussetArcCentre, xyCutOut5)
			nRadius = ARMDIA1.FN_CalcLength(xyGussetArcCentre, xyCutOut5)
			ARMDIA1.PR_CalcPolar(xyGussetArcCentre, aAngle, nRadius + 1.25, xyTmp)
			cSnapCrotchPoly.X(2) = xyTmp.X
			cSnapCrotchPoly.y(2) = xyTmp.y
			cSnapCrotchPoly.X(cSnapCrotchPoly.n) = xyStrap3.X
			cSnapCrotchPoly.y(cSnapCrotchPoly.n) = xyStrap3.y
			ARMDIA1.PR_DrawFitted(cSnapCrotchPoly)
			
			'Draw a circle on the layer construct for use by the user
			sLayer = g_sCurrentLayer
			ARMDIA1.PR_SetLayer("Construct")
			ARMDIA1.PR_DrawCircle(xyCutOut5, 1.25)
			ARMDIA1.PR_SetLayer(sLayer)
			
		Else
			'Draw arc as normal
			ARMDIA1.PR_DrawArc(xyGussetArcCentre, xyProfileBrief, xyStrap3)
		End If
		
		'PR_DrawMarker xyGussetArcCentre
		'PR_DrawText "xyGussetArcCentre", xyGussetArcCentre, .1
		'PR_DrawMarker xyProfileBrief
		'PR_DrawText "xyProfileBrief", xyProfileBrief, .1
		'PR_DrawMarker xyStrap3
		'PR_DrawText "xyStrap3", xyStrap3, .1
		'PR_DrawMarker xyStrap1
		'PR_DrawText "xyStrap1", xyStrap1, .1
		'PR_DrawMarker xyStrap2
		'PR_DrawText "xyStrap2", xyStrap2, .1
		'PR_DrawMarker xyCutOut5
		'PR_DrawText "xyCutOut5", xyCutOut5, .1
		'PR_DrawMarker xyCutOut6
		'PR_DrawText "xyCutOut6", xyCutOut6, .1
		'PR_DrawMarker xyCutOut7
		'PR_DrawText "xyCutOut7", xyCutOut7, .1
		'PR_DrawMarker xyCutOut8
		'PR_DrawText "xyCutOut8", xyCutOut7, .1
		
	End Sub
	
	Private Sub PR_DrawVerticalLineThruPoint(ByRef xyPnt As XY, ByRef nHeight As Double)
		'Function to reduce code in PR_DrawMacro
		Static xy1, xy2 As XY
		ARMDIA1.PR_MakeXY(xy1, xyPnt.X, xyPnt.y - nHeight)
		ARMDIA1.PR_MakeXY(xy2, xyPnt.X, xyPnt.y + nHeight)
		ARMDIA1.PR_DrawLine(xy1, xy2)
		
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
	
	Private Sub PR_Select_Text(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelStart = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
	End Sub
	
	Private Sub PR_SetCrotchSizes()
		
		'Sets two global variables:
		'   g_nFrontCrotchSize (type double)
		'   g_nBackCrotchSize (type double)
		'
		'NOTE: DOES NOT CATER FOR 1 LEG PEOPLE
		
		If g_sSex = "Male" Then
			If g_nAdult Then
				If g_nButtocksCirGiven < 45 Then '45"
					g_nFrontCrotchSize = 4
					g_nBackCrotchSize = 3
				Else
					g_nFrontCrotchSize = 4.5
					g_nBackCrotchSize = 3.5
				End If
			Else
				g_nFrontCrotchSize = 2.25
				g_nBackCrotchSize = 1.5
			End If
		Else
			If g_nAdult Then
				If g_nButtocksCirGiven < 45 Then '45"
					g_nFrontCrotchSize = 3
					g_nBackCrotchSize = 4
				Else
					g_nFrontCrotchSize = 3.5
					g_nBackCrotchSize = 4.5
				End If
			Else
				g_nFrontCrotchSize = 1.5
				g_nBackCrotchSize = 2.25
			End If
		End If
		
	End Sub
	
	Private Sub PR_SetFlySize(ByRef sType As String)
		
		'Sets two global variables:
		'   g_sFlySize (type string)
		'   g_nFlyLength (type double) not at this time
		
		Dim nCutOutDia As Double
		'nCutOutDia is in inches
		nCutOutDia = ARMEDDIA1.fnRoundInches(g_nCutOutRadius * 2)
		'MsgBox "sType=" & sType & "nCutOutDia=" & Str$(nCutOutDia) & "g_nAge=" & Str$(g_nAge)
		If sType = "Hor" Then
			Select Case g_nAge
				Case 3 To 6
					If nCutOutDia <= 1.5 Then
						g_sFlySize = "C-1"
						'g_nFlyLength = "2"
					Else
						g_sFlySize = "C-2"
						'g_nFlyLength = 2.5
					End If
				Case 7 To 11
					If nCutOutDia <= 2 Then
						g_sFlySize = "C-2"
						'g_nFlyLength = 2.5
					Else
						g_sFlySize = "C-3"
						'g_nFlyLength = 2.5
					End If
				Case 12 To 14
					Select Case nCutOutDia
						Case Is <= 2.25
							g_sFlySize = "C-3"
							'g_nFlyLength = 2.5
						Case 2.375 To 2.5
							g_sFlySize = "C-4"
							'g_nFlyLength = 2.5
						Case Is > 2.5
							g_sFlySize = "C-5"
							'g_nFlyLength = 3.125
					End Select
				Case Is >= 15
					Select Case nCutOutDia
						Case Is <= 3.5
							g_sFlySize = "1"
							'g_nFlyLength = 3.5
						Case 3.625 To 4.25
							g_sFlySize = "2"
							'g_nFlyLength = 3.75
						Case 4.375 To 5
							g_sFlySize = "3"
							'g_nFlyLength = 4
						Case Is > 5.125
							g_sFlySize = "Oversize"
							'g_nFlyLength = 4
					End Select
			End Select
		Else
			Select Case g_nAge
				Case Is < 10
					g_sFlySize = "Small"
					'g_nFlyLength = 1.75
				Case 10 To 14
					g_sFlySize = "Medium"
					'g_nFlyLength = 2.75
				Case Is >= 15
					g_sFlySize = "Large"
					'g_nFlyLength = 3.75
			End Select
		End If
		
	End Sub
	
	
	Private Sub PR_SetGussetSize()
		
		'Sets two global variables:
		'   g_sGussetSize (type string)
		'   g_nGussetLength (type double) not at this time
		
		Dim nCutOutDia As Double
		
		'nCutOutDia is in inches
		nCutOutDia = ARMEDDIA1.fnRoundInches(g_nCutOutRadius * 2)
		If g_sSex = "Male" Then
			Select Case g_nAge
				Case Is <= 3
					g_sGussetSize = "1-3/4\"""
					g_nGussetLength = 2.125
				Case 4 To 14
					g_sGussetSize = "Boy's"
					g_nGussetLength = 3.875
				Case Is >= 15
					g_sGussetSize = "Male"
					g_nGussetLength = 5.375
			End Select
		Else
			Select Case g_nAge
				Case Is <= 3
					g_sGussetSize = "1-3/4\"""
					g_nGussetLength = 2.125
				Case 4 To 14
					Select Case nCutOutDia
						Case Is <= 1.75
							g_sGussetSize = "1\"""
							g_nGussetLength = 2.125
						Case 1.875 To 3.125
							g_sGussetSize = "1-1/4\"""
							g_nGussetLength = 2.75
						Case Is >= 3.25
							g_sGussetSize = "Regular"
							g_nGussetLength = 3.625
					End Select
				Case Is >= 15
					If g_nButtocksCirGiven <= 45 Then
						g_sGussetSize = "Regular"
						g_nGussetLength = 3.625
					Else
						g_sGussetSize = "Oversize"
						g_nGussetLength = 4.125
					End If
			End Select
		End If
		
	End Sub
	
	Private Sub PR_UpdateDB()
		'Procedure called from PR_CreateDrawMacro
		
		Dim sSymbol, sMeshData As String
		
		sSymbol = "suitbody"
		
		'Use existing symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("hBody = UID (" & QQ & "find" & QC & Val(txtUidVB.Text) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
		
		'Update the BODY Box symbol
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "AxillaBackNeckRad" & QCQ & g_nAxillaBackNeckRad & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "AxillaFrontNeckRad" & QCQ & g_nAxillaFrontNeckRad & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "ABNRadRight" & QCQ & g_nABNRadRight & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "AFNRadRight" & QCQ & g_nAFNRadRight & QQ & ");")
		
		'Mesh data for use by sleeve drawing programme
		'We use commas as the delimiters to show empty data
		If g_sMeshLeft = "" And g_sMeshRight = "" Then
			sMeshData = ",,,"
			
		ElseIf g_sMeshLeft = "" Then 
			sMeshData = ",," & g_sMeshRight
			
		ElseIf g_sMeshRight = "" Then 
			sMeshData = g_sMeshLeft & ",,"
			
		Else
			sMeshData = g_sMeshLeft & "," & g_sMeshRight
			
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Data" & QCQ & sMeshData & QQ & ");")
		
	End Sub
	
	Private Sub prOptLeftLeg(ByRef i0 As Short, ByRef i1 As Short)
		'Procedure to set the option buttons
		'sets them according to the pattern i0-i2
		
		optLeftLeg(0).Checked = i0
		optLeftLeg(1).Checked = i1
		
	End Sub
	
	Private Sub prOptRightLeg(ByRef i0 As Short, ByRef i1 As Short)
		'Procedure to set the option buttons
		'sets them according to the pattern i0-i2
		
		optRightLeg(0).Checked = i0
		optRightLeg(1).Checked = i1
		
	End Sub
	
	Private Sub PrOutPut()
		
		'    Dim fOutPut As Integer
		'
		'    'Open file
		'    fOutPut = FreeFile
		'    Open "C:\JOBST\VARS.TXT" For Output As fOutPut
		'
		'    Print #fOutPut, "BodySuit variables"
		'    Print #fOutPut,
		'    Print #fOutPut, "g_sCurrentLayer= ", g_sCurrentLayer
		'    Print #fOutPut, "g_nCurrTextHt= ", g_nCurrTextHt
		'    Print #fOutPut, "g_nCurrTextAspect= ", g_nCurrTextAspect
		'    Print #fOutPut, "g_nCurrTextHorizJust= ", g_nCurrTextHorizJust
		'    Print #fOutPut, "g_nCurrTextVertJust= ", g_nCurrTextVertJust
		'    Print #fOutPut, "g_nCurrTextFont= ", g_nCurrTextFont
		'    Print #fOutPut, "g_nCurrTextAngle= ", g_nCurrTextAngle
		'
		'    Print #fOutPut,
		'    Print #fOutPut, "Patient details"
		'    Print #fOutPut,
		'    Print #fOutPut, "g_sFileNo= ", g_sFileNo
		'    Print #fOutPut, "g_sSide= ", g_sSide
		'    Print #fOutPut, "g_sPatient= ", g_sPatient
		'    Print #fOutPut, "g_sID= ", g_sID
		'    Print #fOutPut, "g_sDiagnosis= ", g_sDiagnosis
		'    Print #fOutPut, "g_nAge= ", g_nAge
		'    Print #fOutPut, "g_sSex= ", g_sSex
		'    Print #fOutPut, "g_sWorkOrder= ", g_sWorkOrder
		'    Print #fOutPut, "g_nUnitsFac= ", g_nUnitsFac
		'    Print #fOutPut, "g_sBackNeck= ", g_sBackNeck
		'    Print #fOutPut, "g_sFrontNeck= ", g_sFrontNeck
		'    Print #fOutPut, "g_nUnderBreast= ", g_nUnderBreast
		'    Print #fOutPut, "g_nNipple= ", g_nNipple
		'    Print #fOutPut, "g_nChest= ", g_nChest
		'    Print #fOutPut, "g_sLeftCup= ", g_sLeftCup
		'    Print #fOutPut, "g_sRightCup= ", g_sRightCup
		'    Print #fOutPut, "g_sLeftDisk= ", g_sLeftDisk
		'    Print #fOutPut, "g_sRightDisk= ", g_sRightDisk
		'
		'    Print #fOutPut,
		'    Print #fOutPut, "Response & DiffSides"
		'    Print #fOutPut,
		'   Print #fOutPut, "g_bResponse= ", g_bResponse
		'    Print #fOutPut, "g_bDiffThigh= ", g_bDiffAxillaType
		'
		'    Print #fOutPut,
		'   Print #fOutPut, "Values Given"
		'    Print #fOutPut,
		'    Print #fOutPut, "g_nLeftShldCirGiven= ", g_nLeftShldCirGiven
		'    Print #fOutPut, "g_nRightShldCirGiven= ", g_nRightShldCirGiven
		'    Print #fOutPut, "g_nLeftAxillaCirGiven= ", g_nLeftAxillaCirGiven
		'    Print #fOutPut, "g_nRightAxillaCirGiven= ", g_nRightAxillaCirGiven
		'    Print #fOutPut, "g_nChestCirGiven= ", g_nChestCirGiven
		'    Print #fOutPut, "g_nUnderBreastCirGiven= ", g_nUnderBreastCirGiven
		'    Print #fOutPut, "g_nWaistCirGiven= ", g_nWaistCirGiven
		'    Print #fOutPut, "g_nButtocksCirGiven= ", g_nButtocksCirGiven
		'    Print #fOutPut, "g_nLeftThighCirGiven= ", g_nLeftThighCirGiven
		'    Print #fOutPut, "g_nRightThighCirGiven= ", g_nRightThighCirGiven
		'    Print #fOutPut, "g_nNeckCirGiven= ", g_nNeckCirGiven
		'    Print #fOutPut, "g_nShldToFoldGiven= ", g_nShldToFoldGiven
		'    Print #fOutPut, "g_nShldToWaistGiven= ", g_nShldToWaistGiven
		'    Print #fOutPut, "g_nShldToBreastGiven= ", g_nShldToBreastGiven
		'    Print #fOutPut, "g_nLeftShldToAxilla= ", g_nLeftShldToAxilla
		'    Print #fOutPut, "g_nRightShldToAxilla= ", g_nRightShldToAxilla
		'    Print #fOutPut, "g_nNippleCirGiven= ", g_nNippleCirGiven
		'    Print #fOutPut, "g_sFrontNeckStyle= ", g_sFrontNeckStyle
		'    Print #fOutPut, "g_nFrontNeckSize= ", g_nFrontNeckSize
		'    Print #fOutPut, "g_sBackNeckStyle= ", g_sBackNeckStyle
		'    Print #fOutPut, "g_nBackNeckSize= ", g_nBackNeckSize
		'    Print #fOutPut, "g_sLeftSleeve= ", g_sLeftSleeve
		'    Print #fOutPut, "g_sRightSleeve= ", g_sRightSleeve
		'    Print #fOutPut, "g_sLeftLeg= ", g_sLeftLeg
		'    Print #fOutPut, "g_sRightLeg= ", g_sRightLeg
		'    Print #fOutPut, "g_sClosure= ", g_sClosure
		'    Print #fOutPut, "g_sFabric= ", g_sFabric
		'    Print #fOutPut, "g_sCrotchStyle= ", g_sCrotchStyle
		'
		'    Print #fOutPut,
		'    Print #fOutPut, "Circumferences"
		'    Print #fOutPut,
		'    Print #fOutPut, "g_nShldWidth= ", g_nShldWidth
		'    Print #fOutPut, "g_nLeftThighCir= ", g_nLeftThighCir
		'    Print #fOutPut, "g_nRightThighCir= ", g_nRightThighCir
		'    Print #fOutPut, "g_nShldToAxilla= ", g_nShldToAxilla
		'    Print #fOutPut, "g_nAxillaCir= ", g_nAxillaCir
		'    Print #fOutPut, "g_nThighCir= ", g_nThighCir
		'    Print #fOutPut, "g_nLeftAxillaCir= ", g_nLeftAxillaCir
		'    Print #fOutPut, "g_nRightAxillaCir= ", g_nRightAxillaCir
		'    Print #fOutPut, "g_nChestCir= ", g_nChestCir
		'    Print #fOutPut, "g_nUnderBreastCir= ", g_nUnderBreastCir
		'    Print #fOutPut, "g_nWaistCir= ", g_nWaistCir
		'    Print #fOutPut, "g_nButtocksCir= ", g_nButtocksCir
		'    Print #fOutPut, "g_nNeckCir= ", g_nNeckCir
		'    Print #fOutPut, "g_nShldToFold= ", g_nShldToFold
		'    Print #fOutPut, "g_nShldToWaist= ", g_nShldToWaist
		'    Print #fOutPut, "g_nShldToBreast= ", g_nShldToBreast
		'    Print #fOutPut, "g_nButtocksLength= ", g_nButtocksLength
		'    Print #fOutPut, "g_nButtBackSeamRatio= ", g_nButtBackSeamRatio
		'    Print #fOutPut, "g_nButtFrontSeamRatio= ", g_nButtFrontSeamRatio
		'    Print #fOutPut, "g_nButtRedIncreased= ", g_nButtRedIncreased
		'    Print #fOutPut, "g_nThighRedDecreased= ", g_nThighRedDecreased
		'    Print #fOutPut, "g_nButtCir= ", g_nButtCir
		'    Print #fOutPut, "g_nHalfGroinHeight= ", g_nHalfGroinHeight
		'    Print #fOutPut, "g_nButtRadius= ", g_nButtRadius
		'    Print #fOutPut, "g_nButtCirCalc= ", g_nButtCirCalc
		'    Print #fOutPut, "g_nButtFrontSeam= ", g_nButtFrontSeam
		'    Print #fOutPut, "g_nButtBackSeam= ", g_nButtBackSeam
		'    Print #fOutPut, "g_nButtCutOut= ", g_nButtCutOut
		'    Print #fOutPut, "g_nWaistFrontSeam= ", g_nWaistFrontSeam
		'    Print #fOutPut, "g_nWaistBackSeam= ", g_nWaistBackSeam
		'    Print #fOutPut, "g_nWaistCutOut= ", g_nWaistCutOut
		'    Print #fOutPut, "g_nChestFrontSeam= ", g_nChestFrontSeam
		'    Print #fOutPut, "g_nChestBackSeam= ", g_nChestBackSeam
		'    Print #fOutPut, "g_nChestCutOut= ", g_nChestCutOut
		'
		'    Print #fOutPut,
		'    Print #fOutPut, "Reductions"
		'    Print #fOutPut,
		'    Print #fOutPut, "g_nChestCirRed= ", g_nChestCirRed
		'    Print #fOutPut, "g_nUnderBreastCirRed= ", g_nUnderBreastCirRed
		'    Print #fOutPut, "g_nWaistCirRed= ", g_nWaistCirRed
		'    Print #fOutPut, "g_nLargeButtCirRed= ", g_nLargeButtCirRed
		'    Print #fOutPut, "g_nThighCirRed= ", g_nThighCirRed
		'    Print #fOutPut, "g_nNeckCirRed= ", g_nNeckCirRed
		'    Print #fOutPut, "g_nShldToFoldRed= ", g_nShldToFoldRed
		'    Print #fOutPut, "g_nShldToWaistRed= ", g_nShldToWaistRed
		'    Print #fOutPut, "g_nShldToBreastRed= ", g_nShldToBreastRed
		'    Print #fOutPut, "g_nButtocksCirRed= ", g_nButtocksCirRed
		'
		'    Print #fOutPut,
		'    Print #fOutPut, "Co-ordinate points"
		'    Print #fOutPut,
		'    Print #fOutPut, "xySeamFold= ", xySeamFold.X, xySeamFold.Y
		'    Print #fOutPut, "xySeamThigh= ", xySeamThigh.X, xySeamThigh.Y
		'    Print #fOutPut, "xySeamButt= ", xySeamButt.X, xySeamButt.Y
		'    Print #fOutPut, "xySeamHighShld= ", xySeamHighShld.X, xySeamHighShld.Y
		'    Print #fOutPut, "xySeamLowShld= ", xySeamLowShld.X, xySeamLowShld.Y
		'    Print #fOutPut, "xySeamChest= ", xySeamChest.X, xySeamChest.Y
		'    Print #fOutPut, "xySeamWaist= ", xySeamWaist.X, xySeamWaist.Y
		'    Print #fOutPut, "xyProfileThigh= ", xyProfileThigh.X, xyProfileThigh.Y
		'    Print #fOutPut, "xyFold= ", xyFold.X, xyFold.Y
		'    Print #fOutPut, "xyProfileButt= ", xyProfileButt.X, xyProfileButt.Y
		'    Print #fOutPut, "xyCutOut3= ", xyCutout3.X, xyCutout3.Y
		'    Print #fOutPut, "xyCutOut4= ", xyCutOut4.X, xyCutOut4.Y
		'    Print #fOutPut, "xyCutOut5= ", xyCutOut5.X, xyCutOut5.Y
		'    Print #fOutPut, "xyCutOut6= ", xyCutOut6.X, xyCutOut6.Y
		'    Print #fOutPut, "xyCutOut2= ", xyCutOut2.X, xyCutOut2.Y
		'    Print #fOutPut, "xyCutOut7= ", xyCutOut7.X, xyCutOut7.Y
		'    Print #fOutPut, "xyProfileWaist= ", xyProfileWaist.X, xyProfileWaist.Y
		'    Print #fOutPut, "xyProfileChest= ", xyProfileChest.X, xyProfileChest.Y
		'    Print #fOutPut, "xyProfileNeck= ", xyProfileNeck.X, xyProfileNeck.Y
		'    Print #fOutPut, "xyProfileNeckMirror= ", xyProfileNeckMirror.X, xyProfileNeckMirror.Y
		'    Print #fOutPut, "xyCutOutBackNeck= ", xyCutOutBackNeck.X, xyCutOutBackNeck.Y
		'    Print #fOutPut, "xyCutOut8= ", xyCutOut8.X, xyCutOut8.Y
		'   Print #fOutPut, "xyCutOutFrontNeck= ", xyCutOutFrontNeck.X, xyCutOutFrontNeck.Y
		'    Print #fOutPut, "xyRaglan1= ", xyRaglan1.X, xyRaglan1.Y
		'    Print #fOutPut, "xyRaglan2= ", xyRaglan2.X, xyRaglan2.Y
		'    Print #fOutPut, "xyRaglan3= ", xyRaglan3.X, xyRaglan3.Y
		'    Print #fOutPut, "xyRaglan4= ", xyRaglan4.X, xyRaglan4.Y
		'    Print #fOutPut, "xyRaglan5= ", xyRaglan5.X, xyRaglan5.Y
		'    Print #fOutPut, "xyRaglan6= ", xyRaglan6.X, xyRaglan6.Y
		
		'    Close #fOutPut
		
	End Sub
	
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
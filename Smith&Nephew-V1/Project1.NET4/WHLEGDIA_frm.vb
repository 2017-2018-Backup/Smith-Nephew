Option Strict Off
Option Explicit On
Friend Class whlegdia
	Inherits System.Windows.Forms.Form
	'Project:   Waist height
	'File:      whlegdia.frm
	'Purpose:   Waist height dialogue
	'
	'
	'Version:   3.00
	'Date:      N/A
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	'   '* Windows API Functions Declarations
	'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	
	'   'Constanst used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	'Constants
	Const HEEL_TOL As Short = 9 '9" heel
    Public g_sChangeChecker As String
    Public g_nUnitsFac As Double

    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote

    'Globals

    Public g_nHeelLen As Single
    Public g_nAnkleLen As Single
    Public g_nAnkleTape As Single
    Public g_sPathJOBST As String


    'MsgBox constants
    Const IDYES = 6
    Const IDNO = 7
    Const IDCANCEL = 2




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
                        Return
                    Else
                        MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Leg Data")
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

    Private Sub cmdExtendLegTapes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExtendLegTapes.Click
		'Takes the last tape given and add the required value
		'GOP - 01-02/12, 6.2.1
		'
		Dim ii As Short
		Dim nValue As Double
		
		'Locate last tape
		For ii = 29 To 0 Step -1
			If Val(Text1(ii).Text) > 0 Then Exit For
		Next ii
		
		'Check that there are some tapes.
		'Check that there is room to add a new tape
		If ii = 0 Or ii = 29 Then
			Beep()
			Exit Sub
		End If
		
		'Convert given value to inches
		nValue = ARMEDDIA1.fnDisplayToInches(Val(Text1(ii).Text))
		
		'Extend tape
		If txtSex.Text = "Male" Then
			nValue = nValue + 0.5
		Else
			nValue = nValue + 1
		End If
		
		Text1(ii + 1).Text = CStr(BDLEGDIA1.fnInchesToDisplay(nValue))
		labTapeInInches(ii + 1).Text = ARMEDDIA1.fnInchesToText(nValue)
		
	End Sub
	
	Private Sub cmdShiftTapesDown_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShiftTapesDown.Click
		'The effect of a single click on this button
		'is to shift all of the tapes in the direction of the
		'arrow
		Dim ii As Short
		Dim nValue As Double
		
		'Check that last tape is empty so that the previous
		'tape can be shifted into it
		If Len(Text1(29).Text) > 0 Then
			'Beep and exit function
			Beep()
			Exit Sub
		End If
		
		'Shift all tapes down 1
		For ii = 29 To 1 Step -1
			Text1(ii).Text = Text1(ii - 1).Text
			nValue = ARMEDDIA1.fnDisplayToInches(Val(Text1(ii - 1).Text))
			labTapeInInches(ii).Text = ARMEDDIA1.fnInchesToText(nValue)
		Next ii
		
		'Clean first position
		Text1(0).Text = ""
		labTapeInInches(0).Text = ""
		
		'Check for a heel value
		If Val(Text1(5).Text) > 0 Then
			optBelowKnee.Checked = False
			optAboveKnee.Checked = False
			optBelowKnee.Enabled = False
			optAboveKnee.Enabled = False
		Else
			optBelowKnee.Enabled = True
			optAboveKnee.Enabled = True
		End If
	End Sub
	
	Private Sub cmdShiftTapesUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShiftTapesUp.Click
		'The effect of a single click on this button
		'is to shift all of the tapes in the direction of the
		'arrow
		Dim ii As Short
		Dim nValue As Double
		
		'Check that first tape is empty so that the following
		'tape can be shifted into it
		If Len(Text1(0).Text) > 0 Then
			'Beep and exit function
			Beep()
			Exit Sub
		End If
		
		'Shift all tapes up 1
		For ii = 1 To 29
			Text1(ii - 1).Text = Text1(ii).Text
			nValue = ARMEDDIA1.fnDisplayToInches(Val(Text1(ii).Text))
			labTapeInInches(ii - 1).Text = ARMEDDIA1.fnInchesToText(nValue)
		Next ii
		'Clean last position
		Text1(29).Text = ""
		labTapeInInches(29).Text = ""
		
		'Check for a heel value
		If Val(Text1(5).Text) > 0 Then
			optBelowKnee.Checked = False
			optAboveKnee.Checked = False
			optBelowKnee.Enabled = False
			optAboveKnee.Enabled = False
		Else
			optBelowKnee.Enabled = True
			optAboveKnee.Enabled = True
		End If
		
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
        PrintLine(Project1.ARMDIA1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "//Patient - " & sName & ", " & sFileNo & "")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "//by Visual Basic, Waist Height Leg")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "//type - " & sType & "")

        'Clear user selections etc
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")

        'Display Hour Glass symbol
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & sType & QQ & ");")


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
		For ii = 0 To 29
			sString = sString & Str(Val(Text1(ii).Text))
		Next ii
		
		sString = sString & txtFootPleat1.Text & txtFootPleat2.Text
		
		sString = sString & txtTopLegPleat1.Text & txtTopLegPleat2.Text & txtFootLength.Text
		
		sString = sString & cboToeStyle.Text & Str(optAboveKnee.Checked) & Str(optBelowKnee.Checked)
		
		FN_ValuesString = sString
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Dim nValue As Double
		Dim nAge As Object
		Dim ii As Short
		
		'Disable Time out
		Timer1.Enabled = False
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Foot Length
		nValue = ARMEDDIA1.fnDisplayToInches(Val(txtFootLength.Text))
		If nValue > 0 Then labFootLengthInInches.Text = ARMEDDIA1.fnInchesToText(nValue)
		
		'Set dropdown combo boxes
		'Toe Style
		cboToeStyle.Items.Add("")
		cboToeStyle.Items.Add("Curved")
		cboToeStyle.Items.Add("Cut-Back")
		cboToeStyle.Items.Add("Straight")
		cboToeStyle.Items.Add("Soft Enclosed")
		cboToeStyle.Items.Add("Soft Enclosed B/M")
		cboToeStyle.Items.Add("Self Enclosed")
		'Set value
		For ii = 0 To (cboToeStyle.Items.Count - 1)
			If VB6.GetItemString(cboToeStyle, ii) = txtToeStyle.Text Then
				cboToeStyle.SelectedIndex = ii
			End If
		Next ii
		
		
		'Leg Option buttons
		If txtLeg.Text = "Left" Then
			optLeftLeg.Checked = True
			optRightLeg.Enabled = False
		Else
			optRightLeg.Checked = True
			optLeftLeg.Enabled = False
		End If
		
		'Update tape boxes
		For ii = 0 To 29
			nValue = Val(Mid(txtTapeLengths.Text, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				Text1(ii).Text = CStr(nValue)
				nValue = ARMEDDIA1.fnDisplayToInches(nValue)
				labTapeInInches(ii).Text = ARMEDDIA1.fnInchesToText(nValue)
			End If
		Next ii
		
		'Store original Heel and Ankle tape values
		'Used on DDE_Update check if values have been modified
		'A heel length of 0 => footless
		
		g_nAnkleTape = 0
		g_nAnkleLen = 0
		g_nHeelLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(5).Text))
		
		If (g_nHeelLen <> 0) And (g_nHeelLen < HEEL_TOL) Then
			g_nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
			g_nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(g_nAnkleTape).Text))
		Else
			g_nAnkleTape = 7 'Large heel Ankle at +3 tape
			g_nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(g_nAnkleTape).Text))
		End If
		
		'Set below and above knee option buttons
		'Disable both if a heel value is given
		optBelowKnee.Checked = False
		optAboveKnee.Checked = False
		optBelowKnee.Enabled = False
		optAboveKnee.Enabled = False
		
		If g_nHeelLen = 0 Then
			optBelowKnee.Enabled = True
			optAboveKnee.Enabled = True
			Select Case Val(txtData.Text)
				Case -1
					optBelowKnee.Checked = True
				Case 1
					optAboveKnee.Checked = True
			End Select
		End If
		
		'Store initial values for use in Cancel_Click
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
            Return
        End If
	End Sub
	
	Private Sub whlegdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Hide()
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
		
		g_sPathJOBST = fnPathJOBST()
		
		g_nUnitsFac = 1 'Default to inches
		
		'Start a timer that will ensure that the
		'the VB programme ends if the DRAFIX DDE
		'link fails.
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		'Check that data is all present and insert into drafix
		Dim sTask As String
		
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
                Return
            Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Leg Data")
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hLeg;")
		
		Update_DDE_Text_Boxes()
		
		Dim sSymbol As String
		
		sSymbol = "waistleg"
		
		If txtUidLeg.Text = "" Then
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
			
			'Insert symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			If txtLeg.Text = "Left" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				BDYUTILS.PR_PutLine("  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 1.5,xyMPD_Origin.y );")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				BDYUTILS.PR_PutLine("  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 3,xyMPD_Origin.y );")
			End If
			BDYUTILS.PR_PutLine("  }")
			BDYUTILS.PR_PutLine("else")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
		Else
			'Use existing symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hLeg = UID (" & QQ & "find" & QC & Val(txtUidLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			
		End If
		
		'Update the BODY Box symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "Leg" & QCQ & txtLeg.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid(txtTapeLengths.Text, 1, 60) & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid(txtTapeLengths.Text, 61) & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "AnkleTape" & QCQ & txtAnkleTape.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "FirstTape" & QCQ & txtFirstTape.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "LastTape" & QCQ & txtLastTape.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "ToeStyle" & QCQ & txtToeStyle.Text & QQ & ");")
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "FootPleat1" & QCQ & txtFootPleat1.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "FootPleat2" & QCQ & txtFootPleat2.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "TopLegPleat1" & QCQ & txtTopLegPleat1.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "TopLegPleat2" & QCQ & txtTopLegPleat2.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "FootLength" & QCQ & txtFootLength.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("SetDBData( hLeg" & CQ & "Data" & QCQ & txtData.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(Project1.ARMDIA1.fNum, sLine)

    End Sub
	
	Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
		'Allows the user to use enter as a tab
		
		System.Windows.Forms.SendKeys.Send("{TAB}")
		
	End Sub
	
	Private Sub Text1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Enter
		Dim Index As Short = Text1.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(Text1(Index))
	End Sub
	
	Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
		Dim Index As Short = Text1.GetIndex(eventSender)
		Dim nValue As Double
		nValue = BDLEGDIA1.Validate_And_Display_Text_In_Box(Text1(Index), labTapeInInches(Index))
		
		'If this is for the heel tape ie Index = 5 then
		'Enable or disable the Above or Below option controls
		If nValue >= 0 And Index = 5 Then
			If nValue = 0 Then
				'No heel given Enable
				optBelowKnee.Enabled = True
				optAboveKnee.Enabled = True
			Else
				'Heel value given Disable controls
				optBelowKnee.Checked = False
				optAboveKnee.Checked = False
				optBelowKnee.Enabled = False
				optAboveKnee.Enabled = False
			End If
		End If
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'If this event is called then the
        'DDE link with DRAFIX has not closed
        'Therefor we time out after 6 seconds
        Return

    End Sub
	
	Private Sub txtFootLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFootLength)
	End Sub
	
	Private Sub txtFootLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Leave
		Dim nValue As Double
		nValue = BDLEGDIA1.Validate_And_Display_Text_In_Box(txtFootLength, labFootLengthInInches)
	End Sub
	
	Private Sub txtFootPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat1.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFootPleat1)
	End Sub
	
	Private Sub txtFootPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat1.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtFootPleat1)
	End Sub
	
	Private Sub txtFootPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat2.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFootPleat2)
	End Sub
	
	Private Sub txtFootPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat2.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtFootPleat2)
	End Sub
	
	Private Sub txtTopLegPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat1.Enter
		ARMEDDIA1.Select_Text_In_Box(txtTopLegPleat1)
	End Sub
	
	Private Sub txtTopLegPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat1.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtTopLegPleat1)
	End Sub
	
	Private Sub txtTopLegPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat2.Enter
		ARMEDDIA1.Select_Text_In_Box(txtTopLegPleat2)
	End Sub
	
	Private Sub txtTopLegPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat2.Leave
		BDLEGDIA1.Validate_Text_In_Box(txtTopLegPleat2)
	End Sub
	
	Private Sub Update_DDE_Text_Boxes()
		'Called from OK_Click and Cancel_Click
		'Update the text boxes used for DDE transfer
		Dim ii As Short
		Dim sString, sRightJustified As String
		Dim sJustifiedString As String
		Dim nValue As Single
		Dim nAnkleLen, nAnkleTape, nHeelLen As Single
		
		'Combo Boxes
		txtToeStyle.Text = VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex)
		
		'Option Boxes
		If optLeftLeg.Checked = True Then
			txtLeg.Text = "Left"
		Else
			txtLeg.Text = "Right"
		End If
		
		txtData.Text = "0"
		If optBelowKnee.Checked = True Then txtData.Text = "-1"
		If optAboveKnee.Checked = True Then txtData.Text = "1"
		
		'Tape boxes
		'Assume that data has been validated earlier
		
		'Set initial values
		txtTapeLengths.Text = ""
		txtFirstTape.Text = CStr(-1)
		txtLastTape.Text = CStr(30) 'Take care of special case where last tape is text1(29)
		
		For ii = 0 To 29
			nValue = Val(Text1(ii).Text)
			If nValue <> 0 Then
				nValue = nValue * 10 'Shift decimal place to right
				sJustifiedString = New String(" ", 4)
				sJustifiedString = RSet(Trim(Str(nValue)), Len(sJustifiedString))
			Else
				sJustifiedString = New String(" ", 4)
			End If
			
			'Tape values
			txtTapeLengths.Text = txtTapeLengths.Text & sJustifiedString
			
			'Set first and last tape (assumes no holes in data)
			If CDbl(txtFirstTape.Text) < 0 And nValue > 0 Then txtFirstTape.Text = CStr(ii + 1)
			If CDbl(txtLastTape.Text) = 30 And CDbl(txtFirstTape.Text) > 0 And nValue = 0 Then txtLastTape.Text = CStr(ii)
		Next ii
		
		
		'Check if heel or ankle changed, If they have make the changes and warn
		'that configuring is required
		nAnkleTape = 0
		nAnkleLen = 0
		nHeelLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(5).Text))
		
		If (nHeelLen <> 0) And (nHeelLen < HEEL_TOL) Then
			nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
			nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(nAnkleTape).Text))
		Else
			nAnkleTape = 7 'Large heel Ankle at +3 tape
			nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(nAnkleTape).Text))
		End If
		
		'Note:-
		'        The DRAFIX macros index tape lengths from 1 not 0 as VB does.
		'        Therefor nAnkleTape is incremented by 1
		'
		If (nHeelLen <> g_nHeelLen) Or (nAnkleLen <> g_nAnkleLen) Then
			MsgBox("Figure Ankle for this leg.", 48)
			If nHeelLen = 0 Then txtAnkleTape.Text = "-1 0 0 0 0 0 0 0"
		End If
		
		'Set the txtAnkleTape to reflect the footless state
		'Only if it has not been set before
		If nHeelLen = 0 And txtAnkleTape.Text = "" Then txtAnkleTape.Text = "-1 0 0 0 0 0 0 0"
		
		
	End Sub
	
	Private Function Validate_Data() As Short
		'Called from OK_Click
		
		Dim NL, sError, sTextList As String
		Dim nFirstTape, ii, nLastTape As Short
		Dim iError As Short
		Dim nFootLength, nHeelLength, nTapeX As Double
		Dim nValue As Double
		
		NL = Chr(10) 'new line
		sTextList = "  7  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
		
		'Check Tape length data text boxes for holes (ie missing values)
		'Establish First and last tape
		For ii = 0 To 29 Step 1
			If Val(Text1(ii).Text) > 0 Then Exit For
		Next ii
		nFirstTape = ii
		
		For ii = 29 To 0 Step -1
			If Val(Text1(ii).Text) > 0 Then Exit For
		Next ii
		nLastTape = ii
		
		If nFirstTape = 30 Then
			sError = sError & "No Tape lengths given." & NL
		Else
			For ii = nFirstTape To nLastTape
				If Val(Text1(ii).Text) = 0 Then
					sError = sError & "Missing Tape length " & LTrim(Mid(sTextList, (ii * 3) + 1, 3)) & NL
				End If
			Next ii
		End If
		
		'Check combo boxes, not re-inforcement
		'NB - the first item in each combo box is a blank
		'Toe style need only be given if a heel tape is given, Heel is at tape 5
		If (nFirstTape < 5) And (cboToeStyle.SelectedIndex <= 0) Then
			sError = sError & "Missing Toe Style." & NL
		End If
		
		'Check Tape length data text boxes for holes (ie missing values)
		'Establish First and last tape
		For ii = 0 To 29 Step 1
			If Val(Text1(ii).Text) > 0 Then Exit For
		Next ii
		nFirstTape = ii
		
		For ii = 29 To 0 Step -1
			If Val(Text1(ii).Text) > 0 Then Exit For
		Next ii
		nLastTape = ii
		
		If nFirstTape = 29 Then
			sError = sError & "No Tape lengths given." & NL
		End If
		
		'Check that a minimum of 3 tapes are given
		If (nLastTape - nFirstTape) = 0 Then
			sError = sError & "More than 1 tape must be given." & NL
		End If
		
		'Check that a foot length has been given for self enclosed
		If VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Self Enclosed" And Val(txtFootLength.Text) = 0 Then
			sError = sError & "A foot length must be given for Self Enclosed toes." & NL
		End If
		
		'Get Heel Length if it applies
		If nFirstTape <= 5 Then
			nHeelLength = ARMEDDIA1.fnDisplayToInches(Val(Text1(5).Text))
		Else
			nHeelLength = 0
		End If
		
		'Check if either the "above" or "below" knee options are
		'given for a footless
		If nHeelLength = 0 And optBelowKnee.Checked = False And optAboveKnee.Checked = False Then
			sError = sError & "This is a footless style." & NL
			sError = sError & "ABOVE or BELOW Knee must be selected." & NL
		End If
		
		'Check if -3 tape exists for Large Heel
		If nFirstTape <= 5 And nHeelLength >= 9 And Val(Text1(3).Text) = 0 Then
			sError = sError & "As the Heel is 9 inches and over." & NL
			sError = sError & "and there is no -3 tape the foot will not draw properly " & NL
		End If
		
		'Check if -1 1/2 tape exists for Small Heel
		If nFirstTape <= 5 And nHeelLength < 9 And Val(Text1(4).Text) = 0 Then
			sError = sError & "As the Heel is smaller than 9 inches." & NL
			sError = sError & "and there is no -1 1/2 tape the foot will not draw properly " & NL
		End If
		
		
		'Display Error message (if required) and return
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Fatal - Errors in Data")
			Validate_Data = False
			Exit Function
		Else
			Validate_Data = True
		End If
		
		'For the combo boxes check if a toe selection has been made
		'and a heel tape is not given.
		'This is a none fatal error, the data will be ammended and
		'an information message given
		'will be given
		
		sError = ""
		If nFirstTape > 5 Then
			If cboToeStyle.SelectedIndex > 0 Then
				sError = sError & "Toe style, " & VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) & NL
			End If
			cboToeStyle.SelectedIndex = 0
		End If
		
		'Display Warning message (if required)
		If Len(sError) > 0 Then
			sError = "The TOE style given can't apply. As there is no Heel." & NL
			sError = sError & "Removing TOE style to reflect this."
			MsgBox(sError, 64, "Warning - Errors in Data")
		End If
		
		
		'Yes / No type errors
		'In this case we warn the user that there is a problem!
		'They can continue or they can return to the dialog to make changes
		
		'Initialize error variables
		sError = ""
		iError = False
		
		'Make a stab at seeing if the foot length will be long enough
		'For Toe = "Self enclosed" or Toe = "Soft enclosed" and a foot length is given
		If nHeelLength > 0 And (VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Self Enclosed" Or (VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Soft Enclosed" And Val(txtFootLength.Text) <> 0)) Then
			
			'Revise heel length to worst case
			'Scale heel to the 0 tape of a 30mmHg template
			nHeelLength = nHeelLength * 0.402
			'Reduce heel by 3
			nHeelLength = nHeelLength + (3 * 0.05025)
			
			'
			If nHeelLength >= 9 Then
				nTapeX = 2.71875 'Large heel, distance from heel to -3 tape
			Else
				nTapeX = 1.5 'Small heel, distance from heel to -1 1/2 tape
			End If
			
			nFootLength = ARMEDDIA1.fnDisplayToInches(Val(txtFootLength.Text))
			
			nValue = System.Math.Sqrt(nHeelLength ^ 2 + nTapeX ^ 2)
			
			If nValue > nFootLength Then
				iError = True
				sError = sError & "The given foot length may be too small to correctly" & NL
				sError = sError & "position the toe!" & NL & NL
				sError = sError & "Do you wish to continue anyway ?"
			End If
		End If
		
		'Display Error message (if required) and return
		'These are non-fatal errors
		If iError = True Then
			iError = MsgBox(sError, 36, "Warning, Problems with data")
			If iError = IDYES Then
				Validate_Data = True
			Else
				Validate_Data = False
			End If
		Else
			Validate_Data = True
		End If
		
	End Function
End Class
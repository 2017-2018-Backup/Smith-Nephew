Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Friend Class bdlegdia
	Inherits System.Windows.Forms.Form
	'* Windows API Functions Declarations
	'   Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	'   Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'   Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
	'   Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'   Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
	
	'Constanst used by GetWindow
	'   Const GW_CHILD = 5
	'   Const GW_HWNDFIRST = 0
	'   Const GW_HWNDLAST = 1
	'   Const GW_HWNDNEXT = 2
	'   Const GW_HWNDPREV = 3
	'   Const GW_OWNER = 4
	
	'Constants
	Const HEEL_TOL As Short = 9 '9" heel
	
	'MsgBox constants
	'  Const IDYES = 6
	'  Const IDNO = 7
	'  Const IDCANCEL = 2
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()

        If sCurrentValues <> BDLEGDIA1.g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
            Select Case Response
                Case Project1.Public_Renamed.IDYES
                    ''--------PR_CreateMacro_Save("@C:\ACAD2018_SMITH-NEPHEW\Lookup Tables\jobst\draw.d")
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    saveInfoToDWG()
                    Me.Close()
                Case Project1.Public_Renamed.IDNO
                    Me.Close()
                Case Project1.Public_Renamed.IDCANCEL
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
        BodyLeg.BDLeg.Close()
        BodyMain.BodyMainDlg.Close()
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
        BDLEGDIA1.fNum = FreeFile()
        FileOpen(BDLEGDIA1.fNum, sDrafixFile, VB.OpenMode.Output)
        FN_Open = BDLEGDIA1.fNum

        'Write header information etc. to the DRAFIX macro file
        '
        PrintLine(BDLEGDIA1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(BDLEGDIA1.fNum, "//Patient - " & sName & ", " & sFileNo & "")
        PrintLine(BDLEGDIA1.fNum, "//by Visual Basic, Waist Height Leg")
        PrintLine(BDLEGDIA1.fNum, "//type - " & sType & "")

        'Clear user selections etc
        PrintLine(BDLEGDIA1.fNum, "UserSelection (" & BDLEGDIA1.QQ & "clear" & BDLEGDIA1.QQ & ");")
        PrintLine(BDLEGDIA1.fNum, "Execute (" & BDLEGDIA1.QQ & "menu" & BDLEGDIA1.QCQ & "SetColor" & BDLEGDIA1.QC & "Table(" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QCQ & "color" & BDLEGDIA1.QCQ & "bylayer" & BDLEGDIA1.QQ & "));")
        PrintLine(BDLEGDIA1.fNum, "Execute (" & BDLEGDIA1.QQ & "menu" & BDLEGDIA1.QCQ & "SetStyle" & BDLEGDIA1.QC & "Table(" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QCQ & "style" & BDLEGDIA1.QCQ & "bylayer" & BDLEGDIA1.QQ & "));")

        'Display Hour Glass symbol
        PrintLine(BDLEGDIA1.fNum, "Display (" & BDLEGDIA1.QQ & "cursor" & BDLEGDIA1.QCQ & "wait" & BDLEGDIA1.QCQ & sType & BDLEGDIA1.QQ & ");")
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
            BDLEGDIA1.g_nUnitsFac = 10 / 25.4
        Else
            BDLEGDIA1.g_nUnitsFac = 1
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
            Me.Text = "Body Suit Leg Data - Left"
        Else
            optRightLeg.Checked = True
            optLeftLeg.Enabled = False
            Me.Text = "Body Suit Leg Data - Right"
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

        BDLEGDIA1.g_nAnkleTape = 0
        BDLEGDIA1.g_nAnkleLen = 0
        BDLEGDIA1.g_nHeelLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(5).Text))

        If (BDLEGDIA1.g_nHeelLen <> 0) And (BDLEGDIA1.g_nHeelLen < HEEL_TOL) Then
            BDLEGDIA1.g_nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
            BDLEGDIA1.g_nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(BDLEGDIA1.g_nAnkleTape).Text))
        Else
            BDLEGDIA1.g_nAnkleTape = 7 'Large heel Ankle at +3 tape
            BDLEGDIA1.g_nAnkleLen = ARMEDDIA1.fnDisplayToInches(Val(Text1(BDLEGDIA1.g_nAnkleTape).Text))
        End If

        'Set below and above knee option buttons
        'Disable both if a heel value is given
        optBelowKnee.Checked = False
        optAboveKnee.Checked = True 'default for body suit
        optBelowKnee.Enabled = False
        optAboveKnee.Enabled = False

        If BDLEGDIA1.g_nHeelLen = 0 Then
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
        BDLEGDIA1.g_sChangeChecker = FN_ValuesString()
        Show()
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

    Private Sub bdlegdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Hide()
        'Maintain while loading DDE data
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
        'Reset in Form_LinkClose

        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        BDLEGDIA1.MainForm = Me
        'Initialize globals
        BDLEGDIA1.QQ = Chr(34) 'Double quotes (")
        BDLEGDIA1.NL = Chr(13) 'New Line
        BDLEGDIA1.CC = Chr(44) 'The comma (,)
        BDLEGDIA1.QCQ = BDLEGDIA1.QQ & BDLEGDIA1.CC & BDLEGDIA1.QQ
        BDLEGDIA1.QC = BDLEGDIA1.QQ & BDLEGDIA1.CC
        BDLEGDIA1.CQ = BDLEGDIA1.CC & BDLEGDIA1.QQ

        BDLEGDIA1.g_sPathJOBST = fnPathJOBST()
        BDLEGDIA1.g_nUnitsFac = 1 'Default to inches

        'Start a timer that will ensure that the
        'the VB programme ends if the DRAFIX DDE
        'link fails.
        Timer1.Interval = 6000
        Timer1.Enabled = True
        txtLeg.Text = "Left"
        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("Patient Details have not been entered", 48, "Body Suit Dialog")
            Me.Close()
            Exit Sub
        End If
        obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
        txtPatientName.Text = patient
        txtFileNo.Text = fileNo
        txtDiagnosis.Text = diagnosis
        txtSex.Text = sex
        txtAge.Text = age
        txtUnits.Text = units
        Form_LinkClose()
        readDWGInfo()
    End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
        'Check that data is all present and insert into drafix
        If Validate_Data() Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'Me.Hide()
            'BodyLeg.BDLeg.Hide()
            'BodyMain.BodyMainDlg.Hide()

            ''----------PR_CreateMacro_Save("@C:\ACAD2018_SMITH-NEPHEW\Lookup Tables\jobst\draw.d")
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            PR_CreateMacro_Save(sDrawFile)
            PR_DrawBodyLeftLeg()
            saveInfoToDWG()
            BodyMain.g_bIsBodyLoad = True
            'OK.Enabled = True
            'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Me.Close()
            'BodyLeg.BDLeg.Close()
            'BodyMain.BodyMainDlg.Close()
        End If
    End Sub

    Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
        'fNum is a global variable use in subsequent procedures
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BDLEGDIA1.fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

        'If this is a new drawing of a vest then Define the DATA Base
        'fields for the VEST Body and insert the BODYBOX symbol
        PR_PutLine("HANDLE hMPD, hLeg;")

        Update_DDE_Text_Boxes()

        Dim sSymbol As String

        sSymbol = "waistleg"

        If txtUidLeg.Text = "" Then
            'Define DB Fields
            PR_PutLine("@" & BDLEGDIA1.g_sPathJOBST & "\WAIST\WHFIELDS.D;")

            'Find "mainpatientdetails" and get position
            PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
            PR_PutLine("STRING sMPD_Name;")
            PR_PutLine("ANGLE  aMPD_Angle;")

            PR_PutLine("hMPD = UID (" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QC & Val(txtUidTitle.Text) & ");")
            PR_PutLine("if (hMPD)")
            PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel," & BDLEGDIA1.QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & BDLEGDIA1.QQ & ");")

            'Insert symbol
            PR_PutLine("if ( Symbol(" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QCQ & sSymbol & BDLEGDIA1.QQ & ")){")
            PR_PutLine("  Execute (" & BDLEGDIA1.QQ & "menu" & BDLEGDIA1.QCQ & "SetLayer" & BDLEGDIA1.QC & "Table(" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QCQ & "layer" & BDLEGDIA1.QCQ & "Data" & BDLEGDIA1.QQ & "));")
            If txtLeg.Text = "Left" Then
                PR_PutLine("  hLeg = AddEntity(" & BDLEGDIA1.QQ & "symbol" & BDLEGDIA1.QCQ & sSymbol & BDLEGDIA1.QC & "xyMPD_Origin.x + 1.5,xyMPD_Origin.y );")
            Else
                PR_PutLine("  hLeg = AddEntity(" & BDLEGDIA1.QQ & "symbol" & BDLEGDIA1.QCQ & sSymbol & BDLEGDIA1.QC & "xyMPD_Origin.x + 3,xyMPD_Origin.y );")
            End If
            PR_PutLine("  }")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel, " & BDLEGDIA1.QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & BDLEGDIA1.QQ & ");")
        Else
            'Use existing symbol
            PR_PutLine("hLeg = UID (" & BDLEGDIA1.QQ & "find" & BDLEGDIA1.QC & Val(txtUidLeg.Text) & ");")
            PR_PutLine("if (!hLeg) Exit(%cancel," & BDLEGDIA1.QQ & "Can't find >" & sSymbol & "< symbol to update!" & BDLEGDIA1.QQ & ");")
        End If

        'Update the BODY Box symbol
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "fileno" & BDLEGDIA1.QCQ & txtFileNo.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "Leg" & BDLEGDIA1.QCQ & txtLeg.Text & BDLEGDIA1.QQ & ");")

        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "TapeLengthsPt1" & BDLEGDIA1.QCQ & Mid(txtTapeLengths.Text, 1, 60) & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "TapeLengthsPt2" & BDLEGDIA1.QCQ & Mid(txtTapeLengths.Text, 61) & BDLEGDIA1.QQ & ");")

        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "AnkleTape" & BDLEGDIA1.QCQ & txtAnkleTape.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "FirstTape" & BDLEGDIA1.QCQ & txtFirstTape.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "LastTape" & BDLEGDIA1.QCQ & txtLastTape.Text & BDLEGDIA1.QQ & ");")

        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "ToeStyle" & BDLEGDIA1.QCQ & txtToeStyle.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "FootPleat1" & BDLEGDIA1.QCQ & txtFootPleat1.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "FootPleat2" & BDLEGDIA1.QCQ & txtFootPleat2.Text & BDLEGDIA1.QQ & ");")

        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "TopLegPleat1" & BDLEGDIA1.QCQ & txtTopLegPleat1.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "TopLegPleat2" & BDLEGDIA1.QCQ & txtTopLegPleat2.Text & BDLEGDIA1.QQ & ");")

        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "FootLength" & BDLEGDIA1.QCQ & txtFootLength.Text & BDLEGDIA1.QQ & ");")
        PR_PutLine("SetDBData( hLeg" & BDLEGDIA1.CQ & "Data" & BDLEGDIA1.QCQ & txtData.Text & BDLEGDIA1.QQ & ");")

        FileClose(BDLEGDIA1.fNum)
    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BDLEGDIA1.fNum, sLine)

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
        If (nHeelLen <> BDLEGDIA1.g_nHeelLen) Or (nAnkleLen <> BDLEGDIA1.g_nAnkleLen) Then
            '        MsgBox "Figure Ankle for this leg.", 48    'Not required  for body
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
        '    If cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" And Val(txtFootLength.Text) = 0 Then
        '        sError = sError + "A foot length must be given for Self Enclosed toes." + NL
        '    End If

        'Get Heel Length if it applies
        If nFirstTape <= 5 Then
            nHeelLength = ARMEDDIA1.fnDisplayToInches(Val(Text1(5).Text))
        Else
            nHeelLength = 0
        End If

        'Check if either the "above" or "below" knee options are
        'given for a footless
        '    If nHeelLength = 0 And optBelowKnee.Value = False And optAboveKnee.Value = False Then
        '        sError = sError + "This is a footless style." + NL
        '        sError = sError + "ABOVE or BELOW Knee must be selected." + NL
        '    End If

        'Check if -3 tape exists for Large Heel
        '    If nFirstTape <= 5 And nHeelLength >= 9 And Val(Text1(3)) = 0 Then
        '        sError = sError + "As the Heel is 9 inches and over." + NL
        '        sError = sError + "and there is no -3 tape the foot will not draw properly " + NL
        '    End If

        'Check if -1 1/2 tape exists for Small Heel
        '    If nFirstTape <= 5 And nHeelLength < 9 And Val(Text1(4)) = 0 Then
        '        sError = sError + "As the Heel is smaller than 9 inches." + NL
        '        sError = sError + "and there is no -1 1/2 tape the foot will not draw properly " + NL
        '    End If


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
        '    If nFirstTape > 5 Then
        '        If cboToeStyle.ListIndex > 0 Then
        '            sError = sError + "Toe style, " + cboToeStyle.List(cboToeStyle.ListIndex) + NL
        '        End If
        '        cboToeStyle.ListIndex = 0
        '    End If

        'Display Warning message (if required)
        '   If Len(sError) > 0 Then
        '       sError = "The TOE style given can't apply. As there is no Heel." + NL
        '       sError = sError + "Removing TOE style to reflect this."
        '       MsgBox sError, 64, "Warning - Errors in Data"
        '   End If


        'Yes / No type errors
        'In this case we warn the user that there is a problem!
        'They can continue or they can return to the dialog to make changes

        'Initialize error variables
        sError = ""
        iError = False

        'Make a stab at seeing if the foot length will be long enough
        'For Toe = "Self enclosed" or Toe = "Soft enclosed" and a foot length is given
        '   If nHeelLength > 0 And (cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" Or (cboToeStyle.List(cboToeStyle.ListIndex) = "Soft Enclosed" And Val(txtFootLength.Text) <> 0)) Then

        'Revise heel length to worst case
        'Scale heel to the 0 tape of a 30mmHg template
        '       nHeelLength = nHeelLength * .402
        'Reduce heel by 3
        '       nHeelLength = nHeelLength + (3 * .05025)

        '
        '       If nHeelLength >= 9 Then
        '           nTapeX = 2.71875 'Large heel, distance from heel to -3 tape
        '       Else
        '           nTapeX = 1.5     'Small heel, distance from heel to -1 1/2 tape
        '       End If

        '       nFootLength = fnDisplayToInches(Val(txtFootLength.Text))

        '       nValue = Sqr(nHeelLength ^ 2 + nTapeX ^ 2)
        '
        '        If nValue > nFootLength Then
        '            iError = True
        '            sError = sError + "The given foot length may be too small to correctly" + NL
        '            sError = sError + "position the toe!" + NL + NL
        '            sError = sError + "Do you wish to continue anyway ?"
        '        End If
        '    End If

        'Display Error message (if required) and return
        'These are non-fatal errors
        If iError = True Then
            iError = MsgBox(sError, 36, "Warning, Problems with data")
            If iError = Project1.Public_Renamed.IDYES Then
                Validate_Data = True
            Else
                Validate_Data = False
            End If
        Else
            Validate_Data = True
        End If

    End Function
    Private Sub PR_DrawBodyLeftLeg()
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        Dim xyStart(5), xyEnd(5) As Double
        xyStart(1) = 0
        xyEnd(1) = 1.25
        xyStart(2) = 0
        xyEnd(2) = 1.4375
        xyStart(3) = 1.5
        xyEnd(3) = 1.4375
        xyStart(4) = 1.5
        xyEnd(4) = 1.25
        xyStart(5) = 0
        xyEnd(5) = 1.25

        Dim strTag(14), strTextString(14) As String
        strTag(1) = "fileno"
        strTag(2) = "Leg"
        strTag(3) = "TapeLengthsPt1"
        strTag(4) = "TapeLengthsPt2"
        strTag(5) = "AnkleTape"
        strTag(6) = "FirstTape"
        strTag(7) = "LastTape"
        strTag(8) = "ToeStyle"
        strTag(9) = "FootPleat1"
        strTag(10) = "FootPleat2"
        strTag(11) = "TopLegPleat1"
        strTag(12) = "TopLegPleat2"
        strTag(13) = "FootLength"
        strTag(14) = "Data"

        strTextString(1) = txtFileNo.Text
        strTextString(2) = txtLeg.Text
        strTextString(3) = Mid(txtTapeLengths.Text, 1, 60)
        strTextString(4) = Mid(txtTapeLengths.Text, 61)
        strTextString(5) = txtAnkleTape.Text
        strTextString(6) = txtFirstTape.Text
        strTextString(7) = txtLastTape.Text
        strTextString(8) = cboToeStyle.Text
        strTextString(9) = txtFootPleat1.Text
        strTextString(10) = txtFootPleat2.Text
        strTextString(11) = txtTopLegPleat1.Text
        strTextString(12) = txtTopLegPleat2.Text
        strTextString(13) = txtFootLength.Text
        strTextString(14) = txtData.Text

        Dim XVal As Double = 1.5
        Dim strBlkName As String = "BODYLEFTLEG"
        'If txtLeg.Text = "Right" Then
        '    XVal = 3.0
        '    strBlkName = "WAISTRIGHTLEG"
        'End If

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRefPatient As BlockReference = acTrans.GetObject(blkId, OpenMode.ForRead)
            Dim ptPosition As Point3d = blkRefPatient.Position
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(strBlkName) Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = strBlkName
                Dim acPoly As Polyline = New Polyline()
                Dim ii As Double
                For ii = 1 To 5
                    acPoly.AddVertexAt(ii - 1, New Point2d(xyStart(ii), xyEnd(ii)), 0, 0, 0)
                Next ii
                blkTblRecCross.AppendEntity(acPoly)

                Dim acText As DBText = New DBText()
                acText.Position = New Point3d(0.71875, 1.25, 0)
                acText.Height = 0.1
                acText.TextString = "WH-LEG"
                acText.Rotation = 0
                acText.Justify = AttachmentPoint.BottomCenter
                acText.AlignmentPoint = New Point3d(0.71875, 1.25, 0)
                blkTblRecCross.AppendEntity(acText)
                Dim acAttDef As New AttributeDefinition
                For ii = 1 To 14
                    acAttDef = New AttributeDefinition
                    acAttDef.Position = New Point3d(0, 0, 0)
                    acAttDef.Prompt = strTag(ii)
                    acAttDef.Tag = strTag(ii)
                    acAttDef.TextString = strTextString(ii)
                    acAttDef.Height = 1
                    acAttDef.Justify = AttachmentPoint.BaseLeft
                    acAttDef.Invisible = True
                    blkTblRecCross.AppendEntity(acAttDef)
                Next
                '' Set the layer DATA as current layer
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                If acLyrTbl.Has("DATA") = True Then
                    acCurDb.Clayer = acLyrTbl("DATA")
                End If
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl(strBlkName)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(ptPosition.X + XVal, ptPosition.Y, 0), blkRecId)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(ptPosition.X + XVal, ptPosition.Y, 0)))
                'End If
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                '' Open the Block table record Model space for write
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If Not acAttDef.Constant Then
                            Dim acAttRef As New AttributeReference
                            acAttRef.SetAttributeFromBlock(acAttDef, blkRef.BlockTransform)
                            If acAttRef.Tag.ToUpper().Equals("fileno", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFileNo.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("Leg", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLeg.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt1", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = Mid(txtTapeLengths.Text, 1, 60)
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = Mid(txtTapeLengths.Text, 61)
                            ElseIf acAttRef.Tag.ToUpper().Equals("AnkleTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtAnkleTape.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("FirstTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFirstTape.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LastTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLastTape.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ToeStyle", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboToeStyle.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("FootPleat1", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFootPleat1.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("FootPleat2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFootPleat2.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat1", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtTopLegPleat1.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtTopLegPleat2.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("FootLength", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFootLength.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("Data", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtData.Text
                            Else
                                Continue For
                            End If
                            blkRef.AttributeCollection.AppendAttribute(acAttRef)
                            acTrans.AddNewlyCreatedDBObject(acAttRef, True)
                        End If
                    End If
                Next
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("BodyLeftLeg", "BODYLEFTDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("BodyLeftLeg", "BODYLEFTDIC")
            End If

            Dim resbuf As New ResultBuffer
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFootPleat1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFootPleat2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTopLegPleat2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTopLegPleat1.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_0.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_3.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_9.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_10.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_11.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_12.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_13.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_14.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_15.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_16.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_17.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_18.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_19.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_20.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_21.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_22.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_23.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_24.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_25.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_26.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_27.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_28.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Text1_29.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboToeStyle.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFootLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLeg.Text))

            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optLeftLeg.Checked))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optRightLeg.Checked))

            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optBelowKnee.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optAboveKnee.Checked))
            Dim ii As Double
            For ii = 0 To 29
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), labTapeInInches(ii).Text))
            Next ii

            _sClass.SetXrecord(resbuf, "BodyLeftLeg", "BODYLEFTDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("BodyLeftLeg", "BODYLEFTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                txtFootPleat1.Text = arr(0).Value
                txtFootPleat2.Text = arr(1).Value
                txtTopLegPleat2.Text = arr(2).Value
                txtTopLegPleat1.Text = arr(3).Value

                _Text1_0.Text = arr(4).Value
                _Text1_1.Text = arr(5).Value
                _Text1_2.Text = arr(6).Value
                _Text1_3.Text = arr(7).Value
                _Text1_4.Text = arr(8).Value
                _Text1_5.Text = arr(9).Value
                _Text1_6.Text = arr(10).Value
                _Text1_7.Text = arr(11).Value
                _Text1_8.Text = arr(12).Value
                _Text1_9.Text = arr(13).Value
                _Text1_10.Text = arr(14).Value
                _Text1_11.Text = arr(15).Value
                _Text1_12.Text = arr(16).Value
                _Text1_13.Text = arr(17).Value
                _Text1_14.Text = arr(18).Value
                _Text1_15.Text = arr(19).Value
                _Text1_16.Text = arr(20).Value
                _Text1_17.Text = arr(21).Value
                _Text1_18.Text = arr(22).Value
                _Text1_19.Text = arr(23).Value
                _Text1_20.Text = arr(24).Value
                _Text1_21.Text = arr(25).Value
                _Text1_22.Text = arr(26).Value
                _Text1_23.Text = arr(27).Value
                _Text1_24.Text = arr(28).Value
                _Text1_25.Text = arr(29).Value
                _Text1_26.Text = arr(30).Value
                _Text1_27.Text = arr(31).Value
                _Text1_28.Text = arr(32).Value
                _Text1_29.Text = arr(33).Value

                cboToeStyle.Text = arr(34).Value
                txtFootLength.Text = arr(35).Value
                txtLeg.Text = arr(36).Value
                'optLeftLeg.Checked = arr(37).Value
                'optRightLeg.Checked = arr(38).Value

                optBelowKnee.Checked = arr(37).Value
                optAboveKnee.Checked = arr(38).Value

                Dim ii As Double
                For ii = 0 To 29
                    labTapeInInches(ii).Text = arr(ii + 39).Value
                Next ii
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
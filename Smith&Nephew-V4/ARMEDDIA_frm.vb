Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Friend Class armeddia
    Inherits System.Windows.Forms.Form

    Dim g_strTapeLen(18) As String
    'Project:   ARMEDDIA
    'Purpose:   Editor for Arms and Vest sleeves.
    '
    'Version:   3.01
    'Date:      28.Mar.95
    'Author:    Gary George
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    'Dec 98     GG      Ported to VB 5
    '
    'Notes:-
    '
    ' Known Bugs
    '
    ' 1 With a gauntlet extra or missing points above the wrist
    '   are not checked for.
    '
    ' 2 The vertex mapping that is currently used only with the
    '   gaunlets between the palm and wrist should be extended
    '   to included the whole arm.
    '
    ' 3 Fails if DRAFIX Help is in use.
    '
    '

    'UPGRADE_WARNING: Event cboContracture.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboContracture_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboContracture.SelectedIndexChanged
		Dim sType As String
		
		If g_NoElbowTape Then Exit Sub
		
		If cboContracture.Text <> txtContracture.Text Then
			txtContracture.Text = cboContracture.Text
			sType = VB.Left(txtContracture.Text, 2)
			Select Case sType
				Case "No"
					'Reset mms at elbow to standard and recalculate
					Select Case txtMM.Text
						Case "15mm"
							EditMMs(g_iElbowTape).Text = CStr(12)
						Case "20mm"
							EditMMs(g_iElbowTape).Text = CStr(16)
						Case "25mm"
							EditMMs(g_iElbowTape).Text = CStr(20)
					End Select
					PR_CalculateFromMMs(g_iElbowTape)
					cboContracture.Text = ""
					txtContracture.Text = ""
				Case "10"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 12)
				Case "36"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 15)
				Case "71"
					'set the reduction and back calculate grams and mms
					PR_CalculateFromReduction(g_iElbowTape, 17)
			End Select
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cboLining.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLining_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLining.SelectedIndexChanged
		If cboLining.Text <> txtLining.Text Then
			If cboLining.Text = "None" Then cboLining.Text = ""
			txtLining.Text = cboLining.Text
		End If
	End Sub
	
	Private Sub cboLinings_Change()
		If g_NoElbowTape Then Exit Sub
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        'This cancel will delete the tempory curve then exit.
        'Checks to see if there is an Available drafix instance to
        'sendkeys to.
        If g_ReDrawn = True Then
            PR_CancelProfileEdits()
            Try
                Dim _sClass As New SurroundingClass()
                If (_sClass.GetXrecord("ArmEditInfo", "ARMEDITDIC") IsNot Nothing) Then
                    _sClass.RemoveXrecord("ArmEditInfo", "ARMEDITDIC")
                End If
            Catch ex As Exception
            End Try
        End If
        Me.Close()
        'End
    End Sub
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		'Check that data is all present and commit into drafix
		If FN_Validate_Data() Then
            PR_DrawAndCommitProfileEdits()
            Try
                Dim _sClass As New SurroundingClass()
                If (_sClass.GetXrecord("ArmEditInfo", "ARMEDITDIC") IsNot Nothing) Then
                    _sClass.RemoveXrecord("ArmEditInfo", "ARMEDITDIC")
                End If
            Catch ex As Exception
            End Try
            Me.Close()
        End If
	End Sub
	
	Private Sub cmdDraw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDraw.Click
		'Check that data is all present and insert into drafix
		If FN_Validate_Data() Then
            PR_DrawProfileEdits()
            g_ReDrawn = True
            SaveArmEditInfo()
            Me.Close()
        End If
	End Sub
	
	Private Sub EditMMs_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditMMs.Enter
		Dim Index As Short = EditMMs.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(EditMMs(Index))
	End Sub
	
	Private Sub EditMMs_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditMMs.Leave
		Dim Index As Short = EditMMs.GetIndex(eventSender)
		If FN_CheckValue(EditMMs(Index), "MM's") Then
			PR_CalculateFromMMs(Index)
		End If
	End Sub
	
	Private Sub EditTapes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditTapes.Enter
		Dim Index As Short = EditTapes.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(EditTapes(Index))
	End Sub
	
	Private Sub EditTapes_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditTapes.Leave
		Dim Index As Short = EditTapes.GetIndex(eventSender)
		'Check length given is valid
		'Display value in inches
		If FN_CheckValue(EditTapes(Index), "Tape Length") Then
			InchText(Index).Text = ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(EditTapes(Index).Text)))
			PR_CalculateFromMMs(Index)
		End If
	End Sub
	
	Private Function FN_CheckValue(ByRef TextBox As System.Windows.Forms.Control, ByRef sMessage As String) As Short
		'Check that a valid numeric value has been Entered
		Dim sChar, sText As String
		Dim iLen, nn As Short
		sText = TextBox.Text
		iLen = Len(sText)
		FN_CheckValue = True
		For nn = 1 To iLen
			sChar = Mid(sText, nn, 1)
			If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
				MsgBox("Invalid " & sMessage & " has been entered", 48, "Arm Details")
				TextBox.Focus()
				FN_CheckValue = False
				Exit For
			End If
		Next nn
	End Function
	
	Private Function FN_DrawOpen(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
        'Open the DRAFIX macro file
        'Initialise Global variables

        'Open file
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.fNum = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(ARMEDDIA1.fNum, sDrafixFile, VB.OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FN_DrawOpen = ARMEDDIA1.fNum

        'Initialise String globals
        'UPGRADE_WARNING: Couldn't resolve default property of object cc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.cc = Chr(44) 'The comma ( , )
        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.NL = Chr(10) 'The new line character
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.QQ = Chr(34) 'Double quotes ( " )
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object cc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.QCQ = ARMEDDIA1.QQ & ARMEDDIA1.cc & ARMEDDIA1.QQ 'Quote Comma Quote ( "," )
        'UPGRADE_WARNING: Couldn't resolve default property of object cc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.QC = ARMEDDIA1.QQ & ARMEDDIA1.cc 'Quote Comma ( ", )
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object cc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.CQ = ARMEDDIA1.cc & ARMEDDIA1.QQ 'Comma Quote ( ," )

        'Globals to reduced drafix code written to file
        ARMEDDIA1.g_sCurrentLayer = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMEDDIA1.g_nCurrTextAspect = 0.6

        'Write header information etc. to the DRAFIX macro file
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "//DRAFIX Arm Editing Macro created - " & DateString & "  " & TimeString)
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "//by Visual Basic")

    End Function

    Private Function FN_Validate_Data() As Short

        Dim sError, NL As String
        Dim nFirstTape, ii, nLastTape As Short
        Dim iError As Short
        Dim nValue As Double

        NL = Chr(10) 'new line
        sError = ""

        'Display Error message (if required) and return
        If Len(sError) > 0 Then
            MsgBox(sError, 48, "Errors in Data")
            FN_Validate_Data = False
            Exit Function
        Else
            FN_Validate_Data = True
        End If


    End Function

    'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkClose()
        Dim iValue, ii, nn, filenum As Short
        Dim nValue As Double
        Dim textline As String
        Dim sDistalStyle, sProximalStyle As String

        'Units
        If txtUnits.Text = "cm" Then
            ARMEDDIA1.g_nUnitsFac = 10 / 25.4
        Else
            ARMEDDIA1.g_nUnitsFac = 1
        End If

        ARMEDDIA1.MainForm = Me
        'Other globals
        ARMEDDIA1.g_sID = txtID.Text & txtArm.Text
        ARMEDDIA1.g_sFileNo = Mid(txtID.Text, 5)
        g_sStyle = Mid(txtID.Text, 1, 4)
        ARMEDDIA1.g_sSide = txtArm.Text

        'Proximal and Distal end tape Styles
        sDistalStyle = Mid(txtID.Text, 1, 2)
        sProximalStyle = Mid(txtID.Text, 3, 2)

        'Get gauntlet Details
        g_Gauntlet = ARMEDDIA1.fnGetNumber(txtGauntlet.Text, 1)
        If g_Gauntlet Then
            g_iWristNo = ARMEDDIA1.fnGetNumber(txtGauntlet.Text, 5) - 1
            g_iPalmNo = ARMEDDIA1.fnGetNumber(txtGauntlet.Text, 6) - 1
            g_nPalmWristDist = ARMEDDIA1.fnDisplayToInches(ARMEDDIA1.fnGetNumber(txtGauntlet.Text, 9))
        End If

        'Title bar display
        If txtArm.Text = "Left" Then
            Me.Text = "ARM Edit - Left [" & ARMEDDIA1.g_sFileNo & "]"
        End If
        If txtArm.Text = "Right" Then
            Me.Text = "ARM Edit - Right [" & ARMEDDIA1.g_sFileNo & "]"
        End If

        'Display - Given Length, Length in inches, MMs, Grams and Reduction at each tape
        'Store values into working arrays. Store also the initial values for change checking
        For ii = 0 To 17
            ''----------nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3)) / 10
            ''nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3))
            nValue = Val(FN_GetString(txtTapeLengths.Text, ii + 1))
            ''nValue = Val(g_strTapeLen(ii))
            If nValue > 0 Then

                EditTapes(ii).Text = CStr(nValue)
                g_nLengths(ii) = nValue
                g_nLengthsInit(ii) = nValue
                InchText(ii).Text = ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(nValue))

                iValue = Val(Mid(txtTapeMMs.Text, (ii * 3) + 1, 3))
                EditMMs(ii).Text = CStr(iValue)
                ARMEDDIA1.g_iMMs(ii) = iValue
                g_iMMsInit(ii) = iValue

                iValue = Val(Mid(txtGrams.Text, (ii * 3) + 1, 3))
                EditGrams(ii).Text = Str(iValue)
                ARMEDDIA1.g_iGms(ii) = iValue
                g_iGmsInit(ii) = iValue

                iValue = Val(Mid(txtReduction.Text, (ii * 3) + 1, 3))
                EditReductions(ii).Text = Str(iValue)
                ARMEDDIA1.g_iRed(ii) = iValue
                g_iRedInit(ii) = iValue

            End If
        Next ii

        'Establish First tape
        'Disable unused tapes
        ARMEDDIA1.g_iFirstTape = -1
        nValue = 0
        For ii = 0 To 17
            nValue = Val(EditTapes(ii).Text)
            If nValue > 0 Then Exit For
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        Next ii
        ARMEDDIA1.g_iFirstTape = ii


        'If the first tape is on the elbow then disable
        'first tape editing
        If ARMEDDIA1.g_iFirstTape = 10 Then
            EditTapes(ARMEDDIA1.g_iFirstTape).Enabled = False
            EditMMs(ARMEDDIA1.g_iFirstTape).Enabled = False
            EditMMs(ARMEDDIA1.g_iFirstTape).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        End If

        'Establish Last tape (DISABLE down to Last Tape)
        ARMEDDIA1.g_iLastTape = -1
        nValue = 0
        For ii = 17 To 0 Step -1
            nValue = Val(EditTapes(ii).Text)
            If nValue > 0 Then Exit For
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        Next ii
        ARMEDDIA1.g_iLastTape = ii

        'If the Proximal style is Flap or is a Vest Raglan
        'then disable last tape editing
        If sProximalStyle = "FP" Or sProximalStyle = "VR" Then
            EditTapes(ARMEDDIA1.g_iLastTape).Enabled = False
            EditMMs(ARMEDDIA1.g_iLastTape).Enabled = False
            EditMMs(ARMEDDIA1.g_iLastTape).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        End If

        'Load the original profile from file
        'Setup profile vertex to tape mappings
        ''--PR_GetProfileFromFile("C:\JOBST\ARMCURVE.DAT")
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        PR_GetProfileFromFile(sSettingsPath & "\ARMCURVE.DAT")


        'Load fabric conversion charts from file
        'Establish Fabric
        If Mid(txtFabric.Text, 1, 3) = "Pow" Then
            ''-----ARMEDDIA1.PR_LoadFabricFromFile(g_MATERIAL, ARMEDDIA1.g_sPathJOBST & "\TEMPLTS\POWERNET.DAT")
            ARMEDDIA1.PR_LoadFabricFromFile(g_MATERIAL, sSettingsPath & "\POWERNET.DAT")
        Else
            ''---------ARMEDDIA1.PR_LoadFabricFromFile(g_MATERIAL, ARMEDDIA1.g_sPathJOBST & "\TEMPLTS\BOBINNET.DAT")
            ARMEDDIA1.PR_LoadFabricFromFile(g_MATERIAL, sSettingsPath & "\BOBINNET.DAT")
        End If

        'Establish Fabric modulus
        g_iModulus = Val(Mid(txtFabric.Text, 5, 3))

        'If there is no elbow tape or elbow tape is the last tape
        'then disable linings and contractures
        g_iElbowTape = 10
        g_sOriginalContracture = txtContracture.Text
        g_sOriginalLining = txtLining.Text
        If Val(EditTapes(g_iElbowTape).Text) = 0 Or g_iElbowTape = ARMEDDIA1.g_iFirstTape Or g_iElbowTape = ARMEDDIA1.g_iLastTape Then
            cboContracture.Enabled = False
            cboLining.Enabled = False
            g_NoElbowTape = True
        Else
            cboContracture.Text = txtContracture.Text
            cboLining.Text = txtLining.Text
        End If

        Show()
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
        cmdCancel.Focus()

    End Sub

    'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
        If CmdStr = "Cancel" Then
            Cancel = 0
            'End
        End If
    End Sub

    Private Sub armeddia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            Dim filenum As Short
            Dim textline As String
            'Hide form while loading
            Hide()

            'Check if a previous instance is running
            'If it is warn user and exit
            'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'If App.PrevInstance Then
            '    MsgBox("The Arm Edit Module is already running!" & Chr(13) & "Use ALT-TAB to access it .", 16, "Leg Edit Warning")
            '    Return 'End
            'End If

            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 10)
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)

            'Set Tape Numbers
            num(0).Text = "-6"
            num(1).Text = "-4" & Chr(189)
            num(2).Text = "-3"
            num(3).Text = "-1" & Chr(189)
            num(4).Text = "0"
            num(5).Text = "1" & Chr(189)
            num(6).Text = "3"
            num(7).Text = "4" & Chr(189)
            num(8).Text = "6"
            num(9).Text = "7" & Chr(189)
            num(10).Text = "9"
            num(11).Text = "10" & Chr(189)
            num(12).Text = "12"
            num(13).Text = "13" & Chr(189)
            num(14).Text = "15"
            num(15).Text = "16" & Chr(189)
            num(16).Text = "18"
            num(17).Text = "19" & Chr(189)

            txtUnits.Text = ""
            txtUIDArm.Text = ""
            txtUIDTempArm.Text = ""
            txtUIDCurve.Text = ""
            txtArm.Text = ""
            txtID.Text = ""
            txtTapeLengths.Text = ""
            txtTapeMMs.Text = ""
            txtGrams.Text = ""
            txtReduction.Text = ""
            txtContracture.Text = ""
            txtLining.Text = ""
            txtStump.Text = ""
            txtFabric.Text = ""

            Dim strTapeLen(18) As String
            g_strTapeLen = strTapeLen

            ''Added for #248 in the issue list
            Dim ii As Short
            For ii = 0 To 17
                ARMEDDIA1.g_iGms(17) = 0
                ARMEDDIA1.g_iRed(17) = 0
                ARMEDDIA1.g_iMMs(17) = 0
                g_nLengths(ii) = 0
                g_iChanged(ii) = 0
            Next

            'Linings Combo
            cboLining.Items.Add("Inside Lining")
            cboLining.Items.Add("Outside Lining")
            cboLining.Items.Add("Lining")
            cboLining.Items.Add("Full Lining")
            cboLining.Items.Add("Reinforced Elbow")
            cboLining.Items.Add("None")

            'Contractures Combo
            cboContracture.Items.Add("10-35 Degrees")
            cboContracture.Items.Add("36-70 Degrees")
            cboContracture.Items.Add("71 Degrees and Over")
            cboContracture.Items.Add("None")

            ARMEDDIA1.g_nUnitsFac = 1 'Default to inches
            g_ShortArm = False
            g_NoElbowTape = False
            ARMEDDIA1.g_sPathJOBST = fnPathJOBST()
            Dim strMod(18), strConv(18) As String
            g_MATERIAL.Modulus = strMod
            g_MATERIAL.Conversion_Renamed = strConv
            'Set redraw flag to False
            g_ReDrawn = False
            g_ShortArm = False
            g_sTempArmHandle = ""
            g_sOriginalCurveHandle = ""
            Me.Hide()

            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim obj As New BlockCreation.BlockCreation
            Dim blkId As ObjectId = New ObjectId()
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Can't find Patient Details", 48, "Arm Edit Dialog")
                Me.Close()
                Exit Sub
            End If

            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Arm Edit Dialog")
                Me.Close()
                Exit Sub
            End If
            txtUnits.Text = units
            If PR_GetArmProfile() = False Then
                Me.Close()
                Exit Sub
            End If
            Form_LinkClose()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    Private Sub PR_CalculateFromMMs(ByRef Index As Short)
        'Calculate the reduction based on the given lengths
        'and the given MMs

        'If no change has been made to either length or MMs then exit
        If Val(EditTapes(Index).Text) = g_nLengths(Index) And Val(EditMMs(Index).Text) = ARMEDDIA1.g_iMMs(Index) Then Exit Sub

        Dim sConversion As String
        Dim nLength As Double
        Dim iPrevVal, ii, iVal As Short
        Dim iGrams, iReduction, iMMs As Short

        'Calculate grams
        'NB allows use of decimal MMs to fudge value
        '
        nLength = ARMEDDIA1.fnDisplayToInches(CDbl(EditTapes(Index).Text))
        iGrams = Int(nLength * Val(EditMMs(Index).Text))
        iMMs = Int(Val(EditMMs(Index).Text))

        'Get conversion string based on modulus
        sConversion = ""
        For ii = 0 To 17
            If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
                sConversion = g_MATERIAL.Conversion_Renamed(ii)
                Exit For
            End If
        Next ii

        If sConversion = "" Then
            MsgBox("Fabric Modulus not found in conversion chart", 16, "ARM Edit")
            cmdCancel_Click(cmdCancel, New System.EventArgs())
            Return 'End
        End If

        iPrevVal = 0
        For ii = 0 To 22
            iVal = Val(Mid(sConversion, (ii * 4) + 1, 4))
            If iVal >= iGrams Then Exit For
            iPrevVal = iVal
        Next ii

        Select Case ii
            Case 0
                'Minimum reduction
                iReduction = 10
            Case 23
                'Maximum reduction
                iReduction = 32
            Case Else
                'Get reduction closest to given grams
                If (iGrams - iPrevVal) < (iVal - iGrams) Then
                    iReduction = ii + 9
                Else
                    iReduction = ii + 10
                End If
        End Select

        'Modify stored values
        ARMEDDIA1.g_iMMs(Index) = iMMs
        g_nLengths(Index) = Val(EditTapes(Index).Text)
        ARMEDDIA1.g_iRed(Index) = iReduction
        ARMEDDIA1.g_iGms(Index) = iGrams

        'Change display
        EditMMs(Index).Text = CStr(iMMs) 'Reset MMs to integer value
        EditGrams(Index).Text = Str(iGrams)
        EditReductions(Index).Text = Str(iReduction)

        'Show that this tape has been modified
        g_iChanged(Index) = 1

    End Sub

    Private Sub PR_CalculateFromReduction(ByRef Index As Short, ByRef iReduction As Short)

        'Back calculate from the given reduction revised grams and mms

        Dim sConversion As String
        Dim nLength As Double
        Dim iPrevVal, ii, iVal As Short
        Dim iGrams, iMMs As Short

        'Quiet exit for bad data
        If EditTapes(Index).Text = "" Then
            Exit Sub
        End If
        nLength = ARMEDDIA1.fnDisplayToInches(Val(EditTapes(Index).Text))
        If nLength <= 0 Then Exit Sub

        'Get conversion string based on modulus
        sConversion = ""
        For ii = 0 To 17
            If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
                sConversion = g_MATERIAL.Conversion_Renamed(ii)
                Exit For
            End If
        Next ii

        If sConversion = "" Then
            MsgBox("Fabric Modulus not found in conversion chart", 16, "ARM Edit")
            cmdCancel_Click(cmdCancel, New System.EventArgs())
            Return 'End
        End If


        'Get grams from conversion string based on given reduction
        'N.B. Reductions only run from 10 to 32
        '
        Select Case iReduction
            Case Is < 10
                'Use lowest available grams for those less than 10
                iGrams = Val(Mid(sConversion, 1, 4))
            Case Is > 32
                'Use highest available grams for those greater than 32
                iGrams = Val(Mid(sConversion, (22 * 4) + 1, 4))
            Case Else
                iGrams = Val(Mid(sConversion, ((iReduction - 10) * 4) + 1, 4))
        End Select

        'Back calculate MMs from grams and length
        iMMs = iGrams / nLength

        'Modify stored values
        ARMEDDIA1.g_iMMs(Index) = iMMs
        g_nLengths(Index) = Val(EditTapes(Index).Text)
        ARMEDDIA1.g_iRed(Index) = iReduction
        ARMEDDIA1.g_iGms(Index) = iGrams

        'Change display
        EditMMs(Index).Text = CStr(iMMs) 'Reset MMs to integer value
        EditGrams(Index).Text = Str(iGrams)
        EditReductions(Index).Text = Str(iReduction)

        'Show that this tape has been modified
        g_iChanged(Index) = 1

    End Sub

    Private Sub PR_CancelProfileEdits()
        Dim ii As Short

        If g_ReDrawn = True Then
            ''ARMEDDIA1.fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Cancel", txtArm)
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            ARMEDDIA1.fNum = FN_DrawOpen(sDrawFile, "EDIT", "Cancel", txtArm)

            If g_ShortArm = True Then
                'Restore original curve (Short Arms only)
                PR_PutLine("HANDLE  hCurv;")
                'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(ARMEDDIA1.fNum, "hCurv = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDTempArm.Text) & ");")
                For ii = 0 To g_iProfile
                    'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PrintLine(ARMEDDIA1.fNum, "SetVertex(hCurv," & ii + 1 & "," & xyOriginal(ii).X & "," & xyOriginal(ii).y & ");")
                Next ii
                If g_sTempArmHandle <> "" Then
                    PR_RedrawOriginalCurve(g_sTempArmHandle)
                End If

            Else
                'Delete Curve Copy
                PR_PutLine("HANDLE hTempCurv;")
                'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(ARMEDDIA1.fNum, "hTempCurv = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDTempArm.Text) & ");")
                'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(ARMEDDIA1.fNum, "DeleteEntity(hTempCurv);")
                If g_sTempArmHandle <> "" Then
                    PR_DeleteExistingEntity()
                End If
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileClose(ARMEDDIA1.fNum)
        End If
    End Sub

    Private Sub PR_DeleteTapeLabel(ByRef xyPt As ARMEDDIA1.XY, ByRef nTape As Object)
        'Deletes the tape label and the text at the given point
        Dim slayer As String

        Dim y1, x1, x2, y2 As Double

        x1 = xyPt.X
        x2 = xyPt.X + 1
        y1 = xyPt.y
        y2 = xyPt.y + 2
        ''Changed for #248 in the issue list
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            x2 = xyPt.X + (1 * 2.54)
            y2 = xyPt.y + (2 * 2.54)
        End If
        ''------slayer = ARMEDDIA1.g_sFileNo & Mid(ARMEDDIA1.g_sSide, 1, 1) & nTape
        slayer = "TEMPLATE" & ARMEDDIA1.g_sSide

        PrintLine(ARMEDDIA1.fNum, "hChan=Open(" & ARMEDDIA1.QQ & "selection" & ARMEDDIA1.QCQ & "(type = 'Text' OR type = 'circle') AND layer = '" & slayer & "' AND TOTally INside " & x1 & y1 & x2 & y2 & ARMEDDIA1.QQ & ");")
        PrintLine(ARMEDDIA1.fNum, "if(hChan)")
        PrintLine(ARMEDDIA1.fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}")
        PrintLine(ARMEDDIA1.fNum, "Close(" & ARMEDDIA1.QQ & "selection" & ARMEDDIA1.QC & "hChan);")

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        ''Select Text or Circle with corresponding Layer
        Dim filterType(4) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
        filterType.SetValue(New TypedValue(DxfCode.Start, "TEXT"), 1)
        filterType.SetValue(New TypedValue(DxfCode.Start, "CIRCLE"), 2)
        filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 3)
        filterType.SetValue(New TypedValue(DxfCode.LayerName, slayer), 4)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim ptFstCorner As Point3d = New Point3d(x1, y1, 0)
        Dim ptSecCorner As Point3d = New Point3d(x2, y2, 0)
        Dim selResult As PromptSelectionResult = ed.SelectCrossingWindow(ptFstCorner, ptSecCorner, selFilter)
        If selResult.Status <> PromptStatus.OK Then
            Exit Sub
        End If
        Dim selectionSet As SelectionSet = selResult.Value
        Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim dbObj As DBObject = acTr.GetObject(idObject, OpenMode.ForWrite)
                dbObj.Erase()
            Next
            acTr.Commit()
        End Using
    End Sub

    Private Sub PR_DrawAndCommitProfileEdits()

        'Create a macro to commit the changes to the Original profile

        Dim xyPt As ARMEDDIA1.XY
        Dim ii, iProfileVertex As Short
        Dim ProfileChanged As Short
        Dim nSeam As Double
        Dim sGms, sLen, sMM, sRed As String
        Dim sPackedGms, sPackedLengths, sPackedMMs, sPackedRed As String
        Dim iGms, iLen, iMM, iRed As Short

        nSeam = 0.1875
        ''ARMEDDIA1.fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtArm)
        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        ARMEDDIA1.fNum = FN_DrawOpen(sDrawFile, "EDIT", "Test", txtArm)
        PR_PutLine("HANDLE  hCurv, hTempCurv, hArm, hChan, hEnt;")

        ARMDIA1.PR_SetLayer("Construct")

        'Delete Curve Copy (Only if redraw has been used)
        If g_ReDrawn = True And g_ShortArm = False Then
            PrintLine(ARMEDDIA1.fNum, "hTempCurv = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDTempArm.Text) & ");")
            PrintLine(ARMEDDIA1.fNum, "DeleteEntity(hTempCurv);")
            If g_sTempArmHandle <> "" Then
                PR_DeleteExistingEntity()
            End If
        End If

        'Modify Original curve  (need not have been redrawn first)
        PrintLine(ARMEDDIA1.fNum, "hCurv = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDCurve.Text) & ");")
        ProfileChanged = False
        For ii = 0 To 17
            If g_iChanged(ii) < 0 Or g_iChanged(ii) > 0 Then
                iProfileVertex = g_iVertexMap(ii)
                xyPt.X = xyProfile(iProfileVertex - 1).X
                ''xyPt.y = ((ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    xyPt.y = ((g_nLengths(ii) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
                Else
                    xyPt.y = ((ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
                End If
                PrintLine(ARMEDDIA1.fNum, "SetVertex(hCurv," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
                xyProfile(iProfileVertex - 1) = xyPt
                xyPt.y = xyOtemplate.y
                PR_DeleteTapeLabel(xyPt, ii + 1)
                ''----------ARMDIA1.PR_MakeXY(xyARMPt, xyPt.X, xyPt.y)
                ''----------ARMDIA1.PR_PutTapeLabel(ii + 1, xyARMPt, ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)), ARMEDDIA1.g_iMMs(ii), ARMEDDIA1.g_iGms(ii), ARMEDDIA1.g_iRed(ii))
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    PR_DrawTapeLabel(ii + 1, xyPt, g_nLengths(ii), ARMEDDIA1.g_iMMs(ii), ARMEDDIA1.g_iGms(ii), ARMEDDIA1.g_iRed(ii))
                Else
                    PR_DrawTapeLabel(ii + 1, xyPt, ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)), ARMEDDIA1.g_iMMs(ii), ARMEDDIA1.g_iGms(ii), ARMEDDIA1.g_iRed(ii))
                End If
                ProfileChanged = True
            End If
        Next ii
        If g_sOriginalCurveHandle <> "" Then
            PR_RedrawOriginalCurve(g_sOriginalCurveHandle, True)
        End If

        'Update contracture if it has been changed
        If txtContracture.Text <> g_sOriginalContracture And g_NoElbowTape = False Then
            ''---ARMEUTIL.PR_DeleteByID(ARMEDDIA1.g_sID & "Contracture")
            PR_DeleteByXData(ARMEDDIA1.g_sID & "Contracture")
            PR_DrawContracture()
        End If

        'Add revised lining if it has been changed
        If txtLining.Text <> g_sOriginalLining And g_NoElbowTape = False Then
            PR_DrawLining()
        End If

        'Update Arm Box or Origin Marker (only if Profile has been changed)
        If ProfileChanged = True Then
            For ii = 0 To 17
                iLen = g_nLengths(ii) * 10 'Shift decimal place
                iMM = ARMEDDIA1.g_iMMs(ii)
                iRed = ARMEDDIA1.g_iRed(ii)
                iGms = ARMEDDIA1.g_iGms(ii)

                If iLen <> 0 Then
                    sLen = New String(" ", 3)
                    sLen = RSet(Trim(Str(iLen)), Len(sLen))
                    sMM = New String(" ", 3)
                    sMM = RSet(Trim(Str(iMM)), Len(sMM))
                    sRed = New String(" ", 3)
                    sRed = RSet(Trim(Str(iRed)), Len(sRed))
                    sGms = New String(" ", 3)
                    sGms = RSet(Trim(Str(iGms)), Len(sGms))
                Else
                    sLen = New String(" ", 3)
                    sMM = New String(" ", 3)
                    sRed = New String(" ", 3)
                    sGms = New String(" ", 3)
                End If

                sPackedLengths = sPackedLengths & sLen
                sPackedMMs = sPackedMMs & sMM
                sPackedRed = sPackedRed & sRed
                sPackedGms = sPackedGms & sGms

            Next ii


            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "hArm = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDArm.Text) & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "if (!hArm)Exit(%cancel," & ARMEDDIA1.QQ & "Can't find ARMBOX to Update" & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "TapeLengths" & ARMEDDIA1.QCQ & sPackedLengths & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "TapeMMs" & ARMEDDIA1.QCQ & sPackedMMs & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "Grams" & ARMEDDIA1.QCQ & sPackedGms & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "Reduction" & ARMEDDIA1.QCQ & sPackedRed & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "Contracture" & ARMEDDIA1.QCQ & txtContracture.Text & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "SetDBData( hArm, " & ARMEDDIA1.QQ & "Lining" & ARMEDDIA1.QCQ & txtLining.Text & ARMEDDIA1.QQ & ");")

        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "Execute (" & ARMEDDIA1.QQ & "menu" & ARMEDDIA1.QCQ & "ViewRedraw" & ARMEDDIA1.QQ & ");")

        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(ARMEDDIA1.fNum)

    End Sub

    Private Sub PR_DrawContracture()
        Dim xyContractTop, xyContractBott As ARMEDDIA1.XY
        Dim iElbowVertex As Short
        Dim nSeam, nContractureWidth, nProfileOffset As Double
        Dim sContracture As String
        'UPGRADE_WARNING: Arrays in structure Contracture may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim Contracture As ARMEDDIA1.curve

        If g_NoElbowTape Or txtContracture.Text = "" Then Exit Sub

        'Get elbow position on profile
        nSeam = 0.1875
        nProfileOffset = 0.5
        iElbowVertex = g_iVertexMap(g_iElbowTape)
        xyContractTop.X = xyProfile(iElbowVertex - 1).X
        xyContractTop.y = ((ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iElbowTape)) * (System.Math.Abs(ARMEDDIA1.g_iRed(g_iElbowTape) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
        xyContractTop.y = xyContractTop.y - nProfileOffset

        'Elbow position on the edge of the template at the fold line
        xyContractBott.X = xyContractTop.X
        xyContractBott.y = xyOtemplate.y

        'Contracture width
        sContracture = VB.Left(txtContracture.Text, 2)
        Select Case sContracture
            Case "No"
                'Just in case
                Exit Sub
            Case "10"
                nContractureWidth = 0.5
            Case "36"
                nContractureWidth = 1
            Case "71"
                nContractureWidth = 1.5
        End Select

        'Create contracture polyline
        Contracture.n = 5
        Dim X(5), Y(5) As Double
        Contracture.X = X
        Contracture.y = Y

        'Start at bottom
        Contracture.X(1) = xyContractBott.X
        Contracture.y(1) = xyContractBott.y

        'To the left
        Contracture.X(2) = xyContractBott.X - (nContractureWidth / 2)
        Contracture.y(2) = xyContractBott.y + ((xyContractTop.y - xyContractBott.y) / 2)

        'To the top
        Contracture.X(3) = xyContractTop.X
        Contracture.y(3) = xyContractTop.y

        'To the right
        Contracture.X(4) = xyContractBott.X + (nContractureWidth / 2)
        Contracture.y(4) = Contracture.y(2)

        'Close to the bottom
        Contracture.X(5) = xyContractBott.X
        Contracture.y(5) = xyContractBott.y

        'Draw contracture
        ARMDIA1.PR_SetLayer("Notes")
        ARMEDDIA1.PR_DrawPoly(Contracture)
        ''------------ARMEDDIA1.PR_AddEntityID(ARMEDDIA1.g_sID, "", "Contracture")
        '-------------ARMEDDIA1.PR_SetTextData(1, 32, -1, -1, -1)
        ARMEDDIA1.PR_DrawText("Remove For Contracture", xyContractTop, 0.125, 0, xyContractBott)

    End Sub

    Private Sub PR_DrawLining()
        Dim xyElbow As ARMEDDIA1.XY
        Dim iElbowVertex As Short
        Dim nProfileOffset, nSeam As Double

        If g_NoElbowTape Or txtLining.Text = "" Then Exit Sub

        'Get elbow position on profile
        nProfileOffset = 1
        nSeam = 0.1875
        iElbowVertex = g_iVertexMap(g_iElbowTape)
        xyElbow.X = xyProfile(iElbowVertex - 1).X
        xyElbow.y = ((ARMEDDIA1.fnDisplayToInches(g_nLengths(g_iElbowTape)) * (System.Math.Abs(ARMEDDIA1.g_iRed(g_iElbowTape) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
        xyElbow.y = xyElbow.y - nProfileOffset

        'Draw Lining Text
        ARMDIA1.PR_SetLayer("Notes")
        ''------------ARMEDDIA1.PR_SetTextData(2, 32, -1, -1, -1)
        Dim xyBase As ARMEDDIA1.XY
        ARMEDDIA1.PR_MakeXY(xyBase, xyElbow.X, xyOtemplate.y)
        ''ARMEDDIA1.PR_DrawText(txtLining.Text, xyElbow, 0.125, 0, xyBase)
        ARMEDDIA1.PR_DrawTextCenter(txtLining.Text, xyElbow, 0.125, 0, xyBase)

    End Sub

    Private Sub PR_DrawProfileEdits()
        'Create a macro to copy part of the arm profile and
        'modify this copy
        'Modifications will only be committed to the original profile on "Finish"
        'The exception to this is the ShortArm where due to DRAFIX bugs the edits
        'have to be done directly on the profile.

        Dim xyPt As ARMEDDIA1.XY
        Dim ii, iProfileVertex As Short
        Dim nSeam As Double

        nSeam = 0.1875

        If g_ShortArm = True Then txtUIDTempArm.Text = txtUIDCurve.Text

        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''ARMEDDIA1.fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Profile Edits", txtArm)
        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        ARMEDDIA1.fNum = FN_DrawOpen(sDrawFile, "EDIT", "Profile Edits", txtArm)
        PR_PutLine("HANDLE  hCurv;")

        If g_ReDrawn <> True And g_ShortArm = False Then
            'Make Copy of curve (First time through only)
            'Not for short arms (3 tapes)
            ARMDIA1.PR_SetLayer("Construct")
            PrintLine(ARMEDDIA1.fNum, "hCurv = AddEntity(" & ARMEDDIA1.QQ & "poly" & ARMEDDIA1.QCQ & "fitted" & ARMEDDIA1.QQ)

            For ii = 0 To g_iProfile
                Print(ARMEDDIA1.fNum, ARMEDDIA1.cc & xyProfile(ii).X & ARMEDDIA1.cc & xyProfile(ii).y)
            Next ii
            PrintLine(ARMEDDIA1.fNum, ");")
        Else
            'Modify Curve Copy
            If txtUIDTempArm.Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PrintLine(ARMEDDIA1.fNum, "hCurv = UID (" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QC & Val(txtUIDTempArm.Text) & ");")
            Else

            End If
        End If

        If g_sTempArmHandle <> "" Then
            PR_DeleteExistingEntity()
        End If

        'Loop through all tape values.  If changed then modify profile copy
        For ii = 0 To 17
            If g_iChanged(ii) > 0 Then
                g_iChanged(ii) = -1
                iProfileVertex = g_iVertexMap(ii)
                xyPt.X = xyProfile(iProfileVertex - 1).X
                ''xyPt.y = (((ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100))) / 2) + nSeam + xyOtemplate.y
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    xyPt.y = (((g_nLengths(ii) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100))) / 2) + nSeam + xyOtemplate.y
                Else
                    xyPt.y = (((ARMEDDIA1.fnDisplayToInches(g_nLengths(ii)) * (System.Math.Abs(ARMEDDIA1.g_iRed(ii) - 100) / 100))) / 2) + nSeam + xyOtemplate.y
                End If
                PrintLine(ARMEDDIA1.fNum, "SetVertex(hCurv," & iProfileVertex & "," & xyPt.X & "," & xyPt.y & ");")
                xyProfile(iProfileVertex - 1) = xyPt
            End If
        Next ii
        Dim ptColl As Point3dCollection = New Point3dCollection()
        For ii = 0 To g_iProfile
            ptColl.Add(New Point3d(xyProfile(ii).X, xyProfile(ii).y, 0))
        Next ii
        ARMEDDIA1.PR_DrawSpline(ptColl)

        If g_ReDrawn <> True And g_ShortArm = False Then
            PR_PutLine("HANDLE  hDDE;")

            'Poke Curve UID back to the VB program
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "hDDE = Open (" & ARMEDDIA1.QQ & "dde" & ARMEDDIA1.QCQ & "armeddia" & ARMEDDIA1.QCQ & "armeddia" & ARMEDDIA1.QQ & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMEDDIA1.fNum, "Poke ( hDDE, " & ARMEDDIA1.QQ & "txtUIDTempArm" & ARMEDDIA1.QC & "MakeString(" & ARMEDDIA1.QQ & "long" & ARMEDDIA1.QC & "UID(" & ARMEDDIA1.QQ & "get" & ARMEDDIA1.QQ & ",hCurv)));")

        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "Execute (" & ARMEDDIA1.QQ & "menu" & ARMEDDIA1.QCQ & "ViewRedraw" & ARMEDDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(ARMEDDIA1.fNum)

    End Sub

    Private Sub PR_GetProfileFromFile(ByRef sFileName As String)
        'Procedure to read curve data from file
        'Data is a series of x, y values.
        'In the following format
        '
        '    Line    Type
        '-----------------------------------------------
        '    1       Template Origin (xyOtemplate)
        '    2  }
        '    .  }}
        '    .  }}}  Vertices of profile (Profile)
        '    .  }}
        '    N  }
        '
        '
        'The editable vetices are setup by use of an array that maps the
        'vertex of the profile to the tapes.

        Dim iPalm, ii, iVertex As Short
        Dim fFileNum As Short
        Dim nTolerance As Double

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48)
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, VB.OpenMode.Input)

        'Get control points
        Input(fFileNum, xyOtemplate.X)
        Input(fFileNum, xyOtemplate.y)

        'Get profile points
        g_iProfile = 0
        Do While Not EOF(fFileNum)
            Input(fFileNum, xyProfile(g_iProfile).X)
            Input(fFileNum, xyProfile(g_iProfile).y)
            'Store original profile for use in cancel for short arms
            xyOriginal(g_iProfile).X = xyProfile(g_iProfile).X
            xyOriginal(g_iProfile).y = xyProfile(g_iProfile).y
            g_iProfile = g_iProfile + 1
        Loop
        g_iProfile = g_iProfile - 1
        FileClose(fFileNum)

        'If a gauntlet has been given then the user may have added additional
        'Vertices to smooth the flow of the curve into the gaunlet end.  These
        'will be between the WristTape and the PalmTape.
        '
        'For editing purposes we must find the profile vertex that maps to
        'a tape. We therefor create an array that contains the profile vertex
        'number for each tape
        '
        If g_Gauntlet Then
            'One to One mapping up to Palm
            iVertex = 1
            ii = ARMEDDIA1.g_iFirstTape
            Do
                g_iVertexMap(ii) = iVertex
                iVertex = iVertex + 1
                ii = ii + 1
            Loop Until ii > g_iPalmNo
            iPalm = iVertex - 1

            'Skip Vertex between palm and wrist
            'NB Use of tolerance
            nTolerance = 0.01
            For ii = iPalm To g_iProfile
                If xyProfile(ii).X > xyProfile(iPalm - 1).X + (g_nPalmWristDist + nTolerance) Then Exit For
            Next ii
            iVertex = ii

            'One to One mapping from wrist onwards
            For ii = g_iWristNo To ARMEDDIA1.g_iLastTape
                g_iVertexMap(ii) = iVertex
                iVertex = iVertex + 1
            Next ii

        Else
            'For ordinary arms we assume a 1 to 1 mapping
            iVertex = 1
            For ii = ARMEDDIA1.g_iFirstTape To ARMEDDIA1.g_iLastTape
                g_iVertexMap(ii) = iVertex
                iVertex = iVertex + 1
            Next ii
            'Check to see if the user has inserted extra points
            'Warn and exit, as this will cause errors
            If (ARMEDDIA1.g_iLastTape - ARMEDDIA1.g_iFirstTape) > g_iProfile Then
                MsgBox("There are missing points in the Plain Arm profile.  The editor will not work!  Redraw arm or insert missing point", 16, "Arm EDIT")
                Return
            End If
            If (ARMEDDIA1.g_iLastTape - ARMEDDIA1.g_iFirstTape) < g_iProfile Then
                MsgBox("There are extra points in the Plain Arm profile.  The editor will not work!  Remove extra points and try again", 16, "Arm EDIT")
                Return
            End If
        End If



        'DRAFIX only allows the creation by a macro of an Open Fitted curve if the curve
        'has more than 3 points.
        'Therefor the method of duplicating the curve to display the results is
        'not available.  Edits will be applied directly to the curve
        If g_iProfile = 2 Then g_ShortArm = True


    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    ARMEDDIA1.fNum is global variable
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, sLine)
    End Sub

    Private Sub PR_SetLayer(ByRef sNewLayer As String)
        'To the DRAFIX macro file (given by the global ARMEDDIA1.fNum).
        'Write the syntax to set the current LAYER.
        'For this to work it assumes that hLayer is defined in DRAFIX as
        'a HANDLE.
        '
        'Note:-
        '    ARMEDDIA1.fNum, CC, QQ, NL, g_sCurrentLayer are globals initialised by FN_Open
        '
        'To reduce unessesary writing of DRAFIX code check that the new layer
        'is different from the Current layer, change only if it is different.
        '

        If ARMEDDIA1.g_sCurrentLayer = sNewLayer Then Exit Sub
        ARMEDDIA1.g_sCurrentLayer = sNewLayer

        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "hLayer = Table(" & ARMEDDIA1.QQ & "find" & ARMEDDIA1.QCQ & "layer" & ARMEDDIA1.QCQ & sNewLayer & ARMEDDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object ARMEDDIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMEDDIA1.fNum, "if ( hLayer != %badtable)" & "Execute (" & ARMEDDIA1.QQ & "menu" & ARMEDDIA1.QCQ & "SetLayer" & ARMEDDIA1.QC & "hLayer);")

    End Sub
    Function PR_GetArmProfile() As Boolean
        If readArmEditInfo() = True Then
            If (System.IO.File.Exists(fnGetSettingsPath("LookupTables") + "\\ARMCURVE.DAT")) Then
                Return True
            End If
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Arm Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim strProfile As String = ""
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim bIsReadCurveInfo As Boolean = False
        While (True)
            Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
            If ptEntRes.Status <> PromptStatus.OK Then
                MsgBox("No Arm Profile selected", 16, "Arm Edit")
                Return False
            End If
            Try
                Dim idSpline As ObjectId = ptEntRes.ObjectId
                Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
                    Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ProfileID")
                    If Not IsNothing(rb) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rb
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strProfile = strProfile & typeVal.Value
                        Next
                    End If
                    Dim rbEditInfo As ResultBuffer = acEnt.GetXDataForApplication("EditInfo")
                    If Not IsNothing(rbEditInfo) Then
                        ' Get the values in the xdata
                        Dim strEdifInfo As String = ""
                        For Each typeVal As TypedValue In rbEditInfo
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            If strEdifInfo <> "" Then
                                strEdifInfo = strEdifInfo & "," & typeVal.Value
                            Else
                                strEdifInfo = typeVal.Value
                            End If
                        Next
                        If strEdifInfo <> "" Then
                            ''txtTapeLengths.Text = FN_GetString(strEdifInfo, 1)
                            Dim ii As Short
                            For ii = 1 To 18
                                If ii = 1 Then
                                    txtTapeLengths.Text = FN_GetString(strEdifInfo, ii)
                                Else
                                    txtTapeLengths.Text = txtTapeLengths.Text & "," & FN_GetString(strEdifInfo, ii)
                                End If
                            Next ii
                            txtFabric.Text = FN_GetString(strEdifInfo, 19)
                            txtTapeMMs.Text = FN_GetString(strEdifInfo, 20)
                            txtGrams.Text = FN_GetString(strEdifInfo, 21)
                            txtReduction.Text = FN_GetString(strEdifInfo, 22)
                            cboContracture.Text = FN_GetString(strEdifInfo, 23)
                            cboLining.Text = FN_GetString(strEdifInfo, 24)
                            txtGauntlet.Text = FN_GetString(strEdifInfo, 25)
                            bIsReadCurveInfo = True
                        End If
                    End If
                    Dim fitPtsCount As Integer = acEnt.NumFitPoints
                    Dim i As Integer = 0
                    While (i < fitPtsCount)
                        ptCurveColl.Add(acEnt.GetFitPointAt(i))
                        i = i + 1
                    End While
                    g_sOriginalCurveHandle = acEnt.Handle.ToString()
                    tr.Commit()
                End Using
                Exit While
            Catch ex As Exception
                MsgBox("An ARM Profile was not selected" & Chr(10), 16, "Arm Edit")
                Return False
            End Try
        End While
        strProfile = Trim(strProfile)
        Dim sStyle, sArm As String
        sStyle = ""
        Dim nLen As Integer = 0
        Dim nArmType As Integer = 0
        If strProfile.Contains("ARM") Then
            nLen = 3
            nArmType = 1
        ElseIf strProfile.Contains("VEST") Then
            nLen = 4
            nArmType = 2
        ElseIf strProfile.Contains("BODY") Then
            nLen = 4
            nArmType = 3
        End If

        If strProfile.Contains("LeftProfile") Then
            ''nLen = nLen + 11
            sStyle = Mid(strProfile, nLen + 1, strProfile.Length - (nLen + 11))
            sArm = "Left"
        ElseIf strProfile.Contains("RightProfile") Then
            ''nLen = nLen + 12
            sStyle = Mid(strProfile, nLen + 1, strProfile.Length - (nLen + 12))
            sArm = "Right"
        Else
            MsgBox("An ARM Profile was not selected" & Chr(10), 16, "Arm Edit")
            Return False
        End If
        txtID.Text = sStyle
        txtArm.Text = sArm
        Dim filterType(0) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
        If selResult.Status <> PromptStatus.OK Then
            Return False
        End If

        Dim selectionSet As SelectionSet = selResult.Value
        Dim nMarkerCount As Integer = 0
        Dim xyInsertion As ARMEDDIA1.XY
        Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = acTr.GetObject(idObject, OpenMode.ForRead)
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("ProfileID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Contains(strProfile) Then
                    nMarkerCount = nMarkerCount + 1
                    Dim position As Point3d = blkXMarker.Position
                    ARMEDDIA1.PR_MakeXY(xyInsertion, position.X, position.Y)
                End If
            Next
            acTr.Commit()
        End Using
        If nMarkerCount < 1 Then
            MsgBox("Missing origin marker for selected ARM, data not found!" & Chr(10), 16, "Arm Edit")
            Return False
        End If
        If nMarkerCount > 1 Then
            MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again." & Chr(10), 16, "Arm Edit")
            Return False
        End If
        ''// Write data to file for subsequent use in editor
        Dim fileNum As Object
        fileNum = FreeFile()
        If (Not System.IO.File.Exists(fnGetSettingsPath("LookupTables") + "\\ARMCURVE.DAT")) Then
            FileOpen(fileNum, fnGetSettingsPath("LookupTables") + "\\ARMCURVE.DAT", VB.OpenMode.Append)
        Else
            FileOpen(fileNum, fnGetSettingsPath("LookupTables") + "\\ARMCURVE.DAT", VB.OpenMode.Output)
        End If
        ''---SetData("UnitLinearType", 0);	// "Inches"
        PrintLine(fileNum, Str(xyInsertion.X) & " " & Str(xyInsertion.y) & Chr(10))
        Dim nCount As Integer = 0
        Dim xyPt As ARMEDDIA1.XY
        While (nCount < ptCurveColl.Count)
            ARMEDDIA1.PR_MakeXY(xyPt, ptCurveColl(nCount).X, ptCurveColl(nCount).Y)
            PrintLine(fileNum, Str(xyPt.X) & " " & Str(xyPt.y) & Chr(10))
            nCount = nCount + 1
        End While
        FileClose(fileNum)
        If bIsReadCurveInfo = False Then
            readDWGInfo(nArmType, sArm)
        End If
        Return True
    End Function
    Private Function readArmEditInfo() As Boolean
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("ArmEditInfo", "ARMEDITDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                txtTapeLengths.Text = ""
                txtTapeMMs.Text = ""
                txtGrams.Text = ""
                txtReduction.Text = ""
                Dim ii As Double
                For ii = 0 To 17
                    Dim strTapeLen As String = arr(ii).Value
                    g_strTapeLen(ii) = strTapeLen
                    'If (Val(strTapeLen) * 10) < 100 Then
                    '    strTapeLen = " " & arr(ii).Value
                    'End If
                    'If arr(ii).Value = "" Then
                    '    strTapeLen = strTapeLen & " "
                    'End If
                    'If Val(strTapeLen) < 100 Then
                    '    strTapeLen = strTapeLen & " "
                    'End If
                    If ii = 0 Then
                        txtTapeLengths.Text = txtTapeLengths.Text & strTapeLen
                        Continue For
                    End If
                    txtTapeLengths.Text = txtTapeLengths.Text & "," & strTapeLen
                Next
                txtFabric.Text = arr(18).Value
                For ii = 19 To 36
                    Dim strTapeMMs As String = arr(ii).Value
                    If (Val(strTapeMMs) * 10) < 100 Then
                        strTapeMMs = " " & arr(ii).Value
                    End If
                    If arr(ii).Value = "" Then
                        strTapeMMs = strTapeMMs & " "
                    End If
                    If Val(strTapeMMs) < 100 Then
                        strTapeMMs = strTapeMMs & " "
                    End If
                    txtTapeMMs.Text = txtTapeMMs.Text & strTapeMMs
                Next
                For ii = 37 To 54
                    Dim strTapeGrams As String = arr(ii).Value
                    If (Val(strTapeGrams) * 10) < 100 Then
                        strTapeGrams = " " & arr(ii).Value
                    End If
                    If arr(ii).Value = "" Then
                        strTapeGrams = strTapeGrams & " "
                    End If
                    If Val(strTapeGrams) < 100 Then
                        strTapeGrams = strTapeGrams & " "
                    End If
                    txtGrams.Text = txtGrams.Text & strTapeGrams
                Next
                For ii = 55 To 72
                    Dim strTapeRed As String = arr(ii).Value
                    If (Val(strTapeRed) * 10) < 100 Then
                        strTapeRed = " " & arr(ii).Value
                    End If
                    If arr(ii).Value = "" Then
                        strTapeRed = strTapeRed & " "
                    End If
                    If Val(strTapeRed) < 100 Then
                        strTapeRed = strTapeRed & " "
                    End If
                    txtReduction.Text = txtReduction.Text & strTapeRed
                Next
                txtGauntlet.Text = arr(73).Value
                g_ReDrawn = arr(74).Value
                g_sTempArmHandle = arr(75).Value
                txtID.Text = arr(76).Value
                txtArm.Text = arr(77).Value
                g_sOriginalCurveHandle = arr(78).Value
                cboContracture.Text = arr(79).Value
                ''Dim strCont As String = arr(79).Value
                ''cboContracture.Text = strCont
                cboLining.Text = arr(80).Value
                Dim ln As Long = Convert.ToInt64(g_sTempArmHandle, 16)
                ' Not create a Handle from the long integer
                Dim hn As Handle = New Handle(ln)
                ' And attempt to get an ObjectId for the Handle
                '' Get the current document and database
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Dim acCurDb As Database = acDoc.Database
                Dim id As ObjectId = acCurDb.GetObjectId(False, hn, 0)
                If id.IsValid <> True Or id.IsErased = True Then
                    g_sTempArmHandle = ""
                    MsgBox("Can't find Arm Edit Details" & Chr(10) & "Please select another Arm Profile", 16, "Arm Details")
                    Return False
                End If
                If g_sOriginalCurveHandle <> "" Then
                    Dim lnOriginal As Long = Convert.ToInt64(g_sOriginalCurveHandle, 16)
                    ' Not create a Handle from the long integer
                    Dim hnOriginal As Handle = New Handle(lnOriginal)
                    Dim idOriginal As ObjectId = acCurDb.GetObjectId(False, hnOriginal, 0)
                    If idOriginal.IsValid <> True Or idOriginal.IsErased = True Then
                        g_sOriginalCurveHandle = ""
                        MsgBox("Can't find Arm Edit Details" & Chr(10) & "Please select another Arm Profile", 16, "Arm Details")
                        Return False
                    End If
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Private Sub readDWGInfo(ByRef nArmType As Integer, ByRef sSide As String)
        Dim _sClass As New SurroundingClass()
        Dim resbuf As New ResultBuffer
        Dim sInfo As String = ""
        Dim sDict As String = ""
        Select Case nArmType
            Case 1 'For ARM
                If sSide.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase) Then
                    sInfo = "SNInfo"
                    sDict = "SNDIC"
                Else
                    sInfo = "SNRightInfo"
                    sDict = "SNRIGHTDIC"
                End If
            Case 2 'For Vest - ARM
                sInfo = "VestArmRightInfo"
                sDict = "VESTARMRIGHTDIC"
                If sSide.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase) Then
                    sInfo = "VestArmInfo"
                    sDict = "VESTARMDIC"
                End If
            Case 3 'For Body - Arm
                sInfo = "BodyArmRightInfo"
                sDict = "BODYARMRIGHTDIC"
                If sSide.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase) Then
                    sInfo = "BodyArmInfo"
                    sDict = "BODYARMDIC"
                End If
        End Select
        resbuf = _sClass.GetXrecord(sInfo, sDict)
        If (resbuf IsNot Nothing) Then
            Dim arr() As TypedValue = resbuf.AsArray()
            txtTapeLengths.Text = ""
            txtTapeMMs.Text = ""
            txtGrams.Text = ""
            txtReduction.Text = ""
            Dim ii As Double
            For ii = 0 To 17
                Dim strTapeLen As String = arr(ii).Value
                g_strTapeLen(ii) = strTapeLen
                'Dim nLen As Integer = strTapeLen.Length()
                'If (Val(strTapeLen) * 10) < 100 Then
                '    strTapeLen = " " & arr(ii).Value
                'End If
                'If arr(ii).Value = "" Then
                '    strTapeLen = strTapeLen & " "
                'End If
                'If Val(strTapeLen) < 100 Then
                '    ''If Val(strTapeLen) < 10 Then
                '    If strTapeLen.Contains(".") = False Then
                '        strTapeLen = strTapeLen & " "
                '    End If
                'End If
                If ii = 0 Then
                    txtTapeLengths.Text = txtTapeLengths.Text & strTapeLen
                    Continue For
                End If
                txtTapeLengths.Text = txtTapeLengths.Text & "," & strTapeLen
            Next ii
            txtFabric.Text = arr(18).Value
            For ii = 20 To 37
                Dim strTapeMMs As String = arr(ii).Value
                If (Val(strTapeMMs) * 10) < 100 Then
                    strTapeMMs = " " & arr(ii).Value
                End If
                If arr(ii).Value = "" Then
                    strTapeMMs = strTapeMMs & " "
                End If
                If Val(strTapeMMs) < 100 Then
                    strTapeMMs = strTapeMMs & " "
                End If
                txtTapeMMs.Text = txtTapeMMs.Text & strTapeMMs
            Next ii
            For ii = 38 To 55
                Dim strTapeGrams As String = arr(ii).Value
                If (Val(strTapeGrams) * 10) < 100 Then
                    strTapeGrams = " " & arr(ii).Value
                End If
                If arr(ii).Value = "" Then
                    strTapeGrams = strTapeGrams & " "
                End If
                If Val(strTapeGrams) < 100 Then
                    strTapeGrams = strTapeGrams & " "
                End If
                txtGrams.Text = txtGrams.Text & strTapeGrams
            Next ii
            For ii = 56 To 73
                Dim strTapeRed As String = arr(ii).Value
                If (Val(strTapeRed) * 10) < 100 Then
                    strTapeRed = " " & arr(ii).Value
                End If
                If arr(ii).Value = "" Then
                    strTapeRed = strTapeRed & " "
                End If
                If Val(strTapeRed) < 100 Then
                    strTapeRed = strTapeRed & " "
                End If
                txtReduction.Text = txtReduction.Text & strTapeRed
            Next ii
            If sSide.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase) Then
                txtGauntlet.Text = arr(75).Value
            Else
                Dim sGauntlet As String = "0"
                If arr(77).Value = True Then sGauntlet = "1"
                Dim sPalmNo As String = arr(78).Value
                Dim sWristNo As String = arr(79).Value
                Dim sPalmWrist As String = arr(80).Value
                Dim sCircum As String = arr(81).Value
                Dim sThumb As String = arr(82).Value
                Dim sGauntletExt As String = arr(83).Value
                Dim sDetachable As String = "0"
                If arr(84).Value = True Then sDetachable = "1"
                Dim sNoThumb As String = "0"
                If arr(85).Value = True Then sNoThumb = "1"
                Dim sEnclosed As String = "0"
                If arr(86).Value = True Then sEnclosed = "1"
                txtGauntlet.Text = sGauntlet & " " & sEnclosed & " " & sDetachable & " " & sNoThumb & " " & sWristNo & " " &
                        sPalmNo & " " & sThumb & " " & sCircum & " " & sPalmWrist & " " & sGauntletExt
            End If
        End If
    End Sub
    Private Sub SaveArmEditInfo()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("ArmEditInfo", "ARMEDITDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("ArmEditInfo", "ARMEDITDIC")
            End If

            Dim resbuf As New ResultBuffer
            Dim ii As Double
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Trim(EditTapes(ii).Text)))
            Next
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFabric.Text))
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Trim(EditMMs(ii).Text)))
            Next
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Trim(EditGrams(ii).Text)))
            Next
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Trim(EditReductions(ii).Text)))
            Next
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtGauntlet.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), g_ReDrawn))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), g_sTempArmHandle))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtID.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtArm.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), g_sOriginalCurveHandle))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboContracture.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboLining.Text))
            _sClass.SetXrecord(resbuf, "ArmEditInfo", "ARMEDITDIC")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub PR_DeleteExistingEntity()
        Dim ln As Long = Convert.ToInt64(g_sTempArmHandle, 16)
        ' Not create a Handle from the long integer
        Dim hn As Handle = New Handle(ln)
        ' And attempt to get an ObjectId for the Handle
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim id As ObjectId = acCurDb.GetObjectId(False, hn, 0)
        If id.IsValid <> True Then
            Exit Sub
        End If
        ' Finally let's open the object and erase it
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim obj As DBObject = acTrans.GetObject(id, OpenMode.ForWrite)
            obj.Erase()
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_RedrawOriginalCurve(ByRef sCurveHandle As String, Optional ByRef bIsSetXData As Boolean = False)
        Dim ln As Long = Convert.ToInt64(sCurveHandle, 16)
        ' Not create a Handle from the long integer
        Dim hn As Handle = New Handle(ln)
        ' And attempt to get an ObjectId for the Handle
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim id As ObjectId = acCurDb.GetObjectId(False, hn, 0)
        If id.IsValid <> True Then
            Exit Sub
        End If
        ' Finally let's open the object and erase it
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acSpline As Spline = acTrans.GetObject(id, OpenMode.ForWrite)
            Dim fitPtsCount As Integer = acSpline.NumFitPoints
            Dim i As Integer = 0
            While (i < fitPtsCount)
                acSpline.SetFitPointAt(i, New Point3d(xyProfile(i).X, xyProfile(i).y, 0))
                i = i + 1
            End While
            If bIsSetXData = True Then
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim appName As String = "EditInfo"
                If acRegAppTbl.Has(appName) = False Then
                    Dim acRegAppTblRec As RegAppTableRecord
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                txtTapeLengths.Text = ""
                txtTapeMMs.Text = ""
                txtGrams.Text = ""
                txtReduction.Text = ""
                Dim ii As Double
                For ii = 0 To 17
                    Dim strTapeLen As String = Trim(EditTapes(ii).Text)
                    g_strTapeLen(ii) = strTapeLen
                    'If (Val(strTapeLen) * 10) < 100 Then
                    '    strTapeLen = " " & EditTapes(ii).Text
                    'End If
                    'If EditTapes(ii).Text = "" Then
                    '    strTapeLen = strTapeLen & " "
                    'End If
                    'If Val(strTapeLen) < 100 Then
                    '    ''strTapeLen = strTapeLen & " "
                    '    If strTapeLen.Contains(".") = False Then
                    '        strTapeLen = strTapeLen & " "
                    '    End If
                    'End If
                    If ii = 0 Then
                        txtTapeLengths.Text = txtTapeLengths.Text & strTapeLen
                        Continue For
                    End If
                    txtTapeLengths.Text = txtTapeLengths.Text & "," & strTapeLen
                Next
                For ii = 0 To 17
                    Dim strTapeMMs As String = Trim(EditMMs(ii).Text)
                    If (Val(strTapeMMs) * 10) < 100 Then
                        strTapeMMs = " " & EditMMs(ii).Text
                    End If
                    If EditMMs(ii).Text = "" Then
                        strTapeMMs = strTapeMMs & " "
                    End If
                    If Val(strTapeMMs) < 100 Then
                        strTapeMMs = strTapeMMs & " "
                    End If
                    txtTapeMMs.Text = txtTapeMMs.Text & strTapeMMs
                Next
                For ii = 0 To 17
                    Dim strTapeGrams As String = Trim(EditGrams(ii).Text)
                    If (Val(strTapeGrams) * 10) < 100 Then
                        strTapeGrams = " " & EditGrams(ii).Text
                    End If
                    If EditGrams(ii).Text = "" Then
                        strTapeGrams = strTapeGrams & " "
                    End If
                    If Val(strTapeGrams) < 100 Then
                        strTapeGrams = strTapeGrams & " "
                    End If
                    txtGrams.Text = txtGrams.Text & strTapeGrams
                Next
                For ii = 0 To 17
                    Dim strTapeRed As String = Trim(EditReductions(ii).Text)
                    If (Val(strTapeRed) * 10) < 100 Then
                        strTapeRed = " " & EditReductions(ii).Text
                    End If
                    If EditReductions(ii).Text = "" Then
                        strTapeRed = strTapeRed & " "
                    End If
                    If Val(strTapeRed) < 100 Then
                        strTapeRed = strTapeRed & " "
                    End If
                    txtReduction.Text = txtReduction.Text & strTapeRed
                Next
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtTapeLengths.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtFabric.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtTapeMMs.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtGrams.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtReduction.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, cboContracture.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, cboLining.Text))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, txtGauntlet.Text))
                    acSpline.XData = rb
                End Using
            End If
            acTrans.Commit()
        End Using
    End Sub
    Private Function FN_GetString(ByVal sString As String, ByRef iIndex As Short) As String
        Dim ii, iPos As Short
        'Initial error checking
        ''sString = Trim(sString) 'Remove leading and trailing blanks
        If Len(sString) = 0 Then
            FN_GetString = ""
            Exit Function
        End If

        'Prepare string
        sString = sString & ","
        'Get iIndexth item
        For ii = 1 To iIndex
            iPos = InStr(sString, ",")
            If ii = iIndex Then
                sString = VB.Left(sString, iPos - 1)
                FN_GetString = sString
                Exit Function
            Else
                sString = Mid(sString, iPos + 1)
                If Len(sString) = 0 Then
                    FN_GetString = ""
                    Exit Function
                End If
            End If
        Next ii
        FN_GetString = ""
    End Function
    Private Sub PR_DeleteByXData(ByRef XDataName As String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim filterType(1) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.Start, "Polyline"), 0)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, XDataName), 1)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
        If selResult.Status <> PromptStatus.OK Then
            Exit Sub
        End If
        Dim selectionSet As SelectionSet = selResult.Value
        Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim dbObj As DBObject = acTr.GetObject(idObject, OpenMode.ForWrite)
                dbObj.Erase()
            Next
            acTr.Commit()
        End Using
    End Sub
    Sub PR_DrawTapeLabel(ByRef nTape As Object, ByRef xyStart As ARMEDDIA1.XY, ByRef nLength As Object, ByRef nMM As Object, ByRef nGrm As Object, ByRef nRed As Object)
        Dim nDec As Object
        Dim nInt As Object
        Dim sSymbol As String
        Dim xyPt As ARMEDDIA1.XY
        Dim nSymbolOffSet, nTextHt As Single

        sSymbol = nTape & "tape"
        nSymbolOffSet = 0.6877
        nTextHt = 0.125

        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X, xyStart.y + nSymbolOffSet)
        ''-------Dim slayer As String = ARMEDDIA1.g_sFileNo & Mid(ARMEDDIA1.g_sSide, 1, 1) & nTape
        Dim slayer As String = "TEMPLATE" & ARMEDDIA1.g_sSide
        ARMDIA1.PR_SetLayer(slayer)

        nInt = Int(CDbl(nLength))
        nDec = ARMDIA1.round((nLength - nInt) / 0.125)
        If nDec = 8 Then
            nDec = 0
            nInt = nInt + 1
        End If

        'Draw Integer part
        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 0.75)
        ARMEDDIA1.PR_DrawText(Trim(nInt), xyPt, nTextHt, 0, xyStart)
        'Draw eights part
        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X + 0.0625 + (Len(Trim(nInt)) * nTextHt * 0.8), xyStart.y + 0.75 + nTextHt / 1.5)
        If nDec <> 0 Then ARMEDDIA1.PR_DrawText(Trim(nDec), xyPt, nTextHt / 1.5, 0, xyStart)

        'MMs text
        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1)
        ''----------ARMEDDIA1.PR_DrawText(DirectCast(nMM, TextBox).Text & "mm", xyPt, nTextHt, 0)
        ARMEDDIA1.PR_DrawText(Trim(Str(nMM)) & "mm", xyPt, nTextHt, 0, xyStart)

        'Grams text
        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1.25)
        ''-----------ARMEDDIA1.PR_DrawText(DirectCast(nGrm, Label).Text & "gm", xyPt, nTextHt, 0)
        ARMEDDIA1.PR_DrawText(Trim(Str(nGrm)) & "gm", xyPt, nTextHt, 0, xyStart)

        'Reduction text and circle round the text
        ARMEDDIA1.PR_MakeXY(xyPt, xyStart.X + 0.25, xyStart.y + 1.625)
        ''-----------ARMEDDIA1.PR_DrawTextCenter(Trim(DirectCast(nRed, Label).Text), xyPt, nTextHt, 0)
        ARMEDDIA1.PR_DrawTextCenter(Trim(Str(nRed)), xyPt, nTextHt, 0, xyStart)
        ARMEDDIA1.PR_DrawCircle(xyPt, 0.125, xyStart)
    End Sub
End Class
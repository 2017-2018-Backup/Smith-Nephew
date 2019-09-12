Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Friend Class WHRightLeg
    Inherits System.Windows.Forms.Form
    'Constants
    Const HEEL_TOL As Short = 9 '9" heel
    Public g_sChangeChecker As String
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
    Public g_sTxtLeg As String
    Public g_sTxtData As String
    Public g_sTxtFirstTape As String
    Public g_sTxtLastTape As String
    Public g_sTxtAnkleTape As String
    Public g_sTxtTapeLengths As String
    Public g_sTxtToeStyle As String


    'MsgBox constants
    Const IDYES = 6
    Const IDNO = 7
    Const IDCANCEL = 2

    Structure XY
        Dim X As Double
        Dim Y As Double
    End Structure
    Dim xyOtemplate As XY
    Public Const PI As Double = 3.141592654
    Dim nFabricClass As Double

    Dim sReductionAnkle As String
    Dim sGramsAnkle As String
    Dim sMMAnkle As String
    Dim sWorkOrder As String
    Dim sFabric As String
    Dim sPressure As String
    Dim g_sLegStyle As String
    Dim sTOSCir As String
    Public sTapeLengths As String
    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        Dim Response As Short
        Dim sTask, sCurrentValues As String

        'Check if data has been modified
        sCurrentValues = FN_ValuesString()

        If sCurrentValues <> g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "Waist Leg- Dialogue")
            Select Case Response
                Case IDYES
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    Me.Close()
                Case IDNO
                    Me.Close()
                Case IDCANCEL
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
        WaistLeg.WHLeg.Close()
        WaistMain.WaistMainDlg.Close()
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
        nValue = WHRightLeg1.fnDisplayToInches(Val(Text1(ii).Text))

        'Extend tape
        If txtSex.Text = "Male" Then
            nValue = nValue + 0.5
        Else
            nValue = nValue + 1
        End If

        Text1(ii + 1).Text = CStr(WHRightLeg1.fnInchesToDisplay(nValue))
        labTapeInInches(ii + 1).Text = WHRightLeg1.fnInchesToText(nValue)

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
            nValue = WHRightLeg1.fnDisplayToInches(Val(Text1(ii - 1).Text))
            labTapeInInches(ii).Text = WHRightLeg1.fnInchesToText(nValue)
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
            nValue = WHRightLeg1.fnDisplayToInches(Val(Text1(ii).Text))
            labTapeInInches(ii - 1).Text = WHRightLeg1.fnInchesToText(nValue)
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
        fNum = FreeFile()
        FileOpen(fNum, sDrafixFile, VB.OpenMode.Output)
        FN_Open = fNum

        'Write header information etc. to the DRAFIX macro file
        PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(fNum, "//Patient - " & sName & ", " & sFileNo & "")
        PrintLine(fNum, "//by Visual Basic, Waist Height Leg")
        PrintLine(fNum, "//type - " & sType & "")

        'Clear user selections etc
        PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
        PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
        PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")

        'Display Hour Glass symbol
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
        For ii = 0 To 29
            sString = sString & Str(Val(Text1(ii).Text))
        Next ii

        sString = sString & txtFootPleat1.Text & txtFootPleat2.Text

        sString = sString & txtTopLegPleat1.Text & txtTopLegPleat2.Text & txtFootLength.Text

        sString = sString & cboToeStyle.Text & Str(optAboveKnee.Checked) & Str(optBelowKnee.Checked)

        FN_ValuesString = sString
    End Function
    Private Sub Form_LinkClose()

        Dim nValue As Double
        Dim nAge As Object
        Dim ii As Short

        'Disable Time out
        ''Timer1.Enabled = False

        'Units
        If txtUnits.Text = "cm" Then
            WHRightLeg1.g_nUnitsFac = 10 / 25.4
        Else
            WHRightLeg1.g_nUnitsFac = 1
        End If

        'Foot Length
        nValue = WHRightLeg1.fnDisplayToInches(Val(txtFootLength.Text))
        If nValue > 0 Then labFootLengthInInches.Text = WHRightLeg1.fnInchesToText(nValue)

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
            If VB6.GetItemString(cboToeStyle, ii) = g_sTxtToeStyle Then
                cboToeStyle.SelectedIndex = ii
            End If
        Next ii


        'Leg Option buttons
        'If g_sTxtLeg = "Left" Then
        '    optLeftLeg.Checked = True
        '    optRightLeg.Enabled = False
        'Else
        '    optRightLeg.Checked = True
        '    optLeftLeg.Enabled = False
        'End If

        'Update tape boxes
        For ii = 0 To 29
            nValue = Val(Mid(g_sTxtTapeLengths, (ii * 4) + 1, 4)) / 10
            If nValue > 0 Then
                Text1(ii).Text = CStr(nValue)
                nValue = WHRightLeg1.fnDisplayToInches(nValue)
                labTapeInInches(ii).Text = WHRightLeg1.fnInchesToText(nValue)
            End If
        Next ii

        'Store original Heel and Ankle tape values
        'Used on DDE_Update check if values have been modified
        'A heel length of 0 => footless

        g_nAnkleTape = 0
        g_nAnkleLen = 0
        g_nHeelLen = WHRightLeg1.fnDisplayToInches(Val(Text1(5).Text))

        If (g_nHeelLen <> 0) And (g_nHeelLen < HEEL_TOL) Then
            g_nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
            g_nAnkleLen = WHRightLeg1.fnDisplayToInches(Val(Text1(g_nAnkleTape).Text))
        Else
            g_nAnkleTape = 7 'Large heel Ankle at +3 tape
            g_nAnkleLen = WHRightLeg1.fnDisplayToInches(Val(Text1(g_nAnkleTape).Text))
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
            Select Case Val(g_sTxtData)
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
    Private Sub whlegdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Hide()
        'Maintain while loading DDE data
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
        'Reset in Form_LinkClose

        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        WHRightLeg1.MainForm = Me

        'Initialize globals
        QQ = Chr(34) 'Double quotes (")
        NL = Chr(13) 'New Line
        CC = Chr(44) 'The comma (,)
        QCQ = QQ & CC & QQ
        QC = QQ & CC
        CQ = CC & QQ

        g_sPathJOBST = fnPathJOBST()

        WHRightLeg1.g_nUnitsFac = 1 'Default to inches
        g_sTxtLeg = "Right"
        g_sTxtData = ""
        g_sTxtFirstTape = ""
        g_sTxtLastTape = ""
        g_sTxtAnkleTape = ""
        g_sTxtTapeLengths = ""
        g_sTxtToeStyle = ""

        'Start a timer that will ensure that the
        'the VB programme ends if the DRAFIX DDE
        'link fails.
        ''Timer1.Interval = 6000
        ''Timer1.Enabled = True

        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("No Patient Details have been found in drawing!", 16, "Waist Leg - Dialogue")
            Me.Close()
            Exit Sub
        End If
        obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
        txtFileNo.Text = fileNo
        txtPatientName.Text = patient
        txtDiagnosis.Text = diagnosis
        txtAge.Text = age
        txtSex.Text = sex
        txtUnits.Text = units
        sWorkOrder = workOrder
        Form_LinkClose()
        readDWGInfo()
    End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
        'Check that data is all present and insert into drafix
        Dim sTask As String

        If Validate_Data() Then
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Hide()
            WaistLeg.WHLeg.Hide()
            WaistMain.WaistMainDlg.Hide()

            ''PR_CreateMacro_Save("c:\jobst\draw.d")
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            PR_CreateMacro_Save(sDrawFile)
            PR_DrawWaistLeg()
            saveInfoToDWG()
            'sTask = fnGetDrafixWindowTitleText()
            'If sTask <> "" Then
            '    AppActivate(sTask)
            '    System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
            '    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            '    Return
            'Else
            '    MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Leg Data")
            'End If
        End If
        OK.Enabled = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        WaistLeg.WHLeg.Close()
        WaistMain.g_bIsNeedLoad = True
        WaistMain.WaistMainDlg.Show()
        Me.Show()
    End Sub
    Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
        'fNum is a global variable use in subsequent procedures
        fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

        'If this is a new drawing of a vest then Define the DATA Base
        'fields for the VEST Body and insert the BODYBOX symbol
        PR_PutLine("HANDLE hMPD, hLeg;")

        Update_DDE_Text_Boxes()

        Dim sSymbol As String

        sSymbol = "waistleg"
        Dim strUidLeg As String = ""
        If strUidLeg = "" Then
            'Define DB Fields
            PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFIELDS.D;")

            'Find "mainpatientdetails" and get position
            PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
            PR_PutLine("STRING sMPD_Name;")
            PR_PutLine("ANGLE  aMPD_Angle;")

            Dim strUidTitle As String = ""
            PR_PutLine("hMPD = UID (" & QQ & "find" & QC & Val(strUidTitle) & ");")
            PR_PutLine("if (hMPD)")
            PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")

            'Insert symbol
            PR_PutLine("if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
            PR_PutLine("  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
            If g_sTxtLeg = "Left" Then
                PR_PutLine("  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 1.5,xyMPD_Origin.y );")
            Else
                PR_PutLine("  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 3,xyMPD_Origin.y );")
            End If
            PR_PutLine("  }")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
        Else
            'Use existing symbol
            PR_PutLine("hLeg = UID (" & QQ & "find" & QC & Val(strUidLeg) & ");")
            PR_PutLine("if (!hLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")

        End If

        'Update the BODY Box symbol
        PR_PutLine("SetDBData( hLeg" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "Leg" & QCQ & g_sTxtLeg & QQ & ");")

        PR_PutLine("SetDBData( hLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid(g_sTxtTapeLengths, 1, 60) & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid(g_sTxtTapeLengths, 61) & QQ & ");")

        PR_PutLine("SetDBData( hLeg" & CQ & "AnkleTape" & QCQ & g_sTxtAnkleTape & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "FirstTape" & QCQ & g_sTxtFirstTape & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "LastTape" & QCQ & g_sTxtLastTape & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "ToeStyle" & QCQ & g_sTxtToeStyle & QQ & ");")

        PR_PutLine("SetDBData( hLeg" & CQ & "FootPleat1" & QCQ & txtFootPleat1.Text & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "FootPleat2" & QCQ & txtFootPleat2.Text & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "TopLegPleat1" & QCQ & txtTopLegPleat1.Text & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "TopLegPleat2" & QCQ & txtTopLegPleat2.Text & QQ & ");")

        PR_PutLine("SetDBData( hLeg" & CQ & "FootLength" & QCQ & txtFootLength.Text & QQ & ");")
        PR_PutLine("SetDBData( hLeg" & CQ & "Data" & QCQ & g_sTxtData & QQ & ");")

        FileClose(fNum)
    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable

        PrintLine(fNum, sLine)
    End Sub
    Private Sub Text1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Enter
        Dim Index As Short = Text1.GetIndex(eventSender)
        WHRightLeg1.Select_Text_In_Box(Text1(Index))
    End Sub

    Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
        Dim Index As Short = Text1.GetIndex(eventSender)
        Dim nValue As Double
        nValue = WHRightLeg1.Validate_And_Display_Text_In_Box(Text1(Index), labTapeInInches(Index))

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
    Private Sub txtFootLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Enter
        WHRightLeg1.Select_Text_In_Box(txtFootLength)
    End Sub

    Private Sub txtFootLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Leave
        Dim nValue As Double
        nValue = WHRightLeg1.Validate_And_Display_Text_In_Box(txtFootLength, labFootLengthInInches)
    End Sub
    Private Sub txtFootPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat1.Enter
        WHRightLeg1.Select_Text_In_Box(txtFootPleat1)
    End Sub

    Private Sub txtFootPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat1.Leave
        WHRightLeg1.Validate_Text_In_Box(txtFootPleat1)
    End Sub

    Private Sub txtFootPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat2.Enter
        WHRightLeg1.Select_Text_In_Box(txtFootPleat2)
    End Sub

    Private Sub txtFootPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootPleat2.Leave
        WHRightLeg1.Validate_Text_In_Box(txtFootPleat2)
    End Sub
    Private Sub txtTopLegPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat1.Enter
        WHRightLeg1.Select_Text_In_Box(txtTopLegPleat1)
    End Sub

    Private Sub txtTopLegPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat1.Leave
        WHRightLeg1.Validate_Text_In_Box(txtTopLegPleat1)
    End Sub

    Private Sub txtTopLegPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat2.Enter
        WHRightLeg1.Select_Text_In_Box(txtTopLegPleat2)
    End Sub

    Private Sub txtTopLegPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopLegPleat2.Leave
        WHRightLeg1.Validate_Text_In_Box(txtTopLegPleat2)
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
        g_sTxtToeStyle = VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex)

        'Option Boxes
        'If optLeftLeg.Checked = True Then
        '    g_sTxtLeg = "Left"
        'Else
        '    g_sTxtLeg = "Right"
        'End If

        g_sTxtData = "0"
        If optBelowKnee.Checked = True Then g_sTxtData = "-1"
        If optAboveKnee.Checked = True Then g_sTxtData = "1"

        'Tape boxes
        'Assume that data has been validated earlier

        'Set initial values
        g_sTxtTapeLengths = ""
        g_sTxtFirstTape = CStr(-1)
        g_sTxtLastTape = CStr(30) 'Take care of special case where last tape is text1(29)

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
            g_sTxtTapeLengths = g_sTxtTapeLengths & sJustifiedString

            'Set first and last tape (assumes no holes in data)
            If CDbl(g_sTxtFirstTape) < 0 And nValue > 0 Then g_sTxtFirstTape = CStr(ii + 1)
            If CDbl(g_sTxtLastTape) = 30 And CDbl(g_sTxtFirstTape) > 0 And nValue = 0 Then g_sTxtLastTape = CStr(ii)
        Next ii


        'Check if heel or ankle changed, If they have make the changes and warn
        'that configuring is required
        nAnkleTape = 0
        nAnkleLen = 0
        nHeelLen = WHRightLeg1.fnDisplayToInches(Val(Text1(5).Text))

        If (nHeelLen <> 0) And (nHeelLen < HEEL_TOL) Then
            nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
            nAnkleLen = WHRightLeg1.fnDisplayToInches(Val(Text1(nAnkleTape).Text))
        Else
            nAnkleTape = 7 'Large heel Ankle at +3 tape
            nAnkleLen = WHRightLeg1.fnDisplayToInches(Val(Text1(nAnkleTape).Text))
        End If

        'Note:-
        '        The DRAFIX macros index tape lengths from 1 not 0 as VB does.
        '        Therefor nAnkleTape is incremented by 1
        '
        If (nHeelLen <> g_nHeelLen) Or (nAnkleLen <> g_nAnkleLen) Then
            MsgBox("Figure Ankle for this leg.", 48)
            If nHeelLen = 0 Then g_sTxtAnkleTape = "-1 0 0 0 0 0 0 0"
        End If

        'Set the txtAnkleTape to reflect the footless state
        'Only if it has not been set before
        If nHeelLen = 0 And g_sTxtAnkleTape = "" Then g_sTxtAnkleTape = "-1 0 0 0 0 0 0 0"


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
            nHeelLength = WHRightLeg1.fnDisplayToInches(Val(Text1(5).Text))
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

            nFootLength = WHRightLeg1.fnDisplayToInches(Val(txtFootLength.Text))

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
    Private Sub PR_DrawWaistLeg()
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

        Dim strTag(22), strTextString(22) As String
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
        strTag(15) = "TapeMMs"
        strTag(16) = "TapeMMs2"
        strTag(17) = "Reduction"
        strTag(18) = "Reduction2"
        strTag(19) = "Grams"
        strTag(20) = "Grams2"
        strTag(21) = "Pressure"
        strTag(22) = "Fabric"

        strTextString(1) = txtFileNo.Text
        strTextString(2) = g_sTxtLeg
        strTextString(3) = Mid(g_sTxtTapeLengths, 1, 60)
        strTextString(4) = Mid(g_sTxtTapeLengths, 61)
        strTextString(5) = g_sTxtAnkleTape
        strTextString(6) = g_sTxtFirstTape
        strTextString(7) = g_sTxtLastTape
        strTextString(8) = cboToeStyle.Text
        strTextString(9) = txtFootPleat1.Text
        strTextString(10) = txtFootPleat2.Text
        strTextString(11) = txtTopLegPleat1.Text
        strTextString(12) = txtTopLegPleat2.Text
        strTextString(13) = txtFootLength.Text
        strTextString(14) = g_sTxtData
        strTextString(15) = ""
        strTextString(16) = ""
        strTextString(17) = ""
        strTextString(18) = ""
        strTextString(19) = ""
        strTextString(20) = ""
        strTextString(21) = ""
        strTextString(22) = ""

        Dim XVal As Double = 3.0
        Dim strBlkName As String = "WAISTRIGHTLEG"

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
                For ii = 1 To 22
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
                                acAttRef.TextString = g_sTxtLeg
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt1", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = Mid(g_sTxtTapeLengths, 1, 60)
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = Mid(g_sTxtTapeLengths, 61)
                            ElseIf acAttRef.Tag.ToUpper().Equals("AnkleTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_sTxtAnkleTape
                            ElseIf acAttRef.Tag.ToUpper().Equals("FirstTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_sTxtFirstTape
                            ElseIf acAttRef.Tag.ToUpper().Equals("LastTape", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_sTxtLastTape
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
                                acAttRef.TextString = g_sTxtData
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeMMs", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeMMs2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Reduction", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Reduction2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Grams", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Grams2", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Pressure", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
                            ElseIf acAttRef.Tag.ToUpper().Equals("Fabric", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = ""
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
            If (_sClass.GetXrecord("WaistRightLeg", "WHRIGHTDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("WaistRightLeg", "WHRIGHTDIC")
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
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), g_sTxtLeg))

            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optLeftLeg.Checked))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optRightLeg.Checked))

            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optBelowKnee.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optAboveKnee.Checked))
            Dim ii As Double
            For ii = 0 To 29
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), labTapeInInches(ii).Text))
            Next ii

            _sClass.SetXrecord(resbuf, "WaistRightLeg", "WHRIGHTDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("WaistRightLeg", "WHRIGHTDIC")
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
                g_sTxtLeg = arr(36).Value
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
    Private Function PR_GetBlkAttributeValues() As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If Not acBlkTbl.Has("WAISTBODY") Then
                MsgBox("Can't find WAISTBODY symbol to update!", 16, "Waist - Dialogue")
                Return False
            End If
            Dim blkRecIdBody As ObjectId = ObjectId.Null
            blkRecIdBody = acBlkTbl("WAISTBODY")
            If blkRecIdBody <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecIdBody, OpenMode.ForRead)
                Dim blkRef As BlockReference = New BlockReference(New Point3d(0, 0, 0), blkRecIdBody)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                            g_sLegStyle = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("TOSCir", StringComparison.InvariantCultureIgnoreCase) Then
                            sTOSCir = acAttDef.TextString
                        End If
                    End If
                Next
            End If
            If Not acBlkTbl.Has("WAISTRIGHTLEG") Then
                MsgBox("Can't find WAISTRIGHTLEG symbol to update!", 16, "Waist - Dialogue")
                Return False
            End If
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
            For Each objID As ObjectId In acBlkTblRec
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                If TypeOf dbObj Is BlockReference Then
                    Dim blkRef As BlockReference = dbObj
                    If blkRef.Name = "WAISTRIGHTLEG" Then
                        For Each attributeID As ObjectId In blkRef.AttributeCollection
                            Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                            If TypeOf attRefObj IsNot AttributeReference Then
                                Continue For
                            End If
                            Dim acAttRef As AttributeReference = attRefObj
                            If acAttRef.Tag.ToUpper().Equals("AnkleTape", StringComparison.InvariantCultureIgnoreCase) Then
                                ''nFabricClass = Val(acAttRef.TextString)
                                nFabricClass = WHBODDIA1.fnGetNumber(acAttRef.TextString, 8)
                            ElseIf acAttRef.Tag.ToUpper().Equals("Pressure", StringComparison.InvariantCultureIgnoreCase) Then
                                sPressure = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeMMs", StringComparison.InvariantCultureIgnoreCase) Then
                                sMMAnkle = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeMMs2", StringComparison.InvariantCultureIgnoreCase) Then
                                sMMAnkle = sMMAnkle & acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("Reduction", StringComparison.InvariantCultureIgnoreCase) Then
                                sReductionAnkle = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("Reduction2", StringComparison.InvariantCultureIgnoreCase) Then
                                sReductionAnkle = sReductionAnkle & acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("Grams", StringComparison.InvariantCultureIgnoreCase) Then
                                sGramsAnkle = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("Grams2", StringComparison.InvariantCultureIgnoreCase) Then
                                sGramsAnkle = sGramsAnkle & acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("Fabric", StringComparison.InvariantCultureIgnoreCase) Then
                                sFabric = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt1", StringComparison.InvariantCultureIgnoreCase) Then
                                sTapeLengths = acAttRef.TextString
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt2", StringComparison.InvariantCultureIgnoreCase) Then
                                sTapeLengths = sTapeLengths & acAttRef.TextString
                            End If
                        Next
                    End If
                End If
            Next
        End Using
        If sPressure = "" Or sPressure = " " Then
            MsgBox("Figure RIGHT Leg before drawing", 0, "Second Leg Drawing")
            Return False
        End If
        Return True
    End Function
    Private Function readLegCurveInfo(ByRef ptColl As Point3dCollection) As Boolean
        Dim bIsExist As Boolean = False
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("LegCurve", "LEGCURVEDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim strValue As String
                Dim sFile As String = fnGetSettingsPath("LookupTables") + "\\LEGCURVE.DAT"
                Dim hChan As Object
                hChan = FreeFile()
                ''hChan = Open("file", sFile, "readonly")
                FileOpen(hChan, sFile, VB.OpenMode.Output)
                Dim ii As Double
                For ii = 0 To arr.Length - 1
                    strValue = arr(ii).Value
                    PrintLine(hChan, strValue)
                    Dim xyPoly As XY
                    xyPoly.X = WHBODDIA1.fnGetNumber(strValue, 1)
                    xyPoly.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                    ptColl.Add(New Point3d(xyPoly.X, xyPoly.Y, 0))
                Next ii
                FileClose(hChan)
                bIsExist = True
            End If
        Catch ex As Exception

        End Try
        Return bIsExist
    End Function

    Private Function readCutOutInfo(ByRef xyOBody As XY, ByRef xyTOS As XY, ByRef xyWaist As XY, ByRef xyMidPoint As XY, ByRef xyLargest As XY,
                               ByRef xyFold As XY, ByRef xyFoldPanty As XY, ByRef xyButtockArcCen As XY, ByRef nFoldHt As Double) As Boolean
        Dim bIsExist As Boolean = False
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("WaistCutOut", "WHCUTOUTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim strValue As String
                strValue = arr(0).Value
                xyOBody.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyOBody.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(1).Value
                xyTOS.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyTOS.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(2).Value
                xyWaist.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyWaist.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(3).Value
                xyMidPoint.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyMidPoint.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(4).Value
                xyLargest.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyLargest.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(5).Value
                xyFold.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyFold.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                strValue = arr(6).Value
                xyFoldPanty.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyFoldPanty.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                nFoldHt = Val(arr(7).Value)
                strValue = arr(8).Value
                xyButtockArcCen.X = WHBODDIA1.fnGetNumber(strValue, 1)
                xyButtockArcCen.Y = WHBODDIA1.fnGetNumber(strValue, 2)
                bIsExist = True
            End If
        Catch ex As Exception

        End Try
        Return bIsExist
    End Function
    Private Function FNLegStartHt(ByRef nFirstTape As Double, ByRef nEndTape As Double, ByRef nLastTape As Double, ByRef nFootPleat1 As Double,
                                  ByRef nFootPleat2 As Double, ByRef nTopLegPleat1 As Double, ByRef nTopLegPleat2 As Double, ByRef Footless As Boolean) As Double
        ''File Name : WHPOW.D
        Dim nn, nLength, nSpace As Double
        nn = nFirstTape + 1
        nLength = 0
        While (nn <= nEndTape)
            If (nn <= 4) Then
                nSpace = 1.375
            End If
            If (nn = 5) Then
                nSpace = 1.21875
            End If
            If (nn = 6) Then
                nSpace = 1.5
            End If
            If (nn = 7) Then
                nSpace = 1.5
            End If
            If (nn = 8) Then
                nSpace = 1.21875
            End If
            If (nn > 8) Then
                nSpace = 1.25
            End If
            If ((nn = nFirstTape + 1) And (nFootPleat1 <> 0)) Then
                nSpace = nFootPleat1
            End If
            If ((nn = nFirstTape + 2) And (nFootPleat2 <> 0)) Then
                nSpace = nFootPleat2
            End If
            If ((nn = nLastTape - 1) And (nTopLegPleat2 <> 0)) Then
                nSpace = nTopLegPleat2
            End If
            If ((nn = nLastTape) And (nTopLegPleat1 <> 0) And Footless) Then
                nSpace = nTopLegPleat1
            End If
            nLength = nLength + nSpace
            nn = nn + 1
        End While
        Return (nLength)
    End Function
    Private Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef Y As Double)
        'Utility to return a point based on the X and Y values given
        xyReturn.X = X
        xyReturn.Y = Y
    End Sub
    Private Function PR_GetInsertionPoint(ByRef xyOtemplate As XY) As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Give origin point for Leg "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Return False
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(xyOtemplate, ptStart.X, ptStart.Y)
        Return True
    End Function
    Sub PR_CalcPolar(ByRef xyStart As XY, ByRef nLength As Double, ByVal nAngle As Double, ByRef xyReturn As XY)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double
        'Convert from degees to radians
        nAngle = nAngle * PI / 180

        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.X + A
        xyReturn.Y = xyStart.Y + B
    End Sub
    Sub PR_DrawXMarker(ByRef xyInsertion As XY)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As XY
        PR_CalcPolar(xyBase, 135, 0.0625, xyStart)
        PR_CalcPolar(xyBase, -45, 0.0625, xyEnd)
        PR_CalcPolar(xyBase, 45, 0.0625, xySecSt)
        PR_CalcPolar(xyBase, -135, 0.0625, xySecEnd)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("X Marker") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "X Marker"
                Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.Y, 0), New Point3d(xyEnd.X, xyEnd.Y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xySecSt.X, xySecSt.Y, 0), New Point3d(xySecEnd.X, xySecEnd.Y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl("X Marker")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyInsertion.X, xyInsertion.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Function StringMiddle(ByRef strText As String, ByRef nStart As Integer, ByRef nLength As Integer) As String
        Return Mid(strText, nStart, nLength)
    End Function
    Private Sub ScanLine(ByRef sLine As String, ByRef nNo As Double, ByRef sScale As String,
                         ByRef nSpace As Double, ByRef n20Len As Double, ByRef nReduction As Double)
        nNo = FN_GetNumber(sLine, 1)
        sScale = FN_GetString(sLine, 2)
        nSpace = FN_GetNumber(sLine, 3)
        n20Len = FN_GetNumber(sLine, 4)
        nReduction = FN_GetNumber(sLine, 5)
    End Sub
    Function FN_GetString(ByVal sString As String, ByRef iIndex As Short) As String

        Dim ii, iPos As Short

        'Initial error checking
        sString = Trim(sString) 'Remove leading and trailing blanks

        If Len(sString) = 0 Then
            FN_GetString = ""
            Exit Function
        End If

        'Prepare string
        sString = sString & " " 'Trailing blank as stopper for last item

        'Get iIndexth item
        For ii = 1 To iIndex
            iPos = InStr(sString, " ")
            If ii = iIndex Then
                sString = VB.Left(sString, iPos - 1)
                FN_GetString = sString
                Exit Function
            Else
                sString = LTrim(Mid(sString, iPos))
                If Len(sString) = 0 Then
                    FN_GetString = ""
                    Exit Function
                End If
            End If
        Next ii

        'The function should have exited befor this, however just in case
        '(iIndex = 0) we indicate an error,
        FN_GetString = ""

    End Function
    Private Function FNGetTape(ByRef nIndex As Double) As Double
        ''-----------Dim nInt As Double = Value("scalar", StringMiddle(txtLeftLengths.Text, ((nIndex - 1) * 4) + 1, 3))
        Dim nInt As Double = Val(StringMiddle(sTapeLengths, ((nIndex - 1) * 4) + 1, 3))
        ''----------Dim nDec As Double = Value("scalar", StringMiddle(txtLeftLengths.Text, ((nIndex - 1) * 4) + 4, 1))
        Dim nDec As Double = Val(StringMiddle(sTapeLengths, ((nIndex - 1) * 4) + 4, 1))
        Return (nInt + (nDec * 0.1))
    End Function
    Function FNDecimalise(ByRef nDisplay As Double) As Double
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
        If WHRightLeg1.g_nUnitsFac <> 1 Then
            FNDecimalise = nDisplay * iSign
            Exit Function
        End If

        'Imperial units
        iInt = Int(nDisplay)
        nDec = nDisplay - iInt
        'Check that conversion is possible (return -1 if not)
        If nDec > 0.8 Then
            FNDecimalise = -1
        Else
            FNDecimalise = (iInt + (nDec * 0.125 * 10)) * iSign
        End If

    End Function
    Private Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Object, ByVal nTextmode As Double)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X, xyInsert.Y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.WidthFactor = 0.6
                ''acText.HorizontalMode = nTextmode
                acText.Justify = nTextmode
                ''If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = New Point3d(xyInsert.X, xyInsert.Y, 0)
                ''End If
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawRuler(ByRef sSymbol As String, ByRef xyStart As XY, ByRef sTape As String)
        Dim xyPt As XY
        Dim xyTapeFst, xyTapeSec, xyTapeEnd, xyTapeText As XY
        Dim sTapeSymbol As String = sSymbol
        sTapeSymbol = Trim(sTapeSymbol)
        PR_MakeXY(xyPt, 0, 0)
        PR_MakeXY(xyTapeFst, xyPt.X, xyPt.Y - 0.05)
        PR_MakeXY(xyTapeSec, xyPt.X, xyPt.Y - 0.5)
        PR_MakeXY(xyTapeEnd, xyPt.X, xyTapeSec.Y + 0.05)
        PR_MakeXY(xyTapeText, xyPt.X, xyPt.Y - 0.25)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sTapeSymbol) Then
                Dim blkTblRecTape As BlockTableRecord = New BlockTableRecord()
                blkTblRecTape.Name = sTapeSymbol

                Dim acLine As Line = New Line(New Point3d(xyPt.X, xyPt.Y, 0), New Point3d(xyTapeFst.X, xyTapeFst.Y, 0))
                blkTblRecTape.AppendEntity(acLine)

                acLine = New Line(New Point3d(xyTapeSec.X, xyTapeSec.Y, 0), New Point3d(xyTapeEnd.X, xyTapeEnd.Y, 0))
                blkTblRecTape.AppendEntity(acLine)

                '' Create a single-line text object
                Dim acText As DBText = New DBText()
                acText.Position = New Point3d(xyTapeText.X, xyTapeText.Y, 0)
                acText.Height = 0.125
                acText.TextString = sTape
                acText.Rotation = 0
                acText.Justify = AttachmentPoint.MiddleCenter
                acText.AlignmentPoint = New Point3d(xyTapeText.X, xyTapeText.Y, 0)
                blkTblRecTape.AppendEntity(acText)

                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecTape)
                acTrans.AddNewlyCreatedDBObject(blkTblRecTape, True)
                blkRecId = blkTblRecTape.Id
            Else
                blkRecId = acBlkTbl(sTapeSymbol)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyStart.X, xyStart.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawLine(ByRef xyStart As XY, ByRef xyFinish As XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                      OpenMode.ForWrite)

            '' Create a line that starts at 5,5 and ends at 12,3
            Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.Y, 0),
                                    New Point3d(xyFinish.X, xyFinish.Y, 0))

            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
            End If
            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As XY, ByRef bIsCenter As Boolean)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            Dim mtx As New MText()
            mtx.Location = New Point3d(xyInsert.X, xyInsert.Y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = 0
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            If bIsCenter = True Then
                mtx.Attachment = AttachmentPoint.TopCenter
            End If
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Function FN_CalcAngle(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360
        Dim X, y As Object
        Dim rAngle As Double

        X = xyEnd.X - xyStart.X
        y = xyEnd.Y - xyStart.Y

        'Horizomtal
        If X = 0 Then
            If y > 0 Then
                FN_CalcAngle = 90
            Else
                FN_CalcAngle = 270
            End If
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        If y = 0 Then
            If X > 0 Then
                FN_CalcAngle = 0
            Else
                FN_CalcAngle = 180
            End If
            Exit Function
        End If

        'All other cases
        rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If

    End Function
    Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus
        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)
    End Function
    Private Sub PR_DrawArc(ByRef xyCen As XY, ByRef nRad As Double, ByRef nStartAng As Double, ByRef nDeltaAng As Double)
        ' this procedure draws an arc between two points

        Dim nEndAng As Object
        nEndAng = nStartAng + nDeltaAng

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create an arc that is at 6.25,9.125 with a radius of 6, and
            '' starts at 64 degrees and ends at 204 degrees
            Using acArc As Arc = New Arc(New Point3d(xyCen.X, xyCen.Y, 0),
                                         nRad, (nStartAng * (PI / 180)), (nEndAng * (PI / 180)))

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Function FN_CirLinInt(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyCen As XY, ByRef nRad As Double, ByRef xyInt As XY) As Short
        'Function to calculate the intersection between
        'a line and a circle.
        'Note:-
        '    Returns true if intersection found.
        '    The first intersection (only) is found.
        '    Ported from DRAFIX CAD DLG version.
        '

        Static nM, nC, nA, nSlope, nB, nK, nCalcTmp As Object
        Static nRoot As Double
        Static nSign As Short

        nSlope = FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        If nSlope = 0 Or nSlope = 180 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = xyStart.Y
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Vertical Line
        If nSlope = 90 Or nSlope = 270 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.Y + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                    xyInt.Y = nRoot
                    xyInt.X = xyStart.X
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Non-othogonal line
        If nSlope > 0 Then
            nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X) 'Slope
            nK = xyStart.Y - nM * xyStart.X 'Y-Axis intercept
            nA = (1 + nM ^ 2)
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False
    End Function
    Sub PR_DrawPoly(ByRef PointCollection As Point3dCollection)
        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to draw a POLYLINE through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '
        '
        Dim ii As Short

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                            OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Using acPoly As Polyline = New Polyline()
                For ii = 0 To PointCollection.Count - 1
                    acPoly.AddVertexAt(ii, New Point2d(PointCollection(ii).X, PointCollection(ii).Y), 0, 0, 0)
                Next ii

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''--------------File Name : WHLG2LEG.D --------------------
        Dim ptFstLegColl As Point3dCollection = New Point3dCollection
        If readLegCurveInfo(ptFstLegColl) = False Then
            MsgBox("Can't find First Leg Curve, to copy body curve" & Chr(10) & "Draw or Re-Draw First Leg.", 16, "Waist - Leg Dialog")
            Exit Sub
        End If
        ''--------------File Name : WHLG2ORG.D --------------------
        If PR_GetBlkAttributeValues() = False Then
            Exit Sub
        End If

        Dim xyOBody, xyTOS, xyWaist, xyMidPoint, xyLargest, xyFold, xyFoldPanty, xyButtockArcCen As XY
        Dim BodyFound As Boolean = False
        Dim nFoldHt As Double
        If readCutOutInfo(xyOBody, xyTOS, xyWaist, xyMidPoint, xyLargest, xyFold, xyFoldPanty, xyButtockArcCen, nFoldHt) = False Then
            MsgBox("No BODY found!, draw Cut-Out and try again.", 16, "Waist - Leg Dialog")
            Exit Sub
        End If

        ''// setup values for tape
        ''// Establish leg start position relative to the body 
        Dim nLegStartHt, nBodyLegTapePos, nValue, nFirstTape, nLastTape, ii As Double
        nBodyLegTapePos = Int(nFoldHt / 1.5) + 6
        Dim FootLess As Boolean = WHRightLeg1.fnDisplayToInches(Val(Text1(5).Text))
        Dim nFootPleat1, nFootPleat2, nTopLegPleat1, nTopLegPleat2 As Double
        nFootPleat1 = Val(txtFootPleat1.Text)
        nFootPleat2 = Val(txtFootPleat2.Text)
        nTopLegPleat1 = Val(txtTopLegPleat1.Text)
        nTopLegPleat2 = Val(txtTopLegPleat2.Text)
        nFirstTape = -1
        nLastTape = 30
        Dim sToeStyle As String = cboToeStyle.Text
        For ii = 0 To 29
            nValue = Val(Text1(ii).Text)
            'Set first and last tape (assumes no holes in data)
            If CDbl(nFirstTape) < 0 And nValue > 0 Then
                nFirstTape = CStr(ii + 1)
            End If
            If CDbl(nLastTape) = 30 And CDbl(nFirstTape) > 0 And nValue = 0 Then
                nLastTape = CStr(ii)
            End If
        Next ii
        Dim nExtTempltTol1max As Double = 12.0      ''// as mesured
        Dim nExtTempltTol2max As Double = 23.875
        Dim nExtTempltTol2min As Double = 12.0      ''// at 28 1/2" scale	
        Dim nExtTempltTol3min As Double = 23.875    ''// at  27" scale
        If (nFoldHt < nExtTempltTol1max And FootLess = False) Then
            nLegStartHt = FNLegStartHt(nFirstTape, nBodyLegTapePos, nLastTape, nFootPleat1, nFootPleat2, nTopLegPleat1, nTopLegPleat2, FootLess)
        End If
        If (FootLess) Then
            nLegStartHt = FNLegStartHt(nFirstTape, nBodyLegTapePos, nLastTape, nFootPleat1, nFootPleat2, nTopLegPleat1, nTopLegPleat2, FootLess) '' // Panty
        End If
        If ((nFoldHt >= nExtTempltTol2min And nFoldHt <= nExtTempltTol2max) And FootLess = False) Then
            nLegStartHt = FNLegStartHt(nFirstTape, nBodyLegTapePos - 1, nLastTape, nFootPleat1, nFootPleat2, nTopLegPleat1, nTopLegPleat2, FootLess)
        End If
        If (nFoldHt > nExtTempltTol3min And FootLess = False) Then
            nLegStartHt = FNLegStartHt(nFirstTape, nBodyLegTapePos - 2, nLastTape, nFootPleat1, nFootPleat2, nTopLegPleat1, nTopLegPleat2, FootLess)
        End If

        '' // Get Origin
        Dim xyO As XY
        ''If (!GetUser("xy", "Give origin point for Leg", & xyOtemplate)) Then Exit(%ok, "User Canceled");
        If PR_GetInsertionPoint(xyOtemplate) = False Then
            MsgBox("User Canceled", 16, "Waist Leg - Dialog")
            Exit Sub
        End If
        ''// Offset body origin
        xyO.Y = xyOtemplate.Y
        xyO.X = xyOtemplate.X + nLegStartHt
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("marker", "xmarker", xyOtemplate, 0.125)
        PR_DrawXMarker(xyOtemplate)
        '        If (hEnt) Then {
        '  sID = StringMiddle(MakeString("long", UID("get", hEnt)), 1, 4) ; 
        '  While (StringLength(sID) < 4) sID = sID + " ";
        '  sID = sID + sFileNo + sLeg ;
        '  SetDBData(hEnt, "ID", sID + "Origin");
        '  SetDBData(hEnt, "units", sUnits) ;
        '  If (FootLess) Then
        '                    SetDBData(hEnt, "Data", "WHPT") ;
        '  Else
        '                    SetDBData(hEnt, "Data", "WHFT") ;  
        '  SetDBData(hEnt, "Pressure", sPressure) ;
        '  SetDBData(hEnt, "TapeLengthsPt1", StringMiddle(sTapeLengths, 1, 60));
        '  SetDBData(hEnt, "TapeLengthsPt2", StringMiddle(sTapeLengths, 61, StringLength(sTapeLengths) - 60));
        '  SetDBData(hEnt, "patient", sPatient) ;
        '  SetDBData(hEnt, "age", sAge) ;
        '  SetDBData(hEnt, "WorkOrder", sWorkOrder) ;
        '//  SetDBData (hEnt,"TOSGivenRed", sTOSRed) ;
        '//  SetDBData (hEnt,"WaistGivenRed", sWaistRed) ;
        '//  SetDBData (hEnt,"FoldHt", sFoldHt) ;
        '//  SetDBData (hEnt,"WaistCir", sWaistCir) ;
        '//  SetDBData (hEnt,"WaistHt", sWaistHt) ;
        '//  SetDBData (hEnt,"TOSCir", sTOSCir) ;
        '//  SetDBData (hEnt,"TOSHt", sTOSHt) ;
        '  SetDBData(hEnt, "Body", sLegStyle) ;
        '  SetDBData(hEnt, "FirstTape", MakeString("scalar", nFirstTape));
        '  SetDBData(hEnt, "LastTape", MakeString("scalar", nLastTape));
        '  }

        ''--------------File Name : WHLGxTMP.D --------------------
        ''// set template defaults
        Dim nOffset As Double = 0.5
        ''// Set layer   
        ARMDIA1.PR_SetLayer("Construct")

        ''// Load template data file
        ''PROpenTemplateFile()
        Dim sFile As String = ""
        If (nFabricClass = 0) Then
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "MMHG.DAT"
        Else
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "DS.DAT"
        End If
        Dim hChan As Object
        hChan = FreeFile()
        ''hChan = Open("file", sFile, "readonly")
        FileOpen(hChan, sFile, VB.OpenMode.Input)
        Dim nNo, nSpace, n20Len, nReduction, nn As Double
        Dim sLine As String = ""
        Dim sScale As String = ""
        If (hChan) Then
            sLine = LineInput(hChan)
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
        Else
            MsgBox("Can't open " + sFile + "\nCheck installation", 16, "Waist Leg - Dialog")
            Exit Sub
        End If

        Dim xyPt1, xyTmp, xyText, xyStart, xyEnd, xyInt As XY
        Dim nSeam As Double = 0.1875
        xyPt1 = xyOtemplate
        xyPt1.Y = xyOtemplate.Y + nOffset + nSeam
        ''// Loop Through data file to get to start of leg tapes
        ''// Ignoring the none relevent ones 
        nn = 1
        While (nn < nFirstTape)
            sLine = LineInput(hChan)
            nn = nn + 1
        End While
        ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
        nSpace = 0
        ''// Loop through displaying points And scales
        Dim nLoop, nAnkleTape, nLtLastHeel, nPleatGiven As Double
        nLtLastHeel = WHRightLeg1.fnDisplayToInches(Val(Text1(5).Text))
        If nLtLastHeel <> 0 Then
            If nLtLastHeel < 9 Then
                nAnkleTape = 6 'Small heel Ankle at +1 1/2 tape
            Else
                nAnkleTape = 7 'Large heel Ankle at +3 tape
            End If
        End If
        nAnkleTape = nAnkleTape + 1
        If (nFabricClass = 2) Then
            nLoop = nAnkleTape - 1 ''  // JOBSTEX Gradient
        Else
            nLoop = nLastTape
        End If
        Dim nTapeLen, nLength, jj, nTickStep, nTickHt As Double
        Dim PantyLeg As Boolean = False
        Dim nLegStyle As Double = WHBODDIA1.fnGetNumber(g_sLegStyle, 1)
        If (nLegStyle = 1) Then
            PantyLeg = True
        End If
        While (nn <= nLoop)
            ''// Get data for leg And place on template
            If (FNGetTape(nn)) Then
                ''sSymbol = MakeString("long", nNo) + "tape"
                Dim sSymbol As String = Str(nNo) + "tape"
                ''// Pleats, NB nSpace Is distance of current tape from the previous tape
                nPleatGiven = 0
                If ((nn = nFirstTape + 1) And (nFootPleat1 <> 0)) Then
                    nSpace = nFootPleat1
                    nPleatGiven = 1
                End If
                If ((nn = nFirstTape + 2) And (nFootPleat2 <> 0)) Then
                    nSpace = nFootPleat2
                    nPleatGiven = 1
                End If
                If ((nn = nLastTape - 1) And (nTopLegPleat2 <> 0)) Then
                    nSpace = nTopLegPleat2
                    nPleatGiven = 1
                End If
                If ((nn = nLastTape) And (nTopLegPleat1 <> 0)) Then
                    nSpace = nTopLegPleat1
                    nPleatGiven = 1
                End If
                xyPt1.X = xyPt1.X + nSpace
                xyTmp.X = xyPt1.X
                ''nTapeLen = FNRound(FNDecimalise(FNGetTape(nn)) * nUnitsFac)
                nTapeLen = FNDecimalise(FNGetTape(nn)) * WHRightLeg1.g_nUnitsFac
                nLength = n20Len / 20 * nTapeLen
                xyTmp.Y = xyOtemplate.Y + nSeam + nLength
                ''AddEntity("marker", "xmarker", xyTmp, 0.2, 0.2)
                PR_DrawXMarker(xyTmp)
                '' // Add Pleat indication
                If (nPleatGiven = 1) Then
                    ''SetData("TextHorzJust", 2)
                    '' AddEntity("text", "PLEAT", xyTmp.X - (nSpace / 2), xyTmp.Y + 0.5)
                    PR_MakeXY(xyText, xyTmp.X - (nSpace / 2), xyTmp.Y + 0.5)
                    PR_DrawText("PLEAT", xyText, 0.1, 0, 2)
                End If

                ''// Add Ticks And Text at each scale
                ''// Start ticks 2 below And 2 above
                ii = Int(nLength / (n20Len / 20)) - 1
                jj = ii + 2

                nTickStep = n20Len / (20 * 8)
                nTickHt = xyOtemplate.Y + nSeam + ii * nTickStep * 8
                ''// Add Reduction Text
                ''SetData("TextHeight", 0.125)
                ''SetData("TextHorzJust", 1)
                ''AddEntity("text", MakeString("long", nReduction), xyPt1.X + 0.2, nTickHt - 0.4)
                PR_MakeXY(xyText, xyPt1.X + 0.2, nTickHt - 0.4)
                PR_DrawText(Str(nReduction), xyText, 0.125, 0, 1)
                ''AddEntity("text", Format("length", nTapeLen), xyPt1.X + 0.2, nTickHt - nSeam)
                PR_MakeXY(xyText, xyPt1.X + 0.2, nTickHt - nSeam)
                PR_DrawText(WHRightLeg1.fnInchesToText(nTapeLen), xyText, 0.125, 0, 1)
                ''// Print revised reduction at ankle tape
                If (nn = nAnkleTape And PantyLeg = False) Then
                    ''AddEntity("text", sReductionAnkle, xyPt1.X + 0.4, nTickHt - 0.4)
                    PR_MakeXY(xyText, xyPt1.X + 0.4, nTickHt - 0.4)
                    PR_DrawText(Mid(sReductionAnkle, (nn * 3) + 1, 3), xyText, 0.125, 0, 1)
                    If (nFabricClass = 0) Then
                        ''AddEntity("text", sGramsAnkle + " grams", xyPt1.X + 0.4, nTickHt - 0.6)
                        PR_MakeXY(xyText, xyPt1.X + 0.4, nTickHt - 0.6)
                        PR_DrawText(Mid(sGramsAnkle, (nn * 3) + 1, 3) + " grams", xyText, 0.125, 0, 1)
                    Else
                        ''AddEntity("text", sGramsAnkle + " stretch", xyPt1.X + 0.4, nTickHt - 0.6)
                        PR_MakeXY(xyText, xyPt1.X + 0.4, nTickHt - 0.6)
                        PR_DrawText(Mid(sGramsAnkle, (nn * 3) + 1, 3) + " stretch", xyText, 0.125, 0, 1)
                    End If
                    ''AddEntity("text", sMMAnkle + " mm", xyPt1.X + 0.4, nTickHt - 0.8)
                    PR_MakeXY(xyText, xyPt1.X + 0.4, nTickHt - 0.8)
                    PR_DrawText(Mid(sMMAnkle, (nn * 3) + 1, 3) + " mm", xyText, 0.125, 0, 1)
                    ''AddEntity("marker", "xmarker", xyPt1.X + 0.3, nTickHt - 0.5, 0.2, 0.2)
                    PR_MakeXY(xyText, xyPt1.X + 0.3, nTickHt - 0.5)
                    PR_DrawXMarker(xyText)
                End If
                ''// Add tape ID
                'If (!Symbol("find", sSymbol)) Then
                '    Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists");
                'End If
                'AddEntity("symbol", sSymbol, xyPt1.X, nTickHt)
                Dim strLab As String = Label1(nn - 1).Text
                If nNo = 0 Then
                    strLab = "Pleat"
                End If
                Dim xyTape As XY
                PR_MakeXY(xyTape, xyPt1.X, nTickHt)
                PR_DrawRuler(sSymbol, xyTape, strLab)
                While (ii <= jj)
                    Dim nTickLength As Double = 0.22
                    ''AddEntity("line", xyPt1.X, nTickHt, xyPt1.X + nTickLength, nTickHt)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt)
                    PR_DrawLine(xyStart, xyEnd)
                    ''// Add Tick Text
                    Dim nTextMode As Double
                    If (ii < 10) Then
                        ''SetData("TextHorzJust", 4)
                        nTextMode = 3
                    Else
                        ''SetData("TextHorzJust", 2)
                        nTextMode = 1
                    End If
                    ''AddEntity("text", MakeString("long", ii), xyPt1.X + nSeam, nTickHt - 0.01, 0.06, 0.1, 90)
                    PR_MakeXY(xyText, xyPt1.X + nSeam, nTickHt - 0.01)
                    PR_DrawText(Str(ii), xyText, 0.06, (90 * (PI / 180)), nTextMode)
                    nTickLength = 0.05
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep, xyPt1.X + nTickLength, nTickHt + nTickStep)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep)
                    PR_DrawLine(xyStart, xyEnd)
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 3, xyPt1.X + nTickLength, nTickHt + nTickStep * 3)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 3)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 3)
                    PR_DrawLine(xyStart, xyEnd)
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 5, xyPt1.X + nTickLength, nTickHt + nTickStep * 5)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 5)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 5)
                    PR_DrawLine(xyStart, xyEnd)
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 7, xyPt1.X + nTickLength, nTickHt + nTickStep * 7)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 7)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 7)
                    PR_DrawLine(xyStart, xyEnd)
                    nTickLength = 0.08
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 2, xyPt1.X + nTickLength, nTickHt + nTickStep * 2)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 2)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 2)
                    PR_DrawLine(xyStart, xyEnd)
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 6, xyPt1.X + nTickLength, nTickHt + nTickStep * 6)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 6)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 6)
                    PR_DrawLine(xyStart, xyEnd)
                    nTickLength = 0.12
                    ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 4, xyPt1.X + nTickLength, nTickHt + nTickStep * 4)
                    PR_MakeXY(xyStart, xyPt1.X, nTickHt + nTickStep * 4)
                    PR_MakeXY(xyEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 4)
                    PR_DrawLine(xyStart, xyEnd)
                    ii = ii + 1
                    nTickHt = xyOtemplate.Y + nSeam + ii * nTickStep * 8
                End While
            End If
            nn = nn + 1
            sLine = LineInput(hChan)
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)

            ''// draw patient data at HEEL tape
            ''// NOTE Heel tape Is tape #7 in this case
            Dim nTxtX, nTxtY As Double
            If ((nn = 7) Or (FootLess And nn = nFirstTape + 2)) Then
                ARMDIA1.PR_SetLayer("Notes")
                ''SetData("TextVertJust", 32)     ''// Bottom
                ''SetData("TextHorzJust", 1)      ''// Left
                ''SetData("TextHeight", 0.1)
                nTxtY = xyOtemplate.Y + nSeam + 2.25
                nTxtX = xyPt1.X
                Dim sText As String = "Right Leg" + Chr(10) + " "
                ''// Data
                sText = sText + txtPatientName.Text + Chr(10) + " " + sWorkOrder
                If (nFabricClass = 0) Then
                    sText = sText + Chr(10) + " " + StringMiddle(sFabric, 5, sFabric.Length - 4)
                Else
                    sText = sText + Chr(10) + " " + sFabric
                End If
                ''AddEntity("text", sText, nTxtX, nTxtY)
                PR_MakeXY(xyText, nTxtX, nTxtY)
                PR_DrawMText(sText, xyText, False)

                ARMDIA1.PR_SetLayer("Construct")
                sText = "  " + txtFileNo.Text + Chr(10) + " " + txtDiagnosis.Text + Chr(10) + " " + txtAge.Text +
                    Chr(10) + " " + txtSex.Text + Chr(10) + " " + sPressure
                ''AddEntity("text", sText, nTxtX, nTxtY - 0.75)
                PR_MakeXY(xyText, nTxtX, nTxtY - 0.75)
                PR_DrawMText(sText, xyText, False)
                ''// Reset Setup text defaults 
                ''SetData("TextFont", 0)
                ''SetData("TextVertJust", 8)
                ''SetData("TextHorzJust", 1)
                ''SetData("TextHeight", 0.125)
                ''SetData("TextAspect", 0.6)
            End If
        End While
        FileClose(hChan)

        ''--------------File Name : WHLG2DBD.D --------------------
        ''// For all subsequent calculations the measurements are worked
        ''// from the right edge of the body template #8654
        Dim nXscale, nYscale, aTransAngle, nTrans, aAngle As Double
        nXscale = 1.25 / 1.5
        nYscale = 0.5
        ''// Translate
        ''// Points (Using relpolar, which Is less cumbersome than TransXY)
        Dim xyOFirstBody, xyLargestFirstBody As XY
        xyOFirstBody = xyOBody
        aTransAngle = FN_CalcAngle(xyOFirstBody, xyO)
        nTrans = FN_CalcLength(xyOFirstBody, xyO)
        aAngle = aTransAngle

        ''// Body Points
        '' xyFold = CalcXY("relpolar", xyFold, nTrans, aAngle)
        PR_CalcPolar(xyFold, nTrans, aAngle, xyFold)
        ''xyFoldPanty = CalcXY("relpolar", xyFoldPanty, nTrans, aAngle)
        PR_CalcPolar(xyFoldPanty, nTrans, aAngle, xyFoldPanty)
        ''xyTOS = CalcXY("relpolar", xyTOS, nTrans, aAngle)
        PR_CalcPolar(xyTOS, nTrans, aAngle, xyTOS)
        ''xyWaist = CalcXY("relpolar", xyWaist, nTrans, aAngle)
        PR_CalcPolar(xyWaist, nTrans, aAngle, xyWaist)
        ''xyMidPoint = CalcXY("relpolar", xyMidPoint, nTrans, aAngle)
        PR_CalcPolar(xyMidPoint, nTrans, aAngle, xyMidPoint)
        xyLargestFirstBody = xyLargest
        ''xyLargest = CalcXY("relpolar", xyLargest, nTrans, aAngle)
        PR_CalcPolar(xyLargest, nTrans, aAngle, xyLargest)
        ''xyButtockArcCen = CalcXY("relpolar", xyButtockArcCen, nTrans, aAngle)
        PR_CalcPolar(xyButtockArcCen, nTrans, aAngle, xyButtockArcCen)

        ''//
        ''// DRAW Body  closing lines And Constuction lines 
        ''//
        ''// Always Right leg
        ARMDIA1.PR_SetLayer("TemplateRight")
        ''hTemplateLayer = Table("find", "layer", "TemplateRight")
        Dim strLayer As String = "TemplateRight"

        ''// Closing Line at waist Or TOS & Foldlines
        ''AddEntity("line", xyTOS, xyTOS.X, xyO.Y)
        PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
        PR_DrawLine(xyTOS, xyEnd)
        ''AddEntity("line", xyO, xyTOS.X, xyO.Y)
        PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
        PR_DrawLine(xyO, xyEnd)
        Dim nTOSCir As Double = Val(sTOSCir)
        If (nTOSCir = 0) Then
            ARMDIA1.PR_SetLayer("Notes")
            ''AddEntity("line", xyWaist, xyWaist.X, xyO.Y)
            PR_MakeXY(xyEnd, xyWaist.X, xyO.Y)
            PR_DrawLine(xyWaist, xyEnd)
        End If

        ''// Draw construction lines
        ''// 
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("line", xyFold.X, xyLargest.Y, xyFold.X, xyO.Y)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "FoldLine") ; 
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir) + " " + MakeString("scalar", nGivenWaistCir));     	
        PR_MakeXY(xyStart, xyFold.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyFold.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)

        ''AddEntity("line", xyLargest.X, xyLargest.Y,xyLargest.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyLargest.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyLargest.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''AddEntity("line", xyMidPoint.X, xyLargest.Y,xyMidPoint.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyMidPoint.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyMidPoint.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''hEnt = AddEntity("line", xyWaist.X, xyLargest.Y,xyWaist.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyWaist.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyWaist.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        If (nTOSCir <> 0) Then
            ''hEnt = AddEntity("line", xyTOS.X, xyLargest.Y, xyTOS.X, xyO.Y)
            PR_MakeXY(xyStart, xyTOS.X, xyLargest.Y)
            PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
            PR_DrawLine(xyStart, xyEnd)
        End If
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "EOSLine")
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir) + " " + MakeString("scalar", nGivenWaistCir))

        ''AddEntity("arc", xyButtockArcCen, Calc("length", xyButtockArcCen, xyFold), 60, 70) ;
        PR_DrawArc(xyButtockArcCen, FN_CalcLength(xyButtockArcCen, xyFold), 60, 70)

        ''// Add Body template tapes
        ''// From body position given
        ''//
        If (nBodyLegTapePos > 0) Then
            nn = nBodyLegTapePos - 1
        Else
            nn = nLastTape - 1
        End If
        xyPt1.X = xyO.X
        xyPt1.Y = xyO.Y + nSeam + 0.5
        Dim g_sTextList As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 3637½ 3940½"
        While (xyPt1.X < xyTOS.X)
            'sSymbol = MakeString("long", nn) + "tape"
            Dim sSymbol As String = Str(nn) + "tape"
            'If (!Symbol("find", sSymbol)) Then
            '    MsgBox("Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists", 16, "Waist Leg - Dialog")
            '    Exit Sub
            'End If
            'AddEntity("symbol", sSymbol, xyPt1)
            Dim strTape As String = LTrim(Mid(g_sTextList, ((nn - 1) * 3) + 1, 3))
            PR_DrawRuler(sSymbol, xyPt1, strTape)
            xyPt1.X = xyPt1.X + 1.5 * nXscale
            nn = nn + 1
        End While
        ''// Draw Seam Tram Lines
        ''//
        ARMDIA1.PR_SetLayer("Notes")
        ''AddEntity("line", xyO.X, xyO.Y + nSeam, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5)
        PR_MakeXY(xyStart, xyO.X, xyO.Y + nSeam)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5)
        PR_DrawLine(xyStart, xyEnd)
        ''AddEntity("line", xyO.X, xyO.Y + nSeam + 0.5, xyPt1.X - (1.5 * nXscale), xyPt1.Y)
        PR_MakeXY(xyStart, xyO.X, xyO.Y + nSeam + 0.5)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y)
        PR_DrawLine(xyStart, xyEnd)

        ''--------------File Name : WHLG2DWG.D --------------------
        ''// Load template data file
        If (nFabricClass = 0) Then
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "MMHG.DAT"
        Else
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "DS.DAT"
        End If
        hChan = FreeFile()
        '' hChan = Open("file", sFile, "readonly")
        FileOpen(hChan, sFile, VB.OpenMode.Input)
        If (hChan) Then
            sLine = LineInput(hChan)
            nNo = WHBODDIA1.fnGetNumber(sLine, 1)
            sScale = FN_GetString(sLine, 2)
            nSpace = WHBODDIA1.fnGetNumber(sLine, 3)
            n20Len = WHBODDIA1.fnGetNumber(sLine, 4)
            nReduction = WHBODDIA1.fnGetNumber(sLine, 5)
        Else
            MsgBox("Can't open " + sFile + "\nCheck installation", 16, "Waist Leg - Dialog")
            Exit Sub
        End If
        ''// From data file retain reduction data of first scale for use with
        ''// Straight toes (where a 14 reduction Is required)
        Dim nFirst20Len, nFirstReduction, nHeelTape, nAnkleToHeelOffset, nHt, nTmpltReduction As Double
        nFirst20Len = n20Len
        nFirstReduction = nReduction
        nHeelTape = 6
        ''// Establish type of heel given the ankle position
        ''// NB this comes via WH_FIGUR.D And the Data Base Field AnkleTape 
        Dim SmallHeel As Boolean
        If (nAnkleTape = 7) Then
            SmallHeel = True
            nAnkleToHeelOffset = 1
        Else
            SmallHeel = False
            nAnkleToHeelOffset = 2
        End If
        nHt = xyOtemplate.X
        '' // Loop Through data file to get to start of leg tapes
        ''// Ignoring the none relevent ones 
        nn = 1
        While (nn < nFirstTape)
            sLine = LineInput(hChan)
            nn = nn + 1
        End While
        ScanLine(sLine, nNo, sScale, nSpace, n20Len, nTmpltReduction)
        nSpace = 0
        Dim nLastTape20Len, nLastTapeTmpltReduction, nAnkleMTape, nAnkleM20Len, nAnkleMTmpltReduction As Double
        Dim nAnkleMPrev20Len, nAnkleMPrevTmpltReduction, nAnkleMPrevHt, nPrev20Len, nPrevTmpltReduction As Double
        Dim nHeel20Len, nHeelTmpltReduction, nAnkle20Len, nAnkleTmpltReduction, nPrevHt As Double
        Dim xyFirstTape, xyAnkleM, xyHeel, xyAnkle As XY
        While (nn <= nAnkleTape)
            If ((nn = nFirstTape + 1) And (nFootPleat1 <> 0)) Then
                nSpace = nFootPleat1
            End If
            If ((nn = nFirstTape + 2) And (nFootPleat2 <> 0)) Then
                nSpace = nFootPleat2
            End If
            nHt = nHt + nSpace
            If (nn = nFirstTape) Then
                xyFirstTape.X = nHt
                nLastTape20Len = n20Len
                nLastTapeTmpltReduction = nTmpltReduction
            End If
            If (nn = nHeelTape - nAnkleToHeelOffset) Then
                xyAnkleM.X = nHt
                nAnkleMTape = nHeelTape - nAnkleToHeelOffset
                nAnkleM20Len = n20Len
                nAnkleMTmpltReduction = nTmpltReduction
                If (nn = nFirstTape) Then
                    nAnkleMPrev20Len = n20Len
                    nAnkleMPrevTmpltReduction = nTmpltReduction
                    nAnkleMPrevHt = nHt
                Else
                    nAnkleMPrev20Len = nPrev20Len
                    nAnkleMPrevTmpltReduction = nPrevTmpltReduction
                    nAnkleMPrevHt = nPrevHt
                End If
            End If
            If (nn = nHeelTape) Then
                xyHeel.X = nHt
                nHeel20Len = n20Len
                nHeelTmpltReduction = nTmpltReduction
            End If
            If (nn = nAnkleTape) Then
                xyAnkle.X = nHt
                nAnkle20Len = n20Len
                nAnkleTmpltReduction = nTmpltReduction
            End If
            sLine = LineInput(hChan)
            nPrevHt = nHt   ''// Store wrt AnklePrevM  & Foot Pleats
            nPrev20Len = n20Len
            nPrevTmpltReduction = nTmpltReduction
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nTmpltReduction)
            nn = nn + 1
        End While
        FileClose(hChan)

        '' // Foot Reductions Chart
        ''// Quick And Dirty
        Dim nHeelChartReduction, nAnkleMChartReduction, nAnkleMPrevChartReduction, nToeCutBackReduction, nToeCurvedReduction As Double
        Dim nReductionAnkle As Double = Val(Mid(sReductionAnkle, (nn * 3) + 1, 3))
        If (nReductionAnkle >= 14 And nReductionAnkle <= 18) Then
            nHeelChartReduction = 17
            nAnkleMChartReduction = 14
        End If
        If (nReductionAnkle >= 19 And nReductionAnkle <= 23) Then
            nHeelChartReduction = 19
            nAnkleMChartReduction = 17
        End If
        If (nReductionAnkle >= 24) Then
            nHeelChartReduction = 21
            nAnkleMChartReduction = 17
        End If
        nAnkleMPrevChartReduction = 16
        nToeCutBackReduction = 34
        nToeCurvedReduction = 34

        '' // Choose heel plate
        Dim nHeelR1, nHeelR2, nHeelR3, nHeelD1, nHeelD2, nBackHeelOff, nFrontHeelOff As Double
        If (SmallHeel) Then
            ''// Heel plate #1
            nHeelR1 = 1.1760709
            nHeelR2 = 0.8716718
            nHeelR3 = 1.1760709 '' //Heel Plate now symetrical 27.Sept.94
            nHeelD1 = 2.096372
            nHeelD2 = 2.1843981
            nBackHeelOff = 0.05
            nFrontHeelOff = 0.1875
        Else
            ''// Heel plate #2
            nHeelR1 = 1.6601037
            nHeelR2 = 1.0876233
            nHeelR3 = 1.5053069
            nHeelD1 = 2.7812341
            nHeelD2 = 2.6294854
            nBackHeelOff = 0.25
            nFrontHeelOff = nBackHeelOff
        End If

        '' // Calculate Control points
        ''// For Heel
        ''// At Ankle (ie +3 Or + 1.5 Tape)
        Dim nRedStep, nAnkleMTapeLen As Double
        ''nTapeLen = FNRound(FNDecimalise(FNGetTape(nAnkleTape)) * nUnitsFac) ; 
        nTapeLen = FNDecimalise(FNGetTape(nAnkleTape)) * WHRightLeg1.g_nUnitsFac
        If (nFabricClass = 2) Then
            nLength = (nTapeLen * ((100 - nReductionAnkle) / 100)) / 2
        Else
            nLength = nAnkle20Len / 20 * nTapeLen
            nRedStep = nAnkle20Len / (20 * 8)
            nLength = nLength + ((nAnkleTmpltReduction - nReductionAnkle) * nRedStep)
        End If
        xyAnkle.Y = xyOtemplate.Y + nSeam + nLength

        ''// At Heel (ie at 0 Tape)

        ''nTapeLen = FNRound(FNDecimalise(FNGetTape(nHeelTape)) * nUnitsFac) ; 
        nTapeLen = FNDecimalise(FNGetTape(nHeelTape)) * WHRightLeg1.g_nUnitsFac
        nLength = nHeel20Len / 20 * nTapeLen
        nRedStep = nHeel20Len / (20 * 8)
        nLength = nLength + ((nHeelTmpltReduction - nHeelChartReduction) * nRedStep)
        xyHeel.Y = xyOtemplate.Y + nSeam + nLength

        ''// At Minus Ankle ( ie -3 Or -1.5 Tape)
        ''nAnkleMTapeLen = FNRound(FNDecimalise(FNGetTape(nAnkleMTape)) * nUnitsFac)
        nAnkleMTapeLen = FNDecimalise(FNGetTape(nAnkleMTape)) * WHRightLeg1.g_nUnitsFac
        nLength = nAnkleM20Len / 20 * nAnkleMTapeLen
        nRedStep = nAnkleM20Len / (20 * 8)
        nLength = nLength + ((nAnkleMTmpltReduction - nAnkleMChartReduction) * nRedStep)
        xyAnkleM.Y = xyOtemplate.Y + nSeam + nLength

        Dim xyAnkleMPrev, xyHeelCntrMidDistal As XY
        xyAnkleMPrev.X = nAnkleMPrevHt
        nLength = nAnkleMPrev20Len / 20 * nAnkleMTapeLen   '' // NB "nAnkleMTapeLen" from above
        nRedStep = nAnkleMPrev20Len / (20 * 8)
        nLength = nLength + ((nAnkleMPrevTmpltReduction - nAnkleMPrevChartReduction) * nRedStep)
        xyAnkleMPrev.Y = xyOtemplate.Y + nSeam + nLength

        ''// At Last tape using Minus Ankle Tape Value
        ''// N.B. Quick And dirty
        nLength = nAnkleMTapeLen * 0.66 / 2   ''// NB "nAnkleMTapeLen" from above
        xyFirstTape.Y = xyOtemplate.Y + nLength + nSeam

        ''// Heel Circles
        ''// Control points

        xyHeelCntrMidDistal.X = xyHeel.X - nFrontHeelOff
        xyHeelCntrMidDistal.Y = xyHeel.Y - nHeelR2
        Dim nCalc, nAngle, nToeOffset As Double
        Dim nError As Short
        Dim xyHeelCntrDistal, xyHeelCntrMidProximal, xyHeelCntrProximal As XY
        If (SmallHeel) Then
            nCalc = FN_CalcAngle(xyAnkleM, xyHeelCntrMidDistal)
            nAngle = FN_CalcAngle(xyHeelCntrMidDistal, xyAnkleM) - System.Math.Acos((nHeelR1 ^ 2 - nHeelD1 ^ 2 - nCalc ^ 2) / (-2 * nHeelD1 * nCalc))
            ''xyHeelCntrDistal = CalcXY("relpolar", xyHeelCntrMidDistal, nHeelD1, nAngle)
            PR_CalcPolar(xyHeelCntrMidDistal, nHeelD1, nAngle, xyHeelCntrDistal)
        Else
            'nError = FN_CirLinInt(xyHeelCntrMidDistal.X, xyAnkleM.Y + nHeelR1,
            '           xyHeelCntrMidDistal.X - 10, xyAnkleM.Y + nHeelR1,
            '           xyHeelCntrMidDistal,
            '                nHeelD1)
            PR_MakeXY(xyStart, xyHeelCntrMidDistal.X, xyAnkleM.Y + nHeelR1)
            PR_MakeXY(xyEnd, xyHeelCntrMidDistal.X - 10, xyAnkleM.Y + nHeelR1)
            nError = FN_CirLinInt(xyStart, xyEnd, xyHeelCntrMidDistal, nHeelD1, xyInt)
            xyHeelCntrDistal = xyInt
        End If

        xyHeelCntrMidProximal.X = xyHeel.X + nBackHeelOff
        xyHeelCntrMidProximal.Y = xyHeel.Y - nHeelR2

        Dim BigAnkle As Boolean = False    ''//Flag to indicate Ankle of Lymphdema proportions
        If (SmallHeel) Then
            nCalc = FN_CalcLength(xyAnkle, xyHeelCntrMidProximal)
            nAngle = FN_CalcAngle(xyHeelCntrMidProximal, xyAnkle) + System.Math.Acos((nHeelR3 ^ 2 - nHeelD2 ^ 2 - nCalc ^ 2) / (-2 * nHeelD2 * nCalc))
            ''xyHeelCntrProximal = CalcXY("relpolar", xyHeelCntrMidProximal, nHeelD2, nAngle)
            PR_CalcPolar(xyHeelCntrMidProximal, nHeelD2, nAngle, xyHeelCntrProximal)
        Else
            'nError = FN_CirLinInt(xyHeelCntrMidProximal.x, xyAnkle.Y + nHeelR3,
            '           xyHeelCntrMidProximal.x + 10, xyAnkle.Y + nHeelR3,
            '           xyHeelCntrMidProximal,
            '          nHeelD2)
            PR_MakeXY(xyStart, xyHeelCntrMidProximal.X, xyAnkle.Y + nHeelR3)
            PR_MakeXY(xyEnd, xyHeelCntrMidProximal.X + 10, xyAnkle.Y + nHeelR3)
            nError = FN_CirLinInt(xyStart, xyEnd, xyHeelCntrMidProximal, nHeelD2, xyInt)
            xyHeelCntrProximal = xyInt
            If (nError = False) Then
                BigAnkle = True     ''// IE no intersection found
            End If
        End If

        ''// Toes
        ''// Quick And dirty ;
        Dim xyToeSeam As XY
        xyToeSeam.Y = xyOtemplate.Y + nSeam
        Dim sFootLabel As String = " "
        Dim nAge As Double = Val(txtAge.Text)
        If (nAge <= 10) Then
            If (nAge <= 2) Then
                nToeOffset = 2.75
            End If
            If (nAge = 3) Then
                nToeOffset = 3.0 ''	// Fax 9.Sept.94, Item2
            End If
            If (nAge = 4) Then
                nToeOffset = 3.25
            End If
            If (nAge = 5) Then
                nToeOffset = 3.375
            End If
            If (nAge = 6) Then
                nToeOffset = 3.625
            End If
            If (nAge = 7) Then
                nToeOffset = 3.875
            End If
            If (nAge = 8) Then
                nToeOffset = 4.0
            End If
            If (nAge = 9) Then
                nToeOffset = 4.25
            End If
            If (nAge = 10) Then
                nToeOffset = 4.5
            End If
        End If

        Dim nOtherAnkleMTapeLen As Double
        If (sToeStyle.Equals("Curved")) Then
            If (txtSex.Text = "Male") Then
                xyToeSeam.X = xyHeel.X - 4.75
            Else
                xyToeSeam.X = xyHeel.X - 4.5
            End If
            If (nFirstTape <= 2 Or nAnkleMTapeLen >= 10 Or nOtherAnkleMTapeLen >= 10) Then
                ''//Figuring as a LONG Curved Toe
                xyToeSeam.X = xyHeel.X - 5.5
            End If
            If (nAge <= 10) Then
                ''// Figuring as a EXTRA SHORT  Curved Toe"
                xyToeSeam.X = xyHeel.X - nToeOffset
            End If
        End If
        If (sToeStyle.Equals("Cut-Back") Or sToeStyle.Equals("Soft Enclosed")) Then
            If (txtSex.Text = "Male") Then
                xyToeSeam.X = xyHeel.X - 4.75
            Else
                xyToeSeam.X = xyHeel.X - 4.5
            End If
            If (nAnkleMTapeLen >= 10 Or nOtherAnkleMTapeLen >= 10) Then
                xyToeSeam.X = xyHeel.X - 5.5
            End If
        End If

        If (sToeStyle.Equals("Soft Enclosed")) Then
            sFootLabel = "CAP"
            If (nAge <= 10) Then
                xyToeSeam.X = xyHeel.X - nToeOffset
            End If
        End If
        Dim xyToeOFF As XY
        If (sToeStyle.Equals("Straight") Or sToeStyle.Equals("Soft Enclosed B/M")) Then
            xyToeSeam.X = xyHeel.X - 3.5
            nLength = nFirst20Len / 20 * nAnkleMTapeLen
            nRedStep = nFirst20Len / (20 * 8)
            nLength = nLength + ((nFirstReduction - 14) * nRedStep)
            xyToeOFF.Y = xyOtemplate.Y + nSeam + nLength
            xyToeOFF.X = xyToeSeam.X
            If (sToeStyle.Equals("Soft Enclosed B/M")) Then
                sFootLabel = "CAP"
            End If
        End If
        Dim sFootLength As String = txtFootLength.Text
        Dim nFootLength As Short
        If (sToeStyle.Equals("Self Enclosed") Or (sToeStyle.Equals("Soft Enclosed") And Val(sFootLength))) Then
            If (sToeStyle.Equals("Self Enclosed")) Then
                sFootLabel = "ENCLOSED"
            Else
                sFootLabel = "CAP"
            End If
            ''If (nFootLength = FNRound(FNDecimalise(Value("scalar", sFootLength)) * nUnitsFac)) Then{
            nFootLength = FNDecimalise(Val(sFootLength)) * WHLEGDIA1.g_nUnitsFac
            If (nFootLength <> 0) Then
                If (SmallHeel) Then
                    ''nFootLength = FNRound(nFootLength * 0.9)
                    nFootLength = nFootLength * 0.9
                Else
                    ''nFootLength = FNRound(nFootLength * 0.83)
                    nFootLength = nFootLength * 0.83
                End If
                nLength = (nHeel20Len / 20 * 12.5) + xyOtemplate.Y + nSeam
                If (nLength < xyHeel.Y) Then
                    'nError = FN_CirLinInt(xyOtemplate.X - 20, xyOtemplate.Y + nSeam,
                    '   xyHeel.X, xyOtemplate.Y + nSeam,
                    '          xyHeel.X, nLength,
                    '          nFootLength)
                    PR_MakeXY(xyStart, xyOtemplate.X - 20, xyOtemplate.Y + nSeam)
                    PR_MakeXY(xyEnd, xyHeel.X, xyOtemplate.Y + nSeam)
                    PR_MakeXY(xyText, xyHeel.X, nLength)
                    nError = FN_CirLinInt(xyStart, xyEnd, xyText, nFootLength, xyInt)
                Else
                    'nError = FN_CirLinInt(xyOtemplate.X - 20, xyOtemplate.Y + nSeam,
                    '                   xyHeel.X, xyOtemplate.Y + nSeam,
                    '                          xyHeel,
                    '                          nFootLength)
                    PR_MakeXY(xyStart, xyOtemplate.X - 20, xyOtemplate.Y + nSeam)
                    PR_MakeXY(xyEnd, xyHeel.X, xyOtemplate.Y + nSeam)
                    nError = FN_CirLinInt(xyStart, xyEnd, xyHeel, nFootLength, xyInt)
                End If

                If (nError) Then
                    xyToeSeam = xyInt
                Else
                    MsgBox("Can't position Toe with given Foot length", 16, "Waist Leg - Dialog")
                    MsgBox("Error forming foot", 16, "Waist Leg - Dialog")
                    Exit Sub
                End If
            Else
                MsgBox("A Foot length is required for Self Enclosed Toes", 16, "Waist Leg - Dialog")
                MsgBox("No Foot length given", 16, "Waist Leg - Dialog")
                Exit Sub
            End If
        End If

        ''  // Toe Points to position toe arcs
        ''// Toe circle constants
        ''// Quick And Dirty
        Dim nToeCntrMidToCntrLowY, nToeCntrMidToCntrLowX, nToeCntrMidToCntrHighY, nToeCntrMidToCntrHighX, nToeMidR, nToeLowR, nToeHighR As Double
        nToeCntrMidToCntrLowY = 3.3359296
        nToeCntrMidToCntrLowX = 8.0276805
        nToeCntrMidToCntrHighY = 3.1684888
        nToeCntrMidToCntrHighX = 2.1960337
        nToeMidR = 0.2528866
        nToeLowR = 8.4403592
        nToeHighR = 4.1080042

        ''// Establish Low Toe Arc center

        nLength = xyFirstTape.Y - xyToeSeam.Y       ''// Toe Point line To seam
        Dim xyToeCntrLow, xyToeCntrMid, xyToeCntrHigh, xyToePnt As XY
        If (nToeCntrMidToCntrLowY < nLength) Then
            ''// Simple case
            ''// End of toe curve Is above seam line
            xyToeCntrLow.Y = xyFirstTape.Y - nToeCntrMidToCntrLowY
            xyToeCntrLow.X = xyToeSeam.X - nToeLowR
        Else
            ''// More Complex case
            ''// End of toe curve Is intersected by the seam line
            xyToeCntrLow.Y = xyFirstTape.Y - nToeCntrMidToCntrLowY
            nLength = nToeCntrMidToCntrLowY - nLength
            nLength = System.Math.Sqrt((nToeLowR * nToeLowR) - (nLength * nLength))
            xyToeCntrLow.X = xyToeSeam.X - nLength
        End If

        ''// Having established TOE Low circle center the rest follows 
        ''//
        xyToeCntrMid.Y = xyToeCntrLow.Y + nToeCntrMidToCntrLowY
        xyToeCntrMid.X = xyToeCntrLow.X + nToeCntrMidToCntrLowX
        xyToeCntrHigh.Y = xyToeCntrMid.Y - nToeCntrMidToCntrHighY
        xyToeCntrHigh.X = xyToeCntrMid.X + nToeCntrMidToCntrHighX


        ARMDIA1.PR_SetLayer("Construct")
        Dim ptColl As Point3dCollection = New Point3dCollection
        Dim aPrevAngle, aAngleInc As Double
        Dim xyPlr As XY
        ''If (FootLess = False) Then
        ''// Draw for Straight & Straight types
        If (sToeStyle.Equals("Straight") Or sToeStyle.Equals("Soft Enclosed B/M")) Then
            ''// Draw for straight & Straight types
            ''StartPoly("fitted")
            ''AddVertex(xyToeOFF)
            ptColl.Add(New Point3d(xyToeOFF.X, xyToeOFF.Y, 0))
        Else
            ''// Draw Toe Curve
            ''StartPoly("fitted")
            ''AddVertex(xyToeSeam)
            ptColl.Add(New Point3d(xyToeSeam.X, xyToeSeam.Y, 0))
            ''// First toe curve
            aAngle = FN_CalcAngle(xyToeCntrLow, xyToeCntrMid)
            aPrevAngle = FN_CalcAngle(xyToeCntrLow, xyToeSeam)
            If (aAngle > aPrevAngle) Then
                aAngleInc = (aAngle - aPrevAngle) / 3
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
            End If
            ii = 1
            While (ii <= 3)
                ''AddVertex(CalcXY("relpolar", xyToeCntrLow, nToeLowR, aPrevAngle + aAngleInc * ii))
                PR_CalcPolar(xyToeCntrLow, nToeLowR, aPrevAngle + aAngleInc * ii, xyPlr)
                ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
                ii = ii + 1
            End While
            ''// Toe Point curve
            aAngle = FN_CalcAngle(xyToeCntrMid, xyToeCntrLow)
            aPrevAngle = FN_CalcAngle(xyToeCntrHigh, xyToeCntrMid)
            If (aAngle > aPrevAngle) Then
                aAngleInc = (aAngle - aPrevAngle) / 3
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
            End If
            ii = 1
            ''xyToePnt = CalcXY("relpolar", xyToeCntrMid, nToeMidR, aAngle - aAngleInc * ii)
            PR_CalcPolar(xyToeCntrMid, nToeMidR, aAngle - aAngleInc * ii, xyToePnt)
            While (ii <= 3)
                ''AddVertex(CalcXY("relpolar", xyToeCntrMid, nToeMidR, aAngle - aAngleInc * ii))
                PR_CalcPolar(xyToeCntrMid, nToeMidR, aAngle - aAngleInc * ii, xyPlr)
                ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
                ii = ii + 1
            End While

            ''// Top of toe curve
            aAngle = FN_CalcAngle(xyToeCntrHigh, xyToeCntrMid)
            aPrevAngle = FN_CalcAngle(xyToeCntrHigh, xyAnkleMPrev)
            If (aPrevAngle <= 90) Then aPrevAngle = 90
            If (aAngle > aPrevAngle) Then
                aAngleInc = (aAngle - aPrevAngle) / 3
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
            End If
            ii = 1
            While (ii <= 2)
                ''AddVertex(CalcXY("relpolar", xyToeCntrHigh, nToeHighR, aAngle - aAngleInc * ii))
                PR_CalcPolar(xyToeCntrHigh, nToeHighR, aAngle - aAngleInc * ii, xyPlr)
                ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
                ii = ii + 1
            End While
            ''AddVertex(xyAnkleMPrev)
            ptColl.Add(New Point3d(xyAnkleMPrev.X, xyAnkleMPrev.Y, 0))
        End If

        '' // Draw Heel
        ''// Add Start of heel
        If (xyAnkleMPrev.Y <> 0 And SmallHeel = False) Then
            ''AddVertex(xyAnkleM)
            ptColl.Add(New Point3d(xyAnkleM.X, xyAnkleM.Y, 0))
        End If

        aPrevAngle = 270
        aAngle = FN_CalcAngle(xyHeelCntrDistal, xyHeelCntrMidDistal)
        If (aAngle > aPrevAngle) Then
            aAngleInc = (aAngle - aPrevAngle) / 3
        Else
            aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
        End If
        ii = 1
        While (ii <= 2)
            ''AddVertex(CalcXY("relpolar", xyHeelCntrDistal, nHeelR1, aPrevAngle + aAngleInc * ii))
            PR_CalcPolar(xyHeelCntrDistal, nHeelR1, aPrevAngle + aAngleInc * ii, xyPlr)
            ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
            ii = ii + 1
        End While

        aPrevAngle = 90
        aAngle = FN_CalcAngle(xyHeelCntrMidDistal, xyHeelCntrDistal)
        If (aAngle > aPrevAngle) Then
            aAngleInc = (aAngle - aPrevAngle) / 3
        Else
            aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
        End If
        ii = 1
        While (ii <= 3)
            ''AddVertex(CalcXY("relpolar", xyHeelCntrMidDistal, nHeelR2, aAngle - aAngleInc * ii))
            PR_CalcPolar(xyHeelCntrMidDistal, nHeelR2, aAngle - aAngleInc * ii, xyPlr)
            ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
            ii = ii + 1
        End While

        If (SmallHeel = False Or BigAnkle = True) Then
            ''AddVertex(xyHeel)
            ptColl.Add(New Point3d(xyHeel.X, xyHeel.Y, 0))
        End If

        aAngle = 90
        aPrevAngle = FN_CalcAngle(xyHeelCntrMidProximal, xyHeelCntrProximal)
        If (aAngle > aPrevAngle) Then
            aAngleInc = (aAngle - aPrevAngle) / 3
        Else
            aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
        End If
        Dim nPts As Double
        If (BigAnkle) Then
            nPts = 0
        Else
            nPts = 2
        End If
        ii = 0
        While (ii <= nPts)
            ''AddVertex(CalcXY("relpolar", xyHeelCntrMidProximal, nHeelR2, aAngle - aAngleInc * ii))
            PR_CalcPolar(xyHeelCntrMidProximal, nHeelR2, aAngle - aAngleInc * ii, xyPlr)
            ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
            ii = ii + 1
        End While

        aAngle = 270
        aPrevAngle = FN_CalcAngle(xyHeelCntrProximal, xyHeelCntrMidProximal)
        If (aAngle > aPrevAngle) Then
            aAngleInc = (aAngle - aPrevAngle) / 3
        Else
            aAngleInc = ((aAngle + 360) - aPrevAngle) / 3
        End If
        If (BigAnkle) Then
            nPts = 0
        Else
            nPts = 2
        End If
        ii = 1
        While (ii <= nPts)
            ''xyTmp = CalcXY("relpolar", xyHeelCntrProximal, nHeelR3, aPrevAngle + aAngleInc * ii)
            PR_CalcPolar(xyHeelCntrProximal, nHeelR3, aPrevAngle + aAngleInc * ii, xyTmp)
            If (xyTmp.X < xyAnkle.X) Then
                ''AddVertex(xyTmp)
                ptColl.Add(New Point3d(xyTmp.X, xyTmp.Y, 0))
            End If
            ii = ii + 1
        End While
        ''Else

        ''// Draw as a panty leg
        ''StartPoly("fitted") ;
        ''End If

        ''//
        ''// Draw Leg And Back Body Profile
        ''//
        ''// Load template data file
        If (nFabricClass = 0) Then
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "MMHG.DAT"
        Else
            sFile = fnGetSettingsPath("LookupTables") + "\\WH" + StringMiddle(sPressure, 1, 2) + "DS.DAT"
        End If
        hChan = FreeFile()
        ''hChan = Open("file", sFile, "readonly")
        FileOpen(hChan, sFile, VB.OpenMode.Input)
        If (hChan) Then
            sLine = LineInput(hChan)
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
        Else
            MsgBox("Can't open " + sFile + "\nCheck installation", 16, "Waist Leg - Dialog")
            Exit Sub
        End If

        ''// Skip to FirstTape
        nn = 1
        While (nn < nFirstTape)
            sLine = LineInput(hChan)
            nn = nn + 1
        End While

        ''// Skip to ankle tape 
        xyTmp = xyOtemplate
        While (nn < nAnkleTape)
            sLine = LineInput(hChan)
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
            If ((nn = nFirstTape) And (nFootPleat1 <> 0)) Then
                nSpace = nFootPleat1
            End If
            If ((nn = nFirstTape + 1) And (nFootPleat2 <> 0)) Then
                nSpace = nFootPleat2
            End If
            xyTmp.X = xyTmp.X + nSpace
            nn = nn + 1
        End While

        If (FootLess) Then
            nn = nFirstTape
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
        Else
            nn = nAnkleTape
        End If
        nn = nAnkleTape
        Dim xyLastTape, xyPt2 As XY
        While (nn <= nLastTape)
            ''nTapeLen = FNRound(FNDecimalise(FNGetTape(nn)) * nUnitsFac)
            nTapeLen = FNDecimalise(FNGetTape(nn)) * WHRightLeg1.g_nUnitsFac
            If (nFabricClass = 2) Then
                nLength = (nTapeLen * (100 - Val(StringMiddle(sReductionAnkle, ((nn - 1) * 3) + 1, 3))) / 100) / 2
            Else
                nLength = n20Len / 20 * nTapeLen
            End If
            nRedStep = n20Len / (20 * 8)
            If ((nn = nFirstTape) And FootLess) Then
                ''// Release the distal tape  to a 95% reduction,  92% reduction if +3 Or +1-1/2 tape
                If (nFirstTape > 8) Then
                    nLength = (nTapeLen * 0.95) / 2
                Else
                    nLength = (nTapeLen * 0.92) / 2
                End If
                xyLastTape.X = xyOtemplate.X
                xyLastTape.Y = xyOtemplate.Y + nSeam + nLength
            End If
            If (nn = nAnkleTape And nFabricClass <> 2) Then
                ''// Release the ANKLE tape to the CALCULATED reduction
                nLength = nLength + ((nReduction - nReductionAnkle) * nRedStep)
            End If

            If (nn = nLastTape And nFabricClass <> 2) Then
                ''// Release last tape to a 14 reduction
                nLength = nLength + ((nReduction - 14) * nRedStep)
            End If
            xyTmp.Y = xyOtemplate.Y + nSeam + nLength
            If (xyTmp.X >= xyFold.X) Then
                xyPt2 = xyTmp   ''// Store this For later use
                Exit While
            Else
                ''AddVertex(xyTmp)
                ptColl.Add(New Point3d(xyTmp.X, xyTmp.Y, 0))
                xyPt1 = xyTmp   ''// Store this For later use
            End If
            If (nFabricClass = 2) Then
                ''SetData("TextHorzJust", 2) ''	// Center
                ''sSymbol = MakeString("long", nNo) + "tape"
                Dim sSymbol As String = Str(nNo) + "tape"
                'If (!Symbol("find", sSymbol)) Then
                '    MsgBox("Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists", 16, "Waist Leg - Dialog")
                '    Exit Sub
                'End If
                'AddEntity("symbol", sSymbol, xyTmp)
                Dim strLab As String = Label1(nn - 1).Text
                If nNo = 0 Then
                    strLab = "Pleat"
                End If
                PR_DrawRuler(sSymbol, xyPt1, strLab)
                ''AddEntity("text", Format("length", nTapeLen), xyTmp.X, xyTmp.Y - 0.5)
                PR_MakeXY(xyText, xyTmp.X, xyTmp.Y - 0.5)
                PR_DrawText(WHRightLeg1.fnInchesToText(nTapeLen), xyText, 0.1, 0, 1)
                ''AddEntity("text", StringMiddle(sReduction, ((nn - 1) * 3) + 1, 3), xyTmp.X, xyTmp.Y - 0.7)
                PR_MakeXY(xyText, xyTmp.X, xyTmp.Y - 0.7)
                PR_DrawText(StringMiddle(sReductionAnkle, ((nn - 1) * 3) + 1, 3), xyText, 0.1, 0, 1)
                ''AddEntity("text", StringMiddle(sStretch, ((nn - 1) * 3) + 1, 3), xyTmp.X, xyTmp.Y - 0.9)
                PR_MakeXY(xyText, xyTmp.X, xyTmp.Y - 0.9)
                PR_DrawText(StringMiddle(sGramsAnkle, ((nn - 1) * 3) + 1, 3), xyText, 0.1, 0, 1)
                ''AddEntity("text", StringMiddle(sTapeMMs, ((nn - 1) * 3) + 1, 3), xyTmp.X, xyTmp.Y - 1.1)
                PR_MakeXY(xyText, xyTmp.X, xyTmp.Y - 1.1)
                PR_DrawText(StringMiddle(sMMAnkle, ((nn - 1) * 3) + 1, 3), xyText, 0.1, 0, 1)
            End If

            nn = nn + 1
            sLine = LineInput(hChan)
            ScanLine(sLine, nNo, sScale, nSpace, n20Len, nReduction)
            If ((nn = nFirstTape + 1) And (nFootPleat1 <> 0)) Then
                nSpace = nFootPleat1
            End If
            If ((nn = nFirstTape + 2) And (nFootPleat2 <> 0)) Then
                nSpace = nFootPleat2
            End If
            If ((nn = nLastTape) And (nTopLegPleat1 <> 0)) Then
                nSpace = nTopLegPleat1
            End If
            If ((nn = nLastTape - 1) And (nTopLegPleat2 <> 0)) Then
                nSpace = nTopLegPleat2
            End If
            xyTmp.X = xyTmp.X + nSpace
        End While
        FileClose(hChan)

        ''// check if we use last leg tape Or the fold position
        If (nLegStyle = 1) Then
            ''AddVertex(xyFoldPanty) '' // We must hit this point For panties
            ptColl.Add(New Point3d(xyFoldPanty.X, xyFoldPanty.Y, 0))
        ElseIf (xyTmp.Y < xyFold.Y) Then
            ''AddVertex(xyFold)
            ptColl.Add(New Point3d(xyFold.X, xyFold.Y, 0))
        Else
            '' AddVertex(xyTmp)
            ptColl.Add(New Point3d(xyTmp.X, xyTmp.Y, 0))
        End If

        ''AddVertex(xyLargest) 
        ptColl.Add(New Point3d(xyLargest.X, xyLargest.Y, 0))

        ''// Loop through first leg data file looking for changes that have been made to the curve
        ''// after the largest part of buttocks on the first drawn body 
        ''// nTrans And aTransAngle from WHLG2DBD.D
        Dim hFileCurve As Object
        ''hFileCurve = Open("file", "C:\\JOBST\\LEGCURVE.DAT", "readonly")
        hFileCurve = FreeFile()
        FileOpen(hFileCurve, fnGetSettingsPath("LookupTables") + "\\LEGCURVE.DAT", VB.OpenMode.Input)
        ii = 1
        Dim AddPoint As Boolean = False
        Dim nX, nY As Double
        Dim nLegVertexCount As Double = ptFstLegColl.Count
        While (ii <= nLegVertexCount)
            sLine = LineInput(hFileCurve)
            '' ScanLine(sLine, "blank", & nX, & nY)
            nX = FN_GetNumber(sLine, 1)
            nY = FN_GetNumber(sLine, 2)
            If (nX > xyLargestFirstBody.X) Then AddPoint = True
            If (AddPoint) Then
                ''AddVertex(CalcXY("relpolar", nX, nY, nTrans, aTransAngle))
                PR_MakeXY(xyStart, nX, nY)
                PR_CalcPolar(xyStart, nTrans, aTransAngle, xyPlr)
                ptColl.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
            End If
            ii = ii + 1
        End While
        FileClose()
        ''PR_DrawPoly(ptColl)
        ARMDIA1.PR_SetLayer(strLayer)
        PR_DrawSpline(ptColl)

        '' // Get polyline entity handle
        ''// Change layer And set DB values
        ''hCurv = UID("find", UID("getmax"))
        '' SetEntityData(hCurv, "layer", hTemplateLayer)
        ''SetDBData(hCurv, "ID", sID + "LegCurve")

        If (nTOSCir = 0) Then
            ''Execute("menu", "SetLayer", hTemplateLayer)
            ARMDIA1.PR_SetLayer(strLayer)
            '' AddEntity("line", xyWaist, xyTOS)
            PR_DrawLine(xyWaist, xyTOS)
        End If
        If (FootLess = False) Then
            ''// Draw foot points
            ARMDIA1.PR_SetLayer("Construct")
            ''hEnt = AddEntity("marker", "xmarker", xyAnkle, 0.2, 0.2) ;
            ''SetDBData(hEnt, "ID", sID + "Ankle");
            ''sTmp = MakeString("scalar", xyHeelCntrProximal.X - xyAnkle.X) + " " + MakeString("scalar", xyHeelCntrProximal.Y - xyAnkle.Y)  ;
            ''SetDBData(hEnt, "Data", sTmp);
            PR_DrawXMarker(xyAnkle)
            ''hEnt = AddEntity("marker", "xmarker", xyHeel, 0.2, 0.2) ;
            ''SetDBData(hEnt, "ID", sID + "Heel");
            ''SetDBData(hEnt, "Data", MakeString("long", SmallHeel));
            PR_DrawXMarker(xyHeel)
            ''hEnt = AddEntity("marker", "xmarker", xyAnkleM, 0.2, 0.2) ;
            ''SetDBData(hEnt, "ID", sID + "AnkleM");	
            ''sTmp = MakeString("scalar", xyHeelCntrDistal.X - xyAnkleM.X) + " " + MakeString("scalar", xyHeelCntrDistal.Y - xyAnkleM.Y)  ;
            ''SetDBData(hEnt, "Data", sTmp);
            PR_DrawXMarker(xyAnkleM)
            If (nAge > 10) Then
                ''AddEntity ("marker", "xmarker", xyAnkleMPrev, 0.2, 0.2) 
                PR_DrawXMarker(xyAnkleMPrev)
            End If
            If (SmallHeel) Then
                ''AddEntity("arc", xyHeelCntrProximal, nHeelR3, 180, 90) 
                PR_DrawArc(xyHeelCntrProximal, nHeelR3, 180, 90)
            End If

            ''// Add Fold point for use in editor
            If (nFabricClass = 2) Then
                ''hEnt = AddEntity("marker", "xmarker", xyFold, 0.1, 0.1) ;	
                ''SetDBData(hEnt, "ID", sID + "Fold")
                PR_DrawXMarker(xyFold)
            End If

            ''// Draw rest of it
            ''Execute("menu", "SetLayer", hTemplateLayer)
            ARMDIA1.PR_SetLayer(strLayer)

            ''// Add Closing lines at TOE
            ''AddEntity("line", xyO, xyToeSeam.X, xyOtemplate.Y)
            PR_MakeXY(xyEnd, xyToeSeam.X, xyOtemplate.Y)
            PR_DrawLine(xyO, xyEnd)
            ''AddEntity("line", xyToeSeam.X, xyO.Y, xyToeSeam)
            PR_MakeXY(xyStart, xyToeSeam.X, xyO.Y)
            PR_DrawLine(xyStart, xyToeSeam)

            ''// Toe endings
            Dim xyToeCL As XY
            xyToeCL.X = xyToeSeam.X
            xyToeCL.Y = xyOtemplate.Y

            If (sToeStyle.Equals("Soft Enclosed") And nFootLength = 0 And (nAge <= 10)) Then
                ''AddEntity("line", xyToeSeam, xyToeSeam.X, xyAnkleM.Y)
                PR_MakeXY(xyEnd, xyToeSeam.X, xyAnkleM.Y)
                PR_DrawLine(xyToeSeam, xyEnd)
            End If
            Dim nRightOffset, nLeftOffset As Double
            If (sToeStyle.Contains("Cut-Back") Or (sToeStyle.Equals("Soft Enclosed") And nFootLength = 0 And (nAge > 10))) Then
                nRightOffset = 0.75
                nLeftOffset = 0.25
                xyTmp.X = xyToeCntrMid.X - nToeMidR
                xyTmp.Y = xyToeCntrMid.Y
                aAngle = FN_CalcAngle(xyToeCntrLow, xyTmp)
                aPrevAngle = FN_CalcAngle(xyToeCntrLow, xyToeCL)
                If (aAngle > aPrevAngle) Then
                    aAngleInc = (aAngle - aPrevAngle) / 2
                Else
                    aAngleInc = ((aAngle + 360) - aPrevAngle) / 2
                End If
                '        AddEntity("poly",
                '"openfitted",
                ' xyTmp,
                'CalcXY("relpolar", xyToeCntrLow, nToeLowR - nLeftOffset, aPrevAngle + aAngleInc),
                'xyToeCL)
                PR_CalcPolar(xyToeCntrLow, nToeLowR - nLeftOffset, aPrevAngle + aAngleInc, xyPlr)
                Dim ptColl1 As Point3dCollection = New Point3dCollection
                ptColl1.Add(New Point3d(xyTmp.X, xyTmp.Y, 0))
                ptColl1.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
                ptColl1.Add(New Point3d(xyToeCL.X, xyToeCL.Y, 0))
                PR_DrawPoly(ptColl1)
                ARMDIA1.PR_SetLayer("Notes")
                '        AddEntity("poly",
                '"openfitted",
                ' xyTmp,
                'CalcXY("relpolar", xyToeCntrLow, nToeLowR + nRightOffset, aPrevAngle + aAngleInc),
                'xyToeCL)
                PR_CalcPolar(xyToeCntrLow, nToeLowR + nRightOffset, aPrevAngle + aAngleInc, xyPlr)
                Dim ptColl2 As Point3dCollection = New Point3dCollection
                ptColl2.Add(New Point3d(xyTmp.X, xyTmp.Y, 0))
                ptColl2.Add(New Point3d(xyPlr.X, xyPlr.Y, 0))
                ptColl2.Add(New Point3d(xyToeCL.X, xyToeCL.Y, 0))
                PR_DrawPoly(ptColl2)
            End If

            If (sToeStyle.Equals("Soft Enclosed") And nFootLength <> 0) Then
                If (nAge > 10) Then
                    nRightOffset = 2.25
                Else
                    nRightOffset = 1.75
                End If
                sFootLabel = Format("length", nRightOffset) + " Soft Enclosed"
                xyTmp.X = xyToeCntrMid.X - nToeMidR
                xyTmp.Y = xyToeCntrMid.Y
                aAngle = FN_CalcAngle(xyToeCntrLow, xyTmp)
                aPrevAngle = FN_CalcAngle(xyToeCntrLow, xyToeCL)
                If (aAngle > aPrevAngle) Then
                    aAngleInc = (aAngle - aPrevAngle) / 2
                Else
                    aAngleInc = ((aAngle + 360) - aPrevAngle) / 2
                End If
                ''Execute("menu", "SetLayer", hTemplateLayer)
                ARMDIA1.PR_SetLayer(strLayer)
                ''xyTmp = CalcXY("relpolar", xyToeCntrLow, nToeLowR, aPrevAngle + aAngleInc)
                PR_CalcPolar(xyToeCntrLow, nToeLowR, aPrevAngle + aAngleInc, xyTmp)
                ''AddEntity("line", xyTmp.X + nRightOffset,xyTmp.Y,xyTmp.X + nRightOffset,xyAnkleM.Y)
                PR_MakeXY(xyStart, xyTmp.X + nRightOffset, xyTmp.Y)
                PR_MakeXY(xyEnd, xyTmp.X + nRightOffset, xyAnkleM.Y)
                PR_DrawLine(xyStart, xyEnd)
                ''AddEntity("line", xyTmp.X + nRightOffset,xyTmp.Y,xyTmp)
                PR_MakeXY(xyStart, xyTmp.X + nRightOffset, xyTmp.Y)
                PR_DrawLine(xyStart, xyTmp)
            End If


            If (sToeStyle.Equals("Straight") Or sToeStyle.Equals("Soft Enclosed B/M")) Then
                ''AddEntity("Line", xyToeSeam, xyToeOFF)
                PR_DrawLine(xyToeSeam, xyToeOFF)
            Else
                ARMDIA1.PR_SetLayer("Notes")
                ''AddEntity("line", xyToeSeam.X, xyFirstTape.Y, xyToePnt.X, xyFirstTape.Y)
                PR_MakeXY(xyStart, xyToeSeam.X, xyFirstTape.Y)
                PR_MakeXY(xyEnd, xyToePnt.X, xyFirstTape.Y)
            End If
            ''// Foot lable
            ARMDIA1.PR_SetLayer("Notes")
            ''AddEntity("text", sFootLabel, xyToeSeam.X + 1.5, xyToeSeam.Y + 1)
            PR_MakeXY(xyText, xyToeSeam.X + 1.5, xyToeSeam.Y + 1)
            PR_DrawText(sFootLabel, xyText, 0.1, 0, 1)

            ''// Seam Tramlines
            ''// Add Closing lines at TOE
            ''AddEntity("line", xyO.X, xyO.Y + nSeam, xyToeSeam.X, xyO.Y + nSeam)
            PR_MakeXY(xyStart, xyO.X, xyO.Y + nSeam)
            PR_MakeXY(xyEnd, xyToeSeam.X, xyO.Y + nSeam)
            PR_DrawLine(xyStart, xyEnd)
            ''AddEntity("line", xyO.X, xyO.Y + nSeam + 0.5, xyToeSeam.X, xyO.Y + nSeam + 0.5)
            PR_MakeXY(xyStart, xyO.X, xyO.Y + nSeam + 0.5)
            PR_MakeXY(xyEnd, xyToeSeam.X, xyO.Y + nSeam + 0.5)
            PR_DrawLine(xyStart, xyEnd)
        Else
            ''// Panty
            ARMDIA1.PR_SetLayer("Notes")
            ''AddEntity("line", xyOtemplate.X, xyOtemplate.Y + nSeam, xyO.X, xyOtemplate.Y + nSeam)
            PR_MakeXY(xyStart, xyToeSeam.X, xyOtemplate.Y + nSeam)
            PR_MakeXY(xyEnd, xyO.X, xyOtemplate.Y + nSeam)
            PR_DrawLine(xyStart, xyEnd)
            ''AddEntity("line", xyOtemplate.X, xyOtemplate.Y + nSeam + 0.5, xyO.X, xyOtemplate.Y + nSeam + 0.5);
            PR_MakeXY(xyStart, xyToeSeam.X, xyOtemplate.Y + nSeam + 0.5)
            PR_MakeXY(xyEnd, xyO.X, xyOtemplate.Y + nSeam + 0.5)
            PR_DrawLine(xyStart, xyEnd)
            ''//Elastic note
            Dim nElastic As Integer
            Select Case nLegStyle
                Case 3
                    nElastic = 1
                Case 4, 5
                    nElastic = 1
                    If nLegStyle = 5 Then
                        nElastic = 0
                    End If
            End Select
            If (nElastic = 1) Then
                sFootLabel = "ELASTIC"
            End If
            If (nElastic = -1) Then
                sFootLabel = "NO ELASTIC"
            End If
            ''// Add the text as a symbol (For latter use by the StumpTool)
            'If (!Symbol("find", "TextAsSymbol")) Then
            '    Display("message", "error", "Elastic at distal end of support not added.\nCheck and add text manually")
            'End If
            'hEnt = AddEntity("symbol", "TextAsSymbol", xyLastTape.X + 0.25, xyOtemplate.Y + (xyLastTape.Y - xyOtemplate.Y) / 2, 1, 1, 90)
            'SetDBData(hEnt, "ID", sID + "PantyElasticNote")
            'SetDBData(hEnt, "Data", sFootLabel)

            ''// Panty Closing line	
            ''Execute("menu", "SetLayer", hTemplateLayer)
            ARMDIA1.PR_SetLayer(strLayer)
            ''hEnt = AddEntity("line", xyLastTape, xyOtemplate)
            ''SetDBData(hEnt, "ID", sID + "DistalClosingLine")
            PR_MakeXY(xyLastTape, xyToeSeam.X, xyOtemplate.Y)
            PR_DrawLine(xyLastTape, xyOtemplate)
            PR_MakeXY(xyText, xyLastTape.X, xyLastTape.Y + nSeam)
            PR_DrawLine(xyLastTape, xyText)
            ''AddEntity("line", xyO, xyOtemplate)
            PR_DrawLine(xyO, xyOtemplate)
        End If
        ''// Restore to layer 1
        ''Execute("menu", "SetLayer", Table("find", "layer", "1"))
    End Sub
    'To Draw Spline
    Private Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

End Class
Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
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


    'Structure xy
    '    Dim X As Double
    '    Dim Y As Double
    'End Structure

    'Public MainForm As torsodia

    ''PI as a Global Constant
    'Public Const PI As Double = 3.141592654

    ''Globals set by FN_Open
    'Public TORSODIA1.CC As Object 'Comma
    'Public TORSODIA1.QQ As Object 'Quote
    'Public TORSODIA1.NL As Object 'Newline
    'Public TORSODIA1.fNum As Object 'Macro file number
    'Public TORSODIA1.QCQ As Object 'Quote Comma Quote
    'Public TORSODIA1.QC As Object 'Quote Comma
    'Public TORSODIA1.CQ As Object 'Comma Quote
    ''Public g_nUnitsFac As Double
    'Public g_sChangeChecker As String
    'Public TORSODIA1.g_nCurrTextAngle As Object
    Dim xyInsertion As TORSODIA1.xy


    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        Dim Response As Short
        Dim sTask, sCurrentValues As String

        'Check if data has been modified
        sCurrentValues = FN_ValuesString()

        If sCurrentValues <> TORSODIA1.g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
            Select Case Response
                Case IDYES
                    ''PR_CreateMacro_Save("c:\jobst\draw.d")
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    'sTask = fnGetDrafixWindowTitleText()
                    'If sTask <> "" Then
                    '    AppActivate(sTask)
                    '    System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
                    '    Return
                    'Else
                    '    MsgBox("Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue")
                    'End If
                    Me.Close()
                Case IDNO
                    Me.Close()
                Case IDCANCEL
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
        VestMain.VestMainDlg.Close()
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
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Char_Renamed = TORSODIA1.QQ Then
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
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.fNum = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(TORSODIA1.fNum, sDrafixFile, VB.OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FN_OpenSave = TORSODIA1.fNum

        'Write header information etc. to the DRAFIX macro file
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "//Patient - " & sName & ", " & sFileNo & "")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "//by Visual Basic, TORSO Band")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "//type - " & sType & "")


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
        '            sError = sError & "Missing dimension for " & sCircum(ii) & "!" & TORSODIA1.NL
        '        End If
        '    Next ii

        For ii = 4 To 6
            If Val(txtCir(ii).Text) = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sError = sError & "Missing dimension for " & sCircum(ii) & "!" & TORSODIA1.NL
            End If
        Next ii

        'EOS Measurements (if one given both must be given)
        If Val(txtCir(7).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sError = sError & "Missing dimension for " & sCircum(7) & "!" & TORSODIA1.NL
        End If
        If Val(txtCir(8).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sError = sError & "Missing dimension for " & sCircum(8) & "!" & TORSODIA1.NL
        End If

        If cboClosure.Text = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sError = sError & "Closure not given! " & TORSODIA1.NL
        End If

        If cboFabric.Text = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sError = sError & "Fabric not given! " & TORSODIA1.NL
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

        '------------------------------------------------------------------
        'If txtUidMPD.Text = "" Then
        '    MsgBox("No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue")
        '    Return
        'End If

        If txtCombo(9).Text <> "" Then
            cboClosure.Text = txtCombo(9).Text
        Else
            cboClosure.SelectedIndex = 0
        End If
        cboFabric.Text = txtCombo(10).Text

        'Set up units
        If txtUnits.Text = "cm" Then
            TORSODIA1.g_nUnitsFac = 10 / 25.4
        Else
            TORSODIA1.g_nUnitsFac = 1
        End If

        'Display dimesions sizes in inches
        For ii = 4 To 8
            txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
        Next ii


        'Save the values in the text to a string
        'this can then be used to check if they have changed
        'on use of the close button
        TORSODIA1.g_sChangeChecker = FN_ValuesString()
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
        'If App.PrevInstance Then
        '	MsgBox("TORSO Band Dialogue is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting VEST Body - Dialogue")
        '          Return
        '      End If

        'Maintain while loading DDE data
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
        'Reset in Form_LinkClose

        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        TORSODIA1.MainForm = Me

        'Initialize globals
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.QQ = Chr(34) 'Double quotes (")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.NL = Chr(13) 'New Line
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.CC = Chr(44) 'The comma (,)
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.QCQ = TORSODIA1.QQ & TORSODIA1.CC & TORSODIA1.QQ
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.QC = TORSODIA1.QQ & TORSODIA1.CC
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.CQ = TORSODIA1.CC & TORSODIA1.QQ

        TORSODIA1.g_nUnitsFac = 1
        TORSODIA1.g_PathJOBST = fnPathJOBST()

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

        ''---------PR_GetComboListFromFile(cboFabric, TORSODIA1.g_PathJOBST & "\FABRIC.DAT")
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue")
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
        txtWorkOrder.Text = workOrder
        Form_LinkClose()
        readDWGInfo()

    End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
        Dim sTask As String
        'Don't allow multiple clicking
        '
        ''OK.Enabled = False
        If FN_ValidateData() Then
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Hide()
            VestMain.VestMainDlg.Hide()
            PR_GetInsertionPoint()
            ''PR_CreateMacro_Draw("c:\jobst\draw.d")
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            PR_CreateMacro_Draw(sDrawFile)
            saveInfoToDWG()

            'sTask = fnGetDrafixWindowTitleText()
            'If sTask <> "" Then
            '    AppActivate(sTask)
            '    System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
            '    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            '    Return
            'Else
            '    MsgBox("Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue")
            'End If
            OK.Enabled = True
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            VestMain.VestMainDlg.Close()
        End If
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
        Dim xyO As TORSODIA1.xy
        Dim xyEOS As TORSODIA1.xy
        Dim xyEOSatCL As TORSODIA1.xy
        Dim xyChest As TORSODIA1.xy
        Dim xyChestatCL As TORSODIA1.xy
        Dim xyWaist As TORSODIA1.xy
        Dim xyWaistatCL As TORSODIA1.xy

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

        Dim xyText As TORSODIA1.xy

        'Set figuring factors, flags, defaults etc...
        bEOSGiven = False
        nCirFactor = 0.9
        '    nShoulderCirFactor = 2.5
        nSeamAllowance = 0.1875
        PR_MakeXY(xyO, 0, 0)

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
            nEOSCir = fnDisplayToInches(CDbl(txtCir(8).Text)) * nCirFactor / 4
            nEOStoShoulder = fnDisplayToInches(CDbl(txtCir(7).Text))
        End If

        'Waist
        nWaistCir = fnDisplayToInches(CDbl(txtCir(6).Text)) * nCirFactor / 4
        nWaisttoShoulder = fnDisplayToInches(CDbl(txtCir(4).Text))

        'Chest
        nChestCir = fnDisplayToInches(CDbl(txtCir(5).Text)) * nCirFactor / 4

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
            nZiplength = FN_CalcLength(xyEOSatCL, xyChestatCL)
        Else
            nZiplength = FN_CalcLength(xyChestatCL, xyWaistatCL)
        End If
        nZiplength = (nZiplength - 0.125) / 0.95

        'DRAW DRAW DRAW DRAW DRAW DRAW DRAW DRAW
        '(Join the dots etc..)

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'TORSODIA1.fNum = ARMDIA1.FN_Open(sFile)
        TORSODIA1.fNum = FN_OpenSave(sFile, "Save", txtPatientName.Text, txtFileNo.Text)

        'Draw Text items
        PR_SetLayer("Notes")

        'Zippers
        ARMDIA1.PR_SetTextData(2, 32, -1, -1, -1) 'Horiz-Cen, Vertical-Bottom
        sText = ARMEDDIA1.fnInchesToText(nZiplength) & Chr(34) & cboClosure.Text
        PR_CalcMidPoint(xyChestatCL, xyWaistatCL, xyText)
        PR_MakeXY(xyText, xyText.X, xyText.Y + 0.25)
        PR_DrawText(sText, xyText, 0.1, 0)

        'Elastic Top
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.g_nCurrTextAngle = 270
        PR_CalcMidPoint(xyChestatCL, xyChest, xyText)
        PR_MakeXY(xyText, xyText.X - 0.25, xyText.Y)
        PR_DrawText("TOP ELASTIC", xyText, 0.1, (270 * (TORSODIA1.PI / 180)))

        'Elastic Bottom
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.g_nCurrTextAngle = 90
        If bEOSGiven Then
            PR_CalcMidPoint(xyEOSatCL, xyEOS, xyText)
        Else
            PR_CalcMidPoint(xyWaistatCL, xyWaist, xyText)
        End If
        PR_MakeXY(xyText, xyText.X + 0.25, xyText.Y)
        PR_DrawText("BOTTOM ELASTIC", xyText, 0.1, (90 * (TORSODIA1.PI / 180)))

        'Patient details
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.g_nCurrTextAngle = 0
        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
        PR_CalcMidPoint(xyChest, xyWaistatCL, xyText)
        If txtWorkOrder.Text = "" Then
            sWorkOrder = "-"
        Else
            sWorkOrder = txtWorkOrder.Text
        End If
        sText = txtPatientName.Text & Chr(10) & sWorkOrder & Chr(10) & Trim(Mid(cboFabric.Text, 4))
        'PR_DrawText(sText, xyText, 0.1)
        PR_DrawMText(sText, xyText)

        'Remaining patient details in black on layer construct
        ARMDIA1.PR_SetLayer("Construct")
        PR_MakeXY(xyText, xyText.X, xyText.Y - 0.8)
        sText = txtFileNo.Text & Chr(10) & txtDiagnosis.Text & Chr(10) & txtAge.Text & Chr(10) & txtSex.Text
        'PR_DrawText(sText, xyText, 0.1)
        PR_DrawMText(sText, xyText)

        'Draw construction lines
        If bEOSGiven Then
            PR_DrawLine(xyWaistatCL, xyWaist)
        End If

        'Draw profile
        ARMDIA1.PR_SetLayer("TemplateLeft")
        PR_StartPoly()
        Dim nCount As Integer = 5
        If bEOSGiven Then
            nCount = 6
        End If

        Dim PolyStart(nCount), PolyEnd(nCount) As Double
        If bEOSGiven Then
            PR_AddVertex(xyEOSatCL, 0)
            PR_AddVertex(xyEOS, 0)
            PR_AddVertex(xyWaist, 0)
            PR_AddVertex(xyChest, 0)
            PR_AddVertex(xyChestatCL, 0)
            PR_AddVertex(xyEOSatCL, 0)

            PolyStart(1) = xyEOSatCL.X
            PolyEnd(1) = xyEOSatCL.Y
            PolyStart(2) = xyEOS.X
            PolyEnd(2) = xyEOS.Y
            PolyStart(3) = xyWaist.X
            PolyEnd(3) = xyWaist.Y
            PolyStart(4) = xyChest.X
            PolyEnd(4) = xyChest.Y
            PolyStart(5) = xyChestatCL.X
            PolyEnd(5) = xyChestatCL.Y
            PolyStart(6) = xyEOSatCL.X
            PolyEnd(6) = xyEOSatCL.Y
        Else
            PR_AddVertex(xyWaistatCL, 0)
            PR_AddVertex(xyWaist, 0)
            PR_AddVertex(xyChest, 0)
            PR_AddVertex(xyChestatCL, 0)
            PR_AddVertex(xyWaistatCL, 0)

            PolyStart(1) = xyWaistatCL.X
            PolyEnd(1) = xyWaistatCL.Y
            PolyStart(2) = xyWaist.X
            PolyEnd(2) = xyWaist.Y
            PolyStart(3) = xyChest.X
            PolyEnd(3) = xyChest.Y
            PolyStart(4) = xyChestatCL.X
            PolyEnd(4) = xyChestatCL.Y
            PolyStart(5) = xyWaistatCL.X
            PolyEnd(5) = xyWaistatCL.Y
        End If
        PR_EndPoly()

        ''To Draw PolyLine
        PR_DrawPoly(PolyStart, PolyEnd, nCount)

        ARMDIA1.PR_SetLayer("Construct")
        PR_DrawXMarker()

        PR_UpdateDB()

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(TORSODIA1.fNum)

    End Sub

    Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
        'TORSODIA1.fNum is a global variable use in subsequent procedures
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TORSODIA1.fNum = FN_OpenSave(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

        'If this is a new drawing of a vest then Define the DATA Base
        'fields for the VEST Body and insert the BODYBOX symbol
        PR_PutLine("HANDLE hMPD, hBody;")

        PR_UpdateDB()

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(TORSODIA1.fNum)

    End Sub


    Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.ComboBox, ByRef sFileName As String)
        'General procedure to create the list section of
        'a combo box reading the data from a file

        Dim sLine As String
        Dim fFileNum As Short

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48, "CAD - Glove Dialogue")
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
        Do While Not EOF(fFileNum)
            sLine = LineInput(fFileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Combo_Name.AddItem(sLine)
            Combo_Name.Items.Add(sLine)
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
            PR_PutLine("@" & TORSODIA1.g_PathJOBST & "\VEST\VFIELDS.D;")

            'Find "mainpatientdetails" and get position
            PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
            PR_PutLine("STRING sMPD_Name;")
            PR_PutLine("ANGLE  aMPD_Angle;")

            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("hMPD = UID (" & TORSODIA1.QQ & "find" & TORSODIA1.QC & Val(txtUidMPD.Text) & ");")
            PR_PutLine("if (hMPD)")
            PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PR_PutLine("else")
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("  Exit(%cancel," & TORSODIA1.QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & TORSODIA1.QQ & ");")

            'Insert bodybox
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("if ( Symbol(" & TORSODIA1.QQ & "find" & TORSODIA1.QCQ & sSymbol & TORSODIA1.QQ & ")){")
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("  Execute (" & TORSODIA1.QQ & "menu" & TORSODIA1.QCQ & "SetLayer" & TORSODIA1.QC & "Table(" & TORSODIA1.QQ & "find" & TORSODIA1.QCQ & "layer" & TORSODIA1.QCQ & "Data" & TORSODIA1.QQ & "));")
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("  hBody = AddEntity(" & TORSODIA1.QQ & "symbol" & TORSODIA1.QCQ & sSymbol & TORSODIA1.QC & "xyMPD_Origin);")
            PR_PutLine("  }")
            PR_PutLine("else")
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("  Exit(%cancel, " & TORSODIA1.QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & TORSODIA1.QQ & ");")
        Else
            'Use existing symbol
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("hBody = UID (" & TORSODIA1.QQ & "find" & TORSODIA1.QC & Val(txtUidVB.Text) & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_PutLine("if (!hBody) Exit(%cancel," & TORSODIA1.QQ & "Can't find >" & sSymbol & "< symbol to update!" & TORSODIA1.QQ & ");")

        End If

        'Update the BODY Box symbol
        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "LtSCir" & TORSODIA1.QCQ & txtCir(0).Text & TORSODIA1.QQ & ");"
        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "RtSCir" & TORSODIA1.QCQ & txtCir(1).Text & TORSODIA1.QQ & ");"
        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "NeckCir" & TORSODIA1.QCQ & txtCir(2).Text & TORSODIA1.QQ & ");"
        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "SWidth" & TORSODIA1.QCQ & txtCir(3).Text & TORSODIA1.QQ & ");"
        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "S_Waist" & TORSODIA1.QCQ & txtCir(4).Text & TORSODIA1.QQ & ");"
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "SLgButt" & TORSODIA1.QCQ & txtCir(4).Text & TORSODIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "ChestCir" & TORSODIA1.QCQ & txtCir(5).Text & TORSODIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "WaistCir" & TORSODIA1.QCQ & txtCir(6).Text & TORSODIA1.QQ & ");")

        '    PR_PutLine "SetDBData( hBody" & TORSODIA1.CQ & "S_EOS" & TORSODIA1.QCQ & txtCir(7).Text & TORSODIA1.QQ & ");"
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "SFButt" & TORSODIA1.QCQ & txtCir(7).Text & TORSODIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "EOSCir" & TORSODIA1.QCQ & txtCir(8).Text & TORSODIA1.QQ & ");")

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "Closure" & TORSODIA1.QCQ & cboClosure.Text & TORSODIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "Fabric" & TORSODIA1.QCQ & cboFabric.Text & TORSODIA1.QQ & ");")


        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_PutLine("SetDBData( hBody" & TORSODIA1.CQ & "fileno" & TORSODIA1.QCQ & txtFileNo.Text & TORSODIA1.QQ & ");")

    End Sub

    Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
        'Allows the user to use enter as a tab

        System.Windows.Forms.SendKeys.Send("{TAB}")

    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'It is assumed that the link open from Drafix has failed
        'Therefor we "End" here
        Return
    End Sub

    Private Sub txtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Enter
        Dim Index As Short = txtCir.GetIndex(eventSender)
        TORSODIA1.PR_Select_Text(txtCir(Index))
    End Sub

    Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
        Dim Index As Short = txtCir.GetIndex(eventSender)
        Dim nLen As Double

        nLen = FN_InchesValue(txtCir(Index))
        If nLen > 0 Then
            lblCir(Index).Text = ARMEDDIA1.fnInchesToText(nLen)
        Else
            lblCir(Index).Text = ""
        End If

    End Sub
    Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    TORSODIA1.fNum is global variable
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, sLine)

    End Sub
    Sub PR_CalcMidPoint(ByRef xyStart As TORSODIA1.xy, ByRef xyEnd As TORSODIA1.xy, ByRef xyMid As TORSODIA1.xy)

        Static aAngle, nLength As Double

        aAngle = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)

        If nLength = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object xyMid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyMid = xyStart 'Avoid overflow
        Else
            PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
        End If

    End Sub
    Sub PR_MakeXY(ByRef xyReturn As TORSODIA1.xy, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.Y = y
    End Sub
    Sub PR_CalcPolar(ByRef xyStart As TORSODIA1.xy, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As TORSODIA1.xy)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double

        'Convert from degees to radians
        nAngle = nAngle * TORSODIA1.PI / 180

        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.X + A
        xyReturn.Y = xyStart.Y + B

    End Sub
    Sub PR_StartPoly()
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum)
        'write the syntax to start a POLYLINE the points will
        'be given by the PR_AddVertex and finished by PR_EndPoly
        'For this to work it assumes that the following DRAFIX variables
        'are defined and initialised
        '    XY      xyStart
        '    HANDLE  hEnt
        '    STRING  sID
        '
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL are globals initialised by FN_Open
        '
        '

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "StartPoly (""PolyLine"");")


    End Sub
    Sub PR_AddVertex(ByRef xyPoint As TORSODIA1.xy, ByRef nBulge As Double)
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum)
        'write the syntax to add a Vertex to a polyline opened
        'by PR_StartPoly
        'Allow the use of a bulge factor, but start and end widths
        'will be 0.
        'For this to work it assumes that the following DRAFIX variables
        'are defined and initialised
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL,  are globals initialised by FN_Open
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "  AddVertex(" & "xyStart.x+" & xyPoint.X & TORSODIA1.CC & "xyStart.y+" & xyPoint.Y & TORSODIA1.CC & "0,0," & nBulge & ");")

    End Sub
    Sub PR_EndPoly()
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum)
        'write the syntax to end a POLYLINE
        'For this to work it assumes that the following DRAFIX variables
        'are defined and initialised
        '    XY      xyStart
        '    HANDLE  hEnt
        '    STRING  sID
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL,  are globals initialised by FN_Open
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "EndPoly();")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "hEnt = UID (""find"", UID (""getmax"")) ;")
        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(TORSODIA1.fNum, "SetDBData(hEnt," & TORSODIA1.QQ & "ID" & TORSODIA1.QQ & ",sID);")

    End Sub
    Function FN_CalcAngle(ByRef xyStart As TORSODIA1.xy, ByRef xyEnd As TORSODIA1.xy) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, y As Object
        Dim rAngle As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = xyEnd.X - xyStart.X
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        y = xyEnd.Y - xyStart.Y

        'Horizomtal
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If y > 0 Then
                FN_CalcAngle = 90
            Else
                FN_CalcAngle = 270
            End If
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If y = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If X > 0 Then
                FN_CalcAngle = 0
            Else
                FN_CalcAngle = 180
            End If
            Exit Function
        End If

        'All other cases
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rAngle = System.Math.Atan(y / X) * (180 / TORSODIA1.PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If

    End Function
    Function FN_CalcLength(ByRef xyStart As TORSODIA1.xy, ByRef xyEnd As TORSODIA1.xy) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

    End Function
    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As TORSODIA1.xy, ByRef nHeight As Object, ByRef nAngle As Object)
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum).
        'Write the syntax to draw TEXT at the given height.
        '
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL, g_nCurrTextAspect are globals initialised by FN_Open
        '
        '
        Dim nWidth As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nWidth = nHeight * TORSODIA1.g_nCurrTextAspect
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

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

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsertion.X + xyInsert.X, xyInsertion.Y + xyInsert.Y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.HorizontalMode = TextHorizontalMode.TextCenter
                acText.AlignmentPoint = New Point3d(xyInsertion.X + xyInsert.X, xyInsertion.Y + xyInsert.Y, 0)

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using

        'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' PrintLine(TORSODIA1.fNum, "AddEntity(" & TORSODIA1.QQ & "text" & TORSODIA1.QCQ & sText & TORSODIA1.QC & "xyStart.x+" & Str(xyInsert.X) & TORSODIA1.CC & "xyStart.y+" & Str(xyInsert.y) & TORSODIA1.CC & nWidth & TORSODIA1.CC & nHeight & TORSODIA1.CC & TORSODIA1.g_nCurrTextAngle & ");")

    End Sub
    Sub PR_DrawLine(ByRef xyStart As TORSODIA1.xy, ByRef xyFinish As TORSODIA1.xy)
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL are globals initialised by FN_Open
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

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
            Dim acLine As Line = New Line(New Point3d(xyInsertion.X + xyStart.X, xyInsertion.Y + xyStart.Y, 0),
                                    New Point3d(xyInsertion.X + xyFinish.X, xyInsertion.Y + xyFinish.Y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
        '
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, "hEnt = AddEntity(")
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, TORSODIA1.QQ & "line" & TORSODIA1.QC)
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, "xyStart.x+" & Str(xyStart.X) & TORSODIA1.CC & "xyStart.y+" & Str(xyStart.y) & TORSODIA1.CC)
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, "xyStart.x+" & Str(xyFinish.X) & TORSODIA1.CC & "xyStart.y+" & Str(xyFinish.y))
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, ");")


    End Sub
    Sub PR_DrawPoly(ByRef ChinStrapX As Double(), ByRef ChinStrapY As Double(), ByRef nCount As Short)
        'To the DRAFIX macro file (given by the global TORSODIA1.fNum)
        'write the syntax to draw a POLYLINE through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    TORSODIA1.fNum, TORSODIA1.CC, TORSODIA1.QQ, TORSODIA1.NL are globals initialised by FN_Open
        '
        '
        Dim ii As Short

        'Exit if nothing to draw
        'If Profile.n <= 1 Then Exit Sub

        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, "hEnt = AddEntity(" & TORSODIA1.QQ & "poly" & TORSODIA1.QCQ & "polyline" & TORSODIA1.QQ)
        'For ii = 1 To Profile.n
        '    'UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    PrintLine(TORSODIA1.fNum, TORSODIA1.CC & "xyStart.x+" & Str(Profile.X(ii)) & TORSODIA1.CC & "xyStart.y+" & Str(Profile.y(ii)))
        'Next ii
        ''UPGRADE_WARNING: Couldn't resolve default property of object TORSODIA1.fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(TORSODIA1.fNum, ");")
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
                For ii = 1 To nCount
                    acPoly.AddVertexAt(ii - 1, New Point2d(xyInsertion.X + ChinStrapX(ii), xyInsertion.Y + ChinStrapY(ii)), 0, 0, 0)
                Next ii
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub
    'To Get Insertion point from user
    Private Sub PR_GetInsertionPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Indicate Start Point "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(xyInsertion, ptStart.X, ptStart.Y)
    End Sub

    'Private Sub txtwaist_leave(sender As Object, e As EventArgs) Handles _txtCir_4.Leave

    '    'Check for Numeric Value
    '    If Not IsNumeric(_txtCir_4.Text) And _txtCir_4.Text <> "" Then
    '        MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
    '        _txtCir_4.Focus()
    '    End If

    '    ' Show Values as Inches
    '    Dim inchbit As Double
    '    If Val(_txtCir_4.Text) > 0 Then
    '        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        inchbit = FN_CmToInches(txtUnits, _txtCir_4)
    '        _lblCir_4.Text = FN_InchesToText(inchbit)
    '    Else
    '        _lblCir_4.Text = ""
    '    End If
    'End Sub
    Private Function FN_InchesToText(ByRef nInches As Double) As String

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
                    sString = sString & " " & LTrim(Str(iEighths / 2)) & "/4"
                Case 4
                    sString = sString & " " & "1/2"
                Case Else
                    sString = sString & " " & LTrim(Str(iEighths)) & "/8"
            End Select
        Else
            sString = sString
        End If

        'Return formatted string
        FN_InchesToText = sString

    End Function
    Private Function FN_CmToInches(ByRef flag As System.Windows.Forms.Control, ByRef cm As System.Windows.Forms.Control) As Object

        ' This function converts cm to inches

        Dim temprem, remainder, inch, temp, rnder As Double
        Dim Value, realrem As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If flag.Text = "cm" Then ' value is in cm
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inch = Val(cm.Text) / 2.54
            temp = Int(inch)
            remainder = (inch - temp)
            remainder = remainder / 0.125
            realrem = Int(remainder)
            temprem = remainder - Int(remainder)
            If temprem >= 0.5 Then
                realrem = realrem + 1
            End If
            remainder = realrem * 0.125
            inch = temp + remainder
            FN_CmToInches = System.Math.Abs(inch)
            'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf flag.Text = "inches" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Value = cm.Text
            If Value >= 1 Then
                temp = Int(Value)
                remainder = Value - temp
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FN_CmToInches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    FN_CmToInches = System.Math.Abs(temp + remainder)
                End If
            Else
                temp = Int(Value)
                remainder = Value - temp
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FN_CmToInches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    FN_CmToInches = System.Math.Abs(remainder)
                End If
            End If
        End If

    End Function
    Function FN_InchesValue(ByRef TextBox As System.Windows.Forms.Control) As Double
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
                MsgBox("Invalid - Dimension has been entered", 48, "VEST Body - Dialogue")
                TextBox.Focus()
                FN_InchesValue = -1
                Exit For
            End If
        Next nn

        'Convert to inches
        nLen = fnDisplayToInches(Val(TextBox.Text))
        If nLen = -1 Then
            MsgBox("Invalid - Length has been entered", 48, "Glove Details")
            TextBox.Focus()
            FN_InchesValue = -1
        Else
            FN_InchesValue = nLen
        End If

    End Function

    Function fnDisplayToInches(ByVal nDisplay As Double) As Double
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
        If TORSODIA1.g_nUnitsFac <> 1 Then
            fnDisplayToInches = TORSODIA1.fnRoundInches(nDisplay * TORSODIA1.g_nUnitsFac) * iSign
            Exit Function
        End If

        'Imperial units
        iInt = Int(nDisplay)
        nDec = nDisplay - iInt
        'Check that conversion is possible (return -1 if not)
        If nDec > 0.8 Then
            fnDisplayToInches = -1
        Else
            fnDisplayToInches = TORSODIA1.fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
        End If

    End Function
    Private Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As TORSODIA1.xy)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

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
            mtx.Location = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = 0
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.BottomLeft
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("TorsoBand", "TORSODIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("TorsoBand", "TORSODIC")
            End If

            Dim resbuf As New ResultBuffer
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _txtCir_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _txtCir_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _txtCir_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _txtCir_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _txtCir_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _lblCir_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _lblCir_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _lblCir_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _lblCir_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _lblCir_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboClosure.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboFabric.Text))
            _sClass.SetXrecord(resbuf, "TorsoBand", "TORSODIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("TorsoBand", "TORSODIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                _txtCir_4.Text = arr(0).Value
                _txtCir_5.Text = arr(1).Value
                _txtCir_6.Text = arr(2).Value
                _txtCir_7.Text = arr(3).Value
                _txtCir_8.Text = arr(4).Value
                _lblCir_4.Text = arr(5).Value
                _lblCir_5.Text = arr(6).Value
                _lblCir_6.Text = arr(7).Value
                _lblCir_7.Text = arr(8).Value
                _lblCir_8.Text = arr(9).Value
                cboClosure.Text = arr(10).Value
                cboFabric.Text = arr(11).Value
            End If
            resbuf = _sClass.GetXrecord("NewPatient", "NEWPATIENTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim pattern As String = "dd-MM-yyyy"
                Dim parsedDate As DateTime
                DateTime.TryParseExact(arr(8).Value, pattern, System.Globalization.CultureInfo.CurrentCulture,
                                      System.Globalization.DateTimeStyles.None, parsedDate)
                Dim startTime As DateTime = Convert.ToDateTime(parsedDate)
                Dim endTime As DateTime = DateTime.Today
                Dim span As TimeSpan = endTime.Subtract(startTime)
                Dim totalDays As Double = span.TotalDays
                Dim totalYears As Double = Math.Truncate(totalDays / 365)
                txtAge.Text = totalYears.ToString()
            End If
        Catch ex As Exception
        End Try
    End Sub
    Sub PR_DrawXMarker()
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As TORSODIA1.xy
        PR_CalcPolar(xyBase, 135, 0.0625, xyStart)
        PR_CalcPolar(xyBase, -45, 0.0625, xyEnd)
        PR_CalcPolar(xyBase, 45, 0.0625, xySecSt)
        PR_CalcPolar(xyBase, -135, 0.0625, xySecEnd)

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
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
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

    Private Sub VestToSlv_Click(sender As Object, e As EventArgs) Handles VestToSlv.Click
        ''BODTOSLV.D file
        Me.Hide()
        VestMain.VestMainDlg.Hide()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select Vest Raglan Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.Keywords.Add("Line")
        ptEntOpts.Keywords.Add("Arc")
        ptEntOpts.Keywords.Add("Curve")
        ptEntOpts.Keywords.Add("Polyline")
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("Vest Raglan Profile not selected", 16, "VEST Body - Dialogue")
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If
        Dim sLayer As String = ""
        Try
            Dim idObject As ObjectId = ptEntRes.ObjectId
            Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim ent As Entity = tr.GetObject(idObject, OpenMode.ForRead)
                sLayer = ent.Layer()
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try

        Dim sSleeve As String = Mid(sLayer, 9, (sLayer.Length - 8))
        If (sSleeve.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase) = False And sSleeve.ToUpper().Equals("RIGHT", StringComparison.InvariantCultureIgnoreCase) = False) Then
            MsgBox("Select a Right or Left Raglan profile only", 16, "VEST Body - Dialogue")
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If

        Dim ptOpts As PromptPointOptions = New PromptPointOptions(vbCrLf + "Axilla Point")
        Dim ptRes As PromptPointResult = ed.GetPoint(ptOpts)
        If ptRes.Status <> PromptStatus.OK Then
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If
        Dim ptAxilla As Point3d = ptRes.Value
        Dim xyAxilla As VESTARM.XY
        xyAxilla.X = ptAxilla.X
        xyAxilla.y = ptAxilla.Y

        ptOpts = New PromptPointOptions(vbCrLf + "Front Neck and Raglan intesection")
        ptRes = ed.GetPoint(ptOpts)
        If ptRes.Status <> PromptStatus.OK Then
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If
        Dim ptFrontNeck As Point3d = ptRes.Value
        Dim xyFrontNeck As VESTARM.XY
        xyFrontNeck.X = ptFrontNeck.X
        xyFrontNeck.y = ptFrontNeck.Y

        ptOpts = New PromptPointOptions(vbCrLf + "Back Neck at end of Raglan")
        ptRes = ed.GetPoint(ptOpts)
        If ptRes.Status <> PromptStatus.OK Then
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If
        Dim ptBackNeck As Point3d = ptRes.Value
        Dim xyBackNeck As VESTARM.XY
        xyBackNeck.X = ptBackNeck.X
        xyBackNeck.y = ptBackNeck.Y

        ptOpts = New PromptPointOptions(vbCrLf + "Back Neck at Highest Shoulder line")
        ptRes = ed.GetPoint(ptOpts)
        If ptRes.Status <> PromptStatus.OK Then
            Me.Show()
            VestMain.VestMainDlg.Show()
            Exit Sub
        End If
        Dim ptBackNeckConstruct As Point3d = ptRes.Value
        Dim xyBackNeckConstruct As VESTARM.XY
        xyBackNeckConstruct.X = ptBackNeckConstruct.X
        xyBackNeckConstruct.y = ptBackNeckConstruct.Y

        Dim nAxillaFrontNeckRad, nAxillaBackNeckRad, nShoulderToBackRaglan As Double
        nAxillaFrontNeckRad = VESTARM.FN_CalcLength(xyAxilla, xyFrontNeck)
        nAxillaBackNeckRad = VESTARM.FN_CalcLength(xyAxilla, xyBackNeck)
        nShoulderToBackRaglan = VESTARM.FN_CalcLength(xyBackNeck, xyBackNeckConstruct)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If acBlkTbl.Has("VESTBODY") Then
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkRef As BlockReference = dbObj
                        If blkRef.Name = "VESTBODY" Then
                            For Each attributeID As ObjectId In blkRef.AttributeCollection
                                Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)

                                If TypeOf attRefObj Is AttributeReference Then
                                    Dim acAttRef As AttributeReference = attRefObj
                                    If (sSleeve.ToUpper().Equals("RIGHT", StringComparison.InvariantCultureIgnoreCase)) Then
                                        If acAttRef.Tag.ToUpper().Equals("AFNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nAxillaFrontNeckRad)
                                        ElseIf acAttRef.Tag.ToUpper().Equals("ABNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nAxillaBackNeckRad)
                                        ElseIf acAttRef.Tag.ToUpper().Equals("SBRaglanRight", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nShoulderToBackRaglan)
                                        End If

                                    ElseIf (sSleeve.ToUpper().Equals("LEFT", StringComparison.InvariantCultureIgnoreCase)) Then
                                        If acAttRef.Tag.ToUpper().Equals("AxillaFrontNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nAxillaFrontNeckRad)
                                        ElseIf acAttRef.Tag.ToUpper().Equals("AxillaBackNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nAxillaBackNeckRad)
                                        ElseIf acAttRef.Tag.ToUpper().Equals("ShoulderToBackRaglan", StringComparison.InvariantCultureIgnoreCase) Then
                                            acAttRef.TextString = Str(nShoulderToBackRaglan)
                                        End If
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
        Me.Close()
        VestMain.VestMainDlg.Close()
        MsgBox("Data Transfer Finished", 0, "VEST Body - Dialogue")
    End Sub
End Class
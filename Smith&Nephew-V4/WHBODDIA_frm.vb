Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
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


    'PI as a Global Constant
    Public Const PI As Double = 3.141592654



    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote

    Dim xyInsertion As WHBODDIA1.XY
    Dim xyWHCutInsert As WHBODDIA1.XY
    Dim g_sFoldHt As String
    Dim g_sLegStyle As String
    Dim sCrotchStyle As String
    Dim g_sBody As String
    Dim g_sAnkleTape As String
    Dim LeftLeg As Boolean
    Dim nLastTape As Double

    Dim nTOSCir As Double
    Dim nWaistCir As Double
    Dim nMidPointCir As Double
    Dim nLargestCir As Double
    Dim nWaistHt As Double
    Dim nLargestGivenRed As Double
    Dim nMidPointGivenRed As Double
    Dim nLeftThighCir As Double
    Dim nRightThighCir As Double
    Dim nThighGivenRed As Double
    Dim nWaistGivenRed As Double
    Dim nTOSGivenRed As Double
    Dim nTOSHt As Double
    Dim idLastMarker As ObjectId


    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        Dim Response As Short
        Dim sTask, sCurrentValues As String

        'Check if data has been modified
        sCurrentValues = FN_ValuesString()

        If sCurrentValues <> WHBODDIA1.g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "WAIST Height Body - Dialogue")
            Select Case Response
                Case IDYES
                    'PR_CreateMacro_Save("c:\jobst\draw.d")
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    'sTask = fnGetDrafixWindowTitleText()
                    'If sTask <> "" Then
                    '    AppActivate(sTask)
                    '    System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
                    '    Return
                    'Else
                    '    MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Dialogue")
                    'End If
                    saveInfoToDWG()
                    Me.Close()
                Case IDNO
                    Me.Close()
                Case IDCANCEL
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
        WaistMain.WaistMainDlg.Close()
    End Sub
    Public Function PR_CloseWaistBodyDialog() As Boolean
        Dim Response As Short
        Dim sCurrentValues As String

        'Check if data has been modified
        sCurrentValues = FN_ValuesString()
        If sCurrentValues <> WHBODDIA1.g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "Waist - Body")
            Select Case Response
                Case IDYES
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    saveInfoToDWG()
                Case IDCANCEL
                    Return False
            End Select
        End If
        Return True
    End Function

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
            sError = sError & "Male Crotch type specified for Female patient" + NL
        End If

        'Check for female crotch and male patient
        iIndex = 0
        iIndex = InStr(1, sString, "Female", 0)
        If iIndex <> 0 And txtSex.Text = "Male" Then
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

        If VB.Left(WHBODDIA1.g_sPreviousLegStyle, 4) = "Chap" Then prEnable_WRT_Chap()
        If VB.Left(WHBODDIA1.g_sPreviousLegStyle, 4) = "W4OC" Then prEnable_WRT_W4OC()
        If VB.Left(WHBODDIA1.g_sPreviousLegStyle, 5) = "Panty" Then prRestore_Reductions()
        If VB.Left(WHBODDIA1.g_sPreviousLegStyle, 5) = "Brief" Then prRestore_Reductions()

        'Get the currently selected leg stlye
        Select Case VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)

            Case "Full-Leg"
                prOptLeftLeg(1, 0, 0, 1, 1, 0)
                prOptRightLeg(1, 0, 0, 1, 1, 0)
                'Set dropdown combo boxes for reductions
                'Only set if Leg Style had been changed
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then
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
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then
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
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)

            Case "Chap-Right"
                prDisable_WRT_Chap()
                prOptLeftLeg(0, 0, 0, 0, 0, 0)
                prOptRightLeg(1, 0, 0, 1, 1, 0)
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)

            Case "Chap-Both Legs"
                prDisable_WRT_Chap()
                prOptLeftLeg(1, 0, 0, 1, 1, 0)
                prOptRightLeg(1, 0, 0, 1, 1, 0)
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then prSet_Red_Combo("0.87", cboWaistRed)

            Case "Brief"
                prOptLeftLeg(0, 0, 1, 0, 0, 1)
                prOptRightLeg(0, 0, 1, 0, 0, 1)
                'Set dropdown combo boxes for reductions
                'Only set if Leg Style had been changed
                If WHBODDIA1.g_sPreviousLegStyle <> "" Then
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
                If Val(txtLeftThighCir.Text) = 0 And WHBODDIA1.g_nLeftThighCir Then txtLeftThighCir.Text = CStr(WHBODDIA1.g_nLeftThighCir)
                txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())

                'Disable Right thigh data, save if required for a restore
                If Val(txtRightThighCir.Text) > 0 Then WHBODDIA1.g_nRightThighCir = Val(txtRightThighCir.Text)
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
                If Val(txtRightThighCir.Text) = 0 And WHBODDIA1.g_nRightThighCir Then txtRightThighCir.Text = CStr(WHBODDIA1.g_nRightThighCir)
                txtRightThighCir_Leave(txtRightThighCir, New System.EventArgs())

                'Disable Left thigh data, save if required for a restore
                If Val(txtLeftThighCir.Text) > 0 Then WHBODDIA1.g_nLeftThighCir = Val(txtLeftThighCir.Text)
                prOptLeftLeg(0, 0, 0, 0, 0, 0)
                '        txtLeftThighCir = ""
                '        txtLeftThighCir_LostFocus
                labLeftThigh.Enabled = False
                txtLeftThighCir.Enabled = False
                labLeftThighCir.Enabled = False
        End Select

        'Store leg style for use when it changes
        WHBODDIA1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
    End Sub

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
        fNum = FreeFile()
        FileOpen(fNum, sDrafixFile, VB.OpenMode.Output)
        FN_Open = fNum

        'Write header information etc. to the DRAFIX macro file
        PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(fNum, "//Patient - " & sName & ", " & sFileNo & "")
        PrintLine(fNum, "//by Visual Basic, Waist Height Body")
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
            WHBODDIA1.g_nUnitsFac = 10 / 25.4
        Else
            WHBODDIA1.g_nUnitsFac = 1
        End If

        'Waist Height details
        'Note
        ' The procedure Validate_Text_In_Box also displays the size in
        ' inches.
        '
        WHBODDIA1.Validate_Text_In_Box(txtTOSCir, labTOSCir)
        WHBODDIA1.Validate_Text_In_Box(txtTOSHt, labTOSHt)
        WHBODDIA1.Validate_Text_In_Box(txtWaistCir, labWaistCir)
        WHBODDIA1.Validate_Text_In_Box(txtWaistHt, labWaistHt)
        WHBODDIA1.Validate_Text_In_Box(txtMidPointCir, labMidPointCir)
        WHBODDIA1.Validate_Text_In_Box(txtMidPointHt, labMidPointHt)
        WHBODDIA1.Validate_Text_In_Box(txtLargestCir, labLargestCir)
        WHBODDIA1.Validate_Text_In_Box(txtLargestHt, labLargestHt)
        WHBODDIA1.Validate_Text_In_Box(txtLeftThighCir, labLeftThighCir)
        WHBODDIA1.Validate_Text_In_Box(txtRightThighCir, labRightThighCir)
        WHBODDIA1.Validate_Text_In_Box(txtFoldHt, labFoldHt)

        'Set dropdown combo boxes
        'Crotch Style
        cboCrotchStyle.Items.Add("Open Crotch") 'Both
        cboCrotchStyle.Items.Add("Horizontal Fly") 'Male only
        cboCrotchStyle.Items.Add("Diagonal Fly") 'Male only
        cboCrotchStyle.Items.Add("Gusset") 'Both
        cboCrotchStyle.Items.Add("Mesh Gusset") 'Both
        cboCrotchStyle.Items.Add("Female Mesh Gusset") 'Female only
        cboCrotchStyle.Items.Add("Male Mesh Gusset") 'Male only
        cboCrotchStyle.SelectedIndex = 3
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
        WHBODDIA1.g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
        cboLegStyle.SelectedIndex = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 1)

        'Set this now to avoid problems with cboLegStyle_Click
        If cboLegStyle.SelectedIndex = -1 Then
            WHBODDIA1.g_sPreviousLegStyle = "None"
        Else
            WHBODDIA1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
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

        ii = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 2)
        If ii >= 0 Then optLeftLeg(ii).Checked = True
        ii = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 3)
        If ii >= 0 Then optRightLeg(ii).Checked = True

        'Revised Heights
        'From Body DB field
        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 5)
        If nValue > 0 Then
            txtTOSHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtTOSHtRevised, labTOSHtRevised)
        End If

        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 6)
        If nValue > 0 Then
            txtWaistHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtWaistHtRevised, labWaistHtRevised)
        End If

        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 7)
        If nValue > 0 Then
            txtFoldHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtFoldHtRevised, labFoldHtRevised)
        End If

        If WHBODDIA1.fnGetNumber(txtBody.Text, 4) > 0 Then
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
        End If

        'Store loded values w.r.t cancel button
        '
        WHBODDIA1.g_sChangeChecker = FN_ValuesString()

        Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to default.
        Me.Activate()
    End Sub

    'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
        If CmdStr = "Cancel" Then
            Cancel = 0
            ''---------------------------------------End
        End If
    End Sub

    Private Sub whboddia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            Hide()
            'Check if this is already running
            'If it is then warn the user and exit
            ''==================================================
            'If App.PrevInstance Then
            '	End
            'End If
            '==========================================

            'Maintain while loading DDE data
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to Hourglass.
            'Reset in Form_LinkClose

            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            WHBODDIA1.MainForm = Me

            'Initialize globals
            QQ = Chr(34) 'Double quotes (")
            NL = Chr(13) 'New Line
            CC = Chr(44) 'The comma (,)
            QCQ = QQ & CC & QQ
            QC = QQ & CC
            CQ = CC & QQ
            idLastMarker = New ObjectId

            WHBODDIA1.g_sPathJOBST = fnPathJOBST()

            'Set Units global to default to inches
            WHBODDIA1.g_nUnitsFac = 1

            ''Setup combo box fields
            cboTOSRed.Items.Add("0.78")
            cboTOSRed.Items.Add("0.79")
            cboTOSRed.Items.Add("0.80")
            cboTOSRed.Items.Add("0.81")
            cboTOSRed.Items.Add("0.82")
            cboTOSRed.Items.Add("0.83")
            cboTOSRed.Items.Add("0.84")
            cboTOSRed.Items.Add("0.85")
            cboTOSRed.Items.Add("0.86")
            cboTOSRed.Items.Add("0.87")
            cboTOSRed.Items.Add("0.88")
            cboTOSRed.Items.Add("0.89")
            cboTOSRed.Items.Add("0.90")
            cboTOSRed.Items.Add("0.91")
            cboTOSRed.Items.Add("0.92")
            cboTOSRed.Items.Add("0.93")

            cboWaistRed.Items.Add("0.78")
            cboWaistRed.Items.Add("0.79")
            cboWaistRed.Items.Add("0.80")
            cboWaistRed.Items.Add("0.81")
            cboWaistRed.Items.Add("0.82")
            cboWaistRed.Items.Add("0.83")
            cboWaistRed.Items.Add("0.84")
            cboWaistRed.Items.Add("0.85")
            cboWaistRed.Items.Add("0.86")
            cboWaistRed.Items.Add("0.87")
            cboWaistRed.Items.Add("0.88")
            cboWaistRed.Items.Add("0.89")
            cboWaistRed.Items.Add("0.90")
            cboWaistRed.Items.Add("0.91")
            cboWaistRed.Items.Add("0.92")
            cboWaistRed.Items.Add("0.93")

            cboMidPointRed.Items.Add("0.78")
            cboMidPointRed.Items.Add("0.79")
            cboMidPointRed.Items.Add("0.80")
            cboMidPointRed.Items.Add("0.81")
            cboMidPointRed.Items.Add("0.82")
            cboMidPointRed.Items.Add("0.83")
            cboMidPointRed.Items.Add("0.84")
            cboMidPointRed.Items.Add("0.85")
            cboMidPointRed.Items.Add("0.86")
            cboMidPointRed.Items.Add("0.87")
            cboMidPointRed.Items.Add("0.88")
            cboMidPointRed.Items.Add("0.89")
            cboMidPointRed.Items.Add("0.90")
            cboMidPointRed.Items.Add("0.91")
            cboMidPointRed.Items.Add("0.92")
            cboMidPointRed.Items.Add("0.93")

            cboLargestRed.Items.Add("0.78")
            cboLargestRed.Items.Add("0.79")
            cboLargestRed.Items.Add("0.80")
            cboLargestRed.Items.Add("0.81")
            cboLargestRed.Items.Add("0.82")
            cboLargestRed.Items.Add("0.83")
            cboLargestRed.Items.Add("0.84")
            cboLargestRed.Items.Add("0.85")
            cboLargestRed.Items.Add("0.86")
            cboLargestRed.Items.Add("0.87")
            cboLargestRed.Items.Add("0.88")
            cboLargestRed.Items.Add("0.89")
            cboLargestRed.Items.Add("0.90")
            cboLargestRed.Items.Add("0.91")
            cboLargestRed.Items.Add("0.92")
            cboLargestRed.Items.Add("0.93")

            cboThighRed.Items.Add("0.78")
            cboThighRed.Items.Add("0.79")
            cboThighRed.Items.Add("0.80")
            cboThighRed.Items.Add("0.81")
            cboThighRed.Items.Add("0.82")
            cboThighRed.Items.Add("0.83")
            cboThighRed.Items.Add("0.84")
            cboThighRed.Items.Add("0.85")
            cboThighRed.Items.Add("0.86")
            cboThighRed.Items.Add("0.87")
            cboThighRed.Items.Add("0.88")
            cboThighRed.Items.Add("0.89")
            cboThighRed.Items.Add("0.90")
            cboThighRed.Items.Add("0.91")
            cboThighRed.Items.Add("0.92")
            cboThighRed.Items.Add("0.93")

            'Patient details
            txtFileNo.Text = ""
            txtUnits.Text = ""
            txtPatientName.Text = ""
            txtDiagnosis.Text = ""
            txtAge.Text = ""
            txtSex.Text = ""
            'txtWorkOrder.Text = ""

            ''---------PR_GetComboListFromFile(cboFabric, TORSODIA1.g_PathJOBST & "\FABRIC.DAT")
            Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
            ''PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim blkId As ObjectId = New ObjectId()
            Dim obj As New BlockCreation.BlockCreation
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("No Patient Details have been found in drawing!", 16, "Waist Body - Dialogue")
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
            'txtWorkOrder.Text = workOrder

            'Disable revised Hts (default), set from Link Close event
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked

            'Start a timer that will ensure that the
            'the Vb programme ends if the DRAFIX DDE
            'link fails.

            PR_UpdateAge()
            Form_LinkClose()
            readDWGInfo()
            If cboTOSRed.Text <> "" Then
                cboTOSRed.Enabled = True
            End If
            If cboLegStyle.SelectedIndex = -1 Then
                cboLegStyle.SelectedIndex = 0
            End If
            If cboCrotchStyle.SelectedIndex = -1 Then
                cboCrotchStyle.SelectedIndex = 0
            End If
            'Store loded values w.r.t cancel button
            WHBODDIA1.g_sChangeChecker = FN_ValuesString()
            Timer1.Interval = 6000
            Timer1.Enabled = True
        Catch ex As Exception
            Me.Close()
            WaistMain.WaistMainDlg.Close()
        End Try
    End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
        'Check that data is all present and insert into drafix

        ''Dim sTask As String
        OK.Enabled = False
        If Validate_Data() Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Hide()
            WaistMain.WaistMainDlg.Hide()
            ''PR_CreateMacro_Save("c:\jobst\draw.d")
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            PR_CreateMacro_Save(sDrawFile)
            saveInfoToDWG()
            PR_DrawWaistBodyBlock()
            WHBODDIA1.g_sChangeChecker = FN_ValuesString()
            '         sTask = fnGetDrafixWindowTitleText()
            'If sTask <> "" Then
            '	AppActivate(sTask)
            '	System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
            '	'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            '	System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            '             ''-----------------------------------------End
            '         Else
            '	MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Data")
            'End If
        End If
        OK.Enabled = True
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        WaistMain.WaistMainDlg.Show()
        Me.Show()
    End Sub
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

            If sLegStyle = "W4OC-Left" And WHBODDIA1.g_sPreviousLegStyle <> "" Then
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

            If sLegStyle = "W4OC-Right" And WHBODDIA1.g_sPreviousLegStyle <> "" Then
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
        fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

        'If this is a new drawing of a vest then Define the DATA Base
        'fields for the VEST Body and insert the BODYBOX symbol
        PR_PutLine("HANDLE hMPD, hBody;")

        Update_DDE_Text_Boxes()

        Dim sSymbol As String

        sSymbol = "waistbody"

        If txtUidWHBody.Text = "" Then
            'Define DB Fields
            PR_PutLine("@" & WHBODDIA1.g_sPathJOBST & "\WAIST\WHFIELDS.D;")

            'Find "mainpatientdetails" and get position
            PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
            PR_PutLine("STRING sMPD_Name;")
            PR_PutLine("ANGLE  aMPD_Angle;")

            PR_PutLine("hMPD = UID (" & QQ & "find" & QC & Val(txtUidTitle.Text) & ");")
            PR_PutLine("if (hMPD)")
            PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")

            'Insert body symbol
            PR_PutLine("if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
            PR_PutLine("  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
            PR_PutLine("  hBody = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);")
            PR_PutLine("  }")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
        Else
            'Use existing symbol
            PR_PutLine("hBody = UID (" & QQ & "find" & QC & Val(txtUidWHBody.Text) & ");")
            PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
        End If

        'Update the BODY Box symbol
        PR_PutLine("SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "TOSCir" & QCQ & txtTOSCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "TOSGivenRed" & QCQ & txtTOSRed.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "TOSHt" & QCQ & txtTOSHt.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtWaistCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "WaistGivenRed" & QCQ & txtWaistRed.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "WaistHt" & QCQ & txtWaistHt.Text & QQ & ");")

        PR_PutLine("SetDBData( hBody" & CQ & "MidPointCir" & QCQ & txtMidPointCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "MidPointGivenRed" & QCQ & txtMidPointRed.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "MidPointHt" & QCQ & txtMidPointHt.Text & QQ & ");")

        PR_PutLine("SetDBData( hBody" & CQ & "LargestCir" & QCQ & txtLargestCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "LargestGivenRed" & QCQ & txtLargestRed.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "LargestHt" & QCQ & txtLargestHt.Text & QQ & ");")

        PR_PutLine("SetDBData( hBody" & CQ & "LeftThighCir" & QCQ & txtLeftThighCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "RightThighCir" & QCQ & txtRightThighCir.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "ThighGivenRed" & QCQ & txtThighRed.Text & QQ & ");")

        PR_PutLine("SetDBData( hBody" & CQ & "FoldHt" & QCQ & txtFoldHt.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "CrotchStyle" & QCQ & txtCrotchStyle.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "ID" & QCQ & txtLegStyle.Text & QQ & ");")
        PR_PutLine("SetDBData( hBody" & CQ & "Body" & QCQ & txtBody.Text & QQ & ");")

        FileClose(fNum)
    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable

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

        If WHBODDIA1.g_nLeftThighCir > 0 Then txtLeftThighCir.Text = CStr(WHBODDIA1.g_nLeftThighCir)
        txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())
        txtLeftThighCir.Enabled = True
        labLeftThighCir.Enabled = True
        labLeftThigh.Enabled = True

        If WHBODDIA1.g_nRightThighCir > 0 Then txtRightThighCir.Text = CStr(WHBODDIA1.g_nRightThighCir)
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
        If Val(txtLeftThighCir.Text) = 0 And WHBODDIA1.g_nLeftThighCir Then txtLeftThighCir.Text = CStr(WHBODDIA1.g_nLeftThighCir)
        txtLeftThighCir_Leave(txtLeftThighCir, New System.EventArgs())

        'Enable Right thigh, restore if value available
        prOptRightLeg(0, 1, 0, 1, 1, 0)
        txtRightThighCir.Enabled = True
        labRightThigh.Enabled = True
        If Val(txtRightThighCir.Text) = 0 And WHBODDIA1.g_nRightThighCir Then txtRightThighCir.Text = CStr(WHBODDIA1.g_nRightThighCir)
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
        '--------------------------Combo_Box_Name.AddItem("0.78") '0
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '---------------------------Combo_Box_Name.AddItem("0.79") '1
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-----------------------------------------Combo_Box_Name.AddItem("0.80") '2
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '------------------------------Combo_Box_Name.AddItem("0.81") '3
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-----------------------------------Combo_Box_Name.AddItem("0.82") '4
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '--------------------------------Combo_Box_Name.AddItem("0.83") '5
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-------------------------------Combo_Box_Name.AddItem("0.84") '6
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '----------------------------------Combo_Box_Name.AddItem("0.85") '7
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-----------------------------------Combo_Box_Name.AddItem("0.86") '8
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '------------------------------------Combo_Box_Name.AddItem("0.87") '9
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '------------------------------------Combo_Box_Name.AddItem("0.88") '10
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '------------------------------------Combo_Box_Name.AddItem("0.89") '11
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '------------------------------------Combo_Box_Name.AddItem("0.90") '12
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-------------------------------------Combo_Box_Name.AddItem("0.91") '13
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '-----------------------------------Combo_Box_Name.AddItem("0.92") '15
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''-------------------------Combo_Box_Name.AddItem("0.93") '15

        'If no values given exit
        If Val(sRed) = 0 And Val(Text_Box_Name.Text) = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ''---------------------Combo_Box_Name.ListIndex = -1
            Exit Sub
        End If

        'Set the value of the combo box as given
        If sRed = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ''---------------------------Combo_Box_Name.ListIndex = (Val(Text_Box_Name.Text) * 100) - 78
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ''---------------------------Combo_Box_Name.ListIndex = (Val(sRed) * 100) - 78
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

    Private Sub prSet_Red_Combo(ByRef sRed As String, ByRef Combo_Box_Name As System.Windows.Forms.ComboBox)
        'Set the value of the combo box as given
        'Find index into combo box list

        If Val(sRed) = 0 Then Exit Sub
        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Box_Name.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Combo_Box_Name.SelectedIndex = (Val(sRed) * 100) - 78

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

        If Val(txtWaistCir.Text) > 0 Then WHBODDIA1.g_nWaistCir = Val(txtWaistCir.Text)
        If Val(txtWaistHt.Text) > 0 Then g_nWaistHt = Val(txtWaistHt.Text)
        If Val(txtWaistHtRevised.Text) > 0 Then g_nWaistHtRevised = Val(txtWaistHtRevised.Text)

        If Val(txtMidPointCir.Text) > 0 Then g_nMidPointCir = Val(txtMidPointCir.Text)
        If Val(txtMidPointHt.Text) > 0 Then g_nMidPointHt = Val(txtMidPointHt.Text)

        If Val(txtLargestCir.Text) > 0 Then g_nLargestCir = Val(txtLargestCir.Text)
        If Val(txtLargestHt.Text) > 0 Then g_nLargestHt = Val(txtLargestHt.Text)

        If Val(txtLeftThighCir.Text) > 0 Then WHBODDIA1.g_nLeftThighCir = Val(txtLeftThighCir.Text)
        If Val(txtRightThighCir.Text) > 0 Then WHBODDIA1.g_nRightThighCir = Val(txtRightThighCir.Text)

        If Val(txtFoldHt.Text) > 0 Then g_nFoldHt = Val(txtFoldHt.Text)
        If Val(txtFoldHtRevised.Text) > 0 Then g_nFoldHtRevised = Val(txtFoldHtRevised.Text)

        prStore_Reductions()

    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'If this event is called then the
        'DDE link with DRAFIX has not closed
        'Therefor we time out after 6 seconds
        '-----------End
    End Sub

    Private Sub txtFoldHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHt.Enter
        WHBODDIA1.Select_Text_In_Box(txtFoldHt)
    End Sub

    Private Sub txtFoldHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHt.Leave
        WHBODDIA1.Validate_Text_In_Box(txtFoldHt, labFoldHt)
    End Sub

    Private Sub txtFoldHtRevised_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHtRevised.Enter
        WHBODDIA1.Select_Text_In_Box(txtFoldHtRevised)
    End Sub

    Private Sub txtFoldHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFoldHtRevised.Leave
        WHBODDIA1.Validate_Text_In_Box(txtFoldHtRevised, labFoldHtRevised)
        If Val(txtFoldHtRevised.Text) > 0 Then
            txtFoldHt.Enabled = False
            labFoldHt.Enabled = False
        Else
            txtFoldHt.Enabled = True
            labFoldHt.Enabled = True
        End If
    End Sub

    Private Sub txtLargestCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtLargestCir)
    End Sub

    Private Sub txtLargestCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtLargestCir, labLargestCir)
    End Sub

    Private Sub txtLargestHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestHt.Enter
        WHBODDIA1.Select_Text_In_Box(txtLargestHt)
    End Sub

    Private Sub txtLargestHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLargestHt.Leave
        WHBODDIA1.Validate_Text_In_Box(txtLargestHt, labLargestHt)
    End Sub

    Private Sub txtLeftThighCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftThighCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtLeftThighCir)
    End Sub

    Private Sub txtLeftThighCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftThighCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtLeftThighCir, labLeftThighCir)
    End Sub

    Private Sub txtMidPointCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtMidPointCir)
    End Sub

    Private Sub txtMidPointCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtMidPointCir, labMidPointCir)
    End Sub

    Private Sub txtMidPointHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointHt.Enter
        WHBODDIA1.Select_Text_In_Box(txtMidPointHt)
    End Sub

    Private Sub txtMidPointHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMidPointHt.Leave
        WHBODDIA1.Validate_Text_In_Box(txtMidPointHt, labMidPointHt)
    End Sub

    Private Sub txtRightThighCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightThighCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtRightThighCir)
    End Sub

    Private Sub txtRightThighCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightThighCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtRightThighCir, labRightThighCir)
    End Sub

    Private Sub txtTOSCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtTOSCir)
    End Sub

    Private Sub txtTOSCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtTOSCir, labTOSCir)

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
        WHBODDIA1.Select_Text_In_Box(txtTOSHt)
    End Sub

    Private Sub txtTOSHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHt.Leave
        WHBODDIA1.Validate_Text_In_Box(txtTOSHt, labTOSHt)

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
        WHBODDIA1.Select_Text_In_Box(txtTOSHtRevised)
    End Sub

    Private Sub txtTOSHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTOSHtRevised.Leave
        WHBODDIA1.Validate_Text_In_Box(txtTOSHtRevised, labTOSHtRevised)
        If Val(txtTOSHtRevised.Text) > 0 Then
            txtTOSHt.Enabled = False
            labTOSHt.Enabled = False
        Else
            txtTOSHt.Enabled = True
            labTOSHt.Enabled = True
        End If
    End Sub

    Private Sub txtWaistCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Enter
        WHBODDIA1.Select_Text_In_Box(txtWaistCir)
    End Sub

    Private Sub txtWaistCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Leave
        WHBODDIA1.Validate_Text_In_Box(txtWaistCir, labWaistCir)
    End Sub

    Private Sub txtWaistHt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHt.Enter
        WHBODDIA1.Select_Text_In_Box(txtWaistHt)
    End Sub

    Private Sub txtWaistHt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHt.Leave
        WHBODDIA1.Validate_Text_In_Box(txtWaistHt, labWaistHt)
    End Sub

    Private Sub txtWaistHtRevised_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHtRevised.Enter
        WHBODDIA1.Select_Text_In_Box(txtWaistHtRevised)
    End Sub

    Private Sub txtWaistHtRevised_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistHtRevised.Leave
        WHBODDIA1.Validate_Text_In_Box(txtWaistHtRevised, labWaistHtRevised)
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
        nLargestCir = WHBODDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
        iLegStyle = cboLegStyle.SelectedIndex

        'nOpenBack, nOpenFront
        If txtCrotchStyle.Text = "Open Crotch" Then
            If iLegStyle = 3 Or iLegStyle = 4 Then
                'W4OC-Left and W4OC-Right
                If CDbl(txtAge.Text) <= 10 Then
                    nOpenBack = 2.0#
                    nOpenFront = 2.0#
                Else
                    If nLargestCir >= 45 Then
                        nOpenBack = 4.0#
                        nOpenFront = 4.0#
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
                            nOpenBack = 3.0#
                            nOpenFront = 4.0#
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
                            nOpenBack = 4.0#
                            nOpenFront = 3.0#
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

        Dim xyFold, xyButtockArcCen As WHBODDIA1.XY

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
            Validate_Data = False
            Exit Function
        Else
            Validate_Data = True
        End If

        'NON-FATAL Errors (warnings)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Establish Largest circumferance
        nMaxCir = WHBODDIA1.fnDisplayToInches(Val(txtLargestCir.Text))

        'Establish hts (original or revised)
        'If no revised ht given then use the original
        If chkDrawRevisedHts.CheckState = 1 Then
            nFoldHt = WHBODDIA1.fnDisplayToInches(Val(txtFoldHtRevised.Text))
            nWaistHt = WHBODDIA1.fnDisplayToInches(Val(txtWaistHtRevised.Text))
            nTOSHt = WHBODDIA1.fnDisplayToInches(Val(txtTOSHtRevised.Text))
            'If no revised ht given then use the original
            If nFoldHt = 0 Then nFoldHt = WHBODDIA1.fnDisplayToInches(Val(txtFoldHt.Text))
            If nWaistHt = 0 Then nWaistHt = WHBODDIA1.fnDisplayToInches(Val(txtWaistHt.Text))
            If nTOSHt = 0 Then nTOSHt = WHBODDIA1.fnDisplayToInches(Val(txtTOSHt.Text))
        Else
            nFoldHt = WHBODDIA1.fnDisplayToInches(Val(txtFoldHt.Text))
            nWaistHt = WHBODDIA1.fnDisplayToInches(Val(txtWaistHt.Text))
            nTOSHt = WHBODDIA1.fnDisplayToInches(Val(txtTOSHt.Text))
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
                nRightThighCir = WHBODDIA1.fnDisplayToInches(Val(txtRightThighCir.Text))
                nLeftThighCir = WHBODDIA1.fnDisplayToInches(Val(txtLeftThighCir.Text))

            Case 3, 5 'W4OC-Left and CHAP-Left
                nLeftThighCir = WHBODDIA1.fnDisplayToInches(Val(txtLeftThighCir.Text))
                nRightThighCir = nLeftThighCir

            Case 4, 6 'W4OC-Right and CHAP-Right
                nRightThighCir = WHBODDIA1.fnDisplayToInches(Val(txtRightThighCir.Text))
                nLeftThighCir = nRightThighCir

        End Select

        nThighRed = Val(VB6.GetItemString(cboThighRed, cboThighRed.SelectedIndex))
        nThighCirOriginal = WHBODDIA1.min(nRightThighCir, nLeftThighCir)
        nThighCir = WHBODDIA1.fnRoundInches(nThighCirOriginal * nThighRed)

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
                Validate_Data = True
            Else
                Validate_Data = False
                Exit Function
            End If
        Else
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

            nLargestHt = WHBODDIA1.fnRoundInches((nLargestHt - nDatum) * nXscale)

            nFoldHt = WHBODDIA1.fnRoundInches((nFoldHt - nDatum) * nXscale)

            xyFold.X = nFoldHt
            xyFold.Y = nThighCir * nYscale + nSeam

            xyButtockArcCen.X = nLargestHt
            xyButtockArcCen.Y = ((nThighCir / 2) * nYscale) + nSeam

            nButtockArcRad = WHBODDIA1.FN_CalcLength(xyFold, xyButtockArcCen)

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
                nLargestCir = WHBODDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
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
                nLargestCir = WHBODDIA1.fnDisplayToInches(Val(txtLargestCir.Text))
                nLargestCir = WHBODDIA1.fnRoundInches(nLargestCir * nLargestRed)

                nCO_LargestTopOff = nLargestOff - (((nLargestCir * nYscale) - nLength) * (1 - nCrotchFrontFactor))
                nDistance = WHBODDIA1.fnRoundInches(nLargestOff - nCO_LargestTopOff)

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
                    Validate_Data = True
                Case 1 To 3
                    MsgBox(sError, 64, "Warning - Problems with data")
                    Validate_Data = False
                Case 4 To 100
                    sError = sError & NL & "The above problems have been found in the data do you" & NL
                    sError = sError & "wish to continue ?"
                    iError = MsgBox(sError, 52, "Severe Problems with data")
                    If iError = IDYES Then
                        Validate_Data = True
                    Else
                        Validate_Data = False
                    End If
                Case Else
                    Validate_Data = True
            End Select
        End If 'End of If to exclude CHAP leg styles
    End Function
    Sub PR_MakeXY(ByRef xyReturn As WHBODDIA1.XY, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.Y = y
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("WaistBody", "WHBODDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("WaistBody", "WHBODDIC")
            End If

            Dim resbuf As New ResultBuffer
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTOSCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWaistCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtMidPointCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLargestCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLeftThighCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtRightThighCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTOSHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWaistHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtMidPointHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLargestHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFoldHt.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labTOSCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labWaistCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labMidPointCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labLargestCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labLeftThighCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labRightThighCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labTOSHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labWaistHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labMidPointHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labLargestHt.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), labFoldHt.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboTOSRed.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboWaistRed.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboMidPointRed.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboLargestRed.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboThighRed.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboLegStyle.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboCrotchStyle.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLabel.Checked))

            ''Added for #216 in the issue list
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTOSHtRevised.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWaistHtRevised.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFoldHtRevised.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkDrawRevisedHts.Checked))
            _sClass.SetXrecord(resbuf, "WaistBody", "WHBODDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("WaistBody", "WHBODDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                txtTOSCir.Text = arr(0).Value
                txtWaistCir.Text = arr(1).Value
                txtMidPointCir.Text = arr(2).Value
                txtLargestCir.Text = arr(3).Value
                txtLeftThighCir.Text = arr(4).Value
                txtRightThighCir.Text = arr(5).Value
                txtTOSHt.Text = arr(6).Value
                txtWaistHt.Text = arr(7).Value
                txtMidPointHt.Text = arr(8).Value
                txtLargestHt.Text = arr(9).Value
                txtFoldHt.Text = arr(10).Value

                labTOSCir.Text = arr(11).Value
                labWaistCir.Text = arr(12).Value
                labMidPointCir.Text = arr(13).Value
                labLargestCir.Text = arr(14).Value
                labLeftThighCir.Text = arr(15).Value
                labRightThighCir.Text = arr(16).Value
                labTOSHt.Text = arr(17).Value
                labWaistHt.Text = arr(18).Value
                labMidPointHt.Text = arr(19).Value
                labLargestHt.Text = arr(20).Value
                labFoldHt.Text = arr(21).Value

                cboTOSRed.Text = arr(22).Value
                cboWaistRed.Text = arr(23).Value
                cboMidPointRed.Text = arr(24).Value
                cboLargestRed.Text = arr(25).Value
                cboThighRed.Text = arr(26).Value
                cboLegStyle.Text = arr(27).Value
                cboCrotchStyle.Text = arr(28).Value
                chkLabel.Checked = arr(29).Value
                ''Added for #216 in the issue list
                If arr.Length > 30 Then
                    txtTOSHtRevised.Text = arr(30).Value
                    txtWaistHtRevised.Text = arr(31).Value
                    txtFoldHtRevised.Text = arr(32).Value
                    chkDrawRevisedHts.Checked = arr(33).Value
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
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
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Waist Details")
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
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Waist Details")
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
        nLen = WHBODDIA1.fnDisplayToInches(Val(TextBox.Text))
        If nLen = -1 Then
            MsgBox("Invalid - Length has been entered", 48, "Glove Details")
            TextBox.Focus()
            FN_InchesValue = -1
        Else
            FN_InchesValue = nLen
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
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawWaistBodyBlock()
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
        Dim xyText As WHBODDIA1.XY
        PR_MakeXY(xyText, 0.71875, 1.25)
        Dim strTag(21), strTextString(21) As String
        strTag(1) = "fileno"
        strTag(2) = "TOSCir"
        strTag(3) = "TOSGivenRed"
        strTag(4) = "TOSHt"
        strTag(5) = "WaistCir"
        strTag(6) = "WaistGivenRed"
        strTag(7) = "WaistHt"
        strTag(8) = "MidPointCir"
        strTag(9) = "MidPointGivenRed"
        strTag(10) = "MidPointHt"
        strTag(11) = "LargestCir"
        strTag(12) = "LargestGivenRed"
        strTag(13) = "LargestHt"
        strTag(14) = "LeftThighCir"
        strTag(15) = "RightThighCir"
        strTag(16) = "ThighGivenRed"
        strTag(17) = "FoldHt"
        strTag(18) = "CrotchStyle"
        strTag(19) = "ID"
        strTag(20) = "Body"
        strTag(21) = "Fabric"

        strTextString(1) = txtFileNo.Text
        strTextString(2) = txtTOSCir.Text
        strTextString(3) = cboTOSRed.Text
        strTextString(4) = txtTOSHt.Text
        strTextString(5) = txtWaistCir.Text
        strTextString(6) = cboWaistRed.Text
        strTextString(7) = txtWaistHt.Text
        strTextString(8) = txtMidPointCir.Text
        strTextString(9) = cboMidPointRed.Text
        strTextString(10) = txtMidPointHt.Text
        strTextString(11) = txtLargestCir.Text
        strTextString(12) = cboLargestRed.Text
        strTextString(13) = txtLargestHt.Text
        strTextString(14) = txtLeftThighCir.Text
        strTextString(15) = txtRightThighCir.Text
        strTextString(16) = cboThighRed.Text
        strTextString(17) = txtFoldHt.Text
        strTextString(18) = cboCrotchStyle.Text
        strTextString(19) = txtLegStyle.Text
        strTextString(20) = txtBody.Text
        strTextString(21) = ""

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRefPatient As BlockReference = acTrans.GetObject(blkId, OpenMode.ForRead)
            Dim ptPosition As Point3d = blkRefPatient.Position
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("WAISTBODY") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "WAISTBODY"
                Dim acPoly As Polyline = New Polyline()
                Dim ii As Double
                For ii = 1 To 5
                    acPoly.AddVertexAt(ii - 1, New Point2d(xyStart(ii), xyEnd(ii)), 0, 0, 0)
                Next ii
                blkTblRecCross.AppendEntity(acPoly)

                Dim acText As DBText = New DBText()
                acText.Position = New Point3d(xyText.X, xyText.Y, 0)
                acText.Height = 0.1
                acText.TextString = "WH-BODY"
                acText.Rotation = 0
                acText.Justify = AttachmentPoint.BottomCenter
                acText.AlignmentPoint = New Point3d(xyText.X, xyText.Y, 0)
                blkTblRecCross.AppendEntity(acText)

                Dim acAttDef As New AttributeDefinition
                For ii = 1 To 21
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
                'Dim acLyrTbl As LayerTable
                'acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                'If acLyrTbl.Has("DATA") = True Then
                '    acCurDb.Clayer = acLyrTbl("DATA")
                'End If
                ''Changed for SN-BLOCK issue on 31 July 2019
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForRead)
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkTitleBox As BlockReference = dbObj
                        If blkTitleBox.Name = "MAINPATIENTDETAILS" Then
                            Dim strTitleBoxLayer As String = blkTitleBox.Layer
                            Dim acLyrTbl As LayerTable
                            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(strTitleBoxLayer) = True Then
                                acCurDb.Clayer = acLyrTbl(strTitleBoxLayer)
                            End If
                            Exit For
                        End If
                    End If
                Next
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl("WAISTBODY")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(ptPosition, blkRecId)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    blkRef.TransformBy(Matrix3d.Scaling(2.54, ptPosition))
                'End If
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

                ''--------------Added to set TITLEBOX Layer to all block references on 30-5-2019
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkTitleBox As BlockReference = dbObj
                        If blkTitleBox.Name = "MAINPATIENTDETAILS" Then
                            Dim strTitleBoxLayer As String = blkTitleBox.Layer
                            Dim acLyrTbl As LayerTable
                            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(strTitleBoxLayer) = True Then
                                acCurDb.Clayer = acLyrTbl(strTitleBoxLayer)
                            End If
                            Exit For
                        End If
                    End If
                Next
                ''-----------------------

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
                            ElseIf acAttRef.Tag.ToUpper().Equals("TOSCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtTOSCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("TOSGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboTOSRed.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("TOSHt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtTOSHt.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("WaistCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtWaistCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("WaistGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboWaistRed.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("WaistHt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtWaistHt.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("MidPointCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtMidPointCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("MidPointGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboMidPointRed.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("MidPointHt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtMidPointHt.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LargestCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLargestCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LargestGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboLargestRed.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LargestHt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLargestHt.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LeftThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLeftThighCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("RightThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtRightThighCir.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ThighGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboThighRed.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("FoldHt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFoldHt.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("CrotchStyle", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboCrotchStyle.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLegStyle.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("Body", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtBody.Text
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
    Private Sub PR_UpdateAge()
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
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
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForRead)
            For Each objID As ObjectId In acBlkTblRec
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                If TypeOf dbObj Is BlockReference Then
                    Dim blkRef As BlockReference = dbObj
                    If blkRef.Name = "WAISTBODY" Then
                        For Each attributeID As ObjectId In blkRef.AttributeCollection
                            Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                            If TypeOf attRefObj IsNot AttributeReference Then
                                Continue For
                            End If
                            Dim acAttDef As AttributeReference = attRefObj
                            If acAttDef.Tag.ToUpper().Equals("FoldHt", StringComparison.InvariantCultureIgnoreCase) Then
                                g_sFoldHt = acAttDef.TextString
                            ElseIf acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                                g_sLegStyle = acAttDef.TextString
                            ElseIf acAttDef.Tag.ToUpper().Equals("TOSCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nTOSCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("WaistCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nWaistCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("MidPointCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nMidPointCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("LargestCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nLargestCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("WaistHt", StringComparison.InvariantCultureIgnoreCase) Then
                                nWaistHt = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("LargestGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                nLargestGivenRed = Val(acAttDef.TextString)
                            ElseIf acAttDef.Tag.ToUpper().Equals("MidPointGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                nMidPointGivenRed = Val(acAttDef.TextString)
                            ElseIf acAttDef.Tag.ToUpper().Equals("LeftThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nLeftThighCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("RightThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                                nRightThighCir = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("ThighGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                nThighGivenRed = Val(acAttDef.TextString)
                            ElseIf acAttDef.Tag.ToUpper().Equals("WaistGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                nWaistGivenRed = Val(acAttDef.TextString)
                            ElseIf acAttDef.Tag.ToUpper().Equals("TOSGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                nTOSGivenRed = Val(acAttDef.TextString)
                            ElseIf acAttDef.Tag.ToUpper().Equals("CrotchStyle", StringComparison.InvariantCultureIgnoreCase) Then
                                sCrotchStyle = acAttDef.TextString
                            ElseIf acAttDef.Tag.ToUpper().Equals("TOSHt", StringComparison.InvariantCultureIgnoreCase) Then
                                nTOSHt = WHBODDIA1.fnDisplayToInches(Val(acAttDef.TextString))
                            ElseIf acAttDef.Tag.ToUpper().Equals("Body", StringComparison.InvariantCultureIgnoreCase) Then
                                g_sBody = acAttDef.TextString
                            ElseIf acAttDef.Tag.ToUpper().Equals("AnkleTape", StringComparison.InvariantCultureIgnoreCase) Then
                                g_sAnkleTape = acAttDef.TextString
                            End If
                        Next
                    End If
                End If
            Next

            '    Dim strLegBlkName As String = ""
            '    If acBlkTbl.Has("WAISTLEFTLEG") Then
            '        LeftLeg = True
            '        strLegBlkName = "WAISTLEFTLEG"
            '    ElseIf acBlkTbl.Has("WAISTRIGHTLEG") Then
            '        LeftLeg = False
            '        strLegBlkName = "WAISTRIGHTLEG"
            '    Else
            '        MsgBox("Can't find WAIST LEG symbol to update!", 16, "Waist - Dialogue")
            '        Return False
            '    End If
            '    Dim blkRecIdLeft As ObjectId = acBlkTbl(strLegBlkName)
            '    If blkRecIdLeft <> ObjectId.Null Then
            '        Dim blkTblRec As BlockTableRecord
            '        blkTblRec = acTrans.GetObject(blkRecIdLeft, OpenMode.ForRead)
            '        Dim blkRef As BlockReference = New BlockReference(New Point3d(0, 0, 0), blkRecIdLeft)
            '        For Each objID As ObjectId In blkTblRec
            '            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
            '            If TypeOf dbObj Is AttributeDefinition Then
            '                Dim acAttDef As AttributeDefinition = dbObj
            '                If acAttDef.Tag.ToUpper().Equals("LastTape", StringComparison.InvariantCultureIgnoreCase) Then
            '                    nLastTape = Val(acAttDef.TextString)
            '                End If
            '            End If
            '        Next
            '    End If
        End Using
        Return True
    End Function
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
        PR_MakeXY(xyWHCutInsert, ptStart.X, ptStart.Y)
    End Sub
    Function FNRound(ByVal nNumber As Single) As Double
        'E.G.
        '    round(1.35) = 1
        '    round(1.55) = 2
        '    round(2.50) = 3
        '    round(-2.50) = -3
        '
        Dim nInt, nSign As Short
        nSign = System.Math.Sign(nNumber)
        nNumber = System.Math.Abs(nNumber)
        Dim nPrecision As Double = 0.125
        nInt = Int(nNumber)
        Dim nDec As Double = nNumber - nInt

        'Get decimal part in precision units
        If nDec <> 0 Then
            nDec = nDec / nPrecision 'Avoid overflow
        End If
        nDec = round(nDec)

        'If (nNumber - nInt) >= 0.5 Then
        '    FNRound = (nInt + 1) * nSign
        'Else
        '    FNRound = nInt * nSign
        'End If
        FNRound = (nInt + (nDec * nPrecision)) * nSign
    End Function
    Private Function round(ByVal nValue As Double) As Short
        Dim iInt, iSign As Short

        'Avoid extra work. Return 0 if input is 0
        If nValue = 0 Then
            round = 0
            Exit Function
        End If

        'Split input
        iSign = System.Math.Sign(nValue)
        nValue = System.Math.Abs(nValue)
        iInt = Int(nValue)

        'Effect rounding
        If (nValue - iInt) >= 0.5 Then
            round = (iInt + 1) * iSign
        Else
            round = iInt * iSign
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PR_DrawCutout()
    End Sub
    Public Sub PR_DrawWaistCutout()
        'Initialize globals
        QQ = Chr(34) 'Double quotes (")
        NL = Chr(13) 'New Line
        CC = Chr(44) 'The comma (,)
        QCQ = QQ & CC & QQ
        QC = QQ & CC
        CQ = CC & QQ
        idLastMarker = New ObjectId

        WHBODDIA1.g_sPathJOBST = fnPathJOBST()
        'Set Units global to default to inches
        WHBODDIA1.g_nUnitsFac = 1

        ''Setup combo box fields
        cboTOSRed.Items.Add("0.78")
        cboTOSRed.Items.Add("0.79")
        cboTOSRed.Items.Add("0.80")
        cboTOSRed.Items.Add("0.81")
        cboTOSRed.Items.Add("0.82")
        cboTOSRed.Items.Add("0.83")
        cboTOSRed.Items.Add("0.84")
        cboTOSRed.Items.Add("0.85")
        cboTOSRed.Items.Add("0.86")
        cboTOSRed.Items.Add("0.87")
        cboTOSRed.Items.Add("0.88")
        cboTOSRed.Items.Add("0.89")
        cboTOSRed.Items.Add("0.90")
        cboTOSRed.Items.Add("0.91")
        cboTOSRed.Items.Add("0.92")
        cboTOSRed.Items.Add("0.93")

        cboWaistRed.Items.Add("0.78")
        cboWaistRed.Items.Add("0.79")
        cboWaistRed.Items.Add("0.80")
        cboWaistRed.Items.Add("0.81")
        cboWaistRed.Items.Add("0.82")
        cboWaistRed.Items.Add("0.83")
        cboWaistRed.Items.Add("0.84")
        cboWaistRed.Items.Add("0.85")
        cboWaistRed.Items.Add("0.86")
        cboWaistRed.Items.Add("0.87")
        cboWaistRed.Items.Add("0.88")
        cboWaistRed.Items.Add("0.89")
        cboWaistRed.Items.Add("0.90")
        cboWaistRed.Items.Add("0.91")
        cboWaistRed.Items.Add("0.92")
        cboWaistRed.Items.Add("0.93")

        cboMidPointRed.Items.Add("0.78")
        cboMidPointRed.Items.Add("0.79")
        cboMidPointRed.Items.Add("0.80")
        cboMidPointRed.Items.Add("0.81")
        cboMidPointRed.Items.Add("0.82")
        cboMidPointRed.Items.Add("0.83")
        cboMidPointRed.Items.Add("0.84")
        cboMidPointRed.Items.Add("0.85")
        cboMidPointRed.Items.Add("0.86")
        cboMidPointRed.Items.Add("0.87")
        cboMidPointRed.Items.Add("0.88")
        cboMidPointRed.Items.Add("0.89")
        cboMidPointRed.Items.Add("0.90")
        cboMidPointRed.Items.Add("0.91")
        cboMidPointRed.Items.Add("0.92")
        cboMidPointRed.Items.Add("0.93")

        cboLargestRed.Items.Add("0.78")
        cboLargestRed.Items.Add("0.79")
        cboLargestRed.Items.Add("0.80")
        cboLargestRed.Items.Add("0.81")
        cboLargestRed.Items.Add("0.82")
        cboLargestRed.Items.Add("0.83")
        cboLargestRed.Items.Add("0.84")
        cboLargestRed.Items.Add("0.85")
        cboLargestRed.Items.Add("0.86")
        cboLargestRed.Items.Add("0.87")
        cboLargestRed.Items.Add("0.88")
        cboLargestRed.Items.Add("0.89")
        cboLargestRed.Items.Add("0.90")
        cboLargestRed.Items.Add("0.91")
        cboLargestRed.Items.Add("0.92")
        cboLargestRed.Items.Add("0.93")

        cboThighRed.Items.Add("0.78")
        cboThighRed.Items.Add("0.79")
        cboThighRed.Items.Add("0.80")
        cboThighRed.Items.Add("0.81")
        cboThighRed.Items.Add("0.82")
        cboThighRed.Items.Add("0.83")
        cboThighRed.Items.Add("0.84")
        cboThighRed.Items.Add("0.85")
        cboThighRed.Items.Add("0.86")
        cboThighRed.Items.Add("0.87")
        cboThighRed.Items.Add("0.88")
        cboThighRed.Items.Add("0.89")
        cboThighRed.Items.Add("0.90")
        cboThighRed.Items.Add("0.91")
        cboThighRed.Items.Add("0.92")
        cboThighRed.Items.Add("0.93")

        'Patient details
        txtFileNo.Text = ""
        txtUnits.Text = ""
        txtPatientName.Text = ""
        txtDiagnosis.Text = ""
        txtAge.Text = ""
        txtSex.Text = ""
        'txtWorkOrder.Text = ""

        ''---------PR_GetComboListFromFile(cboFabric, TORSODIA1.g_PathJOBST & "\FABRIC.DAT")
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        ''PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("No Patient Details have been found in drawing!", 16, "Waist Body - Dialogue")
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
        'txtWorkOrder.Text = workOrder

        'Disable revised Hts (default), set from Link Close event
        chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
        chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked

        PR_UpdateAge()
        ''---------------------Form_LinkClose()--------------------
        'On link close.
        Dim sReductionFactor As String
        Dim nAge, ii As Short
        Dim nValue As Double
        'Set global Units Factor
        If txtUnits.Text = "cm" Then
            WHBODDIA1.g_nUnitsFac = 10 / 25.4
        Else
            WHBODDIA1.g_nUnitsFac = 1
        End If

        'Waist Height details
        'Note
        ' The procedure Validate_Text_In_Box also displays the size in
        ' inches.
        '
        WHBODDIA1.Validate_Text_In_Box(txtTOSCir, labTOSCir)
        WHBODDIA1.Validate_Text_In_Box(txtTOSHt, labTOSHt)
        WHBODDIA1.Validate_Text_In_Box(txtWaistCir, labWaistCir)
        WHBODDIA1.Validate_Text_In_Box(txtWaistHt, labWaistHt)
        WHBODDIA1.Validate_Text_In_Box(txtMidPointCir, labMidPointCir)
        WHBODDIA1.Validate_Text_In_Box(txtMidPointHt, labMidPointHt)
        WHBODDIA1.Validate_Text_In_Box(txtLargestCir, labLargestCir)
        WHBODDIA1.Validate_Text_In_Box(txtLargestHt, labLargestHt)
        WHBODDIA1.Validate_Text_In_Box(txtLeftThighCir, labLeftThighCir)
        WHBODDIA1.Validate_Text_In_Box(txtRightThighCir, labRightThighCir)
        WHBODDIA1.Validate_Text_In_Box(txtFoldHt, labFoldHt)

        'Set dropdown combo boxes
        'Crotch Style
        cboCrotchStyle.Items.Add("Open Crotch") 'Both
        cboCrotchStyle.Items.Add("Horizontal Fly") 'Male only
        cboCrotchStyle.Items.Add("Diagonal Fly") 'Male only
        cboCrotchStyle.Items.Add("Gusset") 'Both
        cboCrotchStyle.Items.Add("Mesh Gusset") 'Both
        cboCrotchStyle.Items.Add("Female Mesh Gusset") 'Female only
        cboCrotchStyle.Items.Add("Male Mesh Gusset") 'Male only
        cboCrotchStyle.SelectedIndex = 3
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
        WHBODDIA1.g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
        cboLegStyle.SelectedIndex = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 1)

        'Set this now to avoid problems with cboLegStyle_Click
        If cboLegStyle.SelectedIndex = -1 Then
            WHBODDIA1.g_sPreviousLegStyle = "None"
        Else
            WHBODDIA1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
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

        ii = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 2)
        If ii >= 0 Then optLeftLeg(ii).Checked = True
        ii = WHBODDIA1.fnGetNumber(txtLegStyle.Text, 3)
        If ii >= 0 Then optRightLeg(ii).Checked = True

        'Revised Heights
        'From Body DB field
        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 5)
        If nValue > 0 Then
            txtTOSHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtTOSHtRevised, labTOSHtRevised)
        End If

        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 6)
        If nValue > 0 Then
            txtWaistHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtWaistHtRevised, labWaistHtRevised)
        End If

        nValue = WHBODDIA1.fnGetNumber(txtBody.Text, 7)
        If nValue > 0 Then
            txtFoldHtRevised.Text = CStr(nValue)
            WHBODDIA1.Validate_Text_In_Box(txtFoldHtRevised, labFoldHtRevised)
        End If

        If WHBODDIA1.fnGetNumber(txtBody.Text, 4) > 0 Then
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDrawRevisedHts.CheckState = System.Windows.Forms.CheckState.Checked
        End If

        'Store loded values w.r.t cancel button
        WHBODDIA1.g_sChangeChecker = FN_ValuesString()
        ''-----------------------------------------
        readDWGInfo()
        If cboTOSRed.Text <> "" Then
            cboTOSRed.Enabled = True
        End If
        If cboLegStyle.SelectedIndex = -1 Then
            cboLegStyle.SelectedIndex = 0
        End If
        If cboCrotchStyle.SelectedIndex = -1 Then
            cboCrotchStyle.SelectedIndex = 0
        End If
        PR_DrawCutout()
    End Sub
    Private Sub PR_DrawCutout()
        ''// File Name:	WHCUTDBD.D
        If PR_GetBlkAttributeValues() = False Then
            Exit Sub
        End If
        If WHBODDIA1.fnGetNumber(g_sLegStyle, 1) >= 5 Then
            MsgBox("Can't Draw a CUT-OUT for Chap Styles", 16, "Waist - Dialog")
            Exit Sub
        End If
        If chkDrawRevisedHts.CheckState = 1 Then
            If Val(txtFoldHtRevised.Text) > 0 Then
                g_sFoldHt = txtFoldHtRevised.Text
            End If
            If Val(txtWaistHtRevised.Text) > 0 Then
                nWaistHt = WHBODDIA1.fnDisplayToInches(Val(txtWaistHtRevised.Text))
            End If
            If Val(txtTOSHtRevised.Text) > 0 Then
                nTOSHt = WHBODDIA1.fnDisplayToInches(Val(txtTOSHtRevised.Text))
            End If
        End If

        Dim bIsBrief As Boolean = False
        If WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 2 Then
            bIsBrief = True
        End If
        Dim nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta, nBottArcToLargestOffset As Double
        ''// Get Origin
        Dim xyO, xyStart, xyEnd As WHBODDIA1.XY
        PR_GetInsertionPoint()
        xyO = xyWHCutInsert
        ''xyO.X = xyWHCutInsert.X - 0.625
        Dim nBodyLegTapePos As Integer
        Dim nFoldHt As Double = WHBODDIA1.fnDisplayToInches(Val(g_sFoldHt))
        Dim nDatum As Double = Int(nFoldHt / 1.5) * 1.5
        nBodyLegTapePos = Int(nFoldHt / 1.5) + 6

        ''// NB Explict leg style given
        ''//
        Dim PantyLeg As Boolean = False
        Dim nLegStyle As Double = WHBODDIA1.fnGetNumber(g_sLegStyle, 1)
        Dim sLeg As String = "Left"
        LeftLeg = True
        If (nLegStyle = 1) Then
            PantyLeg = True
        End If

        ''// Calculate Cut Out And TOS
        ''//

        ''// For all subsequent calculations the measurements are worked
        ''// from the right edge of the body template #8654
        Dim nXscale, nYscale, nGivenTOSCir, nGivenWaistCir, nGivenMidPointCir, nGivenLargestCir As Double
        nXscale = 1.25 / 1.5
        nYscale = 0.5

        ''// Scale given values to Decimal inches
        ''// Retain the Given value for later use as the other value will be FIGURED
        ''//
        ''// Circumferences
        nGivenTOSCir = nTOSCir
        nGivenWaistCir = nWaistCir
        nGivenMidPointCir = nMidPointCir
        nGivenLargestCir = nLargestCir

        '' // Start Calculating XY control points
        ''// First Figure Heights And Circumferences
        ''// Largest part of buttocks
        Dim nLength, nLargestHt, nMidPointHt, nThighCirOriginal, nThighCir, nPantyThighRed2, nSeam As Double
        nLength = (nFoldHt + (nWaistHt - nFoldHt) / 3)
        nLargestHt = FNRound((nLength - nDatum) * nXscale)
        ''nLargestHt = ((nLength - nDatum) * nXscale)
        nLargestCir = FNRound(nLargestCir * nLargestGivenRed)

        ''// Mid Point (Cleft)
        nMidPointHt = (nLength + (nWaistHt - nFoldHt) / 4)   ''// N.B. Carry through Of nLength from above
        nMidPointHt = FNRound((nMidPointHt - nDatum) * nXscale)
        Dim xyMidPoint, xyFoldPanty, xyButtockArcCen, xyLargest, xyFold As WHBODDIA1.XY
        xyMidPoint.X = xyO.X + nMidPointHt
        nMidPointCir = FNRound(nMidPointCir * nMidPointGivenRed)

        ''// Fold of Buttocks (11.1)
        nPantyThighRed2 = 0.92
        nSeam = 0.1875
        nFoldHt = FNRound((nFoldHt - nDatum) * nXscale)
        ''nFoldHt = round((nFoldHt - nDatum) * nXscale)
        nThighCirOriginal = WHBODDIA1.min(nLeftThighCir, nRightThighCir)
        If (PantyLeg Or bIsBrief) Then
            '' // 92% mark at thigh
            xyFoldPanty.X = xyO.X + nFoldHt
            xyFoldPanty.Y = xyO.Y + FNRound(nThighCirOriginal * nPantyThighRed2) * nYscale + nSeam
        End If
        nThighCir = FNRound(nThighCirOriginal * nThighGivenRed)

        xyFold.X = xyO.X + nFoldHt
        xyFold.Y = xyO.Y + nThighCir * nYscale + nSeam

        ''//  Largest part of Buttocks (12.1)
        xyButtockArcCen.Y = xyO.Y + (nThighCir / 2) * nYscale + nSeam
        xyButtockArcCen.X = xyO.X + nLargestHt
        nLength = WHBODDIA1.FN_CalcLength(xyButtockArcCen, xyFold)

        xyLargest.Y = xyButtockArcCen.Y + nLength
        xyLargest.X = xyButtockArcCen.X

        ''// WaistCir
        nWaistCir = FNRound(nWaistCir * nWaistGivenRed)

        ''// TOSCir
        nTOSCir = FNRound(nTOSCir * nTOSGivenRed)
        ''// CutOut  (13.1)
        Dim OpenCrotch, ClosedCrotch As Boolean
        OpenCrotch = False
        ClosedCrotch = False
        If (sCrotchStyle.Equals("Open Crotch")) Then
            OpenCrotch = True
        Else
            ClosedCrotch = True
        End If

        ''// Get distance to largest part of buttocks from seam line (13)
        ''// nCrotchFrontFactor from WHBODDIA - VB programme
        Dim nCrotchFrontFactor, nOpenFront, nOpenBack, nOpenOff As Double
        nCrotchFrontFactor = WHBODDIA1.fnGetNumber(g_sBody, 1)
        nOpenFront = WHBODDIA1.fnGetNumber(g_sBody, 2)
        nOpenBack = WHBODDIA1.fnGetNumber(g_sBody, 3)

        Dim nAge As Integer = Val(txtAge.Text)
        'Dim nHeelLength As Double = WHBODDIA1.fnGetNumber(g_sAnkleTape, 1)
        'Dim Footless As Boolean = False
        'If WHBODDIA1.fnGetNumber(g_sAnkleTape, 2) = 1 Then
        '    Footless = True
        'End If
        If sCrotchStyle = "Open Crotch" Then
            If txtSex.Text = "Male" Then
                nOpenOff = 0.75
                'Check for footless style or a footless w40c
                'If nLegStyle = 1 Or nLegStyle = 2 Or (nLegStyle = 3 And Footless = True) Or (nLegStyle = 4 And Footless = True) Then
                '    If nAge <= 10 Then
                '        nOpenOff = 0.375
                '    End If
                'Else
                '    If nHeelLength > 0 And nHeelLength < 9 Then
                '        nOpenOff = 0.375
                '    End If
                'End If
            Else
                nOpenOff = 0.375
            End If
        End If

        'Brief offset
        Dim nBriefOff As Double = 0
        If nLegStyle = 2 Then    'Brief
            nBriefOff = 1.25
            If sCrotchStyle = "Open Crotch" And nOpenOff = 0.375 Then
                nBriefOff = 1.625
                If nOpenOff = 0.75 Then
                    nBriefOff = 2.0#
                End If
            ElseIf sCrotchStyle = "Horizontal Fly" Then
                nBriefOff = 1.625
            End If
        End If

        nLength = xyLargest.Y - xyO.Y - nSeam
        Dim xyCO_LargestBott, xyCO_LargestTop As WHBODDIA1.XY
        xyCO_LargestBott.X = xyLargest.X
        xyCO_LargestBott.Y = xyO.Y + ((nLargestCir * nYscale - nLength) * nCrotchFrontFactor) + nSeam

        xyCO_LargestTop.X = xyLargest.X
        xyCO_LargestTop.Y = xyLargest.Y - ((nLargestCir * nYscale - nLength) * (1 - nCrotchFrontFactor))

        Dim nCO_HalfWidth As Double = WHBODDIA1.FN_CalcLength(xyCO_LargestBott, xyCO_LargestTop) / 2
        ''// Check on Cut Out Diameter
        Dim nCutOutDiaMaxTol As Double = 6.0
        If (FNRound(nCO_HalfWidth * 2) > nCutOutDiaMaxTol) Then
            MsgBox("Warning, Cut Out Diameter is larger than 6 inches" +
                   "\nCalculated diameter is " + Format("length", nCO_HalfWidth * 2), 16, "Waist - Dialog")
        End If

        ''// Cut Out 
        ''// Top Of Support (if given) (Cut Out Front)
        Dim nDiff As Integer
        Dim xyCO_TOSBott, xyCO_WaistBott, xyCO_MidPointBott As WHBODDIA1.XY
        Dim nBodyFrontReduceOff As Double = 0.125
        Dim nBodyFrontIncreaseOff As Double = 0.25
        If (nTOSCir > 0) Then
            nDiff = Int(nGivenTOSCir - nGivenLargestCir)
            If (nDiff < 0) Then
                nLength = nDiff * nBodyFrontReduceOff
            Else
                nLength = nDiff * nBodyFrontIncreaseOff
            End If
            xyCO_TOSBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
            xyCO_TOSBott.X = xyO.X + FNRound((nTOSHt - nDatum) * nXscale)
        End If
        ''// Waist  (Cut Out Front)
        nDiff = Int(nGivenWaistCir - nGivenLargestCir)
        If (nDiff < 0) Then
            nLength = nDiff * nBodyFrontReduceOff
        Else
            nLength = nDiff * nBodyFrontIncreaseOff
        End If
        xyCO_WaistBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
        xyCO_WaistBott.X = xyO.X + FNRound((nWaistHt - nDatum) * nXscale)

        ''// MidPoint  (Cut Out Front)
        ''// Remember that mid point height was established earlier
        nDiff = Int(nGivenMidPointCir - nGivenLargestCir)
        If (nDiff < 0) Then
            nLength = nDiff * nBodyFrontReduceOff
        Else
            nLength = nDiff * nBodyFrontIncreaseOff
        End If
        xyCO_MidPointBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
        xyCO_MidPointBott.X = xyO.X + nMidPointHt

        ''// Cut Out (Back)
        ''// Revise Cut Out Diameter to complete cut out
        Dim nCO_Width, nOriginalCO_WaistTopY, nOriginalCO_TOSTopY As Double
        Dim nCutOutConstructFac_1 As Double = 0.87
        Dim nEndBackBodyOff As Double = 1.25
        Dim xyCO_WaistTop, xyCO_TOSTop, xyCO_MidPointTop As WHBODDIA1.XY
        nCO_Width = nCO_HalfWidth * 2 * nCutOutConstructFac_1

        ''// Waist  (Cut Out Back)
        xyCO_WaistTop.Y = xyCO_WaistBott.Y + nCO_Width
        xyCO_WaistTop.X = xyCO_WaistBott.X

        ''// Top Of Support   (Cut Out Back)
        xyCO_TOSTop.Y = xyCO_WaistTop.Y
        If (nTOSCir > 0) Then
            xyCO_TOSTop.X = xyCO_TOSBott.X
        Else
            xyCO_TOSTop.X = xyCO_WaistTop.X + nEndBackBodyOff
            If bIsBrief Then
                xyCO_TOSBott.Y = xyCO_WaistBott.Y
                xyCO_TOSBott.X = xyCO_TOSTop.X
            End If
        End If

        ''// Back of body
        ''// Get Back body offsets
        ''// Check against Minimun value
        Dim IgnoreMidPoint_CO As Boolean = False
        nOriginalCO_WaistTopY = xyCO_WaistTop.Y
        nOriginalCO_TOSTopY = xyCO_TOSTop.Y

        Dim bIsLoop As Boolean = True
        Dim nWaistBackOff, nTOSBackOff, aAngle, nMidPointBackOff, nMinBackOff As Double
        Dim nBodyBackCutOutMinTol As Double = 3.0
        While (bIsLoop)
            '' // Waist offset
            nWaistBackOff = FNRound(nWaistCir * nYscale)
            nWaistBackOff = (nWaistBackOff - ((xyCO_WaistTop.Y - xyO.Y - nSeam) + (xyCO_WaistBott.Y - xyO.Y - nSeam))) / 2
            ''// TOS Offset
            If (nTOSCir > 0) Then
                nTOSBackOff = FNRound(nTOSCir * nYscale)
                nTOSBackOff = (nTOSBackOff - ((xyCO_TOSTop.Y - xyO.Y - nSeam) + (xyCO_TOSBott.Y - xyO.Y - nSeam))) / 2
            Else
                nTOSBackOff = nWaistBackOff
            End If

            ''// MidPoint Offset
            nLength = xyCO_MidPointBott.X - xyCO_LargestBott.X
            aAngle = FN_CalcAngle(xyCO_LargestTop, xyCO_WaistTop)
            xyCO_MidPointTop.X = xyCO_MidPointBott.X
            ''xyCO_MidPointTop.Y = xyCO_LargestTop.Y + ((System.Math.Tan(aAngle) * (180 / PI)) * nLength)
            xyCO_MidPointTop.Y = xyCO_LargestTop.Y + ((System.Math.Tan(aAngle * (PI / 180))) * nLength)
            nMidPointBackOff = FNRound(nMidPointCir * nYscale)
            nMidPointBackOff = (nMidPointBackOff - ((xyCO_MidPointTop.Y - xyO.Y - nSeam) + (xyCO_MidPointBott.Y - xyO.Y - nSeam))) / 2
            ''// Check that 3" distance is meet
            ''// Note use of 1/2 scale (nYscale) w.r.t.  nBodyBackCutOutMinTol
            bIsLoop = False
            If (IgnoreMidPoint_CO) Then
                nMinBackOff = FNRound(WHBODDIA1.min(nWaistBackOff, nTOSBackOff))
            Else
                nMinBackOff = FNRound(WHBODDIA1.min(WHBODDIA1.min(nWaistBackOff, nTOSBackOff), nMidPointBackOff))
            End If
            If nMinBackOff < (nBodyBackCutOutMinTol * nYscale) Then
                bIsLoop = True
                nDiff = FNRound(((nBodyBackCutOutMinTol * nYscale) - nMinBackOff) * 2)
                If nDiff = 0 Then
                    Exit While
                End If
                xyCO_WaistTop.Y = xyCO_WaistTop.Y - nDiff
                xyCO_TOSTop.Y = xyCO_TOSTop.Y - nDiff
                If xyCO_WaistTop.Y < xyCO_WaistBott.Y Then
                    IgnoreMidPoint_CO = True
                    xyCO_WaistTop.Y = nOriginalCO_WaistTopY
                    xyCO_TOSTop.Y = nOriginalCO_TOSTopY
                    MsgBox("Warning, Mid-Point ignored in calculating back Cut-Out!" +
                           "\nDistance between back of Cut-Out at Mid-Point and Profile may be less than 3", 16, "Waist - Dialog")
                End If
            End If
        End While

        ''// Calculate back body XY points
        ''// Waist
        Dim xyWaist, xyTOS, xyCO_CenterArrow, xyCO_ArcCen As WHBODDIA1.XY
        xyWaist.X = xyCO_WaistTop.X
        xyWaist.Y = xyCO_WaistTop.Y + nWaistBackOff

        ''// Top Of Support
        xyTOS.X = xyCO_TOSTop.X
        xyTOS.Y = xyCO_TOSTop.Y + nTOSBackOff
        ''// MidPoint
        xyMidPoint.X = xyCO_MidPointTop.X
        xyMidPoint.Y = xyCO_MidPointTop.Y + nMidPointBackOff
        ''// Center point of cutout (also used in crotch labeling)
        Dim nCutOutConstructOff_1 As Double = 0.75
        xyCO_CenterArrow.X = xyFold.X + nCutOutConstructOff_1
        xyCO_CenterArrow.Y = xyCO_LargestBott.Y + nCO_HalfWidth

        Dim xyThigh As WHBODDIA1.XY
        If bIsBrief Then
            ''// Thigh position
            xyThigh.x = xyCO_CenterArrow.X - nOpenOff - nBriefOff
            xyThigh.y = xyO.Y + (nThighCirOriginal * 0.95) * nYscale + nSeam
        End If

        ''// Establish center, And angles of cutout arc for drawing purposes	
        xyCO_ArcCen.Y = xyCO_CenterArrow.Y
        Dim nCO_ArcRadius As Double = nCO_HalfWidth
        Dim aCO_ArcStart, aCO_ArcDelta, nArcCenToLargestOffset, nTopArcSegment, aCO_ArcStartOriginal As Double
        xyCO_ArcCen.X = xyCO_CenterArrow.X + nCO_ArcRadius
        If (xyCO_ArcCen.X > xyCO_LargestTop.X) Then
            nLength = WHBODDIA1.FN_CalcLength(xyCO_CenterArrow, xyCO_LargestTop) / 2
            aAngle = FN_CalcAngle(xyCO_CenterArrow, xyCO_LargestTop)
            ''nCO_ArcRadius = nLength / (System.Math.Cos(aAngle) * (180 / PI))
            nCO_ArcRadius = nLength / System.Math.Cos(aAngle * (PI / 180))
            xyCO_ArcCen.X = xyCO_CenterArrow.X + nCO_ArcRadius
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyCO_LargestTop)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyCO_LargestBott) - aCO_ArcStart
        Else
            aCO_ArcStart = 90
            aCO_ArcDelta = 180
        End If

        ''// Additions for OPEN Crotch
        Dim xyCO_OpenArcTop, xyCO_OpenArcBott, xyPt1, xyPt2, xyInt As WHBODDIA1.XY
        '' For W4OC Left and W4OC Right
        Dim xyCO_CutMark, xyCutMark As WHBODDIA1.XY
        Dim nCuttingMarkOff As Double = nOpenOff
        Dim nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta, nTopArcToLargestOffset As Double
        If OpenCrotch Then
            ''//Get Cut out Arc Center to Largest part of buttocks offset
            nArcCenToLargestOffset = xyCO_LargestTop.X - xyCO_ArcCen.X
            If nArcCenToLargestOffset < 0 Then
                nArcCenToLargestOffset = 0 '' //Make sure Not -ve
            End If
            ''// For back / top of crotch establish length of top arc
            nTopArcSegment = (180 - aCO_ArcStart) * PI / 180 * nCO_ArcRadius
            aCO_ArcStartOriginal = aCO_ArcStart '' // Store For possible use With bottom arc
            If nTopArcSegment > nOpenBack Then
                ''	// End Of Top O.C. lands On arc  
                xyCO_OpenArcTop = xyCO_ArcCen
                nCO_OpenArcTopRadius = nCO_ArcRadius
                aCO_OpenArcTopStart = aCO_ArcStart
                aAngle = 180 - (nOpenBack * 180) / (PI * nCO_ArcRadius)
                aCO_OpenArcTopDelta = aAngle - aCO_OpenArcTopStart
                ''// Modify start of Cut-Out arc
                aCO_ArcStart = aCO_OpenArcTopStart + aCO_OpenArcTopDelta
            Else
                ''//Get distance from end of arc to lagest part of buttocks
                nTopArcToLargestOffset = nArcCenToLargestOffset - (nOpenBack - nTopArcSegment)
            End If
            ''// Do Not go past Largest part of buttocks
            If (nTopArcToLargestOffset < 0) Then
                nTopArcToLargestOffset = 0
            End If

            ''// For front / bottom of crotch establish length of arc
            Dim nBottArcSegment As Double
            nBottArcSegment = ((aCO_ArcStartOriginal + aCO_ArcDelta) - 180) * PI / 180 * nCO_ArcRadius
            If (nBottArcSegment > nOpenFront) Then ''	// End Of Bottom O.C. lands On arc  
                xyCO_OpenArcBott = xyCO_ArcCen
                nCO_OpenArcBottRadius = nCO_ArcRadius
                aCO_OpenArcBottStart = 180 + (nOpenFront * 180) / (PI * nCO_ArcRadius)
                aCO_OpenArcBottDelta = aCO_ArcStartOriginal + aCO_ArcDelta - aCO_OpenArcBottStart
                ''// Modify end of Cut-Out arc
                aCO_ArcDelta = aCO_OpenArcBottStart - aCO_ArcStart
            Else
                ''//Get distance from end of arc to lagest part of buttocks
                nBottArcToLargestOffset = nArcCenToLargestOffset - (nOpenFront - nBottArcSegment)
                ''// Change aCO_ArcDelta if Start has moved (New bit)
                If (aCO_ArcStartOriginal <> aCO_ArcStart) Then
                    aCO_ArcDelta = aCO_ArcDelta - (aCO_ArcStart - aCO_ArcStartOriginal)
                End If
            End If
            ''// Do Not go past Largest part of buttocks
            If (nBottArcToLargestOffset < 0) Then
                nBottArcToLargestOffset = 0
            End If

            ''// Establish revised arc for cut out including offset for open crotch
            ''// Ensure that End Points of arc do Not pass largest part of buttocks

            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, xyPt1)
            If (xyPt1.X > xyCO_LargestTop.X) Then
                xyPt1 = xyCO_LargestTop
                nTopArcToLargestOffset = 0
            End If
            If bIsBrief = False Then
                xyCO_CutMark = xyPt1
            End If
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta, xyPt2)
            If (xyPt2.X > xyCO_LargestBott.X) Then
                xyPt2 = xyCO_LargestBott
                nBottArcToLargestOffset = 0
            End If
            PR_MakeXY(xyEnd, xyPt1.X, xyPt1.Y + 2)
            Dim nError As Short = FN_CirLinInt(xyPt1, xyEnd, xyCO_ArcCen, nCO_ArcRadius + nOpenOff, xyInt)
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyInt)
            PR_MakeXY(xyEnd, xyPt2.X, xyPt2.Y - 2)
            nError = FN_CirLinInt(xyPt2, xyEnd, xyCO_ArcCen, nCO_ArcRadius + nOpenOff, xyInt)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyInt) - aCO_ArcStart

            nCO_ArcRadius = nCO_ArcRadius + nOpenOff
        End If

        ''//
        ''// DRAW Cutout And TOS
        ''//
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("marker","xmarker",xyO , 0.1, 0.1) 
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "O")
        PR_DrawXMarker(xyWHCutInsert)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "O")
        ''hEnt = AddEntity("marker", "xmarker", xyTOS, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "TOS")
        PR_DrawXMarker(xyTOS)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "TOS")
        Dim xyLastTOS As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastTOS)
        ''hEnt = AddEntity("marker", "xmarker", xyWaist, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Waist")
        PR_DrawXMarker(xyWaist)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "Waist")
        Dim xyLastWaist As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastWaist)
        ''hEnt = AddEntity("marker", "xmarker", xyMidPoint, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "MidPoint")
        PR_DrawXMarker(xyMidPoint)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "MidPoint")
        Dim xyLastMidPoint As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastMidPoint)
        ''hEnt = AddEntity("marker", "xmarker", xyLargest, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Largest")
        PR_DrawXMarker(xyLargest)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "Largest")
        Dim xyLastLargest As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastLargest)
        ''hEnt = AddEntity("marker", "xmarker", xyButtockArcCen, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "ButtockArcCen")
        PR_DrawXMarker(xyButtockArcCen)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "ButtockArcCen")
        Dim xyLastButtockArcCen As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastButtockArcCen)
        If bIsBrief = False Then
            ''hEnt = AddEntity("marker", "xmarker", xyCO_TOSBott, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_TOSBott")
            PR_DrawXMarker(xyCO_TOSBott)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_TOSBott")
            ''hEnt = AddEntity("marker", "xmarker", xyCO_TOSTop, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_TOSTop")
            PR_DrawXMarker(xyCO_TOSTop)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_TOSTop")
        End If
        ''hEnt = AddEntity("marker", "xmarker", xyCO_WaistBott, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_WaistBott")
        PR_DrawXMarker(xyCO_WaistBott)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_WaistBott")
        If bIsBrief = False Then
            ''hEnt = AddEntity("marker", "xmarker", xyCO_WaistTop, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_WaistTop")
            PR_DrawXMarker(xyCO_WaistTop)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_WaistTop")
            ''hEnt = AddEntity("marker", "xmarker", xyCO_MidPointBott, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_MidPointBott")
            PR_DrawXMarker(xyCO_MidPointBott)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_MidPointBott")
            ''hEnt = AddEntity("marker", "xmarker", xyCO_MidPointTop, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_MidPointTop")
            PR_DrawXMarker(xyCO_MidPointTop)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_MidPointTop")
        End If
        '' hEnt = AddEntity("marker", "xmarker", xyFold, 0.1, 0.1)
        '' SetDBData(hEnt, "ID", sFileNo + sLeg + "Fold")
        PR_DrawXMarker(xyFold)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "Fold")
        Dim xyLastFold As WHBODDIA1.XY
        PR_GetBlkPosition(xyLastFold)
        Dim xyLastFoldPanty As WHBODDIA1.XY = xyFoldPanty
        If (PantyLeg Or bIsBrief) Then
            ''hEnt = AddEntity("marker", "xmarker", xyFoldPanty, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "FoldPanty")
            PR_DrawXMarker(xyFoldPanty)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "FoldPanty")
            PR_GetBlkPosition(xyLastFoldPanty)
        End If
        nFoldHt = WHBODDIA1.fnDisplayToInches(Val(g_sFoldHt))
        saveCutOutInfoToDWG(xyWHCutInsert, xyTOS, xyWaist, xyMidPoint, xyLargest, xyFold, xyFoldPanty, xyButtockArcCen, nFoldHt, xyThigh, xyCO_WaistBott, xyCO_ArcCen)
        ''saveCutOutInfoToDWG(xyWHCutInsert, xyLastTOS, xyLastWaist, xyLastMidPoint, xyLastLargest, xyLastFold, xyLastFoldPanty, xyLastButtockArcCen, nFoldHt)
        If bIsBrief = False Then
            '' hEnt = AddEntity("marker", "xmarker", xyCO_LargestTop, 0.1, 0.1) 	
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_LargestTop")
            ''SetDBData(hEnt, "Data", MakeString("scalar", nTopArcToLargestOffset))
            PR_DrawXMarker(xyCO_LargestTop)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_LargestTop")
            PR_AddDBValueToLast("Data", Str(nTopArcToLargestOffset))

            ''hEnt = AddEntity("marker", "xmarker", xyCO_LargestBott, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_LargestBott") ;
            ''SetDBData(hEnt, "Data", MakeString("scalar", nBottArcToLargestOffset));
            PR_DrawXMarker(xyCO_LargestBott)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_LargestBott")
            PR_AddDBValueToLast("Data", Str(nBottArcToLargestOffset))
        End If

        ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1) ;		
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_ArcCen") ;
        'SetDBData(hEnt, "Data", MakeString("scalar", nCO_ArcRadius) +
        '          " " + MakeString("scalar", aCO_ArcStart) +
        '          " " + MakeString("scalar", aCO_ArcDelta)) ;
        PR_DrawXMarker(xyCO_ArcCen)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_ArcCen")
        PR_AddDBValueToLast("Data", Str(nCO_ArcRadius) + " " + Str(aCO_ArcStart) + " " + Str(aCO_ArcDelta))
        If bIsBrief = False Then
            If xyCO_OpenArcBott.X <> 0 And xyCO_OpenArcBott.Y <> 0 Then
                ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1)
                ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_OpenArcBott")
                'SetDBData(hEnt, "Data", MakeString("scalar", nCO_OpenArcBottRadius) +
                '          " " + MakeString("scalar", aCO_OpenArcBottStart) +
                '          " " + MakeString("scalar", aCO_OpenArcBottDelta))
                PR_DrawXMarker(xyCO_ArcCen)
                PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_OpenArcBott")
                PR_AddDBValueToLast("Data", Str(nCO_OpenArcBottRadius) + " " + Str(aCO_OpenArcBottStart) + " " + Str(aCO_OpenArcBottDelta))
            End If

            If xyCO_OpenArcTop.X <> 0 And xyCO_OpenArcTop.Y <> 0 Then
                ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1) ;		
                ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_OpenArcTop") ;
                'SetDBData(hEnt, "Data", MakeString("scalar", nCO_OpenArcTopRadius) +
                '          " " + MakeString("scalar", aCO_OpenArcTopStart) +
                '          " " + MakeString("scalar", aCO_OpenArcTopDelta)) ;
                PR_DrawXMarker(xyCO_ArcCen)
                PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_OpenArcTop")
                PR_AddDBValueToLast("Data", Str(nCO_OpenArcTopRadius) + " " + Str(aCO_OpenArcTopStart) + " " + Str(aCO_OpenArcTopDelta))
            End If
        End If

        ''hEnt = AddEntity("marker", "xmarker", xyCO_CenterArrow, 0.1, 0.1) ;	
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_CenterArrow");
        ''SetDBData(hEnt, "Data", MakeString("scalar", nOpenOff)) ; 
        PR_DrawXMarker(xyCO_CenterArrow)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "CO_CenterArrow")
        PR_AddDBValueToLast("Data", Str(nOpenOff))
        ''// Check for special case where the user has allowed a difference of more than 5%
        ''// Between reductions add a marker 5% away from existing marker at fold.
        Dim xyTmp As WHBODDIA1.XY
        If System.Math.Abs(nLargestGivenRed - nThighGivenRed) > 0.05 Then
            xyTmp.X = xyFold.X
            xyTmp.Y = xyO.Y + FNRound(nThighCirOriginal * (nThighGivenRed + 0.05)) * nYscale + nSeam
            ''hEnt = AddEntity("marker", "xmarker", xyTmp, 0.1, 0.1) ;	
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Fold+5%");
            PR_DrawXMarker(xyTmp)
            PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "Fold+5%")
        End If

        If (LeftLeg) Then
            ARMDIA1.PR_SetLayer("TemplateLeft")
        Else
            ARMDIA1.PR_SetLayer("TemplateRight")
        End If
        ''//
        ''// DRAW Cutout And TOS
        ''//
        Dim xyFilletTop As WHBODDIA1.XY
        Dim nFilletRadius As Double = 0.1875
        If (OpenCrotch) Then
            ''// Get Start point of arc 
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, xyPt1)
            ''// TOP Fillet radius
            PR_MakeXY(xyStart, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y)
            PR_MakeXY(xyEnd, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y + nCO_ArcRadius)
            Dim nError As Short = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
            xyFilletTop = xyInt
            ''// TOP
            If xyCO_OpenArcTop.X <> 0 And xyCO_OpenArcTop.Y <> 0 Then
                ''AddEntity("arc", xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta)
                PR_DrawArc(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta)
                PR_CalcPolar(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart + aCO_OpenArcTopDelta, xyPt2)
                ''----AddEntity("line", xyPt1.X, xyFilletTop.Y, xyPt2)
                PR_MakeXY(xyStart, xyPt1.X, xyFilletTop.Y)
                PR_DrawLine(xyStart, xyPt2)
                If WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 3 Or WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 4 Then
                    ''// Note:- xyCO_CutMark from above in this case
                    PR_MakeXY(xyStart, xyCO_CutMark.X, xyCO_CutMark.Y + 10)
                    nError = FN_CirLinInt(xyCO_CutMark, xyStart, xyButtockArcCen, WHBODDIA1.FN_CalcLength(xyButtockArcCen, xyFold), xyInt)
                    xyCutMark.Y = xyInt.Y - nCuttingMarkOff
                    xyCutMark.X = xyInt.X - nCuttingMarkOff

                    ''// Top fillet position 
                    'Dim aCutMark As Double = FN_CalcAngle(xyCO_CutMark, xyCutMark)
                    ''nError = FN_CirLinInt(CalcXY("relpolar", xyCO_CutMark, nFilletRadius, aCutMark + 90),
                    ''       CalcXY("relpolar", xyCutMark, nFilletRadius, aCutMark + 90),
                    ''        xyCO_ArcCen, nCO_ArcRadius - nFilletRadius)
                    'PR_CalcPolar(xyCO_CutMark, nFilletRadius, aCutMark + 90, xyStart)
                    'PR_CalcPolar(xyCutMark, nFilletRadius, aCutMark + 90, xyEnd)
                    'nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
                    'xyFilletTop = xyInt

                    ''// Revise cutout arc Start angle And Delta angle
                    'aCO_ArcDelta = aCO_ArcDelta - (FN_CalcAngle(xyCO_ArcCen, xyFilletTop) - aCO_ArcStart)
                    'aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyFilletTop)
                    'PR_CalcPolar(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart + aCO_OpenArcTopDelta, xyPt2)
                    ''AddEntity("line", CalcXY("relpolar", xyFilletTop, nFilletRadius, aCutMark - 90), xyPt2)
                    'PR_CalcPolar(xyFilletTop, nFilletRadius, aCutMark - 90, xyInt)
                    'PR_DrawLine(xyInt, xyPt2)
                End If
                PR_CalcPolar(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, xyTmp)
                If (xyTmp.X < xyCO_LargestTop.X) Then
                    ''AddEntity("line", xyTmp, xyCO_LargestTop)
                    PR_DrawLine(xyTmp, xyCO_LargestTop)
                End If
                ''// Fillet
                ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
            Else
                nLength = xyCO_LargestTop.X - xyPt1.X - nTopArcToLargestOffset
                If WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 3 Or WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 4 Then
                    PR_CalcPolar(xyPt1, nLength, 0, xyPt2)
                    xyCO_CutMark.X = xyPt2.X
                    xyCO_CutMark.Y = xyPt2.Y - nOpenOff
                    ''nError = FN_CirLinInt(xyCO_CutMark, xyCO_CutMark.X, xyCO_CutMark.Y + 10, xyButtockArcCen, Calc("length", xyButtockArcCen, xyFold))
                    PR_MakeXY(xyStart, xyCO_CutMark.X, xyCO_CutMark.Y + 10)
                    FN_CirLinInt(xyCO_CutMark, xyStart, xyButtockArcCen, WHBODDIA1.FN_CalcLength(xyButtockArcCen, xyFold), xyInt)
                    xyCutMark.Y = xyInt.Y - nCuttingMarkOff
                    xyCutMark.X = xyInt.X - nCuttingMarkOff
                End If
                If (nLength > 0) Then
                    PR_CalcPolar(xyPt1, nLength, 0, xyPt2)
                    If (nLength > nFilletRadius) Then
                        xyFilletTop = xyPt1 ''	// gets aCO_ArcStart correct
                        ''AddEntity("arc", xyPt2.X - nFilletRadius, xyPt2.Y - nFilletRadius, nFilletRadius, 0, 90)
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyPt2.Y - nFilletRadius)
                        PR_DrawArc(xyStart, nFilletRadius, 0, 90)
                        ''AddEntity("line", xyPt1, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_DrawLine(xyPt1, xyEnd)
                        ''AddEntity("line", xyPt2.X, xyPt2.Y - nFilletRadius, xyPt2.X, xyCO_LargestTop.Y)
                        PR_MakeXY(xyStart, xyPt2.X, xyPt2.Y - nFilletRadius)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestTop.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    Else
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y + nCO_ArcRadius)
                        nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
                        xyFilletTop = xyInt
                        ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                        PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
                        ''AddEntity("line", xyPt2.X, xyFilletTop.Y, xyPt2.X, xyCO_LargestTop.Y)
                        PR_MakeXY(xyStart, xyPt2.X, xyFilletTop.Y)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestTop.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    End If

                Else
                    ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                    PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
                    ''AddEntity("line", xyPt1.X, xyFilletTop.Y, xyCO_LargestTop)
                    PR_MakeXY(xyStart, xyPt1.X, xyFilletTop.Y)
                    PR_DrawLine(xyStart, xyCO_LargestTop)
                End If

                If (nTopArcToLargestOffset > 0) Then
                    'AddEntity("line",
                    '              CalcXY("relpolar", xyCO_LargestTop, nTopArcToLargestOffset, 180),
                    '                          xyCO_LargestTop)
                    PR_CalcPolar(xyCO_LargestTop, nTopArcToLargestOffset, 180, xyStart)
                    PR_DrawLine(xyStart, xyCO_LargestTop)
                End If
            End If
            ''AddEntity("line", xyCO_LargestTop, xyCO_WaistTop) ;
            PR_DrawLine(xyCO_LargestTop, xyCO_WaistTop)
            '' AddEntity("line", xyCO_WaistTop, xyCO_TOSTop) ;
            PR_DrawLine(xyCO_WaistTop, xyCO_TOSTop)

            ''// BOTTOM
            ''// End points of arc 
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta, xyPt1)

            ''// BOTTOM Fillet radius initial value
            Dim xyFilletBott As WHBODDIA1.XY
            PR_MakeXY(xyStart, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y)
            PR_MakeXY(xyEnd, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y - nCO_ArcRadius)
            nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
            xyFilletBott = xyInt

            If xyCO_OpenArcBott.X <> 0 And xyCO_OpenArcBott.Y <> 0 Then
                ''AddEntity("arc", xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta) ;
                PR_DrawArc(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta)
                PR_CalcPolar(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, xyPt2)
                ''AddEntity("line", xyPt1.X, xyFilletBott.Y, xyPt2) ;
                PR_MakeXY(xyStart, xyPt1.X, xyFilletBott.Y)
                PR_DrawLine(xyStart, xyPt2)
                PR_CalcPolar(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart + aCO_OpenArcBottDelta, xyTmp)
                If (xyTmp.X < xyCO_LargestBott.X) Then
                    ''AddEntity ("line", xyTmp, xyCO_LargestBott)
                    PR_DrawLine(xyTmp, xyCO_LargestBott)
                End If
                ''// Fillet
                aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
            Else
                nLength = xyCO_LargestBott.X - xyPt1.X - nBottArcToLargestOffset
                If (nLength > 0) Then
                    PR_CalcPolar(xyPt1, nLength, 0, xyPt2)
                    If (nLength > nFilletRadius) Then
                        xyFilletBott = xyPt1    ''// gets aCO_ArcStart correct
                        ''AddEntity("arc", xyPt2.X - nFilletRadius, xyPt2.Y + nFilletRadius, nFilletRadius, 270, 90);
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyPt2.Y + nFilletRadius)
                        PR_DrawArc(xyStart, nFilletRadius, 270, 90)
                        ''AddEntity("line", xyPt1, xyPt2.X - nFilletRadius, xyPt2.Y) ;
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_DrawLine(xyPt1, xyEnd)
                        ''AddEntity("line", xyPt2.X, xyPt2.Y + nFilletRadius, xyPt2.X, xyCO_LargestBott.Y);
                        PR_MakeXY(xyStart, xyPt2.X, xyPt2.Y + nFilletRadius)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestBott.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    Else
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y - nCO_ArcRadius)
                        nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
                        xyFilletBott = xyInt
                        aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                        ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                        PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
                        ''AddEntity("line", xyPt2.x, xyFilletBott.y, xyPt2.x, xyCO_LargestBott.y);
                        PR_MakeXY(xyStart, xyPt2.X, xyFilletBott.Y)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestBott.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    End If
                Else
                    aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                    ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                    PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
                    ''AddEntity("line", xyPt1.x, xyFilletBott.y, xyCO_LargestBott);
                    PR_MakeXY(xyStart, xyPt1.X, xyFilletBott.Y)
                    PR_DrawLine(xyStart, xyCO_LargestBott)
                End If

                If (nBottArcToLargestOffset > 0) Then
                    'AddEntity("line",
                    '      CalcXY("relpolar", xyCO_LargestBott, nBottArcToLargestOffset, 180),
                    '                  xyCO_LargestBott);
                    PR_CalcPolar(xyCO_LargestBott, nBottArcToLargestOffset, 180, xyStart)
                    PR_DrawLine(xyStart, xyCO_LargestBott)
                End If
            End If
            ''StartPoly("openfitted")
            '' AddVertex(xyCO_LargestBott)
            ''AddVertex(xyCO_MidPointBott)
            ''AddVertex(xyCO_WaistBott)
            Dim ptColl As Point3dCollection = New Point3dCollection()
            ptColl.Add(New Point3d(xyCO_LargestBott.X, xyCO_LargestBott.Y, 0))
            ptColl.Add(New Point3d(xyCO_MidPointBott.X, xyCO_MidPointBott.Y, 0))
            ptColl.Add(New Point3d(xyCO_WaistBott.X, xyCO_WaistBott.Y, 0))
            If (nTOSCir > 0 Or bIsBrief) Then
                ''AddVertex(xyCO_TOSBott)
                ptColl.Add(New Point3d(xyCO_TOSBott.X, xyCO_TOSBott.Y, 0))
            End If
            ''EndPoly()
            PR_DrawPoly(ptColl)

            ''// Draw arc taking account of fillets
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyFilletTop)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyFilletBott) - aCO_ArcStart
            ''AddEntity("arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
            PR_DrawArc(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta)
        Else
            ''// Closed Crotch
            ''AddEntity("arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
            PR_DrawArc(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta)
            ''// TOP
            ''AddEntity("line", xyCO_LargestTop, xyCO_WaistTop) ;
            PR_DrawLine(xyCO_LargestTop, xyCO_WaistTop)
            '' AddEntity("line", xyCO_WaistTop, xyCO_TOSTop) ;
            PR_DrawLine(xyCO_WaistTop, xyCO_TOSTop)
            ''// FIDDLY BITS, If center of cut out Is below largest part of buttocks then put in joining lines 
            If (xyCO_ArcCen.X < xyCO_LargestTop.X) Then
                ''AddEntity("line", CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, 90), xyCO_LargestTop);
                PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, 90, xyStart)
                PR_DrawLine(xyStart, xyCO_LargestTop)
                ''AddEntity("line", CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, 270), xyCO_LargestBott);
                PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, 270, xyStart)
                PR_DrawLine(xyStart, xyCO_LargestBott)
            End If
            ''// BOTTOM
            ''StartPoly("openfitted")
            ''AddVertex(xyCO_LargestBott)
            ''AddVertex(xyCO_MidPointBott)
            ''AddVertex(xyCO_WaistBott)
            Dim ptColl1 As Point3dCollection = New Point3dCollection()
            ptColl1.Add(New Point3d(xyCO_LargestBott.X, xyCO_LargestBott.Y, 0))
            ptColl1.Add(New Point3d(xyCO_MidPointBott.X, xyCO_MidPointBott.Y, 0))
            ptColl1.Add(New Point3d(xyCO_WaistBott.X, xyCO_WaistBott.Y, 0))
            If (nTOSCir > 0 Or bIsBrief) Then
                ''AddVertex(xyCO_TOSBott)
                ptColl1.Add(New Point3d(xyCO_TOSBott.X, xyCO_TOSBott.Y, 0))
            End If
            ''EndPoly()
            PR_DrawPoly(ptColl1)
        End If
        ''// Closing Line at waist Or TOS
        ''AddEntity("line", xyCO_TOSTop, xyTOS) ;
        PR_DrawLine(xyCO_TOSTop, xyTOS)
        If (nTOSCir > 0 Or bIsBrief) Then
            ''AddEntity("line", xyCO_TOSBott, xyTOS.X, xyO.Y)
            PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
            PR_DrawLine(xyCO_TOSBott, xyEnd)
        Else
            ''AddEntity("line", xyCO_WaistBott, xyCO_WaistBott.X, xyO.Y)
            PR_MakeXY(xyEnd, xyCO_WaistBott.X, xyO.Y)
            PR_DrawLine(xyCO_WaistBott, xyEnd)
        End If
        If bIsBrief Then
            ''AddEntity  ("line", xyOtemplate, xyTOS.x, xyOtemplate.y )
            PR_MakeXY(xyEnd, xyTOS.X, xyWHCutInsert.Y)
            PR_DrawLine(xyWHCutInsert, xyEnd)
        End If

        ''// Draw construction lines
        ''// 
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("line", xyFold.X, xyLargest.Y, xyFold.X, xyO.Y) ;
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "FoldLine") ; 
        PR_MakeXY(xyStart, xyFold.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyFold.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        PR_AddDBValueToLast("ID", txtFileNo.Text + sLeg + "FoldLine")
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir * nYscale + nSeam) + " " + MakeString("scalar", nGivenWaistCir));     	
        PR_AddDBValueToLast("Data", Str(nThighCir * nYscale + nSeam) + " " + Str(nGivenWaistCir))

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
            ''hEnt = AddEntity("line", xyTOS.X, xyLargest.Y, xyTOS.X, xyO.Y) ;
            PR_MakeXY(xyStart, xyTOS.X, xyLargest.Y)
            PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
            PR_DrawLine(xyStart, xyEnd)
        End If
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "EOSLine") ; 
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir * nYscale + nSeam) + " " + MakeString("scalar", nGivenWaistCir));     	
        PR_AddDBValueToLast("EOSID", txtFileNo.Text + sLeg + "EOSLine")
        PR_AddDBValueToLast("Data", Str(nThighCir * nYscale + nSeam) + " " + Str(nGivenWaistCir))

        ''AddEntity("arc", xyButtockArcCen, Calc("length", xyButtockArcCen, xyFold), 60, 70) ;
        PR_DrawArc(xyButtockArcCen, WHBODDIA1.FN_CalcLength(xyButtockArcCen, xyFold), 60, 70)

        ''// Add Body template tapes
        ''// From body position given 
        Dim nn As Double
        nLastTape = 1
        If (nBodyLegTapePos > 0) Then
            nn = nBodyLegTapePos - 1
        Else
            nn = nLastTape - 1
        End If
        ''xyPt1.X = xyO.X + 0.625
        xyPt1.X = xyO.X 
        xyPt1.Y = xyO.Y + nSeam + 0.5
        Dim g_sTextList As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 3637½ 3940½ 42"
        While (xyPt1.X < xyTOS.X)
            'sSymbol = MakeString("long", nn) + "tape"
            'If (!Symbol("find", sSymbol)) Then
            '    MsgBox("Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists\n", 16, "Waist - Dialog")
            '    Exit While
            'End If
            'AddEntity("symbol", sSymbol, xyPt1)
            Dim sSymbol As String = Str(nn) + "tape"
            Dim strTape As String = LTrim(Mid(g_sTextList, ((nn - 1) * 3) + 1, 3))
            PR_DrawRuler(sSymbol, xyPt1, strTape)
            xyPt1.X = xyPt1.X + 1.5 * nXscale
            nn = nn + 1
        End While
        ''// Seam TRAM Lines 
        ARMDIA1.PR_SetLayer("Notes")
        ''AddEntity("line", xyWHCutInsert.x, xyWHCutInsert.y + nSeam + 0.5, xyPt1.X - (1.5 * nXscale), xyPt1.Y) ; //*
        PR_MakeXY(xyStart, xyWHCutInsert.X, xyWHCutInsert.Y + nSeam + 0.5)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''AddEntity("line", xyWHCutInsert.x, xyWHCutInsert.y + nSeam, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5) ; //*
        PR_MakeXY(xyStart, xyWHCutInsert.X, xyWHCutInsert.Y + nSeam)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5)
        PR_DrawLine(xyStart, xyEnd)

        ''// Diagonal Fly Marks
        If (sCrotchStyle.Equals("Diagonal Fly") And sLeg.Equals("Left")) Then
            ''AddEntity("marker", "closed arrow", xyCO_LargestTop, 0.5, 0.125, 90)
            PR_DrawWaistClosedArrow(xyCO_LargestTop, 90)
            ''AddEntity("marker", "closed arrow", xyCO_LargestBott, 0.5, 0.125, 270)
            PR_DrawWaistClosedArrow(xyCO_LargestBott, 270)
        End If

        '' For W4OC Left and W4OC Right
        If WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 3 Or WHBODDIA1.fnGetNumber(g_sLegStyle, 1) = 4 Then
            ''//Cutting marks
            ''AddEntity("line", xyCutMark, xyCutMark.X + nCuttingMarkOff, xyCutMark.Y);
            PR_MakeXY(xyEnd, xyCutMark.X + nCuttingMarkOff, xyCutMark.Y)
            PR_DrawLine(xyCutMark, xyEnd)
            ''AddEntity("line", xyWaist, xyCutMark.X + nCuttingMarkOff, xyCutMark.Y);
            PR_DrawLine(xyWaist, xyEnd)
            ''AddEntity("line", xyCutMark, xyCO_CutMark);
            PR_DrawLine(xyCutMark, xyCO_CutMark)
            ''AddEntity("line", xyCutMark.X, xyO.Y, xyCutMark.X, xyCO_LargestBott.Y - 0.5)
            PR_MakeXY(xyStart, xyCutMark.X, xyO.Y)
            PR_MakeXY(xyEnd, xyCutMark.X, xyCO_LargestBott.Y - 0.5)
            PR_DrawLine(xyStart, xyEnd)
        End If

        If bIsBrief = False Then
            If (LeftLeg) Then
                ARMDIA1.PR_SetLayer("TemplateLeft")
            Else
                ARMDIA1.PR_SetLayer("TemplateRight")
            End If
            If (nTOSCir > 0) Then
                ''AddEntity("line", xyWHCutInsert, xyTOS.X, xyWHCutInsert.y)
                PR_MakeXY(xyEnd, xyTOS.X, xyWHCutInsert.Y)
                PR_DrawLine(xyWHCutInsert, xyEnd)
            Else
                ''AddEntity("line", xyWHCutInsert, xyWaist.X, xyWHCutInsert.y)
                PR_MakeXY(xyEnd, xyWaist.X, xyWHCutInsert.Y)
                PR_DrawLine(xyWHCutInsert, xyEnd)
            End If
        End If
        If chkLabel.Checked = True Then
            PR_DrawWaistLabel()
        End If
    End Sub
    Function FN_CalcAngle(ByRef xyStart As WHBODDIA1.XY, ByRef xyEnd As WHBODDIA1.XY) As Double
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
    Sub PR_CalcPolar(ByRef xyStart As WHBODDIA1.XY, ByRef nLength As Double, ByVal nAngle As Double, ByRef xyReturn As WHBODDIA1.XY)
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
    Function FN_CirLinInt(ByRef xyStart As WHBODDIA1.XY, ByRef xyEnd As WHBODDIA1.XY, ByRef xyCen As WHBODDIA1.XY, ByRef nRad As Double, ByRef xyInt As WHBODDIA1.XY) As Short
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
                If nRoot >= WHBODDIA1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
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
                If nRoot >= WHBODDIA1.min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
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
                If nRoot >= WHBODDIA1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
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
    Sub PR_DrawXMarker(ByRef xyInsertion As WHBODDIA1.XY)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As WHBODDIA1.XY
        PR_CalcPolar(xyBase, 0.0625, 135, xyStart)
        PR_CalcPolar(xyBase, 0.0625, -45, xyEnd)
        PR_CalcPolar(xyBase, 0.0625, 45, xySecSt)
        PR_CalcPolar(xyBase, 0.0625, -135, xySecEnd)
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
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                idLastMarker = blkRef.ObjectId
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawArc(ByRef xyCen As WHBODDIA1.XY, ByRef nRad As Double, ByRef nStartAng As Double, ByRef nDeltaAng As Double)
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
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawLine(ByRef xyStart As WHBODDIA1.XY, ByRef xyFinish As WHBODDIA1.XY)
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
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
            End If
            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            idLastMarker = acLine.ObjectId

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
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
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub
    Private Sub saveCutOutInfoToDWG(ByRef xyO As WHBODDIA1.XY, ByRef xyTOS As WHBODDIA1.XY, ByRef xyWaist As WHBODDIA1.XY, ByRef xyMidPoint As WHBODDIA1.XY, ByRef xyLargest As WHBODDIA1.XY,
                                    ByRef xyFold As WHBODDIA1.XY, ByRef xyFoldPanty As WHBODDIA1.XY, ByRef xyButtockArcCen As WHBODDIA1.XY, ByRef nFoldHt As Double,
                                    ByRef xyThigh As WHBODDIA1.XY, ByRef xyCO_WaistBott As WHBODDIA1.XY, ByRef xyCO_ArcCen As WHBODDIA1.XY)
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("WaistCutOut", "WHCUTOUTDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("WaistCutOut", "WHCUTOUTDIC")
            End If

            Dim resbuf As New ResultBuffer
            Dim sString As String = Str(xyO.X) + " " + Str(xyO.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyTOS.X) + " " + Str(xyTOS.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyWaist.X) + " " + Str(xyWaist.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyMidPoint.X) + " " + Str(xyMidPoint.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyLargest.X) + " " + Str(xyLargest.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyFold.X) + " " + Str(xyFold.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyFoldPanty.X) + " " + Str(xyFoldPanty.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), Str(nFoldHt)))
            sString = Str(xyButtockArcCen.X) + " " + Str(xyButtockArcCen.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyThigh.X) + " " + Str(xyThigh.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyCO_WaistBott.X) + " " + Str(xyCO_WaistBott.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            sString = Str(xyCO_ArcCen.X) + " " + Str(xyCO_ArcCen.Y)
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sString))
            _sClass.SetXrecord(resbuf, "WaistCutOut", "WHCUTOUTDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub PR_DrawRuler(ByRef sSymbol As String, ByRef xyStart As WHBODDIA1.XY, ByRef sTape As String)
        Dim xyPt As WHBODDIA1.XY
        Dim xyTapeFst, xyTapeSec, xyTapeEnd, xyTapeText As WHBODDIA1.XY
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
                Dim blkTblRec As BlockTableRecord = acTrans.GetObject(blkRecId, OpenMode.ForWrite)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As Entity = acTrans.GetObject(objID, OpenMode.ForWrite)
                    dbObj.Layer = "Construct"
                Next
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyStart.X, xyStart.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyWHCutInsert.X, xyWHCutInsert.Y, 0)))
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
    Private Sub PR_DrawWaistLabel()
        Dim nAge As Double = Val(txtAge.Text)
        Dim nUnitsFac As Double = WHBODDIA1.g_nUnitsFac
        Dim sFileNo As String = txtFileNo.Text
        Dim Male As Boolean = False
        Dim Female As Boolean = False
        '' Setup SEX flags
        If (txtSex.Text.Equals("Male")) Then
            Male = True
        Else
            Female = True
        End If

        ''-----------------------------------WHLBLBOD.D----------------------------

        Dim sLeftThighCir As String = txtLeftThighCir.Text
        Dim sRightThighCir As String = txtRightThighCir.Text
        Dim sCrotchStyle As String = cboCrotchStyle.Text
        Dim sLine As String = txtLegStyle.Text
        Dim nLeftThighCir As Double = Val(sLeftThighCir)
        Dim nRightThighCir As Double = Val(sRightThighCir)
        Dim nLegStyle, nMaxCir, nLeftLegStyle, nRightLegStyle As Double
        ScanLine(sLine, nLegStyle, nLeftLegStyle, nRightLegStyle, nMaxCir)

        Dim sLeg As String = ""
        Dim LeftLeg, RightLeg As Boolean
        If ((nLeftThighCir - nRightThighCir) > 1) Then
            sLeg = "Right"
            RightLeg = True
            LeftLeg = False
        Else
            sLeg = "Left"
            LeftLeg = True
            RightLeg = False
        End If

        '' Crotch type
        Dim OpenCrotch As Boolean = False
        Dim ClosedCrotch As Boolean = False
        If sCrotchStyle.Contains("Open Crotch") Then
            OpenCrotch = True
        Else
            ClosedCrotch = True
        End If

        '' Leg  type
        Dim PantyLeg As Boolean
        If (nLegStyle = 1) Then
            PantyLeg = True
        Else
            PantyLeg = False
        End If

        ''--------------------WHLBLMKR.D--------------------------
        Dim LeftCO_CenterArrowFound As Boolean = False
        Dim sCO_LargestTopID As String = sFileNo + sLeg + "CO_LargestTop"
        Dim sCO_LargestBottID As String = sFileNo + sLeg + "CO_LargestBott"
        Dim sCO_CenterArrowID As String = ""
        If (nLegStyle = 4) Then
            sCO_CenterArrowID = sFileNo + "Right" + "CO_CenterArrow" ''// Label Right For W4OC Right
        Else
            sCO_CenterArrowID = sFileNo + "Left" + "CO_CenterArrow" ''/ Label Left Crotch only
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        '' Get the Data From X marker
        Dim filterMarker(1) As TypedValue
        filterMarker.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 1)
        Dim selFilterMarker As SelectionFilter = New SelectionFilter(filterMarker)
        Dim selResultMarker As PromptSelectionResult = ed.SelectAll(selFilterMarker)
        If selResultMarker.Status <> PromptStatus.OK Then
            MsgBox("Can't find Marker Details", 16, "Waist Label")
            Exit Sub
        End If
        Dim selSetCLMarker As SelectionSet = selResultMarker.Value
        Dim xyInsertionXmarker, xyCO_CenterArrow, xyCO_LargestTop, xyCO_LargestBott As WHBODDIA1.XY
        Dim nOpenOff As Double
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selSetCLMarker.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                PR_MakeXY(xyInsertionXmarker, position.X, position.Y)
                If xyInsertionXmarker.X = 0 And xyInsertionXmarker.Y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim sTmp As String = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If sTmp.Contains(sCO_CenterArrowID) Then
                    xyCO_CenterArrow = xyInsertionXmarker
                    LeftCO_CenterArrowFound = True
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    Dim sData As String = ""
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sData = sData & typeVal.Value
                        Next
                    End If
                    nOpenOff = Val(sTmp)
                End If
                If sTmp.Contains(sCO_LargestTopID) Then
                    xyCO_LargestTop = xyInsertionXmarker
                End If
                If sTmp.Contains(sCO_LargestBottID) Then
                    xyCO_LargestBott = xyInsertionXmarker
                End If
            Next
        End Using
        If (LeftCO_CenterArrowFound = False) Then
            If (nLegStyle = 4) Then
                MsgBox("RIGHT leg crotch not found to label!" & Chr(10), 16, "Waist Label")
            Else
                MsgBox("LEFT leg crotch not found to label!" & Chr(10), 16, "Waist Label")
            End If
            Exit Sub
        End If

        ''---------------------------WHLBLCRO.D------------------------------
        Dim sCrotchText As String = ""
        Dim sGussetSize As String = ""
        Dim nCrotchSize As Double = 0
        ARMDIA1.PR_SetLayer("Notes")


        ' if (nLegStyle == 2) {
        '  	SetData("TextVertJust", 16)
        ' 	SetData("TextHorzJust", 2)
        'SetData("TextAngle", 270)	
        '}
        '  else{	
        '  	SetData("TextVertJust", 16)
        ' 	SetData("TextHorzJust", 4)
        '}

        Dim nCO_ArcDiameter As Double = FNRound(WHBODDIA1.FN_CalcLength(xyCO_LargestTop, xyCO_LargestBott))
        If (ClosedCrotch) Then
            If (sCrotchStyle.Contains("Horizontal Fly") = 0) Then
                sCrotchText = "Horizontal Fly, "
                If (nAge >= 3 And nAge <= 6) Then
                    If (nCO_ArcDiameter <= 1.5) Then
                        sCrotchText = sCrotchText + "C-1"
                        nCrotchSize = 2
                    Else
                        sCrotchText = sCrotchText + "C-2"
                        nCrotchSize = 2.5
                    End If
                End If
                If (nAge >= 7 And nAge <= 11) Then
                    If (nCO_ArcDiameter <= 2) Then
                        sCrotchText = sCrotchText + "C-2"
                        nCrotchSize = 2.5
                    Else
                        sCrotchText = sCrotchText + "C-3"
                        nCrotchSize = 2.5
                    End If
                End If
                If (nAge >= 12 And nAge <= 14) Then
                    If (nCO_ArcDiameter <= 2.25) Then
                        sCrotchText = sCrotchText + "C-3"
                        nCrotchSize = 2.5
                    End If
                    If (nCO_ArcDiameter > 2.25 And nCO_ArcDiameter <= 2.5) Then
                        sCrotchText = sCrotchText + "C-4"
                        nCrotchSize = 2.5
                    End If
                    If (nCO_ArcDiameter > 2.5) Then
                        sCrotchText = sCrotchText + "C-5"
                        nCrotchSize = 3.125
                    End If
                End If
                If (nAge >= 15) Then
                    If (nCO_ArcDiameter <= 3.5) Then
                        sCrotchText = sCrotchText + "1"
                        nCrotchSize = 3.5
                    End If
                    If (nCO_ArcDiameter > 3.5 And nCO_ArcDiameter <= 4.25) Then
                        sCrotchText = sCrotchText + "2"
                        nCrotchSize = 3.75
                    End If
                    If (nCO_ArcDiameter > 4.25 And nCO_ArcDiameter <= 5) Then
                        sCrotchText = sCrotchText + "3"
                        nCrotchSize = 4
                    End If
                    If (nCO_ArcDiameter > 5) Then
                        sCrotchText = sCrotchText + "Oversize"
                        nCrotchSize = 4
                    End If
                End If
                If (sCrotchStyle.Contains("Diagonal Fly") = 0) Then
                    sCrotchText = "Diagonal Fly, "
                    If (nAge < 10) Then
                        sCrotchText = sCrotchText + "Small"
                        nCrotchSize = 1.75
                    End If
                    If (nAge >= 10 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Medium"
                        nCrotchSize = 2.75
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Large"
                        nCrotchSize = 3.75
                    End If
                End If
            End If
            If (sCrotchStyle.Contains("Male Mesh Gusset") Or sCrotchStyle.Contains("Mesh Gusset") Or sCrotchStyle.Contains("Female Mesh Gusset")) Then
                sCrotchText = "Mesh Gusset, "
                If (sCrotchStyle.Contains("Female Mesh Gusset") = 0) Then
                    Female = True
                    Male = False
                End If
                If (sCrotchStyle.Contains("Male Mesh Gusset") = 0) Then
                    Female = False
                    Male = True
                End If
                If (Male) Then
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                        nCrotchSize = 2.125
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Boy's"
                        nCrotchSize = 3.875
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Male"
                    End If
                End If
            Else
                If (nAge <= 3) Then
                    sCrotchText = sCrotchText + Format("length", 1.75)
                End If
                If (nAge > 3 And nAge <= 14) Then
                    If (nCO_ArcDiameter <= 1.75) Then
                        sCrotchText = sCrotchText + Format("length", 1)
                    End If
                    If (nCO_ArcDiameter > 1.75 And nCO_ArcDiameter <= 3.125) Then
                        sCrotchText = sCrotchText + Format("length", 1.25)
                    End If
                    If (nCO_ArcDiameter > 3.125) Then
                        sCrotchText = sCrotchText + "Regular"
                    End If
                End If
                If (nAge >= 15) Then
                    If (nMaxCir < 45) Then
                        sCrotchText = sCrotchText + "Regular"
                    Else
                        sCrotchText = sCrotchText + "Oversize"
                    End If
                End If
            End If
            If (String.Compare(sCrotchStyle, "Gusset") = 0) Then
                sCrotchText = sCrotchStyle + ", "
                If (Male) Then
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                        nCrotchSize = 2.125
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Boy's"
                        nCrotchSize = 3.875
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Male"
                    End If
                Else
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        If (nCO_ArcDiameter <= 1.75) Then
                            sCrotchText = sCrotchText + Format("length", 1)
                        End If
                        If (nCO_ArcDiameter > 1.75 And nCO_ArcDiameter <= 3.125) Then
                            sCrotchText = sCrotchText + Format("length", 1.25)
                        End If
                        If (nCO_ArcDiameter > 3.125) Then
                            sCrotchText = sCrotchText + "Regular"
                        End If
                    End If
                    If (nAge >= 15) Then
                        If (nMaxCir < 45) Then
                            sCrotchText = sCrotchText + "Regular"
                        Else
                            sCrotchText = sCrotchText + "Oversize"
                        End If
                    End If
                End If
            End If
            ''hEnt = AddEntity("marker","closed arrow", xyCO_CenterArrow, 0.5 ,0.125, 180 ) ;
            ''AddEntity("text", sCrotchText,xyCO_CenterArrow.X - 0.75,xyCO_CenterArrow.Y)
            PR_DrawWaistClosedArrow(xyCO_CenterArrow, 180)
            Dim xyText As WHBODDIA1.XY
            PR_MakeXY(xyText, xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
            PR_DrawText(sCrotchText, xyText, 0.125, 0, 6)
        Else
            'Open Crotch
            If (nOpenOff = 0.375 Or (nAge <= 10 And PantyLeg)) Then
                ''AddEntity("text", "1/2\" & "Elastic", xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
                Dim xyText As WHBODDIA1.XY
                PR_MakeXY(xyText, xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
                sCrotchText = "1/2" & Chr(34) & "Elastic"
                PR_DrawText(sCrotchText, xyText, 0.125, 0, 6)
            End If
        End If

    End Sub
    Private Sub ScanLine(ByRef sLine As String, ByRef nStyle As Double, ByRef sLeftStyle As Double,
                         ByRef nRightStyle As Double, ByRef nCir As Double)
        nStyle = FN_GetNumber(sLine, 1)
        sLeftStyle = FN_GetNumber(sLine, 2)
        nRightStyle = FN_GetNumber(sLine, 3)
        nCir = FN_GetNumber(sLine, 4)
    End Sub
    Sub PR_DrawWaistClosedArrow(ByRef xyPoint As WHBODDIA1.XY, ByRef nAngle As Object)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim xyStart As WHBODDIA1.XY
        ''PR_MakeXY(xyStart, xyPoint.X, xyPoint.Y + 0.125)
        xyStart = xyPoint
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Dim acPoly As Polyline = New Polyline()
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.Y), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.Y + 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.Y - 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.Y), 0, 0, 0)
            acPoly.TransformBy(Matrix3d.Rotation((nAngle * (ARMDIA1.PI / 180)), Vector3d.ZAxis, New Point3d(xyStart.X, xyStart.Y, 0)))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyStart.X, xyStart.Y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            Dim idPolyline As ObjectId = New ObjectId
            idPolyline = acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
            ''Create Hatch entity
            Dim ObjIds As ObjectIdCollection = New ObjectIdCollection
            ObjIds.Add(idPolyline)
            Dim oHatch As Hatch = New Hatch()
            Dim normal As Vector3d = New Vector3d(0.0, 0.0, 1.0)
            oHatch.Normal = normal
            oHatch.Elevation = 0.0
            oHatch.PatternScale = 2.0
            oHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
            oHatch.ColorIndex = 1
            acBlkTblRec.AppendEntity(oHatch)
            acTrans.AddNewlyCreatedDBObject(oHatch, True)
            oHatch.Associative = True
            oHatch.AppendLoop(CInt(HatchLoopTypes.Default), ObjIds)
            oHatch.Color = acPoly.Color
            oHatch.EvaluateHatch(True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawText(ByRef sText As String, ByRef xyInsert As WHBODDIA1.XY, ByRef nHeight As Object, ByRef nAngle As Double, ByRef nTextmode As Double)
        Dim nWidth As Object
        nWidth = nHeight * g_nWebCurrTextAspect
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
                acText.Position = New Point3d(xyInsert.X, xyInsert.Y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.Justify = nTextmode
                acText.AlignmentPoint = New Point3d(xyInsert.X, xyInsert.Y, 0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X, xyInsert.Y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_AddDBValueToLast(ByRef sDBName As String, ByRef sDBValue As String)
        If (idLastMarker.IsValid = False) Then
            Exit Sub
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As DBObject = acTrans.GetObject(idLastMarker, OpenMode.ForWrite)
            Dim acRegAppTbl As RegAppTable
            acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
            Dim acRegAppTblRec As RegAppTableRecord
            If acRegAppTbl.Has(sDBName) = False Then
                acRegAppTblRec = New RegAppTableRecord
                acRegAppTblRec.Name = sDBName
                acRegAppTbl.UpgradeOpen()
                acRegAppTbl.Add(acRegAppTblRec)
                acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
            End If
            Using rb As New ResultBuffer
                rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, sDBName))
                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, sDBValue))
                acEnt.XData = rb
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_GetBlkPosition(ByRef xyInsertion As WHBODDIA1.XY)
        If (idLastMarker.IsValid = False) Then
            Exit Sub
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkXMarker As BlockReference = acTrans.GetObject(idLastMarker, OpenMode.ForRead)
            xyInsertion.X = acBlkXMarker.Position.X
            xyInsertion.Y = acBlkXMarker.Position.Y
        End Using
    End Sub
End Class
Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
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

    Public MainForm As bodysuit
    Public g_nChestCir As Double
    Public g_nWaistCir As Double
    Public g_nShldToFold As Double
    Public g_nShldToWaist As Double
    Public g_nButtocksLength As Double
    Public g_nUnderBreastCir As Double
    Public g_nNeckCir As Double
    Public g_nShldToBreast As Double
    Public g_nCutOutToBackMinimum As Double
    Public g_nButtocksCir As Double
    Public g_nButtocksCirGiven As Double

    Public g_nChestCirGiven As Double
    Public g_nChestCirRed As Double
    Public g_nWaistCirGiven As Double
    Public g_nWaistCirRed As Double
    Public g_nShldToFoldGiven As Double
    Public g_nShldToFoldRed As Double
    Public g_nShldToWaistGiven As Double
    Public g_nShldToWaistRed As Double
    Public g_nShldToBreastGiven As Double
    Public g_nShldToBreastRed As Double
    Public g_nUnderBreastCirGiven As Double
    Public g_nUnderBreastCirRed As Double
    Public g_nNeckCirGiven As Double
    Public g_nNeckCirRed As Double

    Public g_nGroinHeight As Double
    Public g_nHalfGroinHeight As Double
    Public g_nButtRadius As Double
    Public g_nButtCirCalc As Double
    Public g_nButtFrontSeam As Double
    Public g_nButtBackSeam As Double
    Public g_n85FoldHeight As Double
    Public g_nButtocksCirRed As Double
    Public g_nThighCirRed As Double

    Public g_nLT_GroinHeight As Double
    Public g_nLT_HalfGroinHeight As Double
    Public g_nLT_ButtRadius As Double
    Public g_nButtCutOut As Double
    Public g_nWaistFrontSeam As Double
    Public g_nChestFrontSeam As Double
    Public g_nCutOut As Double
    Public g_nWaistBackSeam As Double
    Public g_nChestBackSeam As Double

    Public g_sLegStyle As String
    Public g_sFileNo As String
    Public g_sPatient As String
    Public g_sDiagnosis As String
    Public g_sWorkOrder As String
    Public g_sSex As String
    Public g_sClosure As String
    Public g_sFabric As String
    Public g_sLeftLeg As String
    Public g_sRightLeg As String
    Public g_nAge As Short
    Public g_nAdult As Short
    Public g_nNippleCirGiven As Double
    Public g_nFrontNeckSize As Double
    Public g_nBackNeckSize As Double

    Public Structure BiArc
        Dim xyStart As BODYSUIT1.XY
        Dim xyTangent As BODYSUIT1.XY
        Dim xyEnd As BODYSUIT1.XY
        Dim xyR1 As BODYSUIT1.XY
        Dim xyR2 As BODYSUIT1.XY
        Dim nR1 As Double
        Dim nR2 As Double
    End Structure
    Public Structure curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim y() As Double
        Public Sub Initialize()
            ReDim X(100)
            ReDim y(100)
        End Sub
    End Structure
    Public g_LeftLegProfile As curve
    Public g_RightLegProfile As curve

    Dim strLeftUidLeg As String
    Dim strRightUidLeg As String
    Dim g_strSide As String
    Dim strLeftLengths As String
    Dim strLeftLegPleats As String
    Dim strRightLengths As String
    Dim strRightLegPleats As String

    Const SEAM As Double = 0.25
    Public Const ONE_and_THREE_QUARTER_GUSSET As Double = 2.58
    Public Const BOYS_GUSSET As Double = 2.8
    Public Const TOP_ As Short = 8
    Public Const LEFT_ As Short = 1

    Dim m_nSpace(30) As Double
    Dim m_n20Len(30) As Double
    Dim m_nReduction(30) As Double
    'Left and right caculated values
    Dim m_nLeftOriginalTapeValues(30) As Double
    Dim m_nRightOriginalTapeValues(30) As Double
    Dim m_nLeftKneePosition As Double
    Dim m_nRightKneePosition As Double

    Dim m_iLeftFirstTape As Short
    Dim m_iLeftLastTape As Short
    Dim m_iRightFirstTape As Short
    Dim m_iRightLastTape As Short
    Dim xyInsertion As BODYSUIT1.XY
    Dim idLastCreated As ObjectId


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()

        If sCurrentValues <> BODYSUIT1.g_sChangeChecker Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "Bodysuit - Glove Dialogue")
            Select Case Response
                Case Public_Renamed.IDYES
                    ''--------PR_CreateMacro_Save("c:\jobst\draw.d")
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    PR_CreateMacro_Save(sDrawFile)
                    saveInfoToDWG()
                    Me.Close()
                Case Public_Renamed.IDNO
                    Me.Close()
                Case Public_Renamed.IDCANCEL
                    Exit Sub
            End Select
        Else
            Me.Close()
        End If
        BodyMain.BodyMainDlg.Close()
    End Sub
	
	Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click

        Dim sCurrentValues As String

        'Check if data has been modified
        sCurrentValues = FN_ValuesString()
        If sCurrentValues <> BODYSUIT1.g_sChangeChecker Then
            'Check all values are present
            ''------PR_GetValuesFromDialogue()
            If FN_ValidateAndCalculateData(True) Then 'else stay in dialog
                ''------------PR_CreateMacro_Save("c:\jobst\draw.d")
                Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                PR_CreateMacro_Save(sDrawFile)
                PR_DrawBodysuitBlock()
                saveInfoToDWG()

                BodyMain.BodyMainDlg.Show()
                Me.Show()
            End If
        Else
            'End '''J'''
        End If
    End Sub

    Private Sub cboBackNeck_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboBackNeck.SelectedIndexChanged
        Select Case VB.Left(cboBackNeck.Text, 1)
            Case "R", "S" 'Regular or Scoop
                txtBackNeck.Text = ""
                lblBackNeck.Text = ""
                txtBackNeck.Enabled = False
                labBackNeck.Enabled = False
            Case Else 'Measured Scoop
                If g_sBodyBackNeck <> "" Then
                    txtBackNeck.Text = BODYSUIT1.g_sFrontNeck
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
            sError = sError & "Male Crotch type specified for Female patient" + BODYSUIT1.NL
        End If

        'Check for female crotch and male patient
        iindex = 0
        iindex = InStr(1, sString, "Female", 0)
        If iindex <> 0 And txtSex.Text = "Male" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sError = sError & "Female Crotch type specified for Male patient" + BODYSUIT1.NL
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
                If BODYSUIT1.g_sFrontNeck <> "" Then
                    txtFrontNeck.Text = BODYSUIT1.g_sFrontNeck
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
    Private Sub cboLeftCup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If Me.Visible And (cboLeftCup.Text = "" Or cboLeftCup.Text = "None") Then
            txtLeftDisk.Text = ""
            BODYSUIT1.g_sLeftCup = ""
            BODYSUIT1.g_sLeftDisk = ""
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
        If sLegStyle = BODYSUIT1.g_sPreviousLegStyle Then Exit Sub

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
        BODYSUIT1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)

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
    Private Sub cboRightCup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If Me.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
            txtRightDisk.Text = ""
            BODYSUIT1.g_sRightCup = ""
            BODYSUIT1.g_sRightDisk = ""
        Else
            PR_EnableCalculateDiskButton()
        End If
    End Sub

    Private Sub cmdCalculateBraDisks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        PR_DoBraCupsAndDisks()
    End Sub

    Private Sub cmdTab_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTab.Click
        'Allows the user to use enter as a tab
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Function FN_CalcLength(ByRef xyStart As BODYSUIT1.XY, ByRef xyEnd As BODYSUIT1.XY) As Double
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
            sError = "No Bra disk is availabe for a Cup Size " & sCup & BODYSUIT1.NL & "For an under breast circumference of " & fnInchestoText(nUnderBreast)
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
                sError = "Warning" & BODYSUIT1.NL & "The Given disk is different in size to the disk that would have been calculated!"
                'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sError = sError & BODYSUIT1.NL & "Given Disk      : " & Str(iDisk)
                'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sError = sError & BODYSUIT1.NL & "Calculated Disk : " & Str(nSelctedDisk)
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
        '        BODYSUIT1.g_nUnitsFac = 1       => nDisplay in Inches
        '        BODYSUIT1.g_nUnitsFac = 10/25.5 => nDisplay in CMs
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
        '        Eg 7.8 is an invalid number if BODYSUIT1.g_nUnitsFac = 1
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
        If BODYSUIT1.g_nUnitsFac <> 1 Then
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
        Dim Char_Renamed As String
        Dim sEscapedString As String = ""

        FN_EscapeQuotesInString = ""

        For ii = 1 To Len(sAssignedString)
            Char_Renamed = Mid(sAssignedString, ii, 1)
            If Char_Renamed = BODYSUIT1.QQ Then
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
        nLen = fnDisplayToInches(Val(TextBox.Text))
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
        BODYSUIT1.fNum = FreeFile()
        FileOpen(BODYSUIT1.fNum, sDrafixFile, VB.OpenMode.Output)
        FN_Open = BODYSUIT1.fNum

        'Write header information etc. to the DRAFIX macro file
        '
        PrintLine(BODYSUIT1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(BODYSUIT1.fNum, "//Patient - " & sName & ", " & sFileNo & "")
        PrintLine(BODYSUIT1.fNum, "//by Visual Basic, BodySuit")
        PrintLine(BODYSUIT1.fNum, "//type - " & sType & "")
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
        nShldWidthGiven = fnDisplayToInches(Val(txtCir(3).Text))
        nShldToWaistGiven = fnDisplayToInches(Val(txtCir(4).Text))
        nChestCirGiven = fnDisplayToInches(Val(txtCir(5).Text))
        nWaistCirGiven = fnDisplayToInches(Val(txtCir(6).Text))
        nShldToFoldGiven = fnDisplayToInches(Val(txtCir(12).Text))
        nButtocksCirGiven = fnDisplayToInches(Val(txtCir(14).Text))
        nLeftThighCirGiven = fnDisplayToInches(Val(txtCir(15).Text))
        nRightThighCirGiven = fnDisplayToInches(Val(txtCir(16).Text))
        sCrotchStyle = cboCrotchStyle.Text

        'Reductions
        nShldToFoldRed = 0.95 '95%
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
            sError = sError & "Shoulder Width is less than 1/2""!" & BODYSUIT1.NL
        End If

        Dim sCircum(16) As Object
        sCircum(0) = "Left shoulder circ."
        sCircum(1) = "Right shoulder circ."
        sCircum(2) = "Neck circ."
        sCircum(3) = "Shoulder width."
        sCircum(4) = "Shoulder to waist."
        sCircum(5) = "Chest circ."
        sCircum(6) = "Waist circ."
        sCircum(9) = "Shoulder to under breast."
        sCircum(10) = "Circ. under breast."
        sCircum(11) = "Circ. over nipple."
        sCircum(12) = "Shoulder to Fold of Buttocks."
        sCircum(13) = "Shoulder to Large Part of Buttocks."
        sCircum(14) = "Circ. of Large Part of Buttocks."
        sCircum(15) = "Left Thigh Circ."
        sCircum(16) = "Right Thigh Circ."

        'FATAL Errors
        '~~~~~~~~~~~~
        'Body measurements (all must be present)
        For ii = 0 To 6
            If Val(txtCir(ii).Text) = 0 Then
                sError = sError & "Missing " & sCircum(ii) & BODYSUIT1.NL
            End If
        Next ii
        For ii = 12 To 16
            If Val(txtCir(ii).Text) = 0 Then
                sError = sError & "Missing " & sCircum(ii) & BODYSUIT1.NL
            End If
        Next ii

        'Bra Cups
        'Note:
        '    The Circumference over nipple is optional unless a cup has been
        '    specified
        '
        If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(9) & BODYSUIT1.NL
        End If
        If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(10) & BODYSUIT1.NL
        End If

        If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
            sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & BODYSUIT1.NL
        End If

        If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
            sError = sError & "Bra disks given! But Missing " & sCircum(9) & BODYSUIT1.NL
        End If

        'If cups or dimensions given then a disk must be present
        'NB
        '    cboXXXXCup.ListIndex = 6 = "None"
        '    cboXXXXCup.ListIndex = 7 = ""

        If cboLeftCup.SelectedIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Left BRA Cup!" & BODYSUIT1.NL
        End If
        If cboRightCup.SelectedIndex < 6 And Val(txtRightDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Right BRA Cup!" & BODYSUIT1.NL
        End If

        'Sex error
        If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.SelectedIndex < 0 Or cboRightCup.SelectedIndex < 0) Then
            sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & BODYSUIT1.NL
        End If

        'Neck at back and front
        Dim sChar As New VB6.FixedLengthString(1)
        sChar.Value = VB.Left(cboBackNeck.Text, 1)
        If sChar.Value = "M" And txtBackNeck.Text = "" Then
            sError = sError & "No dimension for Back Meck Measured Scoop!" & BODYSUIT1.NL
        End If

        sChar.Value = VB.Left(cboFrontNeck.Text, 1)
        If sChar.Value = "M" And txtFrontNeck.Text = "" Then
            sError = sError & "No dimension for Front Neck Measured Scoop!" & BODYSUIT1.NL
        End If

        '
        If cboLeftAxilla.Text = "" Then
            sError = sError & "Left Axilla not given!" & BODYSUIT1.NL
        End If

        If cboRightAxilla.Text = "" Then
            sError = sError & "Right Axilla not given!" & BODYSUIT1.NL
        End If

        If cboFrontNeck.Text = "" Then
            sError = sError & "Neck not given!" & BODYSUIT1.NL
        End If

        If cboBackNeck.Text = "" Then
            sError = sError & "Back neck not given!" & BODYSUIT1.NL
        End If

        If cboClosure.Text = "" Then
            sError = sError & "Closure not given!" & BODYSUIT1.NL
        End If

        If cboFabric.Text = "" Then
            sError = sError & "Fabric not given!" & BODYSUIT1.NL
        End If

        'Extra Values
        If cboLegStyle.Text = "" Then
            sError = sError & "Leg Style not given!" & BODYSUIT1.NL
        End If
        If cboCrotchStyle.Text = "" Then
            sError = sError & "Crotch Style not given!" & BODYSUIT1.NL
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
        nChestCir = fnRoundInches(nChestCirGiven * nChestCirRed) / 2 'half scale
        nWaistCir = fnRoundInches(nWaistCirGiven * nWaistCirRed) / 2 'half-scale
        nShldToFold = fnRoundInches(nShldToFoldGiven * nShldToFoldRed) + BODYSUIT1.BODY_INCH1_2
        nShldToWaist = fnRoundInches(nShldToWaistGiven * nShldToWaistRed) + BODYSUIT1.BODY_INCH1_2
        nButtocksLength = fnRoundInches((nShldToFold - nShldToWaist) / 3)

        'Test Measurements

        'NOTE:-
        '    There can be no more that 5% difference between the reductions
        '    at each circumference.

        'Use smallest thigh size.
        If (nLeftThighCirGiven < nRightThighCirGiven) Then
            nSmallestThighGiven = nLeftThighCirGiven
        Else
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
            nButtocksCir = fnRoundInches(nButtocksCirGiven * nButtocksCirRed) / 2 'half scale
            nGroinHeight = fnRoundInches(nSmallestThighGiven * nThighCirRed) / 2 'half scale
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
            sError = sError & "With given Buttock Reduction, distance from back of Cut-Out" + BODYSUIT1.NL
            sError = sError & "to largest part of buttocks is less than 3 inches!" + BODYSUIT1.NL
            sError = sError & "Buttocks Reduction increased to rectify this." + BODYSUIT1.NL
            'sError = sError + "ref:- " + BODYSUIT1.NL
        End If
        If nThighRedDecreased Then
            cboRed(4).Text = CStr(nThighCirRed)
            sError = sError & "With given Thigh Reduction, distance from back of Cut-Out" + BODYSUIT1.NL
            sError = sError & "to largest part of buttocks is less than 3 inches!" + BODYSUIT1.NL
            sError = sError & "Thigh Reduction decreased to rectify this." + BODYSUIT1.NL
            'sError = sError + "ref:- " + BODYSUIT1.NL
        End If
        'One last time with Reductions
        nButtocksCir = fnRoundInches(nButtocksCirGiven * nButtocksCirRed) / 2 'half scale
        nGroinHeight = fnRoundInches(nSmallestThighGiven * nThighCirRed) / 2 'half scale
        nHalfGroinHeight = (nGroinHeight / 2)
        nButtRadius = System.Math.Sqrt((nButtocksLength ^ 2) + (nHalfGroinHeight ^ 2))
        nButtCirCalc = nButtRadius + nHalfGroinHeight
        nButtBackSeam = (nButtocksCir - nButtCirCalc) * nButtBackSeamRatio

        If nButtBackSeam <= 0 Then
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + BODYSUIT1.NL
            sError = sError & "is still negative or zero!" + BODYSUIT1.NL
            sError = sError & "Even with the modified reductions." + BODYSUIT1.NL
            iFatalError = True
        ElseIf nButtBackSeam < 1.5 Then  'less than 3" full scale
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + BODYSUIT1.NL
            sError = sError & "is still less than 3 inches!" + BODYSUIT1.NL
            sError = sError & "Even with the modified reductions." + BODYSUIT1.NL
        End If

        'All other seams test
        nButtFrontSeam = (nButtocksCir - nButtCirCalc) * nButtFrontSeamRatio
        nButtCutOut = nButtCirCalc - (nButtBackSeam + nButtFrontSeam)
        If (nWaistCirGiven < nButtocksCirGiven) Then
            nWaistFrontSeam = System.Math.Abs(((nButtocksCirGiven - nWaistCirGiven) / 2) - nButtFrontSeam)
        Else
            nWaistFrontSeam = nButtFrontSeam + (Int(nWaistCirGiven - nButtocksCirGiven) * BODYSUIT1.BODY_INCH1_8)
        End If
        If (nChestCirGiven < nButtocksCirGiven) Then
            nChestFrontSeam = System.Math.Abs(((nButtocksCirGiven - nChestCirGiven) / 2) - nButtFrontSeam)
        Else
            nChestFrontSeam = nButtFrontSeam + (Int(nChestCirGiven - nButtocksCirGiven) * BODYSUIT1.BODY_INCH1_8)
        End If
        nWaistCutOut = (nButtCutOut * 0.87)
        nWaistCutOutModified = 0
        Do
            nWaistRem = nWaistCir - (nWaistCutOut + (nWaistFrontSeam * 2))
            nChestRem = nChestCir - (nWaistCutOut + nWaistFrontSeam + nChestFrontSeam)
            If nWaistRem < 1.5 Or nChestRem < 1.5 Then 'less than 3" @ full scale
                If nWaistRem <= 0 Then
                    sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
                    sError = sError & "Distance from back of Cut-Out to Waist" + BODYSUIT1.NL
                    sError = sError & "is negative or zero!" + BODYSUIT1.NL
                    If nWaistCutOutModified > 0 Then
                        sError = sError & "Even with the Waist CutOut decreased." + BODYSUIT1.NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                If nChestRem <= 0 Then
                    sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
                    sError = sError & "Distance from back of Cut-Out to Chest" + BODYSUIT1.NL
                    sError = sError & "is negative or zero!" + BODYSUIT1.NL
                    If nWaistCutOutModified > 0 Then
                        sError = sError & "Even with the Waist CutOut decreased." + BODYSUIT1.NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                'Decreae Waist CutOut
                nWaistCutOut = nWaistCutOut - (2 * (1.5 - nWaistRem))
                nWaistCutOutModified = nWaistCutOutModified + 1
                If nWaistCutOut <= 0 Then
                    sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
                    sError = sError & "Distance from back of Cut-Out to Waist or Chest" + BODYSUIT1.NL
                    sError = sError & "is less than 3 inches!" + BODYSUIT1.NL
                    Select Case nWaistCutOutModified
                        Case 1
                            sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " time" & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
                        Case Is > 1
                            sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " times" & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
                    End Select
                    sError = sError & "but is still negative or zero!" + BODYSUIT1.NL
                    iFatalError = True
                    Exit Do
                End If
            Else
                If nWaistCutOutModified > 0 Then
                    sError = sError & "Distance from back of Cut-Out to Waist or Chest" + BODYSUIT1.NL
                    sError = sError & "was less than 3 inches!" + BODYSUIT1.NL
                    Select Case nWaistCutOutModified
                        Case 1
                            sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " time." & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
                        Case Is > 1
                            sError = sError & "Waist CutOut had to be decreased " & Str(nWaistCutOutModified) & " times." & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
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
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "All Reductions should be within 5% of each other" + BODYSUIT1.NL
            '       iFatalError = True   'gg
        End If

        'Display Error message (if required) and return
        'There could be fatal errors found
        If Len(sError) > 0 Then
            If iFatalError = True Then
                MsgBox(sError, 64, "Warning - Problems with data")
                Fn_ValidateData = False
            Else
                sError = sError + BODYSUIT1.NL + "The above problems have been found in the data do you" + BODYSUIT1.NL
                sError = sError & "wish to continue ?"
                iError = MsgBox(sError, 52, "Severe Problems with data")
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
        '        BODYSUIT1.g_nUnitsFac = 1       => nDisplay in Inches
        '        BODYSUIT1.g_nUnitsFac = 10/25.5 => nDisplay in CMs
        'Returns:-
        '        Single, Inches rounded to the nearest eighth (0.125)
        '        -1,     on conversion error.
        '
        'Errors:-
        '        The returned value is usually +ve. Unless it can't
        '        be sucessfully converted to inches.
        '        Eg 7.8 is an invalid number if BODYSUIT1.g_nUnitsFac = 1
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
        If BODYSUIT1.g_nUnitsFac <> 1 Then
            fnDisplayToInches = fnRoundInches(nDisplay * BODYSUIT1.g_nUnitsFac) * iSign
            Exit Function
        End If

        'Imperial units
        iInt = Int(nDisplay)
        nDec = nDisplay - iInt

        'Check that conversion is possible (return -1 if not)
        If nDec > 0.8 Then
            fnDisplayToInches = -1
        Else
            fnDisplayToInches = fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
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
        'If txtUidMPD.Text = "" Then
        '    MsgBox("No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue")
        '    'End  '''J'''
        'End If

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
                BODYSUIT1.g_sLeftCup = VB6.GetItemString(cboLeftCup, ii)
                cboLeftCup.SelectedIndex = ii
            End If
        Next ii
        For ii = 0 To (cboRightCup.Items.Count - 1)
            If txtCombo(4).Text = VB6.GetItemString(cboRightCup, ii) Then
                BODYSUIT1.g_sRightCup = VB6.GetItemString(cboRightCup, ii)
                cboRightCup.SelectedIndex = ii
            End If
        Next ii

        BODYSUIT1.g_sLeftDisk = txtLeftDisk.Text
        BODYSUIT1.g_sRightDisk = txtRightDisk.Text

        If txtCombo(5).Text <> "" Then cboLeftAxilla.Text = txtCombo(5).Text Else cboLeftAxilla.SelectedIndex = 5
        If txtCombo(6).Text <> "" Then cboRightAxilla.Text = txtCombo(6).Text Else cboRightAxilla.SelectedIndex = 5
        If txtCombo(7).Text <> "" Then cboFrontNeck.Text = txtCombo(7).Text Else cboFrontNeck.SelectedIndex = 0
        If txtCombo(8).Text <> "" Then cboBackNeck.Text = txtCombo(8).Text Else cboBackNeck.SelectedIndex = 0
        If txtCombo(9).Text <> "" Then cboClosure.Text = txtCombo(9).Text Else cboClosure.SelectedIndex = 0
        cboFabric.Text = txtCombo(10).Text

        'Set up units
        If txtUnits.Text = "cm" Then
            BODYSUIT1.g_nUnitsFac = 10 / 25.4
        Else
            BODYSUIT1.g_nUnitsFac = 1
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
        g_nBodyChest = Val(txtCir(5).Text)
        g_nBodyUnderBreast = Val(txtCir(10).Text)
        g_nBodyNipple = Val(txtCir(11).Text)

        'Store values to use on change etc
        Dim sChar As New VB6.FixedLengthString(1)
        g_sBodyBackNeck = txtBackNeck.Text
        sChar.Value = VB.Left(cboBackNeck.Text, 1)
        If sChar.Value = "R" Or sChar.Value = "S" Or sChar.Value = "V" Then
            'Disable length box for Regular, Scoop, V neck
            txtBackNeck.Enabled = False
            labBackNeck.Enabled = False
        Else
            txtBackNeck.Enabled = True
            labBackNeck.Enabled = True
        End If

        BODYSUIT1.g_sFrontNeck = txtFrontNeck.Text
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
        cboCrotchStyle.SelectedIndex = 4

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
        BODYSUIT1.g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
        For ii = 0 To (cboLegStyle.Items.Count - 1)
            If VB6.GetItemString(cboLegStyle, ii) = fnGetString(txtLegStyle.Text, 1) Then
                cboLegStyle.SelectedIndex = ii
            End If
        Next ii

        'Set this now to avoid problems with cboLegStyle_Click
        If cboLegStyle.SelectedIndex = -1 Then
            BODYSUIT1.g_sPreviousLegStyle = "None"
        Else
            BODYSUIT1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
        End If

        If fnGetString(txtLegStyle.Text, 2) <> "" Then
            If fnGetString(txtLegStyle.Text, 2) = "Panty" Then
                optLeftLeg(0).Checked = True
            Else
                optLeftLeg(1).Checked = True 'Brief
            End If
        End If
        If fnGetString(txtLegStyle.Text, 3) <> "" Then
            If fnGetString(txtLegStyle.Text, 3) = "Panty" Then
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
        BODYSUIT1.g_sChangeChecker = FN_ValuesString()

    End Sub

    Private Sub bodysuit_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim ii As Short

        Hide()

        ''MainForm = Me // Commented - 11 - 11 - 2018
        MainForm = Me
        ''g_strSide = VB.Command()
        g_strSide = "Left"
        m_nLeftKneePosition = 0
        m_nRightKneePosition = 0
        idLastCreated = New ObjectId

        'Start a timer
        'The Timer is disabled in LinkClose
        'If after 6 seconds the timer event will "End" the programme
        'This ensures that the dialogue dies in event of a failure
        'on the drafix macro side
        Timer1.Interval = 6000 'Approx 6 Seconds
        Timer1.Enabled = True

        'Maintain while loading DDE data
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
        'Reset in Form_LinkClose

        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.


        'Initialize globals
        BODYSUIT1.QQ = Chr(34) 'Double quotes (")
        BODYSUIT1.NL = Chr(13) 'New Line
        BODYSUIT1.CC = Chr(44) 'The comma (,)
        BODYSUIT1.QCQ = BODYSUIT1.QQ & BODYSUIT1.CC & BODYSUIT1.QQ
        BODYSUIT1.QC = BODYSUIT1.QQ & BODYSUIT1.CC
        BODYSUIT1.CQ = BODYSUIT1.CC & BODYSUIT1.QQ

        BODYSUIT1.g_nUnitsFac = 1
        BODYSUIT1.g_sPathJOBST = fnPathJOBST()

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

        cboLeftAxilla.Items.Add("Regular 2" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Regular 1½" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Regular 2½" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Open")
        cboLeftAxilla.Items.Add("Mesh")
        cboLeftAxilla.Items.Add("Lining")
        cboLeftAxilla.Items.Add("Sleeveless")

        cboRightAxilla.Items.Add("Regular 2" & BODYSUIT1.QQ)
        cboRightAxilla.Items.Add("Regular 1½" & BODYSUIT1.QQ)
        cboRightAxilla.Items.Add("Regular 2½" & BODYSUIT1.QQ)
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

        ''------------LGLEGDIA1.PR_GetComboListFromFile(cboFabric, BODYSUIT1.g_sPathJOBST & "\FABRIC.DAT")
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

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
        txtWorkOrder.Text = workOrder
        Form_LinkClose()
        readDWGInfo()
        'g_nAxillaBackNeckRad = 0
        'g_nAxillaFrontNeckRad = 0
        'g_nABNRadRight = 0
        'g_nAFNRadRight = 0
        Button1.Enabled = False
        PR_EnableDrawBody()
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
        BODYSUIT1.fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

        'If this is a new drawing of a body then Define the DATA Base
        'fields for the BodySuit and insert the BODYSUIT symbol
        PR_PutLine("HANDLE hMPD, hBody;")

        PR_UpdateDB()

        FileClose(BODYSUIT1.fNum)

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
        '    BODYSUIT1.g_nUnitsFac As Double
        '    g_nBodyUnderBreast As Double
        '    g_nBodyChest As Double
        '    g_nBodyNipple As Double
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

        If sLeftCup <> BODYSUIT1.g_sLeftCup Then
            Changed = True
            BODYSUIT1.g_sLeftCup = sLeftCup
        End If

        If sRightCup <> BODYSUIT1.g_sRightCup Then
            Changed = True
            BODYSUIT1.g_sRightCup = sRightCup
        End If

        If Val(txtCir(10).Text) <> g_nBodyUnderBreast Then
            Changed = True
            g_nBodyUnderBreast = Val(txtCir(10).Text)
        End If

        If Val(txtCir(11).Text) <> g_nBodyNipple Then
            Changed = True
            g_nBodyNipple = Val(txtCir(11).Text)
            'Force a recalculation of cup
            If sLeftCup <> "None" Then sLeftCup = ""
            If sRightCup <> "None" Then sRightCup = ""
        End If

        If Val(txtCir(5).Text) <> g_nBodyChest Then
            g_nBodyChest = Val(txtCir(5).Text)
            'Force a recalculation of cup if a nipple circum. is given
            'Changed chest is only significant if a nipple line measurement
            'is given from which the Cup size is calculated
            If g_nBodyNipple <> 0 Then
                Changed = True
                If sLeftCup <> "None" Then sLeftCup = ""
                If sRightCup <> "None" Then sRightCup = ""
            End If
        End If

        'Force a calculation if either disk is empty
        If txtLeftDisk.Text = "" Or txtRightDisk.Text = "" Then Changed = True

        'Force a check if disk has been changed
        If txtLeftDisk.Text <> BODYSUIT1.g_sLeftDisk Or txtRightDisk.Text <> BODYSUIT1.g_sRightDisk Then
            Changed = True
            iRightDisk = Val(txtRightDisk.Text)
            iLeftDisk = Val(txtLeftDisk.Text)
        End If

        'If no modifications then exit
        If Changed <> True Then Exit Sub


        'Recalculate if Changed
        'Convert to inches
        nNipple = fnDisplayToInches(g_nBodyNipple)
        nUnderBreast = fnDisplayToInches(g_nBodyUnderBreast)
        nChest = fnDisplayToInches(g_nBodyChest)

        'Right cup
        If sRightCup <> "" Or nNipple <> 0 Then
            iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sRightCup, iRightDisk, "Right")
            If iError Then
                For ii = 0 To (cboRightCup.Items.Count - 1)
                    If VB6.GetItemString(cboRightCup, ii) = sRightCup Then cboRightCup.SelectedIndex = ii
                Next ii
                If iRightDisk = 0 Then txtRightDisk.Text = "" Else txtRightDisk.Text = CStr(iRightDisk)
                BODYSUIT1.g_sRightDisk = Trim(Str(iRightDisk))
                BODYSUIT1.g_sRightCup = sRightCup
            Else
                txtRightDisk.Text = ""
                cboRightCup.SelectedIndex = -1
                BODYSUIT1.g_sRightDisk = ""
                BODYSUIT1.g_sRightCup = ""
            End If
        End If

        If sLeftCup <> "" Or nNipple <> 0 Then
            iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sLeftCup, iLeftDisk, "Left")
            If iError Then
                For ii = 0 To (cboLeftCup.Items.Count - 1)
                    If VB6.GetItemString(cboLeftCup, ii) = sLeftCup Then cboLeftCup.SelectedIndex = ii
                Next ii
                If iLeftDisk <= 0 Then txtLeftDisk.Text = "" Else txtLeftDisk.Text = CStr(iLeftDisk)
                BODYSUIT1.g_sLeftDisk = Trim(Str(iLeftDisk))
                BODYSUIT1.g_sLeftCup = sLeftCup
            Else
                txtLeftDisk.Text = ""
                cboLeftCup.SelectedIndex = -1
                BODYSUIT1.g_sLeftDisk = ""
                BODYSUIT1.g_sLeftCup = ""
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

    Private Sub PR_GetComboListFromFile(ByRef Combo_Name As ComboBox, ByRef sFileName As String)
        'General procedure to create the list section of
        'a combo box reading the data from a file

        Dim sLine As String
        Dim fFileNum As Short

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48, "BodySuit Dialogue")
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
        Do While Not EOF(fFileNum)
            sLine = LineInput(fFileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Combo_Name.Items.Add(sLine)
        Loop
        FileClose(fFileNum)

    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BODYSUIT1.fNum, sLine)

    End Sub

    Private Sub PR_PutNumberAssign(ByRef sVariableName As String, ByRef nAssignedNumber As Object)
        'Procedure to put a number assignment
        'Adds a semi-colon
        '    fNum is global variable

        'UPGRADE_WARNING: Couldn't resolve default property of object nAssignedNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BODYSUIT1.fNum, sVariableName & "=" & Str(nAssignedNumber) & ";")

    End Sub

    Private Sub PR_PutStringAssign(ByRef sVariableName As String, ByRef sAssignedString As Object)
        'Procedure to put a string assignment
        'Encloses String in quotes and adds a semi-colon
        '    fNum is global variable

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BODYSUIT1.fNum, sVariableName & "=" & BODYSUIT1.QQ & sAssignedString & BODYSUIT1.QQ & ";")

    End Sub

    Private Sub PR_Select_Text(ByRef Text_Box_Name As TextBox)
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
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
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                     OpenMode.ForRead)

            ' Dim sLayerName As String = "Center"

            If acLyrTbl.Has(sNewLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sNewLayer)

                '' Save the changes
                acTrans.Commit()
            End If

            '' Dispose of the transaction
        End Using

        If BODYSUIT1.g_sCurrentLayer = sNewLayer Then Exit Sub
        BODYSUIT1.g_sCurrentLayer = sNewLayer

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BODYSUIT1.fNum, "hLayer = Table(" & BODYSUIT1.QQ & "find" & BODYSUIT1.QCQ & "layer" & BODYSUIT1.QCQ & sNewLayer & BODYSUIT1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(BODYSUIT1.fNum, "if ( hLayer > %zero && hLayer != 32768)" & "Execute (" & BODYSUIT1.QQ & "menu" & BODYSUIT1.QCQ & "SetLayer" & BODYSUIT1.QC & "hLayer);")

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
            PR_PutLine("@" & BODYSUIT1.g_sPathJOBST & "\BODY\BFIELDS.D;")

            'Find "mainpatientdetails" and get position
            PR_PutLine("XY     xyMPD_Origin, xyMPD_Scale ;")
            PR_PutLine("STRING sMPD_Name;")
            PR_PutLine("ANGLE  aMPD_Angle;")

            PR_PutLine("hMPD = UID (" & BODYSUIT1.QQ & "find" & BODYSUIT1.QC & Val(txtUidMPD.Text) & ");")
            PR_PutLine("if (hMPD)")
            PR_PutLine("  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel," & BODYSUIT1.QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & BODYSUIT1.QQ & ");")

            'Insert bodybox
            PR_PutLine("if ( Symbol(" & BODYSUIT1.QQ & "find" & BODYSUIT1.QCQ & sSymbol & BODYSUIT1.QQ & ")){")
            PR_PutLine("  Execute (" & BODYSUIT1.QQ & "menu" & BODYSUIT1.QCQ & "SetLayer" & BODYSUIT1.QC & "Table(" & BODYSUIT1.QQ & "find" & BODYSUIT1.QCQ & "layer" & BODYSUIT1.QCQ & "Data" & BODYSUIT1.QQ & "));")

            PR_PutLine("  hBody = AddEntity(" & BODYSUIT1.QQ & "symbol" & BODYSUIT1.QCQ & sSymbol & BODYSUIT1.QC & "xyMPD_Origin);")
            PR_PutLine("  }")
            PR_PutLine("else")
            PR_PutLine("  Exit(%cancel, " & BODYSUIT1.QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & BODYSUIT1.QQ & ");")
        Else
            'Use existing symbol
            PR_PutLine("hBody = UID (" & BODYSUIT1.QQ & "find" & BODYSUIT1.QC & Val(txtUidVB.Text) & ");")
            PR_PutLine("if (!hBody) Exit(%cancel," & BODYSUIT1.QQ & "Can't find >" & sSymbol & "< symbol to update!" & BODYSUIT1.QQ & ");")

        End If

        'Update the BODY Box symbol
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LtSCir" & BODYSUIT1.QCQ & txtCir(0).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "RtSCir" & BODYSUIT1.QCQ & txtCir(1).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "NeckCir" & BODYSUIT1.QCQ & txtCir(2).Text & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "SWidth" & BODYSUIT1.QCQ & txtCir(3).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "S_Waist" & BODYSUIT1.QCQ & txtCir(4).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "ChestCir" & BODYSUIT1.QCQ & txtCir(5).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "WaistCir" & BODYSUIT1.QCQ & txtCir(6).Text & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "S_Breast" & BODYSUIT1.QCQ & txtCir(9).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BreastCir" & BODYSUIT1.QCQ & txtCir(10).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "NippleCir" & BODYSUIT1.QCQ & txtCir(11).Text & BODYSUIT1.QQ & ");")

        'Extra Values
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "SFButt" & BODYSUIT1.QCQ & txtCir(12).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "SLgButt" & BODYSUIT1.QCQ & txtCir(13).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LgButtCir" & BODYSUIT1.QCQ & txtCir(14).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LtThCir" & BODYSUIT1.QCQ & txtCir(15).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "RtThCir" & BODYSUIT1.QCQ & txtCir(16).Text & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BraLtCup" & BODYSUIT1.QCQ & cboLeftCup.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BraRtCup" & BODYSUIT1.QCQ & cboRightCup.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BraLtDisk" & BODYSUIT1.QCQ & txtLeftDisk.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BraRtDisk" & BODYSUIT1.QCQ & txtRightDisk.Text & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LtAxillaType" & BODYSUIT1.QCQ & FN_EscapeQuotesInString(cboLeftAxilla.Text) & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "RtAxillaType" & BODYSUIT1.QCQ & FN_EscapeQuotesInString(cboRightAxilla.Text) & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "NeckType" & BODYSUIT1.QCQ & cboFrontNeck.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "NeckDimension" & BODYSUIT1.QCQ & txtFrontNeck.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BackNeckType" & BODYSUIT1.QCQ & cboBackNeck.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BackNeckDim" & BODYSUIT1.QCQ & txtBackNeck.Text & BODYSUIT1.QQ & ");")

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "Closure" & BODYSUIT1.QCQ & cboClosure.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "Fabric" & BODYSUIT1.QCQ & cboFabric.Text & BODYSUIT1.QQ & ");")

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

        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LegStyle" & BODYSUIT1.QCQ & txtLegStyle.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "CrotchStyle" & BODYSUIT1.QCQ & cboCrotchStyle.Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "WaistCirUserFac" & BODYSUIT1.QCQ & cboRed(0).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "ChestCirUserFac" & BODYSUIT1.QCQ & cboRed(1).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "BreastCirUserFac" & BODYSUIT1.QCQ & cboRed(2).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "LgButtCirUserFac" & BODYSUIT1.QCQ & cboRed(3).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "ThCirUserFac" & BODYSUIT1.QCQ & cboRed(4).Text & BODYSUIT1.QQ & ");")
        PR_PutLine("SetDBData( hBody" & BODYSUIT1.CQ & "fileno" & BODYSUIT1.QCQ & txtFileNo.Text & BODYSUIT1.QQ & ");")

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

    Private Sub Select_Text_In_Box(ByRef Text_Box_Name As TextBox)
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'It is assumed that the link open from Drafix has failed
        'Therefore we "End" here
        'End    '''J'''
    End Sub

    Private Sub txtBackNeck_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBackNeck.Enter
        PR_Select_Text(txtBackNeck)
    End Sub

    Private Sub txtBackNeck_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBackNeck.Leave
        Dim nLen As Double
        nLen = FN_InchesValue(txtBackNeck)
        If nLen >= 0 Then
            lblBackNeck.Text = fnInchestoText(nLen)
            g_sBodyBackNeck = txtBackNeck.Text
        Else
            lblBackNeck.Text = ""
        End If
    End Sub

    Private Sub txtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Enter
        Dim Index As Short = txtCir.GetIndex(eventSender)
        PR_Select_Text(txtCir(Index))
    End Sub

    Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
        Dim Index As Short = txtCir.GetIndex(eventSender)
        Dim nLen As Double
        nLen = FN_InchesValue(txtCir(Index))
        If nLen > 0 Then
            lblCir(Index).Text = fnInchestoText(nLen)
        Else
            lblCir(Index).Text = ""
        End If

        If Index > 10 Or Index = 5 Then PR_EnableCalculateDiskButton()


    End Sub

    Private Sub txtFrontNeck_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontNeck.Enter
        PR_Select_Text(txtFrontNeck)
    End Sub

    Private Sub txtFrontNeck_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontNeck.Leave
        Dim nLen As Double
        nLen = FN_InchesValue(txtFrontNeck)
        If nLen >= 0 Then
            lblFrontNeck.Text = fnInchestoText(nLen)
            BODYSUIT1.g_sFrontNeck = txtFrontNeck.Text 'Keep this for restore
        Else
            lblFrontNeck.Text = ""
        End If
    End Sub

    Private Sub txtLeftDisk_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        PR_Select_Text(txtLeftDisk)
    End Sub

    Private Sub txtLeftDisk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
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

    Private Sub txtRightDisk_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        PR_Select_Text(txtRightDisk)
    End Sub

    Private Sub txtRightDisk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
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
        rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))

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
        rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))

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

        Label_Name.Text = fnInchestoText(rTextBoxValue)

    End Sub
    Function fnRoundInches(ByVal nNumber As Double) As Double
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
        nDec = round(nDec)

        'Return value
        fnRoundInches = (iInt + (nDec * nPrecision)) * iSign

    End Function
    Sub PR_MakeXY(ByRef xyReturn As BODYSUIT1.XY, ByRef x As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = x
        xyReturn.Y = y
    End Sub
    Private Sub PR_DrawBodysuitBlock()
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
        xyEnd(1) = 0.5
        xyStart(2) = 0
        xyEnd(2) = 0.6875
        xyStart(3) = 1.5
        xyEnd(3) = 0.6875
        xyStart(4) = 1.5
        xyEnd(4) = 0.5
        xyStart(5) = 0
        xyEnd(5) = 0.5

        Dim xyText As BODYSUIT1.XY
        PR_MakeXY(xyText, 0.71875, 0.5)
        Dim strTag(39), strTextString(39) As String
        strTag(1) = "LtSCir"
        strTag(2) = "RtSCir"
        strTag(3) = "NeckCir"
        strTag(4) = "SWidth"
        strTag(5) = "S_Waist"
        strTag(6) = "ChestCir"
        strTag(7) = "WaistCir"
        strTag(8) = "S_Breast"
        strTag(9) = "BreastCir"
        strTag(10) = "NippleCir"
        strTag(11) = "SFButt"
        strTag(12) = "SLgButt"
        strTag(13) = "LgButtCir"
        strTag(14) = "LtThCir"
        strTag(15) = "RtThCir"
        strTag(16) = "BraLtCup"
        strTag(17) = "BraRtCup"
        strTag(18) = "BraLtDisk"
        strTag(19) = "BraRtDisk"
        strTag(20) = "LtAxillaType"
        strTag(21) = "RtAxillaType"
        strTag(22) = "NeckType"
        strTag(23) = "NeckDimension"
        strTag(24) = "BackNeckType"
        strTag(25) = "BackNeckDim"
        strTag(26) = "Closure"
        strTag(27) = "Fabric"
        strTag(28) = "LegStyle"
        strTag(29) = "CrotchStyle"
        strTag(30) = "WaistCirUserFac"
        strTag(31) = "ChestCirUserFac"
        strTag(32) = "BreastCirUserFac"
        strTag(33) = "LgButtCirUserFac"
        strTag(34) = "ThCirUserFac"
        strTag(35) = "fileno"
        strTag(36) = "AxillaBackNeckRad"
        strTag(37) = "AxillaFrontNeckRad"
        strTag(38) = "ABNRadRight"
        strTag(39) = "AFNRadRight"
        Dim n As Double
        For n = 1 To 39
            strTextString(n) = ""
        Next n

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRefPatient As BlockReference = acTrans.GetObject(blkId, OpenMode.ForRead)
            Dim ptPosition As Point3d = blkRefPatient.Position
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("SUITBODY") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "SUITBODY"
                Dim acPoly As Polyline = New Polyline()
                Dim ii As Double
                For ii = 1 To 5
                    acPoly.AddVertexAt(ii - 1, New Point2d(xyStart(ii), xyEnd(ii)), 0, 0, 0)
                Next ii
                blkTblRecCross.AppendEntity(acPoly)

                Dim acText As DBText = New DBText()
                acText.Position = New Point3d(xyText.X, xyText.Y, 0)
                acText.Height = 0.1
                acText.TextString = "SUIT-BODY"
                acText.Rotation = 0
                acText.Justify = AttachmentPoint.BottomCenter
                acText.AlignmentPoint = New Point3d(xyText.X, xyText.Y, 0)
                blkTblRecCross.AppendEntity(acText)

                Dim acAttDef As New AttributeDefinition
                For ii = 1 To 39
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
                blkRecId = acBlkTbl("SUITBODY")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(ptPosition, blkRecId)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(ptPosition.X, ptPosition.Y, 0)))
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
                            If acAttRef.Tag.ToUpper().Equals("LtSCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(0).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("RtSCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(1).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("NeckCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(2).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("SWidth", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(3).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("S_Waist", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(4).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ChestCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(5).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("WaistCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(6).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("S_Breast", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(9).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BreastCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(10).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("NippleCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(11).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("SFButt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(12).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("SLgButt", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(13).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LgButtCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(14).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LtThCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(15).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("RtThCir", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtCir(16).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BraLtCup", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboLeftCup.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BraRtCup", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRightCup.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BraLtDisk", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLeftDisk.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BraRtDisk", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtRightDisk.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LtAxillaType", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboLeftAxilla.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("RtAxillaType", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRightAxilla.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("NeckType", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboFrontNeck.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("NeckDimension", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFrontNeck.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BackNeckType", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboBackNeck.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BackNeckDim", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtBackNeck.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("Closure", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboClosure.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("Fabric", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboFabric.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LegStyle", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtLegStyle.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("CrotchStyle", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboCrotchStyle.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("WaistCirUserFac", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRed(0).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ChestCirUserFac", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRed(1).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("BreastCirUserFac", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRed(2).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("LgButtCirUserFac", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRed(3).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("ThCirUserFac", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = cboRed(4).Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("fileno", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = txtFileNo.Text
                            ElseIf acAttRef.Tag.ToUpper().Equals("AxillaBackNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_nAxillaBackNeckRad
                            ElseIf acAttRef.Tag.ToUpper().Equals("AxillaFrontNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_nAxillaFrontNeckRad
                            ElseIf acAttRef.Tag.ToUpper().Equals("ABNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_nABNRadRight
                            ElseIf acAttRef.Tag.ToUpper().Equals("AFNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = g_nAFNRadRight
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
    Function FN_ValidateAndCalculateData(ByRef bDisplayErrors As Short) As Short
        'This function checks for
        '   1. Missing data
        '   2. Checks the cut-out and modifiys the
        '      reductions to make it fit
        '
        'The argument bDisplayErrors
        '   bDisplayErrors = True   Then display WARNING error message
        '   bDisplayErrors = False  Don't display WARNING errors
        '
        '   This flag is ignored when it is impossible to continue.
        '
        'This function is used by the projects
        '   BODYDRAW.MAK
        '   BODUSUIT.MAK
        '
        Dim sError As String
        Dim iFatalError As Short
        Dim ii As Short
        Dim sCutOutToBackMinimum As String
        Dim nInchesDiff As Double
        Dim nWaistRem As Double
        Dim nChestRem As Double
        Dim nButtRedAtLimit As Short
        Dim nThighRedAtLimit As Short
        Dim nButtRedIncreased As Short
        Dim nThighRedDecreased As Short
        Dim nCutOutModified As Short

        'Initialise
        FN_ValidateAndCalculateData = False
        sError = ""
        iFatalError = False

        Dim sCircum(16) As Object
        sCircum(0) = "Left shoulder circ."
        sCircum(1) = "Right shoulder circ."
        sCircum(2) = "Neck circ."
        sCircum(3) = "Shoulder width."
        sCircum(4) = "Shoulder to waist."
        sCircum(5) = "Chest circ."
        sCircum(6) = "Waist circ."
        sCircum(9) = "Shoulder to under breast."
        sCircum(10) = "Circ. under breast."
        sCircum(11) = "Circ. over nipple."
        sCircum(12) = "Shoulder to Fold of Buttocks."
        sCircum(13) = "Shoulder to Large Part of Buttocks."
        sCircum(14) = "Circ. of Large Part of Buttocks."
        sCircum(15) = "Left Thigh Circ."
        sCircum(16) = "Right Thigh Circ."

        'FATAL Errors
        '~~~~~~~~~~~~
        'Body measurements (all must be present)
        For ii = 0 To 6
            If Val(txtCir(ii).Text) = 0 Then
                sError = sError & "Missing " & sCircum(ii) & BODYSUIT1.NL
            End If
        Next ii
        For ii = 12 To 16
            If Val(txtCir(ii).Text) = 0 Then
                sError = sError & "Missing " & sCircum(ii) & BODYSUIT1.NL
            End If
        Next ii

        'Bra Cups
        'Note:
        '    The Circumference over nipple is optional unless a cup has been
        '    specified
        '
        If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(9) & BODYSUIT1.NL
        End If
        If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(10) & BODYSUIT1.NL
        End If

        If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
            sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & BODYSUIT1.NL
        End If

        If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
            sError = sError & "Bra disks given! But Missing " & sCircum(9) & BODYSUIT1.NL
        End If

        'If cups or dimensions given then a disk must be present
        'NB
        '    cboXXXXCup.ListIndex = 6 = "None"
        '    cboXXXXCup.ListIndex = 7 = ""

        If cboLeftCup.SelectedIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Left BRA Cup!" & BODYSUIT1.NL
        End If
        If cboRightCup.SelectedIndex < 6 And Val(txtRightDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Right BRA Cup!" & BODYSUIT1.NL
        End If

        'Sex error
        If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.SelectedIndex < 0 Or cboRightCup.SelectedIndex < 0) Then
            sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & BODYSUIT1.NL
        End If

        'Neck at back and front
        Dim sChar As New VB6.FixedLengthString(1)
        sChar.Value = VB.Left(cboBackNeck.Text, 1)
        If sChar.Value = "M" And txtBackNeck.Text = "" Then
            sError = sError & "No dimension for Measured Back neck style!" & BODYSUIT1.NL
        End If

        sChar.Value = VB.Left(cboFrontNeck.Text, 1)
        If sChar.Value = "M" And txtFrontNeck.Text = "" Then
            sError = sError & "No dimension for Measured Front neck style" & BODYSUIT1.NL
        End If

        'Get values from dialog
        BODYSUIT1.g_nBodyShldWidth = fnDisplayToInches(Val(txtCir(3).Text))
        BODYSUIT1.g_sBodyBackNeckStyle = cboBackNeck.Text
        BODYSUIT1.g_sBodyFrontNeckStyle = cboFrontNeck.Text
        'Test minimum value for Shoulder width
        If BODYSUIT1.g_nBodyShldWidth < 1.5 Then
            sError = sError & "Shoulder Width is less than 1-1/2""!" & BODYSUIT1.NL
        End If
        If InStr(BODYSUIT1.g_sBodyBackNeckStyle, "Scoop") > 0 And ((BODYSUIT1.g_nBodyShldWidth - 1) < 1.5) Then
            sError = sError & "With a scoop BACK neck the Shoulder Width is less than 1-1/2""!" & BODYSUIT1.NL
        End If

        If (InStr(BODYSUIT1.g_sBodyFrontNeckStyle, "Scoop") > 0 Or InStr(BODYSUIT1.g_sBodyFrontNeckStyle, "V neck") > 0) And (BODYSUIT1.g_nBodyShldWidth - 1) < 1.5 Then
            sError = sError & "With a " & BODYSUIT1.g_sBodyFrontNeckStyle & " FRONT neck the Shoulder Width is less than 1-1/2""!" & BODYSUIT1.NL
        End If

        '
        If cboLeftAxilla.Text = "" Then
            sError = sError & "Left Axilla not given!" & BODYSUIT1.NL
        End If

        If cboRightAxilla.Text = "" Then
            sError = sError & "Right Axilla not given!" & BODYSUIT1.NL
        End If

        If cboFrontNeck.Text = "" Then
            sError = sError & "Neck not given!" & BODYSUIT1.NL
        End If

        If cboBackNeck.Text = "" Then
            sError = sError & "Back neck not given!" & BODYSUIT1.NL
        End If

        If cboClosure.Text = "" Then
            sError = sError & "Closure not given!" & BODYSUIT1.NL
        End If

        If cboFabric.Text = "" Then
            sError = sError & "Fabric not given!" & BODYSUIT1.NL
        End If

        'Extra Values
        If cboLegStyle.Text = "" Then
            sError = sError & "Leg Style not given!" & BODYSUIT1.NL
        End If
        If cboCrotchStyle.Text = "" Then
            sError = sError & "Crotch Style not given!" & BODYSUIT1.NL
        End If

        'Display Error message (if required) and return
        'These are fatal errors
        If Len(sError) > 0 Then
            MsgBox(sError, 48, "Errors in MainForm Data cannot continue!")
            FN_ValidateAndCalculateData = False
            Exit Function
        Else
            FN_ValidateAndCalculateData = True
        End If

        'Possible FATAL Errors and WARNINGS
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        'Get values from the dialogue
        PR_GetValuesFromDialogue()

        'Figured Measurements
        '
        BODYSUIT1.g_bBodyDiffAxillaHeight = False
        BODYSUIT1.g_bBodyDiffThigh = False
        BODYSUIT1.g_bBodySleeveless = False

        'Check if Axilla values are within an inch of each other
        If System.Math.Abs(BODYSUIT1.g_nBodyLeftShldCirGiven - BODYSUIT1.g_nBodyRightShldCirGiven) > 1 Then 'Greater than 1" difference
            'keep separate
            BODYSUIT1.g_bBodyDiffAxillaHeight = True
        Else
            'Use average (sides same at axilla)
            BODYSUIT1.g_nBodyLeftShldCirGiven = ((BODYSUIT1.g_nBodyLeftShldCirGiven + BODYSUIT1.g_nBodyRightShldCirGiven) / 2)
            BODYSUIT1.g_nBodyRightShldCirGiven = BODYSUIT1.g_nBodyLeftShldCirGiven
        End If
        BODYSUIT1.g_nBodyLeftShldToAxilla = fnRoundInches((BODYSUIT1.g_nBodyLeftShldCirGiven * 0.95) / 3.14) + BODYSUIT1.BODY_INCH1_2
        BODYSUIT1.g_nBodyRightShldToAxilla = fnRoundInches((BODYSUIT1.g_nBodyRightShldCirGiven * 0.95) / 3.14) + BODYSUIT1.BODY_INCH1_2


        If BODYSUIT1.g_sBodyLeftSleeve = "Sleeveless" Then
            BODYSUIT1.g_nBodyLeftAxillaCir = fnRoundInches(BODYSUIT1.g_nBodyLeftShldCirGiven * 0.9) '90% figured
            BODYSUIT1.g_bBodySleeveless = True
        Else
            BODYSUIT1.g_nBodyLeftAxillaCir = BODYSUIT1.g_nBodyLeftShldCirGiven
        End If
        If BODYSUIT1.g_sBodyRightSleeve = "Sleeveless" Then
            BODYSUIT1.g_nBodyRightAxillaCir = fnRoundInches(BODYSUIT1.g_nBodyRightShldCirGiven * 0.9) '90% figured
            BODYSUIT1.g_bBodySleeveless = True
        Else
            BODYSUIT1.g_nBodyRightAxillaCir = BODYSUIT1.g_nBodyRightShldCirGiven
        End If


        'Thighs
        If System.Math.Abs(BODYSUIT1.g_nBodyLeftThighCirGiven - BODYSUIT1.g_nBodyRightThighCirGiven) > 1 Then
            BODYSUIT1.g_bBodyDiffThigh = True
        End If
        'Use Smallest Thigh for calculations
        'We will draw both largest and smallest thigh
        If (BODYSUIT1.g_nBodyLeftThighCirGiven <= BODYSUIT1.g_nBodyRightThighCirGiven) Then
            BODYSUIT1.g_nBodySmallestThighGiven = BODYSUIT1.g_nBodyLeftThighCirGiven
            BODYSUIT1.g_sBodySmallestThighGiven = "Left"
            BODYSUIT1.g_nBodyLargestThighGiven = BODYSUIT1.g_nBodyRightThighCirGiven
            BODYSUIT1.g_sBodyLargestThighGiven = "Right"
        Else
            BODYSUIT1.g_nBodySmallestThighGiven = BODYSUIT1.g_nBodyRightThighCirGiven
            BODYSUIT1.g_sBodySmallestThighGiven = "Right"
            BODYSUIT1.g_nBodyLargestThighGiven = BODYSUIT1.g_nBodyLeftThighCirGiven
            BODYSUIT1.g_sBodyLargestThighGiven = "Left"
        End If

        ''        Else
        '          'Within an 1" of each other therefore use smallest thigh size.
        '           BODYSUIT1.g_sBodySmallestThighGiven = "Left"
        '           If (BODYSUIT1.g_nBodyLeftThighCirGiven < BODYSUIT1.g_nBodyRightThighCirGiven) Then
        '                   BODYSUIT1.g_nBodySmallestThighGiven = BODYSUIT1.g_nBodyLeftThighCirGiven
        '               Else
        '                   BODYSUIT1.g_nBodySmallestThighGiven = BODYSUIT1.g_nBodyRightThighCirGiven
        '           End If
        '   End If


        If InStr(BODYSUIT1.g_sBodyCrotchStyle, "Fly") <> 0 And g_sSex = "Male" Then
            BODYSUIT1.g_nBodyButtBackSeamRatio = 0.6 '60%
            BODYSUIT1.g_nBodyButtFrontSeamRatio = 0.4 '40%
        Else
            BODYSUIT1.g_nBodyButtBackSeamRatio = 0.5 '50%
            BODYSUIT1.g_nBodyButtFrontSeamRatio = 0.5 '50%
        End If

        'MsgBox "1"
        'Figure the values
        g_nChestCir = fnRoundInches(g_nChestCirGiven * g_nChestCirRed) / 2 'half scale
        g_nWaistCir = fnRoundInches(g_nWaistCirGiven * g_nWaistCirRed) / 2 'half-scale
        g_nShldToFold = fnRoundInches(g_nShldToFoldGiven * g_nShldToFoldRed) + BODYSUIT1.BODY_INCH1_2
        g_nShldToWaist = fnRoundInches(g_nShldToWaistGiven * g_nShldToWaistRed) + BODYSUIT1.BODY_INCH1_2
        g_nButtocksLength = fnRoundInches((g_nShldToFold - g_nShldToWaist) / 3)
        g_nUnderBreastCir = fnRoundInches(g_nUnderBreastCirGiven * g_nUnderBreastCirRed) / 2 'half scale
        g_nNeckCir = fnRoundInches(g_nNeckCirGiven * g_nNeckCirRed)
        g_nShldToBreast = fnRoundInches(g_nShldToBreastGiven * g_nShldToBreastRed) + BODYSUIT1.BODY_INCH1_2


        'NOTE:-
        '    There can be no more that 5% difference between the reductions
        '    at each circumference.


        'MsgBox "2"

        'CutOut and Back Seam Test
        nButtRedAtLimit = False
        nThighRedAtLimit = False
        nButtRedIncreased = False
        nThighRedDecreased = False
        sCutOutToBackMinimum = "2-1/2"
        g_nCutOutToBackMinimum = (2.5) / 2 'on the half scale

        'Original 85% mark at Fold/Groin
        g_nButtocksCir = fnRoundInches(g_nButtocksCirGiven * 0.85) / 2 'half scale
        g_nGroinHeight = fnRoundInches(BODYSUIT1.g_nBodySmallestThighGiven * 0.85) / 2 'half scale
        g_nHalfGroinHeight = (g_nGroinHeight / 2)
        g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
        g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
        g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * BODYSUIT1.g_nBodyButtBackSeamRatio
        g_n85FoldHeight = g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2))
        'Recalculate below. Above is only to get the original mark

        Do
            If g_nButtocksCirRed = 0.9 Then 'at Reduction limit of 90%
                nButtRedAtLimit = True
            End If
            'I had to use Val and Str$ functions because VB3 has a problem
            'when returning values from the minus operation,
            'i.e. nThighCirRed was not equal to .8 exactly (but it should be!!!)
            If Val(Str(g_nThighCirRed)) = 0.8 Then 'at Reduction limit of 80%
                nThighRedAtLimit = True
            End If
            If nThighRedAtLimit And nButtRedAtLimit Then
                Exit Do
            End If
            g_nButtocksCir = fnRoundInches(g_nButtocksCirGiven * g_nButtocksCirRed) / 2 'half scale
            g_nGroinHeight = fnRoundInches(BODYSUIT1.g_nBodySmallestThighGiven * g_nThighCirRed) / 2 'half scale
            g_nHalfGroinHeight = (g_nGroinHeight / 2)
            g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
            g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
            g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * BODYSUIT1.g_nBodyButtBackSeamRatio
            If g_nButtBackSeam < g_nCutOutToBackMinimum Then 'less than full scale minimum
                If nButtRedAtLimit Then
                    g_nThighCirRed = g_nThighCirRed - 0.05 'decrease the reduction by 5%
                    nThighRedDecreased = True
                Else
                    g_nButtocksCirRed = g_nButtocksCirRed + 0.05 'increase the reduction by 5%
                    nButtRedIncreased = True
                End If
            Else
                Exit Do 'greater than 3" full scale - acceptable
            End If
        Loop
        'MsgBox "3"

        If nButtRedIncreased Then
            cboRed(3).Text = CStr(g_nButtocksCirRed)
            sError = sError & "With given Buttock Reduction, distance from back of Cut-Out" + BODYSUIT1.NL
            sError = sError & "to largest part of buttocks is less than " & sCutOutToBackMinimum & " inches!" + BODYSUIT1.NL
            sError = sError & "Buttocks Reduction increased to rectify this." + BODYSUIT1.NL
            'sError = sError + "ref:- " + BODYSUIT1.NL
        End If
        If nThighRedDecreased Then
            cboRed(4).Text = CStr(g_nThighCirRed)
            sError = sError & "With given Thigh Reduction, distance from back of Cut-Out" + BODYSUIT1.NL
            sError = sError & "to largest part of buttocks is less than " & sCutOutToBackMinimum & " inches!" + BODYSUIT1.NL
            sError = sError & "Thigh Reduction decreased to rectify this." + BODYSUIT1.NL
        End If

        'Check that all of the reductions are within 5% of each other
        Dim nMaxRed, nMinRed, nCurrentValue As Short
        nMinRed = 0
        nMaxRed = 0

        'turn all reductions into integers for easy comparing
        'Waist
        nMinRed = (g_nWaistCirRed * 100)
        nMaxRed = nMinRed
        'Chest
        nCurrentValue = (g_nChestCirRed * 100)
        If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
        End If
        'Buttocks
        nCurrentValue = (g_nButtocksCirRed * 100)
        If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
        End If
        'Thigh
        nCurrentValue = (g_nThighCirRed * 100)
        If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
        End If

        If (nMaxRed - nMinRed) > 5 Then
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "All Reductions should be within 5% of each other" + BODYSUIT1.NL
        End If

        'MsgBox "4"

        'One last time with Reductions found above
        '
        'Calculate for Average / Smallest thigh
        g_nButtocksCir = fnRoundInches(g_nButtocksCirGiven * g_nButtocksCirRed) / 2 'half scale
        g_nGroinHeight = fnRoundInches(BODYSUIT1.g_nBodySmallestThighGiven * g_nThighCirRed) / 2 'half scale
        g_nHalfGroinHeight = (g_nGroinHeight / 2)
        g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
        g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
        g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * BODYSUIT1.g_nBodyButtBackSeamRatio
        g_nGroinHeight = g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength - BODYSUIT1.BODY_INCH3_4) ^ 2))

        'Calculate largest thigh (if given)
        g_nLT_GroinHeight = fnRoundInches(BODYSUIT1.g_nBodyLargestThighGiven * g_nThighCirRed) / 2 'half scale
        g_nLT_HalfGroinHeight = (g_nLT_GroinHeight / 2)
        g_nLT_ButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nLT_HalfGroinHeight ^ 2))
        g_nLT_GroinHeight = g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - ((g_nButtocksLength - BODYSUIT1.BODY_INCH3_4) ^ 2))

        If g_nButtBackSeam <= 0 Then
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + BODYSUIT1.NL
            sError = sError & "is still negative or zero!" + BODYSUIT1.NL
            sError = sError & "Even with the modified reductions." + BODYSUIT1.NL
            iFatalError = True
        ElseIf g_nButtBackSeam < g_nCutOutToBackMinimum Then  'less than required minimum
            sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
            sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + BODYSUIT1.NL
            sError = sError & "is still less than " & sCutOutToBackMinimum & " inches!" + BODYSUIT1.NL
            sError = sError & "Even with the modified reductions." + BODYSUIT1.NL
        End If

        'All other seams test
        'Set initila values
        g_nButtFrontSeam = (g_nButtocksCir - g_nButtCirCalc) * BODYSUIT1.g_nBodyButtFrontSeamRatio
        g_nButtCutOut = g_nButtCirCalc - (g_nButtBackSeam + g_nButtFrontSeam)

        nInchesDiff = Int(System.Math.Abs(g_nWaistCirGiven - g_nButtocksCirGiven))
        g_nWaistFrontSeam = (g_nButtFrontSeam * 2)
        If (g_nWaistCirGiven > g_nButtocksCirGiven) And nInchesDiff > 0 Then
            g_nWaistFrontSeam = g_nWaistFrontSeam + (nInchesDiff * BODYSUIT1.BODY_INCH1_4)
        Else
            g_nWaistFrontSeam = g_nWaistFrontSeam - (nInchesDiff * BODYSUIT1.BODY_INCH1_8)
        End If

        nInchesDiff = round(System.Math.Abs(g_nChestCirGiven - g_nButtocksCirGiven))
        g_nChestFrontSeam = (g_nButtFrontSeam * 2)
        If (g_nChestCirGiven > g_nButtocksCirGiven) And nInchesDiff > 0 Then
            g_nChestFrontSeam = g_nChestFrontSeam + (nInchesDiff * BODYSUIT1.BODY_INCH1_4)
        Else
            g_nChestFrontSeam = g_nChestFrontSeam - (nInchesDiff * BODYSUIT1.BODY_INCH1_8)
        End If

        'Use the half scale for both
        g_nChestFrontSeam = g_nChestFrontSeam / 2
        g_nWaistFrontSeam = g_nWaistFrontSeam / 2

        'MsgBox "5"
        'MsgBox "g_nWaistFrontSeam=" & Str$(g_nWaistFrontSeam) & BODYSUIT1.NL & "g_nChestFrontSeam=" & Str$(g_nChestFrontSeam)


        g_nCutOut = (g_nButtCutOut * 0.87)
        nCutOutModified = 0
        Do
            'MsgBox "Loop Yes"
            nWaistRem = g_nWaistCir - (g_nCutOut + (g_nWaistFrontSeam * 2))
            If (nWaistRem / 2) < g_nCutOutToBackMinimum Then 'less than allowable minimum @ full scale
                If nWaistRem <= 0 Then
                    sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
                    sError = sError & "Distance from back of Cut-Out to profile at WAIST" + BODYSUIT1.NL
                    sError = sError & "is negative or zero!" + BODYSUIT1.NL
                    If nCutOutModified > 0 Then
                        sError = sError & "Even with the front of the Cut-Out at the Waist lowered." + BODYSUIT1.NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                'Lower cut-out at Waist
                '                g_nWaistFrontSeam = g_nWaistFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nWaistRem / 2))) / 2)
                g_nWaistFrontSeam = g_nWaistFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nWaistRem / 2))) / 1)
                nCutOutModified = nCutOutModified + 1
                If nCutOutModified > 100 Then
                    MsgBox("Infinite looping while calculating back profile at waist, Contact Systems Support!")
                    Exit Do
                End If
            Else
                If nCutOutModified > 0 Then
                    sError = sError & "Original distance from back of Cut-Out to profile at the WAIST" + BODYSUIT1.NL
                    sError = sError & "was less than " & sCutOutToBackMinimum & " inches!" + BODYSUIT1.NL
                    Select Case nCutOutModified
                        Case 1
                            sError = sError & "Cut-Out at WAIST had to be lowered " & Str(nCutOutModified) & " time." & BODYSUIT1.NL
                        Case Is > 1
                            sError = sError & "Cut-Out at WAIST had to be lowered " & Str(nCutOutModified) & " times." & BODYSUIT1.NL
                    End Select
                End If
                Exit Do 'greater than g_nCutOutToBackMinimum - acceptable
            End If
        Loop
        g_nWaistBackSeam = nWaistRem / 2
        'MsgBox "6"

        'CHEST
        nCutOutModified = 0
        Do
            nChestRem = g_nChestCir - (g_nCutOut + g_nWaistFrontSeam + g_nChestFrontSeam)
            If (nChestRem / 2) < g_nCutOutToBackMinimum Then 'less than acceptable" @ full scale
                If nChestRem <= 0 Then
                    sError = sError + BODYSUIT1.NL + "Severe - Warning!" + BODYSUIT1.NL
                    sError = sError & "Distance from back of Cut-Out to profile at the CHEST" + BODYSUIT1.NL
                    sError = sError & "is negative or zero!" + BODYSUIT1.NL
                    If nCutOutModified > 0 Then
                        sError = sError & "Even with the Cut Out at the Chest lowered." + BODYSUIT1.NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                'lower cut=out at Chest
                '                g_nChestFrontSeam = g_nChestFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nChestRem / 2))) / 2)
                g_nChestFrontSeam = g_nChestFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nChestRem / 2))) / 1)
                nCutOutModified = nCutOutModified + 1
                If nCutOutModified > 100 Then
                    MsgBox("Infinite looping while calculating back profile at chest, Contact support")
                    Exit Do
                End If
            Else
                If nCutOutModified > 0 Then
                    sError = sError & "Distance from back of Cut-Out at the CHEST" + BODYSUIT1.NL
                    sError = sError & "was less than " & sCutOutToBackMinimum & " inches!" + BODYSUIT1.NL
                    Select Case nCutOutModified
                        Case 1
                            sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str(nCutOutModified) & " time." & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
                        Case Is > 1
                            sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str(nCutOutModified) & " times." & BODYSUIT1.NL
                            'sError = sError + "ref:- " + BODYSUIT1.NL
                    End Select
                End If
                Exit Do 'greater than g_nCutOutToBackMinimum - acceptable
            End If
        Loop
        g_nChestBackSeam = nChestRem / 2

        'Snap Crotch / Gusset warning wrt brief
        If InStr(g_sLegStyle, "Brief") > 0 And BODYSUIT1.g_sBodyCrotchStyle = "Gusset" Then
            sError = sError + BODYSUIT1.NL + "Information!" + BODYSUIT1.NL
            sError = sError & "A crotch style of ""Gusset"" has been selected for a BRIEF." + BODYSUIT1.NL
            sError = sError & "You should draw as a Snap Crotch" + BODYSUIT1.NL
        End If

        If InStr(g_sLegStyle, "Brief") > 0 And BODYSUIT1.g_sBodyCrotchStyle = "Open Crotch" Then
            sError = sError + BODYSUIT1.NL + "Information!" + BODYSUIT1.NL
            sError = sError & "A crotch style of ""Open"" has been selected for a BRIEF." + BODYSUIT1.NL
            sError = sError & "You should not use an Open crotch style with a Brief" + BODYSUIT1.NL
        End If
        If InStr(BODYSUIT1.g_sBodyCrotchStyle, "Hor") > 0 And g_nAge < 3 Then
            sError = sError + BODYSUIT1.NL + "Information!" + BODYSUIT1.NL
            sError = sError & "A Horizontal Fly has been selected for a child under 3." + BODYSUIT1.NL
            sError = sError & "The horizontal fly chart does not contain an entry for children under 3 years old." + BODYSUIT1.NL
        End If


        'Display Error message (if required) and return
        'There could be fatal errors found
        If Len(sError) > 0 And bDisplayErrors Then
            If iFatalError = True Then
                MsgBox(sError, 64, "Warning - Problems with data")
                FN_ValidateAndCalculateData = False
            Else
                sError = sError + BODYSUIT1.NL + "The above problems have been found in the data do you" + BODYSUIT1.NL
                sError = sError & "wish to continue ?"
                If MsgBox(sError, 52, "Severe Problems with data") = IDYES Then
                    FN_ValidateAndCalculateData = True
                Else
                    FN_ValidateAndCalculateData = False
                End If
            End If
        Else
            FN_ValidateAndCalculateData = True
        End If

    End Function
    Sub PR_GetValuesFromDialogue()
        'Get values from dialog box
        g_sFileNo = txtFileNo.Text
        g_sPatient = txtPatientName.Text
        g_sDiagnosis = txtDiagnosis.Text
        g_nAge = Val(txtAge.Text)

        'Set Adult status
        If g_nAge > 10 Then
            g_nAdult = True
        Else
            g_nAdult = False
        End If
        g_sSex = txtSex.Text
        g_sWorkOrder = txtWorkOrder.Text

        BODYSUIT1.g_nBodyLeftShldCirGiven = fnDisplayToInches(Val(txtCir(0).Text))
        BODYSUIT1.g_nBodyRightShldCirGiven = fnDisplayToInches(Val(txtCir(1).Text))
        g_nNeckCirGiven = fnDisplayToInches(Val(txtCir(2).Text))
        BODYSUIT1.g_nBodyShldWidth = fnDisplayToInches(Val(txtCir(3).Text))
        g_nShldToWaistGiven = fnDisplayToInches(Val(txtCir(4).Text))
        g_nChestCirGiven = fnDisplayToInches(Val(txtCir(5).Text))
        g_nWaistCirGiven = fnDisplayToInches(Val(txtCir(6).Text))
        g_nShldToFoldGiven = fnDisplayToInches(Val(txtCir(12).Text))
        g_nButtocksCirGiven = fnDisplayToInches(Val(txtCir(14).Text))
        BODYSUIT1.g_nBodyLeftThighCirGiven = fnDisplayToInches(Val(txtCir(15).Text))
        BODYSUIT1.g_nBodyRightThighCirGiven = fnDisplayToInches(Val(txtCir(16).Text))

        If txtCir(9).Text <> "" Then
            g_nShldToBreastGiven = fnDisplayToInches(Val(txtCir(9).Text))
        End If

        If txtCir(10).Text <> "" Then
            g_nUnderBreastCirGiven = fnDisplayToInches(Val(txtCir(10).Text))
        End If

        If txtCir(11).Text <> "" Then
            g_nNippleCirGiven = fnDisplayToInches(Val(txtCir(11).Text))
        End If

        BODYSUIT1.g_sBodyLeftSleeve = cboLeftAxilla.Text
        BODYSUIT1.g_sBodyRightSleeve = cboRightAxilla.Text
        BODYSUIT1.g_sBodyFrontNeckStyle = cboFrontNeck.Text
        BODYSUIT1.g_sBodyBackNeckStyle = cboBackNeck.Text
        g_sClosure = cboClosure.Text

        If (BODYSUIT1.g_sBodyFrontNeckStyle = "Measured Scoop" Or BODYSUIT1.g_sBodyFrontNeckStyle = "Measured V neck" Or InStr(BODYSUIT1.g_sBodyFrontNeckStyle, "Turtle") > 0) Then
            g_nFrontNeckSize = fnDisplayToInches(Val(txtFrontNeck.Text))
        End If
        If (BODYSUIT1.g_sBodyBackNeckStyle = "Measured Scoop") Then
            g_nBackNeckSize = fnDisplayToInches(Val(txtBackNeck.Text))
        End If
        g_sFabric = cboFabric.Text
        If optLeftLeg(0).Checked = True Then
            g_sLeftLeg = "Panty" 'Options are: Panty & Brief
        ElseIf optLeftLeg(1).Checked = True Then
            g_sLeftLeg = "Brief"
        End If

        If optRightLeg(0).Checked = True Then
            g_sRightLeg = "Panty" 'Options are: Panty & Brief
        ElseIf optRightLeg(1).Checked = True Then
            g_sRightLeg = "Brief"
        End If

        g_sLegStyle = cboLegStyle.Text
        BODYSUIT1.g_sBodyCrotchStyle = cboCrotchStyle.Text

        'Reductions
        g_nNeckCirRed = 0.9 '90%
        g_nShldToWaistRed = 0.95 '95%
        g_nShldToFoldRed = 0.95 '95%
        g_nShldToBreastRed = 0.95 '95%

        If txtSex.Text = "Female" Then
            If cboRed(2).Text <> "" Then
                g_nUnderBreastCirRed = Val(cboRed(2).Text)
            Else
                g_nUnderBreastCirRed = 0.9 '90% default
            End If
        End If

        If cboRed(1).Text <> "" Then
            g_nChestCirRed = Val(cboRed(1).Text)
        Else
            g_nChestCirRed = 0.85 '85% default
        End If
        If cboRed(0).Text <> "" Then
            g_nWaistCirRed = Val(cboRed(0).Text)
        Else
            g_nWaistCirRed = 0.85 '85% default

        End If
        If cboRed(3).Text <> "" Then
            g_nButtocksCirRed = Val(cboRed(3).Text)
        Else
            g_nButtocksCirRed = 0.85 '85% default
        End If

        If cboRed(4).Text <> "" Then
            g_nThighCirRed = Val(cboRed(4).Text)
        Else
            g_nThighCirRed = 0.85 '85% default
        End If
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("BodySuit", "BODYDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("BodySuit", "BODYDIC")
            End If

            Dim resbuf As New ResultBuffer
            Dim ii As Double
            For ii = 0 To 6
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtCir(ii).Text))
            Next ii
            For ii = 9 To 16
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtCir(ii).Text))
            Next ii

            For ii = 0 To 4
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboRed(ii).Text))
            Next ii

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboLeftAxilla.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboRightAxilla.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboFrontNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFrontNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboBackNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtBackNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboClosure.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboFabric.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboLegStyle.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboCrotchStyle.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optLeftLeg(0).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optLeftLeg(1).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optRightLeg(0).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optRightLeg(1).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), lblFrontNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), lblBackNeck.Text))

            For ii = 0 To 6
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), lblCir(ii).Text))
            Next ii
            For ii = 9 To 16
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), lblCir(ii).Text))
            Next ii

            _sClass.SetXrecord(resbuf, "BodySuit", "BODYDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("BodySuit", "BODYDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim ii As Double
                For ii = 0 To 6
                    txtCir(ii).Text = arr(ii).Value
                Next ii
                For ii = 9 To 16
                    txtCir(ii).Text = arr(ii - 2).Value
                Next ii

                For ii = 0 To 4
                    cboRed(ii).Text = arr(ii + 15).Value
                Next ii
                cboLeftAxilla.Text = arr(20).Value
                cboRightAxilla.Text = arr(21).Value
                cboFrontNeck.Text = arr(22).Value
                txtFrontNeck.Text = arr(23).Value
                cboBackNeck.Text = arr(24).Value
                txtBackNeck.Text = arr(25).Value
                cboClosure.Text = arr(26).Value
                cboFabric.Text = arr(27).Value
                cboLegStyle.Text = arr(28).Value
                cboCrotchStyle.Text = arr(29).Value
                optLeftLeg(0).Checked = arr(30).Value
                optLeftLeg(1).Checked = arr(31).Value
                optRightLeg(0).Checked = arr(32).Value
                optRightLeg(1).Checked = arr(33).Value
                lblFrontNeck.Text = arr(34).Value
                lblBackNeck.Text = arr(35).Value

                For ii = 0 To 6
                    lblCir(ii).Text = arr(ii + 36).Value
                Next ii
                For ii = 9 To 16
                    lblCir(ii).Text = arr(ii + 34).Value
                Next ii
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
            resbuf = _sClass.GetXrecord("BodyLeftLeg", "BODYLEFTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                If arr(37).Value = True Then
                    m_nLeftKneePosition = -1
                ElseIf arr(38).Value = True Then
                    m_nLeftKneePosition = 1
                End If
            End If
            resbuf = _sClass.GetXrecord("BodyRightLeg", "BODYRIGHTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                If arr(37).Value = True Then
                    m_nRightKneePosition = -1
                ElseIf arr(38).Value = True Then
                    m_nRightKneePosition = 1
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Sub PR_CalcPolar(ByRef xyStart As BODYSUIT1.XY, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As BODYSUIT1.XY)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double
        'Convert from degees to radians
        nAngle = nAngle * BODYSUIT1.PI / 180
        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.X + A
        xyReturn.Y = xyStart.Y + B
    End Sub
    Function FN_CalcAngle(ByRef xyStart As BODYSUIT1.XY, ByRef xyEnd As BODYSUIT1.XY) As Double
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
        rAngle = System.Math.Atan(y / X) * (180 / BODYSUIT1.PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If
    End Function
    Function FN_LinLinInt(ByRef xyLine1Start As BODYSUIT1.XY, ByRef xyLine1End As BODYSUIT1.XY, ByRef xyLine2Start As BODYSUIT1.XY,
                          ByRef xyLine2End As BODYSUIT1.XY, ByRef xyInt As BODYSUIT1.XY) As Short
        'Function:
        '       BOOLEAN = FN_LinLinInt( xyLine1Start, xyLine1End, xyLine2Start, xyLine2End, xyInt);
        'Parameters:
        '       xyLine1Start = xyLine1Start.X, xyLine1Start.Y
        '       xyLine1End = xyLine1End.X, xyLine1End.Y
        '       xyLine2Start = xyLine2Start.X, xyLine2Start.Y
        '       xyLine2End = xyLine2End.X, xyLine2End.Y
        '
        'Returns:
        '       True if intersection found and lies on the line
        '       False if no intesection
        '       xyInt =  intersection
        '
        Dim nY, nSlope1, nK2, nK1, nM1, nCase, nX As Double
        Dim nM2, nSlope2 As Object

        'Initialy false
        FN_LinLinInt = False

        'Calculate slope of lines
        nCase = 0
        nSlope1 = FN_CalcAngle(xyLine1Start, xyLine1End)
        If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
        If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2

        nSlope2 = FN_CalcAngle(xyLine2Start, xyLine2End)
        If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
        If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

        Select Case nCase
            Case 0
                'Both lines are Non-Orthogonal Lines
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nM2 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                If (nM1 = nM2) Then Exit Function 'Parallel lines
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nK2 = xyLine2Start.Y - (nM2 * xyLine2Start.X) 'Y-Axis intercept
                If (nK1 = nK2) Then Exit Function
                'Find X
                nX = (nK2 - nK1) / (nM1 - nM2)
                'Find Y
                nY = (nM1 * nX) + nK1

            Case 1
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nY = xyLine1Start.Y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1

            Case 2
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nX = xyLine1Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1

            Case 4
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nY = xyLine2Start.Y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1

            Case 5
                'Parallel orthogonal lines, no intersection possible
                Exit Function

            Case 6
                'Line1 is Vertical and the Line2 is Horizontal
                nX = xyLine1Start.X
                nY = xyLine2Start.Y

            Case 8
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nX = xyLine2Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1

            Case 9
                'Line1 is Horizontal and the Line2 is Vertical
                nX = xyLine2Start.X
                nY = xyLine1Start.Y

            Case 10
                'Parallel orthogonal lines, no intersection possible
                Exit Function

            Case Else
                Exit Function
        End Select

        'Ensure that the points X and Y are on the lines
        xyInt.X = nX
        xyInt.Y = nY

        'Line 1
        If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
        If (nY < min(xyLine1Start.Y, xyLine1End.Y) Or nY > max(xyLine1Start.Y, xyLine1End.Y)) Then Exit Function

        'Line 2
        If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
        If (nY < min(xyLine2Start.Y, xyLine2End.Y) Or nY > max(xyLine2Start.Y, xyLine2End.Y)) Then Exit Function

        FN_LinLinInt = True
    End Function
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
    Function FN_CirLinInt(ByRef xyStart As BODYSUIT1.XY, ByRef xyEnd As BODYSUIT1.XY, ByRef xyCen As BODYSUIT1.XY,
                          ByRef nRad As Double, ByRef xyInt As BODYSUIT1.XY) As Short
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
    Private Sub PR_SetFlySize(ByRef sType As String)

        'Sets two global variables:
        '   g_sFlySize (type string)
        '   g_nFlyLength (type double) not at this time

        Dim nCutOutDia As Double
        'nCutOutDia is in inches
        nCutOutDia = fnRoundInches(g_nCutOutRadius * 2)
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
        nCutOutDia = fnRoundInches(g_nCutOutRadius * 2)
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
            If Not acBlkTbl.Has("BODYLEFTLEG") Then
                'MsgBox("Can't find BODYLEFTLEG symbol to update!", 16, "Waist Figure - Dialogue")
                Return False
            End If
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForRead)
            For Each objID As ObjectId In acBlkTblRec
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                If TypeOf dbObj Is BlockReference Then
                    Dim blkRef As BlockReference = dbObj
                    If blkRef.Name = "BODYLEFTLEG" Then
                        strLeftUidLeg = "1"
                        For Each attributeID As ObjectId In blkRef.AttributeCollection
                            Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                            If TypeOf attRefObj Is AttributeReference Then
                                Dim acAttRef As AttributeReference = attRefObj
                                If acAttRef.Tag.ToUpper().Equals("TapeLengthsPt1", StringComparison.InvariantCultureIgnoreCase) Then
                                    strLeftLengths = acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt2", StringComparison.InvariantCultureIgnoreCase) Then
                                    strLeftLengths = strLeftLengths & acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat1", StringComparison.InvariantCultureIgnoreCase) Then
                                    strLeftLegPleats = acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat2", StringComparison.InvariantCultureIgnoreCase) Then
                                    strLeftLegPleats = strLeftLegPleats & acAttRef.TextString
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            If Not acBlkTbl.Has("BODYRIGHTLEG") Then
                'MsgBox("Can't find BODYRIGHTLEG symbol to update!", 16, "Waist Figure - Dialogue")
                Return False
            End If
            For Each objID As ObjectId In acBlkTblRec
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                If TypeOf dbObj Is BlockReference Then
                    Dim blkRef As BlockReference = dbObj
                    If blkRef.Name = "BODYRIGHTLEG" Then
                        strRightUidLeg = "1"
                        For Each attributeID As ObjectId In blkRef.AttributeCollection
                            Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                            If TypeOf attRefObj Is AttributeReference Then
                                Dim acAttRef As AttributeReference = attRefObj
                                If acAttRef.Tag.ToUpper().Equals("TapeLengthsPt1", StringComparison.InvariantCultureIgnoreCase) Then
                                    strRightLengths = acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengthsPt2", StringComparison.InvariantCultureIgnoreCase) Then
                                    strRightLengths = strRightLengths & acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat1", StringComparison.InvariantCultureIgnoreCase) Then
                                    strRightLegPleats = acAttRef.TextString
                                ElseIf acAttRef.Tag.ToUpper().Equals("TopLegPleat2", StringComparison.InvariantCultureIgnoreCase) Then
                                    strRightLegPleats = strRightLegPleats & acAttRef.TextString
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End Using
    End Function
    Function FN_ValidateLegData() As Short
        Dim sError As String
        sError = ""
        If g_sLeftLeg = "Panty" And Val(strLeftUidLeg) = 0 Then
            sError = sError & "Panty requested for Left leg but no data found!" & BODYSUIT1.NL
        End If
        If g_sRightLeg = "Panty" And Val(strRightUidLeg) = 0 Then
            sError = sError & "Panty requested for Right leg but no data found!" & BODYSUIT1.NL
        End If

        'Display errors if found
        If Len(sError) > 0 Then
            MsgBox(sError, 0, "Missing Leg Data")
            FN_ValidateLegData = False
        Else
            FN_ValidateLegData = True
        End If
    End Function
    Function FN_InitiliseLegRoutines() As Short
        '
        Dim ii, fFileNum As Short
        Dim sFileName As String

        'Open leg template data file
        'Check to see if has been loaded from file before,
        'so we don't load it twice
        Dim m_bFileLoaded As Boolean = False
        If Not m_bFileLoaded Then
            fFileNum = FreeFile()
            ''sFileName = g_sPathJOBST & "\TEMPLTS\Wh50mmhg.DAT"
            sFileName = fnGetSettingsPath("LookupTables") & "\Wh50mmhg.DAT"
            If FileLen(sFileName) = 0 Then
                MsgBox("Can't open the template file " & sFileName & ".  Unable to draw the leg curve.", 48)
                FN_InitiliseLegRoutines = False
                Exit Function
            End If
            ii = 0
            FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
            Dim m_nNo(30) As Double
            Dim m_sScale(30) As String
            Do While Not EOF(fFileNum) And ii < 30
                ii = ii + 1
                Input(fFileNum, m_nNo(ii))
                Input(fFileNum, m_sScale(ii))
                Input(fFileNum, m_nSpace(ii))
                Input(fFileNum, m_n20Len(ii))
                Input(fFileNum, m_nReduction(ii))
            Loop
            FileClose(fFileNum)
            m_bFileLoaded = True
        End If
        FN_InitiliseLegRoutines = True
    End Function
    Sub PR_CalculateLeg(ByRef sLeg As String, ByRef nLength As Double)
        'Routine to calculate the leg
        'Returns:   nLength = the length the of the leg
        'Updates:   m_*  (Module level variables)
        '
        'Assumes:   1. That data exists.
        '           2. That FN_InitiliseLegRoutines has been called
        '

        Dim LegProfile As curve
        Dim sTapes As String
        Dim sPleats As String
        Dim nValue, nSpace, nScale As Double
        Dim nTopLegPleat1, nFootPleat1, nFootPleat2, nTopLegPleat2 As Double
        Dim iLastTape, iFirstTape, bAboveKnee As Short
        Dim nn, ii, iValue As Short
        Dim iLowest As Short
        Dim nValueLowest As Double
        Dim xyTmp, xyTmp1 As BODYSUIT1.XY

        If sLeg = "Left" Then
            sTapes = strLeftLengths
            sPleats = strLeftLegPleats
            ''iValue = Val(txtLeftKneePosition.Text)
            iValue = m_nLeftKneePosition
            If iValue = 1 Then
                g_bLeftAboveKnee = True
            Else
                g_bLeftAboveKnee = False
            End If
            bAboveKnee = g_bLeftAboveKnee
        Else
            sTapes = strRightLengths
            sPleats = strRightLegPleats
            ''iValue = Val(txtRightKneePosition.Text)
            iValue = m_nRightKneePosition
            If iValue = 1 Then
                g_bRightAboveKnee = True
            Else
                g_bRightAboveKnee = False
            End If
            bAboveKnee = g_bRightAboveKnee
        End If

        If sTapes = "" Then Exit Sub

        'Split leg tape circumferances and establish start and finish tapes
        LegProfile.n = 0
        iFirstTape = -1
        iLastTape = 30
        Dim X(30), Y(30) As Double
        LegProfile.X = X
        LegProfile.y = Y
        For ii = 0 To 29
            nValue = Val(Mid(sTapes, (ii * 4) + 1, 4)) / 10
            If nValue > 0 Then
                LegProfile.n = LegProfile.n + 1
                nValue = fnDisplayToInches(nValue)
                LegProfile.y(LegProfile.n) = nValue
                If sLeg = "Left" Then m_nLeftOriginalTapeValues(ii) = nValue Else m_nRightOriginalTapeValues(ii) = nValue
            End If

            'Set first and last tape (assumes no holes in data)
            If iFirstTape < 0 And nValue > 0 Then iFirstTape = ii
            If iLastTape = 30 And iFirstTape > 0 And nValue = 0 Then iLastTape = ii - 1
        Next ii

        'Using the 50MMHG chart loaded on initilisation and the pleats (if any) we need to
        'calculate the x values of the profile and the overall length of the leg.
        'NB
        '    These are relative X values only as the absolute values will be calculated
        '    only when the leg is to be drawn.
        nFootPleat1 = Val(ARMDIA1.fnGetString(sPleats, 1, ","))
        nFootPleat2 = Val(ARMDIA1.fnGetString(sPleats, 2, ","))
        nTopLegPleat2 = Val(ARMDIA1.fnGetString(sPleats, 3, ","))
        nTopLegPleat1 = Val(ARMDIA1.fnGetString(sPleats, 4, ","))

        nn = 1 'Counter for LegProfile
        nLength = 0 'Note - length to be returned above
        '    nValueLowest = 10000
        For ii = iFirstTape To iLastTape
            If (ii = iFirstTape) Then
                nSpace = 0
            ElseIf (ii = iFirstTape + 1) And (nFootPleat1 <> 0) Then
                nSpace = nFootPleat1
            ElseIf (ii = iFirstTape + 2) And (nFootPleat2 <> 0) Then
                nSpace = nFootPleat2
            ElseIf (ii = iLastTape - 1) And (nTopLegPleat2 <> 0) Then
                nSpace = nTopLegPleat2
            ElseIf (ii = iLastTape) And (nTopLegPleat1 <> 0) Then
                nSpace = nTopLegPleat1
            Else
                nSpace = m_nSpace(ii - 1) 'Spacing from 50MMHG Chart
            End If
            nLength = nLength + nSpace
            LegProfile.X(nn) = nLength

            'Y scaling (m_n20Len(ii) / 20) from 50MMHG Chart
            If ii = iFirstTape Then
                If bAboveKnee Then
                    nScale = 0.95 / 2 '5 reduction
                Else
                    nScale = 0.92 / 2 '8 reduction
                End If
            Else
                nScale = m_n20Len(ii) / 20
            End If
            LegProfile.y(nn) = (LegProfile.y(nn) * nScale) + SEAM

            nn = nn + 1
        Next ii
        PR_MakeXY(xyTmp, LegProfile.X(1), LegProfile.y(1))
        PR_MakeXY(xyTmp1, LegProfile.X(2), LegProfile.y(2))

        'Assign calculated values to left and right as required
        If sLeg = "Left" Then
            g_LeftLegProfile = LegProfile
            m_iLeftFirstTape = iFirstTape
            m_iLeftLastTape = iLastTape
            PR_CalcMidPoint(xyTmp1, xyTmp, xyBDLeftLegLabelPoint)
        Else
            g_RightLegProfile = LegProfile
            m_iRightFirstTape = iFirstTape
            m_iRightLastTape = iLastTape
            PR_CalcMidPoint(xyTmp1, xyTmp, xyBDRightLegLabelPoint)
        End If
    End Sub
    Sub PR_CalcMidPoint(ByRef xyStart As BODYSUIT1.XY, ByRef xyEnd As BODYSUIT1.XY, ByRef xyMid As BODYSUIT1.XY)

        Static aAngle, nLength As Double
        aAngle = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)

        If nLength = 0 Then
            xyMid = xyStart 'Avoid overflow
        Else
            PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
        End If
    End Sub
    Function FN_AverageLeftAndRightLegs() As Short
        Dim iLeftLegProfileCount, ii, bOutwithTolerance, iRightLegProfileCount As Short
        Dim xyTmp, xyTmp1 As BODYSUIT1.XY
        Dim nAverage, nScale, nDistanceApart As Double

        FN_AverageLeftAndRightLegs = True
        nDistanceApart = 0

        'Compare first tapes
        If m_iLeftFirstTape <> m_iRightFirstTape Then
            FN_AverageLeftAndRightLegs = False
            nDistanceApart = 1000 'Dummy to fool averaging below
        End If

        'Compare lengths (takes care of pleats)
        If g_nLeftLegLength <> g_nRightLegLength Then
            FN_AverageLeftAndRightLegs = False
            nDistanceApart = 1000 'Dummy to fool averaging below
        End If

        'Average legs
        iLeftLegProfileCount = 0
        iRightLegProfileCount = 0
        If ii >= m_iRightFirstTape Then g_RightLegProfile.n = g_RightLegProfile.n + 1

        For ii = min(m_iLeftFirstTape, m_iRightFirstTape) To ARMDIA1.max(m_iLeftLastTape, m_iRightLastTape)
            If ii >= m_iLeftFirstTape Then iLeftLegProfileCount = iLeftLegProfileCount + 1
            If ii >= m_iRightFirstTape Then iRightLegProfileCount = iRightLegProfileCount + 1

            'Y scaling (m_n20Len(ii) / 20) from 50MMHG Chart
            If m_nRightOriginalTapeValues(ii) > 0 And m_nLeftOriginalTapeValues(ii) > 0 Then
                If System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii)) <= INCH3_8 Then
                    'Average leg
                    If ii = m_iLeftFirstTape And ii = m_iRightFirstTape Then
                        If g_bLeftAboveKnee Then
                            nScale = 0.95 / 2 '5 reduction
                        Else
                            nScale = 0.92 / 2 '8 reduction
                        End If
                    ElseIf ii = m_iLeftFirstTape Or ii = m_iRightFirstTape Then
                        'Do nothing
                        'keep original legs tape values
                        GoTo Continue_Loop
                    Else
                        nScale = m_n20Len(ii) / 20
                    End If
                    nAverage = (m_nRightOriginalTapeValues(ii) + m_nLeftOriginalTapeValues(ii)) / 2
                    m_nLeftOriginalTapeValues(ii) = nAverage
                    m_nRightOriginalTapeValues(ii) = nAverage
                    g_LeftLegProfile.y(iLeftLegProfileCount) = (nAverage * nScale) + SEAM
                    g_RightLegProfile.y(iRightLegProfileCount) = (nAverage * nScale) + SEAM

                Else
                    'Do nothing
                    'keep original legs tape values
                    FN_AverageLeftAndRightLegs = False
                    If System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii)) > nDistanceApart Then
                        nDistanceApart = System.Math.Abs(m_nRightOriginalTapeValues(ii) - m_nLeftOriginalTapeValues(ii))
                        If iLeftLegProfileCount >= 1 Then
                            PR_MakeXY(xyTmp, g_LeftLegProfile.X(iLeftLegProfileCount), g_LeftLegProfile.y(iLeftLegProfileCount))
                            PR_MakeXY(xyTmp1, g_LeftLegProfile.X(iLeftLegProfileCount + 1), g_LeftLegProfile.y(iLeftLegProfileCount + 1))
                        ElseIf iLeftLegProfileCount = g_LeftLegProfile.n Then
                            PR_MakeXY(xyTmp, g_LeftLegProfile.X(iLeftLegProfileCount), g_LeftLegProfile.y(iLeftLegProfileCount))
                            PR_MakeXY(xyTmp1, g_LeftLegProfile.X(iLeftLegProfileCount - 1), g_LeftLegProfile.y(iLeftLegProfileCount - 1))
                        End If
                        PR_CalcPolar(xyTmp1, FN_CalcAngle(xyTmp1, xyTmp), FN_CalcLength(xyTmp1, xyTmp) * 2 / 3, xyBDLeftLegLabelPoint)

                        If iRightLegProfileCount >= 1 Then
                            PR_MakeXY(xyTmp, g_RightLegProfile.X(iRightLegProfileCount), g_RightLegProfile.y(iRightLegProfileCount))
                            PR_MakeXY(xyTmp1, g_RightLegProfile.X(iRightLegProfileCount + 1), g_RightLegProfile.y(iRightLegProfileCount + 1))
                        ElseIf iRightLegProfileCount = g_RightLegProfile.n Then
                            PR_MakeXY(xyTmp, g_RightLegProfile.X(iRightLegProfileCount), g_RightLegProfile.y(iRightLegProfileCount))
                            PR_MakeXY(xyTmp1, g_RightLegProfile.X(iRightLegProfileCount - 1), g_RightLegProfile.y(iRightLegProfileCount - 1))
                        End If
                        PR_CalcPolar(xyTmp1, FN_CalcAngle(xyTmp1, xyTmp), FN_CalcLength(xyTmp1, xyTmp) * 1 / 3, xyBDRightLegLabelPoint)

                    End If
                End If
            Else
                'Do nothing
                'keep original legs tape values
                FN_AverageLeftAndRightLegs = False

            End If
Continue_Loop:
        Next ii
        'Revise label points so that they sit inside the lisne
        PR_MakeXY(xyBDRightLegLabelPoint, xyBDRightLegLabelPoint.X - g_nRightLegLength, xyBDRightLegLabelPoint.Y - INCH1_16)
        PR_MakeXY(xyBDLeftLegLabelPoint, xyBDLeftLegLabelPoint.X - g_nLeftLegLength, xyBDLeftLegLabelPoint.Y - INCH1_16)

    End Function
    Sub PR_CopyRightToLeftLeg()
        Dim ii As Short
        g_LeftLegProfile = g_RightLegProfile
        g_nLeftLegLength = g_nRightLegLength
        m_iLeftFirstTape = m_iRightFirstTape
        m_iLeftLastTape = m_iRightLastTape
        xyBDLeftLegLabelPoint = xyBDRightLegLabelPoint

        For ii = 0 To 29
            m_nLeftOriginalTapeValues(ii) = m_nRightOriginalTapeValues(ii)
        Next ii
    End Sub
    Private Function PR_CalculateBodySuit() As Boolean
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
        Static xyTmpCutOut7 As BODYSUIT1.XY
        Static xyTmp1 As BODYSUIT1.XY
        Static xyTmp2 As BODYSUIT1.XY
        Static xyTmp3 As BODYSUIT1.XY
        Static bShldWidthModified As Short
        Static nTangentAngle As Double
        Static nGroinAngle As Double
        Static nThighAngle As Double
        Static nGroinToThighAngle As Double
        Static nGroinToThighMidPoint As Double
        Static sLeg As String
        Static sSleeve As String
        Static xyRegularTmpCutOut8 As BODYSUIT1.XY
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
        If g_bBodyDiffAxillaHeight Then
            'Draw the body using the highest axilla
            If g_nBodyLeftShldToAxilla < g_nBodyRightShldToAxilla Then
                g_nShldToAxilla = g_nBodyLeftShldToAxilla
                g_nShldToAxillaLow = g_nBodyRightShldToAxilla
                g_nAxillaCir = g_nBodyLeftAxillaCir
                g_nAxillaCirLow = g_nBodyRightAxillaCir
                g_sAxillaSide = "Left"
                g_sAxillaSideLow = "Right"
                g_sAxillaType = g_sBodyLeftSleeve
                g_sAxillaTypeLow = g_sBodyRightSleeve
            Else
                g_nShldToAxilla = g_nBodyRightShldToAxilla
                g_nShldToAxillaLow = g_nBodyLeftShldToAxilla
                g_nAxillaCir = g_nBodyRightAxillaCir
                g_nAxillaCirLow = g_nBodyLeftAxillaCir
                g_sAxillaSide = "Right"
                g_sAxillaSideLow = "Left"
                g_sAxillaType = g_sBodyRightSleeve
                g_sAxillaTypeLow = g_sBodyLeftSleeve
            End If
        Else
            'Default to left side
            g_sAxillaSide = "Left" 'Draw axilla by default on left side
            g_sAxillaType = g_sBodyLeftSleeve
            g_nShldToAxilla = g_nBodyLeftShldToAxilla
            g_nAxillaCir = g_nBodyLeftAxillaCir
        End If

        PR_MakeXY(xyBodySeamFold, 0, BODY_INCH1_4)
        PR_MakeXY(xyBodySeamButt, xyBodySeamFold.X + g_nButtocksLength, xyBodySeamFold.Y)
        PR_MakeXY(xyBodySeamHighShld, xyBodySeamFold.X + g_nShldToFold, xyBodySeamFold.Y)
        PR_MakeXY(xyBodySeamLowShld, xyBodySeamHighShld.X - BODY_INCH1_2, xyBodySeamFold.Y)
        PR_MakeXY(xyBodySeamChest, xyBodySeamHighShld.X - g_nShldToAxilla, xyBodySeamFold.Y)
        PR_MakeXY(xyBodySeamChestAxillaLow, xyBodySeamHighShld.X - g_nShldToAxillaLow, xyBodySeamFold.Y)
        PR_MakeXY(xyBodySeamWaist, xyBodySeamHighShld.X - g_nShldToWaist, xyBodySeamFold.Y)
        PR_MakeXY(xyBodyProfileButt, xyBodySeamButt.X, xyBodySeamButt.Y + g_nButtCirCalc)
        PR_MakeXY(xyBodyCutOut3, xyBodyProfileButt.X, xyBodyProfileButt.Y - g_nButtBackSeam)
        PR_MakeXY(xyBodyCutOut4, xyBodySeamFold.X + BODY_INCH3_4, xyBodyCutOut3.Y - (g_nButtCutOut / 2))
        PR_MakeXY(xyBodyCutOut5, xyBodySeamButt.X, xyBodySeamButt.Y + g_nButtFrontSeam)
        PR_MakeXY(xyBodyCutOut6, xyBodySeamWaist.X, xyBodySeamWaist.Y + g_nWaistFrontSeam)
        PR_MakeXY(xyBodyCutOut2, xyBodyCutOut6.X, xyBodyCutOut6.Y + g_nCutOut)
        PR_MakeXY(xyBodyProfileWaist, xyBodyCutOut2.X, xyBodyCutOut2.Y + g_nWaistBackSeam)
        PR_MakeXY(xyBodyCutOut7, xyBodySeamChest.X, xyBodySeamChest.Y + g_nChestFrontSeam)
        PR_MakeXY(xyBodyProfileChest, xyBodySeamChest.X, xyBodyCutOut2.Y + g_nChestBackSeam)

        '    g_nCutOutRadius = (FN_CalcLength(xyBodyCutOut4, xyBodyCutOut3) / 2) / Cos(FN_CalcAngle(xyBodyCutOut4, xyBodyCutOut3) * (PI / 180))
        'Original 85% mark

        If xyBodyCutOut5.X - xyBodyCutOut4.X > (xyBodyCutOut3.Y - xyBodyCutOut5.Y) / 2 Then
            'Extreme measurement case
            g_bExtremeCrotch = True
            g_nCutOutRadius = (FN_CalcLength(xyBodyCutOut5, xyBodyCutOut3) / 2)
            PR_MakeXY(xyBodyCutOut9, xyBodyCutOut4.X + g_nCutOutRadius, xyBodyCutOut3.Y)
            PR_MakeXY(xyBodyCutOut10, xyBodyCutOut9.X, xyBodyCutOut5.Y)
        Else
            'Normal case
            aAngle = FN_CalcAngle(xyBodyCutOut4, xyBodyCutOut3)
            If aAngle <> 90 Then
                g_nCutOutRadius = (FN_CalcLength(xyBodyCutOut4, xyBodyCutOut3) / 2) / System.Math.Cos(aAngle * (BODYSUIT1.PI / 180))
            Else
                g_nCutOutRadius = FN_CalcLength(xyBodyCutOut4, xyBodyCutOut3) / 2
            End If
            xyBodyCutOut9 = xyBodyCutOut3
            xyBodyCutOut10 = xyBodyCutOut5
        End If

        PR_MakeXY(xyBodyCutOutArcCentre, xyBodyCutOut4.X + g_nCutOutRadius, xyBodyCutOut4.Y)

        'Calculate Centre Point of Buttocks Arc for the averaged/smallest thigh
        PR_MakeXY(xyBDButtocksArcCentre, xyBodySeamButt.X, xyBodySeamButt.Y + g_nHalfGroinHeight)
        PR_MakeXY(xyBodyProfileGroin, xyBodySeamFold.X + BODY_INCH3_4, xyBodySeamFold.Y + g_nGroinHeight)

        'Calculate above for largest thigh (if given)
        '    If g_bBodyDiffThigh Then
        PR_MakeXY(xyBodyLT_ButtocksArcCentre, xyBodySeamButt.X, xyBodySeamButt.Y + g_nLT_HalfGroinHeight)
        PR_MakeXY(xyBodyLT_ProfileGroin, xyBodySeamFold.X + BODY_INCH3_4, xyBodySeamFold.Y + g_nLT_GroinHeight)
        '    End If

        'FRONT NECK
        'Needed variables for neck
        If g_bBodySleeveless Then
            nBackNeckOffset = fnRoundInches(g_nNeckCir / 6)
            g_nRadiusFrontNeck = fnRoundInches(((g_nNeckCir - (nBackNeckOffset * 2)) / 3.14) - BODY_INCH1_8) 'Full scale
        Else
            nBackNeckOffset = fnRoundInches(g_nNeckCir / 9)
            g_nRadiusFrontNeck = fnRoundInches(((g_nNeckCir - ((nBackNeckOffset * 2) + 1.5)) / 3.14) - BODY_INCH1_8) 'Full scale
        End If
        If InStr(1, g_sBodyFrontNeckStyle, "Scoop") > 0 Or g_sBodyFrontNeckStyle = "V neck" Then
            'Scoop, Measured scoop, V neck and Measured V neck
            PR_MakeXY(xyRegularTmpCutOut8, xyBodySeamHighShld.X - g_nRadiusFrontNeck, xyBodySeamHighShld.Y + g_nRadiusFrontNeck)
            g_nRadiusFrontNeck = g_nRadiusFrontNeck * 1.2
            g_nBodyShldWidth = g_nBodyShldWidth - 1
            bShldWidthModified = True
        End If
        PR_MakeXY(xyBodyTmpCutOut8, xyBodySeamHighShld.X - g_nRadiusFrontNeck, xyBodySeamHighShld.Y + g_nRadiusFrontNeck)

        PR_MakeXY(xyBDFrontNeckArcCentre, xyBodySeamHighShld.X, xyBodySeamHighShld.Y + g_nRadiusFrontNeck)

        Select Case g_sBodyFrontNeckStyle
            Case "Measured Scoop", "Measured V neck"
                n8_7Length = FN_CalcLength(xyRegularTmpCutOut8, xyBodyCutOut7)
                If (n8_7Length < g_nFrontNeckSize) Then
                    nScoopAngle = FN_CalcAngle(xyBodyCutOut7, xyBodyCutOut6)
                    PR_CalcPolar(xyBodyCutOut7, nScoopAngle, g_nFrontNeckSize - n8_7Length, xyBodyCutOut8)
                    g_bMissCutOut7 = True
                    'Get intersection on ordinary scoop neck line
                    xyTmp1 = xyBodyCutOut8
                    xyTmp1.Y = xyBodyCutOut8.Y + 10 'Extend line
                    ii = FN_LinLinInt(xyBodyCutOut6, xyBodyTmpCutOut8, xyBodyCutOut8, xyTmp1, xyBodyCutOut8)
                Else
                    nScoopAngle = FN_CalcAngle(xyRegularTmpCutOut8, xyBodyCutOut7)
                    PR_CalcPolar(xyRegularTmpCutOut8, nScoopAngle, g_nFrontNeckSize, xyBodyCutOut8)
                    xyTmp1 = xyBodyCutOut8
                    xyTmp1.Y = xyBodyCutOut8.Y + 10 'Extend line
                    ii = FN_LinLinInt(xyBodyCutOut7, xyBodyTmpCutOut8, xyBodyCutOut8, xyTmp1, xyBodyCutOut8)
                End If
                ''revise y of xyBodyCutOut8
                'PR_MakeXY xyBodyCutOut8, xyBodyCutOut8.x, xyBodyTmpCutOut8.y
            Case Else
                'Regular, V neck and Scoop
                xyBodyCutOut8 = xyBodyTmpCutOut8
        End Select

        'BACK NECK Check for scoop
        If InStr(g_sBodyBackNeckStyle, "Scoop") <> 0 Then
            If Not bShldWidthModified Then
                g_nBodyShldWidth = g_nBodyShldWidth - 1
            End If
            'Check for scoops here
            If g_sBodyBackNeckStyle = "Scoop" Then
                If g_nAdult Then
                    nBackScoopDistance = BODY_INCH3_4
                Else
                    nBackScoopDistance = INCH3_8
                End If
            Else
                nBackScoopDistance = g_nBackNeckSize
            End If
        End If
        nFrontAngle = -((360 / (2 * BODYSUIT1.PI * g_nRadiusFrontNeck)) + 90)
        If (g_strSide = "Left") Then
            sSleeve = g_sBodyLeftSleeve
        Else
            sSleeve = g_sBodyRightSleeve
        End If


        If sSleeve = "Sleeveless" Then
            'Neck Co-ordinate points and values
            PR_MakeXY(xyBodyCutOutBackNeck, xyBodySeamHighShld.X - (BODY_INCH1_4 + nBackScoopDistance), xyBodyCutOut2.Y)
            PR_MakeXY(xyBodyProfileNeck, xyBodySeamHighShld.X, xyBodyCutOut2.Y + nBackNeckOffset)
            xyBodyCutOutFrontNeck = xyBodySeamHighShld
            PR_MakeXY(xyBDProfileNeckMirror, xyBodyProfileNeck.X, -xyBodyProfileNeck.Y)

            'Axilla Co-ordinate points and values
            'Because it's sleeveless the axilla length is figured
            'Needed variables for Axilla
            nRaglanOffset = ((xyBodyCutOutFrontNeck.Y - xyBDProfileNeckMirror.Y) - BODY_INCH1_2) / 2
            'This variable is only used for Sleeveless
            nShldWidthOffset = System.Math.Sqrt((g_nBodyShldWidth ^ 2) - (BODY_INCH1_2 ^ 2))
            If (nRaglanOffset < nShldWidthOffset) Then
                nShldWidthOffset = nRaglanOffset
            End If
            PR_MakeXY(xyBodyRaglan1, xyBodySeamHighShld.X, xyBodySeamHighShld.Y - BODY_INCH1_4)
            PR_MakeXY(xyBodyRaglan2, xyBodySeamLowShld.X, xyBodyRaglan1.Y - nShldWidthOffset)
            PR_MakeXY(xyBodyRaglan3, xyBodySeamLowShld.X, xyBodyRaglan1.Y - nRaglanOffset)
            PR_MakeXY(xyBodyRaglan4, xyBodyRaglan3.X - (g_nAxillaCir / 2), xyBodyRaglan3.Y)
            If g_bBodyDiffAxillaHeight Then
                PR_MakeXY(xyBDRaglan4LowAxilla, xyBodyRaglan3.X - (g_nAxillaCirLow / 2), xyBodyRaglan3.Y)
            End If
            PR_MakeXY(xyBodyRaglan6, xyBDProfileNeckMirror.X, xyBDProfileNeckMirror.Y + BODY_INCH1_4)
            PR_MakeXY(xyBodyRaglan5, xyBodySeamLowShld.X, xyBodyRaglan6.Y + nShldWidthOffset)
        Else
            'Neck Co-ordinate points and values
            PR_MakeXY(xyBodyCutOutBackNeck, xyBodySeamLowShld.X - (BODY_INCH1_4 + nBackScoopDistance), xyBodyCutOut2.Y)
            PR_MakeXY(xyBodyProfileNeck, xyBodySeamLowShld.X, xyBodyCutOut2.Y + nBackNeckOffset)
            PR_CalcPolar(xyBDFrontNeckArcCentre, nFrontAngle, g_nRadiusFrontNeck, xyBodyCutOutFrontNeck)
            PR_MakeXY(xyBDProfileNeckMirror, xyBodyProfileNeck.X, -xyBodyProfileNeck.Y)

            'Axilla Co-ordinate points and values
            'Needed variables for Axilla
            nRaglanOffset = ((xyBodyCutOutFrontNeck.Y - xyBDProfileNeckMirror.Y) - BODY_INCH1_2) / 2
            PR_MakeXY(xyBodyRaglan1, xyBodyCutOutFrontNeck.X, xyBodyCutOutFrontNeck.Y - BODY_INCH1_4)
            PR_MakeXY(xyBodyRaglan2, xyBodySeamChest.X, xyBodySeamChest.Y - nRaglanOffset)
            PR_MakeXY(xyBodyRaglan3, xyBDProfileNeckMirror.X, xyBDProfileNeckMirror.Y + BODY_INCH1_4)
            If g_bBodyDiffAxillaHeight Then
                PR_MakeXY(xyBDRaglan2LowAxilla, xyBodySeamHighShld.X - g_nShldToAxillaLow, xyBodySeamChest.Y - nRaglanOffset)
                If g_sAxillaSide = "Left" Then
                    g_nAxillaBackNeckRad = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan3)
                    g_nAxillaFrontNeckRad = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan1)
                    g_nAFNRadRight = FN_CalcLength(xyBDRaglan2LowAxilla, xyBodyRaglan1)
                    g_nABNRadRight = FN_CalcLength(xyBDRaglan2LowAxilla, xyBodyRaglan3)
                Else
                    g_nAxillaBackNeckRad = FN_CalcLength(xyBDRaglan2LowAxilla, xyBodyRaglan3)
                    g_nAxillaFrontNeckRad = FN_CalcLength(xyBDRaglan2LowAxilla, xyBodyRaglan1)
                    g_nAFNRadRight = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan1)
                    g_nABNRadRight = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan3)
                End If
            Else
                g_nAxillaBackNeckRad = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan3)
                g_nAxillaFrontNeckRad = FN_CalcLength(xyBodyRaglan2, xyBodyRaglan1)
                g_nABNRadRight = 0
                g_nAFNRadRight = 0
            End If
        End If


        If g_bBodyDiffThigh Then 'One panty & brief
            'Release extra 5% reduction
            g_nThighCir = fnRoundInches(g_nBodySmallestThighGiven * (g_nThighCirRed + 0.05)) / 2 'half scale
        Else
            g_nThighCir = fnRoundInches(g_nBodySmallestThighGiven * g_nThighCirRed) / 2 'half scale
        End If

        'Calculate Crotch Co-ordinate points and values
        Select Case g_sBodyCrotchStyle
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
                    nCutArcLength = (2 * BODYSUIT1.PI * g_nCutOutRadius) / 4
                    nCutExtendedLength = FN_CalcLength(xyBodyCutOut10, xyBodyCutOut5)
                    'Back crotch
                    If nCutArcLength > g_nBackCrotchSize Then
                        'Open crotch ends on arc
                        xyBodyBackCrotch1 = xyBodyCutOut9
                        PR_CalcPolar(xyBodyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * BODYSUIT1.PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBodyBackCrotch1)
                        xyTmp1.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                        xyTmp1.Y = xyBodyBackCrotch1.Y - nCrotchFilletRadius
                        xyTmp2.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                        xyTmp2.Y = xyBodyBackCrotch1.Y + 1
                        iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDBackCrotchFilletCentre)
                        PR_MakeXY(xyBodyBackCrotch2, xyBDBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBDBackCrotchFilletCentre.Y)
                        PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                        iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyBackCrotch3)
                    ElseIf (nCutArcLength + nCutExtendedLength) < g_nBackCrotchSize Then
                        'Open crotch ends at xyBodyCutOut3
                        'IE we can't go past the Buttock height
                        xyBodyBackCrotch1 = xyBodyCutOut3
                        PR_MakeXY(xyBodyBackCrotch2, xyBodyBackCrotch1.X, xyBodyBackCrotch1.Y + nCrotchFilletRadius)
                        PR_MakeXY(xyBDBackCrotchFilletCentre, xyBodyBackCrotch1.X - nCrotchFilletRadius, xyBodyBackCrotch2.Y)
                        PR_MakeXY(xyBodyBackCrotch3, xyBDBackCrotchFilletCentre.X, xyBDBackCrotchFilletCentre.Y + nCrotchFilletRadius)
                    Else
                        'Open crotch ends between xyBodyCutOut9 and xyCutout3
                        xyBodyBackCrotch1 = xyBodyCutOut3
                        xyBodyBackCrotch1.X = xyBodyCutOut3.X - ((nCutArcLength + nCutExtendedLength) - g_nBackCrotchSize)
                        'Get center of fillet radius(initially relative to the arc)
                        If (xyBodyBackCrotch1.X - nCrotchFilletRadius) < xyBodyCutOut9.X Then
                            'Fillet intersects arc
                            xyTmp1.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                            xyTmp1.Y = xyBodyBackCrotch1.Y - nCrotchFilletRadius
                            xyTmp2.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                            xyTmp2.Y = xyBodyBackCrotch1.Y + 1
                            iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDBackCrotchFilletCentre)
                            PR_MakeXY(xyBodyBackCrotch2, xyBDBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBDBackCrotchFilletCentre.Y)
                            PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                            iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyBackCrotch3)
                        Else
                            'Fillet is on straight bit
                            PR_MakeXY(xyBodyBackCrotch2, xyBodyBackCrotch1.X, xyBodyBackCrotch1.Y + nCrotchFilletRadius)
                            PR_MakeXY(xyBDBackCrotchFilletCentre, xyBodyBackCrotch1.X - nCrotchFilletRadius, xyBodyBackCrotch2.Y)
                            PR_MakeXY(xyBodyBackCrotch3, xyBDBackCrotchFilletCentre.X, xyBDBackCrotchFilletCentre.Y + nCrotchFilletRadius)
                        End If
                    End If

                    'Front crotch
                    If nCutArcLength > g_nFrontCrotchSize Then
                        'Open crotch ends on arc
                        xyBodyFrontCrotch1 = xyBodyCutOut10
                        xyTmp1.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                        xyTmp1.Y = xyBodyFrontCrotch1.Y + nCrotchFilletRadius
                        xyTmp2.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                        xyTmp2.Y = xyBodyFrontCrotch1.Y - 1
                        iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDFrontCrotchFilletCentre)
                        PR_MakeXY(xyBodyFrontCrotch2, xyBDFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyBDFrontCrotchFilletCentre.Y)
                        PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                        iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyFrontCrotch3)
                    ElseIf (nCutArcLength + nCutExtendedLength) < g_nFrontCrotchSize Then
                        'Open crotch ends at xyBodyCutOut5
                        'IE we can't go past the Buttock height
                        xyBodyFrontCrotch1 = xyBodyCutOut5
                        PR_MakeXY(xyBodyFrontCrotch2, xyBodyFrontCrotch1.X, xyBodyFrontCrotch1.Y - nCrotchFilletRadius)
                        PR_MakeXY(xyBDFrontCrotchFilletCentre, xyBodyFrontCrotch1.X - nCrotchFilletRadius, xyBodyFrontCrotch2.Y)
                        PR_MakeXY(xyBodyFrontCrotch3, xyBDFrontCrotchFilletCentre.X, xyBDFrontCrotchFilletCentre.Y - nCrotchFilletRadius)
                    Else
                        'Open crotch ends between xyBodyCutOut10 and xyCutout5
                        xyBodyFrontCrotch1 = xyBodyCutOut5
                        xyBodyFrontCrotch1.X = xyBodyCutOut5.X - ((nCutArcLength + nCutExtendedLength) - g_nFrontCrotchSize)
                        'Get center of fillet radius(initially relative to the arc)
                        If (xyBodyFrontCrotch1.X - nCrotchFilletRadius) < xyBodyCutOut10.X Then
                            'Fillet intersects arc
                            xyTmp1.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                            xyTmp1.Y = xyBodyFrontCrotch1.Y + nCrotchFilletRadius
                            xyTmp2.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                            xyTmp2.Y = xyBodyFrontCrotch1.Y - 1
                            iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDFrontCrotchFilletCentre)
                            PR_MakeXY(xyBodyFrontCrotch2, xyBDFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyBDFrontCrotchFilletCentre.Y)
                            PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                            iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyFrontCrotch3)
                        Else
                            'Fillet is on straight bit
                            PR_MakeXY(xyBodyFrontCrotch2, xyBodyFrontCrotch1.X, xyBodyFrontCrotch1.Y - nCrotchFilletRadius)
                            PR_MakeXY(xyBDFrontCrotchFilletCentre, xyBodyFrontCrotch1.X - nCrotchFilletRadius, xyBodyFrontCrotch2.Y)
                            PR_MakeXY(xyBodyFrontCrotch3, xyBDFrontCrotchFilletCentre.X, xyBDFrontCrotchFilletCentre.Y - nCrotchFilletRadius)
                        End If
                    End If

                Else
                    'Normal crotch
                    nCutSegmentAngle = FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut5) - FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut4)
                    nCutArcLength = fnRoundInches((2 * BODYSUIT1.PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                    If (nCutArcLength <= g_nBackCrotchSize) Then
                        xyBodyBackCrotch1 = xyBodyCutOut3
                    Else
                        PR_CalcPolar(xyBodyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * BODYSUIT1.PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBodyBackCrotch1)
                    End If
                    'Back of crotch
                    xyTmp1.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                    xyTmp1.Y = xyBodyBackCrotch1.Y - nCrotchFilletRadius
                    xyTmp2.X = xyBodyBackCrotch1.X - nCrotchFilletRadius
                    xyTmp2.Y = xyBodyBackCrotch1.Y + 1
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDBackCrotchFilletCentre)
                    PR_MakeXY(xyBodyBackCrotch2, xyBDBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBDBackCrotchFilletCentre.Y)
                    PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                    iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyBackCrotch3)
                    'Front of crotch
                    If (nCutArcLength <= g_nFrontCrotchSize) Then
                        xyBodyFrontCrotch1 = xyBodyCutOut5
                    Else
                        PR_CalcPolar(xyBodyCutOutArcCentre, 180 + ((g_nFrontCrotchSize / (2 * BODYSUIT1.PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBodyFrontCrotch1)
                    End If
                    xyTmp1.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                    xyTmp1.Y = xyBodyFrontCrotch1.Y + nCrotchFilletRadius
                    xyTmp2.X = xyBodyFrontCrotch1.X - nCrotchFilletRadius
                    xyTmp2.Y = xyBodyFrontCrotch1.Y - 1
                    nTmpRadius = g_nCutOutRadius + nCrotchFilletRadius
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyCutOutArcCentre, nTmpRadius, xyBDFrontCrotchFilletCentre)
                    PR_MakeXY(xyBodyFrontCrotch2, xyBDFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyBDFrontCrotchFilletCentre.Y)
                    PR_CalcPolar(xyBodyCutOutArcCentre, FN_CalcAngle(xyBodyCutOutArcCentre, xyBDFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3)
                    iError = FN_CirLinInt(xyBodyCutOutArcCentre, xyTmp3, xyBodyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBodyFrontCrotch3)
                End If
                'Common
                PR_MakeXY(xyBodyLegPoint, BODY_INCH3_4, 0)
                PR_MakeXY(xyBodyCrotchMarker, xyBodyCutOut4.X - (2 * nCrotchFilletRadius), xyBodyCutOut4.Y)
                If g_sSex = "Male" Then
                    '1 1/4" + 3/4" to xyBodyCutOut4 = 2"
                    PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - (1 + BODY_INCH1_4), xyBodySeamFold.Y)
                    If Not g_nAdult Then
                        '7/8" + 3/4" to xyBodyCutOut4 = 1 5/8"
                        PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - INCH7_8, xyBodySeamFold.Y)
                    End If
                Else
                    '7/8" + 3/4" to xyBodyCutOut4 = 1 5/8"
                    PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - INCH7_8, xyBodySeamFold.Y)
                End If
                'Test that at least 1-1/4" between cutout edge and bottom of brief.
                If (xyBodyCrotchMarker.X - xyBodySeamThigh.X) < 1.25 Then 'less than 1-1/4"
                    PR_MakeXY(xyBodySeamThigh, xyBodyCrotchMarker.X - 1.25, xyBodySeamFold.Y)
                End If
                PR_MakeXY(xyBodyProfileBrief, xyBodySeamThigh.X, xyBodySeamThigh.Y + (g_nHalfGroinHeight / 2))

            Case "Horizontal Fly"
                '7/8" + 3/4" to xyBodyCutOut4 = 1 5/8"
                PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - INCH7_8, xyBodySeamFold.Y)
                PR_MakeXY(xyBodyProfileBrief, xyBodySeamThigh.X, xyBodySeamThigh.Y + (g_nHalfGroinHeight / 2))
                PR_SetFlySize("Hor")
                PR_MakeXY(xyBodyLegPoint, 0, 0)
                PR_MakeXY(xyBodyCrotchMarker, xyBodyCutOut4.X, xyBodyCutOut4.Y)
            Case "Diagonal Fly"
                '7/8" + 3/4" to xyBodyCutOut4 = 1 5/8"
                PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - INCH7_8, xyBodySeamFold.Y)
                PR_MakeXY(xyBodyProfileBrief, xyBodySeamThigh.X, xyBodySeamThigh.Y + (g_nHalfGroinHeight / 2))
                PR_SetFlySize("Dia")
                PR_MakeXY(xyBodyLegPoint, 0, 0)
                PR_MakeXY(xyBodyCrotchMarker, xyBodyCutOut4.X, xyBodyCutOut4.Y)

            Case "Gusset"
                '7/8" + 3/4" to xyBodyCutOut4 = 1 5/8"
                PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - INCH7_8, xyBodySeamFold.Y)
                PR_MakeXY(xyBodyProfileBrief, xyBodySeamThigh.X, xyBodySeamThigh.Y + (g_nHalfGroinHeight / 2))
                PR_SetGussetSize()
                PR_MakeXY(xyBodyLegPoint, 0, 0)
                PR_MakeXY(xyBodyCrotchMarker, xyBodyCutOut4.X, xyBodyCutOut4.Y)

            Case "Snap Crotch"
                '1/2" + 3/4" to xyBodyCutOut4 = 1 1/4"
                PR_MakeXY(xyBodySeamThigh, xyBodySeamFold.X - BODY_INCH1_2, xyBodySeamFold.Y)
                PR_MakeXY(xyBodyProfileBrief, xyBodySeamThigh.X, xyBodyCutOut4.Y)
                PR_SetGussetSize()
                nGussetLength = g_nGussetLength + 1.5 '1-1/2"
                If nGussetLength < 2.5 Then '2-1/2" Minimum
                    MsgBox(">>>>>>>>>>>>>>>>>WARNING <<<<<<<<<<<<" & Chr(13) & "Gusset Strap shorter than 2-1/2"" which is unacceptable.")
                    ' End
                End If
                g_nLengthStrap1ToCutOut5 = 0
                If g_bExtremeCrotch Then

                    'Special case for extreme crotch
                    nCutSegmentAngle = FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut10) - FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut4)
                    nCutArcLength = fnRoundInches((2 * BODYSUIT1.PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                    nCutExtendedLength = FN_CalcLength(xyBodyCutOut10, xyBodyCutOut5)

                    If nCutArcLength > nGussetLength Then
                        'End of strap ends on arc
                        nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / BODYSUIT1.PI))
                        PR_CalcPolar(xyBodyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyBodyStrap1)
                        g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength

                    ElseIf (nCutArcLength + nCutExtendedLength) > nGussetLength Then
                        'End of strap ends between xyBodyCutOut10 and xyCutout5
                        nStrapAngle = FN_CalcAngle(xyBodyCutOut10, xyBodyCutOut5)
                        nStrapLength = nGussetLength - nCutArcLength
                        PR_CalcPolar(xyBodyCutOut10, nStrapAngle, nStrapLength, xyBodyStrap1)
                        g_nLengthStrap1ToCutOut5 = FN_CalcLength(xyBodyStrap1, xyBodyCutOut5)
                    Else
                        'End of strap ends between xyBodyCutOut5 and xyCutout6
                        nStrapAngle = FN_CalcAngle(xyBodyCutOut5, xyBodyCutOut6)
                        nStrapLength = nGussetLength - (nCutArcLength + nCutExtendedLength)
                        PR_CalcPolar(xyBodyCutOut5, nStrapAngle, nStrapLength, xyBodyStrap1)
                        g_nLengthStrap1ToCutOut5 = 0
                    End If
                Else
                    'Normal crotch
                    nCutSegmentAngle = FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut5) - FN_CalcAngle(xyBodyCutOutArcCentre, xyBodyCutOut4)
                    nCutArcLength = fnRoundInches((2 * BODYSUIT1.PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                    If (nCutArcLength < nGussetLength) Then
                        nStrapAngle = FN_CalcAngle(xyBodyCutOut5, xyBodyCutOut6)
                        nStrapLength = nGussetLength - nCutArcLength
                        PR_CalcPolar(xyBodyCutOut5, nStrapAngle, nStrapLength, xyBodyStrap1)
                        g_nLengthStrap1ToCutOut5 = 0
                    Else
                        nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / BODYSUIT1.PI))
                        PR_CalcPolar(xyBodyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyBodyStrap1)
                        g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength
                    End If
                End If

                PR_MakeXY(xyBodyStrap2, xyBodyStrap1.X, xyBodyStrap1.Y - 1.75) '1-3/4"
                nGussetRadius = xyBodyCutOutArcCentre.Y - xyBodyStrap2.Y
                PR_MakeXY(xyBodyGussetArcCentre, xyBodyProfileBrief.X + nGussetRadius, xyBodyProfileBrief.Y)

                If xyBodyStrap1.X > xyBodyGussetArcCentre.X Then
                    'Normal end of Snap crotch arc
                    PR_MakeXY(xyBodyStrap3, xyBodyGussetArcCentre.X, xyBodyStrap2.Y) 'Roughly!!!
                Else
                    'End of arc
                    PR_MakeXY(xyBodyStrap3, xyBodyStrap1.X - BODY_INCH3_4, xyBodyStrap2.Y)
                    'Revise xyBodyGussetArcCentre
                    '                aAngle = Fn_CalcAngle(xyBodyStrap3, xyBodyProfileBrief) - 90
                    '                nLength = FN_CalcLength(xyBodyStrap3, xyBodyProfileBrief) / 2
                    '                Pr_CalcMidPoint xyBodyStrap3, xyBodyProfileBrief, xyTmp1
                    '                nLength = Sqr(nGussetRadius ^ 2 - nLength ^ 2)
                    '                PR_CalcPolar xyTmp1, aAngle, nLength, xyBodyGussetArcCentre
                    aAngle = 180 - FN_CalcAngle(xyBodyStrap3, xyBodyProfileBrief)
                    nLength = FN_CalcLength(xyBodyStrap3, xyBodyProfileBrief) / 2
                    nGussetRadius = nLength / System.Math.Cos(aAngle * (BODYSUIT1.PI / 180))
                    PR_MakeXY(xyBodyGussetArcCentre, xyBodyProfileBrief.X + nGussetRadius, xyBodyProfileBrief.Y)
                End If
                PR_MakeXY(xyBodyStrap4, BODY_INCH3_4, xyBodyStrap2.Y)
                PR_MakeXY(xyBodyLegPoint, BODY_INCH3_4, 0)
                PR_MakeXY(xyBodyCrotchMarker, xyBodyCutOut4.X, xyBodyCutOut4.Y)

        End Select

        'Panty
        'NB We check at this point for complete leg data we can't do this in the
        'in the main validation routine as we do not have leg data
        'at that point
        If g_sLeftLeg = "Panty" Or g_sRightLeg = "Panty" Then

            If Not FN_ValidateLegData() Then
                Return False
            End If
            If Not FN_InitiliseLegRoutines() Then
                Return False
            End If

            If g_sLeftLeg = "Panty" Then PR_CalculateLeg("Left", g_nLeftLegLength)
            If g_sRightLeg = "Panty" Then PR_CalculateLeg("Right", g_nRightLegLength)
            If g_sLeftLeg = "Panty" And g_sRightLeg = "Panty" Then g_bDrawSingleLeg = FN_AverageLeftAndRightLegs()
            If g_sLeftLeg <> "Panty" And g_sRightLeg = "Panty" Then
                g_bDrawSingleLeg = True
                PR_CopyRightToLeftLeg()
                xyBDLeftLegLabelPoint.X = xyBDLeftLegLabelPoint.X - g_nLeftLegLength 'This normally a side efect of averaging the legs'Bad practice but what the hell
            End If
            If g_sLeftLeg = "Panty" And g_sRightLeg <> "Panty" Then
                g_bDrawSingleLeg = True
                xyBDLeftLegLabelPoint.X = xyBDLeftLegLabelPoint.X - g_nLeftLegLength 'As above
            End If

            PR_MakeXY(xyBodyLT_Fold, xyBodySeamFold.X, xyBodySeamFold.Y + g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
            PR_MakeXY(xyBodyFold, xyBodySeamFold.X, xyBodySeamFold.Y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2)))

        End If

        If g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief" Then
            'Check to see if it can reach, if not use the tangent from the groin
            'point and join them with a straight line.
            'Calculate points for Average / Smallest thigh
            nGroinAngle = FN_CalcAngle(xyBDButtocksArcCentre, xyBodyProfileGroin)
            If (xyBDButtocksArcCentre.X - xyBodySeamThigh.X) < g_nButtRadius Then
                'Curve reaches bottom of Brief
                g_bDrawBriefCurve = True
                PR_MakeXY(xyTmp1, xyBodySeamThigh.X, xyBDButtocksArcCentre.Y)
                PR_MakeXY(xyTmp2, xyTmp1.X, xyBodyProfileButt.Y)
                iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBDButtocksArcCentre, g_nButtRadius, xyBodyProfileThigh)
                nThighAngle = FN_CalcAngle(xyBDButtocksArcCentre, xyBodyProfileThigh)
                PR_CalcPolar(xyBDButtocksArcCentre, nThighAngle - ((nThighAngle - nGroinAngle) / 3), g_nButtRadius, xyBDProfileThighExtraPT)
            Else
                'Curve doesn't reaches bottom of Brief - draw tangential line
                g_bDrawBriefCurve = False
                nTangentAngle = nGroinAngle + 90
                PR_MakeXY(xyTmp1, xyBodyProfileBrief.X, xyBodyProfileButt.Y)
                PR_CalcPolar(xyBodyProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2)
                iError = FN_LinLinInt(xyBodyProfileBrief, xyTmp1, xyBodyProfileGroin, xyTmp2, xyBodyProfileThigh)
                nGroinToThighAngle = FN_CalcAngle(xyBodyProfileGroin, xyBodyProfileThigh)
                nGroinToThighMidPoint = (FN_CalcLength(xyBodyProfileGroin, xyBodyProfileThigh) / 2)
                PR_CalcPolar(xyBodyProfileGroin, nGroinToThighAngle, nGroinToThighMidPoint, xyBDProfileThighExtraPT)
            End If

            'Calculate points for largest thigh (if required)
            'I know parts of this is identical to above except for a few variable changes but I
            'am trying to retrofit this code without making major changes to code that has
            'already been validated. (Yeah! I know it sound feeble, but this is not paying me enough to
            'really care) GG
            If g_bBodyDiffThigh Then
                nGroinAngle = FN_CalcAngle(xyBodyLT_ButtocksArcCentre, xyBodyLT_ProfileGroin)
                If (xyBodyLT_ButtocksArcCentre.X - xyBodySeamThigh.X) < g_nLT_ButtRadius Then
                    'Curve reaches bottom of Brief
                    PR_MakeXY(xyTmp1, xyBodySeamThigh.X, xyBodyLT_ButtocksArcCentre.Y)
                    PR_MakeXY(xyTmp2, xyTmp1.X, xyBodyProfileButt.Y + 2) 'Ensure it clears
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyBodyLT_ButtocksArcCentre, g_nLT_ButtRadius, xyBodyLT_ProfileThigh)
                Else
                    'Curve doesn't reaches bottom of Brief - draw tangential line
                    nTangentAngle = nGroinAngle + 90
                    PR_MakeXY(xyTmp1, xyBodyProfileBrief.X, xyBodyProfileButt.Y + 2)
                    PR_CalcPolar(xyBodyLT_ProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2)
                    iError = FN_LinLinInt(xyBodyProfileBrief, xyTmp1, xyBodyLT_ProfileGroin, xyTmp2, xyBodyLT_ProfileThigh)
                End If
                PR_MakeXY(xyBodyLT_Fold, xyBodySeamFold.X, xyBodySeamFold.Y + g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
                'Revise xyBodyLT_ButtocksArcCentre so the point can be used construct the back arc
                'Note that we are trying to blend from the largest thigh into the body based on
                'the smallest thigh
                aAngle = FN_CalcAngle(xyBodyLT_Fold, xyBodyProfileButt) - 90 'perpendicular angle to the line xyBodyLT_Fold, xyBodyProfileButt
                nLength = System.Math.Sqrt(g_nButtRadius ^ 2 - (FN_CalcLength(xyBodyLT_Fold, xyBodyProfileButt) / 2) ^ 2)
                PR_CalcMidPoint(xyBodyLT_Fold, xyBodyProfileButt, xyTmp1)
                PR_CalcPolar(xyTmp1, aAngle, nLength, xyBodyLT_ButtocksArcCentre)
            End If
        End If

        'Average/Smallest thigh
        PR_MakeXY(xyBodyFold, xyBodySeamFold.X, xyBodySeamFold.Y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2)))
        PR_MakeXY(xyBody85Fold, xyBodySeamFold.X, xyBodySeamFold.Y + g_n85FoldHeight)
        PR_MakeXY(xyBDProfileThighMirror, xyBodyProfileThigh.X, -xyBodyProfileThigh.Y)
        ' If g_bBodyDiffThigh Then PR_MakeXY xyBodyLT_ProfileThighMirror, xyBodyLT_ProfileThigh.x, -xyBodyLT_ProfileThigh.y
        PR_MakeXY(xyBodyLT_ProfileThighMirror, xyBodyLT_ProfileThigh.X, -xyBodyLT_ProfileThigh.Y)
        PR_MakeXY(xyBDProfileThighMirror1, xyBDProfileThighMirror.X, xyBDProfileThighMirror.Y + BODY_INCH1_4)

        'largest thigh (if required)
        Return True
    End Function
    Private Sub PR_DrawVerticalLineThruPoint(ByRef xyPnt As BODYSUIT1.XY, ByRef nHeight As Double)
        'Function to reduce code in PR_DrawMacro
        Static xy1, xy2 As BODYSUIT1.XY
        PR_MakeXY(xy1, xyPnt.X, xyPnt.Y - nHeight)
        PR_MakeXY(xy2, xyPnt.X, xyPnt.Y + nHeight)
        PR_DrawLine(xy1, xy2)
    End Sub
    Sub PR_DrawLine(ByRef xyStart As BODYSUIT1.XY, ByRef xyFinish As BODYSUIT1.XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
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
            Dim acLine As Line = New Line(New Point3d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y, 0),
                                    New Point3d(xyFinish.X + xyInsertion.X, xyFinish.Y + xyInsertion.Y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            idLastCreated = acLine.ObjectId()

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_AddDBValueToLast(ByRef sDBName As String, ByRef sDBValue As String)
        'The last entity is given by hEnt
        PrintLine(BODYSUIT1.fNum, "if (hEnt) SetDBData( hEnt," & BODYSUIT1.QQ & sDBName & BODYSUIT1.QQ & BODYSUIT1.CC & BODYSUIT1.QQ & sDBValue & BODYSUIT1.QQ & ");")
        If (idLastCreated.IsValid = False) Then
            Exit Sub
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As DBObject = acTrans.GetObject(idLastCreated, OpenMode.ForWrite)
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
    Sub PR_DrawMarker(ByRef xyPoint As BODYSUIT1.XY)
        'Draw a Marker at the given point
        PrintLine(BODYSUIT1.fNum, "hEnt = AddEntity(" & BODYSUIT1.QQ & "marker" & BODYSUIT1.QCQ & "xmarker" & BODYSUIT1.QC & "xyStart.x+" & xyPoint.X & BODYSUIT1.CC & "xyStart.y+" & xyPoint.Y & BODYSUIT1.CC & "0.125);")
        PrintLine(BODYSUIT1.fNum, "SetDBData(hEnt," & BODYSUIT1.QQ & "ID" & BODYSUIT1.QQ & ",sID);")

        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As BODYSUIT1.XY
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
                Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.y, 0), New Point3d(xyEnd.X, xyEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xySecSt.X, xySecSt.y, 0), New Point3d(xySecEnd.X, xySecEnd.y, 0))
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
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyInsertion.X, xyPoint.Y + xyInsertion.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                idLastCreated = blkRef.ObjectId()
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawArc(ByRef xyCen As BODYSUIT1.XY, ByRef xyArcStart As BODYSUIT1.XY, ByRef xyArcEnd As BODYSUIT1.XY)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
        nRad = FN_CalcLength(xyCen, xyArcStart)
        nStartAng = FN_CalcAngle(xyCen, xyArcStart)
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd)

        nDeltaAng = nEndAng - nStartAng
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
            '' Create an arc that is at 6.25,9.125 with a radius of 6, and
            '' starts at 64 degrees and ends at 204 degrees
            Using acArc As Arc = New Arc(New Point3d(xyCen.X + xyInsertion.X, xyCen.Y + xyInsertion.Y, 0),
                                     nRad, (nStartAng * (BODYSUIT1.PI / 180)), (nEndAng * (BODYSUIT1.PI / 180)))
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As BODYSUIT1.XY, ByRef nHeight As Object, ByRef nAngle As Double, ByVal nTextMode As Double)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw TEXT at the given height.
        '
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '
        'Note:-
        '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
        '
        '
        Dim nWidth As Object
        nWidth = nHeight * BODYSUIT1.g_nCurrTextAspect

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
                acText.Position = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
                acText.Height = nHeight
                acText.Rotation = nAngle * (BODYSUIT1.PI / 180)
                acText.TextString = sText
                acText.Justify = nTextMode
                acText.AlignmentPoint = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using
            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
        ' PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC & nWidth & CC & nHeight & CC & g_nCurrTextAngle & ");")
    End Sub
    Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As BODYSUIT1.XY, ByRef dAngle As Double, Optional ByRef bIsCenter As Boolean = False)
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
            mtx.Location = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = dAngle * (BODYSUIT1.PI / 180)
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            If bIsCenter = True Then
                mtx.Attachment = AttachmentPoint.MiddleCenter
            End If
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
    Sub PR_InsertSymbol(ByRef sSymbol As String, ByRef xyInsert As BODYSUIT1.XY, ByRef nScale As Single, ByRef nRotation As Single)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to insert a SYMBOL.
        'Where:-
        '    sSymbol     Symbol name, must exist and be in the symbol library
        '    xyInsert    The insertion point
        '    nScale      Symbol scaling factor, 1 = No scaling
        '    nRotation   Symbol rotation about insertion point
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        'and
        '    The DRAFIX symbol library "C:\JOBST\JOBST.SLB" exists
        '
        'Note:-
        '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
        '
        PR_SetLayer("Notes")
        PrintLine(BODYSUIT1.fNum, "SetSymbolLibrary(sPathJOBST +" & BODYSUIT1.QQ & "\\JOBST.SLB" & BODYSUIT1.QQ & ");")
        PrintLine(BODYSUIT1.fNum, "Symbol(" & BODYSUIT1.QQ & "find" & BODYSUIT1.QCQ & sSymbol & BODYSUIT1.QQ & ");")

        'PR_SetLayer slayer
        PrintLine(BODYSUIT1.fNum, "hEnt = AddEntity(" & BODYSUIT1.QQ & "symbol" & BODYSUIT1.QCQ & sSymbol & BODYSUIT1.QC)
        PrintLine(BODYSUIT1.fNum, "xyStart.x+" & Str(xyInsert.X) & BODYSUIT1.CC & "xyStart.y+" & Str(xyInsert.Y) & BODYSUIT1.CC)
        PrintLine(BODYSUIT1.fNum, Str(nScale) & BODYSUIT1.CC & Str(nScale) & BODYSUIT1.CC & Str(nRotation) & ");")
    End Sub
    Sub PR_AddFormatedDBValueToLast(ByRef sDBName As String, ByRef sLeadingText As String, ByRef nValue As Double, ByRef sTrailingText As String)
        'The last entity is given by hEnt
        ' Format(sDataType, nValue)
        PrintLine(BODYSUIT1.fNum, "if (hEnt) {")
        PrintLine(BODYSUIT1.fNum, "SetDBData( hEnt," & BODYSUIT1.QQ & sDBName & BODYSUIT1.QCQ & sLeadingText & BODYSUIT1.QQ & "+ Format(" & BODYSUIT1.QQ & "length" & BODYSUIT1.QC & nValue & ")+" & BODYSUIT1.QQ & sTrailingText & BODYSUIT1.QQ & ");")
        PrintLine(BODYSUIT1.fNum, "}")
    End Sub
    Sub PR_DrawFitted(ByRef Profile As curve)
        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to draw a FITTED curve through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        Dim ii As Short
        'Return '11-June-2018
        'Draw the profile
        '    If there is no vertex or only one vertex then exit.
        '    For two vertex draw as a polyline (this degenerates to a single line).
        '    For three vertex draw as a polyline (as no fitted curve can be drawn
        '    by a macro).
        Select Case Profile.n
            Case 0 To 1
                Exit Sub
            Case 3
                PR_DrawPoly(Profile)
                'Warn the user to smooth the curve
                'PrintLine(fNum, "Display (" & QQ & "message" & QCQ & "OKquestion" & QCQ & "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line" & QQ & ");")
                ''MsgBox("The Profile has been drawn as a POLYLINE" & Chr(10) & "Edit this line and make it OPEN FITTED," & Chr(10) & " this will then be a smooth line", 16, "BodySuit - Dialog")
            Case Else
                'PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
                'For ii = 1 To Profile.n
                '    PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
                'Next ii
                'PrintLine(fNum, ");")

                ''Dim strLayout As String = BlockTableRecord.ModelSpace
                ''If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
                ''    strLayout = BlockTableRecord.PaperSpace
                ''End If

                ''Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                ''Dim acCurDb As Database = acDoc.Database

                '' Start a transaction
                ''Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                ''    '' Open the Block table for read
                ''    Dim acBlkTbl As BlockTable
                ''    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                ''    '' Open the Block table record Model space for write
                ''    Dim acBlkTblRec As BlockTableRecord
                ''    acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                ''                                    OpenMode.ForWrite)

                ''    '' Create a polyline with two segments (3 points)
                ''    Using acPoly As Polyline = New Polyline()
                ''        For ii = 1 To Profile.n
                ''            acPoly.AddVertexAt(ii - 1, New Point2d(Profile.X(ii) + xyInsertion.X, Profile.y(ii) + xyInsertion.Y), 0, 0, 0)
                ''        Next ii
                ''        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                ''            acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                ''        End If

                ''        '' Add the new object to the block table record and the transaction
                ''        acBlkTblRec.AppendEntity(acPoly)
                ''        acTrans.AddNewlyCreatedDBObject(acPoly, True)
                ''    End Using

                ''    '' Save the new object to the database
                ''    acTrans.Commit()
                ''End Using
                Dim ptColl As Point3dCollection = New Point3dCollection()
                For ii = 1 To Profile.n
                    ptColl.Add(New Point3d(Profile.X(ii) + xyInsertion.X, Profile.y(ii) + xyInsertion.Y, 0))
                Next
                PR_DrawSpline(ptColl)
        End Select
    End Sub
    'To Draw Spline
    Private Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Get the current document and database
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
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                Dim appName As String = "ProfileID"
                Dim xdataStr As String = txtFileNo.Text & "Left"
                If acRegAppTbl.Has(appName) = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                    acSpline.XData = rb
                End Using
                idLastCreated = acSpline.ObjectId()
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawPoly(ByRef Profile As curve)
        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to draw a POLYLINE through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        Dim ii As Short

        'Exit if nothing to draw
        If Profile.n <= 1 Then Exit Sub
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
                For ii = 1 To Profile.n
                    acPoly.AddVertexAt(ii - 1, New Point2d(Profile.X(ii) + xyInsertion.X, Profile.y(ii) + xyInsertion.Y), 0, 0, 0)
                Next ii
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                idLastCreated = acPoly.ObjectId()
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawMarkerNamed(ByRef sName As String, ByRef xyPoint As BODYSUIT1.XY, ByRef nWidth As Object, ByRef nHeight As Object, ByRef nAngle As Object)
        'Draw a Marker at the given point
        PrintLine(BODYSUIT1.fNum, "hEnt = AddEntity(" & BODYSUIT1.QQ & "marker" & BODYSUIT1.QCQ & sName & BODYSUIT1.QC & "xyStart.x+" & xyPoint.X & BODYSUIT1.CC & "xyStart.y+" & xyPoint.Y & BODYSUIT1.CC & nWidth & BODYSUIT1.CC & nHeight & BODYSUIT1.CC & nAngle & ");")
        PrintLine(BODYSUIT1.fNum, "SetDBData(hEnt," & BODYSUIT1.QQ & "ID" & BODYSUIT1.QQ & ",sID);")
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim xyStart As BODYSUIT1.XY
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
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + xyInsertion.X + 0.125, xyStart.Y + xyInsertion.Y + 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + xyInsertion.X + 0.125, xyStart.Y + xyInsertion.Y - 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y), 0, 0, 0)
            acPoly.TransformBy(Matrix3d.Rotation((nAngle * (BODYSUIT1.PI / 180)), Vector3d.ZAxis, New Point3d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y, 0)))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
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
    Sub PR_SetTextData(ByRef nHoriz As Object, ByRef nVert As Object, ByRef nHt As Object, ByRef nAspect As Object, ByRef nFont As Object)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to set the TEXT default attributes, these are
        'based on the values in the arguments.  Where the value is -ve then this
        'attribute is not set.
        'where :-
        '    nHoriz      Horizontal justification (1=Left, 2=Cen, 4=Right)
        '    nVert       Verticalal justification (8=Top, 16=Cen, 32=Bottom)
        '    nHt         Text height
        '    nAspect     Text aspect ratio (heigth/width)
        '    nFont       Text font (0 to 18)
        '
        'N.B. No checking is done on the values given
        '
        'Note:-
        '    fNum, CC, QQ, NL, g_nCurrTextHt, g_nCurrTextAspect,
        '    g_nCurrTextHorizJust, g_nCurrTextVertJust, g_nCurrTextFont
        '    are globals initialised by FN_Open
        Return '11-June-2018
        If nHoriz >= 0 And BODYSUIT1.g_nCurrTextHorizJust <> nHoriz Then
            PrintLine(BODYSUIT1.fNum, "SetData(" & BODYSUIT1.QQ & "TextHorzJust" & BODYSUIT1.QC & nHoriz & ");")
            BODYSUIT1.g_nCurrTextHorizJust = nHoriz
        End If

        If nVert >= 0 And BODYSUIT1.g_nCurrTextVertJust <> nVert Then
            PrintLine(BODYSUIT1.fNum, "SetData(" & BODYSUIT1.QQ & "TextVertJust" & BODYSUIT1.QC & nVert & ");")
            BODYSUIT1.g_nCurrTextVertJust = nVert
        End If

        If nHt >= 0 And BODYSUIT1.g_nCurrTextHt <> nHt Then
            PrintLine(BODYSUIT1.fNum, "SetData(" & BODYSUIT1.QQ & "TextHeight" & BODYSUIT1.QC & nHt & ");")
            BODYSUIT1.g_nCurrTextHt = nHt
        End If

        If nAspect >= 0 And BODYSUIT1.g_nCurrTextAspect <> nAspect Then
            PrintLine(BODYSUIT1.fNum, "SetData(" & BODYSUIT1.QQ & "TextAspect" & BODYSUIT1.QC & nAspect & ");")
            BODYSUIT1.g_nCurrTextAspect = nAspect
        End If

        If nFont >= 0 And BODYSUIT1.g_nCurrTextFont <> nFont Then
            PrintLine(BODYSUIT1.fNum, "SetData(" & BODYSUIT1.QQ & "TextFont" & BODYSUIT1.QC & nFont & ");")
            BODYSUIT1.g_nCurrTextFont = nFont
        End If
    End Sub
    Sub PR_DrawLegTemplate(ByRef sLeg As String, ByRef xyStart As BODYSUIT1.XY)
        'Draw the leg template based on the calculated values.
        'We are actually drawing the tick marks and the labels at
        'each tape.
        '  sLeg    "Left"
        '          "Right"
        '          "Both"  combine the two sets of templates
        '                  this becomes a bit tricky when we have pleats
        'Note the "profile" for the leg will be drawn elsewhere as the leg is a continuation
        'of the back body and is not drawn seperatly.
        '
        Dim iLastTape, iFirstTape, ii As Short
        Dim nMaxHt, nX, nMinHt As Double
        Dim LegProfile As curve

        'get pre-calculated module level variables
        If sLeg = "Left" Then
            iFirstTape = m_iLeftFirstTape
            iLastTape = m_iLeftLastTape
            LegProfile = g_LeftLegProfile
        ElseIf sLeg = "Right" Then
            iFirstTape = m_iRightFirstTape
            iLastTape = m_iRightLastTape
            LegProfile = g_RightLegProfile
        End If

        If sLeg = "Both" Then
            iFirstTape = min(m_iRightFirstTape, m_iLeftFirstTape)
            iLastTape = ARMDIA1.max(m_iRightLastTape, m_iLeftLastTape)

            PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(BODYSUIT1.g_sPathJOBST))
            PR_PutLine("@" & BODYSUIT1.g_sPathJOBST & "\BODY\BD_BOTH.D;")

            nX = xyStart.X
            For ii = iFirstTape To iLastTape
                nMaxHt = ARMDIA1.max(m_nLeftOriginalTapeValues(ii), m_nRightOriginalTapeValues(ii))
                nMinHt = min(m_nLeftOriginalTapeValues(ii), m_nRightOriginalTapeValues(ii))
                If m_nLeftOriginalTapeValues(ii) < m_nRightOriginalTapeValues(ii) Then
                    PR_PutStringAssign("sMinHtColour", "blue")
                    PR_PutStringAssign("sMaxHtColour", "red")
                ElseIf m_nLeftOriginalTapeValues(ii) = m_nRightOriginalTapeValues(ii) Then
                    PR_PutStringAssign("sMinHtColour", "black")
                    PR_PutStringAssign("sMaxHtColour", "black")
                Else
                    PR_PutStringAssign("sMinHtColour", "red")
                    PR_PutStringAssign("sMaxHtColour", "blue")
                End If
                PR_PutStringAssign("sSymbol", Trim(Str(ii)) & "tape")
                PR_PutNumberAssign("n20Len", m_n20Len(ii))
                PR_PutNumberAssign("nReduction", m_nReduction(ii))
                PR_PutNumberAssign("nMaxHt", nMaxHt)
                PR_PutNumberAssign("nMinHt", nMinHt)
                PR_PutNumberAssign("nX", nX)
                nX = nX + m_nSpace(ii)
                PR_PutLine("PR_DrawTapeScale();")
                PR_DrawTapeScale(nX, m_n20Len(ii), nMaxHt, m_nReduction(ii), Trim(Str(ii)) & "tape", ii)
            Next ii
        Else
            'sLeg = Left or sLeg = Right
            PR_PutStringAssign("sPathJOBST", FN_EscapeSlashesInString(BODYSUIT1.g_sPathJOBST))
            PR_PutLine("@" & BODYSUIT1.g_sPathJOBST & "\BODY\BD_SNGLE.D;")
            nX = xyStart.X
            For ii = iFirstTape To iLastTape
                If sLeg = "Left" Then
                    nMaxHt = m_nLeftOriginalTapeValues(ii)
                Else
                    nMaxHt = m_nRightOriginalTapeValues(ii)
                End If
                PR_PutStringAssign("sSymbol", Trim(Str(ii)) & "tape")
                PR_PutNumberAssign("n20Len", m_n20Len(ii))
                PR_PutNumberAssign("nReduction", m_nReduction(ii))
                PR_PutNumberAssign("nHt", nMaxHt)
                PR_PutNumberAssign("nX", nX)
                ''nX = nX + m_nSpace(ii)
                PR_PutLine("PR_DrawTapeScale();")
                PR_DrawTapeScale(nX, m_n20Len(ii), nMaxHt, m_nReduction(ii), Trim(Str(ii)) & "tape", ii)
                nX = nX + m_nSpace(ii)
            Next ii
        End If
    End Sub
    Private Sub PR_DrawRuler(ByRef sSymbol As String, ByRef xyStart As BODYSUIT1.XY, ByRef sTape As String)
        Dim xyPt As BODYSUIT1.XY
        Dim xyTapeFst, xyTapeSec, xyTapeEnd, xyTapeText As BODYSUIT1.XY
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
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y, 0), blkRecId)
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
    Private Sub PR_DrawTapeScale(nX As Double, n20Len As Double, nHt As Double, nReduction As Double, Symbol As String, nTape As Double)
        Dim xyPt1, xyTmp, xyStart, xyText, xyLineSt, xyLineEnd As BODYSUIT1.XY
        ''xyStart = xyInsertion
        xyPt1 = xyStart
        xyPt1.X = xyPt1.X + nX
        xyTmp.X = xyPt1.X
        xyTmp.Y = xyPt1.Y

        Dim nSeam, nLength, ii, jj, nTickStep, nTickHt, nTickLength As Double
        ''// Seam size
        nSeam = 0.25
        nLength = n20Len / 20 * nHt
        xyTmp.Y = xyStart.Y + nSeam + nLength
        ''hEnt = AddEntity("marker", "xmarker", xyTmp, 0.2, 0.2)
        PR_DrawMarker(xyTmp)

        ''// Add Ticks And Text at each scale
        ''// Start ticks 2 below And 2 above
        ii = Int(nLength / (n20Len / 20)) - 1
        jj = ii + 2

        nTickStep = n20Len / (20 * 8)
        nTickHt = xyStart.Y + nSeam + ii * nTickStep * 8

        ''// Add Text
        ''SetData("TextHeight", 0.125);
        ''SetData("TextHorzJust", 1);

        ''// Reduction text	
        ''AddEntity("text", MakeString("long", nReduction), xyPt1.X + 0.2, nTickHt - 0.4)
        PR_MakeXY(xyText, xyPt1.X + 0.2, nTickHt - 0.4)
        PR_DrawText(Str(nReduction), xyText, 0.125, 0, 1)

        ''// Maximum text	
        ''hEnt = AddEntity("text", Format("length", nHt), xyPt1.X + 0.2, nTickHt - nSeam)
        PR_MakeXY(xyText, xyPt1.X + 0.2, nTickHt - nSeam)
        PR_DrawText(fnInchestoText(nHt), xyText, 0.125, 0, 1)
        ''// Add tape ID
        'If (!Symbol("find", sSymbol)) Then
        '    MsgBox("Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists", 16, "BodySuit - Dialog")
        '    Exit Sub
        'End If
        'AddEntity("symbol", sSymbol, xyPt1.X, nTickHt)
        PR_MakeXY(xyText, xyPt1.X, nTickHt)
        Dim g_sTextList As String = "-7½ -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
        Dim strTape As String = LTrim(Mid(g_sTextList, ((nTape - 1) * 3) + 1, 3))
        PR_DrawRuler(Symbol, xyText, strTape)

        ''// Add graticules
        While (ii <= jj)
            nTickLength = 0.22
            ''AddEntity("line", xyPt1.X, nTickHt, xyPt1.X + nTickLength, nTickHt)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt)
            PR_DrawLine(xyLineSt, xyLineEnd)
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
            PR_DrawText(Str(ii), xyText, 0.1, 90, nTextMode)
            nTickLength = 0.05
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep, xyPt1.X + nTickLength, nTickHt + nTickStep)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep)
            PR_DrawLine(xyLineSt, xyLineEnd)
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 3, xyPt1.X + nTickLength, nTickHt + nTickStep * 3)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 3)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 3)
            PR_DrawLine(xyLineSt, xyLineEnd)
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 5, xyPt1.X + nTickLength, nTickHt + nTickStep * 5)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 5)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 5)
            PR_DrawLine(xyLineSt, xyLineEnd)
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 7, xyPt1.X + nTickLength, nTickHt + nTickStep * 7)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 7)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 7)
            PR_DrawLine(xyLineSt, xyLineEnd)
            nTickLength = 0.08
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 2, xyPt1.X + nTickLength, nTickHt + nTickStep * 2)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 2)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 2)
            PR_DrawLine(xyLineSt, xyLineEnd)
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 6, xyPt1.X + nTickLength, nTickHt + nTickStep * 6)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 6)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 6)
            PR_DrawLine(xyLineSt, xyLineEnd)
            nTickLength = 0.12
            ''AddEntity("line", xyPt1.X, nTickHt + nTickStep * 4, xyPt1.X + nTickLength, nTickHt + nTickStep * 4)
            PR_MakeXY(xyLineSt, xyPt1.X, nTickHt + nTickStep * 4)
            PR_MakeXY(xyLineEnd, xyPt1.X + nTickLength, nTickHt + nTickStep * 4)
            PR_DrawLine(xyLineSt, xyLineEnd)
            ii = ii + 1
            nTickHt = xyStart.Y + nSeam + ii * nTickStep * 8
        End While
    End Sub
    Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
        'Search through the string looking for " (double quote characater)
        'If found use \ (Backslash) to escape it
        '
        Dim ii As Short
        Dim Char_Renamed As String
        Dim sEscapedString As String

        FN_EscapeSlashesInString = ""

        For ii = 1 To Len(sAssignedString)
            Char_Renamed = Mid(sAssignedString, ii, 1)
            If Char_Renamed = "\" Then
                sEscapedString = sEscapedString & "\" & Char_Renamed
            Else
                sEscapedString = sEscapedString & Char_Renamed
            End If
        Next ii

        FN_EscapeSlashesInString = sEscapedString

    End Function
    Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the maximum of two numbers
        If nFirst >= nSecond Then
            max = nFirst
        Else
            max = nSecond
        End If
    End Function
    Sub PR_CalcRaglan(ByRef xyStart As BODYSUIT1.XY, ByRef xyEnd As BODYSUIT1.XY, ByRef Flip As Short, ByRef Raglan As curve,
                      ByRef xyOpen As BODYSUIT1.XY, ByRef aOpen As Double, ByRef nRadiusForRequiredPoint As Double)
        'Subroutine to draw a vest raglan curve between the two points
        'given
        'NB
        ' nRadiusForRequiredPoint is used to get the point xyOpen
        ' if nRadiusForRequiredPoint = -1 then xyOpen is at 1/3 the distance along the curve
        ' if nRadiusForRequiredPoint > 0  then xyOpen is at the radius nRadiusForRequiredPoint
        '
        Static nSegLen(100) As Double
        Static nSegAngle(100) As Double
        Static iSegments As Short
        Static sFileName As String
        Static fFileNum As Short
        Static xy2, xy0, xy1, xy4 As BODYSUIT1.XY
        Static nLength, aPrevAngle, aAngle, nTraversedLen, nOpenLength As Double
        Static bOpenFound As Short
        Static aAxis, aRaglan, nRadius As Double
        Static ii As Short

        'On an error the Raglan Curve is returned as a minimum of 2 points
        'these being the given start and end points.
        Raglan.n = 2
        Dim X(2), Y(2) As Double
        Raglan.X = X
        Raglan.y = Y
        Raglan.X(1) = xyStart.X
        Raglan.y(1) = xyStart.Y
        Raglan.X(2) = xyEnd.X
        Raglan.y(2) = xyEnd.Y

        'Check for silly data
        If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then Exit Sub

        'If g_sSleeveType = "Sleeveless" Then
        '    sFileName = BODYSUIT1.g_sPathJOBST & "\TEMPLTS\BODYCURV.DAT"
        'Else
        '    sFileName = BODYSUIT1.g_sPathJOBST & "\TEMPLTS\VESTCURV.DAT"
        'End If
        If g_sSleeveType = "Sleeveless" Then
            sFileName = fnGetSettingsPath("LookupTables") & "\BODYCURV.DAT"
        Else
            sFileName = fnGetSettingsPath("LookupTables") & "\VESTCURV.DAT"
        End If
        'Check to see if the raglan curve has been loaded from file,
        'so we don't load it twice
        If iSegments = 0 Then
            fFileNum = FreeFile()

            If FileLen(sFileName) = 0 Then
                MsgBox("Can't open the file " & sFileName & ".  Unable to draw the raglan curve.", 48)
                Exit Sub
            End If

            FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
            Do While Not EOF(fFileNum)
                iSegments = iSegments + 1
                Input(fFileNum, nSegLen(iSegments))
                Input(fFileNum, nSegAngle(iSegments))
            Loop
            FileClose(fFileNum)
        End If

        'Using the given points loop through the raglan curve to establish the
        'raglan angle.
        'Note:
        'This code is a little bit cumbersome but it is based on DRAFIX macro
        'code that is known to work and is production accepted.
        '
        nRadius = FN_CalcLength(xyStart, xyEnd)

        aRaglan = -1
        xy1.X = 0 : xy1.Y = 0
        xy0.X = 0 : xy0.Y = 0
        aPrevAngle = 0
        aAngle = nSegAngle(1)
        nLength = nSegLen(1)
        nTraversedLen = 0 'For use later

        For ii = 2 To iSegments
            aAngle = aAngle + aPrevAngle
            PR_CalcPolar(xy1, aAngle, nLength, xy2)
            If FN_CirLinInt(xy1, xy2, xy0, nRadius, xy4) Then
                aRaglan = FN_CalcAngle(xy0, xy4)
                nTraversedLen = nTraversedLen + FN_CalcLength(xy1, xy4)
                Exit For
            End If
            nTraversedLen = nTraversedLen + FN_CalcLength(xy1, xy2)
            xy1 = xy2
            aPrevAngle = aAngle
            aAngle = nSegAngle(ii)
            nLength = nSegLen(ii)
        Next ii

        'Check that the intesection has occured
        If aRaglan < 0 Then
            MsgBox("Can't draw the draw the raglan curve with this data.", 48)
            Exit Sub
        End If

        'Create the curve given the angle above and the two point arguments
        '
        nOpenLength = nTraversedLen / 3
        nTraversedLen = 0
        bOpenFound = False

        aAngle = nSegAngle(1)
        nLength = nSegLen(1)
        aAxis = FN_CalcAngle(xyStart, xyEnd)
        aPrevAngle = aAxis - aRaglan
        xy1 = xyStart 'NB Raglan.X(1) and Raglan.Y(1), set at begining to xyStart
        Dim X1(iSegments), Y1(iSegments) As Double
        Raglan.X = X1
        Raglan.y = Y1
        Raglan.X(1) = xyStart.X
        Raglan.y(1) = xyStart.Y
        For ii = 2 To iSegments
            aAngle = aAngle + aPrevAngle
            PR_CalcPolar(xy1, aAngle, nLength, xy2)
            If FN_CirLinInt(xy1, xy2, xyStart, nRadius, xy4) Then
                Raglan.X(ii) = xyEnd.X
                Raglan.y(ii) = xyEnd.Y
                Exit For
            End If
            If nRadiusForRequiredPoint > 0 Then
                If FN_CirLinInt(xy1, xy2, xyStart, nRadiusForRequiredPoint, xy4) Then
                    xyOpen = xy4
                    aOpen = FN_CalcAngle(xy1, xy2)
                    If Flip Then
                        aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
                    Else
                        aOpen = aOpen - 90
                    End If
                End If
            End If
            Raglan.X(ii) = xy2.X
            Raglan.y(ii) = xy2.Y
            If nTraversedLen + nLength >= nOpenLength And Not bOpenFound And nRadiusForRequiredPoint <= 0 Then
                bOpenFound = True
                aOpen = FN_CalcAngle(xy1, xy2)
                PR_CalcPolar(xy1, aOpen, (nTraversedLen + nLength) - nOpenLength, xyOpen)
                If Flip Then
                    aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
                Else
                    aOpen = aOpen - 90
                End If
            End If
            nTraversedLen = nTraversedLen + nLength
            xy1 = xy2
            aPrevAngle = aAngle
            aAngle = nSegAngle(ii)
            nLength = nSegLen(ii)

        Next ii
        Raglan.n = ii

        'If flip is true then mirror Raglan curve in the line (xyStart, xyEnd)
        'The angle of which is given by aAxis
        If Flip Then
            For ii = 1 To Raglan.n
                PR_MakeXY(xy1, Raglan.X(ii), Raglan.y(ii))
                nLength = FN_CalcLength(xyStart, xy1)
                aAngle = FN_CalcAngle(xyStart, xy1)
                aAngle = (360 - (aAngle - aAxis) + aAxis)
                PR_CalcPolar(xyStart, aAngle, nLength, xy2)
                Raglan.X(ii) = xy2.X
                Raglan.y(ii) = xy2.Y
            Next ii
            nLength = FN_CalcLength(xyStart, xyOpen)
            aAngle = FN_CalcAngle(xyStart, xyOpen)
            aAngle = (360 - (aAngle - aAxis) + aAxis)
            PR_CalcPolar(xyStart, aAngle, nLength, xyOpen)
        End If
    End Sub
    Private Sub PR_DrawMeshConstruction(ByRef xyRaglanStart As BODYSUIT1.XY, ByRef xyMeshStart As BODYSUIT1.XY, ByRef sAxillaText As String, ByRef nMeshLength As Double)
        'Draws the custruction lines for a mesh
        'Save code
        'Returns
        '    xyMeshStart point
        '    AxillaText
        '    Length of Mesh
        Dim nOffset As Double
        Dim xyTmp, xyTmp1 As BODYSUIT1.XY
        PR_SetLayer("1")
        If g_nAge < 14 Then
            nOffset = 0.75
            nMeshLength = ONE_and_THREE_QUARTER_GUSSET '2.58
            sAxillaText = "GUSSET 1-3/4\"""
        Else
            nOffset = 0.8125
            nMeshLength = BOYS_GUSSET '2.8
            sAxillaText = "BOY'S GUSSET"
        End If

        PR_MakeXY(xyTmp, xyRaglanStart.X - nOffset, xyRaglanStart.Y)
        PR_DrawLine(xyRaglanStart, xyTmp)
        PR_MakeXY(xyTmp1, xyTmp.X, xyTmp.Y + 0.25)
        PR_MakeXY(xyTmp, xyTmp1.X, xyTmp1.Y - 0.5)
        PR_DrawLine(xyTmp1, xyTmp)
        PR_MakeXY(xyMeshStart, xyRaglanStart.X - nOffset, xyRaglanStart.Y)
    End Sub
    Function FN_CalcAxillaMesh(ByRef xyRaglanStart As BODYSUIT1.XY, ByRef xyRaglanEnd As BODYSUIT1.XY, ByRef xyMeshStart As BODYSUIT1.XY, ByRef nFlip As Short,
                               ByRef nMeshLength As Double, ByRef nAlongRaglan As Double, ByRef MeshProfile As BiArc, ByRef MeshSeam As BiArc) As Short
        'Subroutine to calculate a mesh w.r.t Mesh axilla in the
        'Vest, Body-suit and their respective sleeves

        'Due to problems with the CALC BiARC function we calculate all for the lower
        'half and then flip the result
        'A cheat but what the hell

        Dim aDeltaEnd, aEnd, aInitialStart, aStart, aAngle, aInitialEnd, aMeshEnd, aStartDelta As Double
        Dim nEndLength, nStartLength, nStep As Double
        Dim nRadiusTol, nLength, nTol, nSeam As Double
        Dim bMeshFound, ii As Short
        Dim aAxis, iSign, aEndFactor As Double
        Dim xyMeshEnd, xyTmp As BODYSUIT1.XY
        Dim cDummyCurve As curve

        'intially start as False
        FN_CalcAxillaMesh = False

        'Check for silly data
        If xyRaglanStart.X = xyRaglanEnd.X And xyRaglanStart.Y = xyRaglanEnd.Y Then Exit Function
        If xyMeshStart.X = xyRaglanEnd.X And xyMeshStart.Y = xyRaglanEnd.Y Then Exit Function
        If xyRaglanStart.X = xyMeshStart.X And xyRaglanStart.Y = xyMeshStart.Y Then Exit Function

        If FN_CalcLength(xyMeshStart, xyRaglanStart) >= nMeshLength Then Exit Function


        'initialize variables
        '    nTol = .03125
        nTol = 0.04
        bMeshFound = False
        nStep = -INCH1_16

        If nAlongRaglan > 0 Then
            nStartLength = nAlongRaglan
            nEndLength = nAlongRaglan 'do at least Once
            aDeltaEnd = 160 'Must find
        Else
            nLength = FN_CalcLength(xyMeshStart, xyRaglanStart)
            nEndLength = nLength + INCH3_16
            nStartLength = nLength * 2.5 '

            aDeltaEnd = 90
            nStep = INCH1_16
            If nMeshLength = ONE_and_THREE_QUARTER_GUSSET Then
                nEndLength = nLength * 2.5 '
                nStartLength = 1.25 '
            Else
                nEndLength = nLength * 2.5 '
                nStartLength = 1.5 '
            End If

        End If

        If My.Application.Info.Title <> "meshdraw" Then
            'Body suit Mesh
            aAxis = FN_CalcAngle(xyMeshStart, xyRaglanStart)
            aStartDelta = 40
            '        nRadiusTol = .4
            nRadiusTol = 0
            aInitialStart = 265
        Else
            'Sleeve Mesh
            ''-------aAxis = FN_CalcAngle(xyMeshStart, xyAxilla)
            aAxis = FN_CalcAngle(xyMeshStart, xyRaglanStart)
            aStartDelta = 60
            aInitialStart = 280
            nRadiusTol = 0
        End If

        For aStart = aInitialStart To (aInitialStart + aStartDelta) Step 1
            'PR_CalcPolar xyMeshStart, aStart, 1, xyTmp
            'PR_DrawMarker xyTmp

            For nLength = nStartLength To nEndLength Step nStep
                'Get point on raglan
                PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, xyMeshEnd, aInitialEnd, nLength)
                ' PR_CalcPolar xyMeshEnd, aInitialEnd, 3, xyTmp
                ' PR_DrawLine xyTmp, xyMeshEnd

                'Flip the point and angle
                If nFlip Then
                    'mirror in the line (xyMeshStart, xyRaglanStart)
                    PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, xyMeshEnd) + aAxis), FN_CalcLength(xyMeshStart, xyMeshEnd), xyMeshEnd)
                    aMeshEnd = ((360 - aInitialEnd) + aAxis) + 180
                Else
                    aMeshEnd = (aInitialEnd + 180)
                End If

                'End angle iteration - Get biarc curve that fits the given length to +/- nTol
                'Hold start change end
                For aEnd = (aMeshEnd - (aDeltaEnd / 2)) To (aMeshEnd + (aDeltaEnd / 2)) Step 1
                    ' PR_CalcPolar xyMeshEnd, aEnd, 1, xyTmp
                    ' PR_DrawMarker xyTmp

                    If FN_FitBiArcAtGivenLength(xyMeshStart, aStart, xyMeshEnd, aEnd, nMeshLength, nTol, nRadiusTol, MeshProfile) Then
                        bMeshFound = True
                        Exit For
                    End If
                Next aEnd
                If bMeshFound Then
                    FN_CalcAxillaMesh = True
                    nAlongRaglan = nLength
                    Exit For
                End If

            Next nLength

            If bMeshFound Then Exit For

        Next aStart

        'Don't calculate the seam unless we have a mesh
        If Not bMeshFound Then Exit Function

        'Calculate seam line
        nSeam = 0.125

        'Calculate MeshSeam based on MeshProfile
        If nFlip Then
            'mirror in the line (xyMeshStart, xyRaglanStart)
            PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, MeshProfile.xyEnd) + aAxis), FN_CalcLength(xyMeshStart, MeshProfile.xyEnd), MeshProfile.xyEnd)
            PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, MeshProfile.xyTangent) + aAxis), FN_CalcLength(xyMeshStart, MeshProfile.xyTangent), MeshProfile.xyTangent)
            PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, MeshProfile.xyR1) + aAxis), FN_CalcLength(xyMeshStart, MeshProfile.xyR1), MeshProfile.xyR1)
            PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, MeshProfile.xyR2) + aAxis), FN_CalcLength(xyMeshStart, MeshProfile.xyR2), MeshProfile.xyR2)
        End If

        MeshSeam = MeshProfile

        'Get seam intersection on raglan
        nLength = nLength - nSeam
        PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, MeshSeam.xyEnd, aInitialEnd, nLength)

        'Calculate Start, End and Tangent
        aAngle = FN_CalcAngle(MeshProfile.xyR1, MeshProfile.xyStart)
        MeshSeam.nR1 = MeshProfile.nR1 - nSeam
        PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyStart)

        aAngle = FN_CalcAngle(MeshProfile.xyR1, MeshProfile.xyTangent)
        PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyTangent)
        MeshSeam.nR2 = MeshProfile.nR2 - nSeam

        'PR_DrawMarker MeshSeam.xyR2
        'PR_DrawText "MeshSeam.xyStart", MeshSeam.xyStart, .1

    End Function
    Function FN_FitBiArcAtGivenLength(ByRef xyStart As BODYSUIT1.XY, ByRef aStart As Double, ByRef xyEnd As BODYSUIT1.XY, ByRef aEnd As Double,
                                      ByRef nLength As Double, ByRef nLengthTol As Double, ByRef nRadiusTol As Double, ByRef Mesh As BiArc) As Integer

        Dim aA1#, aA2#, nArc#, aArc#

        FN_FitBiArcAtGivenLength = False
        If FN_BiArcVestCurve(xyStart, aStart, xyEnd, aEnd, Mesh) Then
            'Calculate Length
            'First arc
            aA1 = FN_CalcAngle(Mesh.xyR1, Mesh.xyStart)
            aA2 = FN_CalcAngle(Mesh.xyR1, Mesh.xyTangent)
            aArc = aA2 - aA1
            nArc = (aArc * (BODYSUIT1.PI / 180)) * Mesh.nR1
            'Second arc
            aA1 = FN_CalcAngle(Mesh.xyR2, Mesh.xyTangent)
            aA2 = FN_CalcAngle(Mesh.xyR2, Mesh.xyEnd)
            aArc = aA2 - aA1
            nArc = nArc + ((aArc * (BODYSUIT1.PI / 180)) * Mesh.nR2)

            'Check that the length is within the tolerence
            If nArc > (nLength - nLengthTol) And nArc < (nLength + nLengthTol) Then
                FN_FitBiArcAtGivenLength = True
                If nRadiusTol > 0 And System.Math.Abs(Mesh.nR1 - Mesh.nR2) > nRadiusTol Then FN_FitBiArcAtGivenLength = False
            End If
        End If
    End Function
    Public Function FN_BiArcVestCurve(ByRef xyStart As BODYSUIT1.XY, ByRef aStart As Double, ByRef xyEnd As BODYSUIT1.XY, ByRef aEnd As Double, ByRef Profile As BiArc) As Short
        'Procedure to fit two arcs with a common tangent
        'between two points at the tangent angles specified.
        'Returning a BiArc curve.
        '
        'InputArm
        '        xyStart     Start point (XY System)
        '        xyEnd       End Point   (XY System)
        '        aStart      Start Tangent Angle in degrees
        '        aEnd        End Tangent Angle in degrees
        '
        'Output
        '        Profile     BiArc in (XY System)
        '
        'Restrictions:-
        'Angles must be positive and < 360
        'The data must result in a NON-INFLECTING curve.
        '
        '
        'Notes:-
        'The BiArc is a structure that is used to represent
        'the curve.
        '
        '        Type BiArc
        '                xyStart     as XY
        '                xyTangent   as XY
        '                xyEnd       as XY
        '                xyR1        as XY
        '                xyR2        as XY
        '                nR1         as Double
        '                nR2         as Double
        '        end
        '
        'The point of this represntation is to make
        'the drawing of the curve as simple as possible,
        'without the need to make further calculations.
        '
        '
        'Acknowledgements:-
        '
        '    British Ship Research Association (BSRA)
        '    Technical Memorandum No. 388
        '
        '   "The Fitting of Smooth Curves by Circular
        '    Arcs and Straight Lines"
        '    K.M.Bolton, B.Sc., M.Sc., Grad.I.M.A.
        '    October 1970
        '
        'KNOWN BUGS
        'Needs much more work to make a more general robust tool.
        'Works OK for some special cases (cf THUMB Curves)
        'GG 22.Mar.96
        '
        '
        Dim nTmp, aAxis, nLength As Double
        Dim Phi1, Theta1, Theta2, Phi2 As Double
        Dim C1, S1, d, B, A, c, P, S2, C2 As Double
        Dim R1, Rs, R2 As Double
        Dim Rmin As Object
        Dim uvR2, uvR1, uvOrigin As BODYSUIT1.XY
        '    Dim fError

        'Use a file for printing results for debug
        'Open file
        '    fError = FreeFile
        '    Open "C:\TMP\FIT_ERR.DAT" For Output As fError


        'Initially return false
        FN_BiArcVestCurve = False

        'Check for silly data and return
        If aStart = aEnd Then GoTo Error_Close
        If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then GoTo Error_Close

        'Translate tangent angle in XY system to the UV system.
        'Where the U-Axis is specified by the line xyStart,xyEnd
        '
        aAxis = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)
        'Print #fError, "aAxis degrees ="; aAxis
        'Print #fError, "nLength ="; nLength

        Theta1 = aStart - aAxis
        If Theta1 = 0 Then GoTo Error_Close 'Straight line

        If Theta1 < 0 Then Theta1 = 360 + Theta1
        'Print #fError, "Theta1 degrees ="; Theta1
        Theta1 = Theta1 * (BODYSUIT1.PI / 180)
        'Print #fError, "Theta1="; Theta1

        Theta2 = aEnd - aAxis
        If Theta2 < 0 Then Theta2 = 360 + Theta2
        'Print #fError, "Theta2 degrees="; Theta2
        Theta2 = Theta2 * (BODYSUIT1.PI / 180)
        'Print #fError, "Theta2="; Theta2


        'Check that it is non-inflecting
        'return an error (false) for the inflecting case
        'let the calling routine worry about handling it
        If (Theta1 > 0 And Theta1 < (BODYSUIT1.PI / 2)) And (Theta2 > 0 And Theta2 < (BODYSUIT1.PI / 2)) Then GoTo Error_Close
        If (Theta1 > (3 * (BODYSUIT1.PI / 2)) And Theta1 < (2 * BODYSUIT1.PI)) And (Theta2 > (3 * (BODYSUIT1.PI / 2)) And Theta2 < (2 * BODYSUIT1.PI)) Then GoTo Error_Close

        'Calculate acute unsigned tangent angles to the line in the UV
        'co-ordinate system
        ' Phi1 = Theta1
        If Theta1 < BODYSUIT1.PI Then
            Phi1 = Theta1
        Else
            Phi1 = System.Math.Abs((2 * BODYSUIT1.PI) - Theta1)
        End If
        If Theta2 < BODYSUIT1.PI Then
            Phi2 = Theta2
        Else
            Phi2 = System.Math.Abs((2 * BODYSUIT1.PI) - Theta2)
        End If

        'Print #fError, "Phi1="; Phi1
        'Print #fError, "Phi2="; Phi2


        'Calculate R1 and R2
        S1 = System.Math.Abs(System.Math.Sin(Theta1))
        C1 = (-System.Math.Sin(Theta1) * System.Math.Cos(Theta1)) / S1
        'Print #fError, "S1="; S1; "C1="; C1

        S2 = System.Math.Abs(System.Math.Sin(Theta2))
        C2 = (System.Math.Sin(Theta2) * System.Math.Cos(Theta2)) / S2
        'Print #fError, "S2="; S2; "C2="; C2

        P = FN_CalcLength(xyStart, xyEnd)
        'Print #fError, "P="; P

        If Phi1 <> Phi2 Then
            'Print #fError, "Phi1 <> Phi2"
            A = S1 + S2
            B = (S1 * S2) - (C1 * C2) + 1
            c = S2
            Rs = (P * c) / B
            nTmp = (c ^ 2) - (c * A) + (B / 2)
            'Print #fError, "A="; A; "B="; B; "C="; C
            'Print #fError, "Rs="; Rs
            'Print #fError, "nTmp="; nTmp
            'As this is a root we check that it is not -ve
            If nTmp < 0 Then GoTo Error_Close
            nTmp = (P * System.Math.Sqrt(nTmp)) / B
            'Print #fError, "nTmp="; nTmp
            If Phi1 > Phi2 Then
                R1 = Rs - nTmp
            Else
                R1 = Rs + nTmp
            End If
            'Print #fError, "R1="; R1
            If R1 <= 0 Then GoTo Error_Close
            d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
            'Print #fError, "D="; d
            R2 = R1 - d
            'Print #fError, "R2="; R2
        Else
            'Print #fError, "Phi1 = Phi2"
            A = 2 * System.Math.Sin(Theta1)
            B = (System.Math.Sin(Theta1) ^ 2) - (System.Math.Cos(Theta1) ^ 2) + 1
            R1 = (P * A) / (2 * B)
            R2 = R1
            'Print #fError, "A="; A; "B="; A;
            'Print #fError, "R1="; R1
        End If

        'The radi R1 and R2 must be greater than the specified
        'minimum Rmin
        Rmin = nLength / 3
        If R1 < Rmin Then
            'Print #fError, "R1 < Rmin"
            R1 = Rmin
            d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
            'Print #fError, "D="; d
            R2 = R1 - d
            'Print #fError, "R1="; R1
            'Print #fError, "R2="; R2
        End If
        If R2 < Rmin Then
            'Print #fError, "R2 < Rmin"
            R2 = Rmin
            R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
            'Print #fError, "R1="; R1
            'Print #fError, "R2="; R2
        End If

        'Check for errors
        If R1 < 0 Or R2 < 0 Then GoTo Error_Close

        'Using the calculated radi, create BiArc Curve
        'Start and end points of bi-arc curve
        Profile.xyStart = xyStart
        Profile.xyEnd = xyEnd

        'Centers of arcs
        'Get UV co-ordinates
        PR_MakeXY(uvR1, R1 * S1, R1 * C1)
        'Print #fError, "uvR1="; uvR1.X, uvR1.Y
        'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
        PR_MakeXY(uvR2, P - (R2 * S2), R2 * C2)
        'Print #fError, "uvR2="; uvR2.X, uvR2.Y
        'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
        PR_MakeXY(uvOrigin, 0, 0)
        'Translate to XY co-ordinates
        PR_CalcPolar(xyStart, FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1)
        PR_CalcPolar(xyStart, FN_CalcAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2)
        'Print #fError, "R1="; R1
        'Print #fError, "R2="; R2
        'Tangent point on arc
        If R1 < R2 Then
            PR_CalcPolar(Profile.xyR1, FN_CalcAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent)
        Else
            PR_CalcPolar(Profile.xyR1, FN_CalcAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent)
        End If
        'radi of arcs
        Profile.nR1 = R1
        Profile.nR2 = R2

        'Test that the Tangent point lies between the start and end points
        If FN_CalcLength(xyStart, Profile.xyTangent) > nLength Then GoTo Error_Close
        If FN_CalcLength(xyEnd, Profile.xyTangent) > nLength Then GoTo Error_Close

        'return true as we have a sucessful fit
        'Close #fError
        FN_BiArcVestCurve = True
        Exit Function
Error_Close:
        'Print #fError, "Error and close"
        'Close #fError
    End Function
    Private Sub PR_DrawMesh(ByRef MeshProfile As BiArc, ByRef MeshSeam As BiArc, ByRef sSide As String, Optional ByRef bIsSwap As Boolean = False)
        'Subroutime to draw the given mesh
        'used to save code
        'NB the misnomer. Too complicated to change in the DrawMacro
        'procedure so do it here
        PR_SetLayer("Notes")
        If bIsSwap = True Then
            PR_DrawArc(MeshProfile.xyR1, MeshProfile.xyTangent, MeshProfile.xyStart)
            PR_DrawArc(MeshProfile.xyR2, MeshProfile.xyEnd, MeshProfile.xyTangent)
        Else
            PR_DrawArc(MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent)
            PR_DrawArc(MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd)
        End If
        PR_SetLayer("Template" & sSide)
        If bIsSwap = True Then
            PR_DrawArc(MeshSeam.xyR1, MeshSeam.xyTangent, MeshSeam.xyStart)
            PR_DrawArc(MeshSeam.xyR2, MeshSeam.xyEnd, MeshSeam.xyTangent)
        Else
            PR_DrawArc(MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent)
            PR_DrawArc(MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd)
        End If
    End Sub
    Sub PR_DrawRectangle(ByRef xyLLCorner As BODYSUIT1.XY, ByRef xyURCorner As BODYSUIT1.XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a RECTANGLE using opposite corner points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '

        PrintLine(BODYSUIT1.fNum, "hEnt = AddEntity(" & BODYSUIT1.QQ & "poly" & BODYSUIT1.QCQ & "polygon" & BODYSUIT1.QQ)
        PrintLine(BODYSUIT1.fNum, BODYSUIT1.CC & "xyStart.x+" & Str(xyLLCorner.X) & BODYSUIT1.CC & "xyStart.y+" & Str(xyLLCorner.Y))
        PrintLine(BODYSUIT1.fNum, BODYSUIT1.CC & "xyStart.x+" & Str(xyLLCorner.X) & BODYSUIT1.CC & "xyStart.y+" & Str(xyURCorner.Y))
        PrintLine(BODYSUIT1.fNum, BODYSUIT1.CC & "xyStart.x+" & Str(xyURCorner.X) & BODYSUIT1.CC & "xyStart.y+" & Str(xyURCorner.Y))
        PrintLine(BODYSUIT1.fNum, BODYSUIT1.CC & "xyStart.x+" & Str(xyURCorner.X) & BODYSUIT1.CC & "xyStart.y+" & Str(xyLLCorner.Y))
        PrintLine(BODYSUIT1.fNum, ");")
        PrintLine(BODYSUIT1.fNum, "if (hEnt) SetDBData(hEnt," & BODYSUIT1.QQ & "ID" & BODYSUIT1.QQ & ",sID);")

        Dim ptCurve As curve
        ptCurve.n = 5
        Dim X(5), Y(5) As Double
        ptCurve.X = X
        ptCurve.y = Y
        ptCurve.X(1) = xyLLCorner.X
        ptCurve.y(1) = xyLLCorner.Y
        ptCurve.X(2) = xyLLCorner.X
        ptCurve.y(2) = xyURCorner.Y
        ptCurve.X(3) = xyURCorner.X
        ptCurve.y(3) = xyURCorner.Y
        ptCurve.X(4) = xyURCorner.X
        ptCurve.y(4) = xyLLCorner.Y
        ptCurve.X(5) = xyLLCorner.X
        ptCurve.y(5) = xyLLCorner.Y
        PR_DrawPoly(ptCurve)
    End Sub
    Sub PR_AddEntityArc(ByRef xyCen As BODYSUIT1.XY, ByRef xyArcStart As BODYSUIT1.XY, ByRef xyArcEnd As BODYSUIT1.XY)
        'Draws an arc with original parameters for AddEntity("arc",...) routine

        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object

        nRad = FN_CalcLength(xyCen, xyArcStart)
        nStartAng = FN_CalcAngle(xyCen, xyArcStart)
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
        nDeltaAng = System.Math.Abs(nStartAng - nEndAng)
        PrintLine(BODYSUIT1.fNum, "hEnt = AddEntity(" & BODYSUIT1.QQ & "arc" & BODYSUIT1.QC & "xyStart.x +" & Str(xyCen.X) & BODYSUIT1.CC & "xyStart.y +" & Str(xyCen.Y) & BODYSUIT1.CC & Str(nRad) & BODYSUIT1.CC & Str(nStartAng) & BODYSUIT1.CC & Str(nDeltaAng) & ");")
        PrintLine(BODYSUIT1.fNum, "SetDBData(hEnt," & BODYSUIT1.QQ & "ID" & BODYSUIT1.QQ & ",sID);")
        PR_DrawArc(xyCen, xyArcStart, xyArcEnd)

    End Sub
    Private Sub PR_DrawBriefStandard(ByRef nXOffset As Double)
        Dim cBriefPoly As curve
        Dim cCrotchPoly As curve
        Dim nArc As Double
        Dim xyArc, xyTmp As BODYSUIT1.XY
        Dim aAngle, aArcAngle, aConstruct As Double
        Dim ii As Short

        'To take account of the One leg Panty and One leg Brief
        'we use the nXOffset to move the profile
        'This is a fudge but one that works.
        'Normal brief only not applicable to the Brief-French
        'NB we can reset the Global xy...  points as they are not used again.

        If nXOffset > 0 Then '==>Mixed Panty & Brief
            xyBodyProfileThigh = xyBodyFold
            PR_MakeXY(xyBDProfileThighMirror, xyBodyProfileThigh.X, -xyBodyProfileThigh.Y)
            PR_MakeXY(xyBDProfileThighMirror1, xyBDProfileThighMirror.X, xyBDProfileThighMirror.Y + BODY_INCH1_4)
            xyBodyProfileBrief.X = xyBodyProfileBrief.X + nXOffset

            xyBodyLT_ProfileThigh = xyBodyLT_Fold
            PR_MakeXY(xyBodyLT_ProfileThighMirror, xyBodyLT_ProfileThigh.X, -xyBodyLT_ProfileThigh.Y)
            '        PR_MakeXY xyLT_ProfileThighMirror1, xyBodyLT_ProfileThighMirror.x, xyBodyLT_ProfileThighMirror.y + inch1_4

            xyBodyLegPoint.X = xyBodyLegPoint.X + nXOffset
            xyBodyCutOutArcCentre.X = xyBodyCutOutArcCentre.X + nXOffset
            xyBodySeamThigh.X = xyBodySeamThigh.X + nXOffset

            xyBodyLT_ProfileThighMirror.X = xyBodyLT_ProfileThighMirror.X + nXOffset
            xyBodyStrap4.X = xyBodyStrap4.X + nXOffset
            xyBodyGussetArcCentre.X = xyBodyGussetArcCentre.X + nXOffset

        End If

        If g_bBodyDiffThigh Then
            PR_DrawLine(xyBodyLT_ProfileThigh, xyBodyProfileBrief)
        Else
            PR_DrawLine(xyBodyProfileThigh, xyBodyProfileBrief)
        End If

        If g_sBodyCrotchStyle <> "Snap Crotch" Then
            'All other crotch styles
            cBriefPoly.n = 0
            aConstruct = FN_CalcAngle(xyBodyLegPoint, xyBodyProfileBrief)
            aConstruct = aConstruct - ((aConstruct - 90) / 2)
            nArc = (xyBodyCutOutArcCentre.X - xyBodySeamThigh.X) * (2 / 3)
            ' PR_MakeXY xyArc, xyBodyProfileBrief.X + nArc, xyBodyProfileBrief.Y
            PR_CalcPolar(xyBodyProfileBrief, aConstruct - 90, nArc, xyArc)
            aAngle = FN_CalcAngle(xyArc, xyBodyProfileBrief)
            aArcAngle = 20
            Dim X(3), Y(3) As Double
            cBriefPoly.X = X
            cBriefPoly.y = Y
            For ii = 1 To 3
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle + (aArcAngle / 2)
            Next ii

            nArc = (xyBodyLegPoint.X - xyBodySeamThigh.X) * (3 / 2)
            'PR_MakeXY xyArc, xyBodyLegPoint.X - nArc, xyBodyLegPoint.Y
            PR_CalcPolar(xyBodyLegPoint, aConstruct + 90, nArc, xyArc)
            aArcAngle = 15
            aAngle = FN_CalcAngle(xyArc, xyBodyLegPoint) + aArcAngle
            For ii = 1 To 3
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle - (aArcAngle / 2)
            Next ii

            PR_DrawFitted(cBriefPoly)

            cBriefPoly.X(1) = xyBodyLegPoint.X
            cBriefPoly.y(1) = xyBodyLegPoint.Y
            cBriefPoly.X(2) = xyBDProfileThighMirror1.X
            cBriefPoly.y(2) = xyBDProfileThighMirror1.Y
            If g_bBodyDiffThigh Then
                cBriefPoly.X(3) = xyBodyLT_ProfileThighMirror.X
                cBriefPoly.y(3) = xyBodyLT_ProfileThighMirror.Y
            Else
                cBriefPoly.X(3) = xyBDProfileThighMirror.X
                cBriefPoly.y(3) = xyBDProfileThighMirror.Y
            End If
            cBriefPoly.n = 3
            PR_DrawPoly(cBriefPoly)
        Else
            'Snap Crotch
            PR_DrawSnapCrotchArc()

            Dim X1(3), Y1(3) As Double
            cCrotchPoly.X = X1
            cCrotchPoly.y = Y1
            cCrotchPoly.X(1) = xyBodyStrap1.X
            cCrotchPoly.y(1) = xyBodyStrap1.Y
            cCrotchPoly.X(2) = xyBodyStrap2.X
            cCrotchPoly.y(2) = xyBodyStrap2.Y
            cCrotchPoly.X(3) = xyBodyStrap4.X
            cCrotchPoly.y(3) = xyBodyStrap4.Y
            cCrotchPoly.n = 3
            PR_DrawPoly(cCrotchPoly)

            Dim X2(4), Y2(4) As Double
            cBriefPoly.X = X2
            cBriefPoly.y = Y2
            cBriefPoly.X(1) = xyBodyStrap4.X
            cBriefPoly.y(1) = xyBodyStrap4.Y
            cBriefPoly.X(2) = xyBodyLegPoint.X
            cBriefPoly.y(2) = xyBodyLegPoint.Y
            cBriefPoly.X(3) = xyBDProfileThighMirror1.X
            cBriefPoly.y(3) = xyBDProfileThighMirror1.Y
            If g_bBodyDiffThigh Then
                cBriefPoly.X(4) = xyBodyLT_ProfileThighMirror.X
                cBriefPoly.y(4) = xyBodyLT_ProfileThighMirror.Y
            Else
                cBriefPoly.X(4) = xyBDProfileThighMirror.X
                cBriefPoly.y(4) = xyBDProfileThighMirror.Y
            End If
            cBriefPoly.n = 4
            PR_DrawPoly(cBriefPoly)
        End If
    End Sub
    Private Sub PR_DrawSnapCrotchArc()
        Dim aAngle As Double
        Dim xyTmp, xyTmp1 As BODYSUIT1.XY
        Dim nRadius As Double
        Dim cSnapCrotchPoly As curve
        Dim sLayer As String

        'Check if the arc is too close to the cut-out
        'it should be no closer than 1.25 inches
        nRadius = FN_CalcLength(xyBodyGussetArcCentre, xyBodyProfileBrief)

        If FN_CalcCirCirInt(xyBodyGussetArcCentre, nRadius, xyBodyCutOut5, 1.25, xyTmp, xyTmp1) Then
            'The arc is too close
            'draw as a fitted curve to allow for user adjustment
            MsgBox("The calculated arc for the Snap Crotch was too close to the cut-out. It will be drawn as fitted curve.  Adjust this as required.", 16, "BodySuit Dialog")
            cSnapCrotchPoly.n = 3
            Dim X(3), Y(3) As Double
            cSnapCrotchPoly.X = X
            cSnapCrotchPoly.y = Y
            cSnapCrotchPoly.X(1) = xyBodyProfileBrief.X
            cSnapCrotchPoly.y(1) = xyBodyProfileBrief.Y
            aAngle = FN_CalcAngle(xyBodyGussetArcCentre, xyBodyCutOut5)
            nRadius = FN_CalcLength(xyBodyGussetArcCentre, xyBodyCutOut5)
            PR_CalcPolar(xyBodyGussetArcCentre, aAngle, nRadius + 1.25, xyTmp)
            cSnapCrotchPoly.X(2) = xyTmp.X
            cSnapCrotchPoly.y(2) = xyTmp.Y
            cSnapCrotchPoly.X(cSnapCrotchPoly.n) = xyBodyStrap3.X
            cSnapCrotchPoly.y(cSnapCrotchPoly.n) = xyBodyStrap3.Y
            PR_DrawFitted(cSnapCrotchPoly)

            'Draw a circle on the layer construct for use by the user
            sLayer = BODYSUIT1.g_sCurrentLayer
            PR_SetLayer("Construct")
            PR_DrawCircle(xyBodyCutOut5, 1.25)
            PR_SetLayer(sLayer)
        Else
            'Draw arc as normal
            PR_DrawArc(xyBodyGussetArcCentre, xyBodyProfileBrief, xyBodyStrap3)
        End If

        'PR_DrawMarker xyBodyGussetArcCentre
        'PR_DrawText "xyBodyGussetArcCentre", xyBodyGussetArcCentre, .1
        'PR_DrawMarker xyBodyProfileBrief
        'PR_DrawText "xyBodyProfileBrief", xyBodyProfileBrief, .1
        'PR_DrawMarker xyBodyStrap3
        'PR_DrawText "xyBodyStrap3", xyBodyStrap3, .1
        'PR_DrawMarker xyBodyStrap1
        'PR_DrawText "xyBodyStrap1", xyBodyStrap1, .1
        'PR_DrawMarker xyBodyStrap2
        'PR_DrawText "xyBodyStrap2", xyBodyStrap2, .1
        'PR_DrawMarker xyBodyCutOut5
        'PR_DrawText "xyBodyCutOut5", xyBodyCutOut5, .1
        'PR_DrawMarker xyBodyCutOut6
        'PR_DrawText "xyBodyCutOut6", xyBodyCutOut6, .1
        'PR_DrawMarker xyBodyCutOut7
        'PR_DrawText "xyBodyCutOut7", xyBodyCutOut7, .1
        'PR_DrawMarker xyCutOut8
        'PR_DrawText "xyCutOut8", xyBodyCutOut7, .1

    End Sub
    Function FN_CalcCirCirInt(ByRef xyCen1 As BODYSUIT1.XY, ByRef nRad1 As Double, ByRef xyCen2 As BODYSUIT1.XY,
                              ByRef nRad2 As Double, ByRef xyInt1 As BODYSUIT1.XY, ByRef xyInt2 As BODYSUIT1.XY) As Short
        'Function that will return
        '    TRUE    if two circles intersect
        '    FALSE   if two circles don't intersect
        '
        'The intersection points are returned in the values
        '
        '    xyInt1 & xyInt2
        '
        'with intersection with lowest X value as xyInt1
        '
        Dim aTheta, nLength, aAngle, nCosTheta As Double
        Dim xyTmp As BODYSUIT1.XY

        FN_CalcCirCirInt = False

        'Check that the circles can intersect
        'It is a theorem of plane geometry that no three real numbers
        'a, b and c can be the lenghts of the sides of a triangle unless
        'the sum of any two is greater than the third.
        'We use this as our main test of possible intersection, we also check for silly
        'data.

        'Test for silly data
        If xyCen1.X = xyCen2.X And xyCen1.Y = xyCen2.Y Then Exit Function
        If nRad1 <= 0 Or nRad2 <= 0 Then Exit Function

        'Test for intersection
        nLength = FN_CalcLength(xyCen1, xyCen2)

        If (nLength + nRad1 < nRad2) Or (nLength + nRad2 < nRad1) Or (nRad1 + nRad2 < nLength) Then
            Exit Function
        Else
            FN_CalcCirCirInt = True
        End If

        'Calculate intesection points
        '
        'Special case where circles touch (ie Intersect at one point only)


        'Angle between centers
        'Note: Length between centers from above
        aAngle = FN_CalcAngle(xyCen1, xyCen2)

        'Get angle w.r.t line between centers to the intersection point
        'use cosine rule
        '
        nCosTheta = -((nRad2 ^ 2 - (nLength ^ 2 + nRad1 ^ 2)) / (2 * nLength * nRad1))

        aTheta = System.Math.Atan(-nCosTheta / System.Math.Sqrt(-(nCosTheta ^ 2) + 1)) + 1.5708

        aTheta = aTheta * (180 / BODYSUIT1.PI) 'convert to degrees

        aAngle = aAngle - aTheta
        PR_CalcPolar(xyCen1, aAngle, nRad1, xyInt1)

        aAngle = aAngle + aTheta
        PR_CalcPolar(xyCen1, aAngle, nRad1, xyInt2)

        If xyInt2.X < xyInt1.X Then
            xyTmp = xyInt1
            xyInt1 = xyInt2
            xyInt2 = xyTmp
        End If
    End Function
    Sub PR_DrawCircle(ByRef xyCen As BODYSUIT1.XY, ByRef nRadius As Object)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a CIRCLE at the point given.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '
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

            '' Create a circle that is at 2,3 with a radius of 4.25
            Using acCirc As Circle = New Circle()
                acCirc.Center = New Point3d(xyCen.X + xyInsertion.X, xyCen.Y + xyInsertion.Y, 0)
                acCirc.Radius = nRadius
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acCirc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acCirc)
                acTrans.AddNewlyCreatedDBObject(acCirc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawBriefFrenchCut()
        Dim cBriefPoly As curve
        Dim cCrotchPoly As curve
        Dim nArc As Double
        Dim xyArc As BODYSUIT1.XY
        Dim aAngle, aArcAngle, aConstruct As Double
        Dim ii As Short
        Dim xyTmp As BODYSUIT1.XY

        If g_bBodyDiffThigh Then
            PR_DrawLine(xyBodyLT_ProfileThigh, xyBodyProfileBrief)
            PR_DrawLine(xyBDProfileThighMirror1, xyBodyLT_ProfileThighMirror)
        Else
            PR_DrawLine(xyBodyProfileThigh, xyBodyProfileBrief)
            PR_DrawLine(xyBDProfileThighMirror1, xyBDProfileThighMirror)
        End If
        cBriefPoly.n = 0

        If g_sBodyCrotchStyle <> "Snap Crotch" Then
            'All other crotch styles
            cBriefPoly.n = 0
            aConstruct = FN_CalcAngle(xyBodyLegPoint, xyBodyProfileBrief)
            aConstruct = aConstruct - ((aConstruct - 90) / 2)
            nArc = (xyBodyCutOutArcCentre.X - xyBodySeamThigh.X) * (2 / 3)
            ' PR_MakeXY xyArc, xyBodyProfileBrief.X + nArc, xyBodyProfileBrief.Y
            PR_CalcPolar(xyBodyProfileBrief, aConstruct - 90, nArc, xyArc)
            aAngle = FN_CalcAngle(xyArc, xyBodyProfileBrief)
            aArcAngle = 20
            Dim X(4), Y(4) As Double
            cBriefPoly.X = X
            cBriefPoly.y = Y
            For ii = 1 To 3
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle + (aArcAngle / 2)
            Next ii

            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyBodyLegPoint.X
            cBriefPoly.y(cBriefPoly.n) = xyBodyLegPoint.Y

            nArc = (xyFrenchCut.x - xyBodySeamFold.X)
            If Not g_nAdult Then
                nArc = nArc * 0.5
                aAngle = 40
                aArcAngle = 80
            Else
                aAngle = 45
                aArcAngle = 90
            End If
            PR_MakeXY(xyArc, xyFrenchCut.x - nArc, xyFrenchCut.y)
            Dim X1(5), Y1(5) As Double
            cBriefPoly.X = X1
            cBriefPoly.y = Y1
            For ii = 1 To 5
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle - (aArcAngle / 4)
            Next ii

            nArc = (xyBodyCutOutArcCentre.X - xyBodySeamThigh.X) * (2 / 3)
            If Not g_nAdult Then
                nArc = nArc * 0.25
                PR_MakeXY(xyArc, xyBDProfileThighMirror.X + nArc, xyBDProfileThighMirror.Y)
                nArc = FN_CalcLength(xyBDProfileThighMirror1, xyArc)
                aArcAngle = 40
                aAngle = FN_CalcAngle(xyBDProfileThighMirror1, xyArc) - aArcAngle
            Else
                PR_MakeXY(xyArc, xyBDProfileThighMirror1.X + nArc, xyBDProfileThighMirror1.Y)
                aArcAngle = 40
                aAngle = 180 - aArcAngle
            End If
            PR_MakeXY(xyArc, xyBDProfileThighMirror1.X + nArc, xyBDProfileThighMirror1.Y)
            aArcAngle = 40
            aAngle = 180 - aArcAngle
            For ii = 1 To 5
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle + (aArcAngle / 4)
            Next ii

            PR_DrawFitted(cBriefPoly)

        Else
            'Snap Crotch
            PR_DrawSnapCrotchArc()
            Dim XCro(3), YCro(3) As Double
            cCrotchPoly.X = XCro
            cCrotchPoly.y = YCro
            cCrotchPoly.X(1) = xyBodyStrap1.X
            cCrotchPoly.y(1) = xyBodyStrap1.Y
            cCrotchPoly.X(2) = xyBodyStrap2.X
            cCrotchPoly.y(2) = xyBodyStrap2.Y
            cCrotchPoly.X(3) = xyBodyStrap4.X
            cCrotchPoly.y(3) = xyBodyStrap4.Y
            cCrotchPoly.n = 3
            PR_DrawPoly(cCrotchPoly)

            PR_DrawLine(xyBodyStrap4, xyBodyLegPoint)

            Dim XBr(cBriefPoly.n + 1), YBr(cBriefPoly.n + 1) As Double
            cBriefPoly.X = XBr
            cBriefPoly.y = YBr
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyBodyLegPoint.X
            cBriefPoly.y(cBriefPoly.n) = xyBodyLegPoint.Y

            nArc = (xyFrenchCut.x - xyBodySeamFold.X)
            If Not g_nAdult Then
                nArc = nArc * 0.5
                aAngle = 40
                aArcAngle = 80
            Else
                aAngle = 45
                aArcAngle = 90
            End If
            PR_MakeXY(xyArc, xyFrenchCut.x - nArc, xyFrenchCut.y)
            aAngle = 45
            aArcAngle = 90
            Dim XBr1(5), YBr1(5) As Double
            cBriefPoly.X = XBr1
            cBriefPoly.y = YBr1
            For ii = 1 To 5
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle - (aArcAngle / 4)
            Next ii

            nArc = (xyBodyCutOutArcCentre.X - xyBodySeamThigh.X) * (2 / 3)
            If Not g_nAdult Then
                nArc = nArc * 0.25
                PR_MakeXY(xyArc, xyBDProfileThighMirror.X + nArc, xyBDProfileThighMirror.Y)
                nArc = FN_CalcLength(xyBDProfileThighMirror1, xyArc)
                aArcAngle = 40
                aAngle = FN_CalcAngle(xyBDProfileThighMirror1, xyArc) - aArcAngle
            Else
                PR_MakeXY(xyArc, xyBDProfileThighMirror1.X + nArc, xyBDProfileThighMirror1.Y)
                aArcAngle = 40
                aAngle = 180 - aArcAngle
            End If

            PR_MakeXY(xyArc, xyBDProfileThighMirror1.X + nArc, xyBDProfileThighMirror1.Y)
            aArcAngle = 40
            aAngle = 180 - aArcAngle
            For ii = 1 To 5
                PR_CalcPolar(xyArc, aAngle, nArc, xyTmp)
                cBriefPoly.n = cBriefPoly.n + 1
                cBriefPoly.X(cBriefPoly.n) = xyTmp.X
                cBriefPoly.y(cBriefPoly.n) = xyTmp.Y
                aAngle = aAngle + (aArcAngle / 4)
            Next ii

            PR_DrawFitted(cBriefPoly)
        End If
    End Sub
    Private Sub PR_DrawBra()
        'Assumes that data has been validated before this
        'procedure is called.
        Dim nDiskXoff, nBraCLOffset, nDiskYoff As Double
        Dim iRightDisk, iDisk, ii, iLeftDisk, iNextDisk As Short
        Dim sBraText, sDisk As String
        Dim xyDisk As BODYSUIT1.XY
        Dim xyPt(3) As BODYSUIT1.XY

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
                        nBraCLOffset = 1.0#
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
                    PR_SetLayer("Notes")
                    xyDisk.X = (xySeamChest.x - 0.75) - nDiskXoff
                    xyDisk.Y = (xyCutOut7.y - nBraCLOffset) '- (nDiskYoff * 2)

                    sDisk = "bradisk" & Trim(Str(iDisk)) & "mirrored"
                    PR_InsertSymbol(sDisk, xyDisk, 1, 0)
                    PR_AddDBValueToLast("curvetype", "Bra")

                    'Draw Construction lines
                    PR_SetLayer("Construct")
                    PR_MakeXY(xyPt(1), xySeamChest.x - 0.75, xyCutOut7.y - nBraCLOffset)
                    PR_MakeXY(xyPt(2), xyPt(1).X, (xyCutOut7.y - nBraCLOffset) - (nDiskYoff * 2))
                    PR_MakeXY(xyPt(3), xyPt(1).X - (nDiskXoff * 2), xyPt(1).Y)
                    PR_DrawLine(xyPt(1), xyPt(2))
                    PR_AddDBValueToLast("curvetype", "Bra")
                    PR_DrawLine(xyPt(1), xyPt(3))
                    PR_AddDBValueToLast("curvetype", "Bra")
                    If ii = iNextDisk Then
                        'Insert bra label text
                        PR_SetLayer("Notes")
                        PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
                        PR_CalcMidPoint(xyPt(2), xyPt(3), xyPt(1))
                        PR_DrawText(sBraText, xyPt(1), 0.125, 0, 1)
                    End If
                Else
                    'No Bra cup.
                    'Do nothing as this is dealt with as text.
                    'Except for this very special case
                    If iNextDisk = 1 Then
                        'Insert Text (Special case No Cups Left and Right ????)
                        PR_SetLayer("Notes")
                        PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
                        PR_MakeXY(xyPt(1), xySeamChest.x - 2.75, xySeamChest.y - 2.5)
                        PR_DrawText(sBraText, xyPt(1), 0.125, 0, 1)
                    End If
                End If
            Next ii
        End If
    End Sub
    Private Sub PR_AssignValuesToDlgCtrls()
        Dim ii As Short

        'Stop the timer used to ensure that the Dialogue dies
        'if the DRAFIX macro fails to establish a DDE Link
        Timer1.Enabled = False

        'Check that a "MainPatientDetails" Symbol has been
        'found
        'If txtUidMPD.Text = "" Then
        '    MsgBox("No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue")
        '    'End  '''J'''
        'End If

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
                BODYSUIT1.g_sLeftCup = VB6.GetItemString(cboLeftCup, ii)
                cboLeftCup.SelectedIndex = ii
            End If
        Next ii
        For ii = 0 To (cboRightCup.Items.Count - 1)
            If txtCombo(4).Text = VB6.GetItemString(cboRightCup, ii) Then
                BODYSUIT1.g_sRightCup = VB6.GetItemString(cboRightCup, ii)
                cboRightCup.SelectedIndex = ii
            End If
        Next ii

        BODYSUIT1.g_sLeftDisk = txtLeftDisk.Text
        BODYSUIT1.g_sRightDisk = txtRightDisk.Text

        If txtCombo(5).Text <> "" Then cboLeftAxilla.Text = txtCombo(5).Text Else cboLeftAxilla.SelectedIndex = 5
        If txtCombo(6).Text <> "" Then cboRightAxilla.Text = txtCombo(6).Text Else cboRightAxilla.SelectedIndex = 5
        If txtCombo(7).Text <> "" Then cboFrontNeck.Text = txtCombo(7).Text Else cboFrontNeck.SelectedIndex = 0
        If txtCombo(8).Text <> "" Then cboBackNeck.Text = txtCombo(8).Text Else cboBackNeck.SelectedIndex = 0
        If txtCombo(9).Text <> "" Then cboClosure.Text = txtCombo(9).Text Else cboClosure.SelectedIndex = 0
        cboFabric.Text = txtCombo(10).Text

        'Set up units
        If txtUnits.Text = "cm" Then
            BODYSUIT1.g_nUnitsFac = 10 / 25.4
        Else
            BODYSUIT1.g_nUnitsFac = 1
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
        g_nBodyChest = Val(txtCir(5).Text)
        g_nBodyUnderBreast = Val(txtCir(10).Text)
        g_nBodyNipple = Val(txtCir(11).Text)

        'Store values to use on change etc
        Dim sChar As New VB6.FixedLengthString(1)
        g_sBodyBackNeck = txtBackNeck.Text
        sChar.Value = VB.Left(cboBackNeck.Text, 1)
        If sChar.Value = "R" Or sChar.Value = "S" Or sChar.Value = "V" Then
            'Disable length box for Regular, Scoop, V neck
            txtBackNeck.Enabled = False
            labBackNeck.Enabled = False
        Else
            txtBackNeck.Enabled = True
            labBackNeck.Enabled = True
        End If

        BODYSUIT1.g_sFrontNeck = txtFrontNeck.Text
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
        cboCrotchStyle.SelectedIndex = 4

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
        BODYSUIT1.g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
        For ii = 0 To (cboLegStyle.Items.Count - 1)
            If VB6.GetItemString(cboLegStyle, ii) = fnGetString(txtLegStyle.Text, 1) Then
                cboLegStyle.SelectedIndex = ii
            End If
        Next ii

        'Set this now to avoid problems with cboLegStyle_Click
        If cboLegStyle.SelectedIndex = -1 Then
            BODYSUIT1.g_sPreviousLegStyle = "None"
        Else
            BODYSUIT1.g_sPreviousLegStyle = VB6.GetItemString(cboLegStyle, cboLegStyle.SelectedIndex)
        End If

        If fnGetString(txtLegStyle.Text, 2) <> "" Then
            If fnGetString(txtLegStyle.Text, 2) = "Panty" Then
                optLeftLeg(0).Checked = True
            Else
                optLeftLeg(1).Checked = True 'Brief
            End If
        End If
        If fnGetString(txtLegStyle.Text, 3) <> "" Then
            If fnGetString(txtLegStyle.Text, 3) = "Panty" Then
                optRightLeg(0).Checked = True
            Else
                optRightLeg(1).Checked = True 'Brief
            End If
        End If
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
        BODYSUIT1.g_sChangeChecker = FN_ValuesString()
    End Sub
    Public Sub PR_DrawBody()
        g_strSide = "Left"
        m_nLeftKneePosition = 0
        m_nRightKneePosition = 0
        idLastCreated = New ObjectId
        'Initialize globals
        BODYSUIT1.QQ = Chr(34) 'Double quotes (")
        BODYSUIT1.NL = Chr(13) 'New Line
        BODYSUIT1.CC = Chr(44) 'The comma (,)
        BODYSUIT1.QCQ = BODYSUIT1.QQ & BODYSUIT1.CC & BODYSUIT1.QQ
        BODYSUIT1.QC = BODYSUIT1.QQ & BODYSUIT1.CC
        BODYSUIT1.CQ = BODYSUIT1.CC & BODYSUIT1.QQ

        BODYSUIT1.g_nUnitsFac = 1
        BODYSUIT1.g_sPathJOBST = fnPathJOBST()

        'Clear fields
        'Circumferences and lengths
        Dim ii As Short
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

        cboLeftAxilla.Items.Add("Regular 2" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Regular 1½" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Regular 2½" & BODYSUIT1.QQ)
        cboLeftAxilla.Items.Add("Open")
        cboLeftAxilla.Items.Add("Mesh")
        cboLeftAxilla.Items.Add("Lining")
        cboLeftAxilla.Items.Add("Sleeveless")

        cboRightAxilla.Items.Add("Regular 2" & BODYSUIT1.QQ)
        cboRightAxilla.Items.Add("Regular 1½" & BODYSUIT1.QQ)
        cboRightAxilla.Items.Add("Regular 2½" & BODYSUIT1.QQ)
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

        ''------------LGLEGDIA1.PR_GetComboListFromFile(cboFabric, BODYSUIT1.g_sPathJOBST & "\FABRIC.DAT")
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

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
        txtWorkOrder.Text = workOrder
        PR_AssignValuesToDlgCtrls()
        readDWGInfo()
        PR_DrawBodySuit()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PR_DrawBodySuit()
    End Sub
    Private Sub PR_DrawBodySuit()
        If FN_ValidateAndCalculateData(False) = False Then
            Exit Sub
        End If
        Me.Hide()
        If PR_GetInsertionPoint() = False Then
            Exit Sub
        End If
        PR_GetBlkAttributeValues()
        If PR_CalculateBodySuit() = False Then
            Exit Sub
            Me.Show()
        End If
        Static sText As String
        Static xyText As BODYSUIT1.XY
        Static xyZipMark(4) As BODYSUIT1.XY
        Static xyTmp As BODYSUIT1.XY
        Static xyTmp1 As BODYSUIT1.XY
        Static aAngle As Double
        Static aInc As Double
        Static ii As Short
        Static nLength As Double
        Static nOffset As Double
        Static xyOpen(4) As BODYSUIT1.XY
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
        Static xySeamLineStart As BODYSUIT1.XY
        Static xySeamLineEnd As BODYSUIT1.XY
        Static xyCentreLineStart As BODYSUIT1.XY
        Static xyCentreLineEnd As BODYSUIT1.XY
        Static xyBackNeckMid1 As BODYSUIT1.XY
        Static xyBackNeckMid2 As BODYSUIT1.XY
        Static xyFrontNeckMid1 As BODYSUIT1.XY
        Static xyFrontNeckMid2 As BODYSUIT1.XY
        Static xyGroinExtraPT As BODYSUIT1.XY
        Static xyGroinExtraPT1 As BODYSUIT1.XY
        Static xyTmpRaglanCentre As BODYSUIT1.XY
        Static nBackNeckMidOffsetX As Double
        Static nBackNeckMidOffsetY As Double
        Static nFlip As Short
        Static xyCrotchText As BODYSUIT1.XY
        Static sCrotchText As String
        Static xyGussetOffCut As BODYSUIT1.XY
        Static xyLLGussetRectangle As BODYSUIT1.XY
        Static xyURGussetRectangle As BODYSUIT1.XY
        Static xyNeckText As BODYSUIT1.XY
        Static xyCutOutArcStart As BODYSUIT1.XY
        Static xyCutOutArcEnd As BODYSUIT1.XY
        Static xyProfileThighRelease As BODYSUIT1.XY
        Static xyLegStart As BODYSUIT1.XY
        Static xyMesh As BODYSUIT1.XY

        Static cProfilePoly As curve
        Static cRaglanPoly As curve
        Static cRaglanCurve As curve
        Static cFrontCutOutPoly As curve
        Static cBackCutOutPoly As curve
        Static cFrontNeckArc As curve
        Static cBackNeckArc As curve
        Static sSmallestElasticText As String
        Static sLargestElasticText As String

        cProfilePoly.Initialize()
        cRaglanPoly.Initialize()
        cRaglanCurve.Initialize()
        cFrontCutOutPoly.Initialize()
        cBackCutOutPoly.Initialize()
        cFrontNeckArc.Initialize()
        cBackNeckArc.Initialize()

        Static MeshProfile As BiArc
        Static MeshSeam As BiArc
        Static nMeshSize As Double
        Static nMeshDistanceAlongRaglan As Double

        Static SmallestLegCurve As curve
        Static LargestLegCurve As curve
        SmallestLegCurve.Initialize()
        LargestLegCurve.Initialize()
        Static sLeg As String

        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        BODYSUIT1.fNum = FN_Open(sDrawFile, "Draw", txtPatientName.Text, txtFileNo.Text)
        'Add labels and construction lines
        nVertLinesOffset = xyBodyProfileButt.Y + 1
        nHorLinesOffset = g_nShldToFold * 0.05
        ARMDIA1.PR_SetLayer("Construct")
        If Not g_sLegStyle = "Panty" Then PR_DrawVerticalLineThruPoint(xyBodySeamThigh, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamFold, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamButt, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamWaist, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamLowShld, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamHighShld, nVertLinesOffset)
        PR_DrawVerticalLineThruPoint(xyBodySeamChest, nVertLinesOffset)

        PR_MakeXY(xySeamLineStart, 0 - (ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), BODY_INCH1_4)
        PR_MakeXY(xySeamLineEnd, xyBodySeamHighShld.X + nHorLinesOffset, BODY_INCH1_4)
        PR_MakeXY(xyCentreLineStart, 0 - (ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), 0)
        PR_MakeXY(xyCentreLineEnd, xyBodySeamHighShld.X + nHorLinesOffset, 0)

        'Useful Points
        PR_MakeXY(xyBackNeckMid1, xyBodyCutOutBackNeck.X + ((xyBodyProfileNeck.X - xyBodyCutOutBackNeck.X) / 6), xyBodyCutOutBackNeck.Y + ((xyBodyProfileNeck.Y - xyBodyCutOutBackNeck.Y) / 3))
        PR_MakeXY(xyBackNeckMid2, xyBodyProfileNeck.X - ((xyBodyProfileNeck.X - xyBodyCutOutBackNeck.X) / 2), xyBodyProfileNeck.Y - ((xyBodyProfileNeck.Y - xyBodyCutOutBackNeck.Y) / 3))
        PR_MakeXY(xyFrontNeckMid1, xyBodyCutOutFrontNeck.X - ((xyBodyCutOutFrontNeck.X - xyBodyCutOut8.X) / 3), xyBodyCutOutFrontNeck.Y + ((xyBodyCutOut8.Y - xyBodyCutOutFrontNeck.Y) / 6))
        PR_MakeXY(xyFrontNeckMid2, xyBodyCutOut8.X + ((xyBodyCutOutFrontNeck.X - xyBodyCutOut8.X) / 3), xyBodyCutOut8.Y - ((xyBodyCutOut8.Y - xyBodyCutOutFrontNeck.Y) / 2))
        PR_CalcPolar(xyBDButtocksArcCentre, 45, g_nButtRadius, xyCutOutArcStart)
        PR_CalcPolar(xyBDButtocksArcCentre, 135, g_nButtRadius, xyCutOutArcEnd)
        PR_MakeXY(xyGroinExtraPT, xyBodySeamFold.X + (g_nButtocksLength / 2), xyBodySeamFold.Y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength / 2) ^ 2)))
        PR_MakeXY(xyGroinExtraPT1, xyBodySeamFold.X + ((g_nButtocksLength / 4) * 3), xyBodySeamFold.Y + g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength / 4) ^ 2)))

        'Calculated for manual editing
        PR_MakeXY(xyProfileThighRelease, xyBodySeamThigh.X, xyBodySeamThigh.Y + g_nThighCir) 'this is the only place that g_nThighCir is used
        PR_DrawMarker(xyProfileThighRelease)
        PR_AddDBValueToLast("Data", "Thigh")
        PR_DrawMarker(xyBody85Fold)
        PR_AddDBValueToLast("Data", "At original 85% reduction")
        'Draw Construction Lines
        PR_DrawLine(xySeamLineStart, xySeamLineEnd)
        PR_DrawLine(xyCentreLineStart, xyCentreLineEnd)
        PR_AddDBValueToLast("curvetype", "CentreLine")
        'Draw Buttocks Arc for reference
        PR_DrawArc(xyBDButtocksArcCentre, xyCutOutArcStart, xyCutOutArcEnd)

        'Insert Key Markers for reference
        PR_DrawMarker(xyBodyProfileThigh)
        PR_DrawMarker(xyBodyProfileGroin)
        PR_DrawMarker(xyBodyProfileButt)
        PR_DrawMarker(xyBodyProfileWaist)
        'YES - 2 at same point (one with DB value)!!!!
        PR_DrawMarker(xyBodyProfileChest)
        PR_DrawMarker(xyBodyProfileChest)
        PR_AddDBValueToLast("curvetype", "BackCurveMarker")
        PR_DrawMarker(xyBodyProfileNeck)
        PR_DrawMarker(xyBodyCutOutArcCentre)
        PR_DrawMarker(xyBDFrontNeckArcCentre)
        PR_DrawMarker(xyBodyFold)
        PR_DrawMarker(xyBDButtocksArcCentre)

        'Draw DB markers for Zipper macros
        PR_DrawMarker(xyBodyCutOutBackNeck)
        PR_AddDBValueToLast("Zipper", "BackN")
        PR_DrawMarker(xyBodyCutOut2)
        PR_AddDBValueToLast("Zipper", "BackW")
        PR_DrawMarker(xyBodyCutOut3)
        PR_AddDBValueToLast("Zipper", "BackB")
        PR_DrawMarker(xyBodyCutOut5)
        PR_AddDBValueToLast("Zipper", "FrontB")
        PR_AddDBValueToLast("Data", Str(g_nLengthStrap1ToCutOut5))

        PR_DrawMarker(xyBodyCutOut6)
        PR_AddDBValueToLast("Zipper", "FrontW")
        PR_DrawMarker(xyBodyCutOut8)
        PR_AddDBValueToLast("Zipper", "FrontN")
        If g_sBodyFrontNeckStyle = "Turtle" Then
            PR_AddDBValueToLast("Data", Str(g_nFrontNeckSize))
        End If

        'Other patient details in black on layer construct
        PR_MakeXY(xyText, xyBodySeamButt.X + 1, xyBodySeamButt.Y - 1.8)
        sText = g_sFileNo & Chr(10) & g_sDiagnosis & Chr(10) & Trim(Str(g_nAge)) & Chr(10) & g_sSex
        ''PR_DrawText(sText, xyText, 0.125, 0, 1)
        PR_DrawMText(sText, xyText, 0)

        'Notes Text - e.g. Patient Details, Sewing instructions
        PR_SetLayer("Notes")
        PR_MakeXY(xyText, xyBodySeamButt.X + 1, xyBodySeamButt.Y - 1)
        sText = g_sPatient & Chr(10) & g_sWorkOrder & Chr(10) & Trim(Mid(g_sFabric, 4))
        ''PR_DrawText(sText, xyText, 0.125, 0, 1)
        PR_DrawMText(sText, xyText, 0)

        'Draw Measured Front Neck Text
        If InStr(g_sBodyFrontNeckStyle, "Scoop") > 0 Then
            sText = "SCOOP NECK"
            PR_MakeXY(xyNeckText, xyBodyCutOutFrontNeck.X - 3, xyBodyCutOutFrontNeck.Y + 0.25)
        ElseIf InStr(g_sBodyFrontNeckStyle, "V neck") > 0 Then
            sText = "V NECK"
            PR_MakeXY(xyNeckText, xyBodyCutOutFrontNeck.X - 3, xyBodyCutOutFrontNeck.Y + 0.25)
        ElseIf InStr(g_sBodyFrontNeckStyle, "Turtle") > 0 Then
            sText = "NECK CIRC. " & fnInchestoText(g_nNeckCir) & "\""" & "\nTURTLE NECK WIDTH " & fnInchestoText(g_nFrontNeckSize) & "\"""
            PR_MakeXY(xyNeckText, xyBodyCutOutFrontNeck.X - 3.5, xyBodyCutOutFrontNeck.Y + 0.5)
            If InStr(g_sBodyFrontNeckStyle, "same") Then
                If g_sLegStyle = "Brief" Then
                    sText = sText & "\nIN SAME FABRIC AS BRIEF"
                Else
                    sText = sText & "\nIN SAME FABRIC AS SUIT"
                End If
            ElseIf InStr(g_sBodyFrontNeckStyle, "Detach. Fabric") Then
                sText = sText & "\nDETACHABLE FABRIC"
            Else
                sText = sText & "\nDETACHABLE"
            End If

        Else 'Regular
            sText = "NECK CIRC. " & fnInchestoText(g_nNeckCir) & BODYSUIT1.QQ
            PR_MakeXY(xyNeckText, xyBodyCutOutFrontNeck.X - 3, xyBodyCutOutFrontNeck.Y + 0.25)
        End If
        If InStr(g_sBodyFrontNeckStyle, "Measured") > 0 Then sText = "MEASURED " & sText
        ''PR_DrawText(sText, xyNeckText, 0.125, 0, 1)
        PR_DrawMText(sText, xyNeckText, 0)

        'Draw Measured Back Neck Text
        If InStr(g_sBodyBackNeckStyle, "Scoop") > 0 Then
            If g_sBodyBackNeckStyle = "Measured Scoop" Then
                sText = "MEASURED SCOOP NECK"
            Else
                sText = "SCOOP NECK"
            End If
            PR_MakeXY(xyNeckText, xyBodyCutOutBackNeck.X - 3, xyBodyCutOutBackNeck.Y + 0.25)
            PR_DrawText(sText, xyNeckText, 0.125, 0, 1)
        End If

        'Draw Zipper
        If g_sClosure = "Back Zip" Then
            'Back zipper
            PR_CalcMidPoint(xyBodyCutOut3, xyBodyCutOut2, xyZipMark(1))
            nLength = FN_CalcLength(xyBodyCutOut2, xyBodyCutOutBackNeck) + FN_CalcLength(xyBodyCutOut3, xyZipMark(1))
            aAngle = FN_CalcAngle(xyBodyCutOut3, xyZipMark(1))
            PR_MakeXY(xyZipMark(2), xyZipMark(1).X, xyZipMark(1).Y + 0.125)
            PR_CalcPolar(xyZipMark(1), aAngle, 0.5, xyZipMark(3))
            PR_CalcMidPoint(xyBodyCutOut2, xyBodyCutOutBackNeck, xyZipMark(4)) 'Text insertion point
            xyZipMark(4).Y = xyZipMark(4).Y + 0.125
        End If

        Static nLen As Double
        If g_sClosure = "Front Zip" Then
            'Front Zipper
            If g_bMissCutOut7 Then
                nLength = FN_CalcLength(xyBodyCutOut8, xyBodyCutOut6)
                PR_CalcMidPoint(xyBodyCutOut6, xyBodyCutOut8, xyZipMark(4)) 'Text insertion point
            Else
                nLength = FN_CalcLength(xyBodyCutOut8, xyBodyCutOut7) + FN_CalcLength(xyBodyCutOut7, xyBodyCutOut6)
                PR_CalcMidPoint(xyBodyCutOut6, xyBodyCutOut7, xyZipMark(4)) 'Text insertion point
            End If
            nLength = nLength + FN_CalcLength(xyBodyCutOut5, xyBodyCutOut6)

            xyZipMark(4).Y = xyZipMark(4).Y - 0.25

            aAngle = FN_CalcAngle(xyBodyCutOut5, xyBodyCutOut6)
            If g_sBodyCrotchStyle = "Snap Crotch" Then
                If g_nLengthStrap1ToCutOut5 > 0 Then
                    'This takes care of the special case where the snap crotch lands
                    'before xyBodyCutOut5
                    PR_CalcPolar(xyBodyCutOut5, aAngle, 1.125 - g_nLengthStrap1ToCutOut5, xyZipMark(1))
                    nLength = nLength - FN_CalcLength(xyBodyCutOut5, xyZipMark(1))
                Else
                    'We also need to take care of another special case
                    'where the 1.25" along the cut-out to the start of the
                    'zip mark takes it around the corner (so to speak)
                    nLen = FN_CalcLength(xyBodyStrap1, xyBodyCutOut6)
                    If nLen >= 1.625 Then
                        'Simple case Zip mark is on the line (xyBodyCutOut5, xyBodyCutOut6)
                        PR_CalcPolar(xyBodyStrap1, aAngle, 1.125, xyZipMark(1))
                        nLength = nLength - FN_CalcLength(xyBodyCutOut5, xyZipMark(1))
                    ElseIf nLen > 1.125 And nLen < 1.625 Then
                        'Zip mark is round the corner
                        PR_CalcPolar(xyBodyStrap1, aAngle, 1.125, xyZipMark(1))
                        nLength = nLength - FN_CalcLength(xyBodyCutOut5, xyZipMark(1))
                        aAngle = FN_CalcAngle(xyBodyCutOut6, xyBodyCutOut7)
                    Else
                        'Zip mark is on the line (xyBodyCutOut6, xyBodyCutOut7)
                        nLen = 1.125 - nLen
                        aAngle = FN_CalcAngle(xyBodyCutOut6, xyBodyCutOut7)
                        PR_CalcPolar(xyBodyCutOut6, aAngle, nLen, xyZipMark(1))
                        nLength = nLength - FN_CalcLength(xyBodyCutOut5, xyBodyCutOut6) - nLen
                    End If
                End If
            Else
                PR_CalcMidPoint(xyBodyCutOut5, xyBodyCutOut6, xyZipMark(1))
                nLength = nLength - (FN_CalcLength(xyBodyCutOut5, xyBodyCutOut6) / 2)
            End If
            PR_MakeXY(xyZipMark(2), xyZipMark(1).X, xyZipMark(1).Y - 0.125)
            PR_CalcPolar(xyZipMark(1), aAngle, 0.5, xyZipMark(3))
        End If

        'Draw the zip and text
        If g_sClosure = "Back Zip" Or g_sClosure = "Front Zip" Then
            PR_DrawLine(xyZipMark(1), xyZipMark(2))
            PR_AddDBValueToLast("Zipper", "1")
            PR_DrawLine(xyZipMark(2), xyZipMark(3))
            PR_AddDBValueToLast("Zipper", "1")
            ''PR_InsertSymbol("TextAsSymbol", xyZipMark(4), 1, 0)
            'Adjust length for turtle necks
            If g_sBodyFrontNeckStyle = "Turtle" Then
                nLength = nLength + g_nFrontNeckSize
            End If
            PR_AddDBValueToLast("Zipper", "1")
            nLength = (nLength - 0.125) / 0.95
            sText = fnInchestoText(nLength) & Chr(34) & " " & g_sClosure
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                sText = fnDisplayToCM(nLength) & " " & g_sClosure
            End If
            PR_DrawTextAsSymbol(xyZipMark(4), txtFileNo.Text + g_strSide, sText, 0)
            PR_AddFormatedDBValueToLast("Data", "", nLength, " " & g_sClosure)
            PR_AddDBValueToLast("Zipper", "1")
            PR_AddDBValueToLast("ID", txtFileNo.Text + g_strSide)
            PR_AddDBValueToLast("Data", sText)
        End If

        'Bracups
        If g_sSex = "Female" Then
            PR_DrawBra()
        End If

        'Draw Body Suit/Brief
        PR_SetLayer("Template" & g_strSide)

        'Draw Front Neck
        'Temporary 3 point polyline - will use arc for all options later
        If g_sBodyFrontNeckStyle = "Measured Scoop" Or g_sBodyFrontNeckStyle = "Measured V neck" Then
            cFrontNeckArc.X(1) = xyBodyCutOutFrontNeck.X
            cFrontNeckArc.y(1) = xyBodyCutOutFrontNeck.Y
            cFrontNeckArc.X(2) = xyFrontNeckMid1.X
            cFrontNeckArc.y(2) = xyFrontNeckMid1.Y
            cFrontNeckArc.X(3) = xyFrontNeckMid2.X
            cFrontNeckArc.y(3) = xyFrontNeckMid2.Y
            cFrontNeckArc.X(4) = xyBodyCutOut8.X
            cFrontNeckArc.y(4) = xyBodyCutOut8.Y
            cFrontNeckArc.n = 4
            PR_DrawFitted(cFrontNeckArc)
        Else
            PR_DrawArc(xyBDFrontNeckArcCentre, xyBodyCutOut8, xyBodyCutOutFrontNeck)
        End If

        'Draw Cutout
        If g_bMissCutOut7 Then
            cFrontCutOutPoly.X(1) = xyBodyCutOut5.X
            cFrontCutOutPoly.y(1) = xyBodyCutOut5.Y
            cFrontCutOutPoly.X(2) = xyBodyCutOut6.X
            cFrontCutOutPoly.y(2) = xyBodyCutOut6.Y
            cFrontCutOutPoly.X(3) = xyBodyCutOut8.X
            cFrontCutOutPoly.y(3) = xyBodyCutOut8.Y
            cFrontCutOutPoly.n = 3
            PR_DrawPoly(cFrontCutOutPoly)
            PR_AddDBValueToLast("curvetype", "CutOutCurve")
            PR_AddDBValueToLast("Zipper", "Front")
            PR_AddDBValueToLast("ZipperLength", "%Length")
            'Draw DB markers for Zipper macros
            PR_DrawMarker(xyBodyCutOut8) 'Position this Marker at the neck for simplicity
            PR_AddDBValueToLast("Zipper", "FrontC")
        Else
            cFrontCutOutPoly.X(1) = xyBodyCutOut5.X
            cFrontCutOutPoly.y(1) = xyBodyCutOut5.Y
            cFrontCutOutPoly.X(2) = xyBodyCutOut6.X
            cFrontCutOutPoly.y(2) = xyBodyCutOut6.Y
            cFrontCutOutPoly.X(3) = xyBodyCutOut7.X
            cFrontCutOutPoly.y(3) = xyBodyCutOut7.Y
            cFrontCutOutPoly.X(4) = xyBodyCutOut8.X
            cFrontCutOutPoly.y(4) = xyBodyCutOut8.Y
            cFrontCutOutPoly.n = 4
            PR_DrawPoly(cFrontCutOutPoly)
            PR_AddDBValueToLast("curvetype", "CutOutCurve")
            PR_AddDBValueToLast("Zipper", "Front")
            PR_AddDBValueToLast("ZipperLength", "%Length")
            PR_DrawMarker(xyBodyCutOut7)
            PR_DrawMarker(xyBodyCutOut7)
            PR_AddDBValueToLast("curvetype", "CutOutMarker")
            'Draw DB markers for Zipper macros
            PR_DrawMarker(xyBodyCutOut7)
            PR_AddDBValueToLast("Zipper", "FrontC")
        End If

        cBackCutOutPoly.X(1) = xyBodyCutOut3.X
        cBackCutOutPoly.y(1) = xyBodyCutOut3.Y
        cBackCutOutPoly.X(2) = xyBodyCutOut2.X
        cBackCutOutPoly.y(2) = xyBodyCutOut2.Y
        cBackCutOutPoly.X(3) = xyBodyCutOutBackNeck.X
        cBackCutOutPoly.y(3) = xyBodyCutOutBackNeck.Y
        cBackCutOutPoly.n = 3
        PR_DrawFitted(cBackCutOutPoly)
        PR_AddDBValueToLast("Zipper", "Back")
        PR_AddDBValueToLast("ZipperLength", "%Length")

        'Draw Back Neck
        cBackNeckArc.X(1) = xyBodyCutOutBackNeck.X
        cBackNeckArc.y(1) = xyBodyCutOutBackNeck.Y
        cBackNeckArc.X(2) = xyBackNeckMid1.X
        cBackNeckArc.y(2) = xyBackNeckMid1.Y
        cBackNeckArc.X(3) = xyBackNeckMid2.X
        cBackNeckArc.y(3) = xyBackNeckMid2.Y
        cBackNeckArc.X(4) = xyBodyProfileNeck.X
        cBackNeckArc.y(4) = xyBodyProfileNeck.Y
        cBackNeckArc.n = 4
        PR_DrawFitted(cBackNeckArc)

        If InStr(g_sLegStyle, "Brief") > 0 Then
            'Draw back profile
            cProfilePoly.X(1) = xyBodyProfileNeck.X
            cProfilePoly.y(1) = xyBodyProfileNeck.Y
            cProfilePoly.X(2) = xyBodyProfileChest.X
            cProfilePoly.y(2) = xyBodyProfileChest.Y
            cProfilePoly.X(3) = xyBodyProfileWaist.X
            cProfilePoly.y(3) = xyBodyProfileWaist.Y
            cProfilePoly.X(4) = xyBodyProfileButt.X
            cProfilePoly.y(4) = xyBodyProfileButt.Y
            cProfilePoly.X(5) = xyGroinExtraPT1.X
            cProfilePoly.y(5) = xyGroinExtraPT1.Y
            cProfilePoly.X(6) = xyGroinExtraPT.X
            cProfilePoly.y(6) = xyGroinExtraPT.Y
            cProfilePoly.X(7) = xyBodyProfileGroin.X
            cProfilePoly.y(7) = xyBodyProfileGroin.Y

            If g_bDrawBriefCurve Then
                cProfilePoly.X(8) = xyBodyFold.X
                cProfilePoly.y(8) = xyBodyFold.Y
                cProfilePoly.X(9) = xyBDProfileThighExtraPT.X
                cProfilePoly.y(9) = xyBDProfileThighExtraPT.Y
                cProfilePoly.X(10) = xyBodyProfileThigh.X
                cProfilePoly.y(10) = xyBodyProfileThigh.Y
                cProfilePoly.n = 10
            Else
                cProfilePoly.X(8) = xyBDProfileThighExtraPT.X
                cProfilePoly.y(8) = xyBDProfileThighExtraPT.Y
                cProfilePoly.X(9) = xyBodyProfileThigh.X
                cProfilePoly.y(9) = xyBodyProfileThigh.Y
                cProfilePoly.n = 9
            End If
            PR_DrawFitted(cProfilePoly)
            PR_AddDBValueToLast("curvetype", "BackCurve")
            PR_AddDBValueToLast("Leg", "Left&Right")

            If g_bBodyDiffThigh Then
                'Revise data above
                PR_AddDBValueToLast("Leg", g_sBodySmallestThighGiven)
                'Draw the largest thigh profile

                aAngle = FN_CalcAngle(xyBodyLT_ButtocksArcCentre, xyBodyProfileButt)
                aInc = (aAngle - FN_CalcAngle(xyBodyLT_ButtocksArcCentre, xyBodyLT_Fold)) / 4
                cProfilePoly.n = 0
                For ii = 1 To 5
                    PR_CalcPolar(xyBodyLT_ButtocksArcCentre, aAngle, g_nButtRadius, xyTmp)
                    cProfilePoly.n = cProfilePoly.n + 1
                    cProfilePoly.X(cProfilePoly.n) = xyTmp.X
                    cProfilePoly.y(cProfilePoly.n) = xyTmp.Y
                    aAngle = aAngle - aInc
                Next ii
                cProfilePoly.n = cProfilePoly.n + 1
                cProfilePoly.X(cProfilePoly.n) = xyBodyLT_ProfileThigh.X
                cProfilePoly.y(cProfilePoly.n) = xyBodyLT_ProfileThigh.Y

                ARMDIA1.PR_SetLayer("Template" & g_sBodyLargestThighGiven)
                PR_DrawFitted(cProfilePoly)
                PR_AddDBValueToLast("curvetype", "BackCurveLargest")
                PR_AddDBValueToLast("Leg", g_sBodyLargestThighGiven)

                'Label Left and Right
                ARMDIA1.PR_SetLayer("Notes")
                'Insert ArrowHead symbol
                PR_DrawMarkerNamed("Closed Arrow", xyBodyLT_Fold, BODY_INCH1_4, BODY_INCH1_8, -90)
                PR_DrawMarkerNamed("Closed Arrow", xyBDProfileThighExtraPT, BODY_INCH1_4, BODY_INCH1_8, -90)
                PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
                PR_MakeXY(xyText, xyBDProfileThighExtraPT.X, xyBDProfileThighExtraPT.Y - BODY_INCH1_8)
                PR_DrawText(g_sBodySmallestThighGiven, xyText, 0.1, 0, 1)
                PR_MakeXY(xyText, xyBodyLT_Fold.X, xyBodyLT_Fold.Y - BODY_INCH1_8)
                PR_DrawText(g_sBodyLargestThighGiven, xyText, 0.1, 0, 1)

                'Reset back to original
                ARMDIA1.PR_SetLayer("Template" & g_sBodySmallestThighGiven)

            End If
        End If

        If g_sLegStyle = "Panty" Then
            PR_MakeXY(xyLegStart, -ARMDIA1.max(g_nRightLegLength, g_nLeftLegLength), 0)
            ARMDIA1.PR_SetLayer("Construct")
            If g_bDrawSingleLeg Then
                PR_DrawLegTemplate("Left", xyLegStart)
                SmallestLegCurve = g_LeftLegProfile
                nSmallestLegXOffset = g_nLeftLegLength
                If g_bLeftAboveKnee Then
                    sSmallestElasticText = "ELASTIC"
                Else
                    sSmallestElasticText = "NO ELASTIC"
                End If
            Else
                PR_DrawLegTemplate("Both", xyLegStart)
                If g_sBodySmallestThighGiven = "Left" Then
                    SmallestLegCurve = g_LeftLegProfile
                    LargestLegCurve = g_RightLegProfile
                    nLargestLegXOffset = g_nRightLegLength
                    nSmallestLegXOffset = g_nLeftLegLength
                    If g_bLeftAboveKnee Then
                        sSmallestElasticText = "ELASTIC"
                    Else
                        sSmallestElasticText = "NO ELASTIC"
                    End If
                    If g_bRightAboveKnee Then
                        sLargestElasticText = "ELASTIC"
                    Else
                        sLargestElasticText = "NO ELASTIC"
                    End If
                Else
                    SmallestLegCurve = g_RightLegProfile
                    LargestLegCurve = g_LeftLegProfile
                    nLargestLegXOffset = g_nLeftLegLength
                    nSmallestLegXOffset = g_nRightLegLength
                    If g_bLeftAboveKnee Then
                        sLargestElasticText = "ELASTIC"
                    Else
                        sLargestElasticText = "NO ELASTIC"
                    End If
                    If g_bRightAboveKnee Then
                        sSmallestElasticText = "ELASTIC"
                    Else
                        sSmallestElasticText = "NO ELASTIC"
                    End If
                End If

            End If
            cProfilePoly.X(1) = xyBodyProfileNeck.X
            cProfilePoly.y(1) = xyBodyProfileNeck.Y
            cProfilePoly.X(2) = xyBodyProfileChest.X
            cProfilePoly.y(2) = xyBodyProfileChest.Y
            cProfilePoly.X(3) = xyBodyProfileWaist.X
            cProfilePoly.y(3) = xyBodyProfileWaist.Y
            cProfilePoly.X(4) = xyBodyProfileButt.X
            cProfilePoly.y(4) = xyBodyProfileButt.Y
            cProfilePoly.X(5) = xyGroinExtraPT1.X
            cProfilePoly.y(5) = xyGroinExtraPT1.Y
            cProfilePoly.X(6) = xyGroinExtraPT.X
            cProfilePoly.y(6) = xyGroinExtraPT.Y
            ' cProfilePoly.x(7) = xyBodyProfileGroin.x
            ' cProfilePoly.y(7) = xyBodyProfileGroin.y
            cProfilePoly.n = 6

            'Join leg
            For ii = SmallestLegCurve.n To 1 Step -1
                cProfilePoly.n = cProfilePoly.n + 1
                cProfilePoly.X(cProfilePoly.n) = SmallestLegCurve.X(ii) - nSmallestLegXOffset
                cProfilePoly.y(cProfilePoly.n) = SmallestLegCurve.y(ii)
            Next ii
            If g_bDrawSingleLeg Then
                ARMDIA1.PR_SetLayer("TemplateLeft")
            Else
                ARMDIA1.PR_SetLayer("Template" & g_sBodySmallestThighGiven)
            End If
            PR_DrawFitted(cProfilePoly)

            PR_AddDBValueToLast("curvetype", "BackCurve")

            'Label legs
            ARMDIA1.PR_SetLayer("Notes")
            PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
            If Not g_bDrawSingleLeg Then
                PR_AddDBValueToLast("Leg", g_sBodySmallestThighGiven)
            Else
                PR_AddDBValueToLast("Leg", "Left&Right")
            End If

            'Draw closure
            If g_bDrawSingleLeg Then
                ARMDIA1.PR_SetLayer("TemplateLeft")
            Else
                ARMDIA1.PR_SetLayer("Template" & g_sBodySmallestThighGiven)
            End If
            PR_MakeXY(xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n))
            PR_MakeXY(xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n))
            PR_DrawLine(xyTmp, xyTmp1)

            'Add elastic text for smallest
            ARMDIA1.PR_SetLayer("Notes")
            BODYSUIT1.g_nCurrTextAngle = 270
            PR_MakeXY(xyText, xyTmp1.X + 0.3, 0)
            PR_DrawText(sSmallestElasticText, xyText, 0.125, 270, 1)

            'Draw other leg
            If Not g_bDrawSingleLeg Then
                'Draw the largest thigh profile
                aAngle = FN_CalcAngle(xyBDButtocksArcCentre, xyBodyProfileButt)
                aInc = (aAngle - FN_CalcAngle(xyBDButtocksArcCentre, xyGroinExtraPT)) / 4
                cProfilePoly.n = 0
                For ii = 1 To 5
                    PR_CalcPolar(xyBDButtocksArcCentre, aAngle, g_nButtRadius, xyTmp)
                    cProfilePoly.n = cProfilePoly.n + 1
                    cProfilePoly.X(cProfilePoly.n) = xyTmp.X
                    cProfilePoly.y(cProfilePoly.n) = xyTmp.Y
                    aAngle = aAngle - aInc
                Next ii

                'Join leg
                For ii = LargestLegCurve.n To 1 Step -1
                    cProfilePoly.n = cProfilePoly.n + 1
                    cProfilePoly.X(cProfilePoly.n) = LargestLegCurve.X(ii) - nLargestLegXOffset
                    cProfilePoly.y(cProfilePoly.n) = LargestLegCurve.y(ii)
                Next ii

                ARMDIA1.PR_SetLayer("Template" & g_sBodyLargestThighGiven)
                PR_DrawFitted(cProfilePoly)
                PR_AddDBValueToLast("curvetype", "BackCurveLargest")
                PR_AddDBValueToLast("Leg", g_sBodyLargestThighGiven)
                PR_MakeXY(xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n))
                PR_MakeXY(xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n))
                PR_DrawLine(xyTmp, xyTmp1)

                'Add elastic text
                ARMDIA1.PR_SetLayer("Notes")
                PR_MakeXY(xyText, xyTmp1.X + 0.3, 0)
                PR_DrawText(sLargestElasticText, xyText, 0.125, 0, 1)
            End If
        End If
        'Setup which side is being drawn
        If g_strSide = "Left" Then
            g_sSleeveType = g_sBodyLeftSleeve
        Else
            g_sSleeveType = g_sBodyRightSleeve
        End If
        PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)

        'Draw Axilla
        If g_sSleeveType = "Sleeveless" Then
            'Draw Sleeveless Axilla
            cRaglanPoly.X(1) = xyBodyCutOutFrontNeck.X
            cRaglanPoly.y(1) = xyBodyCutOutFrontNeck.Y
            cRaglanPoly.X(2) = xyBodyRaglan1.X
            cRaglanPoly.y(2) = xyBodyRaglan1.Y
            cRaglanPoly.X(3) = xyBodyRaglan2.X
            cRaglanPoly.y(3) = xyBodyRaglan2.Y
            cRaglanPoly.n = 3
            PR_DrawPoly(cRaglanPoly)
            cRaglanPoly.X(1) = xyBodyRaglan5.X
            cRaglanPoly.y(1) = xyBodyRaglan5.Y
            cRaglanPoly.X(2) = xyBodyRaglan6.X
            cRaglanPoly.y(2) = xyBodyRaglan6.Y
            cRaglanPoly.X(3) = xyBDProfileNeckMirror.X
            cRaglanPoly.y(3) = xyBDProfileNeckMirror.Y
            cRaglanPoly.n = 3
            PR_DrawPoly(cRaglanPoly)

            'Draw Raglan Curves
            PR_SetLayer("Template" & g_sAxillaSide)
            nFlip = True
            PR_CalcRaglan(xyBodyRaglan4, xyBodyRaglan2, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), -1)
            PR_DrawFitted(cRaglanCurve)
            nFlip = False
            PR_CalcRaglan(xyBodyRaglan4, xyBodyRaglan5, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1)
            PR_DrawFitted(cRaglanCurve)
            If g_bBodyDiffAxillaHeight Then
                PR_SetLayer("Template" & g_sAxillaSideLow)
                nFlip = False
                PR_CalcRaglan(xyBDRaglan4LowAxilla, xyBodyRaglan5, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1)
                PR_DrawFitted(cRaglanCurve)
                nFlip = True
                PR_CalcRaglan(xyBDRaglan4LowAxilla, xyBodyRaglan2, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), -1)
                PR_DrawFitted(cRaglanCurve)

                ARMDIA1.PR_SetLayer("Notes")
                PR_DrawMarkerNamed("Closed Arrow", xyBDRaglan4LowAxilla, BODY_INCH1_4, BODY_INCH1_8, 180)
                BODYSUIT1.g_nCurrTextAngle = 270
                PR_MakeXY(xyText, xyBDRaglan4LowAxilla.X - 0.3, xyBDRaglan4LowAxilla.Y)
                PR_DrawText(g_sAxillaSideLow, xyText, 0.125, 270, 1)
                PR_DrawMarkerNamed("Closed Arrow", xyBodyRaglan4, BODY_INCH1_4, BODY_INCH1_8, 180)
                PR_MakeXY(xyText, xyBodyRaglan4.X - 0.3, xyBodyRaglan4.Y)
                PR_DrawText(g_sAxillaSide, xyText, 0.125, 270, 1)
                BODYSUIT1.g_nCurrTextAngle = 0
            End If
        Else
            'Draw Sleeved Axilla
            PR_DrawLine(xyBodyCutOutFrontNeck, xyBodyRaglan1)
            PR_DrawLine(xyBodyRaglan3, xyBDProfileNeckMirror)

            'Draw Raglan Curves
            ARMDIA1.PR_SetLayer("Template" & g_sAxillaSide)
            nFlip = False
            PR_CalcRaglan(xyBodyRaglan2, xyBodyRaglan3, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1)
            PR_DrawFitted(cRaglanCurve)
            nFlip = True
            nLength = FN_CalcLength(xyBodyRaglan2, xyOpen(2))
            PR_CalcRaglan(xyBodyRaglan2, xyBodyRaglan1, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), nLength)
            PR_DrawFitted(cRaglanCurve)

            Select Case g_sAxillaType
                Case "Open"
                    ARMDIA1.PR_SetLayer("Notes")
                    PR_DrawMarkerNamed("Closed Arrow", xyOpen(1), BODY_INCH1_4, BODY_INCH1_8, aOpen(1))
                    PR_DrawMarkerNamed("Closed Arrow", xyOpen(2), BODY_INCH1_4, BODY_INCH1_8, aOpen(2))
                    sAxillaText = "Open"

                Case "Mesh"
                    PR_DrawMeshConstruction(xyBodyRaglan2, xyMesh, sAxillaText, nMeshSize)
                    nFlip = False
                    nMeshDistanceAlongRaglan = 0
                    If Not FN_CalcAxillaMesh(xyBodyRaglan2, xyBodyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                        MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
                    Else
                        PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSide)
                    End If
                    'Calculate the mesh seam line along back neck raglan
                    'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                    nFlip = True
                    If Not FN_CalcAxillaMesh(xyBodyRaglan2, xyBodyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                        MsgBox("Unable to calculate mesh along front neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
                    Else
                        PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSide, True)
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
            BODYSUIT1.g_nCurrTextAngle = 270
            If g_bBodyDiffAxillaHeight Then
                ARMDIA1.PR_SetLayer("Template" & g_sAxillaSideLow)
                nFlip = False
                PR_CalcRaglan(xyBDRaglan2LowAxilla, xyBodyRaglan3, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1)
                PR_DrawFitted(cRaglanCurve)
                nFlip = True
                nLength = FN_CalcLength(xyBDRaglan2LowAxilla, xyOpen(4))
                PR_CalcRaglan(xyBDRaglan2LowAxilla, xyBodyRaglan1, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), nLength)
                PR_DrawFitted(cRaglanCurve)

                Select Case g_sAxillaTypeLow
                    Case "Open"
                        ARMDIA1.PR_SetLayer("Notes")
                        PR_DrawMarkerNamed("Closed Arrow", xyOpen(3), BODY_INCH1_4, BODY_INCH1_8, aOpen(3))
                        PR_DrawMarkerNamed("Closed Arrow", xyOpen(4), BODY_INCH1_4, BODY_INCH1_8, aOpen(4))
                        sAxillaTextLow = "Open"
                    Case "Mesh"
                        PR_DrawMeshConstruction(xyBDRaglan2LowAxilla, xyMesh, sAxillaTextLow, nMeshSize)
                        nFlip = False
                        nMeshDistanceAlongRaglan = 0
                        If Not FN_CalcAxillaMesh(xyBDRaglan2LowAxilla, xyBodyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                            MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
                        Else
                            PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSideLow)
                        End If
                        'Calculate the mesh seam line along back neck raglan
                        'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                        nFlip = True
                        If Not FN_CalcAxillaMesh(xyBDRaglan2LowAxilla, xyBodyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                            MsgBox("Unable to calculate mesh along front neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
                        Else
                            PR_DrawMesh(MeshProfile, MeshSeam, g_sAxillaSideLow)
                        End If
                        sText = Trim(Str(nMeshDistanceAlongRaglan)) & "," & Trim(Str(nMeshSize))
                        If g_sAxillaSideLow = "Left" Then
                            g_sMeshLeft = sText
                        Else
                            g_sMeshRight = sText
                        End If

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
                PR_DrawMarkerNamed("Closed Arrow", xyBDRaglan2LowAxilla, BODY_INCH1_4, BODY_INCH1_8, 180)
                PR_MakeXY(xyText, xyBDRaglan2LowAxilla.X - 0.3, xyBDRaglan2LowAxilla.Y)
                PR_DrawText(g_sAxillaSideLow & " " & sAxillaTextLow, xyText, 0.125, 270, 1)

                PR_DrawMarkerNamed("Closed Arrow", xyBodyRaglan2, BODY_INCH1_4, BODY_INCH1_8, 180)
                PR_MakeXY(xyText, xyBodyRaglan2.X - 0.3, xyBodyRaglan2.Y)
                PR_DrawText(g_sAxillaSide & " " & sAxillaText, xyText, 0.125, 270, 1)

            Else
                'Take care of different axilla types at same axilla height
                If g_sBodyLeftSleeve <> g_sBodyRightSleeve Then
                    'Do the opposite axilla to that done above
                    If g_sBodyLeftSleeve = g_sAxillaType Then
                        g_sAxillaType = g_sBodyRightSleeve
                        sThis = "Right"
                        sOther = "Left"
                    Else
                        g_sAxillaType = g_sBodyLeftSleeve
                        sThis = "Left"
                        sOther = "Right"
                    End If
                    Select Case g_sAxillaType
                        Case "Open"
                            ARMDIA1.PR_SetLayer("Notes")
                            PR_DrawMarkerNamed("Closed Arrow", xyOpen(1), BODY_INCH1_4, BODY_INCH1_8, aOpen(1))
                            PR_DrawMarkerNamed("Closed Arrow", xyOpen(2), BODY_INCH1_4, BODY_INCH1_8, aOpen(2))
                            sAxillaTextThis = "Open"
                        Case "Mesh"
                            PR_DrawMeshConstruction(xyBodyRaglan2, xyMesh, sAxillaText, nMeshSize)
                            nFlip = False
                            nMeshDistanceAlongRaglan = 0
                            If Not FN_CalcAxillaMesh(xyBodyRaglan2, xyBodyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                                MsgBox("Unable to calculate mesh along back neck raglan curve!" & Chr(13) & "Draw manually and refer to supervisor.")
                            Else
                                PR_DrawMesh(MeshProfile, MeshSeam, sThis)
                            End If
                            'Calculate the mesh seam line along back neck raglan
                            'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                            nFlip = True
                            If Not FN_CalcAxillaMesh(xyBodyRaglan2, xyBodyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
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
                    PR_MakeXY(xyText, xyBodyRaglan2.X - 0.3, xyBodyRaglan2.Y)
                    ''PR_DrawText(sOther & " " & sAxillaText & "\n" & sThis & " " & sAxillaTextThis, xyText, 0.125, 270, 1)
                    PR_DrawMText(sOther & " " & sAxillaText & Chr(10) & sThis & " " & sAxillaTextThis, xyText, 270, True)

                Else
                    ARMDIA1.PR_SetLayer("Notes")
                    PR_MakeXY(xyText, xyBodyRaglan2.X - 0.3, xyBodyRaglan2.Y)
                    PR_DrawText(sAxillaText, xyText, 0.125, 270, 1)
                End If
            End If
            BODYSUIT1.g_nCurrTextAngle = 0

            'Update the database for use by sleeve drawing module
            PR_UpdateDB()
        End If
        'Draw Crotch (based on some fixed rules checked in BodySuit dialogue box)
        PR_MakeXY(xyCrotchText, xyBodyCrotchMarker.X - 0.375, xyBodyCrotchMarker.Y)
        PR_DrawMarker(xyBodyCrotchMarker)

        'Draw Crotch styles
        ARMDIA1.PR_SetLayer("Template" & "Left")

        Select Case g_sBodyCrotchStyle
            Case "Open Crotch"
                'Draw Cutout Arc
                PR_DrawLine(xyBodyBackCrotch1, xyBodyBackCrotch2)
                PR_DrawLine(xyBodyFrontCrotch1, xyBodyFrontCrotch2)
                PR_DrawArc(xyBDBackCrotchFilletCentre, xyBodyBackCrotch2, xyBodyBackCrotch3)
                PR_DrawArc(xyBDFrontCrotchFilletCentre, xyBodyFrontCrotch3, xyBodyFrontCrotch2)

                If g_bExtremeCrotch Then
                    'Extreme crotch
                    'Back crotch
                    If xyBodyBackCrotch1.X < xyBodyCutOut9.X Then
                        PR_DrawArc(xyBodyCutOutArcCentre, xyBodyCutOut9, xyBodyBackCrotch1)
                        PR_DrawLine(xyBodyCutOut9, xyBodyCutOut3)
                    ElseIf xyBodyBackCrotch1.X < xyBodyCutOut3.X Then
                        PR_DrawLine(xyBodyBackCrotch1, xyBodyCutOut3)
                    End If
                    'Reuse xyBodyBackCrotch3 here as the start of the arc
                    If xyBodyBackCrotch3.X > xyBodyCutOut9.X Then
                        xyTmp = xyBodyCutOut9 : xyTmp.Y = xyBodyBackCrotch3.Y
                        PR_DrawLine(xyBodyBackCrotch3, xyTmp)
                        xyBodyBackCrotch3.X = xyBodyCutOut9.X
                    Else
                        'Do nothing
                    End If
                    'Front crotch
                    If xyBodyFrontCrotch1.X < xyBodyCutOut10.X Then
                        PR_DrawArc(xyBodyCutOutArcCentre, xyBodyCutOut10, xyBodyFrontCrotch1)
                        PR_DrawLine(xyBodyCutOut10, xyBodyCutOut5)
                    ElseIf xyBodyFrontCrotch1.X < xyBodyCutOut5.X Then
                        PR_DrawLine(xyBodyFrontCrotch1, xyBodyCutOut5)
                    End If
                    'Reuse xyBodyFrontCrotch3 here as the start of the arc
                    If xyBodyFrontCrotch3.X > xyBodyCutOut10.X Then
                        xyTmp = xyBodyCutOut10 : xyTmp.Y = xyBodyFrontCrotch3.Y
                        PR_DrawLine(xyBodyFrontCrotch3, xyTmp)
                        xyBodyFrontCrotch3.X = xyBodyCutOut10.X
                    Else
                        'Do nothing
                    End If
                    'Note:- Reuse of xyFrontCrotch3and xyBodyBackCrotch3 here as the start and end of the arc
                    PR_DrawArc(xyBodyCutOutArcCentre, xyBodyBackCrotch3, xyBodyFrontCrotch3)

                Else
                    'Normal crotch
                    If xyBodyBackCrotch1.X <> xyBodyCutOut3.X And xyBodyBackCrotch1.Y <> xyBodyCutOut3.Y Then
                        PR_DrawArc(xyBodyCutOutArcCentre, xyBodyCutOut3, xyBodyBackCrotch1)
                    End If
                    If xyBodyFrontCrotch1.X <> xyBodyCutOut5.X And xyBodyFrontCrotch1.Y <> xyBodyCutOut5.Y Then
                        PR_DrawArc(xyBodyCutOutArcCentre, xyBodyFrontCrotch1, xyBodyCutOut5)
                    End If
                    PR_DrawArc(xyBodyCutOutArcCentre, xyBodyBackCrotch3, xyBodyFrontCrotch3)
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
                PR_DrawMarkerNamed("Closed Arrow", xyBodyCutOut3, BODY_INCH1_4, BODY_INCH1_8, 90)
                PR_DrawMarkerNamed("Closed Arrow", xyBodyCutOut5, BODY_INCH1_4, BODY_INCH1_8, 270)

            Case "Snap Crotch"

                sCrotchText = "GUSSET, MESH " & g_sGussetSize

                'DRAW Gusset rectangle
                ARMDIA1.PR_SetLayer("Template" & g_strSide)
                PR_MakeXY(xyLLGussetRectangle, xyBodyProfileButt.X - 2, xyBodyProfileButt.Y + 2)
                PR_MakeXY(xyURGussetRectangle, xyLLGussetRectangle.X + (xyBodyStrap2.X - xyBodyStrap4.X), xyLLGussetRectangle.Y + (xyBodyStrap1.Y - xyBodyStrap2.Y))
                PR_DrawRectangle(xyLLGussetRectangle, xyURGussetRectangle)

                PR_SetTextData(LEFT_, TOP_, CURRENT, CURRENT, CURRENT)
                BODYSUIT1.g_nCurrTextAngle = 0
                ARMDIA1.PR_SetLayer("Notes")
                PR_MakeXY(xyText, xyLLGussetRectangle.X + 0.25, xyURGussetRectangle.Y - 0.25)
                sText = g_sPatient & Chr(10) & g_sWorkOrder
                ''PR_DrawText(sText, xyText, 0.125, 0, 1)
                PR_DrawMText(sText, xyText, 0)

                'Draw DB markers for Zipper macros
                ARMDIA1.PR_SetLayer("Construct")
                PR_DrawMarker(xyBodyStrap1)
                PR_AddDBValueToLast("Zipper", "Snap")

            Case "Gusset"
                sCrotchText = "Gusset, " & g_sGussetSize

        End Select
        'Draw Crotch Text and marker
        ARMDIA1.PR_SetLayer("Notes")
        PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
        BODYSUIT1.g_nCurrTextAngle = 270
        If sCrotchText <> "" Then
            PR_DrawText(sCrotchText, xyCrotchText, 0.125, 270, 1)
        End If

        'Insert ArrowHead symbol
        PR_DrawMarkerNamed("Closed Arrow", xyBodyCrotchMarker, BODY_INCH1_4, BODY_INCH1_8, 180)

        'Draw Leg style
        ARMDIA1.PR_SetLayer("Template" & g_strSide)

        'Draw CutOut Arc
        If Not g_sBodyCrotchStyle = "Open Crotch" Then
            If g_bExtremeCrotch Then
                PR_DrawLine(xyBodyCutOut5, xyBodyCutOut10)
                PR_DrawLine(xyBodyCutOut9, xyBodyCutOut3)
                PR_AddEntityArc(xyBodyCutOutArcCentre, xyBodyCutOut9, xyBodyCutOut10)
            Else
                PR_AddEntityArc(xyBodyCutOutArcCentre, xyBodyCutOut3, xyBodyCutOut5)
            End If
        End If

        If g_sLegStyle = "Panty" Then
            BODYSUIT1.g_nCurrTextAngle = 0
            PR_SetTextData(HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT)
            ARMDIA1.PR_SetLayer("Notes")
            If (g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief") Then
                'There is only a single leg draw which is given by the left leg
                PR_MakeXY(xyText, xyBDLeftLegLabelPoint.X, xyBDLeftLegLabelPoint.Y - BODY_INCH1_8)
                PR_DrawMarkerNamed("Closed Arrow", xyBDLeftLegLabelPoint, BODY_INCH1_4, BODY_INCH1_8, -90)
                If g_sLeftLeg = "Brief" Then
                    PR_DrawText("Right", xyText, 0.125, 0, 1)
                    ARMDIA1.PR_SetLayer("TemplateLeft")
                Else
                    PR_DrawText("Left", xyText, 0.125, 0, 1)
                    ARMDIA1.PR_SetLayer("TemplateRight")
                End If
                PR_DrawBriefStandard(0.875)
            Else
                If g_bDrawSingleLeg Then
                    PR_MakeXY(xyText, xyBDLeftLegLabelPoint.X, xyBDLeftLegLabelPoint.Y - BODY_INCH1_8)
                    Dim xyPt As BODYSUIT1.XY
                    PR_MakeXY(xyPt, xyBDLeftLegLabelPoint.X, xyBDLeftLegLabelPoint.Y + BODY_INCH1_8)
                    ''PR_DrawMarkerNamed("Closed Arrow", xyBDLeftLegLabelPoint, BODY_INCH1_4, BODY_INCH1_8, -90)
                    PR_DrawMarkerNamed("Closed Arrow", xyPt, BODY_INCH1_4, BODY_INCH1_8, -90)
                    PR_DrawText("Left & Right", xyText, 0.125, 0, 8)
                Else
                    PR_MakeXY(xyText, xyBDLeftLegLabelPoint.X, xyBDLeftLegLabelPoint.Y - BODY_INCH1_8)
                    PR_DrawMarkerNamed("Closed Arrow", xyBDLeftLegLabelPoint, BODY_INCH1_4, BODY_INCH1_8, -90)
                    PR_DrawText("Left", xyText, 0.125, 0, 1)
                    PR_MakeXY(xyText, xyBDRightLegLabelPoint.X, xyBDRightLegLabelPoint.Y - BODY_INCH1_8)
                    PR_DrawMarkerNamed("Closed Arrow", xyBDRightLegLabelPoint, BODY_INCH1_4, BODY_INCH1_8, -90)
                    PR_DrawText("Right", xyText, 0.125, 0, 1)
                End If
            End If
        End If

        'Draw on body layer
        ARMDIA1.PR_SetLayer("Template" & g_strSide)

        If g_sLegStyle = "Brief" Then
            PR_DrawBriefStandard(0)
        End If

        If g_sLegStyle = "Brief-French" Then
            'calculate additional French-cut points
            If g_sSleeveType = "Sleeveless" Then
                xyFrenchCut.y = xyBodyRaglan4.Y
            Else
                xyFrenchCut.y = xyBodyRaglan2.Y
            End If
            If g_nAdult Then
                xyFrenchCut.x = xyBodySeamFold.X + 2.5
            Else
                xyFrenchCut.x = xyBodySeamFold.X + 2
            End If
            PR_DrawBriefFrenchCut()
        End If

        'Draw Lower axilla construction line
        If g_bBodyDiffAxillaHeight Then
            ARMDIA1.PR_SetLayer("Construct")
            PR_DrawVerticalLineThruPoint(xyBodySeamChestAxillaLow, nVertLinesOffset)
        End If

        'Set Neutral Layer
        ARMDIA1.PR_SetLayer("1")
        'PR_DebugPointText
        FileClose(BODYSUIT1.fNum)
        PR_Update_BlkAttribute()
        Me.Close()
        BodyMain.BodyMainDlg.Close()
    End Sub
    'To Get Insertion point from user
    Private Function PR_GetInsertionPoint() As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Indicate Start Point "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Return False
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(xyInsertion, ptStart.X, ptStart.Y)
        Return True
    End Function
    Private Sub PR_Update_BlkAttribute()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If acBlkTbl.Has("SUITBODY") Then
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkRef As BlockReference = dbObj
                        If blkRef.Name = "SUITBODY" Then
                            For Each attributeID As ObjectId In blkRef.AttributeCollection
                                Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                                If TypeOf attRefObj Is AttributeReference Then
                                    Dim acAttRef As AttributeReference = attRefObj
                                    If acAttRef.Tag.ToUpper().Equals("AxillaBackNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = g_nAxillaBackNeckRad
                                    ElseIf acAttRef.Tag.ToUpper().Equals("AxillaFrontNeckRad", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = g_nAxillaFrontNeckRad
                                    ElseIf acAttRef.Tag.ToUpper().Equals("ABNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = g_nABNRadRight
                                    ElseIf acAttRef.Tag.ToUpper().Equals("AFNRadRight", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = g_nAFNRadRight
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub
    Public Sub PR_EnableDrawBody()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If Not acBlkTbl.Has("BODYLEFTLEG") Then
                Exit Sub
            End If
            If Not acBlkTbl.Has("BODYRIGHTLEG") Then
                Exit Sub
            End If
        End Using
        Button1.Enabled = True
    End Sub
    Private Sub PR_DrawTextAsSymbol(ByRef xyPoint As BODYSUIT1.XY, ByRef sID As String, ByRef sData As String, ByRef nAngle As Double)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim sName As String = "TEXTASSYMBOL"
        Dim strTag(2), strTextString(2) As String
        strTag(1) = "ID"
        strTag(2) = "Data"
        strTextString(1) = sID
        strTextString(2) = sData
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sName) Then
                Dim blkTblRecSymbol As BlockTableRecord = New BlockTableRecord()
                blkTblRecSymbol.Name = sName
                Dim acAttDef As New AttributeDefinition
                Dim ii As Double
                For ii = 1 To 2
                    acAttDef = New AttributeDefinition
                    acAttDef.Position = New Point3d(xyPoint.X, xyPoint.Y, 0)
                    acAttDef.Prompt = strTag(ii)
                    acAttDef.Tag = strTag(ii)
                    acAttDef.TextString = strTextString(ii)
                    acAttDef.Height = 0.125
                    acAttDef.Justify = AttachmentPoint.BottomCenter
                    If ii = 2 Then
                        acAttDef.Invisible = False
                    Else
                        acAttDef.Invisible = True
                    End If
                    blkTblRecSymbol.AppendEntity(acAttDef)
                Next
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecSymbol)
                acTrans.AddNewlyCreatedDBObject(blkTblRecSymbol, True)
                blkRecId = blkTblRecSymbol.Id
            Else
                blkRecId = acBlkTbl(sName)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyInsertion.X, xyPoint.Y + xyInsertion.Y, 0), blkRecId)
                blkRef.Rotation = nAngle * (BODYSUIT1.PI / 180)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                idLastCreated = blkRef.ObjectId
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
                            If acAttRef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sID
                            ElseIf acAttRef.Tag.ToUpper().Equals("Data", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sData
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
    Function fnDisplayToCM(ByRef nInches As Double) As String
        Dim nCMVal, nDec As Double
        Dim iInt As Short
        nCMVal = nInches * 2.54
        nCMVal = System.Math.Abs(nCMVal)
        iInt = Int(nCMVal)
        nDec = nCMVal - iInt
        If nDec >= 0.5 Then
            fnDisplayToCM = Str(ARMDIA1.round(nCMVal))
            Exit Function
        End If
        If iInt <> 0 Then
            nDec = (nCMVal - iInt) * 10
            nDec = ARMDIA1.round(nDec) / 10
            nCMVal = iInt + nDec
            fnDisplayToCM = Str(nCMVal)
        Else
            fnDisplayToCM = Str(nCMVal)
        End If
    End Function

    Private Sub cboFabric_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFabric.SelectedIndexChanged
        BodyMain.g_sBodyFabric = cboFabric.Text
    End Sub
End Class
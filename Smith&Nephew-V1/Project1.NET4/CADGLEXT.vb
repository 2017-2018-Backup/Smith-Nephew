Option Strict Off
Option Explicit On
Module CADGLEXT
    'Project:   CADGLOVE.MAK
    'File:      CADGLEXT.BAS
    'Purpose:   Extend glove to Elbow and Axilla
    '
    'Version:   1.01
    'Date:      17.Jan.96
    'Author:    Gary George
    '
    'Projects:  CADGLOVE.MAK
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    '
    'Notes:-
    '
    '   This module is designed to be common to both the
    '   CAD Glove and the Manual Glove.  Therefor we make use
    '   of the global level variables indicated by g_
    '
    '   This may make the structure of the program a little
    '   more obscure but the concept is that all procedures and
    '   related module variables are in the one module GLVEXTEN.BAS
    '
    '
    Public MainForm As armdia
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote
    Public Const g_sDialogID As String = "MANUAL Glove Dialogue"

    'Flaps
    Public g_nStrapLength As Double
    Public g_nFrontStrapLength As Double
    Public g_nCustFlapLength As Double
    Public g_nWaistCir As Double
    Public g_sFlapType As String
    Public g_iFlapType As Short
    Public g_nWaistCirGiven As Double

    Public g_OnFold As Short
    Public g_nPalm As Double
    Public g_nWrist As Double
    Public g_iInsertSize As Short


    Structure xy
        Dim X As Double
        Dim Y As Double
    End Structure

    Public Structure curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim y() As Double

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim X(100)
            'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim y(100)
        End Sub
    End Structure
    Public Const PI As Double = 3.141592654
    'Public g_nPleats(1 To 4) As Double
    <VBFixedArray(4)> Dim g_nPleats() As Double
    Structure TapeData
        Dim nCir As Double
        Dim iMMs As Short
        Dim iRed As Short
        Dim iGms As Short
        Dim sNote As String
        Dim iTapePos As Short
        Dim sTapeText As String
    End Structure

    'Fingers
    Public Const DIP As Short = 1
    Public Const PIP As Short = 2
    Public Const THUMB As Short = 3
    Public Const THUMB_LEN As Short = 2
    Public Const FINGER_LEN As Short = 1
    Public Const g_sTapeText As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"


    'Tapes etc
    Public Const WRIST As Short = 1
    Public Const PALM As Short = 2
    Public Const TAPE_ONE_HALF As Short = 3
    Public Const TAPE_THREE As Short = 4
    Public Const TAPE_FOUR_HALF As Short = 5
    Public Const TAPE_SIX As Short = 6
    Public Const TAPE_SEVEN_HALF As Short = 7
    Public Const TAPE_NINE As Short = 8

    'Misc
    Public Const EIGHTH As Double = 0.125
    Public Const SIXTEENTH As Double = 0.0625
    Public Const QUARTER As Double = 0.25

    'Reduction Chart constants and Variables
    Const NOFF_FINGER_RED As Short = 32
    Const NOFF_ARM_RED As Short = 76
    Const NOFF_LENGTH_RED As Short = 24
    Const FIGURED_MINIMUM As Short = 6

    'Hand Reduction charts
    Dim iFingerRed(3, NOFF_FINGER_RED) As Short
    Dim nArmRed(8, NOFF_ARM_RED) As Double
    Dim nLengthRed(2, NOFF_LENGTH_RED) As Double

    'Arm reduction chart
    Public Const LOW_MODULUS As Short = 160
    Public Const HIGH_MODULUS As Short = 340
    Const NOFF_MODULUS As Short = 19

    Dim g_iModulusIndex As Short
    Dim g_iPowernet(NOFF_MODULUS, 23) As Short
    Public Const NOFF_ARMTAPES As Short = 16


    'Arm variables
    'Public Const NOFF_ARMTAPES As Short = 16
    Public Const ELBOW_TAPE As Short = 9
    Public Const ARM_PLAIN As Short = 0
    Public Const ARM_FLAP As Short = 1
    Public Const GLOVE_NORMAL As Short = 0
    Public Const GLOVE_ELBOW As Short = 1
    Public Const GLOVE_AXILLA As Short = 2

    Dim g_nCir(1, NOFF_ARMTAPES) As Double
    Dim g_iFirstTape As Integer = -1
    Dim g_iLastTape As Integer = -1
    Public g_iWristPointer As Integer = -1
    Dim g_iEOSPointer As Integer = -1
    Dim g_iNumTotalTapes As Integer = 0
    Dim g_iNumTapesWristToEOS As Integer = 0
    Dim g_iPressure As Integer = 1
    Dim g_ExtendTo As Integer
    Dim g_EOSType As Integer
    Dim g_nFrontStrapLength As Integer
    Dim g_nWaistCir As Integer

    Public g_iMMs(17) As Short
    Public g_nLengths(17) As Double
    Public g_iGms(17) As Short
    Public g_iRed(17) As Short
    Public g_iGmsInit(17) As Short
    Public g_nLengthsInit(17) As Double
    Public g_iChanged(17) As Short
    Public g_iVertexMap(17) As Short
    Public g_sPathJOBST As String

    Public g_sSide As String 'The side Left or right


    Sub PR_EnableFigureArm()

        'Converse of PR_DisableFigureArm

        Static ii As Short
        For ii = 3 To 11
            CType(MainForm.Controls("lblArm"), Object)(ii).Enabled = True
        Next ii

        For ii = 0 To 3
            CType(MainForm.Controls("lblPleat"), Object)(ii).Enabled = True
        Next ii

        CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = True
        CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = True
        If g_nPleats(1) > 0 Then CType(MainForm.Controls("txtWristPleat1"), Object).Text = CStr(g_nPleats(1))
        If g_nPleats(2) > 0 Then CType(MainForm.Controls("txtWristPleat2"), Object).Text = CStr(g_nPleats(2))
        CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = True
        CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = True
        If g_nPleats(3) > 0 Then CType(MainForm.Controls("txtShoulderPleat2"), Object).Text = CStr(g_nPleats(3))
        If g_nPleats(4) > 0 Then CType(MainForm.Controls("txtShoulderPleat1"), Object).Text = CStr(g_nPleats(4))

        CType(MainForm.Controls("frmCalculate"), Object).Enabled = True
        CType(MainForm.Controls("cboPressure"), Object).Enabled = True
        CType(MainForm.Controls("cboPressure"), Object).SelectedIndex = g_iPressure
        CType(MainForm.Controls("cmdCalculate"), Object).Enabled = True

    End Sub

    Sub PR_ExtendTo_Click(ByRef INDEX As Short)
         
        Static iStart, iFirst, ii, iLast, iElbow As Short

        iLast = 23
        iFirst = 8
        iElbow = 16 'Elbow w.r.t. txtExtCir()
        iStart = 8

        Select Case INDEX
            Case 0 'Normal Glove
                'Disable tapes above those required
                For ii = iStart To iLast
                    CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = False
                    CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = False
                    PR_GrdInchesDisplay(ii - 8, 0)
                Next ii

                'Disable all MMs fields etc
                For ii = iFirst To iLast
                    CType(MainForm.Controls("mms"), Object)(ii).Enabled = False
                    CType(MainForm.Controls("mms"), Object)(ii).Text = ""
                    PR_GramRedDisplay(ii - 8, 0, 0)
                Next ii

                PR_DisableFigureArm()
                PR_DisableGloveToAxilla()

                'set disable and disable
                CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = -1
                CType(MainForm.Controls("cboDistalTape"), Object).Enabled = False
                CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = -1
                CType(MainForm.Controls("cboProximalTape"), Object).Enabled = False

            Case 1 'Glove to Elbow
                'Enable MMs fields etc
                For ii = iFirst To iLast
                    If ii <= iElbow Then
                        CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = True
                        CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
                        CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = True
                    Else
                        CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = False
                        CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = False
                        CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
                        CType(MainForm.Controls("mms"), Object)(ii).Text = ""
                        PR_GramRedDisplay(ii - 8, 0, 0)
                        PR_GrdInchesDisplay(ii - 8, 0)
                    End If
                Next ii

                'Enable other bits
                PR_EnableFigureArm()
                PR_DisableGloveToAxilla()

                'Set Wrist and EOS pointers
                If g_iNumTapesWristToEOS > 0 Then
                    'Set wrist to given tape
                    If g_iWristPointer > g_iFirstTape Then
                        CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = g_iWristPointer
                    Else
                        'set to first tape
                        CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
                    End If

                    'Set EOS to given tape or to Elbow if it extends
                    'past the elbow
                    'NB. The order of tapes is reversed in this list
                    'starting at 19-1/2 and finishing at 0
                    If g_iEOSPointer < ELBOW_TAPE Then
                        If g_iEOSPointer >= g_iLastTape Then
                            CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
                        Else
                            CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - g_iEOSPointer
                        End If
                    ElseIf g_iEOSPointer > ELBOW_TAPE Or g_iLastTape > ELBOW_TAPE Then
                        CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - ELBOW_TAPE
                    Else
                        'set to last tape
                        CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
                    End If
                Else
                    CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
                    CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
                End If

                CType(MainForm.Controls("cboDistalTape"), Object).Enabled = True
                CType(MainForm.Controls("cboProximalTape"), Object).Enabled = True

            Case 2 'Glove to Axilla
                For ii = iFirst To iLast
                    CType(MainForm.Controls("mms"), Object)(ii).Enabled = True
                    CType(MainForm.Controls("txtExtCir"), Object)(ii).Enabled = True
                    CType(MainForm.Controls("lblTape"), Object)(ii).Enabled = True
                    '                MainForm!mms(ii) = ""
                    '                PR_GramRedDisplay ii - 8, 0, 0
                Next ii

                'Enable other bits
                PR_EnableFigureArm()

                CType(MainForm.Controls("frmGloveToAxilla"), Object).Enabled = True
                If g_EOSType = ARM_FLAP Then CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Else CType(MainForm.Controls("optProximalTape"), Object)(0).Checked = True


                CType(MainForm.Controls("optProximalTape"), Object)(0).Enabled = True
                CType(MainForm.Controls("optProximalTape"), Object)(1).Enabled = True

                If g_iNumTapesWristToEOS > 0 Then
                    'Set wrist to given tape
                    If g_iWristPointer > g_iFirstTape Then
                        CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = g_iWristPointer
                    Else
                        'set to first tape
                        CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
                    End If

                    'Set EOS to given tape
                    'NB. The order of tapes is reversed in this list
                    'starting at 19-1/2 and finishing at 0
                    If g_iEOSPointer < g_iLastTape Then
                        CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 17 - g_iEOSPointer
                    Else
                        'set to first tape
                        CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
                    End If
                Else
                    For ii = iFirst To iLast
                        CType(MainForm.Controls("mms"), Object)(ii).Text = ""
                        PR_GramRedDisplay(ii - 8, 0, 0)
                    Next ii
                    CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0
                    CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0
                End If

                CType(MainForm.Controls("cboDistalTape"), Object).Enabled = True
                CType(MainForm.Controls("cboProximalTape"), Object).Enabled = True

        End Select

    End Sub

    Sub PR_GetDlgAboveWrist()
        'General procedure to read the data given in the
        'dialogue controls and copy to module level variables.
        'Saves having to do this more than once.
        '
        'Updates:-
        '
        '    g_nCir(1, NOFF_ARMTAPES) As Double
        '    g_iFirstTape            As Integer
        '    g_iLastTape             As Integer
        '    g_iWristPointer         As Integer
        '    g_iEOSPointer           As Integer
        '    g_iNumTotalTapes        As Integer
        '    g_iNumTapesWristToEOS   As Integer
        '    g_EOSType               As Integer
        '    g_OnFold                As Integer
        '    g_ExtendTo              As Integer
        '    g_nPleats(1 To 4)       As Double
        '    g_iPressure             As Integer
        '

        'Dim g_nCir(1, NOFF_ARMTAPES) As Double
        'Dim g_iFirstTape As Integer
        'Dim g_iLastTape As Integer
        'Dim g_iWristPointer As Integer
        'Dim g_iEOSPointer As Integer
        'Dim g_iNumTotalTapes As Integer
        'Dim g_iNumTapesWristToEOS As Integer
        'Dim g_EOSType As Integer
        'Dim g_OnFold As Integer
        'Dim g_ExtendTo As Integer
        'Dim g_nPleats(1 To 4) As Double



        'NOTE:
        '    We ignore the fact that there may be missing tapes.
        '
        '    g_iWristPointer and g_iEOSPointer are used to indicate the
        '    start and finish tapes in the arrays and not in the
        '    txtExtCir() array

        Dim nCir As Double
        Dim ii As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("SSTab1"), Object).Tab <> 1 Then CType(MainForm.Controls("SSTab1"), Object).Tab = 1


        Dim g_iFirstTape As Integer = -1
        Dim g_iLastTape As Integer = -1
        Dim g_iWristPointer As Integer = -1
        Dim g_iEOSPointer As Integer = -1
        Dim g_iNumTotalTapes As Integer = 0
        Dim g_iNumTapesWristToEOS As Integer = 0
        Dim g_iPressure As Integer = 1
        Dim g_ExtendTo As Integer
        'Glove type
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then
            g_ExtendTo = GLOVE_NORMAL
            'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf CType(MainForm.Controls("optExtendTo"), Object)(1).Value = True Then
            g_ExtendTo = GLOVE_ELBOW
        Else
            g_ExtendTo = GLOVE_AXILLA
        End If



        Dim iIndex As Short
        If g_ExtendTo = GLOVE_NORMAL Then
            'Get values from Hand Data TAB
            'We will need to update this w.r.t CAD Glove later
            'Assummes no holes in data
            If g_nCir(0, 2) <= 0 Then
                g_iNumTotalTapes = 1
                g_iNumTapesWristToEOS = 1
            ElseIf g_nCir(0, 3) > 0 Then
                g_iNumTotalTapes = 3
                g_iNumTapesWristToEOS = 3
            Else
                g_iNumTotalTapes = 2
                g_iNumTapesWristToEOS = 2
            End If
            '        g_iWristPointer = 1
            g_iWristPointer = 0
            g_iEOSPointer = 1000 'Stupid value to fool stupid code
            'in PR_ExtendTo_Click
        Else
            'Get First and last tape (if any)
            For ii = 8 To 23
                nCir = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtExtCir"), Object)(ii))
                If nCir > 0 And g_iFirstTape = -1 Then g_iFirstTape = ii - 7
                g_nCir(0, ii - 7) = nCir
            Next ii

            For ii = 23 To 8 Step -1
                If Val(CType(MainForm.Controls("txtExtCir"), Object)(ii).ToString()) > 0 Then
                    g_iLastTape = ii - 7
                    Exit For
                End If
            Next ii

            'Total number of tapes
            If (g_iLastTape = g_iFirstTape) Then
                If g_iLastTape <> -1 Then g_iNumTotalTapes = 1
            Else
                g_iNumTotalTapes = (g_iLastTape - g_iFirstTape) + 1
            End If
            'Get values from "Above Wrist" TAB
            'Wrist tape
            'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
            iIndex = CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex
            If iIndex = 0 Or iIndex = -1 Then
                g_iWristPointer = g_iFirstTape
            Else
                g_iWristPointer = iIndex
            End If

            'EOS tape
            'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
            iIndex = CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex
            If iIndex = 0 Or iIndex = -1 Then
                g_iEOSPointer = g_iLastTape
            Else
                'NB. The order of tapes is reversed in this list
                'starting at 19-1/2 and finishing at 0
                g_iEOSPointer = 17 - iIndex
            End If


            'Number of tapes between
            If (g_iWristPointer = g_iEOSPointer) Then
                'Just in case there are no tapes
                'And "first" and "last" are used for Wrist and EOS
                If g_iLastTape <> -1 Then g_iNumTapesWristToEOS = 1
            Else
                g_iNumTapesWristToEOS = (g_iEOSPointer - g_iWristPointer) + 1
            End If

            'Pleats
            If CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = True Then g_nPleats(1) = Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) Else g_nPleats(1) = 0
            If CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = True Then g_nPleats(2) = Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) Else g_nPleats(2) = 0
            If CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = True Then g_nPleats(3) = Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) Else g_nPleats(3) = 0
            If CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = True Then g_nPleats(4) = Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) Else g_nPleats(4) = 0

            g_iPressure = CType(MainForm.Controls("cboPressure"), Object).SelectedIndex

        End If


        'End of support type
        g_EOSType = ARM_PLAIN
        If g_ExtendTo = GLOVE_AXILLA And CType(MainForm.Controls("optProximalTape"), Object)(1).Checked Then
            g_EOSType = ARM_FLAP
            g_sFlapType = CType(MainForm.Controls("cboFlaps"), Object).Text
            g_iFlapType = CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex
            If CType(MainForm.Controls("txtStrapLength"), Object).Text <> "" Then g_nStrapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtStrapLength"), Object))
            If CType(MainForm.Controls("txtFrontStrapLength"), Object).Text <> "" Then g_nFrontStrapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtFrontStrapLength"), Object))
            If CType(MainForm.Controls("txtCustFlapLength"), Object).Text <> "" Then g_nCustFlapLength = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtCustFlapLength"), Object))
            If CType(MainForm.Controls("txtWaistCir"), Object).Text <> "" Then g_nWaistCir = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtWaistCir"), Object))
        End If

        'On or off fold
        g_OnFold = False


    End Sub

    Sub PR_GetExtensionDDE_Data()
        'Procedure to use the data given in the Form for the
        'glove extension (if any) along the arm
        '
        Dim nAge, Flap As Short
        Dim jj, ii, nn As Short
        Dim nValue As Double
        Dim iGms, iMMs, iRed As Short
        Dim sDiag As String

        Dim g_nCir(1, NOFF_ARMTAPES) As Double
        Dim g_iFirstTape As Integer = -1
        Dim g_iLastTape As Integer = -1
        Dim g_iWristPointer As Integer = -1
        Dim g_iEOSPointer As Integer = -1
        Dim g_iNumTotalTapes As Integer = 0
        Dim g_iNumTapesWristToEOS As Integer = 0
        Dim g_iPressure As Integer = 1


        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("SSTab1"), Object).Tab = 1

        'Defaults
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True 'Normal Glove
        CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = 0 '1st Tape
        CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = 0 'Last Tape
        Flap = False


        'Set default pressure w.r.t diagnosis
        nAge = Val(CType(MainForm.Controls("txtAge"), Object).Text)

        ' if pressure empty set from diagnosis
        sDiag = UCase(Mid(CType(MainForm.Controls("txtDiagnosis"), Object).Text, 1, 6))
        'Arterial Insufficiency
        If sDiag = "ARTERI" Then
            g_iPressure = 0
            'Assist fluid dynamics
        ElseIf sDiag = "ASSIST" Then
            g_iPressure = 1
            'Blood Clots
        ElseIf sDiag = "BLOOD " Then
            g_iPressure = 1
            'Burns
        ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then
            g_iPressure = 0
            'Cancer
        ElseIf sDiag = "CANCER" Then
            g_iPressure = 1
            'Cardial Vascular Arrest
        ElseIf sDiag = "CARDIA" Then
            g_iPressure = 0
            'Carpal Tunnel Syndrome
        ElseIf sDiag = "CARPAL" Then
            g_iPressure = 1
            'Cellulitis
        ElseIf sDiag = "CELLUL" Then
            g_iPressure = 1
            'Chronic Venous Insufficiency
        ElseIf sDiag = "CHRONI" Then
            g_iPressure = 1
            'Heart Condition
        ElseIf sDiag = "HEART " Then
            g_iPressure = 0
            'Hemangioma
        ElseIf sDiag = "HEMANG" Then
            g_iPressure = 0
            'Lymphedema, Lymphedema 1+, Lymphedema 2+
            'N.B. Lymphedema <=> Lymphedema 1+
        ElseIf sDiag = "LYMPHE" Then
            If InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "2") > 0 Then
                If nAge <= 70 Then
                    g_iPressure = 2
                Else
                    g_iPressure = 1
                End If
            Else
                If nAge <= 70 Then
                    g_iPressure = 1
                Else
                    g_iPressure = 0
                End If
            End If
            'Night Wear
        ElseIf sDiag = "NIGHT " Then
            g_iPressure = 0
            'Post Fracture
        ElseIf sDiag = "POST F" Then
            g_iPressure = 1
            'Post Mastectomy
        ElseIf sDiag = "POST M" Then
            g_iPressure = 1
            'Postphlebetic Syndrome
        ElseIf sDiag = "POSTPH" Then
            g_iPressure = 1
            'Renal Disease (Kidney)
        ElseIf sDiag = "RENAL " Then
            g_iPressure = 1
            'Skin Graft
        ElseIf sDiag = "SKIN G" Then
            g_iPressure = 0
            'Stroke
        ElseIf sDiag = "STROKE" Then
            g_iPressure = 0
            'Tendonitis
        ElseIf sDiag = "TENDON" Then
            g_iPressure = 0
            'Thrombophlebitis
        ElseIf sDiag = "THROMB" Then
            g_iPressure = 1
            'Trauma
        ElseIf sDiag = "TRAUMA" Then
            g_iPressure = 1
            'Varicose Veins
        ElseIf sDiag = "VARICO" Then
            g_iPressure = 1
            'Request for 30 m/m, 40 m/m and 50 m/m
        ElseIf sDiag = "REQUES" Then
            If InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "30") > 0 Then
                g_iPressure = 0
            ElseIf InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "40") > 0 Then
                g_iPressure = 1
            ElseIf InStr(CType(MainForm.Controls("txtDiagnosis"), Object).Text, "50") > 0 Then
                g_iPressure = 2
            End If
            'Vein Removal
        ElseIf sDiag = "VEIN R" Then
            g_iPressure = 1
        End If

        'Glove extensions
        'Code for extended glove
        'Glove Option Buttons
        Dim nLen As Double

        If CType(MainForm.Controls("txtDataGlove"), Object).Text <> "" Then
            For ii = 0 To 5
                nValue = Val(Mid(CType(MainForm.Controls("txtDataGlove"), Object).Text, (ii * 2) + 1, 2))
                If ii = 0 Then
                    'Fold options
                    'Do nothing Off fold is default
                ElseIf ii = 1 And nValue >= 0 Then
                    'Pressure w.r.t Figuring and MMs
                    g_iPressure = nValue
                ElseIf ii = 2 Then
                    'Wrist Tape
                    CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex = nValue
                ElseIf ii = 3 Then
                    'Proximal Tape
                    CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex = nValue
                ElseIf ii = 4 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CType(MainForm.Controls("optExtendTo"), Object)(nValue).Value = True
                    If nValue = 2 Then Flap = True
                ElseIf ii = 5 Then
                    'Only if glove to axilla is set do we
                    'do anything
                    If Flap And nValue = 1 Then
                        'Flap, break up flap multiple field
                        CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True
                        PR_ProximalTape_Click((1))
                        For jj = 0 To 4
                            nValue = Val(Mid(CType(MainForm.Controls("txtFlap"), Object).Text, (jj * 3) + 1, 3))
                            If jj = 0 Then
                                CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = nValue
                            ElseIf jj = 1 And nValue > 0 Then
                                g_nStrapLength = nValue / 10
                                CType(MainForm.Controls("txtStrapLength"), Object).Text = CStr(g_nStrapLength)
                                PR_DisplayTextInches(CType(MainForm.Controls("txtStrapLength"), Object), CType(MainForm.Controls("labStrap"), Object))
                            ElseIf jj = 2 And nValue > 0 Then
                                g_nFrontStrapLength = nValue / 10
                                CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = CStr(g_nFrontStrapLength)
                                PR_DisplayTextInches(CType(MainForm.Controls("txtFrontStrapLength"), Object), CType(MainForm.Controls("labFrontStrapLength"), Object))
                            ElseIf jj = 3 And nValue > 0 Then
                                g_nCustFlapLength = nValue / 10
                                CType(MainForm.Controls("txtCustFlapLength"), Object).Text = CStr(g_nCustFlapLength)
                                PR_DisplayTextInches(CType(MainForm.Controls("txtCustFlapLength"), Object), CType(MainForm.Controls("labCustFlapLength"), Object))
                            ElseIf jj = 4 And nValue > 0 Then
                                g_nWaistCir = nValue / 10
                                CType(MainForm.Controls("txtWaistCir"), Object).Text = CStr(g_nWaistCir)
                                PR_DisplayTextInches(CType(MainForm.Controls("txtWaistCir"), Object), CType(MainForm.Controls("labWaistCir"), Object))
                            End If
                        Next jj
                    Else
                        'Disable flaps
                        PR_ProximalTape_Click((0))
                    End If
                End If
            Next ii
        End If

        'Set value for pressure
        CType(MainForm.Controls("cboPressure"), Object).SelectedIndex = g_iPressure

        'Glove to elbow and Glove to axilla
        'These can start at -3 (However to allow for possible
        'changes later we save up to -4-1/2 but ignore -4-1/2 for the
        'mean time)
        If CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text <> "" Then
            ii = 1
            For nn = 8 To 23
                nValue = Val(Mid(CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text, (ii * 3) + 1, 3))
                If nValue > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!txtExtCir(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CType(MainForm.Controls("txtExtCir"), Object)(nn) = nValue / 10
                    nLen = CADGLOVE1.FN_InchesValue(CType(MainForm.Controls("txtExtCir"), Object)(nn))
                    If nLen <> -1 Then PR_GrdInchesDisplay(nn - 8, nLen)
                    g_nCir(0, ii) = nLen
                Else
                    PR_GrdInchesDisplay(nn - 8, 0)
                End If
                iMMs = Val(Mid(CType(MainForm.Controls("txtTapeMMs"), Object).Text, (ii * 3) + 1, 3))
                If iMMs > 0 Then
                    iGms = Val(Mid(CType(MainForm.Controls("txtGrams"), Object).Text, (ii * 3) + 1, 3))
                    iRed = Val(Mid(CType(MainForm.Controls("txtReduction"), Object).Text, (ii * 3) + 1, 3))
                    CType(MainForm.Controls("mms"), Object)(nn).Text = CStr(iMMs)
                    PR_GramRedDisplay(nn - 8, iGms, iRed)
                    g_iMMs(ii) = iMMs
                    g_iRed(ii) = iRed
                    g_iGms(ii) = iGms
                End If
                ii = ii + 1
            Next nn
        End If

        'Pleats
        For ii = 0 To 1
            nValue = Val(Mid(CType(MainForm.Controls("txtWristPleat"), Object).Text, (ii * 3) + 1, 3))
            If ii = 0 And nValue > 0 Then
                CType(MainForm.Controls("txtWristPleat1"), Object).Text = CStr(nValue / 10)
            ElseIf ii = 1 And nValue > 0 Then
                CType(MainForm.Controls("txtWristPleat2"), Object).Text = CStr(nValue / 10)
            End If
            nValue = Val(Mid(CType(MainForm.Controls("txtShoulderPleat"), Object).Text, (ii * 3) + 1, 3))
            If ii = 0 And nValue > 0 Then
                CType(MainForm.Controls("txtShoulderPleat1"), Object).Text = CStr(nValue / 10)
            ElseIf ii = 1 And nValue > 0 Then
                CType(MainForm.Controls("txtShoulderPleat2"), Object).Text = CStr(nValue / 10)
            End If
        Next ii


    End Sub

    Sub PR_GramRedDisplay(ByRef iIndex As Short, ByRef GRAM As Short, ByRef Reduction As Short)
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdDisplay"), Object).Row = iIndex

        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdDisplay"), Object).Col = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If GRAM = 0 Then
            CType(MainForm.Controls("grdDisplay"), Object) = ""
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CType(MainForm.Controls("grdDisplay"), Object) = GRAM
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdDisplay"), Object).Col = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Reduction = 0 Then
            CType(MainForm.Controls("grdDisplay"), Object) = ""
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdDisplay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CType(MainForm.Controls("grdDisplay"), Object) = Reduction
        End If

    End Sub

    Sub PR_GrdInchesDisplay(ByRef iIndex As Short, ByRef nLen As Double)
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdInches"), Object).Row = iIndex
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdInches"), Object).Col = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!grdInches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("grdInches"), Object) = ARMEDDIA1.fnInchesToText(nLen)
    End Sub

    Sub PR_LoadPowernetChart()
        'Procedure to load the reduction charts for the arms from
        'disk
        '
        Static sModulus, sFile, sLine As String
        Static jj, fChart, ii, nn As Short

        sFile = g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(sFile) = "" Then
            MsgBox("Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If

        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, OpenMode.Input)

        'Get ARM chart reductions
        'NB these are fixed format
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_MODULUS
            Input(fChart, sModulus)
            Input(fChart, sLine)

            For jj = 0 To 22
                g_iPowernet(ii, jj + 1) = Val(Mid(sLine, (jj * 4) + 1, 4))
            Next jj

            ii = ii + 1

        End While

        'Close file
        FileClose(fChart)
        sLine = ""
        sModulus = ""

    End Sub

    Sub PR_LoadReductionCharts(ByRef sFabric As String)
        'Procedure to load the reduction charts from
        'disk
        'Two charts are loaded
        '    1. Length chart
        '    2. Circumferences based on the Fabric
        '

        Static sFile, sLine As String
        Static jj, fChart, iModulus, ii, nn As Short

        If sFabric = "" Then Exit Sub
        If UCase(Mid(sFabric, 1, 3)) <> "POW" Then
            MsgBox("Fabric chosen is not Powernet", 48, g_sDialogID)
            Exit Sub
        End If

        'Establish chart to be loaded
        '    Fabric Format  Pow MMM-XX Comment
        '    if mm < 230 use 230 chart
        '    if MM > 280 use 280 chart
        '
        iModulus = Val(Mid(sFabric, 5, 3))
        If iModulus < 230 Then iModulus = 230
        If iModulus > 280 Then iModulus = 280
        sFile = g_sPathJOBST & "\TEMPLTS\GLV_" & Trim(Str(iModulus)) & ".DAT"

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(sFile) = "" Then
            MsgBox("Fabric Chart, Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If

        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, OpenMode.Input)

        'Get finger and thumb reducations
        'NB these are fixed format
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_FINGER_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB use of "." to repeat previous number
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then

                iFingerRed(1, ii) = Val(Mid(sLine, 5, 2))

                If Mid(sLine, 8, 1) = "." Then
                    iFingerRed(2, ii) = iFingerRed(1, ii)
                Else
                    iFingerRed(2, ii) = Val(Mid(sLine, 8, 2))
                End If

                If Mid(sLine, 11, 1) = "." Then
                    iFingerRed(3, ii) = iFingerRed(2, ii)
                Else
                    iFingerRed(3, ii) = Val(Mid(sLine, 11, 2))
                End If

                ii = ii + 1
            End If
        End While

        'Wrist and Palm to 9 tape reductions
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_ARM_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB: use of "." to repeat previous number
            '    also translation from inches and eights to decimal inches)
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
                nArmRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
                jj = 1
                For nn = 9 To 34 Step 4
                    jj = jj + 1
                    If Mid(sLine, nn, 1) = "." Then
                        nArmRed(jj, ii) = nArmRed(jj - 1, ii)
                    Else
                        nArmRed(jj, ii) = Val(Mid(sLine, nn, 2)) + (Val(Mid(sLine, nn + 2, 1)) * 0.125)
                    End If
                Next nn
                ii = ii + 1
            End If
        End While

        'Close file
        FileClose(fChart)


        'Length Reduction chart
        sFile = g_sPathJOBST & "\TEMPLTS\GLV_LEN.DAT"

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(sFile) = "" Then
            MsgBox("Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If

        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, OpenMode.Input)
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_LENGTH_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB: use of "." to repeat previous number
            '    also translation from inches and eights to decimal inches)
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then

                nLengthRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)

                If Mid(sLine, 9, 1) = "." Then
                    nLengthRed(2, ii) = nLengthRed(1, ii)
                Else
                    nLengthRed(2, ii) = Val(Mid(sLine, 9, 2)) + (Val(Mid(sLine, 11, 1)) * 0.125)
                End If

                ii = ii + 1
            End If
        End While

        'Close file
        FileClose(fChart)

        sLine = ""

    End Sub

    Sub PR_ProximalTape_Click(ByRef INDEX As Short)
        Dim ii As Short
        If INDEX = 0 Then
            'Disable flaps
            g_iFlapType = CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex
            CType(MainForm.Controls("cboFlaps"), Object).Enabled = False
            CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = -1
            For ii = 0 To 4
                CType(MainForm.Controls("lblFlap"), Object)(ii).Enabled = False
            Next ii
            CType(MainForm.Controls("txtStrapLength"), Object).Enabled = False
            g_nStrapLength = Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)
            CType(MainForm.Controls("txtStrapLength"), Object).Text = ""
            CType(MainForm.Controls("labStrap"), Object).Text = ""

            CType(MainForm.Controls("txtFrontStrapLength"), Object).Enabled = False
            g_nFrontStrapLength = Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text)
            CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = ""
            CType(MainForm.Controls("labFrontStrapLength"), Object).Text = ""

            CType(MainForm.Controls("txtCustFlapLength"), Object).Enabled = False
            g_nCustFlapLength = Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text)
            CType(MainForm.Controls("txtCustFlapLength"), Object).Text = ""
            CType(MainForm.Controls("labCustFlapLength"), Object).Text = ""

            CType(MainForm.Controls("txtWaistCir"), Object).Enabled = False
            g_nWaistCir = Val(CType(MainForm.Controls("txtWaistCir"), Object).Text)
            CType(MainForm.Controls("txtWaistCir"), Object).Text = ""
            CType(MainForm.Controls("labWaistCir"), Object).Text = ""

        Else
            'Enable flaps
            CType(MainForm.Controls("cboFlaps"), Object).Enabled = True
            CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex = g_iFlapType

            For ii = 0 To 3
                CType(MainForm.Controls("lblFlap"), Object)(ii).Enabled = True
            Next ii
            CType(MainForm.Controls("txtStrapLength"), Object).Enabled = True
            If g_nStrapLength > 0 Then
                CType(MainForm.Controls("txtStrapLength"), Object).Text = CStr(g_nStrapLength)
                PR_DisplayTextInches(CType(MainForm.Controls("txtStrapLength"), Object), CType(MainForm.Controls("labStrap"), Object))
            End If

            CType(MainForm.Controls("txtFrontStrapLength"), Object).Enabled = True
            If g_nFrontStrapLength > 0 Then
                CType(MainForm.Controls("txtFrontStrapLength"), Object).Text = CStr(g_nFrontStrapLength)
                PR_DisplayTextInches(CType(MainForm.Controls("txtFrontStrapLength"), Object), CType(MainForm.Controls("labFrontStrapLength"), Object))
            End If

            CType(MainForm.Controls("txtCustFlapLength"), Object).Enabled = True
            If g_nCustFlapLength > 0 Then
                CType(MainForm.Controls("txtCustFlapLength"), Object).Text = CStr(g_nCustFlapLength)
                PR_DisplayTextInches(CType(MainForm.Controls("txtCustFlapLength"), Object), CType(MainForm.Controls("labCustFlapLength"), Object))
            End If

            If InStr(1, CType(MainForm.Controls("cboFlaps"), Object).Text, "D") > 0 Then
                CType(MainForm.Controls("lblFlap"), Object)(4).Enabled = True
                CType(MainForm.Controls("txtWaistCir"), Object).Enabled = True
                If g_nWaistCir > 0 Then
                    CType(MainForm.Controls("txtWaistCir"), Object).Text = CStr(g_nWaistCir)
                    PR_DisplayTextInches(CType(MainForm.Controls("txtWaistCir"), Object), CType(MainForm.Controls("labWaistCir"), Object))
                End If
            End If

        End If

    End Sub

    Sub PR_SetMMs(ByRef sPressure As String)
        'Set the mms based on the pressure and the wrist and EOS
        'tapes given
        'REF:    GOP 01-02/16, Section 1.4
        '
        Dim ii As Short

        For ii = g_iWristPointer To g_iEOSPointer
            Select Case sPressure
                Case "15"
                    If ii = g_iWristPointer Then
                        g_iMMs(ii) = 12
                    ElseIf ii < ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 15
                    ElseIf ii = ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 12
                    ElseIf ii = ELBOW_TAPE Then
                        g_iMMs(ii) = 8
                    ElseIf ii = ELBOW_TAPE + 1 Then
                        g_iMMs(ii) = 10
                    ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                        g_iMMs(ii) = 12
                    ElseIf ii = g_iEOSPointer - 1 Then
                        g_iMMs(ii) = 10
                    ElseIf ii = g_iEOSPointer Then
                        g_iMMs(ii) = 8
                    End If
                Case "20"
                    If ii = g_iWristPointer Then
                        g_iMMs(ii) = 16
                    ElseIf ii < ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 20
                    ElseIf ii = ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 16
                    ElseIf ii = ELBOW_TAPE Then
                        g_iMMs(ii) = 10
                    ElseIf ii = ELBOW_TAPE + 1 Then
                        g_iMMs(ii) = 13
                    ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                        g_iMMs(ii) = 16
                    ElseIf ii = g_iEOSPointer - 1 Then
                        g_iMMs(ii) = 13
                    ElseIf ii = g_iEOSPointer Then
                        g_iMMs(ii) = 10
                    End If
                Case "25"
                    If ii = g_iWristPointer Then
                        g_iMMs(ii) = 20
                    ElseIf ii < ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 25
                    ElseIf ii = ELBOW_TAPE - 1 Then
                        g_iMMs(ii) = 20
                    ElseIf ii = ELBOW_TAPE Then
                        g_iMMs(ii) = 13
                    ElseIf ii = ELBOW_TAPE + 1 Then
                        g_iMMs(ii) = 17
                    ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                        g_iMMs(ii) = 20
                    ElseIf ii = g_iEOSPointer - 1 Then
                        g_iMMs(ii) = 17
                    ElseIf ii = g_iEOSPointer Then
                        g_iMMs(ii) = 13
                    End If
            End Select
        Next ii

    End Sub

    Sub PR_SetModulusIndex(ByRef sFabric As String)
        'Set the module level variable iModulus Index based on
        'the chosen fabric
        '
        ' Where sFabric = Pow 230-2B
        '
        '    g_iModulusIndex Start = LOW_MODULUS  = 160
        '    g_iModulusIndex End   = HIGH_MODULUS = 340
        '
        ' hence
        '    g_iModulusIndex = ((230 - 160) + 10) / 10
        '
        '
        Dim iMod As Short

        iMod = Val(Mid(sFabric, 5, 3))
        If iMod < LOW_MODULUS Then
            g_iModulusIndex = LOW_MODULUS
        ElseIf iMod > HIGH_MODULUS Then
            g_iModulusIndex = HIGH_MODULUS
        Else
            g_iModulusIndex = ((iMod - LOW_MODULUS) + 10) / 10
        End If

    End Sub

    Function FN_GetGrams(ByVal iReduction As Short) As Short
        'Function that looks up the "Powernet" grams / tension chart
        'and returns the Grams for a given reduction value
        '
        'Used to back calculate from a given reduction
        '
        ' Input
        '        iReduction      Reduction given
        '
        ' Module Level variables
        '        g_iModulusIndex
        '        g_iPowernet()
        '
        If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
            FN_GetGrams = -1
            Exit Function
        End If

        If iReduction < 10 Then iReduction = 10
        If iReduction > 32 Then iReduction = 32

        FN_GetGrams = g_iPowernet(g_iModulusIndex, (iReduction - 10) + 1)


    End Function

    Function FN_GetReduction(ByRef iGrams As Object) As Short
        'Function that looks up the "Powernet" grams / tension chart
        'and returns the reduction value
        '
        ' Input
        '        iGrams      Grams calculated from the data
        '
        ' Module Level variables
        '        g_iModulusIndex
        '        g_iPowernet()
        '
        If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
            FN_GetReduction = -1
            Exit Function
        End If

        Static iValue, ii, iPrevValue As Short

        Select Case iGrams
            Case Is <= g_iPowernet(g_iModulusIndex, 1)
                FN_GetReduction = 10

            Case Is >= g_iPowernet(g_iModulusIndex, 23)
                FN_GetReduction = 32

            Case Else
                'Return value closest
                iPrevValue = 0
                For ii = 1 To 23
                    iValue = g_iPowernet(g_iModulusIndex, ii)
                    'UPGRADE_WARNING: Couldn't resolve default property of object iGrams. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If iValue > iGrams Then Exit For
                    iPrevValue = iValue
                Next ii

                'UPGRADE_WARNING: Couldn't resolve default property of object iGrams. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If System.Math.Abs(iGrams - iPrevValue) < System.Math.Abs(iGrams - iValue) Then
                    FN_GetReduction = ii + 8
                Else
                    FN_GetReduction = ii + 9
                End If

        End Select


    End Function

    Function FN_LengthWristToEOS() As Double
        'Calculates the distance from the EOS to the
        'wrist based on values given in the dialogue
        '
        Static ii As Short
        Static nLen, nSpace As Double

        'Get the data from the dialogue
        '    PR_GetDlgAboveWrist

        nLen = 0

        Select Case g_ExtendTo
            Case GLOVE_NORMAL
                For ii = 1 To g_iNumTapesWristToEOS
                    nSpace = 1.375
                    nLen = nLen + nSpace
                Next ii
                If nLen = 0 Then
                    'If no tapes after wrist then default to 0.625
                    nLen = 0.625
                End If
            Case GLOVE_ELBOW
                For ii = 1 To g_iNumTapesWristToEOS - 1
                    nSpace = 1.375
                    If ii = 1 And g_nPleats(1) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
                    If ii = 2 And g_nPleats(2) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
                    If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
                    If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
                    nLen = nLen + nSpace
                Next ii

            Case GLOVE_AXILLA
                For ii = 1 To g_iNumTapesWristToEOS - 1
                    nSpace = 1.375
                    If ii = 1 And g_nPleats(1) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
                    If ii = 2 And g_nPleats(2) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
                    If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
                    If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))
                    nLen = nLen + nSpace
                Next ii
                If g_EOSType = ARM_FLAP Then
                End If
        End Select

        FN_LengthWristToEOS = nLen

    End Function

    Function FN_ValidateExtensionData() As String
        'Procedure to validated the extension data
        Static sError, nL As String
        Static iNoPleats, ii As Short

        nL = Chr(13)
        sError = "" 'Initialise because of static

        'If normal glove then exit
        If g_ExtendTo = GLOVE_NORMAL Then
            FN_ValidateExtensionData = ""
            Exit Function
        End If

        'Check pleats
        iNoPleats = 0
        For ii = 1 To 4
            If g_nPleats(ii) > 0 Then iNoPleats = iNoPleats + 1
        Next ii
        If iNoPleats + 1 > g_iNumTapesWristToEOS Then
            sError = sError & "Number of pleats exceeds availble spaces between the arm tapes!" & nL & "Disable pleats by Double Clicking on pleat label." & nL
        End If

        'Check that calculate has been used
        '>>>>>>>
        'Not yet implemented
        '>>>>>>>

        'Check that wrist and first tape are the same
        If g_nCir(1, g_iWristPointer) <> g_nWrist Then
            sError = sError & "Wrist tape of the glove and the extension are different!" & nL
        End If

        'Check on style D flaps
        If g_EOSType = ARM_FLAP And InStr(1, g_sFlapType, "D") > 0 And g_nWaistCir = 0 Then
            sError = sError & "A Waist circumference must be given for a D-Style flap" & nL
        End If

        'Return error message
        FN_ValidateExtensionData = sError

    End Function

    Sub PR_CalculateArmTapeReductions()
        'Procedure to calculate the reductions
        'at each arm tape
        Dim ii As Short

        'Don't Calculate if normal glove
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then Exit Sub

        If CType(MainForm.Controls("cboFabric"), Object).Text = "" Then
            MsgBox("Fabric not given! ", 48, "Manual Glove - Calculate ARM Button")
            Exit Sub
        Else
            If UCase(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 1, 3)) <> "POW" Then
                MsgBox("Fabric chosen is not Powernet!", 48, "Manual Glove - Calculate ARM Button")
                Exit Sub
            End If
            If Val(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 5, 3)) < LOW_MODULUS Or Val(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 5, 3)) > HIGH_MODULUS Then
                MsgBox("Modulus of fabric chosen is not available on Gram / Tension reduction chart!", 48, "Manual Glove - Calculate ARM Button")
                Exit Sub
            End If

        End If


        'Update from dialogue
        PR_GetDlgAboveWrist()

        If g_iPressure < 0 Then
            MsgBox("Can't Calculate, missing Pressure", 48, "CAD Glove - Calculate ARM Button")
            Exit Sub
        End If

        If g_iNumTapesWristToEOS <= 1 Then
            MsgBox("Can't Calculate, Only one tape between given Wrist and EOS", 48, "CAD Glove - Calculate ARM Button")
            Exit Sub
        End If

        'Check that there are no holes in the data
        For ii = g_iWristPointer To g_iEOSPointer
            If g_nCir(1, ii) = 0 Then
                MsgBox("Can't Calculate, missing Tape Circumferences between given Wrist and EOS", 48, "CAD Glove - Calculate ARM Button")
                Exit Sub
            End If
        Next ii


        'Set the MMs based on the selected pressure
        PR_SetMMs((CType(MainForm.Controls("cboPressure"), Object).Text))

        'Set the modulus based on the fabric
        PR_SetModulusIndex((CType(MainForm.Controls("cboFabric"), Object).Text))

        'Calculate grams etc
        For ii = 1 To NOFF_ARMTAPES
            If ii >= g_iWristPointer And ii <= g_iEOSPointer Then

                g_iGms(ii) = ARMDIA1.round(g_nCir(1, ii) * g_iMMs(ii))
                g_iRed(ii) = FN_GetReduction(g_iGms(ii))

                'Adjust at wrist
                If ii = g_iEOSPointer Then
                    If ii < ELBOW_TAPE Then
                        'The EOS must lie between 10 and 14
                        'From the given reduction we back calculate the grams and mms
                        If g_iRed(ii) > 14 Then g_iRed(ii) = 14
                    Else
                        'For Flaps always use a 12
                        'From the given reduction we back calculate the grams and mms
                        If g_EOSType = ARM_PLAIN Then
                            g_iRed(ii) = 10
                        Else
                            g_iRed(ii) = 12
                        End If

                    End If
                    'Back calculate
                    g_iGms(ii) = FN_GetGrams(g_iRed(ii))
                    g_iMMs(ii) = ARMDIA1.round(g_iGms(ii) / g_nCir(1, ii))
                End If

                'Adjust at EOS  g_iWristPointer
                If ii = g_iWristPointer Then
                    'Always use a 10 reduction for wrist
                    g_iRed(ii) = 10
                    g_iGms(ii) = FN_GetGrams(g_iRed(ii))
                    g_iMMs(ii) = ARMDIA1.round(g_iGms(ii) / g_nCir(1, ii))
                End If

                'Display Results
                PR_GrdInchesDisplay(ii - 1, g_nCir(1, ii))
                PR_GramRedDisplay(ii - 1, g_iGms(ii), g_iRed(ii))
                CType(MainForm.Controls("mms"), Object)(ii + 7).Text = CStr(g_iMMs(ii))

            Else
                'Blank the displayed values
                PR_GrdInchesDisplay(ii - 1, 0)
                PR_GramRedDisplay(ii - 1, 0, 0)
                CType(MainForm.Controls("mms"), Object)(ii + 7).Text = ""
                g_iMMs(ii) = 0
                g_iRed(ii) = 0
                g_iGms(ii) = 0
            End If
        Next ii

    End Sub

    Sub PR_CalculateExtension()
        'Procedure to calculate the POINTS
        'used to draw the extension of the glove above the
        'wrist
        'The Procedure also supplies the data that is to be
        'printed at each tape
        'N.B.
        'The wrist points will be modified if the wrist has been
        'calculated

        Dim nValue, nMidWristX, nSpacing As Double
        Dim iStartTape, ii, iLastTape As Short
        Dim nFiguredValue As Double
        Dim iVertex As Short
        Dim nOriginalWrist, nInsert, nRevisedWrist As Double

        nInsert = g_iInsertSize * EIGHTH


        If g_iNumTapesWristToEOS = 0 Or g_iNumTapesWristToEOS = 1 Then
            'Extension ends at wrist
            UlnarProfile.n = 1
            UlnarProfile.x(1) = xyPalm(1).x
            UlnarProfile.y(1) = xyPalm(1).y - 0.75

            RadialProfile.n = 1
            RadialProfile.x(1) = xyPalm(6).x
            RadialProfile.y(1) = xyPalm(6).y - 0.75

            Exit Sub

        End If


        If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_AXILLA Then
            'As we have recalculated the wrist we need to start at the
            'wrist
            UlnarProfile.n = g_iNumTapesWristToEOS
            UlnarProfile.y(1) = xyPalm(1).y

            RadialProfile.n = g_iNumTapesWristToEOS
            RadialProfile.y(1) = xyPalm(6).y

            iStartTape = g_iWristPointer
            iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)
            iVertex = 1

            'Get datum point at wrist
            nFiguredValue = (g_nCir(1, iStartTape) * ((100 - g_iRed(iStartTape)) / 100))
            nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
            If g_sSide = "Right" Then
                nMidWristX = xyPalm(6).x + (nValue / 2)
            Else
                nMidWristX = xyPalm(6).x - (nValue / 2)
            End If


            'Note: As we can allow the wrist to be any tape we need a
            'vertex counter
            iVertex = 1

            For ii = iStartTape To iLastTape
                'NB use of 1/2 scale
                nFiguredValue = (g_nCir(1, ii) * ((100 - g_iRed(ii)) / 100))
                nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
                UlnarProfile.x(iVertex) = nMidWristX - (nValue / 2)
                RadialProfile.x(iVertex) = nMidWristX + (nValue / 2)

                'Setup the notes for the vertex
                'We do this here as we have all the data and it simplifies
                'the drawing side
                TapeNote(iVertex).sTapeText = LTrim(Mid(g_sTapeText, ((ii + 1) * 3) + 1, 3))
                TapeNote(iVertex).iTapePos = ii
                TapeNote(iVertex).nCir = g_nCir(1, ii)
                TapeNote(iVertex).iGms = g_iGms(ii)
                TapeNote(iVertex).iRed = g_iRed(ii)
                TapeNote(iVertex).iMMs = g_iMMs(ii)


                'Standard spacing
                nSpacing = 1.375

                'Account all for pleats
                'Wrist (as we have started at the wrist we use, iStartTape + 1
                'and iStartTape + 2 in this case)
                If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(1))
                If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(2))
                'XXXXXXXXXX
                'XXXX Careful now! this won't work if only 3 or less tapes given
                'XXXXXXXXXX
                'Axilla
                If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(3))
                If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = ARMEDDIA1.fnDisplayToInches(g_nPleats(4))

                'We are working from the top (wrist) down
                'iVertex = 1 is set before the For Loop (Values from wrist)
                If iVertex <> 1 Then
                    UlnarProfile.y(iVertex) = UlnarProfile.y(iVertex - 1) - nSpacing
                    RadialProfile.y(iVertex) = RadialProfile.y(iVertex - 1) - nSpacing
                    If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
                End If

                'Increment vertex count
                iVertex = iVertex + 1

            Next ii

            'Reset given wrist points xyPalm(1)
            xyPalm(1).x = UlnarProfile.x(1)
            xyPalm(1).y = UlnarProfile.y(1)
            xyPalm(6).x = RadialProfile.x(1)
            xyPalm(6).y = RadialProfile.y(1)

            PR_CalcMidPoint(xyPalm(1), xyPalm(6), xyPalm(7))

        End If

    End Sub


    Function FN_CalcAngle(ByRef xyStart As CADUTILS.XY, ByRef xyEnd As xy) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, y As Object
        Dim rAngle As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = xyEnd.X - xyStart.x
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        y = xyEnd.Y - xyStart.y

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
        rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If

    End Function

    Function FN_CalcLength(ByRef xyStart As CADUTILS.XY, ByRef xyEnd As xy) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.x) ^ 2 + (xyEnd.Y - xyStart.y) ^ 2)

    End Function

    Sub PR_CalcPolar(ByRef xyStart As CADUTILS.XY, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As xy)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double

        'Convert from degees to radians
        nAngle = nAngle * PI / 180

        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.x + A
        xyReturn.Y = xyStart.y + B

    End Sub


    Sub PR_CalcWristBlendLFS(ByRef xyArcCen As XY, ByRef aArcSweep As Double, ByRef xyArcStart As XY, ByRef xyBSt As XY, ByRef xyBEnd As XY, ByRef ReturnProfile As curve)
        '
        'N.B. Parameters are given in the order of decreasing Y
        '
        Static nR1, rAngle, nL3, nL1, nL2, aA1, aInc, nR2 As Double
        Static xyBotSt, xyTopSt, xyCenter, xyMidPoint, xyTopEnd, xyBotEnd As xy
        Static xyR2, xyArc As CADUTILS.XY
        Static nA, nThirdOfL2, aAngle As Double
        Static xyTEnd, xyPt1, xyPt2, xyArcSt As XY
        Static TopIsArc, MirrorResult, ii, Direction, BottomIsArc As Short
        Static nRadius As Double
        'UPGRADE_WARNING: Lower bound of array xyTmp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Static xyTmp(10) As XY
        Static nTol As Double

        'Do this as we can't use ByVal
        'UPGRADE_WARNING: Couldn't resolve default property of object xyBotSt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyBotSt = xyBSt
        'UPGRADE_WARNING: Couldn't resolve default property of object xyBotEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyBotEnd = xyBEnd
        'UPGRADE_WARNING: Couldn't resolve default property of object xyArc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyArc = xyArcCen
        'UPGRADE_WARNING: Couldn't resolve default property of object xyArcSt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyArcSt = xyArcStart

        'Get the point on the arc that will be
        'used as xyTEnd
        aAngle = FN_CalcAngle(xyArc, xyArcSt) + ((aArcSweep * 3) / 4)
        nRadius = FN_CalcLength(xyArc, xyArcSt)
        PR_CalcPolar(xyArc, aAngle, nRadius, xyTopEnd)
        'UPGRADE_WARNING: Couldn't resolve default property of object xyTEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyTEnd = xyTopEnd

        'Degenerate to a straight line
        'Then exit the sub routine        '????????????????
        If FN_CalcAngle(xyBSt, xyTEnd) = FN_CalcAngle(xyBEnd, xyBSt) Then
            ReturnProfile.n = 2
            ReturnProfile.X(1) = xyTEnd.X
            ReturnProfile.y(1) = xyTEnd.Y
            ReturnProfile.X(2) = xyBEnd.X
            ReturnProfile.y(2) = xyBEnd.Y
            Exit Sub
        End If

        nL2 = FN_CalcLength(xyBSt, xyTEnd)
        nL3 = FN_CalcLength(xyBEnd, xyBSt)

        'Get Included angles & radius & Centers of Arcs
        aAngle = FN_CalcAngle(xyArc, xyArcSt)
        aInc = aArcSweep / 4
        For ii = 2 To 4
            aAngle = aAngle + aInc
            PR_CalcPolar(xyArc, aAngle, nRadius, xyTmp(ii))
        Next ii

        nThirdOfL2 = nL2 / 3

        'Bottom arc
        nA = FN_CalcLength(xyTopEnd, xyBotEnd)
        nA = ((nL2 ^ 2 + nL3 ^ 2) - nA ^ 2) / (2 * nL3 * nL2)
        rAngle = BDYUTILS.Arccos(nA) / 2
        nR2 = nThirdOfL2 * System.Math.Tan(rAngle)
        PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(5))
        PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(8))

        'Establish
        aAngle = FN_CalcAngle(xyBotEnd, xyTopEnd)
        If aAngle < FN_CalcAngle(xyBotEnd, xyBotSt) Then
            aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) - 90
        Else
            aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) + 90
        End If
        PR_CalcPolar(xyTmp(5), aAngle, nR2, xyR2)

        'Check that the gap between the wrist point and the arc is less than
        'or equal to 0.0625 ie. 1/16"
        'If not then make it so.
        '    nTol = .0625
        '    nA = FN_CalcLength(xyBotSt, xyR2)
        '    If nA - nR2 > nTol Then
        '        nR2 = nR2 - ((nA - nR2) - nTol)
        '        'Bit of a misnomer here but it saves on a variable
        '        nThirdOfL2 = nR2 * Tan(rAngle)
        '        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(5)
        '        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(8)
        '        'aAngle from above
        '        PR_CalcPolar xyTmp(5), aAngle, nR2, xyR2
        '    End If

        'Bottom arc points
        aAngle = FN_CalcAngle(xyR2, xyTmp(5))
        aA1 = FN_CalcAngle(xyR2, xyTmp(8))
        If aA1 - aAngle < 180 Then
            aInc = (aA1 - aAngle) / 3
        Else
            aInc = ((aA1 - aAngle) - 360) / 3
        End If

        For ii = 6 To 7
            aAngle = aAngle + aInc
            PR_CalcPolar(xyR2, aAngle, nR2, xyTmp(ii))
        Next ii



        ReturnProfile.n = 9
        ReturnProfile.X(1) = xyArcStart.X
        ReturnProfile.y(1) = xyArcStart.Y
        ReturnProfile.X(9) = xyBotEnd.X
        ReturnProfile.y(9) = xyBotEnd.Y

        For ii = 2 To 8
            ReturnProfile.X(ii) = xyTmp(ii).X
            ReturnProfile.y(ii) = xyTmp(ii).Y
        Next ii

    End Sub

    Sub PR_CalcMidPoint(ByRef xyStart As xy, ByRef xyEnd As xy, ByRef xyMid As xy)

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

    Sub PR_CalcWristBlendThumbSide(ByRef xyTSt As xy, ByRef xyTEnd As xy, ByRef xyBSt As xy, ByRef xyBEnd As xy, ByRef ReturnProfile As curve)
        '
        'N.B. Parameters are given in the order of decreasing Y
        '
        Static xyTmp As xy
        Static nL, nL1, aAngle As Double
        Static ii As Short

        nL1 = FN_CalcLength(xyBSt, xyTEnd) / 3
        aAngle = FN_CalcAngle(xyBSt, xyTEnd)

        PR_CalcMidPoint(xyTSt, xyTEnd, xyTmp)

        ReturnProfile.n = 6

        ReturnProfile.X(1) = xyTSt.X
        ReturnProfile.y(1) = xyTSt.Y

        ReturnProfile.X(2) = xyTmp.X
        ReturnProfile.y(2) = xyTmp.Y

        ReturnProfile.X(3) = xyTEnd.X
        ReturnProfile.y(3) = xyTEnd.Y

        nL = 0
        nL1 = 0.125
        For ii = 5 To 4 Step -1
            nL = nL + nL1
            PR_CalcPolar(xyBSt, aAngle, nL, xyTmp)
            ReturnProfile.X(ii) = xyTmp.X
            ReturnProfile.y(ii) = xyTmp.Y
        Next ii

        ReturnProfile.X(6) = xyBSt.X
        ReturnProfile.y(6) = xyBSt.Y


    End Sub

    Sub PR_DisableFigureArm()

        Dim ii As Short

        For ii = 3 To 11
            CType(MainForm.Controls("lblArm"), Object)(ii).Enabled = False
        Next ii

        'Pleats
        g_nPleats(1) = Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text)
        g_nPleats(2) = Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text)
        g_nPleats(3) = Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text)
        g_nPleats(4) = Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text)

        For ii = 0 To 3
            CType(MainForm.Controls("lblPleat"), Object)(ii).Enabled = False
        Next ii


        CType(MainForm.Controls("txtShoulderPleat1"), Object).Enabled = False
        CType(MainForm.Controls("txtShoulderPleat2"), Object).Enabled = False
        '  MainForm!txtShoulderPleat1 = ""
        '  MainForm!txtShoulderPleat2 = ""
        CType(MainForm.Controls("txtWristPleat1"), Object).Enabled = False
        CType(MainForm.Controls("txtWristPleat2"), Object).Enabled = False
        '  MainForm!txtWristPleat1 = ""
        '  MainForm!txtWristPleat2 = ""

        CType(MainForm.Controls("frmCalculate"), Object).Enabled = False
        CType(MainForm.Controls("cboPressure"), Object).Enabled = False
        CType(MainForm.Controls("cmdCalculate"), Object).Enabled = False

    End Sub

    Sub PR_DisableGloveToAxilla()

        CType(MainForm.Controls("frmGloveToAxilla"), Object).Enabled = False
        CType(MainForm.Controls("optProximalTape"), Object)(0).Checked = False
        CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = False
        CType(MainForm.Controls("optProximalTape"), Object)(0).Enabled = False
        CType(MainForm.Controls("optProximalTape"), Object)(1).Enabled = False
        PR_ProximalTape_Click((0))

    End Sub

    Sub PR_DisplayTextInches(ByRef ctlText As System.Windows.Forms.Control, ByRef ctlCaption As System.Windows.Forms.Control)
        Static nLen As Double
        nLen = CADGLOVE1.FN_InchesValue(ctlText)
        If nLen <> -1 Then ctlCaption.Text = ARMEDDIA1.fnInchesToText(nLen)
    End Sub

    Sub PR_MakeXY(ByRef xyReturn As xy, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.Y = y
    End Sub

    Sub PR_RotateCurve(ByRef xyRotate As CADUTILS.XY, ByRef aRotation As Double, ByRef Profile As CADUTILS.Curve)
        Static ii As Short
        Static xyVertex As xy
        For ii = 1 To Profile.n
            PR_MakeXY(xyVertex, Profile.x(ii), Profile.y(ii))
            PR_RotatePoint(xyRotate, aRotation, xyVertex)
            Profile.x(ii) = xyVertex.X
            Profile.y(ii) = xyVertex.Y
        Next ii
    End Sub

    Sub PR_RotatePoint(ByRef xyRotate As CADUTILS.XY, ByRef aRotation As Double, ByRef xyPoint As xy)
        Static aAngle, nLength As Double

        'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        aAngle = (FN_CalcAngle(xyRotate, xyPoint) + aRotation) Mod 360
        nLength = FN_CalcLength(xyRotate, xyPoint)
        PR_CalcPolar(xyRotate, aAngle, nLength, xyPoint)

    End Sub

    Sub PR_MirrorCurveInYaxis(ByRef nXValue As Double, ByRef nTranslate As Double, ByRef Profile As curve)
        'Procedure to mirror a curve in the y axis about
        'a line given by the X value
        '
        Static ii As Short
        For ii = 1 To Profile.n
            Profile.X(ii) = (nXValue - (Profile.X(ii) - nXValue)) + nTranslate
        Next ii

    End Sub

    Sub PR_DrawShoulderFlaps(ByRef xyS As xy, ByRef xyE As xy)
        'Draws shoulder flaps
        'Extracted from ARMDIA.FRM (Which I sub-contracted)
        'A load of £%&**!&*^$><@~, but it works (just about!)
        'The only reason to keep it, is that it is production proven
        'with the arm.
        '
        'We calculate the curves and point as they would be
        'attached to an arm curve.
        'We then rotate and mirror the points to fit the
        'drawing of a Glove to axilla
        '
        'UPGRADE_WARNING: Arrays in structure RaglanFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'UPGRADE_WARNING: Arrays in structure ShoulderFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Static ShoulderFlap, RaglanFlap As CADUTILS.Curve
        Static LastTapeValue As Object
        Static xyPt1, xyPt2 As xy
        Static xyTopFlap4, xyTopFlap2, xyTopFlap1, xyTopFlap3, xyTopFlap5 As xy
        Static xyTopFlap9, xyTopFlap7, xyTopFlap6, xyTopFlap8, xyTopFlap10 As xy
        Static xyBotFlap4, xyBotFlap2, xyBotFlap1, xyBotFlap3, xyBotFlap5 As xy
        Static xyFlap4, xyFlap2, xyFlap1, xyFlap3, xyFlap5 As xy
        Static Alpha, Theta, Phi, BETA, Omega, Delta, BottomRadius, TopArcIncrement As Object
        Static TopRadius, LittleBit, opp, h1, x1, y1, adj, HYP, Change, FlapMarkerAngle As Object
        Static xyFlapText1 As xy
        Static xyFlapText2, xyCentre, xyMid, xyFlapMarker, xyStrapText As xy
        Static FlapLength, FlapMarkerLength As Object
        Static TopArcAngle, TopArcRadius, TopArcLength, ArcLength, CircleCircum, BottomArcLength, BottomArcRadius, BottomArcAngle As Object
        Static xfablen, TempMarker, xfabby, xfabdist As Object
        Static xyRaglan6, xyRaglan4, xyRaglan2, xyRaglan1, xyRaglan3, xyRaglan5, xyRaglan7 As xy
        Static TopRagLanAngle, RaglanBottom, RagAng2, Rag4, Rag2, Rag1, Rag3, RagAng1, RagAng3, RaglanTip, InnerRaglanTip As Object
        Static nTemplateAngle, nNotchOffset, nTemplateRadius, nNotchToTangent As Double
        Static ii As Short
        Static RagAng5, PhiExtra, RagAng4, aMarker As Double
        Static Strap As String
        Static aAngle, nLength As Double
        Static xyStrt, xyMidPoint, xyTmp As CADUTILS.XY


        PR_CalcMidPoint(xyS, xyE, xyMidPoint)
        nLength = FN_CalcLength(xyS, xyE)
        aAngle = -90
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BottomArcAngle = 47.5
        'All this crap because you can't use ByVal on user defined types (Bastards!)
        'VB3 comment
        If g_sSide = "Right" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object xyStrt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyStrt = xyE
            aMarker = 135
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object xyStrt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyStrt = xyS
            aMarker = 45
        End If


        ' Check for Custom Flap Length
        If Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LastTapeValue = ARMEDDIA1.fnDisplayToInches(CDbl(CType(MainForm.Controls("txtCustFlapLength"), Object).Text))
        Else
            'Standard flap length
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LastTapeValue = g_nCir(1, g_iEOSPointer)
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LastTapeValue = (LastTapeValue / 3.14) * 0.92
        End If


        ' Calculate Bottom part of curve with 5 Points
        If nLength >= 5.5 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BottomRadius = LastTapeValue - 0.625
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanBottom = 3.5
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopRagLanAngle = 17
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanTip = 1.9375
            'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            InnerRaglanTip = 2.173
            nNotchOffset = 0.875 '14/16 ths Raglan Template s287
        ElseIf (nLength < 5.5) And (nLength > 3) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BottomRadius = 2.8
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanBottom = 2.5
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopRagLanAngle = 17
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanTip = 1.05
            'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            InnerRaglanTip = 1.22
            nNotchOffset = 0.75 '12/16 ths Raglan Template s289
        ElseIf (nLength <= 3) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BottomRadius = 2
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanBottom = 2.0625
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RaglanTip = 0.625
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopRagLanAngle = 15
            'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            InnerRaglanTip = 0.727
            nNotchOffset = 0.6875 '11/16 ths, Raglan Template s290
        End If

        nTemplateRadius = 3.5
        nTemplateAngle = 40
        nNotchToTangent = 2.125

        'Calc length of bottom Arc
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BottomArcRadius = BottomRadius
        'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CircleCircum = PI * (2 * (BottomArcRadius))
        'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BottomArcLength = (BottomArcAngle / 360) * CircleCircum

        'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ShoulderFlap.X(1) = xyStrt.X + LastTapeValue
        ShoulderFlap.y(1) = 0

        For ii = 1 To 4
            'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Theta = (ii * 11.875 * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Cos(Theta) * BottomRadius
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Sin(Theta) * BottomRadius
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LittleBit = BottomRadius - adj
            'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ShoulderFlap.X(ii + 1) = xyStrt.X + LastTapeValue - LittleBit
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ShoulderFlap.y(ii + 1) = 0 + opp
        Next ii
        PR_MakeXY(xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.y(5))

        'Calculate 6 angles and points for top part of curve
        PR_MakeXY(xyTopFlap1, xyStrt.X, xyStrt.Y + nLength)
        ShoulderFlap.X(10) = xyTopFlap1.X
        ShoulderFlap.y(10) = xyTopFlap1.Y

        'midpoint
        'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        h1 = FN_CalcLength(xyTopFlap1, xyBotFlap5)
        'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        h1 = h1 / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Omega = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
        'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Omega = 180 - Omega
        'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Omega = (Omega * PI) / 180
        'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        opp = System.Math.Abs(h1 * System.Math.Sin(Omega))
        'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        adj = System.Math.Abs(h1 * System.Math.Cos(Omega))
        'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_MakeXY(xyMid, xyBotFlap5.X - adj, xyBotFlap5.Y + opp)

        'Centre of Top Arc
        'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Delta = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
        'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Delta = Delta - 90
        'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Delta = (Delta * PI) / 180
        'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HYP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HYP = 4 * h1
        'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HYP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        opp = System.Math.Abs(HYP * System.Math.Sin(Delta))
        'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HYP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        adj = System.Math.Abs(HYP * System.Math.Cos(Delta))
        'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_MakeXY(xyCentre, xyMid.X + adj, xyMid.Y + opp)
        'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopRadius = FN_CalcLength(xyCentre, xyTopFlap1)
        'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopArcRadius = TopRadius
        'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Phi = FN_CalcAngle(xyCentre, xyTopFlap1)
        'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Phi = Phi - 180
        'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Alpha = FN_CalcAngle(xyCentre, xyBotFlap5)
        'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Alpha = Alpha - 180
        'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopArcAngle = Alpha - Phi
        'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopArcIncrement = (Alpha - Phi) / 5

        'Calc Flap Length for position of marker
        'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CircleCircum = PI * (2 * (TopArcRadius))
        'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopArcLength = (TopArcAngle / 360) * CircleCircum
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FlapLength = TopArcLength + BottomArcLength
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FlapMarkerLength = FlapLength / 2.7
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
        'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If FlapMarkerLength < TopArcLength Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerAngle = FlapMarkerAngle + Phi
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(TopRadius * System.Math.Sin(FlapMarkerAngle))
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(TopRadius * System.Math.Cos(FlapMarkerAngle))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyFlapMarker, xyCentre.X - adj, xyCentre.Y - opp)
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf FlapMarkerLength > TopArcLength Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TempMarker = FlapMarkerLength - TopArcLength
            'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TempMarker = BottomArcLength - TempMarker
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(BottomArcRadius * System.Math.Sin(FlapMarkerAngle))
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(BottomArcRadius * System.Math.Cos(FlapMarkerAngle))
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LittleBit = BottomRadius - adj
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyFlapMarker, xyStrt.X + LastTapeValue - LittleBit, 0 + opp)
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf FlapMarkerLength = TopArcLength Then
            PR_MakeXY(xyFlapMarker, xyBotFlap5.X, xyBotFlap5.Y)
        End If
        ARMDIA1.PR_SetLayer("Template" & g_sSide)

        For ii = 1 To 4
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PhiExtra = Phi + (ii * TopArcIncrement)
            PhiExtra = (PhiExtra * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(TopRadius * System.Math.Sin(PhiExtra))
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(TopRadius * System.Math.Cos(PhiExtra))
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ShoulderFlap.X(10 - ii) = xyCentre.X - adj
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ShoulderFlap.y(10 - ii) = xyCentre.Y - opp
        Next ii


        If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
            'Copy standard shoulder from profile down
            'only go one vertex past the FlapMarker
            RaglanFlap.n = 0
            For ii = 10 To 1 Step -1
                RaglanFlap.n = RaglanFlap.n + 1
                RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
                RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
                If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
            Next ii
            PR_RotateCurve(xyStrt, aAngle, RaglanFlap)
            If g_sSide = "Left" Then PR_MirrorCurveInYaxis(xyMidPoint.X, 0, RaglanFlap)
            ARMDIA1.PR_DrawFitted(RaglanFlap)
        Else
            ShoulderFlap.n = 10
            CADUTILS.PR_RotateCurve(xyStrt, aAngle, ShoulderFlap)
            If g_sSide = "Left" Then CADUTILS.PR_MirrorCurveInYaxis(xyMidPoint.X, 0, ShoulderFlap)
            ARMDIA1.PR_DrawFitted(ShoulderFlap)
        End If

        'Draw Raglan
        If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag1 = System.Math.Sqrt((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag1 = BottomRadius - Rag1
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan1, xyStrt.X + LastTapeValue - Rag1, nNotchOffset)
            PR_MakeXY(xyRaglan2, xyRaglan1.X - nNotchToTangent, 0)
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng1 = FN_CalcAngle(xyRaglan2, xyRaglan1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag2 = FN_CalcLength(xyRaglan2, xyRaglan1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag2 = Rag2 / 2
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng1 = (RagAng1 * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(Rag2 * System.Math.Sin(RagAng1))
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(Rag2 * System.Math.Cos(RagAng1))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan3, xyRaglan2.X + adj, xyRaglan2.Y + opp)

            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng1 = FN_CalcAngle(xyRaglan3, xyRaglan1)
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng1 = 90 - RagAng1
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng1 = (RagAng1 * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag3 = FN_CalcLength(xyRaglan3, xyRaglan1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rag4 = System.Math.Sqrt((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(Rag4 * System.Math.Sin(RagAng1))
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(Rag4 * System.Math.Cos(RagAng1))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan4, xyRaglan3.X - adj, xyRaglan3.Y + opp)

            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng2 = (nTemplateAngle * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(RaglanBottom * System.Math.Sin(RagAng2))
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(RaglanBottom * System.Math.Cos(RagAng2))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan5, xyRaglan1.X + adj, xyRaglan1.Y + opp)
            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RagAng3 = FN_CalcAngle(xyRaglan5, xyFlapMarker)

            'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If RagAng3 > 180 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng3 = RagAng3 - 180
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf RagAng3 <= 180 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng3 = RagAng3 - 90
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng4 = 90 - RagAng3
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
            End If

            RagAng5 = RagAng4 - 13
            RagAng4 = (RagAng4 * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(RaglanTip * System.Math.Sin(RagAng4))
            'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(RaglanTip * System.Math.Cos(RagAng4))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan6, xyRaglan5.x - adj, xyRaglan5.y + opp)

            'Draw Front line of tip
            RagAng5 = (RagAng5 * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(InnerRaglanTip * System.Math.Sin(RagAng5))
            'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(InnerRaglanTip * System.Math.Cos(RagAng5))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_MakeXY(xyRaglan7, xyRaglan5.x - adj, xyRaglan5.y + opp)


            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan1)
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan2)
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan4)
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan5)
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan6)
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyRaglan7)

            'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyTmp = xyFlapMarker
            CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyTmp)

            If g_sSide = "Left" Then
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan1)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan2)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan4)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan5)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan6)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyRaglan7)
                CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyTmp)
            End If

            ARMDIA1.PR_DrawArc(xyRaglan4, xyRaglan1, xyRaglan2)
            ARMDIA1.PR_DrawLine(xyRaglan1, xyRaglan5)
            ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan6)
            ARMDIA1.PR_DrawLine(xyRaglan6, xyTmp)
            ARMDIA1.PR_SetLayer("Notes")
            ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan7)

        End If

        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
        PR_MakeXY(xyStrapText, xyMidPoint.x, xyMidPoint.y - (2 * EIGHTH))

        If Val(CType(MainForm.Controls("txtStrapLength"), Object).Text) <> 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Strap = "STRAP = " & Trim(ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)))) & "\" & QQ
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Strap = "STRAP = 24\" & QQ
        End If
        If Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text) > 0 Then
            'Display FRONT strap length if given
            'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Strap = "BACK  STRAP = " & Trim(ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text)))) & "\" & QQ
            'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Strap = Strap & "\nFRONT STRAP = " & Trim(ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text)))) & "\" & QQ
        End If
        If Val(CType(MainForm.Controls("txtWaistCir"), Object).Text) > 0 Then
            'Display Waist Circumference for D-Style flaps only
            'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Strap = Strap & "\nWAIST = " & Trim(ARMEDDIA1.fnInchesToText(ARMEDDIA1.fnDisplayToInches(Val(CType(MainForm.Controls("txtWaistCir"), Object).Text)))) & "\" & QQ
        End If

        ARMDIA1.PR_DrawText(Strap, xyStrapText, 0.1)

        'UPGRADE_WARNING: Couldn't resolve default property of object xfabby. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xfabby = CType(MainForm.Controls("cboFlaps"), Object).Text
        'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xfablen = Len(xfabby)
        'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xfabdist = (xfablen * 0.075) + 0.1
        'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_MakeXY(xyFlapText1, xyFlapMarker.x - (xfabdist * 0.75), xyFlapMarker.y - 0.5)
        CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyFlapText1)
        If g_sSide = "Left" Then CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyFlapText1)
        ARMDIA1.PR_DrawText((CType(MainForm.Controls("cboFlaps"), Object).Text), xyFlapText1, 0.1)

        'Flap Marker
        CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyFlapMarker)
        If g_sSide = "Left" Then CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyFlapMarker)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & "closed arrow" & QC & "xyStart.x +" & Str(xyFlapMarker.x) & CC & "xyStart.y +" & Str(xyFlapMarker.y) & ",0.25,0.1," & aMarker & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")

        'draw the fold line
        ARMDIA1.PR_SetLayer("Template" & g_sSide)

        PR_MakeXY(xyPt1, xyStrt.x, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_MakeXY(xyPt2, xyStrt.x + LastTapeValue, 0)

        CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyPt1)
        CADUTILS.PR_RotatePoint(xyStrt, aAngle, xyPt2)
        If g_sSide = "Left" Then
            CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyPt1)
            CADUTILS.PR_MirrorPointInYaxis(xyMidPoint.x, 0, xyPt2)
        End If

        If Left(CType(MainForm.Controls("cboFlaps"), Object).Text, 6) = "Raglan" Then
            ARMDIA1.PR_DrawLine(xyPt1, xyRaglan2)
        Else
            ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        End If

    End Sub

    Sub PR_UpdateDDE_Extension()
        'Procedure to update the fields used when data is transfered
        'from DRAFIX using DDE.
        'Although the transfer back to DRAFIX is not via DDE we use the same controls
        'simply to illustrate the method by which the data is packed into the
        'fields

        'This routine is similar to PR_UpdateDDE in MANGLOVE.FRM#
        'it is split so that we can acccess the Module level variables
        Dim iLen, ii As Short
        Dim sLen, sPacked As String

        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!SSTab1.Tab. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("SSTab1"), Object).Tab = 1


        'Pack Glove extensions
        sPacked = "   " 'Allow for possible extension to -4 tape
        For ii = 8 To 23
            iLen = Val(CType(MainForm.Controls("txtExtCir"), Object)(ii).Text) * 10 'Shift decimal place

            If iLen <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(iLen)), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If

            sPacked = sPacked & sLen

        Next ii
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!txtTapeLengthPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(MainForm.Controls("txtTapeLengthPt1"), Object).Text = sPacked

        'Pack Grams, MMs and Reductions
        Dim sTapeMMs, sGrams, sReduction As String

        sGrams = "   " 'Allow for possible extension to -4 tape
        sTapeMMs = "   " 'Allow for possible extension to -4 tape
        sReduction = "   " 'Allow for possible extension to -4 tape

        For ii = 1 To 16
            'Grams
            If g_iGms(ii) <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(g_iGms(ii))), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sGrams = sGrams & sLen

            'Tape pressures
            If g_iMMs(ii) <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(g_iMMs(ii))), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sTapeMMs = sTapeMMs & sLen

            'Reductions
            If g_iRed(ii) <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(g_iRed(ii))), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sReduction = sReduction & sLen

        Next ii

        CType(MainForm.Controls("txtReduction"), Object).Text = sReduction
        CType(MainForm.Controls("txtTapeMMs"), Object).Text = sTapeMMs
        CType(MainForm.Controls("txtGrams"), Object).Text = sGrams

        'Fold, variations etc etc.
        sPacked = ""

        'Fold
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optFold(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("optFold"), Object)(0).Value = True Then sPacked = " 0" Else sPacked = " 1"

        'Pressure
        sLen = New String(" ", 2)
        sLen = RSet(Trim(Str(CType(MainForm.Controls("cboPressure"), Object).SelectedIndex)), Len(sLen))
        sPacked = sPacked & sLen

        'Wrist Tape
        sLen = RSet(Trim(Str(CType(MainForm.Controls("cboDistalTape"), Object).SelectedIndex)), Len(sLen))
        sPacked = sPacked & sLen

        'EOS Tape
        sLen = RSet(Trim(Str(CType(MainForm.Controls("cboProximalTape"), Object).SelectedIndex)), Len(sLen))
        sPacked = sPacked & sLen

        'Variations
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("optExtendTo"), Object)(0).Value = True Then
            sPacked = sPacked & " 0 0" 'Second Zero is w.r.t Flap
            'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf CType(MainForm.Controls("optExtendTo"), Object)(1).Value = True Then
            sPacked = sPacked & " 1 0" 'Second Zero is w.r.t Flap
        Else
            sPacked = sPacked & " 2"
            'Check for flaps
            If CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Then sPacked = sPacked & " 1" Else sPacked = sPacked & " 0"
        End If

        CType(MainForm.Controls("txtDataGlove"), Object).Text = sPacked

        'Flap
        sPacked = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optExtendTo(2).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(MainForm.Controls("optExtendTo"), Object)(2).Value = True And CType(MainForm.Controls("optProximalTape"), Object)(1).Checked = True Then
            'Flap Type
            sLen = New String(" ", 3)
            sLen = RSet(Trim(Str(CType(MainForm.Controls("cboFlaps"), Object).SelectedIndex)), Len(sLen))
            sPacked = sPacked & sLen

            sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtStrapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
            sPacked = sPacked & sLen

            sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtFrontStrapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
            sPacked = sPacked & sLen

            sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtCustFlapLength"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
            sPacked = sPacked & sLen

            sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWaistCir"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
            sPacked = sPacked & sLen

        End If

        CType(MainForm.Controls("txtFlap"), Object).Text = sPacked

        'Pleats
        sLen = New String(" ", 3)
        If Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) > 0 Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWristPleat1"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
        sPacked = sLen

        sLen = New String(" ", 3)
        If Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtWristPleat2"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
        CType(MainForm.Controls("txtWristPleat"), Object).Text = sPacked & sLen

        sLen = New String(" ", 3)
        If Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) > 0 Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtShoulderPleat1"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
        sPacked = sLen

        sLen = New String(" ", 3)
        If Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) Then sLen = RSet(Trim(Str(Val(CType(MainForm.Controls("txtShoulderPleat2"), Object).Text) * 10)), Len(sLen)) 'Shift decimal place
        CType(MainForm.Controls("txtShoulderPleat"), Object).Text = sPacked & sLen

    End Sub
End Module
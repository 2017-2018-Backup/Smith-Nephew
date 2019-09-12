Attribute VB_Name = "CADGLEXT"
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
    Option Explicit
    
    Type TapeData
        nCir        As Double
        iMMs        As Integer
        iRed        As Integer
        iGms        As Integer
        sNote       As String
        iTapePos    As Integer
        sTapeText   As String
    End Type

   'Fingers
    Global Const DIP = 1
    Global Const PIP = 2
    Global Const THUMB = 3
    Global Const THUMB_LEN = 2
    Global Const FINGER_LEN = 1
    Global Const g_sTapeText = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"


   'Tapes etc
    Global Const WRIST = 1
    Global Const PALM = 2
    Global Const TAPE_ONE_HALF = 3
    Global Const TAPE_THREE = 4
    Global Const TAPE_FOUR_HALF = 5
    Global Const TAPE_SIX = 6
    Global Const TAPE_SEVEN_HALF = 7
    Global Const TAPE_NINE = 8

   'Misc
    Global Const EIGHTH = 0.125
    Global Const SIXTEENTH = 0.0625
    Global Const QUARTER = 0.25

   'Reduction Chart constants and Variables
    Const NOFF_FINGER_RED = 32
    Const NOFF_ARM_RED = 76
    Const NOFF_LENGTH_RED = 24
    Const FIGURED_MINIMUM = 6

   'Hand Reduction charts
    Dim iFingerRed(3, NOFF_FINGER_RED) As Integer
    Dim nArmRed(8, NOFF_ARM_RED) As Double
    Dim nLengthRed(2, NOFF_LENGTH_RED) As Double

   'Arm reduction chart
    Global Const LOW_MODULUS = 160
    Global Const HIGH_MODULUS = 340
    Const NOFF_MODULUS = 19

    Dim g_iModulusIndex As Integer
    Dim g_iPowernet(NOFF_MODULUS, 23)  As Integer


   'Arm variables
    Global Const NOFF_ARMTAPES = 16
    Global Const ELBOW_TAPE = 9
    Global Const ARM_PLAIN = 0
    Global Const ARM_FLAP = 1
    Global Const GLOVE_NORMAL = 0
    Global Const GLOVE_ELBOW = 1
    Global Const GLOVE_AXILLA = 2

Static Sub PR_EnableFigureArm()

   'Converse of PR_DisableFigureArm
    
    Dim ii%
    
    For ii = 3 To 11
        MainForm!lblArm(ii).Enabled = True
    Next ii

    For ii = 0 To 3
        MainForm!lblPleat(ii).Enabled = True
    Next ii
    
    MainForm!txtShoulderPleat1.Enabled = True
    MainForm!txtShoulderPleat2.Enabled = True
    If g_nPleats(1) > 0 Then MainForm!txtWristPleat1 = g_nPleats(1)
    If g_nPleats(2) > 0 Then MainForm!txtWristPleat2 = g_nPleats(2)
    MainForm!txtWristPleat1.Enabled = True
    MainForm!txtWristPleat2.Enabled = True
    If g_nPleats(3) > 0 Then MainForm!txtShoulderPleat2 = g_nPleats(3)
    If g_nPleats(4) > 0 Then MainForm!txtShoulderPleat1 = g_nPleats(4)

    MainForm!frmCalculate.Enabled = True
    MainForm!cboPressure.Enabled = True
    MainForm!cboPressure.ListIndex = g_iPressure
    MainForm!cmdCalculate.Enabled = True

End Sub

Static Sub PR_ExtendTo_Click(INDEX As Integer)
    Dim ii%, iFirst%, iLast%, iStart%, iElbow%
    
    iLast = 23
    iFirst = 8
    iElbow = 16 'Elbow w.r.t. txtExtCir()
    iStart = 8

    Select Case INDEX
        Case 0       'Normal Glove
           'Disable tapes above those required
            For ii = iStart To iLast
                MainForm!txtExtCir(ii).Enabled = False
                MainForm!lblTape(ii).Enabled = False
                PR_GrdInchesDisplay ii - 8, 0
            Next ii

           'Disable all MMs fields etc
            For ii = iFirst To iLast
                MainForm!mms(ii).Enabled = False
                MainForm!mms(ii) = ""
                PR_GramRedDisplay ii - 8, 0, 0
            Next ii

            PR_DisableFigureArm
            PR_DisableGloveToAxilla
           
           'set disable and disable
            MainForm!cboDistalTape.ListIndex = -1
            MainForm!cboDistalTape.Enabled = False
            MainForm!cboProximalTape.ListIndex = -1
            MainForm!cboProximalTape.Enabled = False
        
        Case 1     'Glove to Elbow
           'Enable MMs fields etc
            For ii = iFirst To iLast
                If ii <= iElbow Then
                    MainForm!txtExtCir(ii).Enabled = True
                    MainForm!mms(ii).Enabled = True
                    MainForm!lblTape(ii).Enabled = True
                Else
                    MainForm!txtExtCir(ii).Enabled = False
                    MainForm!lblTape(ii).Enabled = False
                    MainForm!mms(ii).Enabled = True
                    MainForm!mms(ii) = ""
                    PR_GramRedDisplay ii - 8, 0, 0
                    PR_GrdInchesDisplay ii - 8, 0
                End If
            Next ii
           
           'Enable other bits
            PR_EnableFigureArm
            PR_DisableGloveToAxilla
           
           'Set Wrist and EOS pointers
            If g_iNumTapesWristToEOS > 0 Then
               'Set wrist to given tape
                If g_iWristPointer > g_iFirstTape Then
                    MainForm!cboDistalTape.ListIndex = g_iWristPointer
                Else
                   'set to first tape
                    MainForm!cboDistalTape.ListIndex = 0
                End If
    
               'Set EOS to given tape or to Elbow if it extends
               'past the elbow
               'NB. The order of tapes is reversed in this list
               'starting at 19-1/2 and finishing at 0
                If g_iEOSPointer < ELBOW_TAPE Then
                    If g_iEOSPointer >= g_iLastTape Then
                        MainForm!cboProximalTape.ListIndex = 0
                    Else
                        MainForm!cboProximalTape.ListIndex = 17 - g_iEOSPointer
                    End If
                ElseIf g_iEOSPointer > ELBOW_TAPE Or g_iLastTape > ELBOW_TAPE Then
                    MainForm!cboProximalTape.ListIndex = 17 - ELBOW_TAPE
                Else
                   'set to last tape
                    MainForm!cboProximalTape.ListIndex = 0
                End If
            Else
                    MainForm!cboDistalTape.ListIndex = 0
                    MainForm!cboProximalTape.ListIndex = 0
            End If

            MainForm!cboDistalTape.Enabled = True
            MainForm!cboProximalTape.Enabled = True

        Case 2     'Glove to Axilla
            For ii = iFirst To iLast
                MainForm!mms(ii).Enabled = True
                MainForm!txtExtCir(ii).Enabled = True
                MainForm!lblTape(ii).Enabled = True
'                MainForm!mms(ii) = ""
'                PR_GramRedDisplay ii - 8, 0, 0
            Next ii

           'Enable other bits
            PR_EnableFigureArm

            MainForm!frmGloveToAxilla.Enabled = True
            If g_EOSType = ARM_FLAP Then MainForm!optProximalTape(1).Value = True Else MainForm!optProximalTape(0).Value = True
            
            
            MainForm!optProximalTape(0).Enabled = True
            MainForm!optProximalTape(1).Enabled = True
            
            If g_iNumTapesWristToEOS > 0 Then
               'Set wrist to given tape
                If g_iWristPointer > g_iFirstTape Then
                    MainForm!cboDistalTape.ListIndex = g_iWristPointer
                Else
                   'set to first tape
                    MainForm!cboDistalTape.ListIndex = 0
                End If
    
               'Set EOS to given tape
               'NB. The order of tapes is reversed in this list
               'starting at 19-1/2 and finishing at 0
                If g_iEOSPointer < g_iLastTape Then
                    MainForm!cboProximalTape.ListIndex = 17 - g_iEOSPointer
                Else
                   'set to first tape
                    MainForm!cboProximalTape.ListIndex = 0
                End If
            Else
                For ii = iFirst To iLast
                    MainForm!mms(ii) = ""
                    PR_GramRedDisplay ii - 8, 0, 0
                Next ii
                MainForm!cboDistalTape.ListIndex = 0
                MainForm!cboProximalTape.ListIndex = 0
            End If

            MainForm!cboDistalTape.Enabled = True
            MainForm!cboProximalTape.Enabled = True

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
   'NOTE:
   '    We ignore the fact that there may be missing tapes.
   '
   '    g_iWristPointer and g_iEOSPointer are used to indicate the
   '    start and finish tapes in the arrays and not in the
   '    txtExtCir() array

    Dim nCir#
    Dim ii%

    If MainForm!SSTab1.Tab <> 1 Then MainForm!SSTab1.Tab = 1

    g_iFirstTape = -1
    g_iLastTape = -1
    g_iWristPointer = -1
    g_iEOSPointer = -1
    g_iNumTotalTapes = 0
    g_iNumTapesWristToEOS = 0

   'Glove type
    If MainForm!optExtendTo(0).Value = True Then
        g_ExtendTo = GLOVE_NORMAL
    ElseIf MainForm!optExtendTo(1).Value = True Then
        g_ExtendTo = GLOVE_ELBOW
    Else
        g_ExtendTo = GLOVE_AXILLA
    End If

    

    If g_ExtendTo = GLOVE_NORMAL Then
       'Get values from Hand Data TAB
       'We will need to update this w.r.t CAD Glove later
       'Assummes no holes in data
        If g_nCir(2) <= 0 Then
            g_iNumTotalTapes = 1
            g_iNumTapesWristToEOS = 1
        ElseIf g_nCir(3) > 0 Then
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
            nCir = FN_InchesValue(MainForm!txtExtCir(ii))
            If nCir > 0 And g_iFirstTape = -1 Then g_iFirstTape = ii - 7
            g_nCir(ii - 7) = nCir
        Next ii
        
        For ii = 23 To 8 Step -1
            If Val(MainForm!txtExtCir(ii)) > 0 Then
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
        Dim iIndex%
       'Wrist tape
       'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
        iIndex = MainForm!cboDistalTape.ListIndex
        If iIndex = 0 Or iIndex = -1 Then
            g_iWristPointer = g_iFirstTape
        Else
            g_iWristPointer = iIndex
        End If
    
       'EOS tape
       'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
        iIndex = MainForm!cboProximalTape.ListIndex
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
        If MainForm!txtWristPleat1.Enabled = True Then g_nPleats(1) = Val(MainForm!txtWristPleat1.Text) Else g_nPleats(1) = 0
        If MainForm!txtWristPleat2.Enabled = True Then g_nPleats(2) = Val(MainForm!txtWristPleat2.Text) Else g_nPleats(2) = 0
        If MainForm!txtShoulderPleat2.Enabled = True Then g_nPleats(3) = Val(MainForm!txtShoulderPleat2.Text) Else g_nPleats(3) = 0
        If MainForm!txtShoulderPleat1.Enabled = True Then g_nPleats(4) = Val(MainForm!txtShoulderPleat1.Text) Else g_nPleats(4) = 0

        g_iPressure = MainForm!cboPressure.ListIndex
    
    End If


   'End of support type
    g_EOSType = ARM_PLAIN
    If g_ExtendTo = GLOVE_AXILLA And MainForm!optProximalTape(1).Value Then
        g_EOSType = ARM_FLAP
        g_sFlapType = MainForm!cboFlaps.Text
        g_iFlapType = MainForm!cboFlaps.ListIndex
        If MainForm!txtStrapLength.Text <> "" Then g_nStrapLength = FN_InchesValue(MainForm!txtStrapLength)
        If MainForm!txtFrontStrapLength.Text <> "" Then g_nFrontStrapLength = FN_InchesValue(MainForm!txtFrontStrapLength)
        If MainForm!txtCustFlapLength.Text <> "" Then g_nCustFlapLength = FN_InchesValue(MainForm!txtCustFlapLength)
        If MainForm!txtWaistCir.Text <> "" Then g_nWaistCir = FN_InchesValue(MainForm!txtWaistCir)
    End If
    
   'On or off fold
    g_OnFold = False


End Sub

Sub PR_GetExtensionDDE_Data()
   'Procedure to use the data given in the Form for the
   'glove extension (if any) along the arm
   '
    Dim nAge As Integer, Flap%
    Dim ii%, jj%, nValue#, nn%
    Dim iMMs%, iGms%, iRed%
    Dim sDiag As String
    
    MainForm!SSTab1.Tab = 1

   'Defaults
    MainForm!optExtendTo(0).Value = True     'Normal Glove
    MainForm!cboDistalTape.ListIndex = 0     '1st Tape
    MainForm!cboProximalTape.ListIndex = 0   'Last Tape
    Flap = False

    
   'Set default pressure w.r.t diagnosis
    nAge = Val(MainForm!txtAge.Text)

   ' if pressure empty set from diagnosis
    sDiag = UCase$(Mid$(MainForm!txtDiagnosis.Text, 1, 6))
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
        If InStr(MainForm!txtDiagnosis.Text, "2") > 0 Then
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
        If InStr(MainForm!txtDiagnosis.Text, "30") > 0 Then
            g_iPressure = 0
        ElseIf InStr(MainForm!txtDiagnosis.Text, "40") > 0 Then
            g_iPressure = 1
        ElseIf InStr(MainForm!txtDiagnosis.Text, "50") > 0 Then
            g_iPressure = 2
        End If
   'Vein Removal
    ElseIf sDiag = "VEIN R" Then
        g_iPressure = 1
    End If

  'Glove extensions
  'Code for extended glove
  'Glove Option Buttons
   Dim nLen#

   If MainForm!txtDataGlove.Text <> "" Then
     For ii = 0 To 5
        nValue = Val(Mid(MainForm!txtDataGlove.Text, (ii * 2) + 1, 2))
        If ii = 0 Then
           'Fold options
           'Do nothing Off fold is default
        ElseIf ii = 1 And nValue >= 0 Then
           'Pressure w.r.t Figuring and MMs
            g_iPressure = nValue
        ElseIf ii = 2 Then
           'Wrist Tape
            MainForm!cboDistalTape.ListIndex = nValue
        ElseIf ii = 3 Then
           'Proximal Tape
            MainForm!cboProximalTape.ListIndex = nValue
        ElseIf ii = 4 Then
            MainForm!optExtendTo(nValue).Value = True
            If nValue = 2 Then Flap = True
        ElseIf ii = 5 Then
           'Only if glove to axilla is set do we
           'do anything
            If Flap And nValue = 1 Then
                'Flap, break up flap multiple field
                 MainForm!optProximalTape(1).Value = True
                 PR_ProximalTape_Click (1)
                 For jj = 0 To 4
                    nValue = Val(Mid(MainForm!txtFlap.Text, (jj * 3) + 1, 3))
                    If jj = 0 Then
                        MainForm!cboFlaps.ListIndex = nValue
                    ElseIf jj = 1 And nValue > 0 Then
                        g_nStrapLength = nValue / 10
                        MainForm!txtStrapLength = g_nStrapLength
                        PR_DisplayTextInches MainForm!txtStrapLength, MainForm!labStrap
                    ElseIf jj = 2 And nValue > 0 Then
                        g_nFrontStrapLength = nValue / 10
                        MainForm!txtFrontStrapLength = g_nFrontStrapLength
                        PR_DisplayTextInches MainForm!txtFrontStrapLength, MainForm!labFrontStrapLength
                    ElseIf jj = 3 And nValue > 0 Then
                        g_nCustFlapLength = nValue / 10
                        MainForm!txtCustFlapLength = g_nCustFlapLength
                        PR_DisplayTextInches MainForm!txtCustFlapLength, MainForm!labCustFlapLength
                    ElseIf jj = 4 And nValue > 0 Then
                        g_nWaistCir = nValue / 10
                        MainForm!txtWaistCir = g_nWaistCir
                        PR_DisplayTextInches MainForm!txtWaistCir, MainForm!labWaistCir
                    End If
                 Next jj
            Else
                'Disable flaps
                 PR_ProximalTape_Click (0)
            End If
        End If
    Next ii
   End If

  'Set value for pressure
   MainForm!cboPressure.ListIndex = g_iPressure

  'Glove to elbow and Glove to axilla
  'These can start at -3 (However to allow for possible
  'changes later we save up to -4-1/2 but ignore -4-1/2 for the
  'mean time)
   If MainForm!txtTapeLengthPt1.Text <> "" Then
    ii = 1
    For nn = 8 To 23
        nValue = Val(Mid(MainForm!txtTapeLengthPt1.Text, (ii * 3) + 1, 3))
        If nValue > 0 Then
            MainForm!txtExtCir(nn) = nValue / 10
            nLen = FN_InchesValue(MainForm!txtExtCir(nn))
            If nLen <> -1 Then PR_GrdInchesDisplay nn - 8, nLen
            g_nCir(ii) = nLen
        Else
            PR_GrdInchesDisplay nn - 8, 0
        End If
        iMMs = Val(Mid(MainForm!txtTapeMMs.Text, (ii * 3) + 1, 3))
        If iMMs > 0 Then
            iGms = Val(Mid(MainForm!txtGrams.Text, (ii * 3) + 1, 3))
            iRed = Val(Mid(MainForm!txtReduction.Text, (ii * 3) + 1, 3))
            MainForm!mms(nn) = iMMs
            PR_GramRedDisplay nn - 8, iGms, iRed
            g_iMMs(ii) = iMMs
            g_iRed(ii) = iRed
            g_iGms(ii) = iGms
        End If
        ii = ii + 1
    Next nn
   End If

  'Pleats
    For ii = 0 To 1
        nValue = Val(Mid(MainForm!txtWristPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtWristPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtWristPleat2 = nValue / 10
        End If
        nValue = Val(Mid(MainForm!txtShoulderPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtShoulderPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtShoulderPleat2 = nValue / 10
        End If
    Next ii
    

End Sub

Sub PR_GramRedDisplay(iIndex As Integer, GRAM As Integer, Reduction As Integer)
    MainForm!grdDisplay.Row = iIndex
   
    MainForm!grdDisplay.Col = 0
    If GRAM = 0 Then MainForm!grdDisplay = "" Else MainForm!grdDisplay = GRAM
    
    MainForm!grdDisplay.Col = 1
    If Reduction = 0 Then MainForm!grdDisplay = "" Else MainForm!grdDisplay = Reduction

End Sub

Sub PR_GrdInchesDisplay(iIndex As Integer, nLen As Double)
    MainForm!grdInches.Row = iIndex
    MainForm!grdInches.Col = 0
    MainForm!grdInches = fnInchestoText(nLen)
End Sub

Static Sub PR_LoadPowernetChart()
   'Procedure to load the reduction charts for the arms from
   'disk
   '
    Dim sFile As String, sModulus As String, sLine As String
    Dim fChart%, ii%, jj%, nn%

    sFile = g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    
   'Get ARM chart reductions
   'NB these are fixed format
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_MODULUS
        Input #fChart, sModulus, sLine
        
        For jj = 0 To 22
            g_iPowernet(ii, jj + 1) = Val(Mid(sLine, (jj * 4) + 1, 4))
        Next jj

        ii = ii + 1
          
    Wend
    
   'Close file
    Close fChart
    sLine = ""
    sModulus = ""

End Sub

Static Sub PR_LoadReductionCharts(sFabric As String)
   'Procedure to load the reduction charts from
   'disk
   'Two charts are loaded
   '    1. Length chart
   '    2. Circumferences based on the Fabric
   '
   
    Dim sFile As String, sLine As String
    Dim iModulus%, fChart%, ii%, jj%, nn%

    If sFabric = "" Then Exit Sub
    If UCase$(Mid$(sFabric, 1, 3)) <> "POW" Then
        MsgBox "Fabric chosen is not Powernet", 48, g_sDialogID
        Exit Sub
    End If

   'Establish chart to be loaded
   '    Fabric Format  Pow MMM-XX Comment
   '    if mm < 230 use 230 chart
   '    if MM > 280 use 280 chart
   '
    iModulus = Val(Mid$(sFabric, 5, 3))
    If iModulus < 230 Then iModulus = 230
    If iModulus > 280 Then iModulus = 280
    sFile = g_sPathJOBST & "\TEMPLTS\GLV_" & Trim$(Str$(iModulus)) & ".DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Fabric Chart, Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    
   'Get finger and thumb reducations
   'NB these are fixed format
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_FINGER_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB use of "." to repeat previous number
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then

            iFingerRed(1, ii) = Val(Mid(sLine, 5, 2))
            
            If Mid$(sLine, 8, 1) = "." Then
                iFingerRed(2, ii) = iFingerRed(1, ii)
            Else
                iFingerRed(2, ii) = Val(Mid(sLine, 8, 2))
            End If

            If Mid$(sLine, 11, 1) = "." Then
                iFingerRed(3, ii) = iFingerRed(2, ii)
            Else
                iFingerRed(3, ii) = Val(Mid(sLine, 11, 2))
            End If

            ii = ii + 1
        End If
    Wend
    
   'Wrist and Palm to 9 tape reductions
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_ARM_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB: use of "." to repeat previous number
       '    also translation from inches and eights to decimal inches)
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then
            nArmRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
            jj = 1
            For nn = 9 To 34 Step 4
                jj = jj + 1
                If Mid$(sLine, nn, 1) = "." Then
                    nArmRed(jj, ii) = nArmRed(jj - 1, ii)
                Else
                    nArmRed(jj, ii) = Val(Mid(sLine, nn, 2)) + (Val(Mid(sLine, nn + 2, 1)) * 0.125)
                End If
            Next nn
            ii = ii + 1
        End If
    Wend

   'Close file
    Close fChart


   'Length Reduction chart
    sFile = g_sPathJOBST & "\TEMPLTS\GLV_LEN.DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_LENGTH_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB: use of "." to repeat previous number
       '    also translation from inches and eights to decimal inches)
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then
            
            nLengthRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
            
            If Mid$(sLine, 9, 1) = "." Then
                nLengthRed(2, ii) = nLengthRed(1, ii)
            Else
                nLengthRed(2, ii) = Val(Mid(sLine, 9, 2)) + (Val(Mid(sLine, 11, 1)) * 0.125)
            End If

            ii = ii + 1
        End If
    Wend

   'Close file
    Close fChart

    sLine = ""

End Sub

Sub PR_ProximalTape_Click(INDEX As Integer)
   Dim ii%
   If INDEX = 0 Then
       'Disable flaps
        g_iFlapType = MainForm!cboFlaps.ListIndex
        MainForm!cboFlaps.Enabled = False
        MainForm!cboFlaps.ListIndex = -1
        For ii = 0 To 4
            MainForm!lblFlap(ii).Enabled = False
        Next ii
        MainForm!txtStrapLength.Enabled = False
        g_nStrapLength = Val(MainForm!txtStrapLength.Text)
        MainForm!txtStrapLength.Text = ""
        MainForm!labStrap.Caption = ""

        MainForm!txtFrontStrapLength.Enabled = False
        g_nFrontStrapLength = Val(MainForm!txtFrontStrapLength.Text)
        MainForm!txtFrontStrapLength.Text = ""
        MainForm!labFrontStrapLength.Caption = ""
        
        MainForm!txtCustFlapLength.Enabled = False
        g_nCustFlapLength = Val(MainForm!txtCustFlapLength.Text)
        MainForm!txtCustFlapLength.Text = ""
        MainForm!labCustFlapLength.Caption = ""
        
        MainForm!txtWaistCir.Enabled = False
        g_nWaistCir = Val(MainForm!txtWaistCir.Text)
        MainForm!txtWaistCir.Text = ""
        MainForm!labWaistCir.Caption = ""

   Else
       'Enable flaps
        MainForm!cboFlaps.Enabled = True
        MainForm!cboFlaps.ListIndex = g_iFlapType

        For ii = 0 To 3
            MainForm!lblFlap(ii).Enabled = True
        Next ii
        MainForm!txtStrapLength.Enabled = True
        If g_nStrapLength > 0 Then
            MainForm!txtStrapLength = g_nStrapLength
            PR_DisplayTextInches MainForm!txtStrapLength, MainForm!labStrap
        End If

        MainForm!txtFrontStrapLength.Enabled = True
        If g_nFrontStrapLength > 0 Then
            MainForm!txtFrontStrapLength = g_nFrontStrapLength
            PR_DisplayTextInches MainForm!txtFrontStrapLength, MainForm!labFrontStrapLength
        End If

        MainForm!txtCustFlapLength.Enabled = True
        If g_nCustFlapLength > 0 Then
            MainForm!txtCustFlapLength = g_nCustFlapLength
            PR_DisplayTextInches MainForm!txtCustFlapLength, MainForm!labCustFlapLength
        End If

        If InStr(1, MainForm!cboFlaps.Text, "D") > 0 Then
            MainForm!lblFlap(4).Enabled = True
            MainForm!txtWaistCir.Enabled = True
            If g_nWaistCir > 0 Then
                MainForm!txtWaistCir = g_nWaistCir
                PR_DisplayTextInches MainForm!txtWaistCir, MainForm!labWaistCir
            End If
        End If

   End If

End Sub

Sub PR_SetMMs(sPressure As String)
   'Set the mms based on the pressure and the wrist and EOS
   'tapes given
   'REF:    GOP 01-02/16, Section 1.4
   '
    Dim ii%

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

Sub PR_SetModulusIndex(sFabric As String)
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
    Dim iMod%

    iMod = Val(Mid(sFabric, 5, 3))
    If iMod < LOW_MODULUS Then
        g_iModulusIndex = LOW_MODULUS
    ElseIf iMod > HIGH_MODULUS Then
        g_iModulusIndex = HIGH_MODULUS
    Else
        g_iModulusIndex = ((iMod - LOW_MODULUS) + 10) / 10
    End If
    
End Sub

Function FN_GetGrams(ByVal iReduction As Integer) As Integer
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

Static Function FN_GetReduction(iGrams) As Integer
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

    Dim ii%, iValue%, iPrevValue%
    
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
                If iValue > iGrams Then Exit For
                iPrevValue = iValue
            Next ii

            If Abs(iGrams - iPrevValue) < Abs(iGrams - iValue) Then
                FN_GetReduction = ii + 8
            Else
                FN_GetReduction = ii + 9
            End If

    End Select


End Function

Static Function FN_LengthWristToEOS() As Double
   'Calculates the distance from the EOS to the
   'wrist based on values given in the dialogue
   '
    Dim ii%, nLen#, nSpace#

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
            If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
            If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
            If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
            nLen = nLen + nSpace
        Next ii
     
     Case GLOVE_AXILLA
        For ii = 1 To g_iNumTapesWristToEOS - 1
            nSpace = 1.375
            If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
            If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
            If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
            nLen = nLen + nSpace
        Next ii
        If g_EOSType = ARM_FLAP Then
        End If
    End Select

    FN_LengthWristToEOS = nLen

End Function

Static Function FN_ValidateExtensionData() As String
   'Procedure to validated the extension data
    Dim sError As String, nL$
    Dim iNoPleats%, ii%
    
    nL = Chr$(13)
    sError = ""     'Initialise because of static

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
    If g_nCir(g_iWristPointer) <> g_nWrist Then
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
    Dim ii%

   'Don't Calculate if normal glove
    If MainForm!optExtendTo(0).Value = True Then Exit Sub

    If MainForm!cboFabric.Text = "" Then
            MsgBox "Fabric not given! ", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
    Else
        If UCase$(Mid$(MainForm!cboFabric.Text, 1, 3)) <> "POW" Then
            MsgBox "Fabric chosen is not Powernet!", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
        End If
        If Val(Mid$(MainForm!cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid$(MainForm!cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
            MsgBox "Modulus of fabric chosen is not available on Gram / Tension reduction chart!", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
        End If

    End If

   
   'Update from dialogue
    PR_GetDlgAboveWrist
    
    If g_iPressure < 0 Then
        MsgBox "Can't Calculate, missing Pressure", 48, "CAD Glove - Calculate ARM Button"
        Exit Sub
    End If

    If g_iNumTapesWristToEOS <= 1 Then
        MsgBox "Can't Calculate, Only one tape between given Wrist and EOS", 48, "CAD Glove - Calculate ARM Button"
        Exit Sub
    End If
   
   'Check that there are no holes in the data
    For ii = g_iWristPointer To g_iEOSPointer
        If g_nCir(ii) = 0 Then
            MsgBox "Can't Calculate, missing Tape Circumferences between given Wrist and EOS", 48, "CAD Glove - Calculate ARM Button"
            Exit Sub
        End If
    Next ii


   'Set the MMs based on the selected pressure
    PR_SetMMs (MainForm!cboPressure.Text)

   'Set the modulus based on the fabric
    PR_SetModulusIndex (MainForm!cboFabric.Text)

   'Calculate grams etc
    For ii = 1 To NOFF_ARMTAPES
        If ii >= g_iWristPointer And ii <= g_iEOSPointer Then
            
            g_iGms(ii) = round(g_nCir(ii) * g_iMMs(ii))
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
                g_iMMs(ii) = round(g_iGms(ii) / g_nCir(ii))
            End If

           'Adjust at EOS  g_iWristPointer
            If ii = g_iWristPointer Then
               'Always use a 10 reduction for wrist
                g_iRed(ii) = 10
                g_iGms(ii) = FN_GetGrams(g_iRed(ii))
                g_iMMs(ii) = round(g_iGms(ii) / g_nCir(ii))
            End If

           'Display Results
            PR_GrdInchesDisplay ii - 1, g_nCir(ii)
            PR_GramRedDisplay ii - 1, g_iGms(ii), g_iRed(ii)
            MainForm!mms(ii + 7) = g_iMMs(ii)
           
        Else
           'Blank the displayed values
            PR_GrdInchesDisplay ii - 1, 0
            PR_GramRedDisplay ii - 1, 0, 0
            MainForm!mms(ii + 7).Text = ""
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
     
    Dim nMidWristX As Double, nValue As Double, nSpacing As Double
    Dim ii%, nFiguredValue As Double, iStartTape%, iLastTape%
    Dim iVertex%, nInsert#, nOriginalWrist#, nRevisedWrist#

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
        nFiguredValue = (g_nCir(iStartTape) * ((100 - g_iRed(iStartTape)) / 100))
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
            nFiguredValue = (g_nCir(ii) * ((100 - g_iRed(ii)) / 100))
            nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
            UlnarProfile.x(iVertex) = nMidWristX - (nValue / 2)
            RadialProfile.x(iVertex) = nMidWristX + (nValue / 2)

           'Setup the notes for the vertex
           'We do this here as we have all the data and it simplifies
           'the drawing side
            TapeNote(iVertex).sTapeText = LTrim$(Mid$(g_sTapeText, ((ii + 1) * 3) + 1, 3))
            TapeNote(iVertex).iTapePos = ii
            TapeNote(iVertex).nCir = g_nCir(ii)
            TapeNote(iVertex).iGms = g_iGms(ii)
            TapeNote(iVertex).iRed = g_iRed(ii)
            TapeNote(iVertex).iMMs = g_iMMs(ii)


           'Standard spacing
            nSpacing = 1.375
           
           'Account all for pleats
           'Wrist (as we have started at the wrist we use, iStartTape + 1
           'and iStartTape + 2 in this case)
            If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(1))
            If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(2))
'XXXXXXXXXX
'XXXX Careful now! this won't work if only 3 or less tapes given
'XXXXXXXXXX
           'Axilla
            If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(3))
            If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(4))
    
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
     
        PR_CalcMidPoint xyPalm(1), xyPalm(6), xyPalm(7)
        
       End If

End Sub

Static Sub PR_CalcWristBlendLFS(xyArcCen As XY, aArcSweep#, xyArcStart As XY, xyBSt As XY, xyBEnd As XY, ReturnProfile As Curve)
   '
   'N.B. Parameters are given in the order of decreasing Y
   '
    Dim nL1#, nL2#, nL3#, aA1#, rAngle#, aInc#, nR1#, nR2#
    Dim xyCenter As XY, xyMidPoint As XY, xyTopSt As XY, xyTopEnd As XY, xyBotSt As XY, xyBotEnd As XY
    Dim xyR2  As XY, nThirdOfL2#, nA#, aAngle#, xyArc As XY
    Dim xyPt1 As XY, xyPt2 As XY, xyTEnd As XY, xyArcSt As XY
    Dim ii%, MirrorResult%, nRadius#, Direction%, TopIsArc%, BottomIsArc%
    Dim xyTmp(1 To 10) As XY, nTol#

   'Do this as we can't use ByVal
    xyBotSt = xyBSt
    xyBotEnd = xyBEnd
    xyArc = xyArcCen
    xyArcSt = xyArcStart

   'Get the point on the arc that will be
   'used as xyTEnd
    aAngle = FN_CalcAngle(xyArc, xyArcSt) + ((aArcSweep * 3) / 4)
    nRadius = FN_CalcLength(xyArc, xyArcSt)
    PR_CalcPolar xyArc, aAngle, nRadius, xyTopEnd
    xyTEnd = xyTopEnd
     
   'Degenerate to a straight line
   'Then exit the sub routine        '????????????????
    If FN_CalcAngle(xyBSt, xyTEnd) = FN_CalcAngle(xyBEnd, xyBSt) Then
        ReturnProfile.n = 2
        ReturnProfile.x(1) = xyTEnd.x
        ReturnProfile.y(1) = xyTEnd.y
        ReturnProfile.x(2) = xyBEnd.x
        ReturnProfile.y(2) = xyBEnd.y
        Exit Sub
    End If

    nL2 = FN_CalcLength(xyBSt, xyTEnd)
    nL3 = FN_CalcLength(xyBEnd, xyBSt)

   'Get Included angles & radius & Centers of Arcs
    aAngle = FN_CalcAngle(xyArc, xyArcSt)
    aInc = aArcSweep / 4
    For ii = 2 To 4
        aAngle = aAngle + aInc
        PR_CalcPolar xyArc, aAngle, nRadius, xyTmp(ii)
    Next ii

    nThirdOfL2 = nL2 / 3

    'Bottom arc
    nA = FN_CalcLength(xyTopEnd, xyBotEnd)
    nA = ((nL2 ^ 2 + nL3 ^ 2) - nA ^ 2) / (2 * nL3 * nL2)
    rAngle = Arccos(nA) / 2
    nR2 = nThirdOfL2 * Tan(rAngle)
    PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(5)
    PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(8)
    
    'Establish
    aAngle = FN_CalcAngle(xyBotEnd, xyTopEnd)
    If aAngle < FN_CalcAngle(xyBotEnd, xyBotSt) Then
        aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) - 90
    Else
        aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) + 90
    End If
    PR_CalcPolar xyTmp(5), aAngle, nR2, xyR2
    
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
        PR_CalcPolar xyR2, aAngle, nR2, xyTmp(ii)
    Next ii



    ReturnProfile.n = 9
    ReturnProfile.x(1) = xyArcStart.x
    ReturnProfile.y(1) = xyArcStart.y
    ReturnProfile.x(9) = xyBotEnd.x
    ReturnProfile.y(9) = xyBotEnd.y

    For ii = 2 To 8
        ReturnProfile.x(ii) = xyTmp(ii).x
        ReturnProfile.y(ii) = xyTmp(ii).y
    Next ii

End Sub

Static Sub PR_CalcWristBlendThumbSide(xyTSt As XY, xyTEnd As XY, xyBSt As XY, xyBEnd As XY, ReturnProfile As Curve)
   '
   'N.B. Parameters are given in the order of decreasing Y
   '
    Dim xyTmp As XY
    Dim nL1#, nL#, aAngle#, ii%

    nL1 = FN_CalcLength(xyBSt, xyTEnd) / 3
    aAngle = FN_CalcAngle(xyBSt, xyTEnd)
    
    PR_CalcMidPoint xyTSt, xyTEnd, xyTmp

    ReturnProfile.n = 6
    
    ReturnProfile.x(1) = xyTSt.x
    ReturnProfile.y(1) = xyTSt.y

    ReturnProfile.x(2) = xyTmp.x
    ReturnProfile.y(2) = xyTmp.y

    ReturnProfile.x(3) = xyTEnd.x
    ReturnProfile.y(3) = xyTEnd.y

    nL = 0
    nL1 = 0.125
    For ii = 5 To 4 Step -1
        nL = nL + nL1
        PR_CalcPolar xyBSt, aAngle, nL, xyTmp
        ReturnProfile.x(ii) = xyTmp.x
        ReturnProfile.y(ii) = xyTmp.y
    Next ii

    ReturnProfile.x(6) = xyBSt.x
    ReturnProfile.y(6) = xyBSt.y
        

End Sub

Sub PR_DisableFigureArm()
    
    Dim ii%

    For ii = 3 To 11
        MainForm!lblArm(ii).Enabled = False
    Next ii

   'Pleats
    g_nPleats(1) = Val(MainForm!txtWristPleat1.Text)
    g_nPleats(2) = Val(MainForm!txtWristPleat2.Text)
    g_nPleats(3) = Val(MainForm!txtShoulderPleat2.Text)
    g_nPleats(4) = Val(MainForm!txtShoulderPleat1.Text)

    For ii = 0 To 3
        MainForm!lblPleat(ii).Enabled = False
    Next ii
    

    MainForm!txtShoulderPleat1.Enabled = False
    MainForm!txtShoulderPleat2.Enabled = False
  '  MainForm!txtShoulderPleat1 = ""
  '  MainForm!txtShoulderPleat2 = ""
    MainForm!txtWristPleat1.Enabled = False
    MainForm!txtWristPleat2.Enabled = False
  '  MainForm!txtWristPleat1 = ""
  '  MainForm!txtWristPleat2 = ""

    MainForm!frmCalculate.Enabled = False
    MainForm!cboPressure.Enabled = False
    MainForm!cmdCalculate.Enabled = False

End Sub

Sub PR_DisableGloveToAxilla()

    MainForm!frmGloveToAxilla.Enabled = False
    MainForm!optProximalTape(0).Value = False
    MainForm!optProximalTape(1).Value = False
    MainForm!optProximalTape(0).Enabled = False
    MainForm!optProximalTape(1).Enabled = False
    PR_ProximalTape_Click (0)

End Sub

Static Sub PR_DisplayTextInches(ctlText As Control, ctlCaption As Control)
    Dim nLen As Double
    nLen = FN_InchesValue(ctlText)
    If nLen <> -1 Then ctlCaption.Caption = fnInchestoText(nLen)
End Sub

Static Sub PR_DrawShoulderFlaps(xyS As XY, xyE As XY)
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
Dim ShoulderFlap As Curve, RaglanFlap As Curve
Dim LastTapeValue, xyPt1 As XY, xyPt2 As XY
Dim xyTopFlap1 As XY, xyTopFlap2 As XY, xyTopFlap3 As XY, xyTopFlap4 As XY, xyTopFlap5 As XY
Dim xyTopFlap6 As XY, xyTopFlap7 As XY, xyTopFlap8 As XY, xyTopFlap9 As XY, xyTopFlap10 As XY
Dim xyBotFlap1 As XY, xyBotFlap2 As XY, xyBotFlap3 As XY, xyBotFlap4 As XY, xyBotFlap5 As XY
Dim xyFlap1 As XY, xyFlap2 As XY, xyFlap3 As XY, xyFlap4 As XY, xyFlap5 As XY
Dim BETA, Omega, Phi, Delta, Theta, BottomRadius, Alpha, TopArcIncrement
Dim x1, y1, h1, adj, opp, HYP, LittleBit, Change, TopRadius, FlapMarkerAngle, xyFlapText1  As XY
Dim xyMid As XY, xyCentre As XY, FlapLength, FlapMarkerLength, xyFlapMarker As XY, xyFlapText2 As XY, xyStrapText As XY
Dim ArcLength, CircleCircum, TopArcLength, BottomArcLength, TopArcRadius, BottomArcRadius, TopArcAngle, BottomArcAngle
Dim TempMarker, xfabby, xfablen, xfabdist
Dim xyRaglan1 As XY, xyRaglan2 As XY, xyRaglan3 As XY, xyRaglan4 As XY, xyRaglan5 As XY, xyRaglan6 As XY, xyRaglan7 As XY
Dim Rag1, Rag2, Rag3, Rag4, RagAng1, RagAng2, RagAng3, RaglanBottom, RaglanTip, TopRagLanAngle, InnerRaglanTip
Dim nNotchOffset As Double, nTemplateRadius#, nTemplateAngle#, nNotchToTangent#
Dim ii%, PhiExtra#, RagAng4#, RagAng5#, aMarker#
Dim Strap$
Dim aAngle#, nLength#, xyMidPoint As XY, xyStrt As XY, xyTmp As XY


    PR_CalcMidPoint xyS, xyE, xyMidPoint
    nLength = FN_CalcLength(xyS, xyE)
    aAngle = -90
    BottomArcAngle = 47.5
   'All this crap because you can't use ByVal on user defined types (Bastards!)
   'VB3 comment
    If g_sSide = "Right" Then
        xyStrt = xyE
        aMarker = 135
    Else
        xyStrt = xyS
        aMarker = 45
    End If


   ' Check for Custom Flap Length
    If Val(MainForm!txtCustFlapLength) > 0 Then
        LastTapeValue = fnDisplayToInches(MainForm!txtCustFlapLength)
    Else
       'Standard flap length
        LastTapeValue = g_nCir(g_iEOSPointer)
        LastTapeValue = (LastTapeValue / 3.14) * 0.92
    End If

    
    ' Calculate Bottom part of curve with 5 Points
    If nLength >= 5.5 Then
        BottomRadius = LastTapeValue - 0.625
        RaglanBottom = 3.5
        TopRagLanAngle = 17
        RaglanTip = 1.9375
        InnerRaglanTip = 2.173
        nNotchOffset = 0.875   '14/16 ths Raglan Template s287
    ElseIf (nLength < 5.5) And (nLength > 3) Then
        BottomRadius = 2.8
        RaglanBottom = 2.5
        TopRagLanAngle = 17
        RaglanTip = 1.05
        InnerRaglanTip = 1.22
        nNotchOffset = 0.75     '12/16 ths Raglan Template s289
    ElseIf (nLength <= 3) Then
        BottomRadius = 2
        RaglanBottom = 2.0625
        RaglanTip = 0.625
        TopRagLanAngle = 15
        InnerRaglanTip = 0.727
        nNotchOffset = 0.6875   '11/16 ths, Raglan Template s290
    End If

    nTemplateRadius = 3.5
    nTemplateAngle = 40
    nNotchToTangent = 2.125

    'Calc length of bottom Arc
    BottomArcRadius = BottomRadius
    CircleCircum = PI * (2 * (BottomArcRadius))
    BottomArcLength = (BottomArcAngle / 360) * CircleCircum
    
    ShoulderFlap.x(1) = xyStrt.x + LastTapeValue
    ShoulderFlap.y(1) = 0

    For ii = 1 To 4
        Theta = (ii * 11.875 * PI) / 180
        adj = Cos(Theta) * BottomRadius
        opp = Sin(Theta) * BottomRadius
        LittleBit = BottomRadius - adj
        ShoulderFlap.x(ii + 1) = xyStrt.x + LastTapeValue - LittleBit
        ShoulderFlap.y(ii + 1) = 0 + opp
    Next ii
    PR_MakeXY xyBotFlap5, ShoulderFlap.x(5), ShoulderFlap.y(5)
    
    'Calculate 6 angles and points for top part of curve
    PR_MakeXY xyTopFlap1, xyStrt.x, xyStrt.y + nLength
    ShoulderFlap.x(10) = xyTopFlap1.x
    ShoulderFlap.y(10) = xyTopFlap1.y

    'midpoint
    h1 = FN_CalcLength(xyTopFlap1, xyBotFlap5)
    h1 = h1 / 2
    Omega = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Omega = 180 - Omega
    Omega = (Omega * PI) / 180
    opp = Abs(h1 * Sin(Omega))
    adj = Abs(h1 * Cos(Omega))
    PR_MakeXY xyMid, xyBotFlap5.x - adj, xyBotFlap5.y + opp
   
    'Centre of Top Arc
    Delta = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Delta = Delta - 90
    Delta = (Delta * PI) / 180
    HYP = 4 * h1
    opp = Abs(HYP * Sin(Delta))
    adj = Abs(HYP * Cos(Delta))
    PR_MakeXY xyCentre, xyMid.x + adj, xyMid.y + opp
    TopRadius = FN_CalcLength(xyCentre, xyTopFlap1)
    TopArcRadius = TopRadius
    Phi = FN_CalcAngle(xyCentre, xyTopFlap1)
    Phi = Phi - 180
    Alpha = FN_CalcAngle(xyCentre, xyBotFlap5)
    Alpha = Alpha - 180
    TopArcAngle = Alpha - Phi
    TopArcIncrement = (Alpha - Phi) / 5

   'Calc Flap Length for position of marker
    CircleCircum = PI * (2 * (TopArcRadius))
    TopArcLength = (TopArcAngle / 360) * CircleCircum
    FlapLength = TopArcLength + BottomArcLength
    FlapMarkerLength = FlapLength / 2.7
    If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
    If FlapMarkerLength < TopArcLength Then
        FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
        FlapMarkerAngle = FlapMarkerAngle + Phi
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(TopRadius * Sin(FlapMarkerAngle))
        adj = Abs(TopRadius * Cos(FlapMarkerAngle))
        PR_MakeXY xyFlapMarker, xyCentre.x - adj, xyCentre.y - opp
    ElseIf FlapMarkerLength > TopArcLength Then
        TempMarker = FlapMarkerLength - TopArcLength
        TempMarker = BottomArcLength - TempMarker
        FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(BottomArcRadius * Sin(FlapMarkerAngle))
        adj = Abs(BottomArcRadius * Cos(FlapMarkerAngle))
        LittleBit = BottomRadius - adj
        PR_MakeXY xyFlapMarker, xyStrt.x + LastTapeValue - LittleBit, 0 + opp
    ElseIf FlapMarkerLength = TopArcLength Then
        PR_MakeXY xyFlapMarker, xyBotFlap5.x, xyBotFlap5.y
    End If
    PR_SetLayer "Template" & g_sSide

    For ii = 1 To 4
        PhiExtra = Phi + (ii * TopArcIncrement)
        PhiExtra = (PhiExtra * PI) / 180
        opp = Abs(TopRadius * Sin(PhiExtra))
        adj = Abs(TopRadius * Cos(PhiExtra))
        ShoulderFlap.x(10 - ii) = xyCentre.x - adj
        ShoulderFlap.y(10 - ii) = xyCentre.y - opp
    Next ii

    
    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
       'Copy standard shoulder from profile down
       'only go one vertex past the FlapMarker
        RaglanFlap.n = 0
        For ii = 10 To 1 Step -1
            RaglanFlap.n = RaglanFlap.n + 1
            RaglanFlap.x(RaglanFlap.n) = ShoulderFlap.x(ii)
            RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
            If ShoulderFlap.x(ii) > xyFlapMarker.x Then Exit For
        Next ii
        PR_RotateCurve xyStrt, aAngle, RaglanFlap
        If g_sSide = "Left" Then PR_MirrorCurveInYaxis xyMidPoint.x, 0, RaglanFlap
        PR_DrawFitted RaglanFlap
    Else
        ShoulderFlap.n = 10
        PR_RotateCurve xyStrt, aAngle, ShoulderFlap
        If g_sSide = "Left" Then PR_MirrorCurveInYaxis xyMidPoint.x, 0, ShoulderFlap
        PR_DrawFitted ShoulderFlap
    End If
 
    'Draw Raglan
    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
        Rag1 = Sqr((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
        Rag1 = BottomRadius - Rag1
        PR_MakeXY xyRaglan1, xyStrt.x + LastTapeValue - Rag1, nNotchOffset
        PR_MakeXY xyRaglan2, xyRaglan1.x - nNotchToTangent, 0
        RagAng1 = FN_CalcAngle(xyRaglan2, xyRaglan1)
        Rag2 = FN_CalcLength(xyRaglan2, xyRaglan1)
        Rag2 = Rag2 / 2
        RagAng1 = (RagAng1 * PI) / 180
        opp = Abs(Rag2 * Sin(RagAng1))
        adj = Abs(Rag2 * Cos(RagAng1))
        PR_MakeXY xyRaglan3, xyRaglan2.x + adj, xyRaglan2.y + opp
        
        RagAng1 = FN_CalcAngle(xyRaglan3, xyRaglan1)
        RagAng1 = 90 - RagAng1
        RagAng1 = (RagAng1 * PI) / 180
        Rag3 = FN_CalcLength(xyRaglan3, xyRaglan1)
        Rag4 = Sqr((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
        opp = Abs(Rag4 * Sin(RagAng1))
        adj = Abs(Rag4 * Cos(RagAng1))
        PR_MakeXY xyRaglan4, xyRaglan3.x - adj, xyRaglan3.y + opp
        
        RagAng2 = (nTemplateAngle * PI) / 180
        opp = Abs(RaglanBottom * Sin(RagAng2))
        adj = Abs(RaglanBottom * Cos(RagAng2))
        PR_MakeXY xyRaglan5, xyRaglan1.x + adj, xyRaglan1.y + opp
        RagAng3 = FN_CalcAngle(xyRaglan5, xyFlapMarker)
        
        If RagAng3 > 180 Then
            RagAng3 = RagAng3 - 180
            RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
        ElseIf RagAng3 <= 180 Then
            RagAng3 = RagAng3 - 90
            RagAng4 = 90 - RagAng3
            RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
        End If
        
        RagAng5 = RagAng4 - 13
        RagAng4 = (RagAng4 * PI) / 180
        opp = Abs(RaglanTip * Sin(RagAng4))
        adj = Abs(RaglanTip * Cos(RagAng4))
        PR_MakeXY xyRaglan6, xyRaglan5.x - adj, xyRaglan5.y + opp
        
        'Draw Front line of tip
        RagAng5 = (RagAng5 * PI) / 180
        opp = Abs(InnerRaglanTip * Sin(RagAng5))
        adj = Abs(InnerRaglanTip * Cos(RagAng5))
        PR_MakeXY xyRaglan7, xyRaglan5.x - adj, xyRaglan5.y + opp
        

        PR_RotatePoint xyStrt, aAngle, xyRaglan1
        PR_RotatePoint xyStrt, aAngle, xyRaglan2
        PR_RotatePoint xyStrt, aAngle, xyRaglan4
        PR_RotatePoint xyStrt, aAngle, xyRaglan5
        PR_RotatePoint xyStrt, aAngle, xyRaglan6
        PR_RotatePoint xyStrt, aAngle, xyRaglan7

        xyTmp = xyFlapMarker
        PR_RotatePoint xyStrt, aAngle, xyTmp

        If g_sSide = "Left" Then
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan1
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan2
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan4
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan5
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan6
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyRaglan7
            PR_MirrorPointInYaxis xyMidPoint.x, 0, xyTmp
        End If

        PR_DrawArc xyRaglan4, xyRaglan1, xyRaglan2
        PR_DrawLine xyRaglan1, xyRaglan5
        PR_DrawLine xyRaglan5, xyRaglan6
        PR_DrawLine xyRaglan6, xyTmp
        PR_SetLayer "Notes"
        PR_DrawLine xyRaglan5, xyRaglan7

    End If
    
    PR_SetLayer "Notes"
    PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
    PR_MakeXY xyStrapText, xyMidPoint.x, xyMidPoint.y - (2 * EIGHTH)
    
    If Val(MainForm!txtStrapLength) <> 0 Then
        Strap = "STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtStrapLength)))) & "\" & QQ
    Else
        Strap = "STRAP = 24\" & QQ
    End If
    If Val(MainForm!txtFrontStrapLength) > 0 Then
      'Display FRONT strap length if given
       Strap = "BACK  STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtStrapLength)))) & "\" & QQ
       Strap = Strap & "\nFRONT STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtFrontStrapLength)))) & "\" & QQ
    End If
    If Val(MainForm!txtWaistCir) > 0 Then
      'Display Waist Circumference for D-Style flaps only
       Strap = Strap & "\nWAIST = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtWaistCir)))) & "\" & QQ
    End If

    PR_DrawText Strap, xyStrapText, 0.1
    
    xfabby = MainForm!cboFlaps.Text
    xfablen = Len(xfabby)
    xfabdist = (xfablen * 0.075) + 0.1
    PR_MakeXY xyFlapText1, xyFlapMarker.x - (xfabdist * 0.75), xyFlapMarker.y - 0.5
    PR_RotatePoint xyStrt, aAngle, xyFlapText1
    If g_sSide = "Left" Then PR_MirrorPointInYaxis xyMidPoint.x, 0, xyFlapText1
    PR_DrawText (MainForm!cboFlaps.Text), xyFlapText1, 0.1

   'Flap Marker
    PR_RotatePoint xyStrt, aAngle, xyFlapMarker
    If g_sSide = "Left" Then PR_MirrorPointInYaxis xyMidPoint.x, 0, xyFlapMarker
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "closed arrow"; QC; "xyStart.x +"; Str$(xyFlapMarker.x); CC; "xyStart.y +"; Str$(xyFlapMarker.y); ",0.25,0.1,"; aMarker; ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
    
   'draw the fold line
    PR_SetLayer "Template" & g_sSide
    
    PR_MakeXY xyPt1, xyStrt.x, 0
    PR_MakeXY xyPt2, xyStrt.x + LastTapeValue, 0
    
    PR_RotatePoint xyStrt, aAngle, xyPt1
    PR_RotatePoint xyStrt, aAngle, xyPt2
    If g_sSide = "Left" Then
        PR_MirrorPointInYaxis xyMidPoint.x, 0, xyPt1
        PR_MirrorPointInYaxis xyMidPoint.x, 0, xyPt2
    End If

    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
        PR_DrawLine xyPt1, xyRaglan2
    Else
        PR_DrawLine xyPt1, xyPt2                        'Fold Line
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
    Dim iLen%, ii%
    Dim sLen$, sPacked$

    MainForm!SSTab1.Tab = 1


   'Pack Glove extensions
    sPacked = "   "     'Allow for possible extension to -4 tape
    For ii = 8 To 23
        iLen = Val(MainForm!txtExtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii
    MainForm!txtTapeLengthPt1 = sPacked

   'Pack Grams, MMs and Reductions
    Dim sGrams$, sTapeMMs$, sReduction$

    sGrams = "   "      'Allow for possible extension to -4 tape
    sTapeMMs = "   "    'Allow for possible extension to -4 tape
    sReduction = "   "  'Allow for possible extension to -4 tape
    
    For ii = 1 To 16
       'Grams
        If g_iGms(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iGms(ii)))
        Else
            sLen = String(3, " ")
        End If
        sGrams = sGrams & sLen

       'Tape pressures
        If g_iMMs(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iMMs(ii)))
        Else
            sLen = String(3, " ")
        End If
        sTapeMMs = sTapeMMs & sLen

       'Reductions
        If g_iRed(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iRed(ii)))
        Else
            sLen = String(3, " ")
        End If
        sReduction$ = sReduction$ & sLen

    Next ii

    MainForm!txtReduction.Text = sReduction
    MainForm!txtTapeMMs.Text = sTapeMMs
    MainForm!txtGrams.Text = sGrams

   'Fold, variations etc etc.
    sPacked = ""

   'Fold
    If MainForm!optFold(0).Value = True Then sPacked = " 0" Else sPacked = " 1"
    
   'Pressure
    sLen = String(2, " ")
    RSet sLen = Trim$(Str$(MainForm!cboPressure.ListIndex))
    sPacked = sPacked & sLen

   'Wrist Tape
    RSet sLen = Trim$(Str$(MainForm!cboDistalTape.ListIndex))
    sPacked = sPacked & sLen

   'EOS Tape
    RSet sLen = Trim$(Str$(MainForm!cboProximalTape.ListIndex))
    sPacked = sPacked & sLen

   'Variations
    If MainForm!optExtendTo(0).Value = True Then
        sPacked = sPacked & " 0 0" 'Second Zero is w.r.t Flap
    ElseIf MainForm!optExtendTo(1).Value = True Then
        sPacked = sPacked & " 1 0" 'Second Zero is w.r.t Flap
    Else
        sPacked = sPacked & " 2"
       'Check for flaps
        If MainForm!optProximalTape(1).Value = True Then sPacked = sPacked & " 1" Else sPacked = sPacked & " 0"
    End If
    
    MainForm!txtDataGlove.Text = sPacked

   'Flap
    sPacked = ""
    If MainForm!optExtendTo(2).Value = True And MainForm!optProximalTape(1).Value = True Then
       'Flap Type
        sLen = String(3, " ")
        RSet sLen = Trim$(Str$(MainForm!cboFlaps.ListIndex))
        sPacked = sPacked & sLen
  
        RSet sLen = Trim$(Str$(Val(MainForm!txtStrapLength.Text) * 10))    'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtFrontStrapLength.Text) * 10))  'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtCustFlapLength.Text) * 10))   'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtWaistCir.Text) * 10))  'Shift decimal place
        sPacked = sPacked & sLen

    End If

    MainForm!txtFlap.Text = sPacked

   'Pleats
    sLen = String(3, " ")
    If Val(MainForm!txtWristPleat1.Text) > 0 Then RSet sLen = Trim$(Str$(Val(MainForm!txtWristPleat1.Text) * 10)) 'Shift decimal place
    sPacked = sLen

    sLen = String(3, " ")
    If Val(MainForm!txtWristPleat2.Text) Then RSet sLen = Trim$(Str$(Val(MainForm!txtWristPleat2.Text) * 10)) 'Shift decimal place
    MainForm!txtWristPleat.Text = sPacked & sLen
    
    sLen = String(3, " ")
    If Val(MainForm!txtShoulderPleat1.Text) > 0 Then RSet sLen = Trim$(Str$(Val(MainForm!txtShoulderPleat1.Text) * 10)) 'Shift decimal place
    sPacked = sLen

    sLen = String(3, " ")
    If Val(MainForm!txtShoulderPleat2.Text) Then RSet sLen = Trim$(Str$(Val(MainForm!txtShoulderPleat2.Text) * 10)) 'Shift decimal place
    MainForm!txtShoulderPleat.Text = sPacked & sLen

End Sub


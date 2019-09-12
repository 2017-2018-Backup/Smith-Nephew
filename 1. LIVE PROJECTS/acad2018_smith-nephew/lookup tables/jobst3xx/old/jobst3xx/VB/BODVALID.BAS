Attribute VB_Name = "BODVALID"
'Module:   BODVALID.MAK
'Purpose:   Contains data validation and calculation
'
'Projects:  1. MainForm.MAK
'           2. BODYDRAW.MAK
'
'
'Version:   1.00
'Date:      28.August.1997
'Author:    Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
'
'
'
    Option Explicit

   'MsgBox constant
    Const IDCANCEL = 2
    Const IDYES = 6
    Const IDNO = 7

    Global g_bResponse
    Global g_bDiffThigh             As Integer
    Global g_bMissCutOut7           As Integer
    Global g_bExtremeCrotch         As Integer
    Global g_bDiffAxillaHeight      As Integer
    Global g_bSleeveless            As Integer
    Global g_bDrawSingleLeg         As Integer

    Global g_sSex As String
    Global g_sFileNo As String
    Global g_sSide As String
    Global g_sPatient As String
    Global g_sID As String
    Global g_sDiagnosis As String
    Global g_nAge As Integer
    Global g_nAdult As Integer
    Global g_sWorkOrder As String
    Global g_nUnitsFac As Double

    Global g_sBackNeck As String
    Global g_sFrontNeck As String

    Global g_nUnderBreast As Double
    Global g_nNipple As Double
    Global g_nChest As Double
    Global g_sLeftCup As String
    Global g_sRightCup As String

    Global g_sLeftDisk As String
    Global g_sRightDisk As String
    
    Global g_nLeftShldCirGiven As Double
    Global g_nRightShldCirGiven As Double
    Global g_nLeftAxillaCirGiven As Double
    Global g_nRightAxillaCirGiven As Double
    Global g_nChestCirGiven As Double
    Global g_nUnderBreastCirGiven As Double
    Global g_nWaistCirGiven As Double
    Global g_nButtocksCirGiven As Double
    Global g_nLeftThighCirGiven As Double
    Global g_nRightThighCirGiven As Double
    Global g_nNeckCirGiven As Double
    Global g_nShldToFoldGiven As Double
    Global g_nShldToWaistGiven As Double
    Global g_nShldToBreastGiven As Double
    Global g_nLeftShldToAxilla As Double
    Global g_nRightShldToAxilla As Double
    Global g_nNippleCirGiven As Double
    Global g_sFrontNeckStyle As String
    Global g_nFrontNeckSize As Double
    Global g_nRadiusFrontNeck As Double
    Global g_sBackNeckStyle As String
    Global g_nBackNeckSize As Double
    Global g_sClosure As String
    Global g_sFabric As String
    Global g_sCrotchStyle As String

    Global g_sSmallestThighGiven As String
    Global g_nLargestThighGiven As Double
    Global g_sLargestThighGiven As String

    Global g_nShldWidth As Double
    Global g_nLeftThighCir As Double
    Global g_nRightThighCir As Double
    Global g_sLeftSleeve As String
    Global g_sSleeveType As String
    Global g_sRightSleeve As String
    Global g_sLeftLeg As String
    Global g_sRightLeg As String
    Global g_nShldToAxilla As Double
    Global g_nAxillaCir As Double
    Global g_nThighCir As Double
    Global g_nGroinHeight As Double
    Global g_nLeftAxillaCir As Double
    Global g_nRightAxillaCir As Double
    Global g_nChestCir As Double
    Global g_nUnderBreastCir As Double
    Global g_nWaistCir As Double
    Global g_nButtocksCir As Double
    Global g_nNeckCir As Double
    Global g_nShldToFold As Double
    Global g_nShldToWaist As Double
    Global g_nShldToBreast As Double
    Global g_nButtocksLength As Double
    Global g_nButtBackSeamRatio As Double
    Global g_nButtFrontSeamRatio As Double
    Global g_nButtRedIncreased As Double
    Global g_nThighRedDecreased As Double
    Global g_nButtCir As Double
    Global g_nHalfGroinHeight As Double
    Global g_nButtRadius As Double
    Global g_nLT_GroinHeight As Double
    Global g_nLT_HalfGroinHeight As Double
    Global g_nLT_ButtRadius As Double
    Global g_n85FoldHeight As Double
    Global g_nButtCirCalc As Double
    Global g_nButtFrontSeam As Double
    Global g_nButtBackSeam As Double
    Global g_nButtCutOut As Double
    Global g_nWaistFrontSeam As Double
    Global g_nWaistBackSeam As Double
    Global g_nWaistCutOut As Double
    Global g_nChestFrontSeam As Double
    Global g_nChestBackSeam As Double
    Global g_nChestCutOut As Double
    Global g_nSmallestThighGiven As Double
    Global g_nChestCirRed As Double
    Global g_nUnderBreastCirRed As Double
    Global g_nWaistCirRed As Double
    Global g_nLargeButtCirRed As Double
    Global g_nThighCirRed As Double
    Global g_nNeckCirRed As Double
    Global g_nShldToFoldRed As Double
    Global g_nShldToWaistRed As Double
    Global g_nShldToBreastRed As Double
    Global g_nButtocksCirRed As Double
    Global g_nCutOutRadius As Double
    Global g_sFlySize As String
    Global g_sGussetSize As String
    Global g_nGussetLength As Double
    Global g_nFrontCrotchSize As Double
    Global g_nBackCrotchSize As Double
    Global g_bDrawBriefCurve As Integer
    Global g_nCutOut As Double
    Global g_nCutOutToBackMinimum   As Double
    Global g_sLegStyle   As String
    Global g_nRightLegLength As Double
    Global g_nLeftLegLength As Double
    Global g_bRightAboveKnee As Integer
    Global g_bLeftAboveKnee As Integer

    Global g_sMeshLeft As String
    Global g_sMeshRight As String
   

    Global g_nAxillaBackNeckRad As Double
    Global g_nAxillaFrontNeckRad As Double
    Global g_nABNRadRight As Double
    Global g_nAFNRadRight As Double

   'Co-ordinate points
    Global xySeamFold As XY
    Global xySeamThigh As XY
    Global xySeamButt As XY
    Global xySeamHighShld As XY
    Global xySeamLowShld As XY
    Global xySeamChest As XY
    Global xySeamChestAxillaLow As XY

    Global g_sAxillaSide As String
    Global g_sAxillaSideLow As String
    Global g_sAxillaType As String
    Global g_sAxillaTypeLow As String
    Global g_nShldToAxillaLow As Double
    Global g_nAxillaCirLow  As Double
    Global g_nLengthStrap1ToCutOut5  As Double
    

    Global xySeamWaist As XY
    Global xyProfileThigh As XY
    Global xyLT_ProfileThigh As XY
    Global xyProfileThighExtraPT As XY
    Global xyFold As XY
    Global xyLT_Fold As XY
    Global xyProfileButt As XY
    Global xyCutOut3 As XY
    Global xyCutOut4 As XY
    Global xyCutOut5 As XY
    Global xyCutOut6 As XY
    Global xyCutOut2 As XY
    Global xyCutOut9 As XY
    Global xyCutOut10 As XY
    Global xyCutOut7 As XY
    Global xyProfileWaist As XY
    Global xyProfileChest As XY
    Global xyProfileNeck As XY
    Global xyCutOutBackNeck As XY
    Global xyCutOut8 As XY
    Global xyCutOutFrontNeck As XY
    Global xyFrontNeckArcCentre As XY
    Global xyRaglan1 As XY
    Global xyRaglan2 As XY
    Global xyRaglan3 As XY
    Global xyRaglan4LowAxilla As XY
    Global xyRaglan2LowAxilla As XY
    Global xyRaglan4 As XY
    Global xyRaglan5 As XY
    Global xyRaglan6 As XY
    Global xyProfileNeckMirror As XY
    Global xyCutOutArcCentre As XY
    Global xyLegPoint As XY
    Global xy85Fold As XY
    Global xyProfileGroin As XY
    Global xyLT_ProfileGroin As XY

    Global xyProfileBrief As XY
    Global xyProfileThighMirror As XY
    Global xyLT_ProfileThighMirror As XY
    Global xyProfileThighMirror1 As XY
    Global xyStrap1 As XY
    Global xyStrap2 As XY
    Global xyStrap3 As XY
    Global xyStrap4 As XY
    Global xyTmpCutOut8 As XY
    Global xyGussetArcCentre As XY
    Global xyBackCrotchFilletCentre As XY
    Global xyFrontCrotchFilletCentre As XY
    Global xyBackCrotch1 As XY
    Global xyBackCrotch2 As XY
    Global xyBackCrotch3 As XY
    Global xyFrontCrotch1 As XY
    Global xyFrontCrotch2 As XY
    Global xyFrontCrotch3 As XY
    Global xyCrotchMarker As XY
    Global xyButtocksArcCentre As XY
    Global xyLT_ButtocksArcCentre As XY
    Global xyFrenchCut As XY
'    Global xyLeftLegLowestPoint As XY
'    Global xyRightLegLowestPoint As XY
'    Global xyLargestLegLowestPoint As XY
'    Global xySmallestLegLowestPoint As XY
    Global xyLeftLegLabelPoint As XY
    Global xyRightLegLabelPoint As XY
    Global xyBothLegLabelPoint As XY


    Global xyAxilla As XY

    
    Global Const INCH1_16 = 0.0625
    Global Const INCH3_16 = 0.1875
    Global Const INCH5_16 = 0.3125
    Global Const INCH1_8 = 0.125
    Global Const INCH1_4 = 0.25
    Global Const INCH3_8 = 0.375
    Global Const INCH1_2 = 0.5
    Global Const INCH5_8 = 0.625
    Global Const INCH3_4 = 0.75
    Global Const INCH7_8 = 0.875


Function FN_ValidateAndCalculateData(bDisplayErrors As Integer) As Integer
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
    Dim iFatalError As Integer
    Dim ii As Integer
    Dim sCutOutToBackMinimum As String
    Dim nInchesDiff As Double
    Dim nWaistRem As Double
    Dim nChestRem As Double
    Dim nButtRedAtLimit As Integer
    Dim nThighRedAtLimit As Integer
    Dim nButtRedIncreased As Integer
    Dim nThighRedDecreased As Integer
    Dim nCutOutModified As Integer

    'Initialise
    FN_ValidateAndCalculateData = False
    sError = ""
    iFatalError = False
    
    
    
    ReDim sCircum(0 To 16)
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
        If Val(MainForm!txtCir(ii).Text) = 0 Then
            sError = sError & "Missing " & sCircum(ii) & NL
        End If
    Next ii
    For ii = 12 To 16
        If Val(MainForm!txtCir(ii).Text) = 0 Then
            sError = sError & "Missing " & sCircum(ii) & NL
        End If
    Next ii

   'Bra Cups
   'Note:
   '    The Circumference over nipple is optional unless a cup has been
   '    specified
   '
    If Val(MainForm!txtCir(9).Text) = 0 And Val(MainForm!txtCir(10).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(9) & NL
    End If
    If Val(MainForm!txtCir(10).Text) = 0 And Val(MainForm!txtCir(9).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(10) & NL
    End If

    If (Val(MainForm!txtCir(9).Text) <> 0 Or Val(MainForm!txtCir(10).Text) <> 0) And ((MainForm!cboLeftCup.Text <> "None" And MainForm!txtLeftDisk.Text = "") Or (MainForm!txtRightDisk.Text = "" And MainForm!cboRightCup.Text <> "None")) Then
            sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & NL
    End If

    If Val(MainForm!txtCir(9).Text) = 0 And (MainForm!txtLeftDisk.Text <> "" Or MainForm!txtRightDisk.Text <> "") Then
            sError = sError & "Bra disks given! But Missing " & sCircum(9) & NL
    End If

    'If cups or dimensions given then a disk must be present
    'NB
    '    cboXXXXCup.ListIndex = 6 = "None"
    '    cboXXXXCup.ListIndex = 7 = ""

    If MainForm!cboLeftCup.ListIndex < 6 And Val(MainForm!txtLeftDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Left BRA Cup!" & NL
    End If
    If MainForm!cboRightCup.ListIndex < 6 And Val(MainForm!txtRightDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Right BRA Cup!" & NL
    End If
    
   'Sex error
    If MainForm!txtSex.Text = "Male" And (Val(MainForm!txtCir(9).Text) <> 0 Or Val(MainForm!txtCir(10).Text) <> 0 Or MainForm!cboLeftCup.ListIndex < 0 Or MainForm!cboRightCup.ListIndex < 0) Then
            sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & NL
    End If

   'Neck at back and front
    Dim sChar As String * 1
    sChar = Left$(MainForm!cboBackNeck.Text, 1)
    If sChar = "M" And MainForm!txtBackNeck.Text = "" Then
            sError = sError & "No dimension for Measured Back neck style!" & NL
    End If

    sChar = Left$(MainForm!cboFrontNeck.Text, 1)
    If sChar = "M" And MainForm!txtFrontNeck.Text = "" Then
            sError = sError & "No dimension for Measured Front neck style" & NL
    End If

   'Get values from dialog
    g_nShldWidth = fnDisplayToInches(Val(MainForm!txtCir(3).Text))
    g_sBackNeckStyle = MainForm!cboBackNeck.Text
    g_sFrontNeckStyle = MainForm!cboFrontNeck.Text
   'Test minimum value for Shoulder width
    If g_nShldWidth < 1.5 Then
        sError = sError & "Shoulder Width is less than 1-1/2""!" & NL
    End If
    If InStr(g_sBackNeckStyle, "Scoop") > 0 And ((g_nShldWidth - 1) < 1.5) Then
        sError = sError & "With a scoop BACK neck the Shoulder Width is less than 1-1/2""!" & NL
    End If

    If (InStr(g_sFrontNeckStyle, "Scoop") > 0 Or InStr(g_sFrontNeckStyle, "V neck") > 0) And (g_nShldWidth - 1) < 1.5 Then
        sError = sError & "With a " & g_sFrontNeckStyle & " FRONT neck the Shoulder Width is less than 1-1/2""!" & NL
    End If

   '
    If MainForm!cboLeftAxilla.Text = "" Then
            sError = sError & "Left Axilla not given!" & NL
    End If

    If MainForm!cboRightAxilla.Text = "" Then
            sError = sError & "Right Axilla not given!" & NL
    End If

    If MainForm!cboFrontNeck.Text = "" Then
            sError = sError & "Neck not given!" & NL
    End If

    If MainForm!cboBackNeck.Text = "" Then
            sError = sError & "Back neck not given!" & NL
    End If

    If MainForm!cboClosure.Text = "" Then
            sError = sError & "Closure not given!" & NL
    End If

    If MainForm!cboFabric.Text = "" Then
            sError = sError & "Fabric not given!" & NL
    End If
    
   'Extra Values
    If MainForm!cboLegStyle.Text = "" Then
            sError = sError & "Leg Style not given!" & NL
    End If
    If MainForm!cboCrotchStyle.Text = "" Then
            sError = sError & "Crotch Style not given!" & NL
    End If
    
   'Display Error message (if required) and return
   'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in MainForm Data cannot continue!"
        FN_ValidateAndCalculateData = False
        Exit Function
    Else
        FN_ValidateAndCalculateData = True
    End If

   'Possible FATAL Errors and WARNINGS
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   
   'Get values from the dialogue
    PR_GetValuesFromDialogue

   'Figured Measurements
   '
    g_bDiffAxillaHeight = False
    g_bDiffThigh = False
    g_bSleeveless = False

   'Check if Axilla values are within an inch of each other
    If Abs(g_nLeftShldCirGiven - g_nRightShldCirGiven) > 1 Then 'Greater than 1" difference
            'keep separate
            g_bDiffAxillaHeight = True
        Else
            'Use average (sides same at axilla)
            g_nLeftShldCirGiven = ((g_nLeftShldCirGiven + g_nRightShldCirGiven) / 2)
            g_nRightShldCirGiven = g_nLeftShldCirGiven
    End If
    g_nLeftShldToAxilla = fnRoundInches(((g_nLeftShldCirGiven * 0.95) / 3.14)) + INCH1_2
    g_nRightShldToAxilla = fnRoundInches(((g_nRightShldCirGiven * 0.95) / 3.14)) + INCH1_2
    

    If g_sLeftSleeve = "Sleeveless" Then
            g_nLeftAxillaCir = fnRoundInches(g_nLeftShldCirGiven * 0.9) '90% figured
            g_bSleeveless = True
        Else
            g_nLeftAxillaCir = g_nLeftShldCirGiven
    End If
    If g_sRightSleeve = "Sleeveless" Then
            g_nRightAxillaCir = fnRoundInches(g_nRightShldCirGiven * 0.9) '90% figured
            g_bSleeveless = True
        Else
            g_nRightAxillaCir = g_nRightShldCirGiven
    End If

    
   'Thighs
    If Abs(g_nLeftThighCirGiven - g_nRightThighCirGiven) > 1 Then g_bDiffThigh = True
    'Use Smallest Thigh for calculations
    'We will draw both largest and smallest thigh
    If (g_nLeftThighCirGiven <= g_nRightThighCirGiven) Then
            g_nSmallestThighGiven = g_nLeftThighCirGiven
            g_sSmallestThighGiven = "Left"
            g_nLargestThighGiven = g_nRightThighCirGiven
            g_sLargestThighGiven = "Right"
        Else
            g_nSmallestThighGiven = g_nRightThighCirGiven
            g_sSmallestThighGiven = "Right"
            g_nLargestThighGiven = g_nLeftThighCirGiven
            g_sLargestThighGiven = "Left"
    End If
            
''        Else
 '          'Within an 1" of each other therefore use smallest thigh size.
 '           g_sSmallestThighGiven = "Left"
 '           If (g_nLeftThighCirGiven < g_nRightThighCirGiven) Then
 '                   g_nSmallestThighGiven = g_nLeftThighCirGiven
 '               Else
 '                   g_nSmallestThighGiven = g_nRightThighCirGiven
 '           End If
 '   End If
   

    If InStr(g_sCrotchStyle, "Fly") <> 0 And g_sSex = "Male" Then
            g_nButtBackSeamRatio = 0.6 '60%
            g_nButtFrontSeamRatio = 0.4 '40%
        Else
            g_nButtBackSeamRatio = 0.5 '50%
            g_nButtFrontSeamRatio = 0.5 '50%
    End If
    
'MsgBox "1"
   'Figure the values
    g_nChestCir = fnRoundInches((g_nChestCirGiven * g_nChestCirRed)) / 2   'half scale
    g_nWaistCir = fnRoundInches((g_nWaistCirGiven * g_nWaistCirRed)) / 2   'half-scale
    g_nShldToFold = fnRoundInches((g_nShldToFoldGiven * g_nShldToFoldRed)) + INCH1_2
    g_nShldToWaist = fnRoundInches((g_nShldToWaistGiven * g_nShldToWaistRed)) + INCH1_2
    g_nButtocksLength = fnRoundInches((g_nShldToFold - g_nShldToWaist) / 3)
    g_nUnderBreastCir = fnRoundInches((g_nUnderBreastCirGiven * g_nUnderBreastCirRed)) / 2  'half scale
    g_nNeckCir = fnRoundInches(g_nNeckCirGiven * g_nNeckCirRed)
    g_nShldToBreast = fnRoundInches((g_nShldToBreastGiven * g_nShldToBreastRed)) + INCH1_2
        
            
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
    g_nButtocksCir = fnRoundInches((g_nButtocksCirGiven * 0.85)) / 2  'half scale
    g_nGroinHeight = fnRoundInches((g_nSmallestThighGiven * 0.85)) / 2    'half scale
    g_nHalfGroinHeight = (g_nGroinHeight / 2)
    g_nButtRadius = Sqr((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
    g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
    g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
    g_n85FoldHeight = g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2))
   'Recalculate below. Above is only to get the original mark

    Do
        If g_nButtocksCirRed = 0.9 Then 'at Reduction limit of 90%
            nButtRedAtLimit = True
        End If
       'I had to use Val and Str$ functions because VB3 has a problem
       'when returning values from the minus operation,
       'i.e. nThighCirRed was not equal to .8 exactly (but it should be!!!)
        If Val(Str$(g_nThighCirRed)) = 0.8 Then 'at Reduction limit of 80%
                nThighRedAtLimit = True
        End If
        If nThighRedAtLimit And nButtRedAtLimit Then
            Exit Do
        End If
        g_nButtocksCir = fnRoundInches((g_nButtocksCirGiven * g_nButtocksCirRed)) / 2  'half scale
        g_nGroinHeight = fnRoundInches((g_nSmallestThighGiven * g_nThighCirRed)) / 2    'half scale
        g_nHalfGroinHeight = (g_nGroinHeight / 2)
        g_nButtRadius = Sqr((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
        g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
        g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
        If g_nButtBackSeam < g_nCutOutToBackMinimum Then     'less than full scale minimum
                If nButtRedAtLimit Then
                        g_nThighCirRed = g_nThighCirRed - 0.05    'decrease the reduction by 5%
                        nThighRedDecreased = True
                    Else
                        g_nButtocksCirRed = g_nButtocksCirRed + 0.05 'increase the reduction by 5%
                        nButtRedIncreased = True
                End If
            Else
                Exit Do     'greater than 3" full scale - acceptable
        End If
    Loop
'MsgBox "3"

    If nButtRedIncreased Then
        MainForm!cboRed(3).Text = g_nButtocksCirRed
        sError = sError + "With given Buttock Reduction, distance from back of Cut-Out" + NL
        sError = sError + "to largest part of buttocks is less than " + sCutOutToBackMinimum + " inches!" + NL
        sError = sError + "Buttocks Reduction increased to rectify this." + NL
        'sError = sError + "ref:- " + NL
    End If
    If nThighRedDecreased Then
        MainForm!cboRed(4).Text = g_nThighCirRed
        sError = sError + "With given Thigh Reduction, distance from back of Cut-Out" + NL
        sError = sError + "to largest part of buttocks is less than " + sCutOutToBackMinimum + " inches!" + NL
        sError = sError + "Thigh Reduction decreased to rectify this." + NL
        'sError = sError + "ref:- " + NL
    End If

   'Check that all of the reductions are within 5% of each other

    Dim nMinRed%, nMaxRed%, nCurrentValue%
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
        sError = sError + NL + "Severe - Warning!" + NL
        sError = sError + "All Reductions should be within 5% of each other" + NL
    End If

'MsgBox "4"
   
   'One last time with Reductions found above
   '
   'Calculate for Average / Smallest thigh
    g_nButtocksCir = fnRoundInches((g_nButtocksCirGiven * g_nButtocksCirRed)) / 2  'half scale
    g_nGroinHeight = fnRoundInches((g_nSmallestThighGiven * g_nThighCirRed)) / 2    'half scale
    g_nHalfGroinHeight = (g_nGroinHeight / 2)
    g_nButtRadius = Sqr((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
    g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
    g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
    g_nGroinHeight = g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - ((g_nButtocksLength - INCH3_4) ^ 2))
    
   'Calculate largest thigh (if given)
        g_nLT_GroinHeight = fnRoundInches((g_nLargestThighGiven * g_nThighCirRed)) / 2    'half scale
        g_nLT_HalfGroinHeight = (g_nLT_GroinHeight / 2)
        g_nLT_ButtRadius = Sqr((g_nButtocksLength ^ 2) + (g_nLT_HalfGroinHeight ^ 2))
        g_nLT_GroinHeight = g_nLT_HalfGroinHeight + Sqr((g_nLT_ButtRadius ^ 2) - ((g_nButtocksLength - INCH3_4) ^ 2))

    If g_nButtBackSeam <= 0 Then
            sError = sError + NL + "Severe - Warning!" + NL
            sError = sError + "Distance from back of Cut-Out to largest part of buttocks" + NL
            sError = sError + "is still negative or zero!" + NL
            sError = sError + "Even with the modified reductions." + NL
            iFatalError = True
        ElseIf g_nButtBackSeam < g_nCutOutToBackMinimum Then     'less than required minimum
            sError = sError + NL + "Severe - Warning!" + NL
            sError = sError + "Distance from back of Cut-Out to largest part of buttocks" + NL
            sError = sError + "is still less than " + sCutOutToBackMinimum + " inches!" + NL
            sError = sError + "Even with the modified reductions." + NL
    End If
   
   'All other seams test
   'Set initila values
    g_nButtFrontSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtFrontSeamRatio
    g_nButtCutOut = g_nButtCirCalc - (g_nButtBackSeam + g_nButtFrontSeam)
    
    nInchesDiff = Int(Abs(g_nWaistCirGiven - g_nButtocksCirGiven))
    g_nWaistFrontSeam = (g_nButtFrontSeam * 2)
    If (g_nWaistCirGiven > g_nButtocksCirGiven) And nInchesDiff > 0 Then
            g_nWaistFrontSeam = g_nWaistFrontSeam + (nInchesDiff * INCH1_4)
        Else
            g_nWaistFrontSeam = g_nWaistFrontSeam - (nInchesDiff * INCH1_8)
    End If
    
    nInchesDiff = round(Abs(g_nChestCirGiven - g_nButtocksCirGiven))
    g_nChestFrontSeam = (g_nButtFrontSeam * 2)
    If (g_nChestCirGiven > g_nButtocksCirGiven) And nInchesDiff > 0 Then
            g_nChestFrontSeam = g_nChestFrontSeam + (nInchesDiff * INCH1_4)
        Else
            g_nChestFrontSeam = g_nChestFrontSeam - (nInchesDiff * INCH1_8)
    End If
   
   'Use the half scale for both
    g_nChestFrontSeam = g_nChestFrontSeam / 2
    g_nWaistFrontSeam = g_nWaistFrontSeam / 2

'MsgBox "5"
'MsgBox "g_nWaistFrontSeam=" & Str$(g_nWaistFrontSeam) & NL & "g_nChestFrontSeam=" & Str$(g_nChestFrontSeam)

    
    g_nCutOut = (g_nButtCutOut * 0.87)
    nCutOutModified = 0
    Do
'MsgBox "Loop Yes"
        nWaistRem = g_nWaistCir - (g_nCutOut + (g_nWaistFrontSeam * 2))
        If (nWaistRem / 2) < g_nCutOutToBackMinimum Then 'less than allowable minimum @ full scale
                If nWaistRem <= 0 Then
                    sError = sError + NL + "Severe - Warning!" + NL
                    sError = sError + "Distance from back of Cut-Out to profile at WAIST" + NL
                    sError = sError + "is negative or zero!" + NL
                    If nCutOutModified > 0 Then
                        sError = sError + "Even with the front of the Cut-Out at the Waist lowered." + NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
               'Lower cut-out at Waist
'                g_nWaistFrontSeam = g_nWaistFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nWaistRem / 2))) / 2)
                g_nWaistFrontSeam = g_nWaistFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nWaistRem / 2))) / 1)
                nCutOutModified = nCutOutModified + 1
                If nCutOutModified > 100 Then
                    MsgBox "Infinite looping while calculating back profile at waist, Contact Systems Support!"
                    Exit Do
                End If
            Else
                If nCutOutModified > 0 Then
                    sError = sError + "Original distance from back of Cut-Out to profile at the WAIST" + NL
                    sError = sError + "was less than " + sCutOutToBackMinimum + " inches!" + NL
                    Select Case nCutOutModified
                        Case 1
                            sError = sError & "Cut-Out at WAIST had to be lowered " & Str$(nCutOutModified) & " time." & NL
                            'sError = sError + "ref:- " + NL
                        Case Is > 1
                            sError = sError & "Cut-Out at WAIST had to be lowered " & Str$(nCutOutModified) & " times." & NL
                            'sError = sError + "ref:- " + NL
                    End Select
                End If
                Exit Do     'greater than g_nCutOutToBackMinimum - acceptable
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
                    sError = sError + NL + "Severe - Warning!" + NL
                    sError = sError + "Distance from back of Cut-Out to profile at the CHEST" + NL
                    sError = sError + "is negative or zero!" + NL
                    If nCutOutModified > 0 Then
                        sError = sError + "Even with the Cut Out at the Chest lowered." + NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
               'lower cut=out at Chest
'                g_nChestFrontSeam = g_nChestFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nChestRem / 2))) / 2)
                g_nChestFrontSeam = g_nChestFrontSeam - ((2 * (g_nCutOutToBackMinimum - (nChestRem / 2))) / 1)
                nCutOutModified = nCutOutModified + 1
                If nCutOutModified > 100 Then
                    MsgBox "Infinite looping while calculating back profile at chest, Contact support"
                    Exit Do
                End If
            Else
                If nCutOutModified > 0 Then
                    sError = sError + "Distance from back of Cut-Out at the CHEST" + NL
                    sError = sError + "was less than " + sCutOutToBackMinimum + " inches!" + NL
                    Select Case nCutOutModified
                        Case 1
                            sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str$(nCutOutModified) & " time." & NL
                            'sError = sError + "ref:- " + NL
                        Case Is > 1
                            sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str$(nCutOutModified) & " times." & NL
                            'sError = sError + "ref:- " + NL
                    End Select
                End If
                Exit Do     'greater than g_nCutOutToBackMinimum - acceptable
        End If
    Loop
    g_nChestBackSeam = nChestRem / 2

   'Snap Crotch / Gusset warning wrt brief
    If InStr(g_sLegStyle, "Brief") > 0 And g_sCrotchStyle = "Gusset" Then
        sError = sError + NL + "Information!" + NL
        sError = sError + "A crotch style of ""Gusset"" has been selected for a BRIEF." + NL
        sError = sError + "You should draw as a Snap Crotch" + NL
    End If

    If InStr(g_sLegStyle, "Brief") > 0 And g_sCrotchStyle = "Open Crotch" Then
        sError = sError + NL + "Information!" + NL
        sError = sError + "A crotch style of ""Open"" has been selected for a BRIEF." + NL
        sError = sError + "You should not use an Open crotch style with a Brief" + NL
    End If
    If InStr(g_sCrotchStyle, "Hor") > 0 And g_nAge < 3 Then
        sError = sError + NL + "Information!" + NL
        sError = sError + "A Horizontal Fly has been selected for a child under 3." + NL
        sError = sError + "The horizontal fly chart does not contain an entry for children under 3 years old." + NL
    End If

   
   'Display Error message (if required) and return
   'There could be fatal errors found
    If Len(sError) > 0 And bDisplayErrors Then
            If iFatalError = True Then
                    MsgBox sError, 64, "Warning - Problems with data"
                    FN_ValidateAndCalculateData = False
                Else
                    sError = sError + NL + "The above problems have been found in the data do you" + NL
                    sError = sError + "wish to continue ?"
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

    Dim iInt As Integer, nDec As Double, iSign As Integer
   'retain sign
    iSign = Sgn(nDisplay)
    nDisplay = Abs(nDisplay)

   'Simple case where Units are CM
    If g_nUnitsFac <> 1 Then
        fnDisplayToInches = fnRoundInches(nDisplay * g_nUnitsFac) * iSign
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

    Dim iInt As Integer, iSign As Integer
    Dim nPrecision As Double, nDec As Double


   'Return 0 if input is Zero
    If nNumber = 0 Then
       fnRoundInches = 0
       Exit Function
    End If

   'Set precision
    nPrecision = 0.125

   'Break input into components
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)
    nDec = nNumber - iInt

   'Get decimal part in precision units
    If nDec <> 0 Then
        nDec = nDec / nPrecision    'Avoid overflow
    End If
    nDec = round(nDec)

   'Return value
    fnRoundInches = (iInt + (nDec * nPrecision)) * iSign
End Function

Static Sub PR_GetValuesFromDialogue()
    

   'Get values from dialog box
    g_sFileNo = MainForm!txtFileNo.Text
    g_sPatient = MainForm!txtPatientName.Text
    g_sDiagnosis = MainForm!txtDiagnosis.Text
    g_nAge = Val(MainForm!txtAge.Text)

   'Set Adult status
    If g_nAge > 10 Then
            g_nAdult = True
        Else
            g_nAdult = False
    End If
    g_sSex = MainForm!txtSex.Text
    g_sWorkOrder = MainForm!txtWorkOrder.Text
    
    g_nLeftShldCirGiven = fnDisplayToInches(Val(MainForm!txtCir(0).Text))
    g_nRightShldCirGiven = fnDisplayToInches(Val(MainForm!txtCir(1).Text))
    g_nNeckCirGiven = fnDisplayToInches(Val(MainForm!txtCir(2).Text))
    g_nShldWidth = fnDisplayToInches(Val(MainForm!txtCir(3).Text))
    g_nShldToWaistGiven = fnDisplayToInches(Val(MainForm!txtCir(4).Text))
    g_nChestCirGiven = fnDisplayToInches(Val(MainForm!txtCir(5).Text))
    g_nWaistCirGiven = fnDisplayToInches(Val(MainForm!txtCir(6).Text))
    g_nShldToFoldGiven = fnDisplayToInches(Val(MainForm!txtCir(12).Text))
    g_nButtocksCirGiven = fnDisplayToInches(Val(MainForm!txtCir(14).Text))
    g_nLeftThighCirGiven = fnDisplayToInches(Val(MainForm!txtCir(15).Text))
    g_nRightThighCirGiven = fnDisplayToInches(Val(MainForm!txtCir(16).Text))
    If MainForm!txtCir(9).Text <> "" Then g_nShldToBreastGiven = fnDisplayToInches(Val(MainForm!txtCir(9).Text))
    If MainForm!txtCir(10).Text <> "" Then g_nUnderBreastCirGiven = fnDisplayToInches(Val(MainForm!txtCir(10).Text))
    If MainForm!txtCir(11).Text <> "" Then g_nNippleCirGiven = fnDisplayToInches(Val(MainForm!txtCir(11).Text))
    g_sLeftSleeve = MainForm!cboLeftAxilla.Text
    g_sRightSleeve = MainForm!cboRightAxilla.Text
    g_sFrontNeckStyle = MainForm!cboFrontNeck.Text
    g_sBackNeckStyle = MainForm!cboBackNeck.Text
    g_sClosure = MainForm!cboClosure.Text
    
    If (g_sFrontNeckStyle = "Measured Scoop" Or g_sFrontNeckStyle = "Measured V neck" Or InStr(g_sFrontNeckStyle, "Turtle") > 0) Then
                g_nFrontNeckSize = fnDisplayToInches(Val(MainForm!txtFrontNeck.Text))
    End If
    If (g_sBackNeckStyle = "Measured Scoop") Then
                g_nBackNeckSize = fnDisplayToInches(Val(MainForm!txtBackNeck.Text))
    End If
    g_sFabric = MainForm!cboFabric.Text
    If MainForm!optLeftLeg(0).Value Then
        g_sLeftLeg = "Panty"          'Options are: Panty & Brief
    ElseIf MainForm!optLeftLeg(1).Value Then
        g_sLeftLeg = "Brief"
    End If
    If MainForm!optRightLeg(0).Value Then
        g_sRightLeg = "Panty"          'Options are: Panty & Brief
    ElseIf MainForm!optRightLeg(1).Value Then
        g_sRightLeg = "Brief"
    End If

    g_sLegStyle = MainForm!cboLegStyle.Text
    g_sCrotchStyle = MainForm!cboCrotchStyle.Text

   'Reductions
    g_nNeckCirRed = 0.9                 '90%
    g_nShldToWaistRed = 0.95            '95%
    g_nShldToFoldRed = 0.95             '95%
    g_nShldToBreastRed = 0.95           '95%
    
    If MainForm!txtSex.Text = "Female" Then
        If MainForm!cboRed(2).Text <> "" Then
                g_nUnderBreastCirRed = Val(MainForm!cboRed(2).Text)
            Else
                g_nUnderBreastCirRed = 0.9  '90% default
        End If
    End If
    If MainForm!cboRed(1).Text <> "" Then
            g_nChestCirRed = Val(MainForm!cboRed(1).Text)
        Else
            g_nChestCirRed = 0.85 '85% default
    End If
    If MainForm!cboRed(0).Text <> "" Then
            g_nWaistCirRed = Val(MainForm!cboRed(0).Text)
        Else
            g_nWaistCirRed = 0.85 '85% default
    End If
    If MainForm!cboRed(3).Text <> "" Then
            g_nButtocksCirRed = Val(MainForm!cboRed(3).Text)
        Else
            g_nButtocksCirRed = 0.85 '85% default
    End If
    If MainForm!cboRed(4).Text <> "" Then
            g_nThighCirRed = Val(MainForm!cboRed(4).Text)
        Else
            g_nThighCirRed = 0.85   '85% default
    End If

End Sub

Static Function round(ByVal nNumber As Double) As Integer
   'Fuction to return the rounded value of a decimal number
   'E.G.
   '    round(1.35) = 1
   '    round(1.55) = 2
   '    round(2.50) = 3
   '    round(-2.50) = -3
   '
                      
    Dim nInt As Integer, nSign As Integer
    
    nSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    nInt = Int(nNumber)
    If (nNumber - nInt) >= 0.5 Then
    round = (nInt + 1) * nSign
    Else
    round = nInt * nSign
    End If
    
End Function


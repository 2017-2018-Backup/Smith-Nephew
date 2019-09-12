Option Strict Off
Option Explicit On
Module BODVALID
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

    Public Structure XY
        Public x As Double
        Public y As Double
    End Structure

    Public MainForm As armdia



    'Globals set by FN_Open
    'Public CC As Object 'Comma
    'Public QQ As Object 'Quote
    Public NL As Object 'Newline
    'Public fNum As Object 'Macro file number
    'Public QCQ As Object 'Quote Comma Quote
    'Public QC As Object 'Quote Comma
    'Public CQ As Object 'Comma Quote


    'MsgBox constant
    Const IDCANCEL As Short = 2
	Const IDYES As Short = 6
	Const IDNO As Short = 7
	
	Public g_bResponse As Object
	Public g_bDiffThigh As Short
	Public g_bMissCutOut7 As Short
	Public g_bExtremeCrotch As Short
	Public g_bDiffAxillaHeight As Short
	Public g_bSleeveless As Short
    Public g_bDrawSingleLeg As Short


    Public g_sSex As String
	Public g_sFileNo As String
	Public g_sSide As String
	Public g_sPatient As String
	Public g_sID As String
	Public g_sDiagnosis As String
	Public g_nAge As Short
	Public g_nAdult As Short
	Public g_sWorkOrder As String
	Public g_nUnitsFac As Double
	
	Public g_sBackNeck As String
	Public g_sFrontNeck As String
	
	Public g_nUnderBreast As Double
	Public g_nNipple As Double
	Public g_nChest As Double
	Public g_sLeftCup As String
	Public g_sRightCup As String
	
	Public g_sLeftDisk As String
	Public g_sRightDisk As String
	
	Public g_nLeftShldCirGiven As Double
	Public g_nRightShldCirGiven As Double
	Public g_nLeftAxillaCirGiven As Double
	Public g_nRightAxillaCirGiven As Double
	Public g_nChestCirGiven As Double
	Public g_nUnderBreastCirGiven As Double
	Public g_nWaistCirGiven As Double
	Public g_nButtocksCirGiven As Double
	Public g_nLeftThighCirGiven As Double
	Public g_nRightThighCirGiven As Double
	Public g_nNeckCirGiven As Double
	Public g_nShldToFoldGiven As Double
	Public g_nShldToWaistGiven As Double
	Public g_nShldToBreastGiven As Double
	Public g_nLeftShldToAxilla As Double
	Public g_nRightShldToAxilla As Double
	Public g_nNippleCirGiven As Double
	Public g_sFrontNeckStyle As String
	Public g_nFrontNeckSize As Double
	Public g_nRadiusFrontNeck As Double
	Public g_sBackNeckStyle As String
	Public g_nBackNeckSize As Double
	Public g_sClosure As String
	Public g_sFabric As String
	Public g_sCrotchStyle As String
	
	Public g_sSmallestThighGiven As String
	Public g_nLargestThighGiven As Double
	Public g_sLargestThighGiven As String
	
	Public g_nShldWidth As Double
	Public g_nLeftThighCir As Double
	Public g_nRightThighCir As Double
	Public g_sLeftSleeve As String
	Public g_sSleeveType As String
	Public g_sRightSleeve As String
	Public g_sLeftLeg As String
	Public g_sRightLeg As String
	Public g_nShldToAxilla As Double
	Public g_nAxillaCir As Double
	Public g_nThighCir As Double
	Public g_nGroinHeight As Double
	Public g_nLeftAxillaCir As Double
	Public g_nRightAxillaCir As Double
	Public g_nChestCir As Double
	Public g_nUnderBreastCir As Double
	Public g_nWaistCir As Double
	Public g_nButtocksCir As Double
	Public g_nNeckCir As Double
	Public g_nShldToFold As Double
	Public g_nShldToWaist As Double
	Public g_nShldToBreast As Double
	Public g_nButtocksLength As Double
	Public g_nButtBackSeamRatio As Double
	Public g_nButtFrontSeamRatio As Double
	Public g_nButtRedIncreased As Double
	Public g_nThighRedDecreased As Double
	Public g_nButtCir As Double
	Public g_nHalfGroinHeight As Double
	Public g_nButtRadius As Double
	Public g_nLT_GroinHeight As Double
	Public g_nLT_HalfGroinHeight As Double
	Public g_nLT_ButtRadius As Double
	Public g_n85FoldHeight As Double
	Public g_nButtCirCalc As Double
	Public g_nButtFrontSeam As Double
	Public g_nButtBackSeam As Double
	Public g_nButtCutOut As Double
	Public g_nWaistFrontSeam As Double
	Public g_nWaistBackSeam As Double
	Public g_nWaistCutOut As Double
	Public g_nChestFrontSeam As Double
	Public g_nChestBackSeam As Double
	Public g_nChestCutOut As Double
	Public g_nSmallestThighGiven As Double
	Public g_nChestCirRed As Double
	Public g_nUnderBreastCirRed As Double
	Public g_nWaistCirRed As Double
	Public g_nLargeButtCirRed As Double
	Public g_nThighCirRed As Double
	Public g_nNeckCirRed As Double
	Public g_nShldToFoldRed As Double
	Public g_nShldToWaistRed As Double
	Public g_nShldToBreastRed As Double
	Public g_nButtocksCirRed As Double
	Public g_nCutOutRadius As Double
	Public g_sFlySize As String
	Public g_sGussetSize As String
	Public g_nGussetLength As Double
	Public g_nFrontCrotchSize As Double
	Public g_nBackCrotchSize As Double
	Public g_bDrawBriefCurve As Short
	Public g_nCutOut As Double
	Public g_nCutOutToBackMinimum As Double
	Public g_sLegStyle As String
	Public g_nRightLegLength As Double
	Public g_nLeftLegLength As Double
	Public g_bRightAboveKnee As Short
	Public g_bLeftAboveKnee As Short
	
	Public g_sMeshLeft As String
	Public g_sMeshRight As String
	
	
	Public g_nAxillaBackNeckRad As Double
	Public g_nAxillaFrontNeckRad As Double
	Public g_nABNRadRight As Double
	Public g_nAFNRadRight As Double
	
	'Co-ordinate points
	Public xySeamFold As XY
	Public xySeamThigh As XY
	Public xySeamButt As XY
	Public xySeamHighShld As XY
	Public xySeamLowShld As XY
	Public xySeamChest As XY
	Public xySeamChestAxillaLow As XY
	
	Public g_sAxillaSide As String
	Public g_sAxillaSideLow As String
	Public g_sAxillaType As String
	Public g_sAxillaTypeLow As String
	Public g_nShldToAxillaLow As Double
	Public g_nAxillaCirLow As Double
	Public g_nLengthStrap1ToCutOut5 As Double
	
	
	Public xySeamWaist As XY
	Public xyProfileThigh As XY
	Public xyLT_ProfileThigh As XY
	Public xyProfileThighExtraPT As XY
	Public xyFold As XY
	Public xyLT_Fold As XY
	Public xyProfileButt As XY
	Public xyCutOut3 As XY
	Public xyCutOut4 As XY
	Public xyCutOut5 As XY
	Public xyCutOut6 As XY
	Public xyCutOut2 As XY
	Public xyCutOut9 As XY
	Public xyCutOut10 As XY
	Public xyCutOut7 As XY
	Public xyProfileWaist As XY
	Public xyProfileChest As XY
	Public xyProfileNeck As XY
	Public xyCutOutBackNeck As XY
	Public xyCutOut8 As XY
	Public xyCutOutFrontNeck As XY
	Public xyFrontNeckArcCentre As XY
	Public xyRaglan1 As XY
	Public xyRaglan2 As XY
	Public xyRaglan3 As XY
	Public xyRaglan4LowAxilla As XY
	Public xyRaglan2LowAxilla As XY
	Public xyRaglan4 As XY
	Public xyRaglan5 As XY
	Public xyRaglan6 As XY
	Public xyProfileNeckMirror As XY
	Public xyCutOutArcCentre As XY
	Public xyLegPoint As XY
	Public xy85Fold As XY
	Public xyProfileGroin As XY
	Public xyLT_ProfileGroin As XY
	
	Public xyProfileBrief As XY
	Public xyProfileThighMirror As XY
	Public xyLT_ProfileThighMirror As XY
	Public xyProfileThighMirror1 As XY
	Public xyStrap1 As XY
	Public xyStrap2 As XY
	Public xyStrap3 As XY
	Public xyStrap4 As XY
	Public xyTmpCutOut8 As XY
	Public xyGussetArcCentre As XY
	Public xyBackCrotchFilletCentre As XY
	Public xyFrontCrotchFilletCentre As XY
	Public xyBackCrotch1 As XY
	Public xyBackCrotch2 As XY
	Public xyBackCrotch3 As XY
	Public xyFrontCrotch1 As XY
	Public xyFrontCrotch2 As XY
	Public xyFrontCrotch3 As XY
	Public xyCrotchMarker As XY
	Public xyButtocksArcCentre As XY
	Public xyLT_ButtocksArcCentre As XY
	Public xyFrenchCut As XY
	'    Global xyLeftLegLowestPoint As XY
	'    Global xyRightLegLowestPoint As XY
	'    Global xyLargestLegLowestPoint As XY
	'    Global xySmallestLegLowestPoint As XY
	Public xyLeftLegLabelPoint As XY
	Public xyRightLegLabelPoint As XY
	Public xyBothLegLabelPoint As XY
	
	
	Public xyAxilla As XY
	
	
	Public Const INCH1_16 As Double = 0.0625
	Public Const INCH3_16 As Double = 0.1875
	Public Const INCH5_16 As Double = 0.3125
	Public Const INCH1_8 As Double = 0.125
	Public Const INCH1_4 As Double = 0.25
	Public Const INCH3_8 As Double = 0.375
	Public Const INCH1_2 As Double = 0.5
	Public Const INCH5_8 As Double = 0.625
	Public Const INCH3_4 As Double = 0.75
	Public Const INCH7_8 As Double = 0.875
	
	
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
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(0) = "Left shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(1) = "Right shoulder circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(2) = "Neck circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(3) = "Shoulder width."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(4) = "Shoulder to waist."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(5) = "Chest circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(6) = "Waist circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(9) = "Shoulder to under breast."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(10) = "Circ. under breast."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(11) = "Circ. over nipple."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(12) = "Shoulder to Fold of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(13) = "Shoulder to Large Part of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(14) = "Circ. of Large Part of Buttocks."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(15). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(15) = "Left Thigh Circ."
		'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(16). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCircum(16) = "Right Thigh Circ."
		
		'FATAL Errors
		'~~~~~~~~~~~~
		'Body measurements (all must be present)
		For ii = 0 To 6
			If Val(CType(MainForm.Controls("txtCir"), Object)(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing " & sCircum(ii) & NL
			End If
		Next ii
		For ii = 12 To 16
			If Val(CType(MainForm.Controls("txtCir"), Object)(ii).Text) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError & "Missing " & sCircum(ii) & NL
			End If
		Next ii
		
		'Bra Cups
		'Note:
		'    The Circumference over nipple is optional unless a cup has been
		'    specified
		'
		If Val(CType(MainForm.Controls("txtCir"), Object)(9).Text) = 0 And Val(CType(MainForm.Controls("txtCir"), Object)(10).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing " & sCircum(9) & NL
		End If
		If Val(CType(MainForm.Controls("txtCir"), Object)(10).Text) = 0 And Val(CType(MainForm.Controls("txtCir"), Object)(9).Text) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing " & sCircum(10) & NL
		End If
		
		If (Val(CType(MainForm.Controls("txtCir"), Object)(9).Text) <> 0 Or Val(CType(MainForm.Controls("txtCir"), Object)(10).Text) <> 0) And ((CType(MainForm.Controls("cboLeftCup"), Object).Text <> "None" And CType(MainForm.Controls("txtLeftDisk"), Object).Text = "") Or (CType(MainForm.Controls("txtRightDisk"), Object).Text = "" And CType(MainForm.Controls("cboRightCup"), Object).Text <> "None")) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & NL
		End If
		
		If Val(CType(MainForm.Controls("txtCir"), Object)(9).Text) = 0 And (CType(MainForm.Controls("txtLeftDisk"), Object).Text <> "" Or CType(MainForm.Controls("txtRightDisk"), Object).Text <> "") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sCircum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Bra disks given! But Missing " & sCircum(9) & NL
		End If
		
		'If cups or dimensions given then a disk must be present
		'NB
		'    cboXXXXCup.ListIndex = 6 = "None"
		'    cboXXXXCup.ListIndex = 7 = ""
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboLeftCup.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("cboLeftCup"), Object).ListIndex < 6 And Val(CType(MainForm.Controls("txtLeftDisk"), Object).Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Left BRA Cup!" & NL
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboRightCup.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("cboRightCup"), Object).ListIndex < 6 And Val(CType(MainForm.Controls("txtRightDisk"), Object).Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No disk calculated for Right BRA Cup!" & NL
		End If
		
		'Sex error
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboRightCup.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboLeftCup.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("txtSex"), Object).Text = "Male" And (Val(CType(MainForm.Controls("txtCir"), Object)(9).Text) <> 0 Or Val(CType(MainForm.Controls("txtCir"), Object)(10).Text) <> 0 Or CType(MainForm.Controls("cboLeftCup"), Object).ListIndex < 0 Or CType(MainForm.Controls("cboRightCup"), Object).ListIndex < 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & NL
		End If
		
		'Neck at back and front
		Dim sChar As New VB6.FixedLengthString(1)
		sChar.Value = Left(CType(MainForm.Controls("cboBackNeck"), Object).Text, 1)
		If sChar.Value = "M" And CType(MainForm.Controls("txtBackNeck"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Measured Back neck style!" & NL
		End If
		
		sChar.Value = Left(CType(MainForm.Controls("cboFrontNeck"), Object).Text, 1)
		If sChar.Value = "M" And CType(MainForm.Controls("txtFrontNeck"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "No dimension for Measured Front neck style" & NL
		End If
		
		'Get values from dialog
		g_nShldWidth = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(3).Text))
		g_sBackNeckStyle = CType(MainForm.Controls("cboBackNeck"), Object).Text
		g_sFrontNeckStyle = CType(MainForm.Controls("cboFrontNeck"), Object).Text
		'Test minimum value for Shoulder width
		If g_nShldWidth < 1.5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Shoulder Width is less than 1-1/2""!" & NL
		End If
		If InStr(g_sBackNeckStyle, "Scoop") > 0 And ((g_nShldWidth - 1) < 1.5) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With a scoop BACK neck the Shoulder Width is less than 1-1/2""!" & NL
		End If
		
		If (InStr(g_sFrontNeckStyle, "Scoop") > 0 Or InStr(g_sFrontNeckStyle, "V neck") > 0) And (g_nShldWidth - 1) < 1.5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With a " & g_sFrontNeckStyle & " FRONT neck the Shoulder Width is less than 1-1/2""!" & NL
		End If
		
		'
		If CType(MainForm.Controls("cboLeftAxilla"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Left Axilla not given!" & NL
		End If
		
		If CType(MainForm.Controls("cboRightAxilla"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Right Axilla not given!" & NL
		End If
		
		If CType(MainForm.Controls("cboFrontNeck"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Neck not given!" & NL
		End If
		
		If CType(MainForm.Controls("cboBackNeck"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Back neck not given!" & NL
		End If
		
		If CType(MainForm.Controls("cboClosure"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Closure not given!" & NL
		End If
		
		If CType(MainForm.Controls("cboFabric"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Fabric not given!" & NL
		End If
		
		'Extra Values
		If CType(MainForm.Controls("cboLegStyle"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Leg Style not given!" & NL
		End If
		If CType(MainForm.Controls("cboCrotchStyle"), Object).Text = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Crotch Style not given!" & NL
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
		g_bDiffAxillaHeight = False
		g_bDiffThigh = False
		g_bSleeveless = False
		
		'Check if Axilla values are within an inch of each other
		If System.Math.Abs(g_nLeftShldCirGiven - g_nRightShldCirGiven) > 1 Then 'Greater than 1" difference
			'keep separate
			g_bDiffAxillaHeight = True
		Else
			'Use average (sides same at axilla)
			g_nLeftShldCirGiven = ((g_nLeftShldCirGiven + g_nRightShldCirGiven) / 2)
			g_nRightShldCirGiven = g_nLeftShldCirGiven
		End If
		g_nLeftShldToAxilla = fnRoundInches((g_nLeftShldCirGiven * 0.95) / 3.14) + INCH1_2
		g_nRightShldToAxilla = fnRoundInches((g_nRightShldCirGiven * 0.95) / 3.14) + INCH1_2
		
		
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
		If System.Math.Abs(g_nLeftThighCirGiven - g_nRightThighCirGiven) > 1 Then g_bDiffThigh = True
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
		g_nChestCir = fnRoundInches(g_nChestCirGiven * g_nChestCirRed) / 2 'half scale
		g_nWaistCir = fnRoundInches(g_nWaistCirGiven * g_nWaistCirRed) / 2 'half-scale
		g_nShldToFold = fnRoundInches(g_nShldToFoldGiven * g_nShldToFoldRed) + INCH1_2
		g_nShldToWaist = fnRoundInches(g_nShldToWaistGiven * g_nShldToWaistRed) + INCH1_2
		g_nButtocksLength = fnRoundInches((g_nShldToFold - g_nShldToWaist) / 3)
		g_nUnderBreastCir = fnRoundInches(g_nUnderBreastCirGiven * g_nUnderBreastCirRed) / 2 'half scale
		g_nNeckCir = fnRoundInches(g_nNeckCirGiven * g_nNeckCirRed)
		g_nShldToBreast = fnRoundInches(g_nShldToBreastGiven * g_nShldToBreastRed) + INCH1_2
		
		
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
		g_nGroinHeight = fnRoundInches(g_nSmallestThighGiven * 0.85) / 2 'half scale
		g_nHalfGroinHeight = (g_nGroinHeight / 2)
		g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
		g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
		g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
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
			g_nGroinHeight = fnRoundInches(g_nSmallestThighGiven * g_nThighCirRed) / 2 'half scale
			g_nHalfGroinHeight = (g_nGroinHeight / 2)
			g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
			g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
			g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
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
			CType(MainForm.Controls("cboRed"), Object)(3).Text = CStr(g_nButtocksCirRed)
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With given Buttock Reduction, distance from back of Cut-Out" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "to largest part of buttocks is less than " & sCutOutToBackMinimum & " inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Buttocks Reduction increased to rectify this." + NL
			'sError = sError + "ref:- " + NL
		End If
		If nThighRedDecreased Then
			CType(MainForm.Controls("cboRed"), Object)(4).Text = CStr(g_nThighCirRed)
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "With given Thigh Reduction, distance from back of Cut-Out" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "to largest part of buttocks is less than " & sCutOutToBackMinimum & " inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Thigh Reduction decreased to rectify this." + NL
			'sError = sError + "ref:- " + NL
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
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "All Reductions should be within 5% of each other" + NL
		End If
		
		'MsgBox "4"
		
		'One last time with Reductions found above
		'
		'Calculate for Average / Smallest thigh
		g_nButtocksCir = fnRoundInches(g_nButtocksCirGiven * g_nButtocksCirRed) / 2 'half scale
		g_nGroinHeight = fnRoundInches(g_nSmallestThighGiven * g_nThighCirRed) / 2 'half scale
		g_nHalfGroinHeight = (g_nGroinHeight / 2)
		g_nButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nHalfGroinHeight ^ 2))
		g_nButtCirCalc = g_nButtRadius + g_nHalfGroinHeight
		g_nButtBackSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtBackSeamRatio
		g_nGroinHeight = g_nHalfGroinHeight + System.Math.Sqrt((g_nButtRadius ^ 2) - ((g_nButtocksLength - INCH3_4) ^ 2))
		
		'Calculate largest thigh (if given)
		g_nLT_GroinHeight = fnRoundInches(g_nLargestThighGiven * g_nThighCirRed) / 2 'half scale
		g_nLT_HalfGroinHeight = (g_nLT_GroinHeight / 2)
		g_nLT_ButtRadius = System.Math.Sqrt((g_nButtocksLength ^ 2) + (g_nLT_HalfGroinHeight ^ 2))
		g_nLT_GroinHeight = g_nLT_HalfGroinHeight + System.Math.Sqrt((g_nLT_ButtRadius ^ 2) - ((g_nButtocksLength - INCH3_4) ^ 2))
		
		If g_nButtBackSeam <= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "is still negative or zero!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Even with the modified reductions." + NL
			iFatalError = True
		ElseIf g_nButtBackSeam < g_nCutOutToBackMinimum Then  'less than required minimum
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Severe - Warning!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Distance from back of Cut-Out to largest part of buttocks" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "is still less than " & sCutOutToBackMinimum & " inches!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Even with the modified reductions." + NL
		End If
		
		'All other seams test
		'Set initila values
		g_nButtFrontSeam = (g_nButtocksCir - g_nButtCirCalc) * g_nButtFrontSeamRatio
		g_nButtCutOut = g_nButtCirCalc - (g_nButtBackSeam + g_nButtFrontSeam)
		
		nInchesDiff = Int(System.Math.Abs(g_nWaistCirGiven - g_nButtocksCirGiven))
		g_nWaistFrontSeam = (g_nButtFrontSeam * 2)
		If (g_nWaistCirGiven > g_nButtocksCirGiven) And nInchesDiff > 0 Then
			g_nWaistFrontSeam = g_nWaistFrontSeam + (nInchesDiff * INCH1_4)
		Else
			g_nWaistFrontSeam = g_nWaistFrontSeam - (nInchesDiff * INCH1_8)
		End If
		
		nInchesDiff = round(System.Math.Abs(g_nChestCirGiven - g_nButtocksCirGiven))
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
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError + NL + "Severe - Warning!" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to profile at WAIST" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "is negative or zero!" + NL
					If nCutOutModified > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sError = sError & "Even with the front of the Cut-Out at the Waist lowered." + NL
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
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Original distance from back of Cut-Out to profile at the WAIST" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "was less than " & sCutOutToBackMinimum & " inches!" + NL
					Select Case nCutOutModified
						Case 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Cut-Out at WAIST had to be lowered " & Str(nCutOutModified) & " time." & NL
							'sError = sError + "ref:- " + NL
						Case Is > 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Cut-Out at WAIST had to be lowered " & Str(nCutOutModified) & " times." & NL
							'sError = sError + "ref:- " + NL
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
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError + NL + "Severe - Warning!" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out to profile at the CHEST" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "is negative or zero!" + NL
					If nCutOutModified > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sError = sError & "Even with the Cut Out at the Chest lowered." + NL
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
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "Distance from back of Cut-Out at the CHEST" + NL
					'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sError = sError & "was less than " & sCutOutToBackMinimum & " inches!" + NL
					Select Case nCutOutModified
						Case 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str(nCutOutModified) & " time." & NL
							'sError = sError + "ref:- " + NL
						Case Is > 1
							'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sError = sError & "Front of the Cut Out at the CHEST had to be lowered " & Str(nCutOutModified) & " times." & NL
							'sError = sError + "ref:- " + NL
					End Select
				End If
				Exit Do 'greater than g_nCutOutToBackMinimum - acceptable
			End If
		Loop 
		g_nChestBackSeam = nChestRem / 2
		
		'Snap Crotch / Gusset warning wrt brief
		If InStr(g_sLegStyle, "Brief") > 0 And g_sCrotchStyle = "Gusset" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Information!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "A crotch style of ""Gusset"" has been selected for a BRIEF." + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "You should draw as a Snap Crotch" + NL
		End If
		
		If InStr(g_sLegStyle, "Brief") > 0 And g_sCrotchStyle = "Open Crotch" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Information!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "A crotch style of ""Open"" has been selected for a BRIEF." + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "You should not use an Open crotch style with a Brief" + NL
		End If
		If InStr(g_sCrotchStyle, "Hor") > 0 And g_nAge < 3 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError + NL + "Information!" + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "A Horizontal Fly has been selected for a child under 3." + NL
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "The horizontal fly chart does not contain an entry for children under 3 years old." + NL
		End If
		
		
		'Display Error message (if required) and return
		'There could be fatal errors found
		If Len(sError) > 0 And bDisplayErrors Then
			If iFatalError = True Then
				MsgBox(sError, 64, "Warning - Problems with data")
				FN_ValidateAndCalculateData = False
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sError = sError + NL + "The above problems have been found in the data do you" + NL
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
	
	Sub PR_GetValuesFromDialogue()
		
		
		'Get values from dialog box
		g_sFileNo = CType(MainForm.Controls("txtFileNo"), Object).Text
		g_sPatient = CType(MainForm.Controls("txtPatientName"), Object).Text
		g_sDiagnosis = CType(MainForm.Controls("txtDiagnosis"), Object).Text
		g_nAge = Val(CType(MainForm.Controls("txtAge"), Object).Text)
		
		'Set Adult status
		If g_nAge > 10 Then
			g_nAdult = True
		Else
			g_nAdult = False
		End If
		g_sSex = CType(MainForm.Controls("txtSex"), Object).Text
		g_sWorkOrder = CType(MainForm.Controls("txtWorkOrder"), Object).Text
		
		g_nLeftShldCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(0).Text))
		g_nRightShldCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(1).Text))
		g_nNeckCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(2).Text))
		g_nShldWidth = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(3).Text))
		g_nShldToWaistGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(4).Text))
		g_nChestCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(5).Text))
		g_nWaistCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(6).Text))
		g_nShldToFoldGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(12).Text))
		g_nButtocksCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(14).Text))
		g_nLeftThighCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(15).Text))
		g_nRightThighCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(16).Text))
		If CType(MainForm.Controls("txtCir"), Object)(9).Text <> "" Then g_nShldToBreastGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(9).Text))
		If CType(MainForm.Controls("txtCir"), Object)(10).Text <> "" Then g_nUnderBreastCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(10).Text))
		If CType(MainForm.Controls("txtCir"), Object)(11).Text <> "" Then g_nNippleCirGiven = fnDisplayToInches(Val(CType(MainForm.Controls("txtCir"), Object)(11).Text))
		g_sLeftSleeve = CType(MainForm.Controls("cboLeftAxilla"), Object).Text
		g_sRightSleeve = CType(MainForm.Controls("cboRightAxilla"), Object).Text
		g_sFrontNeckStyle = CType(MainForm.Controls("cboFrontNeck"), Object).Text
		g_sBackNeckStyle = CType(MainForm.Controls("cboBackNeck"), Object).Text
		g_sClosure = CType(MainForm.Controls("cboClosure"), Object).Text
		
		If (g_sFrontNeckStyle = "Measured Scoop" Or g_sFrontNeckStyle = "Measured V neck" Or InStr(g_sFrontNeckStyle, "Turtle") > 0) Then
			g_nFrontNeckSize = fnDisplayToInches(Val(CType(MainForm.Controls("txtFrontNeck"), Object).Text))
		End If
		If (g_sBackNeckStyle = "Measured Scoop") Then
			g_nBackNeckSize = fnDisplayToInches(Val(CType(MainForm.Controls("txtBackNeck"), Object).Text))
		End If
		g_sFabric = CType(MainForm.Controls("cboFabric"), Object).Text
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLeftLeg(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optLeftLeg"), Object)(0).Value Then
			g_sLeftLeg = "Panty" 'Options are: Panty & Brief
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLeftLeg(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CType(MainForm.Controls("optLeftLeg"), Object)(1).Value Then 
			g_sLeftLeg = "Brief"
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRightLeg(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("optRightLeg"), Object)(0).Value Then
			g_sRightLeg = "Panty" 'Options are: Panty & Brief
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRightLeg(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CType(MainForm.Controls("optRightLeg"), Object)(1).Value Then 
			g_sRightLeg = "Brief"
		End If
		
		g_sLegStyle = CType(MainForm.Controls("cboLegStyle"), Object).Text
		g_sCrotchStyle = CType(MainForm.Controls("cboCrotchStyle"), Object).Text
		
		'Reductions
		g_nNeckCirRed = 0.9 '90%
		g_nShldToWaistRed = 0.95 '95%
		g_nShldToFoldRed = 0.95 '95%
		g_nShldToBreastRed = 0.95 '95%
		
		If CType(MainForm.Controls("txtSex"), Object).Text = "Female" Then
			If CType(MainForm.Controls("cboRed"), Object)(2).Text <> "" Then
				g_nUnderBreastCirRed = Val(CType(MainForm.Controls("cboRed"), Object)(2).Text)
			Else
				g_nUnderBreastCirRed = 0.9 '90% default
			End If
		End If
		If CType(MainForm.Controls("cboRed"), Object)(1).Text <> "" Then
			g_nChestCirRed = Val(CType(MainForm.Controls("cboRed"), Object)(1).Text)
		Else
			g_nChestCirRed = 0.85 '85% default
		End If
		If CType(MainForm.Controls("cboRed"), Object)(0).Text <> "" Then
			g_nWaistCirRed = Val(CType(MainForm.Controls("cboRed"), Object)(0).Text)
		Else
			g_nWaistCirRed = 0.85 '85% default
		End If
		If CType(MainForm.Controls("cboRed"), Object)(3).Text <> "" Then
			g_nButtocksCirRed = Val(CType(MainForm.Controls("cboRed"), Object)(3).Text)
		Else
			g_nButtocksCirRed = 0.85 '85% default
		End If
		If CType(MainForm.Controls("cboRed"), Object)(4).Text <> "" Then
			g_nThighCirRed = Val(CType(MainForm.Controls("cboRed"), Object)(4).Text)
		Else
			g_nThighCirRed = 0.85 '85% default
		End If
		
	End Sub
	
	Function round(ByVal nNumber As Double) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35) = 1
		'    round(1.55) = 2
		'    round(2.50) = 3
		'    round(-2.50) = -3
		'
		
		Static nInt, nSign As Short
		
		nSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		nInt = Int(nNumber)
		If (nNumber - nInt) >= 0.5 Then
			round = (nInt + 1) * nSign
		Else
			round = nInt * nSign
		End If
		
	End Function
End Module
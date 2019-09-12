Option Strict Off
Option Explicit On
Module BODYSUIT1
    'Globals

    'Public MainForm As bodysuit '

    Public g_sPreviousLegStyle As String
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fNum As Object 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	'Store current layer and text setings to reduce DRAFIX code
	'this value is checked in PR_SetLayer
	Public g_sCurrentLayer As String
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	Public g_nCurrTextAngle As Object
	
	
	Public g_sChangeChecker As String
    Public g_sPathJOBST As String
    Public g_nUnitsFac As Double
    Public g_sFrontNeck As String
    Public g_sLeftDisk As String
    Public g_sRightDisk As String
    Public g_sLeftCup As String
    Public g_sRightCup As String

    Public g_nBodyShldWidth As Double
    Public g_sBodyBackNeckStyle As String
    Public g_sBodyFrontNeckStyle As String
    Public g_bBodyDiffAxillaHeight As Short
    Public g_bBodyDiffThigh As Short
    Public g_bBodySleeveless As Short
    Public g_nBodyLeftShldCirGiven As Double
    Public g_nBodyRightShldCirGiven As Double
    Public g_nBodyLeftAxillaCirGiven As Double
    Public g_nBodyRightAxillaCirGiven As Double

    Public Const BODY_INCH1_2 As Double = 0.5
    Public Const BODY_INCH1_8 As Double = 0.125
    Public Const BODY_INCH1_4 As Double = 0.25
    Public Const BODY_INCH3_4 As Double = 0.75

    Public g_nBodyLeftShldToAxilla As Double
    Public g_nBodyRightShldToAxilla As Double
    Public g_nBodyLeftAxillaCir As Double
    Public g_nBodyRightAxillaCir As Double
    Public g_nBodyLeftThighCirGiven As Double
    Public g_nBodyRightThighCirGiven As Double
    Public g_nBodySmallestThighGiven As Double
    Public g_nBodyLargestThighGiven As Double
    Public g_nBodyButtBackSeamRatio As Double
    Public g_nBodyButtFrontSeamRatio As Double

    Public g_sBodyLeftSleeve As String
    Public g_sBodyRightSleeve As String
    Public g_sBodySmallestThighGiven As String
    Public g_sBodyLargestThighGiven As String
    Public g_sBodyCrotchStyle As String


    'XY data type to represent points
    Structure XY
        Dim X As Double
        Dim Y As Double
    End Structure
    Public xyBDLeftLegLabelPoint As XY
    Public xyBDRightLegLabelPoint As XY
    'Co-ordinate points
    Public xyBodySeamFold As XY
    Public xyBodySeamThigh As XY
    Public xyBodySeamButt As XY
    Public xyBodySeamHighShld As XY
    Public xyBodySeamLowShld As XY
    Public xyBodySeamChest As XY
    Public xyBodySeamChestAxillaLow As XY
    Public xyBodySeamWaist As XY
    Public xyBodyProfileButt As XY
    Public xyBodyProfileWaist As XY
    Public xyBodyProfileChest As XY
    Public xyBodyProfileBrief As XY
    Public xyBodyProfileThigh As XY
    Public xyBDProfileThighExtraPT As XY

    Public xyBodyCutOut3 As XY
    Public xyBodyCutOut4 As XY
    Public xyBodyCutOut5 As XY
    Public xyBodyCutOut6 As XY
    Public xyBodyCutOut2 As XY
    Public xyBodyCutOut7 As XY
    Public xyBodyCutOut8 As XY
    Public xyBodyCutOut9 As XY
    Public xyBodyCutOut10 As XY
    Public xyBodyTmpCutOut8 As XY

    Public xyBDButtocksArcCentre As XY
    Public xyBodyProfileGroin As XY
    Public xyBodyLT_ButtocksArcCentre As XY
    Public xyBodyLT_ProfileGroin As XY
    Public xyBDFrontNeckArcCentre As XY
    Public xyBodyCutOutBackNeck As XY
    Public xyBodyProfileNeck As XY
    Public xyBodyCutOutFrontNeck As XY
    Public xyBDProfileNeckMirror As XY
    Public xyBodyCutOutArcCentre As XY

    Public xyBodyRaglan1 As XY
    Public xyBodyRaglan2 As XY
    Public xyBodyRaglan3 As XY
    Public xyBodyRaglan4 As XY
    Public xyBodyRaglan5 As XY
    Public xyBodyRaglan6 As XY
    Public xyBDRaglan4LowAxilla As XY
    Public xyBDRaglan2LowAxilla As XY

    Public xyBodyBackCrotch1 As XY
    Public xyBodyBackCrotch2 As XY
    Public xyBodyBackCrotch3 As XY
    Public xyBodyFrontCrotch1 As XY
    Public xyBodyFrontCrotch2 As XY
    Public xyBodyFrontCrotch3 As XY
    Public xyBDBackCrotchFilletCentre As XY
    Public xyBDFrontCrotchFilletCentre As XY
    Public xyBodyLegPoint As XY
    Public xyBodyCrotchMarker As XY

    Public xyBodyStrap1 As XY
    Public xyBodyStrap2 As XY
    Public xyBodyStrap3 As XY
    Public xyBodyStrap4 As XY
    Public xyBodyGussetArcCentre As XY
    Public xyBodyLT_Fold As XY
    Public xyBodyFold As XY
    Public xyBodyLT_ProfileThigh As XY
    Public xyBody85Fold As XY
    Public xyBDProfileThighMirror As XY
    Public xyBodyLT_ProfileThighMirror As XY
    Public xyBDProfileThighMirror1 As XY


    'PI as a Global Constant
    Public Const PI As Double = 3.141592654
	
	
	
	'Used to check if any modifications have been
	'done on the data
	Public g_Modified As Short
End Module
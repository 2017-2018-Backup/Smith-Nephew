Option Strict Off
Option Explicit On
Module BODYDRAW1
	
	Public MainForm As bodydia
	
	Public g_nFoldHt As Double
	Public g_nFoldHtRevised As Double
	
	Public g_sPreviousLegStyle As String
	Public sTask As String
	
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
	
	'UPGRADE_WARNING: Arrays in structure g_LeftLegProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public g_LeftLegProfile As curve
	'UPGRADE_WARNING: Arrays in structure g_RightLegProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public g_RightLegProfile As curve
	
	Public g_sChangeChecker As String
	Public g_sPathJOBST As String
	
	
	
	'PI as a Global Constant
	Public Const PI As Double = 3.141592654
	
	
	'Used to check if any modifications have been
	'done on the data
	Public g_Modified As Short
	
	'Global dimension variables
End Module
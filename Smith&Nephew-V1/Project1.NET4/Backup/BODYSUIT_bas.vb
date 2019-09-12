Option Strict Off
Option Explicit On
Module BODYSUIT1
	'Globals
	
	Public MainForm As bodysuit '
	
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
	
	
	'XY data type to represent points
	Structure XY
		Dim X As Double
		Dim Y As Double
	End Structure
	
	
	'PI as a Global Constant
	Public Const PI As Double = 3.141592654
	
	
	
	'Used to check if any modifications have been
	'done on the data
	Public g_Modified As Short
End Module
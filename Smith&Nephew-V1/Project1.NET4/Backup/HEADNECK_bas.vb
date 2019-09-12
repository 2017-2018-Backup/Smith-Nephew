Option Strict Off
Option Explicit On
Module HEADNECK1
	'XY data type to represent points
	Public MainForm As HeadNeck
	
	Structure xy
		Dim X As Double
		Dim Y As Double
	End Structure
	
	Structure curve
		<VBFixedArray(100)> Dim X() As Double
		<VBFixedArray(100)> Dim Y() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim X(100)
			'UPGRADE_WARNING: Lower bound of array Y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Y(100)
		End Sub
	End Structure
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	Public g_sWorkOrder As String
	Public g_sPathJOBST As String
	
	Public Const PI As Double = 3.141593
End Module
Option Strict Off
Option Explicit On
Module Module1
	'XY data type to represent points
	Structure XY
		Dim x As Double
		Dim y As Double
	End Structure
	
	Public Structure BiArc
		Dim xyStart As XY
		Dim xyTangent As XY
		Dim xyEnd As XY
		Dim xyR1 As XY
		Dim xyR2 As XY
		Dim nR1 As Double
		Dim nR2 As Double
	End Structure
End Module
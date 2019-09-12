Option Strict Off
Option Explicit On
Module JOBSTEX
	'Module :   JOBSTEX.BAS
	'Purpose:   Replicates JOBSTEX (HP 42S) Calculator programme
	'
	'Version:   3.01
	'Date:      Jan 1995
	'Author:    Gary George
	'
	'Used in:   WHFIGURE.MAK
	'           LGLEGDIA.MAK
	'           LGEDTDIA.BAS
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Jan 99     GG      Ported to VB5
	'
	'Notes:-
	'   A copy of the function FN_JOBSTEX_Stretch is used in
	'   LGEDTDIA.BAS
	'
	'
	
	
	'Constants used in FN_JOBSTEX_Stretch and FN_JOBSTEX_Pressure
	'Source:
	'   Mary Ann Hettich
	'   Memo to Marian Bourke, dated May 10 1994
	'
	Const A0 As Double = -159.6675
	Const A1 As Double = 7.300195
	Const A2 As Double = 1.021744
	Const A3 As Double = 200.9557
	Const A12 As Double = -0.04440547
	Const A13 As Double = 8.480964
	Const A23 As Double = -1.311007
	
	
	Function FN_JOBSTEX_Pressure(ByRef nAnkleCir As Double, ByRef nHeelCir As Double, ByRef nStretch As Double, ByRef iZipper As Short, ByRef iFabric As Short) As Short
		'Function to calculate stretch of the JOBSTEX material
		'Based on the HP-42S Calculator programs supplied by JOBST Toledo
		'
		'Main Program
		'    Fax from Mary Ann Hettich (Zipper Revisions)
		'    Dated:  May 10 1994
		'
		'Modifications
		'    Authorised Deviation T DEV 008
		'    Dated:  August 31 1994
		'
		'
		'Input
		'    nAnkleCir   Ankle circumference in inches
		'    nHeelCir    Heel circumference in inches
		'    nStretch    Requested stretch at ankle
		'    iZipper     Zipper, 0 => no zipper, 1 => zipper used
		'    iFabric     Jobstex fabric range
		'
		'Output
		'    mmHg        Calculated pressure of JOBSTEX fabric at given stretch.
		'                Integers only.
		'    -1000       Error, Non specific error
		'    -1001       Error, Unsupported fabric
		
		Dim nBit1 As Double 'nBit1 to 2 used to simplify calculations
		Dim nBit2 As Double
		Dim nA As Double
		Dim nB As Double
		Dim nE As Double
		Dim nHA As Double
		Dim nHA_1 As Double 'HA - 1
		Dim nMMHg As Double 'Pressure
		
		'Function is initally false
		FN_JOBSTEX_Pressure = -1000
		
		'Heel to ankle ratio
		'H *
		nHA = nHeelCir / nAnkleCir
		If nHA < 1 Then nHA = 1 'ie Ankle is larger than heel
		If iZipper = 1 Then nHA = 1
		
		nHA_1 = nHA - 1
		
		'F *
		Select Case iFabric
			Case 53
				nE = 0.9
			Case 55
				nE = 1
			Case 57
				nE = 1.1
			Case 63
				nE = 1.28
			Case 65
				nE = 1.4
			Case 67
				nE = 1.49
			Case 73
				nE = 1.64
			Case 75
				nE = 1.71
			Case 77
				nE = 1.84
			Case 83
				nE = 2.24
			Case 85
				nE = 2.31
			Case 87
				nE = 2.45
			Case Else
				FN_JOBSTEX_Pressure = -1001
				Exit Function
		End Select
		
		'A
		nA = A12 * nHA
		
		'B
		nB = A1 + (100 * A12 * nHA_1) + (A13 * nE) + (A2 + (A23 * nE)) * nHA
		
		'Solve for stretch
		
		nBit1 = A3 * nE
		nBit2 = ((100 * (A2 + (A23 * nE))) * nHA_1)
		' nBit3 = ((1.21 * nMMHg * nAnkleCir) - A0)
		
		'Calculate Pressure
		nMMHg = (((((((((nStretch * 2 * nA) + nB) ^ 2) - (nB ^ 2)) / (nA * 4))) + nBit2) + nBit1) + A0) / (1.21 * nAnkleCir)
		
		FN_JOBSTEX_Pressure = ARMDIA1.round(nMMHg)
		
	End Function
	
	Function FN_JOBSTEX_Stretch(ByRef nAnkleCir As Double, ByRef nHeelCir As Double, ByRef nMMHg As Double, ByRef iZipper As Short, ByRef iFabric As Short) As Double
		'Function to calculate stretch of the JOBSTEX material
		'Based on the HP-42S Calculator programs supplied by JOBST Toledo
		'
		'Main Program
		'    Fax from Mary Ann Hettich (Zipper Revisions)
		'    Dated:  May 10 1994
		'
		'Modifications
		'    Authorised Deviation T DEV 008
		'    Dated:  August 31 1994
		'
		'
		'Input
		'    nAnkleCir   Ankle circumference in inches
		'    nHeelCir    Heel circumference in inches
		'    nMMHg       Requested pressure at ankle
		'    iZipper     Zipper, 0 => no zipper, 1 => zipper used
		'    iFabric     Jobstex fabric range
		'
		'Output
		'    Stretch     Calculated strech of JOBSTEX fabric always > 0
		'    -1000       Error, Non specific error
		'    -1001       Error, Unsupported fabric
		'    -1002       Error, Unable to calculate stretch
		
		Dim nBit1 As Double
		Dim nBit2 As Double
		Dim nBit3 As Double
		Dim nA As Double
		Dim nB As Double
		Dim nE As Double
		Dim nHA As Double
		Dim nHA_1 As Double 'HA - 1
		
		'Function is initally false
		FN_JOBSTEX_Stretch = -1000
		
		'Heel to ankle ratio
		'H *
		nHA = nHeelCir / nAnkleCir
		If nHA < 1 Then nHA = 1 'ie Ankle is larger than heel
		If iZipper = 1 Then nHA = 1
		
		nHA_1 = nHA - 1
		
		'F *
		Select Case iFabric
			Case 53
				nE = 0.9
			Case 55
				nE = 1
			Case 57
				nE = 1.1
			Case 63
				nE = 1.28
			Case 65
				nE = 1.4
			Case 67
				nE = 1.49
			Case 73
				nE = 1.64
			Case 75
				nE = 1.71
			Case 77
				nE = 1.84
			Case 83
				nE = 2.24
			Case 85
				nE = 2.31
			Case 87
				nE = 2.45
			Case Else
				FN_JOBSTEX_Stretch = -1001
				Exit Function
		End Select
		
		'A
		nA = A12 * nHA
		
		'B
		nB = A1 + (100 * A12 * nHA_1) + (A13 * nE) + (A2 + (A23 * nE)) * nHA
		
		'Solve for stretch
		
		nBit1 = A3 * nE
		nBit2 = ((100 * (A2 + (A23 * nE))) * nHA_1)
		nBit3 = ((1.21 * nMMHg * nAnkleCir) - A0)
		
		'Check that quadratic can be solved
		If ((nB ^ 2) + (nA * 4 * ((nBit3 - nBit1) - nBit2))) < 0 Then
			FN_JOBSTEX_Stretch = -1002
			Exit Function
		End If
		
		'Calculate stretch
		FN_JOBSTEX_Stretch = (-nB + System.Math.Sqrt((nB ^ 2) + (nA * 4 * ((nBit3 - nBit1) - nBit2)))) / (2 * nA)
		
		
	End Function
End Module
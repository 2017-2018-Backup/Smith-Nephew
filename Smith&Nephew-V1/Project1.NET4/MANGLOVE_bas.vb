Option Strict Off
Option Explicit On
Module MANGLOVE1
    'Project:   MANGLOVE.BAS
    'Purpose:
    '
    '
    'Version:   1.01
    'Date:      Feb.96
    'Author:    Gary George
    '
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    '
    'Notes:-
    '

    Structure XY
        Dim X As Double
        Dim y As Double
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
    Structure XY
        Dim X As Double
        Dim y As Double
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

    'PI as a Global Constant
    Public Const PI As Double = 3.141592654
	
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
	
	
	Public g_sFileNo As String
	Public g_sSide As String
	Public g_sPatient As String
	Public g_sID As String
	
	Public g_nUnitsFac As Double
	
	'Scale to match current CAD-GLOVE output
	'Global Const ISCALE = 1026
	
	Public g_sPathJOBST As String
	
	Public MainForm As manglove
	
	'UPGRADE_WARNING: Lower bound of array g_iInsertValue was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_iInsertValue(6) As Double
	'UPGRADE_WARNING: Lower bound of array g_Finger was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_Finger(5) As Finger 'Includes thumb
	'UPGRADE_WARNING: Lower bound of array g_nFingerPIPCir was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nFingerPIPCir(5) As Double
	'UPGRADE_WARNING: Lower bound of array g_nFingerDIPCir was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nFingerDIPCir(5) As Double
	'UPGRADE_WARNING: Lower bound of array g_nFingerLen was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nFingerLen(5) As Double
	'UPGRADE_WARNING: Lower bound of array g_nWeb was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nWeb(4) As Double
	Public g_OnFold As Short
	Public g_MissingFingers As Short
	'UPGRADE_WARNING: Lower bound of array g_Missing was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_Missing(5) As Short
	Public g_nPalm As Double
	Public g_nWrist As Double
	'UPGRADE_WARNING: Lower bound of array g_iFINGER_CHART was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_iFINGER_CHART(5) As Short
	'UPGRADE_WARNING: Lower bound of array g_nFINGER_FIGURE was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nFINGER_FIGURE(5) As Double
	Public g_nPALM_FIGURE As Double
	Public g_nWRIST_FIGURE As Double
	Public g_iThumbStyle As Short
	Public g_iInsertStyle As Short
	Public g_iInsertSize As Short
	Public g_CalculatedExtension As Short
	Public g_FusedFingers As Short
	Public g_iThumbWebDrop As Short
	
	Public xyDatum As XY
	Public xyLittle As XY
	Public xyLFS As XY
	Public xyRing As XY
	Public xyIndex As XY
	Public xyMiddle As XY
	Public xyThumb As XY
	'UPGRADE_WARNING: Lower bound of array xyPalm was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public xyPalm(6) As XY
	'UPGRADE_WARNING: Lower bound of array xyPalmThumb was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public xyPalmThumb(5) As XY
	
	
	'UPGRADE_WARNING: Lower bound of array g_iGms was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_iGms(NOFF_ARMTAPES) As Short
	'UPGRADE_WARNING: Lower bound of array g_iMMs was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_iMMs(NOFF_ARMTAPES) As Short
	'UPGRADE_WARNING: Lower bound of array g_iRed was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_iRed(NOFF_ARMTAPES) As Short
	'UPGRADE_WARNING: Lower bound of array g_nCir was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nCir(NOFF_ARMTAPES) As Double
	'UPGRADE_WARNING: Lower bound of array g_nPleats was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nPleats(4) As Double
	Public g_iFirstTape As Short
	Public g_iLastTape As Short
	Public g_iWristPointer As Short
	Public g_iEOSPointer As Short
	Public g_iNumTotalTapes As Short
	Public g_iNumTapesWristToEOS As Short
	Public g_EOSType As Short
	Public g_iPressure As Short
	Public g_iInsertOtherGlv As Short
	Public g_OnFoldOtherGlv As Short
	Public g_DataIsCalcuable As Short
	Public g_PrintFold As Short
	Public g_ExtendTo As Short
	
	'Flaps
	Public g_nStrapLength As Double
	Public g_nFrontStrapLength As Double
	Public g_nCustFlapLength As Double
	Public g_nWaistCir As Double
	Public g_sFlapType As String
	Public g_iFlapType As Short
	
	'Curves
	'UPGRADE_WARNING: Arrays in structure g_FingerThumbBlend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public g_FingerThumbBlend As Curve
	Public g_ThumbTopCurve As BiArc
	'UPGRADE_WARNING: Arrays in structure UlnarProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public UlnarProfile As Curve
	'UPGRADE_WARNING: Arrays in structure RadialProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public RadialProfile As Curve
	'UPGRADE_WARNING: Lower bound of array TapeNote was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public TapeNote(NOFF_ARMTAPES) As TapeData
	
	
	Public Const g_sDialogID As String = "MANUAL Glove Dialogue"
	
	
	' MsgBox return values
	' Global Const IDOK = 1                  ' OK button pressed
	'Global Const IDCANCEL = 2              ' Cancel button pressed
	'Global Const IDABORT = 3               ' Abort button pressed
	'Global Const IDRETRY = 4               ' Retry button pressed
	'Global Const IDIGNORE = 5              ' Ignore button pressed
	'Global Const IDYES = 6                 ' Yes button pressed
	'Global Const IDNO = 7                  ' No button pressed
	
	
	Function FN_InchesValue(ByRef TextBox As System.Windows.Forms.Control) As Double
		' Check for numeric values
		Dim sChar, sText As String
		Dim iLen, nn As Short
		Dim nLen As Double
		
		sText = TextBox.Text
		iLen = Len(sText)
		
		'Check the actual structure of the input
		FN_InchesValue = -1
		For nn = 1 To iLen
			sChar = Mid(sText, nn, 1)
			If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
				MsgBox("Invalid - Dimension has been entered", 48, g_sDialogID)
				If MainForm.Visible Then TextBox.Focus()
				FN_InchesValue = -1
				Exit Function
			End If
		Next nn
		
		'Convert to inches
		nLen = fnDisplayToInches(Val(TextBox.Text))
		If nLen = -1 Then
			MsgBox("Invalid - Length has been entered", 48, g_sDialogID)
			If MainForm.Visible Then TextBox.Focus()
			FN_InchesValue = -1
		Else
			FN_InchesValue = nLen
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
		'        Double, Inches rounded to the nearest eighth (0.125)
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
	
	Function fnInchestoText(ByRef nInches As Double) As String
		'Function returns a decimal value in inches as a string
		'
		Dim nPrecision, nDec As Double
		Dim iInt, iEighths As Short
		Dim sString As String
		nPrecision = 0.125
		
		'Split into decimal parts
		iInt = Int(nInches)
		nDec = nInches - iInt
		If nDec <> 0 Then 'Avoid overflow
			iEighths = Int(nDec / nPrecision)
		Else
			iEighths = 0
		End If
		
		'Format string
		If iInt <> 0 Then
			sString = LTrim(Str(iInt))
		Else
			sString = "  "
		End If
		If iEighths <> 0 Then
			Select Case iEighths
				Case 2, 6
					sString = sString & "-" & LTrim(Str(iEighths / 2)) & "/4"
				Case 4
					sString = sString & "-" & "1/2"
				Case Else
					sString = sString & "-" & LTrim(Str(iEighths)) & "/8"
			End Select
		Else
			sString = sString & "   "
		End If
		
		'Return formatted string
		fnInchestoText = sString
		
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
		nDec = ARMDIA1.round(nDec)
		
		'Return value
		fnRoundInches = (iInt + (nDec * nPrecision)) * iSign
		
	End Function

    Sub PR_Select_Text(ByRef Text_Box_Name As System.Windows.Forms.TextBox)
        If Not Text_Box_Name.Enabled Then Exit Sub
        Text_Box_Name.Focus()
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
    End Sub
End Module
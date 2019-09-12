Option Strict Off
Option Explicit On
Module WHBODDIA1
	'   '* Windows API Functions Declarations
	'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	
	'   'Constants used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2'
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	'Globals
	Public g_nUnitsFac As Double
	
	Public MainForm As whboddia
	
	'Globals to hold original values of text boxes
	'to enable restoration of these values when leg style is
	'modified
	Public g_nTOSCir As Double
	Public g_nTOSHt As Double
	Public g_nTOSHtRevised As Double
	Public g_iTOSRed As Short
	
	Public g_nWaistCir As Double
	Public g_nWaistHt As Double
	Public g_nWaistHtRevised As Double
	Public g_iWaistRed As Short
	
	Public g_nMidPointCir As Double
	Public g_nMidPointHt As Double
	Public g_iMidPointRed As Short
	
	Public g_nLargestCir As Double
	Public g_nLargestHt As Double
	Public g_iLargestRed As Short
	
	Public g_nLeftThighCir As Double
	Public g_nRightThighCir As Double
	Public g_iThighRed As Short
	
	Public g_nFoldHt As Double
	Public g_nFoldHtRevised As Double
	
	Public g_sPreviousLegStyle As String
	
	'XY data type to represent points
	Structure XY
		Dim X As Double
		Dim Y As Double
	End Structure
	
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
	
	Public g_sChangeChecker As String
	Public g_sPathJOBST As String
	
	Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Fuction to return the length between two points
		'Greatfull thanks to Pythagorus
		
		FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)
		
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
	
	Function fnGetNumber(ByVal sString As String, ByRef iIndex As Short) As Double
		'Function to return as a numerical value the iIndexth item in a string
		'that uses blanks (spaces) as delimiters.
		'EG
		'    sString = "12.3 65.1 45"
		'    fnGetNumber( sString, 2) = 65.1
		'
		'If the iIndexth item is not found then return -1 to indicate an error.
		'This assumes that the string will not be used to store -ve numbers.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
		If Len(sString) = 0 Then
			fnGetNumber = -1
			Exit Function
		End If
		
		'Prepare string
		sString = sString & " " 'Trailing blank as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iIndex
			iPos = InStr(sString, " ")
			If ii = iIndex Then
				sString = Left(sString, iPos - 1)
				fnGetNumber = Val(sString)
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos))
				If Len(sString) = 0 Then
					fnGetNumber = -1
					Exit Function
				End If
			End If
		Next ii
		
		'The function should have exited befor this, however just in case
		'(iIndex = 0) we indicate an error,
		fnGetNumber = -1
		
	End Function
	
	Function fnInchesToText(ByRef nInches As Double) As String
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
		fnInchesToText = sString
		
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
	
	Function min(ByRef nFirst As Object, ByRef nSecond As Object) As Object
		' Returns the minimum of two numbers
		'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFirst <= nSecond Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			min = nFirst
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			min = nSecond
		End If
	End Function
	
	Function round(ByVal nNumber As Double) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35)  = 1
		'    round(1.55)  = 2
		'    round(2.50)  = 3
		'    round(-2.50) = -3
		'    round(0)     = 0
		'
		
		Dim iInt, iSign As Short
		
		'Avoid extra work. Return 0 if input is 0
		If nNumber = 0 Then
			round = 0
			Exit Function
		End If
		
		'Split input
		iSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		iInt = Int(nNumber)
		
		'Effect rounding
		If (nNumber - iInt) >= 0.5 Then
			round = (iInt + 1) * iSign
		Else
			round = iInt * iSign
		End If
		
	End Function
	
	Sub Select_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelStart = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
	End Sub
	
	Sub Validate_Text(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'Subroutine that is activated when the focus is lost.
		Dim rTextBoxValue As Double
		Dim ii, iTest As Short
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))
		
		'Only if relevant units are Inches
		If rTextBoxValue < 0 Then
			MsgBox("Invalid Format for inches", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		
		'Check that each character is numeric or a decimal point
		'N.B.
		'    Asc("0") = 48
		'    Asc("9") = 57
		'    Asc(".") = 46
		'
		For ii = 1 To Len(Text_Box_Name)
			iTest = Asc(Mid(Text_Box_Name.Text, ii, 1))
			If (iTest < 48 Or iTest > 57) And iTest <> 46 Then
				rTextBoxValue = -1
			End If
		Next ii
		
		If rTextBoxValue < 0 Then
			MsgBox("Invalid or Negative value given", 48, "Data input Error")
			Text_Box_Name.Focus()
		End If
		
		If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
			MsgBox("Zero value given", 48, "Data input Error")
			Text_Box_Name.Focus()
		End If
		
		If rTextBoxValue > 999.9 Then
			MsgBox("Given value too Large", 48, "Data input Error")
			Text_Box_Name.Focus()
		End If
		
	End Sub
	
	Sub Validate_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control, ByRef Label_Name As System.Windows.Forms.Control)
		'Subroutine that is activated when the focus is lost.
		Dim rTextBoxValue, rDec As Double
		Dim ii, iTest As Short
		Dim iEighths, iInt As Short
		Dim sString As String
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))
		
		'Only relevant if units are Inches
		If rTextBoxValue < 0 Then
			MsgBox("Invalid Format for inches", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		'Check that each character is numeric or a decimal point
		'N.B.
		'    Asc("0") = 48
		'    Asc("9") = 57
		'    Asc(".") = 46
		'
		For ii = 1 To Len(Text_Box_Name)
			iTest = Asc(Mid(Text_Box_Name.Text, ii, 1))
			If (iTest < 48 Or iTest > 57) And iTest <> 46 Then
				rTextBoxValue = -1
			End If
		Next ii
		
		If rTextBoxValue < 0 Then
			MsgBox("Invalid or Negative value given", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
			MsgBox("Zero value given", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		If rTextBoxValue > 999.9 Then
			MsgBox("Given value too Large", 48, "Data input Error")
			Text_Box_Name.Focus()
			Exit Sub
		End If
		
		Label_Name.Text = fnInchesToText(rTextBoxValue)
		
	End Sub
End Module
Option Strict Off
Option Explicit On
Module LGLEGDIA1
	'   '* Windows API Functions Declarations
	'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	'
	'   'Constants used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	Public MainForm As lglegdia
	
	'Globals
	Public g_nUnitsFac As Double
	Public g_iLtAnkle As Short
	Public g_JOBSTEX As Short
	Public g_JOBSTEX_FL As Short
    Public g_POWERNET As Short

    'Figuring globals
    'These are used by the sub PR_FigureAnkle to check
    'if re-figuring should occur because the current data
    'does not match the previous data
    Public g_nLtLastHeel As Double
	Public g_nLtLastAnkle As Double
	Public g_iLtLastMM As Double
	Public g_iLtLastStretch As Short
	Public g_iLtLastZipper As Short
	
	Public g_iLtStretch(29) As Short
	Public g_iLtRed(29) As Short
    Public g_iLtMM(29) As Short

    Public g_iStyleFirstTape As Short
	Public g_iStyleLastTape As Short
	Public g_iFirstTape As Short
    Public g_iLastTape As Short

    Public g_iLtLastFabric As Short
	Public g_sTextList As String
	
	Public g_iLastLegStyle As Short
	Public g_iLegStyle As Short
	
	Public g_sStyleString As String
	
	'Constants
	Const HEEL_TOL As Short = 9 '9" heel

    'POWERNET Modulus chart
    'The format was created to allow DRAFIX to index into a string, using it as an array
    'This format is retained here for compatibility.
    Structure Fabric
        <VBFixedArray(18)> Dim Modulus() As String
        'UPGRADE_NOTE: Conversion was upgraded to Conversion_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        <VBFixedArray(18)> Dim Conversion_Renamed() As String

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            ReDim Modulus(18)
        End Sub
    End Structure
    'XY data type to represent points
    Structure XY
        Dim X As Double
        Dim y As Double
    End Structure

    'UPGRADE_WARNING: Arrays in structure POWERNET may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
    Public POWERNET As Fabric
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fnum As Short 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	'Constants to define the drive and the root directory
	'    Const DRIVE = "C:"
	'    Const ROOTDIR = "\JOBST\"
	
	
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	
	Public g_sFileNo As String
	Public g_sSide As String
	Public g_sPatient As String
	Public g_sCurrentLayer As String
	
	'    Global g_sCatNo(1 To 10) As String
	'    Global g_nOptions As Integer
	'    Global g_sBestCatNo As String
	
	
	Public g_sChangeChecker As String
    Public g_sPathJOBST As String

    Public Const PI As Double = 3.141592654
    Public xyLegInsertion As XY


    Function fnDisplaytoInches(ByVal nDisplay As Double) As Double
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
			fnDisplaytoInches = fnRoundInches(nDisplay * g_nUnitsFac) * iSign
			Exit Function
		End If
		
		'Imperial units
		iInt = Int(nDisplay)
		nDec = nDisplay - iInt
		'Check that conversion is possible (return -1 if not)
		If nDec > 0.8 Then
			fnDisplaytoInches = -1
		Else
			fnDisplaytoInches = fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
		End If
		
	End Function
	
	
	Function fnInchesToDisplay(ByRef nInches As Double) As Double
		'This function takes the value given in inches and converts it
		'into the display format.
		'
		'Input:-
		'        nDisplay is the value in inches
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
		'        Single, Value rounded to the nearest tenth 0.1
		'
		'        The convention is that, Metric dimensions use the decimal
		'        point to indicate the division between CMs and mms
		'        ie 7.6 = 7 cm and 6 mm.
		'        Whereas the decimal point for imperial measurements indicates
		'        the division between inches and eighths
		'        ie 7.6 = 7 inches and 6 eighths
		'
		'Errors:-
		'
		
		Dim iInt, iSign As Short
		Dim nDec, nValue As Double
		
		'retain sign
		iSign = System.Math.Sign(nInches)
		nInches = System.Math.Abs(nInches)
		
		If g_nUnitsFac <> 1 Then
			'Units are cms
			nValue = nInches * 2.54 'Convert to cm
			iInt = Int(nValue)
			nDec = (nValue - iInt) * 10 'Shift decimal place right
			If nDec > 0 Then nDec = round(nDec) / 10 'Round and shift decimal place left
		Else
			'Units are inches
			nValue = fnRoundInches(nInches) 'Round to nearest 1/8 th
			iInt = Int(nValue)
			nDec = nValue - iInt
			If nDec > 0 Then
				nDec = (nDec / 0.125) * 10 'get 1/8ths and shift decimal place right
				nDec = round(nDec) / 10 'Round and shift decimal place left
			End If
		End If
		
		fnInchesToDisplay = (iInt + nDec) * iSign
		
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

    Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.ComboBox, ByRef sFileName As String)
        'General procedure to create the list section of
        'a combo box reading the data from a file

        Dim sLine As String
        Dim fFileNum As Short

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48)
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, OpenMode.Input)
        Do While Not EOF(fFileNum)
            sLine = LineInput(fFileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Combo_Name.AddItem(sLine)
            Combo_Name.Items.Add(sLine)
        Loop
        FileClose(fFileNum)

    End Sub

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

    Sub Select_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.TextBox)
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
    End Sub

    Sub Validate_And_Display_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control, ByRef sLabelVal As String)
        'Subroutine that is activated when the focus is lost.
        'Displays value in inches in the grid given
        '
        Dim rTextBoxValue As Double
        Dim iFirstTape, iLastTape As Short
        Dim ii, iTest As Short

        'Checks that input data is valid.
        'If not valid then display a message and returns focus
        'to the text in question

        'Get the text value
        rTextBoxValue = fnDisplaytoInches(Val(Text_Box_Name.Text))

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
        For ii = 1 To Len(Text_Box_Name.Text)
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
        sLabelVal = fnInchesToText(rTextBoxValue)

        'If Index >= 0 Then
        '    'Treat Display as Grid control
        '    'UPGRADE_WARNING: Couldn't resolve default property of object Display_Control.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    '----------Display_Control.Row = Index
        '    Display_Control.Text = fnInchesToText(rTextBoxValue)
        'Else
        '    'Treat Display as Label
        '    Display_Control.Text = fnInchesToText(rTextBoxValue)
        'End If


    End Sub

    Sub Validate_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'Subroutine that is activated when the focus is lost.
		Dim rTextBoxValue As Double
		Dim ii, iTest As Short
		
		'Checks that input data is valid.
		'If not valid then display a message and returns focus
		'to the text in question
		
		'Get the text value
		rTextBoxValue = fnDisplaytoInches(Val(Text_Box_Name.Text))
		
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
		
	End Sub
End Module
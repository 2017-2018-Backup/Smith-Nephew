Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
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

        Public Sub Initialize()
            ReDim X(100)
            ReDim y(100)
        End Sub
    End Structure
    Public Structure Finger
        Dim DIP As Double
        Dim PIP As Double
        Dim Len_Renamed As Double
        Dim WebHt As Double
        Dim Tip As Short
        Dim xyUlnar As XY
        Dim xyRadial As XY
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
    Public Structure TapeData
        Dim nCir As Double
        Dim iMMs As Short
        Dim iRed As Short
        Dim iGms As Short
        Dim sNote As String
        Dim iTapePos As Short
        Dim sTapeText As String
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

    Public g_iInsertValue(6) As Double
    Public g_Finger(5) As Finger 'Includes thumb
    Public g_nFingerPIPCir(5) As Double
    Public g_nFingerDIPCir(5) As Double
    Public g_nFingerLen(5) As Double
    Public g_nWeb(4) As Double
    Public g_OnFold As Short
	Public g_MissingFingers As Short
    Public g_Missing(5) As Short
    Public g_nPalm As Double
	Public g_nWrist As Double
    Public g_iFINGER_CHART(5) As Short
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
    Public xyPalm(6) As XY
    Public xyPalmThumb(5) As XY
    Public xyT(10) As XY

    Public g_iGms(16) As Short
    Public g_iMMs(16) As Short
    Public g_iRed(16) As Short
    ''--------Public g_nCir(16) As Double
    Public g_nCir(1, NOFF_ARMTAPES) As Double
    Public g_nPleats(4) As Double
    Public g_iMGlvFirstTape As Short
    ''---------------Public g_iLastTape As Short
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
    ''--------Public g_FingerThumbBlend As Curve
    Public g_ThumbTopCurve As BiArc
    ''--------Public UlnarProfile As Curve
    ''---------Public RadialProfile As curve
    Public TapeNote(NOFF_ARMTAPES) As TapeData


    Public Const g_sDialogID As String = "Glove Dialogue"
    'Misc
    Public Const EIGHTH As Double = 0.125
    Public Const SIXTEENTH As Double = 0.0625
    Public Const QUARTER As Double = 0.25
    'Tapes etc
    Public Const PALM As Short = 2
    Public Const WRIST As Short = 1
    Public Const TAPE_ONE_HALF As Short = 3
    Public Const TAPE_THREE As Short = 4
    Public Const TAPE_FOUR_HALF As Short = 5
    Public Const TAPE_SIX As Short = 6
    Public Const TAPE_SEVEN_HALF As Short = 7
    Public Const TAPE_NINE As Short = 8

    Public Const LOW_MODULUS As Short = 160
    Public Const HIGH_MODULUS As Short = 340
    Public Const NOFF_MODULUS As Short = 19

    Public Const NOFF_ARMTAPES As Short = 16
    Public Const ELBOW_TAPE As Short = 9
    Public Const ARM_PLAIN As Short = 0
    Public Const ARM_FLAP As Short = 1
    Public Const GLOVE_NORMAL As Short = 0
    Public Const GLOVE_ELBOW As Short = 1
    Public Const GLOVE_AXILLA As Short = 2

    Public g_iModulusIndex As Short
    Dim g_iPowernet(NOFF_MODULUS, 23) As Short

    'Reduction Chart constants and Variables
    Const NOFF_FINGER_RED As Short = 32
    Const NOFF_ARM_RED As Short = 76
    Const NOFF_LENGTH_RED As Short = 24
    Const FIGURED_MINIMUM As Short = 6

    'Hand Reduction charts
    Dim iFingerRed(3, NOFF_FINGER_RED) As Short
    Dim nArmRed(8, NOFF_ARM_RED) As Double
    Dim nLengthRed(2, NOFF_LENGTH_RED) As Double

    'Fingers
    Public Const DIP As Short = 1
    Public Const PIP As Short = 2
    Public Const THUMB As Short = 3
    Public Const THUMB_LEN As Short = 2
    Public Const FINGER_LEN As Short = 1

    Public Const HORIZ_CENTER As Short = 2
    Public Const TOP_ As Short = 8
    Public Const BOTTOM_ As Short = 32
    Public Const LEFT_ As Short = 1
    Public Const RIGHT_ As Short = 4
    Public Const VERT_CENTER As Short = 16
    Public Const CURRENT As Short = -1

    Public Const g_sTapeText As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
    Public xyMGlvInsertion As XY
    Public g_idLastZipper As ObjectId

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
        Text_Box_Name.SelectionStart = 0
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
    End Sub
    Function FN_GetReduction(ByRef iGrams As Object) As Short
        'Function that looks up the "Powernet" grams / tension chart
        'and returns the reduction value
        '
        ' Input
        '        iGrams      Grams calculated from the data
        '
        ' Module Level variables
        '        g_iModulusIndex
        '        g_iPowernet()
        '
        If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
            FN_GetReduction = -1
            Exit Function
        End If

        Static iValue, ii, iPrevValue As Short

        Select Case iGrams
            Case Is <= g_iPowernet(g_iModulusIndex, 1)
                FN_GetReduction = 10

            Case Is >= g_iPowernet(g_iModulusIndex, 23)
                FN_GetReduction = 32

            Case Else
                'Return value closest
                iPrevValue = 0
                For ii = 1 To 23
                    iValue = g_iPowernet(g_iModulusIndex, ii)
                    If iValue > iGrams Then Exit For
                    iPrevValue = iValue
                Next ii

                If System.Math.Abs(iGrams - iPrevValue) < System.Math.Abs(iGrams - iValue) Then
                    FN_GetReduction = ii + 8
                Else
                    FN_GetReduction = ii + 9
                End If

        End Select
    End Function
    Function FN_GetGrams(ByVal iReduction As Short) As Short
        'Function that looks up the "Powernet" grams / tension chart
        'and returns the Grams for a given reduction value
        '
        'Used to back calculate from a given reduction
        '
        ' Input
        '        iReduction      Reduction given
        '
        ' Module Level variables
        '        g_iModulusIndex
        '        g_iPowernet()
        '
        If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
            FN_GetGrams = -1
            Exit Function
        End If

        If iReduction < 10 Then iReduction = 10
        If iReduction > 32 Then iReduction = 32

        FN_GetGrams = g_iPowernet(g_iModulusIndex, (iReduction - 10) + 1)
    End Function
    Sub PR_LoadReductionCharts(ByRef sFabric As String)
        'Procedure to load the reduction charts from
        'disk
        'Two charts are loaded
        '    1. Length chart
        '    2. Circumferences based on the Fabric
        '

        Static sFile, sLine As String
        Static jj, fChart, iModulus, ii, nn As Short

        If sFabric = "" Then Exit Sub
        If UCase(Mid(sFabric, 1, 3)) <> "POW" Then
            MsgBox("Fabric chosen is not Powernet", 48, g_sDialogID)
            Exit Sub
        End If

        'Establish chart to be loaded
        '    Fabric Format  Pow MMM-XX Comment
        '    if mm < 230 use 230 chart
        '    if MM > 280 use 280 chart
        '
        iModulus = Val(Mid(sFabric, 5, 3))
        If iModulus < 230 Then iModulus = 230
        If iModulus > 280 Then iModulus = 280
        ''-------sFile = g_sPathJOBST & "\TEMPLTS\GLV_" & Trim(Str(iModulus)) & ".DAT"
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        sFile = sSettingsPath & "\GLV_" & Trim(Str(iModulus)) & ".DAT"

        If Dir(sFile) = "" Then
            MsgBox("Fabric Chart, Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If

        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, VB.OpenMode.Input)

        'Get finger and thumb reducations
        'NB these are fixed format
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_FINGER_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB use of "." to repeat previous number
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
                iFingerRed(1, ii) = Val(Mid(sLine, 5, 2))

                If Mid(sLine, 8, 1) = "." Then
                    iFingerRed(2, ii) = iFingerRed(1, ii)
                Else
                    iFingerRed(2, ii) = Val(Mid(sLine, 8, 2))
                End If

                If Mid(sLine, 11, 1) = "." Then
                    iFingerRed(3, ii) = iFingerRed(2, ii)
                Else
                    iFingerRed(3, ii) = Val(Mid(sLine, 11, 2))
                End If
                ii = ii + 1
            End If
        End While

        'Wrist and Palm to 9 tape reductions
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_ARM_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB: use of "." to repeat previous number
            '    also translation from inches and eights to decimal inches)
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
                nArmRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
                jj = 1
                For nn = 9 To 34 Step 4
                    jj = jj + 1
                    If Mid(sLine, nn, 1) = "." Then
                        nArmRed(jj, ii) = nArmRed(jj - 1, ii)
                    Else
                        nArmRed(jj, ii) = Val(Mid(sLine, nn, 2)) + (Val(Mid(sLine, nn + 2, 1)) * 0.125)
                    End If
                Next nn
                ii = ii + 1
            End If
        End While

        'Close file
        FileClose(fChart)

        'Length Reduction chart
        ''---------sFile = g_sPathJOBST & "\TEMPLTS\GLV_LEN.DAT"
        sFile = sSettingsPath & "\GLV_LEN.DAT"
        If Dir(sFile) = "" Then
            MsgBox("Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If

        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, VB.OpenMode.Input)
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_LENGTH_RED
            sLine = LineInput(fChart)
            'Ignore comments and blank lines
            'NB: use of "." to repeat previous number
            '    also translation from inches and eights to decimal inches)
            If Mid(sLine, 1, 1) <> "#" And sLine <> "" Then
                nLengthRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)

                If Mid(sLine, 9, 1) = "." Then
                    nLengthRed(2, ii) = nLengthRed(1, ii)
                Else
                    nLengthRed(2, ii) = Val(Mid(sLine, 9, 2)) + (Val(Mid(sLine, 11, 1)) * 0.125)
                End If
                ii = ii + 1
            End If
        End While

        'Close file
        FileClose(fChart)
        sLine = ""
    End Sub
    Function FN_FigureFinger(ByRef iSelect As Short, ByRef nFingerCir As Double, ByRef iInsert As Short) As Double
        'procedure to return the Figured circumference of a finger
        '
        ' iSelect    a constant, one of DIP, PIP, THUMB
        ' nFingerCir Dimension in inches in
        '            1" >= nFingerCir <= 4.825"
        ' iInsert    Insert range to be used
        '            1 >= iNinsert <= 6
        '            iInsert = 1, use looked up figure
        '            iInsert > 1, add 2 to DIP & PIP, add 1 to THUMB
        '                         to looked up figure
        '            iInsert = 6, subtract 1 from DIP, PIP, THUMB
        '                         to looked up figure
        '

        Dim iIndex, iValue As Short

        'Check data is within range
        FN_FigureFinger = 0#
        If iSelect < DIP Or iSelect > THUMB Then Exit Function
        If iInsert < 1 Or iInsert > 6 Then Exit Function
        If nFingerCir < 1 Or nFingerCir > 4.875 Then Exit Function

        'Convert nFingerCir to an index
        'Note integer division \
        iIndex = (((nFingerCir - 1) * 1000) \ 125) + 1

        'Look up value
        iValue = iFingerRed(iSelect, iIndex)

        If iInsert = 6 Then
            iValue = iValue + 1
        ElseIf iSelect = THUMB Then
            iValue = iValue - (iInsert - 1)
        Else
            iValue = iValue - ((iInsert - 1) * 2)
        End If

        If iValue < FIGURED_MINIMUM Then iValue = FIGURED_MINIMUM

        'Convert from 16ths to inches
        FN_FigureFinger = iValue * 0.0625
    End Function
    Function FN_FigureTape(ByRef iSelect As Short, ByRef nTapeCir As Double) As Double
        'Procedure to return the Figured circumference of an ARM tape
        'This also is used to figure the Palm and wrist
        '
        ' iSelect    A constant, one off
        '            WRIST, PALM, TAPE_ONE_HALF ... TAPE_NINE
        ' nTapeCir   Dimension in inches in
        '            3.5" >= nTapeCir <= 12.875"

        'NOTES;
        '    The PALM and WRIST chart only goes up to 12.375
        '    if this value is exceeded then a value is extrapolated
        '    from the 12.875 value.
        '
        Dim iIndex, iValue As Short
        Dim iAdditional As Double

        'Check data is within range
        FN_FigureTape = 0#
        If iSelect < WRIST Or iSelect > TAPE_NINE Then Exit Function
        If nTapeCir < 3.5 Then Exit Function

        'Convert nTapeCir to an index
        'Note integer division \
        If nTapeCir > 12.875 Then
            'Extrapolate a value for this case
            'Get index of the last value in the table (ie 12.875)
            iIndex = (((12.875 - 3.5) * 1000) \ 125) + 1
            'Look up value
            'For each 1/2" greater than 12.875 add an eighth
            'to the returned value
            iAdditional = ((nTapeCir - 12.875) \ 0.5) + 1
            FN_FigureTape = nArmRed(iSelect, iIndex) + (iAdditional * 0.125)
        Else
            iIndex = (((nTapeCir - 3.5) * 1000) \ 125) + 1
            'Look up value
            FN_FigureTape = nArmRed(iSelect, iIndex)
        End If
    End Function
    Function FN_LengthWristToEOS() As Double
        'Calculates the distance from the EOS to the
        'wrist based on values given in the dialogue
        '
        Static ii As Short
        Static nLen, nSpace As Double

        'Get the data from the dialogue
        '    PR_GetDlgAboveWrist

        nLen = 0
        Select Case g_ExtendTo
            Case GLOVE_NORMAL
                ''----- "- 1" is added on 10-11-2018
                For ii = 1 To g_iNumTapesWristToEOS - 1
                    ''Changed for #215 in the issue list
                    ''nSpace = 1.375
                    nSpace = 1.2
                    nLen = nLen + nSpace
                Next ii
                If nLen = 0 Then
                    'If no tapes after wrist then default to 0.625
                    nLen = 0.625
                End If
            Case GLOVE_ELBOW
                For ii = 1 To g_iNumTapesWristToEOS - 1
                    nSpace = 1.375
                    If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
                    If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
                    If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
                    If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
                    nLen = nLen + nSpace
                Next ii

            Case GLOVE_AXILLA
                For ii = 1 To g_iNumTapesWristToEOS - 1
                    nSpace = 1.375
                    If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
                    If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
                    If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
                    If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
                    nLen = nLen + nSpace
                Next ii
                If g_EOSType = ARM_FLAP Then
                End If
        End Select

        FN_LengthWristToEOS = nLen

    End Function
    Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.y = y
    End Sub
    Function FN_FigureLength(ByRef iSelect As Short, ByRef nLen As Double) As Double
        'Procedure to return the Figured length of a finger
        'or thumb
        '
        ' iSelect    A constant, one of THUMB or FINGER
        ' nLen       Dimension in inches in
        '            1" >= nLen <= 3.875

        Dim iIndex, iValue As Short

        'Check data is within range
        FN_FigureLength = 0#
        If iSelect <> THUMB_LEN And iSelect <> FINGER_LEN Then Exit Function
        If nLen >= 1 And nLen <= 3.875 Then
            'Convert nLen to an index
            'Note integer division \
            iIndex = (((nLen - 1) * 1000) \ 125) + 1
            'Look up value
            FN_FigureLength = nLengthRed(iSelect, iIndex)
        ElseIf nLen < 1 Then
            'Return given value unmodified
            FN_FigureLength = nLen
        ElseIf nLen > 3.875 And iSelect = THUMB_LEN Then
            FN_FigureLength = nLen - 0.125
        ElseIf nLen > 3.875 And iSelect = FINGER_LEN Then
            FN_FigureLength = nLen - 0.25
        End If
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
    Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)

    End Function
    Function FN_CalcAngle(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, y As Object
        Dim rAngle As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = xyEnd.X - xyStart.X
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        y = xyEnd.y - xyStart.y

        'Horizomtal
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If y > 0 Then
                FN_CalcAngle = 90
            Else
                FN_CalcAngle = 270
            End If
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If y = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If X > 0 Then
                FN_CalcAngle = 0
            Else
                FN_CalcAngle = 180
            End If
            Exit Function
        End If

        'All other cases
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If
    End Function
    Sub PR_CalcPolar(ByRef xyStart As XY, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As XY)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double

        'Convert from degees to radians
        nAngle = nAngle * PI / 180

        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.X + A
        xyReturn.y = xyStart.y + B

    End Sub
    Function FN_CirLinInt(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyCen As XY, ByRef nRad As Double, ByRef xyInt As XY) As Short
        'Function to calculate the intersection between
        'a line and a circle.
        'Note:-
        '    Returns true if intersection found.
        '    The first intersection (only) is found.
        '    Ported from DRAFIX CAD DLG version.
        '

        Static nM, nC, nA, nSlope, nB, nK, nCalcTmp As Object
        Static nRoot As Double
        Static nSign As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nSlope = FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 0 Or nSlope = 180 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.y = xyStart.y
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Vertical Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 90 Or nSlope = 270 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.y + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
                    xyInt.y = nRoot
                    xyInt.X = xyStart.X
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Non-othogonal line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nM = (xyEnd.y - xyStart.y) / (xyEnd.X - xyStart.X) 'Slope
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nK = xyStart.y - nM * xyStart.X 'Y-Axis intercept
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nA = (1 + nM ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.y * nM))
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xyInt.y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False

    End Function
    Sub PR_CalcFingerBlendingProfile(ByRef xyTop As XY, ByRef xyBottom As XY, ByRef xyFinger As XY, ByRef Digit As Finger, ByRef ReturnProfile As curve, ByRef xyThumbPalm As XY)
        'Procedure to return a smooth inflecting curve between two points
        '
        'The assumptions about the required curve are :-
        '    1. The Top Point is above the Bottom Point
        '    2. The curve leaves the Top point at 270
        '    3. The curve enters the bottom point at 270
        '    4. The curve is aproximated by two arcs
        '       that are tangential at the mid point
        '       between xyTop and xyBottom
        '    5. The curve inflects at the point in 4 above
        '    6. The radius of the arcs are the same
        '
        'The returned profile has the following features :-
        'If there is no missing finger then
        '    1. Decreases in Y from xyTop to xyBottom
        '    2. Has 7 vertex
        '    3. Vertex = 1 = xyTop
        '    4. Vertex = 7 = xyBottom
        '    5. Where xyTop.X = xyBottom.X returns a
        '       two point profile.
        'If there is a missing finger then
        '    1. Decreases in Y from xyTop to xyBottom
        '    2. Has 6 vertex
        '    3. Vertex = 1 = xyFinger
        '    4. Vertex = 6 = xyBottom
        '    4. Vertex = 2 to 5 create a fillet
        '       that can be manually edited
        '
        'Notes :-
        '    This is a fairly restricted routine designed
        '    to blend the profile above the wrist into the
        '    glove with a smooth inflecting curve.
        '
        'Modifications :-
        '    The point given in xyThumbPalm is the current point
        '    of the global variable xyThumbPalm(1) this is modified
        '    to the intersection of the blended curve at the height
        '    given by xyThumbPalm.y
        '
        '
        Static aInc, aAngle, nLength, rAngle, nRadius As Double
        Static xyPt1, xyCenter, xyMidPoint, xyPt2 As XY
        Static xyConstruct, xyInt As XY
        Static iDirection As Short
        Static ii, MirrorResult As Short

        nLength = FN_CalcLength(xyBottom, xyTop)
        aAngle = FN_CalcAngle(xyBottom, xyTop)

        'If the angle is less than 90 Degrees then we just carry on.
        'However if the angle is greater, then we mirror this angle in the
        'y axis and calculate the points as befor, then we mirror the result
        '
        MirrorResult = False
        'UPGRADE_WARNING: Couldn't resolve default property of object xyConstruct. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyConstruct = xyTop
        If aAngle > 90 Then
            MirrorResult = True
            aAngle = 90 - (aAngle - 90)
            xyConstruct.X = xyBottom.X - (xyTop.X - xyBottom.X)
        End If

        'Degenerate to a straight line if xyTop.X = xyBottom.X
        'Then exit the sub routine
        If System.Math.Abs(aAngle) = 90 Or aAngle = 270 Then
            ReturnProfile.n = 2
            Dim X(2), Y(2) As Double
            ReturnProfile.X = X
            ReturnProfile.y = Y
            ReturnProfile.X(1) = xyTop.X
            ReturnProfile.y(1) = xyTop.y
            ReturnProfile.X(2) = xyBottom.X
            ReturnProfile.y(2) = xyBottom.y
            GoTo MissingFingers
        Else
            ReturnProfile.n = 7
            Dim X(7), Y(7) As Double
            ReturnProfile.X = X
            ReturnProfile.y = Y
            ReturnProfile.X(1) = xyConstruct.X
            ReturnProfile.y(1) = xyConstruct.y
            ReturnProfile.X(7) = xyBottom.X
            ReturnProfile.y(7) = xyBottom.y
        End If

        'Get midpoint
        PR_CalcPolar(xyBottom, aAngle, nLength / 2, xyMidPoint)
        ReturnProfile.X(4) = xyMidPoint.X
        ReturnProfile.y(4) = xyMidPoint.y

        'Calculate radius
        rAngle = aAngle * (PI / 180) 'Convert angle to Radians
        nRadius = (nLength / 4) / System.Math.Cos(rAngle)

        'First arc points
        PR_MakeXY(xyCenter, xyConstruct.X - nRadius, xyConstruct.y)
        aAngle = FN_CalcAngle(xyCenter, xyMidPoint)
        aInc = (360 - aAngle) / 3

        'Find Intersection
        If xyThumbPalm.y > xyMidPoint.y And xyThumbPalm.y < xyTop.y Then
            PR_MakeXY(xyPt1, xyCenter.X, xyThumbPalm.y)
            PR_MakeXY(xyPt2, xyCenter.X + 10, xyThumbPalm.y)
            If FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xyThumbPalm = xyInt
                If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
            End If
        End If

        For ii = 3 To 2 Step -1
            aAngle = aAngle + aInc
            PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
            ReturnProfile.X(ii) = xyPt1.X
            ReturnProfile.y(ii) = xyPt1.y
        Next ii

        'Second arc points
        PR_MakeXY(xyCenter, xyBottom.X + nRadius, xyBottom.y)
        aAngle = FN_CalcAngle(xyCenter, xyMidPoint)

        'Find Intersection
        If xyThumbPalm.y <= xyMidPoint.y And xyThumbPalm.y > xyBottom.y Then
            PR_MakeXY(xyPt1, xyCenter.X, xyThumbPalm.y)
            PR_MakeXY(xyPt2, xyCenter.X - 10, xyThumbPalm.y)
            If FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object xyThumbPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xyThumbPalm = xyInt
                If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
            End If
        End If

        For ii = 5 To 6
            aAngle = aAngle + aInc
            PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
            ReturnProfile.X(ii) = xyPt1.X
            ReturnProfile.y(ii) = xyPt1.y
        Next ii

        'Mirroring the result simplifies the code above, as we only have
        'to code for the case where angle < 90
        'N.B. Mirroring in the Y axis along the line X = xyBottom.X
        If MirrorResult Then
            For ii = 1 To ReturnProfile.n
                ReturnProfile.X(ii) = xyBottom.X - (ReturnProfile.X(ii) - xyBottom.X)
            Next ii
        End If

MissingFingers:

        Static nX, nY As Double

        If Digit.Len_Renamed = 0 And (Digit.Tip = 0 Or Digit.Tip = 10) Then

            nY = xyFinger.y - ReturnProfile.y(2)
            nX = xyFinger.X - ReturnProfile.X(1)
            iDirection = System.Math.Sign(xyFinger.X - ReturnProfile.X(1))
            nX = System.Math.Abs(nX)
            'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nRadius = min(nX, nY) - 0.05

            'Straight line
            ReturnProfile.n = 6
            ReturnProfile.X(1) = xyFinger.X
            ReturnProfile.y(1) = xyFinger.y
            PR_MakeXY(xyCenter, xyTop.X + (iDirection * nRadius), xyFinger.y - nRadius)
            aAngle = 90
            aInc = (90 / 3) * iDirection
            For ii = 2 To 5
                PR_CalcPolar(xyCenter, aAngle, nRadius, xyPt1)
                ReturnProfile.X(ii) = xyPt1.X
                ReturnProfile.y(ii) = xyPt1.y
                aAngle = aAngle + aInc
            Next ii
            ReturnProfile.X(6) = xyBottom.X
            ReturnProfile.y(6) = xyBottom.y
        End If
    End Sub
    Public Function FN_BiArcCurve(ByRef xyStart As XY, ByRef aStart As Double, ByRef xyEnd As XY, ByRef aEnd As Double, ByRef Profile As BiArc) As Short
        'Procedure to fit two arcs with a common tangent
        'between two points at the tangent angles specified.
        'Returning a BiArc curve.
        '
        'InputArm
        '        xyStart     Start point (XY System)
        '        xyEnd       End Point   (XY System)
        '        aStart      Start Tangent Angle in degrees
        '        aEnd        End Tangent Angle in degrees
        '
        'Output
        '        Profile     BiArc in (XY System)
        '
        'Restrictions:-
        'Angles must be positive and < 360
        'The data must result in a NON-INFLECTING curve.
        '
        '
        'Notes:-
        'The BiArc is a structure that is used to represent
        'the curve.
        '
        '        Type BiArc
        '                xyStart     as XY
        '                xyTangent   as XY
        '                xyEnd       as XY
        '                xyR1        as XY
        '                xyR2        as XY
        '                nR1         as Double
        '                nR2         as Double
        '        end
        '
        'The point of this represntation is to make
        'the drawing of the curve as simple as possible,
        'without the need to make further calculations.
        '
        '
        'Acknowledgements:-
        '
        '    British Ship Research Association (BSRA)
        '    Technical Memorandum No. 388
        '
        '   "The Fitting of Smooth Curves by Circular
        '    Arcs and Straight Lines"
        '    K.M.Bolton, B.Sc., M.Sc., Grad.I.M.A.
        '    October 1970
        '
        'KNOWN BUGS
        'Needs much more work to make a more general robust tool.
        'Works OK for some special cases (cf THUMB Curves)
        'GG 22.Mar.96
        '
        '
        Dim nTmp, aAxis, nLength As Double
        Dim Phi1, Theta1, Theta2, Phi2 As Double
        Dim C1, S1, d, B, A, c, P, S2, C2 As Double
        Dim R1, Rs, R2 As Double
        Dim Rmin As Object
        Dim uvR2, uvR1, uvOrigin As XY
        '    Dim fError

        'Use a file for printing results for debug
        'Open file
        '    fError = FreeFile
        '    Open "C:\TMP\FIT_ERR.DAT" For Output As fError


        'Initially return false
        FN_BiArcCurve = False

        'Check for silly data and return
        If aStart = aEnd Then GoTo Error_Close
        If xyStart.y = xyEnd.y And xyStart.X = xyEnd.X Then GoTo Error_Close

        'Translate tangent angle in XY system to the UV system.
        'Where the U-Axis is specified by the line xyStart,xyEnd
        '
        aAxis = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)
        'Print #fError, "aAxis degrees ="; aAxis
        'Print #fError, "nLength ="; nLength

        Theta1 = aStart - aAxis
        If Theta1 = 0 Then GoTo Error_Close 'Straight line

        If Theta1 < 0 Then Theta1 = 360 + Theta1
        'Print #fError, "Theta1 degrees ="; Theta1
        Theta1 = Theta1 * (PI / 180)
        'Print #fError, "Theta1="; Theta1

        Theta2 = aEnd - aAxis
        If Theta2 < 0 Then Theta2 = 360 + Theta2
        'Print #fError, "Theta2 degrees="; Theta2
        Theta2 = Theta2 * (PI / 180)
        'Print #fError, "Theta2="; Theta2


        'Check that it is non-inflecting
        'return an error (false) for the inflecting case
        'let the calling routine worry about handling it
        If (Theta1 > 0 And Theta1 < (PI / 2)) And (Theta2 > 0 And Theta2 < (PI / 2)) Then GoTo Error_Close
        If (Theta1 > (3 * (PI / 2)) And Theta1 < (2 * PI)) And (Theta2 > (3 * (PI / 2)) And Theta2 < (2 * PI)) Then GoTo Error_Close

        'Calculate acute unsigned tangent angles to the line in the UV
        'co-ordinate system
        ' Phi1 = Theta1
        If Theta1 < PI Then
            Phi1 = Theta1
        Else
            Phi1 = System.Math.Abs((2 * PI) - Theta1)
        End If
        If Theta2 < PI Then
            Phi2 = Theta2
        Else
            Phi2 = System.Math.Abs((2 * PI) - Theta2)
        End If

        'Print #fError, "Phi1="; Phi1
        'Print #fError, "Phi2="; Phi2


        'Calculate R1 and R2
        S1 = System.Math.Abs(System.Math.Sin(Theta1))
        C1 = (-System.Math.Sin(Theta1) * System.Math.Cos(Theta1)) / S1
        'Print #fError, "S1="; S1; "C1="; C1

        S2 = System.Math.Abs(System.Math.Sin(Theta2))
        C2 = (System.Math.Sin(Theta2) * System.Math.Cos(Theta2)) / S2
        'Print #fError, "S2="; S2; "C2="; C2

        P = FN_CalcLength(xyStart, xyEnd)
        'Print #fError, "P="; P

        If Phi1 <> Phi2 Then
            'Print #fError, "Phi1 <> Phi2"
            A = S1 + S2
            B = (S1 * S2) - (C1 * C2) + 1
            c = S2
            Rs = (P * c) / B
            nTmp = (c ^ 2) - (c * A) + (B / 2)
            'Print #fError, "A="; A; "B="; B; "C="; C
            'Print #fError, "Rs="; Rs
            'Print #fError, "nTmp="; nTmp
            'As this is a root we check that it is not -ve
            If nTmp < 0 Then GoTo Error_Close
            nTmp = (P * System.Math.Sqrt(nTmp)) / B
            'Print #fError, "nTmp="; nTmp
            If Phi1 > Phi2 Then
                R1 = Rs - nTmp
            Else
                R1 = Rs + nTmp
            End If
            'Print #fError, "R1="; R1
            If R1 <= 0 Then GoTo Error_Close
            d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
            If ((2 * R1 * B) - (2 * P * c)) = 0 Then
                d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B))
            End If
            'Print #fError, "D="; d
            R2 = R1 - d
            'Print #fError, "R2="; R2
        Else
            'Print #fError, "Phi1 = Phi2"
            A = 2 * System.Math.Sin(Theta1)
            B = (System.Math.Sin(Theta1) ^ 2) - (System.Math.Cos(Theta1) ^ 2) + 1
            R1 = (P * A) / (2 * B)
            R2 = R1
            'Print #fError, "A="; A; "B="; A;
            'Print #fError, "R1="; R1
        End If

        'The radi R1 and R2 must be greater than the specified
        'minimum Rmin
        'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rmin = nLength / 3
        'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If R1 < Rmin Then
            'Print #fError, "R1 < Rmin"
            'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            R1 = Rmin
            d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
            'Print #fError, "D="; d
            R2 = R1 - d
            'Print #fError, "R1="; R1
            'Print #fError, "R2="; R2
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If R2 < Rmin Then
            'Print #fError, "R2 < Rmin"
            'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            R2 = Rmin
            'UPGRADE_WARNING: Couldn't resolve default property of object Rmin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
            'Print #fError, "R1="; R1
            'Print #fError, "R2="; R2
        End If

        'Check for errors
        If R1 < 0 Or R2 < 0 Then GoTo Error_Close

        'Using the calculated radi, create BiArc Curve
        'Start and end points of bi-arc curve
        'UPGRADE_WARNING: Couldn't resolve default property of object Profile.xyStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Profile.xyStart = xyStart
        'UPGRADE_WARNING: Couldn't resolve default property of object Profile.xyEnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Profile.xyEnd = xyEnd

        'Centers of arcs
        'Get UV co-ordinates
        PR_MakeXY(uvR1, R1 * S1, R1 * C1)
        'Print #fError, "uvR1="; uvR1.X, uvR1.Y
        'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
        PR_MakeXY(uvR2, P - (R2 * S2), R2 * C2)
        'Print #fError, "uvR2="; uvR2.X, uvR2.Y
        'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
        PR_MakeXY(uvOrigin, 0, 0)
        'Translate to XY co-ordinates
        PR_CalcPolar(xyStart, FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1)
        PR_CalcPolar(xyStart, FN_CalcAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2)
        'Print #fError, "R1="; R1
        'Print #fError, "R2="; R2
        'Tangent point on arc
        If R1 < R2 Then
            PR_CalcPolar(Profile.xyR1, FN_CalcAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent)
        Else
            PR_CalcPolar(Profile.xyR1, FN_CalcAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent)
        End If
        'radi of arcs
        Profile.nR1 = R1
        Profile.nR2 = R2

        'Test that the Tangent point lies between the start and end points
        If FN_CalcLength(xyStart, Profile.xyTangent) > nLength Then GoTo Error_Close
        If FN_CalcLength(xyEnd, Profile.xyTangent) > nLength Then GoTo Error_Close

        'return true as we have a sucessful fit
        'Close #fError
        FN_BiArcCurve = True
        Exit Function

Error_Close:
        'Print #fError, "Error and close"
        'Close #fError

    End Function
    Function FN_ValidateExtensionData() As String
        'Procedure to validated the extension data
        Static sError, nL As String
        Static iNoPleats, ii As Short

        nL = Chr(13)
        sError = "" 'Initialise because of static

        'If normal glove then exit
        If g_ExtendTo = GLOVE_NORMAL Then
            FN_ValidateExtensionData = ""
            Exit Function
        End If

        'Check pleats
        iNoPleats = 0
        For ii = 1 To 4
            If g_nPleats(ii) > 0 Then iNoPleats = iNoPleats + 1
        Next ii
        If iNoPleats + 1 > g_iNumTapesWristToEOS Then
            sError = sError & "Number of pleats exceeds availble spaces between the arm tapes!" & nL & "Disable pleats by Double Clicking on pleat label." & nL
        End If

        'Check that calculate has been used
        '>>>>>>>
        'Not yet implemented
        '>>>>>>>

        'Check that wrist and first tape are the same
        ''-----If g_nCir(g_iWristPointer) <> g_nWrist Then
        'If g_nCir(1, g_iWristPointer) <> g_nWrist Then
        '    sError = sError & "Wrist tape of the glove and the extension are different!" & nL
        'End If

        'Check on style D flaps
        If g_EOSType = ARM_FLAP And InStr(1, g_sFlapType, "D") > 0 And g_nWaistCir = 0 Then
            sError = sError & "A Waist circumference must be given for a D-Style flap" & nL
        End If

        'Return error message
        FN_ValidateExtensionData = sError
    End Function
    Sub PR_DisplayTextInches(ByRef ctlText As System.Windows.Forms.Control, ByRef ctlCaption As System.Windows.Forms.Control)
        Static nLen As Double
        nLen = FN_InchesValue(ctlText)
        If nLen <> -1 Then ctlCaption.Text = fnInchestoText(nLen)
    End Sub

    Sub PR_LoadPowernetChart()
        'Procedure to load the reduction charts for the arms from
        'disk
        '
        Static sModulus, sFile, sLine As String
        Static jj, fChart, ii, nn As Short

        ''--------sFile = g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        sFile = sSettingsPath & "\POWERNET.DAT"

        If Dir(sFile) = "" Then
            MsgBox("Template file not found " & sFile, 48, g_sDialogID)
            Exit Sub
        End If
        'Open file
        fChart = FreeFile()
        FileOpen(fChart, sFile, VB.OpenMode.Input)

        'Get ARM chart reductions
        'NB these are fixed format
        ii = 1
        While Not EOF(fChart) And ii <= NOFF_MODULUS
            Input(fChart, sModulus)
            Input(fChart, sLine)

            For jj = 0 To 22
                g_iPowernet(ii, jj + 1) = Val(Mid(sLine, (jj * 4) + 1, 4))
            Next jj
            ii = ii + 1
        End While
        'Close file
        FileClose(fChart)
        sLine = ""
        sModulus = ""
    End Sub
    Sub PR_CalculateExtension(ByRef xyWristUlnar As XY, ByRef xyWristRadial As XY, ByRef nInsert As Double,
                              ByRef UlnarProfile As curve, ByRef RadialProfile As curve)

        'Procedure to calculate the POINTS
        'used to draw the extension of the glove above the
        'wrist
        'The Procedure also supplies the data that is to be
        'printed at each tape
        'N.B.
        'The wrist points will be modified if the wrist has been
        'calculated

        Dim nValue, nMidWristX, nSpacing As Double
        Dim iStartTape, ii, iLastTape As Short
        Dim nFiguredValue As Double
        Dim iVertex As Short

        nMidWristX = xyWristUlnar.X + (System.Math.Abs(xyWristRadial.X - xyWristUlnar.X) / 2)
        UlnarProfile.Initialize()
        RadialProfile.Initialize()
        If g_iNumTapesWristToEOS = 0 Or g_iNumTapesWristToEOS = 1 Then
            'Extension ends at wrist
            UlnarProfile.n = 1
            Dim X(1), Y(1) As Double
            UlnarProfile.X = X
            UlnarProfile.y = Y
            UlnarProfile.X(1) = xyWristUlnar.X
            UlnarProfile.y(1) = xyWristUlnar.y - 0.75

            RadialProfile.n = 1
            Dim RX(1), RY(1) As Double
            RadialProfile.X = RX
            RadialProfile.y = RY
            RadialProfile.X(1) = xyWristRadial.X
            RadialProfile.y(1) = xyWristRadial.y - 0.75

            Exit Sub
        End If

        If g_ExtendTo = GLOVE_NORMAL Then
            'Glove only extends two or three tapes past wrist
            'Use the charts to get the reductions
            '
            'Although given as part of the Extension
            'The normal glove uses data from the Hand Part of the dialogue
            '
            '
            UlnarProfile.n = g_iNumTapesWristToEOS
            UlnarProfile.X(1) = xyWristUlnar.X
            UlnarProfile.y(1) = xyWristUlnar.y

            RadialProfile.n = g_iNumTapesWristToEOS
            RadialProfile.X(1) = xyWristRadial.X
            RadialProfile.y(1) = xyWristRadial.y

            'God knows why this is but it works, and I don't give a !"£$%^&*(
            'any more
            iStartTape = 2
            iLastTape = g_iNumTapesWristToEOS + 1

            For ii = iStartTape To iLastTape

                'NB we add 1 to ii as the charts have wrist at ii=1, palm at ii=2
                'and 1-1/2 tape at ii=3.
                'The palm being out of order is a pain but it made it easier to
                'enter the chart data (swings and roundabouts, sorry!)
                '
                'NB use of 1/2 scale
                ''-------nFiguredValue = FN_FigureTape(ii + 1, g_nCir(ii))
                nFiguredValue = FN_FigureTape(ii + 1, g_nCir(1, ii))

                Select Case g_iInsertStyle
                    Case 0
                        If g_OnFold Then
                            nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
                        Else
                            nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
                        End If
                    Case 1
                        nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
                    Case 2, 3
                        nValue = (nFiguredValue + (2 * EIGHTH)) / 2
                End Select

                If g_OnFold Then
                    UlnarProfile.X(ii) = UlnarProfile.X(1)
                    RadialProfile.X(ii) = UlnarProfile.X(1) + nValue
                Else
                    UlnarProfile.X(ii) = nMidWristX - (nValue / 2)
                    RadialProfile.X(ii) = nMidWristX + (nValue / 2)
                End If

                'Standard spacing
                ''Changed for #215 in the issue list
                ''nSpacing = 1.375
                nSpacing = 1.2

                'We are working from the top (wrist) down
                UlnarProfile.y(ii) = UlnarProfile.y(ii - 1) - nSpacing
                RadialProfile.y(ii) = RadialProfile.y(ii - 1) - nSpacing

            Next ii

        End If


        If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_AXILLA Then
            'As we have recalculated the wrist we need to start at the
            'wrist
            UlnarProfile.n = g_iNumTapesWristToEOS
            UlnarProfile.y(1) = xyWristUlnar.y

            RadialProfile.n = g_iNumTapesWristToEOS
            RadialProfile.y(1) = xyWristRadial.y

            iStartTape = g_iWristPointer
            iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)

            'Note: As we can allow the wrist to be any tape we need a
            'vertex counter
            iVertex = 1

            For ii = iStartTape To iLastTape
                'NB use of 1/2 scale
                ''-----------nFiguredValue = (g_nCir(ii) * ((100 - g_iRed(ii)) / 100))
                nFiguredValue = (g_nCir(1, ii) * ((100 - g_iRed(ii)) / 100))
                Select Case g_iInsertStyle
                    Case 0
                        If g_OnFold Then
                            nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
                        Else
                            nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
                        End If
                    Case 1
                        nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
                    Case 2, 3
                        nValue = (nFiguredValue + (2 * EIGHTH)) / 2
                End Select

                If g_OnFold Then
                    UlnarProfile.X(iVertex) = UlnarProfile.X(1)
                    RadialProfile.X(iVertex) = UlnarProfile.X(1) + nValue
                Else
                    UlnarProfile.X(iVertex) = nMidWristX - (nValue / 2)
                    RadialProfile.X(iVertex) = nMidWristX + (nValue / 2)
                End If


                'Setup the notes for the vertex
                'We do this here as we have all the data and it simplifies
                'the drawing side
                TapeNote(iVertex).sTapeText = LTrim(Mid(g_sTapeText, ((ii + 1) * 3) + 1, 3))
                TapeNote(iVertex).iTapePos = ii
                ''------------TapeNote(iVertex).nCir = g_nCir(ii)
                TapeNote(iVertex).nCir = g_nCir(1, ii)
                TapeNote(iVertex).iGms = g_iGms(ii)
                TapeNote(iVertex).iRed = g_iRed(ii)
                TapeNote(iVertex).iMMs = g_iMMs(ii)


                'Standard spacing
                nSpacing = 1.375

                'Account all for pleats
                'Wrist (as we have started at the wrist we use, iStartTape + 1
                'and iStartTape + 2 in this case)
                If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(1))
                If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(2))
                'XXXXXXXXXX
                'XXXX Careful now! this won't work if only 3 or less tapes given
                'XXXXXXXXXX
                'Axilla
                If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(3))
                If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(4))

                'We are working from the top (wrist) down
                'iVertex = 1 is set before the For Loop (Values from wrist)
                If iVertex <> 1 Then
                    UlnarProfile.y(iVertex) = UlnarProfile.y(iVertex - 1) - nSpacing
                    RadialProfile.y(iVertex) = RadialProfile.y(iVertex - 1) - nSpacing
                    If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
                End If

                'Increment vertex count
                iVertex = iVertex + 1

            Next ii

            'Reset given wrist points  xyWristUlnar and xyWristRadial
            PR_MakeXY(xyWristUlnar, UlnarProfile.X(1), UlnarProfile.y(1))
            PR_MakeXY(xyWristRadial, RadialProfile.X(1), RadialProfile.y(1))
        End If

    End Sub
    Function Arccos(ByRef x As Double) As Double
        Arccos = System.Math.Atan(-x / System.Math.Sqrt(-x * x + 1)) + 1.5708
    End Function
    Sub PR_CalcWristBlendingProfile(ByRef xyTSt As XY, ByRef xyTEnd As XY, ByRef xyBSt As XY, ByRef xyBEnd As XY, ByRef ReturnProfile As curve, ByRef xyThumbPalm As XY)
        'Procedure to return a smooth inflecting curve between two points
        '
        'N.B. Parameters are given in the order of decreasing Y
        '
        Static nR1, rAngle, aA2, nL3, nL1, nL2, aA1, aA3, aInc, nR2 As Double
        Static xyBotSt, xyTopSt, xyCenter, xyMidPoint, xyTopEnd, xyBotEnd As XY
        Static xyR1, xyR2 As XY
        Static nA, nThirdOfL2, aAngle As Double
        Static xyPt2, xyPt1, xyInt As XY
        Static TopIsArc, MirrorResult, ii, Direction, BottomIsArc As Short
        Static Intersection As Double
        Static xyTmp(10) As XY
        Static nTol As Double

        'Do this as we can't use ByVal
        xyTopSt = xyTSt
        xyTopEnd = xyTEnd
        xyBotSt = xyBSt
        xyBotEnd = xyBEnd

        'If the angle is less than 90 Degrees then we just carry on.
        'However if the angle is greater, then we mirror this angle in the
        'y axis and calculate the points as befor, then we mirror the result
        '
        MirrorResult = False
        aA2 = FN_CalcAngle(xyBSt, xyTEnd)
        aA3 = FN_CalcAngle(xyBEnd, xyBSt)
        Direction = 1
        If (aA2 > 90) Or ((aA2 = 90) And (aA3 < 90)) Then
            MirrorResult = True
            Direction = -1
            aA2 = 90 - (aA2 - 90)
            xyTopSt.X = xyBEnd.X - (xyTopSt.X - xyBEnd.X)
            xyTopEnd.X = xyBEnd.X - (xyTopEnd.X - xyBEnd.X)
            xyBotSt.X = xyBEnd.X - (xyBotSt.X - xyBEnd.X)
        End If
        aA1 = FN_CalcAngle(xyTopEnd, xyTopSt)
        aA3 = FN_CalcAngle(xyBotEnd, xyBotSt)

        'Degenerate to a straight line
        'Then exit the sub routine
        If aA1 = aA2 And aA2 = aA3 And (aA1 = 90 Or aA1 = 270) Then
            ReturnProfile.n = 2
            Dim X(2), Y(2) As Double
            ReturnProfile.X = X
            ReturnProfile.y = Y
            ReturnProfile.X(1) = xyTSt.X
            ReturnProfile.y(1) = xyTSt.y
            ReturnProfile.X(2) = xyBEnd.X
            ReturnProfile.y(2) = xyBEnd.y
            Exit Sub
        End If

        nL1 = FN_CalcLength(xyTEnd, xyTSt)
        nL2 = FN_CalcLength(xyBSt, xyTEnd)
        nL3 = FN_CalcLength(xyBEnd, xyBSt)

        'Get Included angles & radius & Centers of Arcs
        nThirdOfL2 = nL2 / 3

        'Top Arc
        If aA1 <> aA2 Then
            'Calculate the points for the ARC for this section
            TopIsArc = True
            nA = FN_CalcLength(xyTopSt, xyBotSt)
            nA = ((nL1 ^ 2 + nL2 ^ 2) - nA ^ 2) / (2 * nL1 * nL2)
            rAngle = Arccos(nA) / 2
            nR1 = nThirdOfL2 * System.Math.Tan(rAngle)
            PR_CalcPolar(xyTopEnd, 90, nThirdOfL2, xyTmp(2))
            PR_CalcPolar(xyTopEnd, FN_CalcAngle(xyTopEnd, xyBotSt), nThirdOfL2, xyTmp(5))
            PR_CalcPolar(xyTmp(2), 180, nR1, xyR1)

            'Top Arc Points
            aAngle = FN_CalcAngle(xyR1, xyTmp(5))
            aInc = (360 - aAngle) / 3
            For ii = 4 To 3 Step -1
                aAngle = aAngle + aInc
                PR_CalcPolar(xyR1, aAngle, nR1, xyTmp(ii))
            Next ii
        Else
            'Calculate the points for the STRAIGHT LINE for this section
            'Note: we use the same noff points for consistancy
            'with respect to the editor and subsequent points
            TopIsArc = False
            PR_CalcPolar(xyTopEnd, 90, nThirdOfL2, xyTmp(2))
            PR_CalcPolar(xyTopEnd, 90, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(3))
            PR_CalcPolar(xyTopEnd, 270, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(4))
            PR_CalcPolar(xyTopEnd, 270, nThirdOfL2, xyTmp(5))
        End If

        If aA2 <> aA3 Then
            'Bottom arc
            BottomIsArc = True
            nA = FN_CalcLength(xyTopEnd, xyBotEnd)
            nA = ((nL2 ^ 2 + nL3 ^ 2) - nA ^ 2) / (2 * nL3 * nL2)
            rAngle = Arccos(nA) / 2
            nR2 = nThirdOfL2 * System.Math.Tan(rAngle)
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))

            'Establish
            aAngle = FN_CalcAngle(xyBotEnd, xyTopEnd)
            If aAngle < FN_CalcAngle(xyBotEnd, xyBotSt) Then
                aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) - 90
            Else
                aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) + 90
            End If
            PR_CalcPolar(xyTmp(6), aAngle, nR2, xyR2)

            'Check that the gap between the wrist point and the arc is less than
            'or equal to 0.0625 ie. 1/16"
            'If not then make it so.
            nTol = 0.0625
            nA = FN_CalcLength(xyBotSt, xyR2)
            If (nA - nR2) > nTol Then
                nR2 = nR2 - ((nA - nR2) - nTol)
                PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
                PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))
                'aAngle from above
                PR_CalcPolar(xyTmp(6), aAngle, nR2, xyR2)
            End If

            'Bottom arc points
            aAngle = FN_CalcAngle(xyR2, xyTmp(6))
            aInc = (FN_CalcAngle(xyR2, xyTmp(9)) - aAngle) / 3
            For ii = 7 To 8
                aAngle = aAngle + aInc
                PR_CalcPolar(xyR2, aAngle, nR2, xyTmp(ii))
            Next ii

        Else
            BottomIsArc = False
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6))
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(7))
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(8))
            PR_CalcPolar(xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9))
        End If

        'Find Intersection
        Intersection = False
        If xyThumbPalm.y >= xyTmp(2).y Then
            nA = xyThumbPalm.y - xyTmp(2).y
            If nA = 0 Then
                xyThumbPalm = xyTmp(2)
            Else
                PR_MakeXY(xyThumbPalm, xyTmp(2).X + (nA / System.Math.Tan(aA1 * (PI / 180))), xyThumbPalm.y)
            End If
            Intersection = True

        ElseIf xyThumbPalm.y < xyTmp(2).y And xyThumbPalm.y > xyTmp(5).y Then
            If TopIsArc Then
                PR_MakeXY(xyPt1, xyR1.X, xyThumbPalm.y)
                PR_MakeXY(xyPt2, xyR1.X + 10, xyThumbPalm.y)
                If FN_CirLinInt(xyPt1, xyPt2, xyR1, nR1, xyInt) Then
                    xyThumbPalm = xyInt
                    Intersection = True
                End If
            Else
                nA = xyThumbPalm.y - xyTmp(5).y
                If nA = 0 Then
                    xyThumbPalm = xyTmp(5)
                Else
                    PR_MakeXY(xyThumbPalm, xyTmp(5).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.y)
                End If
                Intersection = True
            End If

        ElseIf xyThumbPalm.y <= xyTmp(5).y And xyThumbPalm.y >= xyTmp(6).y Then
            nA = xyThumbPalm.y - xyTmp(6).y
            If nA = 0 Then
                xyThumbPalm = xyTmp(6)
            Else
                PR_MakeXY(xyThumbPalm, xyTmp(6).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.y)
            End If
            Intersection = True

        ElseIf xyThumbPalm.y < xyTmp(6).y And xyThumbPalm.y > xyTmp(9).y Then
            If BottomIsArc Then
                PR_MakeXY(xyPt1, xyTmp(6).X + nR2, xyThumbPalm.y)
                PR_MakeXY(xyPt2, xyTmp(6).X - nR2, xyThumbPalm.y)
                If FN_CirLinInt(xyPt1, xyPt2, xyR2, nR2, xyInt) Then
                    xyThumbPalm = xyInt
                    Intersection = True
                End If
            Else
                nA = xyThumbPalm.y - xyTmp(9).y
                If nA = 0 Then
                    xyThumbPalm = xyTmp(9)
                Else
                    PR_MakeXY(xyThumbPalm, xyTmp(9).X + (nA / System.Math.Tan(aA2 * (PI / 180))), xyThumbPalm.y)
                End If
                Intersection = True
            End If
        End If
        If Intersection And MirrorResult Then xyThumbPalm.X = xyBotEnd.X - (xyThumbPalm.X - xyBotEnd.X)

        ReturnProfile.n = 10
        Dim X1(10), Y1(10) As Double
        ReturnProfile.X = X1
        ReturnProfile.y = Y1
        ReturnProfile.X(1) = xyTopSt.X
        ReturnProfile.y(1) = xyTopSt.y
        ReturnProfile.X(10) = xyBotEnd.X
        ReturnProfile.y(10) = xyBotEnd.y

        For ii = 2 To 9
            ReturnProfile.X(ii) = xyTmp(ii).X
            ReturnProfile.y(ii) = xyTmp(ii).y
        Next ii


        'Mirroring the result simplifies the code above, as we only have
        'to code for the case where angle < 90
        'N.B. Mirroring in the Y axis along the line X = xyBottom.X
        If MirrorResult Then
            For ii = 1 To ReturnProfile.n
                ReturnProfile.X(ii) = xyBotEnd.X - (ReturnProfile.X(ii) - xyBotEnd.X)
            Next ii
        End If
    End Sub
    Sub PR_MirrorPointInYaxis(ByRef nXValue As Double, ByRef nTranslate As Double, ByRef xyPoint As XY)
        'Procedure to mirror a Point in the y axis about
        'a line given by the X value
        '
        xyPoint.X = (nXValue - (xyPoint.X - nXValue)) + nTranslate
    End Sub
    Sub PR_MirrorCurveInYaxis(ByRef nXValue As Double, ByRef nTranslate As Double, ByRef Profile As curve)
        'Procedure to mirror a curve in the y axis about
        'a line given by the X value
        '
        Static ii As Short
        For ii = 1 To Profile.n
            Profile.X(ii) = (nXValue - (Profile.X(ii) - nXValue)) + nTranslate
        Next ii

    End Sub
    Sub PR_DrawFitted(ByRef Profile As curve)
        Dim ii As Short
        Select Case Profile.n
            Case 0 To 1
                Exit Sub
            Case 3
                Dim Bulge(Profile.n) As Double
                For ii = 1 To Profile.n
                    Bulge(ii) = 0
                Next ii
                PR_DrawPoly(Profile, Bulge)
                PR_AddXDataValToLast("ID", g_sFileNo + "Left")
            Case Else
                Dim ptColl As Point3dCollection = New Point3dCollection()
                For ii = 1 To Profile.n
                    ptColl.Add(New Point3d(Profile.X(ii) + xyMGlvInsertion.X, Profile.y(ii) + xyMGlvInsertion.y, 0))
                Next
                PR_DrawSpline(ptColl)
                PR_AddXDataValToLast("ID", g_sFileNo + "Left")
        End Select
    End Sub
    Sub PR_DrawPoly(ByRef Profile As curve, ByRef Bulge As Double())
        Dim ii As Short

        'Exit if nothing to draw
        If Profile.n <= 1 Then Exit Sub
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                            OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Using acPoly As Polyline = New Polyline()
                For ii = 1 To Profile.n
                    acPoly.AddVertexAt(ii - 1, New Point2d(Profile.X(ii) + xyMGlvInsertion.X, Profile.y(ii) + xyMGlvInsertion.y), Bulge(ii), 0, 0)
                Next ii

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                g_idLastZipper = acPoly.ObjectId
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    'To Draw Spline
    Private Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
                g_idLastZipper = acSpline.ObjectId
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_AddXDataValToLast(ByRef sDBName As String, ByRef sDBValue As String)
        'The last entity is given by hEnt
        ''--------PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & sDBName & QQ & CC & QQ & sDBValue & QQ & ");")
        If (g_idLastZipper.IsValid = False) Then
            Exit Sub
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As DBObject = acTrans.GetObject(g_idLastZipper, OpenMode.ForWrite)
            Dim acRegAppTbl As RegAppTable
            acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
            Dim acRegAppTblRec As RegAppTableRecord
            If acRegAppTbl.Has(sDBName) = False Then
                acRegAppTblRec = New RegAppTableRecord
                acRegAppTblRec.Name = sDBName
                acRegAppTbl.UpgradeOpen()
                acRegAppTbl.Add(acRegAppTblRec)
                acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
            End If
            Using rb As New ResultBuffer
                rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, sDBName))
                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, sDBValue))
                acEnt.XData = rb
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawArc(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object

        nRad = FN_CalcLength(xyCen, xyArcStart)
        nStartAng = FN_CalcAngle(xyCen, xyArcStart) * (PI / 180)
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd) * (PI / 180)
        nDeltaAng = nEndAng - nStartAng
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Create an arc that is at 6.25,9.125 with a radius of 6, and
            '' starts at 64 degrees and ends at 204 degrees
            Using acArc As Arc = New Arc(New Point3d(xyCen.X + xyMGlvInsertion.X, xyCen.y + xyMGlvInsertion.y, 0),
                                     nRad, nStartAng, nEndAng)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has("ID") = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = "ID"
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, g_sFileNo + "Left"))
                    acArc.XData = rb
                End Using
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawLine(ByRef xyStart As XY, ByRef xyFinish As XY)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                      OpenMode.ForWrite)

            '' Create a line that starts at 5,5 and ends at 12,3
            Dim acLine As Line = New Line(New Point3d(xyStart.X + xyMGlvInsertion.X, xyStart.y + xyMGlvInsertion.y, 0),
                                    New Point3d(xyFinish.X + xyMGlvInsertion.X, xyFinish.y + xyMGlvInsertion.y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
            End If
            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            g_idLastZipper = acLine.ObjectId
            Dim acRegAppTbl As RegAppTable
            acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
            Dim acRegAppTblRec As RegAppTableRecord
            If acRegAppTbl.Has("ID") = False Then
                acRegAppTblRec = New RegAppTableRecord
                acRegAppTblRec.Name = "ID"
                acRegAppTbl.UpgradeOpen()
                acRegAppTbl.Add(acRegAppTblRec)
                acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
            End If
            Using rb As New ResultBuffer
                rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"))
                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, g_sFileNo + "Left"))
                acLine.XData = rb
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
        Dim nWidth As Object
        nWidth = nHeight * g_nCurrTextAspect
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X + xyMGlvInsertion.X, xyInsert.y + xyMGlvInsertion.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_CalcMidPoint(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef xyMid As XY)

        Static aAngle, nLength As Double
        aAngle = FN_CalcAngle(xyStart, xyEnd)
        nLength = FN_CalcLength(xyStart, xyEnd)

        If nLength = 0 Then
            xyMid = xyStart 'Avoid overflow
        Else
            PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
        End If
    End Sub

    'Draws shoulder flaps
    Sub PR_DrawShoulderFlaps(ByRef xyS As XY, ByRef xyE As XY, ByRef strFlaps As String, ByRef strCustFlapLength As String,
                             ByRef strStrapLength As String, ByRef strFrontStrapLength As String, ByRef strWaistCir As String)
        Static ShoulderFlap, RaglanFlap As curve
        Static LastTapeValue As Object
        Static xyPt1, xyPt2 As XY
        Static xyTopFlap4, xyTopFlap2, xyTopFlap1, xyTopFlap3, xyTopFlap5 As XY
        Static xyTopFlap9, xyTopFlap7, xyTopFlap6, xyTopFlap8, xyTopFlap10 As XY
        Static xyBotFlap4, xyBotFlap2, xyBotFlap1, xyBotFlap3, xyBotFlap5 As XY
        Static xyFlap4, xyFlap2, xyFlap1, xyFlap3, xyFlap5 As XY
        Static Alpha, Theta, Phi, BETA, Omega, Delta, BottomRadius, TopArcIncrement As Object
        Static TopRadius, LittleBit, opp, h1, x1, y1, adj, HYP, Change, FlapMarkerAngle As Object
        Static xyFlapText1 As XY
        Static xyFlapText2, xyCentre, xyMid, xyFlapMarker, xyStrapText As XY
        Static FlapLength, FlapMarkerLength As Object
        Static TopArcAngle, TopArcRadius, TopArcLength, ArcLength, CircleCircum, BottomArcLength, BottomArcRadius, BottomArcAngle As Object
        Static xfablen, TempMarker, xfabby, xfabdist As Object
        Static xyRaglan6, xyRaglan4, xyRaglan2, xyRaglan1, xyRaglan3, xyRaglan5, xyRaglan7 As XY
        Static TopRagLanAngle, RaglanBottom, RagAng2, Rag4, Rag2, Rag1, Rag3, RagAng1, RagAng3, RaglanTip, InnerRaglanTip As Object
        Static nTemplateAngle, nNotchOffset, nTemplateRadius, nNotchToTangent As Double
        Static ii As Short
        Static RagAng5, PhiExtra, RagAng4, aMarker As Double
        Static Strap As String
        Static aAngle, nLength As Double
        Static xyStrt, xyMidPoint, xyTmp As XY

        PR_CalcMidPoint(xyS, xyE, xyMidPoint)
        nLength = FN_CalcLength(xyS, xyE)
        aAngle = -90
        BottomArcAngle = 47.5
        'All this crap because you can't use ByVal on user defined types (Bastards!)
        'VB3 comment
        If g_sSide = "Right" Then
            xyStrt = xyE
            aMarker = 135
        Else
            xyStrt = xyS
            aMarker = 45
        End If

        ' Check for Custom Flap Length
        If Val(strCustFlapLength) > 0 Then
            LastTapeValue = MANGLOVE1.fnDisplayToInches(CDbl(strCustFlapLength))
        Else
            'Standard flap length
            ''-------------------LastTapeValue = g_nCir(g_iEOSPointer)
            LastTapeValue = g_nCir(1, g_iEOSPointer)
            LastTapeValue = (LastTapeValue / 3.14) * 0.92
        End If

        ' Calculate Bottom part of curve with 5 Points
        If nLength >= 5.5 Then
            BottomRadius = LastTapeValue - 0.625
            RaglanBottom = 3.5
            TopRagLanAngle = 17
            RaglanTip = 1.9375
            InnerRaglanTip = 2.173
            nNotchOffset = 0.875 '14/16 ths Raglan Template s287
        ElseIf (nLength < 5.5) And (nLength > 3) Then
            BottomRadius = 2.8
            RaglanBottom = 2.5
            TopRagLanAngle = 17
            RaglanTip = 1.05
            InnerRaglanTip = 1.22
            nNotchOffset = 0.75 '12/16 ths Raglan Template s289
        ElseIf (nLength <= 3) Then
            BottomRadius = 2
            RaglanBottom = 2.0625
            RaglanTip = 0.625
            TopRagLanAngle = 15
            InnerRaglanTip = 0.727
            nNotchOffset = 0.6875 '11/16 ths, Raglan Template s290
        End If

        nTemplateRadius = 3.5
        nTemplateAngle = 40
        nNotchToTangent = 2.125

        'Calc length of bottom Arc
        BottomArcRadius = BottomRadius
        CircleCircum = PI * (2 * (BottomArcRadius))
        BottomArcLength = (BottomArcAngle / 360) * CircleCircum

        ''----------To avoid runtime error---------
        ShoulderFlap.n = 10
        Dim X(10), Y(10) As Double
        ShoulderFlap.X = X
        ShoulderFlap.y = Y
        ''------------------

        ShoulderFlap.X(1) = xyStrt.X + LastTapeValue
        ShoulderFlap.y(1) = 0

        For ii = 1 To 4
            Theta = (ii * 11.875 * PI) / 180
            adj = System.Math.Cos(Theta) * BottomRadius
            opp = System.Math.Sin(Theta) * BottomRadius
            LittleBit = BottomRadius - adj
            ShoulderFlap.X(ii + 1) = xyStrt.X + LastTapeValue - LittleBit
            ShoulderFlap.y(ii + 1) = 0 + opp
        Next ii
        PR_MakeXY(xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.y(5))

        'Calculate 6 angles and points for top part of curve
        PR_MakeXY(xyTopFlap1, xyStrt.X, xyStrt.y + nLength)
        ShoulderFlap.X(10) = xyTopFlap1.X
        ShoulderFlap.y(10) = xyTopFlap1.y

        'midpoint
        h1 = FN_CalcLength(xyTopFlap1, xyBotFlap5)
        h1 = h1 / 2
        Omega = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
        Omega = 180 - Omega
        Omega = (Omega * PI) / 180
        opp = System.Math.Abs(h1 * System.Math.Sin(Omega))
        adj = System.Math.Abs(h1 * System.Math.Cos(Omega))
        PR_MakeXY(xyMid, xyBotFlap5.X - adj, xyBotFlap5.y + opp)

        'Centre of Top Arc
        Delta = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
        Delta = Delta - 90
        Delta = (Delta * PI) / 180
        HYP = 4 * h1
        opp = System.Math.Abs(HYP * System.Math.Sin(Delta))
        adj = System.Math.Abs(HYP * System.Math.Cos(Delta))
        PR_MakeXY(xyCentre, xyMid.X + adj, xyMid.y + opp)
        TopRadius = FN_CalcLength(xyCentre, xyTopFlap1)
        TopArcRadius = TopRadius
        Phi = FN_CalcAngle(xyCentre, xyTopFlap1)
        Phi = Phi - 180
        Alpha = FN_CalcAngle(xyCentre, xyBotFlap5)
        Alpha = Alpha - 180
        TopArcAngle = Alpha - Phi
        TopArcIncrement = (Alpha - Phi) / 5

        'Calc Flap Length for position of marker
        CircleCircum = PI * (2 * (TopArcRadius))
        TopArcLength = (TopArcAngle / 360) * CircleCircum
        FlapLength = TopArcLength + BottomArcLength
        FlapMarkerLength = FlapLength / 2.7
        If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
        If FlapMarkerLength < TopArcLength Then
            FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
            FlapMarkerAngle = FlapMarkerAngle + Phi
            FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
            opp = System.Math.Abs(TopRadius * System.Math.Sin(FlapMarkerAngle))
            adj = System.Math.Abs(TopRadius * System.Math.Cos(FlapMarkerAngle))
            PR_MakeXY(xyFlapMarker, xyCentre.X - adj, xyCentre.y - opp)
        ElseIf FlapMarkerLength > TopArcLength Then
            TempMarker = FlapMarkerLength - TopArcLength
            TempMarker = BottomArcLength - TempMarker
            FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
            FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
            opp = System.Math.Abs(BottomArcRadius * System.Math.Sin(FlapMarkerAngle))
            adj = System.Math.Abs(BottomArcRadius * System.Math.Cos(FlapMarkerAngle))
            LittleBit = BottomRadius - adj
            PR_MakeXY(xyFlapMarker, xyStrt.X + LastTapeValue - LittleBit, 0 + opp)
        ElseIf FlapMarkerLength = TopArcLength Then
            PR_MakeXY(xyFlapMarker, xyBotFlap5.X, xyBotFlap5.y)
        End If
        ARMDIA1.PR_SetLayer("Template" & g_sSide)

        For ii = 1 To 4
            PhiExtra = Phi + (ii * TopArcIncrement)
            PhiExtra = (PhiExtra * PI) / 180
            opp = System.Math.Abs(TopRadius * System.Math.Sin(PhiExtra))
            adj = System.Math.Abs(TopRadius * System.Math.Cos(PhiExtra))
            ShoulderFlap.X(10 - ii) = xyCentre.X - adj
            ShoulderFlap.y(10 - ii) = xyCentre.y - opp
        Next ii

        If Left(strFlaps, 6) = "Raglan" Then
            'Copy standard shoulder from profile down
            'only go one vertex past the FlapMarker
            RaglanFlap.n = 0
            ''----------To avoid runtime error---------
            Dim XX(10), YY(10) As Double
            XX(1) = 0
            YY(1) = 0
            RaglanFlap.X = XX
            RaglanFlap.y = YY
            ''-----------------------------
            For ii = 10 To 1 Step -1
                RaglanFlap.n = RaglanFlap.n + 1
                RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
                RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
                If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
            Next ii
            PR_RotateCurve(xyStrt, aAngle, RaglanFlap)
            If g_sSide = "Left" Then PR_MirrorCurveInYaxis(xyMidPoint.X, 0, RaglanFlap)
            PR_DrawFitted(RaglanFlap)
        Else
            ShoulderFlap.n = 10
            PR_RotateCurve(xyStrt, aAngle, ShoulderFlap)
            If g_sSide = "Left" Then PR_MirrorCurveInYaxis(xyMidPoint.X, 0, ShoulderFlap)
            PR_DrawFitted(ShoulderFlap)
        End If

        'Draw Raglan
        If Left(strFlaps, 6) = "Raglan" Then
            Rag1 = System.Math.Sqrt((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
            Rag1 = BottomRadius - Rag1
            PR_MakeXY(xyRaglan1, xyStrt.X + LastTapeValue - Rag1, nNotchOffset)
            PR_MakeXY(xyRaglan2, xyRaglan1.X - nNotchToTangent, 0)
            RagAng1 = FN_CalcAngle(xyRaglan2, xyRaglan1)
            Rag2 = FN_CalcLength(xyRaglan2, xyRaglan1)
            Rag2 = Rag2 / 2
            RagAng1 = (RagAng1 * PI) / 180
            opp = System.Math.Abs(Rag2 * System.Math.Sin(RagAng1))
            adj = System.Math.Abs(Rag2 * System.Math.Cos(RagAng1))
            PR_MakeXY(xyRaglan3, xyRaglan2.X + adj, xyRaglan2.y + opp)

            RagAng1 = FN_CalcAngle(xyRaglan3, xyRaglan1)
            RagAng1 = 90 - RagAng1
            RagAng1 = (RagAng1 * PI) / 180
            Rag3 = FN_CalcLength(xyRaglan3, xyRaglan1)
            Rag4 = System.Math.Sqrt((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
            opp = System.Math.Abs(Rag4 * System.Math.Sin(RagAng1))
            adj = System.Math.Abs(Rag4 * System.Math.Cos(RagAng1))
            PR_MakeXY(xyRaglan4, xyRaglan3.X - adj, xyRaglan3.y + opp)

            RagAng2 = (nTemplateAngle * PI) / 180
            opp = System.Math.Abs(RaglanBottom * System.Math.Sin(RagAng2))
            adj = System.Math.Abs(RaglanBottom * System.Math.Cos(RagAng2))
            PR_MakeXY(xyRaglan5, xyRaglan1.X + adj, xyRaglan1.y + opp)
            RagAng3 = FN_CalcAngle(xyRaglan5, xyFlapMarker)

            If RagAng3 > 180 Then
                RagAng3 = RagAng3 - 180
                RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
            ElseIf RagAng3 <= 180 Then
                RagAng3 = RagAng3 - 90
                RagAng4 = 90 - RagAng3
                RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
            End If

            RagAng5 = RagAng4 - 13
            RagAng4 = (RagAng4 * PI) / 180
            opp = System.Math.Abs(RaglanTip * System.Math.Sin(RagAng4))
            adj = System.Math.Abs(RaglanTip * System.Math.Cos(RagAng4))
            PR_MakeXY(xyRaglan6, xyRaglan5.X - adj, xyRaglan5.y + opp)

            'Draw Front line of tip
            RagAng5 = (RagAng5 * PI) / 180
            opp = System.Math.Abs(InnerRaglanTip * System.Math.Sin(RagAng5))
            adj = System.Math.Abs(InnerRaglanTip * System.Math.Cos(RagAng5))
            PR_MakeXY(xyRaglan7, xyRaglan5.X - adj, xyRaglan5.y + opp)

            PR_RotatePoint(xyStrt, aAngle, xyRaglan1)
            PR_RotatePoint(xyStrt, aAngle, xyRaglan2)
            PR_RotatePoint(xyStrt, aAngle, xyRaglan4)
            PR_RotatePoint(xyStrt, aAngle, xyRaglan5)
            PR_RotatePoint(xyStrt, aAngle, xyRaglan6)
            PR_RotatePoint(xyStrt, aAngle, xyRaglan7)

            xyTmp = xyFlapMarker
            PR_RotatePoint(xyStrt, aAngle, xyTmp)

            If g_sSide = "Left" Then
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan1)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan2)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan4)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan5)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan6)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyRaglan7)
                PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyTmp)
            End If

            PR_DrawArc(xyRaglan4, xyRaglan1, xyRaglan2)
            PR_DrawLine(xyRaglan1, xyRaglan5)

            PR_DrawLine(xyRaglan5, xyRaglan6)
            PR_DrawLine(xyRaglan6, xyTmp)
            ARMDIA1.PR_SetLayer("Notes")
            PR_DrawLine(xyRaglan5, xyRaglan7)

        End If

        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
        PR_MakeXY(xyStrapText, xyMidPoint.X, xyMidPoint.y - (2 * EIGHTH))

        If Val(strStrapLength) <> 0 Then
            ''Strap = "STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strStrapLength)))) & "\" & QQ
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Strap = "STRAP = " & Trim(strStrapLength)
            Else
                Strap = "STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strStrapLength)))) & Chr(34)
            End If
        Else
            ''Strap = "STRAP = 24\" & QQ
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Dim strVal As String = "24"
                Dim nCMVal As Double = 24 * 2.54
                nCMVal = System.Math.Abs(nCMVal)
                Dim iInt As Short = Int(nCMVal)
                If iInt <> 0 Then
                    Dim nDec As Double = (nCMVal - iInt) * 10
                    nDec = ARMDIA1.round(nDec) / 10
                    nCMVal = iInt + nDec
                    strVal = Str(nCMVal)
                Else
                    strVal = Str(1)
                End If
                Strap = "STRAP = " & strVal
            Else
                Strap = "STRAP = 24" & Chr(34)
            End If
        End If
        If Val(strFrontStrapLength) > 0 Then
            'Display FRONT strap length if given
            ''Strap = "BACK  STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strStrapLength)))) & "\" & QQ
            ''Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strFrontStrapLength)))) & "\" & QQ
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Strap = "BACK  STRAP = " & Trim(strStrapLength)
                Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(strFrontStrapLength)
            Else
                Strap = "BACK  STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strStrapLength)))) & Chr(34)
                Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strFrontStrapLength)))) & Chr(34)
            End If
        End If
        If Val(strWaistCir) > 0 Then
            'Display Waist Circumference for D-Style flaps only
            ''Strap = Strap & Chr(10) & "WAIST = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strWaistCir)))) & "\" & QQ
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Strap = Strap & Chr(10) & "WAIST = " & Trim(strWaistCir)
            Else
                Strap = Strap & Chr(10) & "WAIST = " & Trim(MANGLOVE1.fnInchestoText(MANGLOVE1.fnDisplayToInches(Val(strWaistCir)))) & Chr(34)
            End If
        End If

        ''-----------PR_DrawText(Strap, xyStrapText, 0.1, 0)
        PR_DrawMText(Strap, xyStrapText, 0, True)

        xfabby = strFlaps
        xfablen = Len(xfabby)
        xfabdist = (xfablen * 0.075) + 0.1
        PR_MakeXY(xyFlapText1, xyFlapMarker.X - (xfabdist * 0.75), xyFlapMarker.y - 0.5)
        PR_RotatePoint(xyStrt, aAngle, xyFlapText1)
        If g_sSide = "Left" Then PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyFlapText1)
        ''-----------PR_DrawText(strFlaps, xyFlapText1, 0.1, 0)
        PR_DrawMText(strFlaps, xyFlapText1, 0, True)

        'Flap Marker
        PR_RotatePoint(xyStrt, aAngle, xyFlapMarker)
        If g_sSide = "Left" Then PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyFlapMarker)
        ''PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & "closed arrow" & QC & "xyStart.x +" & Str(xyFlapMarker.X) & CC & "xyStart.y +" & Str(xyFlapMarker.y) & ",0.25,0.1," & aMarker & ");")
        Dim xyArrow As XY
        PR_MakeXY(xyArrow, xyMGlvInsertion.X + xyFlapMarker.X, xyMGlvInsertion.y + xyFlapMarker.y)
        PR_DrawManGlvClosedArrow(xyArrow, 45)
        ''PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
        PR_AddXDataValToLast("ID", g_sFileNo + "Left")

        'draw the fold line
        ARMDIA1.PR_SetLayer("Template" & g_sSide)
        PR_MakeXY(xyPt1, xyStrt.X, 0)
        PR_MakeXY(xyPt2, xyStrt.X + LastTapeValue, 0)
        PR_RotatePoint(xyStrt, aAngle, xyPt1)
        PR_RotatePoint(xyStrt, aAngle, xyPt2)

        If g_sSide = "Left" Then
            PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyPt1)
            PR_MirrorPointInYaxis(xyMidPoint.X, 0, xyPt2)
        End If

        If Left(strFlaps, 6) = "Raglan" Then
            PR_DrawLine(xyPt1, xyRaglan2)
        Else
            PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        End If
    End Sub
    Sub PR_DrawManGlvClosedArrow(ByRef xyPoint As MANGLOVE1.XY, ByRef nAngle As Object)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim xyStart As MANGLOVE1.XY
        ''PR_MakeXY(xyStart, xyPoint.X, xyPoint.Y + 0.125)
        xyStart = xyPoint
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Dim acPoly As Polyline = New Polyline()
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.y), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.y + 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.y - 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.y), 0, 0, 0)
            acPoly.TransformBy(Matrix3d.Rotation((nAngle * (BODYSUIT1.PI / 180)), Vector3d.ZAxis, New Point3d(xyStart.X, xyStart.y, 0)))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            Dim idPolyline As ObjectId = New ObjectId
            idPolyline = acBlkTblRec.AppendEntity(acPoly)
            g_idLastZipper = acPoly.ObjectId
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
            ''Create Hatch entity
            Dim ObjIds As ObjectIdCollection = New ObjectIdCollection
            ObjIds.Add(idPolyline)
            Dim oHatch As Hatch = New Hatch()
            Dim normal As Vector3d = New Vector3d(0.0, 0.0, 1.0)
            oHatch.Normal = normal
            oHatch.Elevation = 0.0
            oHatch.PatternScale = 2.0
            oHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
            oHatch.ColorIndex = 1
            acBlkTblRec.AppendEntity(oHatch)
            acTrans.AddNewlyCreatedDBObject(oHatch, True)
            oHatch.Associative = True
            oHatch.AppendLoop(CInt(HatchLoopTypes.Default), ObjIds)
            oHatch.Color = acPoly.Color
            oHatch.EvaluateHatch(True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_RotateCurve(ByRef xyRotate As XY, ByRef aRotation As Double, ByRef Profile As curve)
        Static ii As Short
        Static xyVertex As XY
        For ii = 1 To Profile.n
            PR_MakeXY(xyVertex, Profile.X(ii), Profile.y(ii))
            PR_RotatePoint(xyRotate, aRotation, xyVertex)
            Profile.X(ii) = xyVertex.X
            Profile.y(ii) = xyVertex.y
        Next ii
    End Sub
    Sub PR_RotatePoint(ByRef xyRotate As XY, ByRef aRotation As Double, ByRef xyPoint As XY)
        Static aAngle, nLength As Double

        aAngle = (FN_CalcAngle(xyRotate, xyPoint) + aRotation) Mod 360
        nLength = FN_CalcLength(xyRotate, xyPoint)
        PR_CalcPolar(xyRotate, aAngle, nLength, xyPoint)
    End Sub
    Sub PR_DrawMarkerNamed(ByRef sName As String, ByRef xyPoint As XY, ByRef nWidth As Object, ByRef nHeight As Object, ByRef nAngle As Object)
        'Draw a Marker at the given point
        PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & sName & QC & "xyStart.x+" & xyPoint.X & CC & "xyStart.y+" & xyPoint.y & CC & nWidth & CC & nHeight & CC & nAngle & ");")
        PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyEnd, xyBase, xySecEnd As XY
        PR_MakeXY(xyEnd, xyBase.X + (nWidth / 2), xyBase.y + (nHeight / 2))
        PR_MakeXY(xySecEnd, xyBase.X + (nWidth / 2), xyBase.y - (nHeight / 2))

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sName) Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = sName
                Dim acLine As Line = New Line(New Point3d(xyBase.X, xyBase.y, 0), New Point3d(xyEnd.X, xyEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xyBase.X, xyBase.y, 0), New Point3d(xySecEnd.X, xySecEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl(sName)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyMGlvInsertion.X, xyPoint.y + xyMGlvInsertion.y, 0), blkRecId)
                blkRef.Rotation = nAngle * (PI / 180)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has("ID") = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = "ID"
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, g_sFileNo + "Left"))
                    blkRef.XData = rb
                End Using
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawGloveInsertU(ByRef xyPoint As XY)
        Dim xyFirst, xySec, xyEnd, xySecEnd, xyBase, xyPt1, xyPt2, xyPt3, xyPt4, xyCen As XY
        PR_MakeXY(xyPt1, xyBase.X - 0.125, xyBase.y)
        PR_MakeXY(xyPt2, xyBase.X - 0.125, xyBase.y - 0.1)
        PR_MakeXY(xyPt3, xyBase.X + 0.125, xyBase.y)
        PR_MakeXY(xyPt4, xyBase.X + 0.125, xyBase.y - 0.1)
        PR_MakeXY(xyFirst, xyBase.X - 0.125, xyBase.y)
        ''PR_MakeXY(xyEnd, xyFirst.X + 0.05, xyFirst.y + 0.0625)
        PR_MakeXY(xyEnd, xyBase.X + 0.05, xyBase.y + 0.0625)
        PR_MakeXY(xySec, xyBase.X + 0.125, xyBase.y)
        ''PR_MakeXY(xySecEnd, xySec.X + 0.05, xySec.y - 0.0625)
        PR_MakeXY(xySecEnd, xyBase.X + 0.05, xyBase.y - 0.0625)

        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
        PR_MakeXY(xyCen, xyBase.X, xyBase.y - 0.1)
        nRad = FN_CalcLength(xyCen, xyPt2)
        nStartAng = FN_CalcAngle(xyCen, xyPt2) * (PI / 180)
        nEndAng = FN_CalcAngle(xyCen, xyPt4) * (PI / 180)
        nDeltaAng = nEndAng - nStartAng
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("GloveInsertU") Then
                Dim blkTblRecGloveU As BlockTableRecord = New BlockTableRecord()
                blkTblRecGloveU.Name = "GloveInsertU"
                Dim acLine As Line = New Line(New Point3d(xyPt1.X, xyPt1.y, 0), New Point3d(xyPt2.X, xyPt2.y, 0))
                blkTblRecGloveU.AppendEntity(acLine)
                acLine = New Line(New Point3d(xyPt3.X, xyPt3.y, 0), New Point3d(xyPt4.X, xyPt4.y, 0))
                blkTblRecGloveU.AppendEntity(acLine)
                Dim acArc As Arc = New Arc(New Point3d(xyCen.X, xyCen.y, 0),
                                     nRad, nStartAng, nEndAng)
                blkTblRecGloveU.AppendEntity(acArc)
                Dim blkRecIdArw As ObjectId = ObjectId.Null
                If Not acBlkTbl.Has("Open Short Arrow") Then
                    Dim blkTblRecOpenArw As BlockTableRecord = New BlockTableRecord()
                    blkTblRecOpenArw.Name = "Open Short Arrow"
                    acLine = New Line(New Point3d(xyBase.X, xyBase.y, 0), New Point3d(xyEnd.X, xyEnd.y, 0))
                    blkTblRecOpenArw.AppendEntity(acLine)
                    acLine = New Line(New Point3d(xyBase.X, xyBase.y, 0), New Point3d(xySecEnd.X, xySecEnd.y, 0))
                    blkTblRecOpenArw.AppendEntity(acLine)
                    acBlkTbl.UpgradeOpen()
                    acBlkTbl.Add(blkTblRecOpenArw)
                    acTrans.AddNewlyCreatedDBObject(blkTblRecOpenArw, True)
                    blkRecIdArw = blkTblRecOpenArw.Id
                Else
                    blkRecIdArw = acBlkTbl("Open Short Arrow")
                End If
                If blkRecIdArw <> ObjectId.Null Then
                    'Create new block reference 
                    Dim blkRefOpenArw1 As BlockReference = New BlockReference(New Point3d(xyFirst.X, xyFirst.y, 0), blkRecIdArw)
                    blkRefOpenArw1.Rotation = 270 * (PI / 180)
                    Dim blkRefOpenArw2 As BlockReference = New BlockReference(New Point3d(xySec.X, xySec.y, 0), blkRecIdArw)
                    blkRefOpenArw2.Rotation = 270 * (PI / 180)
                    blkTblRecGloveU.AppendEntity(blkRefOpenArw1)
                    blkTblRecGloveU.AppendEntity(blkRefOpenArw2)
                End If
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecGloveU)
                acTrans.AddNewlyCreatedDBObject(blkTblRecGloveU, True)
                blkRecId = blkTblRecGloveU.Id
            Else
                blkRecId = acBlkTbl("GloveInsertU")
            End If
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyMGlvInsertion.X, xyPoint.y + xyMGlvInsertion.y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has("ID") = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = "ID"
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, g_sFileNo + "Left"))
                    blkRef.XData = rb
                End Using
            End If

            acTrans.Commit()
        End Using
    End Sub
    Sub PR_SetLineStyle(ByRef sStyle As String)
        PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & sStyle & QQ & "));")
    End Sub
    Sub PR_DrawMarker(ByRef xyPoint As XY)
        'Draw a Marker at the given point
        ''----------PrintLine(fNum, "hEnt = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart.x+" & xyPoint.X & CC & "xyStart.y+" & xyPoint.y & CC & "0.125);")
        ''------------PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As XY
        PR_CalcPolar(xyBase, 135, 0.0625, xyStart)
        PR_CalcPolar(xyBase, -45, 0.0625, xyEnd)
        PR_CalcPolar(xyBase, 45, 0.0625, xySecSt)
        PR_CalcPolar(xyBase, -135, 0.0625, xySecEnd)

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("X Marker") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "X Marker"
                Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.y, 0), New Point3d(xyEnd.X, xyEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xySecSt.X, xySecSt.y, 0), New Point3d(xySecEnd.X, xySecEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl("X Marker")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyMGlvInsertion.X, xyPoint.y + xyMGlvInsertion.y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                g_idLastZipper = blkRef.ObjectId
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has("ID") = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = "ID"
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, g_sFileNo + "Left"))
                    blkRef.XData = rb
                End Using
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As XY, ByRef Angle As Double, Optional ByRef bIsCenter As Boolean = False)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            Dim mtx As New MText()
            mtx.Location = New Point3d(xyInsert.X + xyMGlvInsertion.X, xyInsert.y + xyMGlvInsertion.y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = Angle
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            If bIsCenter = True Then
                mtx.Attachment = AttachmentPoint.TopCenter
            End If
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    'To Get Insertion point from user
    Sub PR_GetInsertionPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Indicate Start Point "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(xyMGlvInsertion, ptStart.X, ptStart.Y)
    End Sub
    Sub PR_DrawGlvTapeNotes(ByRef xyPoint As XY, ByRef sData As String, ByRef sMM As String, ByRef sGrams As String, ByRef sReduction As String,
                            ByRef sTapeLen As String, ByRef bIsTapeLen2 As Boolean, ByRef sTapeLen2 As String)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim sName As String = "GLVTAPENOTES"
        Dim xyBase, xyCenter, xyData, xyMM, xyGrams, xyTapeLen, xyTapeLen2 As XY
        PR_MakeXY(xyBase, 0, 0)
        PR_MakeXY(xyCenter, xyBase.X + 0.75 + 0.125 + 0.1, xyBase.y + 0.135)
        PR_MakeXY(xyData, xyBase.X - 0.25, xyBase.y - 0.15)
        PR_MakeXY(xyMM, xyBase.X - 0.275, xyBase.y + 0.1)
        PR_MakeXY(xyGrams, xyBase.X + 0.275, xyBase.y + 0.1)
        PR_MakeXY(xyTapeLen, xyBase.X - 0.6, xyBase.y + 0.1)
        If sTapeLen.Length > 1 Then
            PR_MakeXY(xyTapeLen, xyBase.X - 0.65, xyBase.y + 0.1)
        End If
        PR_MakeXY(xyTapeLen2, xyBase.X - 0.4994, xyBase.y + 0.1655)

        Dim ptCollAtt As Point3dCollection = New Point3dCollection
        ptCollAtt.Add(New Point3d(xyData.X, xyData.y, 0))
        ptCollAtt.Add(New Point3d(xyMM.X, xyMM.y, 0))
        ptCollAtt.Add(New Point3d(xyGrams.X, xyGrams.y, 0))
        ptCollAtt.Add(New Point3d(xyCenter.X, xyCenter.y, 0))
        ptCollAtt.Add(New Point3d(xyTapeLen.X, xyTapeLen.y, 0))

        Dim strTag(5), strTextString(5) As String
        strTag(1) = "Data"
        strTag(2) = "MM"
        strTag(3) = "Grams"
        strTag(4) = "Reduction"
        strTag(5) = "TapeLengths"
        strTextString(1) = sData
        strTextString(2) = sMM
        strTextString(3) = sGrams
        strTextString(4) = sReduction
        strTextString(5) = sTapeLen
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sName) Then
                Dim blkTblRecTape As BlockTableRecord = New BlockTableRecord()
                blkTblRecTape.Name = sName
                Dim acCircle As Circle = New Circle()
                acCircle.Center = New Point3d(xyCenter.X, xyCenter.y, 0)
                acCircle.Radius = 0.125
                blkTblRecTape.AppendEntity(acCircle)

                Dim acAttDef As New AttributeDefinition
                Dim ii As Double
                For ii = 1 To 5
                    acAttDef = New AttributeDefinition
                    ''acAttDef.Position = New Point3d(xyData.X, xyData.y, 0)
                    acAttDef.Position = ptCollAtt(ii - 1)
                    acAttDef.Prompt = strTag(ii)
                    acAttDef.Tag = strTag(ii)
                    acAttDef.TextString = strTextString(ii)
                    acAttDef.Height = 0.125
                    acAttDef.Justify = AttachmentPoint.BaseLeft
                    If ii = 4 Then
                        acAttDef.Justify = AttachmentPoint.MiddleCenter
                        acAttDef.AlignmentPoint = ptCollAtt(ii - 1)
                    End If
                    acAttDef.Invisible = False
                    blkTblRecTape.AppendEntity(acAttDef)
                Next
                If bIsTapeLen2 Then
                    acAttDef = New AttributeDefinition
                    acAttDef.Position = New Point3d(xyTapeLen2.X, xyTapeLen2.y, 0)
                    acAttDef.Prompt = "TapeLengths2"
                    acAttDef.Tag = "TapeLengths2"
                    acAttDef.TextString = sTapeLen2
                    acAttDef.Height = 0.125
                    acAttDef.Justify = AttachmentPoint.BaseLeft
                    acAttDef.Invisible = False
                    blkTblRecTape.AppendEntity(acAttDef)
                End If

                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecTape)
                acTrans.AddNewlyCreatedDBObject(blkTblRecTape, True)
                blkRecId = blkTblRecTape.Id
            Else
                blkRecId = acBlkTbl(sName)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X + xyMGlvInsertion.X, xyPoint.y + xyMGlvInsertion.y, 0), blkRecId)
                ''blkRef.Rotation = nAngle * (PI / 180)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyMGlvInsertion.X, xyMGlvInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                g_idLastZipper = blkRef.ObjectId
                '' Open the Block table record Model space for write
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If Not acAttDef.Constant Then
                            Dim acAttRef As New AttributeReference
                            acAttRef.SetAttributeFromBlock(acAttDef, blkRef.BlockTransform)
                            If acAttRef.Tag.ToUpper().Equals("Data", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sData
                            ElseIf acAttRef.Tag.ToUpper().Equals("MM", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sMM
                            ElseIf acAttRef.Tag.ToUpper().Equals("Grams", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sGrams
                            ElseIf acAttRef.Tag.ToUpper().Equals("Reduction", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sReduction
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengths", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sTapeLen
                            ElseIf acAttRef.Tag.ToUpper().Equals("TapeLengths2", StringComparison.InvariantCultureIgnoreCase) Then
                                If bIsTapeLen2 = False Then
                                    Continue For
                                End If
                                acAttRef.TextString = sTapeLen2
                            End If
                            blkRef.AttributeCollection.AppendAttribute(acAttRef)
                            acTrans.AddNewlyCreatedDBObject(acAttRef, True)
                        End If
                    End If
                Next
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
End Module
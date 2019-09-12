Attribute VB_Name = "MANGLOVE1"
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

    
   'PI as a Global Constant
    Global Const PI = 3.141592654

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

   'Store current layer and text setings to reduce DRAFIX code
   'this value is checked in PR_SetLayer
    Global g_sCurrentLayer As String
    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont
    Global g_nCurrTextAngle

    
    Global g_sFileNo As String
    Global g_sSide As String
    Global g_sPatient As String
    Global g_sID As String

    Global g_nUnitsFac As Double

   'Scale to match current CAD-GLOVE output
   'Global Const ISCALE = 1026

    Global g_sPathJOBST     As String

    Global MainForm As manglove

    Global g_iInsertValue(1 To 6)  As Double
    Global g_Finger(1 To 5)        As Finger 'Includes thumb
    Global g_nFingerPIPCir(1 To 5) As Double
    Global g_nFingerDIPCir(1 To 5) As Double
    Global g_nFingerLen(1 To 5)    As Double
    Global g_nWeb(1 To 4)          As Double
    Global g_OnFold                As Integer
    Global g_MissingFingers        As Integer
    Global g_Missing(1 To 5)       As Integer
    Global g_nPalm                 As Double
    Global g_nWrist                As Double
    Global g_iFINGER_CHART(1 To 5) As Integer
    Global g_nFINGER_FIGURE(1 To 5) As Double
    Global g_nPALM_FIGURE           As Double
    Global g_nWRIST_FIGURE          As Double
    Global g_iThumbStyle            As Integer
    Global g_iInsertStyle           As Integer
    Global g_iInsertSize            As Integer
    Global g_CalculatedExtension    As Integer
    Global g_FusedFingers           As Integer
    Global g_iThumbWebDrop          As Integer
    
    Global xyDatum             As XY
    Global xyLittle            As XY
    Global xyLFS               As XY
    Global xyRing              As XY
    Global xyIndex             As XY
    Global xyMiddle            As XY
    Global xyThumb             As XY
    Global xyPalm(1 To 6)      As XY
    Global xyPalmThumb(1 To 5) As XY


    Global g_iGms(1 To NOFF_ARMTAPES)  As Integer
    Global g_iMMs(1 To NOFF_ARMTAPES)  As Integer
    Global g_iRed(1 To NOFF_ARMTAPES)  As Integer
    Global g_nCir(1 To NOFF_ARMTAPES)  As Double
    Global g_nPleats(1 To 4)           As Double
    Global g_iFirstTape                As Integer
    Global g_iLastTape                 As Integer
    Global g_iWristPointer             As Integer
    Global g_iEOSPointer               As Integer
    Global g_iNumTotalTapes            As Integer
    Global g_iNumTapesWristToEOS       As Integer
    Global g_EOSType                   As Integer
    Global g_iPressure                 As Integer
    Global g_iInsertOtherGlv           As Integer
    Global g_OnFoldOtherGlv            As Integer
    Global g_DataIsCalcuable           As Integer
    Global g_PrintFold                 As Integer
    Global g_ExtendTo                  As Integer

   'Flaps
    Global g_nStrapLength          As Double
    Global g_nFrontStrapLength     As Double
    Global g_nCustFlapLength       As Double
    Global g_nWaistCir             As Double
    Global g_sFlapType             As String
    Global g_iFlapType             As Integer

   'Curves
    Global g_FingerThumbBlend As Curve
    Global g_ThumbTopCurve As BiArc
    Global UlnarProfile As Curve
    Global RadialProfile As Curve
    Global TapeNote(1 To NOFF_ARMTAPES)   As TapeData


    Global Const g_sDialogID = "MANUAL Glove Dialogue"
    

' MsgBox return values
' Global Const IDOK = 1                  ' OK button pressed
'Global Const IDCANCEL = 2              ' Cancel button pressed
'Global Const IDABORT = 3               ' Abort button pressed
'Global Const IDRETRY = 4               ' Retry button pressed
'Global Const IDIGNORE = 5              ' Ignore button pressed
'Global Const IDYES = 6                 ' Yes button pressed
'Global Const IDNO = 7                  ' No button pressed

    Option Explicit

Function FN_InchesValue(TextBox As Control) As Double
' Check for numeric values
    Dim sChar As String, sText As String
    Dim iLen As Integer, nn As Integer
    Dim nLen As Double

    sText = TextBox.Text
    iLen = Len(sText)

   'Check the actual structure of the input
    FN_InchesValue = -1
    For nn = 1 To iLen
        sChar = Mid$(sText, nn, 1)
        If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
            MsgBox "Invalid - Dimension has been entered", 48, g_sDialogID
            If MainForm.Visible Then TextBox.SetFocus
            FN_InchesValue = -1
            Exit Function
        End If
    Next nn

   'Convert to inches
    nLen = fnDisplayToInches(Val(TextBox.Text))
    If nLen = -1 Then
        MsgBox "Invalid - Length has been entered", 48, g_sDialogID
        If MainForm.Visible Then TextBox.SetFocus
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

    Dim iInt As Integer, nDec As Double, iSign As Integer
   'retain sign
    iSign = Sgn(nDisplay)
    nDisplay = Abs(nDisplay)

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

Function fnInchestoText(nInches As Double) As String
   'Function returns a decimal value in inches as a string
   '
    Dim nPrecision As Double, nDec As Double
    Dim iInt As Integer, iEighths As Integer
    Dim sString As String
    nPrecision = 0.125

   'Split into decimal parts
    iInt = Int(nInches)
    nDec = nInches - iInt
    If nDec <> 0 Then           'Avoid overflow
        iEighths = Int(nDec / nPrecision)
    Else
        iEighths = 0
    End If

   'Format string
    If iInt <> 0 Then
        sString = LTrim$(Str$(iInt))
    Else
        sString = "  "
    End If
    If iEighths <> 0 Then
        Select Case iEighths
        Case 2, 6
            sString = sString + "-" + LTrim$(Str$(iEighths / 2)) + "/4"
        Case 4
            sString = sString + "-" + "1/2"
        Case Else
            sString = sString + "-" + LTrim$(Str$(iEighths)) + "/8"
        End Select
    Else
        sString = sString + "   "
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

    Dim iInt As Integer, iSign  As Integer
    Dim nPrecision As Double, nDec As Double

   'Return 0 if input is Zero
    If nNumber = 0 Then
       fnRoundInches = 0
       Exit Function
    End If

   'Set precision
    nPrecision = 0.125

   'Break input into components
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)
    nDec = nNumber - iInt

   'Get decimal part in precision units
    If nDec <> 0 Then
        nDec = nDec / nPrecision    'Avoid overflow
    End If
    nDec = round(nDec)

   'Return value
    fnRoundInches = (iInt + (nDec * nPrecision)) * iSign

End Function

Sub PR_Select_Text(Text_Box_Name As Control)
    If Not Text_Box_Name.Enabled Then Exit Sub
    Text_Box_Name.SetFocus
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub


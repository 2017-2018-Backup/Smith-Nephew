Attribute VB_Name = "WHDRAW1"

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum  As Integer 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

    Public MainForm As whdraw

    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont

    Global g_nUnitsFac  As Double
    Global g_sFileNo    As String
    Global g_sSide      As String
    Global g_sPatient    As String
    Global g_sCurrentLayer As String

    Global g_sPathJOBST As String
    Global g_sWorkOrder As String


    Global xyChapDatum      As xy
    Global xyOrigin         As xy
    Global xyOtemplate      As xy
    Global g_TopLegProfile  As Curve
    Global g_BotLegProfile  As Curve
    Global xyProfileDistal  As xy
    Global xyProfileProximal As xy
    Global g_nFabricClass   As Integer
    Global g_sPressure      As String
    Global g_nFirstTape     As Integer
    Global g_nLastTape      As Integer
    Global g_nFootPleat1    As Double
    Global g_nFootPleat2    As Double
    Global g_nTopLegPleat1  As Double
    Global g_nTopLegPleat2  As Double
    Global g_sTapeLength    As String
    Global g_Footless       As Integer
    Global g_nChapCirGiven  As Double
    Global g_nChapCirFigured As Double
    Global g_nChapLenGiven  As Double
    Global g_nChapLenFigured As Double
    Global g_nLegStyle      As Integer

    Global g_sDialogueID    As String


    Type TemplateScale
        nNo         As Integer
        sScale      As String
        nSpace      As Double
        n20Len      As Double
        nReduction  As Integer
    End Type

    Global TemplateChart(0 To 30)    As TemplateScale



    Option Explicit

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

    Dim iInt As Integer, nDec As Double, iSign As Integer
   'retain sign
    iSign = Sgn(nDisplay)
    nDisplay = Abs(nDisplay)

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


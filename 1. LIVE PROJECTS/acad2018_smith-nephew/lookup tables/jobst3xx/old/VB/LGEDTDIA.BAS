Attribute VB_Name = "LGEDTDIA1"
'   '* Windows API Functions Declarations
'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
   
   'Constants used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

    Public MainForm As lgedtdia

    Type xy
        X As Double
        y As Double
    End Type

   'Globals
    Global g_nUnitsFac      As Double
    Global g_iLtAnkle       As Integer
    Global g_JOBSTEX_FL     As Integer

   'Figuring globals
   'These are used by the sub PR_FigureTape to check
   'if re-figuring should occur because the current data
   'does not match the previous data
    Global g_nLtLastHeel    As Double
    Global g_nLtLastAnkle   As Double
    Global g_iLtLastMM      As Double
    Global g_iLtLastStretch As Integer
    Global g_iLtLastZipper  As Integer
    Global g_iLtLastFabric  As Integer

    Global g_iLtStretch(29) As Integer
    Global g_iLtRed(29)     As Integer
    Global g_iLtMM(29)      As Integer
    Global g_nLtLengths(29)  As Double
   
    Global g_iLtStretchInit(29) As Integer
    Global g_iLtRedInit(29)  As Integer
    Global g_iLtMMInit(29)  As Integer
    Global g_nLtLengthsInit(29) As Double
    Global g_iLtChanged(29) As Integer

    Global g_iFirstTape     As Integer
    Global g_iLastTape      As Integer

    Global g_sTextList      As String

    Global g_ReDrawn        As Integer

   'Profile Globals
    Global xyLegProfile(200)  As xy
    Global g_iLegProfile  As Integer
    Global xyOtemplate      As xy
    Global xyFold           As xy
    Global xyAnkle          As xy
    Global g_iFirstEditTape As Integer
    Global g_iLastEditTape  As Integer


   'Constants
    Const HEEL_TOL = 9 '9" heel

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum  As Integer 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont

    Global g_sFileNo    As String
    Global g_sSide      As String
    Global g_sPatient    As String
    Global g_sCurrentLayer As String

    Global g_sPathJOBST As String

   'Constants used in FN_JOBSTEX_Stretch and FN_JOBSTEX_Per
   'Source: Mary Ann Hettich
   '        Memo to Marian Bourke
   '        May 10 1994
   '
    Const A0 = -159.6675
    Const A1 = 7.300195
    Const A2 = 1.021744
    Const A3 = 200.9557
    Const A12 = -0.04440547
    Const A13 = 8.480964
    Const A23 = -1.311007


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

Function fnInchesToText(nInches As Double) As String
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

Sub PR_GetProfileFromFile(sFileName As String)
   'General procedure to read curve data from file
   'Data is a series of x, y values.
   'In the following format
   '
   '    Line    Type
   '-----------------------------------------------
   '    1       Template Origin (xyOtemplate)
   '    2       Ankle           (xyAnkle)
   '    3       Fold Position   (xyFold)
   '    4  }
   '    .  }}
   '    .  }}}  Vertices of profile (LegProfile)
   '    .  }}
   '    N  }
   '
   'It also establishes the relation between the first and
   'last tapes and the vertices
   'eg.
   '    g_iFirstTape = LegProfile(20) = g_iFirstEditTape
   '    g_iLastTape  = LegProfile(35) = g_iLastEditTape

    Dim ii As Integer
    Dim fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48
        Exit Sub
    End If

    Open sFileName For Input As fFileNum

   'Get control points
    Input #fFileNum, xyOtemplate.X, xyOtemplate.y
    Input #fFileNum, xyAnkle.X, xyAnkle.y
    Input #fFileNum, xyFold.X, xyFold.y

   'Get profile points
    g_iLegProfile = 0
    Do While Not EOF(fFileNum)
        Input #fFileNum, xyLegProfile(g_iLegProfile).X, xyLegProfile(g_iLegProfile).y
        g_iLegProfile = g_iLegProfile + 1
        Loop
    g_iLegProfile = g_iLegProfile - 1 '
    Close #fFileNum

   'Map First and Last tape to points in profile
    For ii = 0 To g_iLegProfile
         If xyLegProfile(ii).X > xyAnkle.X Then Exit For
    Next ii
    
   'Set value but check that it is futher away than an 1/8th
    g_iFirstEditTape = ii
    If Abs(xyLegProfile(ii).X - xyAnkle.X) < 0.125 Then g_iFirstEditTape = ii + 1

    For ii = g_iLegProfile To 0 Step -1
         If xyLegProfile(ii).X < xyFold.X Then Exit For
    Next ii
    g_iLastEditTape = ii
    If Abs(xyLegProfile(ii).X - xyFold.X) < 0.125 Then g_iLastEditTape = ii - 1

End Sub

Function round(ByVal nNumber As Double) As Integer
   'Fuction to return the rounded value of a decimal number
   'E.G.
   '    round(1.35)  = 1
   '    round(1.55)  = 2
   '    round(2.50)  = 3
   '    round(-2.50) = -3
   '    round(0)     = 0
   '

    Dim iInt As Integer, iSign As Integer
    
   'Avoid extra work. Return 0 if input is 0
    If nNumber = 0 Then
        round = 0
        Exit Function
    End If

   'Split input
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)

   'Effect rounding
    If (nNumber - iInt) >= 0.5 Then
        round = (iInt + 1) * iSign
    Else
        round = iInt * iSign
    End If
    
End Function

Sub Select_Text_in_Box(Text_Box_Name As Control)
    Text_Box_Name.SetFocus
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub


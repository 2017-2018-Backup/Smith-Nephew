Attribute VB_Name = "ARMEDDIA1"
   '* Windows API Functions Declarations
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

    Public MainForm As armeddia

    Type XY
        X As Double
        y As Double
    End Type

    Type curve
        n As Integer
        X(1 To 100) As Double
        y(1 To 100) As Double
    End Type

   'Globals
    Global g_nUnitsFac      As Double

    Global g_iGms(17)       As Integer
    Global g_iRed(17)       As Integer
    Global g_iMMs(17)       As Integer
    Global g_nLengths(17)   As Double
   
    Global g_iGmsInit(17)   As Integer
    Global g_iRedInit(17)   As Integer
    Global g_iMMsInit(17)   As Integer
    Global g_nLengthsInit(17) As Double
    Global g_iChanged(17)   As Integer
    Global g_iVertexMap(17) As Integer 'Maps length to profile vertex

    Global g_iFirstTape     As Integer
    Global g_iLastTape      As Integer

    Global g_ReDrawn        As Integer

   'Profile Globals
    Global xyProfile(200)   As XY
    Global xyOriginal(200)   As XY
    Global g_iProfile       As Integer
    Global xyOtemplate      As XY
    Global g_iFirstEditableVertex   As Integer
    Global g_iFirstEditTape As Integer
    Global g_iLastEditTape  As Integer
    Global g_ShortArm       As Integer
    Global g_iElbowTape     As Integer
    Global g_NoElbowTape    As Integer
    Global g_sOriginalContracture   As String
    Global g_sOriginalLining        As String

   'Globals set by FN_Open
    Global cc   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum  As Integer 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote


    Global g_sPathJOBST As String

    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont

    Global g_sFileNo        As String
    Global g_sSide          As String
    Global g_sPatient       As String
    Global g_sCurrentLayer  As String

    Global g_iWristNo       As Integer
    Global g_iPalmNo       As Integer
    
    Global g_nPalmWristDist As Double
    Global g_iModulus       As Integer

   'POWERNET and BOBINNET Modulus charts
   'This format was initially created to allow DRAFIX to index into a string,
   'to mimic an array
   'This format is retained here for compatibility.
    Type fabric
        Modulus(0 To 18) As String
        Conversion(0 To 18) As String
    End Type

    Global g_MATERIAL As fabric

   'Flags
    Global g_Gauntlet       As Integer
    Global g_sID            As String
    Global g_sStyle         As String

    Option Explicit



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



Function fnGetNumber(ByVal sString As String, iIndex As Integer) As Double
   'Function to return as a numerical value the iIndexth item in a string
   'that uses blanks (spaces) as delimiters.
   'EG
   '    sString = "12.3 65.1 45"
   '    fnGetNumber( sString, 2) = 65.1
   '
   'If the iIndexth item is not found then return -1 to indicate an error.
   'This assumes that the string will not be used to store -ve numbers.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

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
            sString = Left$(sString, iPos - 1)
            fnGetNumber = Val(sString)
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos))
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

Sub PR_LoadFabricFromFile(MATERIAL As fabric, sFileName As String)
   'Procedure to load the MATERIAL conversion chart from file
   'N.B. File opening Errors etc. are not handled (so tough titty!)
    
    Dim fNum As Integer, ii As Integer
    fNum = FreeFile
    Open sFileName For Input As fNum
    ii = 0
    Do Until EOF(fNum)
        Input #fNum, MATERIAL.Modulus(ii), MATERIAL.Conversion(ii)
        ii = ii + 1
    Loop
    Close fNum

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

Sub Select_Text_In_Box(Text_Box_Name As Control)
    Text_Box_Name.SetFocus
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub


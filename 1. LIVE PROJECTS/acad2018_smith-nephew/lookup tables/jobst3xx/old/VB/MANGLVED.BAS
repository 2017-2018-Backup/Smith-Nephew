Attribute VB_Name = "MANGLVED1"

'   '* Windows API Functions Declarations
'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
'    Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
   
'   'Constants used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

    Type XY
        X As Double
        y As Double
    End Type
    
    Public MainForm As manglved

    Type Curve
        n As Integer
        X(1 To 100) As Double
        y(1 To 100) As Double
    End Type

   'Globals
    Global g_nUnitsFac      As Double
    Global g_Direction      As Integer

    Global g_iGms(17)       As Integer
    Global g_iRed(17)       As Integer
    Global g_iMMs(17)       As Integer
    Global g_nLengths(17)   As Double
   
    Global g_iGmsInit(17)   As Integer
    Global g_iRedInit(17)   As Integer
    Global g_iMMsInit(17)   As Integer
    Global g_nLengthsInit(17) As Double
    Global g_iChanged(17)   As Integer
    Global g_iVertexETSMap(17) As Integer 'Maps length to editable profile vertex
    Global g_iVertexLFSMap(17) As Integer 'Maps length to editable profile vertex
    
    Global g_iVertexETSBlendMap(17) As Integer 'Maps length to blending vertex
    Global g_iVertexLFSBlendMap(17) As Integer 'Maps length to blending vertex
    Global g_ETSBlendProfile As Curve
    Global g_LFSBlendProfile As Curve
    Global g_BlendingCurveChanged As Integer
    Global g_nInsertSize    As Double
    Global g_iInsertStyle   As Integer

    Global g_iFirstTape     As Integer
    Global g_iLastTape      As Integer

    Global g_ReDrawn        As Integer

   'Profile Globals
    Global xyProfileETS(200)    As XY
    Global xyProfileLFS(200)    As XY
    Global g_iProfileETS        As Integer
    Global g_iProfileLFS        As Integer
    Global xyOtemplate          As XY
    Global xyPALMER             As XY
    Global xyPALM2              As XY
    Global xyPALM3              As XY
    Global xyPALM4              As XY
    Global xyPALM5              As XY
    Global xyPALM6              As XY
    Global g_iFirstEditableVertex   As Integer
    Global g_iFirstEditTape     As Integer
    Global g_iLastEditTape      As Integer
    Global g_OffFold            As Integer
    Global g_iElbowTape         As Integer
    Global g_NoElbowTape        As Integer
    Global g_sOriginalContracture   As String
    Global g_sOriginalLining        As String


   'Globals set by FN_Open
    Global CC   'Comma
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
    Global g_sID            As String
    Global g_sStyle         As String


    Global g_GloveType      As Integer
    Global g_EOSType        As Integer

    Global Const ARM_PLAIN = 0
    Global Const ARM_FLAP = 1
    Global Const GLOVE_NORMAL = 0
    Global Const GLOVE_ELBOW = 1
    Global Const GLOVE_AXILLA = 2
    Global Const PI = 3.141592654
    Global Const ELBOW_TAPE = 10
    Global Const EIGHTH = 0.125


    Option Explicit



Function FN_CalcAngle(xyStart As XY, xyEnd As XY) As Double
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim X, y, rAngle As Double
    
    X = xyEnd.X - xyStart.X
    y = xyEnd.y - xyStart.y

   'Horizontal
    If X = 0 Then
        If y > 0 Then
            FN_CalcAngle = 90
        Else
            FN_CalcAngle = 270
        End If
        Exit Function
    End If
    
   'Vertical (avoid divide by zero later)
    If y = 0 Then
        If X > 0 Then
            FN_CalcAngle = 0
        Else
            FN_CalcAngle = 180
        End If
        Exit Function
    End If

   'All other cases
    rAngle = Atn(y / X) * (180 / PI)         'Convert to degrees
    
    If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
    
    If y > 0 Then
        FN_CalcAngle = rAngle
    Else
        FN_CalcAngle = rAngle + 180
    End If

End Function

Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)

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

Sub PR_AddEntityID(sFileNo As String, sSide As String, sType)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to add to an ENTITY the database information
   'in the DB variable "ID" that will allow the identity of an entity
   'to be retrieved, by other parts of the system.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined.
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Dim sID As String
    
    sID = sFileNo & sSide & sType

    Print #fNum, "if (hEnt) SetDBData( hEnt," & QQ & "ID" & QQ & CC & QQ; sID & QQ & ");"

End Sub

Static Sub PR_CalcMidPoint(xyStart As XY, xyEnd As XY, xyMid As XY)
    
    Dim aAngle As Double, nLength As Double

    aAngle = FN_CalcAngle(xyStart, xyEnd)
    nLength = FN_CalcLength(xyStart, xyEnd)
    
    If nLength = 0 Then
        xyMid = xyStart 'Avoid overflow
    Else
        PR_CalcPolar xyStart, aAngle, nLength / 2, xyMid
    End If

End Sub


Static Sub PR_CalcPolar(xyStart As XY, ByVal nAngle As Double, ByVal nLength As Double, xyReturn As XY)
   'Procedure to return a point at a distance and an angle from a given point
   '
   'PI is a Global Constant

    Dim A As Double, B As Double

   'Ensure angles are +ve and between 0 to 360 degrees
    If nAngle > 360 Then nAngle = nAngle - 360
    If nAngle < 0 Then nAngle = nAngle + 360

    Select Case nAngle
        Case 0
            xyReturn.X = xyStart.X + nLength
            xyReturn.y = xyStart.y
        Case 180
            xyReturn.X = xyStart.X - nLength
            xyReturn.y = xyStart.y
        Case 90
            xyReturn.X = xyStart.X
            xyReturn.y = xyStart.y + nLength
        Case 270
            xyReturn.X = xyStart.X
            xyReturn.y = xyStart.y - nLength
        Case Else
           'Convert from degees to radians
            nAngle = nAngle * PI / 180
            B = Sin(nAngle) * nLength
            A = Cos(nAngle) * nLength
            xyReturn.X = xyStart.X + A
            xyReturn.y = xyStart.y + B
        End Select

End Sub





Sub PR_DeleteByID(sID As String)
    'Procedure to locate and delete all entitie that have the
    'string sID in a DRAFIX data base variable "ID"
     Print #fNum, "hChan=Open("; QQ; "selection"; QCQ; "DB ID = '"; sID; "'"; QQ; ");"
     Print #fNum, "if(hChan)"
     Print #fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}"
     Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"
End Sub

Sub PR_DrawMarker(xyPoint As XY)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; xyPoint.X; CC; xyPoint.y; CC; "0.125);"

End Sub

Sub PR_DrawPoly(Profile As Curve)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a POLYLINE through the points
   'given in Profile.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '
    Dim ii As Integer

   'Exit if nothing to draw
    If Profile.n <= 1 Then Exit Sub

    Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    For ii = 1 To Profile.n
        Print #fNum, CC; Str$(Profile.X(ii)); CC; Str$(Profile.y(ii))
    Next ii
    Print #fNum, ");"

End Sub

Sub PR_DrawText(sText As Variant, xyInsert As XY, nHeight As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw TEXT at the given height.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
   '
   '
    Dim nWidth
    nWidth = nHeight * g_nCurrTextAspect
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; Str$(xyInsert.X); CC; Str$(xyInsert.y); CC; nWidth; CC; nHeight; ",0);"
    
End Sub

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

Sub PR_MakeXY(xyReturn As XY, X As Double, y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.X = X
    xyReturn.y = y
End Sub

Sub PR_PutTapeLabel(nTape As Integer, nLength As Variant, nMM As Variant, nGrm As Variant, nRed As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to add Sleeve Tape details,
   'these details are given explicitly as arguments.
   'Where:-
   '    nTape       Index into sTextList below
   '    nLength     Tape length to be displayed, decimal inches
   '    nMM         MMs to be displayed
   '    nGrm        Grams to be displayed
   '    nRed        Reduction to be displayed
   '
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, g_sFileNo, g_sSide are globals initialised by FN_Open
   '
   '
    Dim sTextList As String, sSymbol As String, nInt%, nDec#
    Dim sInt$, sDec$, sMM$, sRed$, sTape$, sGrams$
    Dim xyPt As XY, nSymbolOffSet As Single, nTextHt As Single

    sTextList = "   -4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
    sTape = Trim(Mid(sTextList, (nTape * 3) + 1, 3))
    

   'Length text
   'N.B. format as Inches and eighths. With eighths offset up and left
    nInt = Int(nLength)     'Integer part of the length (before decimal point)

   'Decimal part of the length (after decimal point)
   'convert to 1/8ths and get nearest by rounding
    nDec = round((nLength - nInt) / 0.125)
    If nDec = 8 Then
        nDec = 0
        nInt = nInt + 1
    End If

   'Integer part
    sInt = Trim(nInt)

   'Eighths part
    If nDec <> 0 Then sDec = Trim(nDec)

   'MMs text
    sMM = Trim(nMM) & "mm"

   'Grams text
    sGrams = Trim(nGrm) & "gm"

   'Reduction text
    sRed = Trim(nRed)

   'Find the symbol and update it's values
    Print #fNum, "hChan=Open("; QQ; "selection"; QCQ; "DB SymbolName = 'glvTapeNotes' AND DB ID = '"; g_sID; "' AND DB Data = '"; sTape; "'"; ; QQ; ");"
    Print #fNum, "if(hChan)"
    Print #fNum, "{ResetSelection(hChan);"
    Print #fNum, " hEnt=GetNextSelection(hChan);"
    Print #fNum, " if(hEnt) {"
    Print #fNum, "  SetDBData(hEnt,"; QQ; "MM"; QCQ; sMM; QQ; ");"
    Print #fNum, "  SetDBData(hEnt,"; QQ; "Grams"; QCQ; sGrams; QQ; ");"
    Print #fNum, "  SetDBData(hEnt,"; QQ; "Reduction"; QCQ; sRed; QQ; ");"
    Print #fNum, "  SetDBData(hEnt,"; QQ; "TapeLenghts"; QCQ; sInt; QQ; ");"
    Print #fNum, "  SetDBData(hEnt,"; QQ; "TapeLengths2"; QCQ; sDec; QQ; ");"
    Print #fNum, " }"
    Print #fNum, "}"
    Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"

End Sub

Sub PR_SetTextData(nHoriz As Variant, nVert As Variant, nHt As Variant, nAspect As Variant, nFont As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the TEXT default attributes, these are
   'based on the values in the arguments.  Where the value is -ve then this
   'attribute is not set.
   'where :-
   '    nHoriz      Horizontal justification (1=Left, 2=Cen, 4=Right)
   '    nVert       Verticalal justification (8=Top, 16=Cen, 32=Bottom)
   '    nHt         Text height
   '    nAspect     Text aspect ratio (heigth/width)
   '    nFont       Text font (0 to 18)
   '
   'N.B. No checking is done on the values given
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextHt, g_nCurrTextAspect,
   '    g_nCurrTextHorizJust, g_nCurrTextVertJust, g_nCurrTextFont
   '    are globals initialised by FN_Open
   '
  
    If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then
        Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; nHoriz; ");"
        g_nCurrTextHorizJust = nHoriz
    End If

    If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
        Print #fNum, "SetData("; QQ; "TextVertJust"; QC; nVert; ");"
        g_nCurrTextVertJust = nVert
    End If

    If nHt >= 0 And g_nCurrTextHt <> nHt Then
        Print #fNum, "SetData("; QQ; "TextHeight"; QC; nHt; ");"
        g_nCurrTextHt = nHt
    End If

    If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
        Print #fNum, "SetData("; QQ; "TextAspect"; QC; nAspect; ");"
        g_nCurrTextAspect = nAspect
    End If

    If nFont >= 0 And g_nCurrTextFont <> nFont Then
        Print #fNum, "SetData("; QQ; "TextFont"; QC; nFont; ");"
        g_nCurrTextFont = nFont
    End If


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


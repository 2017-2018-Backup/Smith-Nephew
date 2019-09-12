Attribute VB_Name = "CADGLOVE1"
'Project:   CADGLOVE.BAS
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

'    Declare Function GetNumTasks Lib "Kernel" () As Integer
    
   'PI as a Global Constant
    Global Const PI = 3.141592654

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global nL   'Newline
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

    Global MainForm As cadglove

    Global g_OnFold                As Integer
    Global g_nPalm                 As Double
    Global g_nWrist                As Double
    Global g_iInsertSize           As Integer
    
    Global xyDatum             As XY
    Global xyLittle            As XY
    Global xyLFS               As XY
    Global xyRing              As XY
    Global xyIndex             As XY
    Global xyMiddle            As XY
    Global xyThumb             As XY
    Global xyPalm(1 To 7)      As XY
    Global aLFAsweep           As Double
    Global xyLFAcen            As XY


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
    Global g_DataIsCalcuable           As Integer
    
    Global g_ExtendTo                  As Integer

   'Flaps
    Global g_nStrapLength          As Double
    Global g_nFrontStrapLength     As Double
    Global g_nCustFlapLength       As Double
    Global g_nWaistCir             As Double
    Global g_sFlapType             As String
    Global g_iFlapType             As Integer

    Global UlnarProfile As Curve
    Global RadialProfile As Curve
    Global TapeNote(1 To NOFF_ARMTAPES)   As TapeData


    Global Const g_sDialogID = "MANUAL Glove Dialogue"
   'Open tip constants
    Global Const AGE_CUTOFF = 10           'Years
    Global Const ADULT_STD_OPEN_TIP = 0.5  'Inches
    Global Const CHILD_STD_OPEN_TIP = 0.375 'Inches
    

'   'MsgBox return values
'    Global Const IDOK = 1                  ' OK button pressed
'    Global Const IDCANCEL = 2              ' Cancel button pressed
'    Global Const IDABORT = 3               ' Abort button pressed
'    Global Const IDRETRY = 4               ' Retry button pressed
'    Global Const IDIGNORE = 5              ' Ignore button pressed
'    Global Const IDYES = 6                 ' Yes button pressed
'    Global Const IDNO = 7                  ' No button pressed
    
   'Scale to match current CAD-GLOVE output
    Global Const ISCALE = 1026
    
    Global XTRANSLATE As Double
    Global YTRANSLATE  As Double
    Global ROTATION  As Double
    Global xyOrigin As XY
    
    Option Explicit
Function FN_InchesValue(TextBox As Control) As Double
   'Check for numeric values
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
            TextBox.SetFocus
            FN_InchesValue = -1
            Exit For
        End If
    Next nn

   'Convert to inches
    nLen = fnDisplayToInches(Val(TextBox.Text))
    If nLen = -1 Then
        MsgBox "Invalid - Length has been entered", 48, g_sDialogID
        TextBox.SetFocus
        FN_InchesValue = -1
    Else
        FN_InchesValue = nLen
    End If

End Function

Function FN_Open(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_Open = fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    nL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName
    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 32    'Bottom
    g_nCurrTextFont = 0         'BLOCK
    g_nCurrTextAngle = 0


   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; g_sPatient; ", "; g_sFileNo; ", Hand - "; g_sSide
    Print #fNum, "//by Visual Basic, GLOVES - Drawing"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;"
    Print #fNum, "XY     xyStart, xyOrigin, xyScale, xyO;"
    Print #fNum, "STRING sFileNo, sSide, sID, sName;"
    Print #fNum, "ANGLE  aAngle;"

   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Data"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ZipperLength"; QCQ; "length"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Zipper"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengths"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengths2"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengthPt1"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "curvetype"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Leg"; QCQ; "string"; QQ; ");"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; "));"
    
   'Set values for use futher on by other macros
    Print #fNum, "sSide = "; QQ; g_sSide; QQ; ";"
    Print #fNum, "sFileNo = "; QQ; g_sFileNo; QQ; ";"
    
   'Get Start point
    Print #fNum, "GetUser ("; QQ; "xy"; QCQ; "Indicate Start Point"; QC; "&xyStart);"
   
   'Place a marker at the start point for later use.
   'Get a UID and create the unique 4 character start to the ID code
   'Note this is a bit dogey if the drawing contains more than 9999 entities
    PR_SetLayer "Construct"
    Print #fNum, "hOrigin = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart"; CC; "0.125);"
    Print #fNum, "if (hOrigin) {"
    Print #fNum, "  sID=StringMiddle(MakeString("; QQ; "long"; QQ; ",UID("; QQ; "get"; QQ; ",hOrigin)), 1, 4) ; "
    Print #fNum, "  while (StringLength(sID) < 4) sID = sID + "; QQ; " "; QQ; ";"
    Print #fNum, "  sID = sID + sFileNo + sSide ;"
    Print #fNum, "  SetDBData(hOrigin,"; QQ; "ID"; QQ; ",sID);"
    Print #fNum, "  }"
    
   'Display Hour Glass Symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
   'Set values for use futher on by other macros
    Print #fNum, "xyOrigin = xyStart"; ";"

    Print #fNum, "hMPD = UID ("; QQ; "find"; QC; Val(MainForm!txtUidMPD.Text); ");"
    Print #fNum, "if (hMPD)"
    Print #fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);"
    Print #fNum, "else"
    Print #fNum, "  Exit(%cancel,"; QQ; "Can't find > mainpatientdetails < symbol, Insert Patient Data"; QQ; ");"

   'Start drawing on correct side
    PR_SetLayer "Template" & g_sSide

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

Sub PR_UpdateDBFields(sType As String)
   'sType = "Save" for use with save only
   'sType = "Draw" for use with a drawing macro

    Dim sSymbol$

   'Glove common
    sSymbol = "glovecommon"
    If MainForm!txtUidGC.Text = "" Then
       'Insert a new symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ")){"
        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x, xyO.y);"
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find >"; sSymbol; "< symbol to insert\nCheck your installation, that JOBST.SLB exists!"; QQ; ");"
    Else
       'Use existing symbol
        Print #fNum, "hSym = UID ("; QQ; "find"; QC; Val(MainForm!txtUidGC.Text); ");"
        Print #fNum, "if (!hSym) Exit(%cancel,"; QQ; "Can't find >"; sSymbol; "< symbol to update!"; QQ; ");"
    End If
    
   'Update DB fields
    Print #fNum, "SetDBData( hSym"; CQ; "Fabric"; QCQ; MainForm!txtFabric.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Data"; QCQ; MainForm!txtDataGC.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; MainForm!txtFileNo.Text; QQ; ");"

   'Glove
    sSymbol = "gloveglove"
    If MainForm!txtUidGlove.Text = "" Then
       'Insert a new symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ")){"
        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        If MainForm!optHand(0) = True Then
           'Left Hand
            Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x+1.5 , xyO.y);"
        Else
           'Right Hand
            Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x+3 , xyO.y);"
        End If
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find >"; sSymbol; "< symbol to insert\nCheck your installation, that JOBST.SLB exists!"; QQ; ");"
    Else
       'Use existing symbol
        Print #fNum, "hSym = UID ("; QQ; "find"; QC; Val(MainForm!txtUidGlove.Text); ");"
        Print #fNum, "if (!hSym) Exit(%cancel,"; QQ; "Can't find >"; sSymbol; "< symbol to update!"; QQ; ");"
    End If
    
   'Update DB fields
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths"; QCQ; MainForm!txtTapeLengths.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths2"; QCQ; MainForm!txtTapeLengths2.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengthPt1"; QCQ; MainForm!txtTapeLengthPt1.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; MainForm!txtFileNo.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Sleeve"; QCQ; MainForm!txtSide.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Grams"; QCQ; MainForm!txtGrams.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Reduction"; QCQ; MainForm!txtReduction.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeMMs"; QCQ; MainForm!txtTapeMMs.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Data"; QCQ; MainForm!txtDataGlove.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "WristPleat"; QCQ; MainForm!txtWristPleat.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "ShoulderPleat"; QCQ; MainForm!txtShoulderPleat.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Flap"; QCQ; MainForm!txtFlap.Text; QQ; ");"

    If sType = "Draw" Then
       'Update the fields associated with the PALMER marker
       'This was given an explicit handle in the procedure PR_CreateDrawMacro
       '
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengths"; QCQ; MainForm!txtTapeLengths.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengths2"; QCQ; MainForm!txtTapeLengths2.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengthPt1"; QCQ; MainForm!txtTapeLengthPt1.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeMMs"; QCQ; MainForm!txtTapeMMs.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "fileno"; QCQ; MainForm!txtFileNo.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Sleeve"; QCQ; MainForm!txtSide.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Data"; QCQ; MainForm!txtDataGlove.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "age"; QCQ; MainForm!txtAge.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "units"; QCQ; MainForm!txtUnits.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Grams"; QCQ; MainForm!txtGrams.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Reduction"; QCQ; MainForm!txtReduction.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeMMs"; QCQ; MainForm!txtTapeMMs.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Data"; QCQ; MainForm!txtDataGlove.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "WristPleat"; QCQ; MainForm!txtWristPleat.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "ShoulderPleat"; QCQ; MainForm!txtShoulderPleat.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Flap"; QCQ; MainForm!txtFlap.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Fabric"; QCQ; MainForm!txtFabric.Text; QQ; ");"
    End If
End Sub

Sub PR_UpdateDDE()
        
   'Procedure to update the fields used when data is transfered
   'from DRAFIX using DDE.
   'Although the transfer back to DRAFIX is not via DDE we use the same controls
   'simply to illustrate the method by which the data is packed into the
   'fields

   'The decisons on format for each is not very neat!
   'The reason is the iteritive development process and the 63 char length
   'limitation on DRAFIX DB Fields.

   'Backward compatibility with the CAD-Glove is required in this
   'procedure

    MainForm!SSTab1.Tab = 0

    Dim iLen%, ii%
    Dim sLen$, sPacked$

   'Initialise
    sPacked = ""
   
   'Pack Circumferences
    For ii = 0 To 10
        iLen = Val(MainForm!txtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Pack Lengths
    For ii = 0 To 8
        iLen = Val(MainForm!txtLen(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Store to DDE text boxes
    MainForm!txtTapeLengths.Text = sPacked

   'Tip options
    ReDim Closed(0 To 4), StdOpen(0 To 4)
    Closed(0) = MainForm!optLittleTip(0).Value
    Closed(1) = MainForm!optRingTip(0).Value
    Closed(2) = MainForm!optMiddleTip(0).Value
    Closed(3) = MainForm!optIndexTip(0).Value
    Closed(4) = MainForm!optThumbTip(0).Value
    
    StdOpen(0) = MainForm!optLittleTip(1).Value
    StdOpen(1) = MainForm!optRingTip(1).Value
    StdOpen(2) = MainForm!optMiddleTip(1).Value
    StdOpen(3) = MainForm!optIndexTip(1).Value
    StdOpen(4) = MainForm!optThumbTip(1).Value

    sPacked = ""
    For ii = 0 To 4
        If Closed(ii) = True Then
            sPacked = sPacked & "0"
        ElseIf StdOpen(ii) = True Then
            sPacked = sPacked & "1"
        Else
            sPacked = sPacked & "2"
       End If
    Next ii

   'CheckBoxes
    sPacked = sPacked & Trim$(Str$(MainForm!chkSlantedInserts.Value))
    sPacked = sPacked & Trim$(Str$(MainForm!chkPalm.Value))
    
   'Tapes past wrist (Store here due to limitations on DB field length
   'Pack Circumferences
    For ii = 11 To 12
        iLen = Val(MainForm!txtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Insert size
    sPacked = sPacked & " " & Str$(g_iInsertSize)

   'Reinforced dorsal
    sPacked = sPacked & Trim$(Str$(MainForm!chkDorsal.Value))

   'Store to DDE text boxes
    MainForm!txtTapeLengths2.Text = sPacked

   'Fabric
    If MainForm!cboFabric.Text <> "" Then MainForm!txtFabric = MainForm!cboFabric.Text

    
End Sub


Attribute VB_Name = "ARMDIA1"
   'XY data type to represent points
    Type XY
        X As Double
        y As Double
    End Type
    
    Type curve
        n As Integer
        X(1 To 100) As Double
        y(1 To 100) As Double
    End Type
    
   'PI as a Global Constant
    Global Const PI = 3.141592654
    
    Public MainForm As armdia

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

   'Constants to define the drive and the root directory
    Global g_sFileNo As String    'The patients file no
    Global g_sSide As String      'The side Left or right
    Global g_sPatient As String   'The patients name
    Global g_sFlapLength As String
    Global g_sFabric As String
    Global g_sFabnam As String
    Global g_sPressureChange As String
    Global g_sFabricChange As String
    Global g_sWpleat1 As String
    Global g_sWpleat2 As String
    Global g_sSpleat1 As String
    Global g_sSpleat2 As String
    Global g_sFlapStrap As String
    Global g_sFlapChk As Variant
    Global g_sinchflag As String
    Global g_sFlap As String
    Global g_sMM As String
    Global g_sGaunt As String
    Global g_sDetGaunt As String
    Global g_sNoThumb As String
    Global g_sPalmNo As String
    Global g_sWristNo As String
    Global g_sPalmWristDist As String
    Global g_sThumbCircum As String
    Global g_sThumbLength As String
    Global g_sEnclosedThumb As String
    Global g_sSecondLastTape As String
    Global g_sSecondTape As String
    Global g_sFirstTape As String
    Global g_sLastTape As String
    Global g_sModulus As String
    Global g_sStump As String
    Global g_sTapeLengths As String
    Global g_sTapeMMs As String
    Global g_sGrams As String
    Global g_sReduction As String
    Global g_sAmm As Variant
    Global g_sBmm As Variant
    Global g_sCmm As Variant
    Global g_sDmm As Variant
    Global g_sWaistCir

    Global g_sWorkOrder As String
    
   'Store current layer and text setings to reduce DRAFIX code
   'this value is checked in PR_SetLayer
    Global g_sCurrentLayer As String
    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont
    Global g_nCurrTextAngle

    Global g_sHoleCheck As String
    Global g_sOld As String

    Global g_sID As String
    Global g_sEditLengths As String
    Global g_sUnits As String


    Global g_iStyleFirstTape As Integer
    Global g_iStyleLastTape As Integer
    Global g_iFirstTape As Integer
    Global g_iLastTape As Integer
    Global g_nFrontStrapLength As Double
    Global g_nGauntletExtension As Double

    Global g_Modified As Integer
    Global g_sPathJOBST As String

    Global g_bDrawVestMesh  As Integer
    Global g_bDrawBodyMesh  As Integer

    Global PalmNo, WristNo, ThumLen, Detachable, EnclosedThm




Sub DataBaseDataUpDate(sType As String)
 Dim sString As String
 Dim sGauntlet As String, sFlap As String
 Dim sBoxType As String, nYoffset As Single


'Choose data box depending on type
 If armdia.txtType.Text = "ARM" Then
    sBoxType = "armarm"
 Else
    sBoxType = "vestarm"
 End If

'Get sleevebox handle
 Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"
 Print #fNum, "hChan = Open("; QQ; "selection"; QCQ; "DB SymbolName = '"; sBoxType; "' AND DB Sleeve ='"; g_sSide; "'"; QQ; ");"
 Print #fNum, "ResetSelection(hChan);"
 Print #fNum, "hSleeve = GetNextSelection(hChan);"
 Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"

'If sleevebox not found then insert one
  Print #fNum, "if(!hSleeve){"
  Print #fNum, "    Close("; QQ; "selection"; QC; "hChan);"
  Print #fNum, "    hChan = Open("; QQ; "selection"; QCQ; "DB SymbolName = 'mainpatientdetails'"; QQ; ");"
  Print #fNum, "    ResetSelection(hChan);"
  Print #fNum, "    hTitle = GetNextSelection(hChan);"
  Print #fNum, "    Close("; QQ; "selection"; QC; "hChan);"
'Get title box origin
  Print #fNum, "    GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);"

'Insert arm or sleeve box
  Print #fNum, "    if ( !Symbol("; QQ; "find"; QCQ; sBoxType; QQ; ")) Exit(%cancel,"; QQ; "Cant find SLEEVEBOX or ARMBOX symbol to insert|nCheck your installation, that JOBST.SLB exists"; QQ; ");"
  Print #fNum, "    if(StringCompare("; QQ; "Left"; QCQ; g_sSide; QQ; ")) xyTitleOrigin.x = xyTitleOrigin.x + 1.5;"
  Print #fNum, "        else xyTitleOrigin.x = xyTitleOrigin.x + 3;"
  Print #fNum, "    Execute("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
  Print #fNum, "    hSleeve = AddEntity("; QQ; "symbol"; QCQ; sBoxType; QC; "xyTitleOrigin);"
  Print #fNum, "}"

   'Insert Arm common
    sSymbol = "armcommon"
    If armdia.txtUidAC.Text = "" Then
      'Insert a new symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ")){"
'        Print #fNum, "if(!hSleeve){"
        Print #fNum, "  Close("; QQ; "selection"; QC; "hChan);"
        Print #fNum, "  hChan = Open("; QQ; "selection"; QCQ; "DB SymbolName = 'mainpatientdetails'"; QQ; ");"
        Print #fNum, "  ResetSelection(hChan);"
        Print #fNum, "  hTitle = GetNextSelection(hChan);"
        Print #fNum, "  Close("; QQ; "selection"; QC; "hChan);"
        'Get title box origin
        Print #fNum, "  GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);"

        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyTitleOrigin.x, xyTitleOrigin.y);"
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find >"; sSymbol; "< symbol to insert\nCheck your installation, that JOBST.SLB exists!"; QQ; ");"
    Else
       'Use existing symbol
        Print #fNum, "hSym = UID ("; QQ; "find"; QC; Val(armdia.txtUidAC.Text); ");"
        Print #fNum, "if (!hSym) Exit(%cancel,"; QQ; "Can't find >"; sSymbol; "< symbol to update!"; QQ; ");"
    End If
   'Update DB fields
    Print #fNum, "SetDBData( hSym"; CQ; "Fabric"; QCQ; armdia.txtFabric.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; armdia.txtFileNo.Text; QQ; ");"



  

 
'Flap multiple field
 sString = String(35, " ")
 LSet sString = g_sFlap
 If Val(armdia.txtCustFlapLength) = 0 Then armdia.txtCustFlapLength = -1
 If Val(armdia.txtWaistCir) = 0 Then armdia.txtWaistCir = "-1"
 If Val(armdia.txtFrontStrapLength) = 0 Then armdia.txtFrontStrapLength = "-1"
 sFlap = sString & armdia.txtStrap.Text & " " & armdia.txtCustFlapLength.Text & " " & armdia.txtWaistCir & " " & armdia.txtFrontStrapLength

'Gauntlet Multiple field
 If g_sGaunt = "False" Then
    sGauntlet = "0"
 Else
   'Set falgs
    sString = "1 "
    If g_sEnclosedThumb = "True" Then
        sString = sString & "1 "
    Else
        sString = sString & "0 "
    End If

    If g_sDetGaunt = "True" Then
        sString = sString & "1 "
    Else
        sString = sString & "0 "
    End If

    If g_sNoThumb = "True" Then
        sString = sString & "1 "
    Else
        sString = sString & "0 "
    End If
    If Val(armdia.txtGauntletExtension) = 0 Then armdia.txtGauntletExtension = "-1"
   'Data
    sGauntlet = sString & Val(g_sWristNo) & " " & Val(g_sPalmNo) & " " & Val(g_sThumbLength) & " " & Val(g_sThumbCircum) & " " & Val(g_sPalmWristDist) & " " & armdia.txtGauntletExtension.Text
 End If
 
'Update data base fields
'
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Sleeve"; QCQ; g_sSide; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Fileno"; QCQ; g_sFileNo; QQ; ");"

'Wrist and Shoulder pleat multiple fields
 Print #fNum, "SetDBData(hSleeve,"; QQ; "WristPleat"; QCQ; g_sWpleat1; " "; g_sWpleat2; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "ShoulderPleat"; QCQ; g_sSpleat1; " "; g_sSpleat2; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Flap"; QCQ; sFlap; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Gauntlet"; QCQ; sGauntlet; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "MM"; QCQ; g_sMM; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Modulus"; QCQ; g_sModulus; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Fabric"; QCQ; g_sFabric; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Stump"; QCQ; g_sStump; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "TapeMMs"; QCQ; g_sTapeMMs; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Reduction"; QCQ; g_sReduction; QQ; ");"
 Print #fNum, "SetDBData(hSleeve,"; QQ; "Grams"; QCQ; g_sGrams; QQ; ");"

'Tape lengths are universal, Store the actually used lengths with the
'Profile Origin Marker
 Print #fNum, "SetDBData(hSleeve,"; QQ; "TapeLengths"; QCQ; g_sTapeLengths; QQ; ");"
'Store ID of last drawn
 Print #fNum, "SetDBData(hSleeve,"; QQ; "ID"; QCQ; g_sID; QQ; ");"


If sType = "Draw" Then
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Fabric"; QCQ; g_sFabric; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Gauntlet"; QCQ; sGauntlet; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "TapeLengths"; QCQ; g_sEditLengths; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "units"; QCQ; g_sUnits; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "WristPleat"; QCQ; g_sWpleat1; " "; g_sWpleat2; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "ShoulderPleat"; QCQ; g_sSpleat1; " "; g_sSpleat2; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Flap"; QCQ; sFlap; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Grams"; QCQ; g_sGrams; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Reduction"; QCQ; g_sReduction; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "TapeMMs"; QCQ; g_sTapeMMs; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Stump"; QCQ; g_sStump; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "Modulus"; QCQ; g_sModulus; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "MM"; QCQ; g_sMM; QQ; ");"
 Print #fNum, "SetDBData(hOrigin,"; QQ; "age"; QCQ; armdia!txtAge.Text; QQ; ");"
End If

End Sub



Function FN_CalcAngle(xyStart As XY, xyEnd As XY) As Double
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim X, y, rAngle As Double
    
    X = xyEnd.X - xyStart.X
    y = xyEnd.y - xyStart.y

   'Horizomtal
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

Function FN_EscapeQuotesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeQuotesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = """" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeQuotesInString = sEscapedString

End Function

Function FN_EscapeSlashesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeSlashesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = "\" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeSlashesInString = sEscapedString

End Function

Function FN_GetNumber(ByVal sString As String, iIndex As Integer) As Double
   'Function to return as a numerical value the iIndexth item in a string
   'that uses blanks (spaces) as delimiters.
   'EG
   '    sString = "12.3 65.1 45"
   '    FN_GetNumber( sString, 2) = 65.1
   '
   'If the iIndexth item is not found then return -1 to indicate an error.
   'This assumes that the string will not be used to store -ve numbers.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

    If Len(sString) = 0 Then
        FN_GetNumber = -1
        Exit Function
    End If

   'Prepare string
    sString = sString & " " 'Trailing blank as stopper for last item

   'Get iIndexth item
    For ii = 1 To iIndex
        iPos = InStr(sString, " ")
        If ii = iIndex Then
            sString = Left$(sString, iPos - 1)
            FN_GetNumber = Val(sString)
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos))
            If Len(sString) = 0 Then
                FN_GetNumber = -1
                Exit Function
            End If
        End If
    Next ii

   'The function should have exited befor this, however just in case
   '(iIndex = 0) we indicate an error,
    FN_GetNumber = -1

End Function

Function FN_Open(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant, sType As String) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables
    Dim sID As String, sAxillaType$
    Dim iPos#
    Dim sStyle As String, sProximalStyle As String, sDistalStyle As String

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_Open = fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName
    If armdia!txtWorkOrder.Text = "" Then
        g_sWorkOrder = "-"
    Else
        g_sWorkOrder = armdia!txtWorkOrder.Text
    End If

    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 32    'Bottom
    g_nCurrTextFont = 0         'BLOCK
    g_nCurrTextAngle = 0

   'Create the 4 character string to identify the type
    PR_GetStyle sStyle

   
   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; g_sPatient; CC; " "; g_sFileNo; CC; " Sleeve-"; g_sSide
    Print #fNum, "//by Visual Basic"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hSleeve, hTitle, hChan, hEnt, hOrigin, hSym;"
    Print #fNum, "XY     xyAxilla, xyAxillaLow, xyTitleOrigin, xyStart, xyTitleScale, xyOrigin, xyElbow;"
    Print #fNum, "ANGLE  aTitleAngle;"
    Print #fNum, "STRING sTitleName, sFileNo, sSleeve, sID, sWorkOrder, sName, sVestID, sData, sDate, sPathJOBST;"
    
   'Set path to JOBST installed directory
    Print #fNum, "sPathJOBST = "; QQ; FN_EscapeSlashesInString(g_sPathJOBST); QQ; ";"

   'Set up ID  data base field
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ID"; QCQ; "string"; QQ; ");"

   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    
    g_sID = sStyle & g_sFileNo & g_sSide
   
   'Find axilla type
    If armdia!txtType <> "ARM" Then
        iPos = InStr(armdia!txtVestRaglan, QQ)   'Regular axilla are awkward as they
                                        'contain a space, look for quote (")
                                        'QQ is a constant set to quote (")
        If iPos <> 0 Then
           'Escape quote for use in drafix
            sAxillaType = Left$(armdia!txtVestRaglan, iPos - 1) & "\" & QQ
        Else
            iPos = InStr(armdia!txtVestRaglan, " ")
            sAxillaType = Left$(armdia!txtVestRaglan, iPos - 1)
        End If
   
    End If

    Print #fNum, "sSleeve = "; QQ; g_sSide; QQ; ";"
    Print #fNum, "sFileNo = "; QQ; g_sFileNo; QQ; ";"

   'Get Start point
    If sType = "Draw" Then
       'Get Start point
        Print #fNum, "SetSymbolLibrary( sPathJOBST + "; QQ; "\\JOBST.SLB"; QQ; ");"
        
        Print #fNum, "GetUser ("; QQ; "xy"; QCQ; "Indicate Start Point"; QC; "&xyStart);"
    
       'Place a marker at the start point for later use.
       'Get a UID and create the unique 4 character start to the ID code
       'Note this is a bit dogey if the drawing contains more than 9999 entities
        PR_SetLayer "Construct"
        Print #fNum, "hOrigin = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart"; CC; "0.125);"
        Print #fNum, "if (hOrigin) {"
        Print #fNum, "  sData=StringMiddle(MakeString("; QQ; "long"; QQ; ",UID("; QQ; "get"; QQ; ",hOrigin)), 1, 4) ; "
        Print #fNum, "  while (StringLength(sData) < 4) sData = sData + "; QQ; " "; QQ; ";"
        Print #fNum, "  sData = sData + sFileNo + sSleeve ;"
        Print #fNum, "  SetDBData( hOrigin," & QQ & "ID" & QQ & CC & QQ; g_sID & "originmark" & QQ & ");"
        Print #fNum, "  SetDBData( hOrigin," & QQ & "Data" & QQ & CC & "sData"; ");"
        Print #fNum, "  SetDBData( hOrigin," & QQ & "curvetype" & QQ & CC & QQ; "sleeveoriginmark" & QQ & ");"
        Print #fNum, "  }"
    End If

   'Display Hour Glass symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
   'Set values for use futher on by other macros
    Print #fNum, "xyOrigin = xyStart"; ";"
    Print #fNum, "sID = "; QQ; g_sID; QQ; ";"
    Print #fNum, "sName = "; QQ; g_sPatient; QQ; ";"
    Print #fNum, "sWorkOrder = "; QQ; armdia.txtWorkOrder.Text; QQ; ";"

End Function

Function fnGetString(ByVal sString As String, iIndex As Integer, sDelimiter As String) As String
   'Function to return as a string the iIndexth item in a string
   'that using the given string sDelimiter as the delimiter.
   'EG
   '    sString = "Sam Spade Hello"
   '    sDelimiter = " " {SPACE}
   '    fnGetNumber( sString, 2) = "Spade"
   '
   'If the iIndexth item is not found then return "" to indicate an error.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

    If Len(sString) = 0 Then
        fnGetString = ""
        Exit Function
    End If

   'Prepare string
    sString = sString & sDelimiter 'Trailing sDelimiter as stopper for last item

   'Get iIndexth item
    For ii = 1 To iIndex
        iPos = InStr(sString, sDelimiter)
        If ii = iIndex Then
            sString = Left$(sString, iPos - 1)
            fnGetString = sString
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos + 1))
            If Len(sString) = 0 Then
                fnGetString = ""
                Exit Function
            End If
        End If
    Next ii

   'The function should have exited befor this, however just in case
   '(iIndex = 0) we indicate an error,
    fnGetString = ""

End Function

Function max(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the maximum of two numbers
    If nFirst >= nSecond Then
        max = nFirst
    Else
        max = nSecond
    End If
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
    
    Print #fNum, "if (hEnt) SetDBData( hEnt," & QQ & "Data" & QQ & CC & "sData"; ");"

End Sub

Sub PR_CalcPolar(xyStart As XY, ByVal nAngle As Double, nLength As Double, xyReturn As XY)
   'Procedure to return a point at a distance and an angle from a given point
   '
    Dim A As Double, B As Double

   'Convert from degees to radians
    nAngle = nAngle * PI / 180

    B = Sin(nAngle) * nLength
    A = Cos(nAngle) * nLength

    xyReturn.X = xyStart.X + A
    xyReturn.y = xyStart.y + B

End Sub

Sub PR_CreateTapeLayer(sFileNo As String, sSide As String, nTape As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to create a layer for a sleeve tape label
   'This will be named by the following convention
   '    <"FileNo"> + <"L"|"R"> + <nTapeNo>
   'E.g. A1234567L10
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined.
   '    HANDLE  hLayer
   '
   'Note:-
   '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
   '
    Dim slayer As String

    slayer = sFileNo & Mid$(sSide, 1, 1) & nTape

    Print #fNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; slayer; QQ; ");"
    Print #fNum, "if ( hLayer != %badtable)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"
    Print #fNum, "else"
    Print #fNum, "Table ("; QQ; "add"; QCQ; "layer"; QCQ; slayer; QCQ; "Tape Layer Data"; QCQ; "current"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "DarkCyan"; QQ; "));"

End Sub

Sub PR_DeleteText(sFileNo As Variant, sSleeve As Variant, nTape As Variant)
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
   '
   Print #fNum, "nTape = "; nTape; ";"
   Print #fNum, "sFileNo = "; QQ; sFileNo; QQ; ";"
   Print #fNum, "sSleeve = "; QQ; sSleeve; QQ; ";"
   Print #fNum, "sLayer="; QQ; "layer = '"; QQ; "+ sFileNo + StringMiddle(sSleeve, 1, 1) + MakeString("; QQ; "long"; QQ; ", nTape) + "; QQ; "'"; QQ; ";"
   Print #fNum, "hChan = Open("; QQ; "selection"; QC; "sLayer);"
   Print #fNum, "UserSelection("; QQ; "clear"; QQ; ");"
   Print #fNum, "if (hChan){"
   Print #fNum, "ResetSelection(hChan);"
   Print #fNum, "while(hEnt = GetNextSelection (hChan)) DeleteEntity(hEnt);"
   Print #fNum, "}"
   Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"
End Sub

Sub PR_DrawArc(xyCen As XY, xyArcStart As XY, xyArcEnd As XY)

Dim nRad, nStartAng, nEndAng, nDeltaAng

    nRad = FN_CalcLength(xyCen, xyArcStart)

    nStartAng = FN_CalcAngle(xyCen, xyArcStart)

    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)

    nDeltaAng = nEndAng - nStartAng
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"

End Sub

Sub PR_DrawAssignDrafixVariable(sName As String, nValue As Double)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to do a variable assignment
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
   '
   '
    Print #fNum, sName; "="; nValue; ";"


End Sub

Sub PR_DrawCircle(xyCen As XY, nRadius As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw a CIRCLE at the point given.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Print #fNum, "hEnt = AddEntity("; QQ; "circle"; QC; "xyStart.x+"; Str$(xyCen.X); CC; "xyStart.y+"; Str$(xyCen.y); CC; nRadius; ");"
End Sub

Sub PR_DrawCircularStump(xyStart As XY, nFiguredLength As Double)
   'To the DRAFIX macro file (given by the global fNum)
   'Write the syntax to draw a CIRCULAR STUMP
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    g_sFileno, g_sPatient, g_sSide are globals initialised by FN_Open
   '

    Dim xyPt As XY, nTextHt As Single, nSeamAllowance As Single
    Dim nRadius As Single

    nTextHt = 0.125
    nSeamAllowance = 0.125  ' 1/8th Inch
    nRadius = (nFiguredLength / 3.14)
        
    PR_SetLayer "Template" & g_sSide

    PR_DrawCircle xyStart, nRadius + nSeamAllowance 'Add seam allowance
    PR_AddEntityID g_sFileNo, g_sSide, "CirStump"   'sFileNo, g_sSide from FN_Open

    PR_SetLayer "Notes"
    PR_DrawCircle xyStart, nRadius
    PR_AddEntityID g_sFileNo, g_sSide, "CirStumpSeam"   'sFileNo, g_sSide from FN_Open

    PR_SetTextData 2, 16, -1, -1, -1                'Horiz center, Vertical center
    PR_DrawText g_sPatient & "\n" & g_sWorkOrder & "\n" & g_sSide, xyStart, nTextHt

End Sub

Sub PR_DrawFitted(Profile As curve)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a FITTED curve through the points
   'given in Profile.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '
    Dim ii As Integer

   'Draw the profile
   '    If there is no vertex or only one vertex then exit.
   '    For two vertex draw as a polyline (this degenerates to a single line).
   '    For three vertex draw as a polyline (as no fitted curve can be drawn
   '    by a macro).
   '
    Select Case Profile.n
    Case 0 To 1
        Exit Sub
    Case 3
        PR_DrawPoly Profile
       'Warn the user to smooth the curve
        Print #fNum, "Display ("; QQ; "message"; QCQ; "OKquestion"; QCQ; "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line"; QQ; ");"
    Case Else
        Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        For ii = 1 To Profile.n
            Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.y(ii))
        Next ii
        Print #fNum, ");"
    End Select

End Sub

Sub PR_DrawLine(xyStart As XY, xyFinish As XY)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw a LINE between two points.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Print #fNum, "hEnt = AddEntity("
    Print #fNum, QQ; "line"; QC
    Print #fNum, "xyStart.x+"; Str$(xyStart.X); CC; "xyStart.y+"; Str$(xyStart.y); CC
    Print #fNum, "xyStart.x+"; Str$(xyFinish.X); CC; "xyStart.y+"; Str$(xyFinish.y)
    Print #fNum, ");"
        

End Sub

Sub PR_DrawLineOffset(xyStart As XY, xyFinish As XY, nOffset As Double)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a line between two points offset a given distance.
   'For this to work it assumes that xyStart is defined in DRAFIX
   'as an XY
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open.
   '    The line is always offset to the LEFT of the direction from start
   '    to finish.
   '    If nOffset is -ve then the line is offset to the RIGHT

    Dim nAngle As Double
    Dim xyPt1 As XY, xyPt2 As XY

    nAngle = FN_CalcAngle(xyStart, xyFinish)
    If nOffset < 0 Then
        nAngle = nAngle - 90
    Else
        nAngle = nAngle + 90
    End If

    PR_CalcPolar xyStart, nAngle, Abs(nOffset), xyPt1
    PR_CalcPolar xyFinish, nAngle, Abs(nOffset), xyPt2

    PR_DrawLine xyPt1, xyPt2

End Sub

Sub PR_DrawMesh()
   'Invokes external Mesh drawing routine
   'As such it must be the last line called by the DRAW Macro
    If g_bDrawBodyMesh Then
        Print #fNum, "Execute ("; QQ; "application"; QC; "sPathJOBST + "; QQ; "\\raglan\\meshdraw"; QCQ; "normal"; QQ; " );"
    End If
    If g_bDrawVestMesh Then
        Print #fNum, "Execute ("; QQ; "application"; QC; "sPathJOBST + "; QQ; "\\raglan\\meshvest"; QCQ; "normal"; QQ; " );"
    End If
End Sub

Sub PR_DrawPoly(Profile As curve)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to draw a POLYLINE through the points
   'given in Profile.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
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
        Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.y(ii))
    Next ii
    Print #fNum, ");"

End Sub

Sub PR_DrawRaglan(sType As String, sVestRaglan As Variant, sAge As Variant, sVestID As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to add a Raglan to the end of the sleeve
   'profile based on the data given in the variable sVestRaglan
   '
   'Note:-
   ' This routine is used to draw the raglan part of the
   ' sleeve based on the data from the vest.
   ' It is called VestRaglan to distinguish it from the
   ' raglan flap.
   '
   'Input:-
   ' sVestRaglan  A composite string consisting of the following:
   '
   '              sAxillaType, either Regular 2"
   '                                  Regular 1½"
   '                                  Regular 2½"
   '                                  Open
   '                                  Mesh
   '                                  Lining
   '                                  Sleeveless
   '              nAxillaFrontNeckRad
   '              nAxillaBackNeckRad
   '              nShoulderToBackRaglan
   '
   '              E.G. "Regular 1½" 5.565656 6.123344 0.3456
   '
   'Output:-
   ' Calls to 3 different macro files written to fNum
   '    Open macro file.
   '
   '        nAxillaFrontNeckRad   }
   '        nAxillaBackNeckRad    } Data from sVestRaglan
   '        nShoulderToBackRaglan } (For vest only)
   '        sAxillaType           }
   '
   '    Axilla macro file (depending on axilla given).
   '    Closing macro file.
   '
   '
   'Note: This routine has been amended to support Bodysuits
   '      For vest read Vest & BodySuit
    Dim sAxillaType As String, sString As String
    Dim iPos As Integer
    
   'Flag to draw the mesh used by PR_DrawMesh
    g_bDrawBodyMesh = False
    g_bDrawVestMesh = False
    

   'Find axilla type
    iPos = InStr(sVestRaglan, QQ)   'Regular axilla are awkward as they
                                    'contain a space, look for quote (")
                                    'QQ is a constant set to quote (")
    If iPos <> 0 Then
       'Escape quote for use in drafix
        sAxillaType = Left$(sVestRaglan, iPos - 1) & "\" & QQ
        sString = Mid$(sVestRaglan, iPos + 1)
    Else
        iPos = InStr(sVestRaglan, " ")
        sAxillaType = Left$(sVestRaglan, iPos - 1)
        sString = Mid$(sVestRaglan, iPos)
    End If

   'Check that an axilla has been given
    If sAxillaType = "None" Or sAxillaType = "Sleeveless" Then Exit Sub
    
   'NB the order of the following is very important
   '
   'Data from vest used in drawing the sleeve
    Print #fNum, "nAxillaFrontNeckRad= "; FN_GetNumber(sString, 1); ";"
    Print #fNum, "nAxillaBackNeckRad= "; FN_GetNumber(sString, 2); ";"
    Print #fNum, "nShoulderToBackRaglan = "; FN_GetNumber(sString, 3); ";"

    If sType = "Vest" Then
       'Initialise raglan drawing
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_INIT.D;"
       
       'Draw for the particular axilla type
        Print #fNum, "sAxillaType = "; QQ; sAxillaType; QQ; ";"

        Select Case sAxillaType
        Case "Open", "Lining"
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_OPEN.D;"
        Case "Mesh"
            g_bDrawVestMesh = True
            Print #fNum, "nAge = "; Val(sAge); ";"
            If g_sSide = "Left" Then
                Print #fNum, "nMeshLength = "; Val(fnGetString(armdia!txtMeshData, 2, ",")); ";"
                Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(armdia!txtMeshData, 1, ",")); ";"
            Else
                If Val(fnGetString(armdia!txtMeshData, 4, ",")) > 0 And Val(fnGetString(armdia!txtMeshData, 3, ",")) > 0 Then
                    Print #fNum, "nMeshLength = "; Val(fnGetString(armdia!txtMeshData, 4, ",")); ";"
                    Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(armdia!txtMeshData, 3, ",")); ";"
                Else
                    Print #fNum, "nMeshLength = "; Val(fnGetString(armdia!txtMeshData, 2, ",")); ";"
                    Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(armdia!txtMeshData, 1, ",")); ";"
                End If
            End If
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_MESH.D;"
        Case Else 'Regular
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_REGLR.D;"
        End Select
    
       'Close raglan drawing
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_CLOSE.D;"
    Else
       'Initialise raglan drawing
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\BR_INIT.D;"
       
       'Draw for the particular axilla type
        Print #fNum, "sAxillaType = "; QQ; sAxillaType; QQ; ";"
    
        Select Case sAxillaType
        Case "Open", "Lining"
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\BR_OPEN.D;"
        Case "Mesh"
            g_bDrawBodyMesh = True
            Print #fNum, "nAge = "; Val(sAge); ";"
            If g_sSide = "Left" Then
                Print #fNum, "nMeshLength = "; Val(fnGetString(armdia!txtMeshData, 2, ",")); ";"
                Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(armdia!txtMeshData, 1, ",")); ";"
            Else
                Print #fNum, "nMeshLength = "; Val(fnGetString(armdia!txtMeshData, 4, ",")); ";"
                Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(armdia!txtMeshData, 3, ",")); ";"
            End If
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\BR_MESH.D;"
        Case Else 'Regular
            Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\BR_REGLR.D;"
        End Select
    
       'Close raglan drawing
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\BR_CLOSE.D;"

    End If
End Sub

Sub PR_DrawText(sText As Variant, xyInsert As XY, nHeight As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw TEXT at the given height.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
   '
   '
    Dim nWidth
    nWidth = nHeight * g_nCurrTextAspect
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.y); CC; nWidth; CC; nHeight; CC; g_nCurrTextAngle; ");"
    
End Sub

Sub PR_GetStyle(sStyle As String)
   'This function returns a value of the following
   'in sWO
   'WorkOrder, ItemNo
   'eg
   '    12345,10
   '
    Dim nOptions%, sDistalStyle$, sProximalStyle$
    Dim sBestCatNo$, sStyleWO$, nResult%

   'Distal
    sDistalStyle = "XX"
    If g_iStyleFirstTape = g_iFirstTape Then
       'Plain
        sDistalStyle = "PL"
    Else
       'Start at tape
        If g_iStyleFirstTape < 10 Then
            sDistalStyle = "0" & LTrim$(Str$(g_iStyleFirstTape))
        Else
            sDistalStyle = LTrim$(Str$(g_iStyleFirstTape))
        End If
    End If
    If g_sGaunt = "True" Then
       'Gauntlet
        sDistalStyle = "GT"
    End If
    If g_sStump = "True" Then
       'Stump
        sDistalStyle = "ST"
    End If

   'Proximal
    sProximalStyle = "XX"
    If g_iStyleLastTape = g_iLastTape Then
       'Plain
        sProximalStyle = "PL"
    Else
       'Start at tape
        If g_iStyleLastTape < 10 Then
            sProximalStyle = "0" & LTrim$(Str$(g_iStyleLastTape))
        Else
            sProximalStyle = LTrim$(Str$(g_iStyleLastTape))
        End If
    End If
    If g_sDetGaunt = "True" And g_iStyleLastTape <= 6 Then
       'Detachable Gauntlet only, if drawn from or less than tape +1-1/2
        sProximalStyle = "GT"
    End If
    If g_sFlapChk = "True" Then
       'Flap
        sProximalStyle = "FP"
    End If
    If Left$(armdia.cboFlaps.Text, 4) = "Vest" Then
       'Vest raglan
        sProximalStyle = "VR"
    End If
    If Left$(armdia.cboFlaps.Text, 4) = "Body" Then
       'Bodysuit raglan
        sProximalStyle = "BR"
    End If

   'Join distal and proximal styles
    sStyle = sDistalStyle & sProximalStyle

    
End Sub

Sub PR_InsertSymbol(sSymbol As String, xyInsert As XY, nScale As Single, nRotation As Single, slayer As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to insert a SYMBOL.
   'Where:-
   '    sSymbol     Symbol name, must exist and be in the symbol library
   '    xyInsert    The insertion point
   '    nScale      Symbol scaling factor, 1 = No scaling
   '    nRotation   Symbol rotation about insertion point
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   'and
   '    The DRAFIX symbol library "C:\JOBST\JOBST.SLB" exists
   '
   'Note:-
   '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
   '
    PR_SetLayer "Notes"
    Print #fNum, "SetSymbolLibrary(sPathJOBST +"; QQ; "\\JOBST.SLB"; QQ; ");"
    Print #fNum, "Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ");"

    'PR_SetLayer slayer

    Print #fNum, "hEnt = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC
    Print #fNum, "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.y); CC
    Print #fNum, Str$(nScale); CC & Str$(nScale); CC; Str$(nRotation); ");"
   
End Sub

Sub PR_MakeXY(xyReturn As XY, X As Double, y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.X = X
    xyReturn.y = y
End Sub

Sub PR_NamedHandle(sHandleName As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to retain the entity handle of a previously
   'added entity.
   '
   'Assumes that hEnt is the entity handle to the just inserted entity.
   '
    Print #fNum, "HANDLE "; sHandleName; ";"
    Print #fNum, sHandleName; " = hEnt;"

End Sub

Sub PR_PutTapeLabel(nTape As Variant, xyStart As XY, nLength As Variant, nMM As Variant, nGrm As Variant, nRed As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to add Sleeve Tape details,
   'these details are given explicitly as arguments.
   'Where:-
   '    nTape       Index into sTextList below
   '    xyStart     Position of tape label on fold
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
    Dim sTextList As String, sSymbol As String
    Dim xyPt As XY, nSymbolOffSet As Single, nTextHt As Single

   'sTextList = "  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
    
    sSymbol = nTape & "tape"
    nSymbolOffSet = 0.6877
    nTextHt = 0.125

    PR_MakeXY xyPt, xyStart.X, xyStart.y + nSymbolOffSet   'Offset because symbol point is at top
    PR_InsertSymbol sSymbol, xyPt, 1, 0, "Notes"

    PR_CreateTapeLayer g_sFileNo, g_sSide, nTape

    PR_SetTextData 1, 32, -1, -1, 0

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

   'Draw Integer part
    PR_MakeXY xyPt, xyStart.X + 0.0625, xyStart.y + 0.75
    PR_DrawText Trim(nInt), xyPt, nTextHt

   'Draw eights part
    PR_MakeXY xyPt, xyStart.X + 0.0625 + (Len(Trim(nInt)) * nTextHt * 0.8), xyStart.y + 0.75 + nTextHt / 1.5
    If nDec <> 0 Then PR_DrawText Trim(nDec), xyPt, nTextHt / 1.5

   'MMs text
    PR_MakeXY xyPt, xyStart.X + 0.0625, xyStart.y + 1
    PR_DrawText nMM & "mm", xyPt, nTextHt

   'Grams text
    PR_MakeXY xyPt, xyStart.X + 0.0625, xyStart.y + 1.25
    PR_DrawText nGrm & "gm", xyPt, nTextHt

   'Reduction text and circle round the text
    PR_SetTextData 2, 16, -1, -1, -1
    PR_MakeXY xyPt, xyStart.X + 0.25, xyStart.y + 1.625
    PR_DrawText Trim(nRed), xyPt, nTextHt
    PR_DrawCircle xyPt, 0.125


    
End Sub

Sub PR_RightThumbHole(xyCen As XY, xyArcStart As XY, xyArcEnd As XY)
' This Calculates and Draws the curve at the
' right of the ThumbHole on the template

Dim nRad, nStartAng, nEndAng, nDeltaAng
    nRad = FN_CalcLength(xyCen, xyArcStart)
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
    nDeltaAng = nEndAng - nStartAng
    nDeltaAng = -nDeltaAng
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
End Sub

Sub PR_SetLayer(sNewLayer As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the current LAYER.
   'For this to work it assumes that hLayer is defined in DRAFIX as
   'a HANDLE.
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_sCurrentLayer are globals initialised by FN_Open
   '
   'To reduce unessesary writing of DRAFIX code check that the new layer
   'is different from the Current layer, change only if it is different.
   '

   ' If g_sCurrentLayer = sNewLayer Then Exit Sub
    'g_sCurrentLayer = sNewLayer
    
    Print #fNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; sNewLayer; QQ; ");"
    Print #fNum, "if ( hLayer != %badtable)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

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

Function round(ByVal nNumber As Single) As Integer
   'Fuction to return the rounded value of a decimal number
   'E.G.
   '    round(1.35) = 1
   '    round(1.55) = 2
   '    round(2.50) = 3
   '    round(-2.50) = -3
   '

    Dim nInt As Integer, nSign As Integer
    
    nSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    nInt = Int(nNumber)
    If (nNumber - nInt) >= 0.5 Then
        round = (nInt + 1) * nSign
    Else
        round = nInt * nSign
    End If
    
End Function


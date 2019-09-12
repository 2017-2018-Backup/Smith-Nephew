Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Public Module VESTARM
    'XY data type to represent points
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
    Public Structure BiArcVest
        Dim xyStart As XY
        Dim xyTangent As XY
        Dim xyEnd As XY
        Dim xyR1 As XY
        Dim xyR2 As XY
        Dim nR1 As Double
        Dim nR2 As Double
    End Structure

    'PI as a Global Constant
    Public Const PI As Double = 3.141592654

    Public MainForm As VESTARM_frm

    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public fNum As Object 'Macro file number
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote

    'Constants to define the drive and the root directory
    Public g_sFileNo As String 'The patients file no
    Public g_sSide As String 'The side Left or right
    Public g_sPatient As String 'The patients name
    Public g_sVestFlapLength As String
    Public g_sFabric As String
    Public g_svestFabnam As String
    Public g_sVestPressureChange As String
    Public g_sVestFabricChange As String
    Public g_sVestWpleat1 As String
    Public g_sVestWpleat2 As String
    Public g_sVestSpleat1 As String
    Public g_sVestSpleat2 As String
    Public g_vestsFlapStrap As String
    Public g_sVestFlapChk As Object
    Public g_sVestinchflag As String
    Public g_svestFlap As String
    Public g_sVestMM As String
    Public g_sVestGaunt As String
    Public g_sVestDetGaunt As String
    Public g_sVestNoThumb As String
    Public g_sVestPalmNo As String
    Public g_sVestWristNo As String
    Public g_sVestPalmWristDist As String
    Public g_sVestThumbCircum As String
    Public g_sVestThumbLength As String
    Public g_sVestEnclosedThumb As String
    Public g_sVestSecondLastTape As String
    Public g_sVestSecondTape As String
    Public g_sVestFirstTape As String
    Public g_sVestLastTape As String
    Public g_sVestModulus As String
    Public g_sVestStump As String
    Public g_sVestTapeLengths As String
    Public g_sVestTapeMMs As String
    Public g_sVestGrams As String
    Public g_sVestReduction As String
    Public g_sVestAmm As Object
    Public g_sVestBmm As Object
    Public g_sVestCmm As Object
    Public g_sVestDmm As Object
    Public g_sVestWaistCir As Object

    Public g_sVestWorkOrder As String

    'Store current layer and text setings to reduce DRAFIX code
    'this value is checked in PR_SetVestLayer
    Public g_sCurrentLayer As String
    Public g_nCurrTextHt As Object
    Public g_nCurrTextAspect As Object
    Public g_nCurrTextHorizJust As Object
    Public g_nCurrTextVertJust As Object
    Public g_nCurrTextFont As Object
    Public g_nCurrTextAngle As Object

    Public g_svestHoleCheck As String
    Public g_vestsOld As String

    Public g_sID As String
    Public g_vestsEditLengths As String
    Public g_sUnits As String


    Public g_iStyleFirstTape As Short
    Public g_iStyleLastTape As Short
    Public g_iFirstTape As Short
    Public g_iLastTape As Short
    Public g_nFrontStrapLength As Double
    Public g_nGauntletExtension As Double

    Public g_Modified As Short
    Public g_sPathJOBST As String

    Public g_bDrawVestMesh As Short
    Public g_bDrawBodyMesh As Short

    Public Detachable, WristNo, PalmNo, ThumLen, EnclosedThm As Object
    Public xyVestInsertion As XY

    Public g_iRightStyleFirstTape As Short
    Public g_iRightStyleLastTape As Short
    Public g_iRightFirstTape As Short
    Public g_iRightLastTape As Short
    Public g_sVestRightModulus As String
    Public g_sVestRightAmm As Object
    Public g_sVestRightBmm As Object
    Public g_sVestRightCmm As Object
    Public g_sVestRightDmm As Object
    Public g_sVestRightHoleCheck As String
    Public g_sVestRightPressureChange As String
    Dim xyElbow As XY

    Sub vestDataBaseDataUpDate(ByRef sType As String)
        Dim sSymbol As Object
        Dim sString As String
        Dim sGauntlet, sFlap As String
        Dim sBoxType As String
        Dim nYoffset As Single

        Return '11-June-2018
        Dim _frmarmdia As New armdia()      'Choose data box depending on type
        If _frmarmdia.txtType.Text = "ARM" Then
            sBoxType = "armarm"
        Else
            sBoxType = "vestarm"
        End If

        'Get sleevebox handle
        PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
        PrintLine(fNum, "hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = '" & sBoxType & "' AND DB Sleeve ='" & g_sSide & "'" & QQ & ");")
        PrintLine(fNum, "ResetSelection(hChan);")
        PrintLine(fNum, "hSleeve = GetNextSelection(hChan);")
        PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")

        'If sleevebox not found then insert one
        PrintLine(fNum, "if(!hSleeve){")
        PrintLine(fNum, "    Close(" & QQ & "selection" & QC & "hChan);")
        PrintLine(fNum, "    hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = 'mainpatientdetails'" & QQ & ");")
        PrintLine(fNum, "    ResetSelection(hChan);")
        PrintLine(fNum, "    hTitle = GetNextSelection(hChan);")
        PrintLine(fNum, "    Close(" & QQ & "selection" & QC & "hChan);")
        'Get title box origin
        PrintLine(fNum, "    GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);")

        'Insert arm or sleeve box
        PrintLine(fNum, "    if ( !Symbol(" & QQ & "find" & QCQ & sBoxType & QQ & ")) Exit(%cancel," & QQ & "Cant find SLEEVEBOX or ARMBOX symbol to insert|nCheck your installation, that JOBST.SLB exists" & QQ & ");")
        PrintLine(fNum, "    if(StringCompare(" & QQ & "Left" & QCQ & g_sSide & QQ & ")) xyTitleOrigin.x = xyTitleOrigin.x + 1.5;")
        PrintLine(fNum, "        else xyTitleOrigin.x = xyTitleOrigin.x + 3;")
        PrintLine(fNum, "    Execute(" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
        PrintLine(fNum, "    hSleeve = AddEntity(" & QQ & "symbol" & QCQ & sBoxType & QC & "xyTitleOrigin);")
        PrintLine(fNum, "}")

        'Insert Arm common
        sSymbol = "armcommon"
        If _frmarmdia.txtUidAC.Text = "" Then
            'Insert a new symbol
            PrintLine(fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
            '        Print #fNum, "if(!hSleeve){"
            PrintLine(fNum, "  Close(" & QQ & "selection" & QC & "hChan);")
            PrintLine(fNum, "  hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = 'mainpatientdetails'" & QQ & ");")
            PrintLine(fNum, "  ResetSelection(hChan);")
            PrintLine(fNum, "  hTitle = GetNextSelection(hChan);")
            PrintLine(fNum, "  Close(" & QQ & "selection" & QC & "hChan);")
            'Get title box origin
            PrintLine(fNum, "  GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);")
            PrintLine(fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
            PrintLine(fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyTitleOrigin.x, xyTitleOrigin.y);")
            PrintLine(fNum, "  }")
            PrintLine(fNum, "else")
            PrintLine(fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
        Else
            'Use existing symbol
            PrintLine(fNum, "hSym = UID (" & QQ & "find" & QC & Val(_frmarmdia.txtUidAC.Text) & ");")
            PrintLine(fNum, "if (!hSym) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
        End If
        'Update DB fields
        PrintLine(fNum, "SetDBData( hSym" & CQ & "Fabric" & QCQ & _frmarmdia.txtFabric.Text & QQ & ");")
        PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & _frmarmdia.txtFileNo.Text & QQ & ");")

        'Flap multiple field
        sString = New String(" ", 35)
        sString = LSet(g_svestFlap, Len(sString))
        If Val(_frmarmdia.txtCustFlapLength.Text) = 0 Then _frmarmdia.txtCustFlapLength.Text = CStr(-1)
        If Val(_frmarmdia.txtWaistCir.Text) = 0 Then _frmarmdia.txtWaistCir.Text = "-1"
        If Val(_frmarmdia.txtFrontStrapLength.Text) = 0 Then _frmarmdia.txtFrontStrapLength.Text = "-1"
        sFlap = sString & _frmarmdia.txtStrap.Text & " " & _frmarmdia.txtCustFlapLength.Text & " " & _frmarmdia.txtWaistCir.Text & " " & _frmarmdia.txtFrontStrapLength.Text

        'Gauntlet Multiple field
        If g_sVestGaunt = "False" Then
            sGauntlet = "0"
        Else
            'Set falgs
            sString = "1 "
            If g_sVestEnclosedThumb = "True" Then
                sString = sString & "1 "
            Else
                sString = sString & "0 "
            End If

            If g_sVestDetGaunt = "True" Then
                sString = sString & "1 "
            Else
                sString = sString & "0 "
            End If

            If g_sVestNoThumb = "True" Then
                sString = sString & "1 "
            Else
                sString = sString & "0 "
            End If
            If Val(_frmarmdia.txtGauntletExtension.Text) = 0 Then _frmarmdia.txtGauntletExtension.Text = "-1"
            'Data
            sGauntlet = sString & Val(g_sVestWristNo) & " " & Val(g_sVestPalmNo) & " " & Val(g_sVestThumbLength) & " " & Val(g_sVestThumbCircum) & " " & Val(g_sVestPalmWristDist) & " " & _frmarmdia.txtGauntletExtension.Text
        End If

        'Update data base fields
        '
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Sleeve" & QCQ & g_sSide & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Fileno" & QCQ & g_sFileNo & QQ & ");")

        'Wrist and Shoulder pleat multiple fields
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "WristPleat" & QCQ & g_sVestWpleat1 & " " & g_sVestWpleat2 & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "ShoulderPleat" & QCQ & g_sVestSpleat1 & " " & g_sVestSpleat2 & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Flap" & QCQ & sFlap & QQ & ");")

        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Gauntlet" & QCQ & sGauntlet & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "MM" & QCQ & g_sVestMM & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Modulus" & QCQ & g_sVestModulus & QQ & ");")

        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Fabric" & QCQ & g_sFabric & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Stump" & QCQ & g_sVestStump & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "TapeMMs" & QCQ & g_sVestTapeMMs & QQ & ");")

        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Reduction" & QCQ & g_sVestReduction & QQ & ");")
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Grams" & QCQ & g_sVestGrams & QQ & ");")

        'Tape lengths are universal, Store the actually used lengths with the
        'Profile Origin Marker
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "TapeLengths" & QCQ & g_sVestTapeLengths & QQ & ");")
        'Store ID of last drawn
        PrintLine(fNum, "SetDBData(hSleeve," & QQ & "ID" & QCQ & g_sID & QQ & ");")

        If sType = "Draw" Then
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Fabric" & QCQ & g_sFabric & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Gauntlet" & QCQ & sGauntlet & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "TapeLengths" & QCQ & g_vestsEditLengths & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "units" & QCQ & g_sUnits & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "WristPleat" & QCQ & g_sVestWpleat1 & " " & g_sVestWpleat2 & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "ShoulderPleat" & QCQ & g_sVestSpleat1 & " " & g_sVestSpleat2 & QQ & ");")

            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Flap" & QCQ & sFlap & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Grams" & QCQ & g_sVestGrams & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Reduction" & QCQ & g_sVestReduction & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "TapeMMs" & QCQ & g_sVestTapeMMs & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Stump" & QCQ & g_sVestStump & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Modulus" & QCQ & g_sVestModulus & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "MM" & QCQ & g_sVestMM & QQ & ");")
            PrintLine(fNum, "SetDBData(hOrigin," & QQ & "age" & QCQ & CType(_frmarmdia.Controls("txtAge"), Object).Text & QQ & ");")
        End If

    End Sub

    Function FN_CalcAngle(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, y As Object
        Dim rAngle As Double

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
        rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees
        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If
    End Function

    Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)

    End Function

    Function FN_EscapeQuotesInString(ByRef sAssignedString As Object) As String
        'Search through the string looking for " (double quote characater)
        'If found use \ (Backslash) to escape it
        '
        Dim ii As Short
        Dim Char_Renamed As String
        Dim sEscapedString As String
        FN_EscapeQuotesInString = ""

        For ii = 1 To Len(sAssignedString)

            Char_Renamed = Mid(sAssignedString, ii, 1)
            If Char_Renamed = """" Then
                sEscapedString = sEscapedString & "\" & Char_Renamed
            Else
                sEscapedString = sEscapedString & Char_Renamed
            End If
        Next ii

        FN_EscapeQuotesInString = sEscapedString

    End Function

    Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
        'Search through the string looking for " (double quote characater)
        'If found use \ (Backslash) to escape it
        '
        Dim ii As Short
        Dim Char_Renamed As String
        Dim sEscapedString As String

        FN_EscapeSlashesInString = ""

        For ii = 1 To Len(sAssignedString)
            Char_Renamed = Mid(sAssignedString, ii, 1)
            If Char_Renamed = "\" Then
                sEscapedString = sEscapedString & "\" & Char_Renamed
            Else
                sEscapedString = sEscapedString & Char_Renamed
            End If
        Next ii
        FN_EscapeSlashesInString = sEscapedString

    End Function

    Function FN_GetVestNumber(ByVal sString As String, ByRef iIndex As Short) As Double
        'Function to return as a numerical value the iIndexth item in a string
        'that uses blanks (spaces) as delimiters.
        'EG
        '    sString = "12.3 65.1 45"
        '    FN_GetVestNumber( sString, 2) = 65.1
        '
        'If the iIndexth item is not found then return -1 to indicate an error.
        'This assumes that the string will not be used to store -ve numbers.
        'Indexing starts from 1

        Dim ii, iPos As Short
        Dim sItem As String

        'Initial error checking
        sString = Trim(sString) 'Remove leading and trailing blanks
        If Len(sString) = 0 Then
            FN_GetVestNumber = -1
            Exit Function
        End If

        'Prepare string
        sString = sString & " " 'Trailing blank as stopper for last item
        'Get iIndexth item
        For ii = 1 To iIndex
            iPos = InStr(sString, " ")
            If ii = iIndex Then
                sString = Left(sString, iPos - 1)
                FN_GetVestNumber = Val(sString)
                Exit Function
            Else
                sString = LTrim(Mid(sString, iPos))
                If Len(sString) = 0 Then
                    FN_GetVestNumber = -1
                    Exit Function
                End If
            End If
        Next ii

        'The function should have exited befor this, however just in case
        '(iIndex = 0) we indicate an error,
        FN_GetVestNumber = -1
    End Function

    Function FN_Open(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object, ByRef sType As String) As Short
        'Open the DRAFIX macro file
        'Initialise Global variables
        Dim sID, sAxillaType As String
        Dim iPos As Double
        Dim sProximalStyle, sStyle, sDistalStyle As String

        'Open file
        fNum = FreeFile()
        FileOpen(fNum, sDrafixFile, Microsoft.VisualBasic.OpenMode.Output)
        FN_Open = fNum

        'Initialise String globals
        CC = Chr(44) 'The comma (,)
        NL = Chr(10) 'The new line character
        QQ = Chr(34) 'Double quotes (")
        QCQ = QQ & CC & QQ
        QC = QQ & CC
        CQ = CC & QQ
        Dim _frmarmdia As New armdia()
        'Initialise patient globals
        g_sFileNo = sPatientFile
        g_sSide = sLeftorRight
        g_sPatient = sName
        If CType(_frmarmdia.Controls("txtWorkOrder"), Object).Text = "" Then
            g_sVestWorkOrder = "-"
        Else
            g_sVestWorkOrder = CType(_frmarmdia.Controls("txtWorkOrder"), Object).Text
        End If


        'Globals to reduced drafix code written to file
        g_sCurrentLayer = ""
        g_nCurrTextHt = 0.125
        g_nCurrTextAspect = 0.6
        g_nCurrTextHorizJust = 1 'Left
        g_nCurrTextVertJust = 32 'Bottom
        g_nCurrTextFont = 0 'BLOCK
        g_nCurrTextAngle = 0

        'Create the 4 character string to identify the type
        PR_GetStyle(sStyle)
        'Write header information etc. to the DRAFIX macro file
        '
        PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(fNum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " Sleeve-" & g_sSide)
        PrintLine(fNum, "//by Visual Basic")

        'Define DRAFIX variables

        PrintLine(fNum, "HANDLE hLayer, hSleeve, hTitle, hChan, hEnt, hOrigin, hSym;")
        PrintLine(fNum, "XY     xyAxilla, xyAxillaLow, xyTitleOrigin, xyStart, xyTitleScale, xyOrigin, xyElbow;")
        PrintLine(fNum, "ANGLE  aTitleAngle;")
        PrintLine(fNum, "STRING sTitleName, sFileNo, sSleeve, sID, sWorkOrder, sName, sVestID, sData, sDate, sPathJOBST;")

        'Set path to JOBST installed directory
        PrintLine(fNum, "sPathJOBST = " & QQ & FN_EscapeSlashesInString(g_sPathJOBST) & QQ & ";")

        'Set up ID  data base field
        PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "ID" & QCQ & "string" & QQ & ");")

        'Text data
        PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & g_nCurrTextHorizJust & ");")
        PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & g_nCurrTextVertJust & ");")
        PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & g_nCurrTextHt & ");")
        PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & g_nCurrTextAspect & ");")
        PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & g_nCurrTextFont & ");")

        'Clear user selections etc
        PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
        PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
        g_sID = sStyle & g_sFileNo & g_sSide

        'Find axilla type
        If CType(_frmarmdia.Controls("txtType"), Object).Text <> "ARM" Then
            iPos = InStr(CType(_frmarmdia.Controls("txtVestRaglan"), Object).Text, QQ) 'Regular axilla are awkward as they
            'contain a space, look for quote (")
            'QQ is a constant set to quote (")
            If iPos <> 0 Then
                'Escape quote for use in drafix
                sAxillaType = Left(CType(_frmarmdia.Controls("txtVestRaglan"), Object).Text, iPos - 1) & "\" & QQ
            Else
                iPos = InStr(CType(_frmarmdia.Controls("txtVestRaglan"), Object).Text, " ")
                If iPos <> 0 Then
                    sAxillaType = Left(CType(_frmarmdia.Controls("txtVestRaglan"), Object).Text, iPos - 1)
                End If
            End If
        End If

        PrintLine(fNum, "sSleeve = " & QQ & g_sSide & QQ & ";")
        PrintLine(fNum, "sFileNo = " & QQ & g_sFileNo & QQ & ";")

        'Get Start point
        If sType = "Draw" Then
            'Get Start point
            PrintLine(fNum, "SetSymbolLibrary( sPathJOBST + " & QQ & "\\JOBST.SLB" & QQ & ");")
            PrintLine(fNum, "GetUser (" & QQ & "xy" & QCQ & "Indicate Start Point" & QC & "&xyStart);")

            'Place a marker at the start point for later use.
            'Get a UID and create the unique 4 character start to the ID code
            'Note this is a bit dogey if the drawing contains more than 9999 entities
            PR_SetVestLayer("Construct")
            PrintLine(fNum, "hOrigin = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart" & CC & "0.125);")
            PrintLine(fNum, "if (hOrigin) {")
            PrintLine(fNum, "  sData=StringMiddle(MakeString(" & QQ & "long" & QQ & ",UID(" & QQ & "get" & QQ & ",hOrigin)), 1, 4) ; ")
            PrintLine(fNum, "  while (StringLength(sData) < 4) sData = sData + " & QQ & " " & QQ & ";")
            PrintLine(fNum, "  sData = sData + sFileNo + sSleeve ;")

            PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "ID" & QQ & CC & QQ & g_sID & "originmark" & QQ & ");")
            PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "Data" & QQ & CC & "sData" & ");")
            PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "curvetype" & QQ & CC & QQ & "sleeveoriginmark" & QQ & ");")
            PrintLine(fNum, "  }")
        End If

        'Display Hour Glass symbol
        PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Drawing" & QQ & ");")

        'Set values for use futher on by other macros
        PrintLine(fNum, "xyOrigin = xyStart" & ";")
        PrintLine(fNum, "sID = " & QQ & g_sID & QQ & ";")
        PrintLine(fNum, "sName = " & QQ & g_sPatient & QQ & ";")
        PrintLine(fNum, "sWorkOrder = " & QQ & _frmarmdia.txtWorkOrder.Text & QQ & ";")
    End Function

    Function fnGetString(ByVal sString As String, ByRef iIndex As Short, ByRef sDelimiter As String) As String
        'Function to return as a string the iIndexth item in a string
        'that using the given string sDelimiter as the delimiter.
        'EG
        '    sString = "Sam Spade Hello"
        '    sDelimiter = " " {SPACE}
        '    fnGetNumber( sString, 2) = "Spade"
        '
        'If the iIndexth item is not found then return "" to indicate an error.
        'Indexing starts from 1

        Dim ii, iPos As Short
        Dim sItem As String

        'Initial error checking
        sString = Trim(sString) 'Remove leading and trailing blanks

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
                sString = Left(sString, iPos - 1)
                fnGetString = sString
                Exit Function
            Else
                sString = LTrim(Mid(sString, iPos + 1))
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

    Function maxVest(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the maximum of two numbers
        If nFirst >= nSecond Then
            maxVest = nFirst
        Else
            maxVest = nSecond
        End If
    End Function

    Sub PR_AddEntityID(ByRef sFileNo As String, ByRef sSide As String, ByRef sType As Object)
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
        Return '11-June-2018
        Dim sID As String
        sID = sFileNo & sSide & sType
        PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & "ID" & QQ & CC & QQ & sID & QQ & ");")
        PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & "Data" & QQ & CC & "sData" & ");")
    End Sub

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

    Sub PR_CreateTapeLayer(ByRef sFileNo As String, ByRef sSide As String, ByRef nTape As Object)
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
        slayer = LTrim(sFileNo & Mid(sSide, 1, 1) & nTape)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                     OpenMode.ForRead)

            'Dim sLayerName As String = "Center"
            If acLyrTbl.Has(slayer) = False Then
                Using acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                    '' Assign the layer the ACI color 3 and a name
                    acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                    acLyrTblRec.Name = slayer

                    '' Upgrade the Layer table for write
                    acLyrTbl.UpgradeOpen()

                    '' Append the new layer to the Layer table and the transaction
                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                End Using
            End If

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create a circle object
            Using acCirc As Circle = New Circle()
                acCirc.Center = New Point3d(2, 2, 0)
                acCirc.Radius = 1
                acCirc.Layer = slayer

                'acBlkTblRec.AppendEntity(acCirc)
                'acTrans.AddNewlyCreatedDBObject(acCirc, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_DeleteText(ByRef sFileNo As Object, ByRef sSleeve As Object, ByRef nTape As Object)
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '
        'Note:-
        '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
        '

        Return '11-June -2018

        PrintLine(fNum, "nTape = " & nTape & ";")
        PrintLine(fNum, "sFileNo = " & QQ & sFileNo & QQ & ";")
        PrintLine(fNum, "sSleeve = " & QQ & sSleeve & QQ & ";")
        PrintLine(fNum, "sLayer=" & QQ & "layer = '" & QQ & "+ sFileNo + StringMiddle(sSleeve, 1, 1) + MakeString(" & QQ & "long" & QQ & ", nTape) + " & QQ & "'" & QQ & ";")
        PrintLine(fNum, "hChan = Open(" & QQ & "selection" & QC & "sLayer);")

        PrintLine(fNum, "UserSelection(" & QQ & "clear" & QQ & ");")
        PrintLine(fNum, "if (hChan){")
        PrintLine(fNum, "ResetSelection(hChan);")
        PrintLine(fNum, "while(hEnt = GetNextSelection (hChan)) DeleteEntity(hEnt);")
        PrintLine(fNum, "}")
        PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
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
            Using acArc As Arc = New Arc(New Point3d(xyCen.X + xyVestInsertion.X, xyCen.y + xyVestInsertion.y, 0),
                                     nRad, nStartAng, nEndAng)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    ''Sub PR_DrawVestAssignDrafixVariable(ByRef sName As String, ByRef nValue As Double)
    Sub PR_DrawVestAssignDrafixVariable(ByRef nXValue As Double, ByRef nYValue As Double)
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

        ''Return '11-06-2018    PrintLine(fNum, sName & "=" & nValue & ";")
        xyElbow.X = nXValue
        xyElbow.y = nYValue
    End Sub

    Sub PR_DrawCircle(ByRef xyCen As XY, ByRef nRadius As Object)
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

            '' Create a circle that is at 2,3 with a radius of 4.25
            Using acCirc As Circle = New Circle()
                acCirc.Center = New Point3d(xyCen.X + xyVestInsertion.X, xyCen.y + xyVestInsertion.y, 0)
                acCirc.Radius = nRadius
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acCirc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acCirc)
                acTrans.AddNewlyCreatedDBObject(acCirc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

        'PrintLine(fNum, "hEnt = AddEntity(" & QQ & "circle" & QC & "xyStart.x+" & Str(xyCen.X) & CC & "xyStart.y+" & Str(xyCen.y) & CC & nRadius & ");")
    End Sub

    Sub PR_DrawVestCircularStump(ByRef xyStart As XY, ByRef nFiguredLength As Double, ByRef sAge As String)
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

        Dim xyPt As XY
        Dim nTextHt, nSeamAllowance As Single
        Dim nRadius As Single

        nTextHt = 0.125
        nSeamAllowance = 0.125 ' 1/8th Inch
        nRadius = (nFiguredLength / 3.14)

        PR_SetVestLayer("Template" & g_sSide)

        PR_DrawCircle(xyStart, nRadius + nSeamAllowance) 'Add seam allowance
        PR_AddEntityID(g_sFileNo, g_sSide, "CirStump") 'sFileNo, g_sSide from FN_Open

        PR_SetVestLayer("Notes")
        PR_DrawCircle(xyStart, nRadius)
        PR_AddEntityID(g_sFileNo, g_sSide, "CirStumpSeam") 'sFileNo, g_sSide from FN_Open

        PR_SetTextData(2, 16, -1, -1, -1) 'Horiz center, Vertical center
        ''-----PR_DrawText(g_sPatient & "\n" & g_sVestWorkOrder & "\n" & g_sSide, xyStart, nTextHt, 0)
        PR_DrawMText(g_sPatient & Chr(10) & sAge & Chr(10) & g_sSide, xyStart, True)

    End Sub

    Sub PR_DrawFitted(ByRef Profile As curve, Optional ByVal bIsSetXDATA As Boolean = False, Optional ByVal g_sbIsBody As Boolean = False)
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
        Dim ii As Short
        'Return '11-June-2018
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
                PR_DrawPoly(Profile)
            Case Else
                Dim ptColl As Point3dCollection = New Point3dCollection()
                For ii = 1 To Profile.n
                    ptColl.Add(New Point3d(Profile.X(ii) + xyVestInsertion.X, Profile.y(ii) + xyVestInsertion.y, 0))
                Next
                PR_DrawSpline(ptColl, bIsSetXDATA, g_sbIsBody)
        End Select
    End Sub

    Sub PR_DrawLine(ByRef xyStart As XY, ByRef xyFinish As XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
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
            Dim acLine As Line = New Line(New Point3d(xyStart.X + xyVestInsertion.X, xyStart.y + xyVestInsertion.y, 0),
                                    New Point3d(xyFinish.X + xyVestInsertion.X, xyFinish.y + xyVestInsertion.y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_DrawVestLineOffset(ByRef xyStart As XY, ByRef xyFinish As XY, ByRef nOffset As Double)
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
        Dim xyPt1, xyPt2 As XY

        nAngle = FN_CalcAngle(xyStart, xyFinish)
        If nOffset < 0 Then
            nAngle = nAngle - 90
        Else
            nAngle = nAngle + 90
        End If

        PR_CalcPolar(xyStart, nAngle, System.Math.Abs(nOffset), xyPt1)
        PR_CalcPolar(xyFinish, nAngle, System.Math.Abs(nOffset), xyPt2)
        PR_DrawLine(xyPt1, xyPt2)
    End Sub

    Sub PR_DrawVestMesh()
        'Invokes external Mesh drawing routine
        'As such it must be the last line called by the DRAW Macro
        If g_bDrawBodyMesh Then
            PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshdraw" & QCQ & "normal" & QQ & " );")
        End If
        If g_bDrawVestMesh Then
            PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshvest" & QCQ & "normal" & QQ & " );")
        End If
    End Sub

    Sub PR_DrawPoly(ByRef Profile As curve)
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
                    acPoly.AddVertexAt(ii - 1, New Point2d(Profile.X(ii) + xyVestInsertion.X, Profile.y(ii) + xyVestInsertion.y), 0, 0, 0)
                Next ii
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_DrawVestRaglan(ByRef sType As String, ByRef sVestRaglan As Object, ByRef sAge As Object,
                          ByRef sSleeve As String, ByRef Profile As curve, ByRef strAxillaType As String)
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
        Dim sAxillaType, sString As String
        Dim iPos As Short

        'Flag to draw the mesh used by PR_DrawMesh
        g_bDrawBodyMesh = False
        g_bDrawVestMesh = False


        'Find axilla type
        ''iPos = InStr(sVestRaglan, QQ) 'Regular axilla are awkward as they
        'iPos = InStr(sVestRaglan, " ")
        ''contain a space, look for quote (")
        ''QQ is a constant set to quote (")
        'If iPos <> 0 Then
        '    'Escape quote for use in drafix
        '    sAxillaType = Left(sVestRaglan, iPos - 1) & "\" & QQ
        '    sString = Mid(sVestRaglan, iPos + 1)
        'Else
        '    iPos = InStr(sVestRaglan, " ")
        '    sAxillaType = Left(sVestRaglan, iPos - 1)
        '    sString = Mid(sVestRaglan, iPos)
        'End If
        sAxillaType = strAxillaType
        sString = sVestRaglan

        'Check that an axilla has been given
        If sAxillaType = "None" Or sAxillaType = "Sleeveless" Then Exit Sub

        'NB the order of the following is very important
        '
        'Data from vest used in drawing the sleeve

        PrintLine(fNum, "nAxillaFrontNeckRad= " & FN_GetVestNumber(sString, 1) & ";")
        PrintLine(fNum, "nAxillaBackNeckRad= " & FN_GetVestNumber(sString, 2) & ";")
        PrintLine(fNum, "nShoulderToBackRaglan = " & FN_GetVestNumber(sString, 3) & ";")

        If sType = "Vest" Then
            'Initialise raglan drawing
            PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_INIT.D;")
            Dim xyLowerAxilla, xyProfileStart, xyOrigin, xyAxilla, xyRaglanNeck As XY
            xyOrigin = xyVestInsertion
            PR_MakeXY(xyAxilla, xyOrigin.X + 2.75, xyOrigin.y + 4.63125)
            PR_MakeXY(xyProfileStart, xyOrigin.X + 0, xyOrigin.y + 4.878125)
            ''PR_MakeXY(xyProfileEnd, xyOrigin.X + 1.375, xyOrigin.y + 4.285625)
            Dim ptCurveColl As Point3dCollection = New Point3dCollection
            'ptCurveColl.Add(New Point3d(xyProfileStart.X, xyProfileStart.y, 0))
            'ptCurveColl.Add(New Point3d(xyProfileEnd.X, xyProfileEnd.y, 0))
            'ptCurveColl.Add(New Point3d(xyAxilla.X, xyAxilla.y, 0))
            Dim ii As Double
            For ii = 1 To Profile.n
                ptCurveColl.Add(New Point3d(Profile.X(ii) + xyVestInsertion.X, Profile.y(ii) + xyVestInsertion.y, 0))
            Next
            PR_MakeXY(xyAxilla, Profile.X(Profile.n) + xyVestInsertion.X, Profile.y(Profile.n) + xyVestInsertion.y)
            PR_MakeXY(xyProfileStart, Profile.X(1) + xyVestInsertion.X, Profile.y(1) + xyVestInsertion.y)
            Dim nTape As Double = ptCurveColl.Count
            Dim sEntClass As String = ""
            Dim DrawSegment As Boolean
            'If (hSleeveProfile) Then
            '    ''// Establish proximal control points from Sleeve profile
            '    GetEntityClass(hSleeveProfile, & sEntClass)
            '    If (sEntClass.Equals("curve") Or sEntClass.Equals("polyline")) Then
            '        nTape = GetVertexCount(hSleeveProfile)
            '        GetVertex(hSleeveProfile, nTape, & xyAxilla)
            '    Else
            '        GetGeometry(hSleeveProfile, & xyProfileStart, & xyAxilla)
            '    End If

            '    xyLowerAxilla.y = xyOrigin.y
            '    xyLowerAxilla.x = xyAxilla.x
            'End If
            xyLowerAxilla.y = xyOrigin.y
            xyLowerAxilla.X = xyAxilla.X

            ''// Get 1/3 rd line down from upper proximal scale And intersect with 
            ''//
            Dim nSeam As Double = 0.1875
            Dim nLen As Double = (FN_CalcLength(xyLowerAxilla, xyAxilla) - nSeam) / 3
            Dim xyStart, xyEnd, xyInt, xyBackNeck, xyEndBottomCurve As XY
            PR_MakeXY(xyStart, xyAxilla.X, xyAxilla.y - nLen)
            PR_MakeXY(xyEnd, xyAxilla.X + 100, xyAxilla.y - nLen)
            Dim nAxillaBackNeckRad As Double = FN_GetVestNumber(sString, 2)
            Dim nShoulderToBackRaglan As Double = FN_GetVestNumber(sString, 3)
            If (sSleeve.Equals("Right")) Then
                If FN_GetVestNumber(sString, 5) > 0 Then
                    nAxillaBackNeckRad = FN_GetVestNumber(sString, 5)
                End If
                If FN_GetVestNumber(sString, 6) > 0 Then
                    nShoulderToBackRaglan = FN_GetVestNumber(sString, 6)
                End If
            End If

            If (FN_CirLinInt(xyStart, xyEnd, xyAxilla, nAxillaBackNeckRad, xyInt)) Then
                xyBackNeck = xyInt
                xyEndBottomCurve.y = xyBackNeck.y - nShoulderToBackRaglan
                xyEndBottomCurve.X = xyBackNeck.X
            Else
                MsgBox("Can't form raglan curve with this sleeve data", 16, "Vest Arm Dialog")
                Exit Sub
            End If
            ''// set layer
            If (sSleeve.Equals("Right")) Then
                PR_SetLayer("TemplateRight")
            Else
                PR_SetLayer("TemplateLeft")
            End If
            Dim nLowerCurveRadius As Double = FN_CalcLength(xyLowerAxilla, xyEndBottomCurve)
            Dim aUpperCurve As Double = 100000    ''// Impossible value used To test For non-intersecton
            Dim aLowerCurve As Double = 100000
            Dim aPrevAngle As Double = 0
            Dim xyPt1, xyPt2, xyCenter As XY
            xyPt1.X = 0
            xyPt1.y = 0
            Dim hCurve As Object
            hCurve = FreeFile()
            FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
            Dim sLine As String = ""
            Dim aAngle, nLength As Double
            While Not EOF(hCurve)
                sLine = LineInput(hCurve)
                ''FN_GetNumber(sLine, nLength, aAngle)
                nLength = FN_GetNumber(sLine, 1)
                aAngle = FN_GetNumber(sLine, 2)
                aAngle = aAngle + aPrevAngle
                ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                If (FN_CirLinInt(xyPt1, xyPt2, xyCenter, nAxillaBackNeckRad, xyInt)) Then
                    aUpperCurve = FN_CalcAngle(xyCenter, xyInt)
                End If
                If (FN_CirLinInt(xyPt1, xyPt2, xyCenter, nLowerCurveRadius, xyInt)) Then
                    aLowerCurve = FN_CalcAngle(xyCenter, xyInt)
                End If
                xyPt1 = xyPt2
                aPrevAngle = aAngle
                ''ScanLine(sLine, "blank", nLength, aAngle)
            End While
            If (aUpperCurve = 100000 Or aLowerCurve = 100000) Then
                MsgBox("Sleeve drawing error" & Chr(10) & "Can't make Ragalan curve with this data", 16, "Vest Arm Dialog")
            End If
            FileClose(hCurve)

            'Draw for the particular axilla type
            PrintLine(fNum, "sAxillaType = " & QQ & sAxillaType & QQ & ";")
            Dim _frmarmdia As New armdia()
            Dim nSeamAllowance As Double = 0.125
            Dim nRadius, nCount As Double
            Dim nAxillaFrontNeckRad As Double = FN_GetVestNumber(sString, 1)
            If (sSleeve.Equals("Right")) Then
                If FN_GetVestNumber(sString, 4) > 0 Then
                    nAxillaFrontNeckRad = FN_GetVestNumber(sString, 4)
                End If
            End If
            Dim nFrontNeckRedFac As Double = 0.125          ''// GOP 01-02/18, 15.5
            Select Case sAxillaType
                Case "Open", "Lining"
                    PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_OPEN.D;")
                    Dim nLengthToOpenAxillaPt As Double = 0.0
                    ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                    ''GetLine(hCurve, & sLine)
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    hCurve = FreeFile()
                    FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                    aPrevAngle = (FN_CalcAngle(xyAxilla, xyBackNeck) - aUpperCurve)    ''// Rotate curve To correct start angle
                    xyPt1 = xyAxilla
                    DrawSegment = True                  ''// Set draw segments flag on
                    Dim ptOpenColl As Point3dCollection = New Point3dCollection
                    ''StartPoly("polyline")
                    ''AddVertex(xyPt1)
                    ptOpenColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
                    nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                    Dim xyOpenCen As XY = xyAxilla
                    nCount = 1
                    While Not EOF(hCurve)
                        sLine = LineInput(hCurve)
                        nLength = FN_GetNumber(sLine, 1)
                        aAngle = FN_GetNumber(sLine, 2)
                        aAngle = aAngle + aPrevAngle
                        ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                        PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                        If (FN_CirLinInt(xyPt1, xyPt2, xyOpenCen, nRadius, xyInt)) Then
                            If (nCount = 1) Then
                                xyRaglanNeck = xyInt    ''// Found Raglan End
                                nRadius = nAxillaBackNeckRad
                                ''// Check if it ends on same segment
                                If (FN_CirLinInt(xyPt1, xyPt2, xyOpenCen, nRadius, xyInt)) Then nCount = 2
                            End If
                            If (nCount = 2) Then
                                nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + FN_CalcLength(xyPt1, xyInt)
                                ''AddVertex(xyInt)
                                ''EndPoly()
                                ptOpenColl.Add(New Point3d(xyInt.X, xyInt.y, 0))
                                'hEnt = UID("find", UID("getmax"))
                                'SetDBData(hEnt, "ID", sID)
                                'SetDBData(hEnt, "curvetype", "sleeveraglan")
                                'SetDBData(hEnt, "Data", sData)
                                Exit While
                            End If
                            nCount = nCount + 1
                        End If
                        If (DrawSegment) Then
                            ''AddVertex(xyPt2)
                            ptOpenColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                        End If
                        nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + FN_CalcLength(xyPt1, xyPt2)
                        xyPt1 = xyPt2
                        aPrevAngle = aAngle
                        ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    End While
                    FileClose(hCurve)
                    PR_DrawPoly(ptOpenColl)
                    PR_SetLayer("Notes")
                    Dim nTraversedLength As Double
                    Dim nError1 As Integer
                    If (sAxillaType.Equals("Open")) Then
                        ''// Locate 1/3 raglan length And add an arrow at that point
                        ''// Label axilla as open
                        nLengthToOpenAxillaPt = nLengthToOpenAxillaPt / 3
                        nTraversedLength = 0.0
                        ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                        ''GetLine(hCurve, & sLine)
                        ''ScanLine(sLine, "blank", & nLength, & aAngle)
                        hCurve = FreeFile()
                        FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                        aPrevAngle = (FN_CalcAngle(xyAxilla, xyBackNeck) - aUpperCurve)    ''// Rotate curve To correct start angle
                        xyPt1 = xyAxilla
                        While Not EOF(hCurve)
                            sLine = LineInput(hCurve)
                            nLength = FN_GetNumber(sLine, 1)
                            aAngle = FN_GetNumber(sLine, 2)
                            aAngle = aAngle + aPrevAngle
                            ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                            PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                            nTraversedLength = nTraversedLength + nLength
                            If (nTraversedLength > nLengthToOpenAxillaPt) Then
                                nError1 = FN_CirLinInt(xyPt1, xyPt2, xyPt1, nTraversedLength - nLengthToOpenAxillaPt, xyInt)
                                ''hEnt = AddEntity("marker", "medarrow", xyInt, 0.2, 0.1, FN_CalcAngle(xyPt1, xyPt2) - 90)
                                'SetDBData(hEnt, "ID", sID)
                                'SetDBData(hEnt, "curvetype", "sleeveopenaxillamarker")
                                'SetDBData(hEnt, "Data", sData)
                                ''// Add text
                                ''AddEntity("text","OPEN",xyAxilla.X,xyAxilla.y - 0.5)
                                PR_MakeXY(xyStart, xyAxilla.X, xyAxilla.y - 0.5)
                                PR_DrawVestText("OPEN", xyStart, 0.1, 0)
                                ''AddEntity("text", "AXILLA", xyAxilla.X, xyAxilla.y - 0.7)
                                PR_MakeXY(xyStart, xyAxilla.X, xyAxilla.y - 0.7)
                                PR_DrawVestText("AXILLA", xyStart, 0.1, 0)
                                Exit While
                            End If
                            xyPt1 = xyPt2
                            aPrevAngle = aAngle
                            ''ScanLine(sLine, "blank", & nLength, & aAngle)
                        End While
                        FileClose(hCurve)

                    Else
                        ''AddEntity("text","LINING",xyAxilla.X,xyAxilla.y - 0.5)
                        PR_MakeXY(xyStart, xyAxilla.X, xyAxilla.y - 0.5)
                        PR_DrawVestText("LINING", xyStart, 0.1, 0)
                    End If
                Case "Mesh"
                    g_bDrawVestMesh = True

                    PrintLine(fNum, "nAge = " & Val(sAge) & ";")
                    'If g_sSide = "Left" Then
                    '    PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
                    '    PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
                    'Else
                    '    If Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 4, ",")) > 0 And Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 3, ",")) > 0 Then
                    '        PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 4, ",")) & ";")
                    '        PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 3, ",")) & ";")
                    '    Else
                    '        PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
                    '        PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
                    '    End If
                    'End If

                    PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_MESH.D;")

                    ''// From the Axilla point traverse the VESTCURV picking up the following intersections
                    ''// in the order given (ie Increasing X)
                    ''//
                    ''//	1. Start of ragalan curve (1/8th) back from next point
                    ''//	2. Intersection of end of Mesh symbol And raglan curve, defines symbol angle
                    Dim nAge As Integer = Val(sAge)
                    Dim nMeshAxillaFac, nMeshSymbolLength, nMeshLength, nDistanceAlongRaglan As Double
                    Dim nMeshAxillaGussetFac As Double = 0.75
                    Dim nMSLengthGusset As Double = 2.1875
                    Dim nMeshAxillaGussetBoysFac As Double = 0.8125
                    Dim nMSLengthGussetBoys As Double = 2.375
                    Dim sMeshSymbol As String = ""
                    If (nAge <= 14) Then
                        sMeshSymbol = "mesh0to14"
                        nMeshAxillaFac = nMeshAxillaGussetFac
                        nMeshSymbolLength = nMSLengthGusset
                        nMeshLength = 2.58
                    Else
                        sMeshSymbol = "mesh15andUP"
                        nMeshAxillaFac = nMeshAxillaGussetBoysFac
                        nMeshSymbolLength = nMSLengthGussetBoys
                        nMeshLength = 2.8
                    End If
                    ''// Find point on sleeve profile (note - this Is an aproximation only)
                    ''// From above nTape Is currenty set to last vertex on the profile
                    Dim xyAxillaLow1, xyCen1, xyRaglanAxilla1, xyMeshAtSeamAllowance As XY
                    If (sEntClass.Equals("curve") Or sEntClass.Equals("polyline")) Then
                        nTape = nTape - 1
                        xyPt1 = xyAxilla
                        While (nTape > 0)
                            ''GetVertex(hSleeveProfile, nTape, & xyPt2)
                            PR_MakeXY(xyPt2, ptCurveColl(nTape).X, ptCurveColl(nTape).Y)
                            If (FN_CirLinInt(xyPt1, xyPt2, xyAxilla, nMeshAxillaFac, xyInt)) Then
                                xyAxillaLow1 = xyInt
                                Exit While
                            End If
                            nTape = nTape - 1
                            xyPt1 = xyPt2
                        End While
                    Else
                        If (FN_CirLinInt(xyProfileStart, xyAxilla, xyAxilla, nMeshAxillaFac, xyInt)) Then
                            xyAxillaLow1 = xyInt
                        End If
                    End If
                    ''// Open Curve for Reading
                    ''// Assume that if the tests for Existance/Corruption are passed. Then no need to repeat them

                    ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                    ''GetLine(hCurve, & sLine)
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    hCurve = FreeFile()
                    FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                    aPrevAngle = (FN_CalcAngle(xyAxilla, xyBackNeck) - aUpperCurve)
                    xyCen1 = xyAxillaLow1
                    nRadius = nMeshSymbolLength - nSeamAllowance
                    nCount = 1
                    xyPt1 = xyAxilla
                    ''//   DrawSegment = %false 		// Set draw segments flag off
                    DrawSegment = True      ''// Draw entire curve for the moment
                    Dim ptMeshPoly As Point3dCollection = New Point3dCollection
                    ''StartPoly("polyline")
                    ''AddVertex(xyAxilla)
                    ptMeshPoly.Add(New Point3d(xyAxilla.X, xyAxilla.y, 0))
                    nDistanceAlongRaglan = 0
                    While Not EOF(hCurve)
                        sLine = LineInput(hCurve)
                        nLength = FN_GetNumber(sLine, 1)
                        aAngle = FN_GetNumber(sLine, 2)
                        aAngle = aAngle + aPrevAngle
                        ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                        PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                        If (DrawSegment = False) Then
                            nDistanceAlongRaglan = nDistanceAlongRaglan + nLength
                        End If
                        If (FN_CirLinInt(xyPt1, xyPt2, xyCen1, nRadius, xyInt)) Then
                            If (nCount = 1) Then
                                DrawSegment = True      ''// Draw from seam allowance
                                xyRaglanAxilla1 = xyInt
                                nRadius = nMeshSymbolLength
                                ''// Check if seam allowance on same segment
                                If (FN_CirLinInt(xyPt1, xyPt2, xyCen1, nRadius, xyInt)) Then
                                    nCount = 2
                                End If
                                xyPt1 = xyRaglanAxilla1
                                nDistanceAlongRaglan = nDistanceAlongRaglan + FN_CalcLength(xyPt1, xyPt2)
                                ''//			StartPoly("polyline");
                                ''//			AddVertex(xyRaglanAxilla); 		
                            End If
                            If (nCount = 2) Then
                                xyMeshAtSeamAllowance = xyInt   ''// Mesh symbol End For angle
                                xyCen1 = xyAxilla
                                nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                            End If
                            If (nCount = 3) Then
                                xyRaglanNeck = xyInt        ''// Found Raglan End
                                xyCen1 = xyAxilla
                                nRadius = nAxillaBackNeckRad
                                ''// Check if it ends on same segment
                                If (FN_CirLinInt(xyPt1, xyPt2, xyCen1, nRadius, xyInt)) Then
                                    nCount = 4
                                End If
                            End If
                            If (nCount = 4) Then
                                ''AddVertex(xyBackNeck)
                                ''EndPoly()
                                ptMeshPoly.Add(New Point3d(xyBackNeck.X, xyBackNeck.y, 0))
                                'hEnt = UID("find", UID("getmax"))
                                'SetDBData(hEnt, "ID", sID)
                                'SetDBData(hEnt, "curvetype", "sleeveraglan")
                                'SetDBData(hEnt, "Data", sData)
                                Exit While
                            End If
                            nCount = nCount + 1
                        End If

                        If (DrawSegment) Then
                            ''AddVertex(xyPt2)      ''// Draw raglan
                            ptMeshPoly.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                        End If
                        xyPt1 = xyPt2
                        aPrevAngle = aAngle
                        ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    End While
                    FileClose(hCurve)
                    PR_DrawPoly(ptMeshPoly)

                    ''// Insert Mesh axilla symbol at correct angle 
                    aAngle = FN_CalcAngle(xyAxillaLow1, xyMeshAtSeamAllowance)

                    ''// Insert  Text
                    PR_SetLayer("Notes")
                    ''SetData("TextHorzJust", 4 )
                    ''SetData("TextVertJust", 32 )
                    If (nAge <= 14) Then
                        ''AddEntity("text", "1-3/4\" + "GUSSET", xyAxillaLow1.X - 0.25, xyAxillaLow1.y - 0.65)
                        PR_MakeXY(xyStart, xyAxillaLow1.X - 0.25, xyAxillaLow1.y - 0.65)
                        PR_DrawVestText("1-3/4" + QQ + "GUSSET", xyStart, 0.1, 0)
                    Else
                        ''AddEntity("text", "BOYS GUSSET", xyAxillaLow1.X - 0.25, xyAxillaLow1.y - 0.65)
                        PR_MakeXY(xyStart, xyAxillaLow1.X - 0.25, xyAxillaLow1.y - 0.65)
                        PR_DrawVestText("BOYS GUSSET", xyStart, 0.1, 0)
                    End If
                    ''// Revised code to write a datafile and start the MESHDRAW
                    ''// 
                    ''hCurve = Open("file", "C:\\JOBST\\MESHDRAW.DAT", "write")
                    hCurve = FreeFile()
                    FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\MESHDRAW.DAT", VB.OpenMode.Output)
                    ''SetData("UnitLinearType", 0)    ''// "Inches"
                    PrintLine(hCurve, "sleevemesh", "\n")
                    PrintLine(hCurve, Str(xyAxillaLow1.X) & " " & Str(xyAxillaLow1.y) & Chr(10))
                    PrintLine(hCurve, Str(xyAxilla.X) & " " & Str(xyAxilla.y) & Chr(10))
                    PrintLine(hCurve, Str(xyAxilla.X) & " " & Str(xyAxilla.y) & Chr(10))
                    PrintLine(hCurve, Str(xyBackNeck.X) & " " & Str(xyBackNeck.y) & Chr(10))
                    PrintLine(hCurve, Str(nMeshLength) & Chr(10))
                    PrintLine(hCurve, Str(nDistanceAlongRaglan) & Chr(10))
                    PrintLine(hCurve, sSleeve & Chr(10))
                    ''PrintLine(hCurve, sData, "\n")
                    ''SetData("UnitLinearType", 6)    ''// "Inches/Fraction"
                    FileClose(hCurve)
                    ''------
                    PrintLine(fNum, "nMeshLength = " & Str(nMeshLength) & ";")
                    PrintLine(fNum, "nDistanceAlongRaglan = " & Str(nDistanceAlongRaglan) & ";")
                    ''-------
                    Dim xyRaglanStart As XY = xyAxilla
                    Dim xyRaglanEnd As XY = xyBackNeck
                    Dim xyMeshStart As XY = xyAxillaLow1
                    Dim xyMarker As XY
                    PR_CalcPolar(xyMeshStart, aAngle, 0.125, xyMarker)
                    PR_SetLayer("Construct")
                    PR_DrawVestXMarker(xyMarker)
                    Dim MeshSeam, MeshProfile As BiArcVest
                    Dim iError As Integer = FN_CalcAxillaMesh(xyRaglanStart, xyRaglanEnd, xyProfileStart, False, nMeshLength, nDistanceAlongRaglan, MeshSeam, MeshProfile)
                    If iError = True Then
                        'Draw the mesh
                        PR_SetLayer("Template" & sSleeve)
                        PR_DrawArc(MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent)
                        PR_DrawArc(MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd)
                        Dim xyTmp As XY
                        iError = FN_CirLinInt(xyMeshStart, xyAxilla, MeshProfile.xyR1, MeshProfile.nR1, xyTmp)
                        PR_DrawVestLine(MeshProfile.xyStart, xyTmp)
                        'Draw closing line
                        PR_DrawVestLine(xyRaglanStart, xyAxilla)

                        PR_SetLayer("Notes")
                        PR_DrawArc(MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent)
                        PR_DrawArc(MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd)
                        PR_CalcPolar(xyMeshStart, FN_CalcAngle(xyAxilla, xyMeshStart), 3, xyTmp)
                        iError = FN_CirLinInt(xyTmp, xyAxilla, MeshSeam.xyR1, MeshSeam.nR1, xyTmp)
                        PR_DrawVestLine(MeshSeam.xyStart, xyTmp)
                        PR_SetLayer("Construct")
                        PR_DrawVestXMarker(MeshSeam.xyEnd)
                    Else
                        MsgBox("Failure to Calculate mesh. Draw manually and contact supervisor", 48, "Vest Arm - Dialog")
                        ''PR_DrawVestXMarker(MeshSeam.xyEnd) 'Bad practice here (using a side effect)
                        'but I know that this point
                        'is at the nDistanceAlongRaglan
                    End If
                    ''// End  MESH Axilla

                Case Else 'Regular
                    PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_REGLR.D;")
                    Dim nRegAxillaNormal As Double = 2.0
                    Dim nRegAxillaChild As Double = 1.5
                    Dim nRegAxillaLargeAdult As Double = 2.5
                    Dim nRegAxillaFac_1 As Double = 1.5             ''// GOP 01-02/18, 16.1
                    Dim nRegAxillaFac_2 As Double = 0.75            ''// GOP 01-02/18, 16.3

                    If (sAxillaType.Contains("Regular 2" & QQ)) Then
                        nRegAxillaFac_1 = nRegAxillaNormal
                    End If
                    If (sAxillaType.Contains("Regular 1½" & QQ)) Then
                        nRegAxillaFac_1 = nRegAxillaChild
                    End If
                    If (sAxillaType.Contains("Regular 2½" & QQ)) Then
                        nRegAxillaFac_1 = nRegAxillaLargeAdult
                    End If

                    Dim nXInsert, nYInsert As Double
                    ''nXInsert = -10
                    nXInsert = 10
                    nYInsert = -3.5
                    Dim xyCen As XY = xyAxilla

                    ''// Open Curve for Reading
                    hCurve = FreeFile()
                    FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                    ''GetLine(hCurve, & sLine)
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)

                    aPrevAngle = (FN_CalcAngle(xyAxilla, xyBackNeck) - aUpperCurve)      ''// Rotate curve To correct start angle

                    xyPt1 = xyAxilla
                    DrawSegment = False         ''// Set draw segments flag off
                    Dim DrawInsert As Boolean = False
                    nRadius = nSeamAllowance
                    nCount = 1
                    Dim xyInsertConstruct_3, xyRaglanAxilla As XY
                    Dim ptColl As Point3dCollection = New Point3dCollection
                    While Not EOF(hCurve)
                        sLine = LineInput(hCurve)
                        nLength = FN_GetNumber(sLine, 1)
                        aAngle = FN_GetNumber(sLine, 2)
                        aAngle = aAngle + aPrevAngle
                        ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                        PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                        If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then
                            If (nCount = 1) Then
                                DrawInsert = True
                                xyPt1 = xyInt           ''// Found Construction point
                                xyInsertConstruct_3 = xyInt
                                nRadius = nRegAxillaFac_1
                            End If
                            If (nCount = 2) Then
                                xyCen = xyInt           ''// Found Construction point
                                nRadius = nRegAxillaFac_2
                            End If
                            If (nCount = 3) Then
                                xyCen = xyAxilla            ''// Found raglan Start
                                nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                                xyRaglanAxilla = xyInt
                                '     hEnt = AddEntity("line", TransXY(xyPt1, "trs", nXInsert, nYInsert),
                                'TransXY(xyInt, "trs", nXInsert, nYInsert)) ''// Draw insert
                                PR_MakeXY(xyStart, xyPt1.X + nXInsert, xyPt1.y + nYInsert)
                                PR_MakeXY(xyEnd, xyInt.X + nXInsert, xyInt.y + nYInsert)
                                ''PR_DrawLine(xyStart, xyEnd)
                                PR_DrawVestLine(xyStart, xyEnd)
                                xyPt1 = xyInt
                                'SetDBData(hEnt, "ID", sID)
                                'SetDBData(hEnt, "curvetype", "sleeveinsert")
                                'SetDBData(hEnt, "Data", sData)
                                DrawSegment = True      ''// Start Drawing from here
                                ''StartPoly("polyline")
                                ''AddVertex(xyPt1)
                                ptColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
                                DrawInsert = False
                            End If
                            If (nCount = 4) Then
                                xyRaglanNeck = xyInt        ''// Found Raglan End
                                xyCen = xyAxilla
                                nRadius = nAxillaBackNeckRad
                                ''// Check if it ends on same segment
                                If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then
                                    nCount = 5
                                End If
                            End If
                            If (nCount = 5) Then
                                ''AddVertex(xyBackNeck)
                                ptColl.Add(New Point3d(xyBackNeck.X, xyBackNeck.y, 0))
                                ''EndPoly()
                                'hEnt = UID("find", UID("getmax"))
                                'SetDBData(hEnt, "curvetype", "sleeveraglan")
                                'SetDBData(hEnt, "Data", sData)
                                Exit While
                            End If
                            nCount = nCount + 1
                        End If
                        If (DrawSegment) Then
                            ''AddVertex(xyPt2)
                            ptColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                        End If
                        If (DrawInsert) Then
                            'hEnt = AddEntity("line", TransXY(xyPt1, "trs", nXInsert, nYInsert),
                            'TransXY(xyPt2, "trs", nXInsert, nYInsert)) ''// Draw insert
                            PR_MakeXY(xyStart, xyPt1.X + nXInsert, xyPt1.y + nYInsert)
                            PR_MakeXY(xyEnd, xyPt2.X + nXInsert, xyPt2.y + nYInsert)
                            ''PR_DrawLine(xyStart, xyEnd)
                            PR_DrawVestLine(xyStart, xyEnd)
                            'SetDBData(hEnt, "ID", sID)
                            'SetDBData(hEnt, "curvetype", "sleeveinsert")
                            'SetDBData(hEnt, "Data", sData)
                        End If
                        xyPt1 = xyPt2
                        aPrevAngle = aAngle
                        ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    End While
                    FileClose(hCurve)
                    PR_DrawPoly(ptColl)
                    '' // Draw insert 
                    '' // Calculate second Regular Axilla Construction point And insert points
                    ''//
                    aAngle = (System.Math.Acos(nRegAxillaFac_1 / FN_CalcLength(xyAxilla, xyRaglanAxilla)) * (180 / PI))
                    aAngle = FN_CalcAngle(xyAxilla, xyRaglanAxilla) - aAngle
                    Dim xyAxillaConstruct_2, xyInsertConstruct_1, xyInsertConstruct_2 As XY
                    ''xyAxillaConstruct_2 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1, aAngle)
                    PR_CalcPolar(xyAxilla, aAngle, nRegAxillaFac_1, xyAxillaConstruct_2)
                    ''xyInsertConstruct_1 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1 + nSeamAllowance, aAngle)
                    PR_CalcPolar(xyAxilla, aAngle, nRegAxillaFac_1 + nSeamAllowance, xyInsertConstruct_1)
                    ''xyInsertConstruct_2 = CalcXY("relpolar", xyRaglanAxilla, nSeamAllowance, aAngle)
                    PR_CalcPolar(xyRaglanAxilla, aAngle, nSeamAllowance, xyInsertConstruct_2)

                    ''// Find point on sleeve profile (note - this Is an aproximation only)
                    ''// From above nTape Is currenty set to last vertex on the profile
                    ''// For curve profiles only, Not short sleeves 3 tapes Or less
                    Dim xyAxillaLow, xyInsertConstruct_4 As XY
                    If (sEntClass.Equals("curve") Or sEntClass.Equals("polyline")) Then
                        nTape = nTape - 1
                        xyPt1 = xyAxilla
                        While (nTape > 0)
                            ''GetVertex(hSleeveProfile, nTape, xyPt2)
                            PR_MakeXY(xyPt2, ptCurveColl(nTape).X, ptCurveColl(nTape).Y)
                            If (FN_CirLinInt(xyPt1, xyPt2, xyAxilla, nRegAxillaFac_1, xyInt)) Then
                                xyAxillaLow = xyInt
                                'hEnt = AddEntity("line", xyAxillaConstruct_2, xyRaglanAxilla)
                                'SetDBData(hEnt, "ID", sID + "RegularAxillaLine")
                                'SetDBData(hEnt, "Data", sData)
                                'SetDBData(hEnt, "curvetype", "sleevecutout")
                                ''PR_DrawLine(xyAxillaConstruct_2, xyRaglanAxilla)
                                PR_DrawVestLine(xyAxillaConstruct_2, xyRaglanAxilla)
                                aPrevAngle = FN_CalcAngle(xyAxilla, xyAxillaLow)
                                aAngle = FN_CalcAngle(xyAxilla, xyAxillaConstruct_2) - aPrevAngle
                                'hEnt = AddEntity("arc", xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle)
                                'SetDBData(hEnt, "ID", sID + "RegularAxillaArc")
                                'SetDBData(hEnt, "Data", sData)
                                'SetDBData(hEnt, "curvetype", "sleevecutout")
                                PR_DrawArc(xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle)
                                Exit While
                            End If
                            nTape = nTape - 1
                            xyPt1 = xyPt2
                        End While
                    Else
                        If (FN_CirLinInt(xyProfileStart, xyAxilla, xyAxilla, nRegAxillaFac_1, xyInt)) Then
                            xyAxillaLow = xyInt
                            'hEnt = AddEntity("line", xyAxillaConstruct_2, xyRaglanAxilla)
                            'SetDBData(hEnt, "ID", sID + "RegularAxillaLine")
                            'SetDBData(hEnt, "Data", sData)
                            'SetDBData(hEnt, "curvetype", "sleevecutout")
                            ''PR_DrawLine(xyAxillaConstruct_2, xyRaglanAxilla)
                            PR_DrawVestLine(xyAxillaConstruct_2, xyRaglanAxilla)
                            aPrevAngle = FN_CalcAngle(xyAxilla, xyAxillaLow)
                            aAngle = FN_CalcAngle(xyAxilla, xyAxillaConstruct_2) - aPrevAngle
                            'hEnt = AddEntity("arc", xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle)
                            'SetDBData(hEnt, "ID", sID + "RegularAxillaArc")
                            'SetDBData(hEnt, "Data", sData)
                            'SetDBData(hEnt, "curvetype", "sleevecutout")
                            'SetDBData(hEnt, "Data", sData)
                            PR_DrawArc(xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle)
                            If (FN_CalcLength(xyAxillaLow, xyProfileStart) < 0.75) Then
                                MsgBox("Could not make an Axilla Insert" & Chr(10) & "Too close to end of Sleeve, need minimum of 3/4" & QQ, 16, "Vest Arm Dialog")
                                Exit Sub
                            End If
                        Else
                            MsgBox("Could not make an Axilla Insert" & Chr(10) & "Try again with a smaller Axilla Insert size", 16, "Vest Arm Dialog")
                            Exit Sub
                        End If
                    End If
                    '' // Draw insert
                    ''xyInsertConstruct_4 = CalcXY("relpolar", xyAxillaLow, sqrt(2 * (nSeamAllowance * nSeamAllowance)), aPrevAngle + 45)
                    PR_CalcPolar(xyAxillaLow, aPrevAngle + 45, System.Math.Sqrt(2 * (nSeamAllowance * nSeamAllowance)), xyInsertConstruct_4)
                    aPrevAngle = FN_CalcAngle(xyAxilla, xyInsertConstruct_4)
                    aAngle = FN_CalcAngle(xyAxilla, xyAxillaConstruct_2) - aPrevAngle

                    ''hEnt = AddEntity("arc", TransXY(xyAxilla, "trs", nXInsert, nYInsert), nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle)
                    PR_MakeXY(xyStart, xyAxilla.X + nXInsert, xyAxilla.y + nYInsert)
                    PR_DrawArc(xyStart, nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle)
                    'SetDBData(hEnt, "ID", sID + "RegularAxillaInsertArc")
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")
                    ''hEnt = AddEntity("line", TransXY(xyInsertConstruct_4, "trs", nXInsert, nYInsert), TransXY(xyInsertConstruct_3, "trs", nXInsert, nYInsert))
                    PR_MakeXY(xyStart, xyInsertConstruct_4.X + nXInsert, xyInsertConstruct_4.y + nYInsert)
                    PR_MakeXY(xyEnd, xyInsertConstruct_3.X + nXInsert, xyInsertConstruct_3.y + nYInsert)
                    ''PR_DrawLine(xyStart, xyEnd)
                    PR_DrawVestLine(xyStart, xyEnd)
                    'SetDBData(hEnt, "ID", sID + "RegularAxillaInsertLine")
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")
                    ''hEnt = AddEntity("line", TransXY(xyInsertConstruct_1, "trs", nXInsert, nYInsert), TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert))
                    PR_MakeXY(xyStart, xyInsertConstruct_1.X + nXInsert, xyInsertConstruct_1.y + nYInsert)
                    PR_MakeXY(xyEnd, xyInsertConstruct_2.X + nXInsert, xyInsertConstruct_2.y + nYInsert)
                    ''PR_DrawLine(xyStart, xyEnd)
                    PR_DrawVestLine(xyStart, xyEnd)
                    'SetDBData(hEnt, "ID", sID)
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")
                    ''hEnt = AddEntity("line", TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert), TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert))
                    PR_MakeXY(xyStart, xyInsertConstruct_2.X + nXInsert, xyInsertConstruct_2.y + nYInsert)
                    PR_MakeXY(xyEnd, xyRaglanAxilla.X + nXInsert, xyRaglanAxilla.y + nYInsert)
                    ''PR_DrawLine(xyStart, xyEnd)
                    PR_DrawVestLine(xyStart, xyEnd)
                    'SetDBData(hEnt, "ID", sID)
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")

                    ''// Draw seam line And stamp 
                    ''// Note recalculation because of drawing on the notes layer
                    PR_SetLayer("Notes")
                    ''hEnt = AddEntity("line", TransXY(xyAxillaConstruct_2, "trs", nXInsert, nYInsert), TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert))
                    PR_MakeXY(xyStart, xyAxillaConstruct_2.X + nXInsert, xyAxillaConstruct_2.y + nYInsert)
                    PR_MakeXY(xyEnd, xyRaglanAxilla.X + nXInsert, xyRaglanAxilla.y + nYInsert)
                    ''PR_DrawLine(xyStart, xyEnd)
                    PR_DrawVestLine(xyStart, xyEnd)
                    'SetDBData(hEnt, "ID", sID)
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")
                    aPrevAngle = FN_CalcAngle(xyAxilla, xyInsertConstruct_4)
                    aAngle = FN_CalcAngle(xyAxilla, xyAxillaConstruct_2) - aPrevAngle
                    ''hEnt = AddEntity("arc", TransXY(xyAxilla, "trs", nXInsert, nYInsert), nRegAxillaFac_1, aPrevAngle, aAngle)
                    PR_MakeXY(xyStart, xyAxilla.X + nXInsert, xyAxilla.y + nYInsert)
                    PR_DrawArc(xyStart, nRegAxillaFac_1, aPrevAngle, aAngle)
                    'SetDBData(hEnt, "ID", sID + "RegularAxillaInsertSeamArc")
                    'SetDBData(hEnt, "Data", sData)
                    'SetDBData(hEnt, "curvetype", "sleeveinsert")
                    If (sSleeve.Equals("Left")) Then
                        ''PRDataStamp(TransXY(xyAxilla.X - 1, xyAxilla.y - 1, "trs", nXInsert, nYInsert), 21)
                        PR_MakeXY(xyStart, xyAxilla.X - 1 + nXInsert, xyAxilla.y - 1 + nYInsert)
                        Dim sText As String = "Sleeve Left" + Chr(10) + g_sPatient + Chr(10) + VESTARM.g_sVestWorkOrder
                        PR_DrawVestMText(sText, xyStart, False)
                    Else
                        ''PRDataStamp(TransXY(xyAxilla.X - 1, xyAxilla.y - 1, "trs", nXInsert, nYInsert), 22)
                        PR_MakeXY(xyStart, xyAxilla.X - 1 + nXInsert, xyAxilla.y - 1 + nYInsert)
                        Dim sText As String = "Sleeve Right" + Chr(10) + g_sPatient + Chr(10) + VESTARM.g_sVestWorkOrder
                        PR_DrawVestMText(sText, xyStart, False)
                    End If
                    ''// End  Regular Axilla

            End Select

            'Close raglan drawing
            PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_CLOSE.D;")
            ''// Draw fold line at raglan curve point
            PR_SetLayer("Notes")
            Dim nShouldertoBackRagRedFac As Double = 0.1875 ''// GOP 01-02/18, 15.6
            ''hEnt = AddEntity("line", xyEndBottomCurve.X, xyEndBottomCurve.y + nShouldertoBackRagRedFac, xyRaglanNeck)
            PR_MakeXY(xyStart, xyEndBottomCurve.X, xyEndBottomCurve.y + nShouldertoBackRagRedFac)
            ''PR_DrawLine(xyStart, xyRaglanNeck)
            PR_DrawVestLine(xyStart, xyRaglanNeck)
            'SetDBData(hEnt, "ID", sID)
            'SetDBData(hEnt, "Data", sData)
            'SetDBData(hEnt, "curvetype", "sleevecrease")

            ''// Before drawing the curve we must establish the Tangent point to the bottom edge of
            ''// the template.
            ''// We can then use this to calculate a New Origin And Angle for the drawing of the curve
            ''//
            nLowerCurveRadius = FN_CalcLength(xyLowerAxilla, xyEndBottomCurve)
            Dim xyTangent, xyStartBottomCurve As XY
            xyTangent.X = 0
            xyTangent.y = 100000
            ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
            ''GetLine(hCurve, & sLine)
            ''ScanLine(sLine, "blank", & nLength, & aAngle)
            hCurve = FreeFile()
            FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
            aPrevAngle = (FN_CalcAngle(xyLowerAxilla, xyEndBottomCurve) - aLowerCurve)
            xyPt1 = xyLowerAxilla
            While Not EOF(hCurve)
                sLine = LineInput(hCurve)
                nLength = FN_GetNumber(sLine, 1)
                aAngle = FN_GetNumber(sLine, 2)
                aAngle = aAngle + aPrevAngle
                ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                If (xyPt2.y > xyTangent.y) Then
                    Exit While           ''// break When lowest point passed
                End If
                If (xyPt2.y <= xyTangent.y) Then
                    xyTangent = xyPt2  ''// Set initial tangent point
                End If
                xyPt1 = xyPt2
                aPrevAngle = aAngle
                ''ScanLine(sLine, "blank", & nLength, & aAngle)
            End While
            FileClose(hCurve)

            ''// Having got the initial tangent ( ie the lowest point on the curve) we must use this to
            ''// establish the actual tangent point on the lower edge of the template And get the
            ''// angle to rotate the startpoint of the bottonm curve to
            PR_MakeXY(xyEnd, xyEndBottomCurve.X, xyLowerAxilla.y)
            Dim nError As Short = FN_CirLinInt(xyLowerAxilla, xyEnd, xyEndBottomCurve, FN_CalcLength(xyTangent, xyEndBottomCurve), xyInt)
            aAngle = FN_CalcAngle(xyEndBottomCurve, xyTangent) - FN_CalcAngle(xyEndBottomCurve, xyInt)
            ''xyStartBottomCurve = CalcXY("relpolar",xyEndBottomCurve,Calc("length", xyEndBottomCurve, xyLowerAxilla),Calc("angle", xyEndBottomCurve, xyLowerAxilla) - aAngle) ;
            PR_CalcPolar(xyEndBottomCurve, FN_CalcAngle(xyEndBottomCurve, xyLowerAxilla) - aAngle, FN_CalcLength(xyEndBottomCurve, xyLowerAxilla), xyStartBottomCurve)
            xyTangent = xyInt

            ''//
            ''// DRAW BOTTOM CURVE
            ''//
            ''// set layer

            If (sSleeve.Equals("Right")) Then
                PR_SetLayer("TemplateRight")
            Else
                PR_SetLayer("TemplateLeft")
            End If
            ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
            ''GetLine(hCurve, & sLine)
            ''ScanLine(sLine, "blank", & nLength, & aAngle)
            hCurve = FreeFile()
            FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
            ''// Rotate curve to correct start angle
            aPrevAngle = (FN_CalcAngle(xyStartBottomCurve, xyEndBottomCurve) - aLowerCurve)
            xyPt1 = xyStartBottomCurve
            DrawSegment = False         ''// Set draw segments flag off
            Dim ptColl1 As Point3dCollection = New Point3dCollection
            While Not EOF(hCurve)
                sLine = LineInput(hCurve)
                nLength = FN_GetNumber(sLine, 1)
                aAngle = FN_GetNumber(sLine, 2)
                aAngle = aAngle + aPrevAngle
                ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                If (FN_CirLinInt(xyPt1, xyPt2, xyLowerAxilla, nLowerCurveRadius, xyInt)) Then
                    ''AddVertex(xyEndBottomCurve)
                    ptColl1.Add(New Point3d(xyEndBottomCurve.X, xyEndBottomCurve.y, 0))
                    'EndPoly()
                    'hEnt = UID("find", UID("getmax"))
                    'SetDBData(hEnt, "ID", sID)
                    'SetDBData(hEnt, "curvetype", "sleeveraglanbottom")
                    'SetDBData(hEnt, "Data", sData)
                    'break
                    Exit While
                End If
                ''// Pick up curve after it has passed tangent point
                If (xyPt2.X > xyTangent.X And DrawSegment = False) Then
                    DrawSegment = True
                    xyPt1 = xyTangent
                    ''StartPoly("polyline")
                    ''AddVertex(xyLowerAxilla)
                    ptColl1.Add(New Point3d(xyLowerAxilla.X, xyLowerAxilla.y, 0))
                    ''AddVertex(xyPt1)
                    ptColl1.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
                End If

                If (DrawSegment) Then
                    ''AddVertex(xyPt2)
                    ptColl1.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                End If
                xyPt1 = xyPt2
                aPrevAngle = aAngle
                ''ScanLine(sLine, "blank", & nLength, & aAngle)
            End While
            FileClose(hCurve)
            PR_DrawPoly(ptColl1)

            ''// Draw Closing lines
            ''hEnt = AddEntity("line", xyEndBottomCurve, xyBackNeck)
            ''PR_DrawLine(xyEndBottomCurve, xyBackNeck)
            PR_DrawVestLine(xyEndBottomCurve, xyBackNeck)
            'SetDBData(hEnt, "ID", sID)
            'SetDBData(hEnt, "curvetype", "sleeveclosing")
            'SetDBData(hEnt, "Data", sData)

            ''// Reset to layer 1
            PR_SetLayer("1")

        Else
            PR_DrawVestBody(sVestRaglan, sAge, sSleeve, Profile, strAxillaType)
            'Initialise raglan drawing
            'PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_INIT.D;")

            ''Draw for the particular axilla type
            'PrintLine(fNum, "sAxillaType = " & QQ & sAxillaType & QQ & ";")
            'Dim _frmarmdia As New armdia
            'Select Case sAxillaType
            '    Case "Open", "Lining"
            '        PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_OPEN.D;")
            '    Case "Mesh"
            '        g_bDrawBodyMesh = True
            '        PrintLine(fNum, "nAge = " & Val(sAge) & ";")
            '        If g_sSide = "Left" Then
            '            PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
            '            PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
            '        Else
            '            PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 4, ",")) & ";")
            '            PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 3, ",")) & ";")
            '        End If
            '        PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_MESH.D;")
            '    Case Else 'Regular
            '        PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_REGLR.D;")
            'End Select

            ''Close raglan drawing
            'PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_CLOSE.D;")
        End If
    End Sub
    Sub PR_DrawVestBody(ByRef sVestRaglan As Object, ByRef sAge As Object, ByRef sSleeve As String,
                        ByRef Profile As curve, ByRef strAxillaType As String)
        'Initialise raglan drawing
        PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_INIT.D;")
        ''// Initialise raglan drawing
        ''//
        ''// Establish proximal control points from Sleeve profile
        ''// 
        'GetEntityClass(hSleeveProfile, & sEntClass)
        'If (StringCompare("curve", sEntClass) Or StringCompare("polyline", sEntClass)) Then
        '    nTape = GetVertexCount(hSleeveProfile)
        '    GetVertex(hSleeveProfile, nTape, & xyAxilla)
        'Else

        '    GetGeometry(hSleeveProfile, & xyProfileStart, & xyAxilla)
        'End If
        Dim xyOrigin, xyAxilla, xyProfileStart, xyLowerAxilla As XY
        xyOrigin = xyVestInsertion
        PR_MakeXY(xyAxilla, xyOrigin.X + 2.75, xyOrigin.y + 4.63125)
        PR_MakeXY(xyProfileStart, xyOrigin.X + 0, xyOrigin.y + 4.878125)
        ''PR_MakeXY(xyProfileEnd, xyOrigin.X + 1.375, xyOrigin.y + 4.285625)
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        'ptCurveColl.Add(New Point3d(xyProfileStart.X, xyProfileStart.y, 0))
        'ptCurveColl.Add(New Point3d(xyProfileEnd.X, xyProfileEnd.y, 0))
        'ptCurveColl.Add(New Point3d(xyAxilla.X, xyAxilla.y, 0))
        Dim ii As Double
        For ii = 1 To Profile.n
            ptCurveColl.Add(New Point3d(Profile.X(ii) + xyVestInsertion.X, Profile.y(ii) + xyVestInsertion.y, 0))
        Next
        PR_MakeXY(xyAxilla, Profile.X(Profile.n) + xyVestInsertion.X, Profile.y(Profile.n) + xyVestInsertion.y)
        PR_MakeXY(xyProfileStart, Profile.X(1) + xyVestInsertion.X, Profile.y(1) + xyVestInsertion.y)
        Dim nTape As Double = ptCurveColl.Count
        Dim sEntClass As String = ""
        Dim DrawSegment As Boolean

        xyLowerAxilla.y = xyOrigin.y
        xyLowerAxilla.X = xyAxilla.X
        ''// Get 1/3 rd line down from upper proximal scale And intersect with 
        ''//
        Dim sAxillaType, sString As String
        sAxillaType = strAxillaType
        sString = sVestRaglan
        Dim nAxillaBackNeckRad As Double = FN_GetVestNumber(sString, 1)
        Dim nAxillaFrontNeckRad As Double = FN_GetVestNumber(sString, 2)
        If (sSleeve.Equals("Right")) Then
            If FN_GetVestNumber(sString, 3) > 0 Then
                nAxillaBackNeckRad = FN_GetVestNumber(sString, 3)
            End If
            If FN_GetVestNumber(sString, 4) > 0 Then
                nAxillaFrontNeckRad = FN_GetVestNumber(sString, 4)
            End If
        End If
        Dim xyAxillaBodySuit, xyBackNeck As XY
        xyAxillaBodySuit = xyAxilla
        xyAxillaBodySuit.y = xyAxilla.y - 0.125
        xyBackNeck.X = xyAxillaBodySuit.X + nAxillaBackNeckRad
        xyBackNeck.y = xyAxillaBodySuit.y

        ''// set layer
        If (sSleeve.Equals("Right")) Then
            PR_SetLayer("TemplateRight")
        Else
            PR_SetLayer("TemplateLeft")
        End If
        ''// Get vest curve angle

        ''// Calculate the required vestcurve angle.
        ''// Open the BODYCURVE data file
        ''// Check for 1. Existance And 2. Corruption  
        ''//

        ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
        Dim hCurve As Object
        hCurve = FreeFile()
        FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
        Dim sLine As String = ""
        Dim nLength, aAngle, aUpperCurve, aPrevAngle As Double
        'If (hCurve) Then
        '    '          If (GetLine(hCurve, & sLine)) Then
        '    '              ScanLine(sLine, "blank", & nLength, & aAngle)
        '    '          Else
        '    '              Exit (%abort, "Can't read VESTCURV.DAT\nFile maybe corrupted")
        '    'End If
        '    sLine = LineInput(hCurve)
        '    nLength = FN_GetNumber(sLine, 1)
        '    aAngle = FN_GetNumber(sLine, 2)
        'Else
        '    MsgBox("Can't open BODYCURV.DAT\nCheck installation", 48, "Body Arm - Dialog")
        '    Exit Sub
        'End If
        ''// Intersect the unrotated vest curve with a radius, that Is centered on the 
        ''// 0,0, to get a curve angle.
        ''// This angle will be used to calculate bodycurve rotation.
        ''// For only the upper curve, lower curve Is calculated seperatly
        ''//

        aUpperCurve = 100000    ''// Impossible value used To test For non-intersecton
        aPrevAngle = 0
        Dim xyPt1, xyPt2, xyCen, xyInt As XY
        xyPt1.X = 0
        xyPt1.y = 0
        While Not EOF(hCurve)
            sLine = LineInput(hCurve)
            nLength = FN_GetNumber(sLine, 1)
            aAngle = FN_GetNumber(sLine, 2)
            aAngle = aAngle + aPrevAngle
            ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
            PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
            If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nAxillaBackNeckRad, xyInt)) Then
                aUpperCurve = FN_CalcAngle(xyCen, xyInt)
            End If
            xyPt1 = xyPt2
            aPrevAngle = aAngle
            ''ScanLine(sLine, "blank", & nLength, & aAngle)
        End While
        If (aUpperCurve = 100000) Then
            MsgBox("Sleeve drawing error" & Chr(10) & "Can't make Upper Raglan curve with this data", 48, "Body Arm - Dialog")
        End If
        FileClose(hCurve)
        'Draw for the particular axilla type
        PrintLine(fNum, "sAxillaType = " & QQ & sAxillaType & QQ & ";")
        ''Dim _frmarmdia As New armdia
        Dim nRadius, nCount As Double
        Dim nFrontNeckRedFac As Double = 0.125            ''// 210.236, 16.K
        Dim nError As Short
        Dim xyStart, xyEnd, xyAxillaLow, xyRaglanAxilla, xyRaglanNeck, xyEndBottomCurve, xyAxillaConstruct_2 As XY
        Dim nSeamAllowance As Double = 0.125
        Select Case sAxillaType
            Case "Open", "Lining"
                PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_OPEN.D;")
                Dim nLengthToOpenAxillaPt As Double = 0.0
                ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                ''GetLine(hCurve, & sLine)
                ''ScanLine(sLine, "blank", & nLength, & aAngle)
                hCurve = FreeFile()
                FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                aPrevAngle = -aUpperCurve                    '' // Rotate curve To correct start angle 
                xyPt1 = xyAxillaBodySuit
                DrawSegment = True                  ''// Set draw segments flag on
                Dim ptOpenColl As Point3dCollection = New Point3dCollection
                ''StartPoly("polyline")
                ''AddVertex(xyPt1)
                ptOpenColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
                nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                xyCen = xyAxillaBodySuit
                nCount = 1
                While Not EOF(hCurve)
                    sLine = LineInput(hCurve)
                    nLength = FN_GetNumber(sLine, 1)
                    aAngle = FN_GetNumber(sLine, 2)
                    aAngle = aAngle + aPrevAngle
                    ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                    PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                    If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then
                        If (nCount = 1) Then
                            xyRaglanNeck = xyInt    ''// Found Raglan End
                            nRadius = nAxillaBackNeckRad
                            ''// Check if it ends on same segment
                            If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then nCount = 2
                        End If
                        If (nCount = 2) Then
                            nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + FN_CalcLength(xyPt1, xyInt)
                            ''AddVertex(xyInt)
                            ''EndPoly()
                            ptOpenColl.Add(New Point3d(xyInt.X, xyInt.y, 0))
                            ''// Revise xyEndBottomCurve 
                            aAngle = FN_CalcAngle(xyPt1, xyInt) - 90
                            ''xyEndBottomCurve = CalcXY("relpolar", xyInt, 0.75, aAngle)
                            PR_CalcPolar(xyInt, aAngle, 0.75, xyEndBottomCurve)
                            Exit While
                        End If
                        nCount = nCount + 1
                    End If
                    If (DrawSegment) Then
                        ''AddVertex(xyPt2)
                        ptOpenColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                    End If
                    nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + FN_CalcLength(xyPt1, xyPt2)
                    xyPt1 = xyPt2
                    aPrevAngle = aAngle
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)
                End While
                FileClose(hCurve)
                PR_DrawPoly(ptOpenColl)

                ''// Draw Closing line at Axilla
                ''AddEntity("line", xyAxilla, xyAxillaBodySuit)
                PR_DrawVestLine(xyAxilla, xyAxillaBodySuit)
                PR_SetLayer("Notes")
                Dim nTraversedLength As Double
                If (sAxillaType.Equals("Open")) Then
                    ''// Locate 1/3 raglan length And add an arrow at that point
                    ''// Label axilla as open
                    nLengthToOpenAxillaPt = nLengthToOpenAxillaPt / 3
                    nTraversedLength = 0.0
                    'hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                    'GetLine(hCurve, & sLine)
                    'ScanLine(sLine, "blank", & nLength, & aAngle)
                    hCurve = FreeFile()
                    FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
                    aPrevAngle = (FN_CalcAngle(xyAxillaBodySuit, xyBackNeck) - aUpperCurve)    ''// Rotate curve To correct start angle
                    xyPt1 = xyAxillaBodySuit
                    While Not EOF(hCurve)
                        sLine = LineInput(hCurve)
                        nLength = FN_GetNumber(sLine, 1)
                        aAngle = FN_GetNumber(sLine, 2)
                        aAngle = aAngle + aPrevAngle
                        ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                        PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                        nTraversedLength = nTraversedLength + nLength
                        If (nTraversedLength > nLengthToOpenAxillaPt) Then
                            nError = FN_CirLinInt(xyPt1, xyPt2, xyPt1, nTraversedLength - nLengthToOpenAxillaPt, xyInt)
                            ''AddEntity("marker", "medarrow", xyInt, 0.2, 0.1, Calc("angle", xyPt1, xyPt2) - 90)
                            ''PR_DrawVestXMarker(xyInt)
                            ''// Add text
                            ''AddEntity("text", "OPEN", xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.5)
                            PR_MakeXY(xyStart, xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.5)
                            PR_DrawVestText("OPEN", xyStart, 0.1, 0)
                            ''AddEntity("text", "AXILLA", xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.7)
                            PR_MakeXY(xyStart, xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.7)
                            PR_DrawVestText("AXILLA", xyStart, 0.1, 0)
                            Exit While
                        End If
                        xyPt1 = xyPt2
                        aPrevAngle = aAngle
                        ''ScanLine(sLine, "blank", & nLength, & aAngle)
                    End While
                    FileClose(hCurve)
                Else
                    ''AddEntity("text", "LINING", xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.5)
                    PR_MakeXY(xyStart, xyAxillaBodySuit.X, xyAxillaBodySuit.y - 0.5)
                    PR_DrawVestText("LINING", xyStart, 0.1, 0)
                End If
            Case "Mesh"
                g_bDrawBodyMesh = True
                PrintLine(fNum, "nAge = " & Val(sAge) & ";")
                'If g_sSide = "Left" Then
                '    PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
                '    PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
                'Else
                '    PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 4, ",")) & ";")
                '    PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(_frmarmdia.Controls("txtMeshData"), Object).Text, 3, ",")) & ";")
                'End If
                PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_MESH.D;")
                Dim nAge As Integer = Val(sAge)
                Dim sMeshSymbol As String
                Dim nMeshAxillaFac, nMeshSymbolLength, nMeshLength, nDistanceAlongRaglan As Double
                Dim nMeshAxillaGussetFac As Double = 0.75
                Dim nMSLengthGusset As Double = 2.1875
                Dim nMeshAxillaGussetBoysFac As Double = 0.8125
                Dim nMSLengthGussetBoys As Double = 2.375
                If (nAge <= 14) Then
                    sMeshSymbol = "mesh0to14"
                    nMeshAxillaFac = nMeshAxillaGussetFac
                    nMeshSymbolLength = nMSLengthGusset
                    nMeshLength = 2.58
                Else
                    sMeshSymbol = "mesh15andUP"
                    nMeshAxillaFac = nMeshAxillaGussetBoysFac
                    nMeshSymbolLength = nMSLengthGussetBoys
                    nMeshLength = 2.8
                End If
                ''// Find point on sleeve profile (note - this Is an aproximation only)
                ''// From above nTape Is currenty set to last vertex on the profile
                Dim xyMeshAtSeamAllowance As XY
                If (sEntClass.Equals("curve") Or sEntClass.Equals("polyline")) Then
                    nTape = nTape - 1
                    xyPt1 = xyAxilla
                    While (nTape > 0)
                        ''GetVertex(hSleeveProfile, nTape, & xyPt2)
                        PR_MakeXY(xyPt2, ptCurveColl(nTape).X, ptCurveColl(nTape).Y)
                        If (FN_CirLinInt(xyPt1, xyPt2, xyAxilla, nMeshAxillaFac - 0.125, xyInt)) Then
                            xyAxillaLow = xyInt
                        End If
                        ''// Get seam allowance point
                        If (FN_CirLinInt(xyPt1, xyPt2, xyAxilla, nMeshAxillaFac, xyInt)) Then
                            xyAxillaConstruct_2 = xyInt
                            Exit While
                        End If
                        nTape = nTape - 1
                        xyPt1 = xyPt2
                    End While
                Else
                    If (FN_CirLinInt(xyProfileStart, xyAxilla, xyAxilla, nMeshAxillaFac - 0.125, xyInt)) Then
                        xyAxillaLow = xyInt
                    End If
                    If (FN_CirLinInt(xyProfileStart, xyAxilla, xyAxilla, nMeshAxillaFac, xyInt)) Then
                        xyAxillaConstruct_2 = xyInt
                    End If
                End If
                ''// Open Curve for Reading
                ''// Assume that if the tests for Existance/Corruption are passed. Then no need to repeat them

                'hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                'GetLine(hCurve, & sLine)
                'ScanLine(sLine, "blank", & nLength, & aAngle)
                hCurve = FreeFile()
                FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)

                aPrevAngle = -aUpperCurve
                xyCen = xyAxillaLow
                nRadius = nMeshSymbolLength - nSeamAllowance
                nCount = 1
                xyPt1 = xyAxillaBodySuit
                ''//   DrawSegment = %false ;		// Set draw segments flag off
                DrawSegment = True      ''// Draw entire curve for the moment
                ''StartPoly("polyline")
                Dim ptMeshColl As Point3dCollection = New Point3dCollection
                ''AddVertex(xyAxillaBodySuit)
                ptMeshColl.Add(New Point3d(xyAxillaBodySuit.X, xyAxillaBodySuit.y, 0))
                nDistanceAlongRaglan = 0
                While Not EOF(hCurve)
                    sLine = LineInput(hCurve)
                    nLength = FN_GetNumber(sLine, 1)
                    aAngle = FN_GetNumber(sLine, 2)
                    aAngle = aAngle + aPrevAngle
                    ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle)
                    PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                    If (DrawSegment = False) Then
                        nDistanceAlongRaglan = nDistanceAlongRaglan + nLength
                    End If
                    If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then
                        If (nCount = 1) Then
                            DrawSegment = True      ''// Draw from seam allowance
                            xyRaglanAxilla = xyInt
                            nRadius = nMeshSymbolLength
                            ''// Check if seam allowance on same segment
                            If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then nCount = 2
                            xyPt1 = xyRaglanAxilla
                            nDistanceAlongRaglan = nDistanceAlongRaglan + FN_CalcLength(xyPt1, xyPt2)
                            ''//			StartPoly("polyline");
                            ''//			AddVertex(xyRaglanAxilla); 		
                        End If
                        If (nCount = 2) Then
                            xyMeshAtSeamAllowance = xyInt   ''// Mesh symbol End For angle
                            xyCen = xyAxillaBodySuit
                            nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                        End If
                        If (nCount = 3) Then
                            xyRaglanNeck = xyInt        ''// Found Raglan End
                            xyCen = xyAxillaBodySuit
                            nRadius = nAxillaBackNeckRad
                            ''// Check if it ends on same segment
                            If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then nCount = 4
                        End If
                        If (nCount = 4) Then
                            ''AddVertex(xyBackNeck)
                            ''EndPoly()
                            ptMeshColl.Add(New Point3d(xyBackNeck.X, xyBackNeck.y, 0))
                            ''// Revise xyEndBottomCurve 
                            aAngle = FN_CalcAngle(xyPt1, xyInt) - 90
                            ''xyEndBottomCurve = CalcXY("relpolar", xyInt, 0.75, aAngle)
                            PR_CalcPolar(xyInt, aAngle, 0.75, xyEndBottomCurve)
                            Exit While
                        End If
                        nCount = nCount + 1
                    End If

                    If (DrawSegment) Then
                        ''AddVertex(xyPt2)    ''// Draw raglan
                        ptMeshColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                    End If
                    xyPt1 = xyPt2
                    aPrevAngle = aAngle
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)
                End While
                FileClose(hCurve)
                PR_DrawPoly(ptMeshColl)
                ''// Insert Mesh axilla symbol at correct angle 
                ''// Insert a marker to show start point for manual editing
                ''//   AddEntity ("marker", "xmarker", xyAxillaLow,  0.0625,  0.0625, 0);		// Mesh point
                PR_SetLayer("Notes")
                ''//   AddEntity ("marker", "xmarker", xyAxillaConstruct_2, 0.0625,  0.0625, 0);	// Seam allowance

                ''//   aAngle = Calc("angle", xyAxillaLow, xyMeshAtSeamAllowance);

                ''//   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB"); 
                ''//   if ( !Symbol("find", sMeshSymbol+"seam")) Exit(%cancel, "Can't find MESH Axilla symbol to insert\nCheck your installation, that JOBST.SLB exists");
                ''//   if ( !Symbol("find", sMeshSymbol+"pro")) Exit(%cancel, "Can't find MESH Axilla symbol to insert\nCheck your installation, that JOBST.SLB exists");
                ''// Insert  Seam
                ''//   hMesh = AddEntity("symbol", sMeshSymbol + "seam", xyAxillaLow, 1, 1, aAngle);
                ''SetData("TextHorzJust", 4);
                ''SetData("TextVertJust", 32);
                If (nAge <= 14) Then
                    ''AddEntity("text","1-3/4\" + " GUSSET",xyAxillaLow.X - 0.25,xyAxillaLow.y - 0.75);
                    PR_MakeXY(xyStart, xyAxillaLow.X - 0.25, xyAxillaLow.y - 0.75)
                    PR_DrawVestText("1-3/4" + QQ + " GUSSET", xyStart, 0.1, 0)
                Else
                    ''AddEntity("text", "BOYS GUSSET", xyAxillaLow.X - 0.25, xyAxillaLow.y - 0.75)
                    PR_MakeXY(xyStart, xyAxillaLow.X - 0.25, xyAxillaLow.y - 0.75)
                    PR_DrawVestText("BOYS GUSSET", xyStart, 0.1, 0)
                End If

                ''// Insert profile
                ''//    if (StringCompare(sSleeve, "Right"))
                ''//	PRSetLayerTemplateRight() ;
                ''//    else
                ''//	PRSetLayerTemplateLeft() ;
                ''//    hMesh = AddEntity("symbol", sMeshSymbol + "pro", xyAxillaLow, 1, 1, aAngle);

                ''// Revised code to write a datafile and start the MESHDRAW
                ''// 
                ''hCurve = Open("file", "C: \\JOBST\\MESHDRAW.DAT", "write")
                hCurve = FreeFile()
                FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\MESHDRAW.DAT", VB.OpenMode.Output)
                ''SetData("UnitLinearType", 0)	''// "Inches"
                PrintLine(hCurve, Str(xyAxillaConstruct_2.X) & " " & Str(xyAxillaConstruct_2.y) & Chr(10))
                PrintLine(hCurve, Str(xyAxilla.X) & " " & Str(xyAxilla.y) & Chr(10))
                PrintLine(hCurve, Str(xyAxillaBodySuit.X) & " " & Str(xyAxillaBodySuit.y) & Chr(10))
                PrintLine(hCurve, Str(xyBackNeck.X) & " " & Str(xyBackNeck.y) & Chr(10))
                PrintLine(hCurve, nMeshLength & Chr(10))
                PrintLine(hCurve, nDistanceAlongRaglan & Chr(10))
                PrintLine(hCurve, sSleeve & Chr(10))
                ''PrintLine(hCurve, sID, "\n")
                ''SetData("UnitLinearType", 6);	// "Inches/Fraction"
                FileClose(hCurve)

                ''// End  MESH Axilla

            Case Else 'Regular
                PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_REGLR.D;")
                Dim nRegAxillaFac_1 As Double = 1.5             ''// GOP 01-02/18, 16.1
                Dim nRegAxillaFac_2 As Double = 0.75            ''// GOP 01-02/18, 16.3
                Dim nRegAxillaNormal As Double = 2.0
                Dim nRegAxillaChild As Double = 1.5
                Dim nRegAxillaLargeAdult As Double = 2.5
                If (sAxillaType.Contains("Regular 2" & QQ)) Then
                    nRegAxillaFac_1 = nRegAxillaNormal
                End If
                If (sAxillaType.Contains("Regular 1½" & QQ)) Then
                    nRegAxillaFac_1 = nRegAxillaChild
                End If
                If (sAxillaType.Contains("Regular 2½" & QQ)) Then
                    nRegAxillaFac_1 = nRegAxillaLargeAdult
                End If
                Dim nXInsert, nYInsert As Double
                nXInsert = 10
                nYInsert = -3.5
                xyCen = xyAxillaBodySuit
                ''// Open Curve for Reading
                'hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
                'GetLine(hCurve, & sLine)
                'ScanLine(sLine, "blank", & nLength, & aAngle)
                hCurve = FreeFile()
                FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)

                ''//   	aPrevAngle = (Calc("angle",xyAxilla, xyBackNeck ) - aUpperCurve);      // Rotate curve to correct start angle
                aPrevAngle = -aUpperCurve      ''// Rotate curve To correct start angle

                xyPt1 = xyAxillaBodySuit
                DrawSegment = False         ''// Set draw segments flag off
                Dim DrawInsert As Boolean = False
                nRadius = nSeamAllowance
                nCount = 1
                Dim xyInsertConstruct_3 As XY
                Dim ptRegularColl As Point3dCollection = New Point3dCollection
                While Not EOF(hCurve)
                    sLine = LineInput(hCurve)
                    nLength = FN_GetNumber(sLine, 1)
                    aAngle = FN_GetNumber(sLine, 2)
                    aAngle = aAngle + aPrevAngle
                    ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
                    PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
                    If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then
                        If (nCount = 1) Then
                            DrawInsert = True
                            xyPt1 = xyInt           ''// Found Construction point
                            xyInsertConstruct_3 = xyInt
                            nRadius = nRegAxillaFac_1
                        End If
                        If (nCount = 2) Then
                            xyCen = xyInt           ''// Found Construction point
                            nRadius = nRegAxillaFac_2
                        End If
                        If (nCount = 3) Then
                            xyCen = xyAxillaBodySuit            ''// Found raglan Start
                            nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac
                            xyRaglanAxilla = xyInt
                            ''AddEntity("line", TransXY(xyPt1, "trs", nXInsert, nYInsert), TransXY(xyInt, "trs", nXInsert, nYInsert)) ''// Draw insert
                            PR_MakeXY(xyStart, xyPt1.X + nXInsert, xyPt1.y + nYInsert)
                            PR_MakeXY(xyEnd, xyInt.X + nXInsert, xyInt.y + nYInsert)
                            PR_DrawVestLine(xyStart, xyEnd)
                            xyPt1 = xyInt
                            DrawSegment = True      ''// Start Drawing from here
                            ''StartPoly("polyline")
                            ''AddVertex(xyPt1)
                            ptRegularColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
                            DrawInsert = False
                        End If
                        If (nCount = 4) Then
                            xyRaglanNeck = xyInt        ''// Found Raglan End
                            xyCen = xyAxillaBodySuit
                            nRadius = nAxillaBackNeckRad
                            ''// Check if it ends on same segment
                            If (FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)) Then nCount = 5
                        End If
                        If (nCount = 5) Then
                            ''//				AddVertex(xyBackNeck);
                            ''AddVertex(xyInt)
                            ''EndPoly()
                            ptRegularColl.Add(New Point3d(xyInt.X, xyInt.y, 0))
                            ''// Revise xyEndBottomCurve 
                            aAngle = FN_CalcAngle(xyPt1, xyInt) - 90
                            ''xyEndBottomCurve = CalcXY("relpolar", xyInt, 0.75, aAngle)
                            PR_CalcPolar(xyInt, aAngle, 0.75, xyEndBottomCurve)
                            Exit While
                        End If
                        nCount = nCount + 1
                    End If
                    If (DrawSegment) Then
                        ''AddVertex(xyPt2)
                        ptRegularColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
                    End If
                    If (DrawInsert) Then
                        ''AddEntity("line", TransXY(xyPt1, "trs", nXInsert, nYInsert), TransXY(xyPt2, "trs", nXInsert, nYInsert)) ''// Draw insert
                        PR_MakeXY(xyStart, xyPt1.X + nXInsert, xyPt1.y + nYInsert)
                        PR_MakeXY(xyEnd, xyPt2.X + nXInsert, xyPt2.y + nYInsert)
                        PR_DrawVestLine(xyStart, xyEnd)
                    End If
                    xyPt1 = xyPt2
                    aPrevAngle = aAngle
                    ''ScanLine(sLine, "blank", & nLength, & aAngle)
                End While
                FileClose(hCurve)
                PR_DrawPoly(ptRegularColl)
                ''// Draw insert 
                ''// Calculate second Regular Axilla Construction point And insert points
                ''//
                Dim xyInsertConstruct_1, xyInsertConstruct_2 As XY
                aAngle = System.Math.Acos(nRegAxillaFac_1 / FN_CalcLength(xyAxillaBodySuit, xyRaglanAxilla)) * (180 / PI)
                aAngle = FN_CalcAngle(xyAxillaBodySuit, xyRaglanAxilla) - aAngle
                ''xyAxillaConstruct_2 = CalcXY("relpolar", xyAxillaBodySuit, nRegAxillaFac_1, aAngle)
                PR_CalcPolar(xyAxillaBodySuit, aAngle, nRegAxillaFac_1, xyAxillaConstruct_2)
                ''xyInsertConstruct_1 = CalcXY("relpolar", xyAxillaBodySuit, nRegAxillaFac_1 + nSeamAllowance, aAngle)
                PR_CalcPolar(xyAxillaBodySuit, aAngle, nRegAxillaFac_1 + nSeamAllowance, xyInsertConstruct_1)
                ''xyInsertConstruct_2 = CalcXY("relpolar", xyRaglanAxilla, nSeamAllowance, aAngle)
                PR_CalcPolar(xyRaglanAxilla, aAngle, nSeamAllowance, xyInsertConstruct_2)

                ''// Find point on sleeve profile (note - this Is an aproximation only)
                ''// From above nTape Is currenty set to last vertex on the profile
                ''// For curve profiles only, Not short sleeves 3 tapes Or less
                If (sEntClass.Equals("curve") Or sEntClass.Equals("polyline")) Then
                    nTape = nTape - 1
                    xyPt1 = xyAxilla
                    While (nTape > 0)
                        ''GetVertex(hSleeveProfile, nTape, & xyPt2)
                        PR_MakeXY(xyPt2, ptCurveColl(nTape).X, ptCurveColl(nTape).Y)
                        If (FN_CirLinInt(xyPt1, xyPt2, xyAxillaBodySuit, nRegAxillaFac_1, xyInt)) Then
                            xyAxillaLow = xyInt
                            ''hEnt = AddEntity("line", xyAxillaConstruct_2, xyRaglanAxilla)
                            ''SetDBData(hEnt, "ID", sID + "RegularAxillaLine")
                            PR_DrawVestLine(xyAxillaConstruct_2, xyRaglanAxilla)
                            aPrevAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaLow)
                            aAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle
                            ''hEnt = AddEntity("arc", xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle)
                            ''SetDBData(hEnt, "ID", sID + "RegularAxillaArc")
                            PR_DrawArc(xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle)
                            Exit While
                        End If
                        nTape = nTape - 1
                        xyPt1 = xyPt2
                    End While
                Else
                    If (FN_CirLinInt(xyProfileStart, xyAxilla, xyAxilla, nRegAxillaFac_1, xyInt)) Then
                        xyAxillaLow = xyInt
                        ''hEnt = AddEntity("line", xyAxillaConstruct_2, xyRaglanAxilla)
                        ''SetDBData(hEnt, "ID", sID + "RegularAxillaLine")
                        PR_DrawVestLine(xyAxillaConstruct_2, xyRaglanAxilla)
                        aPrevAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaLow)
                        aAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle
                        ''hEnt = AddEntity("arc", xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle)
                        ''SetDBData(hEnt, "ID", sID + "RegularAxillaArc")
                        PR_DrawArc(xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle)
                        ''// Warn if too close to end of sleev
                        If (FN_CalcLength(xyAxillaLow, xyProfileStart) < 0.5) Then
                            MsgBox("Axilla Insert is too close to end of Sleeve, need minimum of 1/2" & QQ, 16, "Body Arm - Dialog")
                            ''//			Exit (%cancel, "Could Not make an Axilla Insert\nToo close to end of Sleeve, need minimum Of 3/4\"") ;
                        End If
                    Else
                        MsgBox("Could not make an Axilla Insert" & Chr(10) & "Try again with a smaller Axilla Insert size", 16, "Body Arm - Dialog")
                        Exit Sub
                    End If
                End If
                '' // Warn if too close to elbow
                ''//
                If (FN_CalcLength(xyAxillaLow, xyElbow) < 0.75) Then
                    MsgBox("Axilla Insert is too close to the Elbow, need minimum of 3/4" & QQ, 16, "Body Arm - Dialog")
                End If
                ''// Draw insert
                ''//
                Dim xyInsertConstruct_4 As XY
                ''xyInsertConstruct_4 = CalcXY("relpolar", xyAxillaLow, sqrt(2 * (nSeamAllowance * nSeamAllowance)), aPrevAngle + 45) 
                PR_CalcPolar(xyAxillaLow, aPrevAngle + 45, System.Math.Sqrt(2 * (nSeamAllowance * nSeamAllowance)), xyInsertConstruct_4)
                aPrevAngle = FN_CalcAngle(xyAxillaBodySuit, xyInsertConstruct_4)
                aAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle

                ''hEnt = AddEntity("arc", TransXY(xyAxillaBodySuit, "trs", nXInsert, nYInsert), nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle)
                ''SetDBData(hEnt, "ID", sID + "RegularAxillaInsertArc")
                PR_MakeXY(xyStart, xyAxillaBodySuit.X + nXInsert, xyAxillaBodySuit.y + nYInsert)
                PR_DrawArc(xyStart, nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle)
                ''hEnt = AddEntity("line", TransXY(xyInsertConstruct_4, "trs", nXInsert, nYInsert), TransXY(xyInsertConstruct_3, "trs", nXInsert, nYInsert)) ;
                ''SetDBData(hEnt, "ID", sID + "RegularAxillaInsertLine") 
                PR_MakeXY(xyStart, xyInsertConstruct_4.X + nXInsert, xyInsertConstruct_4.y + nYInsert)
                PR_MakeXY(xyEnd, xyInsertConstruct_3.X + nXInsert, xyInsertConstruct_3.y + nYInsert)
                PR_DrawVestLine(xyStart, xyEnd)
                ''AddEntity("line", TransXY(xyInsertConstruct_1, "trs", nXInsert, nYInsert), TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)) ;	
                PR_MakeXY(xyStart, xyInsertConstruct_1.X + nXInsert, xyInsertConstruct_1.y + nYInsert)
                PR_MakeXY(xyEnd, xyInsertConstruct_2.X + nXInsert, xyInsertConstruct_2.y + nYInsert)
                PR_DrawVestLine(xyStart, xyEnd)
                ''AddEntity("line", TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert), TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)) ;	
                PR_MakeXY(xyStart, xyInsertConstruct_2.X + nXInsert, xyInsertConstruct_2.y + nYInsert)
                PR_MakeXY(xyEnd, xyRaglanAxilla.X + nXInsert, xyRaglanAxilla.y + nYInsert)
                PR_DrawVestLine(xyStart, xyEnd)
                ''// Draw seam line And stamp 
                ''// Note recalculation because of drawing on the notes layer
                PR_SetLayer("Notes")
                ''AddEntity ( "line", TransXY(xyAxillaConstruct_2, "trs", nXInsert, nYInsert), TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)); 
                PR_MakeXY(xyStart, xyAxillaConstruct_2.X + nXInsert, xyAxillaConstruct_2.y + nYInsert)
                PR_MakeXY(xyEnd, xyRaglanAxilla.X + nXInsert, xyRaglanAxilla.y + nYInsert)
                PR_DrawVestLine(xyStart, xyEnd)
                aPrevAngle = FN_CalcAngle(xyAxilla, xyInsertConstruct_4)
                aAngle = FN_CalcAngle(xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle
                ''hEnt = AddEntity ( "arc", TransXY( xyAxillaBodySuit, "trs", nXInsert, nYInsert), nRegAxillaFac_1, aPrevAngle, aAngle)
                ''SetDBData(hEnt, "ID", sID + "RegularAxillaInsertSeamArc") 
                PR_MakeXY(xyStart, xyAxillaBodySuit.X + nXInsert, xyAxillaBodySuit.y + nYInsert)
                PR_DrawArc(xyStart, nRegAxillaFac_1, aPrevAngle, aAngle)
                If (sSleeve.Equals("Left")) Then
                    ''PRDataStamp(TransXY(xyAxillaBodySuit.X - 1, xyAxillaBodySuit.y - 1, "trs", nXInsert, nYInsert), 21)
                    PR_MakeXY(xyStart, xyAxillaBodySuit.X - 1 + nXInsert, xyAxillaBodySuit.y - 1 + nYInsert)
                    ''AddEntity("text", "Sleeve Left" + "\n" + sName + "\n" + sWorkOrder,%1 ,%2);
                    Dim sLeftText As String = "Sleeve Left" + Chr(10) + g_sPatient + Chr(10) + g_sVestWorkOrder
                    PR_DrawVestMText(sLeftText, xyStart, False)
                Else
                    ''PRDataStamp(TransXY(xyAxillaBodySuit.X - 1, xyAxillaBodySuit.y - 1, "trs", nXInsert, nYInsert), 22)
                    PR_MakeXY(xyStart, xyAxillaBodySuit.X - 1 + nXInsert, xyAxillaBodySuit.y - 1 + nYInsert)
                    ''AddEntity("text", "Sleeve Right" + "\n" + sName + "\n" + sWorkOrder ,%1, %2);
                    Dim sRightLeft As String = "Sleeve Right" + Chr(10) + g_sPatient + Chr(10) + g_sVestWorkOrder
                    PR_DrawVestMText(sRightLeft, xyStart, False)
                End If
                ''// End  Regular Axilla
        End Select

        'Close raglan drawing
        PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_CLOSE.D;")
        Dim nShouldertoBackRagRedFac As Double = 0.1875 ''// GOP 01-02/18, 15.6
        ''// Draw fold line at raglan curve point
        PR_SetLayer("Notes")
        aAngle = FN_CalcAngle(xyEndBottomCurve, xyBackNeck)
        ''xyPt1 = CalcXY("relpolar", xyEndBottomCurve, nShouldertoBackRagRedFac, aAngle);		
        PR_CalcPolar(xyEndBottomCurve, aAngle, nShouldertoBackRagRedFac, xyPt1)
        ''AddEntity("line", xyRaglanNeck, xyPt1) ;
        PR_DrawVestLine(xyRaglanNeck, xyPt1)

        ''// Calculate for lower curve angle
        ''//     
        ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
        hCurve = FreeFile()
        FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
        Dim aLowerCurve As Double = 100000
        aPrevAngle = 0
        Dim nLowerCurveRadius As Double = FN_CalcLength(xyLowerAxilla, xyEndBottomCurve)
        xyPt1.X = 0
        xyPt1.y = 0
        ''GetLine(hCurve,& sLine)
        ''ScanLine(sLine, "blank", & nLength, & aAngle)
        Dim xyCen1 As XY
        While Not EOF(hCurve)
            sLine = LineInput(hCurve)
            nLength = FN_GetNumber(sLine, 1)
            aAngle = FN_GetNumber(sLine, 2)
            aAngle = aAngle + aPrevAngle
            ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
            PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
            If (FN_CirLinInt(xyPt1, xyPt2, xyCen1, nLowerCurveRadius, xyInt)) Then
                aLowerCurve = FN_CalcAngle(xyCen1, xyInt)
            End If
            xyPt1 = xyPt2
            aPrevAngle = aAngle
            ''ScanLine(sLine, "blank", & nLength, & aAngle)
        End While
        If (aLowerCurve = 100000) Then
            MsgBox("Sleeve drawing error" & Chr(10) & "Can't make Lower Ragalan curve with this data", 16, "Body Arm - Dialog")
        End If
        FileClose(hCurve)
        '' // Before drawing the curve we must establish the Tangent point to the bottom edge of
        ''// the template.
        ''// We can then use this to calculate a New Origin And Angle for the drawing of the curve
        ''//
        Dim xyTangent As XY
        xyTangent.X = 0
        xyTangent.y = 100000
        Dim bTangentFound As Boolean = False
        ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly")
        ''GetLine(hCurve, & sLine)
        ''ScanLine(sLine, "blank", & nLength, & aAngle)
        hCurve = FreeFile()
        FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
        aPrevAngle = (FN_CalcAngle(xyLowerAxilla, xyEndBottomCurve) - aLowerCurve)
        xyPt1 = xyLowerAxilla
        While Not EOF(hCurve)
            sLine = LineInput(hCurve)
            nLength = FN_GetNumber(sLine, 1)
            aAngle = FN_GetNumber(sLine, 2)
            aAngle = aAngle + aPrevAngle
            ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
            PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
            If (xyPt2.y > xyTangent.y) Then
                bTangentFound = True
                Exit While
            End If      ''// break when lowest point passed
            If (xyPt2.y <= xyTangent.y) Then
                xyTangent = xyPt2   ''// Set initial tangent point
            End If
            xyPt1 = xyPt2
            aPrevAngle = aAngle
            ''ScanLine(sLine, "blank", & nLength, & aAngle);
        End While
        FileClose(hCurve)
        ''// Check that we have actully found a tangent
        If (bTangentFound = False) Then
            MsgBox("Can't form bottom part of sleeve raglan", 16, "Body Arm - Dialog")
            Exit Sub
        End If

        ''// Having got the initial tangent ( ie the lowest point on the curve) we must use this to
        ''// establish the actual tangent point on the lower edge of the template And get the
        ''// angle to rotate the startpoint of the bottonm curve to
        PR_MakeXY(xyStart, xyEndBottomCurve.X, xyLowerAxilla.y)
        nError = FN_CirLinInt(xyLowerAxilla, xyStart, xyEndBottomCurve, FN_CalcLength(xyTangent, xyEndBottomCurve), xyInt)
        aAngle = FN_CalcAngle(xyEndBottomCurve, xyTangent) - FN_CalcAngle(xyEndBottomCurve, xyInt)
        Dim xyStartBottomCurve As XY
        ''xyStartBottomCurve = CalcXY("relpolar", xyEndBottomCurve, FN_CalcLength(xyEndBottomCurve, xyLowerAxilla), FN_CalcAngle(xyEndBottomCurve, xyLowerAxilla) - aAngle)
        PR_CalcPolar(xyEndBottomCurve, FN_CalcAngle(xyEndBottomCurve, xyLowerAxilla) - aAngle, FN_CalcLength(xyEndBottomCurve, xyLowerAxilla), xyStartBottomCurve)
        xyTangent = xyInt
        ''//
        ''// DRAW BOTTOM CURVE
        ''//
        ''// set layer

        If (sSleeve.Equals("Right")) Then
            PR_SetLayer("TemplateRight")
        Else
            PR_SetLayer("TemplateLeft")
        End If
        ''hCurve = Open("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
        ''GetLine(hCurve, & sLine);
        ''ScanLine(sLine, "blank", & nLength, & aAngle);
        hCurve = FreeFile()
        FileOpen(hCurve, fnGetSettingsPath("LookupTables") + "\\VESTCURV.DAT", VB.OpenMode.Input)
        ''// Rotate curve to correct start angle
        aPrevAngle = (FN_CalcAngle(xyStartBottomCurve, xyEndBottomCurve) - aLowerCurve)
        xyPt1 = xyStartBottomCurve
        DrawSegment = False         ''// Set draw segments flag off
        Dim ptColl As Point3dCollection = New Point3dCollection
        While Not EOF(hCurve)
            sLine = LineInput(hCurve)
            nLength = FN_GetNumber(sLine, 1)
            aAngle = FN_GetNumber(sLine, 2)
            aAngle = aAngle + aPrevAngle
            ''xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
            PR_CalcPolar(xyPt1, aAngle, nLength, xyPt2)
            If (FN_CirLinInt(xyPt1, xyPt2, xyLowerAxilla, nLowerCurveRadius, xyInt)) Then
                ''AddVertex(xyEndBottomCurve)
                ''EndPoly()
                ptColl.Add(New Point3d(xyEndBottomCurve.X, xyEndBottomCurve.y, 0))
                Exit While
            End If
            ''// Pick up curve after it has passed tangent point
            If (xyPt2.X > xyTangent.X And DrawSegment = False) Then
                DrawSegment = True
                xyPt1 = xyTangent
                ''StartPoly("polyline")
                ''AddVertex(xyLowerAxilla)
                ptColl.Add(New Point3d(xyLowerAxilla.X, xyLowerAxilla.y, 0))
                ''AddVertex(xyPt1)
                ptColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
            End If

            If (DrawSegment) Then
                ''AddVertex(xyPt2)
                ptColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
            End If
            xyPt1 = xyPt2
            aPrevAngle = aAngle
            ''ScanLine(sLine, "blank", & nLength, & aAngle);
        End While
        FileClose(hCurve)
        PR_DrawPoly(ptColl)
        ''// Draw Closing lines
        ''AddEntity("line", xyEndBottomCurve, xyBackNeck);
        PR_DrawVestLine(xyEndBottomCurve, xyBackNeck)
        ''// Reset to layer 1
        PR_SetLayer("1")
    End Sub

    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
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
                acText.Position = New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_GetStyle(ByRef sStyle As String)
        'This function returns a value of the following
        'in sWO
        'WorkOrder, ItemNo
        'eg
        '    12345,10
        '
        Dim nOptions As Short
        Dim sDistalStyle, sProximalStyle As String
        Dim sBestCatNo, sStyleWO As String
        Dim nResult As Short

        'Distal
        sDistalStyle = "XX"
        If g_iStyleFirstTape = g_iFirstTape Then
            'Plain
            sDistalStyle = "PL"
        Else
            'Start at tape
            If g_iStyleFirstTape < 10 Then
                sDistalStyle = "0" & LTrim(Str(g_iStyleFirstTape))
            Else
                sDistalStyle = LTrim(Str(g_iStyleFirstTape))
            End If
        End If
        If g_sVestGaunt = "True" Then
            'Gauntlet
            sDistalStyle = "GT"
        End If
        If g_sVestStump = "True" Then
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
                sProximalStyle = "0" & LTrim(Str(g_iStyleLastTape))
            Else
                sProximalStyle = LTrim(Str(g_iStyleLastTape))
            End If
        End If
        If g_sVestDetGaunt = "True" And g_iStyleLastTape <= 6 Then
            'Detachable Gauntlet only, if drawn from or less than tape +1-1/2
            sProximalStyle = "GT"
        End If
        If g_sVestFlapChk = "True" Then
            'Flap
            sProximalStyle = "FP"
        End If
        Dim _frmarmdia As New armdia
        If Left(_frmarmdia.cboFlaps.Text, 4) = "Vest" Then
            'Vest raglan
            sProximalStyle = "VR"
        End If
        If Left(_frmarmdia.cboFlaps.Text, 4) = "Body" Then
            'Bodysuit raglan
            sProximalStyle = "BR"
        End If

        'Join distal and proximal styles
        sStyle = sDistalStyle & sProximalStyle
    End Sub

    Sub PR_InsertSymbol(ByRef sSymbol As String, ByRef xyInsert As XY, ByRef nScale As Single, ByRef nRotation As Single, ByRef slayer As String)
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
        '    The DRAFIX symbol library "C:   \JOBST\JOBST.SLB" exists
        '
        'Note:-
        '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
        '
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        PR_SetVestLayer("Notes")
        Dim blkIdCollection As ObjectIdCollection = New ObjectIdCollection()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim OpenDb As Database = New Database()
        Try
            OpenDb.ReadDwgFile(g_sPathJOBST & QQ & "\\JOBST.SLB", System.IO.FileShare.Read, True, "")
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            ''MsgBox("Bad Tape File")
            Exit Sub
        End Try
        Using acTrans As Transaction = OpenDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(OpenDb.BlockTableId, OpenMode.ForRead)
            If acBlkTbl.Has(sSymbol) Then
                blkIdCollection.Add(acBlkTbl(sSymbol))
            End If
            acTrans.Commit()
        End Using
        If blkIdCollection.Count > 0 Then
            Dim iMap As IdMapping = New IdMapping()
            acCurDb.WblockCloneObjects(blkIdCollection, acCurDb.BlockTableId, iMap, DuplicateRecordCloning.Ignore, False)
        End If

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim blkRecId As ObjectId = ObjectId.Null
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If acBlkTbl.Has(sSymbol) Then
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0), acBlkTbl(sSymbol))
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0)))
                End If
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
            End If
            acTrans.Commit()
        End Using
        PrintLine(fNum, "SetSymbolLibrary(sPathJOBST +" & QQ & "\\JOBST.SLB" & QQ & ");")
        PrintLine(fNum, "Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ");")

        'PR_SetVestLayer slayer
        PrintLine(fNum, "hEnt = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC)
        PrintLine(fNum, "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC)
        PrintLine(fNum, Str(nScale) & CC & Str(nScale) & CC & Str(nRotation) & ");")
    End Sub

    Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.y = y
    End Sub

    Sub PR_NamedHandle(ByRef sHandleName As String)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to retain the entity handle of a previously
        'added entity.
        '
        'Assumes that hEnt is the entity handle to the just inserted entity.
        Return '11-June-2018

        PrintLine(fNum, "HANDLE " & sHandleName & ";")
        PrintLine(fNum, sHandleName & " = hEnt;")
    End Sub

    Sub PR_PutTapeLabel(ByRef nTape As Object, ByRef xyStart As XY, ByRef nLength As Object, ByRef nMM As Object, ByRef nGrm As Object, ByRef nRed As Object)
        Dim nDec As Object
        Dim nInt As Object
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
        Dim sTextList, sSymbol As String
        Dim xyPt As XY
        Dim nSymbolOffSet, nTextHt As Single

        'sTextList = "  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

        sSymbol = nTape & "tape"
        nSymbolOffSet = 0.6877
        nTextHt = 0.125

        PR_MakeXY(xyPt, xyStart.X, xyStart.y + nSymbolOffSet) 'Offset because symbol point is at top
        ''PR_InsertSymbol(sSymbol, xyPt, 1, 0, "Notes") '' Commented on 23-06-2018

        PR_CreateTapeLayer(g_sFileNo, g_sSide, nTape)

        PR_SetTextData(1, 32, -1, -1, 0)

        'Length text
        'N.B. format as Inches and eighths. With eighths offset up and left
        nInt = Int(CDbl(nLength)) 'Integer part of the length (before decimal point)

        'Decimal part of the length (after decimal point)
        'convert to 1/8ths and get nearest by rounding
        nDec = round((nLength - nInt) / 0.125)
        If nDec = 8 Then
            nDec = 0
            nInt = nInt + 1
        End If
        ''Added for #169, #170, & #173 in the issue list on 25th June 2019
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            ''nDec = Int((nLength - nInt) * 10)
            ''Changed on 12th July 2019
            nDec = ARMDIA1.round((nLength - nInt) * 10)
        End If

        'Draw Integer part
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 0.75)
        PR_DrawText(Trim(nInt), xyPt, nTextHt, 0)

        'Draw eights part
        PR_MakeXY(xyPt, xyStart.X + 0.0625 + (Len(Trim(nInt)) * nTextHt * 0.8), xyStart.y + 0.75 + nTextHt / 1.5)
        If nDec <> 0 Then PR_DrawText(Trim(nDec), xyPt, nTextHt / 1.5, 0)

        'MMs text
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1)
        PR_DrawText(DirectCast(nMM, TextBox).Text & "mm", xyPt, nTextHt, 0)

        'Grams text
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1.25)
        PR_DrawText(DirectCast(nGrm, Label).Text & "gm", xyPt, nTextHt, 0)

        'Reduction text and circle round the text
        PR_SetTextData(2, 16, -1, -1, -1)
        PR_MakeXY(xyPt, xyStart.X + 0.25, xyStart.y + 1.625)
        PR_DrawTextCenter(Trim(DirectCast(nRed, Label).Text), xyPt, nTextHt, 0)
        PR_DrawCircle(xyPt, 0.125)
    End Sub

    Sub PR_VestRightThumbHole(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
        ' This Calculates and Draws the curve at the
        ' right of the ThumbHole on the template

        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
        nRad = FN_CalcLength(xyCen, xyArcStart)
        nStartAng = FN_CalcAngle(xyCen, xyArcStart) * (PI / 180)
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd) * (PI / 180)
        nDeltaAng = nEndAng - nStartAng
        nDeltaAng = -nDeltaAng
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
            Using acArc As Arc = New Arc(New Point3d(xyCen.X + xyVestInsertion.X, xyCen.y + xyVestInsertion.y, 0),
                                         nRad, nStartAng, nEndAng)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    Sub PR_SetVestLayer(ByRef sNewLayer As String)
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

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                     OpenMode.ForRead)

            ' Dim sLayerName As String = "Center"

            If acLyrTbl.Has(sNewLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sNewLayer)

                '' Save the changes
                acTrans.Commit()
            End If

            '' Dispose of the transaction
        End Using
    End Sub

    Sub PR_SetTextData(ByRef nHoriz As Object, ByRef nVert As Object, ByRef nHt As Object, ByRef nAspect As Object, ByRef nFont As Object)
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
        Return '11-June-2018
        If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then

            PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & nHoriz & ");")
            g_nCurrTextHorizJust = nHoriz
        End If

        If nVert >= 0 And g_nCurrTextVertJust <> nVert Then

            PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & nVert & ");")
            g_nCurrTextVertJust = nVert
        End If

        If nHt >= 0 And g_nCurrTextHt <> nHt Then

            PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & nHt & ");")
            g_nCurrTextHt = nHt
        End If

        If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then

            PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & nAspect & ");")
            g_nCurrTextAspect = nAspect
        End If

        If nFont >= 0 And g_nCurrTextFont <> nFont Then
            PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & nFont & ");")
            g_nCurrTextFont = nFont
        End If
    End Sub

    Function round(ByVal nNumber As Single) As Short
        'Fuction to return the rounded value of a decimal number
        'E.G.
        '    round(1.35) = 1
        '    round(1.55) = 2
        '    round(2.50) = 3
        '    round(-2.50) = -3
        '

        Dim nInt, nSign As Short

        nSign = System.Math.Sign(nNumber)
        nNumber = System.Math.Abs(nNumber)
        nInt = Int(nNumber)
        If (nNumber - nInt) >= 0.5 Then
            round = (nInt + 1) * nSign
        Else
            round = nInt * nSign
        End If

    End Function
    Sub PR_DrawTextCenter(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
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
                acText.Position = New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.HorizontalMode = TextHorizontalMode.TextMid
                acText.AlignmentPoint = New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    'To Draw Spline
    Private Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection, Optional ByVal bIsSetXDATA As Boolean = False, Optional ByVal g_sbIsBody As Boolean = False)
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
                    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
                If bIsSetXDATA = True Then
                    Dim acRegAppTbl As RegAppTable
                    acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                    Dim acRegAppTblRec As RegAppTableRecord
                    Dim appName As String = "ProfileID"
                    Dim xdataStr As String = "VEST" & VESTARM.g_sID & "Profile"
                    If g_sbIsBody = True Then
                        xdataStr = "BODY" & VESTARM.g_sID & "Profile"
                    End If
                    If acRegAppTbl.Has(appName) = False Then
                        acRegAppTblRec = New RegAppTableRecord
                        acRegAppTblRec.Name = appName
                        acRegAppTbl.UpgradeOpen()
                        acRegAppTbl.Add(acRegAppTblRec)
                        acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                    End If
                    Using rb As New ResultBuffer
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                        acSpline.XData = rb
                    End Using
                End If
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As XY, ByRef bIsCenter As Boolean)
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
            mtx.Location = New Point3d(xyInsert.X + xyVestInsertion.X, xyInsert.y + xyVestInsertion.y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = 0
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            If bIsCenter = True Then
                mtx.Attachment = AttachmentPoint.TopCenter
            End If
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawVestXMarker(ByRef xyBlkInsert As XY, Optional ByVal bIsSetXDATA As Boolean = False, Optional ByVal g_sbIsBody As Boolean = False)
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
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyBlkInsert.X, xyBlkInsert.y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                If bIsSetXDATA Then
                    Dim acRegAppTbl As RegAppTable
                    acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                    Dim acRegAppTblRec As RegAppTableRecord
                    Dim appName As String = "ProfileID"
                    Dim xdataStr As String = "VEST" & VESTARM.g_sID & "Profile"
                    If g_sbIsBody = True Then
                        xdataStr = "BODY" & VESTARM.g_sID & "Profile"
                    End If
                    If acRegAppTbl.Has(appName) = False Then
                        acRegAppTblRec = New RegAppTableRecord
                        acRegAppTblRec.Name = appName
                        acRegAppTbl.UpgradeOpen()
                        acRegAppTbl.Add(acRegAppTblRec)
                        acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                    End If
                    Using rb As New ResultBuffer
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                        blkRef.XData = rb
                    End Using
                End If
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
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

        nSlope = FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        If nSlope = 0 Or nSlope = 180 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
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
        If nSlope = 90 Or nSlope = 270 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.y + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
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
        If nSlope > 0 Then
            nM = (xyEnd.y - xyStart.y) / (xyEnd.X - xyStart.X) 'Slope
            nK = xyStart.y - nM * xyStart.X 'Y-Axis intercept
            nA = (1 + nM ^ 2)
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.y * nM))
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
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
    Private Sub PR_DrawArc(ByRef xyCen As XY, ByRef nRad As Double, ByRef nStartAng As Double, ByRef nDeltaAng As Double)

        ' this procedure draws an arc between two points
        'Dim nDeltaAng As Object
        'nDeltaAng = nEndAng - nStartAng
        Dim nEndAng As Object
        nEndAng = nStartAng + nDeltaAng
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
            Using acArc As Arc = New Arc(New Point3d(xyCen.X, xyCen.y, 0),
                                         nRad, (nStartAng * (PI / 180)), (nEndAng * (PI / 180)))
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawPoly(ByRef PointCollection As Point3dCollection)
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
        Dim ii As Short
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

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
                For ii = 0 To PointCollection.Count - 1
                    acPoly.AddVertexAt(ii, New Point2d(PointCollection(ii).X, PointCollection(ii).Y), 0, 0, 0)
                Next ii

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub
    Sub PR_DrawVestLine(ByRef xyStart As XY, ByRef xyFinish As XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
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
            Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.y, 0),
                                    New Point3d(xyFinish.X, xyFinish.y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawVestMText(ByRef sText As Object, ByRef xyInsert As XY, ByRef bIsCenter As Boolean)
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
            mtx.Location = New Point3d(xyInsert.X, xyInsert.y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = 0
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            If bIsCenter = True Then
                mtx.Attachment = AttachmentPoint.TopCenter
            End If
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawVestText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
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
                acText.Position = New Point3d(xyInsert.X, xyInsert.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyVestInsertion.X, xyVestInsertion.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_CalcRaglan(xyStart As XY, xyEnd As XY, Flip As Integer, Raglan As curve, xyOpen As XY, aOpen As Double, nRadiusForRequiredPoint As Double)
        'Subroutine to draw a vest raglan curve between the two points
        'given
        'NB
        ' nRadiusForRequiredPoint is used to get the point xyOpen
        ' if nRadiusForRequiredPoint = -1 then xyOpen is at 1/3 the distance along the curve
        ' if nRadiusForRequiredPoint > 0  then xyOpen is at the radius nRadiusForRequiredPoint
        '
        Dim nSegLen(100) As Double, nSegAngle(100) As Double, iSegments As Integer
        Dim sFileName As String, fFileNum As Integer
        Dim xy0 As XY, xy1 As XY, xy2 As XY, xy4 As XY
        Dim aAngle, aPrevAngle, nTraversedLen, nLength, nOpenLength, bOpenFound As Integer
        Dim aRaglan, ii, aAxis, nRadius As Double

        'On an error the Raglan Curve is returned as a minimum of 2 points
        'these being the given start and end points.
        Raglan.n = 2
        Dim X(2), Y(2) As Double
        Raglan.X = X
        Raglan.y = Y
        Raglan.X(1) = xyStart.X
        Raglan.y(1) = xyStart.y
        Raglan.X(2) = xyEnd.X
        Raglan.y(2) = xyEnd.y

        'Check for silly data
        If xyStart.y = xyEnd.y And xyStart.X = xyEnd.X Then
            Exit Sub
        End If

        If g_sSleeveType = "Sleeveless" Then
            sFileName = g_sPathJOBST & "\TEMPLTS\BODYCURV.DAT"
        Else
            sFileName = g_sPathJOBST & "\TEMPLTS\VESTCURV.DAT"
        End If
        'Check to see if the raglan curve has been loaded from file,
        'so we don't load it twice
        If iSegments = 0 Then
            fFileNum = FreeFile()
            If FileLen(sFileName) = 0 Then
                MsgBox("Can't open the file " & sFileName & ".  Unable to draw the raglan curve.", 48, "Vest Arm - Dialog")
                Exit Sub
            End If

            ''Open sFileName For Input As fFileNum
            FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
            Dim sLine As String = ""
            Do While Not EOF(fFileNum)
                iSegments = iSegments + 1
                ''Input #fFileNum, nSegLen(iSegments), nSegAngle(iSegments)
                sLine = LineInput(fFileNum)
                nSegLen(iSegments) = FN_GetNumber(sLine, 1)
                nSegAngle(iSegments) = FN_GetNumber(sLine, 2)
            Loop
            ''Close #fFileNum
            FileClose(fFileNum)
        End If

        'Using the given points loop through the raglan curve to establish the
        'raglan angle.
        'Note:
        'This code is a little bit cumbersome but it is based on DRAFIX macro
        'code that is known to work and is production accepted.
        '
        nRadius = FN_CalcLength(xyStart, xyEnd)
        aRaglan = -1
        xy1.X = 0 : xy1.y = 0
        xy0.X = 0 : xy0.y = 0
        aPrevAngle = 0
        aAngle = nSegAngle(1)
        nLength = nSegLen(1)
        nTraversedLen = 0       'For use later

        For ii = 2 To iSegments
            aAngle = aAngle + aPrevAngle
            PR_CalcPolar(xy1, aAngle, nLength, xy2)
            If FN_CirLinInt(xy1, xy2, xy0, nRadius, xy4) Then
                aRaglan = FN_CalcAngle(xy0, xy4)
                nTraversedLen = nTraversedLen + FN_CalcLength(xy1, xy4)
                Exit For
            End If
            nTraversedLen = nTraversedLen + FN_CalcLength(xy1, xy2)
            xy1 = xy2
            aPrevAngle = aAngle
            aAngle = nSegAngle(ii)
            nLength = nSegLen(ii)
        Next ii
        'Check that the intesection has occured
        If aRaglan < 0 Then
            MsgBox("Can't draw the draw the raglan curve with this data.", 48, "Vest Arm - Dialog")
            Exit Sub
        End If

        'Create the curve given the angle above and the two point arguments
        '
        nOpenLength = nTraversedLen / 3
        nTraversedLen = 0
        bOpenFound = False

        aAngle = nSegAngle(1)
        nLength = nSegLen(1)
        aAxis = FN_CalcAngle(xyStart, xyEnd)
        aPrevAngle = aAxis - aRaglan
        xy1 = xyStart   'NB Raglan.X(1) and Raglan.Y(1), set at begining to xyStart
        Dim X1(iSegments), Y1(iSegments) As Double
        Raglan.X = X1
        Raglan.y = Y1
        Raglan.X(1) = xyStart.X
        Raglan.y(1) = xyStart.y
        For ii = 2 To iSegments
            aAngle = aAngle + aPrevAngle
            PR_CalcPolar(xy1, aAngle, nLength, xy2)
            If FN_CirLinInt(xy1, xy2, xyStart, nRadius, xy4) Then
                Raglan.X(ii) = xyEnd.X
                Raglan.y(ii) = xyEnd.y
                Exit For
            End If
            If nRadiusForRequiredPoint > 0 Then
                If FN_CirLinInt(xy1, xy2, xyStart, nRadiusForRequiredPoint, xy4) Then
                    xyOpen = xy4
                    aOpen = FN_CalcAngle(xy1, xy2)
                    If Flip Then
                        aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
                    Else
                        aOpen = aOpen - 90
                    End If
                End If
            End If
            Raglan.X(ii) = xy2.X
            Raglan.y(ii) = xy2.y
            If nTraversedLen + nLength >= nOpenLength And Not bOpenFound And nRadiusForRequiredPoint <= 0 Then
                bOpenFound = True
                aOpen = FN_CalcAngle(xy1, xy2)
                PR_CalcPolar(xy1, aOpen, (nTraversedLen + nLength) - nOpenLength, xyOpen)
                If Flip Then
                    aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
                Else
                    aOpen = aOpen - 90
                End If
            End If
            nTraversedLen = nTraversedLen + nLength
            xy1 = xy2
            aPrevAngle = aAngle
            aAngle = nSegAngle(ii)
            nLength = nSegLen(ii)
        Next ii
        Raglan.n = ii

        'If flip is true then mirror Raglan curve in the line (xyStart, xyEnd)
        'The angle of which is given by aAxis
        If Flip Then
            For ii = 1 To Raglan.n
                PR_MakeXY(xy1, Raglan.X(ii), Raglan.y(ii))
                nLength = FN_CalcLength(xyStart, xy1)
                aAngle = FN_CalcAngle(xyStart, xy1)
                aAngle = (360 - (aAngle - aAxis) + aAxis)
                PR_CalcPolar(xyStart, aAngle, nLength, xy2)
                Raglan.X(ii) = xy2.X
                Raglan.y(ii) = xy2.y
            Next ii
            nLength = FN_CalcLength(xyStart, xyOpen)
            aAngle = FN_CalcAngle(xyStart, xyOpen)
            aAngle = (360 - (aAngle - aAxis) + aAxis)
            PR_CalcPolar(xyStart, aAngle, nLength, xyOpen)
        End If
    End Sub
    Public Function FN_BiArcVestCurve(ByRef xyStart As XY, ByRef aStart As Double, ByRef xyEnd As XY, ByRef aEnd As Double, ByRef Profile As BiArcVest) As Short
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
        FN_BiArcVestCurve = False

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
        FN_BiArcVestCurve = True
        Exit Function

Error_Close:
        'Print #fError, "Error and close"
        'Close #fError

    End Function
    Function FN_FitBiArcAtGivenLength(xyStart As XY, aStart#, xyEnd As XY, aEnd#, nLength#, nLengthTol#, nRadiusTol#, Mesh As BiArcVest) As Integer

        Dim aA1#, aA2#, nArc#, aArc#

        FN_FitBiArcAtGivenLength = False
        If FN_BiArcVestCurve(xyStart, aStart, xyEnd, aEnd, Mesh) Then
            'Calculate Length
            'First arc
            aA1 = FN_CalcAngle(Mesh.xyR1, Mesh.xyStart)
            aA2 = FN_CalcAngle(Mesh.xyR1, Mesh.xyTangent)
            aArc = aA2 - aA1
            nArc = (aArc * (PI / 180)) * Mesh.nR1
            'Second arc
            aA1 = FN_CalcAngle(Mesh.xyR2, Mesh.xyTangent)
            aA2 = FN_CalcAngle(Mesh.xyR2, Mesh.xyEnd)
            aArc = aA2 - aA1
            nArc = nArc + ((aArc * (PI / 180)) * Mesh.nR2)

            'Check that the length is within the tolerence
            If nArc > (nLength - nLengthTol) And nArc < (nLength + nLengthTol) Then
                FN_FitBiArcAtGivenLength = True
                If nRadiusTol > 0 And System.Math.Abs(Mesh.nR1 - Mesh.nR2) > nRadiusTol Then FN_FitBiArcAtGivenLength = False
            End If
        End If
    End Function
    Function FN_CalcAxillaMesh(xyRaglanStart As XY, xyRaglanEnd As XY, xyMeshStart As XY, nFlip As Integer,
                               nMeshLength As Double, nAlongRaglan As Double, MeshProfile As BiArcVest, MeshSeam As BiArcVest) As Integer
        'Check for silly data
        If xyRaglanStart.X = xyRaglanEnd.X And xyRaglanStart.y = xyRaglanEnd.y Then
            Exit Function
        End If
        If xyMeshStart.X = xyRaglanEnd.X And xyMeshStart.y = xyRaglanEnd.y Then
            Exit Function
        End If
        If xyRaglanStart.X = xyMeshStart.X And xyRaglanStart.y = xyMeshStart.y Then
            Exit Function
        End If

        If FN_CalcLength(xyMeshStart, xyRaglanStart) >= nMeshLength Then
            Exit Function
        End If
        'initialize variables
        '    nTol = .03125
        Dim nTol As Double = 0.04
        Dim bMeshFound As Boolean = False
        Dim nStep As Double = -INCH1_16
        Dim nStartLength, nEndLength, aDeltaEnd, nLength As Double
        If nAlongRaglan > 0 Then
            nStartLength = nAlongRaglan
            nEndLength = nAlongRaglan                      'do at least Once
            aDeltaEnd = 160                                'Must find
        Else
            nLength = FN_CalcLength(xyMeshStart, xyRaglanStart)
            nEndLength = nLength + INCH3_16
            nStartLength = nLength * 2.5

            aDeltaEnd = 90
            nStep = INCH1_16
            If nMeshLength = 2.58 Then
                nEndLength = nLength * 2.5
                nStartLength = 1.25
            Else
                nEndLength = nLength * 2.5
                nStartLength = 1.5
            End If
        End If
        Dim aAxis, aStartDelta, aInitialStart, nRadiusTol, aStart, aInitialEnd, aAngle As Double
        aAxis = FN_CalcAngle(xyMeshStart, xyRaglanStart)
        aStartDelta = 60
        aInitialStart = 280
        nRadiusTol = 0
        FN_CalcAxillaMesh = False
        Dim cDummyCurve As curve
        For aStart = aInitialStart To (aInitialStart + aStartDelta) Step 1
            'PR_CalcPolar xyMeshStart, aStart, 1, xyTmp
            'PR_DrawMarker xyTmp
            Dim xyMeshEnd As XY
            Dim aMeshEnd, aEnd As Double
            For nLength = nStartLength To nEndLength Step nStep
                'Get point on raglan
                PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, xyMeshEnd, aInitialEnd, nLength)
                ' PR_CalcPolar xyMeshEnd, aInitialEnd, 3, xyTmp
                ' PR_DrawLine xyTmp, xyMeshEnd

                'Flip the point and angle
                If nFlip Then
                    'mirror in the line (xyMeshStart, xyRaglanStart)
                    PR_CalcPolar(xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, xyMeshEnd) + aAxis), FN_CalcLength(xyMeshStart, xyMeshEnd), xyMeshEnd)
                    aMeshEnd = ((360 - aInitialEnd) + aAxis) + 180
                Else
                    aMeshEnd = (aInitialEnd + 180)
                End If

                'End angle iteration - Get biarc curve that fits the given length to +/- nTol
                'Hold start change end
                For aEnd = (aMeshEnd - (aDeltaEnd / 2)) To (aMeshEnd + (aDeltaEnd / 2)) Step 1
                    ' PR_CalcPolar xyMeshEnd, aEnd, 1, xyTmp
                    ' PR_DrawMarker xyTmp
                    If FN_FitBiArcAtGivenLength(xyMeshStart, aStart, xyMeshEnd, aEnd, nMeshLength, nTol, nRadiusTol, MeshProfile) Then
                        bMeshFound = True
                        Exit For
                    End If
                Next aEnd
                If bMeshFound Then
                    FN_CalcAxillaMesh = True
                    nAlongRaglan = nLength
                    Exit For
                End If
            Next nLength
            If bMeshFound Then
                Exit For
            End If
        Next aStart
        'Don't calculate the seam unless we have a mesh
        If Not bMeshFound Then
            Exit Function
        End If

        'Calculate seam line
        Dim nSeam As Double = 0.125
        MeshSeam = MeshProfile
        'Get seam intersection on raglan
        nLength = nLength - nSeam
        PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, MeshSeam.xyEnd, aInitialEnd, nLength)

        'Calculate Start, End and Tangent
        aAngle = FN_CalcAngle(MeshProfile.xyR1, MeshProfile.xyStart)
        MeshSeam.nR1 = MeshProfile.nR1 - nSeam
        PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyStart)
        aAngle = FN_CalcAngle(MeshProfile.xyR1, MeshProfile.xyTangent)
        PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyTangent)
        MeshSeam.nR2 = MeshProfile.nR2 - nSeam
    End Function
End Module

Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Friend Class WEBSPACRRight_frm
    Inherits System.Windows.Forms.Form

    'Globals set by FN_Open
    Public CC As Object 'Comma
    Public QQ As Object 'Quote
    Public NL As Object 'Newline
    Public QCQ As Object 'Quote Comma Quote
    Public QC As Object 'Quote Comma
    Public CQ As Object 'Comma Quote


    'Open tip constants
    Const AGE_CUTOFF As Short = 10 'Years
    Const ADULT_STD_OPEN_TIP As Double = 0.5 'Inches
    Const CHILD_STD_OPEN_TIP As Double = 0.375 'Inches

    Public Const g_sDialogID As String = "Glove Dialogue"

    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        Me.Close()
        ManGlvMain.ManGlvMainDlg.Close()
    End Sub

    Private Function FN_CalculateWebSpacer() As Short
        'Procedure to setup Module level variables based on the data in the
        'dialogue
        'Returns true if there is enough data to work with
        'else returns false if there isn't

        Dim nWristCir, nThumbWeb, nWeb, nHighWeb, nPalmCir, nThumbCir As Double
        Dim nStrapLengthLittle, nWebToWeb, nStrapLengthMiddleAndIndex, nStrapWidth As Double
        Dim nFigure, nStrapConstruct, nGaunletCrook, nExtendThumb As Double
        Dim nSpace, nSeam, nInsert, nLength As Double
        Dim nBotRadFac, nWristNotch As Double
        Dim nThumbOffset, aThumb, aAngle, nTaperNotchRad, nWristOffset As Double
        Dim xyTmp, xyTmp1 As WEBSPACRRight_bas.xy
        Dim iAge, ii, iError As Short
        Dim sError, NL As String
        Dim nWebHt(4) As Double
        Dim nThumbHt1 As Double
        'Initialise
        NL = Chr(13) 'New Line character
        FN_CalculateWebSpacer = False 'Default to false

        'Patients age
        iAge = Val(txtAge.Text)

        'Circumferences
        nThumbCir = WEBSPACRRight_bas.fnDisplayToInches(Val(txtCir(8).Text)) 'Thumb
        nPalmCir = WEBSPACRRight_bas.fnDisplayToInches(Val(txtCir(9).Text)) 'Palm
        nWristCir = WEBSPACRRight_bas.fnDisplayToInches(Val(txtCir(10).Text)) 'Wrist

        'Webs
        'The thumb web can be lowered when the thumb
        'is being calculated in the Manual glove drawing
        'programme
        'Therefor we use the value that has been calculated and
        'stored

        nHighWeb = 0
        For ii = 1 To 4
            nWeb = WEBSPACRRight_bas.fnDisplayToInches(Val(txtLen(ii + 4).Text))
            If ii = 4 Then nWeb = nWeb - g_nRightThumbWebDrop 'Calculated thumb drop
            nWebHt(ii) = nWeb
            If nWeb > nHighWeb Then nHighWeb = nWeb
        Next ii
        'Last web is thumb web
        nThumbWeb = nWeb


        'Check on available data
        sError = ""
        If nThumbCir = 0 Then
            sError = sError & "Missing THUMB circumference!" & NL
        End If
        If nPalmCir = 0 Then
            sError = sError & "Missing PALM circumference!" & NL
        End If
        If nWristCir = 0 Then
            sError = sError & "Missing WRIST circumference!" & NL
        End If
        If nHighWeb = 0 Then
            sError = sError & "Missing WEB lengths!" & NL
        End If
        If nThumbWeb = 0 Then
            sError = sError & "Missing THUMB WEB length!" & NL
        End If

        If g_iRightWebInsertSize = 0 Then
            sError = sError & "Insert not specified!" & NL
        End If

        If g_iRightWebInsertSize < 3 Or g_iRightWebInsertSize > 8 Then
            sError = sError & "Specified insert is not found on table ""ONE INSERT WITH LARGE SEAMS"", ref GOP 01-02/28, page 4!" & NL
        End If

        If sError <> "" Then
            sError = "Unable to draw Web Spacer because of missing or incomplete data." & NL & sError
            MsgBox(sError, 16, g_sDialogID)
            FN_CalculateWebSpacer = False
            Exit Function
        End If

        'Setup figuring and construction factors
        'NB age related construction factors
        nGaunletCrook = 3 * RIGHT_WEB_SIXTEENTH
        nFigure = 0.86
        nSeam = 0.25
        nExtendThumb = 0.5
        nInsert = g_iRightWebInsertSize * RIGHT_WEB_EIGHTH

        If iAge > 10 Then
            'Adults
            nStrapLengthMiddleAndIndex = 1.25
            nStrapLengthLittle = 1
            nStrapWidth = 0.5
            nStrapConstruct = 0.25
        Else
            'Children
            nStrapLengthMiddleAndIndex = 1
            nStrapLengthLittle = 0.875
            nStrapWidth = 0.375
            nStrapConstruct = 0.1875
        End If

        'Figure dimensions
        nPalmCir = ((nPalmCir * nFigure) - nInsert) + RIGHT_WEB_QUARTER
        nWristCir = nWristCir * nFigure
        nWebToWeb = (nHighWeb - nThumbWeb) * nFigure
        nHighWeb = nHighWeb * nFigure
        nThumbCir = FN_FigureThumb(nThumbCir)
        If nThumbCir <= 0 Then
            sError = "Unable to figure the thumb with the given insert and Thumb circumference." & NL
            MsgBox(sError, 16, g_sDialogID)
            FN_CalculateWebSpacer = False
            Exit Function
        End If

        'Calculate points
        'Work w.r.t the right hand first
        'when we have caculated the thumb and we know the
        'full extents of the drawing then we can mirror and
        'translate the final image etc.

        'NOTE
        'The construction points xyRW(n) are completly
        'arbitary and given im no particular order
        'refer to drawing

        WEBSPACRRight_bas.PR_MakeXY(xyRightWebFold, 10, 0)

        WEBSPACRRight_bas.PR_MakeXY(xyRW(1), xyRightWebFold.X - (nWristCir / 2), 0)
        WEBSPACRRight_bas.PR_MakeXY(xyRW(4), xyRightWebFold.X + (nWristCir / 2), xyRW(1).Y)

        WEBSPACRRight_bas.PR_MakeXY(xyRW(2), xyRightWebFold.X - (nPalmCir / 2), nHighWeb)
        WEBSPACRRight_bas.PR_MakeXY(xyRW(3), xyRightWebFold.X + (nPalmCir / 2), xyRW(2).Y)

        WEBSPACRRight_bas.PR_MakeXY(xyRW(8), xyRW(3).X - 0.5, xyRW(3).Y)
        WEBSPACRRight_bas.PR_MakeXY(xyRW(9), xyRightWebFold.X + nSeam, xyRW(3).Y - nSeam)
        WEBSPACRRight_bas.PR_MakeXY(xyRW(5), xyRW(2).X + nSeam, xyRW(2).Y)

        WEBSPACRRight_bas.PR_MakeXY(xyRW(6), xyRW(3).X, xyRW(3).Y - nWebToWeb)

        'Finger strap points
        'Assume for the mean time no missing fingers
        nLength = WEBSPACRRight_bas.FN_CalcLength(xyRW(5), xyRW(9))
        nSpace = (nLength - (3 * nStrapWidth)) / 4

        WEBSPACRRight_bas.PR_MakeXY(xyRW(10), xyRW(5).X + nSpace, xyRW(2).Y)

        WEBSPACRRight_bas.PR_CalcPolar(xyRW(10), 90, nStrapLengthMiddleAndIndex, xyRW(20))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(20), 0, nStrapWidth, xyRW(21))

        nLength = (nStrapLengthMiddleAndIndex + nSeam) - (nSpace / 2)
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(21), 270, nLength, xyRW(11))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(11), 0, nSpace, xyRW(12))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(12), 90, nLength, xyRW(22))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(22), 0, nStrapWidth, xyRW(23))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(23), 270, nLength, xyRW(13))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(13), 0, nSpace, xyRW(14))

        nLength = (nStrapLengthLittle + nSeam) - (nSpace / 2)
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(14), 90, nLength, xyRW(24))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(24), 0, nStrapWidth, xyRW(25))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(25), 270, nLength, xyRW(15))

        WEBSPACRRight_bas.PR_MakeXY(xyRW(26), xyRW(15).X + (nSpace / 2), xyRW(9).Y)


        'Thumb points, construction circles etc.
        '    nRightNotchRad = .09375
        '    nBotRadFac = 1.3
        '    aThumb = 53
        nRightNotchRad = 0.05
        nBotRadFac = 1.4
        aThumb = 70
        nThumbOffset = 0.5
        nWristOffset = 0.5
        nWristNotch = 0.125

        nRightTopThumbRad = nThumbCir + nRightNotchRad
        nRightBotThumbRad = ((xyRW(6).Y - xyRW(4).Y) - nThumbCir)

        If ((xyRW(6).Y - xyRW(4).Y) - nThumbCir) <= 0 Then
            sError = "Query thumb circumference as it is impossible to draw the thumb with the given dimension!." & NL
            MsgBox(sError, 16, g_sDialogID)
            FN_CalculateWebSpacer = False
            Exit Function
        Else
            nRightBotThumbRad = nRightBotThumbRad * nBotRadFac
        End If

        xyRightTopThumbRad.X = xyRW(6).X + nRightNotchRad
        xyRightTopThumbRad.Y = xyRW(6).Y + nRightNotchRad

        'Calculate notch points
        'Note use of factor to create a taper of the thumb and the
        'seam into the notch
        nTaperNotchRad = nRightNotchRad * 0.95
        xyRightWebT(1).X = xyRightTopThumbRad.X - nTaperNotchRad
        xyRightWebT(1).Y = xyRightTopThumbRad.Y
        WEBSPACRRight_bas.PR_CalcPolar(xyRightTopThumbRad, 270 + aThumb, nRightTopThumbRad + 1, xyTmp)

        iError = WEBSPACRRight_bas.FN_CirLinInt(xyRightTopThumbRad, xyTmp, xyRightTopThumbRad, nTaperNotchRad, xyRightWebT(2))

        'Other thumb points at top
        WEBSPACRRight_bas.PR_CalcPolar(xyRightTopThumbRad, 270 + aThumb, nRightTopThumbRad, xyRightWebT(5))
        WEBSPACRRight_bas.PR_CalcPolar(xyRightWebT(5), aThumb, nThumbOffset, xyRightWebT(4))
        WEBSPACRRight_bas.PR_CalcPolar(xyRightWebT(4), aThumb + 90, nThumbCir, xyRightWebT(3))

        'Calculate bottom circle
        'Note use of cosine rule
        Dim nC, nA, nB, nCosA As Double
        nA = nRightTopThumbRad + nRightBotThumbRad
        nB = WEBSPACRRight_bas.FN_CalcLength(xyRightTopThumbRad, xyRW(4))
        nC = nRightBotThumbRad
        nCosA = ((nB ^ 2 + nC ^ 2) - nA ^ 2) / (2 * nB * nC)
        aAngle = WEBSPACRRight_bas.FN_CalcAngle(xyRW(4), xyRightTopThumbRad) - (Arccos(nCosA) * (180 / WEBSPACRRight_bas.PI))
        WEBSPACRRight_bas.PR_CalcPolar(xyRW(4), aAngle, nRightBotThumbRad, xyRightBotThumbRad)

        'Notch at EOS
        PR_GetDlgAboveWrist()
        If g_RightWebExtendTo = RIGHT_WEB_GLOVE_NORMAL Then
            'Calculate notch
            xyTmp.X = xyRightBotThumbRad.X
            xyTmp.Y = xyRightWebFold.Y + nWristOffset
            xyTmp1.X = xyRightWebFold.X
            xyTmp1.Y = xyTmp.Y

            If WEBSPACRRight_bas.FN_CirLinInt(xyTmp, xyTmp1, xyRightBotThumbRad, nRightBotThumbRad, xyRN(3)) Then
                'Result in xyRN(3)
            Else
                'No intesection on bottom circle
                aAngle = WEBSPACRRight_bas.FN_CalcAngle(xyRW(4), xyRightWebT(4))
                nLength = nWristOffset / System.Math.Sin(aAngle * (WEBSPACRRight_bas.PI / 180))
                WEBSPACRRight_bas.PR_CalcPolar(xyRW(4), aAngle, nLength, xyRN(3))
            End If

            xyRN(2) = xyRN(3)
            xyRN(2).X = xyRN(2).X - nWristNotch
            xyRN(1) = xyRW(4)
        Else
            'If there is data for an extension then we calculate it here
            PR_CalculateExtension(RightWebUlnarProfile, RightWebRadialProfile)
        End If

        'Exit with true if we get this far
        FN_CalculateWebSpacer = True

    End Function

    Private Function FN_FigureThumb(ByRef nThumbCir As Object) As Double
        'Figure the thumb
        'Based on the Table
        '    "ONE INSERT WITH LARGE SEAMS"
        'GOP 01-02/28, page 4
        'Date May 1991
        '
        ' Globals
        '    g_iRightWebInsertSize   -   Given insert size on EIGHTHs
        '
        ' Returns
        '    Figured dimension in inches
        '    0 if not able to calculate

        Dim iThumb As Short
        FN_FigureThumb = 0

        If g_iRightWebInsertSize < 3 Or g_iRightWebInsertSize > 8 Then Exit Function

        'Translate thumb circumference to EIGHTHs
        iThumb = ARMDIA1.round(nThumbCir / RIGHT_WEB_EIGHTH)
        'Check that thumb lies within range
        If iThumb < 8 Then Exit Function

        'Calculate Table entry for given thumb and insert
        iThumb = (3 + (7 - g_iRightWebInsertSize)) + (iThumb - 8)

        If iThumb < 0 Then Exit Function

        FN_FigureThumb = iThumb * RIGHT_WEB_SIXTEENTH
    End Function

    Private Function FN_LengthWristToEOS() As Double
        'Calculates the distance from the EOS to the
        'wrist based on values given in the dialogue
        '
        Static ii As Short
        Static nLen, nSpace As Double

        'Get the data from the dialogue
        nLen = 0

        Select Case g_RightWebExtendTo
            Case RIGHT_WEB_GLOVE_NORMAL
                nLen = 0
            Case RIGHT_WEB_GLOVE_ELBOW, RIGHT_GLOVE_PASTWRIST
                For ii = 1 To g_iRightWebNumTapesWristToEOS - 1
                    nSpace = 1.375
                    If ii = 1 And g_nRightWebPleats(1) <> 0 Then nSpace = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(1))
                    If ii = 2 And g_nRightWebPleats(2) <> 0 Then nSpace = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(2))
                    If ii = g_iRightWebNumTapesWristToEOS - 2 And ii <> 1 And g_nRightWebPleats(3) <> 0 Then nSpace = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(3))
                    If ii = g_iRightWebNumTapesWristToEOS - 1 And ii <> 2 And g_nRightWebPleats(4) <> 0 Then nSpace = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(4))
                    nLen = nLen + nSpace
                Next ii

        End Select
        FN_LengthWristToEOS = nLen
    End Function

    Private Function FN_Open(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
        'Open the DRAFIX macro file
        'Initialise Global variables

        'Open file
        WEBSPACRRight_bas.fNum = FreeFile()
        FileOpen(WEBSPACRRight_bas.fNum, sDrafixFile, VB.OpenMode.Output)
        FN_Open = WEBSPACRRight_bas.fNum

        'Initialise String globals
        CC = Chr(44) 'The comma (,)
        NL = Chr(10) 'The new line character
        QQ = Chr(34) 'Double quotes (")
        QCQ = QQ & CC & QQ
        QC = QQ & CC
        CQ = CC & QQ

        'Initialise patient globals
        g_sRightWebFileNo = sPatientFile
        g_sRightWebSide = sLeftorRight
        g_sRightWebPatient = sName

        'Globals to reduced drafix code written to file
        g_sRightWebCurrentLayer = ""
        g_nRightWebCurrTextHt = 0.125
        g_nRightWebCurrTextAspect = 0.6
        g_nRightWebCurrTextHorizJust = 1 'Left
        g_nRightWebCurrTextVertJust = 32 'Bottom
        g_nRightWebCurrTextFont = 0 'BLOCK
        g_nRightWebCurrTextAngle = 0


        'Write header information etc. to the DRAFIX macro file
        '
        PrintLine(WEBSPACRRight_bas.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(WEBSPACRRight_bas.fNum, "//Patient - " & g_sRightWebPatient & ", " & g_sRightWebFileNo & ", Hand - " & g_sRightWebSide)
        PrintLine(WEBSPACRRight_bas.fNum, "//by Visual Basic, GLOVES - Drawing")

        'Define DRAFIX variables
        PrintLine(WEBSPACRRight_bas.fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;")
        PrintLine(WEBSPACRRight_bas.fNum, "XY     xyStart, xyOrigin, xyScale, xyO;")
        PrintLine(WEBSPACRRight_bas.fNum, "STRING sFileNo, sSide, sID, sName;")
        PrintLine(WEBSPACRRight_bas.fNum, "ANGLE  aAngle;")

        'Text data
        PrintLine(WEBSPACRRight_bas.fNum, "SetData(" & QQ & "TextHorzJust" & QC & g_nRightWebCurrTextHorizJust & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "SetData(" & QQ & "TextVertJust" & QC & g_nRightWebCurrTextVertJust & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "SetData(" & QQ & "TextHeight" & QC & g_nRightWebCurrTextHt & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "SetData(" & QQ & "TextAspect" & QC & g_nRightWebCurrTextAspect & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "SetData(" & QQ & "TextFont" & QC & g_nRightWebCurrTextFont & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Data" & QCQ & "string" & QQ & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "ZipperLength" & QCQ & "length" & QQ & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Zipper" & QCQ & "string" & QQ & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "curvetype" & QCQ & "string" & QQ & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "Leg" & QCQ & "string" & QQ & ");")

        'Clear user selections etc
        PrintLine(WEBSPACRRight_bas.fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
        PrintLine(WEBSPACRRight_bas.fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")

        'Set values for use futher on by other macros
        PrintLine(WEBSPACRRight_bas.fNum, "sSide = " & QQ & g_sRightWebSide & QQ & ";")
        PrintLine(WEBSPACRRight_bas.fNum, "sFileNo = " & QQ & g_sRightWebFileNo & QQ & ";")

        'Get Start point
        PrintLine(WEBSPACRRight_bas.fNum, "GetUser (" & QQ & "xy" & QCQ & "Indicate Start Point" & QC & "&xyStart);")

        'Place a marker at the start point for later use.
        'Get a UID and create the unique 4 character start to the ID code
        'Note this is a bit dodgey if the drawing contains more than 9999 entities
        ARMDIA1.PR_SetLayer("Construct")
        PrintLine(WEBSPACRRight_bas.fNum, "hOrigin = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart" & CC & "0.125);")
        PrintLine(WEBSPACRRight_bas.fNum, "if (hOrigin) {")
        PrintLine(WEBSPACRRight_bas.fNum, "  sID=StringMiddle(MakeString(" & QQ & "long" & QQ & ",UID(" & QQ & "get" & QQ & ",hOrigin)), 1, 4) ; ")
        PrintLine(WEBSPACRRight_bas.fNum, "  while (StringLength(sID) < 4) sID = sID + " & QQ & " " & QQ & ";")
        PrintLine(WEBSPACRRight_bas.fNum, "  sID = sID + sFileNo + sSide ;")
        PrintLine(WEBSPACRRight_bas.fNum, "  SetDBData(hOrigin," & QQ & "ID" & QQ & ",sID);")
        PrintLine(WEBSPACRRight_bas.fNum, "  }")

        'Display Hour Glass Symbol
        PrintLine(WEBSPACRRight_bas.fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Drawing" & QQ & ");")

        'Set values for use futher on by other macros
        PrintLine(WEBSPACRRight_bas.fNum, "xyOrigin = xyStart" & ";")
        PrintLine(WEBSPACRRight_bas.fNum, "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");")
        PrintLine(WEBSPACRRight_bas.fNum, "if (hMPD)")
        PrintLine(WEBSPACRRight_bas.fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);")
        PrintLine(WEBSPACRRight_bas.fNum, "else")
        PrintLine(WEBSPACRRight_bas.fNum, "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");")

        'Start drawing on correct side
        ARMDIA1.PR_SetLayer("Template" & g_sRightWebSide)
    End Function

    Private Sub Form_LinkClose()

        Dim nAge, ii, nn, jj As Short
        Dim nValue As Double
        Dim sTapesBeyondWrist As String

        'Scaling factor
        If txtUnits.Text = "cm" Then
            g_nRightWebUnitsFac = 10 / 25.4
        Else
            g_nRightWebUnitsFac = 1
        End If

        If txtUidGlove.Text = "" Then
            'Disable text boxes that will not recieve data

            For ii = 20 To 23
                labCir_DoubleClick(labCir.Item(ii), New System.EventArgs())
            Next ii

            labCir_DoubleClick(labCir.Item(11), New System.EventArgs())
            txtLen(4).Text = CStr(0.5 / g_nRightWebUnitsFac)
            txtLen_Leave(txtLen.Item(4), New System.EventArgs())

            optExtendTo(1).Enabled = False
        End If

        'Hand Option buttons
        If txtSide.Text = "Left" Then
            optHand(0).Checked = True
            optHand(1).Enabled = False
            g_sRightWebSide = "Left"
        End If
        If txtSide.Text = "Right" Then
            optHand(1).Checked = True
            optHand(0).Enabled = False
            g_sRightWebSide = "Right"
        End If

        'Circumferences
        For ii = 0 To 10
            nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3))
            If nValue > 0 Then
                txtCir(ii).Text = CStr(nValue / 10)
                txtCir_Leave(txtCir.Item(ii), New System.EventArgs())
            End If
        Next ii

        'Lengths
        For ii = 0 To 8
            nn = ii + 11
            nValue = Val(Mid(txtTapeLengths.Text, (nn * 3) + 1, 3))
            If nValue > 0 Then
                txtLen(ii).Text = CStr(nValue / 10)
                txtLen_Leave(txtLen.Item(ii), New System.EventArgs())
            End If
        Next ii

        'Tapes beyond wrist
        'normal glove
        ii = 0
        sTapesBeyondWrist = Mid(txtTapeLengths2.Text, 8)
        For nn = 11 To 12
            nValue = Val(Mid(sTapesBeyondWrist, (ii * 3) + 1, 3))
            If nValue > 0 Then
                txtCir(nn).Text = CStr(nValue / 10)
                txtCir_Leave(txtCir.Item(nn), New System.EventArgs())
            End If
            ii = ii + 1
        Next nn

        cboInsert.Items.Add("Calculate")
        cboInsert.Items.Add("1/8")
        cboInsert.Items.Add("1/4")
        cboInsert.Items.Add("3/8")
        cboInsert.Items.Add("1/2")
        cboInsert.Items.Add("5/8")
        cboInsert.Items.Add("3/4")
        cboInsert.Items.Add("7/8")
        cboInsert.Items.Add("1")
        cboInsert.Items.Add("1-1/8")
        cboInsert.Items.Add("1-1/4")
        cboInsert.Items.Add("1-3/8")
        cboInsert.Items.Add("1-1/2")
        cboInsert.Items.Add("1-5/8")
        cboInsert.Items.Add("1-3/4")
        cboInsert.Items.Add("1-7/8")

        'Stored insert value (This is stored in 8ths)
        Dim sInsert As String
        Dim nInsert As Short

        sInsert = Mid(txtTapeLengths2.Text, 15, 2)
        If sInsert <> "" Then
            g_iRightWebInsertSize = Val(sInsert)
        Else
            g_iRightWebInsertSize = 4
        End If
        If g_iRightWebInsertSize <= cboInsert.Items.Count Then
            cboInsert.SelectedIndex = g_iRightWebInsertSize
        Else
            cboInsert.SelectedIndex = 4
        End If

        cboFabric.Text = txtFabric.Text

        'Default always to normal webspacer
        optExtendTo(0).Checked = True

        If Val(txtTapeLengthPt1.Text) > 0 Then
            PR_GetExtensionDDE_Data()
            If Not g_RightTapesFound Then
                'disable access to the extension
                SSTab1.TabPages(1).Enabled = False
                optExtendTo(1).Enabled = False
            End If
        Else
            'disable access to the extension
            SSTab1.TabPages(1).Enabled = False
            optExtendTo(1).Enabled = False
        End If

        'Revised Web hts
        If txtDataGlove.Text <> "" Then
            ii = 6
            nValue = Val(Mid(txtDataGlove.Text, (ii * 2) + 1, 2))
            'Amount thumbweb was dropped in 1/16ths
            g_nRightThumbWebDrop = nValue * 0.0625
        End If

        'Show after link is closed
        Show()
        'Use ok click to draw the web spacer automatically
        ' OK_Click
    End Sub

    Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
        If CmdStr = "Cancel" Then
            Cancel = 0
            Return
        End If
    End Sub

    Private Sub webspacr_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim ii As Short
        'Hide during DDE transfer
        Hide()

        WEBSPACRRight_bas.MainForm = Me

        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(WEBSPACRRight_bas.MainForm.Width)) / 2)
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(WEBSPACRRight_bas.MainForm.Height)) / 2)

        'Get the path to the JOBST system installation directory
        g_sRightWebPathJOBST = fnPathJOBST()

        'Clear DDE transfer boxes
        txtUidGlove.Text = ""
        txtTapeLengths.Text = ""
        txtTapeLengths2.Text = ""
        txtTapeLengthPt1.Text = ""
        txtWristPleat.Text = ""
        txtShoulderPleat.Text = ""
        txtSide.Text = "Right"
        txtWorkOrder.Text = ""
        txtUidGC.Text = ""
        txtFabric.Text = ""
        txtUidMPD.Text = ""
        txtDataGlove.Text = ""

        'Default units
        'g_nRightWebUnitsFac = 1             'Metric
        g_nRightWebUnitsFac = 10 / 25.4 'Inches

        'Set up fabric list
        Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
        LGLEGDIA1.PR_GetComboListFromFile(cboFabric, sSettingsPath & "\FABRIC.DAT")

        'Setup display inches grid
        'grdInches.set_ColWidth(0, 615)
        'grdInches.set_ColAlignment(0, 2)
        'For ii = 0 To 15
        '    grdInches.set_RowHeight(ii, 270)
        'Next ii

        cboDistalTape.Items.Add("1st")
        For ii = 2 To 17
            cboDistalTape.Items.Add(LTrim(Mid(g_sRightWebTapeText, (ii * 3) + 1, 3)))
        Next ii
        cboDistalTape.SelectedIndex = 0

        cboProximalTape.Items.Add("Last")
        For ii = 17 To 2 Step -1
            cboProximalTape.Items.Add(LTrim(Mid(g_sRightWebTapeText, (ii * 3) + 1, 3)))
        Next ii
        cboProximalTape.SelectedIndex = 0
        PR_GetBlkAttributeValues()
        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("No Patient Details have been found in drawing!", 16, "Web Spacers Draw")
            Me.Close()
            Exit Sub
        End If
        obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
        txtFileNo.Text = fileNo
        txtPatientName.Text = patient
        txtDiagnosis.Text = diagnosis
        txtAge.Text = age
        txtSex.Text = sex
        txtUnits.Text = units
        txtWorkOrder.Text = workOrder

        PR_UpdateAge()
        Form_LinkClose()

    End Sub

    Private Sub labCir_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labCir.DoubleClick
        Dim Index As Short = labCir.GetIndex(eventSender)
        'Procedure to disable the circumference fields
        'to enable the use to select only the relevent
        'data to be drawn
        'Saves having to delele and re-enter to get variations
        'on a glove
        Dim iDIP, iPIP As Short
        If System.Drawing.ColorTranslator.ToOle(labCir(Index).ForeColor) = &HC0C0C0 Then
            Select Case Index
                Case 0 To 7, 11 To 12
                    PR_EnableCir(Index)
                Case 8
                    PR_EnableCir(Index)
                    'Disable thumb length
                    labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
                    txtLen(4).Enabled = True
                    lblLen(4).Enabled = True
                Case 20 To 23
                    iDIP = (Index - 20) * 2
                    iPIP = iDIP + 1
                    labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
                    PR_EnableCir(iDIP)
                    PR_EnableCir(iPIP)
                    'Enable finger length
                    labLen(Index - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
                    txtLen(Index - 20).Enabled = True
                    lblLen(Index - 20).Enabled = True
            End Select
        Else
            Select Case Index
                Case 0 To 7, 12
                    PR_DisableCir(Index)
                Case 8
                    PR_DisableCir(Index)
                    'Disable thumb length
                    labLen(4).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    txtLen(4).Enabled = False
                    lblLen(4).Enabled = False
                Case 11
                    PR_DisableCir(Index)
                    PR_DisableCir(Index + 1)
                Case 20 To 23
                    iDIP = (Index - 20) * 2
                    iPIP = iDIP + 1
                    labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    PR_DisableCir(iDIP)
                    PR_DisableCir(iPIP)
                    'Disable finger length
                    labLen(Index - 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    txtLen(Index - 20).Enabled = False
                    lblLen(Index - 20).Enabled = False
            End Select
        End If
    End Sub

    Private Sub labLen_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labLen.DoubleClick
        Dim Index As Short = labLen.GetIndex(eventSender)
        If Index = 5 Then Exit Sub
        If System.Drawing.ColorTranslator.ToOle(labLen(Index).ForeColor) = &HC0C0C0 Then
            labLen(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
            txtLen(Index).Enabled = True
            lblLen(Index).Enabled = True
        Else
            labLen(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            txtLen(Index).Enabled = False
            lblLen(Index).Enabled = False
        End If
    End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click

        'Don't allow multiple clicking
        '
        OK.Enabled = False

        If FN_CalculateWebSpacer() Then
            'Display and hourglass
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Me.Hide()
            ManGlvMain.ManGlvMainDlg.Hide()
            WEBSPACRRight_bas.PR_GetInsertionPoint()
            Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
            PR_CreateDrawMacro(sDrawFile)
        End If

        OK.Enabled = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Me.Close()
        ManGlvMain.ManGlvMainDlg.Close()
        Exit Sub
    End Sub

    Private Sub optHand_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHand.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optHand.GetIndex(eventSender)
            If Index = 0 Then
                g_sRightWebSide = "Left"
            Else
                g_sRightWebSide = "Right"
            End If
        End If
    End Sub

    Private Sub PR_CalculateExtension(ByRef RightWebUlnarProfile As WEBSPACRRight_bas.curve, ByRef RightWebRadialProfile As WEBSPACRRight_bas.curve)
        'Procedure to calculate the POINTS
        'used to draw the extension of the glove above the
        'wrist
        'The Procedure also supplies the data that is to be
        'printed at each tape
        'N.B.

        Dim nValue, nMidWristX, nSpacing As Double
        Dim iLastTape, ii, iStartTape, iError As Short
        Dim nFiguredValue As Double
        Dim iVertex As Short
        Dim nWristOffset, nHt, nWristNotch As Double
        Dim xyTmp As WEBSPACRRight_bas.xy
        RightWebUlnarProfile.Initialize()
        RightWebRadialProfile.Initialize()

        If g_RightWebExtendTo = RIGHT_WEB_GLOVE_ELBOW Or g_RightWebExtendTo = RIGHT_GLOVE_PASTWRIST Then
            'As we have recalculated the wrist we need to start at the
            'wrist
            nWristOffset = 0.5
            nWristNotch = 0.125

            nHt = FN_LengthWristToEOS()
            Dim X(g_iRightWebNumTapesWristToEOS), Y(g_iRightWebNumTapesWristToEOS) As Double
            RightWebUlnarProfile.X = X
            RightWebUlnarProfile.y = Y
            RightWebUlnarProfile.n = g_iRightWebNumTapesWristToEOS
            RightWebUlnarProfile.y(1) = nHt

            Dim RX(g_iRightWebNumTapesWristToEOS), RY(g_iRightWebNumTapesWristToEOS) As Double
            RightWebRadialProfile.X = RX
            RightWebRadialProfile.y = RY
            RightWebRadialProfile.n = g_iRightWebNumTapesWristToEOS
            RightWebRadialProfile.y(1) = nHt

            iStartTape = g_iRightWebWristPointer
            iLastTape = g_iRightWebWristPointer + (g_iRightWebNumTapesWristToEOS - 1)

            'Note: As we can allow the wrist to be any tape we need a
            'vertex counter
            iVertex = 1
            For ii = iStartTape To iLastTape
                'NB use of 1/2 scale
                nValue = g_nRightWebCir(ii) * 0.86
                RightWebUlnarProfile.X(iVertex) = xyRightWebFold.X - (nValue / 2)
                RightWebRadialProfile.X(iVertex) = xyRightWebFold.X + (nValue / 2)

                'Setup the notes for the vertex
                'We do this here as we have all the data and it simplifies
                'the drawing side
                RightWebTapeNote(iVertex).sTapeText = LTrim(Mid(g_sRightWebTapeText, ((ii + 1) * 3) + 1, 3))
                RightWebTapeNote(iVertex).iTapePos = ii
                RightWebTapeNote(iVertex).nCir = g_nRightWebCir(ii)

                'Standard spacing
                nSpacing = 1.375

                'Account all for pleats
                'Wrist (as we have started at the wrist we use, iStartTape + 1
                'and iStartTape + 2 in this case)
                If ii = iStartTape + 1 And g_nRightWebPleats(1) > 0 Then nSpacing = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(1))
                If ii = iStartTape + 2 And g_nRightWebPleats(2) > 0 Then nSpacing = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(2))

                'Axilla
                If ii = iLastTape - 1 And g_nRightWebPleats(3) <> 0 Then nSpacing = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(3))
                If ii = iLastTape And g_nRightWebPleats(4) <> 0 Then nSpacing = WEBSPACRRight_bas.fnDisplayToInches(g_nRightWebPleats(4))

                'We are working from the top (wrist) down
                'iVertex = 1 is set before the For Loop (Values from wrist)
                If iVertex <> 1 Then
                    RightWebUlnarProfile.y(iVertex) = RightWebUlnarProfile.y(iVertex - 1) - nSpacing
                    RightWebRadialProfile.y(iVertex) = RightWebRadialProfile.y(iVertex - 1) - nSpacing
                    If nSpacing <> 1.375 Then RightWebTapeNote(iVertex).sNote = "PLEAT"
                End If

                'Increment vertex count
                iVertex = iVertex + 1
            Next ii

            '
            'Calculate Notch
            WEBSPACRRight_bas.PR_MakeXY(xyRN(1), RightWebRadialProfile.X(RightWebRadialProfile.n), RightWebRadialProfile.y(RightWebRadialProfile.n))
            WEBSPACRRight_bas.PR_MakeXY(xyTmp, RightWebRadialProfile.X(RightWebRadialProfile.n - 1), RightWebRadialProfile.y(RightWebRadialProfile.n - 1))
            If RightWebTapeNote(RightWebRadialProfile.n).iTapePos = RIGHT_WEB_ELBOW_TAPE Then
                'Cutback by 3/4 of an inch
                iError = WEBSPACRRight_bas.FN_CirLinInt(xyRN(1), xyTmp, xyRN(1), 0.75, xyRN(1))
            End If

            iError = WEBSPACRRight_bas.FN_CirLinInt(xyRN(1), xyTmp, xyRN(1), nWristOffset, xyRN(3))
            xyRN(2) = xyRN(3)
            xyRN(2).X = xyRN(2).X - nWristNotch

            'Reset end of profile
            RightWebRadialProfile.X(RightWebRadialProfile.n) = xyRN(3).X
            RightWebRadialProfile.y(RightWebRadialProfile.n) = xyRN(3).Y
            RightWebUlnarProfile.X(RightWebUlnarProfile.n) = xyRightWebFold.X - (xyRN(3).X - xyRightWebFold.X)
            RightWebUlnarProfile.y(RightWebUlnarProfile.n) = xyRN(3).Y
        End If
    End Sub

    Private Sub PR_CreateDrawMacro(ByRef sFileName As String)
        '
        ' GLOBALS
        '            g_sRightWebSide
        '

        Dim nInsert, nFingerSpace As Double
        Dim ii As Short
        Dim xyPt3, xyPt1, xyPt2, xyTextAlign As WEBSPACRRight_bas.xy
        Dim xyWristPt1, xyWristPt2 As WEBSPACRRight_bas.xy
        Dim xyTmp As WEBSPACRRight_bas.xy
        Dim aInc, aAngle, nTranslate As Double
        Dim Thumb As WEBSPACRRight_bas.curve
        Dim sText, sWorkOrder As String

        'Initialise
        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        WEBSPACRRight_bas.fNum = FN_Open(sDrawFile, txtPatientName.Text, txtFileNo.Text, g_sRightWebSide)

        'Draw on correct layer
        ARMDIA1.PR_SetLayer("Template" & g_sRightWebSide)

        '    PR_DebugPoint "xyRightWebFold", xyRightWebFold
        '    For ii = 1 To 20
        '        PR_DebugPoint "W" & Trim$(Str$(ii)), xyRW(ii)
        '    Next ii
        '    For ii = 1 To 8
        '        PR_DebugPoint "T" & Trim$(Str$(ii)), xyRightWebT(ii)
        '    Next ii

        'PR_DebugPoint "xyRightTopThumbRad", xyRightTopThumbRad
        'PR_DebugPoint "xyRightBotThumbRad", xyRightBotThumbRad

        'Having established the points we now draw them
        'We must mirror for the right hand side and
        'translate all points relative to our temporary datum xyRightWebFold

        'Translate points relative to xyRightWebFold
        Dim nYTranslate As Double
        nYTranslate = FN_LengthWristToEOS()
        nTranslate = -(xyRightWebFold.X + (xyRightWebFold.X - xyRightWebT(4).X))
        xyRightWebFold.X = xyRightWebFold.X + nTranslate
        xyRightWebFold.Y = xyRightWebFold.Y + nYTranslate

        For ii = 1 To 5
            xyRightWebT(ii).X = xyRightWebT(ii).X + nTranslate
            xyRightWebT(ii).Y = xyRightWebT(ii).Y + nYTranslate
        Next ii

        For ii = 1 To 3
            xyRN(ii).X = xyRN(ii).X + nTranslate
        Next ii

        For ii = 1 To 26
            xyRW(ii).X = xyRW(ii).X + nTranslate
            xyRW(ii).Y = xyRW(ii).Y + nYTranslate
            If g_sRightWebSide = "Left" Then WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyRW(ii))
        Next ii
        xyRightTopThumbRad.X = xyRightTopThumbRad.X + nTranslate
        xyRightTopThumbRad.Y = xyRightTopThumbRad.Y + nYTranslate
        xyRightBotThumbRad.X = xyRightBotThumbRad.X + nTranslate
        xyRightBotThumbRad.Y = xyRightBotThumbRad.Y + nYTranslate


        'Drawfingerstraps
        Dim Polycurve As WEBSPACRRight_bas.curve
        Dim X(17), Y(17), Bulge(17) As Double
        X(1) = xyRW(2).X
        Y(1) = xyRW(2).Y
        Bulge(1) = 0
        X(2) = xyRW(10).X
        Y(2) = xyRW(10).Y
        Bulge(2) = 0
        X(3) = xyRW(20).X
        Y(3) = xyRW(20).Y
        Bulge(3) = 0
        X(4) = xyRW(21).X
        Y(4) = xyRW(21).Y
        Bulge(4) = 0
        If g_sRightWebSide = "Left" Then
            ''-------BDYUTILS.PR_AddVertex(xyRW(11), -1)
            X(5) = xyRW(11).X
            Y(5) = xyRW(11).Y
            Bulge(5) = -1
        Else
            ''----------BDYUTILS.PR_AddVertex(xyRW(11), 1)
            X(5) = xyRW(11).X
            Y(5) = xyRW(11).Y
            Bulge(5) = 1
        End If
        X(6) = xyRW(12).X
        Y(6) = xyRW(12).Y
        Bulge(6) = 0
        X(7) = xyRW(22).X
        Y(7) = xyRW(22).Y
        Bulge(7) = 0
        X(8) = xyRW(23).X
        Y(8) = xyRW(23).Y
        Bulge(8) = 0
        If g_sRightWebSide = "Left" Then
            X(9) = xyRW(13).X
            Y(9) = xyRW(13).Y
            Bulge(9) = -1
        Else
            X(9) = xyRW(13).X
            Y(9) = xyRW(13).Y
            Bulge(9) = 1
        End If
        X(10) = xyRW(14).X
        Y(10) = xyRW(14).Y
        Bulge(10) = 0
        X(11) = xyRW(24).X
        Y(11) = xyRW(24).Y
        Bulge(11) = 0
        X(12) = xyRW(25).X
        Y(12) = xyRW(25).Y
        Bulge(12) = 0
        If g_sRightWebSide = "Left" Then
            X(13) = xyRW(15).X
            Y(13) = xyRW(15).Y
            Bulge(13) = -0.414
        Else
            X(13) = xyRW(15).X
            Y(13) = xyRW(15).Y
            Bulge(13) = 0.414
        End If
        X(14) = xyRW(26).X
        Y(14) = xyRW(26).Y
        Bulge(14) = 0
        X(15) = xyRW(9).X
        Y(15) = xyRW(9).Y
        Bulge(15) = 0
        X(16) = xyRW(8).X
        Y(16) = xyRW(8).Y
        Bulge(16) = 0
        X(17) = xyRW(3).X
        Y(17) = xyRW(3).Y
        Bulge(17) = 0
        Polycurve.n = 17
        Polycurve.X = X
        Polycurve.y = Y
        WEBSPACRRight_bas.PR_DrawPoly(Polycurve, Bulge)

        Dim nLength As Double
        If g_RightWebExtendTo = RIGHT_WEB_GLOVE_ELBOW Or g_RightWebExtendTo = RIGHT_GLOVE_PASTWRIST Then
            Dim BulgeWeb(RightWebRadialProfile.n) As Double
            For ii = 1 To RightWebRadialProfile.n
                RightWebUlnarProfile.X(ii) = RightWebUlnarProfile.X(ii) + nTranslate
                RightWebRadialProfile.X(ii) = RightWebRadialProfile.X(ii) + nTranslate
                BulgeWeb(ii) = 0
            Next ii

            WEBSPACRRight_bas.PR_DrawPoly(RightWebUlnarProfile, BulgeWeb)
            WEBSPACRRight_bas.PR_DrawPoly(RightWebRadialProfile, BulgeWeb)

            'Draw notes for each tape
            If g_RightWebExtendTo = RIGHT_WEB_GLOVE_ELBOW Then
                nLength = -1
                For ii = 1 To RightWebRadialProfile.n
                    WEBSPACRRight_bas.PR_MakeXY(xyPt1, RightWebUlnarProfile.X(ii), RightWebUlnarProfile.y(ii))
                    WEBSPACRRight_bas.PR_MakeXY(xyPt2, RightWebRadialProfile.X(ii), RightWebRadialProfile.y(ii))
                    If ii = 1 Then
                        xyWristPt1 = xyPt1
                        xyWristPt2 = xyPt2
                    End If
                    If ii <> RightWebRadialProfile.n Then
                        PR_DrawTapeNotes(ii, xyPt1, xyPt2, nLength, "")
                    Else
                        If RightWebTapeNote(RightWebRadialProfile.n).iTapePos <> RIGHT_WEB_ELBOW_TAPE Then
                            xyTmp = xyRN(1)
                            xyPt3 = xyRN(1)
                        Else
                            xyTmp = xyRN(1)
                            xyTmp.Y = 0
                            xyPt3 = xyTmp
                        End If
                        WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyPt3)
                        PR_DrawTapeNotes(ii, xyTmp, xyPt3, nLength, "")
                    End If

                    If RightWebTapeNote(ii).iTapePos = RIGHT_WEB_ELBOW_TAPE And RightWebTapeNote(RightWebRadialProfile.n).iTapePos <> RIGHT_WEB_ELBOW_TAPE Then
                        'Draw elbow line
                        ARMDIA1.PR_SetTextData(HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT)
                        ARMDIA1.PR_SetLayer("Notes")
                        WEBSPACRRight_bas.PR_DrawLine(xyPt1, xyPt2)
                        WEBSPACRRight_bas.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
                        xyPt3.Y = xyPt3.Y + 0.0625
                        WEBSPACRRight_bas.PR_DrawText("E", xyPt3, 0.1, 0)
                    End If
                Next ii

                'Center line
                ARMDIA1.PR_SetLayer("Construct")
                WEBSPACRRight_bas.PR_CalcMidPoint(xyWristPt1, xyWristPt2, xyTmp)
                WEBSPACRRight_bas.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
                WEBSPACRRight_bas.PR_DrawLine(xyPt3, xyTmp)
            End If

            'Wrist line
            ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)

            ARMDIA1.PR_SetLayer("Notes")
            WEBSPACRRight_bas.PR_DrawLine(xyRW(1), xyRW(4))
            WEBSPACRRight_bas.PR_CalcMidPoint(xyRW(1), xyRW(4), xyPt3)
            xyPt3.Y = xyPt3.Y + 0.0625
            WEBSPACRRight_bas.PR_DrawText("W", xyPt3, 0.1, 0)

            'Elastic Text
            xyPt1 = xyRN(1)
            xyPt2 = xyRN(1)

            WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyPt2)

            sText = Trim(ARMEDDIA1.fnInchesToText(WEBSPACRRight_bas.FN_CalcLength(xyPt1, xyPt2) + 1.5))
            WEBSPACRRight_bas.PR_CalcMidPoint(xyPt1, xyPt2, xyPt3)
            xyPt3.Y = xyPt3.Y + 0.375
            sText = sText & Chr(34) & " ELASTIC"
            WEBSPACRRight_bas.PR_DrawText(sText, xyPt3, 0.1, 0)

            'Elastic for children
            If Val(txtAge.Text) <= 10 Then
                '           PR_Setlayer "Notes"
                '           PR_CalcMidPoint xyPt1, xyPt2, xyPt3
                xyPt3.Y = xyPt3.Y - 0.25
                ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
                WEBSPACRRight_bas.PR_DrawText("1/2" & Chr(34) & " ELASTIC", xyPt3, 0.1, 0)
            End If

            'EOS
            ARMDIA1.PR_SetLayer("Template" & g_sRightWebSide)
            WEBSPACRRight_bas.PR_DrawLine(xyPt1, xyPt2)
        Else
            'Elastic text at EOS
            ARMDIA1.PR_SetLayer("Notes")
            sText = Trim(ARMEDDIA1.fnInchesToText(WEBSPACRRight_bas.FN_CalcLength(xyRW(1), xyRW(4)) + 1.5))
            ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT)
            WEBSPACRRight_bas.PR_CalcMidPoint(xyRW(1), xyRW(4), xyPt3)
            xyPt3.Y = xyPt3.Y + 0.25
            sText = sText & Chr(34) & " ELASTIC"
            WEBSPACRRight_bas.PR_DrawText(sText, xyPt3, 0.1, 0)

            'EOS at wrist
            ARMDIA1.PR_SetLayer("Template" & g_sRightWebSide)
            WEBSPACRRight_bas.PR_DrawLine(xyRW(1), xyRW(4))
        End If

        'Drawthumb (LEFT)
        If g_sRightWebSide = "Left" Then WEBSPACRRight_bas.PR_DrawLine(xyRightWebT(1), xyRW(2)) Else WEBSPACRRight_bas.PR_DrawLine(xyRightWebT(1), xyRW(3))

        WEBSPACRRight_bas.PR_DrawArc(xyRightTopThumbRad, xyRightWebT(1), xyRightWebT(2))

        Dim X1(4), Y1(4), Bulge1(4) As Double
        X1(1) = xyRightWebT(2).X
        Y1(1) = xyRightWebT(2).Y
        Bulge1(1) = 0
        X1(2) = xyRightWebT(3).X
        Y1(2) = xyRightWebT(3).Y
        Bulge1(2) = 0
        X1(3) = xyRightWebT(4).X
        Y1(3) = xyRightWebT(4).Y
        Bulge1(3) = 0
        X1(4) = xyRightWebT(5).X
        Y1(4) = xyRightWebT(5).Y
        Bulge1(4) = 0
        Polycurve.n = 4
        Polycurve.X = X1
        Polycurve.y = Y1
        WEBSPACRRight_bas.PR_DrawPoly(Polycurve, Bulge1)

        'Notch
        Dim X2(3), Y2(3), Bulge2(3) As Double
        X2(1) = xyRN(1).X
        Y2(1) = xyRN(1).Y
        Bulge2(1) = 0
        X2(2) = xyRN(2).X
        Y2(2) = xyRN(2).Y
        Bulge2(2) = 0
        X2(3) = xyRN(3).X
        Y2(3) = xyRN(3).Y
        Bulge2(3) = 0
        Polycurve.n = 3
        Polycurve.X = X2
        Polycurve.y = Y2
        WEBSPACRRight_bas.PR_DrawPoly(Polycurve, Bulge2)

        'Calculate thumb curve
        'For the left then mirror for right
        Thumb.n = 8
        Dim TX(8), TY(8) As Double
        Thumb.X = TX
        Thumb.y = TY

        If g_RightWebExtendTo = RIGHT_WEB_GLOVE_NORMAL Then
            xyTmp = xyRN(3)
        Else
            If g_sRightWebSide = "Left" Then
                xyTmp = xyRW(1)
            Else
                xyTmp = xyRW(4)
            End If
        End If
        Thumb.n = 1
        Thumb.X(Thumb.n) = xyTmp.X
        Thumb.y(Thumb.n) = xyTmp.Y

        aAngle = WEBSPACRRight_bas.FN_CalcAngle(xyRightBotThumbRad, xyTmp)
        aInc = (aAngle - WEBSPACRRight_bas.FN_CalcAngle(xyRightBotThumbRad, xyRightTopThumbRad)) / 5
        For ii = 2 To 4
            aAngle = aAngle - aInc
            WEBSPACRRight_bas.PR_CalcPolar(xyRightBotThumbRad, aAngle, nRightBotThumbRad, xyTmp)
            If xyTmp.Y > xyRN(3).Y Then
                Thumb.n = Thumb.n + 1
                Thumb.X(Thumb.n) = xyTmp.X
                Thumb.y(Thumb.n) = xyTmp.Y
            End If
        Next ii

        aAngle = WEBSPACRRight_bas.FN_CalcAngle(xyRightTopThumbRad, xyRightBotThumbRad)
        aInc = (WEBSPACRRight_bas.FN_CalcAngle(xyRightTopThumbRad, xyRightWebT(5)) - aAngle) / 4
        For ii = 5 To 8
            aAngle = aAngle + aInc
            WEBSPACRRight_bas.PR_CalcPolar(xyRightTopThumbRad, aAngle, nRightTopThumbRad, xyTmp)
            If xyTmp.Y > xyRN(3).Y Then
                Thumb.n = Thumb.n + 1
                Thumb.X(Thumb.n) = xyTmp.X
                Thumb.y(Thumb.n) = xyTmp.Y
            End If
        Next ii

        WEBSPACRRight_bas.PR_DrawFitted(Thumb)

        'Draw thumb (Right)
        For ii = 1 To 5
            WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyRightWebT(ii))
        Next ii
        For ii = 1 To 3
            WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyRN(ii))
        Next ii
        WEBSPACRRight_bas.PR_MirrorPointInYaxis(xyRightWebFold.X, xyRightTopThumbRad)
        WEBSPACRRight_bas.PR_MirrorCurveInYaxis(xyRightWebFold.X, Thumb)
        If g_sRightWebSide = "Left" Then WEBSPACRRight_bas.PR_DrawLine(xyRightWebT(1), xyRW(3)) Else WEBSPACRRight_bas.PR_DrawLine(xyRightWebT(1), xyRW(2))

        WEBSPACRRight_bas.PR_DrawArc(xyRightTopThumbRad, xyRightWebT(1), xyRightWebT(2))
        X1(1) = xyRightWebT(2).X
        Y1(1) = xyRightWebT(2).Y
        X1(2) = xyRightWebT(3).X
        Y1(2) = xyRightWebT(3).Y
        X1(3) = xyRightWebT(4).X
        Y1(3) = xyRightWebT(4).Y
        X1(4) = xyRightWebT(5).X
        Y1(4) = xyRightWebT(5).Y
        Polycurve.n = 4
        Polycurve.X = X1
        Polycurve.y = Y1
        WEBSPACRRight_bas.PR_DrawPoly(Polycurve, Bulge1)

        'notch
        X2(1) = xyRN(1).X
        Y2(1) = xyRN(1).Y
        X2(2) = xyRN(2).X
        Y2(2) = xyRN(2).Y
        X2(3) = xyRN(3).X
        Y2(3) = xyRN(3).Y
        Polycurve.n = 3
        Polycurve.X = X2
        Polycurve.y = Y2
        WEBSPACRRight_bas.PR_DrawPoly(Polycurve, Bulge2)

        WEBSPACRRight_bas.PR_DrawFitted(Thumb)

        'PR_DrawCircle xyRightBotThumbRad, nRightBotThumbRad
        'PR_DrawCircle xyRightTopThumbRad, nRightTopThumbRad
        'PR_DrawCircle xyRightTopThumbRad, nRightNotchRad

        'Add notes etc.
        ARMDIA1.PR_SetLayer("Notes")

        'Insert Size (Text)
        'Position Insert text Size w.r.t Little Finger web
        nInsert = g_iRightWebInsertSize * RIGHT_WEB_EIGHTH
        If nInsert = 0.375 Then
            'Substitute 1/2" for 3/8" inch for printing only
            sText = Trim(ARMEDDIA1.fnInchesToText(0.5))
        Else
            sText = Trim(ARMEDDIA1.fnInchesToText(nInsert))
        End If
        If Mid(sText, 1, 1) = "-" Then sText = Mid(sText, 2) 'Strip leading "-" sign
        sText = sText & QQ
        sText = "INSERT " & sText

        If g_sRightWebSide = "Left" Then
            WEBSPACRRight_bas.PR_CalcMidPoint(xyRW(15), xyRW(14), xyTextAlign)
        Else
            WEBSPACRRight_bas.PR_CalcMidPoint(xyRW(10), xyRW(11), xyTextAlign)
        End If
        xyTextAlign.Y = xyTextAlign.Y - 0.5

        xyPt3 = xyTextAlign
        WEBSPACRRight_bas.PR_DrawText(sText, xyPt3, 0.1, 0)

        'Arrow
        WEBSPACRRight_bas.PR_DrawMarkerNamed("closed arrow", xyRW(9), 0.125, 0.05, 270)

        'Patient Details
        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
        WEBSPACRRight_bas.PR_MakeXY(xyPt3, xyPt3.X, xyPt3.Y - 0.3)
        If txtWorkOrder.Text = "" Then
            sWorkOrder = "-"
        Else
            sWorkOrder = txtWorkOrder.Text
        End If
        sText = txtSide.Text & Chr(10) & txtPatientName.Text & Chr(10) & sWorkOrder & Chr(10) & Trim(Mid(txtFabric.Text, 4))
        WEBSPACRRight_bas.PR_DrawMText(sText, xyPt3, 0)

        'Other patient details in black on layer construct
        ARMDIA1.PR_SetLayer("Construct")
        WEBSPACRRight_bas.PR_MakeXY(xyPt3, xyPt3.X, xyPt3.Y - 0.8)
        sText = txtFileNo.Text & Chr(10) & txtDiagnosis.Text & Chr(10) & txtAge.Text & Chr(10) & txtSex.Text
        WEBSPACRRight_bas.PR_DrawMText(sText, xyPt3, 0)
        WEBSPACRRight_bas.PR_DrawMarker()

        'This will only save if there is no existing glove data
        PR_UpdateDDE()
        PR_UpdateDBFields()
        FileClose(WEBSPACRRight_bas.fNum)
    End Sub

    Private Sub PR_DisableCir(ByRef Index As Short)
        'Read as part of labCir_DblClick
        labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        txtCir(Index).Enabled = False
        lblCir(Index).Enabled = False
    End Sub

    Private Sub PR_DrawTapeNotes(ByRef iVertex As Short, ByRef xyStart As WEBSPACRRight_bas.xy, ByRef xyEnd As WEBSPACRRight_bas.xy, ByRef nOffSet As Double, ByRef sZipperID As String)
        'Draws the the notes at ecah tape
        Dim xyDatum, xyInsert As WEBSPACRRight_bas.xy
        Dim nLength, aAngle, nDec As Double
        Dim sText As String
        Dim iInt As Short

        'Exit from sub if no Tape position Given as
        'this implies that the RightWebTapeNote is Empty
        '(EG with WEBSPACRRight_bas.GLOVE_NORMAL)
        If RightWebTapeNote(iVertex).iTapePos = 0 Then Exit Sub

        'Draw Construction line
        ARMDIA1.PR_SetLayer("Construct")
        WEBSPACRRight_bas.PR_DrawLine(xyStart, xyEnd)

        'Get center point on which text positions are based
        'Doing it this way accounts for case when xyStart.x > xyEnd.x
        If nOffSet > 0 Then
            aAngle = WEBSPACRRight_bas.FN_CalcAngle(xyStart, xyEnd)
            WEBSPACRRight_bas.PR_CalcPolar(xyStart, aAngle, nOffSet, xyDatum)
        Else
            WEBSPACRRight_bas.PR_CalcMidPoint(xyStart, xyEnd, xyDatum)
        End If

        'Insert Symbol
        ''------------------ARMDIA1.PR_InsertSymbol("glvTapeNotes", xyDatum, 1, 0)

        'Setup for Tape Lable
        WEBSPACRRight_bas.PR_AddDBValueToLast("Data", RightWebTapeNote(iVertex).sTapeText)
        'Circumference  in Inches and Eighths format
        iInt = Int(RightWebTapeNote(iVertex).nCir)
        WEBSPACRRight_bas.PR_AddDBValueToLast("TapeLengths", Trim(Str(iInt)))

        nDec = (RightWebTapeNote(iVertex).nCir - iInt)
        If nDec > 0 Then
            iInt = ARMDIA1.round((nDec / RIGHT_WEB_EIGHTH) * 10) / 10
            If iInt <> 0 Then WEBSPACRRight_bas.PR_AddDBValueToLast("TapeLengths2", Trim(Str(iInt)))
        End If
        WEBSPACRRight_bas.PR_AddDBValueToLast("Reduction", "14")

        'Check for Notes
        If RightWebTapeNote(iVertex).sNote <> "" Then
            'set text justification etc.
            ARMDIA1.PR_SetTextData(RIGHT_, VERT_CENTER, RIGHT_WEB_EIGHTH, CURRENT, CURRENT)
            WEBSPACRRight_bas.PR_MakeXY(xyInsert, xyDatum.X - 0.1, xyDatum.Y + 0.35)
            If RightWebTapeNote(iVertex).sNote <> "PLEAT" Then ARMDIA1.PR_SetLayer("Notes")
            WEBSPACRRight_bas.PR_DrawText(RightWebTapeNote(iVertex).sNote, xyInsert, 0.1, 0)
        End If
    End Sub

    Private Sub PR_EnableCir(ByRef Index As Short)
        'Read as part of labCir_DblClick
        labCir(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
        txtCir(Index).Enabled = True
        lblCir(Index).Enabled = True
    End Sub

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
                MsgBox("Invalid - Dimension has been entered", 48, "VEST Body - Dialogue")
                TextBox.Focus()
                FN_InchesValue = -1
                Exit For
            End If
        Next nn

        'Convert to inches
        nLen = fnDisplayToInches(Val(TextBox.Text))
        If nLen = -1 Then
            MsgBox("Invalid - Length has been entered", 48, "Glove Details")
            TextBox.Focus()
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
        If g_nRightWebUnitsFac <> 1 Then
            fnDisplayToInches = fnRoundInches(nDisplay * g_nRightWebUnitsFac) * iSign
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

    Private Sub PR_GetDlgAboveWrist()
        'General procedure to read the data given in the
        'dialogue controls and copy to module level variables.
        'Saves having to do this more than once.
        '
        '
        'NOTE:
        '    We ignore the fact that there may be missing tapes.
        '
        '    g_iRightWebWristPointer and g_iRightWebEOSPointer are used to indicate the
        '    start and finish tapes in the arrays and not in the
        '    txtExtCir() array

        Dim nCir As Double
        Dim ii As Short

        g_iRightWebFirstTape = -1
        g_iRightWebLastTape = -1
        g_iRightWebWristPointer = -1
        g_iRightWebEOSPointer = -1
        g_iRightWebNumTotalTapes = 0
        g_iRightWebNumTapesWristToEOS = 0

        'Glove type
        If optExtendTo(0).Checked = True Then
            g_RightWebExtendTo = RIGHT_WEB_GLOVE_NORMAL
        ElseIf optExtendTo(1).Checked = True Then
            g_RightWebExtendTo = RIGHT_WEB_GLOVE_ELBOW
        End If

        Dim iIndex As Short
        If g_RightWebExtendTo = RIGHT_WEB_GLOVE_NORMAL Then
            'Get values from Hand Data TAB
            For ii = 10 To 12
                nCir = FN_InchesValue(txtCir(ii))
                g_nRightWebCir(ii - 9) = nCir
            Next ii
            g_iRightWebFirstTape = 1

            'Assumes no holes in data
            If g_nRightWebCir(2) <= 0 Then
                g_iRightWebNumTotalTapes = 1
                g_iRightWebNumTapesWristToEOS = 1
                g_iRightWebLastTape = 1
            ElseIf g_nRightWebCir(3) > 0 Then
                g_iRightWebNumTotalTapes = 3
                g_iRightWebNumTapesWristToEOS = 3
                g_iRightWebLastTape = 3
                g_RightWebExtendTo = RIGHT_GLOVE_PASTWRIST
            Else
                g_iRightWebNumTotalTapes = 2
                g_iRightWebNumTapesWristToEOS = 2
                g_iRightWebLastTape = 2
                g_RightWebExtendTo = RIGHT_GLOVE_PASTWRIST
            End If
            '        g_iRightWebWristPointer = 1
            g_iRightWebWristPointer = 1
        Else
            If SSTab1.SelectedIndex <> 1 Then SSTab1.SelectedIndex = 1

            'Get First and last tape (if any)
            For ii = 8 To 23
                nCir = FN_InchesValue(txtExtCir(ii))
                If nCir > 0 And g_iRightWebFirstTape = -1 Then g_iRightWebFirstTape = ii - 7
                g_nRightWebCir(ii - 7) = nCir
            Next ii

            For ii = 23 To 8 Step -1
                If Val(txtExtCir(ii).Text.ToString()) > 0 Then
                    g_iRightWebLastTape = ii - 7
                    Exit For
                End If
            Next ii

            'Total number of tapes
            If (g_iRightWebLastTape = g_iRightWebFirstTape) Then
                If g_iRightWebLastTape <> -1 Then g_iRightWebNumTotalTapes = 1
            Else
                g_iRightWebNumTotalTapes = (g_iRightWebLastTape - g_iRightWebFirstTape) + 1
            End If
            'Get values from "Above Wrist" TAB
            'Wrist tape
            'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
            iIndex = cboDistalTape.SelectedIndex
            If iIndex = 0 Or iIndex = -1 Then
                g_iRightWebWristPointer = g_iRightWebFirstTape
            Else
                g_iRightWebWristPointer = iIndex
            End If

            'EOS tape
            'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
            iIndex = cboProximalTape.SelectedIndex
            If iIndex = 0 Or iIndex = -1 Then
                g_iRightWebEOSPointer = g_iRightWebLastTape
            Else
                'NB. The order of tapes is reversed in this list
                'starting at 19-1/2 and finishing at 0
                g_iRightWebEOSPointer = 17 - iIndex
            End If

            'Number of tapes between
            If (g_iRightWebWristPointer = g_iRightWebEOSPointer) Then
                'Just in case there are no tapes
                'And "first" and "last" are used for Wrist and EOS
                If g_iRightWebLastTape <> -1 Then g_iRightWebNumTapesWristToEOS = 1
            Else
                g_iRightWebNumTapesWristToEOS = (g_iRightWebEOSPointer - g_iRightWebWristPointer) + 1
            End If

            'Pleats
            If txtWristPleat1.Enabled = True Then g_nRightWebPleats(1) = Val(txtWristPleat1.Text) Else g_nRightWebPleats(1) = 0
            If txtWristPleat2.Enabled = True Then g_nRightWebPleats(2) = Val(txtWristPleat2.Text) Else g_nRightWebPleats(2) = 0
            If txtShoulderPleat2.Enabled = True Then g_nRightWebPleats(3) = Val(txtShoulderPleat2.Text) Else g_nRightWebPleats(3) = 0
            If txtShoulderPleat1.Enabled = True Then g_nRightWebPleats(4) = Val(txtShoulderPleat1.Text) Else g_nRightWebPleats(4) = 0

        End If
    End Sub

    Private Sub PR_GetExtensionDDE_Data()
        'Procedure to use the data given in the Form for the
        'glove extension (if any) along the arm
        '
        Dim nAge, Flap As Short
        Dim jj, ii, nn As Short
        Dim nValue As Double
        Dim iGms, iMMs, iRed As Short
        Dim sDiag As String

        SSTab1.SelectedIndex = 1

        'Defaults
        optExtendTo(0).Checked = True 'Normal Glove
        cboDistalTape.SelectedIndex = 0 '1st Tape
        cboProximalTape.SelectedIndex = 0 'Last Tape
        Flap = False

        'Set default pressure w.r.t diagnosis
        nAge = Val(txtAge.Text)

        'Glove extensions
        'Code for extended glove
        'Glove Option Buttons
        Dim nLen As Double
        If txtDataGlove.Text <> "" Then
            For ii = 0 To 5
                nValue = Val(Mid(txtDataGlove.Text, (ii * 2) + 1, 2))
                If ii = 0 Then
                    'Fold options
                    'Do nothing Off fold is default
                ElseIf ii = 1 And nValue >= 0 Then
                    'Pressure w.r.t Figuring and MMs
                    'ignore
                ElseIf ii = 2 Then
                    'Wrist Tape
                    cboDistalTape.SelectedIndex = nValue
                ElseIf ii = 3 Then
                    'Proximal Tape
                    cboProximalTape.SelectedIndex = nValue
                ElseIf ii = 4 Then
                    'Ignore
                ElseIf ii = 5 Then
                    'Ignore flaps
                End If
            Next ii
        End If

        'Glove to elbow and Glove to axilla
        'These can start at -3 (However to allow for possible
        'changes later we save up to -4-1/2 but ignore -4-1/2 for the
        'mean time)
        g_RightTapesFound = False
        If txtTapeLengthPt1.Text <> "" Then
            ii = 1
            For nn = 8 To 23
                nValue = Val(Mid(txtTapeLengthPt1.Text, (ii * 3) + 1, 3))
                If nValue > 0 Then
                    txtExtCir(nn).Text = CStr(nValue / 10)
                    nLen = FN_InchesValue(txtExtCir(nn))
                    If nLen <> -1 Then
                        ''--------------PR_GrdInchesDisplay(nn - 8, nLen)
                        g_RightTapesFound = True
                    End If
                    g_nRightWebCir(ii) = nLen
                Else
                    ''--------------PR_GrdInchesDisplay(nn - 8, 0)
                End If
                ii = ii + 1
            Next nn
        End If

        'Pleats
        For ii = 0 To 1
            ''-------nValue = Val(Mid(CType(WEBSPACRRight_bas.MainForm.Controls("txtWristPleat"), Object).Text, (ii * 3) + 1, 3))
            nValue = Val(Mid(txtWristPleat.Text, (ii * 3) + 1, 3))
            If ii = 0 And nValue > 0 Then
                ''----------CType(WEBSPACRRight_bas.MainForm.Controls("txtWristPleat1"), Object).Text = CStr(nValue / 10)
                txtWristPleat1.Text = CStr(nValue / 10)
            ElseIf ii = 1 And nValue > 0 Then
                ''----------CType(WEBSPACRRight_bas.MainForm.Controls("txtWristPleat2"), Object).Text = CStr(nValue / 10)
                txtWristPleat2.Text = CStr(nValue / 10)
            End If
            ''------------nValue = Val(Mid(CType(WEBSPACRRight_bas.MainForm.Controls("txtShoulderPleat"), Object).Text, (ii * 3) + 1, 3))
            nValue = Val(Mid(txtShoulderPleat.Text, (ii * 3) + 1, 3))
            If ii = 0 And nValue > 0 Then
                ''----------CType(WEBSPACRRight_bas.MainForm.Controls("txtShoulderPleat1"), Object).Text = CStr(nValue / 10)
                txtShoulderPleat1.Text = CStr(nValue / 10)
            ElseIf ii = 1 And nValue > 0 Then
                ''----------CType(WEBSPACRRight_bas.MainForm.Controls("txtShoulderPleat2"), Object).Text = CStr(nValue / 10)
                txtShoulderPleat2.Text = CStr(nValue / 10)
            End If
        Next ii

        ''---------CType(WEBSPACRRight_bas.MainForm.Controls("SSTab1"), Object).Tab = 0
        SSTab1.SelectedIndex = 0
    End Sub

    Private Sub PR_GrdInchesDisplay(ByRef iIndex As Short, ByRef nLen As Double)
        'CType(WEBSPACRRight_bas.MainForm.Controls("grdInches"), Object).Row = iIndex
        'CType(WEBSPACRRight_bas.MainForm.Controls("grdInches"), Object).Col = 0
        'CType(WEBSPACRRight_bas.MainForm.Controls("grdInches"), Object).Text = WEBSPACRRight_bas.fnInchesToText(nLen)
    End Sub

    Private Sub PR_UpdateDBFields()
        Dim sSymbol As String

        'Glove common
        sSymbol = "glovecommon"
        If txtUidGC.Text = "" Then
            'Insert a new symbol
            PrintLine(WEBSPACRRight_bas.fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
            PrintLine(WEBSPACRRight_bas.fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
            PrintLine(WEBSPACRRight_bas.fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x, xyO.y);")
            PrintLine(WEBSPACRRight_bas.fNum, "  }")
            PrintLine(WEBSPACRRight_bas.fNum, "else")
            PrintLine(WEBSPACRRight_bas.fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
            'Update DB fields
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
        End If

        'Glove
        sSymbol = "gloveglove"
        If txtUidGlove.Text = "" Then
            'Insert a new symbol
            PrintLine(WEBSPACRRight_bas.fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
            PrintLine(WEBSPACRRight_bas.fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
            If optHand(0).Checked = True Then
                'Left Hand
                PrintLine(WEBSPACRRight_bas.fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x+1.5 , xyO.y);")
            Else
                'Right Hand
                PrintLine(WEBSPACRRight_bas.fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyO.x+3 , xyO.y);")
            End If
            PrintLine(WEBSPACRRight_bas.fNum, "  }")
            PrintLine(WEBSPACRRight_bas.fNum, "else")
            PrintLine(WEBSPACRRight_bas.fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
            'Update DB fields
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "TapeLengths" & QCQ & txtTapeLengths.Text & QQ & ");")
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "TapeLengths2" & QCQ & txtTapeLengths2.Text & QQ & ");")
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "Sleeve" & QCQ & txtSide.Text & QQ & ");")
            PrintLine(WEBSPACRRight_bas.fNum, "SetDBData( hSym" & CQ & "Data" & QCQ & txtDataGlove.Text & QQ & ");")
        End If
    End Sub

    Private Sub PR_UpdateDDE()

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

        SSTab1.SelectedIndex = 0

        Dim iLen, ii As Short
        Dim sLen, sPacked As String

        'Initialise
        sPacked = ""

        'Pack Circumferences
        For ii = 0 To 10
            iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
            If iLen <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(iLen)), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sPacked = sPacked & sLen
        Next ii

        'Pack Lengths
        For ii = 0 To 8
            iLen = Val(txtLen(ii).Text) * 10 'Shift decimal place
            If iLen <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(iLen)), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sPacked = sPacked & sLen
        Next ii

        'Store to DDE text boxes
        txtTapeLengths.Text = sPacked

        'Tip options
        sPacked = ""
        For ii = 0 To 4
            sPacked = sPacked & "0"
        Next ii

        'CheckBoxes
        sPacked = sPacked & "0"
        sPacked = sPacked & "0"

        'Tapes past wrist (Store here due to limitations on DB field length
        'Pack Circumferences
        For ii = 11 To 12
            iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
            If iLen <> 0 Then
                sLen = New String(" ", 3)
                sLen = RSet(Trim(Str(iLen)), Len(sLen))
            Else
                sLen = New String(" ", 3)
            End If
            sPacked = sPacked & sLen
        Next ii

        'Insert size
        sLen = New String(" ", 3)
        sLen = RSet(Trim(Str(g_iRightWebInsertSize)), Len(sLen))
        sPacked = sPacked & sLen

        'Reinforced dorsal
        sPacked = sPacked & "0"
        'Thumb Options
        sPacked = sPacked & "0"

        'Insert Options
        sPacked = sPacked & "0"

        'Store to DDE text boxes
        txtTapeLengths2.Text = sPacked
        'Fabric
        If cboFabric.Text <> "" Then txtFabric.Text = cboFabric.Text
    End Sub

    Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
        'Allows the user to use enter as a tab
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Sub PR_Select_Text(ByRef Text_Box_Name As System.Windows.Forms.TextBox)
        If Not Text_Box_Name.Enabled Then Exit Sub
        Text_Box_Name.Focus()
        Text_Box_Name.SelectionStart = 0
        Text_Box_Name.SelectionLength = Len(Text_Box_Name.Text)
    End Sub

    Private Sub txtCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Enter
        Dim Index As Short = txtCir.GetIndex(eventSender)
        PR_Select_Text(txtCir(Index))
    End Sub

    Private Sub txtCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCir.Leave
        Dim Index As Short = txtCir.GetIndex(eventSender)

        Dim nLen As Double
        'Check that number is valid (-ve value => invalid)
        nLen = FN_InchesValue(txtCir(Index))

        'Display value in inches if Valid
        If nLen <> -1 Then lblCir(Index).Text = ARMEDDIA1.fnInchesToText(nLen)
    End Sub

    Private Sub txtLen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Enter
        Dim Index As Short = txtLen.GetIndex(eventSender)
        PR_Select_Text(txtLen(Index))
    End Sub

    Private Sub txtLen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLen.Leave
        Dim Index As Short = txtLen.GetIndex(eventSender)
        Dim nLen As Double
        Dim nTipCutBack As Double

        'Check that number is valid (-ve value => invalid)
        nLen = FN_InchesValue(txtLen(Index))
        If nLen < 0 Then
            lblLen(Index).Text = ""
            Exit Sub
        End If

        'Display length text
        If nLen > 0 Then
            lblLen(Index).Text = ARMEDDIA1.fnInchesToText(nLen)
        Else
            lblLen(Index).Text = ""
        End If
    End Sub
    Function Arccos(ByRef x As Double) As Double
        Arccos = System.Math.Atan(-x / System.Math.Sqrt(-x * x + 1)) + 1.5708
    End Function
    Private Sub PR_GetBlkAttributeValues()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecIdCommon As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("GLOVECOMMON") Then
                Return
            End If
            blkRecIdCommon = acBlkTbl("GLOVECOMMON")
            txtUidGlove.Text = "GLOVECOMMON"
            If blkRecIdCommon <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecIdCommon, OpenMode.ForRead)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If acAttDef.Tag.ToUpper().Equals("Fabric", StringComparison.InvariantCultureIgnoreCase) Then
                            txtFabric.Text = acAttDef.TextString
                        End If
                    End If
                Next
            End If
            Dim blkRecIdRight As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("GLOVERIGHT") Then
                Return
            End If
            blkRecIdRight = acBlkTbl("GLOVERIGHT")
            If blkRecIdRight <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecIdRight, OpenMode.ForRead)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If acAttDef.Tag.ToUpper().Equals("TapeLengths", StringComparison.InvariantCultureIgnoreCase) Then
                            txtTapeLengths.Text = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("TapeLengths2", StringComparison.InvariantCultureIgnoreCase) Then
                            txtTapeLengths2.Text = acAttDef.TextString
                        End If
                    End If
                Next
            End If
        End Using
    End Sub
    Private Sub PR_UpdateAge()
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("NewPatient", "NEWPATIENTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim pattern As String = "dd-MM-yyyy"
                Dim parsedDate As DateTime
                DateTime.TryParseExact(arr(8).Value, pattern, System.Globalization.CultureInfo.CurrentCulture,
                                      System.Globalization.DateTimeStyles.None, parsedDate)
                Dim startTime As DateTime = Convert.ToDateTime(parsedDate)
                Dim endTime As DateTime = DateTime.Today
                Dim span As TimeSpan = endTime.Subtract(startTime)
                Dim totalDays As Double = span.TotalDays
                Dim totalYears As Double = Math.Truncate(totalDays / 365)
                txtAge.Text = totalYears.ToString()
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
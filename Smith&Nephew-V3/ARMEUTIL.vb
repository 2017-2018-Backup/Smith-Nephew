Option Strict Off
Option Explicit On
Module ARMEUTIL
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
		Dim sID As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sID = sFileNo & sSide & sType

        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "if (hEnt) SetDBData( hEnt," & ARMDIA1.QQ & "ID" & ARMDIA1.QQ & ARMDIA1.CC & ARMDIA1.QQ & sID & ARMDIA1.QQ & ");")

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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		slayer = sFileNo & Mid(sSide, 1, 1) & nTape

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hLayer = Table(" & ARMDIA1.QQ & "find" & ARMDIA1.QCQ & "layer" & ARMDIA1.QCQ & slayer & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "if ( hLayer != %badtable)" & "Execute (" & ARMDIA1.QQ & "menu" & ARMDIA1.QCQ & "SetLayer" & ARMDIA1.QC & "hLayer);")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "else")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table (" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "layer" & ARMDIA1.QCQ & slayer & ARMDIA1.QCQ & "Tape Layer Data" & ARMDIA1.QCQ & "current" & ARMDIA1.QC & "Table(" & ARMDIA1.QQ & "find" & ARMDIA1.QCQ & "color" & ARMDIA1.QCQ & "DarkCyan" & ARMDIA1.QQ & "));")

    End Sub
	
	Sub PR_DeleteByID(ByRef sID As String)
        'Procedure to locate and delete all entitie that have the
        'string sID in a DRAFIX data base variable "ID"
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hChan=Open(" & ARMDIA1.QQ & "selection" & ARMDIA1.QCQ & "DB ID = '" & sID & "'" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "if(hChan)")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Close(" & ARMDIA1.QQ & "selection" & ARMDIA1.QC & "hChan);")
    End Sub

    Sub PR_DrawCircle(ByRef xyCen As ARMDIA1.XY, ByRef nRadius As Object)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a CIRCLE at the point given.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "circle" & ARMDIA1.QC & Str(xyCen.X) & ARMDIA1.CC & Str(xyCen.y) & ARMDIA1.CC & nRadius & ");")
    End Sub

    Sub PR_DrawPoly(ByRef Profile As ARMDIA1.curve)
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
        Dim ii As Short

        'Exit if nothing to draw
        If Profile.n <= 1 Then Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "poly" & ARMDIA1.QCQ & "polyline" & ARMDIA1.QQ)
        For ii = 1 To Profile.n
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, ARMDIA1.CC & Str(Profile.X(ii)) & ARMDIA1.CC & Str(Profile.y(ii)))
        Next ii
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, ");")

    End Sub

    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As ARMDIA1.XY, ByRef nHeight As Object)
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
        Dim nWidth As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nWidth = nHeight * ARMDIA1.g_nCurrTextAspect
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "AddEntity(" & ARMDIA1.QQ & "text" & ARMDIA1.QCQ & sText & ARMDIA1.QC & Str(xyInsert.X) & ARMDIA1.CC & Str(xyInsert.y) & ARMDIA1.CC & nWidth & ARMDIA1.CC & nHeight & ",0);")

    End Sub

    Sub PR_MakeXY(ByRef xyReturn As ARMDIA1.XY, ByRef x As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = x
        xyReturn.y = y
    End Sub

    Sub PR_PutTapeLabel(ByRef nTape As Short, ByRef xyStart As ARMDIA1.XY, ByRef nLength As Object, ByRef nMM As Object, ByRef nGrm As Object, ByRef nRed As Object)
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
        Dim xyPt As ARMDIA1.XY
        Dim nSymbolOffSet, nTextHt As Single

        'sTextList = "  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

        nSymbolOffSet = 0.6877
        nTextHt = 0.125

        PR_MakeXY(xyPt, xyStart.X, xyStart.y + nSymbolOffSet) 'Offset because symbol point is at top

        PR_CreateTapeLayer(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, nTape)

        PR_SetTextData(1, 32, 0.125, 0.6, 0)

        'Length text
        'N.B. format as Inches and eighths. With eighths offset up and left
        'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nInt = Int(CDbl(nLength)) 'Integer part of the length (before decimal point)

        'Decimal part of the length (after decimal point)
        'convert to 1/8ths and get nearest by rounding
        'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nDec = ARMDIA1.round((nLength - nInt) / 0.125)
        'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nDec = 8 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nDec = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nInt = nInt + 1
        End If

        'Draw Integer part
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 0.75)
        'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_DrawText(Trim(nInt), xyPt, nTextHt)

        'Draw eights part
        'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_MakeXY(xyPt, xyStart.X + 0.0625 + (Len(Trim(nInt)) * nTextHt * 0.8), xyStart.y + 0.75 + nTextHt / 1.5)
        'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nDec <> 0 Then PR_DrawText(Trim(nDec), xyPt, nTextHt / 1.5)

        'MMs text
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object nMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_DrawText(nMM & "mm", xyPt, nTextHt)

        'Grams text
        PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1.25)
        'UPGRADE_WARNING: Couldn't resolve default property of object nGrm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_DrawText(nGrm & "gm", xyPt, nTextHt)

        'Reduction text and circle round the text
        PR_SetTextData(2, 16, -1, -1, -1)
        PR_MakeXY(xyPt, xyStart.X + 0.25, xyStart.y + 1.625)
        'UPGRADE_WARNING: Couldn't resolve default property of object nRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_DrawText(Trim(nRed), xyPt, nTextHt)
        PR_DrawCircle(xyPt, 0.125)

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
        '

        'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nHoriz >= 0 And ARMDIA1.g_nCurrTextHorizJust <> nHoriz Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextHorzJust" & ARMDIA1.QC & nHoriz & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextHorizJust = nHoriz
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nVert >= 0 And ARMDIA1.g_nCurrTextVertJust <> nVert Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextVertJust" & ARMDIA1.QC & nVert & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextVertJust = nVert
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nHt >= 0 And ARMDIA1.g_nCurrTextHt <> nHt Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextHeight" & ARMDIA1.QC & nHt & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextHt = nHt
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nAspect >= 0 And ARMDIA1.g_nCurrTextAspect <> nAspect Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextAspect" & ARMDIA1.QC & nAspect & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextAspect = nAspect
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nFont >= 0 And ARMDIA1.g_nCurrTextFont <> nFont Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(ARMDIA1.fNum, "SetData(" & ARMDIA1.QQ & "TextFont" & ARMDIA1.QC & nFont & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.g_nCurrTextFont = nFont
        End If


    End Sub
End Module
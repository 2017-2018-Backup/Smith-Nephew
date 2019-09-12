Sub PR_AddEntityID (sFileNo As String, sSide As String, sType)
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

Sub PR_CreateTapeLayer (sFileNo As String, sSide As String, nTape As Variant)
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

Sub PR_DeleteByID (sID As String)
    'Procedure to locate and delete all entitie that have the
    'string sID in a DRAFIX data base variable "ID"
     Print #fNum, "hChan=Open("; QQ; "selection"; QCQ; "DB ID = '"; sID; "'"; QQ; ");"
     Print #fNum, "if(hChan)"
     Print #fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}"
     Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"
End Sub

Sub PR_DrawCircle (xyCen As XY, nRadius As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to draw a CIRCLE at the point given.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Print #fNum, "hEnt = AddEntity("; QQ; "circle"; QC; Str$(xyCen.x); CC; Str$(xyCen.y); CC; nRadius; ");"
End Sub

Sub PR_DrawPoly (Profile As curve)
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
        Print #fNum, CC; Str$(Profile.x(ii)); CC; Str$(Profile.y(ii))
    Next ii
    Print #fNum, ");"

End Sub

Sub PR_DrawText (sText As Variant, xyInsert As XY, nHeight As Variant)
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
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; Str$(xyInsert.x); CC; Str$(xyInsert.y); CC; nWidth; CC; nHeight; ",0);"
    
End Sub

Sub PR_MakeXY (xyReturn As XY, x As Double, y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.x = x
    xyReturn.y = y
End Sub

Sub PR_PutTapeLabel (nTape As Integer, xyStart As XY, nLength As Variant, nMM As Variant, nGrm As Variant, nRed As Variant)
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
    
    nSymbolOffSet = .6877
    nTextHt = .125

    PR_MakeXY xyPt, xyStart.x, xyStart.y + nSymbolOffSet   'Offset because symbol point is at top
    
    PR_CreateTapeLayer g_sFileNo, g_sSide, nTape

    PR_SetTextData 1, 32, .125, .6, 0

   'Length text
   'N.B. format as Inches and eighths. With eighths offset up and left
    nInt = Int(nLength)     'Integer part of the length (before decimal point)

   'Decimal part of the length (after decimal point)
   'convert to 1/8ths and get nearest by rounding
    nDec = round((nLength - nInt) / .125)
    If nDec = 8 Then
        nDec = 0
        nInt = nInt + 1
    End If

   'Draw Integer part
    PR_MakeXY xyPt, xyStart.x + .0625, xyStart.y + .75
    PR_DrawText Trim(nInt), xyPt, nTextHt

   'Draw eights part
    PR_MakeXY xyPt, xyStart.x + .0625 + (Len(Trim(nInt)) * nTextHt * .8), xyStart.y + .75 + nTextHt / 1.5
    If nDec <> 0 Then PR_DrawText Trim(nDec), xyPt, nTextHt / 1.5

   'MMs text
    PR_MakeXY xyPt, xyStart.x + .0625, xyStart.y + 1
    PR_DrawText nMM & "mm", xyPt, nTextHt

   'Grams text
    PR_MakeXY xyPt, xyStart.x + .0625, xyStart.y + 1.25
    PR_DrawText nGrm & "gm", xyPt, nTextHt

   'Reduction text and circle round the text
    PR_SetTextData 2, 16, -1, -1, -1
    PR_MakeXY xyPt, xyStart.x + .25, xyStart.y + 1.625
    PR_DrawText Trim(nRed), xyPt, nTextHt
    PR_DrawCircle xyPt, .125

End Sub

Sub PR_SetTextData (nHoriz As Variant, nVert As Variant, nHt As Variant, nAspect As Variant, nFont As Variant)
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


Attribute VB_Name = "MANUTILS"
'File:      MANUTILS.BAS
'Purpose:   Module of utilities for use by the manual
'           Glove
'
'
'Version:   1.01
'Date:      30.Jun.95
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-


   'Global constants for use with PR_SetTextData
   '
    Global Const HELVETICA = 10
    Global Const BLOCK = 0
    Global Const LEFT_ = 1
    Global Const RIGHT_ = 4
    Global Const HORIZ_CENTER = 2
    Global Const TOP_ = 8
    Global Const BOTTOM_ = 32
    Global Const VERT_CENTER = 16
    Global Const CURRENT = -1
   
    Option Explicit

Function Arccos(X As Double) As Double
     Arccos = Atn(-X / Sqr(-X * X + 1)) + 1.5708
End Function

Sub DebugMarker(xyPt As XY, sName As String)
    
    PR_DrawMarker xyPt
    PR_DrawText sName, xyPt, 0.1

End Sub


Function FN_CalcAngle(xyStart As XY, xyEnd As XY) As Double
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim X, Y, rAngle As Double
    
    X = xyEnd.X - xyStart.X
    Y = xyEnd.Y - xyStart.Y

   'Horizontal
    If X = 0 Then
        If Y > 0 Then
            FN_CalcAngle = 90
        Else
            FN_CalcAngle = 270
        End If
        Exit Function
    End If
    
   'Vertical (avoid divide by zero later)
    If Y = 0 Then
        If X > 0 Then
            FN_CalcAngle = 0
        Else
            FN_CalcAngle = 180
        End If
        Exit Function
    End If

   'All other cases
    rAngle = Atn(Y / X) * (180 / PI)         'Convert to degrees
    
    If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
    
    If Y > 0 Then
        FN_CalcAngle = rAngle
    Else
        FN_CalcAngle = rAngle + 180
    End If

End Function

Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

End Function

Static Function FN_CirLinInt(xyStart As XY, xyEnd As XY, xyCen As XY, nRad As Double, xyInt As XY) As Integer
   'Function to calculate the intersection between
   'a line and a circle.
   'Note:-
   '    Returns true if intersection found.
   '    The first intersection (only) is found.
   '    Ported from DRAFIX CAD DLG version.
   '

    Dim nSlope, nA, nB, nC, nK, nM, nCalcTmp, nRoot As Double
    Dim nSign As Integer

    nSlope = FN_CalcAngle(xyStart, xyEnd)

   'Horizontal Line
    If nSlope = 0 Or nSlope = 180 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.X + Sqr(nC) * nSign
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.Y = xyStart.Y
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Vertical Line
    If nSlope = 90 Or nSlope = 270 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                    'test each root
        While nSign > -2
            nRoot = xyCen.Y + Sqr(nC) * nSign
            If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                xyInt.Y = nRoot
                xyInt.X = xyStart.X
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Non-othogonal line
    If nSlope > 0 Then
        nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X)  'Slope
        nK = xyStart.Y - nM * xyStart.X                     'Y-Axis intercept
        nA = (1 + nM ^ 2)
        nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
        nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
        nCalcTmp = (nB ^ 2) - (4 * nC * nA)

        If (nCalcTmp < 0) Then
            FN_CirLinInt = False         'No Roots
            Exit Function
        End If
        nSign = 1
        While nSign > -2
            nRoot = (-nB + (Sqr(nCalcTmp) / nSign)) / (2 * nA)
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.Y = nM * nRoot + nK
                FN_CirLinInt = True
                Exit Function   'Return first root found
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False 'Should never get to here
    End If
    FN_CirLinInt = False

End Function

Static Function FN_EscapeSlashesInString(sAssignedString As Variant) As String
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
    sEscapedString = ""

End Function

Function max(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the maximum of two numbers
    If nFirst >= nSecond Then
        max = nFirst
    Else
        max = nSecond
    End If
End Function

Function min(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the minimum of two numbers
    If nFirst <= nSecond Then
        min = nFirst
    Else
        min = nSecond
    End If
End Function

Sub PR_AddDBValueToLast(sDBName As String, sDBValue As String)
   'The last entity is given by hEnt
   '
    Print #fNum, "if (hEnt) SetDBData( hEnt," & QQ & sDBName & QQ & CC & QQ; sDBValue & QQ & ");"

End Sub


Sub PR_AddVertex(xyPoint As XY, nBulge As Double)
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to add a Vertex to a polyline opened
   'by PR_StartPoly
   'Allow the use of a bulge factor, but start and end widths
   'will be 0.
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL,  are globals initialised by FN_Open
   '
    Print #fNum, "  AddVertex("; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; "0,0,"; nBulge; ");"

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
            xyReturn.Y = xyStart.Y
        Case 180
            xyReturn.X = xyStart.X - nLength
            xyReturn.Y = xyStart.Y
        Case 90
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y + nLength
        Case 270
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y - nLength
        Case Else
           'Convert from degees to radians
            nAngle = nAngle * PI / 180
            B = Sin(nAngle) * nLength
            A = Cos(nAngle) * nLength
            xyReturn.X = xyStart.X + A
            xyReturn.Y = xyStart.Y + B
        End Select

End Sub

Static Sub PR_DrawArc(xyCen As XY, xyArcStart As XY, xyArcEnd As XY)
   'Draws an arc
   'Restrictions
   '    1. Arc must be 180 degrees or less
    
Dim nRad, nStartAng, nEndAng, nDeltaAng, nOpp, nAdj, nSign
Dim xyMidPoint As XY

' ORIGINAL CODE ********************
'    nRad = FN_CalcLength(xyCen, xyArcStart)
'
'    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
'
'    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
'
'    If Side > 0 Then nDeltaAng = Abs(nEndAng - nStartAng) Else nDeltaAng = nEndAng - nStartAng
' ORIGINAL CODE ********************
'

    nRad = FN_CalcLength(xyCen, xyArcStart)
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)

   'Direction of arc
    nSign = Sgn(nEndAng - nStartAng)
    If Abs(nEndAng - nStartAng) > 180 Then
        nSign = 0 - nSign
    End If
   
   'Included angle
    PR_CalcMidPoint xyArcStart, xyArcEnd, xyMidPoint
    nAdj = FN_CalcLength(xyCen, xyMidPoint)
    nOpp = FN_CalcLength(xyArcStart, xyArcEnd) / 2

    If nAdj = 0 Then
        nDeltaAng = 180
    Else
        nDeltaAng = (Atn(nOpp / nAdj) * 2) * (180 / PI)
    End If

    nDeltaAng = nDeltaAng * nSign
    
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawFitted(Profile As Curve)
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
   '    For two vertex draw as a polyline (this degenerates to a Single line).
   '    For three vertex draw as a polyline (as no fitted curve can be drawn
   '    by a macro).
   '
    Select Case Profile.n
        Case 0 To 1
            Exit Sub
        Case 3
            PR_DrawPoly Profile
            Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
           'Warn the user to smooth the curve
           'Not required by Glove
 '           Print #fNum, "Display ("; QQ; "message"; QCQ; "OKquestion"; QCQ; "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line"; QQ; ");"
        Case Else
            Print #fNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
            For ii = 1 To Profile.n
                Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.Y(ii))
            Next ii
            Print #fNum, ");"
            Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
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
   'Ignore 0 length lines
    If xyStart.X = xyFinish.X And xyStart.Y = xyFinish.Y Then Exit Sub
    Print #fNum, "hEnt = AddEntity("
    Print #fNum, QQ; "line"; QC
    Print #fNum, "xyStart.x+"; Str$(xyStart.X); CC; "xyStart.y+"; Str$(xyStart.Y); CC
    Print #fNum, "xyStart.x+"; Str$(xyFinish.X); CC; "xyStart.y+"; Str$(xyFinish.Y)
    Print #fNum, ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
    

End Sub

Sub PR_DrawMarker(xyPoint As XY)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; "0.0625);"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawMarkerNamed(sName$, xyPoint As XY, nWidth, nHeight, nAngle)
   'Draw a Marker at the given point
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; sName; QC; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; nWidth; CC; nHeight; CC; nAngle; ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_DrawPoly(Profile As Curve)
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
    Print #fNum, CC; "xyStart.x+"; Str$(Profile.X(ii)); CC; "xyStart.y+"; Str$(Profile.Y(ii))
    Next ii
    Print #fNum, ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

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
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.Y); CC; nWidth; CC; nHeight; CC; g_nCurrTextAngle; ");"
    
End Sub


Sub PR_EndPoly()
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to end a POLYLINE
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '    STRING  sID
   '
   'Note:-
   '    fNum, CC, QQ, NL,  are globals initialised by FN_Open
   '
    Print #fNum, "EndPoly();"
    Print #fNum, "hEnt = UID (""find"", UID (""getmax"")) ;"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"

End Sub

Sub PR_InsertSymbol(sSymbol As String, xyInsert As XY, nScale As Single, nRotation As Single)
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
   '    The DRAFIX symbol library sPathJOBST + "\JOBST.SLB" exists
   '
   'Note:-
   '    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
   '    g_sPathJOBST is path to JOBST CAD System
   '
    Print #fNum, "SetSymbolLibrary("; QQ; FN_EscapeSlashesInString(g_sPathJOBST); "\\JOBST.SLB"; QQ; ");"
    Print #fNum, "Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ");"
    Print #fNum, "hEnt = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC
    Print #fNum, "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.Y); CC
    Print #fNum, Str$(nScale); CC & Str$(nScale); CC; Str$(nRotation); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
   
End Sub

Sub PR_MakeXY(xyReturn As XY, X As Double, Y As Double)
   'Utility to return a point based on the X and Y values
   'given
    xyReturn.X = X
    xyReturn.Y = Y

End Sub

Static Sub PR_MirrorCurveInYaxis(nXValue As Double, nTranslate As Double, Profile As Curve)
   'Procedure to mirror a curve in the y axis about
   'a line given by the X value
   '
    Dim ii%
    For ii = 1 To Profile.n
        Profile.X(ii) = (nXValue - (Profile.X(ii) - nXValue)) + nTranslate
    Next ii
    
End Sub

Sub PR_MirrorPointInYaxis(nXValue As Double, nTranslate As Double, xyPoint As XY)
   'Procedure to mirror a Point in the y axis about
   'a line given by the X value
   '
   xyPoint.X = (nXValue - (xyPoint.X - nXValue)) + nTranslate

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

Static Sub PR_RotateCurve(xyRotate As XY, aRotation#, Profile As Curve)
    Dim ii%, xyVertex As XY
    For ii = 1 To Profile.n
        PR_MakeXY xyVertex, Profile.X(ii), Profile.Y(ii)
        PR_RotatePoint xyRotate, aRotation, xyVertex
        Profile.X(ii) = xyVertex.X
        Profile.Y(ii) = xyVertex.Y
    Next ii
End Sub

Static Sub PR_RotatePoint(xyRotate As XY, aRotation#, xyPoint As XY)
    Dim aAngle#, nLength#

    aAngle = (FN_CalcAngle(xyRotate, xyPoint) + aRotation) Mod 360
    nLength = FN_CalcLength(xyRotate, xyPoint)
    PR_CalcPolar xyRotate, aAngle, nLength, xyPoint
    
End Sub

Sub PR_Setlayer(sNewLayer As String)
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

    If g_sCurrentLayer = sNewLayer Then Exit Sub
    g_sCurrentLayer = sNewLayer
    
    Print #fNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; sNewLayer; QQ; ");"
    Print #fNum, "if ( hLayer > %zero && hLayer != 32768)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Sub PR_SetLineStyle(sStyle As String)
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; sStyle; QQ; "));"
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

Static Sub PR_StartPoly()
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to start a POLYLINE the points will
   'be given by the PR_AddVertex and finished by PR_EndPoly
   'For this to work it assumes that the following DRAFIX variables
   'are defined and initialised
   '    XY      xyStart
   '    HANDLE  hEnt
   '    STRING  sID
   '
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
   '

    Print #fNum, "StartPoly (""PolyLine"");"


End Sub

Static Function round(ByVal nNumber As Double) As Integer
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


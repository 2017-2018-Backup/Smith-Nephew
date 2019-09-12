'Module:    MESHCALC.MAK
'Purpose:   Mesh axilla calculation
'
'Projects:  1. BODYDRAW.MAK
'           2. MESHDRAW.MAK
'
'Version:   1.00
'Date:      5.Nov.1997
'Author:    Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
'
'
'

    Option Explicit

    Type BiArc
        xyStart     As XY
        xyTangent   As XY
        xyEnd       As XY
        xyR1        As XY
        xyR2        As XY
        nR1         As Double
        nR2         As Double
    End Type

    Global Const ONE_and_THREE_QUARTER_GUSSET = 2.58
    Global Const BOYS_GUSSET = 2.8

Function FN_BiArcCurveVest (xyStart As XY, aStart As Double, xyEnd As XY, aEnd As Double, Profile As BiArc) As Integer
   '>>>>>>>>> WARNING <<<<<<<<<<<<<
   'This procedure FN_BiArcCurveVest is specific to the MESHVEST.MAK project
   '>>>>>>>>> WARNING <<<<<<<<<<<<<<
   '
   'Procedure to fit two arcs with a common tangent
   'between two points at the tangent angles specified.
   'Returning a BiArc curve.
   '
   'Input
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
    Dim aAxis#, nTmp#, nLength#
    Dim Theta1#, Theta2#, Phi1#, Phi2#
    Dim A#, B#, c#, d#, P#, S1#, S2#, C1#, C2#
    Dim Rs#, R1#, R2#, Rmin
    Dim uvR1 As XY, uvR2 As XY, uvOrigin As XY
'    Dim fError

   'Use a file for printing results for debug
   'Open file
'    fError = FreeFile
'    Open "C:\TMP\FIT_ERR.DAT" For Output As fError


   'Initially return false
    FN_BiArcCurveVest = False
    
   'Check for silly data and return
    If aStart = aEnd Then GoTo Error_Close
    If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then GoTo Error_Close

   'Translate tangent angle in XY system to the UV system.
   'Where the U-Axis is specified by the line xyStart,xyEnd
   '
    aAxis = FN_CalcAngle(xyStart, xyEnd)
    nLength = FN_CalcLength(xyStart, xyEnd)

    Theta1 = aStart - aAxis
    If Theta1 = 0 Then GoTo Error_Close 'Straight line

    If Theta1 < 0 Then Theta1 = 360 + Theta1
    Theta1 = Theta1 * (PI / 180)

    Theta2 = aEnd - aAxis
    If Theta2 < 0 Then Theta2 = 360 + Theta2
    Theta2 = Theta2 * (PI / 180)
   
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
        Phi1 = Abs((2 * PI) - Theta1)
    End If
    If Theta2 < PI Then
        Phi2 = Theta2
    Else
        Phi2 = Abs((2 * PI) - Theta2)
    End If

   'Calculate R1 and R2
    S1 = Abs(Sin(Theta1))
    C1 = (-Sin(Theta1) * Cos(Theta1)) / S1
    S2 = Abs(Sin(Theta2))
    C2 = (Sin(Theta2) * Cos(Theta2)) / S2

    P = FN_CalcLength(xyStart, xyEnd)

    If Phi1 <> Phi2 Then
        A = S1 + S2
        B = (S1 * S2) - (C1 * C2) + 1
        c = S2
        Rs = (P * c) / B
        nTmp = (c ^ 2) - (c * A) + (B / 2)
       'As this is a root we check that it is not -ve
        If nTmp < 0 Then GoTo Error_Close
        nTmp = (P * Sqr(nTmp)) / B
        If Phi1 > Phi2 Then
            R1 = Rs - nTmp
        Else
            R1 = Rs + nTmp
        End If
        If R1 <= 0 Then GoTo Error_Close
        d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
        R2 = R1 - d
    Else
        A = 2 * Sin(Theta1)
        B = (Sin(Theta1) ^ 2) - (Cos(Theta1) ^ 2) + 1
        R1 = (P * A) / (2 * B)
        R2 = R1
    End If

   'The radi R1 and R2 must be greater than the specified
   'minimum Rmin
    Rmin = nLength * 3 / 7   '<<<<<<<<<<<<<<<< N.B.  <<<<<<<<<<<<<<<<<
    If R1 < Rmin Then
        R1 = Rmin
        d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
        R2 = R1 - d
    End If
    If R2 < Rmin Then
        R2 = Rmin
        R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
    End If

   'Check for errors
    If R1 < 0 Or R2 < 0 Then GoTo Error_Close

   'Using the calculated radi, create BiArc Curve
   'Start and end points of bi-arc curve
    Profile.xyStart = xyStart
    Profile.xyEnd = xyEnd

   'Centers of arcs
   'Get UV co-ordinates
    PR_MakeXY uvR1, R1 * S1, R1 * C1
    PR_MakeXY uvR2, P - (R2 * S2), R2 * C2
    PR_MakeXY uvOrigin, 0, 0

   'Translate to XY co-ordinates
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2
   'Tangent point on arc
    If R1 < R2 Then
        PR_CalcPolar Profile.xyR1, FN_CalcAngle(Profile.xyR2, Profile.xyR1), R1, Profile.xyTangent
    Else
        PR_CalcPolar Profile.xyR1, FN_CalcAngle(Profile.xyR1, Profile.xyR2), R1, Profile.xyTangent
    End If
   'radi of arcs
    Profile.nR1 = R1
    Profile.nR2 = R2
    
   'Test that the Tangent point lies between the start and end points
    If FN_CalcLength(xyStart, Profile.xyTangent) > nLength Then GoTo Error_Close
    If FN_CalcLength(xyEnd, Profile.xyTangent) > nLength Then GoTo Error_Close

   'return true as we have a sucessful fit
    FN_BiArcCurveVest = True

    Exit Function
    
Error_Close:

End Function

Function FN_CalcAxillaMeshVest (iIteration As Integer, xyRaglanStart As XY, xyRaglanEnd As XY, xyMeshStart As XY, nFlip%, nMeshLength#, nAlongRaglan#, Meshprofile As BiArc, MeshSeam As BiArc) As Integer
   'Subroutine to calculate a mesh w.r.t Mesh axilla in the
   'Vest and Vest sleeves

   'Due to problems with the CALC BiARC function we calculate all for the lower
   'half and then flip the result
   'A cheat but what the hell

    Dim aStart#, aAngle#, aInitialStart#, aInitialEnd#, aEnd#, aMeshEnd#, aDeltaEnd#, aStartDelta#
    Dim nStartLength#, nEndLength#, nStep#
    Dim nLength#, nTol#, nRadiusTol#, nSeam#
    Dim bMeshFound%, ii%, iSign#, aAxis#, aEndFactor#, iError%
    Dim xyMeshEnd As XY, xyTmp  As XY
    Dim cDummyCurve As curve

   'intially start as False
    FN_CalcAxillaMeshVest = False

   'Check for silly data
    If xyRaglanStart.X = xyRaglanEnd.X And xyRaglanStart.Y = xyRaglanEnd.Y Then Exit Function
    If xyMeshStart.X = xyRaglanEnd.X And xyMeshStart.Y = xyRaglanEnd.Y Then Exit Function
    If xyRaglanStart.X = xyMeshStart.X And xyRaglanStart.Y = xyMeshStart.Y Then Exit Function
    If FN_CalcLength(xyMeshStart, xyRaglanStart) >= nMeshLength Then Exit Function

   'initialize variables
'    nTol = .03125
    nTol = .04
    bMeshFound = False
    nStep = -INCH1_16
    
    If nAlongRaglan > 0 Then
        nStartLength = nAlongRaglan
        nEndLength = nAlongRaglan                      'do at least Once
        aDeltaEnd = 160                                'Must find
    Else
        aDeltaEnd = 90
        nStep = INCH1_16
        If nMeshLength = ONE_and_THREE_QUARTER_GUSSET Then
            nStartLength = 1.75'
        Else
            nStartLength = 1.875'
        End If
        nEndLength = nStartLength + 1  '

    End If
    
   'Vest Mesh
    aAxis = FN_CalcAngle(xyMeshStart, xyAxilla)
    If g_sCallingApplication = "vestmesh" Then
        aStartDelta = 15
        aInitialStart = 267
        If nMeshLength = ONE_and_THREE_QUARTER_GUSSET Then
            nRadiusTol = 0
        Else
            nRadiusTol = 0
        End If
    End If

   'Sleeve mesh
    If g_sCallingApplication = "sleevemesh" Then
        aStartDelta = 60
        If nMeshLength = ONE_and_THREE_QUARTER_GUSSET Then
            nRadiusTol = .4
            If iIteration > 0 Then nRadiusTol = 0
            aInitialStart = 297
        Else
            nRadiusTol = 0
            aInitialStart = 290
        End If
    End If
   
   For nLength = nStartLength To nEndLength Step nStep

    For aStart = aInitialStart To (aInitialStart + aStartDelta) Step 1
'PR_CalcPolar xyMeshStart, aStart, 1, xyTmp
'PR_DrawMarker xyTmp
   
           'Get point on raglan
            PR_CalcRaglan xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, xyMeshEnd, aInitialEnd, nLength
'    PR_CalcPolar xyMeshEnd, aInitialEnd, 3, xyTmp
'    PR_DrawLine xyTmp, xyMeshEnd
    
           'Flip the point and angle
            If nFlip Then
               'mirror in the line (xyMeshStart, xyRaglanStart)
                PR_CalcPolar xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, xyMeshEnd) + aAxis), FN_CalcLength(xyMeshStart, xyMeshEnd), xyMeshEnd
                aMeshEnd = ((360 - aInitialEnd) + aAxis) + 180
            Else
                aMeshEnd = (aInitialEnd + 180)
            End If
    
           'End angle iteration - Get biarc curve that fits the given length to +/- nTol
           'Hold start change end
            For aEnd = (aMeshEnd - (aDeltaEnd / 2)) To (aMeshEnd + (aDeltaEnd / 2)) Step 1
'PR_CalcPolar xyMeshEnd, aEnd, 1, xyTmp
'PR_DrawMarker xyTmp
    
                If FN_FitBiArcAtGivenLength(xyMeshStart, aStart, xyMeshEnd, aEnd, nMeshLength, nTol, nRadiusTol, Meshprofile) Then
                    bMeshFound = True
                    Exit For
                End If
            Next aEnd
            If bMeshFound Then
                FN_CalcAxillaMeshVest = True
                nAlongRaglan = nLength
                Exit For
            End If
    

  
    Next aStart
    If bMeshFound Then Exit For
    Next nLength


   'Calculate seam line
    nSeam = .125

   'Calculate MeshSeam based on MeshProfile
    If nFlip Then
       'mirror in the line (xyMeshStart, xyRaglanStart)
        PR_CalcPolar xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, Meshprofile.xyEnd) + aAxis), FN_CalcLength(xyMeshStart, Meshprofile.xyEnd), Meshprofile.xyEnd
        PR_CalcPolar xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, Meshprofile.xyTangent) + aAxis), FN_CalcLength(xyMeshStart, Meshprofile.xyTangent), Meshprofile.xyTangent
        PR_CalcPolar xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, Meshprofile.xyR1) + aAxis), FN_CalcLength(xyMeshStart, Meshprofile.xyR1), Meshprofile.xyR1
        PR_CalcPolar xyMeshStart, (360 - FN_CalcAngle(xyMeshStart, Meshprofile.xyR2) + aAxis), FN_CalcLength(xyMeshStart, Meshprofile.xyR2), Meshprofile.xyR2
    End If

    MeshSeam = Meshprofile

   'Get seam intersection on raglan
    nLength = nLength - nSeam
    PR_CalcRaglan xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, MeshSeam.xyEnd, aInitialEnd, nLength
   
   'Calculate Start, End and Tangent
    aAngle = FN_CalcAngle(Meshprofile.xyR1, Meshprofile.xyStart)
    MeshSeam.nR1 = Meshprofile.nR1 - nSeam
    PR_CalcPolar Meshprofile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyStart

    aAngle = FN_CalcAngle(Meshprofile.xyR1, Meshprofile.xyTangent)
    PR_CalcPolar Meshprofile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyTangent
    MeshSeam.nR2 = Meshprofile.nR2 - nSeam
    
'PR_DrawMarker MeshSeam.xyR2
'PR_DrawText "MeshSeam.xyStart", MeshSeam.xyStart, .1


End Function

Function FN_FitBiArcAtGivenLength (xyStart As XY, aStart#, xyEnd As XY, aEnd#, nLength#, nLengthTol#, nRadiusTol#, Mesh As BiArc) As Integer
    
    Dim aA1#, aA2#, nArc#, aArc#
        
    FN_FitBiArcAtGivenLength = False
    If FN_BiArcCurveVest(xyStart, aStart, xyEnd, aEnd, Mesh) Then
       'Calculate Length
       'First arc
        aA1 = FN_CalcAngle(Mesh.xyR1, Mesh.xyStart)
        aA2 = FN_CalcAngle(Mesh.xyR1, Mesh.xyTangent)
        aArc = aA2 - aA1
        If Abs(aArc) > 180 Then aArc = 360 - Abs(aArc)
        nArc = (aArc * (PI / 180)) * Mesh.nR1
       'Second arc
        aA1 = FN_CalcAngle(Mesh.xyR2, Mesh.xyTangent)
        aA2 = FN_CalcAngle(Mesh.xyR2, Mesh.xyEnd)
        aArc = aA2 - aA1
        If Abs(aArc) > 180 Then aArc = 360 - Abs(aArc)
        nArc = nArc + ((aArc * (PI / 180)) * Mesh.nR2)

       'Check that the length is within the tolerence
        If nArc > (nLength - nLengthTol) And nArc < (nLength + nLengthTol) Then
            FN_FitBiArcAtGivenLength = True
            If nRadiusTol > 0 And Abs(Mesh.nR1 - Mesh.nR2) > nRadiusTol Then FN_FitBiArcAtGivenLength = False
        End If

    End If

End Function

Static Sub PR_CalcRaglan (xyStart As XY, xyEnd As XY, Flip%, Raglan As curve, xyOpen As XY, aOpen As Double, nRadiusForRequiredPoint As Double)
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
    Dim aAngle#, aPrevAngle#, nTraversedLen#, nLength#, nOpenLength#, bOpenFound  As Integer
    Dim aRaglan#, ii%, aAxis#, nRadius#

   'On an error the Raglan Curve is returned as a minimum of 2 points
   'these being the given start and end points.
    Raglan.n = 2
    Raglan.X(1) = xyStart.X
    Raglan.Y(1) = xyStart.Y
    Raglan.X(2) = xyEnd.X
    Raglan.Y(2) = xyEnd.Y

   'Check for silly data
    If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then Exit Sub
    
    If g_sSleeveType = "Sleeveless" Then
        sFileName = g_sPathJOBST & "\TEMPLTS\BODYCURV.DAT"
    Else
        sFileName = g_sPathJOBST & "\TEMPLTS\VESTCURV.DAT"
    End If
   'Check to see if the raglan curve has been loaded from file,
   'so we don't load it twice
    If iSegments = 0 Then
        fFileNum = FreeFile
        If FileLen(sFileName) = 0 Then
            MsgBox "Can't open the file " & sFileName & ".  Unable to draw the raglan curve.", 48
            Exit Sub
        End If
        
        Open sFileName For Input As fFileNum
        Do While Not EOF(fFileNum)
            iSegments = iSegments + 1
            Input #fFileNum, nSegLen(iSegments), nSegAngle(iSegments)
            Loop
        Close #fFileNum
    End If

   'Using the given points loop through the raglan curve to establish the
   'raglan angle.
   'Note:
   'This code is a little bit cumbersome but it is based on DRAFIX macro
   'code that is known to work and is production accepted.
   '
    nRadius = FN_CalcLength(xyStart, xyEnd)
    
    aRaglan = -1
    xy1.X = 0: xy1.Y = 0
    xy0.X = 0: xy0.Y = 0
    aPrevAngle = 0
    aAngle = nSegAngle(1)
    nLength = nSegLen(1)
    nTraversedLen = 0       'For use later

    For ii = 2 To iSegments
        aAngle = aAngle + aPrevAngle
        PR_CalcPolar xy1, aAngle, nLength, xy2
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
        MsgBox "Can't draw the draw the raglan curve with this data.", 48
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
    Raglan.X(1) = xyStart.X
    Raglan.Y(1) = xyStart.Y
    For ii = 2 To iSegments
        aAngle = aAngle + aPrevAngle
        PR_CalcPolar xy1, aAngle, nLength, xy2
        If FN_CirLinInt(xy1, xy2, xyStart, nRadius, xy4) Then
            Raglan.X(ii) = xyEnd.X
            Raglan.Y(ii) = xyEnd.Y
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
        Raglan.Y(ii) = xy2.Y
        If nTraversedLen + nLength >= nOpenLength And Not bOpenFound And nRadiusForRequiredPoint <= 0 Then
            bOpenFound = True
            aOpen = FN_CalcAngle(xy1, xy2)
            PR_CalcPolar xy1, aOpen, (nTraversedLen + nLength) - nOpenLength, xyOpen
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
            PR_MakeXY xy1, Raglan.X(ii), Raglan.Y(ii)
            nLength = FN_CalcLength(xyStart, xy1)
            aAngle = FN_CalcAngle(xyStart, xy1)
            aAngle = (360 - (aAngle - aAxis) + aAxis)
            PR_CalcPolar xyStart, aAngle, nLength, xy2
            Raglan.X(ii) = xy2.X
            Raglan.Y(ii) = xy2.Y
        Next ii
        nLength = FN_CalcLength(xyStart, xyOpen)
        aAngle = FN_CalcAngle(xyStart, xyOpen)
        aAngle = (360 - (aAngle - aAxis) + aAxis)
        PR_CalcPolar xyStart, aAngle, nLength, xyOpen

    End If
            

End Sub


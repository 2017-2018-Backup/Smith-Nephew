Attribute VB_Name = "GLVEXTEN"
'File:      GLVEXTEN.BAS
'Purpose:   Extend glove to Elbow and Axilla
'
'Version:   1.01
'Date:      17.Jan.96
'Author:    Gary George
'
'Projects:  CADGLOVE.MAK
'           MANGLOVE.MAK
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
'
'   This module is designed to be common to both the
'   CAD Glove and the Manual Glove.  Therefor we make use
'   of the global level variables indicated by g_
'
'   This may make the structure of the program a little
'   more obscure but the concept is that all procedures and
'   related module variables are in the one module GLVEXTEN.BAS
'
'
    Option Explicit
    
    Public Type TapeData
        nCir        As Double
        iMMs        As Integer
        iRed        As Integer
        iGms        As Integer
        sNote       As String
        iTapePos    As Integer
        sTapeText   As String
    End Type
    
        
    Public Type Curve
        n As Integer
        X(1 To 100) As Double
        Y(1 To 100) As Double
    End Type


       'XY data type to represent points
    Public Type XY
        X As Double
        Y As Double
    End Type

    
    Public Type BiArc
        xyStart     As XY
        xyTangent   As XY
        xyEnd       As XY
        xyR1        As XY
        xyR2        As XY
        nR1         As Double
        nR2         As Double
    End Type
    
    Public Type Finger
        DIP As Double
        PIP As Double
        Len As Double
        WebHt As Double
        Tip As Integer
        xyUlnar As XY
        xyRadial As XY
    End Type

   'Fingers
    Global Const DIP = 1
    Global Const PIP = 2
    Global Const THUMB = 3
    Global Const THUMB_LEN = 2
    Global Const FINGER_LEN = 1
    Global Const g_sTapeText = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"


   'Tapes etc
    Global Const WRIST = 1
    Global Const PALM = 2
    Global Const TAPE_ONE_HALF = 3
    Global Const TAPE_THREE = 4
    Global Const TAPE_FOUR_HALF = 5
    Global Const TAPE_SIX = 6
    Global Const TAPE_SEVEN_HALF = 7
    Global Const TAPE_NINE = 8

   'Misc
    Global Const EIGHTH = 0.125
    Global Const SIXTEENTH = 0.0625
    Global Const QUARTER = 0.25

   'Reduction Chart constants and Variables
    Const NOFF_FINGER_RED = 32
    Const NOFF_ARM_RED = 76
    Const NOFF_LENGTH_RED = 24
    Const FIGURED_MINIMUM = 6

   'Hand Reduction charts
    Dim iFingerRed(3, NOFF_FINGER_RED) As Integer
    Dim nArmRed(8, NOFF_ARM_RED) As Double
    Dim nLengthRed(2, NOFF_LENGTH_RED) As Double

   'Arm reduction chart
    Global Const LOW_MODULUS = 160
    Global Const HIGH_MODULUS = 340
    Const NOFF_MODULUS = 19

    Dim g_iModulusIndex As Integer
    Dim g_iPowernet(NOFF_MODULUS, 23)  As Integer


   'Arm variables
    Global Const NOFF_ARMTAPES = 16
    Global Const ELBOW_TAPE = 9
    Global Const ARM_PLAIN = 0
    Global Const ARM_FLAP = 1
    Global Const GLOVE_NORMAL = 0
    Global Const GLOVE_ELBOW = 1
    Global Const GLOVE_AXILLA = 2


        

Function FN_FigureFinger(iSelect As Integer, nFingerCir As Double, iInsert As Integer) As Double
   'procedure to return the Figured circumference of a finger
   '
   ' iSelect    a constant, one of DIP, PIP, THUMB
   ' nFingerCir Dimension in inches in
   '            1" >= nFingerCir <= 4.825"
   ' iInsert    Insert range to be used
   '            1 >= iNinsert <= 6
   '            iInsert = 1, use looked up figure
   '            iInsert > 1, add 2 to DIP & PIP, add 1 to THUMB
   '                         to looked up figure
   '            iInsert = 6, subtract 1 from DIP, PIP, THUMB
   '                         to looked up figure
   '

    Dim iIndex%, iValue%

   'Check data is within range
    FN_FigureFinger = 0#
    If iSelect < DIP Or iSelect > THUMB Then Exit Function
    If iInsert < 1 Or iInsert > 6 Then Exit Function
    If nFingerCir < 1 Or nFingerCir > 4.875 Then Exit Function
           
   'Convert nFingerCir to an index
   'Note integer division \
    iIndex = (((nFingerCir - 1) * 1000) \ 125) + 1

   'Look up value
    iValue = iFingerRed(iSelect, iIndex)

    If iInsert = 6 Then
        iValue = iValue + 1
    ElseIf iSelect = THUMB Then
        iValue = iValue - (iInsert - 1)
    Else
        iValue = iValue - ((iInsert - 1) * 2)
    End If

    If iValue < FIGURED_MINIMUM Then iValue = FIGURED_MINIMUM

   'Convert from 16ths to inches
    FN_FigureFinger = iValue * 0.0625
        
End Function




Function FN_FigureLength(iSelect As Integer, nLen As Double) As Double
   'Procedure to return the Figured length of a finger
   'or thumb
   '
   ' iSelect    A constant, one of THUMB or FINGER
   ' nLen       Dimension in inches in
   '            1" >= nLen <= 3.875
    
    Dim iIndex%, iValue%

   'Check data is within range
    FN_FigureLength = 0#
    If iSelect <> THUMB_LEN And iSelect <> FINGER_LEN Then Exit Function
    If nLen >= 1 And nLen <= 3.875 Then
       'Convert nLen to an index
       'Note integer division \
        iIndex = (((nLen - 1) * 1000) \ 125) + 1
       'Look up value
        FN_FigureLength = nLengthRed(iSelect, iIndex)

    ElseIf nLen < 1 Then
       'Return given value unmodified
        FN_FigureLength = nLen

    ElseIf nLen > 3.875 And iSelect = THUMB_LEN Then
        FN_FigureLength = nLen - 0.125
        
    ElseIf nLen > 3.875 And iSelect = FINGER_LEN Then
        FN_FigureLength = nLen - 0.25
        
    End If


End Function

Function FN_FigureTape(iSelect As Integer, nTapeCir As Double) As Double
   'Procedure to return the Figured circumference of an ARM tape
   'This also is used to figure the Palm and wrist
   '
   ' iSelect    A constant, one off
   '            WRIST, PALM, TAPE_ONE_HALF ... TAPE_NINE
   ' nTapeCir   Dimension in inches in
   '            3.5" >= nTapeCir <= 12.875"
    
   'NOTES;
   '    The PALM and WRIST chart only goes up to 12.375
   '    if this value is exceeded then a value is extrapolated
   '    from the 12.875 value.
   '

    Dim iIndex%, iValue%, iAdditional#

   'Check data is within range
    FN_FigureTape = 0#
    If iSelect < WRIST Or iSelect > TAPE_NINE Then Exit Function
    If nTapeCir < 3.5 Then Exit Function
           
   'Convert nTapeCir to an index
   'Note integer division \
    If nTapeCir > 12.875 Then
       'Extrapolate a value for this case
       'Get index of the last value in the table (ie 12.875)
        iIndex = (((12.875 - 3.5) * 1000) \ 125) + 1
       'Look up value
       'For each 1/2" greater than 12.875 add an eighth
       'to the returned value
        iAdditional = ((nTapeCir - 12.875) \ 0.5) + 1

        FN_FigureTape = nArmRed(iSelect, iIndex) + (iAdditional * 0.125)

    Else
        iIndex = (((nTapeCir - 3.5) * 1000) \ 125) + 1
       'Look up value
        FN_FigureTape = nArmRed(iSelect, iIndex)
    End If

End Function

Function FN_GetGrams(ByVal iReduction As Integer) As Integer
   'Function that looks up the "Powernet" grams / tension chart
   'and returns the Grams for a given reduction value
   '
   'Used to back calculate from a given reduction
   '
   ' Input
   '        iReduction      Reduction given
   '
   ' Module Level variables
   '        g_iModulusIndex
   '        g_iPowernet()
   '
    If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
        FN_GetGrams = -1
        Exit Function
    End If

    If iReduction < 10 Then iReduction = 10
    If iReduction > 32 Then iReduction = 32
    
    FN_GetGrams = g_iPowernet(g_iModulusIndex, (iReduction - 10) + 1)

    
End Function

Function FN_GetReduction(iGrams) As Integer
   'Function that looks up the "Powernet" grams / tension chart
   'and returns the reduction value
   '
   ' Input
   '        iGrams      Grams calculated from the data
   '
   ' Module Level variables
   '        g_iModulusIndex
   '        g_iPowernet()
   '
    If g_iModulusIndex < 1 Or g_iModulusIndex > NOFF_MODULUS Then
        FN_GetReduction = -1
        Exit Function
    End If

    Dim ii%, iValue%, iPrevValue%
    
    Select Case iGrams
        Case Is <= g_iPowernet(g_iModulusIndex, 1)
            FN_GetReduction = 10

        Case Is >= g_iPowernet(g_iModulusIndex, 23)
            FN_GetReduction = 32

        Case Else
           'Return value closest
            iPrevValue = 0
            For ii = 1 To 23
                iValue = g_iPowernet(g_iModulusIndex, ii)
                If iValue > iGrams Then Exit For
                iPrevValue = iValue
            Next ii

            If Abs(iGrams - iPrevValue) < Abs(iGrams - iValue) Then
                FN_GetReduction = ii + 8
            Else
                FN_GetReduction = ii + 9
            End If

    End Select


End Function

Static Function FN_LengthWristToEOS() As Double
   'Calculates the distance from the EOS to the
   'wrist based on values given in the dialogue
   '
    Dim ii%, nLen#, nSpace#

   'Get the data from the dialogue
    PR_GetDlgAboveWrist

    nLen = 0

    Select Case g_ExtendTo
     Case GLOVE_NORMAL
        For ii = 1 To g_iNumTapesWristToEOS - 1
            nSpace = 1.375
 '           If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
 '           If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            nLen = nLen + nSpace
        Next ii
        If nLen = 0 Then
           'If no tapes after wrist then default to 0.75
            nLen = 0.75
        End If
     Case GLOVE_ELBOW
        For ii = 1 To g_iNumTapesWristToEOS - 1
            nSpace = 1.375
            If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
            If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
            If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
            nLen = nLen + nSpace
        Next ii
     
     Case GLOVE_AXILLA
        For ii = 1 To g_iNumTapesWristToEOS - 1
            nSpace = 1.375
            If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
            If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
            If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
            nLen = nLen + nSpace
        Next ii
        If g_EOSType = ARM_FLAP Then
        End If
    End Select

    FN_LengthWristToEOS = nLen

End Function

Static Function FN_ValidateExtensionData() As String
   'Procedure to validated the extension data
    Dim sError As String, NL$
    Dim iNoPleats%, ii%
    
    NL = Chr$(13)
    sError = ""     'Initialise because of static

   'If normal glove then exit
    If g_ExtendTo = GLOVE_NORMAL Then
        FN_ValidateExtensionData = ""
        Exit Function
    End If

   'Check pleats
    iNoPleats = 0
    For ii = 1 To 4
        If g_nPleats(ii) > 0 Then iNoPleats = iNoPleats + 1
    Next ii
    If iNoPleats + 1 > g_iNumTapesWristToEOS Then
        sError = sError & "Number of pleats exceeds availble spaces between the arm tapes!" & NL & "Disable pleats by Double Clicking on pleat label." & NL
    End If

   'Check that calculate has been used
    If g_CalculatedExtension = False Then
        sError = sError & "The arm has not been calculated" & NL
    End If

   'Check on style D flaps
    If g_EOSType = ARM_FLAP And InStr(1, g_sFlapType, "D") > 0 And g_nWaistCir = 0 Then
        sError = sError & "A Waist circumference must be given for a D-Style flap" & NL
    End If

   'Return error message
    FN_ValidateExtensionData = sError


End Function
Function FN_BiArcCurve(xyStart As XY, aStart As Double, xyEnd As XY, aEnd As Double, Profile As BiArc) As Integer
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
    FN_BiArcCurve = False
    
   'Check for silly data and return
    If aStart = aEnd Then GoTo Error_Close
    If xyStart.Y = xyEnd.Y And xyStart.X = xyEnd.X Then GoTo Error_Close

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
        Phi1 = Abs((2 * PI) - Theta1)
    End If
    If Theta2 < PI Then
        Phi2 = Theta2
    Else
        Phi2 = Abs((2 * PI) - Theta2)
    End If

'Print #fError, "Phi1="; Phi1
'Print #fError, "Phi2="; Phi2


   'Calculate R1 and R2
    S1 = Abs(Sin(Theta1))
    C1 = (-Sin(Theta1) * Cos(Theta1)) / S1
'Print #fError, "S1="; S1; "C1="; C1

    S2 = Abs(Sin(Theta2))
    C2 = (Sin(Theta2) * Cos(Theta2)) / S2
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
        nTmp = (P * Sqr(nTmp)) / B
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
        A = 2 * Sin(Theta1)
        B = (Sin(Theta1) ^ 2) - (Cos(Theta1) ^ 2) + 1
        R1 = (P * A) / (2 * B)
        R2 = R1
'Print #fError, "A="; A; "B="; A;
'Print #fError, "R1="; R1
    End If

   'The radi R1 and R2 must be greater than the specified
   'minimum Rmin
    Rmin = nLength / 3
    If R1 < Rmin Then
'Print #fError, "R1 < Rmin"
        R1 = Rmin
        d = ((P ^ 2) - (2 * P * R1 * A) + (2 * (R1 ^ 2) * B)) / ((2 * R1 * B) - (2 * P * c))
'Print #fError, "D="; d
        R2 = R1 - d
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
    End If
    If R2 < Rmin Then
'Print #fError, "R2 < Rmin"
        R2 = Rmin
        R1 = ((P * Rmin * c) - ((P ^ 2) / 2)) / ((B * Rmin) + (P * c) - (P * A))
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
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
'Print #fError, "uvR1="; uvR1.X, uvR1.Y
'Print #fError, "R1 angle"; FN_CalcAngle(uvOrigin, uvR1)
    PR_MakeXY uvR2, P - (R2 * S2), R2 * C2
'Print #fError, "uvR2="; uvR2.X, uvR2.Y
'Print #fError, "R2 angle"; FN_CalcAngle(uvOrigin, uvR2)
    PR_MakeXY uvOrigin, 0, 0
   'Translate to XY co-ordinates
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR1) + aAxis, R1, Profile.xyR1
    PR_CalcPolar xyStart, FN_CalcAngle(uvOrigin, uvR2) + aAxis, FN_CalcLength(uvOrigin, uvR2), Profile.xyR2
'Print #fError, "R1="; R1
'Print #fError, "R2="; R2
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
'Close #fError
    FN_BiArcCurve = True

    Exit Function
    
Error_Close:
'Print #fError, "Error and close"
'Close #fError

End Function

Static Sub PR_CalcFingerBlendingProfile(xyTop As XY, xyBottom As XY, xyFinger As XY, Digit As Finger, ReturnProfile As Curve, xyThumbPalm As XY)
   'Procedure to return a smooth inflecting curve between two points
   '
   'The assumptions about the required curve are :-
   '    1. The Top Point is above the Bottom Point
   '    2. The curve leaves the Top point at 270
   '    3. The curve enters the bottom point at 270
   '    4. The curve is aproximated by two arcs
   '       that are tangential at the mid point
   '       between xyTop and xyBottom
   '    5. The curve inflects at the point in 4 above
   '    6. The radius of the arcs are the same
   '
   'The returned profile has the following features :-
   'If there is no missing finger then
   '    1. Decreases in Y from xyTop to xyBottom
   '    2. Has 7 vertex
   '    3. Vertex = 1 = xyTop
   '    4. Vertex = 7 = xyBottom
   '    5. Where xyTop.X = xyBottom.X returns a
   '       two point profile.
   'If there is a missing finger then
   '    1. Decreases in Y from xyTop to xyBottom
   '    2. Has 6 vertex
   '    3. Vertex = 1 = xyFinger
   '    4. Vertex = 6 = xyBottom
   '    4. Vertex = 2 to 5 create a fillet
   '       that can be manually edited
   '
   'Notes :-
   '    This is a fairly restricted routine designed
   '    to blend the profile above the wrist into the
   '    glove with a smooth inflecting curve.
   '
   'Modifications :-
   '    The point given in xyThumbPalm is the current point
   '    of the global variable xyThumbPalm(1) this is modified
   '    to the intersection of the blended curve at the height
   '    given by xyThumbPalm.y
   '
   '
    Dim nLength#, aAngle#, rAngle#, aInc#, nRadius#
    Dim xyCenter As XY, xyMidPoint As XY, xyPt1 As XY, xyPt2 As XY
    Dim xyConstruct As XY, xyInt As XY
    Dim iDirection%
    Dim ii%, MirrorResult%

    nLength = FN_CalcLength(xyBottom, xyTop)
    aAngle = FN_CalcAngle(xyBottom, xyTop)
     
   'If the angle is less than 90 Degrees then we just carry on.
   'However if the angle is greater, then we mirror this angle in the
   'y axis and calculate the points as befor, then we mirror the result
   '
    MirrorResult = False
    xyConstruct = xyTop
    If aAngle > 90 Then
        MirrorResult = True
        aAngle = 90 - (aAngle - 90)
        xyConstruct.X = xyBottom.X - (xyTop.X - xyBottom.X)
    End If

   'Degenerate to a straight line if xyTop.X = xyBottom.X
   'Then exit the sub routine
    If Abs(aAngle) = 90 Or aAngle = 270 Then
        ReturnProfile.n = 2
        ReturnProfile.X(1) = xyTop.X
        ReturnProfile.Y(1) = xyTop.Y
        ReturnProfile.X(2) = xyBottom.X
        ReturnProfile.Y(2) = xyBottom.Y
        GoTo MissingFingers:
    Else
        ReturnProfile.n = 7
        ReturnProfile.X(1) = xyConstruct.X
        ReturnProfile.Y(1) = xyConstruct.Y
        ReturnProfile.X(7) = xyBottom.X
        ReturnProfile.Y(7) = xyBottom.Y
    End If

   'Get midpoint
    PR_CalcPolar xyBottom, aAngle, nLength / 2, xyMidPoint
    ReturnProfile.X(4) = xyMidPoint.X
    ReturnProfile.Y(4) = xyMidPoint.Y
    
   'Calculate radius
    rAngle = aAngle * (PI / 180) 'Convert angle to Radians
    nRadius = (nLength / 4) / Cos(rAngle)
   
   'First arc points
    PR_MakeXY xyCenter, xyConstruct.X - nRadius, xyConstruct.Y
    aAngle = FN_CalcAngle(xyCenter, xyMidPoint)
    aInc = (360 - aAngle) / 3

   'Find Intersection
    If xyThumbPalm.Y > xyMidPoint.Y And xyThumbPalm.Y < xyTop.Y Then
        PR_MakeXY xyPt1, xyCenter.X, xyThumbPalm.Y
        PR_MakeXY xyPt2, xyCenter.X + 10, xyThumbPalm.Y
        If FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
            xyThumbPalm = xyInt
            If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
        End If
    End If

    For ii = 3 To 2 Step -1
        aAngle = aAngle + aInc
        PR_CalcPolar xyCenter, aAngle, nRadius, xyPt1
        ReturnProfile.X(ii) = xyPt1.X
        ReturnProfile.Y(ii) = xyPt1.Y
    Next ii
    
   'Second arc points
    PR_MakeXY xyCenter, xyBottom.X + nRadius, xyBottom.Y
    aAngle = FN_CalcAngle(xyCenter, xyMidPoint)

   'Find Intersection
    If xyThumbPalm.Y <= xyMidPoint.Y And xyThumbPalm.Y > xyBottom.Y Then
        PR_MakeXY xyPt1, xyCenter.X, xyThumbPalm.Y
        PR_MakeXY xyPt2, xyCenter.X - 10, xyThumbPalm.Y
        If FN_CirLinInt(xyPt1, xyPt2, xyCenter, nRadius, xyInt) Then
            xyThumbPalm = xyInt
            If MirrorResult Then xyThumbPalm.X = xyBottom.X - (xyThumbPalm.X - xyBottom.X)
        End If
    End If

    For ii = 5 To 6
        aAngle = aAngle + aInc
        PR_CalcPolar xyCenter, aAngle, nRadius, xyPt1
        ReturnProfile.X(ii) = xyPt1.X
        ReturnProfile.Y(ii) = xyPt1.Y
    Next ii

   'Mirroring the result simplifies the code above, as we only have
   'to code for the case where angle < 90
   'N.B. Mirroring in the Y axis along the line X = xyBottom.X
    If MirrorResult Then
        For ii = 1 To ReturnProfile.n
            ReturnProfile.X(ii) = xyBottom.X - (ReturnProfile.X(ii) - xyBottom.X)
        Next ii
    End If

MissingFingers:

    Dim nX#, nY#
    
    If Digit.Len = 0 And (Digit.Tip = 0 Or Digit.Tip = 10) Then
    
        nY = xyFinger.Y - ReturnProfile.Y(2)
        nX = xyFinger.X - ReturnProfile.X(1)
        iDirection = Sgn(xyFinger.X - ReturnProfile.X(1))
        nX = Abs(nX)
        nRadius = min(nX, nY) - 0.05
        
       'Straight line
        ReturnProfile.n = 6
        ReturnProfile.X(1) = xyFinger.X
        ReturnProfile.Y(1) = xyFinger.Y
        PR_MakeXY xyCenter, xyTop.X + (iDirection * nRadius), xyFinger.Y - nRadius
        aAngle = 90
        aInc = (90 / 3) * iDirection
        For ii = 2 To 5
            PR_CalcPolar xyCenter, aAngle, nRadius, xyPt1
            ReturnProfile.X(ii) = xyPt1.X
            ReturnProfile.Y(ii) = xyPt1.Y
            aAngle = aAngle + aInc
        Next ii
        ReturnProfile.X(6) = xyBottom.X
        ReturnProfile.Y(6) = xyBottom.Y

    End If

End Sub


Sub PR_CalculateArmTapeReductions()
'MsgBox "PR_CalculateArmTapeReductions"
   'Procedure to calculate the reductions
   'at each arm tape
    Dim ii%

   'Don't Calculate if normal glove
    If MainForm!optExtendTo(0).Value = True Then Exit Sub

    If MainForm!cboFabric.Text = "" Then
            MsgBox "Fabric not given! ", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
    Else
        If UCase$(Mid$(MainForm!cboFabric.Text, 1, 3)) <> "POW" Then
            MsgBox "Fabric chosen is not Powernet!", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
        End If
        If Val(Mid$(MainForm!cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid$(MainForm!cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
            MsgBox "Modulus of fabric chosen is not available on Gram / Tension reduction chart!", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
        End If

    End If

   
   'Update from dialogue
    PR_GetDlgAboveWrist
'PR_PrintDlgAboveWrist
    If g_iPressure < 0 Then
        MsgBox "Can't Calculate, missing Pressure", 48, "Manual Glove - Calculate ARM Button"
        Exit Sub
    End If

    If g_iNumTapesWristToEOS <= 1 Then
        MsgBox "Can't Calculate, Only one tape between given Wrist and EOS", 48, "Manual Glove - Calculate ARM Button"
        Exit Sub
    End If
   
   'Check that there are no holes in the data
    For ii = g_iWristPointer To g_iEOSPointer
        If g_nCir(ii) = 0 Then
            MsgBox "Can't Calculate, missing Tape Circumferences between given Wrist and EOS", 48, "Manual Glove - Calculate ARM Button"
            Exit Sub
        End If
    Next ii


   'Set the MMs based on the selected pressure
    PR_SetMMs (MainForm!cboPressure.Text)

   'Set the modulus based on the fabric
    PR_SetModulusIndex (MainForm!cboFabric.Text)

   'Calculate grams etc
    For ii = 1 To NOFF_ARMTAPES
        If ii >= g_iWristPointer And ii <= g_iEOSPointer Then
            
            g_iGms(ii) = round(g_nCir(ii) * g_iMMs(ii))
            g_iRed(ii) = FN_GetReduction(g_iGms(ii))

           'Adjust at wrist
            If ii = g_iWristPointer Then
               'The wrist must lie between 10 and 14
               'From the given reduction we back calculate the grams and mms
                If g_iRed(ii) > 14 Then g_iRed(ii) = 14
                g_iGms(ii) = FN_GetGrams(g_iRed(ii))
                g_iMMs(ii) = round(g_iGms(ii) / g_nCir(ii))
            End If

           'Adjust at EOS
            If ii = g_iEOSPointer Then
               'Always use a 10 reduction for plain EOS
               'For Flaps always use a 12
               'From the given reduction we back calculate the grams and mms
                If g_EOSType = ARM_PLAIN Then
                    g_iRed(ii) = 10
                Else
                    g_iRed(ii) = 12
                End If
                g_iGms(ii) = FN_GetGrams(g_iRed(ii))
                g_iMMs(ii) = round(g_iGms(ii) / g_nCir(ii))
            End If

           'Display Results
            PR_GrdInchesDisplay ii - 1, g_nCir(ii)
            PR_GramRedDisplay ii - 1, g_iGms(ii), g_iRed(ii)
            MainForm!mms(ii + 7) = g_iMMs(ii)
           
        Else
           'Blank the displayed values
            PR_GrdInchesDisplay ii - 1, 0
            PR_GramRedDisplay ii - 1, 0, 0
            MainForm!mms(ii + 7).Text = ""
        End If
    Next ii

    g_CalculatedExtension = True
End Sub

Sub PR_CalculateExtension(xyWristUlnar As XY, xyWristRadial As XY, nInsert As Double)

   'Procedure to calculate the POINTS
   'used to draw the extension of the glove above the
   'wrist
   'The Procedure also supplies the data that is to be
   'printed at each tape
   'N.B.
   'The wrist points will be modified if the wrist has been
   'calculated
     
    Dim nMidWristX As Double, nValue As Double, nSpacing As Double
    Dim ii%, nFiguredValue As Double, iStartTape%, iLastTape%
    Dim iVertex%
    
    nMidWristX = xyWristUlnar.X + (Abs(xyWristRadial.X - xyWristUlnar.X) / 2)

    If g_iNumTapesWristToEOS = 0 Or g_iNumTapesWristToEOS = 1 Then
       'Extension ends at wrist
        UlnarProfile.n = 1
        UlnarProfile.X(1) = xyWristUlnar.X
        UlnarProfile.Y(1) = xyWristUlnar.Y - 0.75

        RadialProfile.n = 1
        RadialProfile.X(1) = xyWristRadial.X
        RadialProfile.Y(1) = xyWristRadial.Y - 0.75

        Exit Sub

    End If

    If g_ExtendTo = GLOVE_NORMAL Then
       'Glove only extends two or three tapes past wrist
       'Use the charts to get the reductions
       '
       'Although given as part of the Extension
       'The normal glove uses data from the Hand Part of the dialogue
       '
       '
        UlnarProfile.n = g_iNumTapesWristToEOS
        UlnarProfile.X(1) = xyWristUlnar.X
        UlnarProfile.Y(1) = xyWristUlnar.Y
    
        RadialProfile.n = g_iNumTapesWristToEOS
        RadialProfile.X(1) = xyWristRadial.X
        RadialProfile.Y(1) = xyWristRadial.Y

       'God knows why this is but it works, and I don't give a !"£$%^&*(
       'any more
        iStartTape = 2
        iLastTape = g_iNumTapesWristToEOS + 1

        For ii = iStartTape To iLastTape
            
           'NB we add 1 to ii as the charts have wrist at ii=1, palm at ii=2
           'and 1-1/2 tape at ii=3.
           'The palm being out of order is a pain but it made it easier to
           'enter the chart data (swings and roundabouts, sorry!)
           '
           'NB use of 1/2 scale
            nFiguredValue = FN_FigureTape(ii + 1, g_nCir(ii))

            Select Case g_iInsertStyle
             Case 0
                If g_OnFold Then
                    nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
                Else
                    nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
                End If
             Case 1
                nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
             Case 2, 3
                nValue = (nFiguredValue + (2 * EIGHTH)) / 2
            End Select

            If g_OnFold Then
                UlnarProfile.X(ii) = UlnarProfile.X(1)
                RadialProfile.X(ii) = UlnarProfile.X(1) + nValue
            Else
                UlnarProfile.X(ii) = nMidWristX - (nValue / 2)
                RadialProfile.X(ii) = nMidWristX + (nValue / 2)
            End If
           
           'Standard spacing
            nSpacing = 1.375
           
           'We are working from the top (wrist) down
            UlnarProfile.Y(ii) = UlnarProfile.Y(ii - 1) - nSpacing
            RadialProfile.Y(ii) = RadialProfile.Y(ii - 1) - nSpacing

        Next ii

    End If


    If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_AXILLA Then
       'As we have recalculated the wrist we need to start at the
       'wrist
        UlnarProfile.n = g_iNumTapesWristToEOS
        UlnarProfile.Y(1) = xyWristUlnar.Y
        
        RadialProfile.n = g_iNumTapesWristToEOS
        RadialProfile.Y(1) = xyWristRadial.Y
        
        iStartTape = g_iWristPointer
        iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)

       'Note: As we can allow the wrist to be any tape we need a
       'vertex counter
        iVertex = 1
        
        For ii = iStartTape To iLastTape
           'NB use of 1/2 scale
            nFiguredValue = (g_nCir(ii) * ((100 - g_iRed(ii)) / 100))
            Select Case g_iInsertStyle
             Case 0
                If g_OnFold Then
                    nValue = ((nFiguredValue + EIGHTH) - nInsert) / 2
                Else
                    nValue = ((nFiguredValue + (3 * EIGHTH)) - (2 * nInsert)) / 2
                End If
             Case 1
                nValue = ((nFiguredValue + (3 * EIGHTH)) - nInsert) / 2
             Case 2, 3
                nValue = (nFiguredValue + (2 * EIGHTH)) / 2
            End Select

            If g_OnFold Then
                UlnarProfile.X(iVertex) = UlnarProfile.X(1)
                RadialProfile.X(iVertex) = UlnarProfile.X(1) + nValue
            Else
                UlnarProfile.X(iVertex) = nMidWristX - (nValue / 2)
                RadialProfile.X(iVertex) = nMidWristX + (nValue / 2)
            End If
        

           'Setup the notes for the vertex
           'We do this here as we have all the data and it simplifies
           'the drawing side
            TapeNote(iVertex).sTapeText = LTrim$(Mid$(g_sTapeText, ((ii + 1) * 3) + 1, 3))
            TapeNote(iVertex).iTapePos = ii
            TapeNote(iVertex).nCir = g_nCir(ii)
            TapeNote(iVertex).iGms = g_iGms(ii)
            TapeNote(iVertex).iRed = g_iRed(ii)
            TapeNote(iVertex).iMMs = g_iMMs(ii)


           'Standard spacing
            nSpacing = 1.375
           
           'Account all for pleats
           'Wrist (as we have started at the wrist we use, iStartTape + 1
           'and iStartTape + 2 in this case)
            If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(1))
            If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(2))
'XXXXXXXXXX
'XXXX Careful now! this won't work if only 3 or less tapes given
'XXXXXXXXXX
           'Axilla
            If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(3))
            If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(4))
    
           'We are working from the top (wrist) down
           'iVertex = 1 is set before the For Loop (Values from wrist)
            If iVertex <> 1 Then
                UlnarProfile.Y(iVertex) = UlnarProfile.Y(iVertex - 1) - nSpacing
                RadialProfile.Y(iVertex) = RadialProfile.Y(iVertex - 1) - nSpacing
                If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
            End If

           'Increment vertex count
            iVertex = iVertex + 1
            
        Next ii

       'Reset given wrist points  xyWristUlnar and xyWristRadial
        PR_MakeXY xyWristUlnar, UlnarProfile.X(1), UlnarProfile.Y(1)
        PR_MakeXY xyWristRadial, RadialProfile.X(1), RadialProfile.Y(1)
    End If

End Sub

Static Sub PR_CalcWristBlendingProfile(xyTSt As XY, xyTEnd As XY, xyBSt As XY, xyBEnd As XY, ReturnProfile As Curve, xyThumbPalm As XY)
   'Procedure to return a smooth inflecting curve between two points
   '
   'N.B. Parameters are given in the order of decreasing Y
   '
    Dim nL1#, nL2#, nL3#, aA1#, aA2#, aA3#, rAngle#, aInc#, nR1#, nR2#
    Dim xyCenter As XY, xyMidPoint As XY, xyTopSt As XY, xyTopEnd As XY, xyBotSt As XY, xyBotEnd As XY
    Dim xyR1 As XY, xyR2 As XY, nThirdOfL2#, nA#, aAngle#
    Dim xyPt1 As XY, xyPt2 As XY, xyInt As XY
    Dim ii%, MirrorResult%, Intersection#, Direction%, TopIsArc%, BottomIsArc%
    Dim xyTmp(1 To 10) As XY, nTol#

   'Do this as we can't use ByVal
    xyTopSt = xyTSt
    xyTopEnd = xyTEnd
    xyBotSt = xyBSt
    xyBotEnd = xyBEnd
     
   'If the angle is less than 90 Degrees then we just carry on.
   'However if the angle is greater, then we mirror this angle in the
   'y axis and calculate the points as befor, then we mirror the result
   '
    MirrorResult = False
    aA2 = FN_CalcAngle(xyBSt, xyTEnd)
    aA3 = FN_CalcAngle(xyBEnd, xyBSt)
    Direction = 1
    If (aA2 > 90) Or ((aA2 = 90) And (aA3 < 90)) Then
        MirrorResult = True
        Direction = -1
        aA2 = 90 - (aA2 - 90)
        xyTopSt.X = xyBEnd.X - (xyTopSt.X - xyBEnd.X)
        xyTopEnd.X = xyBEnd.X - (xyTopEnd.X - xyBEnd.X)
        xyBotSt.X = xyBEnd.X - (xyBotSt.X - xyBEnd.X)
    End If
    aA1 = FN_CalcAngle(xyTopEnd, xyTopSt)
    aA3 = FN_CalcAngle(xyBotEnd, xyBotSt)

   'Degenerate to a straight line
   'Then exit the sub routine
    If aA1 = aA2 And aA2 = aA3 And (aA1 = 90 Or aA1 = 270) Then
        ReturnProfile.n = 2
        ReturnProfile.X(1) = xyTSt.X
        ReturnProfile.Y(1) = xyTSt.Y
        ReturnProfile.X(2) = xyBEnd.X
        ReturnProfile.Y(2) = xyBEnd.Y
        Exit Sub
    End If

    nL1 = FN_CalcLength(xyTEnd, xyTSt)
    nL2 = FN_CalcLength(xyBSt, xyTEnd)
    nL3 = FN_CalcLength(xyBEnd, xyBSt)

   'Get Included angles & radius & Centers of Arcs
    nThirdOfL2 = nL2 / 3
   
   'Top Arc
    If aA1 <> aA2 Then
       'Calculate the points for the ARC for this section
        TopIsArc = True
        nA = FN_CalcLength(xyTopSt, xyBotSt)
        nA = ((nL1 ^ 2 + nL2 ^ 2) - nA ^ 2) / (2 * nL1 * nL2)
        rAngle = Arccos(nA) / 2
        nR1 = nThirdOfL2 * Tan(rAngle)
        PR_CalcPolar xyTopEnd, 90, nThirdOfL2, xyTmp(2)
        PR_CalcPolar xyTopEnd, FN_CalcAngle(xyTopEnd, xyBotSt), nThirdOfL2, xyTmp(5)
        PR_CalcPolar xyTmp(2), 180, nR1, xyR1
    
       'Top Arc Points
        aAngle = FN_CalcAngle(xyR1, xyTmp(5))
        aInc = (360 - aAngle) / 3
        For ii = 4 To 3 Step -1
            aAngle = aAngle + aInc
            PR_CalcPolar xyR1, aAngle, nR1, xyTmp(ii)
        Next ii
    Else
       'Calculate the points for the STRAIGHT LINE for this section
       'Note: we use the same noff points for consistancy
       'with respect to the editor and subsequent points
        TopIsArc = False
        PR_CalcPolar xyTopEnd, 90, nThirdOfL2, xyTmp(2)
        PR_CalcPolar xyTopEnd, 90, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(3)
        PR_CalcPolar xyTopEnd, 270, ((nThirdOfL2 * 2) / 3) / 2, xyTmp(4)
        PR_CalcPolar xyTopEnd, 270, nThirdOfL2, xyTmp(5)
    End If

    If aA2 <> aA3 Then
       'Bottom arc
        BottomIsArc = True
        nA = FN_CalcLength(xyTopEnd, xyBotEnd)
        nA = ((nL2 ^ 2 + nL3 ^ 2) - nA ^ 2) / (2 * nL3 * nL2)
        rAngle = Arccos(nA) / 2
        nR2 = nThirdOfL2 * Tan(rAngle)
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6)
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9)
       
       'Establish
        aAngle = FN_CalcAngle(xyBotEnd, xyTopEnd)
        If aAngle < FN_CalcAngle(xyBotEnd, xyBotSt) Then
            aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) - 90
        Else
            aAngle = FN_CalcAngle(xyBotSt, xyTopEnd) + 90
        End If
        PR_CalcPolar xyTmp(6), aAngle, nR2, xyR2
       
       'Check that the gap between the wrist point and the arc is less than
       'or equal to 0.0625 ie. 1/16"
       'If not then make it so.
        nTol = 0.0625
        nA = FN_CalcLength(xyBotSt, xyR2)
        If (nA - nR2) > nTol Then
            nR2 = nR2 - ((nA - nR2) - nTol)
            PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6)
            PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9)
           'aAngle from above
            PR_CalcPolar xyTmp(6), aAngle, nR2, xyR2
        End If
        
       'Bottom arc points
        aAngle = FN_CalcAngle(xyR2, xyTmp(6))
        aInc = (FN_CalcAngle(xyR2, xyTmp(9)) - aAngle) / 3
        For ii = 7 To 8
            aAngle = aAngle + aInc
            PR_CalcPolar xyR2, aAngle, nR2, xyTmp(ii)
        Next ii
   
    Else
        BottomIsArc = False
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), nThirdOfL2, xyTmp(6)
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyTopEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(7)
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), ((nThirdOfL2 * 2) / 3) / 2, xyTmp(8)
        PR_CalcPolar xyBotSt, FN_CalcAngle(xyBotSt, xyBotEnd), nThirdOfL2, xyTmp(9)
    End If

   'Find Intersection
    Intersection = False
    If xyThumbPalm.Y >= xyTmp(2).Y Then
        nA = xyThumbPalm.Y - xyTmp(2).Y
        If nA = 0 Then
                xyThumbPalm = xyTmp(2)
        Else
            PR_MakeXY xyThumbPalm, xyTmp(2).X + (nA / Tan(aA1 * (PI / 180))), xyThumbPalm.Y
        End If
        Intersection = True

    ElseIf xyThumbPalm.Y < xyTmp(2).Y And xyThumbPalm.Y > xyTmp(5).Y Then
        If TopIsArc Then
            PR_MakeXY xyPt1, xyR1.X, xyThumbPalm.Y
            PR_MakeXY xyPt2, xyR1.X + 10, xyThumbPalm.Y
            If FN_CirLinInt(xyPt1, xyPt2, xyR1, nR1, xyInt) Then
                xyThumbPalm = xyInt
                Intersection = True
            End If
        Else
            nA = xyThumbPalm.Y - xyTmp(5).Y
            If nA = 0 Then
                 xyThumbPalm = xyTmp(5)
            Else
                PR_MakeXY xyThumbPalm, xyTmp(5).X + (nA / Tan(aA2 * (PI / 180))), xyThumbPalm.Y
            End If
            Intersection = True
        End If

    ElseIf xyThumbPalm.Y <= xyTmp(5).Y And xyThumbPalm.Y >= xyTmp(6).Y Then
        nA = xyThumbPalm.Y - xyTmp(6).Y
        If nA = 0 Then
             xyThumbPalm = xyTmp(6)
        Else
            PR_MakeXY xyThumbPalm, xyTmp(6).X + (nA / Tan(aA2 * (PI / 180))), xyThumbPalm.Y
        End If
        Intersection = True

    ElseIf xyThumbPalm.Y < xyTmp(6).Y And xyThumbPalm.Y > xyTmp(9).Y Then
        If BottomIsArc Then
            PR_MakeXY xyPt1, xyTmp(6).X + nR2, xyThumbPalm.Y
            PR_MakeXY xyPt2, xyTmp(6).X - nR2, xyThumbPalm.Y
            If FN_CirLinInt(xyPt1, xyPt2, xyR2, nR2, xyInt) Then
                xyThumbPalm = xyInt
                Intersection = True
            End If
        Else
            nA = xyThumbPalm.Y - xyTmp(9).Y
            If nA = 0 Then
                 xyThumbPalm = xyTmp(9)
            Else
                PR_MakeXY xyThumbPalm, xyTmp(9).X + (nA / Tan(aA2 * (PI / 180))), xyThumbPalm.Y
            End If
            Intersection = True
        End If
    End If
    If Intersection And MirrorResult Then xyThumbPalm.X = xyBotEnd.X - (xyThumbPalm.X - xyBotEnd.X)

    ReturnProfile.n = 10
    ReturnProfile.X(1) = xyTopSt.X
    ReturnProfile.Y(1) = xyTopSt.Y
    ReturnProfile.X(10) = xyBotEnd.X
    ReturnProfile.Y(10) = xyBotEnd.Y

    For ii = 2 To 9
        ReturnProfile.X(ii) = xyTmp(ii).X
        ReturnProfile.Y(ii) = xyTmp(ii).Y
    Next ii
    

   'Mirroring the result simplifies the code above, as we only have
   'to code for the case where angle < 90
   'N.B. Mirroring in the Y axis along the line X = xyBottom.X
    If MirrorResult Then
        For ii = 1 To ReturnProfile.n
            ReturnProfile.X(ii) = xyBotEnd.X - (ReturnProfile.X(ii) - xyBotEnd.X)
        Next ii
    End If

End Sub

Sub PR_DisableFigureArm()
    
    Dim ii%

    For ii = 3 To 11
        MainForm!lblArm(ii).Enabled = False
    Next ii

   'Pleats
    g_nPleats(1) = Val(MainForm!txtWristPleat1.Text)
    g_nPleats(2) = Val(MainForm!txtWristPleat2.Text)
    g_nPleats(3) = Val(MainForm!txtShoulderPleat2.Text)
    g_nPleats(4) = Val(MainForm!txtShoulderPleat1.Text)

    For ii = 0 To 3
        MainForm!lblPleat(ii).Enabled = False
    Next ii
    

    MainForm!txtShoulderPleat1.Enabled = False
    MainForm!txtShoulderPleat2.Enabled = False
  '  MainForm!txtShoulderPleat1 = ""
  '  MainForm!txtShoulderPleat2 = ""
    MainForm!txtWristPleat1.Enabled = False
    MainForm!txtWristPleat2.Enabled = False
  '  MainForm!txtWristPleat1 = ""
  '  MainForm!txtWristPleat2 = ""

    MainForm!frmCalculate.Enabled = False
    MainForm!cboPressure.Enabled = False
    MainForm!cmdCalculate.Enabled = False

End Sub

Sub PR_DisableGloveToAxilla()

    MainForm!frmGloveToAxilla.Enabled = False
    MainForm!optProximalTape(0).Value = False
    MainForm!optProximalTape(1).Value = False
    MainForm!optProximalTape(0).Enabled = False
    MainForm!optProximalTape(1).Enabled = False
    PR_ProximalTape_Click (0)

End Sub

Static Sub PR_DisplayTextInches(ctlText As Control, ctlCaption As Control)
    Dim nLen As Double
    nLen = FN_InchesValue(ctlText)
    If nLen <> -1 Then ctlCaption.Caption = fnInchestoText(nLen)
End Sub

Static Sub PR_DrawShoulderFlaps(xyS As XY, xyE As XY)
    'Draws shoulder flaps
    'Extracted from ARMDIA.FRM
    'A load of £%&**!&*^$><@~, but it works (just about!)
    'The only reason to keep it, is that it is production proven
    'with the arm.
    '
    'We calculate the curves and point as they would be
    'attached to an arm curve.
    'We then rotate and mirror the points to fit the
    'drawing of a Glove to axilla
    '
Dim ShoulderFlap As Curve, RaglanFlap As Curve
Dim LastTapeValue, xyPt1 As XY, xyPt2 As XY
Dim xyTopFlap1 As XY, xyTopFlap2 As XY, xyTopFlap3 As XY, xyTopFlap4 As XY, xyTopFlap5 As XY
Dim xyTopFlap6 As XY, xyTopFlap7 As XY, xyTopFlap8 As XY, xyTopFlap9 As XY, xyTopFlap10 As XY
Dim xyBotFlap1 As XY, xyBotFlap2 As XY, xyBotFlap3 As XY, xyBotFlap4 As XY, xyBotFlap5 As XY
Dim xyFlap1 As XY, xyFlap2 As XY, xyFlap3 As XY, xyFlap4 As XY, xyFlap5 As XY
Dim Beta, Omega, Phi, Delta, Theta, BottomRadius, Alpha, TopArcIncrement
Dim x1, y1, h1, adj, opp, hyp, LittleBit, Change, TopRadius, FlapMarkerAngle, xyFlapText1  As XY
Dim xyMid As XY, xyCentre As XY, FlapLength, FlapMarkerLength, xyFlapMarker As XY, xyFlapText2 As XY, xyStrapText As XY
Dim ArcLength, CircleCircum, TopArcLength, BottomArcLength, TopArcRadius, BottomArcRadius, TopArcAngle, BottomArcAngle
Dim TempMarker, xfabby, xfablen, xfabdist
Dim xyRaglan1 As XY, xyRaglan2 As XY, xyRaglan3 As XY, xyRaglan4 As XY, xyRaglan5 As XY, xyRaglan6 As XY, xyRaglan7 As XY
Dim Rag1, Rag2, Rag3, Rag4, RagAng1, RagAng2, RagAng3, RaglanBottom, RaglanTip, TopRagLanAngle, InnerRaglanTip
Dim nNotchOffset As Double, nTemplateRadius#, nTemplateAngle#, nNotchToTangent#
Dim ii%, PhiExtra#, RagAng4#, RagAng5#, aMarker#
Dim Strap$
Dim aAngle#, nLength#, xyMidPoint As XY, xyStrt As XY, xyTmp As XY


    PR_CalcMidPoint xyS, xyE, xyMidPoint
    nLength = FN_CalcLength(xyS, xyE)
    aAngle = -90
    BottomArcAngle = 47.5
   'All this crap because you can't use ByVal on user defined types (Bastards!)
    If g_sSide = "Right" Then
        xyStrt = xyE
        aMarker = 135
    Else
        xyStrt = xyS
        aMarker = 45
    End If


   ' Check for Custom Flap Length
    If Val(MainForm!txtCustFlapLength) > 0 Then
        LastTapeValue = fnDisplayToInches(MainForm!txtCustFlapLength)
    Else
       'Standard flap length
        LastTapeValue = g_nCir(g_iEOSPointer)
        LastTapeValue = (LastTapeValue / 3.14) * 0.92
    End If

    
    ' Calculate Bottom part of curve with 5 Points
    If nLength >= 5.5 Then
        BottomRadius = LastTapeValue - 0.625
        RaglanBottom = 3.5
        TopRagLanAngle = 17
        RaglanTip = 1.9375
        InnerRaglanTip = 2.173
        nNotchOffset = 0.875   '14/16 ths Raglan Template s287
    ElseIf (nLength < 5.5) And (nLength > 3) Then
        BottomRadius = 2.8
        RaglanBottom = 2.5
        TopRagLanAngle = 17
        RaglanTip = 1.05
        InnerRaglanTip = 1.22
        nNotchOffset = 0.75     '12/16 ths Raglan Template s289
    ElseIf (nLength <= 3) Then
        BottomRadius = 2
        RaglanBottom = 2.0625
        RaglanTip = 0.625
        TopRagLanAngle = 15
        InnerRaglanTip = 0.727
        nNotchOffset = 0.6875   '11/16 ths, Raglan Template s290
    End If

    nTemplateRadius = 3.5
    nTemplateAngle = 40
    nNotchToTangent = 2.125

    'Calc length of bottom Arc
    BottomArcRadius = BottomRadius
    CircleCircum = PI * (2 * (BottomArcRadius))
    BottomArcLength = (BottomArcAngle / 360) * CircleCircum
    
    ShoulderFlap.X(1) = xyStrt.X + LastTapeValue
    ShoulderFlap.Y(1) = 0

    For ii = 1 To 4
        Theta = (ii * 11.875 * PI) / 180
        adj = Cos(Theta) * BottomRadius
        opp = Sin(Theta) * BottomRadius
        LittleBit = BottomRadius - adj
        ShoulderFlap.X(ii + 1) = xyStrt.X + LastTapeValue - LittleBit
        ShoulderFlap.Y(ii + 1) = 0 + opp
    Next ii
    PR_MakeXY xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.Y(5)
    
    'Calculate 6 angles and points for top part of curve
    PR_MakeXY xyTopFlap1, xyStrt.X, xyStrt.Y + nLength
    ShoulderFlap.X(10) = xyTopFlap1.X
    ShoulderFlap.Y(10) = xyTopFlap1.Y

    'midpoint
    h1 = FN_CalcLength(xyTopFlap1, xyBotFlap5)
    h1 = h1 / 2
    Omega = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Omega = 180 - Omega
    Omega = (Omega * PI) / 180
    opp = Abs(h1 * Sin(Omega))
    adj = Abs(h1 * Cos(Omega))
    PR_MakeXY xyMid, xyBotFlap5.X - adj, xyBotFlap5.Y + opp
   
    'Centre of Top Arc
    Delta = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Delta = Delta - 90
    Delta = (Delta * PI) / 180
    hyp = 4 * h1
    opp = Abs(hyp * Sin(Delta))
    adj = Abs(hyp * Cos(Delta))
    PR_MakeXY xyCentre, xyMid.X + adj, xyMid.Y + opp
    TopRadius = FN_CalcLength(xyCentre, xyTopFlap1)
    TopArcRadius = TopRadius
    Phi = FN_CalcAngle(xyCentre, xyTopFlap1)
    Phi = Phi - 180
    Alpha = FN_CalcAngle(xyCentre, xyBotFlap5)
    Alpha = Alpha - 180
    TopArcAngle = Alpha - Phi
    TopArcIncrement = (Alpha - Phi) / 5

   'Calc Flap Length for position of marker
    CircleCircum = PI * (2 * (TopArcRadius))
    TopArcLength = (TopArcAngle / 360) * CircleCircum
    FlapLength = TopArcLength + BottomArcLength
    FlapMarkerLength = FlapLength / 2.7
    If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
    If FlapMarkerLength < TopArcLength Then
        FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
        FlapMarkerAngle = FlapMarkerAngle + Phi
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(TopRadius * Sin(FlapMarkerAngle))
        adj = Abs(TopRadius * Cos(FlapMarkerAngle))
        PR_MakeXY xyFlapMarker, xyCentre.X - adj, xyCentre.Y - opp
    ElseIf FlapMarkerLength > TopArcLength Then
        TempMarker = FlapMarkerLength - TopArcLength
        TempMarker = BottomArcLength - TempMarker
        FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(BottomArcRadius * Sin(FlapMarkerAngle))
        adj = Abs(BottomArcRadius * Cos(FlapMarkerAngle))
        LittleBit = BottomRadius - adj
        PR_MakeXY xyFlapMarker, xyStrt.X + LastTapeValue - LittleBit, 0 + opp
    ElseIf FlapMarkerLength = TopArcLength Then
        PR_MakeXY xyFlapMarker, xyBotFlap5.X, xyBotFlap5.Y
    End If
    PR_Setlayer "Template" & g_sSide

    For ii = 1 To 4
        PhiExtra = Phi + (ii * TopArcIncrement)
        PhiExtra = (PhiExtra * PI) / 180
        opp = Abs(TopRadius * Sin(PhiExtra))
        adj = Abs(TopRadius * Cos(PhiExtra))
        ShoulderFlap.X(10 - ii) = xyCentre.X - adj
        ShoulderFlap.Y(10 - ii) = xyCentre.Y - opp
    Next ii

    
    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
       'Copy standard shoulder from profile down
       'only go one vertex past the FlapMarker
        RaglanFlap.n = 0
        For ii = 10 To 1 Step -1
            RaglanFlap.n = RaglanFlap.n + 1
            RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
            RaglanFlap.Y(RaglanFlap.n) = ShoulderFlap.Y(ii)
            If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
        Next ii
        PR_RotateCurve xyStrt, aAngle, RaglanFlap
        If g_sSide = "Left" Then PR_MirrorCurveInYaxis xyMidPoint.X, 0, RaglanFlap
        PR_DrawFitted RaglanFlap
    Else
        ShoulderFlap.n = 10
        PR_RotateCurve xyStrt, aAngle, ShoulderFlap
        If g_sSide = "Left" Then PR_MirrorCurveInYaxis xyMidPoint.X, 0, ShoulderFlap
        PR_DrawFitted ShoulderFlap
    End If
 
    'Draw Raglan
    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
        Rag1 = Sqr((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
        Rag1 = BottomRadius - Rag1
        PR_MakeXY xyRaglan1, xyStrt.X + LastTapeValue - Rag1, nNotchOffset
        PR_MakeXY xyRaglan2, xyRaglan1.X - nNotchToTangent, 0
        RagAng1 = FN_CalcAngle(xyRaglan2, xyRaglan1)
        Rag2 = FN_CalcLength(xyRaglan2, xyRaglan1)
        Rag2 = Rag2 / 2
        RagAng1 = (RagAng1 * PI) / 180
        opp = Abs(Rag2 * Sin(RagAng1))
        adj = Abs(Rag2 * Cos(RagAng1))
        PR_MakeXY xyRaglan3, xyRaglan2.X + adj, xyRaglan2.Y + opp
        
        RagAng1 = FN_CalcAngle(xyRaglan3, xyRaglan1)
        RagAng1 = 90 - RagAng1
        RagAng1 = (RagAng1 * PI) / 180
        Rag3 = FN_CalcLength(xyRaglan3, xyRaglan1)
        Rag4 = Sqr((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
        opp = Abs(Rag4 * Sin(RagAng1))
        adj = Abs(Rag4 * Cos(RagAng1))
        PR_MakeXY xyRaglan4, xyRaglan3.X - adj, xyRaglan3.Y + opp
        
        RagAng2 = (nTemplateAngle * PI) / 180
        opp = Abs(RaglanBottom * Sin(RagAng2))
        adj = Abs(RaglanBottom * Cos(RagAng2))
        PR_MakeXY xyRaglan5, xyRaglan1.X + adj, xyRaglan1.Y + opp
        RagAng3 = FN_CalcAngle(xyRaglan5, xyFlapMarker)
        
        If RagAng3 > 180 Then
            RagAng3 = RagAng3 - 180
            RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
        ElseIf RagAng3 <= 180 Then
            RagAng3 = RagAng3 - 90
            RagAng4 = 90 - RagAng3
            RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
        End If
        
        RagAng5 = RagAng4 - 13
        RagAng4 = (RagAng4 * PI) / 180
        opp = Abs(RaglanTip * Sin(RagAng4))
        adj = Abs(RaglanTip * Cos(RagAng4))
        PR_MakeXY xyRaglan6, xyRaglan5.X - adj, xyRaglan5.Y + opp
        
        'Draw Front line of tip
        RagAng5 = (RagAng5 * PI) / 180
        opp = Abs(InnerRaglanTip * Sin(RagAng5))
        adj = Abs(InnerRaglanTip * Cos(RagAng5))
        PR_MakeXY xyRaglan7, xyRaglan5.X - adj, xyRaglan5.Y + opp
        

        PR_RotatePoint xyStrt, aAngle, xyRaglan1
        PR_RotatePoint xyStrt, aAngle, xyRaglan2
        PR_RotatePoint xyStrt, aAngle, xyRaglan4
        PR_RotatePoint xyStrt, aAngle, xyRaglan5
        PR_RotatePoint xyStrt, aAngle, xyRaglan6
        PR_RotatePoint xyStrt, aAngle, xyRaglan7

        xyTmp = xyFlapMarker
        PR_RotatePoint xyStrt, aAngle, xyTmp

        If g_sSide = "Left" Then
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan1
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan2
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan4
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan5
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan6
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyRaglan7
            PR_MirrorPointInYaxis xyMidPoint.X, 0, xyTmp
        End If

        PR_DrawArc xyRaglan4, xyRaglan1, xyRaglan2
        PR_DrawLine xyRaglan1, xyRaglan5
        PR_DrawLine xyRaglan5, xyRaglan6
        PR_DrawLine xyRaglan6, xyTmp
        PR_Setlayer "Notes"
        PR_DrawLine xyRaglan5, xyRaglan7

    End If
    
    PR_Setlayer "Notes"
    PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
    PR_MakeXY xyStrapText, xyMidPoint.X, xyMidPoint.Y - (2 * EIGHTH)
    
    If Val(MainForm!txtStrapLength) <> 0 Then
        Strap = "STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtStrapLength)))) & "\" & QQ
    Else
        Strap = "STRAP = 24\" & QQ
    End If
    If Val(MainForm!txtFrontStrapLength) > 0 Then
      'Display FRONT strap length if given
       Strap = "BACK  STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtStrapLength)))) & "\" & QQ
       Strap = Strap & "\nFRONT STRAP = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtFrontStrapLength)))) & "\" & QQ
    End If
    If Val(MainForm!txtWaistCir) > 0 Then
      'Display Waist Circumference for D-Style flaps only
       Strap = Strap & "\nWAIST = " & Trim(fnInchestoText(fnDisplayToInches(Val(MainForm!txtWaistCir)))) & "\" & QQ
    End If

    PR_DrawText Strap, xyStrapText, 0.1
    
    xfabby = MainForm!cboFlaps.Text
    xfablen = Len(xfabby)
    xfabdist = (xfablen * 0.075) + 0.1
    PR_MakeXY xyFlapText1, xyFlapMarker.X - (xfabdist * 0.75), xyFlapMarker.Y - 0.5
    PR_RotatePoint xyStrt, aAngle, xyFlapText1
    If g_sSide = "Left" Then PR_MirrorPointInYaxis xyMidPoint.X, 0, xyFlapText1
    PR_DrawText (MainForm!cboFlaps.Text), xyFlapText1, 0.1

   'Flap Marker
    PR_RotatePoint xyStrt, aAngle, xyFlapMarker
    If g_sSide = "Left" Then PR_MirrorPointInYaxis xyMidPoint.X, 0, xyFlapMarker
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "closed arrow"; QC; "xyStart.x +"; Str$(xyFlapMarker.X); CC; "xyStart.y +"; Str$(xyFlapMarker.Y); ",0.25,0.1,"; aMarker; ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"
    
   'draw the fold line
    PR_Setlayer "Template" & g_sSide
    
    PR_MakeXY xyPt1, xyStrt.X, 0
    PR_MakeXY xyPt2, xyStrt.X + LastTapeValue, 0
    
    PR_RotatePoint xyStrt, aAngle, xyPt1
    PR_RotatePoint xyStrt, aAngle, xyPt2
    If g_sSide = "Left" Then
        PR_MirrorPointInYaxis xyMidPoint.X, 0, xyPt1
        PR_MirrorPointInYaxis xyMidPoint.X, 0, xyPt2
    End If

    If Left$(MainForm!cboFlaps.Text, 6) = "Raglan" Then
        PR_DrawLine xyPt1, xyRaglan2
    Else
        PR_DrawLine xyPt1, xyPt2                        'Fold Line
    End If

End Sub

Static Sub PR_EnableFigureArm()

   'Converse of PR_DisableFigureArm
    
    Dim ii%
    
    For ii = 3 To 11
        MainForm!lblArm(ii).Enabled = True
    Next ii

    For ii = 0 To 3
        MainForm!lblPleat(ii).Enabled = True
    Next ii
    
    MainForm!txtShoulderPleat1.Enabled = True
    MainForm!txtShoulderPleat2.Enabled = True
    If g_nPleats(1) > 0 Then MainForm!txtWristPleat1 = g_nPleats(1)
    If g_nPleats(2) > 0 Then MainForm!txtWristPleat2 = g_nPleats(2)
    MainForm!txtWristPleat1.Enabled = True
    MainForm!txtWristPleat2.Enabled = True
    If g_nPleats(3) > 0 Then MainForm!txtShoulderPleat2 = g_nPleats(3)
    If g_nPleats(4) > 0 Then MainForm!txtShoulderPleat1 = g_nPleats(4)

    MainForm!frmCalculate.Enabled = True
    MainForm!cboPressure.Enabled = True
    MainForm!cboPressure.ListIndex = g_iPressure
    MainForm!cmdCalculate.Enabled = True

End Sub

Static Sub PR_ExtendTo_Click(Index As Integer)
    Dim ii%, iFirst%, iLast%, iStart%, iElbow%
    
    iLast = 23
    iFirst = 8
    iElbow = 16 'Elbow w.r.t. txtExtCir()
    iStart = 8

    Select Case Index
        Case 0       'Normal Glove
           'Disable tapes above those required
            For ii = iStart To iLast
                MainForm!txtExtCir(ii).Enabled = False
                MainForm!lblTape(ii).Enabled = False
                PR_GrdInchesDisplay ii - 8, 0
            Next ii

           'Disable all MMs fields etc
            For ii = iFirst To iLast
                MainForm!mms(ii).Enabled = False
                MainForm!mms(ii) = ""
                PR_GramRedDisplay ii - 8, 0, 0
            Next ii

            PR_DisableFigureArm
            PR_DisableGloveToAxilla
           
           'set disable and disable
            MainForm!cboDistalTape.ListIndex = -1
            MainForm!cboDistalTape.Enabled = False
            MainForm!cboProximalTape.ListIndex = -1
            MainForm!cboProximalTape.Enabled = False
        
        Case 1     'Glove to Elbow
           'Enable MMs fields etc
            For ii = iFirst To iLast
                If ii <= iElbow Then
                    MainForm!txtExtCir(ii).Enabled = True
                    MainForm!mms(ii).Enabled = True
                    MainForm!lblTape(ii).Enabled = True
                Else
                    MainForm!txtExtCir(ii).Enabled = False
                    MainForm!lblTape(ii).Enabled = False
                    MainForm!mms(ii).Enabled = True
                    MainForm!mms(ii) = ""
                    PR_GramRedDisplay ii - 8, 0, 0
                    PR_GrdInchesDisplay ii - 8, 0
                End If
            Next ii
           
           'Enable other bits
            PR_EnableFigureArm
            PR_DisableGloveToAxilla
           
           'Set Wrist and EOS pointers
            If g_iNumTapesWristToEOS > 0 Then
               'Set wrist to given tape
                If g_iWristPointer > g_iFirstTape Then
                    MainForm!cboDistalTape.ListIndex = g_iWristPointer
                Else
                   'set to first tape
                    MainForm!cboDistalTape.ListIndex = 0
                End If
    
               'Set EOS to given tape or to Elbow if it extends
               'past the elbow
               'NB. The order of tapes is reversed in this list
               'starting at 19-1/2 and finishing at 0
                If g_iEOSPointer < ELBOW_TAPE Then
                    If g_iEOSPointer >= g_iLastTape Then
                        MainForm!cboProximalTape.ListIndex = 0
                    Else
                        MainForm!cboProximalTape.ListIndex = 17 - g_iEOSPointer
                    End If
                ElseIf g_iEOSPointer > ELBOW_TAPE Or g_iLastTape > ELBOW_TAPE Then
                    MainForm!cboProximalTape.ListIndex = 17 - ELBOW_TAPE
                Else
                   'set to last tape
                    MainForm!cboProximalTape.ListIndex = 0
                End If
            Else
                    MainForm!cboDistalTape.ListIndex = 0
                    MainForm!cboProximalTape.ListIndex = 0
            End If

            MainForm!cboDistalTape.Enabled = True
            MainForm!cboProximalTape.Enabled = True

        Case 2     'Glove to Axilla
            For ii = iFirst To iLast
                MainForm!mms(ii).Enabled = True
                MainForm!txtExtCir(ii).Enabled = True
                MainForm!lblTape(ii).Enabled = True
'                MainForm!mms(ii) = ""
'                PR_GramRedDisplay ii - 8, 0, 0
            Next ii

           'Enable other bits
            PR_EnableFigureArm

            MainForm!frmGloveToAxilla.Enabled = True
            If g_EOSType = ARM_FLAP Then MainForm!optProximalTape(1).Value = True Else MainForm!optProximalTape(0).Value = True
            
            
            MainForm!optProximalTape(0).Enabled = True
            MainForm!optProximalTape(1).Enabled = True
            
            If g_iNumTapesWristToEOS > 0 Then
               'Set wrist to given tape
                If g_iWristPointer > g_iFirstTape Then
                    MainForm!cboDistalTape.ListIndex = g_iWristPointer
                Else
                   'set to first tape
                    MainForm!cboDistalTape.ListIndex = 0
                End If
    
               'Set EOS to given tape
               'NB. The order of tapes is reversed in this list
               'starting at 19-1/2 and finishing at 0
                If g_iEOSPointer < g_iLastTape Then
                    MainForm!cboProximalTape.ListIndex = 17 - g_iEOSPointer
                Else
                   'set to first tape
                    MainForm!cboProximalTape.ListIndex = 0
                End If
            Else
                For ii = iFirst To iLast
                    MainForm!mms(ii) = ""
                    PR_GramRedDisplay ii - 8, 0, 0
                Next ii
                MainForm!cboDistalTape.ListIndex = 0
                MainForm!cboProximalTape.ListIndex = 0
            End If

            MainForm!cboDistalTape.Enabled = True
            MainForm!cboProximalTape.Enabled = True

    End Select

End Sub
Sub PR_PrintDlgAboveWrist()
   'Debug routine only
    Dim sMessage As String
    Dim ii As Integer
   'Prints
   '    g_nCir(1, NOFF_ARMTAPES) As Double
   '    g_iFirstTape            As Integer
   '    g_iLastTape             As Integer
   '    g_iWristPointer         As Integer
   '    g_iEOSPointer           As Integer
   '    g_iNumTotalTapes        As Integer
   '    g_iNumTapesWristToEOS   As Integer
   '    g_EOSType               As Integer
   '    g_OnFold                As Integer
   '    g_ExtendTo              As Integer
   '    g_nPleats(1 To 4)       As Double
   '    g_iPressure             As Integer
    For ii = 1 To NOFF_ARMTAPES
        sMessage = sMessage & "g_nCir(" & Str(ii) & ")=" & Str(g_nCir(ii)) & Chr$(13)
    Next ii
    MsgBox sMessage
    sMessage = ""
    sMessage = "g_iFirstTape=" & Str(g_iFirstTape) & Chr$(13)
    sMessage = sMessage & "g_iLastTape=" & Str(g_iLastTape) & Chr$(13)
    sMessage = sMessage & "g_iWristPointer=" & Str(g_iWristPointer) & Chr$(13)
    MsgBox sMessage

End Sub
Sub PR_GetDlgAboveWrist()
'MsgBox "PR_GetDlgAboveWrist"
   'General procedure to read the data given in the
   'dialogue controls and copy to module level variables.
   'Saves having to do this more than once.
   '
   'Updates:-
   '
   '    g_nCir(1, NOFF_ARMTAPES) As Double
   '    g_iFirstTape            As Integer
   '    g_iLastTape             As Integer
   '    g_iWristPointer         As Integer
   '    g_iEOSPointer           As Integer
   '    g_iNumTotalTapes        As Integer
   '    g_iNumTapesWristToEOS   As Integer
   '    g_EOSType               As Integer
   '    g_OnFold                As Integer
   '    g_ExtendTo              As Integer
   '    g_nPleats(1 To 4)       As Double
   '    g_iPressure             As Integer
   '
   'NOTE:
   '    We ignore the fact that there may be missing tapes.
   '
   '    g_iWristPointer and g_iEOSPointer are used to indicate the
   '    start and finish tapes in the arrays and not in the
   '    txtExtCir() array

    Dim nCir#
    Dim ii%

    If MainForm!SSTab1.Tab <> 1 Then MainForm!SSTab1.Tab = 1

    g_iFirstTape = -1
    g_iLastTape = -1
    g_iWristPointer = -1
    g_iEOSPointer = -1
    g_iNumTotalTapes = 0
    g_iNumTapesWristToEOS = 0

   'Glove type
    If MainForm!optExtendTo(0).Value = True Then
        g_ExtendTo = GLOVE_NORMAL
    ElseIf MainForm!optExtendTo(1).Value = True Then
        g_ExtendTo = GLOVE_ELBOW
    Else
        g_ExtendTo = GLOVE_AXILLA
    End If

    

    If g_ExtendTo = GLOVE_NORMAL Then
       'Get values from Hand Data TAB
       'We will need to update this w.r.t CAD Glove later
       'Assummes no holes in data
        If g_nCir(2) <= 0 Then
            g_iNumTotalTapes = 1
            g_iNumTapesWristToEOS = 1
        ElseIf g_nCir(3) > 0 Then
            g_iNumTotalTapes = 3
            g_iNumTapesWristToEOS = 3
        Else
            g_iNumTotalTapes = 2
            g_iNumTapesWristToEOS = 2
        End If
'        g_iWristPointer = 1
        g_iWristPointer = 0
        g_iEOSPointer = 1000 'Stupid value to fool stupid code
                             'in PR_ExtendTo_Click
    Else
       'Get First and last tape (if any)
        For ii = 8 To 23
            nCir = FN_InchesValue(MainForm!txtExtCir(ii))
            If nCir > 0 And g_iFirstTape = -1 Then g_iFirstTape = ii - 7
            g_nCir(ii - 7) = nCir
        Next ii
        
        For ii = 23 To 8 Step -1
            If Val(MainForm!txtExtCir(ii)) > 0 Then
                g_iLastTape = ii - 7
                Exit For
            End If
        Next ii
    
       'Total number of tapes
        If (g_iLastTape = g_iFirstTape) Then
            If g_iLastTape <> -1 Then g_iNumTotalTapes = 1
        Else
            g_iNumTotalTapes = (g_iLastTape - g_iFirstTape) + 1
        End If
       'Get values from "Above Wrist" TAB
        Dim iIndex%
       'Wrist tape
       'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
        iIndex = MainForm!cboDistalTape.ListIndex
        If iIndex = 0 Or iIndex = -1 Then
            g_iWristPointer = g_iFirstTape
        Else
            g_iWristPointer = iIndex
        End If
    
       'EOS tape
       'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
        iIndex = MainForm!cboProximalTape.ListIndex
        If iIndex = 0 Or iIndex = -1 Then
            g_iEOSPointer = g_iLastTape
        Else
           'NB. The order of tapes is reversed in this list
           'starting at 19-1/2 and finishing at 0
            g_iEOSPointer = 17 - iIndex
        End If
    
        
       'Number of tapes between
        If (g_iWristPointer = g_iEOSPointer) Then
           'Just in case there are no tapes
           'And "first" and "last" are used for Wrist and EOS
            If g_iLastTape <> -1 Then g_iNumTapesWristToEOS = 1
        Else
            g_iNumTapesWristToEOS = (g_iEOSPointer - g_iWristPointer) + 1
        End If
    
       'Pleats
        If MainForm!txtWristPleat1.Enabled = True Then g_nPleats(1) = Val(MainForm!txtWristPleat1.Text) Else g_nPleats(1) = 0
        If MainForm!txtWristPleat2.Enabled = True Then g_nPleats(2) = Val(MainForm!txtWristPleat2.Text) Else g_nPleats(2) = 0
        If MainForm!txtShoulderPleat2.Enabled = True Then g_nPleats(3) = Val(MainForm!txtShoulderPleat2.Text) Else g_nPleats(3) = 0
        If MainForm!txtShoulderPleat1.Enabled = True Then g_nPleats(4) = Val(MainForm!txtShoulderPleat1.Text) Else g_nPleats(4) = 0

        g_iPressure = MainForm!cboPressure.ListIndex
    
    End If


   'End of support type
    g_EOSType = ARM_PLAIN
    If g_ExtendTo = GLOVE_AXILLA And MainForm!optProximalTape(1).Value Then
        g_EOSType = ARM_FLAP
        g_sFlapType = MainForm!cboFlaps.Text
        g_iFlapType = MainForm!cboFlaps.ListIndex
        If MainForm!txtStrapLength.Text <> "" Then g_nStrapLength = FN_InchesValue(MainForm!txtStrapLength)
        If MainForm!txtFrontStrapLength.Text <> "" Then g_nFrontStrapLength = FN_InchesValue(MainForm!txtFrontStrapLength)
        If MainForm!txtCustFlapLength.Text <> "" Then g_nCustFlapLength = FN_InchesValue(MainForm!txtCustFlapLength)
        If MainForm!txtWaistCir.Text <> "" Then g_nWaistCir = FN_InchesValue(MainForm!txtWaistCir)
    End If
    
   'On or off fold
    If MainForm!optFold(1).Value = True Then g_OnFold = True Else g_OnFold = False


End Sub

Sub PR_GetExtensionDDE_Data()
   'Procedure to use the data given in the Form for the
   'glove extension (if any) along the arm
   '
    Dim nAge As Integer, Flap%
    Dim ii%, jj%, nValue#, nn%
    Dim iMMs%, iGms%, iRed%
    Dim sDiag As String
    
    MainForm!SSTab1.Tab = 1

   'Defaults
    MainForm!optFold(1).Value = True         'On fold
    MainForm!optExtendTo(0).Value = True     'Normal Glove
    MainForm!cboDistalTape.ListIndex = 0     '1st Tape
    MainForm!cboProximalTape.ListIndex = 0   'Last Tape
    Flap = False

    
   'Set default pressure w.r.t diagnosis
    nAge = Val(MainForm!txtAge.Text)

   ' if pressure empty set from diagnosis
    sDiag = UCase$(Mid$(MainForm!txtDiagnosis.Text, 1, 6))
   'Arterial Insufficiency
    If sDiag = "ARTERI" Then
        g_iPressure = 0
   'Assist fluid dynamics
    ElseIf sDiag = "ASSIST" Then
        g_iPressure = 1
   'Blood Clots
    ElseIf sDiag = "BLOOD " Then
        g_iPressure = 1
   'Burns
    ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then
        g_iPressure = 0
   'Cancer
    ElseIf sDiag = "CANCER" Then
        g_iPressure = 1
   'Cardial Vascular Arrest
    ElseIf sDiag = "CARDIA" Then
        g_iPressure = 0
   'Carpal Tunnel Syndrome
    ElseIf sDiag = "CARPAL" Then
        g_iPressure = 1
   'Cellulitis
    ElseIf sDiag = "CELLUL" Then
        g_iPressure = 1
   'Chronic Venous Insufficiency
    ElseIf sDiag = "CHRONI" Then
        g_iPressure = 1
   'Heart Condition
    ElseIf sDiag = "HEART " Then
        g_iPressure = 0
   'Hemangioma
    ElseIf sDiag = "HEMANG" Then
        g_iPressure = 0
   'Lymphedema, Lymphedema 1+, Lymphedema 2+
   'N.B. Lymphedema <=> Lymphedema 1+
    ElseIf sDiag = "LYMPHE" Then
        If InStr(MainForm!txtDiagnosis.Text, "2") > 0 Then
            If nAge <= 70 Then
                g_iPressure = 2
            Else
                g_iPressure = 1
            End If
        Else
            If nAge <= 70 Then
                g_iPressure = 1
            Else
                g_iPressure = 0
            End If
        End If
   'Night Wear
    ElseIf sDiag = "NIGHT " Then
        g_iPressure = 0
   'Post Fracture
    ElseIf sDiag = "POST F" Then
        g_iPressure = 1
   'Post Mastectomy
    ElseIf sDiag = "POST M" Then
        g_iPressure = 1
   'Postphlebetic Syndrome
    ElseIf sDiag = "POSTPH" Then
        g_iPressure = 1
   'Renal Disease (Kidney)
    ElseIf sDiag = "RENAL " Then
        g_iPressure = 1
   'Skin Graft
    ElseIf sDiag = "SKIN G" Then
        g_iPressure = 0
   'Stroke
    ElseIf sDiag = "STROKE" Then
        g_iPressure = 0
   'Tendonitis
    ElseIf sDiag = "TENDON" Then
        g_iPressure = 0
   'Thrombophlebitis
    ElseIf sDiag = "THROMB" Then
        g_iPressure = 1
   'Trauma
    ElseIf sDiag = "TRAUMA" Then
        g_iPressure = 1
   'Varicose Veins
    ElseIf sDiag = "VARICO" Then
        g_iPressure = 1
   'Request for 30 m/m, 40 m/m and 50 m/m
    ElseIf sDiag = "REQUES" Then
        If InStr(MainForm!txtDiagnosis.Text, "30") > 0 Then
            g_iPressure = 0
        ElseIf InStr(MainForm!txtDiagnosis.Text, "40") > 0 Then
            g_iPressure = 1
        ElseIf InStr(MainForm!txtDiagnosis.Text, "50") > 0 Then
            g_iPressure = 2
        End If
   'Vein Removal
    ElseIf sDiag = "VEIN R" Then
        g_iPressure = 1
    End If

  'Glove extensions
  'Code for extended glove
  'Glove Option Buttons
   Dim nLen#
   
   If MainForm!txtDataGlove.Text <> "" Then
     For ii = 0 To 5
        nValue = Val(Mid(MainForm!txtDataGlove.Text, (ii * 2) + 1, 2))
        If ii = 0 Then
           'Fold options
            MainForm!optFold(nValue).Value = True
        ElseIf ii = 1 And nValue >= 0 Then
           'Pressure w.r.t Figuring and MMs
            g_iPressure = nValue
        ElseIf ii = 2 Then
           'Wrist Tape
            MainForm!cboDistalTape.ListIndex = nValue
        ElseIf ii = 3 Then
           'Proximal Tape
            MainForm!cboProximalTape.ListIndex = nValue
        ElseIf ii = 4 Then
            MainForm!optExtendTo(nValue).Value = True
            If nValue = 2 Then Flap = True
        ElseIf ii = 5 Then
           'Only if glove to axilla is set do we
           'do anything
            If Flap And nValue = 1 Then
                'Flap, break up flap multiple field
                 MainForm!optProximalTape(1).Value = True
                 PR_ProximalTape_Click (1)
                 For jj = 0 To 4
                    nValue = Val(Mid(MainForm!txtFlap.Text, (jj * 3) + 1, 3))
                    If jj = 0 Then
                        MainForm!cboFlaps.ListIndex = nValue
                    ElseIf jj = 1 And nValue > 0 Then
                        g_nStrapLength = nValue / 10
                        MainForm!txtStrapLength = g_nStrapLength
                        PR_DisplayTextInches MainForm!txtStrapLength, MainForm!labStrap
                    ElseIf jj = 2 And nValue > 0 Then
                        g_nFrontStrapLength = nValue / 10
                        MainForm!txtFrontStrapLength = g_nFrontStrapLength
                        PR_DisplayTextInches MainForm!txtFrontStrapLength, MainForm!labFrontStrapLength
                    ElseIf jj = 3 And nValue > 0 Then
                        g_nCustFlapLength = nValue / 10
                        MainForm!txtCustFlapLength = g_nCustFlapLength
                        PR_DisplayTextInches MainForm!txtCustFlapLength, MainForm!labCustFlapLength
                    ElseIf jj = 4 And nValue > 0 Then
                        g_nWaistCir = nValue / 10
                        MainForm!txtWaistCir = g_nWaistCir
                        PR_DisplayTextInches MainForm!txtWaistCir, MainForm!labWaistCir
                    End If
                 Next jj
            Else
                'Disable flaps
                 PR_ProximalTape_Click (0)
            End If
        End If
    Next ii
   End If

  'Set value for pressure
   MainForm!cboPressure.ListIndex = g_iPressure

  'Glove to elbow and Glove to axilla
  'These can start at -3 (However to allow for possible
  'changes later we save up to -4-1/2 but ignore -4-1/2 for the
  'mean time)
   If MainForm!txtTapeLengthPt1.Text <> "" Then
    ii = 1
    For nn = 8 To 23
        nValue = Val(Mid(MainForm!txtTapeLengthPt1.Text, (ii * 3) + 1, 3))
        If nValue > 0 Then
            MainForm!txtExtCir(nn) = nValue / 10
            nLen = FN_InchesValue(MainForm!txtExtCir(nn))
            If nLen <> -1 Then PR_GrdInchesDisplay nn - 8, nLen
            g_nCir(ii) = nLen
        Else
            PR_GrdInchesDisplay nn - 8, 0
        End If
        iMMs = Val(Mid(MainForm!txtTapeMMs.Text, (ii * 3) + 1, 3))
        If iMMs > 0 Then
            g_CalculatedExtension = True
            iGms = Val(Mid(MainForm!txtGrams.Text, (ii * 3) + 1, 3))
            iRed = Val(Mid(MainForm!txtReduction.Text, (ii * 3) + 1, 3))
            MainForm!mms(nn) = iMMs
            PR_GramRedDisplay nn - 8, iGms, iRed
            g_iMMs(ii) = iMMs
            g_iRed(ii) = iRed
            g_iGms(ii) = iGms
        End If
        ii = ii + 1
    Next nn
   End If

  'Pleats
    For ii = 0 To 1
        nValue = Val(Mid(MainForm!txtWristPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtWristPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtWristPleat2 = nValue / 10
        End If
        nValue = Val(Mid(MainForm!txtShoulderPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtShoulderPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtShoulderPleat2 = nValue / 10
        End If
    Next ii
    

End Sub


Sub PR_GramRedDisplay(iIndex As Integer, Gram As Integer, Reduction As Integer)
    MainForm!grdDisplay.Row = iIndex
   
    MainForm!grdDisplay.Col = 0
    If Gram = 0 Then MainForm!grdDisplay.Text = "" Else MainForm!grdDisplay.Text = Str$(Gram)
    
    MainForm!grdDisplay.Col = 1
    If Reduction = 0 Then MainForm!grdDisplay.Text = "" Else MainForm!grdDisplay.Text = Str$(Reduction)

End Sub

Sub PR_GrdInchesDisplay(iIndex As Integer, nLen As Double)
    MainForm!grdInches.Row = iIndex
    MainForm!grdInches.Col = 0
    MainForm!grdInches.Text = fnInchestoText(nLen)
End Sub

Static Sub PR_LoadPowernetChart()
   'Procedure to load the reduction charts for the arms from
   'disk
   '
    Dim sFile As String, sModulus As String, sLine As String
    Dim fChart%, ii%, jj%, nn%

    sFile = g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    
   'Get ARM chart reductions
   'NB these are fixed format
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_MODULUS
        Input #fChart, sModulus, sLine
        
        For jj = 0 To 22
            g_iPowernet(ii, jj + 1) = Val(Mid(sLine, (jj * 4) + 1, 4))
        Next jj

        ii = ii + 1
          
    Wend
    
   'Close file
    Close fChart
    sLine = ""
    sModulus = ""

End Sub

Static Sub PR_LoadReductionCharts(sFabric As String)
   'Procedure to load the reduction charts from
   'disk
   'Two charts are loaded
   '    1. Length chart
   '    2. Circumferences based on the Fabric
   '
   
    Dim sFile As String, sLine As String
    Dim iModulus%, fChart%, ii%, jj%, nn%

    If sFabric = "" Then Exit Sub
    If UCase$(Mid$(sFabric, 1, 3)) <> "POW" Then
        MsgBox "Fabric chosen is not Powernet", 48, g_sDialogID
        Exit Sub
    End If

   'Establish chart to be loaded
   '    Fabric Format  Pow MMM-XX Comment
   '    if mm < 230 use 230 chart
   '    if MM > 280 use 280 chart
   '
    iModulus = Val(Mid$(sFabric, 5, 3))
    If iModulus < 230 Then iModulus = 230
    If iModulus > 280 Then iModulus = 280
    sFile = g_sPathJOBST & "\TEMPLTS\GLV_" & Trim$(Str$(iModulus)) & ".DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Fabric Chart, Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    
   'Get finger and thumb reducations
   'NB these are fixed format
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_FINGER_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB use of "." to repeat previous number
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then

            iFingerRed(1, ii) = Val(Mid(sLine, 5, 2))
            
            If Mid$(sLine, 8, 1) = "." Then
                iFingerRed(2, ii) = iFingerRed(1, ii)
            Else
                iFingerRed(2, ii) = Val(Mid(sLine, 8, 2))
            End If

            If Mid$(sLine, 11, 1) = "." Then
                iFingerRed(3, ii) = iFingerRed(2, ii)
            Else
                iFingerRed(3, ii) = Val(Mid(sLine, 11, 2))
            End If

            ii = ii + 1
        End If
    Wend
    
   'Wrist and Palm to 9 tape reductions
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_ARM_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB: use of "." to repeat previous number
       '    also translation from inches and eights to decimal inches)
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then
            nArmRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
            jj = 1
            For nn = 9 To 34 Step 4
                jj = jj + 1
                If Mid$(sLine, nn, 1) = "." Then
                    nArmRed(jj, ii) = nArmRed(jj - 1, ii)
                Else
                    nArmRed(jj, ii) = Val(Mid(sLine, nn, 2)) + (Val(Mid(sLine, nn + 2, 1)) * 0.125)
                End If
            Next nn
            ii = ii + 1
        End If
    Wend

   'Close file
    Close fChart


   'Length Reduction chart
    sFile = g_sPathJOBST & "\TEMPLTS\GLV_LEN.DAT"

    If Dir$(sFile) = "" Then
        MsgBox "Template file not found " & sFile, 48, g_sDialogID
        Exit Sub
    End If

   'Open file
    fChart = FreeFile
    Open sFile For Input As fChart
    ii = 1
    While Not EOF(fChart) And ii <= NOFF_LENGTH_RED
        Line Input #fChart, sLine
       'Ignore comments and blank lines
       'NB: use of "." to repeat previous number
       '    also translation from inches and eights to decimal inches)
        If Mid$(sLine, 1, 1) <> "#" And sLine <> "" Then
            
            nLengthRed(1, ii) = Val(Mid(sLine, 5, 2)) + (Val(Mid(sLine, 7, 1)) * 0.125)
            
            If Mid$(sLine, 9, 1) = "." Then
                nLengthRed(2, ii) = nLengthRed(1, ii)
            Else
                nLengthRed(2, ii) = Val(Mid(sLine, 9, 2)) + (Val(Mid(sLine, 11, 1)) * 0.125)
            End If

            ii = ii + 1
        End If
    Wend

   'Close file
    Close fChart

    sLine = ""

End Sub

Sub PR_ProximalTape_Click(Index As Integer)
   Dim ii%
   If Index = 0 Then
       'Disable flaps
        g_iFlapType = MainForm!cboFlaps.ListIndex
        MainForm!cboFlaps.Enabled = False
        MainForm!cboFlaps.ListIndex = -1
        For ii = 0 To 4
            MainForm!lblFlap(ii).Enabled = False
        Next ii
        MainForm!txtStrapLength.Enabled = False
        g_nStrapLength = Val(MainForm!txtStrapLength.Text)
        MainForm!txtStrapLength.Text = ""
        MainForm!labStrap.Caption = ""

        MainForm!txtFrontStrapLength.Enabled = False
        g_nFrontStrapLength = Val(MainForm!txtFrontStrapLength.Text)
        MainForm!txtFrontStrapLength.Text = ""
        MainForm!labFrontStrapLength.Caption = ""
        
        MainForm!txtCustFlapLength.Enabled = False
        g_nCustFlapLength = Val(MainForm!txtCustFlapLength.Text)
        MainForm!txtCustFlapLength.Text = ""
        MainForm!labCustFlapLength.Caption = ""
        
        MainForm!txtWaistCir.Enabled = False
        g_nWaistCir = Val(MainForm!txtWaistCir.Text)
        MainForm!txtWaistCir.Text = ""
        MainForm!labWaistCir.Caption = ""

   Else
       'Enable flaps
        MainForm!optFold(0) = True   'Ensure off the fold only
        MainForm!cboFlaps.Enabled = True
        MainForm!cboFlaps.ListIndex = g_iFlapType

        For ii = 0 To 3
            MainForm!lblFlap(ii).Enabled = True
        Next ii
        MainForm!txtStrapLength.Enabled = True
        If g_nStrapLength > 0 Then
            MainForm!txtStrapLength = g_nStrapLength
            PR_DisplayTextInches MainForm!txtStrapLength, MainForm!labStrap
        End If

        MainForm!txtFrontStrapLength.Enabled = True
        If g_nFrontStrapLength > 0 Then
            MainForm!txtFrontStrapLength = g_nFrontStrapLength
            PR_DisplayTextInches MainForm!txtFrontStrapLength, MainForm!labFrontStrapLength
        End If

        MainForm!txtCustFlapLength.Enabled = True
        If g_nCustFlapLength > 0 Then
            MainForm!txtCustFlapLength = g_nCustFlapLength
            PR_DisplayTextInches MainForm!txtCustFlapLength, MainForm!labCustFlapLength
        End If

        If InStr(1, MainForm!cboFlaps.Text, "D") > 0 Then
            MainForm!lblFlap(4).Enabled = True
            MainForm!txtWaistCir.Enabled = True
            If g_nWaistCir > 0 Then
                MainForm!txtWaistCir = g_nWaistCir
                PR_DisplayTextInches MainForm!txtWaistCir, MainForm!labWaistCir
            End If
        End If

   End If

End Sub

Sub PR_SetMMs(sPressure As String)
   'Set the mms based on the pressure and the wrist and EOS
   'tapes given
   'REF:    GOP 01-02/16, Section 1.4
   '
    Dim ii%

    For ii = g_iWristPointer To g_iEOSPointer
     Select Case sPressure
        Case "15"
            If ii = g_iWristPointer Then
                g_iMMs(ii) = 12
            ElseIf ii < ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 15
            ElseIf ii = ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 12
            ElseIf ii = ELBOW_TAPE Then
                g_iMMs(ii) = 8
            ElseIf ii = ELBOW_TAPE + 1 Then
                g_iMMs(ii) = 10
            ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                g_iMMs(ii) = 12
            ElseIf ii = g_iEOSPointer - 1 Then
                g_iMMs(ii) = 10
            ElseIf ii = g_iEOSPointer Then
                g_iMMs(ii) = 8
            End If
        Case "20"
            If ii = g_iWristPointer Then
                g_iMMs(ii) = 16
            ElseIf ii < ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 20
            ElseIf ii = ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 16
            ElseIf ii = ELBOW_TAPE Then
                g_iMMs(ii) = 10
            ElseIf ii = ELBOW_TAPE + 1 Then
                g_iMMs(ii) = 13
            ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                g_iMMs(ii) = 16
            ElseIf ii = g_iEOSPointer - 1 Then
                g_iMMs(ii) = 13
            ElseIf ii = g_iEOSPointer Then
                g_iMMs(ii) = 10
            End If
        Case "25"
            If ii = g_iWristPointer Then
                g_iMMs(ii) = 20
            ElseIf ii < ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 25
            ElseIf ii = ELBOW_TAPE - 1 Then
                g_iMMs(ii) = 20
            ElseIf ii = ELBOW_TAPE Then
                g_iMMs(ii) = 13
            ElseIf ii = ELBOW_TAPE + 1 Then
                g_iMMs(ii) = 17
            ElseIf (ii > ELBOW_TAPE + 1) And (ii <> g_iEOSPointer - 1) And (ii <> g_iEOSPointer) Then
                g_iMMs(ii) = 20
            ElseIf ii = g_iEOSPointer - 1 Then
                g_iMMs(ii) = 17
            ElseIf ii = g_iEOSPointer Then
                g_iMMs(ii) = 13
            End If
     End Select
    Next ii

End Sub

Sub PR_SetModulusIndex(sFabric As String)
   'Set the module level variable iModulus Index based on
   'the chosen fabric
   '
   ' Where sFabric = Pow 230-2B
   '
   '    g_iModulusIndex Start = LOW_MODULUS  = 160
   '    g_iModulusIndex End   = HIGH_MODULUS = 340
   '
   ' hence
   '    g_iModulusIndex = ((230 - 160) + 10) / 10
   '
   '
    Dim iMod%

    iMod = Val(Mid(sFabric, 5, 3))
    If iMod < LOW_MODULUS Then
        g_iModulusIndex = LOW_MODULUS
    ElseIf iMod > HIGH_MODULUS Then
        g_iModulusIndex = HIGH_MODULUS
    Else
        g_iModulusIndex = ((iMod - LOW_MODULUS) + 10) / 10
    End If
    
End Sub

Sub PR_UpdateDDE_Extension()
   'Procedure to update the fields used when data is transfered
   'from DRAFIX using DDE.
   'Although the transfer back to DRAFIX is not via DDE we use the same controls
   'simply to illustrate the method by which the data is packed into the
   'fields

   'This routine is similar to PR_UpdateDDE in MANGLOVE.FRM#
   'it is split so that we can acccess the Module level variables
    Dim iLen%, ii%
    Dim sLen$, sPacked$

    MainForm!SSTab1.Tab = 1


   'Pack Glove extensions
    sPacked = "   "     'Allow for possible extension to -4 tape
    For ii = 8 To 23
        iLen = Val(MainForm!txtExtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii
    MainForm!txtTapeLengthPt1 = sPacked

   'Pack Grams, MMs and Reductions
    Dim sGrams$, sTapeMMs$, sReduction$

    sGrams = "   "      'Allow for possible extension to -4 tape
    sTapeMMs = "   "    'Allow for possible extension to -4 tape
    sReduction = "   "  'Allow for possible extension to -4 tape
    
    For ii = 1 To 16
       'Grams
        If g_iGms(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iGms(ii)))
        Else
            sLen = String(3, " ")
        End If
        sGrams = sGrams & sLen

       'Tape pressures
        If g_iMMs(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iMMs(ii)))
        Else
            sLen = String(3, " ")
        End If
        sTapeMMs = sTapeMMs & sLen

       'Reductions
        If g_iRed(ii) <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(g_iRed(ii)))
        Else
            sLen = String(3, " ")
        End If
        sReduction$ = sReduction$ & sLen

    Next ii

    MainForm!txtReduction.Text = sReduction
    MainForm!txtTapeMMs.Text = sTapeMMs
    MainForm!txtGrams.Text = sGrams

   'Fold, variations etc etc.
    sPacked = ""

   'Fold
    If MainForm!optFold(0).Value = True Then sPacked = " 0" Else sPacked = " 1"
    
   'Pressure
    sLen = String(2, " ")
    RSet sLen = Trim$(Str$(MainForm!cboPressure.ListIndex))
    sPacked = sPacked & sLen

   'Wrist Tape
    RSet sLen = Trim$(Str$(MainForm!cboDistalTape.ListIndex))
    sPacked = sPacked & sLen

   'EOS Tape
    RSet sLen = Trim$(Str$(MainForm!cboProximalTape.ListIndex))
    sPacked = sPacked & sLen

   'Variations
    If MainForm!optExtendTo(0).Value = True Then
        sPacked = sPacked & " 0 0" 'Second Zero is w.r.t Flap
    ElseIf MainForm!optExtendTo(1).Value = True Then
        sPacked = sPacked & " 1 0" 'Second Zero is w.r.t Flap
    Else
        sPacked = sPacked & " 2"
       'Check for flaps
        If MainForm!optProximalTape(1).Value = True Then sPacked = sPacked & " 1" Else sPacked = sPacked & " 0"
    End If
    
   'Thumb Web Drop in 1/16ths for use only with web spacers
   'Ignored by every thing else
    RSet sLen = Trim$(Str$(g_iThumbWebDrop))
    sPacked = sPacked & sLen
    
    
    MainForm!txtDataGlove.Text = sPacked

   'Flap
    sPacked = ""
    If MainForm!optExtendTo(2).Value = True And MainForm!optProximalTape(1).Value = True Then
       'Flap Type
        sLen = String(3, " ")
        RSet sLen = Trim$(Str$(MainForm!cboFlaps.ListIndex))
        sPacked = sPacked & sLen
  
        RSet sLen = Trim$(Str$(Val(MainForm!txtStrapLength.Text) * 10))    'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtFrontStrapLength.Text) * 10))  'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtCustFlapLength.Text) * 10))   'Shift decimal place
        sPacked = sPacked & sLen

        RSet sLen = Trim$(Str$(Val(MainForm!txtWaistCir.Text) * 10))  'Shift decimal place
        sPacked = sPacked & sLen

    End If

    MainForm!txtFlap.Text = sPacked

   'Pleats
    sLen = String(3, " ")
    If Val(MainForm!txtWristPleat1.Text) > 0 Then RSet sLen = Trim$(Str$(Val(MainForm!txtWristPleat1.Text) * 10)) 'Shift decimal place
    sPacked = sLen

    sLen = String(3, " ")
    If Val(MainForm!txtWristPleat2.Text) Then RSet sLen = Trim$(Str$(Val(MainForm!txtWristPleat2.Text) * 10)) 'Shift decimal place
    MainForm!txtWristPleat.Text = sPacked & sLen
    
    sLen = String(3, " ")
    If Val(MainForm!txtShoulderPleat1.Text) > 0 Then RSet sLen = Trim$(Str$(Val(MainForm!txtShoulderPleat1.Text) * 10)) 'Shift decimal place
    sPacked = sLen

    sLen = String(3, " ")
    If Val(MainForm!txtShoulderPleat2.Text) Then RSet sLen = Trim$(Str$(Val(MainForm!txtShoulderPleat2.Text) * 10)) 'Shift decimal place
    MainForm!txtShoulderPleat.Text = sPacked & sLen

End Sub


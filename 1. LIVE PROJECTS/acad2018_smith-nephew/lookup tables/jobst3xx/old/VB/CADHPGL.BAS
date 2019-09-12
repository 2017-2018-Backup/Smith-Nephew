Attribute VB_Name = "CADHPGL"
'Project:   CADGLOVE.VBP
'File:      CADHPGL.BAS
'Purpose:   HPGL Interpreter and DRAFIX translator
'           module.
'
'Version:   3.01
'Date:      17.Jan.96
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'07 Dec 98  GG      Move the interaction with the CAD
'                   Glove external programme to the module
'                   Kevin
'Notes:-
'
    
    Private miFileHPGL As Integer
    Private Value(1 To 10) As Variant

   'The TextLabels are use to identify points
   '
    Type TextLabel
        Found As Integer
        sString As String      'If any
        X As Double
        y As Double
    End Type
    
    Dim g_Labels(1 To 30) As TextLabel

    
    Option Explicit


'
'Notes:-
'
'   This implementation of a stack as an array is lifted directly
'   from
'       Data Structures and Program Design
'       Robert L. Kruse, 1984
'       2nd Edition, Prentice Hall International Editions
'
'

    Const MAXSTACK = 100

    Type stack
        Top As Integer
        X(1 To MAXSTACK) As Double
        y(1 To MAXSTACK) As Double
    End Type


    Global S As stack


Sub ClearStack()
    S.Top = 0
End Sub

Sub DrawMarker()
    Dim xyStart As XY

   'Draw a Marker at the top of the current stack point

    Dim Xlast#, Ylast#

   'Return if empty
    If EmptyStack() Then Exit Sub

   'Get last point
    StackCurrent Xlast, Ylast
    PR_MakeXY xyStart, Xlast / ISCALE, Ylast / ISCALE
    PR_RevisePoint xyStart

    PR_DrawMarker xyStart

End Sub

Sub DrawText(Text)
    Dim xyStart As XY
    Dim Xstart#, Ystart#

    StackCurrent Xstart, Ystart
    PR_MakeXY xyStart, Xstart / ISCALE, Ystart / ISCALE
    PR_RevisePoint xyStart

    PR_DrawText Text, xyStart, 0.1

End Sub

Function EmptyStack() As Integer
    If S.Top = 0 Then EmptyStack = True Else EmptyStack = False
End Function



Function FullStack()
    If S.Top = MAXSTACK Then FullStack = True Else FullStack = False
End Function

Sub PopStack(X As Double, y As Double)
    If S.Top = 0 Then
        MsgBox "Stack is Empty", 16, "Stack Operations"
        End
    Else
        X = S.X(S.Top)
        y = S.y(S.Top)
        S.Top = S.Top - 1
    End If

End Sub

Static Sub PR_AddNotesEtc()
   'This is a catchall procedure to draw everything that is not
   'part of the glove outline.
   'It uses the "hooks" created in the Labels "LB" statements in the
   'HPGL code.  It's ugly but it works!
   '
    Dim ii%
    Dim sText$, sSymbol$
    Dim TmpX#, TmpY#
    Dim xyInsert As XY, xyStart As XY, xyEnd As XY, Reinforced As Curve
    Dim xyPt As XY
    Dim nMaxX#, nMinX#, nMaxY#, nMinY#, nInsert#, nXStep#, nYStep#, nSpacing#
    Dim sWorkOrder As String

   'Patient Details
    If g_Labels(6).Found = True Then
        g_nCurrTextAngle = 0

        PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
       
       'Main details in Green on layer notes
        PR_SetLayer "Notes"

        PR_MakeXY xyInsert, (g_Labels(6).X / ISCALE) - 1.25, (g_Labels(6).y / ISCALE) - 0.25
        PR_RevisePoint xyInsert
        If MainForm!txtWorkOrder.Text = "" Then
            sWorkOrder = "-"
        Else
            sWorkOrder = MainForm!txtWorkOrder.Text
        End If
        sText = MainForm!txtSide.Text & "\n" & MainForm!txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim$(Mid$(MainForm!cboFabric.Text, 4))
        PR_DrawText sText, xyInsert, 0.125


       'Other patient details in black on layer construct
        PR_SetLayer "Construct"

        PR_MakeXY xyInsert, xyInsert.X + 0.8, xyInsert.y
        sText = MainForm!txtFileNo.Text & "\n" & MainForm!txtDiagnosis.Text & "\n" & MainForm!txtAge.Text & "\n" & MainForm!txtSex.Text
        PR_DrawText sText, xyInsert, 0.125
    
    End If
    
    PR_SetLayer "Notes"
    PR_SetTextData 2, 16, -1, -1, -1
    
   'Thumb and palm pieces
    g_nCurrTextAngle = 90
    For ii = 7 To 8
      If g_Labels(ii).Found = True Then
        sText = MainForm!txtSide.Text & "\n" & MainForm!txtPatientName.Text & "\n" & MainForm!txtWorkOrder.Text
        PR_MakeXY xyInsert, g_Labels(ii).X / ISCALE, g_Labels(ii).y / ISCALE
        PR_RevisePoint xyInsert
        PR_DrawText sText, xyInsert, 0.125
      End If
    Next ii
    g_nCurrTextAngle = 0

   'Position Insert text Size w.r.t lowest web position
    nMaxX = 0
    nMinX = 1.23456789E+18
    nMaxY = 0
    nMinY = 1.23456789E+18
    For ii = 1 To 5
        If g_Labels(ii).Found = True Then
            nMaxX = max(nMaxX, g_Labels(ii).X)
            nMinX = min(nMinX, g_Labels(ii).X)
            nMaxY = max(nMaxY, g_Labels(ii).y)
            nMinY = min(nMinY, g_Labels(ii).y)
        End If
    Next ii
   
   'Get insert size
    sText = "Unknown "
    If g_Labels(9).Found = True Then
       'Use the given insert value
       'else use caculated value returned from the CAD-Glove DOS Programme
        If MainForm!cboInsert.ListIndex > 0 Then
            g_iInsertSize = MainForm!cboInsert.ListIndex
            nInsert = g_iInsertSize * 0.125
        Else
            nInsert = Val(Mid(g_Labels(9).sString, 3))
            If nInsert <> 0 Then g_iInsertSize = round(nInsert / 0.125)
        End If
        sText = Trim$(fnInchestoText(nInsert))
        If Mid$(sText, 1, 1) = "-" Then sText = Mid$(sText, 2) 'Strip leading "-" sign
        sText = sText & "\" & QQ
    End If
    sText = "INSERT " & sText
    g_nCurrTextAngle = 0
    PR_MakeXY xyInsert, (nMaxX / ISCALE) + nInsert, (nMinY + ((nMaxY - nMinY) / 2)) / ISCALE
    PR_RevisePoint xyInsert
    PR_DrawText sText, xyInsert, 0.125
    g_nCurrTextAngle = 0

   'Elastic
    If Val(MainForm!txtAge.Text) <= 10 And g_ExtendTo = GLOVE_NORMAL Then
        PR_MakeXY xyInsert, g_Labels(6).X / ISCALE, (nMinY + ((nMaxY - nMinY) / 2)) / ISCALE
        PR_RevisePoint xyInsert
        sText = "1/2\" & QQ & " Elastic"
        g_nCurrTextAngle = 0
        PR_DrawText sText, xyInsert, 0.125
        g_nCurrTextAngle = 0
    End If

   'Slanted inserts
    If MainForm!chkSlantedInserts.Value = 1 Then
        nInsert = nInsert - 0.125
        For ii = 2 To 4
            If g_Labels(ii).Found = True Then
                TmpX = g_Labels(ii).X / ISCALE
                TmpY = g_Labels(ii).y / ISCALE

               'Horizontal line
                PR_MakeXY xyStart, TmpX, TmpY
                PR_RevisePoint xyStart
                PR_MakeXY xyEnd, TmpX + nInsert, TmpY
                PR_RevisePoint xyEnd
                PR_DrawLine xyStart, xyEnd
                PR_AddDBValueToLast "ID", g_sID

               'Vertical line
                PR_MakeXY xyStart, TmpX + nInsert, TmpY + 0.125
                PR_RevisePoint xyStart
                PR_MakeXY xyEnd, TmpX + nInsert, TmpY - 0.125
                PR_RevisePoint xyEnd
                PR_DrawLine xyStart, xyEnd
                PR_AddDBValueToLast "ID", g_sID

            End If
        Next ii
    End If

   'Reinforced Palms
    If MainForm!chkPalm.Value = 1 Then
        For ii = 1 To 5
            If g_Labels(ii).Found = True Then
                Reinforced.n = Reinforced.n + 1
                PR_MakeXY xyPt, (g_Labels(ii).X / ISCALE) + 0.25, g_Labels(ii).y / ISCALE
                PR_RevisePoint xyPt
                Reinforced.X(Reinforced.n) = xyPt.X
                Reinforced.y(Reinforced.n) = xyPt.y
            End If
        Next ii

       'Draw reinforcing lines using Long Dash line type
        Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "Long Dash"; QQ; "));"
        
        PR_DrawFitted Reinforced
        PR_MakeXY xyStart, g_Labels(10).X / ISCALE, g_Labels(10).y / ISCALE
        PR_MakeXY xyEnd, g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE
        PR_RevisePoint xyStart
        PR_RevisePoint xyEnd
        PR_DrawLine xyStart, xyEnd
        
        Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    
       'Add REINFORCED text
        TmpX = (g_Labels(5).X + ((g_Labels(11).X - g_Labels(5).X) / 2)) / ISCALE
        TmpY = g_Labels(11).y / ISCALE

        PR_MakeXY xyInsert, TmpX, TmpY
        PR_RevisePoint xyInsert
        PR_DrawText "REINFORCED\nPALM", xyInsert, 0.125

    End If

   'Reinforced Dorsal
    If MainForm!chkDorsal.Value = 1 Then
        Dim nDorsalOffset#
        nDorsalOffset = 0
        Reinforced.n = 0
        If MainForm!chkSlantedInserts.Value = 1 Then nDorsalOffset = nInsert
        For ii = 1 To 5
            If g_Labels(ii).Found = True Then
                Reinforced.n = Reinforced.n + 1
                PR_MakeXY xyPt, (g_Labels(ii).X / ISCALE) + 0.25 + nDorsalOffset, g_Labels(ii).y / ISCALE
                PR_RevisePoint xyPt
                Reinforced.X(Reinforced.n) = xyPt.X
                Reinforced.y(Reinforced.n) = xyPt.y
            End If
        Next ii

       'Draw reinforcing lines using Long Dash line type
        Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "Long Dash"; QQ; "));"
        
        PR_DrawFitted Reinforced
        If MainForm!chkPalm.Value <> 1 Then
           'Draw wrist line only if not drawn above
            PR_MakeXY xyStart, g_Labels(10).X / ISCALE, g_Labels(10).y / ISCALE
            PR_MakeXY xyEnd, g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE
            PR_RevisePoint xyStart
            PR_RevisePoint xyEnd
            PR_DrawLine xyStart, xyEnd
        End If
        
        Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    
       'Add REINFORCED text
        TmpX = (g_Labels(5).X + ((g_Labels(11).X - g_Labels(5).X) / 2)) / ISCALE
        TmpY = g_Labels(11).y / ISCALE
        If MainForm!chkPalm.Value = 1 Then
           'Shift text so that it misses PALM Text
           If MainForm!txtSide.Text = "Right" Then
                TmpY = TmpY - 0.4
           Else
                TmpY = TmpY + 0.4
           End If
        End If
        PR_MakeXY xyInsert, TmpX, TmpY
        PR_RevisePoint xyInsert
        PR_DrawText "REINFORCED\nDORSAL", xyInsert, 0.125

    End If

   'Missing Fingers
   'Check for Fingers tip options
    ReDim Closed(0 To 4)
    Closed(0) = MainForm!optLittleTip(0).Value
    Closed(1) = MainForm!optRingTip(0).Value
    Closed(2) = MainForm!optMiddleTip(0).Value
    Closed(3) = MainForm!optIndexTip(0).Value
    Closed(4) = MainForm!optThumbTip(0).Value
    
    nXStep = ((g_Labels(1).X - g_Labels(5).X) / 4) / ISCALE
    nYStep = ((g_Labels(1).y - g_Labels(5).y) / 4) / ISCALE
    g_nCurrTextAngle = 0
    PR_MakeXY xyStart, (g_Labels(5).X / ISCALE) + (nXStep / 2), (g_Labels(5).y / ISCALE) + (nYStep / 2)
    PR_RevisePoint xyStart

   'Loop through the fingers
    For ii = 0 To 3
'        xyInsert.x = xyStart.x + (nXStep * ii) + .25
'        xyInsert.y = xyStart.y + (nYStep * ii)
'       Because I know about the translation and rotation
        xyInsert.X = xyStart.X + (nYStep * ii)
        xyInsert.y = xyStart.y + (nXStep * ii) + 0.25

        If Val(MainForm!txtLen(ii).Text) = 0 Then
           'Missing Finger
            If Closed(ii) = True Then
                PR_DrawText "CLOSED", xyInsert, 0.125
            Else
                PR_DrawText "OPEN", xyInsert, 0.125
            End If
        End If
    
    Next ii
    g_nCurrTextAngle = 0

   'Place markers at lowest Web positions
   'For latter use with Zips, if a slant insert given then
   'supply the length of the slant insert
    PR_SetLayer "Construct"

   'Palmer and LFS zipper
   'based on little finger web
    If g_Labels(4).Found = True Then
       'Use little finger web
        PR_MakeXY xyInsert, g_Labels(4).X / ISCALE, g_Labels(4).y / ISCALE
        PR_RevisePoint xyInsert
    Else
       'Use start of Little Finger Side Arc
        PR_MakeXY xyInsert, g_Labels(5).X / ISCALE, g_Labels(5).y / ISCALE
        PR_RevisePoint xyInsert
    End If
    PR_MakeXY xyPt, xyInsert.X, xyInsert.y
    PR_DrawMarker xyPt
    PR_AddDBValueToLast "Zipper", "PALMER-WEB"
    If MainForm!chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If

   'Get diff between Little finger side and thumb side
   'to be used if a finger is missing
    nXStep = (g_Labels(1).X - g_Labels(5).X) / ISCALE
    nYStep = (g_Labels(1).y - g_Labels(5).y) / ISCALE

   'Dorsal
    If g_Labels(3).Found = True Then
        PR_MakeXY xyInsert, (g_Labels(3).X / ISCALE), g_Labels(3).y / ISCALE
        PR_RevisePoint xyInsert
    Else
       'Use mid point of line joining Little finger side and thumb side
        PR_MakeXY xyInsert, (g_Labels(5).X / ISCALE) + (nXStep / 2), (g_Labels(5).y / ISCALE) + (nYStep / 2)
        PR_RevisePoint xyInsert
    End If
    PR_MakeXY xyPt, xyInsert.X, xyInsert.y
    PR_DrawMarker xyPt
    PR_AddDBValueToLast "Zipper", "DORSAL-WEB"
    If MainForm!chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If
    
   'Outside (Thumb Side)
    If g_Labels(2).Found = True Then
        PR_MakeXY xyInsert, g_Labels(2).X / ISCALE, g_Labels(2).y / ISCALE
        PR_RevisePoint xyInsert
    Else
       'Use end of thumb side
        PR_MakeXY xyInsert, g_Labels(1).X / ISCALE, g_Labels(1).y / ISCALE
        PR_RevisePoint xyInsert
    End If
    PR_MakeXY xyPt, xyInsert.X, xyInsert.y
    PR_DrawMarker xyPt
    PR_AddDBValueToLast "Zipper", "OUTSIDE-WEB"
    If MainForm!chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If


   'Wrist points for extension
    PR_MakeXY xyPalm(6), g_Labels(19).X / ISCALE, g_Labels(19).y / ISCALE
    PR_RevisePoint xyPalm(6)

    PR_MakeXY xyPalm(5), g_Labels(22).X / ISCALE, g_Labels(22).y / ISCALE
    PR_RevisePoint xyPalm(5)

    PR_MakeXY xyPalm(4), g_Labels(9).X / ISCALE, g_Labels(9).y / ISCALE
    PR_RevisePoint xyPalm(4)

    PR_MakeXY xyPalm(3), g_Labels(5).X / ISCALE, g_Labels(5).y / ISCALE
    PR_RevisePoint xyPalm(3)

    PR_MakeXY xyPalm(1), g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE
    PR_RevisePoint xyPalm(1)

    PR_DrawMarker xyPalm(1)
    PR_AddDBValueToLast "Zipper", "PALMER"
    PR_NamedHandle "hPalmer"

    For ii = 3 To 6
        PR_DrawMarker xyPalm(ii)
        PR_AddDBValueToLast "Data", "PALM" & Trim$(Str$(ii))
    Next ii

End Sub



Sub PR_DrawAA(xCenter#, Ycenter#, nSweepAngle#)
   'Draw the Arc based on the above
   'Absolute co-ordinates
   '
   'INPUT
   '
   '  S         Stack of points. (Global)
   '            The Top point is always the current point
   '
   '  Xcenter#  Center of Arc in absolute co-ordinates
   '  Ycenter#     "    "   "   "   "   "   "   "
   '            in 1000th of an Inch
   '
   '  SweepAngle
   '            Angle in degrees through which the arc is
   '            swept
   '            -ve Clockwise
   '            +ve Anti-Clockwise
   '
   'OUTPUT
   '  Drafix Macro command to draw an arc
   '
    Dim Xstart As Double, Ystart As Double
    Dim xyStart As XY, xyCenter As XY, xyLast As XY
    Dim nStartAngle#, nEndAngle#, nRadius#
    
   
   'Get start point from stack
    StackCurrent Xstart, Ystart
    PR_MakeXY xyStart, Xstart / ISCALE, Ystart / ISCALE
    PR_RevisePoint xyStart

   'Make center point
    PR_MakeXY xyCenter, xCenter / ISCALE, Ycenter / ISCALE
    PR_RevisePoint xyCenter

   'Get the angle
    nStartAngle = FN_CalcAngle(xyCenter, xyStart)

   'Calculate the position of the last point
    nRadius = FN_CalcLength(xyCenter, xyStart)
   
   'Draw the arc
   'Assumes that FN_Open has been called first

    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCenter.X); CC; "xyStart.y +"; Str$(xyCenter.y); CC; Str$(nRadius); CC; Str$(nStartAngle); CC; Str$(nSweepAngle); ");"
    Print #fNum, "SetDBData(hEnt,"; QQ; "ID"; QQ; ",sID);"


End Sub

Sub PR_DrawEOS()
   'Draw the points stored in the stack as a line.
   'This is a special variant of the procedure PR_DrawStack
   '
   'INPUT
   '
   '  S         Stack of points. (Global)
   '            The Top point is always the current point
   '
   'OUTPUT
   '  S         The stack will contain one point
   '            this will be the current point
   '
   '  DRAFIX Macro
   '            By using a stack the line is actually drawn
   '            in reverse to the input.
   '
   'NOTES
   '    If stack is empty then or contains only one point
   '    then nothing is drawn
   '

    Dim Xlast#, Ylast#, Xprev#, Yprev#, X#, y#
    Dim xyPt1 As XY, xyPt2  As XY

   'Return if empty
    If EmptyStack() Then Exit Sub

   'Store last point
    PopStack Xlast, Ylast

   'Check if empty (=> only one point)
    If EmptyStack() Then
        PushStack Xlast, Ylast  'Restore stack to original state
        Exit Sub
    End If

   'Restore stack to original state
    PR_MakeXY xyPt1, Xlast / ISCALE, Ylast / ISCALE
    PR_RevisePoint xyPt1

    PopStack X, y
    PR_MakeXY xyPt2, X / ISCALE, y / ISCALE
    PR_RevisePoint xyPt2

    ClearStack

   'Draw line in reverse
    PR_DrawLine xyPt2, xyPt1

   'Stack will now contain only one point
    PushStack Xlast, Ylast

End Sub
Public Sub PR_DrawGloveHPGL_File(sFileName As String)
   'Using the HPGL data file created by the Glove Basic
   'programme create a macro to draw the data
   '
   'Modules;
   '    miFileHPGL
   'NOTE
   '    This is not a complete HPGL interpreter. It is
   '    designed only to draw the sub-set of HPGL commands
   '    used by version 3.0 of the CAD Glove Basic Programme
   '
    Dim TmpX As Double, TmpY As Double, ii%
    Dim PU As Integer   'Boolean Pen Up
    Dim PD As Integer   'Boolean Pen Down
    Dim xyArcCen As XY
    Dim aArcSweep#
    

   'Values from FN_GetCommand
    Dim sMNemomic As String, nParams As Integer
    
   'Initialise
    ClearStack

    miFileHPGL = FreeFile
    Open sFileName For Input As #miFileHPGL

   'Intialise
    PU = True
    PD = False

    While FN_GetCommand(sMNemomic, nParams, Value())
        Select Case sMNemomic
        Case "IN"
            'Ignore
        
        Case "IP"
            'Ignore

        Case "PU"
            If PD Then
                PR_DrawStack
            End If
            PU = True
            PD = False
            If nParams > 0 Then
                ClearStack
                PushStack Val(Value(nParams - 1)), Val(Value(nParams))
            End If

        Case "PD"
            PU = False
            PD = True
            If nParams > 0 Then
                For ii = 1 To nParams Step 2
                    PushStack Val(Value(ii)), Val(Value(ii + 1))
                Next ii
            End If

        Case "AA"
            PR_EndPointAA Val(Value(1)), Val(Value(2)), Val(Value(3)), TmpX, TmpY
            If PD Then
                PR_DrawStack
                PR_DrawAA Val(Value(1)), Val(Value(2)), Val(Value(3))
               'Keep these values to allow the LFA arc to be stored
                xyArcCen.X = Val(Value(1))
                xyArcCen.y = Val(Value(2))
                aArcSweep = Val(Value(3))
            End If
            ClearStack
            PushStack TmpX, TmpY
        
        Case "PA"
            If nParams > 0 Then
                If PD Then
                    If nParams > 0 Then
                        For ii = 1 To nParams Step 2
                            PushStack Val(Value(ii)), Val(Value(ii + 1))
                        Next ii
                    End If
                Else
                    ClearStack
                    PushStack Val(Value(nParams - 1)), Val(Value(nParams))
                End If
            End If
        
        Case "SI"
           'Ignore

        Case "DI"
           'Ignore
        
        Case "LB"
           'This is used to pass info to the program.
           'The actual LB text value is not used except for
           'ii = 9 where this is the INSERT value
           '
           'Enable for De-Bug only
           'DrawText Value(1)
            
            Select Case Value(1)
            Case "EOS1", "EOS2", "EOS3"
                'End of Support Line
                'EOS1 => 2 Tapes beyond wrist
                'EOS2 => 1 Tape beyond wrist
                'EOS3 => 0 Tapes beyond wrist
                 If PD Then
                    If g_ExtendTo <> GLOVE_NORMAL Then
                        PR_SetLayer "Construct"
                        PR_DrawStack
                        PR_SetLayer "Template" & g_sSide
                    Else
                       'This is a very special case as
                       'the direction of the EOS line is
                       'very important to the zipper modules.
                       'We complicate here to simplify later (OK!)
                        PR_DrawEOS
                        PR_AddDBValueToLast "Zipper", "EOS"
                    End If
                 End If

            Case "LFS"
                'End of Little Finger Side line
                 If PD Then
                    If g_ExtendTo <> GLOVE_NORMAL Then
                        PR_SetLayer "Construct"
                        PR_DrawStack
                        PR_SetLayer "Template" & g_sSide
                    Else
                        PR_DrawStack
                        PR_AddDBValueToLast "Zipper", "LFS"
                        PR_AddDBValueToLast "ZipperLength", "%Length"
                    End If
                 End If

            Case "ETS"
                'End of Thumb Side line
                 If PD Then
                    If g_ExtendTo <> GLOVE_NORMAL Then
                        PR_SetLayer "Construct"
                        PR_DrawStack
                        PR_SetLayer "Template" & g_sSide
                    Else
                        PR_DrawStack
                        PR_AddDBValueToLast "Zipper", "ETS"
                        PR_AddDBValueToLast "ZipperLength", "%Length"
                    End If
                 End If

            Case "LFA"
                'End of Little Finger Side Arc
                    If g_ExtendTo <> GLOVE_NORMAL Then
                         PR_SetEntityData "layer", "Construct", Null, Null
                         PR_AddDBValueToLast "curvetype", "LFA-CONSTRUCT"
                    Else
                         PR_AddDBValueToLast "Zipper", "LFA"
                         PR_AddDBValueToLast "ZipperLength", "%Length"
                    End If
                    xyLFAcen.X = xyArcCen.X / ISCALE
                    xyLFAcen.y = xyArcCen.y / ISCALE
                    PR_RevisePoint xyLFAcen
                    aLFAsweep = aArcSweep
                    PR_SetLayer "Template" & g_sSide

            Case "MTA"
                'Middle of Thumb Arc
                 PR_SetLayer "Construct"
                 DrawMarker
                 PR_AddDBValueToLast "Zipper", "DORSAL"
                 PR_SetLayer "Notes" 'Bit of a bodge here

            Case Else
               'Numeric values stored and used in PR_AddNotesEtc
                ii = Val(Mid(Value(1), 1, 2))
                g_Labels(ii).Found = True
                g_Labels(ii).sString = Value(1)
                StackCurrent g_Labels(ii).X, g_Labels(ii).y

            End Select

        Case "LT"
            If PD Then
                PR_DrawStack
            End If
            SetLineType Value(1)
        
        Case Else
           'Ignore
        
        End Select

    Wend

    Close #miFileHPGL
    
   'Draw last bit
    If Not PU Then
        PR_DrawStack
    End If

   'Add Labels etc
    PR_AddNotesEtc

'    Close #fNum

End Sub

Sub PR_DrawStack()
   'Draw the point stored in the stack as a Polyline
   '
   'INPUT
   '
   '  S         Stack of points. (Global)
   '            The Top point is always the current point
   '
   'OUTPUT
   '  S         The stack will contain one point
   '            this will be the current point
   '
   '  DRAFIX Macro
   '            By using a stack the line is actually drawn
   '            in reverse to the input.
   '
   'NOTES
   '    If stack is empty then or contains only one point
   '    then nothing is drawn
   '

    Dim Xlast#, Ylast#, Xprev#, Yprev#, X#, y#
    Dim xyPt As XY

    Dim PolyLine As Curve

   'Return if empty
    If EmptyStack() Then Exit Sub

   'Store last point
    PopStack Xlast, Ylast

   'Check if empty (=> only one point)
    If EmptyStack() Then
        PushStack Xlast, Ylast  'Restore stack to original state
        Exit Sub
    End If

   'Restore stack to original state
    PushStack Xlast, Ylast

   'Impossible values for testing (I hope)
    PolyLine.n = 0  'Curve is set to empty

    Do While Not EmptyStack()
        PopStack X, y
       'Add point to curve
        PolyLine.n = PolyLine.n + 1
        PR_MakeXY xyPt, X / ISCALE, y / ISCALE
        PR_RevisePoint xyPt
        PolyLine.X(PolyLine.n) = xyPt.X
        PolyLine.y(PolyLine.n) = xyPt.y
    Loop

   'Draw only if more than one point
   'We don't need to explicty check this as it is part of PR_DrawPoly
    PR_DrawPoly PolyLine

   'Stack will now contain only one point
    PushStack Xlast, Ylast

End Sub

Sub PR_EndPointAA(xCenter#, Ycenter#, SweepAngle#, Xlast#, Ylast#)
   'Returns in the variables Xlast and Ylast the value of
   'the endpoint of the arc.
   'Absolute co-ordinates
   '
   'INPUT
   '
   '  S         Stack of points. (Global)
   '            The Top point is always the current point
   '
   '  Xcenter#  Center of Arc in absolute co-ordinates
   '  Ycenter#     "    "   "   "   "   "   "   "
   '
   '  SweepAngle
   '            Angle in degrees through which the arc is
   '            swept
   '            -ve Clockwise
   '            +ve Anti-Clockwise
   '
   'OUTPUT
   '  Xlast     X and Y co-ordinate of the end point of the
   '  Ylast     arc.  The start point being the current
   '            point given by the Stack S
   '
    Dim Xstart As Double, Ystart As Double
    Dim xyStart As XY, xyCenter As XY, xyLast As XY
    Dim nAngle#, nRadius#
    
   'Get start point from stack
    StackCurrent Xstart, Ystart
    PR_MakeXY xyStart, Xstart, Ystart

   'Make center point
    PR_MakeXY xyCenter, xCenter, Ycenter

   'Get the angle
    nAngle = FN_CalcAngle(xyCenter, xyStart)

   'Revise angle wrt to sweep angle
    nAngle = nAngle + SweepAngle
    If nAngle > 360 Then
        nAngle = nAngle - 360
    ElseIf nAngle < 0 Then
        nAngle = nAngle + 360
    End If
   
   'Calculate the position of the last point
    nRadius = FN_CalcLength(xyCenter, xyStart)
    PR_CalcPolar xyCenter, nAngle, nRadius, xyLast

   'Return values
    Xlast = xyLast.X
    Ylast = xyLast.y

End Sub

Sub PR_RevisePoint(xyPt As XY)
   'Procedure to translate and rotate a point so that
   'the final drawing of the glove is vertical rather than
   'the default horizontal
   '
   '
   'Constants in use
   '  XTRANSLATE
   '  YTRANSLATE
   '  ROTATION
   '  xyOrigin

    xyPt.X = xyPt.X + XTRANSLATE
    PR_CalcPolar xyOrigin, FN_CalcAngle(xyOrigin, xyPt) + ROTATION, FN_CalcLength(xyOrigin, xyPt), xyPt
    xyPt.y = xyPt.y + YTRANSLATE

End Sub


Sub PushStack(ByVal X As Double, ByVal y As Double)

    If S.Top = MAXSTACK Then
        MsgBox "Stack is Full", 16, "Stack Operations"
        End
    Else
        S.Top = S.Top + 1
        S.X(S.Top) = X
        S.y(S.Top) = y

    End If
    
End Sub

Sub SetLineType(Style)

    Select Case Style
        Case 5
           'Line on layer notes
            PR_SetLayer "Notes"
        Case Else
           'Default to solid line
            PR_SetLayer "Template" & g_sSide
    End Select

End Sub

Sub StackCurrent(X As Double, y As Double)
   'Retrieves the current position on the stack
    If Not EmptyStack() Then
        X = S.X(S.Top)
        y = S.y(S.Top)
    Else
       'Default to 0, Not particularly usefull
       'but saves an error message
        X = 0
        y = 0
    End If
End Sub

Public Function FN_GetCommand(Mnemonic As String, nParameter As Integer, Parameter() As Variant) As Boolean
   'Purpose:    Return an next HPGL command and it's parameters
   'Inputs:     miFileHPGL  file number of the sequentially opened HPGL file.
   '
   'Returns     True if a command found
   '            False if EOF HPGL file of HPGL file not openned
   'Arguments:  Mnemonic   - The HPGL Command string
   '            nParameter - Number of parameters for the command
   '            Parameter  - An array containg the command parameters
   '
   'Notes:      This procedure is designed to process only the restricted
   '            number of HPGL commands used by Kevin.  It is not a general
   '            tool!
   '
   'Known Bugs:
   '            Doesn't handle LB command very well (see notes)
   '

   'Scan the input stream breaking it into
   'Command Mnemonic and the Arguments to that command
   '
    Static CharPrev$

    Dim Char$, sParameter$, nAsc%, nParamLen%

   'Initialise
    Mnemonic = ""
    sParameter = ""
    nParamLen = 0
    nParameter = 0

   'Scan input stream.
   'First 2 characters will form the Mnemonic
   'get the rest of the stream up to the command terminator
   'the difficulty is that the explict terminator is ";"
   'however it can also be terminated by a new Mnemonic
   '
   'Thus
   '    PD;PA 1000 1000;PA 2000 2000 ...
   '
   'is equivelent to
   '
   '    PDPA1000 1000PA 2000 2000 ...
   '
   'The other problem is that the LB command can be followed
   'by any ASCI character and is terminated only by the ETX
   'Chr$(3)
   '
    
    FN_GetCommand = True
    
   'Exit if file not open
    If miFileHPGL < 0 Then
        FN_GetCommand = False
        Exit Function
    End If
    
    Do While Not EOF(miFileHPGL)
        If CharPrev = "" Then
            Char = Input(1, #miFileHPGL)
        Else
            Char = CharPrev
            CharPrev = ""
        End If
        nAsc = Asc(Char)
        Select Case nAsc
            Case 65 To 90, 97 To 122
               'A-Z, a-z
                If Len(Mnemonic) = 2 Then
                    If Mnemonic = "LB" Then
                       'Only a single parameter for "LB"
                        nParameter = 1
                        sParameter = sParameter & Char
                        nParamLen = nParamLen + 1
                Else
                       'Store this for later use
                        CharPrev = Char
                       'Save Prameter value (if any)
                        If nParamLen <> 0 Then
                            nParameter = nParameter + 1
                            Parameter(nParameter) = sParameter
                        End If
                        Exit Do
                    End If
                Else
                    Mnemonic = Mnemonic & Char
                End If

            Case 59
               ';
                If Len(Mnemonic) = 2 Then
                    If nParamLen <> 0 Then
                        nParameter = nParameter + 1
                        Parameter(nParameter) = sParameter
                    End If
                    Exit Do
                End If

            Case 3
               'ETX end of text, Indicates end of LB command
                nParameter = 1
                Parameter(1) = sParameter
                Exit Do
            
            Case Else
                If Mnemonic <> "LB" Then
                    If (nAsc = 32) Or (nAsc = 44) Then
                       'Space or Comma as parameter seperators
                       'Ignore leading space or empty parameter lists
                        If nParamLen <> 0 Then
                            nParameter = nParameter + 1
                            Parameter(nParameter) = sParameter
                            sParameter = ""
                            nParamLen = 0
                        End If
                    ElseIf (nAsc >= 48 And nAsc <= 57) Or nAsc = 46 Or nAsc = 45 Or nAsc = 43 Then
                       'Only 0-9 and .
                        sParameter = sParameter & Char
                        nParamLen = nParamLen + 1
                    End If
                Else
                   'All characters for LB Mnemonic
                    sParameter = sParameter & Char
                    nParamLen = nParamLen + 1
                End If

        End Select

    Loop

    If EOF(miFileHPGL) Then FN_GetCommand = False


End Function


Attribute VB_Name = "Kevin"
'Module:    Kevin
'File:      KEVIN.BAS
'Purpose:   Supply HPGL commands to draw a CAD Glove based
'           on the CAD Glove QuickBASIC 4.5 code developed
'           by Kevin Tucker - Beiersdorf Jobst Inc.
'
'
'Version:   1.01
'Date:      07.Dec.98
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Notes:-
'   What is all this I hear you exclaim! Sit back put your feet up and I shall
'   tell you the story of the CAD glove.
'
'   1. Original
'      The original CAD glove program was developed by Kevin Tucker for JOBST
'      either in Toledo OH or Charlotte NC.  The final programme was written in
'      Microsoft QuickBASIC 4.5 consisting of two chained modules KEVIN1.EXE
'      and KEVIN.EXE.
'      This was an interactive program running under DOS on an
'      IBM PC.  The operator was prompted for input and the output was displayed
'      on the screen.  Final output was sent to an local HP A3 size plotter and a
'      bundle sheet printed on a local printer.  No data was saved which meant that
'      repeat orders had to be re-input.
'
'   2. CAD System
'      A DRAFIX Cad system was introduced and it was decided to incorporate the
'      CAD Glove into this system. The DRAFIX CAD system ran under 3.1 using VB3
'      Dialogs to interact with the operators.
'      I decided not to recode the QuickBASIC into VB3 because the output from
'      the CAD Glove worked very well and I wanted to make the least possible
'      changes (if it ain't broke don't fix it)
'      The QuickBASIC code was modified to read from, and write to, files.  The CAD
'      Glove VB3 programme ran the QB code in a shell.  I then interpreted the HPGL
'      output file creating DRAFIX macros to draw the glove shape and any extension
'      along the arm.
'      Printing of the bundle sheet was disabled.
'
'   3. 32 bit upgrade.
'      The CAD glove uses BRUN45.EXE, this is the runtime required for QuickBASIC
'      programs.  It also runs the QB programs in a shell using the .PIF file which
'      makes it fiddly to relocate etc.
'      These pose potential problems as Y2K approaches and the program is
'      moved to a 32 bit platform.
'      I decided to bring the code from KEVIN1.BAS and KEVIN.BAS into VB5.  The
'      approach is still to modify the code as little as possible, hence the line
'      numbers and gosub commands et.al. are retained.
'      The name KEVIN for the module being retained to acknowledge the originator
'      Kevin Tucker.
'
'
 
Private C(15) As Single, L(15), LR(15), T(15), S(15), W(20), DT(8), P(70, 2), CC(10, 2) As Single, TR(70, 2), CCT(10, 2), HYP(200), ANGL(200) As Single, I6(100), J6(100), D$(20), Z$(30), DL(15)

' Variables (formerly declared as COMMON) used by the procedures Kevin and Kevin1
Private FLAG6, BDSHEET As Single, M$
Private B$, sn$, OD#, sOD$, HAND$, SUBFLAG, TYPES, PNO, sP$, NBR, PD$, CAT$, OZ$, LOCNO, ZL$, RsP$, SLI$, BOTH, n, RQTY, LQTY, INSERT, A2$, sINSERT$, SW, OpenTips, STDTIPS, F1, F2, F3, F4
Private LONGEST, XBASE, YBASE, INDEX, MIDDLE, RING, LITTLE, THUMB, WRISTL, XTI, YTI, X10, BETA, DIPI, DIPM, DIPR, DIPL, DESIGNER, sDESIGNER$, FLAG4, FLAG5, FLAG7, FLAG8, FLAG9
Private SEMICIRC, QUARTCIRC1, QUARTCIRC2, CONST1, CONST2, CONST3, YC, XR10, XRTI, JKR9, JK9, JK8, JK7, JUMPFLAG, NOPLOT, PI, MISSING, GA3, GA4, TBA2, GA10, XFLAG
Private TAPES, PLOTONCE, LOCY, CONST6, MINW, FBASE, SPACE, FINGER, SCREENPLOT, K, YC7, GA5, FLAG100, PLOTPALM, PG, ZIPLEN, sZIPLEN$, GA8, LTW, GA7, OZ2$, XXC, GRAM$
Private H21, A21, R21, TA21, TA22, T21, TIRATIO, LTW2, ITDIF
'C-Gem
Private JOBST$

Private Sub Kevin()
'******************************** ORIGINAL ***********************************
'FILENAME:      KEVIN.BAS
'REVISION:      UNKNOWN
'EXE DATE:      19/11/93 10:41:48
'BY:            KEVIN TUCKER - JOBST
'
'******************************** MODIFIED ***********************************
'MODIFIED:      G.GEORGE
'DATE:          28.Jun.95
'REASON:        Input is from file.
'               Output is to file.
'
'               The CAD Glove is now run from a Windows VB shell.
'               This shell takes the input from the user and creates
'               the data file used by this programme.  The output from
'               this programme is taken by the VB shell and input into
'               the JOBST CAD System as a DRAFIX drawing.
'
'******************************** MODIFIED ***********************************
'MODIFIED:      G.GEORGE
'DATE:          07.Dec.98
'REASON:        The CAD glove code has been modifed to run within VB 5 rather
'               than as an external procedure.
'               Minimal modifications have been made to the code and it still
'               reads input from a file and outputs the result to a file.
'


'$STATIC
'Dim C(15) As Single, L(15), LR(15), T(15), S(15), W(20), DT(8), P(70, 2), CC(10, 2) As Single, TR(70, 2), CCT(10, 2), HYP(200), ANGL(200) As Single, I6(100), J6(100), D$(20), Z$(30), DL(15)

'These have now been declared as module level variables
'COMMON SHARED FLAG6, BDSHEET, C() AS SINGLE, L(), LR(), T(), S(), W(), DT(), P(), CC() AS SINGLE, TR(), CCT(), HYP(), ANGL() AS SINGLE, I6(), J6(), D$(), Z$(), DL(), M$
'COMMON SHARED B$, n$, OD#, sOD$, HAND$, SUBFLAG, TYPES, PNO, sP$, NBR, PD$, CAT$, OZ$, LOCNO, ZL$, RsP$, SLI$, BOTH, n, RQTY, LQTY, INSERT, A2$, sINSERT$, SW, OPENTIPS, STDTIPS, F1, F2, F3, F4
'COMMON SHARED LONGEST, XBASE, YBASE, INDEX, MIDDLE, RING, LITTLE, THUMB, WRISTL, XTI, YTI, X10, BETA, DIPI, DIPM, DIPR, DIPL, DESIGNER, sDESIGNER$, FLAG4, FLAG5, FLAG7, FLAG8, FLAG9
'COMMON SHARED SEMICIRC, QUARTCIRC1, QUARTCIRC2, CONST1, CONST2, CONST3, YC, XR10, XRTI, JKR9, JK9, JK8, JK7, JUMPFLAG, NOPLOT, PI, MISSING, GA3, GA4, TBA2, GA10, XFLAG
'COMMON SHARED TAPES, PLOTONCE, LOCY, CONST6, MINW, FBASE, SPACE, FINGER, SCREENPLOT, K, YC7, GA5, FLAG100, PLOTPALM, PG, ZIPLEN, sZIPLEN$, GA8, LTW, GA7, OZ2$, XXC, GRAM$
'COMMON SHARED H21, A21, R21, TA21, TA22, T21, TIRATIO, LTW2, ITDIF
''GG
'COMMON SHARED JOBST$

      If JUMPFLAG = 1 Then
            GoTo 4800
      Else
            GoTo 4008
      End If

4002 SUBFLAG = 4000: Close #1: GoTo 4008 'Plotter Routine
4008 If SUBFLAG = 4000 And (Left$(HAND$, 1) = "R" Or Left$(HAND$, 1) = "r") Then GoSub 8000 'REINSTATES SIGNS FOR REPLOTS


4020 If Asc(M$) = 77 Or Asc(M$) = 109 Then
      
        K = CInt(1021 / 2.54)
        XAX = 13 * 2.54 * K
        YAX = 2.25 * 2.54 * K
        GoTo 4022
    
     End If

4021 K = 1021: XAX = 13 * K: YAX = Int(2.25 * K) 'User Units per inch,X-axis,Y-axis
4022 SEMICIRC = -180: CONST1 = -2: CONST2 = 1: CONST3 = -1: QUARTCIRC1 = -90: YC7 = 60: QUARTCIRC2 = 90
4023 If SUBFLAG = 4000 And PLOTONCE = 0 Then YC7 = 60: GoTo 4036
4025 P(59, 1) = XBASE: P(59, 2) = YBASE
4026 P(60, 1) = XBASE: P(60, 2) = YBASE - SW
4027 P(61, 2) = P(31, 2) - SW * Cos(GA5): P(61, 1) = P(31, 1) + CONST3 * SW * Sin(GA5)
4028 P(62, 2) = P(31, 2) - SW * (Cos(GA5) + Sin(GA5)): P(62, 1) = P(31, 1) + SW * (Sin(GA5) + CONST2 * Cos(GA5))
4029 'P(63, 2) = P(38, 2) - SW * SIN(GA7): P(63, 1) = P(38, 1) + CONST2 * SW * COS(GA7)
     LTW = Sqr((P(31, 1) - P(41, 1)) ^ 2 + P(41, 2) ^ 2)
     LTW2 = Sqr((P(38, 1) - P(31, 1)) ^ 2 + P(38, 2) ^ 2)
     ITDIF = LTW - LTW2
     P(63, 1) = P(31, 1) + Abs(LTW * Sin(GA7)) + Abs(SW * Cos(GA7)): P(63, 2) = P(31, 2) + LTW * Cos(GA7) - SW * Sin(GA7)
4030 'P(64, 2) = P(38, 2) + SW * SIN(GA6): P(64, 1) = P(38, 1) + CONST3 * SW * COS(GA6)
     GA8 = Atn((P(63, 2) - P(37, 2)) / Abs(P(37, 1) - P(63, 1)))
     'GA8 = ATN((P(38, 2) - P(37, 2)) / ABS(P(37, 1) - P(38, 1)))
     P(70, 1) = P(38, 1) - SW * Sin(GA8): P(70, 2) = P(38, 2) + SW * Cos(GA8)
     P(64, 1) = P(63, 1) - SW * Sin(GA8): P(64, 2) = P(63, 2) + SW * Cos(GA8)
4031 'P(66, 1) = P(37, 1) - SW / COS(GA6): P(66, 2) = P(37, 2)
     P(66, 1) = P(37, 1): P(66, 2) = P(37, 2) + SW
     'P(66, 1) = P(37, 1) - SW * SIN(GA8): P(66, 2) = P(37, 2) + SW * COS(GA8)

4032 P(67, 1) = P(46, 1): P(67, 2) = P(46, 2) - SW
4033 P(68, 1) = X10: P(68, 2) = P(37, 2)
4034 P(69, 1) = XTI: P(69, 2) = YTI
4035 P(65, 1) = P(38, 1) + SW * (Sin(GA6) - Cos(GA6)): P(65, 2) = P(38, 2) + SW * (Sin(GA6) + Cos(GA6))
4036 If (Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l") And PLOTONCE = 0 Then GoSub 8000
4038 GoSub 7000
' GG
'4040 OPEN "COM1:9600,S,7,1,RS,CS65535,DS,CD" FOR RANDOM AS #1
4040 Open "C:\JOBST\GLOVDATA.PLT" For Append As #1

4042 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 4043 Else GoTo 4044

4043 Print #1, "IN;SP2;IP1326,658,39934,26058;": GoTo 4047
4044 Print #1, "IN;SP2;IP522,259,15722,10259;"


4047 If PLOTONCE = 1000 Then GoTo 4275
     Print #1, "PU;PA"; XAX - Int(P(1, 2) * K), YAX + Int(P(1, 1) * K) + YC7; "PD"
    'GG
     Print #1, "LB9 " + Str$(INSERT) + Chr$(3)
     Print #1, "LB1 " + Chr$(3)
     
4048 If F1 = MISSING Then GoTo 4111

4050 For I = 2 To 4
4052   If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 4062
4054   If I = 4 And INDEX = 1000 Then GoTo 4070
4060   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4061   If DIPI = 100 And INDEX = 1000 Then I = 4
4062   If DIPI = 100 Then I = 4
4065   If I = 3 And STDTIPS = 1000 Then I = 4
4070 Next I

4075 If STDTIPS = 1000 Or INDEX = 1000 Then GoTo 4090
4080 Print #1, "AA"; XAX - Int(CC(3, 2) * K), YAX + Int(CC(3, 1) * K) + YC7, SEMICIRC, 10
4090 For I = 5 To 8
4092   If I = 5 And INDEX = 1000 Then GoTo 4110
4093   If I = 5 And STDTIPS = 1000 Then I = 6
4094   If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 4110
4096   If (I >= 5 And I <> 8) And DIPI = 100 And INDEX = 1000 Then I = 7
4100   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4110 Next I
    'GG
     Print #1, "LB2 " + Chr$(3)
4111 If F2 = MISSING Then GoTo 4151
4112 For I = 8 To 11
4113   If (P(I, 1) = 0 And P(I, 2) = 0) Or (I = 11 And MIDDLE = 1000) Then GoTo 4117
4114   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4115   If I = 10 And STDTIPS = 1000 Then I = 11
4116   If I = 9 And DIPM = 100 And MIDDLE = 1000 Then I = 11
4117 Next I
4118 If STDTIPS = 1000 Or MIDDLE = 1000 Then GoTo 4130
4120 Print #1, "AA"; XAX - Int(CC(4, 2) * K), YAX + Int(CC(4, 1) * K) + YC7, SEMICIRC, 10
4130 For I = 12 To 15
4132   If I = 12 And MIDDLE = 1000 Then GoTo 4150
4133   If I = 12 And STDTIPS = 1000 Then I = 13
4135   If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 4150
4136   If (I >= 12 And I <> 15) And DIPM = 100 And MIDDLE = 1000 Then I = 14
4140   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4150 Next I
    'GG
     Print #1, "LB3 " + Chr$(3)

4151 If F3 = MISSING Then GoTo 4191
4152 For I = 15 To 18
4153   If (P(I, 1) = 0 And P(I, 2) = 0) Or (I = 18 And RING = 1000) Then GoTo 4157
4154   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4155   If I = 17 And STDTIPS = 1000 Then I = 18
4156   If I = 16 And DIPR = 100 And RING = 1000 Then I = 18
4157 Next I
4158 If STDTIPS = 1000 Or RING = 1000 Then GoTo 4170
4160 Print #1, "AA"; XAX - Int(CC(5, 2) * K), YAX + Int(CC(5, 1) * K) + YC7, SEMICIRC, 10
4170 For I = 19 To 22
4172   If I = 19 And RING = 1000 Then GoTo 4190
4173   If I = 19 And STDTIPS = 1000 Then I = 20
4175   If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 4190
4176   If (I >= 19 And I <> 22) And DIPR = 100 And RING = 1000 Then I = 21
4180   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4190 Next I
    'GG
     Print #1, "LB4 " + Chr$(3)

4191 If F4 = MISSING Then GoTo 4235
4192 For I = 22 To 25
4193   If (P(I, 1) = 0 And P(I, 2) = 0) Or (I = 25 And LITTLE = 1000) Then GoTo 4197
4194   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4195   If I = 24 And STDTIPS = 1000 Then I = 25
4196   If I = 23 And DIPL = 100 And LITTLE = 1000 Then I = 25
4197 Next I
4198 If STDTIPS = 1000 Or LITTLE = 1000 Then GoTo 4210
4200 Print #1, "AA"; XAX - Int(CC(6, 2) * K), YAX + Int(CC(6, 1) * K) + YC7, SEMICIRC, 10
4210 For I = 26 To 29
4212   If I = 26 And LITTLE = 1000 Then GoTo 4230
4213   If I = 26 And STDTIPS = 1000 Then I = 27
4215   If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 4230
4216   If (I >= 26 And I <> 29) And DIPL = 100 And LITTLE = 1000 Then I = 28
4220   Print #1, "PA"; XAX - Int(P(I, 2) * K), YAX + Int(P(I, 1) * K) + YC7
4230 Next I
4233 GoTo 4236
4235 Print #1, "PA"; XAX - Int(P(29, 2) * K), YAX + Int(P(29, 1) * K) + YC7
4236
    'GG
     Print #1, "LB5 " + Chr$(3)
4237 Print #1, "AA"; XAX - Int(CC(7, 2) * K), YAX + Int(CC(7, 1) * K) + YC7, CONST1 * GA3 * 180 / PI, 3
    'GG
     Print #1, "LBLFA" + Chr$(3) 'End of Little Finger Side Arc

4238 Print #1, "PA"; XAX - Int(P(30, 2) * K), YAX + Int(P(30, 1) * K) + YC7
    'GG
     Print #1, "LB11" + Chr$(3) 'Wrist At Little Finger Side"

4239 If TAPES = 0 Then GoTo 4245
4240 Print #1, "PD;PA"; XAX - Int(P(52, 2) * K), YAX + Int(P(52, 1) * K) + YC7
    'GG
     If TAPES = 1 Then Print #1, "LBLFS" + Chr$(3) 'End of Little Finger Side ( 1 Tape above Wrist)
4241 If TAPES = 1 Then GoTo 4244
4242 Print #1, "PD;PA"; XAX - Int(P(54, 2) * K), YAX + Int(P(54, 1) * K) + YC7
    'GG
     Print #1, "LBLFS" + Chr$(3) 'End of Little Finger Side ( 2 Tape above Wrist)
4243 Print #1, "PA"; XAX - Int(P(55, 2) * K), YAX + Int(P(55, 1) * K) + YC7
    'GG
     Print #1, "LBEOS1" + Chr$(3) 'End Of Support
   
    'GG
'4244 PRINT #1, "PA"; XAX - INT(P(53, 2) * K), YAX + INT(P(53, 1) * K) + YC7: GOTO 4249'End of Support
4244 Print #1, "PA"; XAX - Int(P(53, 2) * K), YAX + Int(P(53, 1) * K) + YC7
     If TAPES = 1 Then
        Print #1, "LBEOS2" + Chr$(3)
     End If
     GoTo 4249
     
4245 Print #1, "PD;PA"; XAX - Int(P(40, 2) * K), YAX + Int(P(40, 1) * K) + YC7
    'GG
     Print #1, "LBLFS" + Chr$(3) 'End of Little Finger Side (0 Tapes above wrist)

4248 Print #1, "PA"; XAX - Int(P(39, 2) * K), YAX + Int(P(39, 1) * K) + YC7
    'GG
     Print #1, "LBEOS3" + Chr$(3) 'End of Support
4249 Print #1, "PA"; XAX - Int(P(31, 2) * K), YAX + Int(P(31, 1) * K) + YC7
    'GG
     Print #1, "LB19" + Chr$(3)

4257 If PLOTONCE = 1000 Then
      GoTo 4262
     Else
      Print #1, "PA"; XAX - Int(P(37, 2) * K), YAX + Int(P(37, 1) * K) + YC7
      Print #1, "LB22" + Chr$(3)
     End If
4258 Print #1, "PA"; XAX - Int(P(41, 2) * K), YAX + Int(P(41, 1) * K) + YC7: ' GOTO 4272'******
    'GG
     Print #1, "LBETS" + Chr$(3) 'End of Thumb Side
4262 Print #1, "PU;PA"; XAX - Int(P(31, 2) * K), YAX + Int(P(31, 1) * K) + YC7
    'GG
     Print #1, "LB10" + Chr$(3)   'Wrist at Thumb Side
4263 If LR(9) / 2 <= W(11) / 2 Then GoTo 4267
4264 Print #1, "PU;PA"; XAX - Int(YBASE * K), YAX + Int(XBASE * K) + YC7
    'GG
     Print #1, "LB10" + Chr$(3)   'Wrist at Thumb Side
4265 If PLOTONCE = 1000 Then Print #1, "LT" Else Print #1, "LT 5"
     Print #1, "PD;AA"; XAX - Int(CC(1, 2) * K), YAX + Int(CC(1, 1) * K) + YC7, CONST2 * (90 - TBA2 * 180 / PI), 5
4266 GoTo 4269
4267 Print #1, "PD;PA"; XAX - Int(P(46, 2) * K), YAX + Int(P(46, 1) * K) + YC7
4268 If PLOTONCE = 1000 Then Print #1, "LT" Else Print #1, "LT 5"
     Print #1, "AA"; XAX - Int(CC(1, 2) * K), YAX + Int(CC(1, 1) * K) + YC7, QUARTCIRC2, 5
    'GG
4269 Print #1, "LBMTA" + Chr$(3)   'Thumb Side Middle point on thumb palm arc
     Print #1, "AA"; XAX - Int(CC(2, 2) * K), YAX + Int(CC(2, 1) * K) + YC7, CONST2 * GA4 * 180 / PI, 5
'4269 PRINT #1, "AA"; XAX - INT(CC(2, 2) * K), YAX + INT(CC(2, 1) * K) + YC7, CONST2 * GA4 * 180 / PI, 5
4270 'IF PLOTONCE = 1000 THEN PRINT #1, "PD";  ELSE PRINT #1, "PU";
4271 Print #1, "PA"; XAX - Int(P(41, 2) * K), YAX + Int(P(41, 1) * K) + YC7
4272 Print #1, "PD;LT;PA"; XAX - Int(P(1, 2) * K), YAX + Int(P(1, 1) * K) + YC7
4273 If HAND$ = "RIGHT" Then I = 8 Else I = 22
     Print #1, "PU;PA"; XAX - Int(P(40, 2) / 2 * K), YAX + Int(P(I, 1) * K) + YC7

'4274 PRINT #1, "SI.150,.250;DI0,1;LB" + LEFT$(HAND$, 1) + sOD$ + CHR$(3)
' GG
4274 Print #1, "SI.150,.250;DI0,1;LB6" + Chr$(3) 'Position for Patient Details Symbol
     Print #1, "PU"; XAX - Int(CC(1, 2) * K), YAX + Int((2 * CONST2 * SW) * K) + YC7
     Print #1, "PD,PA"; XAX - Int((CC(1, 2) + INSERT + ITDIF) * K), YAX + Int((2 * CONST2 * SW) * K) + YC7
     Print #1, "PA"; XAX - Int((CC(1, 2) + INSERT) * K), YAX + Int((2 * CONST2 * SW + CONST2 * INSERT) * K) + YC7
     Print #1, "PA"; XAX - Int(CC(1, 2) * K), YAX + Int((2 * CONST2 * SW + CONST2 * INSERT) * K) + YC7, "PU"

4275 'Draw the thumb separately
4277 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 4278 Else GoTo 4279
4278 If PLOTONCE = 1000 Then XXC = 500 Else XXC = 50
     XC = Int((P(15, 2) + LONGEST) * K) - XXC: YC = 1510: SW = 0.3175: GoTo 4280 '**********

4279 If PLOTONCE = 1000 Then XXC = 200 Else XXC = 20
     XC = Int((P(15, 2) + 2 * LONGEST / 3) * K) - XXC: YC = 2 * K: SW = 0.125

4280 GoSub 7130
4290 'Close #1 ,Reverse signs and Reopen
4292 Print #1, "PU;SP0;PA0,0;": FLAG100 = 6: Close #1
'4294 CLS : IF PLOTPALM = 1000 THEN GOTO 4800
4294 If PLOTPALM = 1000 Then GoTo 4800
4300 'Draw the Palm surface pieces
4313 GoSub 8000 '***************
4314 GoSub 8200 '***************
4320 GoTo 4040
4325 GoSub 7130
4400 Print #1, "PU;SP0;PA0,0";: Close #1: GoTo 4800 '*******************
4599 PLOTONCE = 1000: GoSub 8000
4600 Print #1, "PU;SP0";: Close #1
4602 GoTo 4800 'PRINT THE PATIENT'S MEASUREMENTS

4800 PLOTONCE = 0: PLOTPALM = 0: GoSub 9000: GoTo 5300
5100 'ERROR CHECKING ROUTINE FOR THE PLOTTER
5300 ' CLOSING STATEMENT TO CHECK FOR GOOD PLOT

5302 'CLS : COLOR 15, 9, 11
5304 On KEY(4) GoSub 5400  '350: FLAG4 = 1'Data Check and Corrections, then Replot
5305 On KEY(5) GoSub 5450  '300: FLAG5 = 1'Designate an Insert size

5306 On KEY(6) GoSub 4002: FLAG6 = 1 'Replot without changes

5307 On KEY(7) GoSub 8500: FLAG7 = 1 'Run the left hand

5308 On KEY(8) GoSub 5460 '8    'Run another Patient

5310 On KEY(10) GoSub 5367 'End of Program Statement
     On KEY(9) GoSub 9000: FLAG9 = 1 'Reprint Bundle Sheet
'GG
'5312 KEY(4) ON: KEY(5) ON: KEY(6) ON: KEY(7) ON: KEY(8) ON: KEY(9) ON: KEY(10) ON
5312 'KEY(4) ON: KEY(5) ON: KEY(6) ON: KEY(7) ON: KEY(8) ON: KEY(9) ON: KEY(10) ON
5315 LOCATE 5, 20: Print #3, "SELECT ONE OF THESE FUNCTIONS"
5320 LOCATE 8, 15: Print #3, "TO CHANGE ANY MEASUREMENTS AND REPLOT, PRESS F4"
5322 LOCATE 10, 15: Print #3, "TO DESIGNATE AN INSERT AND REPLOT, PRESS F5"
5325 LOCATE 12, 15: Print #3, "TO REPLOT THIS SAME GLOVE , PRESS F6"
5327 LOCATE 14, 15: Print #3, "TO NOW DESIGN "; sn$; "'S LEFT GLOVE , PRESS F7"

5330 LOCATE 16, 15: Print #3, "TO RERUN THE PROGRAM FOR ANOTHER PATIENT, PRESS F8"
     LOCATE 18, 15: Print #3, "TO REPRINT THE BUNDLE SHEET, PRESS F9."
5335 LOCATE 20, 15: Print #3, "TO END THIS PROGRAM, PRESS F10"

'Exit
'5355    A1$ = INKEY$
'        IF A1$ = "" THEN GOTO 5355

5367 'CLS
'Exit this sub
5370 LOCATE 12, 10: Print #3, "Thank you. Enjoy what's left of your day. ": Print #3,: Print #3,: Exit Sub  'End

5400    FLAG4 = 1
       'GG
        Call Kevin1

5450    FLAG5 = 1
       'GG
5460    Call Kevin1

7000 'SUB TO ROTATE THUMBS
7010 For I = 1 To 40
7015   HYP(I) = Sqr(P(30 + I, 2) ^ 2 + P(30 + I, 1) ^ 2)
7016   If P(30 + I, 2) = 0 And P(30 + I, 1) > 0 Then ANGL(I) = 0: GoTo 7039
7017   If P(30 + I, 2) = 0 And P(30 + I, 1) < 0 Then ANGL(I) = PI: GoTo 7039
7018   If P(30 + I, 2) = 0 And P(30 + I, 1) = 0 Then GoTo 7050
7019   If P(30 + I, 1) = 0 And P(30 + I, 2) > 0 Then ANGL(I) = PI / 2: GoTo 7039
7020   If P(30 + I, 1) = 0 And P(30 + I, 2) < 0 Then ANGL(I) = 3 * PI / 2: GoTo 7039
7025   ANGL(I) = Atn(Abs(P(30 + I, 2)) / Abs(P(30 + I, 1)))
7030   If P(30 + I, 1) < 0 And P(30 + I, 2) > 0 Then ANGL(I) = PI - ANGL(I)
7035   If P(30 + I, 1) > 0 And P(30 + I, 2) < 0 Then ANGL(I) = 2 * PI - ANGL(I)
7037   If P(30 + I, 1) < 0 And P(30 + I, 2) < 0 Then ANGL(I) = PI + ANGL(I)
7039   If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then GoTo 7045
7040   TR(30 + I, 1) = HYP(I) * Cos(ANGL(I) - PI / 4)
7042   TR(30 + I, 2) = HYP(I) * Sin(ANGL(I) - PI / 4): GoTo 7050
7045   TR(30 + I, 1) = HYP(I) * Cos(ANGL(I) + PI / 4)
7047   TR(30 + I, 2) = HYP(I) * Sin(ANGL(I) + PI / 4)
7050 Next I

7060 For I = 1 To 2
7065   HYP(I) = Sqr(CC(I, 2) ^ 2 + CC(I, 1) ^ 2)

7070   ANGL(I) = Atn(Abs(CC(I, 2)) / Abs(CC(I, 1)))
7072   If CC(I, 1) < 0 And CC(I, 2) > 0 Then ANGL(I) = PI - ANGL(I): GoTo 7079
7074   If CC(I, 1) > 0 And CC(I, 2) < 0 Then ANGL(I) = 2 * PI - ANGL(I): GoTo 7079
7076   If CC(I, 1) < 0 And CC(I, 2) < 0 Then ANGL(I) = PI + ANGL(I)
7079   If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then GoTo 7085
7080   CCT(I, 1) = HYP(I) * Cos(ANGL(I) - PI / 4)
7082   CCT(I, 2) = HYP(I) * Sin(ANGL(I) - PI / 4): GoTo 7090
7085   CCT(I, 1) = HYP(I) * Cos(ANGL(I) + PI / 4)
7087   CCT(I, 2) = HYP(I) * Sin(ANGL(I) + PI / 4)
7090 Next I
7100 For I = 8 To 9 Step 1
7103   HYP(I) = Sqr(CC(I, 2) ^ 2 + CC(I, 1) ^ 2)
7105   ANGL(I) = Atn(Abs(CC(I, 2)) / Abs(CC(I, 1)))
7107   If CC(I, 1) < 0 And CC(I, 2) > 0 Then ANGL(I) = PI - ANGL(I): GoTo 7111
7109   If CC(I, 1) > 0 And CC(I, 2) < 0 Then ANGL(I) = 2 * PI - ANGL(I): GoTo 7111
7110   If CC(I, 1) < 0 And CC(I, 2) < 0 Then ANGL(I) = PI + ANGL(I)
7111   If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then GoTo 7115
7112   CCT(I, 1) = HYP(I) * Cos(ANGL(I) - PI / 4)
7114   CCT(I, 2) = HYP(I) * Sin(ANGL(I) - PI / 4): GoTo 7120
7115   CCT(I, 1) = HYP(I) * Cos(ANGL(I) + PI / 4)
7117   CCT(I, 2) = HYP(I) * Sin(ANGL(I) + PI / 4)
7120 Next I
7121 'Draw the top portion of the glove as a separate piece
7125 GoSub 8200
7127 Return  'TO 4040
7130 'Sub to draw the top portion of the glove as a separate piece
     If PLOTONCE = 0 Then
       XC = 0: YC7 = 5 * K
       If Asc(M$) = 77 Or Asc(M$) = 109 Then YC7 = 2.5 * YC7
     End If
7131 If LR(9) / 2 <= W(11) / 2 Then GoTo 7134
7132 Print #1, "PU;PA"; XAX - Int(TR(60, 2) * K) - XC, YAX + Int(TR(60, 1) * K) + YC7
7134 Print #1, "PU;PA"; XAX - Int(TR(31, 2) * K) - XC, YAX + Int(TR(31, 1) * K) + YC7
7135 Print #1, "PD;PA"; XAX - Int(TR(36, 2) * K) - XC, YAX + Int(TR(36, 1) * K) + YC7
7140 Print #1, "PA"; XAX - Int(TR(35, 2) * K) - XC, YAX + Int(TR(35, 1) * K) + YC7
7145 If STDTIPS = 1000 Or THUMB = 1000 And LR(5) <> 0 Then Print #1, "PA"; XAX - Int(TR(33, 2) * K) - XC, YAX + Int(TR(33, 1) * K) + YC7: GoTo 7165
7150 Print #1, "AA"; XAX - Int(CCT(9, 2) * K) - XC, YAX + Int(CCT(9, 1) * K) + YC7, CONST3 * 90, 10; "PU;PA"; XAX - Int(TR(33, 2) * K) - XC, YAX + Int(TR(33, 1) * K) + YC7; "PD";
7154 For I = 1 To 91
7155   If I6(I) = 0 And J6(I) = 0 Then GoTo 7160
7156   Print #1, "PA"; XAX - Int((CCT(8, 2) + CONST2 * J6(I)) * K) - XC, YAX + Int((CCT(8, 1) + CONST2 * I6(I)) * K) + YC7
7160 Next I
7165 Print #1, "PU;PA"; XAX - Int(TR(33, 2) * K) - XC, YAX + Int(TR(33, 1) * K) + YC7, "PD;PA"; XAX - Int(TR(45, 2) * K) - XC, YAX + Int(TR(45, 1) * K) + YC7

7170 Print #1, "PU;PA"; XAX - Int(TR(31, 2) * K) - XC, YAX + Int(TR(31, 1) * K) + YC7; "PD";
7175 If PLOTONCE = 1000 Then GoTo 7225
7180 Print #1, "PA"; XAX - Int(TR(61, 2) * K) - XC, YAX + Int(TR(61, 1) * K) + YC7
7190 Print #1, "PA"; XAX - Int(TR(62, 2) * K) - XC, YAX + Int(TR(62, 1) * K) + YC7
7200 Print #1, "PA"; XAX - Int(TR(63, 2) * K) - XC, YAX + Int(TR(63, 1) * K) + YC7
7210 Print #1, "PA"; XAX - Int(TR(64, 2) * K) - XC, YAX + Int(TR(64, 1) * K) + YC7
7214 ' IF XFLAG = 1000 THEN GOTO 7218
7216 'PRINT #1, "PA"; XAX - INT(TR(69, 2) * K) - XC, YAX + INT(TR(69, 1) * K) + YC7: GOTO 7239
7218 ' PRINT #1, "PA"; XAX - INT(TR(68, 2) * K) - XC, YAX + INT(TR(68, 1) * K) + YC7
     'IF ABS(SW / COS(GA6)) > ABS(P(32, 1) - P(37, 1)) THEN
      ' PRINT #1, "PA"; XAX - INT(TR(44, 2) * K) - XC, YAX + INT(TR(44, 1) * K) + YC7
      ' PRINT #1, "PA"; XAX - INT(TR(45, 2) * K) - XC, YAX + INT(TR(45, 1) * K): GOTO 7222
     'END IF
7219 Print #1, "PA"; XAX - Int(TR(66, 2) * K) - XC, YAX + Int(TR(66, 1) * K) + YC7
7220 Print #1, "PA"; XAX - Int(TR(45, 2) * K) - XC, YAX + Int(TR(45, 1) * K) + YC7
     'PRINT #1, "PA"; XAX - INT(TR(32, 2) * K) - XC, YAX + INT(TR(32, 1) * K) + YC7

7222 'PRINT #1, "PU;PA"; XAX - INT(TR(38, 2) * K) - XC, YAX + INT(TR(38, 1) * K) + YC7
     'PRINT #1, "PD;PA"; XAX - INT(TR(37, 2) * K) - XC, YAX + INT(TR(37, 1) * K) + YC7
     'PRINT #1, "PA"; XAX - INT(TR(32, 2) * K) - XC, YAX + INT(TR(32, 1) * K) + YC7
     If PLOTONCE = 0 Then GoTo 7282
7225 Print #1, "PD;PA"; XAX - Int(TR(42, 2) * K) - XC, YAX + Int(TR(42, 1) * K) + YC7
7227 If LR(9) / 2 <= W(11) / 2 Then GoTo 7235
7229 Print #1, "PD;PA"; XAX - Int(TR(60, 2) * K) - XC, YAX + Int(TR(60, 1) * K) + YC7
7231 Print #1, "AA"; XAX - Int(CCT(1, 2) * K) - XC, YAX + Int(CCT(1, 1) * K) + YC7, CONST2 * (90 - TBA2 * 180 / PI), 5
7233 GoTo 7239
7235 Print #1, XAX - Int(TR(67, 2) * K) - XC, YAX + Int(TR(67, 1) * K) + YC7
7237 Print #1, "AA"; XAX - Int(CCT(1, 2) * K) - XC, YAX + Int(CCT(1, 1) * K) + YC7, CONST2 * 90, 5
7239 Print #1, "AA"; XAX - Int(CCT(2, 2) * K) - XC, YAX + Int(CCT(2, 1) * K) + YC7, CONST2 * GA10 * 180 / PI, 5
7270 'PRINT #1, "PA"; XAX - INT(TR(65, 2) * K) - XC, YAX + INT(TR(65, 1) * K) + YC7
     Print #1, "PA"; XAX - Int(TR(70, 2) * K) - XC, YAX + Int(TR(70, 1) * K) + YC7
7275 Print #1, "PA"; XAX - Int(TR(66, 2) * K) - XC, YAX + Int(TR(66, 1) * K) + YC7
7280 Print #1, "PA"; XAX - Int(TR(45, 2) * K) - XC, YAX + Int(TR(45, 1) * K) + YC7
     'PRINT #1, "PA"; XAX - INT(TR(32, 2) * K) - XC, YAX + INT(TR(32, 1) * K) + YC7

7282 Print #1, "PU"; XAX - Int(TR(32, 2) * K) - XC, YAX + YC7
'GG
'7284 PRINT #1, "SI.15,.25;DI1,0;LB" + LEFT$(HAND$, 1) + sOD$ + CHR$(3)
7284 If PLOTONCE = 1000 Then
        Print #1, "SI.15,.25;DI1,0;LB8 " + Chr$(3)
     Else
        Print #1, "SI.15,.25;DI1,0;LB7 " + Chr$(3)
     End If
7285 'NOW PLOT THE PALM PORTIONS OF THE GLOVE
7287 If PLOTONCE = 1000 Then PLOTPALM = 1000: GoTo 7290
7288 If Left$(HAND$, 1) = "R" Or Left$(HAND$, 1) = "r" Then PLOTONCE = 1000: 'GOSUB 8000
7289 If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then PLOTONCE = 1000: 'GOSUB 800
7290 Return  'GOSUB 4290
8000 'SUB TO PLOT PALM SURFACE PARTS
8005 XBASE = -XBASE: X10 = -X10: XTI = -XTI: SEMICIRC = -SEMICIRC: QUARTCIRC1 = -QUARTCIRC1: QUARTCIRC2 = -QUARTCIRC2
8006 CONST1 = -CONST1: CONST2 = -CONST2: CONST3 = -CONST3: YC = 5 * K: 'YC2=-4084
8007 XR10 = -XR10: XRTI = -XRTI: JKR9 = -JKR9: JK9 = -JK9: JK8 = -JK8: JK7 = -JK7
8010 For I = 1 To 70
8020  P(I, 1) = -P(I, 1): TR(I, 1) = -TR(I, 1)
8030 Next I
8040 For I = 1 To 9
8050  CC(I, 1) = -CC(I, 1): CCT(I, 1) = -CCT(I, 1)
8060 Next I
8070 If PLOTONCE = 1000 Then YC7 = 5 * K: 'RETURN 4042
     If Asc(M$) = 77 Or Asc(M$) = 109 Then YC7 = 2.5 * YC7

8075 'IF PLOTLEFT=1000 THEN YC7=5*K : RETURN 4042 ELSE GOTO 6040
8080 Return

8200 ' SUB TO CALCULATE ELLIPTIC ARC
8205 For I = PI / 2 To PI Step 2 * PI / 180
        I5 = ((LR(5) / 2) * Cos(I))
        J5 = ((2 * W(9) / 3) * Sin(I))
        iI6 = (I5 * Cos(CONST3 * PI / 4) + CONST3 * J5 * Sin(CONST3 * PI / 4))
        iJ6 = (I5 * Sin(CONST3 * PI / 4) + CONST2 * J5 * Cos(CONST3 * PI / 4))
8210   CNT = CNT + 1: HYP(CNT) = Sqr(iJ6 ^ 2 + iI6 ^ 2)
8215   ANGL(CNT) = Atn(Abs(iJ6) / Abs(iI6))
8220   If iI6 < 0 And iJ6 > 0 Then ANGL(CNT) = PI - ANGL(CNT)
8225   If iJ6 < 0 And iI6 > 0 Then ANGL(CNT) = 2 * PI - ANGL(CNT)
8230   If iJ6 < 0 And iI6 < 0 Then ANGL(CNT) = PI + ANGL(CNT)
8235   I6(CNT) = HYP(CNT) * Cos(ANGL(CNT) - CONST2 * PI / 4)
8240   J6(CNT) = HYP(CNT) * Sin(ANGL(CNT) - CONST2 * PI / 4)
8245 Next I

8250 CNT = 0
        Return 'TO 7125 OR 4314

8500 'SUB TO RESET ALL DATA ARRAYS

8505 SUBFLAG = 8500

Erase P, TR, C, L, LR, T, S, W, DT, CC, CCT, HYP, ANGL, I6, J6

'8510 FOR I = 1 TO 15
'8520   C(I) = 0: L(I) = 0: LR(I) = 0: T(I) = 0: S(I) = 0
'8530 NEXT I
'8540 FOR I = 1 TO 20
'8550   W(I) = 0
'8560 NEXT I
'8565 FOR I = 1 TO 8
'8570   DT(I) = 0
'8575 NEXT I

'8580 FOR I = 1 TO 70
'8585   P(I, 1) = 0: P(I, 2) = 0: TR(I, 1) = 0: TR(I, 2) = 0
'8590 NEXT I

'8595 FOR I = 1 TO 9
'8600  CC(I, 1) = 0
 '       CC(I, 2) = 0
 '       CCT(I, 1) = 0
 '       CCT(I, 2) = 0
'8605 NEXT I

'8610 FOR I = 1 TO 200
'8615  HYP(I) = 0: ANGL(I) = 0
'8620 NEXT I

'8625 FOR I = 1 TO 100
'8630  I6(I) = 0: J6(I) = 0
'8635 NEXT I

'REDIM THE ARAYS WHICH WERE CLEARED

'REDIM C(15), L(15), LR(15), T(15), S(15), W(20), DT(8), P(70, 2), CC(10, 2), TR(70, 2), CCT(10, 2), HYP(200), ANGL(200), I6(100), J6(100)

8640 TAPES = 0: SPACE = 0: OpenTips = 0: STDTIPS = 0: INDEX = 0: MIDDLE = 0: RING = 0: LITTLE = 0
8650 THUMB = 0: INSERT = 0: DIPI = 0: DIPM = 0: DIPR = 0: DIPL = 0: CHILD = 0
8660 FLAG4 = 0: FLAG5 = 0: FLAG6 = 0: FLAG7 = 0: FLAG8 = 0: LONGEST = 0

8665 PLOTONCE = 0: PLOTPALM = 0: YC7 = 0

8680 'COLOR 15, 9, 11: CLS
8685 LOCATE 4, 3: Print #3, "  Customer's name :"; sn$
8690 LOCATE 6, 5: Print #3, sn$; "'s  File Number :"; sOD$

8695 LOCATE 8, 3: Print #3, "  Now, let's do "; sn$; "'s Left hand."
        HAND$ = "LEFT"
'GG
'8700 RETURN 8705
8705  JUMPFLAG = 2
      'GG
      Call Kevin1


9000 'SUB TO PRINT THE BUNDLE SHEET
     'GG - Don't print the bundle sheet at all
     GoTo 9200 'DONT PRINT BUNDLE SHEET
     
'GG Start
'     If FLAG9 = 1 Then GoTo 9002
'9001 If SUBFLAG = 4000 Or BDSHEET = 1000 Or NOPLOT = 1000 Then GoTo 9200
'
'9002 If BOTH <> 2 Then P1$ = "1": P2$ = "1": LPRINT: GoTo 9006
'9004 LPRINT
'9005 P2$ = "2": P1$ = Str$(PG + 1): PG = 1
'9006 LPRINT TAB(2); "PAGE "; P1$; " OF "; P2$; TAB(20); "TOTAL PRODUCTS ON ORDER "; TYPES
'9010 LPRINT : LPRINT TAB(2); sP$
'9011 If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then QTY = LQTY Else QTY = RQTY
'9012 LPRINT : LPRINT TAB(2); "TOTAL "; HAND$; " GLOVES ON ORDER "; QTY; TAB(57); CAT$
'9020 LPRINT TAB(57); HAND$
'
'9025 LPRINT : LPRINT TAB(2); sn$: LPRINT
'
'9040 LPRINT TAB(2); "DESIGN DATE "; DATE$; TAB(30); sOD$
'9050 LPRINT TAB(47); "CIRCUMFERENCES"; TAB(65); "LENGTHS"
'9055 LPRINT
'9060 For I = 1 To 11
'9065   If I = 2 Then Z$ = sn$ Else Z$ = ""
'9067   If I = 4 Then Z$ = PD$ + "     " + CAT$ + " " + HAND$ + " " + Str$(QTY) Else Z$ = ""
'9070   LPRINT TAB(2); Z$; TAB(49); I; TAB(53); STR$(C(I)); TAB(65); I + 11; TAB(69); STR$(L(I))
'9080 Next I
'     LPRINT
'     For I = 1 To TAPES
'       LPRINT TAB(49); 11 + I; TAB(53); STR$(C(11 + I)); TAB(61); "( "; 1.5 * I; " tape )"
'     Next I
'9085 LPRINT:  LPRINT
'9088 If Asc(M$) = 69 Or Asc(M$) = 101 Then GoTo 9092
'9090 LPRINT TAB(43); "INSERT WIDTH : "; STR$(INSERT); " CM ": GOTO 9094
'9092 LPRINT TAB(43); "INSERT WIDTH : "; INSERT * 8; " EIGHTHS "
'9094 LPRINT:  LPRINT
'9095 Z$ = "OPTIONS"
'     sZIPLEN$ = Str$(ZIPLEN)
'
'9100 LPRINT TAB(2); Z$
'     OZ2$ = Chr$(40) + " " + sZIPLEN$ + " " + Chr$(41)'
'     If ZIPLEN = 0 Then OZ2$ = ""
'9105 LPRINT TAB(10); ZL$ + " " + OZ$ + OZ2$
'9107 LPRINT TAB(10); SLI$
'9109 If RsP$ = "" Then
'       GoTo 9120
'     Else
'       LPRINT TAB(10); "REINFORCE " + RsP$
'     End If
'9120 LPRINT:  LPRINT: LPRINT
'9130 LPRINT TAB(40); "FABRIC PIECE # _____________________"
'9135 LPRINT
'9140 LPRINT TAB(40); "FABRIC GRAMS : USE_" + GRAM$ + "_RANGE FABRIC ONLY!"
'     LPRINT : LPRINT : LPRINT TAB(2); "SPECIAL INSTRUCTIONS : ________________________________________________________"
'     LPRINT : LPRINT TAB(2); "________________________________________________________________________________"
'
'9150 LPRINT:  LPRINT: LPRINT
'9160 LPRINT TAB(2); "MATCHER"; TAB(10); "CUTTER"; TAB(17); "SEWER"; TAB(23); "REINF"; TAB(29); "TRIM"; TAB(34); "ELASTIC"; TAB(42); "OPTIONS"; TAB(50); "EDGER"; TAB(56); "TOES"; TAB(61); "TACKER"; TAB(68); "INSP"; TAB(73); "PACKER"
'9180 LPRINT Chr$(12)
'9190 If BOTH = 2 And Val(P1$) = 2 Then BDSHEET = 1000
'9195 OZ$ = "": ZL$ = "": SLI$ = "": RsP$ = ""
'9197 FLAG9 = 0
'GG End
9200 Return
End Sub

Public Function FN_CADGloveFormat(nInches As Double) As String

    Dim nPrecision As Double, nDec As Double
    Dim iInt As Integer, iEighths As Integer
    Dim sString As String
    
    If nInches < 0 Then
        FN_CADGloveFormat = "0"
        Exit Function
    End If

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
    If iInt >= 10 Then
        sString = LTrim$(Str$(iInt))
    ElseIf iInt > 0 Then
        sString = "0" & LTrim$(Str$(iInt))
    Else
        sString = ""
    End If

    sString = sString & Trim$(Str$(iEighths))

       
   'Return formatted string
    FN_CADGloveFormat = sString

End Function
Public Sub PR_RunCADGlove()

    Const PLOTDATA = "C:\JOBST\GLOVDATA.PLT"

   'Use an error handler for Kill
   '
    On Error GoTo ErrorHandler
    
   'Delete existing plot data file this is because
   'the CAD GLOVE programme opens this file
   'with APPEND
    
    Kill PLOTDATA   'NB ErrorHandler for when no file is
                    'found
   'Call the VB5 port of the CAD Glove programme
    Call Kevin1
   
   'Close any open files
    Close
    
    Exit Sub
    
ErrorHandler:
   'Ignore error 53
   'Warn for all others
    If Err <> 53 Then    ' Error 53 is "File not Found".
        MsgBox "Error Running CAD Glove  programme.", 48, "CAD - Glove Dialogue"
    End If
    Resume Next
 
End Sub

Public Sub PR_CreateCADGloveData()
    Dim fGloveData As Integer
    Dim X As Integer

    Const GLOVEDATA = "C:\JOBST\GLOVDATA.DAT"

   'Create Glove data file that is input to
   'the CAD glove programme
    
    fGloveData = FreeFile
    Open GLOVEDATA For Output As #fGloveData
       'Path to executables
        Print #fGloveData, g_sPathJOBST; "\GLOVECAD"

       'Standard header
        Print #fGloveData, "VBDATA" 'Name
        Print #fGloveData, "1234"   'File Number
        Print #fGloveData, "1"      'Product types 1 only
        Print #fGloveData, "2"      'Burns
        Print #fGloveData, "N"      'One Glove only

       'Which glove Left or Right
        If MainForm!optHand(0).Value = True Then
            Print #fGloveData, "L"  'Left Hand
        Else
            Print #fGloveData, "R"  'Right Hand
        End If

        Print #fGloveData, "N"      'No Zipper
        Print #fGloveData, "N"      'No Slant Inserts
        Print #fGloveData, "N"      'No Reinforcements

       'Number of Fingers (Excluding thumb)
        Dim ii%, iFingers%
        iFingers = 0
        For ii = 0 To 3
            If Val(MainForm!txtLen(ii).Text) <> 0 Then iFingers = iFingers + 1
        Next ii
        Print #fGloveData, Trim$(Str$(iFingers))  'Number of fingers

        Print #fGloveData, "Y"      'Name is correct
        Print #fGloveData, "Y"      'File# is correct
        Print #fGloveData, "Y"      'Hand is correct
        Print #fGloveData, "1"      '1 glove only
        Print #fGloveData, "N"      'N to Lymphedema

       'Fabric range
        Dim sFabric$, nFabric%
'        sFabric = MainForm!cboFabric.List(MainForm!cboFabric.ListIndex)
        sFabric = MainForm!cboFabric.Text
        nFabric = Val(Mid(sFabric, 5, 3))

        If nFabric <= 190 Then
            Print #fGloveData, "A"  'A Range Fabric
        Else
            Print #fGloveData, "B"  'B Range Fabric
        End If

        Print #fGloveData, "E"      'English measurments

       'Circumferences and Lengths
        For ii = 0 To 10
            Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(ii).Text)))
        Next ii
        For ii = 0 To 8
            Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtLen(ii).Text)))
        Next ii

       'Tapes extending past wrist
        If Val(MainForm!txtCir(11).Text) <> 0 Then
            Print #fGloveData, "Y"      'Does extend beyond wrist
            If Val(MainForm!txtCir(12).Text) = 0 Then
                Print #fGloveData, "1"  'Only one tape past wrist
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(11).Text)))
            Else
                Print #fGloveData, "2"  'Two tapes past wrist
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(11).Text)))
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(12).Text)))
            End If
        Else
            Print #fGloveData, "N"      'Doesn't extend beyond wrist
        End If
        
        Print #fGloveData, "Y"      'Everything correct
        Print #fGloveData, "VB"     'Designers Name
        
       'Tips
        Dim nClosed%, nStdOpen%
        ReDim Closed(0 To 4) As Integer
        ReDim StdOpen(0 To 4) As Integer

        nClosed = 0
        nStdOpen = 0
        nClosed = Abs(MainForm!optLittleTip(0).Value) + Abs(MainForm!optRingTip(0).Value) + Abs(MainForm!optMiddleTip(0).Value) + Abs(MainForm!optIndexTip(0).Value) + Abs(MainForm!optThumbTip(0).Value)
        nStdOpen = Abs(MainForm!optLittleTip(1).Value) + Abs(MainForm!optRingTip(1).Value) + Abs(MainForm!optMiddleTip(1).Value) + Abs(MainForm!optIndexTip(1).Value) + Abs(MainForm!optThumbTip(1).Value)
       
       'The simple case is either all open or all closed
        If nClosed = 5 Then
           'All tips closed
            Print #fGloveData, "N"  'Open tips Not required
' The following commented out in response to the fact that
' the cutback for standard tips when performed by
' KEVIN.EXE produces strange and varing results.
' Now force all lengths to desired
'
'        ElseIf nStdOpen = 5 Then
'           'All tips open
'            Print #fGloveData, "Y"  'Open tips required
'            Print #fGloveData, "Y"  'On all 4 fingers and thumb
        Else
           'Some open, some closed, some desired
           'If a std tip is required then the finger length is
           'the given - std length
            Print #fGloveData, "Y"  'Open tips required
            Print #fGloveData, "N"  'On all fingers and thumb
            Dim nLen#
            Dim nTipCutBack#
        
           'Setup
            If Val(MainForm!txtAge.Text) <= 10 Then
                nTipCutBack = CHILD_STD_OPEN_TIP
            Else
                nTipCutBack = ADULT_STD_OPEN_TIP
            End If
            
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
            
            For ii = 0 To 4
                nLen = fnDisplayToInches(Val(MainForm!txtLen(ii).Text))

                If nLen > 0 And Closed(ii) <> True Then
                    Print #fGloveData, "Y"  'Open tip required
                    If StdOpen(ii) = True Then
                       'Cut back for a standard tip
                        Print #fGloveData, "N" 'No length is wrong
                        Print #fGloveData, FN_CADGloveFormat(nLen - nTipCutBack)
                    Else
                       'Yes this is the desired length
                        Print #fGloveData, "Y"
                    End If
                Else
                    Print #fGloveData, "N"  'Open tip not required
                End If
            Next ii
        End If

        Print #fGloveData, "Y"  'Force plot

    Close #fGloveData

End Sub
Private Sub Kevin1()
'******************************** ORIGINAL ***********************************
'FILENAME:      KEVIN1.BAS
'REVISION:      UNKNOWN
'EXE DATE:      19/11/93 10:41:48
'BY:            KEVIN TUCKER - JOBST
'
'******************************** MODIFIED ***********************************
'MODIFIED:      G.GEORGE
'DATE:          28.Jun.95
'REASON:        Input is from file.
'               Output is to file.
'
'               The CAD Glove is now run from a Windows VB shell.
'               This shell takes the input from the user and creates
'               the data file used by this programme.  The output from
'               this programme is taken by the VB shell and input into
'               the JOBST CAD System as a DRAFIX drawing.
'
'******************************** MODIFIED ***********************************
'MODIFIED:      G.GEORGE
'DATE:          07.Dec.98
'REASON:        The CAD glove code has been modifed to run within VB 5 rather
'               than as an external procedure.
'               Minimal modifications have been made to the code and it still
'               reads input from a file and outputs the result to a file.
'

'$STATIC
'Dim C(15) As Single, L(15), LR(15), T(15), S(15), W(20), DT(8), P(70, 2), CC(10, 2) As Single, TR(70, 2), CCT(10, 2), HYP(200), ANGL(200) As Single, I6(100), J6(100), D$(20), Z$(30), DL(15)
'Dim C(15) As Single, L(15), LR(15), T(15), S(15), W(20), DT(8), P(70, 2) As Single, TR(70, 2), CCT(10, 2), HYP(200), ANGL(200) As Single, I6(100), J6(100), D$(20), Z$(30), DL(15)
Dim iI6, iJ6, iW

'These have now been declared as module level variables
'COMMON SHARED FLAG6, BDSHEET, C() AS SINGLE, L(), LR(), T(), S(), W(), DT(), P(), CC() AS SINGLE, TR(), CCT(), HYP(), ANGL() AS SINGLE, I6(), J6(), D$(), Z$(), DL(), M$
'COMMON SHARED B$, N$, OD#, sOD$, HAND$, SUBFLAG, TYPES, PNO, sP$, NBR, PD$, CAT$, OZ$, LOCNO, ZL$, RsP$, SLI$, BOTH, N, RQTY, LQTY, INSERT, A2$, sINSERT$, SW, OPENTIPS, STDTIPS, F1, F2, F3, F4
'COMMON SHARED LONGEST, XBASE, YBASE, INDEX, MIDDLE, RING, LITTLE, THUMB, WRISTL, XTI, YTI, X10, BETA, DIPI, DIPM, DIPR, DIPL, DESIGNER, sDESIGNER$, FLAG4, FLAG5, FLAG7, FLAG8, FLAG9
'COMMON SHARED SEMICIRC, QUARTCIRC1, QUARTCIRC2, CONST1, CONST2, CONST3, YC, XR10, XRTI, JKR9, JK9, JK8, JK7, JUMPFLAG, NOPLOT, PI, MISSING, GA3, GA4, TBA2, GA10, XFLAG
'COMMON SHARED TAPES, PLOTONCE, LOCY, CONST6, MINW, FBASE, SPACE, FINGER, SCREENPLOT, K, YC7, GA5, FLAG100, PLOTPALM, PG, ZIPLEN, sZIPLEN$, GA8, LTW, GA7, OZ2$, XXC, GRAM$
'COMMON SHARED H21, A21, R21, TA21, TA22, T21, TIRATIO, LTW2, ITDIF
''C-GEM
'COMMON SHARED JOBST$

'ENTRY POINT FOR CHAINING AND PROGRAM START POINT
                
        If FLAG4 = 1 Then GoTo 350
        If FLAG5 = 1 Then GoTo 300
        If JUMPFLAG = 2 Then GoTo 54

3
'5 MIDSECT = 1000:CHAIN MERGE "A:MIDSECT",1000,ALL
'6 ENDSECT = 1000:CHAIN MERGE "A:ENDSECT",4000,ALL
7 MIDSECT = 1: ENDSECT = 1
8 FLAG6 = 0: DESIGNER = 0: BDSHEET = 0
9 CLEAR
10 PI = 3.14159263589793: MISSING = 100

14 B$ = "                                                                     "
16 COLOR 15, 9, 11:
'C-GEM
'Input and Output are local
18 Open "C:\JOBST\GLOVDATA.DAT" For Input As #2
   Open "C:\JOBST\GLOVDATA.TMP" For Output As #3
   
'C-GEM
'Get the path to the EXECUTEABLES for use in chain command.
'Allows for use of a networked drive.
   Input #2, JOBST$
20 LOCATE 8, 3: Print #3, "What is this Customer's name ";: Input #2, sn$
22 Print #3,: Print #3, "  What is this Customer's File Number ";: Input #2, OD#: sOD$ = Str$(OD#)
24 If SUBFLAG = 8500 Then GoTo 58
26 LOCATE 12, 3: Print #3, "How many product types are in this order ";: Input #2, TYPES
28 LOCATE 14, 3: Print #3, "Enter the appropriate number from the following list of Priority Statuses."
30 LOCATE 16, 15: Print #3, " (1) RUSH"
32 LOCATE 17, 15: Print #3, " (2) BURN"
34 LOCATE 18, 15: Print #3, " (3) HOTBURN"
36 LOCATE 19, 15: Print #3, " (4) HOTLINE"
38 LOCATE 20, 15: Print #3, " (5) BLUELINE"
40 LOCATE 22, 17: Input #2, PNO
42 If PNO = 1 Then sP$ = "RUSH": GoTo 51
44 If PNO = 2 Then sP$ = "BURN": GoTo 51
46 If PNO = 3 Then sP$ = "HOTBURN": GoTo 51
48 If PNO = 4 Then sP$ = "HOTLINE": GoTo 51
50 If PNO = 5 Then sP$ = "BLUELINE"
51 LOCATE 6, 3: Print #3, "Does "; sn$; " want gloves for both hands ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then BOTH = 2 Else GoTo 53
52 Print #3,: Print #3, "  Start with "; sn$; "'s Right hand.": HAND$ = "RIGHT": GoTo 57
53 Print #3,: Print #3, "  Is this for the Customer's Right or Left hand ";: Input #2, HAND$
54 If Left$(HAND$, 1) = "R" Or Left$(HAND$, 1) = "r" Then HAND$ = "RIGHT"
55 If Left$(HAND$, 1) = "L" Or Left$(HAND$, 1) = "l" Then HAND$ = "LEFT"
57 PD$ = "GLOVE TO WRIST": CAT$ = "110100"
   If JUMPFLAG = 2 Then CLS
58 LOCATE 10, 3: Print #3, "Please answer Y or N for each Option :"
60 LOCATE 12, 20: Print #3, "ZIPPERS ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then OZ$ = "ZIPPER": LOCY = 9 Else OZ$ = "": GoTo 74
62 LOCATE 13, 31: Print #3, "(1) Ulnar"
63 LOCATE 14, 13: Print #3, "Select the "; Tab(31); "(2) Palmar"
64 LOCATE 15, 9: Print #3, "desired location :"; Tab(31); "(3) Dorsal"


65 LOCATE 16, 31: Print #3, "(4) Outside"
   LOCATE 17, 31: Print #3, "(5) Other "

66 LOCATE 19, 31: Print #3, "";: Input #2, LOCNO
68 If LOCNO = 1 Then ZL$ = "ULNAR": GoTo 73
69 If LOCNO = 2 Then ZL$ = "PALMAR": GoTo 73
70 If LOCNO = 3 Then ZL$ = "DORSAL": GoTo 73
72 If LOCNO = 4 Then ZL$ = "OUTSIDE": GoTo 73
   If LOCNO = 5 Then ZL$ = "OTHER"
73 LOCATE 20, 20: Print #3, "What length zipper ";: Input #2, ZIPLEN
74 LOCATE 13 + LOCY, 20: Print #3, "SLANT INSERTS ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then SLI$ = "SLANT INSERTS" Else SLI$ = ""
   LOCATE 14 + LOCY, 20: Print #3, "REINFORCEMENTS";: Input #2, A$
   If Asc(A$) = 89 Or Asc(A$) = 121 Then
     LOCATE 16 + LOCY, 20: Print #3, "WHAT AREA DO THEY WANT REINFORCED ";
     Input #2, RsP$
    
   Else
     RsP$ = ""
  
   End If


79      If JUMPFLAG <> 2 Then CLS
        
        JUMPFLAG = 0
        LOCATE 14, 3: Print #3, "How many fingers,not counting the thumb, will this glove have ";: Input #2, n

80 If n > 4 Then Beep: LOCATE 14, 65: Print #3, "        ": GoTo 79
81 If n = 0 Then Beep: LOCATE 20, 4: Print #3, "  This program is not able to accomodate fingerless hands,only missing fingers.   Please reenter.": GoTo 79
82 If BOTH = 2 And HAND$ = "LEFT" Then GoTo 106
84   LOCATE 7, 10: Print #3, "Patient's name : "; sn$: LOCATE 9, 10: Print #3, "File Number : "; Left$(sOD$, 8); "-"; Right$(sOD$, 1)
86 LOCATE 11, 10: Print #3, "Right or Left : "; HAND$
88 If NAMCHECK = 1000 And ODCHECK = 1000 Then GoTo 100
89 If NAMCHECK = 1000 And ODCHECK = 0 Then GoTo 96
90 LOCATE 16, 10: Print #3, "Is the name correct (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then NAMCHECK = 1000: GoTo 93
91 LOCATE 16, 6: Print #3, B$
92 LOCATE 16, 10: Print #3, "Retype the correct name.";: Input #2, sn$
93 LOCATE 16, 6: Print #3, B$: GoTo 84
94 If ODCHECK = 1000 And NAMCHECK = 0 Then GoTo 84
95 If ODCHECK = 1000 And NAMCHECK = 1000 Then GoTo 100
96 LOCATE 16, 10: Print #3, "Is the File Number correct (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then ODCHECK = 1000: GoTo 99
97 LOCATE 16, 6: Print #3, B$
98 LOCATE 16, 10: Print #3, "Retype the correct file number.";: Input #2, OD#: sOD$ = Str$(OD#)
99 If NAMCHECK = 1000 And ODCHECK = 1000 Then GoTo 100 Else GoTo 82
100 LOCATE 16, 10: Print #3, "Is "; HAND$; " the correct hand for this glove (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then HDCHECK = 1000: GoTo 106
101 LOCATE 16, 6: Print #3, B$
102 LOCATE 16, 10: Print #3, "Retype the correct hand ";: Input #2, HAND$
104 LOCATE 16, 6: Print #3, B$: LOCATE 11, 6: Print #3, B$: GoTo 84
106 'DEF FNCONVIE(X) = ((X / 10) - Int(X / 10)) / 0.8 + Int(X / 10)
107 If HAND$ = "LEFT" Then GoTo 108 Else Print #3,: Print #3, "         How many RIGHT gloves do you want to make ";: Input #2, RQTY: GoTo 109
108 Print #3,: Print #3, "         How many LEFT gloves do you want to make ";: Input #2, LQTY


109     CLS
        LOCATE 10, 5
        Print #3, "Is this order for a Lymphadema diagnosis (Y or N)";
        Input #2, A$

        If Asc(A$) = 89 Or Asc(A$) = 121 Then
           GoTo SKIP
        Else: GoTo QUERRY
       
        End If


QUERRY: CLS
        LOCATE 10, 5
        Print #3, "Will you be using A or B gram range fabric (A or B)";
        Input #2, GRAM$
       
        If (Left$(GRAM$, 1) = "B" Or Left$(GRAM$, 1) = "b") Then
          GRAM$ = "B"
       
        Else: GRAM$ = "A"
       
        End If
       
SKIP:   CLS
        LOCATE 10, 5
        Print #3, "Will you be using English or Metric units (E or M)";
        Input #2, M$
       
        M$ = Left$(M$, 1)
       
        If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 112
        
110   Print #3,: Print #3,: Print #3,: Print #3,: Print #3, "   Now enter the finger circumferences and lengths, 1 thru 20.": GoTo 116
112   LOCATE 10, 4: Print #3, "Please enter all circumference and length measurements "
113 LOCATE 11, 4: Print #3, "in Centimeters (ie. 10.5 cm).": Print #3,
116 Print #3, "   Press ENTER after each input. If a circumference is not given,"
117 Print #3, "   just type 0 and then ENTER."
118 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 119 Else GoTo 120
119 Print #3,: Print #3,: For I = 1 To 11: Print #3, Tab(20); "   C("; I; ") = ";: Input #2, C(I): Next I: GoTo 130
120 Print #3,: Print #3,: For I = 1 To 11: Print #3, Tab(20); "   C("; I; ") = ";: Input #2, C(I): C(I) = FNCONVIE(C(I)): Next I
130 Print #3,: Print #3, Tab(10); "Now enter the length dimensions.": Print #3,
140 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 145 Else GoTo 160
145 For I = 1 To 9: Print #3, Tab(20); "   L("; I + 11; ") = ";: Input #2, L(I): Next I: GoTo 162
160 For I = 1 To 9: Print #3, Tab(20); "   L("; I + 11; ") = ";: Input #2, L(I): L(I) = FNCONVIE(L(I)): Next I
    ZIPLEN = FNCONVIE(ZIPLEN)
162   LOCATE 10, 5: Print #3, "    Will this glove extend beyond the wrist (Y/N)";: Input #2, A$: If Asc(A$) = 78 Or Asc(A$) = 110 Then GoTo 172
163 LOCATE 10, 5: Print #3, B$
164 LOCATE 10, 5: Print #3, "    How many tapes, 1 or 2 ";: Input #2, TAPES
165 If TAPES > 2 Then Beep: Print #3, "      Bad INPUT #2, "
166 LOCATE 10, 5: Print #3, B$
167 If TAPES > 2 Then GoTo 164
169 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 170 Else GoTo 172
170 For I = 1 To TAPES: Print #3, "    Enter the circumference for the "; 1.5 * 2.54 * I; " tape.";: Input #2, C(11 + I)
171 L(9 + I) = 1.5 * 2.54 * I: Print #3,: Next I: GoTo 178
172 For I = 1 To TAPES: Print #3, "    Enter the circumference for the "; 1.5 * I; " tape.";: Input #2, C(11 + I): C(11 + I) = FNCONVIE(C(11 + I))
174 L(9 + I) = 1.5 * I: Print #3,: Next I
178   Print #3, "  Please check that the input values are correct.": Print #3,
179 'IF SCREENPLOT>0  THEN SUBFLAG=178
180 Print #3, Tab(13); "Circumferences"; Tab(36); "Lengths"
190 For I = 1 To 11: Print #3, Tab(15); "C"; I; Tab(20); C(I); Tab(35); "L"; I + 11; Tab(40); L(I): Next I
192 Print #3,: Print #3, Tab(15); "C"; 12; Tab(20); C(12); " ( 1 1/2 (3.81 cm) Tape )"
194 Print #3, Tab(15); "C"; 13; Tab(20); C(13); "   ( 3 (7.62 cm) Tape )"
196 If L(9) / L(8) > 1.8 Then
     LOCATE 17, 5: Print #3, " The Index to Thumb ratio is unusualy large. Please check the hand outline "
     LOCATE 18, 5: Print #3, "for finger webs or for a bad thumb height."
    End If
200 LOCATE 20, 5: Print #3, " Is everything correct (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 295
205 BADDATA = 1
210 LOCATE 20, 5: Print #3, " Are the circumferences correct (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 250
215 BADDATA = 2
220 LOCATE 20, 5: Print #3, B$
222 LOCATE 20, 5: Print #3, " How many circumferences are wrong ";: Input #2, iW
224 For J = 1 To iW
226 LOCATE 20, 5: Print #3, " Type in the number of a wrong circumference.  ";: Input #2, I
228 LOCATE 20, 5: Print #3, B$
229 If I > 13 Then GoTo 226
230 LOCATE 20, 5: Print #3, " Please retype the correct value for C"; I;: Input #2, C(I)
231 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 232 Else C(I) = FNCONVIE(C(I))
232 If I = 12 Then LOCATE 4 + I, 20: Print #3, "          ": TAPES = 1: GoTo 235
233 If I = 13 Then LOCATE 4 + I, 20: Print #3, "          ": TAPES = 2: GoTo 235
234 LOCATE 3 + I, 20: Print #3, "          "
235 If Asc(M$) = 77 Or Asc(M$) = 109 Then CONST6 = 2.54 Else CONST6 = 1
237 If I = 12 Then L(10) = 1.5 * CONST6: LOCATE 4 + I, 20: COLOR 12, 9, 11: Print #3, C(I): LOCATE 13, 40: Print #3, L(10): COLOR 15, 9, 11: GoTo 240
238 If I = 13 Then L(11) = 3! * CONST6: LOCATE 4 + I, 20: COLOR 12, 9, 11: Print #3, C(I): LOCATE 14, 40: Print #3, L(11): COLOR 15, 9, 11: GoTo 240
239 LOCATE 3 + I, 20: COLOR 12, 9, 11: Print #3, C(I): COLOR 15, 9, 11
240 LOCATE 20, 5: Print #3, B$
241 Next J
242 LOCATE 20, 5: Print #3, " Are there any other circumference corrections ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 220
245   Print #3,: Print #3,: Print #3, Tab(13); "Circumferences"; Tab(36); "Lengths"
246 For I = 1 To 10: Print #3, Tab(15); "C"; I; Tab(20); C(I); Tab(35); "L"; I + 11; Tab(40); L(I): Next I
247 If C(12) <> 0 Then Print #3, Tab(35); "L"; 22; Tab(40); L(11)
250 If BADDATA = 1 Then GoTo 270
255 LOCATE 20, 5: Print #3, " Are the lengths correct (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 295
270 LOCATE 20, 5: Print #3, B$
272 LOCATE 20, 5: Print #3, " How many lengths are wrong ";: Input #2, iW
273 For J = 1 To iW
274 LOCATE 20, 5: Print #3, " Type in the number of an incorrect length.  ";: Input #2, I: I = I - 11
280 LOCATE 20, 5: Print #3, B$
282 LOCATE 20, 5: Print #3, " Please retype the correct value for L"; I + 11;: Input #2, L(I)
283 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 284 Else L(I) = FNCONVIE(L(I))
284 LOCATE 3 + I, 40: Print #3, "          "
285 LOCATE 3 + I, 40: COLOR 12, 9, 11: Print #3, L(I): COLOR 15, 9, 11
290 LOCATE 20, 5: Print #3, B$
291 Next J
292 LOCATE 20, 5: Print #3, " Are there any other length corrections ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 270
293 For I = 1 To TAPES
294 If C(11 + I) = 0 Then TAPES = 0
       
        Next I
295 GoSub 1300
299 GoTo 380

300 ' SUB TO DESIGNATE AN INSERT
310   COLOR 15, 9, 11
311 SUBFLAG = 300
312 If Asc(M$) = 77 Or Asc(M$) = 109 Then A2$ = "cm.": INS = INSERT Else A2$ = "eighths": INS = 8 * INSERT
313 LOCATE 12, 5: Print #3, "The current INSERT width is "; INS; A2$; "."
315 LOCATE 10, 5: Print #3, "What size Insert ( "; A2$; " ) do you want to try ";: Input #2, INSERT
317 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 319
318 INSERT = FNCONVIE(INSERT): sINSERT$ = Str$(INSERT) + Chr$(34) + " INSERT": GoTo 320
319 ssINSERT$ = Str$(INSERT) + " cm" + " INSERT"
320 If SUBFLAG = 300 And (Left$(HAND$, 1) = "R" Or Left$(HAND$, 1) = "r") Then GoSub 8000  'REINSTATES SIGNS FOR REPLOTS
        FLAG5 = 0
325 GoTo 452

350 'SUB TO REINSTATE SIGNS FOR DATA CORRECTION
352 SUBFLAG = 178
354 If SUBFLAG = 178 And (Left$(HAND$, 1) = "R" Or Left$(HAND$, 1) = "r") Then GoSub 8000   'REINSTATES SIGNS FOR REPLOTS
        FLAG4 = 0
360 GoTo 178

380 For I = 1 To 7 Step 2: T(I) = 1.08 * 30 * C(I): Next I
390 For I = 2 To 8 Step 2: T(I) = 1.08 * 25 * C(I): Next I
400 T(9) = 1.08 * 25 * C(9): T(10) = 1.08 * 15 * C(10): T(11) = 1.08 * 15 * C(11)

402   If (Left$(GRAM$, 1) = "B" Or Left$(GRAM$, 1) = "b") Then

          For I = 1 To 7 Step 2: S(I) = 0.155: Next I
          For I = 2 To 8 Step 2: S(I) = 0.13: Next I
          S(9) = 0.13: S(10) = 0.105: S(11) = 0.1: S(12) = 0.1: S(13) = 0.1
          GoTo 410
       
      End If


406 For I = 1 To 7 Step 2: S(I) = 0.18: Next I
407 For I = 2 To 8 Step 2: S(I) = 0.15: Next I
408 S(9) = 0.15: S(10) = 0.12: S(11) = 0.1: S(12) = 0.1: S(13) = 0.1
410 'FOR I=1 TO 11 : S(I)=(28-(GRAMS-T(I))/MODULUS)/100:NEXT I
425 SW = 0.125: DTSUM = 0
427 If Asc(M$) = 77 Or Asc(M$) = 109 Then SW = 0.3175
430 For I = 1 To 8
432   If C(I) = O Then GoTo 438
434   DT(I) = (C(I) - 4 * SW) / (4 * (1 + S(I))) + 2 * SW: DTSUM = DTSUM + DT(I): DTCOUNT = DTCOUNT + 1
438 Next I
439 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 445
440 INSERT = DTSUM / DTCOUNT: INSERT = Int((INSERT / 0.125) + 0.5) * 0.125: sINSERT$ = Str$(INSERT) + Chr$(34) + " INSERT": DTCOUNT = 0
442 GoTo 450
445 INSERT = DTSUM / DTCOUNT: INSERT = CInt(INSERT * 1000) / 1000: sINSERT$ = Str$(INSERT) + " cm" + " INSERT": DTCOUNT = 0
450 If Asc(M$) = 77 Or Asc(M$) = 109 Then MINW = 0.9525 Else MINW = 0.375
452 For I = 1 To 8
453   W(I) = (C(I) - 4 * SW) / (2 * (1 + S(I))) - INSERT + 4 * SW
454   If W(I) < MINW Then W(I) = MINW
456   If C(I) = 0 Then W(I) = 0
458 Next I
460 W(16) = (0.88 * C(1) - 4 * SW) / (2 * (1 + S(1))) - INSERT + 4 * SW: If W(16) < MINW Then W(16) = MINW
462 W(17) = (0.88 * C(3) - 4 * SW) / (2 * (1 + S(3))) - INSERT + 4 * SW: If W(17) < MINW Then W(17) = MINW
464 W(18) = (0.88 * C(5) - 4 * SW) / (2 * (1 + S(5))) - INSERT + 4 * SW: If W(18) < MINW Then W(18) = MINW
466 W(19) = (0.88 * C(7) - 4 * SW) / (2 * (1 + S(7))) - INSERT + 4 * SW: If W(19) < MINW Then W(19) = MINW
467 For I = 2 To 8 Step 2: FBASE = FBASE + W(I): Next I
468 W(10) = (C(10) - 4 * SW) / (2 * (1 + S(10))) - INSERT + 4 * SW
469 W(11) = (C(11) - 4 * SW) / (2 * (1 + S(11))) - INSERT + 4 * SW
470 For I = 1 To TAPES: W(11 + I) = (C(11 + I) - 4 * SW) / (2 * (1 + S(11 + I))) - INSERT + 4 * SW: Next I
473 If n = 4 Then GoTo 480
474 SPACE = (W(10) - 2 * SW - FBASE) / (4 - n)
476 For I = 2 To 8 Step 2
477   If W(I) = 0 Then W(I) = SPACE
479 Next I
480 W(9) = (C(9) - 3 * SW) / (2 * (1 + S(9))) - INSERT / 2 + 3 * SW
492 'GOSUB 1300 (See 295)
494 If OpenTips = 0 Then GoTo 500
496 For I = 6 To 9: LR(I) = L(I) / 1.07: Next I
497 For I = 1 To 5
        If STDTIPS = 1000 Then LR(I) = L(I) / 1.1
    Next I

498 GoTo 501
500 For I = 1 To 5: LR(I) = L(I) / 1.1: Next I: For I = 6 To 9: LR(I) = L(I) / 1.07: Next I
501 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 504
502 For I = 1 To TAPES: LR(11 + I) = (1.5 * I) / 1.07: Next I
503 GoTo 505
504 For I = 1 To TAPES: LR(11 + I) = (3.81 * I) / 1.07: Next I
505 If LR(4) = 0 Then F1 = MISSING: FINGER = FINGER + 1
506 If LR(3) = 0 Then F2 = MISSING: FINGER = FINGER + 1
507 If LR(2) = 0 Then F3 = MISSING: FINGER = FINGER + 1
508 If LR(1) = 0 Then F4 = MISSING: FINGER = FINGER + 1
509 For I = 1 To 4
510  If LR(I) >= LONGEST Then LONGEST = LR(I)
511 Next I
520 TIRATIO = LR(8) / LR(9)
    If TIRATIO < 1.5 Then LR(9) = LR(8) / 1.6: GoTo 590
    If TIRATIO > 1.8 Then WEB = 1000

590 T1 = 1 - ((W(6) / 2 + W(8)) / LR(9)): TA1 = PI / 2 - Atn(T1 / Sqr(1 - T1 ^ 2))
630 T2 = ((LR(8) - LR(9)) / 4 + LR(9) - (W(6) / 2 + W(8))) / 2
632 If LR(9) / 2 <= W(11) / 2 Then GoTo 660
634 TB1 = Sqr((LR(9) / 2) ^ 2 - (LR(9) / 2 - W(11) / 2) ^ 2): TB2 = 1 - (W(11) / LR(9))
636 XBASE = -W(11) / 2: YBASE = LR(9) / 2 - TB1
638 TBA2 = Atn(TB2 / Sqr(1 - TB2 ^ 2))
660  G1 = Atn((W(6) / 2 + W(4) + W(2) - W(11) / 2) / LR(6)): L1 = LR(6) / Cos(G1)
670 C1 = Sqr((W(11) / 2) ^ 2 + (L1 / 2) ^ 2 - 2 * (W(11) / 2) * (L1 / 2) * Cos(PI / 2 + G1))
680 C2 = ((L1 / 2) * Sin(PI / 2 + G1)) / C1: Phi = Atn(C2 / Sqr(1 - C2 ^ 2))
690 A1 = PI / 2 - Phi: A2 = PI / 2 - G1: A3 = PI / 2 + G1: A4 = PI / 4 - G1
700 H2 = C1 * Sin(3 * PI / 4 - Phi) / Sin(A4): R2 = C1 * Sin(G1 + Phi) / Sin(PI / 4 - G1)
710 RP = Sqr(H2 ^ 2 + (L1 / 2) ^ 2): RP2 = Sqr((H2 + RP) ^ 2 + (L1 / 2) ^ 2)
715 G2 = ((LR(6) - R2 * Sin(PI / 4)) / RP): GA2 = Atn(G2 / Sqr(1 - G2 ^ 2)): GA3 = Atn((L1 / 2) / (H2 + RP))
735 If OpenTips = 1000 Then GoSub 1500
736 If STDTIPS = 1000 Then GoTo 780
737 If INDEX = 1000 Then GoTo 750
740 P(1, 1) = -(W(8) + W(6) / 2): P(1, 2) = LR(8)
741 If LR(4) = 0 Then GoTo 750
742 P(2, 1) = -(W(8) + W(6) / 2): P(2, 2) = LR(8) + 0.4 * LR(4)
744 P(3, 1) = -((W(8) + W(7) + W(6)) / 2): P(3, 2) = P(2, 2) + 0.3 * LR(4)
745 P(4, 1) = -((W(6) + W(8) + W(19)) / 2): P(4, 2) = P(3, 2) + 0.3 * LR(4) - W(19) / 2
746 P(5, 1) = -((W(8) + W(6) - W(19)) / 2): P(5, 2) = P(4, 2)
747 P(6, 1) = -((W(8) + W(6) - W(7)) / 2): P(6, 2) = P(3, 2)
748 P(7, 1) = -(W(6) / 2): P(7, 2) = P(2, 2)
750 P(8, 1) = -(W(6) / 2): P(8, 2) = LR(8)
751 If LR(3) = 0 Or MIDDLE = 1000 Then GoTo 760
752 P(9, 1) = -(W(6) / 2): P(9, 2) = LR(8) + 0.4 * LR(3)
754 P(10, 1) = -(W(5) / 2): P(10, 2) = P(9, 2) + 0.3 * LR(3)
755 P(11, 1) = -(W(18) / 2): P(11, 2) = P(10, 2) + 0.3 * LR(3) - W(18) / 2
756 P(12, 1) = W(18) / 2: P(12, 2) = P(11, 2)
757 P(13, 1) = W(5) / 2: P(13, 2) = P(10, 2)
758 P(14, 1) = W(6) / 2: P(14, 2) = LR(8) + 0.4 * LR(3)
760 P(15, 1) = W(6) / 2: P(15, 2) = LR(7)
761 If LR(2) = 0 Or RING = 1000 Then GoTo 769
762 P(16, 1) = W(6) / 2: P(16, 2) = LR(7) + 0.4 * LR(2)
764 P(17, 1) = (W(6) + W(4) - W(3)) / 2: P(17, 2) = P(16, 2) + 0.3 * LR(2)
765 P(18, 1) = ((W(6) + W(4) - W(17)) / 2): P(18, 2) = P(17, 2) + 0.3 * LR(2) - W(17) / 2
766 P(19, 1) = (W(6) + W(4) + W(17)) / 2: P(19, 2) = P(18, 2)
767 P(20, 1) = (W(6) + W(4) + W(3)) / 2: P(20, 2) = P(17, 2)
768 P(21, 1) = W(6) / 2 + W(4): P(21, 2) = LR(7) + 0.4 * LR(2)
769 P(22, 1) = W(6) / 2 + W(4): P(22, 2) = LR(6)
770 If LR(1) = 0 Or LITTLE = 1000 Then GoTo 778
771 P(23, 1) = W(6) / 2 + W(4): P(23, 2) = LR(6) + 0.4 * LR(1)
772 P(24, 1) = (W(6) + W(2) - W(1)) / 2 + W(4): P(24, 2) = P(23, 2) + 0.3 * LR(1)
773 P(25, 1) = ((W(6) + W(2) - W(16)) / 2) + W(4): P(25, 2) = P(24, 2) + 0.3 * LR(1) - W(16) / 2
774 P(26, 1) = (W(6) + W(2) + W(16)) / 2 + W(4): P(26, 2) = P(25, 2)
775 P(27, 1) = (W(6) + W(2) + W(1)) / 2 + W(4): P(27, 2) = P(24, 2)
776 P(28, 1) = W(6) / 2 + W(2) + W(4): P(28, 2) = LR(6) + 0.4 * LR(1)
778 P(29, 1) = W(6) / 2 + W(2) + W(4): P(29, 2) = LR(6)
780 P(30, 1) = W(11) / 2: P(30, 2) = 0
782 P(31, 1) = -(W(11) / 2): P(31, 2) = 0
784 P(32, 1) = -(W(6) / 2 + W(8) + T2 - (W(9) / 2) * Cos(PI / 4)): P(32, 2) = (LR(8) - LR(9)) / 4 + LR(9) - T2 + (W(9) / 2) * Sin(PI / 4)
785 If LR(5) = 0 Or THUMB = 1000 Or STDTIPS = 1000 Then GoTo 792
786 P(33, 1) = P(32, 1) - (LR(5) / 2) * Cos(PI / 4): P(33, 2) = P(32, 2) + (LR(5) / 2) * Sin(PI / 4)
788 P(34, 1) = P(33, 1) - (2 * W(9) / 3 + LR(5) / 2) * Cos(PI / 4): P(34, 2) = P(33, 2) - (2 * W(9) / 3 - LR(5) / 2) * Sin(PI / 4)
790 P(35, 1) = P(32, 1) - W(9) * Cos(PI / 4) - (LR(5) - W(9) / 3) * Cos(PI / 4): P(35, 2) = P(32, 2) - W(9) * Sin(PI / 4) + (LR(5) - W(9) / 3) * Sin(PI / 4)
792 P(36, 1) = P(32, 1) - W(9) * Cos(PI / 4): P(36, 2) = P(32, 2) - W(9) * Sin(PI / 4)
794 P(37, 1) = -(W(8) + W(6) / 2): P(37, 2) = P(32, 2)

796 K1 = ((W(11) / 2 + LR(9) / 2) - (W(8) + W(6) / 2)) / LR(9): IA2 = Atn(K1 / Sqr(1 - K1 ^ 2))
    H21 = 0.77 * LR(8) - LR(9) / 2
    A21 = Atn(Abs(H21 / P(1, 1)))
    R21 = H21 / Sin(2 * (PI / 2 - A21))
    T21 = INSERT / (2 * R21)
    TA21 = (T21 / Sqr(1 - T21 ^ 2))
798 K2 = (INSERT / 2) / R21
    TA22 = Atn((R21 - Abs(P(1, 1))) / H21)
    'K2 = (INSERT / 2) / LR(9)
    IA3 = PI / 2 - Atn(K2 / Sqr(1 - K2 ^ 2))
    'IA4 = IA3 + TA1 - PI / 2
    IA4 = IA3 - TA22

800  P(41, 1) = P(37, 1): P(41, 2) = LR(9) / 2 + H21
    'P(41, 1) = P(37, 1): P(41, 2) = LR(9) * (SIN(TA1) + .5)

805 P(38, 1) = -(W(8) + W(6) / 2 - INSERT * Sin(IA4)): P(38, 2) = P(41, 2) - INSERT * Cos(IA4)
806 GA7 = Abs(Atn((P(31, 1) - P(38, 1)) / P(38, 2)))
807 GA5 = Abs(Atn((P(37, 1) - P(38, 1)) / P(38, 2)))
809 GA6 = Abs(Atn((P(38, 1) - P(37, 1)) / (P(38, 2) - P(37, 2))))
810 If Asc(M$) = 77 Or Asc(M$) = 109 Then WRISTL = -1.5875: XW = 0.0792 Else WRISTL = -0.625: XW = 0.0312
811 P(39, 1) = -(W(11) / 2 + XW): P(39, 2) = WRISTL
812 P(40, 1) = W(11) / 2 + XW: P(40, 2) = WRISTL
814 GA4 = PI / 2 - TA22 - 2 * TA21
    'GA4 = TA1 - 2 * (PI / 2 - IA3)

820 CC(1, 1) = -(LR(9) / 2): CC(1, 2) = LR(9) / 2
822 CC(2, 1) = -R21: CC(2, 2) = LR(9) / 2
     'CC(2, 1) = -LR(9): CC(2, 2) = LR(9) / 2

823 If LR(4) = 0 Or STDTIPS = 1000 Or INDEX = 1000 Then GoTo 825
824 CC(3, 1) = -(W(6) + W(8)) / 2: CC(3, 2) = LR(8) + LR(4) - W(19) / 2
825 If LR(3) = 0 Or STDTIPS = 1000 Or MIDDLE = 1000 Then GoTo 827
826 CC(4, 1) = 0: CC(4, 2) = LR(8) + LR(3) - W(18) / 2
827 If LR(2) = 0 Or STDTIPS = 1000 Or RING = 1000 Then GoTo 829
828 CC(5, 1) = (W(6) + W(4)) / 2: CC(5, 2) = LR(7) + LR(2) - W(17) / 2
829 If LR(1) = 0 Or STDTIPS = 1000 Or LITTLE = 1000 Then GoTo 845
830 CC(6, 1) = (W(6) + W(2)) / 2 + W(4): CC(6, 2) = LR(6) + LR(1) - W(16) / 2
845 CC(7, 1) = -R2 * Cos(PI / 4) - RP * Sin(PI / 2 - G1): CC(7, 2) = R2 * Sin(PI / 4) + RP * Cos(PI / 2 - G1)
850 CC(8, 1) = P(33, 1) - (2 * W(9) / 3) * Cos(PI / 4): CC(8, 2) = P(33, 2) - (2 * W(9) / 3) * Sin(PI / 4)
855 CC(9, 1) = CC(8, 1) - (LR(5) / 2 - W(9) / 3) * Cos(PI / 4): CC(9, 2) = CC(8, 2) + (LR(5) / 2 - W(9) / 3) * Sin(PI / 4)
870 P(42, 1) = P(31, 1): P(42, 2) = P(31, 2) - SW
880 P(43, 1) = P(38, 1) + SW * Cos(GA4): P(43, 2) = P(38, 2) + SW * Sin(GA4)
890 P(44, 1) = P(37, 1): P(44, 2) = P(37, 2) + SW / Sin(GA6)
900 P(45, 1) = P(32, 1) - SW: P(45, 2) = P(32, 2) + SW
910 P(46, 1) = -LR(9) / 2: P(46, 2) = 0
915 P(47, 1) = W(10) / 2: P(47, 2) = LR(9)
920 P(52, 1) = W(12) / 2: P(52, 2) = -LR(12)
925 P(54, 1) = W(13) / 2: P(54, 2) = -LR(13)
930 P(55, 1) = -W(13) / 2: P(55, 2) = -LR(13)
935 P(53, 1) = -W(12) / 2: P(53, 2) = -LR(12)
950 H9 = CC(2, 1): R9 = Abs(CC(2, 1)) + SW: K9 = CC(2, 2)
952 M1 = Tan(PI / 2 - GA6): Y9 = P(37, 2) + SW / Cos(PI / 2 - GA6): X9 = P(37, 1)
954 B1 = Y9 - (X9 * M1): CZ1 = (2 * M1 * B1 - 2 * H9 - 2 * K9 * M1) / (1 + M1 ^ 2): CZ2 = (H9 ^ 2 + B1 ^ 2 - 2 * B1 * K9 + K9 ^ 2 - R9 ^ 2) / (1 + M1 ^ 2)
956 XX = (-CZ1 + Sqr(CZ1 ^ 2 - 4 * CZ2)) / 2: YY = M1 * XX + B1: XX1 = (-CZ1 - Sqr(CZ1 ^ 2 - 4 * CZ2)) / 2: YY1 = XX1 * M1 + B1
958 P(50, 1) = XX: P(50, 2) = YY
960 K10 = (P(50, 2) - K9) / R9: GA10 = Atn(K10 / Sqr(1 - K10 ^ 2))
962 X10 = P(37, 1) + SW / Cos(GA6)
964 M2 = Tan(3 * PI / 4): B2 = P(33, 2) - M2 * P(33, 1)
966 XTI = (B2 - B1) / (M1 - M2): YTI = M2 * XTI + B2
968 If XTI < X10 Then XFLAG = 1000 '***************************
'1000 IF MIDSECT = 1000 THEN GOTO 6
1001 GoTo 1260 'SKIP Screen drawing
1002 SF = 4 / 3
1005 If Asc(M$) = 77 Or Asc(M$) = 109 Then VIEW "(0, 0)-(639, 199)": WINDOW "(-10, -5)-(10, 20)": SF = 4 / 3: VASPECT = 1 / 2: HASPECT = 1 / 4: GoTo 1020
1010 VIEW "(0, 0)-(639, 199)": WINDOW "(-5, -2)-(5, 8)": VASPECT = 1 / 3: HASPECT = 1 / 3
1020 For I = 1 To 4
1021    If STDTIPS = 1000 Then I = 4: GoTo 1028
1022    If I = 1 And INDEX = 1000 Then GoTo 1028
1023    If I = 2 And MIDDLE = 1000 Then GoTo 1028
1024    If I = 3 And RING = 1000 Then GoTo 1028
1025    If I = 4 And LITTLE = 1000 Then GoTo 1028
1026    If CC(I + 2, 1) = 0 And CC(I + 2, 2) = 0 Then Count = Count + 1: GoTo 1028
1027    'Circle (CC(I + 2, 1), CC(I + 2, 2)), W(20 - I) / 2, , 0, PI, VASPECT
1028 Next I
1029 BETA = 7 * PI / 4 + A4 + GA3: If BETA > 2 * PI Then BETA = BETA - 2 * PI
1030 'Circle (CC(1, 1), CC(1, 2)), LR(9) / 2, , 3 * PI / 2, 2 * PI, HASPECT
1032 'Circle (CC(2, 1), CC(2, 2)), R21, , 0, (2 * (PI / 2 - A21)), HASPECT
     'CIRCLE (CC(2, 1), CC(2, 2)), LR(9), , 0, TA1, HASPECT

1034 'Circle (CC(7, 1), CC(7, 2)), RP2, , 7 * PI / 4 + A4 - GA3, BETA, HASPECT
1035 If LR(5) = 0 Or STDTIPS = 1000 Or THUMB = 1000 Then GoTo 1038
1036 'Circle (CC(9, 1), CC(9, 2)), W(9) / 3, , 3 * PI / 4, 5 * PI / 4, HASPECT
1038 'PSet (P(1, 1), P(1, 2))
1040 For I = 2 To 8
1042    If LR(4) = 0 And L(4) = 0 Then I = 8
1044    If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 1050
1045    If (STDTIPS = 1000 Or INDEX = 1000) And (I = 4 Or I = 5) Then GoTo 1050
1046    'Line -(P(I, 1), P(I, 2))
1049    If INDEX = 1000 And DIPI = 100 And I = 2 Then I = 6
1050 Next I
1052 For I = 8 To 15
1054    If LR(3) = 0 And L(3) = 0 Then I = 15
1055    If (STDTIPS = 1000 Or MIDDLE = 1000) And (I = 11 Or I = 12) Then GoTo 1060
1056    If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 1060
1058    'Line -(P(I, 1), P(I, 2))
1059    If MIDDLE = 1000 And DIPM = 100 And I = 9 Then I = 13
1060 Next I
1062 For I = 15 To 22
1064    If LR(2) = 0 And L(2) = 0 Then I = 22
1065    If (STDTIPS = 1000 Or RING = 1000) And (I = 18 Or I = 19) Then GoTo 1070
1066    If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 1070
1068    'Line -(P(I, 1), P(I, 2))
1069    If RING = 1000 And DIPR = 100 And I = 16 Then I = 20
1070 Next I
1072 For I = 22 To 29
1074    If LR(1) = 0 And L(1) = 0 Then I = 29
1075    If (STDTIPS = 1000 Or LITTLE = 1000) And (I = 25 Or I = 26) Then GoTo 1080
1076    If P(I, 1) = 0 And P(I, 2) = 0 Then GoTo 1080
1078    'Line -(P(I, 1), P(I, 2))
1079    If LITTLE = 1000 And DIPL = 100 And I = 23 Then I = 27
1080 Next I
1095 'Line (P(32, 1), P(32, 2))-(P(33, 1), P(33, 2))
1096 'If STDTIPS = 1000 Or THUMB = 1000 And LR(5) <> 0 Then Line -(P(35, 1), P(35, 2))
1097 'Line (P(35, 1), P(35, 2))-(P(36, 1), P(36, 2)): Line -(P(31, 1), P(31, 2)): Line -(P(39, 1), P(39, 2)): Line -(P(40, 1), P(40, 2)): Line -(P(30, 1), P(30, 2))
1100 'Line -(P(31, 1), P(31, 2))
1105 If LR(4) = 0 And L(4) = 0 Then GoTo 1112
1110 'PSet (P(4, 1), P(4, 2)): Line -(P(5, 1), P(5, 2)), 0
1112 If LR(3) = 0 And L(3) = 0 Then GoTo 1117
1115 'PSet (P(11, 1), P(11, 2)): Line -(P(12, 1), P(12, 2)), 0
1117 If LR(2) = 0 And L(2) = 0 Then GoTo 1123
1120 'PSet (P(18, 1), P(18, 2)): Line -(P(19, 1), P(19, 2)), 0
1123 If LR(1) = 0 And L(1) = 0 Then GoTo 1130
1125 'PSet (P(25, 1), P(25, 2)): Line -(P(26, 1), P(26, 2)), 0
1130 'PSet (P(32, 1), P(32, 2)): Line -(P(37, 1), P(37, 2)): Line -(P(1, 1), P(1, 2))
1135 'PSet (P(37, 1), P(37, 2)): Line -(P(38, 1), P(38, 2)): Line -(P(41, 1), P(41, 2))
1140 If OpenTips = 1000 Then GoTo 1165
1150 'For I = 4 To 25 Step 7: PSet (P(I, 1), P(I, 2)): Next I
1160 'For I = 5 To 26 Step 7: PSet (P(I, 1), P(I, 2)): Next I
1165 If LR(5) = 0 Or STDTIPS = 1000 Or THUMB = 1000 Then GoTo 1233
1170 'Circle (CC(9, 1), CC(9, 2)), W(9) / 3, , 0.8 * PI, 1.22 * PI
1229 For I = PI / 2 To PI Step PI / 180: I5 = ((LR(5) / 2) * Cos(I)): J5 = ((2 * W(9) / 3) * Sin(I)): iI6 = (I5 * Cos(-PI / 4) - J5 * Sin(-PI / 4)): iJ6 = (I5 * Sin(-PI / 4) + J5 * Cos(-PI / 4))
1230 'PSet (I6 + CC(8, 1), J6 + CC(8, 2))
     Next I
1233 K = Int(20 - LR(7) / 2) - 2
1235 LOCATE K, 43: Print #3, sn$
1240 LOCATE K + 1, 42: Print #3, Left$(sOD$, 8); "-"; Right$(sOD$, 1)
1245 LOCATE K + 2, 43: Print #3, HAND$
1246 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 1247 Else GoTo 1250
1247 LOCATE K + 3, 42: Print #3, Str$(INSERT); " cm": GoTo 1260
1250 LOCATE K + 3, 42: Print #3, Str$(INSERT); Chr$(34)
1260 SCREENPLOT = SCREENPLOT + 1: NOPLOT = 0
1263 LOCATE K + 6, 10: Print #3, "Would you like a plot of this glove";
1265 Input #2, A$
    
    If A$ = "Y" Or A$ = "y" Then
        JUMPFLAG = 0  'PLOT GLOVE
    ElseIf A$ = "N" Or A$ = "n" Then
            NOPLOT = 1000
            JUMPFLAG = 1  'FUNCTION KEY OPTIONS
        Else
            GoTo 1265
    End If
    
    ' END IF
       
    Call Kevin

   'Get out now
    Exit Sub
    
1300 'SUB
1301 OpenTips = 0: STDTIPS = 0: LITTLE = 0: RING = 0: MIDDLE = 0: INDEX = 0: THUMB = 0
1302 If DESIGNER = 1000 Then LOCATE 10, 6: Print #3, sDESIGNER$; " are you still the"; Chr$(13); "     Designer of "; sn$; "'s glove (Y/N)";: Input #2, A$ Else GoTo 1306
1303 LOCATE 10, 6: Print #3, B$: LOCATE 11, 6: Print #3, B$
1305 If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1308
1306 LOCATE 10, 6: Print #3, "Please type in your name. ";: Input #2, sDESIGNER$: DESIGNER = 1000
1307 LOCATE 10, 6: Print #3, B$
1308 LOCATE 10, 6: Print #3, "Does "; sn$; " want OPEN TIPS (Y/N)";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then OpenTips = 1000 Else Return
1310   LOCATE 10, 6: Print #3, "Does "; sn$; " want STANDARD Open Tips"
1312 LOCATE 12, 6: Print #3, "on ALL "; n; " fingers and thumb (Y/N)";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then STDTIPS = 1000: Return Else GoTo 1315
1315 LOCATE 10, 6: Print #3, B$
1317 LOCATE 12, 6: Print #3, B$
1320 LOCATE 10, 6: Print #3, "Will the LITTLE finger be OPEN TIPPED (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1322 Else GoTo 1329
1322 LOCATE 10, 6: Print #3, B$: LITTLE = 1000
1324 LOCATE 10, 6: Print #3, "Is L(12) "; L(1); " the desired length of the LITTLE finger";: Input #2, A$
1326 If Asc(A$) = 89 Or Asc(A$) = 121 Then LR(1) = L(1): LOCATE 10, 6: Print #3, B$: GoTo 1330
1327 LOCATE 10, 6: Print #3, B$: LOCATE 10, 6: Print #3, "Please enter the desired length for the LITTLE finger. ";: Input #2, DL(1): If Asc(M$) = 77 Or Asc(M$) = 109 Then LR(1) = DL(1): LOCATE 10, 6: Print #3, B$: GoTo 1330
1328 DL(1) = FNCONVIE(DL(1)): LR(1) = DL(1): LOCATE 10, 6: Print #3, B$: GoTo 1330
1329 LR(1) = L(1) / 1.1: LOCATE 10, 6: Print #3, B$
1330 LOCATE 10, 6: Print #3, "Will the RING finger be OPEN TIPPED (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1332 Else GoTo 1339
1332 LOCATE 10, 6: Print #3, B$: RING = 1000
1334 LOCATE 10, 6: Print #3, "Is L(13) "; L(2); " the desired length of the RING finger";: Input #2, A$
1336 If Asc(A$) = 89 Or Asc(A$) = 121 Then LR(2) = L(2): LOCATE 10, 6: Print #3, B$: GoTo 1340
1337 LOCATE 10, 6
        Print #3, B$
        LOCATE 10, 6
        Print #3, "Please enter the desired length for the RING finger. ";
        Input #2, DL(2)
        If Asc(M$) = 77 Or Asc(M$) = 109 Then LR(2) = DL(2): LOCATE 10, 6: Print #3, B$: GoTo 1340

1338 DL(2) = FNCONVIE(DL(2)): LR(2) = DL(2): LOCATE 10, 6: Print #3, B$: GoTo 1340
1339 LR(2) = L(2) / 1.1: LOCATE 10, 6: Print #3, B$
1340 LOCATE 10, 6: Print #3, "Will the MIDDLE finger be OPEN TIPPED (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1342 Else GoTo 1349
1342 LOCATE 10, 6: Print #3, B$: MIDDLE = 1000
1344 LOCATE 10, 6: Print #3, "Is L(14) "; L(3); " the desired length of the MIDDLE finger";: Input #2, A$
1346 If Asc(A$) = 89 Or Asc(A$) = 121 Then LR(3) = L(3): LOCATE 10, 6: Print #3, B$: GoTo 1350
1347 LOCATE 10, 6: Print #3, B$: LOCATE 10, 6: Print #3, "Please enter the desired length for the MIDDLE finger. ";: Input #2, DL(3): If Asc(M$) = 77 Or Asc(M$) = 109 Then LR(3) = DL(3): LOCATE 10, 6: Print #3, B$: GoTo 1350
1348 DL(3) = FNCONVIE(DL(3)): LR(3) = DL(3): LOCATE 10, 6: Print #3, B$: GoTo 1350
1349 LR(3) = L(3) / 1.1: LOCATE 10, 6: Print #3, B$
1350 LOCATE 10, 6: Print #3, "Will the INDEX finger be OPEN TIPPED (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1352 Else GoTo 1359
1352 LOCATE 10, 6: Print #3, B$: INDEX = 1000
1354 LOCATE 10, 6: Print #3, "Is L(15) "; L(4); " the desired length of the INDEX finger";: Input #2, A$
1356 If Asc(A$) = 89 Or Asc(A$) = 121 Then LR(4) = L(4): LOCATE 10, 6: Print #3, B$: GoTo 1360
1357 LOCATE 10, 6: Print #3, B$: LOCATE 10, 6: Print #3, "Please enter the desired length for the INDEX finger. ";: Input #2, DL(4): If Asc(M$) = 77 Or Asc(M$) = 109 Then LR(4) = DL(4): LOCATE 10, 6: Print #3, B$: GoTo 1360
1358 DL(4) = FNCONVIE(DL(4)): LR(4) = DL(4): LOCATE 10, 6: Print #3, B$: GoTo 1360
1359 LR(4) = L(4) / 1.1: LOCATE 10, 6: Print #3, B$
1360 LOCATE 10, 6: Print #3, "Will the THUMB be OPEN TIPPED (Y/N) ";: Input #2, A$: If Asc(A$) = 89 Or Asc(A$) = 121 Then GoTo 1362 Else GoTo 1369
1362 LOCATE 10, 6: Print #3, B$: THUMB = 1000
1364 LOCATE 10, 6: Print #3, "Is L(16) "; L(5); " the desired length of the THUMB";: Input #2, A$
1366 If Asc(A$) = 89 Or Asc(A$) = 121 Then LR(5) = L(5): GoTo 1370
1367 LOCATE 10, 6: Print #3, B$: LOCATE 10, 6: Print #3, "Please enter the desired length for the THUMB. ";: Input #2, DL(5): If Asc(M$) = 77 Or Asc(M$) = 109 Then LR(5) = DL(5): LOCATE 10, 6: Print #3, B$: GoTo 1370
1368 DL(5) = FNCONVIE(DL(5)): LR(5) = DL(5): LOCATE 10, 6: Print #3, B$: GoTo 1370
1369 LR(5) = L(5) / 1.1: LOCATE 10, 6: Print #3, B$
1370 Return

1500 'SUB
1505 If L(8) < 3.5 Or L(7) < 3.5 Then CHILD = 1
1510 If STDTIPS = 1000 Then GoTo 1600
1515 If INDEX = 0 Then GoTo 1529
1519 If LR(4) = 0 Then GoTo 1529
1520 P(1, 1) = -(W(8) + W(6) / 2): P(8, 1) = -W(6) / 2: P(1, 2) = LR(8): P(8, 2) = LR(8)
1521 P(2, 1) = -(W(8) + W(6) / 2): P(7, 1) = -(W(6) / 2)
1522 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 1523 Else GoTo 1524
1523 If LR(4) < 0.492 Or C(7) = 0 Then P(2, 2) = LR(8) + LR(4): P(7, 2) = P(2, 2): DIPI = 100: GoTo 1529
1524 If LR(4) < 1.25 Or C(7) = 0 Then P(2, 2) = LR(8) + LR(4): P(7, 2) = P(2, 2): DIPI = 100: GoTo 1529
1526 P(2, 2) = LR(8) + 0.4 * LR(4): P(7, 2) = P(2, 2): P(3, 2) = LR(8) + LR(4): P(6, 2) = P(3, 2)
1527 P(3, 1) = -((W(8) + 0.88 * W(7) + W(6)) / 2): P(6, 1) = -((W(8) + W(6) - 0.88 * W(7)) / 2)
1529 If LR(3) = 0 Then GoTo 1539
1530 If MIDDLE = 0 Then GoTo 1539
1532 P(9, 1) = -(W(6) / 2): P(14, 1) = W(6) / 2: P(15, 1) = P(14, 1): P(15, 2) = P(8, 2)
1533 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 1534 Else GoTo 1536
1534 If LR(3) < 0.492 Or C(5) = 0 Then P(9, 2) = LR(8) + LR(3): P(14, 2) = P(9, 2): DIPM = 100: GoTo 1539
1536 If LR(3) < 1.25 Or C(5) = 0 Then P(9, 2) = LR(8) + LR(3): P(14, 2) = P(9, 2): DIPM = 100: GoTo 1539
1537 P(10, 1) = -(0.88 * W(5) / 2): P(13, 1) = 0.88 * W(5) / 2
1538 P(9, 2) = LR(8) + 0.4 * LR(3): P(14, 2) = P(9, 2): P(10, 2) = LR(8) + LR(3): P(13, 2) = P(10, 2)
1539 If LR(2) = 0 Then GoTo 1549
1540 If RING = 0 Then GoTo 1549
1542 P(16, 1) = W(6) / 2: P(21, 1) = W(6) / 2 + W(4): P(22, 1) = P(21, 1): P(22, 2) = LR(6)
1543 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 1544 Else GoTo 1546
1544 If LR(2) < 0.492 Or C(3) = 0 Then P(16, 2) = LR(7) + LR(2): P(21, 2) = P(16, 2): DIPR = 100: GoTo 1549
1546 If LR(2) < 1.25 Or C(3) = 0 Then P(16, 2) = LR(7) + LR(2): P(21, 2) = P(16, 2): DIPR = 100: GoTo 1549
1547 P(17, 1) = (W(6) + W(4) - 0.88 * W(3)) / 2: P(20, 1) = (W(6) + W(4) + 0.88 * W(3)) / 2
1548 P(16, 2) = LR(7) + 0.4 * LR(2): P(21, 2) = P(16, 2): P(17, 2) = LR(7) + LR(2): P(20, 2) = P(17, 2)
1549 If LR(1) = 0 Then GoTo 1560
1550 If LITTLE = 0 Then GoTo 1560
1552 P(23, 1) = W(6) / 2 + W(4): P(28, 1) = W(6) / 2 + W(4) + W(2): P(29, 1) = P(28, 1): P(29, 2) = LR(6)
1553 If Asc(M$) = 77 Or Asc(M$) = 109 Then GoTo 1554 Else GoTo 1556
1554 If LR(1) < 0.394 Or C(1) = 0 Then P(23, 2) = LR(6) + LR(1): P(28, 2) = P(23, 2): DIPL = 100: GoTo 1560
1556 If LR(1) < 1! Or C(1) = 0 Then P(23, 2) = LR(6) + LR(1): P(28, 2) = P(23, 2): DIPL = 100: GoTo 1560
1558 P(23, 2) = LR(6) + 0.4 * LR(1): P(28, 2) = P(23, 2): P(24, 2) = LR(6) + LR(1): P(27, 2) = P(24, 2)
1559 P(24, 1) = (W(6) + W(2) - 0.88 * W(1)) / 2 + W(4): P(27, 1) = (W(6) + W(2) + 0.88 * W(1)) / 2 + W(4)
1560 If THUMB = 0 Then GoTo 1590
1561 If LR(5) = 0 Then GoTo 1590
1562 P(35, 1) = -(W(6) / 2 + W(8) + T2 - (W(9) / 2) * Cos(PI / 4)) - (W(9) + LR(5)) * Cos(PI / 4): P(35, 2) = (LR(8) - LR(9)) / 4 + LR(9) - T2 + (W(9) / 2) * Sin(PI / 4) - (W(9) - LR(5)) * Sin(PI / 4)
1564 P(33, 1) = -(W(6) / 2 + W(8) + T2 - (W(9) / 2) * Cos(PI / 4)) - LR(5) * Cos(PI / 4): P(33, 2) = (LR(8) - LR(9)) / 4 + LR(9) - T2 + (W(9) / 2) * Sin(PI / 4) + LR(5) * Sin(PI / 4)
1590 Return

1600 'STD OPEN TIPS
1602 P(1, 1) = -(W(8) + W(6) / 2): P(1, 2) = LR(8)
1604 If LR(4) = 0 Then GoTo 1614
1606 P(2, 1) = -(W(8) + W(6) / 2): P(2, 2) = LR(8) + 0.4 * L(4)
1608 P(3, 1) = -((W(8) + 0.88 * W(7) + W(6)) / 2): P(3, 2) = LR(8) + 2 / 3 * L(4)
1610 P(6, 1) = -((W(8) + W(6) - 0.88 * W(7)) / 2): P(6, 2) = P(3, 2)
1612 P(7, 1) = -(W(6) / 2): P(7, 2) = P(2, 2)
1614 P(8, 1) = -(W(6) / 2): P(8, 2) = LR(8)
1620 If LR(3) = 0 Then GoTo 1628
1622 P(9, 1) = -(W(6) / 2): P(9, 2) = LR(8) + 0.4 * L(3): P(14, 1) = W(6) / 2: P(14, 2) = P(9, 2)
1624 P(10, 1) = -(0.88 * W(5) / 2): P(10, 2) = LR(8) + 2 / 3 * L(3)
1626 P(13, 1) = 0.88 * W(5) / 2: P(13, 2) = P(10, 2)
1628 P(15, 1) = W(6) / 2: P(15, 2) = LR(7)
1630 If LR(2) = 0 Then GoTo 1638
1632 P(16, 1) = W(6) / 2: P(16, 2) = LR(7) + 0.4 * L(2): P(21, 2) = P(16, 2)
1634 P(17, 1) = (W(6) + W(4) - 0.88 * W(3)) / 2: P(17, 2) = LR(7) + 2 / 3 * L(2)
1636 P(20, 1) = (W(6) + W(4) + 0.88 * W(3)) / 2: P(20, 2) = P(17, 2): P(21, 1) = W(6) / 2 + W(4)
1638 P(22, 1) = W(6) / 2 + W(4): P(22, 2) = LR(6)
1640 If LR(1) = 0 Then GoTo 1648
1642 P(23, 1) = W(6) / 2 + W(4): P(23, 2) = LR(6) + 0.4 * L(1): P(28, 2) = P(23, 2)
1644 P(24, 1) = (W(6) + W(2) - 0.88 * W(1)) / 2 + W(4): P(24, 2) = LR(6) + 2 / 3 * L(1)
1646 P(27, 1) = (W(6) + W(2) + 0.88 * W(1)) / 2 + W(4): P(27, 2) = P(24, 2): P(28, 1) = W(6) / 2 + W(2) + W(4)
1648 P(29, 1) = W(6) / 2 + W(2) + W(4): P(29, 2) = LR(6)
1650 If LR(5) = 0 Then GoTo 1690
1652 P(35, 1) = -(W(6) / 2 + W(8) + T2 - (W(9) / 2) * Cos(PI / 4)) - (W(9) + 2 / 3 * L(5)) * Cos(PI / 4): P(35, 2) = (LR(8) - LR(9)) / 4 + LR(9) - T2 + (W(9) / 2) * Sin(PI / 4) - (W(9) - 2 / 3 * L(5)) * Sin(PI / 4)
1654 P(33, 1) = -(W(6) / 2 + W(8) + T2 - (W(9) / 2) * Cos(PI / 4)) - 2 / 3 * L(5) * Cos(PI / 4): P(33, 2) = (LR(8) - LR(9)) / 4 + LR(9) - T2 + (W(9) / 2) * Sin(PI / 4) + 2 / 3 * L(5) * Sin(PI / 4)
1690 Return

8000 'SUB TO PLOT PALM SURFACE PARTS
8005 XBASE = -XBASE: X10 = -X10: XTI = -XTI: SEMICIRC = -SEMICIRC: QUARTCIRC1 = -QUARTCIRC1: QUARTCIRC2 = -QUARTCIRC2
8006 CONST1 = -CONST1: CONST2 = -CONST2: CONST3 = -CONST3: YC = 5 * K: 'YC2=-4084
8007 XR10 = -XR10: XRTI = -XRTI: JKR9 = -JKR9: JK9 = -JK9: JK8 = -JK8: JK7 = -JK7
8010 For I = 1 To 70
8020  P(I, 1) = -P(I, 1): TR(I, 1) = -TR(I, 1)
8030 Next I
8040 For I = 1 To 9
8050  CC(I, 1) = -CC(I, 1): CCT(I, 1) = -CCT(I, 1)
8060 Next I
8070 If PLOTONCE = 1000 Then YC7 = 5 * K
8080 Return
8081
End Sub

Private Sub CLS()
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub
Private Sub CLEAR()
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub

Private Sub LOCATE(pDummy1, pDummy2)
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub

Private Sub WINDOW(pDummy)
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub

Private Sub VIEW(pDummy)
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub
Private Function KEY(pDummy) As Boolean
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
    KEY = False
End Function

Private Function FNCONVIE(X) As Single
   'Replicates a QB 4.5 statment
    FNCONVIE = ((X / 10) - Int(X / 10)) / 0.8 + Int(X / 10)
End Function

Private Sub COLOR(pDummy1, pDummy2, pDummy3)
    'Dummy subroutine
    'This is used to save having to edit this QB 4.5 statment out of the code
End Sub



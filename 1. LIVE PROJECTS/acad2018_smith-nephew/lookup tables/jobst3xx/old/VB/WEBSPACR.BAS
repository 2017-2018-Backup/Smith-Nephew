Attribute VB_Name = "WEBSPACR1"
'Project:   WEBSPACR.BAS
'Purpose:
'
'
'Version:   1.01
'Date:      21 July 1996
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
'

    
   'PI as a Global Constant
    Global Const PI = 3.141592654

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

   'Store current layer and text setings to reduce DRAFIX code
   'this value is checked in PR_SetLayer
    Global g_sCurrentLayer As String
    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont
    Global g_nCurrTextAngle

    Global g_iInsertSize As Integer
    Global g_sFileNo As String
    Global g_sSide As String
    Global g_sPatient As String

    Global g_nUnitsFac As Double
    Global g_ExtendTo As Integer
    Global g_iNumTapesWristToEOS As Integer
    Global g_nPleats(1 To 4) As Double
    Global g_iFirstTape%
    Global g_iLastTape%
    Global g_iWristPointer%
    Global g_iEOSPointer%
    Global g_iNumTotalTapes%
    Global g_nCir(1 To 17) As Double
    Global g_sPathJOBST     As String
    Global g_TapesFound As Integer
    Global g_nThumbWebDrop  As Double
           
    Global MainForm As webspacr
    
    Global UlnarProfile     As Curve
    Global RadialProfile     As Curve

    Global xyFold             As XY
    Global xyTopThumbRad      As XY
    Global xyBotThumbRad      As XY
    Global xyW(1 To 30)       As XY
    Global xyT(1 To 10)       As XY
    Global xyN(1 To 3)       As XY
    Global nNotchRad#, nTopThumbRad#, nBotThumbRad#

    Global Const g_sDialogID = "Web Spacers - Draw"

    Global Const g_sTapeText = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"


    Global Const SIXTEENTH = 0.0625
    Global Const EIGHTH = 0.125
    Global Const QUARTER = 0.25

   'MsgBox return values
    Global Const IDOK = 1                  ' OK button pressed
    Global Const IDCANCEL = 2              ' Cancel button pressed
    Global Const IDABORT = 3               ' Abort button pressed
    Global Const IDRETRY = 4               ' Retry button pressed
    Global Const IDIGNORE = 5              ' Ignore button pressed
    Global Const IDYES = 6                 ' Yes button pressed
    Global Const IDNO = 7                  ' No button pressed
    Type TapeData
        nCir        As Double
        iMMs        As Integer
        iRed        As Integer
        iGms        As Integer
        sNote       As String
        iTapePos    As Integer
        sTapeText   As String
    End Type


   'Arm variables
    Global Const NOFF_ARMTAPES = 16
    Global Const ELBOW_TAPE = 9
    Global Const ARM_PLAIN = 0
    Global Const ARM_FLAP = 1
    Global Const GLOVE_NORMAL = 0
    Global Const GLOVE_ELBOW = 1
    Global Const GLOVE_AXILLA = 2
    Global Const GLOVE_PASTWRIST = 3

    Global TapeNote(1 To NOFF_ARMTAPES)   As TapeData
    
    Option Explicit



Option Strict Off
Option Explicit On
Module WEBSPACR1
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



    Structure xy
        Dim X As Double
        Dim Y As Double
    End Structure

    Public Structure curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim y() As Double

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim X(100)
            'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim y(100)
        End Sub
    End Structure

    'PI as a Global Constant
    Public Const PI As Double = 3.141592654
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fNum As Object 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	'Store current layer and text setings to reduce DRAFIX code
	'this value is checked in PR_SetLayer
	Public g_sCurrentLayer As String
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	Public g_nCurrTextAngle As Object
	
	Public g_iInsertSize As Short
	Public g_sFileNo As String
	Public g_sSide As String
	Public g_sPatient As String
	
	Public g_nUnitsFac As Double
	Public g_ExtendTo As Short
	Public g_iNumTapesWristToEOS As Short
	'UPGRADE_WARNING: Lower bound of array g_nPleats was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nPleats(4) As Double
	Public g_iFirstTape As Short
	Public g_iLastTape As Short
	Public g_iWristPointer As Short
	Public g_iEOSPointer As Short
	Public g_iNumTotalTapes As Short
	'UPGRADE_WARNING: Lower bound of array g_nCir was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public g_nCir(17) As Double
	Public g_sPathJOBST As String
	Public g_TapesFound As Short
	Public g_nThumbWebDrop As Double
	
	Public MainForm As webspacr
	
	'UPGRADE_WARNING: Arrays in structure UlnarProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public UlnarProfile As Curve
	'UPGRADE_WARNING: Arrays in structure RadialProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public RadialProfile As Curve
	
	Public xyFold As XY
	Public xyTopThumbRad As XY
	Public xyBotThumbRad As XY
	'UPGRADE_WARNING: Lower bound of array xyW was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public xyW(30) As XY
	'UPGRADE_WARNING: Lower bound of array xyT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public xyT(10) As XY
	'UPGRADE_WARNING: Lower bound of array xyN was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public xyN(3) As XY
	Public nTopThumbRad, nNotchRad, nBotThumbRad As Double
	
	Public Const g_sDialogID As String = "Web Spacers - Draw"
	
	Public Const g_sTapeText As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
	
	
	Public Const SIXTEENTH As Double = 0.0625
	Public Const EIGHTH As Double = 0.125
	Public Const QUARTER As Double = 0.25
	
	'MsgBox return values
	Public Const IDOK As Short = 1 ' OK button pressed
	Public Const IDCANCEL As Short = 2 ' Cancel button pressed
	Public Const IDABORT As Short = 3 ' Abort button pressed
	Public Const IDRETRY As Short = 4 ' Retry button pressed
	Public Const IDIGNORE As Short = 5 ' Ignore button pressed
	Public Const IDYES As Short = 6 ' Yes button pressed
	Public Const IDNO As Short = 7 ' No button pressed
	Structure TapeData
		Dim nCir As Double
		Dim iMMs As Short
		Dim iRed As Short
		Dim iGms As Short
		Dim sNote As String
		Dim iTapePos As Short
		Dim sTapeText As String
	End Structure
	
	
	'Arm variables
	Public Const NOFF_ARMTAPES As Short = 16
	Public Const ELBOW_TAPE As Short = 9
	Public Const ARM_PLAIN As Short = 0
	Public Const ARM_FLAP As Short = 1
	Public Const GLOVE_NORMAL As Short = 0
	Public Const GLOVE_ELBOW As Short = 1
	Public Const GLOVE_AXILLA As Short = 2
	Public Const GLOVE_PASTWRIST As Short = 3
	
	'UPGRADE_WARNING: Lower bound of array TapeNote was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public TapeNote(NOFF_ARMTAPES) As TapeData
	
End Module
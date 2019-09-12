Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class scoopdia
	Inherits System.Windows.Forms.Form
    'Project:   SCOOPDIA.MAK
    'Purpose:   VEST side scoop dialogue
    '
    '
    'Version:   3.01
    'Date:      20th Jan 1998
    'Author:    Gary George
    'Copyright  C-Gem Ltd
    '
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    'Dec 98     GG      Ported to VB5
    '
    'Notes:-
    '
    '
    '
    '
    '   'Windows API Functions Declarations
    '    Private Declare Function GetActiveWindow Lib "User" () As Integer
    '    Private Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer
    '    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
    '    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
    '    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
    '    Private Declare Function GetNumTasks Lib "Kernel" () As Integer
    '    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
    '    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
    '
    '
    '   'Constanst used by GetWindow
    '    Const GW_CHILD = 5
    '    Const GW_HWNDFIRST = 0
    '    Const GW_HWNDLAST = 1
    '    Const GW_HWNDNEXT = 2
    '    Const GW_HWNDPREV = 3
    '    Const GW_OWNER = 4'
    '
    '   'MsgBox constant'
    '    Const IDCANCEL = 2
    '    Const IDYES = 6
    '    Const IDNO = 7

    Structure XY
        Dim X As Double
        Dim y As Double
    End Structure

    Structure curve
        Dim n As Short
        <VBFixedArray(100)> Dim X() As Double
        <VBFixedArray(100)> Dim Y() As Double

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim X(100)
            'UPGRADE_WARNING: Lower bound of array Y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim Y(100)
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

    'mODULE LEVEL VARIABLES
    Dim m_sDialogueTitle As String
	
	'Vest details
	Dim xyBackNeckConstruct As xy
	Dim m_iBackNeckConstructUID As Short
	Dim m_iVestBodyUID As Short
	Dim m_iVestRightRaglanUID As Short
	Dim m_iVestLeftRaglanUID As Short
	Dim m_iVestBackNeckUID As Short
	Dim m_iVestFrontNeckUID As Short
	Dim m_iVestLeftOpenAxillaMarkerUID As Short
	Dim m_iVestRightOpenAxillaMarkerUID As Short
	
	'UPGRADE_WARNING: Arrays in structure m_VestLeftRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_VestLeftRaglan As curve
	'UPGRADE_WARNING: Arrays in structure m_NewVestLeftRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_NewVestLeftRaglan As curve
	'UPGRADE_WARNING: Arrays in structure m_ConstructLeftRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_ConstructLeftRaglan As curve
	'UPGRADE_WARNING: Arrays in structure m_VestRightRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_VestRightRaglan As curve
	'UPGRADE_WARNING: Arrays in structure m_NewVestRightRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_NewVestRightRaglan As curve
	'UPGRADE_WARNING: Arrays in structure m_ConstructRightRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_ConstructRightRaglan As curve
	Dim m_sVestRightRaglanDBData As String
	Dim m_sVestLeftRaglanDBData As String
	
	Dim xyVestLeftRaglanAtNeck As xy
	Dim xyVestLeftRaglanAtAxilla As xy
	Dim xyVestLeftRaglanStart As xy
	Dim xyVestRightRaglanAtNeck As xy
	Dim xyVestRightRaglanAtAxilla As xy
	Dim xyVestRightRaglanStart As xy
	Dim xyVestLeftopenaxillamarker As xy
	Dim xyVestRightopenaxillamarker As xy
	
	'UPGRADE_WARNING: Arrays in structure m_FrontNeckProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_FrontNeckProfile As curve
	'UPGRADE_WARNING: Arrays in structure m_NewFrontNeckProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_NewFrontNeckProfile As curve
	Dim m_bFrontNeckIsCurve As Short
	Dim m_nFrontNeckRadius As Double
	Dim xyFrontNeckCen As xy
	Dim xyNewFrontNeckCen As xy
	Dim xyFrontNeckOnRaglan As xy
	Dim xyFrontNeckOnCL As xy
	Dim xyNewFrontNeckOnRaglan As xy
	
	'UPGRADE_WARNING: Arrays in structure m_BackNeckProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_BackNeckProfile As curve
	'UPGRADE_WARNING: Arrays in structure m_NewBackNeckProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim m_NewBackNeckProfile As curve
	Dim m_bBackNeckIsCurve As Short
	Dim m_nBackNeckRadius As Double
	Dim xyBackNeckCen As xy
	Dim xyNewBackNeckCen As xy
	Dim xyBackNeckOnRaglan As xy
	Dim xyBackNeckOnCL As xy
	Dim xyNewBackNeckOnRaglan As xy
	
	Dim m_iReDrawSleeveTop As Short
	Dim m_iReDrawSleeveProfile As Short
	Dim m_iReDrawSleeveOrigin As Short
	
	Dim m_iAxilla As Short
	
	Dim m_nAFNRadRight As Double
	Dim m_nABNRadRight As Double
	Dim m_nSBRaglanRight As Double
	Dim m_nAxillaBackNeckRad As Double
	Dim m_nAxillaFrontNeckRad As Double
	Dim m_nShoulderToBackRaglan As Double
	
	Const MAX_SLEEVE_ReDRAW As Short = 10
	
	'UPGRADE_WARNING: Lower bound of array m_sSleeveProfileDBData was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_sSleeveProfileDBData(MAX_SLEEVE_ReDRAW) As String
	'UPGRADE_WARNING: Lower bound of array m_sSleeveOriginDBData was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_sSleeveOriginDBData(MAX_SLEEVE_ReDRAW) As String
	'UPGRADE_WARNING: Lower bound of array m_sSleeveTopDBData was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_sSleeveTopDBData(MAX_SLEEVE_ReDRAW) As String
	'UPGRADE_WARNING: Lower bound of array m_sSleeveProfileID was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_sSleeveProfileID(MAX_SLEEVE_ReDRAW) As String
	'UPGRADE_WARNING: Lower bound of array m_iSleeveProfileUID was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_iSleeveProfileUID(MAX_SLEEVE_ReDRAW) As Short
	'UPGRADE_WARNING: Lower bound of array m_iSleeveOriginUID was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_iSleeveOriginUID(MAX_SLEEVE_ReDRAW) As Short
	
	'UPGRADE_WARNING: Lower bound of array m_xySleeveOrigin was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_xySleeveOrigin(MAX_SLEEVE_ReDRAW) As xy
	
	
	Const MAX_ENTITIES_ON_DELETE_LIST As Short = 200
	
	'UPGRADE_WARNING: Lower bound of array m_iDeletedUID_List was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_iDeletedUID_List(MAX_ENTITIES_ON_DELETE_LIST) As Short
	'UPGRADE_WARNING: Lower bound of array m_sDeletedUID_List was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim m_sDeletedUID_List(MAX_ENTITIES_ON_DELETE_LIST) As String
	Dim m_iDeletedUID_Count As Short
	
	
	Dim m_bSleeve As Short
	Dim m_bVest As Short
	
	Dim m_bExecuteMeshDraw As Short
	
	Dim m_sRtAxillaType As String
	Dim m_sLtAxillaType As String
	Dim m_sMeshData As String
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click

        Return

    End Sub
	
	Private Function FN_CalcCirCurveInt(ByRef Profile As curve, ByRef xyCen As xy, ByRef nRadius As Double, ByRef ProfileLargePart As curve, ByRef ProfileSmallPart As curve, ByRef xyIntersection As xy) As Short
		
		Dim xyPt1 As xy
		Dim xyPt2 As xy
		Dim xyInt As xy
		Dim ii As Short
		Dim bIntersectionFound As Short
		'UPGRADE_WARNING: Arrays in structure ProfilePart1 may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ProfilePart1 As curve
		'UPGRADE_WARNING: Arrays in structure ProfilePart2 may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ProfilePart2 As curve
		
		
		bIntersectionFound = False
		FN_CalcCirCurveInt = False
		ProfilePart2.n = 0
		
		If Profile.n < 2 Then Exit Function
		
		ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), Profile.Y(1))
		
		ProfilePart1.X(1) = Profile.X(1)
		ProfilePart1.Y(1) = Profile.Y(1)
		ProfilePart1.n = 1
		
		For ii = 2 To Profile.n
			ARMDIA1.PR_MakeXY(xyPt2, Profile.X(ii), Profile.Y(ii))
			If Not bIntersectionFound Then bIntersectionFound = BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)
			If bIntersectionFound Then
				'Check that we have the upper intersection
				If xyInt.Y < xyCen.Y Then
					'If we don't have the upper intersection then we check that the intersection
					'is not on the same segment.  This is due to the nature of the function FN_CirLinInt
					'that returns only the first intesection
					bIntersectionFound = False
					If xyCen.Y < xyPt1.Y Then
						bIntersectionFound = BDYUTILS.FN_CirLinInt(xyPt1, xyCen, xyCen, nRadius, xyInt)
					Else
						bIntersectionFound = BDYUTILS.FN_CirLinInt(xyPt2, xyCen, xyCen, nRadius, xyInt)
					End If
				End If
				If bIntersectionFound Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyIntersection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyIntersection = xyInt
					ProfilePart1.n = ProfilePart1.n + 1
					ProfilePart1.X(ProfilePart1.n) = xyInt.X
					ProfilePart1.Y(ProfilePart1.n) = xyInt.Y
				End If
			End If
			
			If Not bIntersectionFound Then
				ProfilePart1.n = ProfilePart1.n + 1
				ProfilePart1.X(ProfilePart1.n) = Profile.X(ii)
				ProfilePart1.Y(ProfilePart1.n) = Profile.Y(ii)
			Else
				ProfilePart2.n = ProfilePart2.n + 1
				If ProfilePart2.n = 1 Then
					ProfilePart2.X(1) = xyIntersection.X
					ProfilePart2.Y(1) = xyIntersection.Y
				Else
					ProfilePart2.X(ProfilePart2.n) = Profile.X(ii)
					ProfilePart2.Y(ProfilePart2.n) = Profile.Y(ii)
				End If
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt1 = xyPt2
			
		Next ii
		
		If ProfilePart1.n > ProfilePart2.n Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ProfileLargePart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ProfileLargePart = ProfilePart1
			'UPGRADE_WARNING: Couldn't resolve default property of object ProfileSmallPart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ProfileSmallPart = ProfilePart2
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object ProfileLargePart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ProfileLargePart = ProfilePart2
			'UPGRADE_WARNING: Couldn't resolve default property of object ProfileSmallPart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ProfileSmallPart = ProfilePart1
		End If
		If bIntersectionFound Then FN_CalcCirCurveInt = True
		
	End Function
	
	Private Function FN_EscapeQuotesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Dim ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As String
		Dim sEscapedString As String
		
		FN_EscapeQuotesInString = ""
		
		For ii = 1 To Len(sAssignedString)
			'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Char_Renamed = Mid(sAssignedString, ii, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Char_Renamed = QQ Then
				sEscapedString = sEscapedString & "\" & Char_Renamed
			Else
				sEscapedString = sEscapedString & Char_Renamed
			End If
		Next ii
		
		FN_EscapeQuotesInString = sEscapedString
		
	End Function
	
	Private Function FN_OpenAxillaMarker(ByRef Profile As curve, ByRef xyIntersection As xy, ByRef aAngle As Double) As Short
		
		Dim xyPt1 As xy
		Dim xyPt2 As xy
		Dim xyInt As xy
		Dim ii As Short
		Dim iError As Short
		Dim nLength As Double
		Dim nRequired As Double
		
		
		FN_OpenAxillaMarker = False
		
		If Profile.n < 2 Then Exit Function
		
		ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), Profile.Y(1))
		nLength = 0
		For ii = 2 To Profile.n
			ARMDIA1.PR_MakeXY(xyPt2, Profile.X(ii), Profile.Y(ii))
			nLength = nLength + ARMDIA1.FN_CalcLength(xyPt1, xyPt2)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt1 = xyPt2
		Next ii
		
		nRequired = nLength / 3
		
		ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), Profile.Y(1))
		nLength = 0
		For ii = 2 To Profile.n
			ARMDIA1.PR_MakeXY(xyPt2, Profile.X(ii), Profile.Y(ii))
			nLength = nLength + ARMDIA1.FN_CalcLength(xyPt1, xyPt2)
			
			If nLength >= nRequired Then
				nLength = nLength - nRequired
				If nLength > 0 Then
					iError = BDYUTILS.FN_CirLinInt(xyPt1, xyPt2, xyPt1, nLength, xyIntersection)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object xyIntersection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyIntersection = xyPt2
				End If
				FN_OpenAxillaMarker = True
				aAngle = ARMDIA1.FN_CalcAngle(xyPt1, xyPt2) - 90
				Exit For
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt1 = xyPt2
		Next ii
		
		
	End Function
	
	
	Private Function FN_ValidateData() As Short
		'This function is used only to make gross checks
		'for missing data.
		'It does not perform any sensibility checks on the
		'data
		Dim sError As String
		Dim ii, nn As Short
		
		'Initialise
		FN_ValidateData = False
		sError = ""
		
		'Dimensions
		If Val(txtDim(0).Text) = 0 And Val(txtDim(1).Text) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sError = sError & "Missing dimensions for Side Scoop!" & NL
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Val(txtDim(0).Text) > 0 And Val(txtDim(0).Text) < 0.1 Then sError = sError & "Front Side Scoop must be greater than 1/10th Inch!" & NL
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Val(txtDim(1).Text) > 0 And Val(txtDim(1).Text) < 0.1 Then sError = sError & "Back Side Scoop must be greater than 1/10th Inch!" & NL
		
		'Check for errors
		If sError <> "" And m_bVest Then
			MsgBox(sError, 16, m_sDialogueTitle)
			FN_ValidateData = False
		Else
			FN_ValidateData = True
		End If
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Dim ii As Short
		Dim iEntityUID As Short
		Dim fFileNum As Short
		Dim sEntityname As String
		Dim sEntityType As String
		Dim sTemplateToScoop As String
		Dim sEntityID As String
		Dim aStartAngle As Double
		Dim aDeltaAngle As Double
		
		'Dummys
		'UPGRADE_WARNING: Arrays in structure Dummy may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Dummy As curve
		Dim xyPt As xy
		Dim xyPt1 As xy
		Dim xyPt2 As xy
		Dim nRadius As Double
		Dim sDummy As String
		
		
		'Stop the timer used to ensure that the Dialogue dies
		'if the DRAFIX macro fails to establish a DDE Link
		Timer1.Enabled = False
		
		'Check that a Data File name has been passed
		If txtScoopDataFile.Text = "" Then
			MsgBox("No VEST details have been found in drawing!", 16, m_sDialogueTitle)
            Return
        End If
		
		'Get ID to mark the nee draw entities
		g_sID = txtID.Text
		
		'Load from file the vest details
		'Setup error handling
		On Error GoTo ErrorStarting
		
		'Check that Data file exist
		'Note that this is a fixed file name
		If FileLen(txtScoopDataFile.Text) = 0 Then
			MsgBox("Can't open the file " & txtScoopDataFile.Text & ".  Unable to modify the VEST Neck.", 48, m_sDialogueTitle)
			End
		End If
		
		m_iReDrawSleeveTop = 0
		m_iReDrawSleeveProfile = 0
		m_iReDrawSleeveOrigin = 0
		
		m_iVestRightRaglanUID = -1
		m_iVestLeftRaglanUID = -1
		m_iVestLeftOpenAxillaMarkerUID = -1
		m_iVestRightOpenAxillaMarkerUID = -1
		
		m_bSleeve = False
		m_bVest = False
		
		'Open file for reading only
		fFileNum = FreeFile
		FileOpen(fFileNum, txtScoopDataFile.Text, OpenMode.Input, OpenAccess.Read)
		
		Input(fFileNum, sTemplateToScoop)
		If sTemplateToScoop = "SLEEVE" Then
			m_bSleeve = True
		ElseIf sTemplateToScoop = "VEST" Then 
			m_bVest = True
		Else
			GoTo ErrorStarting
		End If
		
		Do While Not EOF(fFileNum)
			Input(fFileNum, sEntityname)
			Input(fFileNum, sEntityType)
			Input(fFileNum, iEntityUID)
			sEntityID = LineInput(fFileNum)
			Select Case sEntityname
				Case "vestbackneckconstruct"
					m_iBackNeckConstructUID = iEntityUID
					Input(fFileNum, xyBackNeckConstruct.X)
					Input(fFileNum, xyBackNeckConstruct.Y)
					
				Case "vestbody"
					m_iVestBodyUID = iEntityUID
					
				Case "vestfrontneck"
					If sEntityType = "curve" Or sEntityType = "polyline" Then
						'It is a curve
						m_bFrontNeckIsCurve = True
						PR_GetCurveFromFile(fFileNum, m_FrontNeckProfile, xyFrontNeckOnCL, xyFrontNeckOnRaglan)
					Else
						'Assume that it is an Arc
						m_bFrontNeckIsCurve = False
						Input(fFileNum, xyFrontNeckCen.X)
						Input(fFileNum, xyFrontNeckCen.Y)
						Input(fFileNum, m_nFrontNeckRadius)
						Input(fFileNum, aStartAngle)
						Input(fFileNum, aDeltaAngle)
						ARMDIA1.PR_CalcPolar(xyFrontNeckCen, aStartAngle, m_nFrontNeckRadius, xyFrontNeckOnRaglan)
						ARMDIA1.PR_CalcPolar(xyFrontNeckCen, aStartAngle + aDeltaAngle, m_nFrontNeckRadius, xyFrontNeckOnCL)
					End If
					m_iVestFrontNeckUID = iEntityUID
					
				Case "vestprofile"
					'Read as dummy
					PR_GetCurveFromFile(fFileNum, Dummy, xyPt, xyPt)
					
				Case "vestLeftopenaxillamarker"
					Input(fFileNum, xyVestLeftopenaxillamarker.X)
					Input(fFileNum, xyVestLeftopenaxillamarker.Y)
					If m_iVestLeftOpenAxillaMarkerUID = -1 Then
						'Save only the first one
						m_iVestLeftOpenAxillaMarkerUID = iEntityUID
					Else
						'Delete any extranous ones
						PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					End If
					
				Case "vestRightopenaxillamarker"
					Input(fFileNum, xyVestRightopenaxillamarker.X)
					Input(fFileNum, xyVestRightopenaxillamarker.Y)
					If m_iVestRightOpenAxillaMarkerUID = -1 Then
						'Save only the first one
						m_iVestRightOpenAxillaMarkerUID = iEntityUID
					Else
						'Delete any extranous ones
						PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					End If
					
				Case "vestLeftraglan"
					PR_GetCurveFromFile(fFileNum, m_VestLeftRaglan, xyVestLeftRaglanAtNeck, xyVestLeftRaglanStart)
					If m_iVestLeftRaglanUID = -1 Then
						'Save only the first one
						m_iVestLeftRaglanUID = iEntityUID
						m_sVestLeftRaglanDBData = sEntityID
					Else
						'Delete any extranous ones
						PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					End If
					
				Case "vestRightraglan"
					PR_GetCurveFromFile(fFileNum, m_VestRightRaglan, xyVestRightRaglanAtNeck, xyVestRightRaglanStart)
					If m_iVestRightRaglanUID = -1 Then
						'Save only the first one
						m_iVestRightRaglanUID = iEntityUID
						m_sVestRightRaglanDBData = sEntityID
					Else
						'Delete any extranous ones
						PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					End If
					
				Case "vestRightaxillaconstruct"
					Input(fFileNum, xyPt1.X)
					Input(fFileNum, xyPt1.Y)
					Input(fFileNum, xyPt2.X)
					Input(fFileNum, xyPt2.Y)
					If xyPt1.Y < xyPt2.Y Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyVestRightRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyVestRightRaglanAtAxilla = xyPt2
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object xyVestRightRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyVestRightRaglanAtAxilla = xyPt1
					End If
					m_iAxilla = m_iAxilla + 1
					
				Case "vestLeftaxillaconstruct"
					Input(fFileNum, xyPt1.X)
					Input(fFileNum, xyPt1.Y)
					Input(fFileNum, xyPt2.X)
					Input(fFileNum, xyPt2.Y)
					If xyPt1.Y < xyPt2.Y Then
						'UPGRADE_WARNING: Couldn't resolve default property of object xyVestLeftRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyVestLeftRaglanAtAxilla = xyPt2
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object xyVestLeftRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyVestLeftRaglanAtAxilla = xyPt1
					End If
					m_iAxilla = m_iAxilla + 1
					
				Case "sleeveraglan"
					'Note we may have multiple sleeves based on the one vest
					'We use seperate counters for top and bottom curves
					'The data held in the string sEntityID allows us to link the various sleeve
					'components in the and contains enough data to redraw
					m_iReDrawSleeveTop = m_iReDrawSleeveTop + 1
					m_sSleeveTopDBData(m_iReDrawSleeveTop) = sEntityID
					PR_GetCurveFromFile(fFileNum, Dummy, xyPt, xyPt)
					PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					
				Case "sleeveprofile"
					'Note we may have multiple sleeves based on the one vest
					m_iReDrawSleeveProfile = m_iReDrawSleeveProfile + 1
					m_sSleeveProfileDBData(m_iReDrawSleeveProfile) = sEntityID
					If InStr(sEntityID, "RightProfile") > 0 Then
						m_sSleeveProfileID(m_iReDrawSleeveProfile) = Mid(sEntityID, 1, Len(sEntityID) - 12)
					Else
						m_sSleeveProfileID(m_iReDrawSleeveProfile) = Mid(sEntityID, 1, Len(sEntityID) - 11)
					End If
					m_iSleeveProfileUID(m_iReDrawSleeveProfile) = iEntityUID
					PR_GetCurveFromFile(fFileNum, Dummy, xyPt, xyPt)
					
				Case "sleeveoriginmark"
					m_iReDrawSleeveOrigin = m_iReDrawSleeveOrigin + 1
					m_sSleeveOriginDBData(m_iReDrawSleeveOrigin) = sEntityID
					m_iSleeveOriginUID(m_iReDrawSleeveOrigin) = iEntityUID
					Input(fFileNum, xyPt.X)
					Input(fFileNum, xyPt.Y)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_xySleeveOrigin(m_iReDrawSleeveOrigin). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_xySleeveOrigin(m_iReDrawSleeveOrigin) = xyPt
					
				Case "vestbackneck"
					If sEntityType = "curve" Or sEntityType = "polyline" Then
						'It is a curve
						m_bBackNeckIsCurve = True
						PR_GetCurveFromFile(fFileNum, m_BackNeckProfile, xyBackNeckOnCL, xyBackNeckOnRaglan)
					Else
						'Assume that it is an Arc
						m_bBackNeckIsCurve = False
						Input(fFileNum, xyBackNeckCen.X)
						Input(fFileNum, xyBackNeckCen.Y)
						Input(fFileNum, m_nBackNeckRadius)
						Input(fFileNum, aStartAngle)
						Input(fFileNum, aDeltaAngle)
						ARMDIA1.PR_CalcPolar(xyBackNeckCen, aStartAngle, m_nBackNeckRadius, xyBackNeckOnRaglan)
						ARMDIA1.PR_CalcPolar(xyBackNeckCen, aStartAngle + aDeltaAngle, m_nBackNeckRadius, xyBackNeckOnCL)
					End If
					m_iVestBackNeckUID = iEntityUID
					
				Case Else
					'Input the curve depending on it's class into dummy variables.
					'Add them to the deleted list as we know that these entites are part of sleeve
					'raglans that will be redrawn
					
					Select Case sEntityType
						Case "curve", "polyline"
							PR_GetCurveFromFile(fFileNum, Dummy, xyPt, xyPt)
						Case "arc"
							Input(fFileNum, xyPt.X)
							Input(fFileNum, xyPt.Y)
							Input(fFileNum, nRadius)
							Input(fFileNum, aStartAngle)
							Input(fFileNum, aDeltaAngle)
						Case "line"
							Input(fFileNum, xyPt.X)
							Input(fFileNum, xyPt.Y)
							Input(fFileNum, xyPt.X)
							Input(fFileNum, xyPt.Y)
						Case "marker"
							Input(fFileNum, xyPt.X)
							Input(fFileNum, xyPt.Y)
					End Select
					
					If Not sEntityname = "NoCurveTypeFound" Then PR_AddEntityToDeletedUID_List(iEntityUID, sEntityname)
					
			End Select
			
		Loop  'End of DO While
		
		'Close file
		FileClose(fFileNum)
		
		m_sLtAxillaType = txtLtAxillaType.Text
		m_sRtAxillaType = txtRtAxillaType.Text
		m_sMeshData = txtMeshData.Text
		
		m_nAxillaBackNeckRad = Val(txtAxillaBackNeckRad.Text)
		m_nAxillaFrontNeckRad = Val(txtAxillaFrontNeckRad.Text)
		m_nShoulderToBackRaglan = Val(txtShoulderToBackRaglan.Text)
		m_nAFNRadRight = Val(txtAFNRadRight.Text)
		m_nABNRadRight = Val(txtABNRadRight.Text)
		m_nSBRaglanRight = Val(txtSBRaglanRight.Text)
		
		If m_bSleeve Then
			OK_Click(OK, New System.EventArgs())
		Else
			Show()
		End If
		
		'Change pointer to normal
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorStarting: 
		
		MsgBox("Error Reading Input file")
        Return

    End Sub
	
	
	
	Private Sub scoopdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		
		Hide()
		
		'Start a timer
		'The Timer is disabled in LinkClose
		'If after 6 seconds the timer event will "End" the programme
		'This ensures that the dialogue dies in event of a failure
		'on the drafix macro side
		Timer1.Interval = 6000 'Approx 6 Seconds
		Timer1.Enabled = True
		
		'Initialize globals
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes (")
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(13) 'New Line
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCQ = QQ & CC & QQ
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QC = QQ & CC
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CQ = CC & QQ
		
		'Check if a previous instance is running
		'If it is warn user and exit
		m_sDialogueTitle = "VEST Side Scoop Neck"
		
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MsgBox(m_sDialogueTitle & " is already running!" & NL & "Use ALT-TAB and Cancel it.", 16, m_sDialogueTitle)
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		MainForm = Me
		
		g_nUnitsFac = 1
		g_sPathJOBST = fnPathJOBST()
		
		'Clear fields
		'Circumferences and lengths
		txtDim(0).Text = ""
		txtScoopDataFile.Text = ""
		txtID.Text = ""
		txtAxillaBackNeckRad.Text = ""
		txtAxillaFrontNeckRad.Text = ""
		txtShoulderToBackRaglan.Text = ""
		txtAFNRadRight.Text = ""
		txtABNRadRight.Text = ""
		txtSBRaglanRight.Text = ""
		
		'Debug "Read from file"
		'    txtScoopDataFile = "C:\JOBST\SCOOPNCK.DAT"
		'    Form_LinkClose
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim sTask As String
		'Don't allow multiple clicking
		'
		OK.Enabled = False
		If FN_ValidateData() Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			Hide()
			
			PR_CreateMacro_Draw("c:\jobst\draw.d")
			
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Return
            Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, m_sDialogueTitle)
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub PR_AddEntityToDeletedUID_List(ByRef iEntityUID As Short, ByRef sEntity As String)
		m_iDeletedUID_Count = m_iDeletedUID_Count + 1
		m_iDeletedUID_List(m_iDeletedUID_Count) = iEntityUID
		m_sDeletedUID_List(m_iDeletedUID_Count) = sEntity
	End Sub
	
	Private Sub PR_CreateMacro_Draw(ByRef sFile As String)
		'Since this is a fairly straight forward module we
		'will keep calculation and drawing in the same procedure
		'
		'The assumption is that all data has been validated before
		'this procedure is called
		
		Dim iError As Short
		Dim ii, jj As Short
		Dim nA As Double
		Dim nB As Double
		Dim nC As Double
		Dim aA As Double
		Dim xyTmp As xy
		Dim xySleeveOrigin As xy
		Dim xyVestRaglanAtAxilla As xy
		
		Dim iProfileUID As Short
		Dim iOriginUID As Short
		Dim iLowestOpenAxillaMarkerUID As Short
		Dim iHighestOpenAxillaMarkerUID As Short
		Dim sTimeStampTop As String
		Dim sTimeStamp As String
		Dim sID As String
		Dim sHighVestDBData As String
		Dim sHighVestCurveType As String
		Dim sHighestSide As String
		Dim sLowestSide As String
		Dim sLowVestDBData As String
		Dim sLowVestCurveType As String
		
		
		'UPGRADE_WARNING: Arrays in structure DummyProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim DummyProfile As curve
		'UPGRADE_WARNING: Arrays in structure HighestVestRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim HighestVestRaglan As curve
		'UPGRADE_WARNING: Arrays in structure LowestVestRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim LowestVestRaglan As curve
		'UPGRADE_WARNING: Arrays in structure ConstructRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ConstructRaglan As curve
		'UPGRADE_WARNING: Arrays in structure NewVestRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim NewVestRaglan As curve
		'UPGRADE_WARNING: Arrays in structure NewLowVestRaglan may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim NewLowVestRaglan As curve
		
		Dim nFrontSideScoop As Double
		Dim nBackSideScoop As Double
		
		Dim xyOpenAxillaMaker As xy
		Dim aAngle As Double
		
		
		
		nFrontSideScoop = Val(txtDim(0).Text)
		nBackSideScoop = Val(txtDim(1).Text)
		
		If m_bVest Then
			'Where we have two raglan curves then we use the one with the highest axilla to
			'Calculate the new front neck
			If m_iAxilla = 2 Then
				If xyVestLeftRaglanAtAxilla.X > xyVestRightRaglanAtAxilla.X Then
					sHighestSide = "Left"
					sLowestSide = "Right"
					'UPGRADE_WARNING: Couldn't resolve default property of object HighestVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					HighestVestRaglan = m_VestLeftRaglan
					'UPGRADE_WARNING: Couldn't resolve default property of object LowestVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LowestVestRaglan = m_VestRightRaglan
					iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
					iLowestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
					'UPGRADE_WARNING: Couldn't resolve default property of object xyVestRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyVestRaglanAtAxilla = xyVestLeftRaglanAtAxilla
					sHighVestCurveType = "vestLeftraglan"
					sLowVestCurveType = "vestRightraglan"
					sLowVestDBData = m_sVestLeftRaglanDBData
					sHighVestDBData = m_sVestRightRaglanDBData
				Else
					sHighestSide = "Right"
					sLowestSide = "Left"
					'UPGRADE_WARNING: Couldn't resolve default property of object HighestVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					HighestVestRaglan = m_VestRightRaglan
					'UPGRADE_WARNING: Couldn't resolve default property of object LowestVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LowestVestRaglan = m_VestLeftRaglan
					iHighestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
					iLowestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
					'UPGRADE_WARNING: Couldn't resolve default property of object xyVestRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyVestRaglanAtAxilla = xyVestRightRaglanAtAxilla
					sHighVestCurveType = "vestRightraglan"
					sLowVestCurveType = "vestLeftraglan"
					sLowVestDBData = m_sVestRightRaglanDBData
					sHighVestDBData = m_sVestLeftRaglanDBData
				End If
			Else
				sHighestSide = "Left"
				'UPGRADE_WARNING: Couldn't resolve default property of object HighestVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				HighestVestRaglan = m_VestLeftRaglan
				iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
				If m_iVestLeftOpenAxillaMarkerUID > -1 Then iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
				If m_iVestRightOpenAxillaMarkerUID > -1 Then iHighestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
				'UPGRADE_WARNING: Couldn't resolve default property of object xyVestRaglanAtAxilla. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				xyVestRaglanAtAxilla = xyVestLeftRaglanAtAxilla
				sHighVestCurveType = "vestLeftraglan"
				sHighVestDBData = m_sVestLeftRaglanDBData
			End If
			
			'Calculate New Front Neck on the Raglan
			If nFrontSideScoop > 0 Then
				iError = FN_CalcCirCurveInt(HighestVestRaglan, xyFrontNeckOnRaglan, nFrontSideScoop, DummyProfile, DummyProfile, xyNewFrontNeckOnRaglan)
				If m_bFrontNeckIsCurve Then
					
					'Calcluate new front  neck polyline using 1/3 and 1/2 rule
					m_NewFrontNeckProfile.n = 3
					m_NewFrontNeckProfile.X(1) = xyFrontNeckOnCL.X
					m_NewFrontNeckProfile.Y(1) = xyFrontNeckOnCL.Y
					
					m_NewFrontNeckProfile.X(2) = xyFrontNeckOnCL.X + ((xyNewFrontNeckOnRaglan.X - xyFrontNeckOnCL.X) / 3)
					m_NewFrontNeckProfile.Y(2) = xyFrontNeckOnCL.Y + ((xyNewFrontNeckOnRaglan.Y - xyFrontNeckOnCL.Y) / 2)
					
					m_NewFrontNeckProfile.X(m_NewFrontNeckProfile.n) = xyNewFrontNeckOnRaglan.X
					m_NewFrontNeckProfile.Y(m_NewFrontNeckProfile.n) = xyNewFrontNeckOnRaglan.Y
					
				Else
					'Calculate New front neck arc center
					'Using the new start and end points and the radius
					nA = m_nFrontNeckRadius + nFrontSideScoop
					nC = ARMDIA1.FN_CalcLength(xyNewFrontNeckOnRaglan, xyFrontNeckOnCL) / 2
					nB = System.Math.Sqrt(nA ^ 2 - nC ^ 2)
					aA = ARMDIA1.FN_CalcAngle(xyFrontNeckOnCL, xyNewFrontNeckOnRaglan)
					ARMDIA1.PR_CalcPolar(xyFrontNeckOnCL, aA, nC, xyTmp)
					ARMDIA1.PR_CalcPolar(xyTmp, aA - 90, nB, xyNewFrontNeckCen)
					
				End If
			End If
			
			'Calculate New Back Neck
			If nBackSideScoop > 0 Then
				iError = FN_CalcCirCurveInt(HighestVestRaglan, xyBackNeckOnRaglan, nBackSideScoop, NewVestRaglan, ConstructRaglan, xyNewBackNeckOnRaglan)
				
				If m_bBackNeckIsCurve Then
					'Calcluate new back neck polyline using 1/3 and 1/2 rule
					m_NewBackNeckProfile.n = 3
					m_NewBackNeckProfile.X(1) = xyBackNeckOnCL.X
					m_NewBackNeckProfile.Y(1) = xyBackNeckOnCL.Y
					
					m_NewBackNeckProfile.X(2) = xyBackNeckOnCL.X + ((xyNewBackNeckOnRaglan.X - xyBackNeckOnCL.X) / 3)
					m_NewBackNeckProfile.Y(2) = xyBackNeckOnCL.Y + ((xyNewBackNeckOnRaglan.Y - xyBackNeckOnCL.Y) / 2)
					
					m_NewBackNeckProfile.X(m_NewBackNeckProfile.n) = xyNewBackNeckOnRaglan.X
					m_NewBackNeckProfile.Y(m_NewBackNeckProfile.n) = xyNewBackNeckOnRaglan.Y
					
				Else
					'Calculate New back neck arc center
					'Using the new start and end points and the radius
					nA = m_nBackNeckRadius + nBackSideScoop
					nC = ARMDIA1.FN_CalcLength(xyNewBackNeckOnRaglan, xyBackNeckOnCL) / 2
					nB = System.Math.Sqrt(nA ^ 2 - nC ^ 2)
					aA = ARMDIA1.FN_CalcAngle(xyBackNeckOnCL, xyNewBackNeckOnRaglan)
					ARMDIA1.PR_CalcPolar(xyBackNeckOnCL, aA, nC, xyTmp)
					ARMDIA1.PR_CalcPolar(xyTmp, aA - 90, nB, xyNewBackNeckCen)
				End If
			End If
			
			'Highest
			If nFrontSideScoop > 0 Then m_nAxillaFrontNeckRad = ARMDIA1.FN_CalcLength(xyVestRaglanAtAxilla, xyNewFrontNeckOnRaglan)
			If nBackSideScoop > 0 Then
				m_nAxillaBackNeckRad = ARMDIA1.FN_CalcLength(xyVestRaglanAtAxilla, xyNewBackNeckOnRaglan)
				m_nShoulderToBackRaglan = ARMDIA1.FN_CalcLength(xyBackNeckConstruct, xyNewBackNeckOnRaglan)
			End If
			m_nAFNRadRight = m_nAxillaFrontNeckRad
			m_nABNRadRight = m_nAxillaBackNeckRad
			m_nSBRaglanRight = m_nShoulderToBackRaglan
			
			'adjust if 2 axilla hts
			If m_iAxilla = 2 Then
				If nBackSideScoop > 0 Then iError = FN_CalcCirCurveInt(LowestVestRaglan, xyBackNeckOnRaglan, nBackSideScoop, NewLowVestRaglan, ConstructRaglan, xyTmp)
				If nFrontSideScoop > 0 Then
					m_nAFNRadRight = ARMDIA1.FN_CalcLength(xyVestRightRaglanAtAxilla, xyNewFrontNeckOnRaglan)
					m_nAxillaFrontNeckRad = ARMDIA1.FN_CalcLength(xyVestLeftRaglanAtAxilla, xyNewFrontNeckOnRaglan)
				End If
				If nBackSideScoop > 0 Then
					m_nABNRadRight = ARMDIA1.FN_CalcLength(xyVestRightRaglanAtAxilla, xyNewBackNeckOnRaglan)
					m_nAxillaBackNeckRad = ARMDIA1.FN_CalcLength(xyVestLeftRaglanAtAxilla, xyNewBackNeckOnRaglan)
					m_nSBRaglanRight = ARMDIA1.FN_CalcLength(xyBackNeckConstruct, xyNewBackNeckOnRaglan)
				End If
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fNum = ARMDIA1.FN_Open(sFile)
			
			'PR_DrawMarker xyFrontNeckOnRaglan
			'PR_DrawText "xyFrontNeckOnRaglan", xyFrontNeckOnRaglan, .1
			'PR_DrawMarker xyNewFrontNeckOnRaglan
			'PR_DrawText "xyNewFrontNeckOnRaglan", xyNewFrontNeckOnRaglan, .1
			'PR_DrawMarker xyVestRightRaglanAtAxilla
			'PR_DrawText "xyVestRightRaglanAtAxilla", xyVestRightRaglanAtAxilla, .1
			'PR_DrawMarker xyVestLeftRaglanAtAxilla
			'PR_DrawText "xyVestLeftRaglanAtAxilla", xyVestLeftRaglanAtAxilla, .1
			'PR_DrawMarker xyBackNeckConstruct
			'PR_DrawText "xyBackNeckConstruct", xyBackNeckConstruct, .1
			
			
			If nFrontSideScoop > 0 Then
				ARMDIA1.PR_SetLayer("Notes")
				If m_bFrontNeckIsCurve Then
					ARMDIA1.PR_DrawFitted(m_NewFrontNeckProfile)
				Else
					ARMDIA1.PR_DrawArc(xyNewFrontNeckCen, xyFrontNeckOnCL, xyNewFrontNeckOnRaglan)
				End If
				BDYUTILS.PR_AddDBValueToLast("curvetype", "vestfrontneck")
				BDYUTILS.PR_AddDBValueToLast("Data", sHighVestDBData)
				PR_AddEntityToDeletedUID_List(m_iVestFrontNeckUID, "vestfrontneck")
			End If
			
			If nBackSideScoop > 0 Then
				
				iError = FN_CalcCirCurveInt(HighestVestRaglan, xyBackNeckOnRaglan, nBackSideScoop, NewVestRaglan, ConstructRaglan, xyNewBackNeckOnRaglan)
				
				ARMDIA1.PR_SetLayer("TemplateLeft")
				If m_bBackNeckIsCurve Then
					ARMDIA1.PR_DrawFitted(m_NewBackNeckProfile)
				Else
					ARMDIA1.PR_DrawArc(xyNewBackNeckCen, xyBackNeckOnCL, xyNewBackNeckOnRaglan)
				End If
				BDYUTILS.PR_AddDBValueToLast("curvetype", "vestbackneck")
				PR_AddEntityToDeletedUID_List(m_iVestBackNeckUID, "vestbackneck")
				BDYUTILS.PR_AddDBValueToLast("Data", sHighVestDBData)
				
				ARMDIA1.PR_SetLayer("Template" & sHighestSide)
				ARMDIA1.PR_DrawPoly(NewVestRaglan)
				BDYUTILS.PR_AddDBValueToLast("curvetype", sHighVestCurveType)
				BDYUTILS.PR_AddDBValueToLast("Data", sHighVestDBData)
				If iHighestOpenAxillaMarkerUID > -1 Then
					iError = FN_OpenAxillaMarker(NewVestRaglan, xyOpenAxillaMaker, aAngle)
					ARMDIA1.PR_SetLayer("Notes")
					BDYUTILS.PR_DrawMarkerNamed("medarrow", xyOpenAxillaMaker, 0.2, 0.1, aAngle)
					BDYUTILS.PR_AddDBValueToLast("Data", sHighVestDBData)
					BDYUTILS.PR_AddDBValueToLast("curvetype", "vest" & sHighestSide & "openaxillamarker")
				End If
				If m_iAxilla = 2 Then
					ARMDIA1.PR_SetLayer("Template" & sLowestSide)
					ARMDIA1.PR_DrawPoly(NewLowVestRaglan)
					BDYUTILS.PR_AddDBValueToLast("curvetype", sLowVestCurveType)
					BDYUTILS.PR_AddDBValueToLast("Data", sLowVestDBData)
					If iLowestOpenAxillaMarkerUID > -1 Then
						iError = FN_OpenAxillaMarker(NewLowVestRaglan, xyOpenAxillaMaker, aAngle)
						ARMDIA1.PR_SetLayer("Notes")
						BDYUTILS.PR_DrawMarkerNamed("medarrow", xyOpenAxillaMaker, 0.2, 0.1, aAngle)
						BDYUTILS.PR_AddDBValueToLast("Data", sLowVestDBData)
						BDYUTILS.PR_AddDBValueToLast("curvetype", "vest" & sLowestSide & "openaxillamarker")
					End If
				End If
				
				PR_AddEntityToDeletedUID_List(m_iVestLeftRaglanUID, "vestLeftraglan")
				PR_AddEntityToDeletedUID_List(m_iVestRightRaglanUID, "vestRightraglan")
				If m_iVestLeftOpenAxillaMarkerUID > -1 Then PR_AddEntityToDeletedUID_List(m_iVestLeftOpenAxillaMarkerUID, "vestLeftopenaxillamarker")
				If m_iVestRightOpenAxillaMarkerUID > -1 Then PR_AddEntityToDeletedUID_List(m_iVestRightOpenAxillaMarkerUID, "vestRightopenaxillamarker")
				
			End If
			
			'Delete entities and others
			PR_DoDeleteEntityByUID_List()
			
			'Update database
			PR_UpdateDB()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
			
			
		End If
		
		'ReDraw sleeves
		'Check that there are enough sleeve bits etc
		If m_bSleeve Then
			m_bExecuteMeshDraw = False
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fNum = ARMDIA1.FN_Open(sFile)
			If m_iReDrawSleeveTop > 0 And m_iReDrawSleeveTop = m_iReDrawSleeveProfile And m_iReDrawSleeveTop = m_iReDrawSleeveOrigin Then
				'loop through these entities finding the matches
				For ii = 1 To m_iReDrawSleeveTop
					iProfileUID = -1
					sTimeStampTop = ARMDIA1.fnGetString(m_sSleeveTopDBData(ii), 2, ",")
					
					For jj = 1 To m_iReDrawSleeveProfile
						sTimeStamp = ARMDIA1.fnGetString(m_sSleeveProfileDBData(jj), 2, ",")
						If sTimeStamp = sTimeStampTop Then
							iProfileUID = m_iSleeveProfileUID(jj)
							sID = m_sSleeveProfileID(jj)
						End If
					Next jj
					
					For jj = 1 To m_iReDrawSleeveOrigin
						sTimeStamp = ARMDIA1.fnGetString(m_sSleeveOriginDBData(jj), 2, ",")
						'UPGRADE_WARNING: Couldn't resolve default property of object xySleeveOrigin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If sTimeStamp = sTimeStampTop Then xySleeveOrigin = m_xySleeveOrigin(jj)
					Next jj
					
					If iProfileUID <> -1 Then
						PR_DrawRaglan(m_sSleeveTopDBData(ii), xySleeveOrigin, iProfileUID, sID)
					End If
					
				Next ii
				
			Else
				'Give a warning
			End If
			
			'Delete these entities and others
			PR_DoDeleteEntityByUID_List()
			
			'Start the mesh drawing programme as last action
			If m_bExecuteMeshDraw = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshvest" & QCQ & "normal" & QQ & " );")
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
			
		End If
		
		'Update the Database
		'    PR_DrawMarker xyNewFrontNeckOnRaglan
		'PR_DrawText "xyNewFrontNeckOnRaglan", xyNewFrontNeckOnRaglan, .1
		'
		'    PR_DrawMarker xyNewBackNeckOnRaglan
		'PR_DrawText "xyNewBackNeckOnRaglan", xyNewBackNeckOnRaglan, .1
		'
		'    PR_DrawMarker xyBackNeckOnRaglan
		'PR_DrawText "xyBackNeckOnRaglan", xyBackNeckOnRaglan, .1
		'
		'    PR_DrawMarker xyBackNeckOnCL
		'PR_DrawText "xyBackNeckOnCL", xyBackNeckOnCL, .1
		'
		'    PR_DrawMarker xyNewBackNeckCen
		'PR_DrawText "xyNewBackNeckCen", xyNewBackNeckCen, .1
		'
		'    PR_DrawMarker xyBackNeckOnCL
		'PR_DrawText "xyBackNeckOnCL", xyBackNeckOnCL, .1
		
		
	End Sub
	
	Private Sub PR_DeleteEntityByUID(ByRef iEntityUID As Short, ByRef sCurveType As Object)
		Dim sDeleteText As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sCurveType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sDeleteText = "DELETED-" & sCurveType
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("hEnt = UID(" & QQ & "find" & QC & iEntityUID & ");")
		BDYUTILS.PR_PutLine("if (hEnt){")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("    SetDBData(hEnt," & QQ & "curvetype" & QCQ & sDeleteText & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("    SetEntityData(hEnt," & QQ & "layer" & QC & "hLayerConstruct);")
		BDYUTILS.PR_PutLine("    }")
		
	End Sub
	
	Private Sub PR_DoDeleteEntityByUID_List()
		Dim ii As Object
		For ii = 1 To m_iDeletedUID_Count
			'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PR_DeleteEntityByUID(m_iDeletedUID_List(ii), m_sDeletedUID_List(ii))
		Next ii
	End Sub
	
	Private Sub PR_DrawRaglan(ByVal sSleeveData As String, ByRef xyOrigin As xy, ByRef iProfileUID As Short, ByRef sID As String)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to add a Raglan to the end of the sleeve
		'profile based on the data given in the variable sVestRaglan
		
		Dim sAxillaType, sString As String
		Dim iPos As Short
		Dim sVestID, sSide As String
		Dim iMeshPos As Short
		
		
		If VB.Right(sSleeveData, 5) = "Right" Then
			sSide = "Right"
			sAxillaType = m_sRtAxillaType
		End If
		
		If VB.Right(sSleeveData, 4) = "Left" Then
			sSide = "Left"
			sAxillaType = m_sLtAxillaType
		End If
		
		'    sAxillaType = FN_EscapeQuotesInString(fnGetString(sSleeveData, 5, ","))
		'    sVestID = fnGetString(sSleeveData, 1, ",")
		
		'Check that an axilla has been given
		If sAxillaType = "None" Or sAxillaType = "" Or sAxillaType = "Sleeveless" Then Exit Sub
		
		'Data from vest used in drawing the sleeve
		If sSide = "Left" Then
			BDYUTILS.PR_PutNumberAssign("nAxillaFrontNeckRad", m_nAxillaFrontNeckRad)
			BDYUTILS.PR_PutNumberAssign("nAxillaBackNeckRad", m_nAxillaBackNeckRad)
			BDYUTILS.PR_PutNumberAssign("nShoulderToBackRaglan", m_nShoulderToBackRaglan)
		Else
			BDYUTILS.PR_PutNumberAssign("nAxillaFrontNeckRad", m_nAFNRadRight)
			BDYUTILS.PR_PutNumberAssign("nAxillaBackNeckRad", m_nABNRadRight)
			BDYUTILS.PR_PutNumberAssign("nShoulderToBackRaglan", m_nSBRaglanRight)
		End If
		
		'Load subroutines etc but ony once
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\SR_DEF.D;")
		
		'Draw for the particular axilla type
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sAxillaType = " & QQ & FN_EscapeQuotesInString(sAxillaType) & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sVestID = " & QQ & sVestID & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sData = " & QQ & FN_EscapeQuotesInString(sSleeveData) & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sSleeve = " & QQ & sSide & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sID = " & QQ & sID & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hSleeveProfile = UID(" & QQ & "find" & QC & iProfileUID & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyOrigin.X= " & xyOrigin.X & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyOrigin.Y= " & xyOrigin.Y & ";")
		
		'Initialise raglan drawing
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\SR_INIT.D;")
		
		
		Select Case sAxillaType
			Case "Open", "Lining"
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\SR_OPEN.D;")
			Case "Mesh"
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "nAge = " & CType(Me.Controls("txtAge"), Object).Text & ";")
				If sSide = "Left" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nMeshLength = " & Val(ARMDIA1.fnGetString(m_sMeshData, 2, ",")) & ";")
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nDistanceAlongRaglan = " & Val(ARMDIA1.fnGetString(m_sMeshData, 1, ",")) & ";")
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nMeshLength = " & Val(ARMDIA1.fnGetString(m_sMeshData, 4, ",")) & ";")
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nDistanceAlongRaglan = " & Val(ARMDIA1.fnGetString(m_sMeshData, 3, ",")) & ";")
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\SR_MESH.D;")
				m_bExecuteMeshDraw = True
				
			Case Else 'Regular
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\SR_REGLR.D;")
				
		End Select
		
		'Close raglan drawing
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_CLOSE.D;")
		
	End Sub

    Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.ComboBox, ByRef sFileName As String)
        'General procedure to create the list section of
        'a combo box reading the data from a file

        Dim sLine As String
        Dim fFileNum As Short

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48, "CAD - Glove Dialogue")
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, OpenMode.Input)
        Do While Not EOF(fFileNum)
            sLine = LineInput(fFileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Combo_Name.AddItem(sLine)
            Combo_Name.Items.Add(sLine)
        Loop
        FileClose(fFileNum)

    End Sub

    Private Sub PR_GetCurveFromFile(ByVal fFileNum As Short, ByRef Profile As curve, ByRef xyLowestYpoint As xy, ByRef xyHighestYpoint As xy)
		
		
		Dim ii As Short
		
		Input(fFileNum, Profile.n)
		For ii = 1 To Profile.n
			Input(fFileNum, Profile.X(ii))
			Input(fFileNum, Profile.Y(ii))
		Next ii
		
		'This is done here as it is convienient to do so.
		If Profile.Y(1) > Profile.Y(Profile.n) Then
			ARMDIA1.PR_MakeXY(xyHighestYpoint, Profile.X(1), Profile.Y(1))
			ARMDIA1.PR_MakeXY(xyLowestYpoint, Profile.X(Profile.n), Profile.Y(Profile.n))
		Else
			ARMDIA1.PR_MakeXY(xyHighestYpoint, Profile.X(Profile.n), Profile.Y(Profile.n))
			ARMDIA1.PR_MakeXY(xyLowestYpoint, Profile.X(1), Profile.Y(1))
		End If
		
	End Sub
	
	Private Sub PR_UpdateDB()
		'Procedure called from
		'    PR_CreateMacro_Save
		'and
		'    PR_CreateMacro_Draw
		'
		'Used to stop duplication on code
		
		Dim sSymbol As String
		
		sSymbol = "vestbody"
		
		'Use existing symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("hBody = UID (" & QQ & "find" & QC & m_iVestBodyUID & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BDYUTILS.PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
		
		'Update the VESTBODY Box symbol
		BDYUTILS.PR_PutLine("hEnt = hBody;")
		BDYUTILS.PR_AddDBValueToLast("AFNRadRight", Str(m_nAFNRadRight))
		BDYUTILS.PR_AddDBValueToLast("ABNRadRight", Str(m_nABNRadRight))
		BDYUTILS.PR_AddDBValueToLast("SBRaglanRight", Str(m_nSBRaglanRight))
		BDYUTILS.PR_AddDBValueToLast("AxillaFrontNeckRad", Str(m_nAxillaFrontNeckRad))
		BDYUTILS.PR_AddDBValueToLast("AxillaBackNeckRad", Str(m_nAxillaBackNeckRad))
		BDYUTILS.PR_AddDBValueToLast("ShoulderToBackRaglan", Str(m_nShoulderToBackRaglan))
		
		
		'   PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'It is assumed that the link open from Drafix has failed
        'Therefor we "End" here
        Return
    End Sub
	
	Private Sub txtDim_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDim.Enter
		Dim Index As Short = txtDim.GetIndex(eventSender)
		CADGLOVE1.PR_Select_Text(txtDim(Index))
	End Sub
	
	Private Sub txtDim_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDim.Leave
		Dim Index As Short = txtDim.GetIndex(eventSender)
		Dim nLen As Double
		Dim sString As String
		
		'Use this function to test validity of entered dimension
		nLen = CADGLOVE1.FN_InchesValue(txtDim(Index))
		'    If nLen > 0 Then
		'        sString = Trim$(fnInchesToText(nLen))
		'        If Left$(sString, 1) = "-" Then sString = Mid$(sString, 2)
		'        lblDim(Index).Caption = sString
		'    Else
		'        lblDim(Index).Caption = ""
		'    End If
		
	End Sub
End Class
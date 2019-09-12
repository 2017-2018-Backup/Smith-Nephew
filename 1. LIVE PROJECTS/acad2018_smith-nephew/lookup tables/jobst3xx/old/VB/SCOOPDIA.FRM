VERSION 5.00
Begin VB.Form scoopdia 
   Caption         =   "VEST Side Scoop Neck - Dialogue"
   ClientHeight    =   2400
   ClientLeft      =   2085
   ClientTop       =   2835
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "scoopdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   3375
   Begin VB.TextBox txtDim 
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   2
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtMeshData 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   17
      Text            =   "1.9434646536216,2.8,,"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtRtAxillaType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   16
      Text            =   "Open"
      Top             =   2100
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtLtAxillaType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   15
      Text            =   "Open"
      Top             =   1740
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtSBRaglanRight 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   14
      Text            =   " .604812479758342"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtAxillaBackNeckRad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Text            =   " 8.55764708056567"
      Top             =   1740
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtABNRadRight 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Text            =   " 8.55764708056567"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtAFNRadRight 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Text            =   " 7.4745475498778"
      Top             =   2940
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtShoulderToBackRaglan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   10
      Text            =   " .604812479758342"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtAxillaFrontNeckRad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Text            =   " 7.4745475498778"
      Top             =   2100
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Text            =   "25"
      Top             =   960
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Text            =   "3423VEST/NECKRight"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtScoopDataFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Text            =   "C:\JOBST\SCOOPNCK.DAT"
      Top             =   960
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtDim 
      Height          =   315
      Index           =   0
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   360
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1275
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Back Scoop Dim."
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDim 
      Caption         =   "Inches"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   18
      Top             =   960
      Width           =   675
   End
   Begin VB.Label lblDim 
      Caption         =   "Inches"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   420
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Front Scoop Dim."
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   420
      Width           =   1575
   End
End
Attribute VB_Name = "scoopdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

   'mODULE LEVEL VARIABLES
    Dim m_sDialogueTitle As String

   'Vest details
    Dim xyBackNeckConstruct         As xy
    Dim m_iBackNeckConstructUID     As Integer
    Dim m_iVestBodyUID              As Integer
    Dim m_iVestRightRaglanUID       As Integer
    Dim m_iVestLeftRaglanUID        As Integer
    Dim m_iVestBackNeckUID          As Integer
    Dim m_iVestFrontNeckUID         As Integer
    Dim m_iVestLeftOpenAxillaMarkerUID         As Integer
    Dim m_iVestRightOpenAxillaMarkerUID        As Integer

    Dim m_VestLeftRaglan            As curve
    Dim m_NewVestLeftRaglan         As curve
    Dim m_ConstructLeftRaglan       As curve
    Dim m_VestRightRaglan           As curve
    Dim m_NewVestRightRaglan        As curve
    Dim m_ConstructRightRaglan      As curve
    Dim m_sVestRightRaglanDBData    As String
    Dim m_sVestLeftRaglanDBData     As String

    Dim xyVestLeftRaglanAtNeck      As xy
    Dim xyVestLeftRaglanAtAxilla    As xy
    Dim xyVestLeftRaglanStart       As xy
    Dim xyVestRightRaglanAtNeck     As xy
    Dim xyVestRightRaglanAtAxilla   As xy
    Dim xyVestRightRaglanStart      As xy
    Dim xyVestLeftopenaxillamarker  As xy
    Dim xyVestRightopenaxillamarker  As xy

    Dim m_FrontNeckProfile          As curve
    Dim m_NewFrontNeckProfile       As curve
    Dim m_bFrontNeckIsCurve         As Integer
    Dim m_nFrontNeckRadius          As Double
    Dim xyFrontNeckCen              As xy
    Dim xyNewFrontNeckCen           As xy
    Dim xyFrontNeckOnRaglan         As xy
    Dim xyFrontNeckOnCL             As xy
    Dim xyNewFrontNeckOnRaglan      As xy

    Dim m_BackNeckProfile           As curve
    Dim m_NewBackNeckProfile        As curve
    Dim m_bBackNeckIsCurve          As Integer
    Dim m_nBackNeckRadius           As Double
    Dim xyBackNeckCen               As xy
    Dim xyNewBackNeckCen            As xy
    Dim xyBackNeckOnRaglan          As xy
    Dim xyBackNeckOnCL              As xy
    Dim xyNewBackNeckOnRaglan       As xy

    Dim m_iReDrawSleeveTop          As Integer
    Dim m_iReDrawSleeveProfile      As Integer
    Dim m_iReDrawSleeveOrigin       As Integer

    Dim m_iAxilla                   As Integer

    Dim m_nAFNRadRight              As Double
    Dim m_nABNRadRight              As Double
    Dim m_nSBRaglanRight            As Double
    Dim m_nAxillaBackNeckRad        As Double
    Dim m_nAxillaFrontNeckRad       As Double
    Dim m_nShoulderToBackRaglan     As Double

    Const MAX_SLEEVE_ReDRAW = 10
    
    Dim m_sSleeveProfileDBData(1 To MAX_SLEEVE_ReDRAW)    As String
    Dim m_sSleeveOriginDBData(1 To MAX_SLEEVE_ReDRAW)     As String
    Dim m_sSleeveTopDBData(1 To MAX_SLEEVE_ReDRAW)        As String
    Dim m_sSleeveProfileID(1 To MAX_SLEEVE_ReDRAW)        As String
    Dim m_iSleeveProfileUID(1 To MAX_SLEEVE_ReDRAW)       As Integer
    Dim m_iSleeveOriginUID(1 To MAX_SLEEVE_ReDRAW)        As Integer
    
    Dim m_xySleeveOrigin(1 To MAX_SLEEVE_ReDRAW)          As xy

            
    Const MAX_ENTITIES_ON_DELETE_LIST = 200

    Dim m_iDeletedUID_List(1 To MAX_ENTITIES_ON_DELETE_LIST) As Integer
    Dim m_sDeletedUID_List(1 To MAX_ENTITIES_ON_DELETE_LIST) As String
    Dim m_iDeletedUID_Count                                  As Integer


    Dim m_bSleeve                       As Integer
    Dim m_bVest                         As Integer

    Dim m_bExecuteMeshDraw              As Integer

    Dim m_sRtAxillaType                 As String
    Dim m_sLtAxillaType                 As String
    Dim m_sMeshData                     As String
    Option Explicit

Private Sub Cancel_Click()
    
    End

End Sub

Private Function FN_CalcCirCurveInt(Profile As curve, xyCen As xy, nRadius As Double, ProfileLargePart As curve, ProfileSmallPart As curve, xyIntersection As xy) As Integer
    
    Dim xyPt1               As xy
    Dim xyPt2               As xy
    Dim xyInt               As xy
    Dim ii                  As Integer
    Dim bIntersectionFound  As Integer
    Dim ProfilePart1        As curve
    Dim ProfilePart2        As curve


    bIntersectionFound = False
    FN_CalcCirCurveInt = False
    ProfilePart2.n = 0

    If Profile.n < 2 Then Exit Function
    
    PR_MakeXY xyPt1, Profile.X(1), Profile.Y(1)

    ProfilePart1.X(1) = Profile.X(1)
    ProfilePart1.Y(1) = Profile.Y(1)
    ProfilePart1.n = 1
    
    For ii = 2 To Profile.n
        PR_MakeXY xyPt2, Profile.X(ii), Profile.Y(ii)
        If Not bIntersectionFound Then bIntersectionFound = FN_CirLinInt(xyPt1, xyPt2, xyCen, nRadius, xyInt)
        If bIntersectionFound Then
           'Check that we have the upper intersection
            If xyInt.Y < xyCen.Y Then
               'If we don't have the upper intersection then we check that the intersection
               'is not on the same segment.  This is due to the nature of the function FN_CirLinInt
               'that returns only the first intesection
                bIntersectionFound = False
                If xyCen.Y < xyPt1.Y Then
                    bIntersectionFound = FN_CirLinInt(xyPt1, xyCen, xyCen, nRadius, xyInt)
                Else
                    bIntersectionFound = FN_CirLinInt(xyPt2, xyCen, xyCen, nRadius, xyInt)
                End If
            End If
            If bIntersectionFound Then
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
        
        xyPt1 = xyPt2

    Next ii

    If ProfilePart1.n > ProfilePart2.n Then
        ProfileLargePart = ProfilePart1
        ProfileSmallPart = ProfilePart2
    Else
        ProfileLargePart = ProfilePart2
        ProfileSmallPart = ProfilePart1
    End If
    If bIntersectionFound Then FN_CalcCirCurveInt = True

End Function

Private Function FN_EscapeQuotesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeQuotesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = QQ Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeQuotesInString = sEscapedString

End Function

Private Function FN_OpenAxillaMarker(Profile As curve, xyIntersection As xy, aAngle As Double) As Integer

    Dim xyPt1               As xy
    Dim xyPt2               As xy
    Dim xyInt               As xy
    Dim ii                  As Integer
    Dim iError              As Integer
    Dim nLength             As Double
    Dim nRequired           As Double


    FN_OpenAxillaMarker = False
    
    If Profile.n < 2 Then Exit Function
    
    PR_MakeXY xyPt1, Profile.X(1), Profile.Y(1)
    nLength = 0
    For ii = 2 To Profile.n
        PR_MakeXY xyPt2, Profile.X(ii), Profile.Y(ii)
        nLength = nLength + FN_CalcLength(xyPt1, xyPt2)
        xyPt1 = xyPt2
    Next ii
    
    nRequired = nLength / 3

    PR_MakeXY xyPt1, Profile.X(1), Profile.Y(1)
    nLength = 0
    For ii = 2 To Profile.n
        PR_MakeXY xyPt2, Profile.X(ii), Profile.Y(ii)
        nLength = nLength + FN_CalcLength(xyPt1, xyPt2)

        If nLength >= nRequired Then
            nLength = nLength - nRequired
            If nLength > 0 Then
                iError = FN_CirLinInt(xyPt1, xyPt2, xyPt1, nLength, xyIntersection)
            Else
                xyIntersection = xyPt2
            End If
            FN_OpenAxillaMarker = True
            aAngle = FN_CalcAngle(xyPt1, xyPt2) - 90
            Exit For
        End If
        xyPt1 = xyPt2
    Next ii


End Function


Private Function FN_ValidateData() As Integer
   'This function is used only to make gross checks
   'for missing data.
   'It does not perform any sensibility checks on the
   'data
    Dim sError$, ii%, nn%
    
   'Initialise
    FN_ValidateData = False
    sError = ""

   'Dimensions
    If Val(txtDim(0).Text) = 0 And Val(txtDim(1).Text) = 0 Then
        sError = sError & "Missing dimensions for Side Scoop!" & NL
    End If
    
    If Val(txtDim(0).Text) > 0 And Val(txtDim(0).Text) < 0.1 Then sError = sError & "Front Side Scoop must be greater than 1/10th Inch!" & NL
    If Val(txtDim(1).Text) > 0 And Val(txtDim(1).Text) < 0.1 Then sError = sError & "Back Side Scoop must be greater than 1/10th Inch!" & NL
    
   'Check for errors
    If sError <> "" And m_bVest Then
        MsgBox sError, 16, m_sDialogueTitle
        FN_ValidateData = False
    Else
        FN_ValidateData = True
    End If

End Function

Private Sub Form_LinkClose()

    Dim ii                  As Integer
    Dim iEntityUID          As Integer
    Dim fFileNum            As Integer
    Dim sEntityname         As String
    Dim sEntityType         As String
    Dim sTemplateToScoop    As String
    Dim sEntityID           As String
    Dim aStartAngle         As Double
    Dim aDeltaAngle         As Double

   'Dummys
    Dim Dummy               As curve
    Dim xyPt                As xy
    Dim xyPt1               As xy
    Dim xyPt2               As xy
    Dim nRadius             As Double
    Dim sDummy              As String


   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

   'Check that a Data File name has been passed
    If txtScoopDataFile = "" Then
        MsgBox "No VEST details have been found in drawing!", 16, m_sDialogueTitle
        End
    End If
    
   'Get ID to mark the nee draw entities
    g_sID = txtID.Text

   'Load from file the vest details
   'Setup error handling
    On Error GoTo ErrorStarting

   'Check that Data file exist
   'Note that this is a fixed file name
    If FileLen(txtScoopDataFile.Text) = 0 Then
        MsgBox "Can't open the file " & txtScoopDataFile.Text & ".  Unable to modify the VEST Neck.", 48, m_sDialogueTitle
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
    Open txtScoopDataFile.Text For Input Access Read As fFileNum

    Input #fFileNum, sTemplateToScoop
    If sTemplateToScoop = "SLEEVE" Then
        m_bSleeve = True
    ElseIf sTemplateToScoop = "VEST" Then
        m_bVest = True
    Else
        GoTo ErrorStarting
    End If
        
    Do While Not EOF(fFileNum)
        Input #fFileNum, sEntityname
        Input #fFileNum, sEntityType
        Input #fFileNum, iEntityUID
        Line Input #fFileNum, sEntityID
        Select Case sEntityname
         Case "vestbackneckconstruct"
            m_iBackNeckConstructUID = iEntityUID
            Input #fFileNum, xyBackNeckConstruct.X, xyBackNeckConstruct.Y

         Case "vestbody"
            m_iVestBodyUID = iEntityUID

         Case "vestfrontneck"
            If sEntityType = "curve" Or sEntityType = "polyline" Then
               'It is a curve
                m_bFrontNeckIsCurve = True
                PR_GetCurveFromFile fFileNum, m_FrontNeckProfile, xyFrontNeckOnCL, xyFrontNeckOnRaglan
            Else
               'Assume that it is an Arc
                m_bFrontNeckIsCurve = False
                Input #fFileNum, xyFrontNeckCen.X, xyFrontNeckCen.Y, m_nFrontNeckRadius, aStartAngle, aDeltaAngle
                PR_CalcPolar xyFrontNeckCen, aStartAngle, m_nFrontNeckRadius, xyFrontNeckOnRaglan
                PR_CalcPolar xyFrontNeckCen, aStartAngle + aDeltaAngle, m_nFrontNeckRadius, xyFrontNeckOnCL
            End If
            m_iVestFrontNeckUID = iEntityUID
         
         Case "vestprofile"
           'Read as dummy
            PR_GetCurveFromFile fFileNum, Dummy, xyPt, xyPt

         Case "vestLeftopenaxillamarker"
            Input #fFileNum, xyVestLeftopenaxillamarker.X, xyVestLeftopenaxillamarker.Y
            If m_iVestLeftOpenAxillaMarkerUID = -1 Then
               'Save only the first one
                m_iVestLeftOpenAxillaMarkerUID = iEntityUID
            Else
               'Delete any extranous ones
                PR_AddEntityToDeletedUID_List iEntityUID, sEntityname
            End If

         Case "vestRightopenaxillamarker"
            Input #fFileNum, xyVestRightopenaxillamarker.X, xyVestRightopenaxillamarker.Y
            If m_iVestRightOpenAxillaMarkerUID = -1 Then
               'Save only the first one
                m_iVestRightOpenAxillaMarkerUID = iEntityUID
            Else
               'Delete any extranous ones
                PR_AddEntityToDeletedUID_List iEntityUID, sEntityname
            End If

         Case "vestLeftraglan"
            PR_GetCurveFromFile fFileNum, m_VestLeftRaglan, xyVestLeftRaglanAtNeck, xyVestLeftRaglanStart
            If m_iVestLeftRaglanUID = -1 Then
               'Save only the first one
                m_iVestLeftRaglanUID = iEntityUID
                m_sVestLeftRaglanDBData = sEntityID
            Else
               'Delete any extranous ones
                PR_AddEntityToDeletedUID_List iEntityUID, sEntityname
            End If
         
         Case "vestRightraglan"
            PR_GetCurveFromFile fFileNum, m_VestRightRaglan, xyVestRightRaglanAtNeck, xyVestRightRaglanStart
            If m_iVestRightRaglanUID = -1 Then
               'Save only the first one
                m_iVestRightRaglanUID = iEntityUID
                m_sVestRightRaglanDBData = sEntityID
            Else
               'Delete any extranous ones
                PR_AddEntityToDeletedUID_List iEntityUID, sEntityname
            End If

         Case "vestRightaxillaconstruct"
            Input #fFileNum, xyPt1.X, xyPt1.Y, xyPt2.X, xyPt2.Y
            If xyPt1.Y < xyPt2.Y Then
                xyVestRightRaglanAtAxilla = xyPt2
            Else
                xyVestRightRaglanAtAxilla = xyPt1
            End If
            m_iAxilla = m_iAxilla + 1

         Case "vestLeftaxillaconstruct"
            Input #fFileNum, xyPt1.X, xyPt1.Y, xyPt2.X, xyPt2.Y
            If xyPt1.Y < xyPt2.Y Then
                xyVestLeftRaglanAtAxilla = xyPt2
            Else
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
            PR_GetCurveFromFile fFileNum, Dummy, xyPt, xyPt
            PR_AddEntityToDeletedUID_List iEntityUID, sEntityname

         Case "sleeveprofile"
           'Note we may have multiple sleeves based on the one vest
            m_iReDrawSleeveProfile = m_iReDrawSleeveProfile + 1
            m_sSleeveProfileDBData(m_iReDrawSleeveProfile) = sEntityID
            If InStr(sEntityID, "RightProfile") > 0 Then
                m_sSleeveProfileID(m_iReDrawSleeveProfile) = Mid$(sEntityID, 1, Len(sEntityID) - 12)
            Else
                m_sSleeveProfileID(m_iReDrawSleeveProfile) = Mid$(sEntityID, 1, Len(sEntityID) - 11)
            End If
            m_iSleeveProfileUID(m_iReDrawSleeveProfile) = iEntityUID
            PR_GetCurveFromFile fFileNum, Dummy, xyPt, xyPt

         Case "sleeveoriginmark"
            m_iReDrawSleeveOrigin = m_iReDrawSleeveOrigin + 1
            m_sSleeveOriginDBData(m_iReDrawSleeveOrigin) = sEntityID
            m_iSleeveOriginUID(m_iReDrawSleeveOrigin) = iEntityUID
            Input #fFileNum, xyPt.X, xyPt.Y
            m_xySleeveOrigin(m_iReDrawSleeveOrigin) = xyPt
         
         Case "vestbackneck"
            If sEntityType = "curve" Or sEntityType = "polyline" Then
               'It is a curve
                m_bBackNeckIsCurve = True
                PR_GetCurveFromFile fFileNum, m_BackNeckProfile, xyBackNeckOnCL, xyBackNeckOnRaglan
            Else
               'Assume that it is an Arc
                m_bBackNeckIsCurve = False
                Input #fFileNum, xyBackNeckCen.X, xyBackNeckCen.Y, m_nBackNeckRadius, aStartAngle, aDeltaAngle
                PR_CalcPolar xyBackNeckCen, aStartAngle, m_nBackNeckRadius, xyBackNeckOnRaglan
                PR_CalcPolar xyBackNeckCen, aStartAngle + aDeltaAngle, m_nBackNeckRadius, xyBackNeckOnCL
            End If
            m_iVestBackNeckUID = iEntityUID
         
         Case Else
           'Input the curve depending on it's class into dummy variables.
           'Add them to the deleted list as we know that these entites are part of sleeve
           'raglans that will be redrawn
            
            Select Case sEntityType
                Case "curve", "polyline"
                    PR_GetCurveFromFile fFileNum, Dummy, xyPt, xyPt
                Case "arc"
                    Input #fFileNum, xyPt.X, xyPt.Y, nRadius, aStartAngle, aDeltaAngle
                Case "line"
                    Input #fFileNum, xyPt.X, xyPt.Y, xyPt.X, xyPt.Y
                Case "marker"
                    Input #fFileNum, xyPt.X, xyPt.Y
            End Select

            If Not sEntityname = "NoCurveTypeFound" Then PR_AddEntityToDeletedUID_List iEntityUID, sEntityname

        End Select
        
    Loop 'End of DO While

   'Close file
    Close #fFileNum

    m_sLtAxillaType = txtLtAxillaType
    m_sRtAxillaType = txtRtAxillaType
    m_sMeshData = txtMeshData

    m_nAxillaBackNeckRad = Val(txtAxillaBackNeckRad)
    m_nAxillaFrontNeckRad = Val(txtAxillaFrontNeckRad)
    m_nShoulderToBackRaglan = Val(txtShoulderToBackRaglan)
    m_nAFNRadRight = Val(txtAFNRadRight)
    m_nABNRadRight = Val(txtABNRadRight)
    m_nSBRaglanRight = Val(txtSBRaglanRight)

    If m_bSleeve Then
        OK_Click
    Else
        Show
    End If

   'Change pointer to normal
    Screen.MousePointer = 0

    Exit Sub

ErrorStarting:
    
    MsgBox "Error Reading Input file"
    End

End Sub



Private Sub Form_Load()
    Dim ii%

    Hide

   'Start a timer
   'The Timer is disabled in LinkClose
   'If after 6 seconds the timer event will "End" the programme
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 6000  'Approx 6 Seconds
    Timer1.Enabled = True

   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

   'Check if a previous instance is running
   'If it is warn user and exit
    m_sDialogueTitle = "VEST Side Scoop Neck"

    If App.PrevInstance Then
        MsgBox m_sDialogueTitle & " is already running!" & NL & "Use ALT-TAB and Cancel it.", 16, m_sDialogueTitle
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - scoopdia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - scoopdia.Height) / 2  ' Center form vertically.
    
    Set MainForm = scoopdia

    g_nUnitsFac = 1
    g_sPathJOBST = fnPathJOBST()

   'Clear fields
   'Circumferences and lengths
    txtDim(0) = ""
    txtScoopDataFile = ""
    txtID = ""
    txtAxillaBackNeckRad = ""
    txtAxillaFrontNeckRad = ""
    txtShoulderToBackRaglan = ""
    txtAFNRadRight = ""
    txtABNRadRight = ""
    txtSBRaglanRight = ""
    
'Debug "Read from file"
'    txtScoopDataFile = "C:\JOBST\SCOOPNCK.DAT"
'    Form_LinkClose

End Sub

Private Sub OK_Click()
    Dim sTask$
   'Don't allow multiple clicking
   '
    OK.Enabled = False
    If FN_ValidateData() Then
        Screen.MousePointer = 11
        Hide

        PR_CreateMacro_Draw "c:\jobst\draw.d"
        
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            AppActivate sTask
                    SendKeys "@c:\jobst\draw.d{enter}", True
            Screen.MousePointer = 0
            End
        Else
            MsgBox "Can't find a Drafix Drawing to update!", 16, m_sDialogueTitle
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub PR_AddEntityToDeletedUID_List(iEntityUID%, sEntity$)
    m_iDeletedUID_Count = m_iDeletedUID_Count + 1
    m_iDeletedUID_List(m_iDeletedUID_Count) = iEntityUID
    m_sDeletedUID_List(m_iDeletedUID_Count) = sEntity
End Sub

Private Sub PR_CreateMacro_Draw(sFile As String)
   'Since this is a fairly straight forward module we
   'will keep calculation and drawing in the same procedure
   '
   'The assumption is that all data has been validated before
   'this procedure is called

    Dim iError          As Integer
    Dim ii%, jj%
    Dim nA              As Double
    Dim nB              As Double
    Dim nC              As Double
    Dim aA              As Double
    Dim xyTmp           As xy
    Dim xySleeveOrigin  As xy
    Dim xyVestRaglanAtAxilla As xy

    Dim iProfileUID     As Integer
    Dim iOriginUID      As Integer
    Dim iLowestOpenAxillaMarkerUID     As Integer
    Dim iHighestOpenAxillaMarkerUID    As Integer
    Dim sTimeStampTop   As String
    Dim sTimeStamp      As String
    Dim sID             As String
    Dim sHighVestDBData     As String
    Dim sHighVestCurveType  As String
    Dim sHighestSide        As String
    Dim sLowestSide         As String
    Dim sLowVestDBData      As String
    Dim sLowVestCurveType   As String


    Dim DummyProfile        As curve
    Dim HighestVestRaglan   As curve
    Dim LowestVestRaglan    As curve
    Dim ConstructRaglan     As curve
    Dim NewVestRaglan       As curve
    Dim NewLowVestRaglan    As curve
    
    Dim nFrontSideScoop     As Double
    Dim nBackSideScoop      As Double

    Dim xyOpenAxillaMaker   As xy
    Dim aAngle              As Double
    


    nFrontSideScoop = Val(txtDim(0).Text)
    nBackSideScoop = Val(txtDim(1).Text)

    If m_bVest Then
       'Where we have two raglan curves then we use the one with the highest axilla to
       'Calculate the new front neck
        If m_iAxilla = 2 Then
            If xyVestLeftRaglanAtAxilla.X > xyVestRightRaglanAtAxilla.X Then
                sHighestSide = "Left"
                sLowestSide = "Right"
                HighestVestRaglan = m_VestLeftRaglan
                LowestVestRaglan = m_VestRightRaglan
                iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
                iLowestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
                xyVestRaglanAtAxilla = xyVestLeftRaglanAtAxilla
                sHighVestCurveType = "vestLeftraglan"
                sLowVestCurveType = "vestRightraglan"
                sLowVestDBData = m_sVestLeftRaglanDBData
                sHighVestDBData = m_sVestRightRaglanDBData
            Else
                sHighestSide = "Right"
                sLowestSide = "Left"
                HighestVestRaglan = m_VestRightRaglan
                LowestVestRaglan = m_VestLeftRaglan
                iHighestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
                iLowestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
                xyVestRaglanAtAxilla = xyVestRightRaglanAtAxilla
                sHighVestCurveType = "vestRightraglan"
                sLowVestCurveType = "vestLeftraglan"
                sLowVestDBData = m_sVestRightRaglanDBData
                sHighVestDBData = m_sVestLeftRaglanDBData
            End If
        Else
                sHighestSide = "Left"
                HighestVestRaglan = m_VestLeftRaglan
                iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
                If m_iVestLeftOpenAxillaMarkerUID > -1 Then iHighestOpenAxillaMarkerUID = m_iVestLeftOpenAxillaMarkerUID
                If m_iVestRightOpenAxillaMarkerUID > -1 Then iHighestOpenAxillaMarkerUID = m_iVestRightOpenAxillaMarkerUID
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
                nC = FN_CalcLength(xyNewFrontNeckOnRaglan, xyFrontNeckOnCL) / 2
                nB = Sqr(nA ^ 2 - nC ^ 2)
                aA = FN_CalcAngle(xyFrontNeckOnCL, xyNewFrontNeckOnRaglan)
                PR_CalcPolar xyFrontNeckOnCL, aA, nC, xyTmp
                PR_CalcPolar xyTmp, aA - 90, nB, xyNewFrontNeckCen
                
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
                nC = FN_CalcLength(xyNewBackNeckOnRaglan, xyBackNeckOnCL) / 2
                nB = Sqr(nA ^ 2 - nC ^ 2)
                aA = FN_CalcAngle(xyBackNeckOnCL, xyNewBackNeckOnRaglan)
                PR_CalcPolar xyBackNeckOnCL, aA, nC, xyTmp
                PR_CalcPolar xyTmp, aA - 90, nB, xyNewBackNeckCen
            End If
        End If

       'Highest
        If nFrontSideScoop > 0 Then m_nAxillaFrontNeckRad = FN_CalcLength(xyVestRaglanAtAxilla, xyNewFrontNeckOnRaglan)
        If nBackSideScoop > 0 Then
            m_nAxillaBackNeckRad = FN_CalcLength(xyVestRaglanAtAxilla, xyNewBackNeckOnRaglan)
            m_nShoulderToBackRaglan = FN_CalcLength(xyBackNeckConstruct, xyNewBackNeckOnRaglan)
        End If
        m_nAFNRadRight = m_nAxillaFrontNeckRad
        m_nABNRadRight = m_nAxillaBackNeckRad
        m_nSBRaglanRight = m_nShoulderToBackRaglan
    
       'adjust if 2 axilla hts
        If m_iAxilla = 2 Then
            If nBackSideScoop > 0 Then iError = FN_CalcCirCurveInt(LowestVestRaglan, xyBackNeckOnRaglan, nBackSideScoop, NewLowVestRaglan, ConstructRaglan, xyTmp)
            If nFrontSideScoop > 0 Then
                m_nAFNRadRight = FN_CalcLength(xyVestRightRaglanAtAxilla, xyNewFrontNeckOnRaglan)
                m_nAxillaFrontNeckRad = FN_CalcLength(xyVestLeftRaglanAtAxilla, xyNewFrontNeckOnRaglan)
            End If
            If nBackSideScoop > 0 Then
                m_nABNRadRight = FN_CalcLength(xyVestRightRaglanAtAxilla, xyNewBackNeckOnRaglan)
                m_nAxillaBackNeckRad = FN_CalcLength(xyVestLeftRaglanAtAxilla, xyNewBackNeckOnRaglan)
                m_nSBRaglanRight = FN_CalcLength(xyBackNeckConstruct, xyNewBackNeckOnRaglan)
            End If
        End If

       fNum = FN_Open(sFile)

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
            PR_SetLayer "Notes"
            If m_bFrontNeckIsCurve Then
                PR_DrawFitted m_NewFrontNeckProfile
            Else
                PR_DrawArc xyNewFrontNeckCen, xyFrontNeckOnCL, xyNewFrontNeckOnRaglan
            End If
            PR_AddDBValueToLast "curvetype", "vestfrontneck"
            PR_AddDBValueToLast "Data", sHighVestDBData
            PR_AddEntityToDeletedUID_List m_iVestFrontNeckUID, "vestfrontneck"
        End If

        If nBackSideScoop > 0 Then

        iError = FN_CalcCirCurveInt(HighestVestRaglan, xyBackNeckOnRaglan, nBackSideScoop, NewVestRaglan, ConstructRaglan, xyNewBackNeckOnRaglan)

            PR_SetLayer "TemplateLeft"
            If m_bBackNeckIsCurve Then
                PR_DrawFitted m_NewBackNeckProfile
            Else
                PR_DrawArc xyNewBackNeckCen, xyBackNeckOnCL, xyNewBackNeckOnRaglan
            End If
            PR_AddDBValueToLast "curvetype", "vestbackneck"
            PR_AddEntityToDeletedUID_List m_iVestBackNeckUID, "vestbackneck"
            PR_AddDBValueToLast "Data", sHighVestDBData
            
            PR_SetLayer "Template" & sHighestSide
            PR_DrawPoly NewVestRaglan
            PR_AddDBValueToLast "curvetype", sHighVestCurveType
            PR_AddDBValueToLast "Data", sHighVestDBData
            If iHighestOpenAxillaMarkerUID > -1 Then
                iError = FN_OpenAxillaMarker(NewVestRaglan, xyOpenAxillaMaker, aAngle)
                PR_SetLayer "Notes"
                PR_DrawMarkerNamed "medarrow", xyOpenAxillaMaker, 0.2, 0.1, aAngle
                PR_AddDBValueToLast "Data", sHighVestDBData
                PR_AddDBValueToLast "curvetype", "vest" & sHighestSide & "openaxillamarker"
            End If
            If m_iAxilla = 2 Then
                PR_SetLayer "Template" & sLowestSide
                PR_DrawPoly NewLowVestRaglan
                PR_AddDBValueToLast "curvetype", sLowVestCurveType
                PR_AddDBValueToLast "Data", sLowVestDBData
                If iLowestOpenAxillaMarkerUID > -1 Then
                    iError = FN_OpenAxillaMarker(NewLowVestRaglan, xyOpenAxillaMaker, aAngle)
                    PR_SetLayer "Notes"
                    PR_DrawMarkerNamed "medarrow", xyOpenAxillaMaker, 0.2, 0.1, aAngle
                    PR_AddDBValueToLast "Data", sLowVestDBData
                    PR_AddDBValueToLast "curvetype", "vest" & sLowestSide & "openaxillamarker"
                End If
            End If
            
            PR_AddEntityToDeletedUID_List m_iVestLeftRaglanUID, "vestLeftraglan"
            PR_AddEntityToDeletedUID_List m_iVestRightRaglanUID, "vestRightraglan"
            If m_iVestLeftOpenAxillaMarkerUID > -1 Then PR_AddEntityToDeletedUID_List m_iVestLeftOpenAxillaMarkerUID, "vestLeftopenaxillamarker"
            If m_iVestRightOpenAxillaMarkerUID > -1 Then PR_AddEntityToDeletedUID_List m_iVestRightOpenAxillaMarkerUID, "vestRightopenaxillamarker"

        End If
       
       'Delete entities and others
        PR_DoDeleteEntityByUID_List

       'Update database
        PR_UpdateDB
        
       Close #fNum


    End If

   'ReDraw sleeves
   'Check that there are enough sleeve bits etc
    If m_bSleeve Then
        m_bExecuteMeshDraw = False
       fNum = FN_Open(sFile)
        If m_iReDrawSleeveTop > 0 And m_iReDrawSleeveTop = m_iReDrawSleeveProfile And m_iReDrawSleeveTop = m_iReDrawSleeveOrigin Then
           'loop through these entities finding the matches
            For ii = 1 To m_iReDrawSleeveTop
                iProfileUID = -1
                sTimeStampTop = fnGetString(m_sSleeveTopDBData(ii), 2, ",")
                
                For jj = 1 To m_iReDrawSleeveProfile
                    sTimeStamp = fnGetString(m_sSleeveProfileDBData(jj), 2, ",")
                    If sTimeStamp = sTimeStampTop Then
                        iProfileUID = m_iSleeveProfileUID(jj)
                        sID = m_sSleeveProfileID(jj)
                    End If
                Next jj
                
                For jj = 1 To m_iReDrawSleeveOrigin
                    sTimeStamp = fnGetString(m_sSleeveOriginDBData(jj), 2, ",")
                    If sTimeStamp = sTimeStampTop Then xySleeveOrigin = m_xySleeveOrigin(jj)
                Next jj
                
                If iProfileUID <> -1 Then
                    PR_DrawRaglan m_sSleeveTopDBData(ii), xySleeveOrigin, iProfileUID, sID
                End If
    
            Next ii
            
        Else
           'Give a warning
        End If
    
       'Delete these entities and others
        PR_DoDeleteEntityByUID_List

       'Start the mesh drawing programme as last action
        If m_bExecuteMeshDraw = True Then
            Print #fNum, "Execute ("; QQ; "application"; QC; "sPathJOBST + "; QQ; "\\raglan\\meshvest"; QCQ; "normal"; QQ; " );"
        End If

       Close #fNum

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

Private Sub PR_DeleteEntityByUID(iEntityUID As Integer, sCurveType)
    Dim sDeleteText As String

    sDeleteText = "DELETED-" & sCurveType

    PR_PutLine "hEnt = UID(" & QQ & "find" & QC & iEntityUID & ");"
    PR_PutLine "if (hEnt){"
    PR_PutLine "    SetDBData(hEnt," & QQ & "curvetype" & QCQ & sDeleteText & QQ & ");"
    PR_PutLine "    SetEntityData(hEnt," & QQ & "layer" & QC & "hLayerConstruct);"
    PR_PutLine "    }"

End Sub

Private Sub PR_DoDeleteEntityByUID_List()
    Dim ii
    For ii = 1 To m_iDeletedUID_Count
        PR_DeleteEntityByUID m_iDeletedUID_List(ii), m_sDeletedUID_List(ii)
    Next ii
End Sub

Private Sub PR_DrawRaglan(ByVal sSleeveData As String, xyOrigin As xy, iProfileUID As Integer, sID As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to add a Raglan to the end of the sleeve
   'profile based on the data given in the variable sVestRaglan
    
    Dim sAxillaType As String, sString As String
    Dim iPos As Integer
    Dim sVestID As String, sSide    As String
    Dim iMeshPos  As Integer
   
    
    If Right$(sSleeveData, 5) = "Right" Then
        sSide = "Right"
        sAxillaType = m_sRtAxillaType
    End If

    If Right$(sSleeveData, 4) = "Left" Then
        sSide = "Left"
        sAxillaType = m_sLtAxillaType
    End If

'    sAxillaType = FN_EscapeQuotesInString(fnGetString(sSleeveData, 5, ","))
'    sVestID = fnGetString(sSleeveData, 1, ",")

   'Check that an axilla has been given
    If sAxillaType = "None" Or sAxillaType = "" Or sAxillaType = "Sleeveless" Then Exit Sub

   'Data from vest used in drawing the sleeve
    If sSide = "Left" Then
        PR_PutNumberAssign "nAxillaFrontNeckRad", m_nAxillaFrontNeckRad
        PR_PutNumberAssign "nAxillaBackNeckRad", m_nAxillaBackNeckRad
        PR_PutNumberAssign "nShoulderToBackRaglan", m_nShoulderToBackRaglan
    Else
        PR_PutNumberAssign "nAxillaFrontNeckRad", m_nAFNRadRight
        PR_PutNumberAssign "nAxillaBackNeckRad", m_nABNRadRight
        PR_PutNumberAssign "nShoulderToBackRaglan", m_nSBRaglanRight
    End If

   'Load subroutines etc but ony once
    Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\SR_DEF.D;"
    
   'Draw for the particular axilla type
    Print #fNum, "sAxillaType = "; QQ; FN_EscapeQuotesInString(sAxillaType); QQ; ";"
    Print #fNum, "sVestID = "; QQ; sVestID; QQ; ";"
    Print #fNum, "sData = "; QQ; FN_EscapeQuotesInString(sSleeveData); QQ; ";"
    Print #fNum, "sSleeve = "; QQ; sSide; QQ; ";"
    Print #fNum, "sID = "; QQ; sID; QQ; ";"
    Print #fNum, "hSleeveProfile = UID("; QQ; "find"; QC; iProfileUID; ");"
    Print #fNum, "xyOrigin.X= "; xyOrigin.X; ";"
    Print #fNum, "xyOrigin.Y= "; xyOrigin.Y; ";"

   'Initialise raglan drawing
    Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\SR_INIT.D;"

              
    Select Case sAxillaType
    Case "Open", "Lining"
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\SR_OPEN.D;"
    Case "Mesh"
        Print #fNum, "nAge = "; scoopdia!txtAge.Text; ";"
        If sSide = "Left" Then
            Print #fNum, "nMeshLength = "; Val(fnGetString(m_sMeshData, 2, ",")); ";"
            Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(m_sMeshData, 1, ",")); ";"
        Else
            Print #fNum, "nMeshLength = "; Val(fnGetString(m_sMeshData, 4, ",")); ";"
            Print #fNum, "nDistanceAlongRaglan = "; Val(fnGetString(m_sMeshData, 3, ",")); ";"
        End If

        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\SR_MESH.D;"
        m_bExecuteMeshDraw = True
    
    Case Else 'Regular
        Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\SR_REGLR.D;"

    End Select

   'Close raglan drawing
    Print #fNum, "@"; g_sPathJOBST; "\RAGLAN\VR_CLOSE.D;"

End Sub

Private Sub PR_GetComboListFromFile(Combo_Name As Control, sFileName As String)
   'General procedure to create the list section of
   'a combo box reading the data from a file
    
    Dim sLine As String, fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48, "CAD - Glove Dialogue"
        Exit Sub
    End If

    Open sFileName For Input As fFileNum
    Do While Not EOF(fFileNum)
        Line Input #fFileNum, sLine
        Combo_Name.AddItem sLine
        Loop
    Close #fFileNum

End Sub

Private Sub PR_GetCurveFromFile(ByVal fFileNum As Integer, Profile As curve, xyLowestYpoint As xy, xyHighestYpoint As xy)
  

    Dim ii  As Integer

    Input #fFileNum, Profile.n
    For ii = 1 To Profile.n
        Input #fFileNum, Profile.X(ii), Profile.Y(ii)
    Next ii

   'This is done here as it is convienient to do so.
    If Profile.Y(1) > Profile.Y(Profile.n) Then
        PR_MakeXY xyHighestYpoint, Profile.X(1), Profile.Y(1)
        PR_MakeXY xyLowestYpoint, Profile.X(Profile.n), Profile.Y(Profile.n)
    Else
        PR_MakeXY xyHighestYpoint, Profile.X(Profile.n), Profile.Y(Profile.n)
        PR_MakeXY xyLowestYpoint, Profile.X(1), Profile.Y(1)
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
    PR_PutLine "hBody = UID (" & QQ & "find" & QC & m_iVestBodyUID & ");"
    PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"
   
   'Update the VESTBODY Box symbol
    PR_PutLine "hEnt = hBody;"
    PR_AddDBValueToLast "AFNRadRight", Str$(m_nAFNRadRight)
    PR_AddDBValueToLast "ABNRadRight", Str$(m_nABNRadRight)
    PR_AddDBValueToLast "SBRaglanRight", Str$(m_nSBRaglanRight)
    PR_AddDBValueToLast "AxillaFrontNeckRad", Str$(m_nAxillaFrontNeckRad)
    PR_AddDBValueToLast "AxillaBackNeckRad", Str$(m_nAxillaBackNeckRad)
    PR_AddDBValueToLast "ShoulderToBackRaglan", Str$(m_nShoulderToBackRaglan)

    
 '   PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"

End Sub

Private Sub Timer1_Timer()
    'It is assumed that the link open from Drafix has failed
    'Therefor we "End" here
     End
End Sub

Private Sub txtDim_GotFocus(Index As Integer)
    PR_Select_Text txtDim(Index)
End Sub

Private Sub txtDim_LostFocus(Index As Integer)
    Dim nLen#
    Dim sString$
    
   'Use this function to test validity of entered dimension
    nLen = FN_InchesValue(txtDim(Index))
'    If nLen > 0 Then
'        sString = Trim$(fnInchesToText(nLen))
'        If Left$(sString, 1) = "-" Then sString = Mid$(sString, 2)
'        lblDim(Index).Caption = sString
'    Else
'        lblDim(Index).Caption = ""
'    End If

End Sub


VERSION 5.00
Begin VB.Form whboddia 
   Caption         =   "Waist Height Body - Data"
   ClientHeight    =   5250
   ClientLeft      =   105
   ClientTop       =   1440
   ClientWidth     =   10455
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
   LinkTopic       =   "whboddia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdTab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   7500
      TabIndex        =   95
      Top             =   5640
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Left            =   6900
      Top             =   5580
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   6900
      TabIndex        =   70
      Top             =   90
      Width           =   3435
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   76
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2580
         TabIndex        =   75
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2400
         TabIndex        =   74
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   73
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   72
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   71
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2040
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   1920
         TabIndex        =   63
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   300
         TabIndex        =   69
         Top             =   1620
         Width           =   375
      End
   End
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   64
      Top             =   6180
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7980
      TabIndex        =   62
      Top             =   6180
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidWHBody 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6000
      TabIndex        =   60
      Top             =   6180
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6000
      TabIndex        =   58
      Top             =   5820
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtCrotchStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6000
      TabIndex        =   56
      Top             =   5460
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtTOSRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   45
      Top             =   5460
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLargestRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   44
      Top             =   5460
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtThighRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   43
      Top             =   5820
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtWaistRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   42
      Top             =   5820
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtMidPointRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   41
      Top             =   6180
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame2 
      Caption         =   "Styles/Modifications"
      Height          =   2775
      Left            =   6900
      TabIndex        =   38
      Top             =   2250
      Width           =   3435
      Begin VB.Frame Frame4 
         Caption         =   "Right Leg"
         Height          =   1335
         Left            =   1980
         TabIndex        =   15
         Top             =   780
         Width           =   1155
         Begin VB.OptionButton optRightLeg 
            Caption         =   "Brief"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   94
            Top             =   1020
            Width           =   795
         End
         Begin VB.OptionButton optRightLeg 
            Caption         =   "Panty"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   93
            Top             =   690
            Width           =   915
         End
         Begin VB.OptionButton optRightLeg 
            Caption         =   "Full"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Left Leg"
         Height          =   1335
         Left            =   360
         TabIndex        =   14
         Top             =   780
         Width           =   1155
         Begin VB.OptionButton optLeftLeg 
            Caption         =   "Brief"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   1020
            Width           =   795
         End
         Begin VB.OptionButton optLeftLeg 
            Caption         =   "Panty"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   91
            Top             =   690
            Width           =   915
         End
         Begin VB.OptionButton optLeftLeg 
            Caption         =   "Full"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.ComboBox cboCrotchStyle 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2280
         Width           =   1995
      End
      Begin VB.ComboBox cboLegStyle 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label14 
         Caption         =   "Crotch Style:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Leg Style:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Torso Body Measurements"
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   90
      Width           =   6675
      Begin VB.TextBox txtFoldHtRevised 
         Height          =   315
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   53
         Top             =   3600
         Width           =   600
      End
      Begin VB.CheckBox chkDrawRevisedHts 
         Caption         =   "Draw with revised heights"
         Height          =   255
         Left            =   2760
         TabIndex        =   67
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtTOSHtRevised 
         Height          =   315
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   11
         Top             =   660
         Width           =   600
      End
      Begin VB.TextBox txtWaistHtRevised 
         Height          =   315
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1140
         Width           =   600
      End
      Begin VB.TextBox txtFoldHt 
         Height          =   315
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   10
         Top             =   3600
         Width           =   600
      End
      Begin VB.TextBox txtRightThighCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   9
         Top             =   3060
         Width           =   600
      End
      Begin VB.ComboBox cboMidPointRed 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1620
         Width           =   700
      End
      Begin VB.TextBox txtMidPointHt 
         Height          =   315
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1620
         Width           =   600
      End
      Begin VB.TextBox txtMidPointCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1620
         Width           =   600
      End
      Begin VB.ComboBox cboWaistRed 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1140
         Width           =   700
      End
      Begin VB.TextBox txtWaistHt 
         Height          =   315
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1140
         Width           =   600
      End
      Begin VB.TextBox txtWaistCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1140
         Width           =   600
      End
      Begin VB.ComboBox cboThighRed 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2820
         Width           =   700
      End
      Begin VB.TextBox txtLeftThighCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2580
         Width           =   600
      End
      Begin VB.ComboBox cboLargestRed 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2100
         Width           =   700
      End
      Begin VB.TextBox txtLargestHt 
         Height          =   315
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2100
         Width           =   600
      End
      Begin VB.TextBox txtLargestCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2100
         Width           =   600
      End
      Begin VB.ComboBox cboTOSRed 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   660
         Width           =   700
      End
      Begin VB.TextBox txtTOSHt 
         Height          =   315
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   1
         Top             =   660
         Width           =   600
      End
      Begin VB.TextBox txtTOSCir 
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   0
         Top             =   660
         Width           =   600
      End
      Begin VB.Label labFoldHtRevised 
         Height          =   225
         Left            =   4920
         TabIndex        =   78
         Top             =   3660
         Width           =   720
      End
      Begin VB.Label labWaistHtRevised 
         Height          =   225
         Left            =   4860
         TabIndex        =   90
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label labTOSHtRevised 
         Height          =   225
         Left            =   4860
         TabIndex        =   89
         Top             =   690
         Width           =   735
      End
      Begin VB.Label labFoldHt 
         Height          =   225
         Left            =   3480
         TabIndex        =   88
         Top             =   3660
         Width           =   600
      End
      Begin VB.Label labLargestHt 
         Height          =   225
         Left            =   3480
         TabIndex        =   87
         Top             =   2130
         Width           =   600
      End
      Begin VB.Label labMidPointHt 
         Height          =   225
         Left            =   3480
         TabIndex        =   86
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label labWaistHt 
         Height          =   225
         Left            =   3480
         TabIndex        =   85
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label labTOSHt 
         Height          =   225
         Left            =   3480
         TabIndex        =   84
         Top             =   690
         Width           =   600
      End
      Begin VB.Label labRightThighCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   83
         Top             =   3090
         Width           =   600
      End
      Begin VB.Label labLeftThighCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   82
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label labLargestCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   81
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label labMidPointCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   80
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label labWaistCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   79
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label labTOSCir 
         Height          =   225
         Left            =   2040
         TabIndex        =   77
         Top             =   720
         Width           =   600
      End
      Begin VB.Label labRevisedHts 
         Caption         =   "Revised Hts"
         Height          =   195
         Left            =   4140
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labFold 
         Caption         =   "Fold height"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   3660
         Width           =   1000
      End
      Begin VB.Label labRightThigh 
         Caption         =   "Thigh - Right"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label labMidPoint 
         Caption         =   "Midpoint"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   1680
         Width           =   1000
      End
      Begin VB.Label labWaist 
         Caption         =   "Waist"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   1200
         Width           =   1000
      End
      Begin VB.Label labLeftThigh 
         Caption         =   "Thigh - Left"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   2640
         Width           =   1000
      End
      Begin VB.Label labLargest 
         Caption         =   "Buttocks"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Label Label6 
         Caption         =   "REDUC."
         Height          =   195
         Left            =   5640
         TabIndex        =   29
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "HEIGHTS"
         Height          =   195
         Left            =   2760
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "CIRCUMFER."
         Height          =   195
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label labTOS 
         Caption         =   "TOS"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   720
         Width           =   1000
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "DDE Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   66
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Body"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   65
      Top             =   6240
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6720
      TabIndex        =   52
      Top             =   6240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Body UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4740
      TabIndex        =   61
      Top             =   6240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4740
      TabIndex        =   59
      Top             =   5880
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Crotch Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4740
      TabIndex        =   57
      Top             =   5520
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TOS reduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   51
      Top             =   5520
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Largest reduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   49
      Top             =   5520
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Thigh reduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   48
      Top             =   5880
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Waist reduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Midpoint reduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   46
      Top             =   6240
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "whboddia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   '* Windows API Functions Declarations
'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
'
'   'Constants used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4''
'
'   'MsgBox constant
'    Const IDYES = 6
'    Const IDNO = 7
'    Const IDCANCEL = 2

    Option Explicit

    

Private Sub Cancel_Click()
    Dim Response As Integer, sTask$, sCurrentValues$
    
   'Check if data has been modified
    sCurrentValues = FN_ValuesString()

    If sCurrentValues <> g_sChangeChecker Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "WAIST Height Body - Dialogue")
        Select Case Response
            Case IDYES
                PR_CreateMacro_Save "c:\jobst\draw.d"
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate sTask
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Dialogue"
                End If
            Case IDNO
                End
            Case IDCANCEL
                Exit Sub
        End Select
    Else
        End
    End If
End Sub

Private Sub cboCrotchStyle_Click()
    Dim sString As String, sError As String
    Dim iIndex As Integer
    
    sError = ""

   'Check Crotch style against sex
    sString = cboCrotchStyle.List(cboCrotchStyle.ListIndex)

   'Check for male crotch and female patient
    iIndex = 0
    iIndex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
    If iIndex <> 0 And txtSex = "Female" Then
        sError = sError + "Male Crotch type specified for Female patient" + NL
    End If

   'Check for female crotch and male patient
    iIndex = 0
    iIndex = InStr(1, sString, "Female", 0)
    If iIndex <> 0 And txtSex = "Male" Then
        sError = sError + "Female Crotch type specified for Male patient" + NL
    End If


   'Display Error message (if required) and return
   'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 0, "Crotch Style Error"
    End If

End Sub

Private Sub cboLegStyle_Click()
   '
    Dim sLegStyle As String, sReductionFactor  As String
    Dim iAge As Integer

    If Left$(g_sPreviousLegStyle, 4) = "Chap" Then prEnable_WRT_Chap
    If Left$(g_sPreviousLegStyle, 4) = "W4OC" Then prEnable_WRT_W4OC
    If Left$(g_sPreviousLegStyle, 5) = "Panty" Then prRestore_Reductions
    If Left$(g_sPreviousLegStyle, 5) = "Brief" Then prRestore_Reductions
  
   'Get the currently selected leg stlye
    Select Case cboLegStyle.List(cboLegStyle.ListIndex)

    Case "Full-Leg"
        prOptLeftLeg 1, 0, 0, 1, 1, 0
        prOptRightLeg 1, 0, 0, 1, 1, 0
       'Set dropdown combo boxes for reductions
       'Only set if Leg Style had been changed
        If g_sPreviousLegStyle <> "" Then
            prStore_Reductions
           'Reductions, set defaults
            sReductionFactor = "0.83"
            iAge = Val(txtAge.Text)
            If iAge <= 3 Then sReductionFactor = "0.88"
            If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
           
           'NB. Special case for TOS as TOS need not be given.
            If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
                prSet_Red_Combo sReductionFactor, cboTOSRed
            End If
            prSet_Red_Combo sReductionFactor, cboWaistRed
            prSet_Red_Combo sReductionFactor, cboMidPointRed
            prSet_Red_Combo sReductionFactor, cboLargestRed
            prSet_Red_Combo sReductionFactor, cboThighRed
        End If


    Case "Panty"
        prOptLeftLeg 0, 1, 0, 0, 1, 1
        prOptRightLeg 0, 1, 0, 0, 1, 1
       'Set dropdown combo boxes for reductions
       'Only set if Leg Style had been changed
        If g_sPreviousLegStyle <> "" Then
            prStore_Reductions
            sReductionFactor = "0.85"
           'NB. Special case for TOS as TOS need not be given.
            If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
                prSet_Red_Combo sReductionFactor, cboTOSRed
            End If
            prSet_Red_Combo sReductionFactor, cboWaistRed
            prSet_Red_Combo sReductionFactor, cboMidPointRed
            prSet_Red_Combo sReductionFactor, cboLargestRed
            prSet_Red_Combo sReductionFactor, cboThighRed
        End If

    Case "Chap-Left"
        prDisable_WRT_Chap
        prOptLeftLeg 1, 0, 0, 1, 1, 0
        prOptRightLeg 0, 0, 0, 0, 0, 0
        If g_sPreviousLegStyle <> "" Then prSet_Red_Combo "0.87", cboWaistRed
    
    Case "Chap-Right"
        prDisable_WRT_Chap
        prOptLeftLeg 0, 0, 0, 0, 0, 0
        prOptRightLeg 1, 0, 0, 1, 1, 0
        If g_sPreviousLegStyle <> "" Then prSet_Red_Combo "0.87", cboWaistRed
    
    Case "Chap-Both Legs"
        prDisable_WRT_Chap
        prOptLeftLeg 1, 0, 0, 1, 1, 0
        prOptRightLeg 1, 0, 0, 1, 1, 0
        If g_sPreviousLegStyle <> "" Then prSet_Red_Combo "0.87", cboWaistRed

    Case "Brief"
        prOptLeftLeg 0, 0, 1, 0, 0, 1
        prOptRightLeg 0, 0, 1, 0, 0, 1
       'Set dropdown combo boxes for reductions
       'Only set if Leg Style had been changed
        If g_sPreviousLegStyle <> "" Then
            prStore_Reductions
            sReductionFactor = "0.85"
           'NB. Special case for TOS as TOS need not be given.
            If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
                prSet_Red_Combo sReductionFactor, cboTOSRed
            End If
            prSet_Red_Combo sReductionFactor, cboWaistRed
            prSet_Red_Combo sReductionFactor, cboMidPointRed
            prSet_Red_Combo sReductionFactor, cboLargestRed
            prSet_Red_Combo sReductionFactor, cboThighRed
        End If


    Case "W4OC-Left"
        prDisable_WRT_W4OC
       'Enable Left thigh, restore if value available
        prOptLeftLeg 1, 0, 0, 1, 1, 0
        txtLeftThighCir.Enabled = True
        labLeftThighCir.Enabled = True
        labLeftThigh.Enabled = True
        If Val(txtLeftThighCir.Text) = 0 And g_nLeftThighCir Then txtLeftThighCir = g_nLeftThighCir
        txtLeftThighCir_LostFocus
       
       'Disable Right thigh data, save if required for a restore
        If Val(txtRightThighCir.Text) > 0 Then g_nRightThighCir = Val(txtRightThighCir.Text)
        prOptRightLeg 0, 0, 0, 0, 0, 0
'        txtRightThighCir = ""
'        txtRightThighCir_LostFocus
        labRightThigh.Enabled = False
        txtRightThighCir.Enabled = False
        labRightThighCir.Enabled = False

    Case "W4OC-Right"
        prDisable_WRT_W4OC
       'Enable Right thigh, restore if value available
        prOptRightLeg 1, 0, 0, 1, 1, 0
        txtRightThighCir.Enabled = True
        labRightThighCir.Enabled = True
        labRightThigh.Enabled = True
        If Val(txtRightThighCir.Text) = 0 And g_nRightThighCir Then txtRightThighCir = g_nRightThighCir
        txtRightThighCir_LostFocus

       'Disable Left thigh data, save if required for a restore
        If Val(txtLeftThighCir.Text) > 0 Then g_nLeftThighCir = Val(txtLeftThighCir.Text)
        prOptLeftLeg 0, 0, 0, 0, 0, 0
'        txtLeftThighCir = ""
'        txtLeftThighCir_LostFocus
        labLeftThigh.Enabled = False
        txtLeftThighCir.Enabled = False
        labLeftThighCir.Enabled = False

    End Select

   'Store leg style for use when it changes
    g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

End Sub

Private Sub chkDrawRevisedHts_Click()
    If chkDrawRevisedHts.Value = 1 Then
        If Val(txtTOSHtRevised) > 0 Then
            txtTOSHt.Enabled = False
            labTOSHt.Enabled = False
        End If
        If Val(txtWaistHtRevised) > 0 Then
            txtWaistHt.Enabled = False
            labWaistHt.Enabled = False
        End If
        If Val(txtFoldHtRevised) > 0 Then
            txtFoldHt.Enabled = False
            labFoldHt.Enabled = False
        End If
        txtWaistHtRevised.Enabled = True
        txtTOSHtRevised.Enabled = True
        txtFoldHtRevised.Enabled = True
        labWaistHtRevised.Enabled = True
        labTOSHtRevised.Enabled = True
        labFoldHtRevised.Enabled = True
        labRevisedHts.Enabled = True
    Else
        txtTOSHtRevised.Enabled = False
        txtWaistHtRevised.Enabled = False
        txtFoldHtRevised.Enabled = False
        labTOSHtRevised.Enabled = False
        labWaistHtRevised.Enabled = False
        labFoldHtRevised.Enabled = False
        txtTOSHt.Enabled = True
        txtWaistHt.Enabled = True
        txtFoldHt.Enabled = True
        labTOSHt.Enabled = True
        labWaistHt.Enabled = True
        labFoldHt.Enabled = True
        labRevisedHts.Enabled = False
    End If
End Sub

Private Sub cmdTab_Click()
   'Allows the user to use enter as a tab
    SendKeys "{TAB}", False

End Sub

Private Function FN_Open(sDrafixFile As String, sType As String, sName As Variant, sFileNo As Variant) As Integer
   'Open the DRAFIX macro file
   'Return the file number

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_Open = fNum

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; sName; ", "; sFileNo; ""
    Print #fNum, "//by Visual Basic, Waist Height Body"
    Print #fNum, "//type - "; sType; ""

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"

   'Display Hour Glass symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; sType; QQ; ");"


End Function


Private Function FN_ValuesString() As String
   'Create a string of all the values given in the
   'Text and Combo boxes.
   '
   'Ignore patient details
   '
    Dim ii%, sString$
    
    FN_ValuesString = ""
    sString = ""

   'Dimensions
    sString = sString & txtTOSCir & txtTOSHt & txtWaistCir & txtWaistHt & txtMidPointCir
    sString = sString & txtMidPointHt & txtLargestCir & txtLargestHt & txtLeftThighCir
    sString = sString & txtRightThighCir & txtFoldHt
    
    sString = sString & cboLegStyle.Text & cboCrotchStyle.Text
    
    sString = sString & chkDrawRevisedHts.Value

    sString = sString & optLeftLeg(0).Value & optLeftLeg(1).Value & optLeftLeg(2).Value

    sString = sString & optRightLeg(0).Value & optRightLeg(1).Value & optRightLeg(2).Value

    FN_ValuesString = sString

End Function

Private Sub Form_LinkClose()

  'On link close.
   Dim sReductionFactor As String
   Dim nAge As Integer, ii As Integer
   Dim nValue As Double

  'Disable Time out
   Timer1.Enabled = False


  
  'Set global Units Factor
   If txtUnits = "cm" Then
        g_nUnitsFac = 10 / 25.4
   Else
        g_nUnitsFac = 1
   End If
        
  'Waist Height details
  'Note
  ' The procedure Validate_Text_In_Box also displays the size in
  ' inches.
  '
   Validate_Text_In_Box txtTOSCir, labTOSCir
   Validate_Text_In_Box txtTOSHt, labTOSHt
   Validate_Text_In_Box txtWaistCir, labWaistCir
   Validate_Text_In_Box txtWaistHt, labWaistHt
   Validate_Text_In_Box txtMidPointCir, labMidPointCir
   Validate_Text_In_Box txtMidPointHt, labMidPointHt
   Validate_Text_In_Box txtLargestCir, labLargestCir
   Validate_Text_In_Box txtLargestHt, labLargestHt
   Validate_Text_In_Box txtLeftThighCir, labLeftThighCir
   Validate_Text_In_Box txtRightThighCir, labRightThighCir
   Validate_Text_In_Box txtFoldHt, labFoldHt

  'Set dropdown combo boxes
  'Crotch Style
   cboCrotchStyle.AddItem "Open Crotch"         'Both
   cboCrotchStyle.AddItem "Horizontal Fly"      'Male only
   cboCrotchStyle.AddItem "Diagonal Fly"        'Male only
   cboCrotchStyle.AddItem "Gusset"              'Both
   cboCrotchStyle.AddItem "Mesh Gusset"         'Both
   cboCrotchStyle.AddItem "Female Mesh Gusset"  'Female only
   cboCrotchStyle.AddItem "Male Mesh Gusset"    'Male only


  'Set value
   For ii = 0 To (cboCrotchStyle.ListCount - 1)
        If cboCrotchStyle.List(ii) = txtCrotchStyle.Text Then
            cboCrotchStyle.ListIndex = ii
        End If
   Next ii
   
  'Leg Style
   cboLegStyle.AddItem "Full-Leg"
   cboLegStyle.AddItem "Panty"
   cboLegStyle.AddItem "Brief"
   cboLegStyle.AddItem "W4OC-Left"
   cboLegStyle.AddItem "W4OC-Right"
   cboLegStyle.AddItem "Chap-Left"
   cboLegStyle.AddItem "Chap-Right"
   cboLegStyle.AddItem "Chap-Both Legs"


  'Set values of combo and option buttons
   g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
   cboLegStyle.ListIndex = fnGetNumber(txtLegStyle.Text, 1)

  'Set this now to avoid problems with cboLegStyle_Click
   If cboLegStyle.ListIndex = -1 Then
        g_sPreviousLegStyle = "None"
   Else
        g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)
   End If

'   ii = fnGetNumber(txtLegStyle.Text, 2)
'   If ii >= 0 Then optLeftLeg(ii).Value = True
'   ii = fnGetNumber(txtLegStyle.Text, 3)
'   If ii >= 0 Then optRightLeg(ii).Value = True
  
  'Reductions (set defaults if no handle to the body)
   sReductionFactor = ""
   nAge = Val(txtAge.Text)
   If Val(txtUidWHBody.Text) = 0 Then
        sReductionFactor = "0.83"
        If nAge <= 3 Then
            sReductionFactor = "0.88"
        End If
        If nAge <= 10 And nAge > 3 Then
            sReductionFactor = "0.86"
        End If
   End If
   
  'Set dropdown combo boxes for reductions
  'NB. Special case for TOS as TOS need not be given.
  '    Also if this is the first time through then display reduction.
   prInitialise_Red_Combo sReductionFactor, txtTOSRed, cboTOSRed
   If (Len(txtTOSCir.Text) = 0 And Len(txtTOSHt.Text) = 0) Or Len(txtUidWHBody.Text) = 0 Then
      cboTOSRed.ListIndex = -1
      cboTOSRed.Enabled = False
   End If
   prInitialise_Red_Combo sReductionFactor, txtWaistRed, cboWaistRed
   prInitialise_Red_Combo sReductionFactor, txtMidPointRed, cboMidPointRed
   prInitialise_Red_Combo sReductionFactor, txtLargestRed, cboLargestRed
   prInitialise_Red_Combo sReductionFactor, txtThighRed, cboThighRed
   
   ii = fnGetNumber(txtLegStyle.Text, 2)
   If ii >= 0 Then optLeftLeg(ii).Value = True
   ii = fnGetNumber(txtLegStyle.Text, 3)
   If ii >= 0 Then optRightLeg(ii).Value = True


  'Revised Heights
  'From Body DB field
   nValue = fnGetNumber(txtBody.Text, 5)
   If nValue > 0 Then
        txtTOSHtRevised = nValue
        Validate_Text_In_Box txtTOSHtRevised, labTOSHtRevised
   End If

   nValue = fnGetNumber(txtBody.Text, 6)
   If nValue > 0 Then
        txtWaistHtRevised = nValue
        Validate_Text_In_Box txtWaistHtRevised, labWaistHtRevised
   End If

   nValue = fnGetNumber(txtBody.Text, 7)
   If nValue > 0 Then
        txtFoldHtRevised = nValue
        Validate_Text_In_Box txtFoldHtRevised, labFoldHtRevised
   End If

   If fnGetNumber(txtBody.Text, 4) > 0 Then
        chkDrawRevisedHts.Value = 0
        chkDrawRevisedHts.Value = 1
   End If

  'Store loded values w.r.t cancel button
  '
   g_sChangeChecker = FN_ValuesString()
    
   Show
   Screen.MousePointer = 0    ' Change pointer to default.
   whboddia.SetFocus

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Hide
   'Check if this is already running
   'If it is then warn the user and exit
    If App.PrevInstance Then
        End
    End If
   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to Hourglass.
   'Reset in Form_LinkClose

   'Position to center of screen
    Left = (Screen.Width - whboddia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - whboddia.Height) / 2  ' Center form vertically.

    Set MainForm = whboddia
    
   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    g_sPathJOBST = fnPathJOBST()

   'Set Units global to default to inches
    g_nUnitsFac = 1

   'Disable revised Hts (default), set from Link Close event
    chkDrawRevisedHts.Value = 1
    chkDrawRevisedHts.Value = 0

   'Start a timer that will ensure that the
   'the Vb programme ends if the DRAFIX DDE
   'link fails.
    Timer1.Interval = 6000
    Timer1.Enabled = True
End Sub

Private Sub OK_Click()
   'Check that data is all present and insert into drafix

    Dim sTask$

    If Validate_Data() Then
        Screen.MousePointer = 11
        Hide

        PR_CreateMacro_Save "c:\jobst\draw.d"
        
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            AppActivate sTask
            SendKeys "@C:\JOBST\DRAW.D{enter}", True
            Screen.MousePointer = 0
            End
        Else
            MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height Body - Data"
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub


Private Sub optLeftLeg_Click(Index As Integer)
   '0 = Full-Leg
   '1 = Panty
   '2 = Brief

    Dim sLegStyle As String, sReductionFactor   As String
    Dim iAge As Integer

    sLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

    If sLegStyle = "Full-Leg" Then
        If Index = 0 Then
            optRightLeg(1).Enabled = True
            optRightLeg(2).Enabled = True
        Else
            optRightLeg(1).Enabled = False
            optRightLeg(2).Enabled = False
        End If
    End If

    If sLegStyle = "Panty" Then
        If Index = 1 Then
            optRightLeg(2).Enabled = True
        Else
            optRightLeg(2).Enabled = False
        End If
    End If

    If sLegStyle = "W4OC-Left" And g_sPreviousLegStyle <> "" Then
       'Reductions, set defaults
        prStore_Reductions
        If Index = 0 Then
           'Reductions, for full leg
            sReductionFactor = "0.83"
            iAge = Val(txtAge.Text)
            If iAge <= 3 Then sReductionFactor = "0.88"
            If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
        Else
           'Reductions panty leg
            sReductionFactor = "0.85"
        End If
 
       'NB. Special case for TOS as TOS need not be given.
        If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
            prSet_Red_Combo sReductionFactor, cboTOSRed
        End If
        prSet_Red_Combo sReductionFactor, cboWaistRed
        prSet_Red_Combo sReductionFactor, cboMidPointRed
        prSet_Red_Combo sReductionFactor, cboLargestRed
        prSet_Red_Combo sReductionFactor, cboThighRed
    End If

End Sub

Private Sub optRightLeg_Click(Index As Integer)
   '0 = Full-Leg
   '1 = Panty
   '2 = Brief

    Dim sLegStyle As String, sReductionFactor  As String
    Dim iAge As Integer

    sLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

    If sLegStyle = "Full-Leg" Then
        If Index = 0 Then
            optLeftLeg(1).Enabled = True
            optLeftLeg(2).Enabled = True
        Else
            optLeftLeg(1).Enabled = False
            optLeftLeg(2).Enabled = False
        End If
    End If

    If sLegStyle = "Panty" Then
        If Index = 1 Then
            optLeftLeg(2).Enabled = True
        Else
            optLeftLeg(2).Enabled = False
        End If
    End If

    If sLegStyle = "W4OC-Right" And g_sPreviousLegStyle <> "" Then
       'Reductions, set defaults
        prStore_Reductions
        If Index = 0 Then
           'Reductions, for full leg
            sReductionFactor = "0.83"
            iAge = Val(txtAge.Text)
            If iAge <= 3 Then sReductionFactor = "0.88"
            If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
        Else
           'Reductions panty leg
            sReductionFactor = "0.85"
        End If

       'NB. Special case for TOS as TOS need not be given.
        If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
            prSet_Red_Combo sReductionFactor, cboTOSRed
        End If
        prSet_Red_Combo sReductionFactor, cboWaistRed
        prSet_Red_Combo sReductionFactor, cboMidPointRed
        prSet_Red_Combo sReductionFactor, cboLargestRed
        prSet_Red_Combo sReductionFactor, cboThighRed
    End If

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
    PR_PutLine "HANDLE hMPD, hBody;"
   
    Update_DDE_Text_Boxes

    Dim sSymbol As String
    
    sSymbol = "waistbody"

    If txtUidWHBody.Text = "" Then
       'Define DB Fields
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHFIELDS.D;"

       'Find "mainpatientdetails" and get position
        PR_PutLine "XY     xyMPD_Origin, xyMPD_Scale ;"
        PR_PutLine "STRING sMPD_Name;"
        PR_PutLine "ANGLE  aMPD_Angle;"

        PR_PutLine "hMPD = UID (" & QQ & "find" & QC & Val(txtUidTitle.Text) & ");"
        PR_PutLine "if (hMPD)"
        PR_PutLine "  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");"

       'Insert body symbol
        PR_PutLine "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){"
        PR_PutLine "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));"
        PR_PutLine "  hBody = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);"
        PR_PutLine "  }"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");"
    Else
       'Use existing symbol
        PR_PutLine "hBody = UID (" & QQ & "find" & QC & Val(txtUidWHBody.Text) & ");"
        PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

    End If

   'Update the BODY Box symbol
    PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "TOSCir" & QCQ & txtTOSCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "TOSGivenRed" & QCQ & txtTOSRed.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "TOSHt" & QCQ & txtTOSHt.Text & QQ & ");"

    PR_PutLine "SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtWaistCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "WaistGivenRed" & QCQ & txtWaistRed.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "WaistHt" & QCQ & txtWaistHt.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "MidPointCir" & QCQ & txtMidPointCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "MidPointGivenRed" & QCQ & txtMidPointRed.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "MidPointHt" & QCQ & txtMidPointHt.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "LargestCir" & QCQ & txtLargestCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "LargestGivenRed" & QCQ & txtLargestRed.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "LargestHt" & QCQ & txtLargestHt.Text & QQ & ");"

    PR_PutLine "SetDBData( hBody" & CQ & "LeftThighCir" & QCQ & txtLeftThighCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "RightThighCir" & QCQ & txtRightThighCir.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ThighGivenRed" & QCQ & txtThighRed.Text & QQ & ");"

    PR_PutLine "SetDBData( hBody" & CQ & "FoldHt" & QCQ & txtFoldHt.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "CrotchStyle" & QCQ & txtCrotchStyle.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ID" & QCQ & txtLegStyle.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "Body" & QCQ & txtBody.Text & QQ & ");"

    Close #fNum
   
End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable

    Print #fNum, sLine

End Sub

Private Sub prDisable_WRT_Chap()
   'Disables the text boxes with respect to the Chap Style
   'Called from cboLegStyle_Click

   'Store existing values (nb 0 values are not stored)
    prStoreValues

'    txtTOSCir = ""
'    txtTOSCir_LostFocus
'    txtTOSHt = ""
'    txtTOSHt_LostFocus
'    txtTOSHtRevised = ""
'    txtTOSHtRevised_LostFocus
    
'    txtTOSCir.Enabled = False
'    labTOSCir.Enabled = False
'    txtTOSHt.Enabled = False
'    labTOSHt.Enabled = False
'    txtTOSHtRevised.Enabled = False
'    labTOSHtRevised.Enabled = False
'    labTOS.Enabled = False
'    cboTOSRed.ListIndex = -1
'    cboTOSRed.Enabled = False
    
'    txtMidPointCir = ""
'    txtMidPointCir_LostFocus
'    txtMidPointHt = ""
'    txtMidPointHt_LostFocus
    txtMidPointCir.Enabled = False
    labMidPointCir.Enabled = False
    txtMidPointHt.Enabled = False
    labMidPointHt.Enabled = False
    labMidPoint.Enabled = False
    cboMidPointRed.ListIndex = -1
    cboMidPointRed.Enabled = False

'    txtLargestCir = ""
'    txtLargestCir_LostFocus
'    txtLargestHt = ""
'    txtLargestHt_LostFocus
    txtLargestCir.Enabled = False
    labLargestCir.Enabled = False
    txtLargestHt.Enabled = False
    labLargestHt.Enabled = False
    labLargest.Enabled = False
    cboLargestRed.ListIndex = -1
    cboLargestRed.Enabled = False

'    txtLeftThighCir = ""
'    txtLeftThighCir_LostFocus
    txtLeftThighCir.Enabled = False
    labLeftThighCir.Enabled = False
    labLeftThigh.Enabled = False

'    txtRightThighCir = ""
'    txtRightThighCir_LostFocus
    txtRightThighCir.Enabled = False
    labRightThighCir.Enabled = False
    labRightThigh.Enabled = False
    cboThighRed.ListIndex = -1
    cboThighRed.Enabled = False

    cboCrotchStyle.ListIndex = -1    'No crotch for a chap style
    cboCrotchStyle.Enabled = False

End Sub

Private Sub prDisable_WRT_W4OC()
    
    cboCrotchStyle.ListIndex = 0    'Open Crotch
    cboCrotchStyle.Enabled = False

End Sub

Private Sub prEnable_WRT_Chap()
   'Enable all the Text boxes that where disabled when the
   'Chap style was chosen before
   '
   'Restore the original values
    Dim ii As Integer

    If g_nTOSCir > 0 Then txtTOSCir = g_nTOSCir
    txtTOSCir_LostFocus
    If g_nTOSHt > 0 Then txtTOSHt = g_nTOSHt
    txtTOSHt_LostFocus
    If g_nTOSHtRevised > 0 Then txtTOSHtRevised = g_nTOSHtRevised
    txtTOSHtRevised_LostFocus
    txtTOSCir.Enabled = True
    labTOSCir.Enabled = True
    txtTOSHt.Enabled = True
    labTOSHt.Enabled = True
    txtTOSHtRevised.Enabled = True
    labTOSHtRevised.Enabled = True
    labTOS.Enabled = True
    
    If g_nMidPointCir > 0 Then txtMidPointCir = g_nMidPointCir
    txtMidPointCir_LostFocus
    If g_nMidPointHt > 0 Then txtMidPointHt = g_nMidPointHt
    txtMidPointHt_LostFocus
    txtMidPointCir.Enabled = True
    labMidPointCir.Enabled = True
    txtMidPointHt.Enabled = True
    labMidPointHt.Enabled = True
    labMidPoint.Enabled = True

    If g_nLargestCir > 0 Then txtLargestCir = g_nLargestCir
    txtLargestCir_LostFocus
    If g_nLargestHt > 0 Then txtLargestHt = g_nLargestHt
    txtLargestHt_LostFocus
    txtLargestCir.Enabled = True
    txtLargestHt.Enabled = True
    labLargestCir.Enabled = True
    labLargestHt.Enabled = True
    labLargest.Enabled = True
    
    If g_nLeftThighCir > 0 Then txtLeftThighCir = g_nLeftThighCir
    txtLeftThighCir_LostFocus
    txtLeftThighCir.Enabled = True
    labLeftThighCir.Enabled = True
    labLeftThigh.Enabled = True

    If g_nRightThighCir > 0 Then txtRightThighCir = g_nRightThighCir
    txtRightThighCir_LostFocus
    txtRightThighCir.Enabled = True
    labRightThighCir.Enabled = True
    labRightThigh.Enabled = True

    cboTOSRed.Enabled = True
    cboMidPointRed.Enabled = True
    cboLargestRed.Enabled = True
    cboThighRed.Enabled = True

    prRestore_Reductions

  'Set value
   For ii = 0 To (cboCrotchStyle.ListCount - 1)
        If cboCrotchStyle.List(ii) = txtCrotchStyle.Text Then
            cboCrotchStyle.ListIndex = ii
        End If
   Next ii
   cboCrotchStyle.Enabled = True


End Sub

Private Sub prEnable_WRT_W4OC()
    Dim ii As Integer

   'Enable Left thigh, restore if value available
    prOptLeftLeg 0, 1, 0, 1, 1, 0
    txtLeftThighCir.Enabled = True
    labLeftThigh.Enabled = True
    If Val(txtLeftThighCir.Text) = 0 And g_nLeftThighCir Then txtLeftThighCir = g_nLeftThighCir
    txtLeftThighCir_LostFocus

   'Enable Right thigh, restore if value available
    prOptRightLeg 0, 1, 0, 1, 1, 0
    txtRightThighCir.Enabled = True
    labRightThigh.Enabled = True
    If Val(txtRightThighCir.Text) = 0 And g_nRightThighCir Then txtRightThighCir = g_nRightThighCir
    txtRightThighCir_LostFocus
        
  'Set value for crotch style
    For ii = 0 To (cboCrotchStyle.ListCount - 1)
        If cboCrotchStyle.List(ii) = txtCrotchStyle.Text Then
            cboCrotchStyle.ListIndex = ii
        End If
    Next ii
    cboCrotchStyle.Enabled = True

End Sub

Private Sub prInitialise_Red_Combo(sRed As String, Text_Box_Name As Control, Combo_Box_Name As Control)
    Dim nIndex As Integer

   'Set of combo box
    Combo_Box_Name.AddItem "0.78"   '0
    Combo_Box_Name.AddItem "0.79"   '1
    Combo_Box_Name.AddItem "0.80"   '2
    Combo_Box_Name.AddItem "0.81"   '3
    Combo_Box_Name.AddItem "0.82"   '4
    Combo_Box_Name.AddItem "0.83"   '5
    Combo_Box_Name.AddItem "0.84"   '6
    Combo_Box_Name.AddItem "0.85"   '7
    Combo_Box_Name.AddItem "0.86"   '8
    Combo_Box_Name.AddItem "0.87"   '9
    Combo_Box_Name.AddItem "0.88"   '10
    Combo_Box_Name.AddItem "0.89"   '11
    Combo_Box_Name.AddItem "0.90"   '12
    Combo_Box_Name.AddItem "0.91"   '13
    Combo_Box_Name.AddItem "0.92"   '15
    Combo_Box_Name.AddItem "0.93"   '15

   'If no values given exit
    If Val(sRed) = 0 And Val(Text_Box_Name.Text) = 0 Then
        Combo_Box_Name.ListIndex = -1
        Exit Sub
    End If

   'Set the value of the combo box as given
    If sRed = "" Then
        Combo_Box_Name.ListIndex = (Val(Text_Box_Name.Text) * 100) - 78
    Else
        Combo_Box_Name.ListIndex = (Val(sRed) * 100) - 78
    End If

End Sub

Private Sub prOptLeftLeg(i0 As Integer, i1 As Integer, i2 As Integer, j0 As Integer, j1 As Integer, j2 As Integer)
   'Procedure to set the option buttons
   'sets them according to the pattern i0-i2
   'enables or disables according to the pattern j0-j2

    optLeftLeg(0).Value = i0
    optLeftLeg(1).Value = i1
    optLeftLeg(2).Value = i2
    optLeftLeg(0).Enabled = j0
    optLeftLeg(1).Enabled = j1
    optLeftLeg(2).Enabled = j2

End Sub

Private Sub prOptRightLeg(i0 As Integer, i1 As Integer, i2 As Integer, j0 As Integer, j1 As Integer, j2 As Integer)
   'Procedure to set the option buttons
   'sets them according to the pattern i0-i2
   'enables or disables according to the pattern j0-j2

    optRightLeg(0).Value = i0
    optRightLeg(1).Value = i1
    optRightLeg(2).Value = i2
    optRightLeg(0).Enabled = j0
    optRightLeg(1).Enabled = j1
    optRightLeg(2).Enabled = j2

End Sub

Private Sub prRestore_Reductions()

   'Restore reduction values from stored ones
    cboThighRed.ListIndex = g_iThighRed
    cboLargestRed.ListIndex = g_iLargestRed
    cboMidPointRed.ListIndex = g_iMidPointRed
    cboWaistRed.ListIndex = g_iWaistRed
    cboTOSRed.ListIndex = g_iTOSRed

End Sub

Private Sub prSet_Red_Combo(sRed As String, Combo_Box_Name As Control)
   'Set the value of the combo box as given
   'Find index into combo box list
    
    If Val(sRed) = 0 Then Exit Sub
    Combo_Box_Name.ListIndex = (Val(sRed) * 100) - 78

End Sub

Private Sub prStore_Reductions()
   'Store Reduction settings for later restore
    g_iThighRed = cboThighRed.ListIndex
    g_iLargestRed = cboLargestRed.ListIndex
    g_iMidPointRed = cboMidPointRed.ListIndex
    g_iWaistRed = cboWaistRed.ListIndex
    g_iTOSRed = cboTOSRed.ListIndex

End Sub

Private Sub prStoreValues()
   'Stores the current values of the text boxes.
   '
   'Called from cboLegStyle_Click
   '
   'Used when the leg style changes so that the text box
   'values can be restored if the operator changes the leg style
   'again
    
    If Val(txtTOSCir.Text) > 0 Then g_nTOSCir = Val(txtTOSCir.Text)
    If Val(txtTOSHt.Text) > 0 Then g_nTOSHt = Val(txtTOSHt.Text)
    If Val(txtTOSHtRevised) > 0 Then g_nTOSHtRevised = Val(txtTOSHtRevised.Text)
    
    If Val(txtWaistCir.Text) > 0 Then g_nWaistCir = Val(txtWaistCir.Text)
    If Val(txtWaistHt.Text) > 0 Then g_nWaistHt = Val(txtWaistHt.Text)
    If Val(txtWaistHtRevised.Text) > 0 Then g_nWaistHtRevised = Val(txtWaistHtRevised.Text)
    
    If Val(txtMidPointCir.Text) > 0 Then g_nMidPointCir = Val(txtMidPointCir.Text)
    If Val(txtMidPointHt.Text) > 0 Then g_nMidPointHt = Val(txtMidPointHt.Text)
    
    If Val(txtLargestCir.Text) > 0 Then g_nLargestCir = Val(txtLargestCir.Text)
    If Val(txtLargestHt.Text) > 0 Then g_nLargestHt = Val(txtLargestHt.Text)

    If Val(txtLeftThighCir.Text) > 0 Then g_nLeftThighCir = Val(txtLeftThighCir.Text)
    If Val(txtRightThighCir.Text) > 0 Then g_nRightThighCir = Val(txtRightThighCir.Text)

    If Val(txtFoldHt.Text) > 0 Then g_nFoldHt = Val(txtFoldHt.Text)
    If Val(txtFoldHtRevised.Text) > 0 Then g_nFoldHtRevised = Val(txtFoldHtRevised.Text)

    prStore_Reductions

End Sub

Private Sub Timer1_Timer()
    'If this event is called then the
    'DDE link with DRAFIX has not closed
    'Therefor we time out after 6 seconds
     End
End Sub

Private Sub txtFoldHt_GotFocus()
    Select_Text_In_Box txtFoldHt
End Sub

Private Sub txtFoldHt_LostFocus()
    Validate_Text_In_Box txtFoldHt, labFoldHt
End Sub

Private Sub txtFoldHtRevised_GotFocus()
    Select_Text_In_Box txtFoldHtRevised
End Sub

Private Sub txtFoldHtRevised_LostFocus()
    Validate_Text_In_Box txtFoldHtRevised, labFoldHtRevised
    If Val(txtFoldHtRevised) > 0 Then
        txtFoldHt.Enabled = False
        labFoldHt.Enabled = False
    Else
        txtFoldHt.Enabled = True
        labFoldHt.Enabled = True
    End If
End Sub

Private Sub txtLargestCir_GotFocus()
    Select_Text_In_Box txtLargestCir
End Sub

Private Sub txtLargestCir_LostFocus()
    Validate_Text_In_Box txtLargestCir, labLargestCir
End Sub

Private Sub txtLargestHt_GotFocus()
    Select_Text_In_Box txtLargestHt
End Sub

Private Sub txtLargestHt_LostFocus()
    Validate_Text_In_Box txtLargestHt, labLargestHt
End Sub

Private Sub txtLeftThighCir_GotFocus()
    Select_Text_In_Box txtLeftThighCir
End Sub

Private Sub txtLeftThighCir_LostFocus()
    Validate_Text_In_Box txtLeftThighCir, labLeftThighCir
End Sub

Private Sub txtMidPointCir_GotFocus()
    Select_Text_In_Box txtMidPointCir
End Sub

Private Sub txtMidPointCir_LostFocus()
    Validate_Text_In_Box txtMidPointCir, labMidPointCir
End Sub

Private Sub txtMidPointHt_GotFocus()
    Select_Text_In_Box txtMidPointHt
End Sub

Private Sub txtMidPointHt_LostFocus()
    Validate_Text_In_Box txtMidPointHt, labMidPointHt
End Sub

Private Sub txtRightThighCir_GotFocus()
    Select_Text_In_Box txtRightThighCir
End Sub

Private Sub txtRightThighCir_LostFocus()
    Validate_Text_In_Box txtRightThighCir, labRightThighCir
End Sub

Private Sub txtTOSCir_GotFocus()
    Select_Text_In_Box txtTOSCir
End Sub

Private Sub txtTOSCir_LostFocus()
    Validate_Text_In_Box txtTOSCir, labTOSCir
    
   'Special case where we are adding a TOS to existing
    Dim iAge As Integer
    Dim sReductionFactor As String
    
    If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Then
        cboTOSRed.Enabled = True
       'Reductions, set defaults
        sReductionFactor = "0.83"
        iAge = Val(txtAge.Text)
        If iAge <= 3 Then sReductionFactor = "0.88"
        If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
        If cboLegStyle.List(cboLegStyle.ListIndex) = "Panty" Then sReductionFactor = "0.85"
        If cboTOSRed.ListIndex = -1 Then
            prSet_Red_Combo sReductionFactor, cboTOSRed
        End If
    End If

End Sub

Private Sub txtTOSHt_GotFocus()
    Select_Text_In_Box txtTOSHt
End Sub

Private Sub txtTOSHt_LostFocus()
    Validate_Text_In_Box txtTOSHt, labTOSHt

   'Special case where we are adding a TOS to existing
    Dim iAge As Integer
    Dim sReductionFactor As String
    
    If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Then
        cboTOSRed.Enabled = True
       'Reductions, set defaults
        sReductionFactor = "0.83"
        iAge = Val(txtAge.Text)
        If iAge <= 3 Then sReductionFactor = "0.88"
        If iAge <= 10 And iAge > 3 Then sReductionFactor = "0.86"
        If cboLegStyle.List(cboLegStyle.ListIndex) = "Panty" Then sReductionFactor = "0.85"
        If cboTOSRed.ListIndex = -1 Then
            prSet_Red_Combo sReductionFactor, cboTOSRed
        End If
    End If

End Sub

Private Sub txtTOSHtRevised_GotFocus()
    Select_Text_In_Box txtTOSHtRevised
End Sub

Private Sub txtTOSHtRevised_LostFocus()
    Validate_Text_In_Box txtTOSHtRevised, labTOSHtRevised
    If Val(txtTOSHtRevised) > 0 Then
        txtTOSHt.Enabled = False
        labTOSHt.Enabled = False
    Else
        txtTOSHt.Enabled = True
        labTOSHt.Enabled = True
    End If
End Sub

Private Sub txtWaistCir_GotFocus()
    Select_Text_In_Box txtWaistCir
End Sub

Private Sub txtWaistCir_LostFocus()
    Validate_Text_In_Box txtWaistCir, labWaistCir
End Sub

Private Sub txtWaistHt_GotFocus()
    Select_Text_In_Box txtWaistHt
End Sub

Private Sub txtWaistHt_LostFocus()
    Validate_Text_In_Box txtWaistHt, labWaistHt
End Sub

Private Sub txtWaistHtRevised_GotFocus()
    Select_Text_In_Box txtWaistHtRevised
End Sub

Private Sub txtWaistHtRevised_LostFocus()
    Validate_Text_In_Box txtWaistHtRevised, labWaistHtRevised
    If Val(txtWaistHtRevised) > 0 Then
       txtWaistHt.Enabled = False
       labWaistHt.Enabled = False
    Else
       txtWaistHt.Enabled = True
       labWaistHt.Enabled = True
    End If
End Sub

Private Sub Update_DDE_Text_Boxes()
   'Called from OK_Click, assumes data has been validated
   'Updates the text boxes used to transfer DDE data

    Dim nLargestCir As Double, nCir As Double, nCrotchFrontFactor As Double
    Dim nOpenBack As Double, nOpenFront As Double
    Dim sString As String
    Dim iLeftleg As Integer, iRightleg As Integer, ii  As Integer
    Dim iLegStyle As Integer

   'Update from COMBO boxes
   'NB. Special case for TOS
    If txtTOSCir <> "" Then
        txtTOSRed.Text = cboTOSRed.List(cboTOSRed.ListIndex)
    End If
    txtWaistRed.Text = cboWaistRed.List(cboWaistRed.ListIndex)
    txtMidPointRed.Text = cboMidPointRed.List(cboMidPointRed.ListIndex)
    txtLargestRed.Text = cboLargestRed.List(cboLargestRed.ListIndex)
    txtThighRed.Text = cboThighRed.List(cboThighRed.ListIndex)
    txtLegStyle.Text = cboLegStyle.List(cboLegStyle.ListIndex)
    txtCrotchStyle.Text = cboCrotchStyle.List(cboCrotchStyle.ListIndex)

   'Calculate:-
   '    nCrotchFrontFactor
   '    nOpenBack
   '    nOpenFront
   'These are used in the DRAFIX macro WHCUTDBD.D, transfered via the
   'DB Field "Body".
   'The point of this is to Offload conditional decisions to VB where
   'it is faster, also to reduce the DRAFIX macro code.
   '
   'Establish Largest circumferance
   'NOTE:- this value is also used by the macro WH_LABL.D, passed through "ID"
   '
    nLargestCir = fnDisplayToInches(Val(txtLargestCir.Text))
    iLegStyle = cboLegStyle.ListIndex
   
   'nOpenBack, nOpenFront
    If txtCrotchStyle = "Open Crotch" Then
      If iLegStyle = 3 Or iLegStyle = 4 Then
       'W4OC-Left and W4OC-Right
        If txtAge <= 10 Then
            nOpenBack = 2#
            nOpenFront = 2#
        Else
            If nLargestCir >= 45 Then
                nOpenBack = 4#
                nOpenFront = 4#
            Else
                nOpenBack = 3.5
                nOpenFront = 3.5
            End If
        End If
      Else
       'All other leg styles
        If txtSex = "Male" Then
            If txtAge <= 10 Then
                nOpenBack = 1.5
                nOpenFront = 2.25
            Else
                If nLargestCir >= 45 Then
                    nOpenBack = 3.5
                    nOpenFront = 4.5
                Else
                    nOpenBack = 3#
                    nOpenFront = 4#
                End If
            End If
        Else
            If txtAge <= 10 Then
                nOpenBack = 2.25
                nOpenFront = 1.5
            Else
                If nLargestCir >= 45 Then
                    nOpenBack = 4.5
                    nOpenFront = 3.5
                Else
                    nOpenBack = 4#
                    nOpenFront = 3#
                End If
            End If
        End If
      End If
    End If

   'nCrotchFrontFactor
    If txtSex = "Male" And txtAge >= 3 Then
        If txtCrotchStyle = "Open Crotch" Then
            nCrotchFrontFactor = 0.5
        Else
            nCrotchFrontFactor = 0.4
        End If
    End If

    If txtSex = "Female" Or txtAge < 3 Then
            nCrotchFrontFactor = 0.5
    End If
    
   'Update DDE data field "Body". Overwrite previous values
   'As this field will be read using ScanLine(sBody, "blank", ...
   'Accept the leading blank given by Str$, as txtBody will be
   'blank delimited as required by ScanLine.
   '
   'Add the data for revised Waist and TOS heights
    sString = Str$(nCrotchFrontFactor) & Str$(nOpenFront) & Str$(nOpenBack)
    
    If chkDrawRevisedHts.Value = 1 Then
        sString = sString & " 1"    'Draw with revised hts
    Else
        sString = sString & " 0"    'Don't draw with revised hts
    End If

    sString = sString & Str$(Val(txtTOSHtRevised.Text)) & Str$(Val(txtWaistHtRevised.Text)) & Str$(Val(txtFoldHtRevised.Text))
    txtBody = sString

   'Leg Style data field
   'Update DDE data field "ID". Overwrite previous values
   'NOTE:- Includes largest circumferance data, used by WH_LABL.D
   '       GG 8.Nov.94 - Production modifications
   'This is now revised to incorporate data about the Left and Right Legs
   '                CODE       LEFT        RIGHT
   'Full-Leg        0       0 | 1 | 2       0 | 1 | 2
   'Panty           1       1 | 2           1 | 2
   'Brief           2       2               2
   'W4OC-Left       3       0 | 1           -1
   'W4OC-Right      4       -1              0 | 1
   'Chap-Left       5       0 | 1           -1
   'Chap-Right      6       -1              0 | 1
   'Chap Both Legs  7       0 | 1           0 | 1
   'Where:-
   '           -1   =>  No leg style
   '            0   =>  Full-Leg
   '            1   =>  Panty
   '            2   =>  Brief
   
    iRightleg = -1
    iLeftleg = -1
    For ii = 0 To 2
        If optLeftLeg(ii).Value = True Then iLeftleg = ii
        If optRightLeg(ii).Value = True Then iRightleg = ii
    Next ii
    sString = Str$(cboLegStyle.ListIndex) & " " & Str$(iLeftleg) & " " & Str$(iRightleg) & " " & Str$(nLargestCir)
    txtLegStyle = sString
    
End Sub

Private Function Validate_Data()
   'Checks data for holes etc.
   'Called from OK_Click

   'This breaks into two section the first is for fatal errors
   'that must be corrected

   'The second is for non-fatal error that require warning to the user
   'who can then fix them or not depending on requirements.
    

    Dim sError As String, NL As String, sString As String, sLegStyle As String
    Dim sHt As String

    Dim iIndex As Integer, iError As Integer, ii As Integer

    Dim nWaistHt As Double, nTOSHt As Double, nMaxCir As Double
    Dim nFoldHt As Double, nCir As Double, nDistance As Double
    
    Dim nLargestCir As Double, nLargestHt As Double, nLargestRed As Double
    Dim nLargestOff As Double, nCO_LargestBottOff As Double, nCO_LargestTopOff As Double
    Dim nButtockArcRad As Double, nCO_Diameter As Double

    Dim nDiff As Double, nCrotchFrontFactor As Double

    Dim nYscale As Double, nXscale As Double, nSeam As Double, nLength As Double
    Dim nDatum As Double

    Dim nThighCir As Double, nLeftThighCir As Double, nRightThighCir As Double
    Dim nThighRed As Double, nThighCirOriginal As Double

    Dim xyFold As XY, xyButtockArcCen As XY

    Dim nBodyBackCutOutMinTol As Double, nCutOutDiaMaxTol As Double


    NL = Chr$(10) 'new line
    sLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

   'FATAL Errors
   '~~~~~~~~~~~~
    If (txtTOSCir.Text = "" And txtTOSHt.Text <> "") Or (txtTOSCir.Text <> "" And txtTOSHt.Text = "") Then
        sError = sError + "Missing Top of support value/s." + NL
    End If
    
    If txtWaistCir.Text = "" Or txtWaistHt.Text = "" Then
        sError = sError + "Missing Waist value/s." + NL
    End If
    
    If txtMidPointCir.Text = "" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
        sError = sError + "Missing Mid point Circumferance." + NL
    End If
    
    If txtLargestCir.Text = "" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
        sError = sError + "Missing Largest part of buttocks Circumferance." + NL
    End If
    
    If txtLeftThighCir.Text = "" And sLegStyle <> "W4OC-Right" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
        sError = sError + "Missing Left Thigh value." + NL
    End If

    If txtRightThighCir.Text = "" And sLegStyle <> "W4OC-Left" And sLegStyle <> "Chap-Left" And sLegStyle <> "Chap-Right" And sLegStyle <> "Chap-Both Legs" Then
        sError = sError + "Missing Right Thigh value." + NL
    End If

    If txtFoldHt.Text = "" Then
        sError = sError + "Missing Fold Height value." + NL
    End If

   'Except for the Chap style
    If cboCrotchStyle.ListIndex < 0 And cboLegStyle.ListIndex < 5 Then
        sError = sError + "Missing Crotch Style." + NL
    End If

    If cboLegStyle.ListIndex < 0 Then
        sError = sError + "Missing Leg Style." + NL
    End If
        
   'Check Crotch style against sex
    sString = cboCrotchStyle.List(cboCrotchStyle.ListIndex)

   'Display Error message (if required) and return
   'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        Validate_Data = False
        Exit Function
    Else
        Validate_Data = True
    End If

   'NON-FATAL Errors (warnings)
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~
   'Establish Largest circumferance
    nMaxCir = fnDisplayToInches(Val(txtLargestCir.Text))

   'Establish hts (original or revised)
   'If no revised ht given then use the original
    If chkDrawRevisedHts.Value = 1 Then
        nFoldHt = fnDisplayToInches(Val(txtFoldHtRevised.Text))
        nWaistHt = fnDisplayToInches(Val(txtWaistHtRevised.Text))
        nTOSHt = fnDisplayToInches(Val(txtTOSHtRevised.Text))
       'If no revised ht given then use the original
        If nFoldHt = 0 Then nFoldHt = fnDisplayToInches(Val(txtFoldHt.Text))
        If nWaistHt = 0 Then nWaistHt = fnDisplayToInches(Val(txtWaistHt.Text))
        If nTOSHt = 0 Then nTOSHt = fnDisplayToInches(Val(txtTOSHt.Text))
    Else
        nFoldHt = fnDisplayToInches(Val(txtFoldHt.Text))
        nWaistHt = fnDisplayToInches(Val(txtWaistHt.Text))
        nTOSHt = fnDisplayToInches(Val(txtTOSHt.Text))
    End If

   'Get difference between TOS and fold (note case for explicit TOS)
    If nTOSHt > 0 Then
        nDiff = nTOSHt - nFoldHt
        sHt = "TOS"
    Else
        nDiff = nWaistHt - nFoldHt
        sHt = "Waist"
    End If

   'Check Fold Height to TOS/Waist Height
    If nMaxCir <= 39 And nDiff < 5 And nDiff > 0 Then
        sError = sError + "Largest Part of Buttocks Circumference is" + Str$(nMaxCir) + " inches" + NL
        sError = sError + "Fold to " + sHt + " distance is" + Str$(nDiff) + " inches" + NL
        sError = sError + NL
        sError = sError + "If the buttocks circumferance is :-" + NL
        sError = sError + "less than 39 inches, there must be at least 5 inches" + NL
        sError = sError + "from fold to the " + sHt + "." + NL
        sError = sError + "ref:- 8.1.1.1, GOP 01-02/12, page 5 of 24" + NL + NL
        iError = True
    End If

   'Check Fold Height to TOS/Waist Height
    If nMaxCir > 39 And nMaxCir <= 44.875 And ((nDiff < 10 And txtSex = "Male") Or (nDiff < 12 And txtSex = "Female")) And nDiff > 0 Then
        sError = sError + "Largest Part of Buttocks Circumference is" + Str$(nMaxCir) + " inches" + NL
        sError = sError + "Fold to " + sHt + " distance is" + Str$(nDiff) + " inches" + NL
        sError = sError + NL
        sError = sError + "If the buttocks circumferance is :-" + NL
        sError = sError + "between 39 and 44-7/8 inches, there must be at least" + NL
        sError = sError + "10 inches for Males and 12 inches for Females from the" + NL
        sError = sError + "fold to the " + sHt + "." + NL
        sError = sError + "ref:- 8.1.1.2, GOP 01-02/12, page 5 of 24" + NL + NL
        iError = True
    End If

   'Check Fold Height to TOS/Waist Height
    If nMaxCir > 44.875 And nDiff < 15 And nDiff > 0 Then
        sError = sError + "Largest Part of Buttocks Circumference is" + Str$(nMaxCir) + " inches" + NL
        sError = sError + "Fold to " + sHt + " distance is" + Str$(nDiff) + " inches" + NL
        sError = sError + NL
        sError = sError + "If the buttocks circumferance is :-" + NL
        sError = sError + "greater than 45 inches, there must be at least 15 inches" + NL
        sError = sError + "from the fold to the " + sHt + "." + NL
        sError = sError + "ref:- 8.1.1.3, GOP 01-02/12, page 5 of 24" + NL + NL
        iError = True
    End If

   'Check distance from waist to Explicity given TOS
    nDiff = nTOSHt - nWaistHt

    If (nDiff < 1.5 Or nDiff > 5) And nTOSHt > 0 Then
        sError = sError + "TOS to Waist distance is" + Str$(nDiff) + " inches" + NL
        sError = sError + NL
        sError = sError + "The distance between the Waist and the TOS should be no" + NL
        sError = sError + "less than 1-1/2 inches or more than 5 inches" + NL
        sError = sError + "ref:- 8.2 GOP 01-02/12, page 5 of 24" + NL + NL
        iError = True
    End If

   'Thigh figuring
   'Take account of W4OC and CHAP w.r.t. thigh circumferences
    Select Case cboLegStyle.ListIndex
        Case 0, 1, 2, 7
            nRightThighCir = fnDisplayToInches(Val(txtRightThighCir.Text))
            nLeftThighCir = fnDisplayToInches(Val(txtLeftThighCir.Text))
        
        Case 3, 5   'W4OC-Left and CHAP-Left
            nLeftThighCir = fnDisplayToInches(Val(txtLeftThighCir.Text))
            nRightThighCir = nLeftThighCir

        Case 4, 6  'W4OC-Right and CHAP-Right
            nRightThighCir = fnDisplayToInches(Val(txtRightThighCir.Text))
            nLeftThighCir = nRightThighCir

    End Select

    nThighRed = Val(cboThighRed.List(cboThighRed.ListIndex))
    nThighCirOriginal = min(nRightThighCir, nLeftThighCir)
    nThighCir = fnRoundInches(nThighCirOriginal * nThighRed)
   
   'Check for male crotch and female patient
    iIndex = 0
    iIndex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
    If iIndex <> 0 And txtSex = "Female" Then
        sError = sError + "Male Crotch type specified for Female patient" + NL
    End If

   'Check for female crotch and male patient
    iIndex = 0
    iIndex = InStr(1, sString, "Female", 0)
    If iIndex <> 0 And txtSex = "Male" Then
        sError = sError + "Female Crotch type specified for Male patient" + NL
    End If


   'Display Error message (if required) and return
   'These are non-fatal errors
    sError = sError + NL + "The above problems have been found in the data do you" + NL
    sError = sError + "wish to continue?"
    
    If iError = True Then
        iError = MsgBox(sError, 36, "Problems with data")
        If iError = IDYES Then
            Validate_Data = True
        Else
            Validate_Data = False
            Exit Function
        End If
    Else
        Validate_Data = True
    End If

   'nCrotchFrontFactor
    If txtSex.Text = "Male" And Val(txtAge.Text) >= 3 Then
        If cboCrotchStyle.List(cboCrotchStyle.ListIndex) = "Open Crotch" Then
            nCrotchFrontFactor = 0.5
        Else
            nCrotchFrontFactor = 0.4
        End If
    End If

    If txtSex.Text = "Female" Or Val(txtAge.Text) < 3 Then
            nCrotchFrontFactor = 0.5
    End If

   'Check Cut-Out values (excluding CHAP styles)

   If cboLegStyle.ListIndex < 5 Then
   'Heights are figured with the fold height as the datum
    nSeam = 0.1875
    nYscale = 0.5
    nXscale = 1.25 / 1.5
    iError = 0
    sError = ""
    iIndex = -1 'Also acts as a flag

    nDatum = Int(nFoldHt / 1.5) * 1.5

   'Largest part of buttocks figuring
    nLargestHt = ((nWaistHt - nFoldHt) / 3) + nFoldHt

    nLargestHt = fnRoundInches((nLargestHt - nDatum) * nXscale)
   
    nFoldHt = fnRoundInches((nFoldHt - nDatum) * nXscale)

    xyFold.X = nFoldHt
    xyFold.Y = nThighCir * nYscale + nSeam

    xyButtockArcCen.X = nLargestHt
    xyButtockArcCen.Y = ((nThighCir / 2) * nYscale) + nSeam

    nButtockArcRad = FN_CalcLength(xyFold, xyButtockArcCen)

    nLargestOff = xyButtockArcCen.Y + nButtockArcRad

   'NOTE:-
   '    In both of the following tests we loop only twice.
   '    There can be no more that 5% difference between the reductions
   '    at each circumference.

   'Loop to test for 6" cut out
   'We also test to see if the cut-out is -ve, this will only occur
   'with v.small body
   'Test and fix the first time - give information message
   'The second time - give error message
    For ii = 1 To 2
        nLargestRed = Val(cboLargestRed.List(cboLargestRed.ListIndex))
        nLargestCir = fnDisplayToInches(Val(txtLargestCir.Text))
        'nLargestCir = fnRoundInches(nLargestCir * nLargestRed)
        nLargestCir = nLargestCir * nLargestRed

        nLength = nLargestOff - nSeam

        nCO_LargestBottOff = (((nLargestCir * nYscale) - nLength) * nCrotchFrontFactor) + nSeam
        nCO_LargestTopOff = nLargestOff - (((nLargestCir * nYscale) - nLength) * (1 - nCrotchFrontFactor))
        'nCO_Diameter = fnRoundInches(nCO_LargestTopOff - nCO_LargestBottOff)
        nCO_Diameter = nCO_LargestTopOff - nCO_LargestBottOff

        nCutOutDiaMaxTol = 6
        If nCO_Diameter < nCutOutDiaMaxTol And nCO_Diameter > 0 Then
            Exit For
        End If
        If nCO_Diameter <= 0 Then
           'Negative or Zero Cut-Out
            If ii = 1 Then
                iIndex = cboLargestRed.ListIndex - 5
                If iIndex < 0 Then
                    cboLargestRed.ListIndex = 0
                Else
                    cboLargestRed.ListIndex = iIndex
                End If
                sError = sError + "With the given Buttock Reduction, Cut-Out diameter" + NL
                sError = sError + "is negative or zero!" + NL
                sError = sError + "Reduction decreased by 5% to rectify this." + NL
                iError = iError + 1
            Else
                sError = sError + NL + "Severe - Warning!" + NL
                sError = sError + "Cut-Out is still negative or zero!" + NL
                sError = sError + "Even with the reduction decreased by 5%." + NL
                iError = iError + 4
            End If

        Else
           'Cut-Out greater than 6"
            If ii = 1 Then
                iIndex = cboLargestRed.ListIndex + 5
                If iIndex >= cboLargestRed.ListCount Then
                    cboLargestRed.ListIndex = cboLargestRed.ListCount - 1
                Else
                    cboLargestRed.ListIndex = iIndex
                End If
                sError = sError + "With given Buttock Reduction, Cut-Out is larger than 6 inches" + NL
                sError = sError + "Reduction increased by 5% to rectify this." + NL
                sError = sError + "ref:- 13.4.3.1 GOP 01-02/12, page 7 of 24" + NL
                iError = iError + 1
            Else
                sError = sError + NL + "Severe - Warning!" + NL
                sError = sError + "Cut-Out is still larger than 6 inches." + NL
                sError = sError + "Even with reduction already increased by 5%." + NL
                iError = iError + 4
            End If
        End If
    Next ii
   
   'Check that minimum distance of 3" from Back Cut-Out to Largest part of
   'buttocks is met
   'Loop to test for 3" minimum
   'Test and fix the first time - give information message.
   '    Except if the reduction has been already modified above wrt
   '    6" Cut-Out maximum.
   '    Uses iIndex as the flag for this.
   'The second time - give error message.
    For ii = 1 To 2
        nLargestRed = Val(cboLargestRed.List(cboLargestRed.ListIndex))
        nLargestCir = fnDisplayToInches(Val(txtLargestCir.Text))
        nLargestCir = fnRoundInches(nLargestCir * nLargestRed)
        
        nCO_LargestTopOff = nLargestOff - (((nLargestCir * nYscale) - nLength) * (1 - nCrotchFrontFactor))
        nDistance = fnRoundInches(nLargestOff - nCO_LargestTopOff)
        
        nBodyBackCutOutMinTol = 3
        If nDistance >= (nBodyBackCutOutMinTol * nYscale) Then
            Exit For
        Else
            If ii = 1 Then
                If iIndex < 0 Then
                    iIndex = cboLargestRed.ListIndex + 5
                    sError = sError + "With given Buttock Reduction, distance from back of Cut-Out" + NL
                    sError = sError + "to largest part of buttocks is less than 3 inches!" + NL
                    sError = sError + "Reduction increased by 5% to rectify this." + NL
                    sError = sError + "ref:- 16.1 GOP 01-02/12, page 8 of 24" + NL
                End If
                If iIndex >= cboLargestRed.ListCount Then
                    cboLargestRed.ListIndex = cboLargestRed.ListCount - 1
                Else
                    cboLargestRed.ListIndex = iIndex
                End If
                iError = iError + 2
            Else
                sError = sError + NL + "Severe - Warning!" + NL
                sError = sError + "Distance from back of Cut-Out to largest part of buttocks" + NL
                sError = sError + "is still less than 3 inches!" + NL
                sError = sError + "Even with reduction already increased by 5%." + NL
                iError = iError + 8
            End If
        End If
    Next ii

   'Check that all of the reductions are within 5% of each other
    Dim nMinRed%, nMaxRed%
    nMinRed = 15
    nMaxRed = 0
    
    If cboTOSRed.ListIndex <> -1 And cboTOSRed.ListIndex < nMinRed And Val(txtTOSCir.Text) <> 0 Then nMinRed = cboTOSRed.ListIndex
    If cboTOSRed.ListIndex <> -1 And cboTOSRed.ListIndex > nMaxRed And Val(txtTOSCir.Text) <> 0 Then nMaxRed = cboTOSRed.ListIndex
    
    If cboWaistRed.ListIndex <> -1 And cboWaistRed.ListIndex < nMinRed Then nMinRed = cboWaistRed.ListIndex
    If cboWaistRed.ListIndex <> -1 And cboWaistRed.ListIndex > nMaxRed Then nMaxRed = cboWaistRed.ListIndex
    
    If cboMidPointRed.ListIndex <> -1 And cboMidPointRed.ListIndex < nMinRed Then nMinRed = cboMidPointRed.ListIndex
    If cboMidPointRed.ListIndex <> -1 And cboMidPointRed.ListIndex > nMaxRed Then nMaxRed = cboMidPointRed.ListIndex
    
    If cboLargestRed.ListIndex <> -1 And cboLargestRed.ListIndex < nMinRed Then nMinRed = cboLargestRed.ListIndex
    If cboLargestRed.ListIndex <> -1 And cboLargestRed.ListIndex > nMaxRed Then nMaxRed = cboLargestRed.ListIndex

    If cboThighRed.ListIndex <> -1 And cboThighRed.ListIndex < nMinRed Then nMinRed = cboThighRed.ListIndex
    If cboThighRed.ListIndex <> -1 And cboThighRed.ListIndex > nMaxRed Then nMaxRed = cboThighRed.ListIndex
   
    If (nMaxRed - nMinRed) > 5 Then
        sError = sError + NL + "Severe - Warning!" + NL
        sError = sError + "All Reductions should be within 5% of each other" + NL
        iError = iError + 8
    End If
   
   'Display Error message (if required) and return
   'These are non-fatal errors
    Select Case iError
        Case 0
            Validate_Data = True
        Case 1 To 3
            MsgBox sError, 64, "Warning - Problems with data"
            Validate_Data = False
        Case 4 To 100
            sError = sError + NL + "The above problems have been found in the data do you" + NL
            sError = sError + "wish to continue ?"
            iError = MsgBox(sError, 52, "Severe Problems with data")
            If iError = IDYES Then
                Validate_Data = True
            Else
                Validate_Data = False
            End If
        Case Else
            Validate_Data = True
    End Select

   End If  'End of If to exclude CHAP leg styles

End Function


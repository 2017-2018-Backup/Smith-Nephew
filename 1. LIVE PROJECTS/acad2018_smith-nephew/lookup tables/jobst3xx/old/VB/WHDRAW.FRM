VERSION 5.00
Begin VB.Form whdraw 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Waist Height - DRAW"
   ClientHeight    =   7560
   ClientLeft      =   1815
   ClientTop       =   1200
   ClientWidth     =   8670
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
   LinkTopic       =   "whdraw"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7560
   ScaleWidth      =   8670
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6240
      TabIndex        =   117
      Top             =   5640
      Width           =   435
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      TabIndex        =   115
      Top             =   1170
      Width           =   400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   6210
   End
   Begin VB.TextBox txtleftFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   112
      Top             =   7050
      Width           =   400
   End
   Begin VB.TextBox txtRightFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   111
      Top             =   7080
      Width           =   400
   End
   Begin VB.TextBox txtRightData 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   6690
      Width           =   400
   End
   Begin VB.TextBox txtLeftData 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   6660
      Width           =   400
   End
   Begin VB.TextBox txtRightFootLength 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   6330
      Width           =   400
   End
   Begin VB.TextBox txtLeftFootLength 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   6300
      Width           =   400
   End
   Begin VB.TextBox txtRightToeStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   5970
      Width           =   400
   End
   Begin VB.TextBox txtLeftToeStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   5940
      Width           =   400
   End
   Begin VB.TextBox txtRightTopLegPleat2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   5610
      Width           =   400
   End
   Begin VB.TextBox txtLeftTopLegPleat2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   5580
      Width           =   400
   End
   Begin VB.TextBox txtRightTopLegPleat1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   5250
      Width           =   400
   End
   Begin VB.TextBox txtLeftTopLegPleat1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   5220
      Width           =   400
   End
   Begin VB.TextBox txtDrawType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      TabIndex        =   10
      Top             =   750
      Width           =   400
   End
   Begin VB.TextBox txtTOSCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2370
      Width           =   400
   End
   Begin VB.TextBox txtTOSHt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2370
      Width           =   400
   End
   Begin VB.TextBox txtLargestCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3450
      Width           =   400
   End
   Begin VB.TextBox txtLargestHt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   14
      Top             =   3450
      Width           =   400
   End
   Begin VB.TextBox txtLeftThighCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   15
      Top             =   3810
      Width           =   400
   End
   Begin VB.TextBox txtWaistCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   16
      Top             =   2730
      Width           =   400
   End
   Begin VB.TextBox txtWaistHt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   17
      Top             =   2730
      Width           =   400
   End
   Begin VB.TextBox txtMidPointCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   18
      Top             =   3090
      Width           =   400
   End
   Begin VB.TextBox txtMidPointHt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   19
      Top             =   3090
      Width           =   400
   End
   Begin VB.TextBox txtRightThighCir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   20
      Top             =   4170
      Width           =   400
   End
   Begin VB.TextBox txtFoldHt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   21
      Top             =   4500
      Width           =   400
   End
   Begin VB.TextBox txtMidPointRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   22
      Top             =   3090
      Width           =   400
   End
   Begin VB.TextBox txtWaistRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   23
      Top             =   2730
      Width           =   400
   End
   Begin VB.TextBox txtThighRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   24
      Top             =   3990
      Width           =   400
   End
   Begin VB.TextBox txtLargestRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   25
      Top             =   3450
      Width           =   400
   End
   Begin VB.TextBox txtTOSRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   26
      Top             =   2370
      Width           =   400
   End
   Begin VB.TextBox txtCrotchStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      TabIndex        =   27
      Top             =   5250
      Width           =   400
   End
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      TabIndex        =   28
      Top             =   4890
      Width           =   400
   End
   Begin VB.TextBox txtLeftFootPleat1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   29
      Top             =   4500
      Width           =   400
   End
   Begin VB.TextBox txtRightFootPleat1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   36
      Top             =   4530
      Width           =   400
   End
   Begin VB.TextBox txtLeftFootPleat2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   39
      Top             =   4860
      Width           =   400
   End
   Begin VB.TextBox txtRightFootPleat2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   40
      Top             =   4890
      Width           =   400
   End
   Begin VB.TextBox txtUidBody 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   48
      Top             =   630
      Width           =   400
   End
   Begin VB.TextBox txtRightRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   96
      Top             =   4170
      Width           =   400
   End
   Begin VB.TextBox txtLeftRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   97
      Top             =   4140
      Width           =   400
   End
   Begin VB.TextBox txtRightStr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   98
      Top             =   3810
      Width           =   400
   End
   Begin VB.TextBox txtLeftStr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   99
      Top             =   3780
      Width           =   400
   End
   Begin VB.TextBox txtLeftMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   100
      Top             =   3420
      Width           =   400
   End
   Begin VB.TextBox txtRightMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   101
      Top             =   3450
      Width           =   400
   End
   Begin VB.TextBox txtRightTemplate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   95
      Top             =   3090
      Width           =   400
   End
   Begin VB.TextBox txtUidRightLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   93
      Top             =   990
      Width           =   400
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7980
      TabIndex        =   90
      Top             =   4890
      Width           =   400
   End
   Begin VB.TextBox txtLeftTemplate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   88
      Top             =   3060
      Width           =   400
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   360
      TabIndex        =   51
      Top             =   150
      Width           =   3435
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   49
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   50
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   57
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   58
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtUnits 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   59
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   60
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Age:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File No:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sex:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   1980
         TabIndex        =   55
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Units:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Diag:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7980
      TabIndex        =   46
      Top             =   5250
      Width           =   400
   End
   Begin VB.TextBox txtRightLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   43
      Top             =   2370
      Width           =   400
   End
   Begin VB.TextBox txtRightAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   42
      Top             =   2730
      Width           =   400
   End
   Begin VB.TextBox txtLeftAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   38
      Top             =   2700
      Width           =   400
   End
   Begin VB.TextBox txtLeftLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   35
      Top             =   2340
      Width           =   400
   End
   Begin VB.TextBox txtUidLeftLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   31
      Top             =   1350
      Width           =   400
   End
   Begin VB.TextBox txtUidTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   30
      Top             =   270
      Width           =   400
   End
   Begin VB.CommandButton OK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   61
      Top             =   6330
      Width           =   1215
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ID"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4980
      TabIndex        =   118
      Top             =   5730
      Width           =   1215
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "txtWorkOrder"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   5640
      TabIndex        =   116
      Top             =   1230
      Width           =   1155
   End
   Begin VB.Label Label46 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   114
      Top             =   7110
      Width           =   1275
   End
   Begin VB.Label Label45 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   113
      Top             =   7140
      Width           =   1275
   End
   Begin VB.Label Label44 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   62
      Top             =   6750
      Width           =   1275
   End
   Begin VB.Label Label43 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   63
      Top             =   6720
      Width           =   1275
   End
   Begin VB.Label Label42 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootLength"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   64
      Top             =   6390
      Width           =   1275
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootLength"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   65
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Label Label40 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ToeStyle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   66
      Top             =   6030
      Width           =   1275
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TopLegPleat2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   67
      Top             =   5670
      Width           =   1275
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TopLegPleat1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   68
      Top             =   5310
      Width           =   1275
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootPleat2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   69
      Top             =   4950
      Width           =   1275
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootPleat1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   70
      Top             =   4590
      Width           =   1275
   End
   Begin VB.Label Label39 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ToeStyle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   71
      Top             =   6000
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TopLegPleat2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   72
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TopLegPleat1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   73
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Draw Type"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   5880
      TabIndex        =   74
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3900
      TabIndex        =   75
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label Label38 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RED"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7200
      TabIndex        =   76
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label labTOS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TOS"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   77
      Top             =   2430
      Width           =   1005
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CIR."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6300
      TabIndex        =   78
      Top             =   2130
      Width           =   315
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HTS"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6720
      TabIndex        =   79
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label labLargest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Buttocks"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   80
      Top             =   3510
      Width           =   1005
   End
   Begin VB.Label labLeftThigh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Thigh - Left"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   81
      Top             =   3870
      Width           =   1005
   End
   Begin VB.Label labWaist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Waist"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   82
      Top             =   2790
      Width           =   1005
   End
   Begin VB.Label labMidPoint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Midpoint"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   83
      Top             =   3150
      Width           =   1005
   End
   Begin VB.Label labRightThigh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Thigh - Right"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   84
      Top             =   4230
      Width           =   1185
   End
   Begin VB.Label labFold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fold height"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   85
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Crotch Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4980
      TabIndex        =   86
      Top             =   5310
      Width           =   1275
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Body"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4980
      TabIndex        =   87
      Top             =   4950
      Width           =   1095
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootPleat1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   102
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FootPleat2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   103
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Body UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3900
      TabIndex        =   104
      Top             =   690
      Width           =   1275
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   107
      Top             =   4230
      Width           =   1275
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   108
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   109
      Top             =   3870
      Width           =   1275
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   110
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   105
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   106
      Top             =   3510
      Width           =   1275
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Templt"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   92
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3900
      TabIndex        =   94
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6720
      TabIndex        =   91
      Top             =   4950
      Width           =   1275
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Templt"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   89
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6720
      TabIndex        =   47
      Top             =   5310
      Width           =   1275
   End
   Begin VB.Label lable14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   45
      Top             =   2430
      Width           =   1275
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   44
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   41
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   37
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3900
      TabIndex        =   33
      Top             =   1410
      Width           =   1275
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   32
      Top             =   -390
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "whdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   WHDRAW.FRM
'Purpose:   Drawing Arms and Vest sleeves.
'
'Version:   3.0
'Date:      94/95
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Feb.95     GG      TRITON to CADLINK
'
'Dec 98     GG      Ported to VB5
'
'
'Note:-
'
   'This module is used as the decision making front end to
   'all of the drawing functions for the Waist Height
   'This means that the actual DRAFIX macros can be tailored to
   'load only those Macros that are relevent to the required drawing.
   'E.G.
   '    Drawing a footless leg means that WHFTPNTS.D need not be used.
   '    Drawing of a 2nd leg with no cutout means that the WHLG2DBD.D
   '    is not required.
   '    Etc.
   '
   '* Windows API Functions Declarations
'    Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
'
'   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

   'MsgBox constants
'    Const IDYES = 6
'    Const IDNO = 7

    Option Explicit

Private Function FN_ChapOpen(sDrafixFile As String) As Integer
   'Open the DRAFIX macro file that draws the CHAP
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_ChapOpen = fNum


   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX CHAP Drawing Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; g_sPatient; CC; " "; g_sFileNo; CC; " SIDE - "; g_sSide
    Print #fNum, "//by Visual Basic"
    Print #fNum, "HANDLE hEnt, hLayer;"
    Print #fNum, "XY xyStart;"
    Print #fNum, "STRING sID;"
    
    Print #fNum, "xyStart.x=0;"
    Print #fNum, "xyStart.y=0;"
    PR_PutStringAssign "sID", txtID


End Function


Private Function FN_WHOpen(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_WHOpen = fNum

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Waist Height Drawing Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; g_sPatient; CC; " "; g_sFileNo; CC; " SIDE - "; g_sSide
    Print #fNum, "//by Visual Basic"

   'Ensure that style and color are by layer
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; ")) ;"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; ")) ;"
   
   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"

End Function


Private Function fnGetNumber(ByVal sString As String, iIndex As Integer) As Double
   'Function to return as a numerical value the iIndexth item in a string
   'that uses blanks (spaces) as delimiters.
   'EG
   '    sString = "12.3 65.1 45"
   '    fnGetNumber( sString, 2) = 65.1
   '
   'If the iIndexth item is not found then return -1 to indicate an error.
   'This assumes that the string will not be used to store -ve numbers.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

    If Len(sString) = 0 Then
        fnGetNumber = -1
        Exit Function
    End If

   'Prepare string
    sString = sString & " " 'Trailing blank as stopper for last item

   'Get iIndexth item
    For ii = 1 To iIndex
        iPos = InStr(sString, " ")
        If ii = iIndex Then
            sString = Left$(sString, iPos - 1)
            fnGetNumber = Val(sString)
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos))
            If Len(sString) = 0 Then
                fnGetNumber = -1
                Exit Function
            End If
        End If
    Next ii

   'The function should have exited befor this, however just in case
   '(iIndex = 0) we indicate an error,
    fnGetNumber = -1

End Function

Private Sub Form_LinkClose()
   Dim ii As Integer, nn As Integer, nAge  As Integer, iFabricClass As Integer
   Dim iMMHg As Integer, iZipper As Integer, iValue As Integer
   Dim nValue As Double, nStretch As Double
   Dim iLegStyle%

   Timer1.Enabled = False
  
  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

  'Work order
   If txtWorkOrder.Text = "" Then
        g_sWorkOrder = "-"
   Else
        g_sWorkOrder = txtWorkOrder.Text
   End If
     
   iLegStyle = fnGetNumber(txtLegStyle, 1)
   Select Case txtDrawType
        Case "WH1LG"
            g_sDialogueID = "First Leg Drawing"
            If iLegStyle = 2 Then
               'Brief
                PR_WH_Brief1stLeg
            ElseIf iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then
               'ChapStyle
                PR_WH_ChapLeg (txtDrawType.Text)
            Else
               'Panty or full leg
                PR_WH_1stLeg
            End If

        Case "WH2LG"
            g_sDialogueID = "Second Leg Drawing"
            If iLegStyle = 2 Then
               'Brief
                PR_WH_Brief2ndLeg
            ElseIf iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then
               'ChapStyle
                PR_WH_ChapLeg (txtDrawType.Text)
            Else
               'Panty or full leg
                PR_WH_2ndLeg
            End If
        
        Case "WHCUT"
            g_sDialogueID = "Cut Out Drawing"
            If iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then
                MsgBox "Can't Draw a CUT-OUT for Chap styles", 16, g_sDialogueID
                End
            End If
            PR_WH_CutOut

        Case "CHAPBODYLeft", "CHAPBODYRight"
            g_sDialogueID = "Chap Drawing"
            PR_WH_ChapBody Mid$(txtDrawType.Text, 9)

        Case Else
            End
   End Select
   Screen.MousePointer = 0    ' Change pointer to default.
   AppActivate fnGetDrafixWindowTitleText()
   SendKeys "@C:\JOBST\DRAW.D{enter}", True
   End
'Show
End Sub


Private Sub Form_LinkExecute(CmdStr As String, cancel As Integer)
    If CmdStr = "Cancel" Then
        cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Dim ii As Integer
    Hide
   'Check if a previous instance is running
   'If it is warn user and exit

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - whdraw.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - whdraw.Height) / 2  ' Center form vertically.
    
    Set MainForm = whdraw

    g_nUnitsFac = 1 'Default to inches

    g_sPathJOBST = fnPathJOBST()

   'Initialise String globals
    CC = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & CC & QQ  'Quote Comma Quote ( "," )
    QC = QQ & CC        'Quote Comma ( ", )
    CQ = CC & QQ        'Comma Quote ( ," )

    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 8    'Bottom
    g_nCurrTextFont = 0         'BLOCK


   'Process will Time Out after approx 10 secs
    Timer1.Interval = 10000
    Timer1.Enabled = True

End Sub

Private Sub OK_Click()
   'Not available in normal use
    Dim sTask$
    sTask = fnGetDrafixWindowTitleText()
    AppActivate sTask
    SendKeys "@C:\JOBST\DRAW.D{enter}", True
    End
End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable
   '
    Print #fNum, sLine

End Sub

Private Sub PR_PutNumberAssign(sVariableName As String, nAssignedNumber As Variant)

   'Procedure to put a number assignment
   'Adds a semi-colon

    Print #fNum, sVariableName; "="; Str$(nAssignedNumber); ";"


End Sub

Private Sub PR_PutStringAssign(sVariableName As String, sAssignedString As Variant)
   'Procedure to put a string assignment
   'Encloses String in quotes and adds a semi-colon

    Print #fNum, sVariableName; "="; QQ; sAssignedString; QQ; ";"

End Sub

Private Sub PR_StandardDetails()
    Dim nAge%

   'Path to JOBST
    PR_PutStringAssign "sPathJOBST", FN_EscapeSlashesInString(g_sPathJOBST)

   'Patient Details
    PR_PutStringAssign "sPatient", txtPatientName
    PR_PutStringAssign "sFileNo", txtFileNo
    PR_PutStringAssign "sWorkOrder", g_sWorkOrder
    
    PR_PutStringAssign "sAge", txtAge
    nAge = Val(txtAge.Text)
    PR_PutNumberAssign "nAge", nAge
    
    PR_PutStringAssign "sSEX", txtSex
    If txtSex.Text = "Male" Then
        PR_PutLine "Male = %true;"
        PR_PutLine "Female = %false;"
    Else
        PR_PutLine "Male = %false;"
        PR_PutLine "Female = %true;"
    End If

    PR_PutStringAssign "sUnits", txtUnits
    PR_PutNumberAssign "nUnitsFac", g_nUnitsFac
    
    PR_PutStringAssign "sDiagnosis", txtDiagnosis

End Sub

Private Sub PR_WH_1stLeg()
   'Procedure to create a Macro to draw the first leg of a
   'Waist Height
    Dim sFile As String, sSide  As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim ii As Integer, Footless As Integer, itemplate As Integer
    Dim nFoldHt As Double, nValue As Double

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    nLeftLegStyle = fnGetNumber(txtLegStyle.Text, 2)
    nRightLegStyle = fnGetNumber(txtLegStyle.Text, 3)
    
    sFile = "C:\JOBST\DRAW.D"

   'Establish which is the First Leg
    Select Case nLegStyle
        Case 0, 1
            If fnDisplaytoInches(Val(txtLeftThighCir)) - fnDisplaytoInches(Val(txtRightThighCir)) > 1 Then
                sSide = "Right"
            Else
                sSide = "Left"
            End If
        Case 2
            sSide = "Left"
        Case 3, 5
            sSide = "Left"
        Case 4, 6
            sSide = "Right"
    End Select

   'Having decided on leg do a swift check to see if there is enough
   'data to draw the leg
    If sSide = "Left" And txtLeftAnkle = "" And txtLeftLengths = "" Then
        MsgBox "No data for LEFT Leg", 16, g_sDialogueID
        End
    End If

    If sSide = "Left" And txtLeftTemplate = "" Then
        MsgBox "Figure LEFT Leg before drawing", 16, "First Leg Drawing"
        End
    End If

    If sSide = "Right" And txtRightAnkle = "" And txtRightLengths = "" Then
        MsgBox "No data for RIGHT Leg", 16, g_sDialogueID
        End
    End If

    If sSide = "Right" And txtRightTemplate = "" Then
        MsgBox "Figure RIGHT Leg before drawing", 16, g_sDialogueID
        End
    End If


'    If sSide = "Left" And fnGetNumber(txtLeftAnkle.Text, 1) = -1 And nLeftLegStyle <> 1 Then
'        MsgBox "A Full leg has been chosen in the Body, but the data in the leg is footless.  ", 16, "First Leg Drawing"
'        End
'    End If


   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)

    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_LEG1.D;"
   
   'Put standard patient details
    PR_StandardDetails

   'Body details
    PR_PutStringAssign "sTOSCir", txtTOSCir
   'Fold Height (Height to fold of buttocks), Used revised if given
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    If fnGetNumber(txtBody.Text, 4) > 0 And fnGetNumber(txtBody.Text, 7) > 0 Then
        nFoldHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 7))
    End If
    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutStringAssign "sCrotchStyle", txtCrotchStyle
    PR_PutStringAssign "sFabric", txtFabric

    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutNumberAssign "nLeftLegStyle", nLeftLegStyle
    PR_PutNumberAssign "nRightLegStyle", nRightLegStyle
    
   'Leg Details
    If sSide = "Left" Then
        PR_PutLine "LeftLeg = %true;"
        PR_PutLine "RightLeg = %false;"
        sLengths = txtLeftLengths.Text  'Use later to establish 1st & Last tapes
       'Heel value
       'Overide the LegStyle if there is no heel value
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        
        If nLeftLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtLeftAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtLeftAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtLeftAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtLeftAnkle.Text, 4)
        End If
        
        nFabricClass = fnGetNumber(txtLeftAnkle.Text, 8)
        If nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtLeftRed
            PR_PutStringAssign "sTapeMMs", txtLeftMMs
            PR_PutStringAssign "sStretch", txtLeftStr
        End If

        PR_PutStringAssign "sFirstLeg", "Left"
        PR_PutStringAssign "sLeg", "Left"
        PR_PutStringAssign "sSecondLeg", "Right"

        PR_PutStringAssign "sPressure", txtLeftTemplate
        itemplate = Val(Left$(txtLeftTemplate.Text, 2))
        
        PR_PutStringAssign "sTapeLengths", txtLeftLengths
        PR_PutStringAssign "sToeStyle", txtLeftToeStyle
        
        PR_PutNumberAssign "nFootPleat1", fnDisplaytoInches(Val(txtLeftFootPleat1.Text))
        PR_PutNumberAssign "nFootPleat2", fnDisplaytoInches(Val(txtLeftFootPleat2.Text))
        PR_PutNumberAssign "nTopLegPleat1", fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text))
        PR_PutNumberAssign "nTopLegPleat2", fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text))
        
        PR_PutStringAssign "sFootLength", txtLeftFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtLeftData.Text, 1)

       'Other leg details used to control toe
        nValue = fnGetNumber(txtRightAnkle.Text, 1)
        If nRightLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If
                                                      
    Else
        PR_PutLine "LeftLeg = %false;"
        PR_PutLine "RightLeg = %true;"
        sLengths = txtRightLengths.Text  'Use later to establish 1st & Last tapes
       'Heel value
       'Overide the LegStyle if there is no heel value
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        
        If nRightLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtRightAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtRightAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtRightAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtRightAnkle.Text, 4)
        End If

        nFabricClass = fnGetNumber(txtRightAnkle.Text, 8)
        If nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtRightRed
            PR_PutStringAssign "sTapeMMs", txtRightMMs
            PR_PutStringAssign "sStretch", txtRightStr
        End If

        PR_PutStringAssign "sFirstLeg", "Right"
        PR_PutStringAssign "sLeg", "Right"
        PR_PutStringAssign "sSecondLeg", "Left"

        PR_PutStringAssign "sPressure", txtRightTemplate
        itemplate = Val(Left$(txtRightTemplate.Text, 2))
        
        
        PR_PutStringAssign "sTapeLengths", txtRightLengths
        PR_PutStringAssign "sToeStyle", txtRightToeStyle
        
        PR_PutNumberAssign "nFootPleat1", fnDisplaytoInches(Val(txtRightFootPleat1.Text))
        PR_PutNumberAssign "nFootPleat2", fnDisplaytoInches(Val(txtRightFootPleat2.Text))
        PR_PutNumberAssign "nTopLegPleat1", fnDisplaytoInches(Val(txtRightTopLegPleat1.Text))
        PR_PutNumberAssign "nTopLegPleat2", fnDisplaytoInches(Val(txtRightTopLegPleat2.Text))
        
        PR_PutStringAssign "sFootLength", txtRightFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtRightData.Text, 1)

       'Other leg details used to control toe
        nValue = fnGetNumber(txtLeftAnkle.Text, 1)
        If nLeftLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If


    End If

   'First and LastTapes
    nFirstTape = -1
    nLastTape = 30
    For ii = 0 To 29
        nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
       'Set first and last tape (assumes no holes in data)
        If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
        If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
    Next ii
    PR_PutNumberAssign "nFirstTape", nFirstTape
    PR_PutNumberAssign "nLastTape", nLastTape

   'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
   'For a footless style the only valid classes are 0 and 1
    If Footless And nFabricClass = 2 Then nFabricClass = 1
    PR_PutNumberAssign "nFabricClass", nFabricClass
   

   'Draw leg except for BREIF
    If nLegStyle = 2 Then
        PR_PutStringAssign "sPressure", "-"
       'Draw Back body
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1BRF.D;"

    Else
       'Load universal Procedures and Defaults
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1DEF.D;"
       
       'Template type
        Select Case itemplate
            Case 13
                PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH13DS.D;"
            Case 9
                PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH09DS.D;"
            Case Else
                PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHPOW.D;"
        End Select


       'Get Origin
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1ORG.D;"

       'Draw template
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;"

       'Calculate Foot Points (If a foot exists)
        If Not Footless Then PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;"
    
       'Draw Leg
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1DWG.D;"
    End If

   'Close macro file
    Close #fNum
    
End Sub

Private Sub PR_WH_2ndLeg()
   'Procedure to create a Macro to draw the 2nd leg of a
   'Waist Height
    Dim sFile As String, sSide  As String, sBody As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim ii As Integer, Footless As Integer, itemplate As Integer
    Dim nFoldHt As Double, nValue As Double
    Dim nThighCir As Double

   'Establish which is the First Leg
    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    Select Case nLegStyle
        Case 0, 1
            If fnDisplaytoInches(Val(txtLeftThighCir)) - fnDisplaytoInches(Val(txtRightThighCir)) > 1 Then
                sSide = "Left"
                nThighCir = fnDisplaytoInches(Val(txtRightThighCir)) * Val(txtThighRed) / 2
            Else
                sSide = "Right"
                nThighCir = fnDisplaytoInches(Val(txtLeftThighCir)) * Val(txtThighRed) / 2
            End If
        Case 2
            MsgBox "Brief Style Not yet supported", 16, "g_sDialogueID"
            End
        Case 3 To 6
            MsgBox "Can't draw a second leg for this style", 16, g_sDialogueID
            End
    End Select


   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)

    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_LEG2.D;"

   'Put standard patient details
    PR_StandardDetails

   'Body details
    PR_PutStringAssign "sTOSCir", txtTOSCir
    PR_PutNumberAssign "nTOSCir", Val(txtTOSCir.Text)

   'Fold Height (Height to fold of buttocks), Used revised if given
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    sBody = txtBody.Text
    If fnGetNumber(sBody, 4) > 0 And fnGetNumber(sBody, 7) > 0 Then
        nFoldHt = fnDisplaytoInches(fnGetNumber(sBody, 7))
    End If

    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutStringAssign "sCrotchStyle", txtCrotchStyle
    PR_PutStringAssign "sFabric", txtFabric

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    nLeftLegStyle = fnGetNumber(txtLegStyle.Text, 2)
    nRightLegStyle = fnGetNumber(txtLegStyle.Text, 3)
    
    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutNumberAssign "nLeftLegStyle", nLeftLegStyle
    PR_PutNumberAssign "nRightLegStyle", nRightLegStyle
    

   'Leg Details
    If sSide = "Left" Then
       'The Cut-Out is always drawn for the Left side
        PR_PutLine "LeftLeg = %true;"
        PR_PutLine "RightLeg = %false;"
        sLengths = txtLeftLengths.Text  'Use later to establish 1st & Last tapes
       'Overide the LegStyle if there is no heel value
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        If nLeftLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtLeftAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtLeftAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtLeftAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtLeftAnkle.Text, 4)
        End If
        
        nFabricClass = fnGetNumber(txtLeftAnkle.Text, 8)
        If nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtLeftRed
            PR_PutStringAssign "sTapeMMs", txtLeftMMs
            PR_PutStringAssign "sStretch", txtLeftStr
        End If

        PR_PutStringAssign "sFirstLeg", "Right"
        PR_PutStringAssign "sLeg", "Left"
        PR_PutStringAssign "sSecondLeg", "Left"

        PR_PutStringAssign "sPressure", txtLeftTemplate
        itemplate = Val(Left$(txtLeftTemplate.Text, 2))
        
        PR_PutStringAssign "sTapeLengths", txtLeftLengths
        PR_PutStringAssign "sToeStyle", txtLeftToeStyle
        
        PR_PutNumberAssign "nFootPleat1", fnDisplaytoInches(Val(txtLeftFootPleat1.Text))
        PR_PutNumberAssign "nFootPleat2", fnDisplaytoInches(Val(txtLeftFootPleat2.Text))
        PR_PutNumberAssign "nTopLegPleat1", fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text))
        PR_PutNumberAssign "nTopLegPleat2", fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text))
        
        PR_PutStringAssign "sFootLength", txtLeftFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtLeftData.Text, 1)
       
       'Crotch style and data
        If txtCrotchStyle.Text = "Open Crotch" Then
            PR_PutLine "OpenCrotch  = %true;"
            PR_PutLine "ClosedCrotch = %false;"
        Else
            PR_PutLine "OpenCrotch  = %false;"
            PR_PutLine "ClosedCrotch = %true;"
        End If

        PR_PutNumberAssign "nCrotchFrontFactor", fnGetNumber(sBody, 1)
        PR_PutNumberAssign "nOpenFront", fnGetNumber(sBody, 2)
        PR_PutNumberAssign "nOpenBack", fnGetNumber(sBody, 3)

       'Other leg details used to control toe
        nValue = fnGetNumber(txtRightAnkle.Text, 1)
        If nRightLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If
    
    Else
        PR_PutLine "LeftLeg = %false;"
        PR_PutLine "RightLeg = %true;"
        sLengths = txtRightLengths.Text  'Use later to establish 1st & Last tapes
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        If nRightLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtRightAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtRightAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtRightAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtRightAnkle.Text, 4)
        End If

        nFabricClass = fnGetNumber(txtRightAnkle.Text, 8)
        If nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtRightRed
            PR_PutStringAssign "sTapeMMs", txtRightMMs
            PR_PutStringAssign "sStretch", txtRightStr
        End If

        PR_PutStringAssign "sFirstLeg", "Left"
        PR_PutStringAssign "sLeg", "Right"
        PR_PutStringAssign "sSecondLeg", "Right"

        PR_PutStringAssign "sPressure", txtRightTemplate
        itemplate = Val(Left$(txtRightTemplate.Text, 2))
        
        PR_PutStringAssign "sTapeLengths", txtRightLengths
        PR_PutStringAssign "sToeStyle", txtRightToeStyle
        
        PR_PutNumberAssign "nFootPleat1", fnDisplaytoInches(Val(txtRightFootPleat1.Text))
        PR_PutNumberAssign "nFootPleat2", fnDisplaytoInches(Val(txtRightFootPleat2.Text))
        PR_PutNumberAssign "nTopLegPleat1", fnDisplaytoInches(Val(txtRightTopLegPleat1.Text))
        PR_PutNumberAssign "nTopLegPleat2", fnDisplaytoInches(Val(txtRightTopLegPleat2.Text))
        
        PR_PutStringAssign "sFootLength", txtRightFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtRightData.Text, 1)
       
       'Other leg details used to control toe
        nValue = fnGetNumber(txtLeftAnkle.Text, 1)
        If nLeftLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If


    End If

   'First and LastTapes
    nFirstTape = -1
    nLastTape = 30
    For ii = 0 To 29
        nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
       'Set first and last tape (assumes no holes in data)
        If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
        If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
    Next ii
    PR_PutNumberAssign "nFirstTape", nFirstTape
    PR_PutNumberAssign "nLastTape", nLastTape

   'Actual Waist and Figured thigh Circumference
   'Used for body and lateral zips
    PR_PutNumberAssign "nGivenWaistCir", fnDisplaytoInches(Val(txtWaistCir))
    PR_PutNumberAssign "nThighCir", nThighCir + 0.1875

   'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
   'For a footless style the only valid classes are 0 and 1
    If Footless And nFabricClass = 2 Then nFabricClass = 1
    PR_PutNumberAssign "nFabricClass", nFabricClass
   
   'Load universal Procedures and Defaults
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2DEF.D;"
   
   'Get 1st leg profile
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2LEG.D;"

   'Template type
    Select Case itemplate
        Case 13
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH13DS.D;"
        Case 9
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH09DS.D;"
        Case Else
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHPOW.D;"
    End Select

   'Get Origin
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2ORG.D;"

   'Draw template
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;"

   'Draw Body
    If sSide = "Left" Then
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2CUT.D;"
    Else
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2DBD.D;"
    End If
    
   'Calculate Foot Points (If a foot exists)
    If Not Footless Then PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;"
    
   'Draw Leg
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2DWG.D;"

   'Close macro file
    Close #fNum
    
End Sub

Private Sub PR_WH_Brief1stLeg()
   'Procedure to create a Macro to draw the first leg of a
   'Waist Height
    Dim sFile As String, sSide  As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim ii As Integer, Footless As Integer, itemplate As Integer
    Dim nFoldHt As Double, nValue As Double

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    sFile = "C:\JOBST\DRAW.D"

   'For a breif the side is always LEFT
    sSide = "Left"

   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)

    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_LEG1.D;"
   
   'Put standard patient details
    PR_StandardDetails

   'Body details
    PR_PutStringAssign "sTOSCir", txtTOSCir
   
   'Fold Height (Height to fold of buttocks), Used revised if given
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    If fnGetNumber(txtBody.Text, 4) > 0 And fnGetNumber(txtBody.Text, 7) > 0 Then
        nFoldHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 7))
    End If
    
    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutStringAssign "sCrotchStyle", txtCrotchStyle
    PR_PutStringAssign "sFabric", txtFabric

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    
   'Leg Details
    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutLine "LeftLeg = %true;"
    PR_PutLine "RightLeg = %false;"
    PR_PutLine "FootLess = %true;"
    PR_PutStringAssign "sFirstLeg", "Left"
    PR_PutStringAssign "sLeg", "Left"
    PR_PutStringAssign "sSecondLeg", "Right"
    PR_PutNumberAssign "nFabricClass", 0
    PR_PutStringAssign "sPressure", "-"

   'Draw Back body
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1BRF.D;"

   'Close macro file
    Close #fNum

End Sub

Private Sub PR_WH_Brief2ndLeg()
   'Procedure to create a Macro to draw the Second leg of a
   'Waist Height Breif
    Dim sFile As String, sSide  As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim ii As Integer, Footless As Integer, itemplate As Integer
    Dim nFoldHt As Double, nValue As Double

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    sFile = "C:\JOBST\DRAW.D"

   'For a breif the side is always LEFT
    sSide = "Left"

   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)

    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_LEG2.D;"
    
   'Put standard patient details
    PR_StandardDetails

   'Body details
    PR_PutStringAssign "sTOSCir", txtTOSCir
   
   'Fold Height (Height to fold of buttocks), Used revised if given
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    If fnGetNumber(txtBody.Text, 4) > 0 And fnGetNumber(txtBody.Text, 7) > 0 Then
        nFoldHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 7))
    End If
    
    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutStringAssign "sCrotchStyle", txtCrotchStyle
    PR_PutStringAssign "sFabric", txtFabric

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    
   'Leg Details
    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutLine "LeftLeg = %true;"
    PR_PutLine "RightLeg = %false;"
    PR_PutLine "FootLess = %true;"
    PR_PutStringAssign "sLeg", "Left"
    PR_PutStringAssign "sFirstLeg", "Left"
    PR_PutStringAssign "sSecondLeg", "Right"
    PR_PutNumberAssign "nFabricClass", 0
    PR_PutStringAssign "sPressure", "-"

    If txtCrotchStyle.Text = "Open Crotch" Then
        PR_PutLine "OpenCrotch = %true;"
        PR_PutNumberAssign "nBriefOff", 1.625
    Else
        PR_PutLine "OpenCrotch = %false;"
        PR_PutNumberAssign "nBriefOff", 1.25
    End If

   'Fold offset
    If txtCrotchStyle.Text = "Horizontal Fly" Then
        PR_PutNumberAssign "nFoldOff", 0
    Else
        PR_PutNumberAssign "nFoldOff", 0.75
    End If

   'Draw Back body
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2LEG.D;"
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG2BRF.D;"

   'Close macro file
    Close #fNum


End Sub

Private Sub PR_WH_ChapBody(sSide As String)
   'Procedure to draw the chap style BODY
    Dim xyPt As xy, xyPt1 As xy, xyPt2 As xy, xyPt3 As xy, xyStartCL As xy, xyStartOFF As xy, xyStartBotOFF As xy, xyInt As xy
    Dim xyEOSCL As xy, xyChapDatumOFF As xy, xyWaistBandProx As xy, xyWaistBandDist As xy
    Dim xyVelcroOverLapProx  As xy, xyVelcroOverLapDist As xy
    Dim iDistCase%, iProxCase%, xyChapBotDatumOFF As xy
    Dim xyCenTopProx As xy, xyCenTopDist As xy, xyCenTop As xy, xyArcBottomPt  As xy
    Dim xyCenBotProx As xy, xyCenBotDist As xy, xyCenBot As xy, xyTmpDist  As xy, xyTmpProx As xy
    Dim nTopRadius#, nTopRaduisProx#, nTopRaduisDist#, nLength#, aAngle#, nEighthOfWaistCir#
    Dim nBotRadius#, nBotRaduisProx#, nBotRaduisDist#
    Dim TmpBotlegProfile  As Curve
    Dim nA#, nB#, nC#, aTheta#, nCosTheta#
    Dim iStart%, ii%, Sucess%
    Dim nWaistCir#, nTOSCir#, nWaistHt#, nFoldHt#, nTOSHt#, nReduction#
    Dim sLegFile$, fFile%, nn#
    

   'Get Leg curve from file
    If FileLen("C:\JOBST\LEGCURVE.DAT") = 0 Then
        MsgBox "C:\JOBST\LEGCURVE.DAT" & "Not found", 48, g_sDialogueID
        Exit Sub
    End If

    fFile = FreeFile
    Open "C:\JOBST\LEGCURVE.DAT" For Input As fFile

   'Get control points
    Input #fFile, xyOtemplate.X, xyOtemplate.Y

   'Get profile points
    g_TopLegProfile.n = 0
    Do While Not EOF(fFile)
        g_TopLegProfile.n = g_TopLegProfile.n + 1
        Input #fFile, g_TopLegProfile.X(g_TopLegProfile.n), g_TopLegProfile.Y(g_TopLegProfile.n)
        Loop
    Close #fFile


   'Revise xyChapDatum wrt leg and last tape
    xyChapDatum.X = g_TopLegProfile.X(g_TopLegProfile.n) - 0.75
    xyChapDatum.Y = xyOtemplate.Y

   'Get the leg points that straddle x = xyChapDatum.x.
   'This is done to take account of pleats

    If g_TopLegProfile.n - 5 <= 0 Then iStart = 1 Else iStart = g_TopLegProfile.n - 5
   
   'From top
    For nn = g_TopLegProfile.n To iStart Step -1
        If g_TopLegProfile.X(nn) > xyChapDatum.X Then
            xyProfileProximal.X = g_TopLegProfile.X(nn)
            xyProfileProximal.Y = g_TopLegProfile.Y(nn)
        End If
    Next nn
   
   'From bottom
   'NB
   ' 1. We also create the mirrored bottom leg profile
   ' 2. Special case where g_ToplegProfile.X(nn) = xyChapDatum.X
    g_BotLegProfile.n = 0
    For nn = iStart To g_TopLegProfile.n
        g_BotLegProfile.n = g_BotLegProfile.n + 1
        g_BotLegProfile.X(g_BotLegProfile.n) = g_TopLegProfile.X(nn)
        nLength = g_TopLegProfile.Y(nn) - xyChapDatum.Y
        g_BotLegProfile.Y(g_BotLegProfile.n) = xyChapDatum.Y - nLength
        If g_TopLegProfile.X(nn) <= xyChapDatum.X Then
            If g_TopLegProfile.X(nn) = xyChapDatum.X Then
                xyProfileDistal.X = g_TopLegProfile.X(nn) - 0.001
            Else
                xyProfileDistal.X = g_TopLegProfile.X(nn)
            End If
            xyProfileDistal.Y = g_TopLegProfile.Y(nn)
        End If
    Next nn

   'Circumferences
    nTOSCir = fnDisplaytoInches(Val(txtTOSCir.Text))
    nWaistCir = fnDisplaytoInches(Val(txtWaistCir.Text))
    nTOSHt = fnDisplaytoInches(Val(txtTOSHt.Text))
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    nWaistHt = fnDisplaytoInches(Val(txtWaistHt.Text))
    If nTOSCir <> 0 And nTOSCir < nWaistCir Then
        nReduction = Val(txtTOSRed.Text)
        g_nChapCirGiven = nTOSCir
        g_nChapCirFigured = (nTOSCir * nReduction) / 2
        g_nChapLenGiven = nTOSHt - nFoldHt
    Else
        nReduction = Val(txtWaistRed.Text)
        g_nChapCirGiven = nWaistCir
        g_nChapCirFigured = (nWaistCir * nReduction) / 2
        g_nChapLenGiven = nWaistHt - nFoldHt
    End If
    g_nChapLenFigured = (g_nChapLenGiven / 1.2) + 0.75

    g_nLegStyle = fnGetNumber(txtLegStyle.Text, 1)


    'Get intersections on leg profile
    'Note special case for only 3 points in leg profile
     If g_TopLegProfile.n - 5 <= 0 Then iStart = 1 Else iStart = g_TopLegProfile.n - 5

    'PR_MakeXY xyOrigin, 0, xyChapDatum.y
     xyOrigin = xyOtemplate
     PR_MakeXY xyStartCL, xyChapDatum.X - 2.5, xyChapDatum.Y
     PR_MakeXY xyStartOFF, xyStartCL.X, xyChapDatum.Y + 100

     PR_MakeXY xyPt1, g_TopLegProfile.X(iStart), g_TopLegProfile.Y(iStart)
     
     Sucess = False
     If g_TopLegProfile.n = 3 Then iStart = 0 'NB 2 tape leg
     For ii = iStart + 1 To g_TopLegProfile.n
        PR_MakeXY xyPt2, g_TopLegProfile.X(ii), g_TopLegProfile.Y(ii)
        If FN_LinLinInt(xyStartCL, xyStartOFF, xyPt1, xyPt2, xyStartOFF) Then
            Sucess = True
            Exit For
        End If
        xyPt1 = xyPt2
     Next ii
     
    'assume a short 3 tape leg
    'and extend the bottom leg profile
     If Not Sucess Then
        PR_MakeXY xyStartOFF, xyStartCL.X, g_TopLegProfile.Y(1)
        
        For ii = 1 To g_BotLegProfile.n
            TmpBotlegProfile.X(ii + 1) = g_BotLegProfile.X(ii)
            TmpBotlegProfile.Y(ii + 1) = g_BotLegProfile.Y(ii)
        Next ii
        
        nLength = xyStartOFF.Y - xyChapDatum.Y
        TmpBotlegProfile.Y(1) = xyChapDatum.Y - nLength
        TmpBotlegProfile.X(1) = xyStartOFF.X
        TmpBotlegProfile.n = g_BotLegProfile.n + 1

        g_BotLegProfile = TmpBotlegProfile

     End If

    'Get intesection at X = xyChapDatum.X
     Sucess = False
     PR_MakeXY xyChapDatumOFF, xyChapDatum.X, xyChapDatum.Y + 100
     If FN_LinLinInt(xyChapDatum, xyChapDatumOFF, xyProfileDistal, xyProfileProximal, xyChapDatumOFF) Then
        Sucess = True
     End If

    'Give an error message here if we can't find a leg intesection
     If Not Sucess Then
        MsgBox "Can't form upper part of back edge of chap. Unable to find an intesection on the leg profile", 16, g_sDialogueID
        Close #fNum
        End
     End If

    'Calculate key points
     nEighthOfWaistCir = g_nChapCirFigured / 4
     PR_MakeXY xyEOSCL, xyChapDatum.X + g_nChapLenFigured, xyChapDatum.Y
     PR_MakeXY xyVelcroOverLapProx, xyEOSCL.X, xyChapDatum.Y + nEighthOfWaistCir
     PR_MakeXY xyVelcroOverLapDist, xyEOSCL.X - 1, xyChapDatum.Y + nEighthOfWaistCir
          
     PR_MakeXY xyWaistBandProx, xyVelcroOverLapProx.X, xyVelcroOverLapProx.Y - g_nChapCirFigured
     PR_MakeXY xyWaistBandDist, xyVelcroOverLapDist.X, xyVelcroOverLapDist.Y - g_nChapCirFigured


    'Draw CHAP
     fNum = FN_ChapOpen("C:\JOBST\DRAW.D")

     PR_SetLayer "Template" & sSide
     PR_DrawLine xyWaistBandDist, xyWaistBandProx
     PR_DrawLine xyWaistBandProx, xyVelcroOverLapProx

    'Draw velcro OVERLAP
    'If two legs and we are drawing the Left then don't draw the
    'velcro overlap
     If g_nLegStyle = 7 And sSide = "Left" Then
        PR_DrawLine xyVelcroOverLapProx, xyVelcroOverLapDist
     Else
       'Label overlap so it can be deleted by the Zipper programme
        PR_MakeXY xyPt1, xyVelcroOverLapProx.X, xyVelcroOverLapProx.Y + 2.5
        PR_MakeXY xyPt2, xyVelcroOverLapDist.X, xyVelcroOverLapDist.Y + 2.5
        PR_DrawLine xyVelcroOverLapProx, xyPt1
        PR_AddDBValueToLast "Zipper", "VelcroOverlap"
        PR_DrawLine xyPt1, xyPt2
        PR_AddDBValueToLast "Zipper", "VelcroOverlap"
        PR_DrawLine xyVelcroOverLapDist, xyPt2
        PR_AddDBValueToLast "Zipper", "VelcroOverlap"

        PR_SetLayer "Notes"
        PR_DrawLine xyVelcroOverLapProx, xyVelcroOverLapDist
        PR_AddDBValueToLast "Zipper", "VelcroOverlap"
        PR_CalcMidPoint xyPt1, xyVelcroOverLapDist, xyPt3
        PR_InsertSymbol "TextAsSymbol", xyPt3, 1, 90
        PR_AddDBValueToLast "Data", "2-1/2\"" VELCRO"
        PR_AddDBValueToLast "Zipper", "VelcroOverlap"
        PR_SetLayer "Template" & sSide
     
     End If
     
    'Add One waist band for Chap Both Legs
     If g_nLegStyle = 7 Then
        PR_SetLayer "Notes"
        PR_SetTextData HORIZ_CENTER, VERT_CENTER, 0.125, CURRENT, CURRENT
        g_nCurrTextAngle = 90
        PR_CalcMidPoint xyEOSCL, xyVelcroOverLapProx, xyPt3
        xyPt3.X = xyPt3.X - 0.5
        PR_DrawText "ONE WAIST BAND", xyPt3, 0.125
        g_nCurrTextAngle = 0
        PR_SetLayer "Template" & sSide
     End If


    'Get the first aproximation for the Top and Bottom arcs
    'Note we work on the +ve side of the line Y = xyChapDatum.Y
    'and then mirror the results for the bottom arc

    'Top arc/s
     nTopRadius = (xyVelcroOverLapDist.X - xyChapDatum.X) / 2
     PR_MakeXY xyCenTop, xyChapDatum.X + nTopRadius, xyChapDatum.Y + nTopRadius

    'Adjust to fit if the center of the arc is above the velcro line
    'or the end of the profile
     PR_MakeXY xyArcBottomPt, xyCenTop.X, xyChapDatum.Y
     If xyCenTop.Y > xyVelcroOverLapDist.Y Then
       'Revise center of arc and Raduis
        PR_MakeXY xyArcBottomPt, xyCenTop.X, xyChapDatum.Y
        nLength = FN_CalcLength(xyArcBottomPt, xyVelcroOverLapDist) / 2
        aAngle = 90 - FN_CalcAngle(xyArcBottomPt, xyVelcroOverLapDist)
        nTopRaduisProx = nLength / Cos(aAngle * (PI / 180))
        PR_MakeXY xyCenTopProx, xyArcBottomPt.X, xyArcBottomPt.Y + nTopRaduisProx
        PR_DrawArc xyCenTopProx, xyArcBottomPt, xyVelcroOverLapDist
     Else
        PR_MakeXY xyPt1, xyCenTop.X + nTopRadius, xyCenTop.Y
        PR_DrawLine xyVelcroOverLapDist, xyPt1
        PR_DrawArc xyCenTop, xyArcBottomPt, xyPt1
     End If
     
     If xyCenTop.Y > xyChapDatumOFF.Y Then
       'Revise center of arc and Raduis
        nLength = FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
        aAngle = FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF) - 90
        nTopRaduisDist = nLength / Cos(aAngle * (PI / 180))
        PR_MakeXY xyCenTopDist, xyArcBottomPt.X, xyArcBottomPt.Y + nTopRaduisDist
        PR_DrawArc xyCenTopDist, xyChapDatumOFF, xyArcBottomPt
     Else
        PR_MakeXY xyPt1, xyCenTop.X - nTopRadius, xyCenTop.Y
        PR_DrawLine xyChapDatumOFF, xyPt1
        PR_DrawArc xyCenTop, xyPt1, xyArcBottomPt
     End If


    'Bottom arc/s
    'Calculate for +ve Y and then mirror results
     PR_MakeXY xyArcBottomPt, xyCenTop.X, xyWaistBandDist.Y + nEighthOfWaistCir
    

    'Mirror our two control points in the line Y = xyChapDatum.Y
     xyArcBottomPt.Y = xyChapDatum.Y + Abs(xyArcBottomPt.Y - xyChapDatum.Y)
     xyWaistBandDist.Y = xyChapDatum.Y + Abs(xyWaistBandDist.Y - xyChapDatum.Y)
    
     nBotRadius = nTopRadius
     PR_MakeXY xyCenBot, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRadius

     If xyCenBot.Y > xyWaistBandDist.Y Then
        nLength = FN_CalcLength(xyArcBottomPt, xyWaistBandDist) / 2
        aAngle = FN_CalcAngle(xyArcBottomPt, xyWaistBandDist) - 90
        nBotRaduisProx = nLength / Cos(aAngle * (PI / 180))
        PR_MakeXY xyCenBotProx, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRaduisProx
        xyTmpProx = xyWaistBandDist
        iProxCase = 1
     Else
        xyCenBotProx = xyCenBot
        nBotRaduisProx = nBotRadius
        PR_MakeXY xyTmpProx, xyCenBot.X + nBotRadius, xyCenBot.Y
        iProxCase = 2
     End If

     If xyCenBot.Y < xyChapDatumOFF.Y Then
        PR_MakeXY xyTmpDist, xyCenBot.X - nBotRadius, xyCenBot.Y
        xyCenBotDist = xyCenBot
        iDistCase = 1

     ElseIf xyArcBottomPt.Y < xyChapDatumOFF.Y Then
       'Revise center of arc and Raduis
        nLength = FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
        aAngle = 90 - FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF)
        nBotRaduisDist = nLength / Cos(aAngle * (PI / 180))
        PR_MakeXY xyCenBotDist, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRaduisDist
        iDistCase = 2
     ElseIf xyArcBottomPt.Y = xyChapDatumOFF.Y Then
        iDistCase = 3
     Else
        nA = FN_CalcLength(xyCenBotProx, xyChapDatumOFF)
        nC = nBotRaduisProx
        nB = Sqr(Abs(nA ^ 2 - nC ^ 2))
       'Using the cosine rule to get the angle
        nCosTheta = (nC ^ 2 - (nA ^ 2 + nB ^ 2)) / -(2 * nA * nB)
        aTheta = Arccos(nCosTheta) * (180 / PI)
        aAngle = FN_CalcAngle(xyChapDatumOFF, xyCenBotProx) - aTheta
        PR_CalcPolar xyChapDatumOFF, aAngle, nB, xyArcBottomPt
        iDistCase = 4
     End If

    'Mirror our points in the line Y = xyChapDatum.Y
     xyArcBottomPt.Y = xyChapDatum.Y - Abs(xyArcBottomPt.Y - xyChapDatum.Y)
     xyWaistBandDist.Y = xyChapDatum.Y - Abs(xyWaistBandDist.Y - xyChapDatum.Y)
     xyCenBotProx.Y = xyChapDatum.Y - Abs(xyCenBotProx.Y - xyChapDatum.Y)
     xyCenBotDist.Y = xyChapDatum.Y - Abs(xyCenBotDist.Y - xyChapDatum.Y)
     xyChapBotDatumOFF = xyChapDatumOFF
     xyChapBotDatumOFF.Y = xyChapDatum.Y - Abs(xyChapBotDatumOFF.Y - xyChapDatum.Y)
     xyStartBotOFF = xyStartOFF
     xyStartBotOFF.Y = xyChapDatum.Y - Abs(xyStartBotOFF.Y - xyChapDatum.Y)
     xyTmpProx.Y = xyChapDatum.Y - Abs(xyTmpProx.Y - xyChapDatum.Y)
     xyTmpDist.Y = xyChapDatum.Y - Abs(xyTmpDist.Y - xyChapDatum.Y)

    'Draw the arcs and lines based on the draw case
     If iProxCase = 2 Then PR_DrawLine xyWaistBandDist, xyTmpProx
     PR_DrawArc xyCenBotProx, xyArcBottomPt, xyTmpProx

     Select Case iDistCase
        Case 1
            PR_DrawLine xyChapBotDatumOFF, xyTmpDist
            PR_DrawArc xyCenBotDist, xyTmpDist, xyArcBottomPt
        Case 2
            PR_DrawArc xyCenBotDist, xyChapBotDatumOFF, xyArcBottomPt
        Case 3, 4
            PR_DrawLine xyArcBottomPt, xyChapBotDatumOFF
     End Select

    'Bottom leg profile
     PR_DrawFitted g_BotLegProfile
     PR_DrawLine xyStartCL, xyEOSCL

    'Draw markers for use with zippers
     PR_SetLayer "Construct"
     PR_DrawMarker xyChapDatum, "Fold"

     PR_DrawMarker xyEOSCL, "EOS"
    
    'Draw Closing lines
     PR_SetLayer "Template" & sSide
     PR_DrawLine xyStartCL, xyStartBotOFF
     PR_DrawLine xyOrigin, xyStartCL

    'Tram lines
     PR_SetLayer "Notes"

     xyPt1 = xyOrigin
     xyPt1.Y = xyOrigin.Y + 0.1875
     xyPt2 = xyStartCL
     xyPt2.Y = xyStartCL.Y + 0.1875
     PR_DrawLine xyPt1, xyPt2
     
     xyPt1.Y = xyPt2.Y + 0.5
     xyPt2.Y = xyPt2.Y + 0.5
     PR_DrawLine xyPt1, xyPt2


'     PR_DrawMarker xyCenTop
'     PR_DrawText "xyCenTop", xyCenTop, .1
'     PR_DrawMarker xyStartOFF
'     PR_DrawText "xyStartOFF", xyStartOFF, .1
'     PR_DrawMarker xyChapDatumOFF
'     PR_DrawText "xyChapDatumOFF", xyChapDatumOFF, .1
'     PR_DrawMarker xyChapDatum
'     PR_DrawText "xyChapDatum", xyChapDatum, .1
'     PR_DrawMarker xyEOSCL
'     PR_DrawText "xyEOSCL", xyEOSCL, .1
'     PR_DrawMarker xyProfileDistal
'     PR_DrawText "xyProfileDistal", xyProfileDistal, .1
'     PR_DrawMarker xyProfileProximal
'     PR_DrawText "xyProfileProximal", xyProfileProximal, .1
     
     Close #fNum

End Sub

Private Sub PR_WH_ChapLeg(sDrawType$)
   'Draw a chap style
    Dim sFile As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer
    Dim ii As Integer, Footless As Integer, itemplate As Integer
    Dim nFoldHt As Double, nValue As Double

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    g_nLegStyle = nLegStyle
    nLeftLegStyle = fnGetNumber(txtLegStyle.Text, 2)
    nRightLegStyle = fnGetNumber(txtLegStyle.Text, 3)
    
    sFile = "C:\JOBST\DRAW.D"

   'Establish which is the First Leg
    Select Case nLegStyle
        Case 5
            g_sSide = "Left"
        Case 6
            g_sSide = "Right"
        Case 7
            If sDrawType = "WH2LG" Then
                g_sSide = "Right"
            Else
                g_sSide = "Left"
            End If
    End Select

   'Having decided on leg do a swift check to see if there is enough
   'data to draw the leg
    If g_sSide = "Left" And ((txtLeftAnkle = "" And txtLeftLengths = "") Or txtUidLeftLeg = "") Then
        MsgBox "No data for LEFT Leg", 16, g_sDialogueID
        End
    End If

    If g_sSide = "Left" And txtLeftTemplate = "" Then
        MsgBox "Figure LEFT Leg before drawing", 16, g_sDialogueID
        End
    End If

    If g_sSide = "Right" And ((txtRightAnkle = "" And txtRightLengths = "") Or txtUidRightLeg = "") Then
        MsgBox "No data for RIGHT Leg", 16, g_sDialogueID
        End
    End If

    If g_sSide = "Right" And txtRightTemplate = "" Then
        MsgBox "Figure RIGHT Leg before drawing", 16, g_sDialogueID
        End
    End If



'    If g_sSide = "Left" And fnGetNumber(txtLeftAnkle.Text, 1) = -1 And nLeftLegStyle <> 1 Then
'        MsgBox "A Full leg has been chosen in the Body, but the data in the leg is footless.  ", 16, "First Leg Drawing"
'        End
'    End If


   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, g_sSide)

    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_LEGC.D;"
   
   'Put standard patient details
    PR_StandardDetails

   'Get Lengths and Circumferences
    Dim UseRevisedHts%
   'Heights
    If fnGetNumber(txtBody.Text, 4) > 0 Then UseRevisedHts = True Else UseRevisedHts = False
    If UseRevisedHts And fnGetNumber(txtBody.Text, 7) > 0 Then
        PR_PutStringAssign "sFoldHt", fnGetNumber(txtBody.Text, 7)
    Else
        PR_PutStringAssign "sFoldHt", txtFoldHt
    End If
    
    If UseRevisedHts And fnGetNumber(txtBody.Text, 6) > 0 Then
        PR_PutStringAssign "sWaistHt", fnGetNumber(txtBody.Text, 6)
    Else
        PR_PutStringAssign "sWaistHt", txtWaistHt
    End If

    If UseRevisedHts And fnGetNumber(txtBody.Text, 5) > 0 Then
        PR_PutStringAssign "sTOSHt", fnGetNumber(txtBody.Text, 5)
    Else
        PR_PutStringAssign "sTOSHt", txtTOSHt
    End If


   'Body details
    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutStringAssign "sWaistCir", txtWaistCir
    PR_PutStringAssign "sTOSCir", txtTOSCir
    PR_PutStringAssign "sTOSRed", txtTOSRed
    PR_PutStringAssign "sWaistRed", txtWaistRed
    PR_PutStringAssign "sFabric", txtFabric
    PR_PutStringAssign "sLegStyle", txtLegStyle

    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutNumberAssign "nLeftLegStyle", nLeftLegStyle
    PR_PutNumberAssign "nRightLegStyle", nRightLegStyle
    
   'Leg Details
    If g_sSide = "Left" Then
        PR_PutLine "LeftLeg = %true;"
        PR_PutLine "RightLeg = %false;"
        sLengths = txtLeftLengths.Text  'Use later to establish 1st & Last tapes
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        If nLeftLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtLeftAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtLeftAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtLeftAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtLeftAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtLeftAnkle.Text, 4)
        End If
        
        g_nFabricClass = fnGetNumber(txtLeftAnkle.Text, 8)
        If g_nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtLeftRed
            PR_PutStringAssign "sTapeMMs", txtLeftMMs
            PR_PutStringAssign "sStretch", txtLeftStr
        End If

        PR_PutStringAssign "sFirstLeg", "Left"
        PR_PutStringAssign "sLeg", "Left"
        PR_PutStringAssign "sSecondLeg", "Right"

        g_sPressure = txtLeftTemplate.Text
        PR_PutStringAssign "sPressure", txtLeftTemplate

        itemplate = Val(Left$(txtLeftTemplate.Text, 2))
        
        PR_PutStringAssign "sTapeLengths", txtLeftLengths
        g_sTapeLength = txtLeftLengths.Text
        
        PR_PutStringAssign "sToeStyle", txtLeftToeStyle
        
        g_nFootPleat1 = fnDisplaytoInches(Val(txtLeftFootPleat1.Text))
        PR_PutNumberAssign "nFootPleat1", g_nFootPleat1
        
        g_nFootPleat2 = fnDisplaytoInches(Val(txtLeftFootPleat2.Text))
        PR_PutNumberAssign "nFootPleat2", g_nFootPleat2

        g_nTopLegPleat1 = fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text))
        PR_PutNumberAssign "nTopLegPleat1", g_nTopLegPleat1
        
        g_nTopLegPleat2 = fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text))
        PR_PutNumberAssign "nTopLegPleat2", g_nTopLegPleat2
        
        PR_PutStringAssign "sFootLength", txtLeftFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtLeftData.Text, 1)

       'Other leg details used to control toe
        nValue = fnGetNumber(txtRightAnkle.Text, 1)
        If nRightLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If
                                                      
    Else
        PR_PutLine "LeftLeg = %false;"
        PR_PutLine "RightLeg = %true;"
        sLengths = txtRightLengths.Text  'Use later to establish 1st & Last tapes
        nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
        If nRightLegStyle <> 0 Or nValue = 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
           'Ankle tape values
            PR_PutNumberAssign "nAnkleTape", fnGetNumber(txtRightAnkle.Text, 1)
            PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(txtRightAnkle.Text, 1))
            PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 2))
            PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(txtRightAnkle.Text, 3))
            PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(txtRightAnkle.Text, 4))
            PR_PutNumberAssign "nReductionAnkle", fnGetNumber(txtRightAnkle.Text, 4)
        End If

        g_nFabricClass = fnGetNumber(txtRightAnkle.Text, 8)
        If g_nFabricClass = 2 Then
            PR_PutStringAssign "sReduction", txtRightRed
            PR_PutStringAssign "sTapeMMs", txtRightMMs
            PR_PutStringAssign "sStretch", txtRightStr
        End If

        PR_PutStringAssign "sFirstLeg", "Right"
        PR_PutStringAssign "sLeg", "Right"
        PR_PutStringAssign "sSecondLeg", "Left"

        PR_PutStringAssign "sPressure", txtRightTemplate
        g_sPressure = txtRightTemplate
        itemplate = Val(Left$(txtRightTemplate.Text, 2))
        
        PR_PutStringAssign "sTapeLengths", txtRightLengths
        g_sTapeLength = txtRightLengths.Text
        PR_PutStringAssign "sToeStyle", txtRightToeStyle
        
        g_nFootPleat1 = fnDisplaytoInches(Val(txtRightFootPleat1.Text))
        g_nFootPleat2 = fnDisplaytoInches(Val(txtRightFootPleat2.Text))
        g_nTopLegPleat1 = fnDisplaytoInches(Val(txtRightTopLegPleat1.Text))
        g_nTopLegPleat2 = fnDisplaytoInches(Val(txtRightTopLegPleat2.Text))
        PR_PutNumberAssign "nFootPleat1", g_nFootPleat1
        PR_PutNumberAssign "nFootPleat2", g_nFootPleat2
        PR_PutNumberAssign "nTopLegPleat1", g_nTopLegPleat1
        PR_PutNumberAssign "nTopLegPleat2", g_nTopLegPleat2
        
        PR_PutStringAssign "sFootLength", txtRightFootLength
        PR_PutNumberAssign "nElastic", fnGetNumber(txtRightData.Text, 1)

       'Other leg details used to control toe
        nValue = fnGetNumber(txtLeftAnkle.Text, 1)
        If nLeftLegStyle <> 0 Or nValue < 0 Then
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", 0
        Else
           'Establish - Minus Ankle tape value
            If nValue = 8 Then
               'Ankle at +3
                nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
            Else
               'Ankle at +1-1/2
                nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
            End If
            PR_PutNumberAssign "nOtherAnkleMTapeLen ", fnDisplaytoInches(nValue)
        End If
    End If

    g_Footless = Footless

   'First and LastTapes
    g_nFirstTape = -1
    g_nLastTape = 30
    For ii = 0 To 29
        nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
       'Set first and last tape (assumes no holes in data)
        If g_nFirstTape < 0 And nValue > 0 Then g_nFirstTape = ii + 1
        If g_nLastTape = 30 And g_nFirstTape > 0 And nValue = 0 Then g_nLastTape = ii
    Next ii
    PR_PutNumberAssign "nFirstTape", g_nFirstTape
    PR_PutNumberAssign "nLastTape", g_nLastTape

   'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
   'For a footless style the only valid classes are 0 and 1
    If Footless And g_nFabricClass = 2 Then g_nFabricClass = 1
    PR_PutNumberAssign "nFabricClass", g_nFabricClass

    

   'Draw leg
   'Load universal Procedures and Defaults
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLG1DEF.D;"
    
   'Get Origin
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLGCORG.D;"

   'Draw template
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;"

   'Calculate Foot Points (If a foot exists)
    If Not Footless Then PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;"

   'Draw Leg
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHLGCDWG.D;"
   
   'Close macro file
    Close #fNum

    
End Sub

Private Sub PR_WH_CutOut()
   'Procedure to draw a cutout for a waist height
    Dim sFile As String, sSide  As String
    Dim sLengths As String
    Dim nLegStyle As Integer, nLeftLegStyle As Integer, nRightLegStyle As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim ii As Integer, Footless As Integer, nUseRevisedHts As Integer, nAge  As Integer
    Dim nFoldHt As Double, nTOSHt As Double, nWaistHt As Double, nValue As Double
    Dim nCuttingMarkOff As Double, nOpenOff As Double, nHeelLength As Double
    Dim nBriefOff As Double

   'Establish Leg Style
   'and Open Macro file (fNum is declared as Global)
    nAge = Val(txtAge.Text)

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    sFile = "C:\JOBST\DRAW.D"

    Select Case nLegStyle
        Case 0, 1
            If fnDisplaytoInches(Val(txtLeftThighCir)) - fnDisplaytoInches(Val(txtRightThighCir)) > 1 Then
                sSide = "Right"
            Else
                sSide = "Left"
            End If
            fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_CUT.D;"
            PR_PutNumberAssign "nRightThighCir", fnDisplaytoInches(Val(txtRightThighCir.Text))
            PR_PutNumberAssign "nLeftThighCir", fnDisplaytoInches(Val(txtLeftThighCir.Text))
            
            If txtCrotchStyle.Text = "Open Crotch" Then
                 PR_PutLine "OpenCrotch = %true;"
            Else
                 PR_PutLine "OpenCrotch = %false;"
            End If

        Case 2
            fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
            sSide = "Left"
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_CUT.D;"
            PR_PutNumberAssign "nRightThighCir", fnDisplaytoInches(Val(txtRightThighCir.Text))
            PR_PutNumberAssign "nLeftThighCir", fnDisplaytoInches(Val(txtLeftThighCir.Text))
            
            If txtCrotchStyle.Text = "Open Crotch" Then
                 PR_PutLine "OpenCrotch = %true;"
            Else
                 PR_PutLine "OpenCrotch = %false;"
            End If
        
        Case 3
            sSide = "Left"
            fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_W4OC.D;"
            
            PR_PutLine "OpenCrotch = %true;"
            
            PR_PutNumberAssign "nLeftThighCir", fnDisplaytoInches(Val(txtLeftThighCir.Text))
            PR_PutNumberAssign "nRightThighCir", fnDisplaytoInches(Val(txtLeftThighCir.Text))
        
        Case 4
            sSide = "Right"
            fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH_W4OC.D;"
            
            PR_PutLine "OpenCrotch = %true;"

            PR_PutNumberAssign "nRightThighCir", fnDisplaytoInches(Val(txtRightThighCir.Text))
            PR_PutNumberAssign "nLeftThighCir", fnDisplaytoInches(Val(txtRightThighCir.Text))
    
    End Select

   'Put standard patient details
    PR_StandardDetails
    
   'Body details
   'Circumferences
    PR_PutNumberAssign "nTOSCir", fnDisplaytoInches(Val(txtTOSCir.Text))
    PR_PutNumberAssign "nLargestCir", fnDisplaytoInches(Val(txtLargestCir.Text))
    PR_PutNumberAssign "nMidPointCir", fnDisplaytoInches(Val(txtMidPointCir.Text))
    PR_PutNumberAssign "nWaistCir", fnDisplaytoInches(Val(txtWaistCir.Text))

   'Heights
    nUseRevisedHts = fnGetNumber(txtBody.Text, 4)
    nFoldHt = fnDisplaytoInches(Val(txtFoldHt.Text))
    If nUseRevisedHts > 0 And fnGetNumber(txtBody.Text, 7) > 0 Then
        nFoldHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 7))
    End If
    
    nWaistHt = fnDisplaytoInches(Val(txtWaistHt.Text))
    If nUseRevisedHts > 0 And fnGetNumber(txtBody.Text, 6) > 0 Then
        nWaistHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 6))
    End If

    nTOSHt = fnDisplaytoInches(Val(txtTOSHt.Text))
    If nUseRevisedHts > 0 And fnGetNumber(txtBody.Text, 5) > 0 Then
        nTOSHt = fnDisplaytoInches(fnGetNumber(txtBody.Text, 5))
    End If
    
    PR_PutNumberAssign "nFoldHt", nFoldHt
    PR_PutNumberAssign "nTOSHt", nTOSHt
    PR_PutNumberAssign "nWaistHt", nWaistHt

   'Reductions
    PR_PutNumberAssign "nTOSGivenRed", Val(txtTOSRed.Text)
    PR_PutNumberAssign "nWaistGivenRed", Val(txtWaistRed.Text)
    PR_PutNumberAssign "nThighGivenRed", Val(txtThighRed.Text)
    PR_PutNumberAssign "nLargestGivenRed", Val(txtLargestRed.Text)
    PR_PutNumberAssign "nMidPointGivenRed", Val(txtMidPointRed.Text)

   'Open Crotch factors
    PR_PutNumberAssign "nCrotchFrontFactor", fnGetNumber(txtBody.Text, 1)
    PR_PutNumberAssign "nOpenFront", fnGetNumber(txtBody.Text, 2)
    PR_PutNumberAssign "nOpenBack", fnGetNumber(txtBody.Text, 3)
 

    PR_PutStringAssign "sCrotchStyle", txtCrotchStyle

    nLegStyle = fnGetNumber(txtLegStyle.Text, 1)
    nLeftLegStyle = fnGetNumber(txtLegStyle.Text, 2)
    nRightLegStyle = fnGetNumber(txtLegStyle.Text, 3)
    
    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutNumberAssign "nLeftLegStyle", nLeftLegStyle
    PR_PutNumberAssign "nRightLegStyle", nRightLegStyle
    

   'Leg Details
    If sSide = "Left" Then
        PR_PutLine "LeftLeg = %true;"
        PR_PutLine "RightLeg = %false;"
        sLengths = txtLeftLengths.Text  'Use later to establish 1st & Last tapes
        If nLeftLegStyle <> 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
        End If
        PR_PutStringAssign "sLeg", "Left"
    Else
        PR_PutLine "LeftLeg = %false;"
        PR_PutLine "RightLeg = %true;"
        sLengths = txtRightLengths.Text  'Use later to establish 1st & Last tapes
        If nRightLegStyle <> 0 Then
            PR_PutLine "FootLess = %true;"
            Footless = True
        Else
            PR_PutLine "FootLess = %false;"
        End If
        PR_PutStringAssign "sLeg", "Right"
    End If
    

   'Open crotch factors
    If txtCrotchStyle.Text = "Open Crotch" Then
        nHeelLength = max(fnGetNumber(txtRightAnkle.Text, 6), fnGetNumber(txtLeftAnkle.Text, 6))
        If txtSex = "Male" Then
            nOpenOff = 0.75
           'Check for footless style or a footless w40c
            If nLegStyle = 1 Or nLegStyle = 2 Or (nLegStyle = 3 And Footless = True) Or (nLegStyle = 4 And Footless = True) Then
                If nAge <= 10 Then
                    nOpenOff = 0.375
                End If
            Else
                If nHeelLength > 0 And nHeelLength < 9 Then
                    nOpenOff = 0.375
                End If
            End If
        Else
            nOpenOff = 0.375
        End If
        PR_PutNumberAssign "nOpenOff", nOpenOff
    End If

   'Brief offset
    If nLegStyle = 2 Then    'Brief
        nBriefOff = 1.25
        If txtCrotchStyle.Text = "Open Crotch" And nOpenOff = 0.375 Then nBriefOff = 1.625
        If txtCrotchStyle.Text = "Open Crotch" And nOpenOff = 0.75 Then nBriefOff = 2#
        If txtCrotchStyle.Text = "Horizontal Fly" Then nBriefOff = 1.625
        PR_PutNumberAssign "nBriefOff", nBriefOff
    End If

   'Cutting marks for W4OC
    If nLegStyle = 3 Or nLegStyle = 4 Then      'W4OC-Left and W4OC-Right
        nCuttingMarkOff = 0.75
        If Footless = True Then 'ie panty
            If nAge <= 10 Then
                nCuttingMarkOff = 0.375
            End If
        Else
            If nHeelLength > 0 And nHeelLength < 9 Then
                nCuttingMarkOff = 0.375
            End If
        End If
        PR_PutNumberAssign "nCuttingMarkOff", nCuttingMarkOff
    End If


   'First and LastTapes
    nFirstTape = -1
    nLastTape = 30
    For ii = 0 To 29
        nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
       'Set first and last tape (assumes no holes in data)
        If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
        If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
    Next ii
    PR_PutNumberAssign "nFirstTape", nFirstTape
    PR_PutNumberAssign "nLastTape", nLastTape

    
   'Load defaults and Procedures
    PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHCUTDEF.D;"

   'Draw cutout
    Select Case nLegStyle
        Case 0, 1
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHCUTDBD.D;"
        Case 2
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHCUTBRF.D;"
        Case 3, 4
            PR_PutLine "@" & g_sPathJOBST & "\WAIST\WH4OCDBD.D;"
            
    End Select

   'Close macro file
    Close #fNum

End Sub

Private Sub Timer1_Timer()
   'The timer is enabled in Load.
   'A Timer event happens after 10 seconds have passed.
   'If this event happens we assume that there is problem and we kill
   'this instance
   '
   End
End Sub


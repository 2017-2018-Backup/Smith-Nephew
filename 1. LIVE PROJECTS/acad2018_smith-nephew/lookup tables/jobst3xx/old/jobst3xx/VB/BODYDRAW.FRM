VERSION 5.00
Begin VB.Form bodydia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Body Draw - Dialogue"
   ClientHeight    =   8220
   ClientLeft      =   75
   ClientTop       =   930
   ClientWidth     =   12345
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
   LinkTopic       =   "bodydia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   12345
   Begin VB.TextBox txtRightLegPleats 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   99
      Text            =   "txtRightLegPleats"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtLeftLegPleats 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   98
      Text            =   "txtLeftLegPleats"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtLeftUidLeg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   16
      Text            =   "txtLeftUidLeg"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtRightUidLeg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   17
      Text            =   "txtRightUidLeg"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtLeftKneePosition 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   21
      Text            =   "txtLeftKneePosition"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtRightKneePosition 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   23
      Text            =   "txtRightKneePosition"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtLeftLengths 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   24
      Text            =   "txtLeftLengths"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtRightLengths 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      TabIndex        =   25
      Text            =   "txtRightLengths"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   12
      Left            =   9840
      TabIndex        =   36
      Text            =   "Thigh Red."
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   11
      Left            =   9840
      TabIndex        =   38
      Text            =   "Lg. Butt. Red."
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCrotchStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11880
      TabIndex        =   40
      Text            =   "txtCrotchStyle"
      Top             =   690
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11880
      TabIndex        =   41
      Text            =   "txtLegStyle"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9840
      Top             =   6240
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   9840
      TabIndex        =   43
      Text            =   "Fabric"
      Top             =   5220
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   9840
      TabIndex        =   44
      Text            =   "Closure"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   9840
      TabIndex        =   46
      Text            =   "Back Neck"
      Top             =   4620
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   9840
      TabIndex        =   47
      Text            =   "Front Neck"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   9840
      TabIndex        =   51
      Text            =   "Right Axilla"
      Top             =   4020
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   9840
      TabIndex        =   53
      Text            =   "Left Axilla"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   9840
      TabIndex        =   57
      Text            =   "Right Bracup"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   9840
      TabIndex        =   59
      Text            =   "Left Bra Cup"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   9840
      TabIndex        =   63
      Text            =   "Nipple Reduction"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   9840
      TabIndex        =   64
      Text            =   "Chest reduction"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   9840
      TabIndex        =   71
      Text            =   "Waist Reduction"
      Top             =   2220
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidVB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9840
      TabIndex        =   72
      Text            =   "txtUidVB"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9840
      TabIndex        =   73
      Text            =   "txtWorkOrder"
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9840
      TabIndex        =   75
      Text            =   "txtUidMPD"
      Top             =   675
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Design Choices"
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   6060
      TabIndex        =   96
      Top             =   2115
      Width           =   3435
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Left Leg"
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   420
         TabIndex        =   76
         Top             =   4095
         Width           =   1035
         Begin VB.OptionButton optLeftLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Panty"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton optLeftLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Brief"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   78
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.ComboBox cboLegStyle 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3735
         Width           =   1995
      End
      Begin VB.ComboBox cboCrotchStyle 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   5295
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Right Leg"
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   2040
         TabIndex        =   79
         Top             =   4095
         Width           =   1035
         Begin VB.OptionButton optRightLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Panty"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton optRightLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Brief"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.ComboBox cboFabric 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   180
         TabIndex        =   33
         Text            =   "cboFabric"
         Top             =   3180
         Width           =   3075
      End
      Begin VB.ComboBox cboClosure 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   32
         Text            =   "cboClosure"
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtBackNeck 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2190
         Width           =   555
      End
      Begin VB.ComboBox cboBackNeck 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   30
         Text            =   "cboBackNeck"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtFrontNeck 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1410
         Width           =   555
      End
      Begin VB.ComboBox cboFrontNeck 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   28
         Text            =   "cboFrontNeck"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.ComboBox cboRightAxilla 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   27
         Text            =   "cboRightAxilla"
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cboLeftAxilla 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   26
         Text            =   "cboLeftAxilla"
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label lblBackNeck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2340
         TabIndex        =   74
         Top             =   2220
         Width           =   600
      End
      Begin VB.Label lblFrontNeck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2340
         TabIndex        =   77
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   26
         Left            =   1980
         TabIndex        =   97
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame frmBra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Measurements for Bra"
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      TabIndex        =   62
      Top             =   4500
      Width           =   5835
      Begin VB.CommandButton cmdCalculateBraDisks 
         Appearance      =   0  'Flat
         Caption         =   "Calculate Disks"
         Height          =   375
         Left            =   2100
         TabIndex        =   86
         Top             =   2430
         Width           =   1455
      End
      Begin VB.TextBox txtRightDisk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4680
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cboRightCup 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1995
         Width           =   1035
      End
      Begin VB.TextBox txtLeftDisk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cboLeftCup 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1995
         Width           =   1035
      End
      Begin VB.ComboBox cboRed 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   4920
         TabIndex        =   37
         Text            =   "cboRed"
         Top             =   690
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   14
         Top             =   990
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   13
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   12
         Text            =   "9"
         Top             =   330
         Width           =   555
      End
      Begin VB.Label LabBra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   3300
         TabIndex        =   88
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   11
         Left            =   4260
         TabIndex        =   87
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label LabBra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "B"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   3300
         TabIndex        =   85
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   10
         Left            =   4260
         TabIndex        =   84
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LabBra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3300
         TabIndex        =   82
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   9
         Left            =   4260
         TabIndex        =   81
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Measurments for Body"
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   120
      TabIndex        =   61
      Top             =   45
      Width           =   5835
      Begin VB.ComboBox cboRed 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   4920
         TabIndex        =   67
         Text            =   "cboRed"
         Top             =   3780
         Width           =   850
      End
      Begin VB.ComboBox cboRed 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   4920
         TabIndex        =   68
         Text            =   "cboRed"
         Top             =   3293
         Width           =   850
      End
      Begin VB.ComboBox cboRed 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   4920
         TabIndex        =   69
         Text            =   "cboRed"
         Top             =   1973
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2625
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   8
         Top             =   2955
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   14
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   9
         Top             =   3285
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   15
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   10
         Top             =   3615
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   16
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   11
         Top             =   3945
         Width           =   555
      End
      Begin VB.ComboBox cboRed 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   4920
         TabIndex        =   70
         Text            =   "cboRed"
         Top             =   2303
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   6
         Top             =   2295
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1965
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1635
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1305
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   975
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   1
         Top             =   645
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "0"
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   12
         Left            =   4260
         TabIndex        =   89
         Top             =   2685
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   13
         Left            =   4260
         TabIndex        =   90
         Top             =   3015
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   14
         Left            =   4260
         TabIndex        =   91
         Top             =   3345
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   15
         Left            =   4260
         TabIndex        =   92
         Top             =   3675
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   16
         Left            =   4260
         TabIndex        =   93
         Top             =   4005
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   6
         Left            =   4260
         TabIndex        =   65
         Top             =   2355
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   5
         Left            =   4260
         TabIndex        =   60
         Top             =   2025
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   4
         Left            =   4260
         TabIndex        =   55
         Top             =   1695
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   3
         Left            =   4260
         TabIndex        =   49
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   4260
         TabIndex        =   45
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   4260
         TabIndex        =   42
         Top             =   705
         Width           =   600
      End
      Begin VB.Label lblCir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "inches"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   4260
         TabIndex        =   39
         Top             =   375
         Width           =   600
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   6060
      TabIndex        =   15
      Top             =   45
      Width           =   3435
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   54
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtUnits 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2580
         TabIndex        =   50
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   58
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   56
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   48
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   52
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shoulder to Fold"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   19
      Left            =   600
      TabIndex        =   94
      Top             =   3690
      Width           =   1635
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Combo Box DDE "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   9840
      TabIndex        =   95
      Top             =   2010
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "bodydia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   bodydraw
'Purpose:   Body Drawing function
'
'
'Version:   3.00
'Date:      3.May.96
'Author:    Paul O'Rawe
'           Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Port to VB5
'
'Notes:-
'
'
'
'
    Option Explicit

   'Windows API Functions Declarations
'    Private Declare Function GetActiveWindow Lib "User" () As Integer
'    Private Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
'    Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
'    Private Declare Function GetNumTasks Lib "Kernel" () As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
   
   
   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

Private Sub cboCrotchStyle_Click()
    Dim sString As String, sError As String
    Dim iIndex As Integer
    
    sError = ""

   'Check Crotch style against sex
    sString = cboCrotchStyle.List(cboCrotchStyle.ListIndex)

   'Check for male crotch and female patient
    iIndex = 0
    iIndex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
    If iIndex <> 0 And bodydia!txtSex = "Female" Then
        sError = sError + "Male Crotch type specified for Female patient" + NL
    End If

   'Check for female crotch and male patient
    iIndex = 0
    iIndex = InStr(1, sString, "Female", 0)
    If iIndex <> 0 And bodydia!txtSex = "Male" Then
        sError = sError + "Female Crotch type specified for Male patient" + NL
    End If


   'Display Error message (if required) and return
   'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 0, "Crotch Style Error"
    End If

End Sub

Private Sub cboLeftCup_Click()
    If bodydia.Visible And (cboLeftCup.Text = "" Or cboLeftCup.Text = "None") Then
        txtLeftDisk = ""
        g_sLeftCup = ""
        g_sLeftDisk = ""
    Else
        PR_EnableCalculateDiskButton
    End If
End Sub

Private Sub cboLegStyle_Click()
   '
    Dim sLegStyle As String
    'Dim sReductionFactor  As String
    'Dim iAge As Integer

    'Get the currently selected leg style
    Select Case cboLegStyle.List(cboLegStyle.ListIndex)

    Case "Panty"
        prOptLeftLeg 1, 0
        prOptRightLeg 1, 0
       'Set dropdown combo boxes for reductions
       'Only set if Leg Style had been changed
       'If g_sPreviousLegStyle <> "" Then
            'prStore_Reductions
            'sReductionFactor = "0.85"
           'NB. Special case for TOS as TOS need not be given.
            'If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
            '    prSet_Red_Combo sReductionFactor, cboTOSRed
            'End If
            'prSet_Red_Combo sReductionFactor, cboWaistRed
            'prSet_Red_Combo sReductionFactor, cboMidPointRed
            'prSet_Red_Combo sReductionFactor, cboLargestRed
            'prSet_Red_Combo sReductionFactor, cboThighRed
        'End If

    Case "Brief", "Brief-French"
        prOptLeftLeg 0, 1
        prOptRightLeg 0, 1
       'Set dropdown combo boxes for reductions
       'Only set if Leg Style had been changed
       'If g_sPreviousLegStyle <> "" Then
            'prStore_Reductions
            'sReductionFactor = "0.85"
           'NB. Special case for TOS as TOS need not be given.
            'If Len(txtTOSCir.Text) > 0 And Len(txtTOSHt.Text) > 0 Or Len(txtUidWHBody.Text) = 0 Then
            '    prSet_Red_Combo sReductionFactor, cboTOSRed
            'End If
            'prSet_Red_Combo sReductionFactor, cboWaistRed
            'prSet_Red_Combo sReductionFactor, cboMidPointRed
            'prSet_Red_Combo sReductionFactor, cboLargestRed
            'prSet_Red_Combo sReductionFactor, cboThighRed
        'End If
    
    End Select

   'Store leg style for use when it changes
    g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

End Sub

Private Sub cboRightCup_Click()
    If bodydia.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
        txtRightDisk = ""
        g_sRightCup = ""
        g_sRightDisk = ""
    Else
        PR_EnableCalculateDiskButton
    End If
End Sub

Private Sub cmdCalculateBraDisks_Click()
    PR_DoBraCupsAndDisks
End Sub

Private Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)

End Function

Private Function FN_CalculateBra(nChest As Double, nUnderBreast As Double, nNipple As Double, sCup As String, iDisk As Integer, sSide As String) As Integer
   'Procedure to calcluate the size of a BRA
   'Returning the CUP size and the Bra disk size.
   '
   'INPUT
   '    nChest   Chest Circumference (Inches)
   '
   '    nNipple  Circumference over nipple line (Inches)
   '
   '    nUnderBreast
   '             Circumference just under the breast (Inches)
   '
   '
   '    sCup        Cup size
   '                If a cup size is given then disk size
   '                is based on the given cup.
   '
   '    sSide       Used when displaying error & warning messages
   '
   'OUTPUT
   '    sCup        If no cup is given then a cup is calculated
   '                and the disk size calculated from this cup.
   '
   '    iDisk       Size of disk to be used on template
   '
   '    FN_CalculateBra
   '                True if no errors
   '                False if errors, Message returned in sError
   '
   '    If sCup [and | or] iDisk are given then these are not changed
   '    the values are calculated and a warning given if the calculated
   '    values are different from the given
   '
   'NOTES
   '    Cup sizes   TRAINING, A, B, C, D, E, NONE
   '    Cup sizes   -1, 0 and 1 to 9
   '                Where 0 => a specified missing cup.
   '                and  -1 => failure to calculate.
   '
   'SPECIFICATIONS
   '    GOP 01-02/17    VEST WITHOUT SLEEVES
   '                    Pages 11,12 29.October.1991
   '    GOP 01-02/17    VEST WITH SLEEVES
   '                    Pages 17,18 29.October.1991
   '
   '    BODYBRA.D       DRAFIX macro version 1.04
   '
   '    BRACHART.DAT    Data used in conjunction with
   '                    the macro BODYBRA.D
   '

   'Variables
    Dim nDiff As Double, nSelctedDisk As Integer
    Dim sError$

   'Initially set to false
    FN_CalculateBra = False
    sError = ""

   'Simple case for Cup type = "None"
    If sCup = "None" Then
        FN_CalculateBra = True
        iDisk = 0
        Exit Function
    End If


   'Calulate bracup
    If sCup = "" Then
        If nNipple = 0 Then
            FN_CalculateBra = False
            sError = "Can't calculate a cup as Circumference over Nipple line is missing"
            MsgBox sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")"
            Exit Function
        End If
        
        nDiff = nNipple - nChest
        If nDiff < 0 Then
            FN_CalculateBra = False
            sError = "Circumference over Nipple line is smaller than Chest circumference"
            MsgBox sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")"
            Exit Function
        End If

       'Calculate bra cup
        If nDiff <= 1.375 Then
            sCup = "A"
        ElseIf nDiff <= 2.375 Then
            sCup = "B"
        ElseIf nDiff <= 3.375 Then
            sCup = "C"
        ElseIf nDiff <= 4.375 Then
            sCup = "D"
        Else
            sCup = "E"
        End If
    End If

    
   'Bra cup to disk mappings
    nSelctedDisk = -1
    Select Case nUnderBreast
        Case 26 To 28.99
            If sCup = "Training" Then
                nSelctedDisk = 1
            ElseIf sCup = "A" Then nSelctedDisk = 1
            ElseIf sCup = "B" Then nSelctedDisk = 2
            ElseIf sCup = "C" Then nSelctedDisk = 3
            ElseIf sCup = "D" Then nSelctedDisk = 4
            ElseIf sCup = "E" Then nSelctedDisk = 5
            Else: nSelctedDisk = -1
            End If
        Case 29 To 31.99
            If sCup = "Training" Then
                nSelctedDisk = 1
            ElseIf sCup = "A" Then nSelctedDisk = 2
            ElseIf sCup = "B" Then nSelctedDisk = 3
            ElseIf sCup = "C" Then nSelctedDisk = 4
            ElseIf sCup = "D" Then nSelctedDisk = 5
            ElseIf sCup = "E" Then nSelctedDisk = 6
            Else: nSelctedDisk = -1
            End If
        Case 32 To 34.99
            If sCup = "Training" Then
                nSelctedDisk = 2
            ElseIf sCup = "A" Then nSelctedDisk = 3
            ElseIf sCup = "B" Then nSelctedDisk = 4
            ElseIf sCup = "C" Then nSelctedDisk = 5
            ElseIf sCup = "D" Then nSelctedDisk = 6
            ElseIf sCup = "E" Then nSelctedDisk = 7
            Else: nSelctedDisk = -1
            End If
        Case 35 To 37.99
            If sCup = "Training" Then
                nSelctedDisk = 3
            ElseIf sCup = "A" Then nSelctedDisk = 4
            ElseIf sCup = "B" Then nSelctedDisk = 5
            ElseIf sCup = "C" Then nSelctedDisk = 6
            ElseIf sCup = "D" Then nSelctedDisk = 7
            ElseIf sCup = "E" Then nSelctedDisk = 8
            Else: nSelctedDisk = -1
            End If
        Case 38 To 40.99
            If sCup = "A" Then
                nSelctedDisk = 5
            ElseIf sCup = "B" Then nSelctedDisk = 6
            ElseIf sCup = "C" Then nSelctedDisk = 7
            ElseIf sCup = "D" Then nSelctedDisk = 8
            ElseIf sCup = "E" Then nSelctedDisk = 9
            Else: nSelctedDisk = -1
            End If
        Case 41 To 44
            If sCup = "A" Then
                nSelctedDisk = 6
            ElseIf sCup = "B" Then nSelctedDisk = 7
            ElseIf sCup = "C" Then nSelctedDisk = 8
            ElseIf sCup = "D" Then nSelctedDisk = 9
            Else: nSelctedDisk = -1
            End If
    End Select

    If nSelctedDisk = -1 Then
        sError = "No Bra disk is availabe for a Cup Size " & sCup & NL & "For an under breast circumference of " & fnInchestoText(nUnderBreast)
        MsgBox sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")"
        FN_CalculateBra = False
        Exit Function
    End If

   'Cut back disk by 1 if it is over 5
    If nSelctedDisk > 5 Then nSelctedDisk = nSelctedDisk - 1
    
    
   'Set disk size.
   'If disk is given then only check that size is the same, warn if not!
    If iDisk > 0 Then
        If iDisk <> nSelctedDisk Then
            sError = "Warning" & NL & "The Given disk is different in size to the disk that would have been calculated!"
            sError = sError & NL & "Given Disk      : " & Str$(iDisk)
            sError = sError & NL & "Calculated Disk : " & Str$(nSelctedDisk)
            MsgBox sError, 0, "BodySuit - Bra Cup" & "(" & sSide & ")"
        End If
    Else
        iDisk = nSelctedDisk
    End If
    
    FN_CalculateBra = True


End Function

Private Function FN_Decimalise(nDisplay As Double) As Double
   'This function takes the value given and converts it to its
   'decimal equivelent.
   '
   '
   'Input:-
   '        nDisplay is the value as input by the operator in the
   '        dialog.
   '        The convention is that, Metric dimensions use the decimal
   '        point to indicate the division between CMs and mms
   '        ie 7.6 = 7 cm and 6 mm.
   '        Whereas the decimal point for imperial measurements indicates
   '        the division between inches and eighths
   '        ie 7.6 = 7 inches and 6 eighths
   '
   'Globals:-
   '        g_nUnitsFac = 1       => nDisplay in Inches
   '        g_nUnitsFac = 10/25.5 => nDisplay in CMs
   '
   'Returns:-
   '        Single,
   '        For CMs value is returned unaltered.
   '        For Inches value is returned in decimal inches
   '        -1,     on conversion error.
   '
   'Errors:-
   '        The returned value is usually +ve. Unless it can't
   '        be sucessfully converted to inches.
   '        Eg 7.8 is an invalid number if g_nUnitsFac = 1
   '
   '                            WARNING
   '                            ~~~~~~~
   'In most cases the input is a +ve number.  This function will handle a
   '-ve number but in this case the error checking is invalid.  This
   'is done to provide a general conversion tool.  Where the input is
   'likley to be -ve then the calling subroutine or function should check
   'the sensibility of the returned value for that specific case.
   '

    Dim iInt As Integer, nDec As Double, iSign As Integer
   'retain sign
    iSign = Sgn(nDisplay)
    nDisplay = Abs(nDisplay)

   'Simple case where Units are CM
    If g_nUnitsFac <> 1 Then
        FN_Decimalise = nDisplay * iSign
        Exit Function
    End If

   'Imperial units
    iInt = Int(nDisplay)
    nDec = nDisplay - iInt
   'Check that conversion is possible (return -1 if not)
    If nDec > 0.8 Then
        FN_Decimalise = -1
    Else
        FN_Decimalise = (iInt + (nDec * 0.125 * 10)) * iSign
    End If

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
        If Char = """" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeQuotesInString = sEscapedString

End Function

Private Function FN_InchesValue(TextBox As Control) As Double
' Check for numeric values
    Dim sChar As String, sText As String
    Dim iLen As Integer, nn As Integer
    Dim nLen As Double

    sText = TextBox.Text
    iLen = Len(sText)

   'Check the actual structure of the input
    FN_InchesValue = -1
    For nn = 1 To iLen
        sChar = Mid$(sText, nn, 1)
        If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
            MsgBox "Invalid - Dimension has been entered", 48, "BodySuit - Dialogue"
            TextBox.SetFocus
            FN_InchesValue = -1
            Exit For
        End If
    Next nn

   'Convert to inches
    nLen = fnDisplayToInches(Val(TextBox.Text))
    If nLen = -1 Then
        MsgBox "Invalid - Length has been entered", 48, "BodySuit Details"
        TextBox.SetFocus
        FN_InchesValue = -1
    Else
        FN_InchesValue = nLen
    End If

End Function

Private Function FN_Open(sDrafixFile As String) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_Open = fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(13) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 32    'Bottom
    g_nCurrTextFont = 0         'BLOCK
    g_nCurrTextAngle = 0


   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; g_sPatient; ", "; g_sFileNo; ", Hand - "; g_sSide
    Print #fNum, "//by Visual Basic, BODYSUIT - Drawing"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;"
    Print #fNum, "XY     xyStart, xyOrigin, xyScale, xyO;"
    Print #fNum, "STRING sFileNo, sSide, sID, sPathJOBST;"
    Print #fNum, "ANGLE  aAngle;"

   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ZipperLength"; QCQ; "length"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Zipper"; QCQ; "string"; QQ; ");"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; "));"
    
   'Set values for use futher on by other macros
    Print #fNum, "sSide = "; QQ; g_sSide; QQ; ";"
    Print #fNum, "sFileNo = "; QQ; g_sFileNo; QQ; ";"
    
   'Get Start point
    Print #fNum, "GetUser ("; QQ; "xy"; QCQ; "Indicate Start Point"; QC; "&xyStart);"
   
   'Place a marker at the start point for later use.
   'Get a UID and create the unique 4 character start to the ID code
   'Note this is a bit dogey if the drawing contains more than 9999 entities
    PR_SetLayer "Construct"
    Print #fNum, "hOrigin = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart"; CC; "0.125);"
    Print #fNum, "if (hOrigin) {"
    Print #fNum, "  sID=StringMiddle(MakeString("; QQ; "long"; QQ; ",UID("; QQ; "get"; QQ; ",hOrigin)), 1, 4) ; "
    Print #fNum, "  while (StringLength(sID) < 4) sID = sID + "; QQ; " "; QQ; ";"
    Print #fNum, "  sID = sID + sFileNo + sSide ;"
    Print #fNum, "  SetDBData(hOrigin,"; QQ; "ID"; QQ; ",sID);"
    Print #fNum, "  }"
    
   'Display Hour Glass Symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
   'Start drawing on correct side
    PR_SetLayer "Template" & g_sSide

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

    For ii = 0 To 6
        sString = sString & txtCir(ii)
    Next ii
    'Reason for skip is that 7 & 8 were taken out during development
    For ii = 9 To 16
        sString = sString & txtCir(ii)
    Next ii
    
    For ii = 0 To 4
        sString = sString & cboRed(ii).Text
    Next ii
    
    sString = sString & cboLeftCup.Text & txtLeftDisk.Text
    
    sString = sString & cboRightCup.Text & txtRightDisk.Text

    sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text

    sString = sString & cboFrontNeck.Text & txtFrontNeck.Text

    sString = sString & cboBackNeck.Text & txtBackNeck.Text
    
    sString = sString & cboClosure.Text

    'Extra Values

    sString = sString & cboLegStyle.Text
   
   'Index 0   =>  Panty
   'Index 1   =>  Brief
    If optLeftLeg(0).Value = True Then
        sString = sString & "Panty"
    Else
        sString = sString & "Brief"
    End If
    If optRightLeg(0).Value = True Then
        sString = sString & "Panty"
    Else
        sString = sString & "Brief"
    End If
    
    sString = sString & cboCrotchStyle.Text
    
    
    FN_ValuesString = sString

End Function

Private Function fnInchestoText(nInches As Double) As String
   'Function returns a decimal value in inches as a string
   '
    Dim nPrecision As Double, nDec As Double
    Dim iInt As Integer, iEighths As Integer
    Dim sString As String
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
    If iInt <> 0 Then
        sString = LTrim$(Str$(iInt))
    Else
        sString = "  "
    End If
    If iEighths <> 0 Then
        Select Case iEighths
        Case 2, 6
            sString = sString + "-" + LTrim$(Str$(iEighths / 2)) + "/4"
        Case 4
            sString = sString + "-" + "1/2"
        Case Else
            sString = sString + "-" + LTrim$(Str$(iEighths)) + "/8"
        End Select
    Else
        sString = sString + "   "
    End If
       
   'Return formatted string
    fnInchestoText = sString

End Function

Private Function fnRoundInches(ByVal nNumber As Double) As Double
   'Function to return the rounded value in decimal inches
   'returns to the nearest eighth (0.125)
   'E.G.
   '    5.67         = 5 inches and 0.67 inches
   '                   0.67 / 0.125 = 5.36 eighths
   '                   5.36 eighths = 5 eighths (rounded to nearest eighth)
   '    5.67         = 5 inches and 5 eighths
   '    5.67         = 5 + ( 5 * 0.125)
   '    5.67         = 6.625 inches
   '

    Dim iInt As Integer, iSign As Integer
    Dim nPrecision As Double, nDec As Double


   'Return 0 if input is Zero
    If nNumber = 0 Then
       fnRoundInches = 0
       Exit Function
    End If

   'Set precision
    nPrecision = 0.125

   'Break input into components
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)
    nDec = nNumber - iInt

   'Get decimal part in precision units
    If nDec <> 0 Then
        nDec = nDec / nPrecision    'Avoid overflow
    End If
    nDec = round(nDec)

   'Return value
    fnRoundInches = (iInt + (nDec * nPrecision)) * iSign
End Function

Private Sub Form_LinkClose()
    Dim ii%, iError%

   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

   'Check that a "MainPatientDetails" Symbol has been
   'found
    If txtUidMPD = "" Then
        MsgBox "No Patient Details have been found in drawing!", 16, "Error, Bodysuit Drawing"
        End
    End If

   'Assign combo boxes (set the defaults if empty)
   'Reductions
    If txtCombo(0) <> "" Then cboRed(0).Text = txtCombo(0) Else cboRed(0).Text = ""
    If txtCombo(1) <> "" Then cboRed(1).Text = txtCombo(1) Else cboRed(1).Text = ""
    If txtCombo(2) <> "" Then cboRed(2).Text = txtCombo(2) Else cboRed(2).Text = ""
    If txtCombo(11) <> "" Then cboRed(3).Text = txtCombo(11) Else cboRed(3).Text = ""
    If txtCombo(12) <> "" Then cboRed(4).Text = txtCombo(12) Else cboRed(4).Text = ""
   
   'Bra cups and disk
    For ii = 0 To (cboLeftCup.ListCount - 1)
        If txtCombo(3).Text = cboLeftCup.List(ii) Then
            g_sLeftCup = cboLeftCup.List(ii)
            cboLeftCup.ListIndex = ii
        End If
    Next ii
    For ii = 0 To (cboRightCup.ListCount - 1)
        If txtCombo(4).Text = cboRightCup.List(ii) Then
            g_sRightCup = cboRightCup.List(ii)
            cboRightCup.ListIndex = ii
        End If
    Next ii

    g_sLeftDisk = txtLeftDisk.Text
    g_sRightDisk = txtRightDisk.Text

    If txtCombo(5) <> "" Then cboLeftAxilla.Text = txtCombo(5) Else cboLeftAxilla.ListIndex = 0
    If txtCombo(6) <> "" Then cboRightAxilla.Text = txtCombo(6) Else cboRightAxilla.ListIndex = 0
    If txtCombo(7) <> "" Then cboFrontNeck.Text = txtCombo(7) Else cboFrontNeck.ListIndex = 0
    If txtCombo(8) <> "" Then cboBackNeck.Text = txtCombo(8) Else cboBackNeck.ListIndex = 0
    If txtCombo(9) <> "" Then cboClosure.Text = txtCombo(9) Else cboClosure.ListIndex = 0
    cboFabric.Text = txtCombo(10)

   'Set up units
    If txtUnits = "cm" Then
        g_nUnitsFac = 10 / 25.4
    Else
        g_nUnitsFac = 1
    End If

   'Display dimesions sizes in inches
    For ii = 0 To 6
        txtCir_LostFocus (ii)
    Next ii
    For ii = 9 To 16
        txtCir_LostFocus (ii)
    Next ii
    txtFrontNeck_LostFocus
    txtBackNeck_LostFocus

   'Store values used to check for mods to bra cups
    g_nChest = Val(txtCir(5).Text)
    g_nUnderBreast = Val(txtCir(10).Text)
    g_nNipple = Val(txtCir(11).Text)

   'Store values to use on change etc
    Dim sChar As String * 1
    g_sBackNeck = txtBackNeck
    sChar = Left$(cboBackNeck.Text, 1)

    g_sFrontNeck = txtFrontNeck
    sChar = Left$(cboFrontNeck.Text, 1)
   
   'Set dropdown combo boxes
   'Crotch Style
    cboCrotchStyle.AddItem "Open Crotch"         'Both
    cboCrotchStyle.AddItem "Horizontal Fly"      'Male only
    cboCrotchStyle.AddItem "Diagonal Fly"        'Male only
    cboCrotchStyle.AddItem "Snap Crotch"         'Both
    cboCrotchStyle.AddItem "Gusset"              'Both

   'Set value
    For ii = 0 To (cboCrotchStyle.ListCount - 1)
         If cboCrotchStyle.List(ii) = txtCrotchStyle.Text Then
             cboCrotchStyle.ListIndex = ii
         End If
    Next ii
   
   'Leg Style
    cboLegStyle.AddItem "Panty"
    cboLegStyle.AddItem "Brief"
    cboLegStyle.AddItem "Brief-French"

   'Set values of combo and option buttons
    g_sPreviousLegStyle = "" 'this is used by cboLegStyle_Click
    For ii = 0 To (cboLegStyle.ListCount - 1)
         If cboLegStyle.List(ii) = fnGetString(txtLegStyle.Text, 1, " ") Then
             cboLegStyle.ListIndex = ii
         End If
    Next ii
 
   'Set this now to avoid problems with cboLegStyle_Click
    If cboLegStyle.ListIndex = -1 Then
         g_sPreviousLegStyle = "None"
    Else
         g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)
    End If

    If fnGetString(txtLegStyle.Text, 2, " ") <> "" Then
        If fnGetString(txtLegStyle.Text, 2, " ") = "Panty" Then
            optLeftLeg(0).Value = True
        Else
            optLeftLeg(1).Value = True 'Brief
        End If
    End If
    If fnGetString(txtLegStyle.Text, 3, " ") <> "" Then
        If fnGetString(txtLegStyle.Text, 3, " ") = "Panty" Then
            optRightLeg(0).Value = True
        Else
            optRightLeg(1).Value = True 'Brief
        End If
    End If
    
    'Create and check all values are present to allow creation of drawing.
    iError = FN_ValidateAndCalculateData(False)        'Sets variables from dialog box
                    'N.B. Use of local JOBST Directory C:\JOBST within this Routine
                    PR_CalculateBodySuit
                    PR_CreateDrawMacro
                    sTask = fnGetDrafixWindowTitleText()
                    If sTask <> "" Then
                            AppActivate sTask
                            SendKeys "@c:\jobst\draw.d{enter}", True
                        Else
                            MsgBox "Can't find a Drafix Drawing to update! - Terminating."
                    End If

      'Missing values - User is informed of which in FnAllValuesPresent
       Screen.MousePointer = 0    ' Change pointer back to normal.
    End

End Sub

Private Sub Form_Load()
    Dim ii%

    Set MainForm = bodydia  'load

    NL = Chr$(13) 'The new line character
    g_sSide = Command$
    
    Hide

   'Start a timer
   'The Timer is disabled in LinkClose
   'If after 6 seconds the timer event will "End" the programme
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 6000  'Approx 6 Seconds
    Timer1.Enabled = True

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "Body Draw Dialogue is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting BodySuit - Dialogue"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - bodydia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - bodydia.Height) / 2  ' Center form vertically.

    g_nUnitsFac = 1
    g_sPathJOBST = fnPathJOBST()

   'Clear fields
   'Circumferences and lengths
    For ii = 0 To 6
        txtCir(ii) = ""
    Next ii
    For ii = 9 To 16
        txtCir(ii) = ""
    Next ii

   'The data from these DDE text boxes is copied
   'to the combo boxes on Link close
   '
    For ii = 0 To 12
        txtCombo(ii).Text = ""
    Next ii

   'Bra cups
    txtLeftDisk = ""
    txtRightDisk.Text = ""
   
   'Patient details
    txtFileNo = ""
    txtUnits = ""
    txtPatientName = ""
    txtDiagnosis = ""
    txtAge = ""
    txtSex = ""
    txtWorkOrder = ""

   'Design choices
    txtFrontNeck = ""
    txtBackNeck = ""

   'UID of symbols
    txtUidMPD = ""
    txtUidVB = ""
    
   'Setup combo box fields
   'Waist Default 85% +/- 5%
    cboRed(0).AddItem "0.90"
    cboRed(0).AddItem "0.80"
    cboRed(0).AddItem ""
    
    'Chest Default 90% +/- 5%
    cboRed(1).AddItem "0.95"
    cboRed(1).AddItem "0.85"
    cboRed(1).AddItem ""

    'Under Breast 90% +/- 5%
    cboRed(2).AddItem "0.95"
    cboRed(2).AddItem "0.85"
    cboRed(2).AddItem ""

    'Large part of Buttocks 85% +/- 5%
    cboRed(3).AddItem "0.90"
    cboRed(3).AddItem "0.80"
    cboRed(3).AddItem ""
    
    'Left Thigh 85% +/- 5%
    cboRed(4).AddItem "0.90"
    cboRed(4).AddItem "0.80"
    cboRed(4).AddItem ""
    'Right Thigh uses Left Thigh cboRed value
    
    cboLeftCup.AddItem "Training"
    cboLeftCup.AddItem "A"
    cboLeftCup.AddItem "B"
    cboLeftCup.AddItem "C"
    cboLeftCup.AddItem "D"
    cboLeftCup.AddItem "E"
    cboLeftCup.AddItem "None"
    cboLeftCup.AddItem ""

    cboRightCup.AddItem "Training"
    cboRightCup.AddItem "A"
    cboRightCup.AddItem "B"
    cboRightCup.AddItem "C"
    cboRightCup.AddItem "D"
    cboRightCup.AddItem "E"
    cboRightCup.AddItem "None"
    cboRightCup.AddItem ""

    cboLeftAxilla.AddItem "Regular 2" & QQ
    cboLeftAxilla.AddItem "Regular 1½" & QQ
    cboLeftAxilla.AddItem "Regular 2½" & QQ
    cboLeftAxilla.AddItem "Open"
    cboLeftAxilla.AddItem "Mesh"
    cboLeftAxilla.AddItem "Lining"
    cboLeftAxilla.AddItem "Sleeveless"

    cboRightAxilla.AddItem "Regular 2" & QQ
    cboRightAxilla.AddItem "Regular 1½" & QQ
    cboRightAxilla.AddItem "Regular 2½" & QQ
    cboRightAxilla.AddItem "Open"
    cboRightAxilla.AddItem "Mesh"
    cboRightAxilla.AddItem "Lining"
    cboRightAxilla.AddItem "Sleeveless"

    cboFrontNeck.AddItem "Regular"
    cboFrontNeck.AddItem "Scoop"
    cboFrontNeck.AddItem "Measured Scoop"
    cboFrontNeck.AddItem "V neck"
    cboFrontNeck.AddItem "Turtle"
    cboFrontNeck.AddItem "Turtle - Fabric same as Vest"
    cboFrontNeck.AddItem "Turtle Detachable"
    cboFrontNeck.AddItem "Turtle Detach. Fabric"

    cboBackNeck.AddItem "Regular"
    cboBackNeck.AddItem "Scoop"
    cboBackNeck.AddItem "Measured Scoop"

    cboClosure.AddItem "Front Zip"
    cboClosure.AddItem "Back Zip"
  
    PR_GetComboListFromFile cboFabric, g_sPathJOBST & "\FABRIC.DAT"


End Sub

Private Function min(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the minimum of two numbers
    If nFirst <= nSecond Then
        min = nFirst
    Else
        min = nSecond
    End If
End Function

Private Static Sub PR_CalculateBodySuit()
'This function caclulates the body suit
'
'This function is used by the project
'   BODYDRAW.MAK
'

    Dim nBackNeckOffset As Double
    Dim nFrontAngle As Double
    Dim nRaglanOffset As Double
    Dim nShldWidthOffset As Double
    Dim nFrontScoopDistance As Double
    Dim nBackScoopDistance As Double
    Dim nScoopAngle As Double
    Dim nGussetSize As Double
    Dim nGussetLength As Double
    Dim nGussetRadius As Double
    Dim nCutSegmentAngle As Double
    Dim nCutArcLength As Double
    Dim nStrapAngle As Double
    Dim nStrapLength As Double
    Dim nGussetAngle As Double
    Dim n8_7Length As Double
    Dim iError As Integer
    Dim nCutOut7Angle As Double
    Dim nTmpRadius As Double
    Dim nCrotchFilletRadius As Double
    Dim nCutExtendedLength As Double
    Dim xyTmpCutOut7 As XY
    Dim xyTmp1 As XY
    Dim xyTmp2 As XY
    Dim xyTmp3 As XY
    Dim bShldWidthModified As Integer
    Dim nTangentAngle As Double
    Dim nGroinAngle As Double
    Dim nThighAngle As Double
    Dim nGroinToThighAngle As Double
    Dim nGroinToThighMidPoint As Double
    Dim sLeg As String
    Dim sSleeve As String
    Dim xyRegularTmpCutOut8 As XY
    Dim nRegularRadiusFrontNeck  As Double
    Dim aAngle As Double
    Dim nLength As Double
    Dim ii As Integer

   'Drawing variables
    g_bMissCutOut7 = False
    nScoopAngle = 0
    nFrontScoopDistance = 0
    nBackScoopDistance = 0
   
   'Body Co-ordinate points and values
   'Setup which side to be drawn
    If g_bDiffAxillaHeight Then
       'Draw the body using the highest axilla
        If g_nLeftShldToAxilla < g_nRightShldToAxilla Then
                g_nShldToAxilla = g_nLeftShldToAxilla
                g_nShldToAxillaLow = g_nRightShldToAxilla
                g_nAxillaCir = g_nLeftAxillaCir
                g_nAxillaCirLow = g_nRightAxillaCir
                g_sAxillaSide = "Left"
                g_sAxillaSideLow = "Right"
                g_sAxillaType = g_sLeftSleeve
                g_sAxillaTypeLow = g_sRightSleeve

            Else
                g_nShldToAxilla = g_nRightShldToAxilla
                g_nShldToAxillaLow = g_nLeftShldToAxilla
                g_nAxillaCir = g_nRightAxillaCir
                g_nAxillaCirLow = g_nLeftAxillaCir
                g_sAxillaSide = "Right"
                g_sAxillaSideLow = "Left"
                g_sAxillaType = g_sRightSleeve
                g_sAxillaTypeLow = g_sLeftSleeve
        End If
    Else
       'Default to left side
        g_sAxillaSide = "Left" 'Draw axilla by default on left side
        g_sAxillaType = g_sLeftSleeve
        g_nShldToAxilla = g_nLeftShldToAxilla
        g_nAxillaCir = g_nLeftAxillaCir
    End If
    
    PR_MakeXY xySeamFold, 0, INCH1_4
    PR_MakeXY xySeamButt, xySeamFold.X + g_nButtocksLength, xySeamFold.y
    PR_MakeXY xySeamHighShld, xySeamFold.X + g_nShldToFold, xySeamFold.y
    PR_MakeXY xySeamLowShld, xySeamHighShld.X - INCH1_2, xySeamFold.y
    PR_MakeXY xySeamChest, xySeamHighShld.X - g_nShldToAxilla, xySeamFold.y
    PR_MakeXY xySeamChestAxillaLow, xySeamHighShld.X - g_nShldToAxillaLow, xySeamFold.y
    PR_MakeXY xySeamWaist, xySeamHighShld.X - g_nShldToWaist, xySeamFold.y
    PR_MakeXY xyProfileButt, xySeamButt.X, xySeamButt.y + g_nButtCirCalc
    PR_MakeXY xyCutOut3, xyProfileButt.X, xyProfileButt.y - g_nButtBackSeam
    PR_MakeXY xyCutOut4, xySeamFold.X + INCH3_4, xyCutOut3.y - (g_nButtCutOut / 2)
    PR_MakeXY xyCutOut5, xySeamButt.X, xySeamButt.y + g_nButtFrontSeam
    PR_MakeXY xyCutOut6, xySeamWaist.X, xySeamWaist.y + g_nWaistFrontSeam
    PR_MakeXY xyCutOut2, xyCutOut6.X, xyCutOut6.y + g_nCutOut
    PR_MakeXY xyProfileWaist, xyCutOut2.X, xyCutOut2.y + g_nWaistBackSeam
    PR_MakeXY xyCutOut7, xySeamChest.X, xySeamChest.y + g_nChestFrontSeam
    PR_MakeXY xyProfileChest, xySeamChest.X, xyCutOut2.y + g_nChestBackSeam

'    g_nCutOutRadius = (FN_CalcLength(xyCutOut4, xyCutOut3) / 2) / Cos(FN_CalcAngle(xyCutOut4, xyCutOut3) * (PI / 180))
   'Original 85% mark

    If xyCutOut5.X - xyCutOut4.X > (xyCutOut3.y - xyCutOut5.y) / 2 Then
       'Extreme measurement case
        g_bExtremeCrotch = True
        g_nCutOutRadius = (FN_CalcLength(xyCutOut5, xyCutOut3) / 2)
        PR_MakeXY xyCutOut9, xyCutOut4.X + g_nCutOutRadius, xyCutOut3.y
        PR_MakeXY xyCutOut10, xyCutOut9.X, xyCutOut5.y
    Else
       'Normal case
        aAngle = FN_CalcAngle(xyCutOut4, xyCutOut3)
        If aAngle <> 90 Then
            g_nCutOutRadius = (FN_CalcLength(xyCutOut4, xyCutOut3) / 2) / Cos(aAngle * (PI / 180))
        Else
            g_nCutOutRadius = FN_CalcLength(xyCutOut4, xyCutOut3) / 2
        End If
        xyCutOut9 = xyCutOut3
        xyCutOut10 = xyCutOut5
    End If
    
    PR_MakeXY xyCutOutArcCentre, xyCutOut4.X + g_nCutOutRadius, xyCutOut4.y
   
   'Calculate Centre Point of Buttocks Arc for the averaged/smallest thigh
    PR_MakeXY xyButtocksArcCentre, xySeamButt.X, xySeamButt.y + g_nHalfGroinHeight
    PR_MakeXY xyProfileGroin, xySeamFold.X + INCH3_4, xySeamFold.y + g_nGroinHeight

   'Calculate above for largest thigh (if given)
'    If g_bDiffThigh Then
        PR_MakeXY xyLT_ButtocksArcCentre, xySeamButt.X, xySeamButt.y + g_nLT_HalfGroinHeight
        PR_MakeXY xyLT_ProfileGroin, xySeamFold.X + INCH3_4, xySeamFold.y + g_nLT_GroinHeight
'    End If


    'FRONT NECK
    'Needed variables for neck
    If g_bSleeveless Then
            nBackNeckOffset = fnRoundInches(g_nNeckCir / 6)
            g_nRadiusFrontNeck = fnRoundInches(((g_nNeckCir - (nBackNeckOffset * 2)) / 3.14) - INCH1_8)  'Full scale
        Else
            nBackNeckOffset = fnRoundInches(g_nNeckCir / 9)
            g_nRadiusFrontNeck = fnRoundInches(((g_nNeckCir - ((nBackNeckOffset * 2) + 1.5)) / 3.14) - INCH1_8)  'Full scale
    End If
    If InStr(1, g_sFrontNeckStyle, "Scoop") > 0 Or g_sFrontNeckStyle = "V neck" Then
        'Scoop, Measured scoop, V neck and Measured V neck
        PR_MakeXY xyRegularTmpCutOut8, xySeamHighShld.X - g_nRadiusFrontNeck, xySeamHighShld.y + g_nRadiusFrontNeck
        g_nRadiusFrontNeck = g_nRadiusFrontNeck * 1.2
        g_nShldWidth = g_nShldWidth - 1
        bShldWidthModified = True
    End If
    PR_MakeXY xyTmpCutOut8, xySeamHighShld.X - g_nRadiusFrontNeck, xySeamHighShld.y + g_nRadiusFrontNeck

    PR_MakeXY xyFrontNeckArcCentre, xySeamHighShld.X, xySeamHighShld.y + g_nRadiusFrontNeck
    
    Select Case g_sFrontNeckStyle
        Case "Measured Scoop", "Measured V neck"
            n8_7Length = FN_CalcLength(xyRegularTmpCutOut8, xyCutOut7)
            If (n8_7Length < g_nFrontNeckSize) Then
                    nScoopAngle = FN_CalcAngle(xyCutOut7, xyCutOut6)
                    PR_CalcPolar xyCutOut7, nScoopAngle, g_nFrontNeckSize - n8_7Length, xyCutOut8
                    g_bMissCutOut7 = True
                   'Get intersection on ordinary scoop neck line
                    xyTmp1 = xyCutOut8
                    xyTmp1.y = xyCutOut8.y + 10  'Extend line
                    ii = FN_LinLinInt(xyCutOut6, xyTmpCutOut8, xyCutOut8, xyTmp1, xyCutOut8)
                Else
                    nScoopAngle = FN_CalcAngle(xyRegularTmpCutOut8, xyCutOut7)
                    PR_CalcPolar xyRegularTmpCutOut8, nScoopAngle, g_nFrontNeckSize, xyCutOut8
                    xyTmp1 = xyCutOut8
                    xyTmp1.y = xyCutOut8.y + 10  'Extend line
                    ii = FN_LinLinInt(xyCutOut7, xyTmpCutOut8, xyCutOut8, xyTmp1, xyCutOut8)
            End If
           ''revise y of xyCutOut8
           'PR_MakeXY xyCutOut8, xyCutOut8.x, xyTmpCutOut8.y
        Case Else
            'Regular, V neck and Scoop
            xyCutOut8 = xyTmpCutOut8
    End Select


   'BACK NECK Check for scoop
    If InStr(g_sBackNeckStyle, "Scoop") <> 0 Then
        If Not bShldWidthModified Then
            g_nShldWidth = g_nShldWidth - 1
        End If
        'Check for scoops here
        If g_sBackNeckStyle = "Scoop" Then
                If g_nAdult Then
                        nBackScoopDistance = INCH3_4
                    Else
                        nBackScoopDistance = INCH3_8
                End If
            Else
                nBackScoopDistance = g_nBackNeckSize
        End If
    End If
    nFrontAngle = -((360 / (2 * PI * g_nRadiusFrontNeck)) + 90)
    If (g_sSide = "Left") Then
            sSleeve = g_sLeftSleeve
        Else
            sSleeve = g_sRightSleeve
    End If


    If sSleeve = "Sleeveless" Then
            'Neck Co-ordinate points and values
            PR_MakeXY xyCutOutBackNeck, xySeamHighShld.X - (INCH1_4 + nBackScoopDistance), xyCutOut2.y
            PR_MakeXY xyProfileNeck, xySeamHighShld.X, xyCutOut2.y + nBackNeckOffset
            xyCutOutFrontNeck = xySeamHighShld
            PR_MakeXY xyProfileNeckMirror, xyProfileNeck.X, -xyProfileNeck.y
            
            'Axilla Co-ordinate points and values
            'Because it's sleeveless the axilla length is figured
            'Needed variables for Axilla
            nRaglanOffset = ((xyCutOutFrontNeck.y - xyProfileNeckMirror.y) - INCH1_2) / 2
            'This variable is only used for Sleeveless
            nShldWidthOffset = Sqr((g_nShldWidth ^ 2) - (INCH1_2 ^ 2))
            If (nRaglanOffset < nShldWidthOffset) Then
                nShldWidthOffset = nRaglanOffset
            End If
            PR_MakeXY xyRaglan1, xySeamHighShld.X, xySeamHighShld.y - INCH1_4
            PR_MakeXY xyRaglan2, xySeamLowShld.X, xyRaglan1.y - nShldWidthOffset
            PR_MakeXY xyRaglan3, xySeamLowShld.X, xyRaglan1.y - nRaglanOffset
            PR_MakeXY xyRaglan4, xyRaglan3.X - (g_nAxillaCir / 2), xyRaglan3.y
            If g_bDiffAxillaHeight Then
                PR_MakeXY xyRaglan4LowAxilla, xyRaglan3.X - (g_nAxillaCirLow / 2), xyRaglan3.y
            End If
            PR_MakeXY xyRaglan6, xyProfileNeckMirror.X, xyProfileNeckMirror.y + INCH1_4
            PR_MakeXY xyRaglan5, xySeamLowShld.X, xyRaglan6.y + nShldWidthOffset
    Else
            'Neck Co-ordinate points and values
            PR_MakeXY xyCutOutBackNeck, xySeamLowShld.X - (INCH1_4 + nBackScoopDistance), xyCutOut2.y
            PR_MakeXY xyProfileNeck, xySeamLowShld.X, xyCutOut2.y + nBackNeckOffset
            PR_CalcPolar xyFrontNeckArcCentre, nFrontAngle, g_nRadiusFrontNeck, xyCutOutFrontNeck
            PR_MakeXY xyProfileNeckMirror, xyProfileNeck.X, -xyProfileNeck.y
            
            'Axilla Co-ordinate points and values
            'Needed variables for Axilla
            nRaglanOffset = ((xyCutOutFrontNeck.y - xyProfileNeckMirror.y) - INCH1_2) / 2
            PR_MakeXY xyRaglan1, xyCutOutFrontNeck.X, xyCutOutFrontNeck.y - INCH1_4
            PR_MakeXY xyRaglan2, xySeamChest.X, xySeamChest.y - nRaglanOffset
            PR_MakeXY xyRaglan3, xyProfileNeckMirror.X, xyProfileNeckMirror.y + INCH1_4
            If g_bDiffAxillaHeight Then
                PR_MakeXY xyRaglan2LowAxilla, xySeamHighShld.X - g_nShldToAxillaLow, xySeamChest.y - nRaglanOffset
                If g_sAxillaSide = "Left" Then
                    g_nAxillaBackNeckRad = FN_CalcLength(xyRaglan2, xyRaglan3)
                    g_nAxillaFrontNeckRad = FN_CalcLength(xyRaglan2, xyRaglan1)
                    g_nAFNRadRight = FN_CalcLength(xyRaglan2LowAxilla, xyRaglan1)
                    g_nABNRadRight = FN_CalcLength(xyRaglan2LowAxilla, xyRaglan3)
                Else
                    g_nAxillaBackNeckRad = FN_CalcLength(xyRaglan2LowAxilla, xyRaglan3)
                    g_nAxillaFrontNeckRad = FN_CalcLength(xyRaglan2LowAxilla, xyRaglan1)
                    g_nAFNRadRight = FN_CalcLength(xyRaglan2, xyRaglan1)
                    g_nABNRadRight = FN_CalcLength(xyRaglan2, xyRaglan3)
                End If
            Else
                g_nAxillaBackNeckRad = FN_CalcLength(xyRaglan2, xyRaglan3)
                g_nAxillaFrontNeckRad = FN_CalcLength(xyRaglan2, xyRaglan1)
                g_nABNRadRight = 0
                g_nAFNRadRight = 0
            End If
    End If


    If g_bDiffThigh Then     'One panty & brief
            'Release extra 5% reduction
            g_nThighCir = fnRoundInches((g_nSmallestThighGiven * (g_nThighCirRed + 0.05))) / 2 'half scale
        Else
            g_nThighCir = fnRoundInches((g_nSmallestThighGiven * g_nThighCirRed)) / 2  'half scale
    End If
    
   'Calculate Crotch Co-ordinate points and values
    Select Case g_sCrotchStyle
        Case "Open Crotch"
            'Set sizes
            PR_SetCrotchSizes
            If g_sSex = "Male" Then
                    nCrotchFilletRadius = INCH3_8
                    If Not g_nAdult Then
                        nCrotchFilletRadius = INCH3_16
                    End If
                Else
                    nCrotchFilletRadius = INCH3_16
            End If
            nTmpRadius = g_nCutOutRadius + nCrotchFilletRadius

            If g_bExtremeCrotch Then
                'Extreme crotch
                nCutArcLength = (2 * PI * g_nCutOutRadius) / 4
                nCutExtendedLength = FN_CalcLength(xyCutOut10, xyCutOut5)
                'Back crotch
                If nCutArcLength > g_nBackCrotchSize Then
                    'Open crotch ends on arc
                    xyBackCrotch1 = xyCutOut9
                    PR_CalcPolar xyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBackCrotch1
                    xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
                    xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
                    xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
                    xyTmp2.y = xyBackCrotch1.y + 1
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
                    PR_MakeXY xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y
                    PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                    iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
                ElseIf (nCutArcLength + nCutExtendedLength) < g_nBackCrotchSize Then
                    'Open crotch ends at xyCutOut3
                    'IE we can't go past the Buttock height
                    xyBackCrotch1 = xyCutOut3
                    PR_MakeXY xyBackCrotch2, xyBackCrotch1.X, xyBackCrotch1.y + nCrotchFilletRadius
                    PR_MakeXY xyBackCrotchFilletCentre, xyBackCrotch1.X - nCrotchFilletRadius, xyBackCrotch2.y
                    PR_MakeXY xyBackCrotch3, xyBackCrotchFilletCentre.X, xyBackCrotchFilletCentre.y + nCrotchFilletRadius
                Else
                    'Open crotch ends between xyCutOut9 and xyCutout3
                    xyBackCrotch1 = xyCutOut3
                    xyBackCrotch1.X = xyCutOut3.X - ((nCutArcLength + nCutExtendedLength) - g_nBackCrotchSize)
                    'Get center of fillet radius(initially relative to the arc)
                    If (xyBackCrotch1.X - nCrotchFilletRadius) < xyCutOut9.X Then
                        'Fillet intersects arc
                        xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
                        xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
                        xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
                        xyTmp2.y = xyBackCrotch1.y + 1
                        iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
                        PR_MakeXY xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y
                        PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                        iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
                    Else
                        'Fillet is on straight bit
                        PR_MakeXY xyBackCrotch2, xyBackCrotch1.X, xyBackCrotch1.y + nCrotchFilletRadius
                        PR_MakeXY xyBackCrotchFilletCentre, xyBackCrotch1.X - nCrotchFilletRadius, xyBackCrotch2.y
                        PR_MakeXY xyBackCrotch3, xyBackCrotchFilletCentre.X, xyBackCrotchFilletCentre.y + nCrotchFilletRadius
                    End If
                End If
                
                'Front crotch
                If nCutArcLength > g_nFrontCrotchSize Then
                    'Open crotch ends on arc
                    xyFrontCrotch1 = xyCutOut10
                    xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
                    xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
                    xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
                    xyTmp2.y = xyFrontCrotch1.y - 1
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
                    PR_MakeXY xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y
                    PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                    iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
                ElseIf (nCutArcLength + nCutExtendedLength) < g_nFrontCrotchSize Then
                    'Open crotch ends at xyCutOut5
                    'IE we can't go past the Buttock height
                    xyFrontCrotch1 = xyCutOut5
                    PR_MakeXY xyFrontCrotch2, xyFrontCrotch1.X, xyFrontCrotch1.y - nCrotchFilletRadius
                    PR_MakeXY xyFrontCrotchFilletCentre, xyFrontCrotch1.X - nCrotchFilletRadius, xyFrontCrotch2.y
                    PR_MakeXY xyFrontCrotch3, xyFrontCrotchFilletCentre.X, xyFrontCrotchFilletCentre.y - nCrotchFilletRadius
                Else
                    'Open crotch ends between xyCutOut10 and xyCutout5
                    xyFrontCrotch1 = xyCutOut5
                    xyFrontCrotch1.X = xyCutOut5.X - ((nCutArcLength + nCutExtendedLength) - g_nFrontCrotchSize)
                    'Get center of fillet radius(initially relative to the arc)
                    If (xyFrontCrotch1.X - nCrotchFilletRadius) < xyCutOut10.X Then
                        'Fillet intersects arc
                        xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
                        xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
                        xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
                        xyTmp2.y = xyFrontCrotch1.y - 1
                        iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
                        PR_MakeXY xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y
                        PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                        iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
                    Else
                        'Fillet is on straight bit
                        PR_MakeXY xyFrontCrotch2, xyFrontCrotch1.X, xyFrontCrotch1.y - nCrotchFilletRadius
                        PR_MakeXY xyFrontCrotchFilletCentre, xyFrontCrotch1.X - nCrotchFilletRadius, xyFrontCrotch2.y
                        PR_MakeXY xyFrontCrotch3, xyFrontCrotchFilletCentre.X, xyFrontCrotchFilletCentre.y - nCrotchFilletRadius
                    End If
                End If
                
            Else
                'Normal crotch
                nCutSegmentAngle = FN_CalcAngle(xyCutOutArcCentre, xyCutOut5) - FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
                nCutArcLength = fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                If (nCutArcLength <= g_nBackCrotchSize) Then
                        xyBackCrotch1 = xyCutOut3
                    Else
                        PR_CalcPolar xyCutOutArcCentre, 180 - ((g_nBackCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyBackCrotch1
                End If
                'Back of crotch
                xyTmp1.X = xyBackCrotch1.X - nCrotchFilletRadius
                xyTmp1.y = xyBackCrotch1.y - nCrotchFilletRadius
                xyTmp2.X = xyBackCrotch1.X - nCrotchFilletRadius
                xyTmp2.y = xyBackCrotch1.y + 1
                iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyBackCrotchFilletCentre)
                PR_MakeXY xyBackCrotch2, xyBackCrotchFilletCentre.X + nCrotchFilletRadius, xyBackCrotchFilletCentre.y
                PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyBackCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyBackCrotch3)
                'Front of crotch
                If (nCutArcLength <= g_nFrontCrotchSize) Then
                        xyFrontCrotch1 = xyCutOut5
                    Else
                        PR_CalcPolar xyCutOutArcCentre, 180 + ((g_nFrontCrotchSize / (2 * PI * g_nCutOutRadius)) * 360), g_nCutOutRadius, xyFrontCrotch1
                End If
                xyTmp1.X = xyFrontCrotch1.X - nCrotchFilletRadius
                xyTmp1.y = xyFrontCrotch1.y + nCrotchFilletRadius
                xyTmp2.X = xyFrontCrotch1.X - nCrotchFilletRadius
                xyTmp2.y = xyFrontCrotch1.y - 1
                nTmpRadius = g_nCutOutRadius + nCrotchFilletRadius
                iError = FN_CirLinInt(xyTmp1, xyTmp2, xyCutOutArcCentre, nTmpRadius, xyFrontCrotchFilletCentre)
                PR_MakeXY xyFrontCrotch2, xyFrontCrotchFilletCentre.X + nCrotchFilletRadius, xyFrontCrotchFilletCentre.y
                PR_CalcPolar xyCutOutArcCentre, FN_CalcAngle(xyCutOutArcCentre, xyFrontCrotchFilletCentre), g_nCutOutRadius + 1, xyTmp3
                iError = FN_CirLinInt(xyCutOutArcCentre, xyTmp3, xyCutOutArcCentre, g_nCutOutRadius + (2 * nCrotchFilletRadius), xyFrontCrotch3)
            End If
            'Common
            PR_MakeXY xyLegPoint, INCH3_4, 0
            PR_MakeXY xyCrotchMarker, xyCutOut4.X - (2 * nCrotchFilletRadius), xyCutOut4.y
            If g_sSex = "Male" Then
                    '1 1/4" + 3/4" to xyCutOut4 = 2"
                    PR_MakeXY xySeamThigh, xySeamFold.X - (1 + INCH1_4), xySeamFold.y
                    If Not g_nAdult Then
                        '7/8" + 3/4" to xyCutOut4 = 1 5/8"
                        PR_MakeXY xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y
                    End If
                Else
                    '7/8" + 3/4" to xyCutOut4 = 1 5/8"
                    PR_MakeXY xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y
            End If
            'Test that at least 1-1/4" between cutout edge and bottom of brief.
            If (xyCrotchMarker.X - xySeamThigh.X) < 1.25 Then   'less than 1-1/4"
                PR_MakeXY xySeamThigh, xyCrotchMarker.X - 1.25, xySeamFold.y
            End If
            PR_MakeXY xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2)
        
        Case "Horizontal Fly"
            '7/8" + 3/4" to xyCutOut4 = 1 5/8"
            PR_MakeXY xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y
            PR_MakeXY xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2)
            PR_SetFlySize "Hor"
            PR_MakeXY xyLegPoint, 0, 0
            PR_MakeXY xyCrotchMarker, xyCutOut4.X, xyCutOut4.y
        Case "Diagonal Fly"
            '7/8" + 3/4" to xyCutOut4 = 1 5/8"
            PR_MakeXY xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y
            PR_MakeXY xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2)
            PR_SetFlySize "Dia"
            PR_MakeXY xyLegPoint, 0, 0
            PR_MakeXY xyCrotchMarker, xyCutOut4.X, xyCutOut4.y
        
        Case "Gusset"
            '7/8" + 3/4" to xyCutOut4 = 1 5/8"
            PR_MakeXY xySeamThigh, xySeamFold.X - INCH7_8, xySeamFold.y
            PR_MakeXY xyProfileBrief, xySeamThigh.X, xySeamThigh.y + (g_nHalfGroinHeight / 2)
            PR_SetGussetSize
            PR_MakeXY xyLegPoint, 0, 0
            PR_MakeXY xyCrotchMarker, xyCutOut4.X, xyCutOut4.y
        
        Case "Snap Crotch"
            '1/2" + 3/4" to xyCutOut4 = 1 1/4"
            PR_MakeXY xySeamThigh, xySeamFold.X - INCH1_2, xySeamFold.y
            PR_MakeXY xyProfileBrief, xySeamThigh.X, xyCutOut4.y
            PR_SetGussetSize
            nGussetLength = g_nGussetLength + 1.5   '1-1/2"
            If nGussetLength < 2.5 Then     '2-1/2" Minimum
                MsgBox ">>>>>>>>>>>>>>>>>WARNING <<<<<<<<<<<<" & Chr$(13) & "Gusset Strap shorter than 2-1/2"" which is unacceptable."
                ' End
            End If
            g_nLengthStrap1ToCutOut5 = 0
            If g_bExtremeCrotch Then

                'Special case for extreme crotch
                nCutSegmentAngle = FN_CalcAngle(xyCutOutArcCentre, xyCutOut10) - FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
                nCutArcLength = fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                nCutExtendedLength = FN_CalcLength(xyCutOut10, xyCutOut5)

                If nCutArcLength > nGussetLength Then
                        'End of strap ends on arc
                        nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / PI))
                        PR_CalcPolar xyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyStrap1
                        g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength

                ElseIf (nCutArcLength + nCutExtendedLength) > nGussetLength Then
                        'End of strap ends between xyCutOut10 and xyCutout5
                        nStrapAngle = FN_CalcAngle(xyCutOut10, xyCutOut5)
                        nStrapLength = nGussetLength - nCutArcLength
                        PR_CalcPolar xyCutOut10, nStrapAngle, nStrapLength, xyStrap1
                        g_nLengthStrap1ToCutOut5 = FN_CalcLength(xyStrap1, xyCutOut5)
                Else
                        'End of strap ends between xyCutOut5 and xyCutout6
                        nStrapAngle = FN_CalcAngle(xyCutOut5, xyCutOut6)
                        nStrapLength = nGussetLength - (nCutArcLength + nCutExtendedLength)
                        PR_CalcPolar xyCutOut5, nStrapAngle, nStrapLength, xyStrap1
                        g_nLengthStrap1ToCutOut5 = 0
                End If
            Else
               'Normal crotch
                nCutSegmentAngle = FN_CalcAngle(xyCutOutArcCentre, xyCutOut5) - FN_CalcAngle(xyCutOutArcCentre, xyCutOut4)
                nCutArcLength = fnRoundInches((2 * PI * g_nCutOutRadius) * (nCutSegmentAngle / 360))
                If (nCutArcLength < nGussetLength) Then
                        nStrapAngle = FN_CalcAngle(xyCutOut5, xyCutOut6)
                        nStrapLength = nGussetLength - nCutArcLength
                        PR_CalcPolar xyCutOut5, nStrapAngle, nStrapLength, xyStrap1
                        g_nLengthStrap1ToCutOut5 = 0
                Else
                        nGussetAngle = 180 + ((nGussetLength / g_nCutOutRadius) * (180 / PI))
                        PR_CalcPolar xyCutOutArcCentre, nGussetAngle, g_nCutOutRadius, xyStrap1
                        g_nLengthStrap1ToCutOut5 = nCutArcLength - nGussetLength
                End If
            End If

            PR_MakeXY xyStrap2, xyStrap1.X, xyStrap1.y - 1.75   '1-3/4"
            nGussetRadius = xyCutOutArcCentre.y - xyStrap2.y
            PR_MakeXY xyGussetArcCentre, xyProfileBrief.X + nGussetRadius, xyProfileBrief.y
            
            If xyStrap1.X > xyGussetArcCentre.X Then
               'Normal end of Snap crotch arc
                PR_MakeXY xyStrap3, xyGussetArcCentre.X, xyStrap2.y  'Roughly!!!
            Else
               'End of arc
                PR_MakeXY xyStrap3, xyStrap1.X - INCH3_4, xyStrap2.y
               'Revise xyGussetArcCentre
'                aAngle = Fn_CalcAngle(xyStrap3, xyProfileBrief) - 90
'                nLength = FN_CalcLength(xyStrap3, xyProfileBrief) / 2
'                Pr_CalcMidPoint xyStrap3, xyProfileBrief, xyTmp1
'                nLength = Sqr(nGussetRadius ^ 2 - nLength ^ 2)
'                PR_CalcPolar xyTmp1, aAngle, nLength, xyGussetArcCentre
                aAngle = 180 - FN_CalcAngle(xyStrap3, xyProfileBrief)
                nLength = FN_CalcLength(xyStrap3, xyProfileBrief) / 2
                nGussetRadius = nLength / Cos(aAngle * (PI / 180))
                PR_MakeXY xyGussetArcCentre, xyProfileBrief.X + nGussetRadius, xyProfileBrief.y
            End If
            PR_MakeXY xyStrap4, INCH3_4, xyStrap2.y
            PR_MakeXY xyLegPoint, INCH3_4, 0
            PR_MakeXY xyCrotchMarker, xyCutOut4.X, xyCutOut4.y

    End Select
    
   'Panty
   'NB We check at this point for complete leg data we can't do this in the
   'in the main validation routine as we do not have leg data
   'at that point
    If g_sLeftLeg = "Panty" Or g_sRightLeg = "Panty" Then

        If Not FN_ValidateLegData() Then End
        If Not FN_InitiliseLegRoutines() Then End
        
        If g_sLeftLeg = "Panty" Then PR_CalculateLeg "Left", g_nLeftLegLength
        If g_sRightLeg = "Panty" Then PR_CalculateLeg "Right", g_nRightLegLength
        If g_sLeftLeg = "Panty" And g_sRightLeg = "Panty" Then g_bDrawSingleLeg = FN_AverageLeftAndRightLegs()
        If g_sLeftLeg <> "Panty" And g_sRightLeg = "Panty" Then
            g_bDrawSingleLeg = True
            PR_CopyRightToLeftLeg
            xyLeftLegLabelPoint.X = xyLeftLegLabelPoint.X - g_nLeftLegLength 'This normally a side efect of averaging the legs'Bad practice but what the hell
        End If
        If g_sLeftLeg = "Panty" And g_sRightLeg <> "Panty" Then
            g_bDrawSingleLeg = True
            xyLeftLegLabelPoint.X = xyLeftLegLabelPoint.X - g_nLeftLegLength 'As above
        End If
        
        PR_MakeXY xyLT_Fold, xySeamFold.X, xySeamFold.y + g_nLT_HalfGroinHeight + Sqr((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2))
        PR_MakeXY xyFold, xySeamFold.X, xySeamFold.y + g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2))
    
    End If

    If g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief" Then
        'Check to see if it can reach, if not use the tangent from the groin
        'point and join them with a straight line.
        'Calculate points for Average / Smallest thigh
        nGroinAngle = FN_CalcAngle(xyButtocksArcCentre, xyProfileGroin)
        If (xyButtocksArcCentre.X - xySeamThigh.X) < g_nButtRadius Then
                'Curve reaches bottom of Brief
                g_bDrawBriefCurve = True
                PR_MakeXY xyTmp1, xySeamThigh.X, xyButtocksArcCentre.y
                PR_MakeXY xyTmp2, xyTmp1.X, xyProfileButt.y
                iError = FN_CirLinInt(xyTmp1, xyTmp2, xyButtocksArcCentre, g_nButtRadius, xyProfileThigh)
                nThighAngle = FN_CalcAngle(xyButtocksArcCentre, xyProfileThigh)
                PR_CalcPolar xyButtocksArcCentre, nThighAngle - ((nThighAngle - nGroinAngle) / 3), g_nButtRadius, xyProfileThighExtraPT
            Else
                'Curve doesn't reaches bottom of Brief - draw tangential line
                g_bDrawBriefCurve = False
                nTangentAngle = nGroinAngle + 90
                PR_MakeXY xyTmp1, xyProfileBrief.X, xyProfileButt.y
                PR_CalcPolar xyProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2
                iError = FN_LinLinInt(xyProfileBrief, xyTmp1, xyProfileGroin, xyTmp2, xyProfileThigh)
                nGroinToThighAngle = FN_CalcAngle(xyProfileGroin, xyProfileThigh)
                nGroinToThighMidPoint = (FN_CalcLength(xyProfileGroin, xyProfileThigh) / 2)
                PR_CalcPolar xyProfileGroin, nGroinToThighAngle, nGroinToThighMidPoint, xyProfileThighExtraPT
        End If
    
        'Calculate points for largest thigh (if required)
        'I know parts of this is identical to above except for a few variable changes but I
        'am trying to retrofit this code without making major changes to code that has
        'already been validated. (Yeah! I know it sound feeble, but this is not paying me enough to
        'really care) GG
        If g_bDiffThigh Then
            nGroinAngle = FN_CalcAngle(xyLT_ButtocksArcCentre, xyLT_ProfileGroin)
            If (xyLT_ButtocksArcCentre.X - xySeamThigh.X) < g_nLT_ButtRadius Then
                    'Curve reaches bottom of Brief
                    PR_MakeXY xyTmp1, xySeamThigh.X, xyLT_ButtocksArcCentre.y
                    PR_MakeXY xyTmp2, xyTmp1.X, xyProfileButt.y + 2 'Ensure it clears
                    iError = FN_CirLinInt(xyTmp1, xyTmp2, xyLT_ButtocksArcCentre, g_nLT_ButtRadius, xyLT_ProfileThigh)
                Else
                    'Curve doesn't reaches bottom of Brief - draw tangential line
                    nTangentAngle = nGroinAngle + 90
                    PR_MakeXY xyTmp1, xyProfileBrief.X, xyProfileButt.y + 2
                    PR_CalcPolar xyLT_ProfileGroin, nTangentAngle, g_nShldToFold, xyTmp2
                    iError = FN_LinLinInt(xyProfileBrief, xyTmp1, xyLT_ProfileGroin, xyTmp2, xyLT_ProfileThigh)
            End If
            PR_MakeXY xyLT_Fold, xySeamFold.X, xySeamFold.y + g_nLT_HalfGroinHeight + Sqr((g_nLT_ButtRadius ^ 2) - (g_nButtocksLength ^ 2))
            'Revise xyLT_ButtocksArcCentre so the point can be used construct the back arc
            'Note that we are trying to blend from the largest thigh into the body based on
            'the smallest thigh
            aAngle = FN_CalcAngle(xyLT_Fold, xyProfileButt) - 90  'perpendicular angle to the line xyLT_Fold, xyProfileButt
            nLength = Sqr(g_nButtRadius ^ 2 - (FN_CalcLength(xyLT_Fold, xyProfileButt) / 2) ^ 2)
            PR_CalcMidPoint xyLT_Fold, xyProfileButt, xyTmp1
            PR_CalcPolar xyTmp1, aAngle, nLength, xyLT_ButtocksArcCentre
        End If
    End If

   'Average/Smallest thigh
    PR_MakeXY xyFold, xySeamFold.X, xySeamFold.y + g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - (g_nButtocksLength ^ 2))
    PR_MakeXY xy85Fold, xySeamFold.X, xySeamFold.y + g_n85FoldHeight
    PR_MakeXY xyProfileThighMirror, xyProfileThigh.X, -xyProfileThigh.y
   ' If g_bDiffThigh Then PR_MakeXY xyLT_ProfileThighMirror, xyLT_ProfileThigh.x, -xyLT_ProfileThigh.y
    PR_MakeXY xyLT_ProfileThighMirror, xyLT_ProfileThigh.X, -xyLT_ProfileThigh.y
    PR_MakeXY xyProfileThighMirror1, xyProfileThighMirror.X, xyProfileThighMirror.y + INCH1_4

   'largest thigh (if required)

End Sub

Private Static Sub PR_CreateDrawMacro()
    Dim sText As String
    Dim xyText As XY
    Dim xyZipMark(1 To 4)  As XY
    Dim xyTmp As XY
    Dim xyTmp1 As XY
    Dim aAngle As Double
    Dim aInc As Double
    Dim ii As Integer
    Dim nLength As Double
    Dim nOffset As Double
    Dim xyOpen(1 To 4)  As XY
    Dim aOpen(1 To 4)  As Double
    
    Dim sAxillaText As String
    Dim sAxillaTextLow As String

    Dim sAxillaTextThis  As String
    Dim sThis  As String
    Dim sOther As String

    Dim nVertLinesOffset As Double
    Dim nHorLinesOffset As Integer
    Dim nSmallestLegXOffset As Double
    Dim nLargestLegXOffset As Double
    Dim xySeamLineStart As XY
    Dim xySeamLineEnd As XY
    Dim xyCentreLineStart As XY
    Dim xyCentreLineEnd As XY
    Dim xyBackNeckMid1 As XY
    Dim xyBackNeckMid2 As XY
    Dim xyFrontNeckMid1 As XY
    Dim xyFrontNeckMid2 As XY
    Dim xyGroinExtraPT As XY
    Dim xyGroinExtraPT1 As XY
    Dim xyTmpRaglanCentre As XY
    Dim nBackNeckMidOffsetX As Double
    Dim nBackNeckMidOffsetY As Double
    Dim nFlip As Integer
    Dim xyCrotchText As XY
    Dim sCrotchText As String
    Dim xyGussetOffCut As XY
    Dim xyLLGussetRectangle As XY
    Dim xyURGussetRectangle As XY
    Dim xyNeckText As XY
    Dim xyCutOutArcStart As XY
    Dim xyCutOutArcEnd As XY
    Dim xyProfileThighRelease As XY
    Dim xyLegStart As XY
    Dim xyMesh As XY
    
    

    Dim cProfilePoly As curve
    Dim cRaglanPoly As curve
    Dim cRaglanCurve As curve
    Dim cFrontCutOutPoly As curve
    Dim cBackCutOutPoly As curve
    Dim cFrontNeckArc As curve
    Dim cBackNeckArc As curve
    Dim sSmallestElasticText As String
    Dim sLargestElasticText As String

    Dim MeshProfile As BiArc
    Dim MeshSeam As BiArc
    Dim nMeshSize As Double
    Dim nMeshDistanceAlongRaglan As Double

    Dim SmallestLegCurve As curve
    Dim LargestLegCurve As curve


    Dim sLeg As String

    'Initialise
    fNum = FN_Open("C:\JOBST\DRAW.D")

    'Add labels and construction lines
    nVertLinesOffset = xyProfileButt.y + 1
    nHorLinesOffset = g_nShldToFold * 0.05
    PR_SetLayer "Construct"
    If Not g_sLegStyle = "Panty" Then PR_DrawVerticalLineThruPoint xySeamThigh, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamFold, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamButt, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamWaist, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamLowShld, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamHighShld, nVertLinesOffset
    PR_DrawVerticalLineThruPoint xySeamChest, nVertLinesOffset

    PR_MakeXY xySeamLineStart, 0 - (max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), INCH1_4
    PR_MakeXY xySeamLineEnd, xySeamHighShld.X + nHorLinesOffset, INCH1_4
    PR_MakeXY xyCentreLineStart, 0 - (max(g_nRightLegLength, g_nLeftLegLength) + nHorLinesOffset), 0
    PR_MakeXY xyCentreLineEnd, xySeamHighShld.X + nHorLinesOffset, 0

    
    'Useful Points
    PR_MakeXY xyBackNeckMid1, xyCutOutBackNeck.X + ((xyProfileNeck.X - xyCutOutBackNeck.X) / 6), xyCutOutBackNeck.y + ((xyProfileNeck.y - xyCutOutBackNeck.y) / 3)
    PR_MakeXY xyBackNeckMid2, xyProfileNeck.X - ((xyProfileNeck.X - xyCutOutBackNeck.X) / 2), xyProfileNeck.y - ((xyProfileNeck.y - xyCutOutBackNeck.y) / 3)
    PR_MakeXY xyFrontNeckMid1, xyCutOutFrontNeck.X - ((xyCutOutFrontNeck.X - xyCutOut8.X) / 3), xyCutOutFrontNeck.y + ((xyCutOut8.y - xyCutOutFrontNeck.y) / 6)
    PR_MakeXY xyFrontNeckMid2, xyCutOut8.X + ((xyCutOutFrontNeck.X - xyCutOut8.X) / 3), xyCutOut8.y - ((xyCutOut8.y - xyCutOutFrontNeck.y) / 2)
    PR_CalcPolar xyButtocksArcCentre, 45, g_nButtRadius, xyCutOutArcStart
    PR_CalcPolar xyButtocksArcCentre, 135, g_nButtRadius, xyCutOutArcEnd
    PR_MakeXY xyGroinExtraPT, xySeamFold.X + (g_nButtocksLength / 2), xySeamFold.y + g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - ((g_nButtocksLength / 2) ^ 2))
    PR_MakeXY xyGroinExtraPT1, xySeamFold.X + ((g_nButtocksLength / 4) * 3), xySeamFold.y + g_nHalfGroinHeight + Sqr((g_nButtRadius ^ 2) - ((g_nButtocksLength / 4) ^ 2))
    
    'Calculated for manual editing
    PR_MakeXY xyProfileThighRelease, xySeamThigh.X, xySeamThigh.y + g_nThighCir   'this is the only place that g_nThighCir is used
    PR_DrawMarker xyProfileThighRelease
    PR_AddDBValueToLast "Data", "Thigh"
    PR_DrawMarker xy85Fold
    PR_AddDBValueToLast "Data", "At original 85% reduction"
   'Draw Construction Lines
    PR_DrawLine xySeamLineStart, xySeamLineEnd
    PR_DrawLine xyCentreLineStart, xyCentreLineEnd
    PR_AddDBValueToLast "curvetype", "CentreLine"
    'Draw Buttocks Arc for reference
    PR_DrawArc xyButtocksArcCentre, xyCutOutArcStart, xyCutOutArcEnd
    'Insert Key Markers for reference
    PR_DrawMarker xyProfileThigh
    PR_DrawMarker xyProfileGroin
    PR_DrawMarker xyProfileButt
    PR_DrawMarker xyProfileWaist
    'YES - 2 at same point (one with DB value)!!!!
    PR_DrawMarker xyProfileChest
    PR_DrawMarker xyProfileChest
    PR_AddDBValueToLast "curvetype", "BackCurveMarker"
    PR_DrawMarker xyProfileNeck
    PR_DrawMarker xyCutOutArcCentre
    PR_DrawMarker xyFrontNeckArcCentre
    PR_DrawMarker xyFold
    PR_DrawMarker xyButtocksArcCentre
    
    'Draw DB markers for Zipper macros
    PR_DrawMarker xyCutOutBackNeck
    PR_AddDBValueToLast "Zipper", "BackN"
    PR_DrawMarker xyCutOut2
    PR_AddDBValueToLast "Zipper", "BackW"
    PR_DrawMarker xyCutOut3
    PR_AddDBValueToLast "Zipper", "BackB"
    PR_DrawMarker xyCutOut5
    PR_AddDBValueToLast "Zipper", "FrontB"
    PR_AddDBValueToLast "Data", Str$(g_nLengthStrap1ToCutOut5)
    
    PR_DrawMarker xyCutOut6
    PR_AddDBValueToLast "Zipper", "FrontW"
    PR_DrawMarker xyCutOut8
    PR_AddDBValueToLast "Zipper", "FrontN"
    If g_sFrontNeckStyle = "Turtle" Then
        PR_AddDBValueToLast "Data", Str$(g_nFrontNeckSize)
    End If


    'Other patient details in black on layer construct
    PR_MakeXY xyText, xySeamButt.X + 1, xySeamButt.y - 1.8
    sText = g_sFileNo & "\n" & g_sDiagnosis & "\n" & Trim$(Str$(g_nAge)) & "\n" & g_sSex
    Pr_DrawText sText, xyText, 0.125
   
    'Notes Text - e.g. Patient Details, Sewing instructions
    PR_SetLayer "Notes"
    PR_MakeXY xyText, xySeamButt.X + 1, xySeamButt.y - 1
    sText = g_sPatient & "\n" & g_sWorkOrder & "\n" & Trim$(Mid$(g_sFabric, 4))
    Pr_DrawText sText, xyText, 0.125
    

    
   'Draw Measured Front Neck Text
    If InStr(g_sFrontNeckStyle, "Scoop") > 0 Then
            sText = "SCOOP NECK"
            PR_MakeXY xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25
        ElseIf InStr(g_sFrontNeckStyle, "V neck") > 0 Then
            sText = "V NECK"
            PR_MakeXY xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25
        ElseIf InStr(g_sFrontNeckStyle, "Turtle") > 0 Then
            sText = "NECK CIRC. " & fnInchestoText(g_nNeckCir) & "\""" & "\nTURTLE NECK WIDTH " & fnInchestoText(g_nFrontNeckSize) & "\"""
            PR_MakeXY xyNeckText, xyCutOutFrontNeck.X - 3.5, xyCutOutFrontNeck.y + 0.5
            If InStr(g_sFrontNeckStyle, "same") Then
                If g_sLegStyle = "Brief" Then
                    sText = sText & "\nIN SAME FABRIC AS BRIEF"
                Else
                    sText = sText & "\nIN SAME FABRIC AS SUIT"
                End If
            ElseIf InStr(g_sFrontNeckStyle, "Detach. Fabric") Then
                sText = sText & "\nDETACHABLE FABRIC"
            Else
                sText = sText & "\nDETACHABLE"
            End If

        Else    'Regular
            sText = "NECK CIRC. " & fnInchestoText(g_nNeckCir) & "\"""
            PR_MakeXY xyNeckText, xyCutOutFrontNeck.X - 3, xyCutOutFrontNeck.y + 0.25
    End If
    If InStr(g_sFrontNeckStyle, "Measured") > 0 Then sText = "MEASURED " & sText
    Pr_DrawText sText, xyNeckText, 0.125
    
   'Draw Measured Back Neck Text
    If InStr(g_sBackNeckStyle, "Scoop") > 0 Then
        If g_sBackNeckStyle = "Measured Scoop" Then
            sText = "MEASURED SCOOP NECK"
        Else
            sText = "SCOOP NECK"
        End If
        PR_MakeXY xyNeckText, xyCutOutBackNeck.X - 3, xyCutOutBackNeck.y + 0.25
        Pr_DrawText sText, xyNeckText, 0.125
    End If
    
    'Draw Zipper
     If g_sClosure = "Back Zip" Then
       'Back zipper
        PR_CalcMidPoint xyCutOut3, xyCutOut2, xyZipMark(1)
        nLength = FN_CalcLength(xyCutOut2, xyCutOutBackNeck) + FN_CalcLength(xyCutOut3, xyZipMark(1))
        aAngle = FN_CalcAngle(xyCutOut3, xyZipMark(1))
        PR_MakeXY xyZipMark(2), xyZipMark(1).X, xyZipMark(1).y + 0.125
        PR_CalcPolar xyZipMark(1), aAngle, 0.5, xyZipMark(3)
        PR_CalcMidPoint xyCutOut2, xyCutOutBackNeck, xyZipMark(4) 'Text insertion point
        xyZipMark(4).y = xyZipMark(4).y + 0.125
     End If

     If g_sClosure = "Front Zip" Then
       'Front Zipper
        If g_bMissCutOut7 Then
            nLength = FN_CalcLength(xyCutOut8, xyCutOut6)
            PR_CalcMidPoint xyCutOut6, xyCutOut8, xyZipMark(4) 'Text insertion point
        Else
            nLength = FN_CalcLength(xyCutOut8, xyCutOut7) + FN_CalcLength(xyCutOut7, xyCutOut6)
            PR_CalcMidPoint xyCutOut6, xyCutOut7, xyZipMark(4) 'Text insertion point
        End If
        nLength = nLength + FN_CalcLength(xyCutOut5, xyCutOut6)

        xyZipMark(4).y = xyZipMark(4).y - 0.25

        aAngle = FN_CalcAngle(xyCutOut5, xyCutOut6)
        If g_sCrotchStyle = "Snap Crotch" Then
            If g_nLengthStrap1ToCutOut5 > 0 Then
               'This takes care of the special case where the snap crotch lands
               'before xyCutOut5
                PR_CalcPolar xyCutOut5, aAngle, 1.125 - g_nLengthStrap1ToCutOut5, xyZipMark(1)
                nLength = nLength - FN_CalcLength(xyCutOut5, xyZipMark(1))
            Else
               'We also need to take care of another special case
               'where the 1.25" along the cut-out to the start of the
               'zip mark takes it around the corner (so to speak)
                Dim nLen As Double
                nLen = FN_CalcLength(xyStrap1, xyCutOut6)
                If nLen >= 1.625 Then
                   'Simple case Zip mark is on the line (xyCutOut5, xyCutOut6)
                    PR_CalcPolar xyStrap1, aAngle, 1.125, xyZipMark(1)
                    nLength = nLength - FN_CalcLength(xyCutOut5, xyZipMark(1))
                ElseIf nLen > 1.125 And nLen < 1.625 Then
                   'Zip mark is round the corner
                    PR_CalcPolar xyStrap1, aAngle, 1.125, xyZipMark(1)
                    nLength = nLength - FN_CalcLength(xyCutOut5, xyZipMark(1))
                    aAngle = FN_CalcAngle(xyCutOut6, xyCutOut7)
                Else
                   'Zip mark is on the line (xyCutOut6, xyCutOut7)
                    nLen = 1.125 - nLen
                    aAngle = FN_CalcAngle(xyCutOut6, xyCutOut7)
                    PR_CalcPolar xyCutOut6, aAngle, nLen, xyZipMark(1)
                    nLength = nLength - FN_CalcLength(xyCutOut5, xyCutOut6) - nLen
                End If
            End If
        Else
            PR_CalcMidPoint xyCutOut5, xyCutOut6, xyZipMark(1)
            nLength = nLength - (FN_CalcLength(xyCutOut5, xyCutOut6) / 2)
        End If
        PR_MakeXY xyZipMark(2), xyZipMark(1).X, xyZipMark(1).y - 0.125
        PR_CalcPolar xyZipMark(1), aAngle, 0.5, xyZipMark(3)

     End If

    'Draw the zip and text
     If g_sClosure = "Back Zip" Or g_sClosure = "Front Zip" Then
         PR_DrawLine xyZipMark(1), xyZipMark(2)
         PR_AddDBValueToLast "Zipper", "1"
         PR_DrawLine xyZipMark(2), xyZipMark(3)
         PR_AddDBValueToLast "Zipper", "1"
         PR_InsertSymbol "TextAsSymbol", xyZipMark(4), 1, 0
        'Adjust length for turtle necks
         If g_sFrontNeckStyle = "Turtle" Then
             nLength = nLength + g_nFrontNeckSize
         End If
         PR_AddDBValueToLast "Zipper", "1"
         nLength = (nLength - 0.125) / 0.95
         sText = fnInchestoText(nLength) & "\"" " & g_sClosure
         PR_AddDBValueToLast "Data", sText
         PR_AddFormatedDBValueToLast "Data", "", nLength, " " & g_sClosure
     End If

    'Bracups
     If g_sSex = "Female" Then PR_DrawBra

    'Draw Body Suit/Brief
     PR_SetLayer "Template" & g_sSide
    
    'Draw Front Neck
    'Temporary 3 point polyline - will use arc for all options later
     If g_sFrontNeckStyle = "Measured Scoop" Or g_sFrontNeckStyle = "Measured V neck" Then
            cFrontNeckArc.X(1) = xyCutOutFrontNeck.X
            cFrontNeckArc.y(1) = xyCutOutFrontNeck.y
            cFrontNeckArc.X(2) = xyFrontNeckMid1.X
            cFrontNeckArc.y(2) = xyFrontNeckMid1.y
            cFrontNeckArc.X(3) = xyFrontNeckMid2.X
            cFrontNeckArc.y(3) = xyFrontNeckMid2.y
            cFrontNeckArc.X(4) = xyCutOut8.X
            cFrontNeckArc.y(4) = xyCutOut8.y
            cFrontNeckArc.n = 4
            PR_DrawFitted cFrontNeckArc
        Else
            PR_DrawArc xyFrontNeckArcCentre, xyCutOutFrontNeck, xyCutOut8
    End If
    
    'Draw Cutout
    If g_bMissCutOut7 Then
            cFrontCutOutPoly.X(1) = xyCutOut5.X
            cFrontCutOutPoly.y(1) = xyCutOut5.y
            cFrontCutOutPoly.X(2) = xyCutOut6.X
            cFrontCutOutPoly.y(2) = xyCutOut6.y
            cFrontCutOutPoly.X(3) = xyCutOut8.X
            cFrontCutOutPoly.y(3) = xyCutOut8.y
            cFrontCutOutPoly.n = 3
            PR_DrawPoly cFrontCutOutPoly
            PR_AddDBValueToLast "curvetype", "CutOutCurve"
            PR_AddDBValueToLast "Zipper", "Front"
            PR_AddDBValueToLast "ZipperLength", "%Length"
            'Draw DB markers for Zipper macros
            PR_DrawMarker xyCutOut8 'Position this Marker at the neck for simplicity
            PR_AddDBValueToLast "Zipper", "FrontC"
        Else
            cFrontCutOutPoly.X(1) = xyCutOut5.X
            cFrontCutOutPoly.y(1) = xyCutOut5.y
            cFrontCutOutPoly.X(2) = xyCutOut6.X
            cFrontCutOutPoly.y(2) = xyCutOut6.y
            cFrontCutOutPoly.X(3) = xyCutOut7.X
            cFrontCutOutPoly.y(3) = xyCutOut7.y
            cFrontCutOutPoly.X(4) = xyCutOut8.X
            cFrontCutOutPoly.y(4) = xyCutOut8.y
            cFrontCutOutPoly.n = 4
            PR_DrawPoly cFrontCutOutPoly
            PR_AddDBValueToLast "curvetype", "CutOutCurve"
            PR_AddDBValueToLast "Zipper", "Front"
            PR_AddDBValueToLast "ZipperLength", "%Length"
            PR_DrawMarker xyCutOut7
            PR_DrawMarker xyCutOut7
            PR_AddDBValueToLast "curvetype", "CutOutMarker"
            'Draw DB markers for Zipper macros
            PR_DrawMarker xyCutOut7
            PR_AddDBValueToLast "Zipper", "FrontC"
    End If
    
    cBackCutOutPoly.X(1) = xyCutOut3.X
    cBackCutOutPoly.y(1) = xyCutOut3.y
    cBackCutOutPoly.X(2) = xyCutOut2.X
    cBackCutOutPoly.y(2) = xyCutOut2.y
    cBackCutOutPoly.X(3) = xyCutOutBackNeck.X
    cBackCutOutPoly.y(3) = xyCutOutBackNeck.y
    cBackCutOutPoly.n = 3
    PR_DrawFitted cBackCutOutPoly
    PR_AddDBValueToLast "Zipper", "Back"
    PR_AddDBValueToLast "ZipperLength", "%Length"
    
    'Draw Back Neck
    cBackNeckArc.X(1) = xyCutOutBackNeck.X
    cBackNeckArc.y(1) = xyCutOutBackNeck.y
    cBackNeckArc.X(2) = xyBackNeckMid1.X
    cBackNeckArc.y(2) = xyBackNeckMid1.y
    cBackNeckArc.X(3) = xyBackNeckMid2.X
    cBackNeckArc.y(3) = xyBackNeckMid2.y
    cBackNeckArc.X(4) = xyProfileNeck.X
    cBackNeckArc.y(4) = xyProfileNeck.y
    cBackNeckArc.n = 4
    PR_DrawFitted cBackNeckArc
    
    If InStr(g_sLegStyle, "Brief") > 0 Then
       'Draw back profile
        cProfilePoly.X(1) = xyProfileNeck.X
        cProfilePoly.y(1) = xyProfileNeck.y
        cProfilePoly.X(2) = xyProfileChest.X
        cProfilePoly.y(2) = xyProfileChest.y
        cProfilePoly.X(3) = xyProfileWaist.X
        cProfilePoly.y(3) = xyProfileWaist.y
        cProfilePoly.X(4) = xyProfileButt.X
        cProfilePoly.y(4) = xyProfileButt.y
        cProfilePoly.X(5) = xyGroinExtraPT1.X
        cProfilePoly.y(5) = xyGroinExtraPT1.y
        cProfilePoly.X(6) = xyGroinExtraPT.X
        cProfilePoly.y(6) = xyGroinExtraPT.y
        cProfilePoly.X(7) = xyProfileGroin.X
        cProfilePoly.y(7) = xyProfileGroin.y
    
        If g_bDrawBriefCurve Then
                cProfilePoly.X(8) = xyFold.X
                cProfilePoly.y(8) = xyFold.y
                cProfilePoly.X(9) = xyProfileThighExtraPT.X
                cProfilePoly.y(9) = xyProfileThighExtraPT.y
                cProfilePoly.X(10) = xyProfileThigh.X
                cProfilePoly.y(10) = xyProfileThigh.y
                cProfilePoly.n = 10
            Else
                cProfilePoly.X(8) = xyProfileThighExtraPT.X
                cProfilePoly.y(8) = xyProfileThighExtraPT.y
                cProfilePoly.X(9) = xyProfileThigh.X
                cProfilePoly.y(9) = xyProfileThigh.y
                cProfilePoly.n = 9
        End If
        PR_DrawFitted cProfilePoly
        PR_AddDBValueToLast "curvetype", "BackCurve"
        PR_AddDBValueToLast "Leg", "Left&Right"
    
        If g_bDiffThigh Then
           'Revise data above
            PR_AddDBValueToLast "Leg", g_sSmallestThighGiven
           'Draw the largest thigh profile
         
            aAngle = FN_CalcAngle(xyLT_ButtocksArcCentre, xyProfileButt)
            aInc = (aAngle - FN_CalcAngle(xyLT_ButtocksArcCentre, xyLT_Fold)) / 4
            cProfilePoly.n = 0
            For ii = 1 To 5
                PR_CalcPolar xyLT_ButtocksArcCentre, aAngle, g_nButtRadius, xyTmp
                cProfilePoly.n = cProfilePoly.n + 1
                cProfilePoly.X(cProfilePoly.n) = xyTmp.X
                cProfilePoly.y(cProfilePoly.n) = xyTmp.y
                aAngle = aAngle - aInc
            Next ii
            cProfilePoly.n = cProfilePoly.n + 1
            cProfilePoly.X(cProfilePoly.n) = xyLT_ProfileThigh.X
            cProfilePoly.y(cProfilePoly.n) = xyLT_ProfileThigh.y
    
            PR_SetLayer "Template" & g_sLargestThighGiven
            PR_DrawFitted cProfilePoly
            PR_AddDBValueToLast "curvetype", "BackCurveLargest"
            PR_AddDBValueToLast "Leg", g_sLargestThighGiven
    
           'Label Left and Right
            PR_SetLayer "Notes"
           'Insert ArrowHead symbol
            PR_DrawMarkerNamed "Closed Arrow", xyLT_Fold, INCH1_4, INCH1_8, -90
            PR_DrawMarkerNamed "Closed Arrow", xyProfileThighExtraPT, INCH1_4, INCH1_8, -90
            PR_SetTextData HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT
            PR_MakeXY xyText, xyProfileThighExtraPT.X, xyProfileThighExtraPT.y - INCH1_8
            Pr_DrawText g_sSmallestThighGiven, xyText, 0.1
            PR_MakeXY xyText, xyLT_Fold.X, xyLT_Fold.y - INCH1_8
            Pr_DrawText g_sLargestThighGiven, xyText, 0.1
    
           'Reset back to original
            PR_SetLayer "Template" & g_sSmallestThighGiven
    
        End If
    End If

    If g_sLegStyle = "Panty" Then

        PR_MakeXY xyLegStart, -max(g_nRightLegLength, g_nLeftLegLength), 0
        PR_SetLayer "Construct"
        If g_bDrawSingleLeg Then
            PR_DrawLegTemplate "Left", xyLegStart
            SmallestLegCurve = g_LeftLegProfile
            nSmallestLegXOffset = g_nLeftLegLength
            If g_bLeftAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
        Else
            PR_DrawLegTemplate "Both", xyLegStart
            If g_sSmallestThighGiven = "Left" Then
                SmallestLegCurve = g_LeftLegProfile
                LargestLegCurve = g_RightLegProfile
                nLargestLegXOffset = g_nRightLegLength
                nSmallestLegXOffset = g_nLeftLegLength
                If g_bLeftAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
                If g_bRightAboveKnee Then sLargestElasticText = "ELASTIC" Else sLargestElasticText = "NO ELASTIC"
            Else
                SmallestLegCurve = g_RightLegProfile
                LargestLegCurve = g_LeftLegProfile
                nLargestLegXOffset = g_nLeftLegLength
                nSmallestLegXOffset = g_nRightLegLength
                If g_bLeftAboveKnee Then sLargestElasticText = "ELASTIC" Else sLargestElasticText = "NO ELASTIC"
                If g_bRightAboveKnee Then sSmallestElasticText = "ELASTIC" Else sSmallestElasticText = "NO ELASTIC"
            End If

        End If
        cProfilePoly.X(1) = xyProfileNeck.X
        cProfilePoly.y(1) = xyProfileNeck.y
        cProfilePoly.X(2) = xyProfileChest.X
        cProfilePoly.y(2) = xyProfileChest.y
        cProfilePoly.X(3) = xyProfileWaist.X
        cProfilePoly.y(3) = xyProfileWaist.y
        cProfilePoly.X(4) = xyProfileButt.X
        cProfilePoly.y(4) = xyProfileButt.y
        cProfilePoly.X(5) = xyGroinExtraPT1.X
        cProfilePoly.y(5) = xyGroinExtraPT1.y
        cProfilePoly.X(6) = xyGroinExtraPT.X
        cProfilePoly.y(6) = xyGroinExtraPT.y
       ' cProfilePoly.x(7) = xyProfileGroin.x
       ' cProfilePoly.y(7) = xyProfileGroin.y
        cProfilePoly.n = 6
    
       'Join leg
        For ii = SmallestLegCurve.n To 1 Step -1
            cProfilePoly.n = cProfilePoly.n + 1
            cProfilePoly.X(cProfilePoly.n) = SmallestLegCurve.X(ii) - nSmallestLegXOffset
            cProfilePoly.y(cProfilePoly.n) = SmallestLegCurve.y(ii)
        Next ii
        If g_bDrawSingleLeg Then PR_SetLayer "TemplateLeft" Else PR_SetLayer "Template" & g_sSmallestThighGiven
        PR_DrawFitted cProfilePoly

        PR_AddDBValueToLast "curvetype", "BackCurve"

       'Label legs
        PR_SetLayer "Notes"
        PR_SetTextData HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT
        If Not g_bDrawSingleLeg Then
            PR_AddDBValueToLast "Leg", g_sSmallestThighGiven
        Else
            PR_AddDBValueToLast "Leg", "Left&Right"
        End If

       'Draw closure
        If g_bDrawSingleLeg Then PR_SetLayer "TemplateLeft" Else PR_SetLayer "Template" & g_sSmallestThighGiven
        PR_MakeXY xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n)
        PR_MakeXY xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n)
        PR_DrawLine xyTmp, xyTmp1

       'Add elastic text for smallest
        PR_SetLayer "Notes"
        g_nCurrTextAngle = 270
        PR_MakeXY xyText, xyTmp1.X + 0.3, 0
        Pr_DrawText sSmallestElasticText, xyText, 0.125
        
       'Draw other leg
        If Not g_bDrawSingleLeg Then

           'Draw the largest thigh profile
            aAngle = FN_CalcAngle(xyButtocksArcCentre, xyProfileButt)
            aInc = (aAngle - FN_CalcAngle(xyButtocksArcCentre, xyGroinExtraPT)) / 4
            cProfilePoly.n = 0
            For ii = 1 To 5
                PR_CalcPolar xyButtocksArcCentre, aAngle, g_nButtRadius, xyTmp
                cProfilePoly.n = cProfilePoly.n + 1
                cProfilePoly.X(cProfilePoly.n) = xyTmp.X
                cProfilePoly.y(cProfilePoly.n) = xyTmp.y
                aAngle = aAngle - aInc
            Next ii

           'Join leg
            For ii = LargestLegCurve.n To 1 Step -1
                cProfilePoly.n = cProfilePoly.n + 1
                cProfilePoly.X(cProfilePoly.n) = LargestLegCurve.X(ii) - nLargestLegXOffset
                cProfilePoly.y(cProfilePoly.n) = LargestLegCurve.y(ii)
            Next ii
            
            PR_SetLayer "Template" & g_sLargestThighGiven
            PR_DrawFitted cProfilePoly
            PR_AddDBValueToLast "curvetype", "BackCurveLargest"
            PR_AddDBValueToLast "Leg", g_sLargestThighGiven
            PR_MakeXY xyTmp, cProfilePoly.X(cProfilePoly.n), cProfilePoly.y(cProfilePoly.n)
            PR_MakeXY xyTmp1, cProfilePoly.X(cProfilePoly.n), -cProfilePoly.y(cProfilePoly.n)
            PR_DrawLine xyTmp, xyTmp1

           'Add elastic text
            PR_SetLayer "Notes"
            PR_MakeXY xyText, xyTmp1.X + 0.3, 0
            Pr_DrawText sLargestElasticText, xyText, 0.125
        End If
    End If

    
    'Setup which side is being drawn
    If g_sSide = "Left" Then
            g_sSleeveType = g_sLeftSleeve
        Else
            g_sSleeveType = g_sRightSleeve
    End If
    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT

    'Draw Axilla
    If g_sSleeveType = "Sleeveless" Then
            'Draw Sleeveless Axilla
            cRaglanPoly.X(1) = xyCutOutFrontNeck.X
            cRaglanPoly.y(1) = xyCutOutFrontNeck.y
            cRaglanPoly.X(2) = xyRaglan1.X
            cRaglanPoly.y(2) = xyRaglan1.y
            cRaglanPoly.X(3) = xyRaglan2.X
            cRaglanPoly.y(3) = xyRaglan2.y
            cRaglanPoly.n = 3
            PR_DrawPoly cRaglanPoly
            cRaglanPoly.X(1) = xyRaglan5.X
            cRaglanPoly.y(1) = xyRaglan5.y
            cRaglanPoly.X(2) = xyRaglan6.X
            cRaglanPoly.y(2) = xyRaglan6.y
            cRaglanPoly.X(3) = xyProfileNeckMirror.X
            cRaglanPoly.y(3) = xyProfileNeckMirror.y
            cRaglanPoly.n = 3
            PR_DrawPoly cRaglanPoly

           'Draw Raglan Curves
            PR_SetLayer "Template" & g_sAxillaSide
            nFlip = True
            PR_CalcRaglan xyRaglan4, xyRaglan2, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), -1
            PR_DrawFitted cRaglanCurve
            nFlip = False
            PR_CalcRaglan xyRaglan4, xyRaglan5, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1
            PR_DrawFitted cRaglanCurve
            If g_bDiffAxillaHeight Then
                PR_SetLayer "Template" & g_sAxillaSideLow
                nFlip = False
                PR_CalcRaglan xyRaglan4LowAxilla, xyRaglan5, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1
                PR_DrawFitted cRaglanCurve
                nFlip = True
                PR_CalcRaglan xyRaglan4LowAxilla, xyRaglan2, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), -1
                PR_DrawFitted cRaglanCurve

                PR_SetLayer "Notes"
                PR_DrawMarkerNamed "Closed Arrow", xyRaglan4LowAxilla, INCH1_4, INCH1_8, 180
                g_nCurrTextAngle = 270
                PR_MakeXY xyText, xyRaglan4LowAxilla.X - 0.3, xyRaglan4LowAxilla.y
                Pr_DrawText g_sAxillaSideLow, xyText, 0.125
                PR_DrawMarkerNamed "Closed Arrow", xyRaglan4, INCH1_4, INCH1_8, 180
                PR_MakeXY xyText, xyRaglan4.X - 0.3, xyRaglan4.y
                Pr_DrawText g_sAxillaSide, xyText, 0.125
                g_nCurrTextAngle = 0
            End If
        Else
           'Draw Sleeved Axilla
            PR_DrawLine xyCutOutFrontNeck, xyRaglan1
            PR_DrawLine xyRaglan3, xyProfileNeckMirror
 
           'Draw Raglan Curves
            PR_SetLayer "Template" & g_sAxillaSide
            nFlip = False
            PR_CalcRaglan xyRaglan2, xyRaglan3, nFlip, cRaglanCurve, xyOpen(2), aOpen(2), -1
            PR_DrawFitted cRaglanCurve
            nFlip = True
            nLength = FN_CalcLength(xyRaglan2, xyOpen(2))
            PR_CalcRaglan xyRaglan2, xyRaglan1, nFlip, cRaglanCurve, xyOpen(1), aOpen(1), nLength
            PR_DrawFitted cRaglanCurve
            
            Select Case g_sAxillaType
             Case "Open"
                PR_SetLayer "Notes"
                PR_DrawMarkerNamed "Closed Arrow", xyOpen(1), INCH1_4, INCH1_8, aOpen(1)
                PR_DrawMarkerNamed "Closed Arrow", xyOpen(2), INCH1_4, INCH1_8, aOpen(2)
                sAxillaText = "Open"
    
             Case "Mesh"
                PR_DrawMeshConstruction xyRaglan2, xyMesh, sAxillaText, nMeshSize
                nFlip = False
                nMeshDistanceAlongRaglan = 0
                If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                    MsgBox "Unable to calculate mesh along back neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                Else
                    PR_DrawMesh MeshProfile, MeshSeam, g_sAxillaSide
                End If
               'Calculate the mesh seam line along back neck raglan
               'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                nFlip = True
                If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                    MsgBox "Unable to calculate mesh along front neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                Else
                    PR_DrawMesh MeshProfile, MeshSeam, g_sAxillaSide
                End If
                sText = Trim$(Str$(nMeshDistanceAlongRaglan)) & "," & Trim$(Str$(nMeshSize))
                g_sMeshLeft = sText
                g_sMeshRight = sText

           'regular axilla (because CAD Block can't handle 1/2's)
            Case "Regular 1½"""
                    sAxillaText = "Regular 1-1/2\"""
            Case "Regular 2"""
                    sAxillaText = "Regular 2\"""
            Case "Regular 2½"""
                    sAxillaText = "Regular 2-1/2\"""

            Case Else
                    sAxillaText = g_sAxillaType
            End Select

           'Draw Lowest Axilla and Label
            g_nCurrTextAngle = 270
            If g_bDiffAxillaHeight Then
                PR_SetLayer "Template" & g_sAxillaSideLow
                nFlip = False
                PR_CalcRaglan xyRaglan2LowAxilla, xyRaglan3, nFlip, cRaglanCurve, xyOpen(4), aOpen(4), -1
                PR_DrawFitted cRaglanCurve
                nFlip = True
                nLength = FN_CalcLength(xyRaglan2LowAxilla, xyOpen(4))
                PR_CalcRaglan xyRaglan2LowAxilla, xyRaglan1, nFlip, cRaglanCurve, xyOpen(3), aOpen(3), nLength
                PR_DrawFitted cRaglanCurve

                Select Case g_sAxillaTypeLow
                 Case "Open"
                    PR_SetLayer "Notes"
                    PR_DrawMarkerNamed "Closed Arrow", xyOpen(3), INCH1_4, INCH1_8, aOpen(3)
                    PR_DrawMarkerNamed "Closed Arrow", xyOpen(4), INCH1_4, INCH1_8, aOpen(4)
                    sAxillaTextLow = "Open"
                 Case "Mesh"
                    PR_DrawMeshConstruction xyRaglan2LowAxilla, xyMesh, sAxillaTextLow, nMeshSize
                    nFlip = False
                    nMeshDistanceAlongRaglan = 0
                    If Not FN_CalcAxillaMesh(xyRaglan2LowAxilla, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                        MsgBox "Unable to calculate mesh along back neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                    Else
                        PR_DrawMesh MeshProfile, MeshSeam, g_sAxillaSideLow
                    End If
                   'Calculate the mesh seam line along back neck raglan
                   'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                    nFlip = True
                    If Not FN_CalcAxillaMesh(xyRaglan2LowAxilla, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                        MsgBox "Unable to calculate mesh along front neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                    Else
                        PR_DrawMesh MeshProfile, MeshSeam, g_sAxillaSideLow
                    End If
                    sText = Trim$(Str$(nMeshDistanceAlongRaglan)) & "," & Trim$(Str$(nMeshSize))
                    If g_sAxillaSideLow = "Left" Then g_sMeshLeft = sText Else g_sMeshRight = sText

    
                'regular axilla
                 Case "Regular 1½"""
                        sAxillaTextLow = "Regular 1-1/2\"""
                 Case "Regular 2"""
                        sAxillaTextLow = "Regular 2\"""
                 Case "Regular 2½"""
                        sAxillaTextLow = "Regular 2-1/2\"""
                 Case Else
                        sAxillaTextLow = g_sAxillaTypeLow
                         
                End Select

                PR_SetLayer "Notes"
                PR_DrawMarkerNamed "Closed Arrow", xyRaglan2LowAxilla, INCH1_4, INCH1_8, 180
                PR_MakeXY xyText, xyRaglan2LowAxilla.X - 0.3, xyRaglan2LowAxilla.y
                Pr_DrawText g_sAxillaSideLow & " " & sAxillaTextLow, xyText, 0.125

                PR_DrawMarkerNamed "Closed Arrow", xyRaglan2, INCH1_4, INCH1_8, 180
                PR_MakeXY xyText, xyRaglan2.X - 0.3, xyRaglan2.y
                Pr_DrawText g_sAxillaSide & " " & sAxillaText, xyText, 0.125

            Else
               'Take care of different axilla types at same axilla height
                If g_sLeftSleeve <> g_sRightSleeve Then
                   'Do the opposite axilla to that done above
                    If g_sLeftSleeve = g_sAxillaType Then
                        g_sAxillaType = g_sRightSleeve
                        sThis = "Right"
                        sOther = "Left"
                    Else
                        g_sAxillaType = g_sLeftSleeve
                        sThis = "Left"
                        sOther = "Right"
                    End If
                    Select Case g_sAxillaType
                     Case "Open"
                        PR_SetLayer "Notes"
                        PR_DrawMarkerNamed "Closed Arrow", xyOpen(1), INCH1_4, INCH1_8, aOpen(1)
                        PR_DrawMarkerNamed "Closed Arrow", xyOpen(2), INCH1_4, INCH1_8, aOpen(2)
                        sAxillaTextThis = "Open"
                     Case "Mesh"
                        PR_DrawMeshConstruction xyRaglan2, xyMesh, sAxillaText, nMeshSize
                        nFlip = False
                        nMeshDistanceAlongRaglan = 0
                        If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan3, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                            MsgBox "Unable to calculate mesh along back neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                        Else
                            PR_DrawMesh MeshProfile, MeshSeam, sThis
                        End If
                       'Calculate the mesh seam line along back neck raglan
                       'NB nMeshDistanceAlongRaglan is from  FN_CalcAxillaMesh
                        nFlip = True
                        If Not FN_CalcAxillaMesh(xyRaglan2, xyRaglan1, xyMesh, nFlip, nMeshSize, nMeshDistanceAlongRaglan, MeshProfile, MeshSeam) Then
                            MsgBox "Unable to calculate mesh along front neck raglan curve!" & Chr$(13) & "Draw manually and refer to supervisor."
                        Else
                            PR_DrawMesh MeshProfile, MeshSeam, sThis
                        End If

                        sText = Trim$(Str$(nMeshDistanceAlongRaglan)) & "," & Trim$(Str$(nMeshSize))
                        If sThis = "Left" Then g_sMeshLeft = sText Else g_sMeshRight = sText
                    
                    'regular axilla
                     Case "Regular 1½"""
                            sAxillaTextThis = "Regular 1-1/2\"""
                     Case "Regular 2"""
                            sAxillaTextThis = "Regular 2\"""
                     Case "Regular 2½"""
                            sAxillaTextThis = "Regular 2-1/2\"""
                     Case Else
                            sAxillaTextThis = g_sAxillaTypeLow
        
                    End Select

                    PR_SetLayer "Notes"
                    PR_MakeXY xyText, xyRaglan2.X - 0.3, xyRaglan2.y
                    Pr_DrawText sOther & " " & sAxillaText & "\n" & sThis & " " & sAxillaTextThis, xyText, 0.125

                Else
                    PR_SetLayer "Notes"
                    PR_MakeXY xyText, xyRaglan2.X - 0.3, xyRaglan2.y
                    Pr_DrawText sAxillaText, xyText, 0.125
                End If
            End If
            g_nCurrTextAngle = 0

           'Update the database for use by sleeve drawing module
            PR_UpdateDB
            
    End If
    
   'Draw Crotch (based on some fixed rules checked in BodySuit dialogue box)
    PR_MakeXY xyCrotchText, xyCrotchMarker.X - 0.375, xyCrotchMarker.y
    PR_DrawMarker xyCrotchMarker

   'Draw Crotch styles
    PR_SetLayer "Template" & "Left"

    Select Case g_sCrotchStyle
        Case "Open Crotch"
           'Draw Cutout Arc
            PR_DrawLine xyBackCrotch1, xyBackCrotch2
            PR_DrawLine xyFrontCrotch1, xyFrontCrotch2
            PR_DrawArc xyBackCrotchFilletCentre, xyBackCrotch2, xyBackCrotch3
            PR_DrawArc xyFrontCrotchFilletCentre, xyFrontCrotch3, xyFrontCrotch2

            If g_bExtremeCrotch Then
               'Extreme crotch
               'Back crotch
                If xyBackCrotch1.X < xyCutOut9.X Then
                    PR_DrawArc xyCutOutArcCentre, xyCutOut9, xyBackCrotch1
                    PR_DrawLine xyCutOut9, xyCutOut3
                ElseIf xyBackCrotch1.X < xyCutOut3.X Then
                    PR_DrawLine xyBackCrotch1, xyCutOut3
                End If
               'Reuse xyBackCrotch3 here as the start of the arc
                If xyBackCrotch3.X > xyCutOut9.X Then
                    xyTmp = xyCutOut9: xyTmp.y = xyBackCrotch3.y
                    PR_DrawLine xyBackCrotch3, xyTmp
                    xyBackCrotch3.X = xyCutOut9.X
                Else
                   'Do nothing
                End If
               'Front crotch
                If xyFrontCrotch1.X < xyCutOut10.X Then
                    PR_DrawArc xyCutOutArcCentre, xyCutOut10, xyFrontCrotch1
                    PR_DrawLine xyCutOut10, xyCutOut5
                ElseIf xyFrontCrotch1.X < xyCutOut5.X Then
                    PR_DrawLine xyFrontCrotch1, xyCutOut5
                End If
               'Reuse xyFrontCrotch3 here as the start of the arc
                If xyFrontCrotch3.X > xyCutOut10.X Then
                    xyTmp = xyCutOut10: xyTmp.y = xyFrontCrotch3.y
                    PR_DrawLine xyFrontCrotch3, xyTmp
                    xyFrontCrotch3.X = xyCutOut10.X
                Else
                   'Do nothing
                End If
               'Note:- Reuse of xyFrontCrotch3and xyBackCrotch3 here as the start and end of the arc
                PR_DrawArc xyCutOutArcCentre, xyBackCrotch3, xyFrontCrotch3


            Else
               'Normal crotch
                If xyBackCrotch1.X <> xyCutOut3.X And xyBackCrotch1.y <> xyCutOut3.y Then
                    PR_DrawArc xyCutOutArcCentre, xyCutOut3, xyBackCrotch1
                End If
                If xyFrontCrotch1.X <> xyCutOut5.X And xyFrontCrotch1.y <> xyCutOut5.y Then
                    PR_DrawArc xyCutOutArcCentre, xyFrontCrotch1, xyCutOut5
                End If
                PR_DrawArc xyCutOutArcCentre, xyBackCrotch3, xyFrontCrotch3
            End If
            
            If g_sSex = "Male" Then
                    sCrotchText = ""
                    If Not g_nAdult Then
                        sCrotchText = "1/2\"" ELASTIC"
                    End If
                Else
                    sCrotchText = "1/2\"" ELASTIC"
            End If

        Case "Horizontal Fly"
            sCrotchText = "Horizontal Fly," & g_sFlySize

        Case "Diagonal Fly"
            sCrotchText = "Diagonal Fly," & g_sFlySize
            PR_SetLayer "Notes"
            PR_DrawMarkerNamed "Closed Arrow", xyCutOut3, INCH1_4, INCH1_8, 90
            PR_DrawMarkerNamed "Closed Arrow", xyCutOut5, INCH1_4, INCH1_8, 270

        Case "Snap Crotch"

            sCrotchText = "GUSSET, MESH " & g_sGussetSize
            
           'DRAW Gusset rectangle
            PR_SetLayer "Template" & g_sSide
            PR_MakeXY xyLLGussetRectangle, xyProfileButt.X - 2, xyProfileButt.y + 2
            PR_MakeXY xyURGussetRectangle, xyLLGussetRectangle.X + (xyStrap2.X - xyStrap4.X), xyLLGussetRectangle.y + (xyStrap1.y - xyStrap2.y)
            PR_DrawRectangle xyLLGussetRectangle, xyURGussetRectangle
            
            PR_SetTextData LEFT_, TOP_, CURRENT, CURRENT, CURRENT
            g_nCurrTextAngle = 0
            PR_SetLayer "Notes"
            PR_MakeXY xyText, xyLLGussetRectangle.X + 0.25, xyURGussetRectangle.y - 0.25
            sText = g_sPatient & "\n" & g_sWorkOrder
            Pr_DrawText sText, xyText, 0.125

           'Draw DB markers for Zipper macros
            PR_SetLayer "Construct"
            PR_DrawMarker xyStrap1
            PR_AddDBValueToLast "Zipper", "Snap"

        Case "Gusset"
            sCrotchText = "Gusset, " & g_sGussetSize

    End Select

   'Draw Crotch Text and marker
    PR_SetLayer "Notes"
    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
    g_nCurrTextAngle = 270
    If sCrotchText <> "" Then Pr_DrawText sCrotchText, xyCrotchText, 0.125

   'Insert ArrowHead symbol
    PR_DrawMarkerNamed "Closed Arrow", xyCrotchMarker, INCH1_4, INCH1_8, 180

   'Draw Leg style
    PR_SetLayer "Template" & g_sSide

   'Draw CutOut Arc
    If Not g_sCrotchStyle = "Open Crotch" Then
        If g_bExtremeCrotch Then
            PR_DrawLine xyCutOut5, xyCutOut10
            PR_DrawLine xyCutOut9, xyCutOut3
            PR_AddEntityArc xyCutOutArcCentre, xyCutOut9, xyCutOut10
        Else
            PR_AddEntityArc xyCutOutArcCentre, xyCutOut3, xyCutOut5
        End If
    End If

    If g_sLegStyle = "Panty" Then
        g_nCurrTextAngle = 0
        PR_SetTextData HORIZ_CENTER, TOP_, CURRENT, CURRENT, CURRENT
        PR_SetLayer "Notes"
        If (g_sLeftLeg = "Brief" Or g_sRightLeg = "Brief") Then
           'There is only a single leg draw which is given by the left leg
            PR_MakeXY xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8
            PR_DrawMarkerNamed "Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90
            If g_sLeftLeg = "Brief" Then
                Pr_DrawText "Right", xyText, 0.125
                PR_SetLayer "TemplateLeft"
            Else
                Pr_DrawText "Left", xyText, 0.125
                PR_SetLayer "TemplateRight"
            End If
            PR_DrawBriefStandard 0.875
        Else
            If g_bDrawSingleLeg Then
                PR_MakeXY xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8
                PR_DrawMarkerNamed "Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90
                Pr_DrawText "Left & Right", xyText, 0.125
            Else
                PR_MakeXY xyText, xyLeftLegLabelPoint.X, xyLeftLegLabelPoint.y - INCH1_8
                PR_DrawMarkerNamed "Closed Arrow", xyLeftLegLabelPoint, INCH1_4, INCH1_8, -90
                Pr_DrawText "Left", xyText, 0.125
                PR_MakeXY xyText, xyRightLegLabelPoint.X, xyRightLegLabelPoint.y - INCH1_8
                PR_DrawMarkerNamed "Closed Arrow", xyRightLegLabelPoint, INCH1_4, INCH1_8, -90
                Pr_DrawText "Right", xyText, 0.125
            End If
        End If
    End If

   'Draw on body layer
    PR_SetLayer "Template" & g_sSide

    If g_sLegStyle = "Brief" Then
        PR_DrawBriefStandard 0
    End If
    
    If g_sLegStyle = "Brief-French" Then
       'calculate additional French-cut points
        If g_sSleeveType = "Sleeveless" Then xyFrenchCut.y = xyRaglan4.y Else xyFrenchCut.y = xyRaglan2.y
        If g_nAdult Then xyFrenchCut.X = xySeamFold.X + 2.5 Else xyFrenchCut.X = xySeamFold.X + 2
        PR_DrawBriefFrenchCut
    End If
                              

   'Draw Lower axilla construction line
    If g_bDiffAxillaHeight Then
        PR_SetLayer "Construct"
        PR_DrawVerticalLineThruPoint xySeamChestAxillaLow, nVertLinesOffset
    End If

   'Set Neutral Layer
    PR_SetLayer "1"

'PR_DebugPointText
   
    Close #fNum





End Sub

Private Sub PR_DebugPointText()
Dim nn#
    nn = 0.1
    PR_SetTextData RIGHT_, BOTTOM_, CURRENT, CURRENT, CURRENT
'    Pr_DrawText "xyCrotchMarker", xyCrotchMarker, nn
'    PR_DrawText "xyProfileThigh", xyProfileThigh, nn
'    PR_DrawText "xyProfileGroin", xyProfileGroin, nn
'    PR_DrawText "xyFold", xyFold, nn
'    PR_DrawText "xyLT_ProfileThigh", xyLT_ProfileThigh, nn
'    PR_DrawText "xyLT_ProfileGroin", xyLT_ProfileGroin, nn
'    PR_DrawText "xyLT_Fold", xyLT_Fold, nn
'    PR_DrawText "xyLT_ButtocksArcCentre", xyLT_ButtocksArcCentre, nn
'    PR_DrawText "xyProfileButt", xyProfileButt, nn
'    PR_DrawText "xyProfileBrief", xyProfileBrief, nn
    
'    Pr_DrawText "xyProfileWaist", xyProfileWaist, nn
'    Pr_DrawText "xyProfileChest", xyProfileChest, nn
'    Pr_DrawText "xyProfileNeck", xyProfileNeck, nn
'    PR_DrawText "xyCutOutArcCentre", xyCutOutArcCentre, nn
'    Pr_DrawText "xyFold", xyFold, nn
'    PR_DrawText "xyButtocksArcCentre", xyButtocksArcCentre, nn
'    Pr_DrawText "xyCutOutBackNeck", xyCutOutBackNeck, nn
'    Pr_DrawText "xyCutOut2", xyCutOut2, nn
    Pr_DrawText "xyCutOut3", xyCutOut3, nn
    Pr_DrawText "xyCutOut4", xyCutOut4, nn
'    Pr_DrawText "xyCutOut5", xyCutOut5, nn
'    Pr_DrawText "xyCutOut6", xyCutOut6, nn
'    Pr_DrawText "xyCutOut7", xyCutOut7, nn
 '   Pr_DrawText "xyCutOut8", xyCutOut8, nn
'
'    Pr_DrawText "xyBackCrotch1", xyBackCrotch1, nn
'    Pr_DrawText "xyBackCrotch2", xyBackCrotch2, nn
'    Pr_DrawText "xyBackCrotch3", xyBackCrotch3, nn
'    Pr_DrawText "xyBackCrotchFilletCentre", xyBackCrotchFilletCentre, nn
 '   Pr_DrawText "xyFrontCrotch1", xyFrontCrotch1, nn
'    Pr_DrawText "xyFrontCrotch2", xyFrontCrotch2, nn
'    Pr_DrawText "xyFrontCrotch3", xyFrontCrotch3, nn
'    Pr_DrawText "xyFrontCrotchFilletCentre", xyFrontCrotchFilletCentre, nn

'    Pr_DrawText "xyRaglan1", xyRaglan1, nn
'    Pr_DrawText "xyRaglan2", xyRaglan2, nn
'    Pr_DrawText "xyRaglan3", xyRaglan3, nn
'    Pr_DrawText "xyRaglan4", xyRaglan4, nn
'    Pr_DrawText "xyRaglan5", xyRaglan5, nn
'    Pr_DrawText "xyFrenchCut", xyFrenchCut, nn

                   
End Sub

Private Sub PR_DoBraCupsAndDisks()
   'Procedure to either
   '    1. Check if there is enough data to enable a
   '       calculation.
   ' or
   '    2. Check if a change has been made to the data
   '       requiring a recalculation.
   '
   'Then calculate the bra cups and disks.
   '
   'NOTE
   '    txtCir(5) == Chest Circumference
   '    txtCir(10) == Under Breast Circumference
   '    txtCir(11) == Circumference over nipple line
   '
   'GLOBALS
   '    g_nUnitsFac As Double
   '    g_nUnderBreast As Double
   '    g_nChest As Double
   '    g_nNipple As Double
   '    g_sLeftCup As String
   '    g_sRightCup As String
   '    g_sLeftDisk As String
   '    g_sRightDisk As String

    Dim sLeftCup$, sRightCup$, sSide%
    Dim ii%, iError%
    Dim nNipple#, nUnderBreast#, nChest#
    Dim sError$
    Dim iLeftDisk%, iRightDisk%

    sLeftCup = cboLeftCup.Text
    sRightCup = cboRightCup.Text

   'Check if enough data to caclulate
   'Exit sub if not
    If Val(txtCir(5).Text) = 0 Then
       'No chest cir.
        Exit Sub
    ElseIf Val(txtCir(10).Text) = 0 Then
       'No under breast cir.
        Exit Sub
    ElseIf Val(txtCir(11).Text) = 0 And ((sLeftCup = "" Or sLeftCup = "None") And (sRightCup = "" Or sRightCup = "None")) Then
       'No over nipple circumference and no bra cups given for either side
        Exit Sub
    End If

   'Check if changed
    Dim Changed%

    Changed = False

    If sLeftCup <> g_sLeftCup Then
        Changed = True
        g_sLeftCup = sLeftCup
    End If
    
    If sRightCup <> g_sRightCup Then
        Changed = True
        g_sRightCup = sRightCup
    End If
    
    If Val(txtCir(10).Text) <> g_nUnderBreast Then
        Changed = True
        g_nUnderBreast = Val(txtCir(10).Text)
    End If
    
    If Val(txtCir(11).Text) <> g_nNipple Then
        Changed = True
        g_nNipple = Val(txtCir(11).Text)
       'Force a recalculation of cup
        If sLeftCup <> "None" Then sLeftCup = ""
        If sRightCup <> "None" Then sRightCup = ""
    End If

    If Val(txtCir(5).Text) <> g_nChest Then
        g_nChest = Val(txtCir(5).Text)
       'Force a recalculation of cup if a nipple circum. is given
       'Changed chest is only significant if a nipple line measurement
       'is given from which the Cup size is calculated
        If g_nNipple <> 0 Then
            Changed = True
            If sLeftCup <> "None" Then sLeftCup = ""
            If sRightCup <> "None" Then sRightCup = ""
        End If
    End If

   'Force a calculation if either disk is empty
    If txtLeftDisk.Text = "" Or txtRightDisk = "" Then Changed = True

   'Force a check if disk has been changed
    If txtLeftDisk.Text <> g_sLeftDisk Or txtRightDisk <> g_sRightDisk Then
        Changed = True
        iRightDisk = Val(txtRightDisk)
        iLeftDisk = Val(txtLeftDisk)
    End If
 
   'If no modifications then exit
    If Changed <> True Then Exit Sub
    

   'Recalculate if Changed
   'Convert to inches
    nNipple = fnDisplayToInches(g_nNipple)
    nUnderBreast = fnDisplayToInches(g_nUnderBreast)
    nChest = fnDisplayToInches(g_nChest)
   
   'Right cup
    If sRightCup <> "" Or nNipple <> 0 Then
        iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sRightCup, iRightDisk, "Right")
        If iError Then
            For ii = 0 To (cboRightCup.ListCount - 1)
                If cboRightCup.List(ii) = sRightCup Then cboRightCup.ListIndex = ii
            Next ii
            If iRightDisk = 0 Then txtRightDisk.Text = "" Else txtRightDisk = iRightDisk
            g_sRightDisk = Trim$(Str$(iRightDisk))
            g_sRightCup = sRightCup
        Else
            txtRightDisk.Text = ""
            cboRightCup.ListIndex = -1
            g_sRightDisk = ""
            g_sRightCup = ""
        End If
    End If

    If sLeftCup <> "" Or nNipple <> 0 Then
        iError = FN_CalculateBra(nChest, nUnderBreast, nNipple, sLeftCup, iLeftDisk, "Left")
        If iError Then
            For ii = 0 To (cboLeftCup.ListCount - 1)
                If cboLeftCup.List(ii) = sLeftCup Then cboLeftCup.ListIndex = ii
            Next ii
            If iLeftDisk <= 0 Then txtLeftDisk.Text = "" Else txtLeftDisk = iLeftDisk
            g_sLeftDisk = Trim$(Str$(iLeftDisk))
            g_sLeftCup = sLeftCup
        Else
            txtLeftDisk.Text = ""
            cboLeftCup.ListIndex = -1
            g_sLeftDisk = ""
            g_sLeftCup = ""
        End If
    End If

    PR_EnableCalculateDiskButton
    

End Sub

Private Sub PR_DrawBra()
   'Assumes that data has been validated before this
   'procedure is called.
    Dim nBraCLOffset#, nDiskXoff#, nDiskYoff#
    Dim ii%, iDisk%, iLeftDisk%, iRightDisk%, iNextDisk%
    Dim sBraText As String, sDisk As String
    Dim xyDisk As XY
    ReDim xyPt(1 To 3) As XY

    If cboLeftCup.ListIndex < 7 Or cboLeftCup.ListIndex < 7 Or txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "" Then
       'A cup or a disk of some sort has been given
        If cboLeftCup.Text = "None" Then iLeftDisk = -1 Else iLeftDisk = Val(txtLeftDisk.Text)
        If cboRightCup.Text = "None" Then iRightDisk = -1 Else iRightDisk = Val(txtRightDisk.Text)

        If iLeftDisk = iRightDisk Then
            iNextDisk = 1
            If iLeftDisk = -1 Then
                sBraText = "NO CUP LEFT & RIGHT"
            Else
                sBraText = "No." & Str$(iLeftDisk) & " LEFT & RIGHT"
            End If
        Else
            iNextDisk = 2
            If iLeftDisk = -1 Then
                sBraText = "NO CUP LEFT"
            Else
                sBraText = "No." & Str$(iLeftDisk) & " LEFT"
            End If

            If iRightDisk = -1 Then
                sBraText = sBraText & "\nNO CUP RIGHT"
            Else
                sBraText = sBraText & "\nNo." & Str$(iRightDisk) & " RIGHT"
            End If
        End If

       'Loop through both cups and give positioning information
        For ii = 1 To iNextDisk
            If ii = 1 Then iDisk = iLeftDisk Else iDisk = iRightDisk
            Select Case iDisk
                Case -1, 0
                   nBraCLOffset = 0
                   nDiskXoff = 0
                   nDiskYoff = 0
                Case 1
                   nBraCLOffset = 1.25
                   nDiskXoff = 1.45
                   nDiskYoff = 1.797
                Case 2
                   nBraCLOffset = 1.125
                   nDiskXoff = 1.652
                   nDiskYoff = 2.029
                Case 3
                   nBraCLOffset = 1#
                   nDiskXoff = 1.893
                   nDiskYoff = 2.288
                Case 4
                   nBraCLOffset = 0.875
                   nDiskXoff = 2.175
                   nDiskYoff = 2.572
                Case 5
                   nBraCLOffset = 0.75
                   nDiskXoff = 2.293 + 0.104
                   nDiskYoff = 2.882
                Case 6
                   nBraCLOffset = 0.625
                   nDiskXoff = 2.625 + 0.0272
                   nDiskYoff = 3.025
                Case 7
                   nBraCLOffset = 0.5
                   nDiskXoff = 2.94 + 0.02
                   nDiskYoff = 3.335
                Case 8
                   nBraCLOffset = 0.5
                   nDiskXoff = 3.139
                   nDiskYoff = 3.518
                Case 9
                   nBraCLOffset = 0.5
                   nDiskXoff = 3.393
                   nDiskYoff = 3.938
            End Select
            
            If iDisk > 0 Then
               'Insert disk
                PR_SetLayer "Notes"
                xyDisk.X = (xySeamChest.X - 0.75) - nDiskXoff
                xyDisk.y = (xyCutOut7.y - nBraCLOffset) '- (nDiskYoff * 2)
    
                sDisk = "bradisk" & Trim$(Str$(iDisk)) & "mirrored"
                PR_InsertSymbol sDisk, xyDisk, 1, 0
                PR_AddDBValueToLast "curvetype", "Bra"
               
               'Draw Construction lines
                PR_SetLayer "Construct"
                PR_MakeXY xyPt(1), xySeamChest.X - 0.75, xyCutOut7.y - nBraCLOffset
                PR_MakeXY xyPt(2), xyPt(1).X, (xyCutOut7.y - nBraCLOffset) - (nDiskYoff * 2)
                PR_MakeXY xyPt(3), xyPt(1).X - (nDiskXoff * 2), xyPt(1).y
                PR_DrawLine xyPt(1), xyPt(2)
                PR_AddDBValueToLast "curvetype", "Bra"
                PR_DrawLine xyPt(1), xyPt(3)
                PR_AddDBValueToLast "curvetype", "Bra"
                If ii = iNextDisk Then
                   'Insert bra label text
                    PR_SetLayer "Notes"
                    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
                    PR_CalcMidPoint xyPt(2), xyPt(3), xyPt(1)
                    Pr_DrawText sBraText, xyPt(1), 0.125
                End If
            Else
               'No Bra cup.
               'Do nothing as this is dealt with as text.
               'Except for this very special case
                If iNextDisk = 1 Then
                   'Insert Text (Special case No Cups Left and Right ????)
                    PR_SetLayer "Notes"
                    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
                    PR_MakeXY xyPt(1), xySeamChest.X - 2.75, xySeamChest.y - 2.5
                    Pr_DrawText sBraText, xyPt(1), 0.125
                End If
            End If
        Next ii
    End If
End Sub

Private Sub PR_DrawBriefFrenchCut()
    Dim cBriefPoly As curve
    Dim cCrotchPoly As curve
    Dim nArc As Double
    Dim xyArc As XY
    Dim aArcAngle As Double, aAngle As Double, ii As Integer, aConstruct   As Double
    Dim xyTmp As XY

    If g_bDiffThigh Then
        PR_DrawLine xyLT_ProfileThigh, xyProfileBrief
        PR_DrawLine xyProfileThighMirror1, xyLT_ProfileThighMirror
    Else
        PR_DrawLine xyProfileThigh, xyProfileBrief
        PR_DrawLine xyProfileThighMirror1, xyProfileThighMirror
    End If
    cBriefPoly.n = 0

    If g_sCrotchStyle <> "Snap Crotch" Then
       'All other crotch styles
        cBriefPoly.n = 0
        aConstruct = FN_CalcAngle(xyLegPoint, xyProfileBrief)
        aConstruct = aConstruct - ((aConstruct - 90) / 2)
        nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
       ' PR_MakeXY xyArc, xyProfileBrief.X + nArc, xyProfileBrief.Y
        PR_CalcPolar xyProfileBrief, aConstruct - 90, nArc, xyArc
        aAngle = FN_CalcAngle(xyArc, xyProfileBrief)
        aArcAngle = 20
        For ii = 1 To 3
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle + (aArcAngle / 2)
        Next ii

        cBriefPoly.n = cBriefPoly.n + 1
        cBriefPoly.X(cBriefPoly.n) = xyLegPoint.X
        cBriefPoly.y(cBriefPoly.n) = xyLegPoint.y

        nArc = (xyFrenchCut.X - xySeamFold.X)
        If Not g_nAdult Then
            nArc = nArc * 0.5
            aAngle = 40
            aArcAngle = 80
        Else
            aAngle = 45
            aArcAngle = 90
        End If
        PR_MakeXY xyArc, xyFrenchCut.X - nArc, xyFrenchCut.y
        For ii = 1 To 5
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle - (aArcAngle / 4)
        Next ii

        nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
        If Not g_nAdult Then
            nArc = nArc * 0.25
            PR_MakeXY xyArc, xyProfileThighMirror.X + nArc, xyProfileThighMirror.y
            nArc = FN_CalcLength(xyProfileThighMirror1, xyArc)
            aArcAngle = 40
            aAngle = FN_CalcAngle(xyProfileThighMirror1, xyArc) - aArcAngle
        Else
            PR_MakeXY xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y
            aArcAngle = 40
            aAngle = 180 - aArcAngle
        End If
        PR_MakeXY xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y
        aArcAngle = 40
        aAngle = 180 - aArcAngle
        For ii = 1 To 5
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle + (aArcAngle / 4)
        Next ii

        PR_DrawFitted cBriefPoly
    
    Else
       'Snap Crotch
        PR_DrawSnapCrotchArc
        cCrotchPoly.X(1) = xyStrap1.X
        cCrotchPoly.y(1) = xyStrap1.y
        cCrotchPoly.X(2) = xyStrap2.X
        cCrotchPoly.y(2) = xyStrap2.y
        cCrotchPoly.X(3) = xyStrap4.X
        cCrotchPoly.y(3) = xyStrap4.y
        cCrotchPoly.n = 3
        PR_DrawPoly cCrotchPoly
        
        PR_DrawLine xyStrap4, xyLegPoint

        cBriefPoly.n = cBriefPoly.n + 1
        cBriefPoly.X(cBriefPoly.n) = xyLegPoint.X
        cBriefPoly.y(cBriefPoly.n) = xyLegPoint.y

        nArc = (xyFrenchCut.X - xySeamFold.X)
        If Not g_nAdult Then
            nArc = nArc * 0.5
            aAngle = 40
            aArcAngle = 80
        Else
            aAngle = 45
            aArcAngle = 90
        End If
        PR_MakeXY xyArc, xyFrenchCut.X - nArc, xyFrenchCut.y
        aAngle = 45
        aArcAngle = 90
        For ii = 1 To 5
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle - (aArcAngle / 4)
        Next ii

        nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
        If Not g_nAdult Then
            nArc = nArc * 0.25
            PR_MakeXY xyArc, xyProfileThighMirror.X + nArc, xyProfileThighMirror.y
            nArc = FN_CalcLength(xyProfileThighMirror1, xyArc)
            aArcAngle = 40
            aAngle = FN_CalcAngle(xyProfileThighMirror1, xyArc) - aArcAngle
        Else
            PR_MakeXY xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y
            aArcAngle = 40
            aAngle = 180 - aArcAngle
        End If

        PR_MakeXY xyArc, xyProfileThighMirror1.X + nArc, xyProfileThighMirror1.y
        aArcAngle = 40
        aAngle = 180 - aArcAngle
        For ii = 1 To 5
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle + (aArcAngle / 4)
        Next ii

        PR_DrawFitted cBriefPoly
    End If
End Sub

Private Sub PR_DrawBriefStandard(nXOffset As Double)
    Dim cBriefPoly As curve
    Dim cCrotchPoly As curve
    Dim nArc As Double
    Dim xyArc As XY
    Dim aArcAngle As Double, aAngle As Double, ii As Integer, aConstruct As Double
    Dim xyTmp As XY, xyTmp1    As XY
    Dim nRadius As Double
    Dim cSnapCrotchPoly As curve

   'To take account of the One leg Panty and One leg Brief
   'we use the nXOffset to move the profile
   'This is a fudge but one that works.
   'Normal brief only not applicable to the Brief-French
   'NB we can reset the Global xy...  points as they are not used again.
   
    If nXOffset > 0 Then '==>Mixed Panty & Brief
        xyProfileThigh = xyFold
        PR_MakeXY xyProfileThighMirror, xyProfileThigh.X, -xyProfileThigh.y
        PR_MakeXY xyProfileThighMirror1, xyProfileThighMirror.X, xyProfileThighMirror.y + INCH1_4
        xyProfileBrief.X = xyProfileBrief.X + nXOffset
    

        xyLT_ProfileThigh = xyLT_Fold
        PR_MakeXY xyLT_ProfileThighMirror, xyLT_ProfileThigh.X, -xyLT_ProfileThigh.y
'        PR_MakeXY xyLT_ProfileThighMirror1, xyLT_ProfileThighMirror.x, xyLT_ProfileThighMirror.y + inch1_4
        
        xyLegPoint.X = xyLegPoint.X + nXOffset
        xyCutOutArcCentre.X = xyCutOutArcCentre.X + nXOffset
        xySeamThigh.X = xySeamThigh.X + nXOffset

        xyLT_ProfileThighMirror.X = xyLT_ProfileThighMirror.X + nXOffset
        xyStrap4.X = xyStrap4.X + nXOffset
        xyGussetArcCentre.X = xyGussetArcCentre.X + nXOffset

    End If

    If g_bDiffThigh Then
        PR_DrawLine xyLT_ProfileThigh, xyProfileBrief
    Else
        PR_DrawLine xyProfileThigh, xyProfileBrief
    End If

    If g_sCrotchStyle <> "Snap Crotch" Then
       'All other crotch styles
        cBriefPoly.n = 0
        aConstruct = FN_CalcAngle(xyLegPoint, xyProfileBrief)
        aConstruct = aConstruct - ((aConstruct - 90) / 2)
        nArc = (xyCutOutArcCentre.X - xySeamThigh.X) * (2 / 3)
       ' PR_MakeXY xyArc, xyProfileBrief.X + nArc, xyProfileBrief.Y
        PR_CalcPolar xyProfileBrief, aConstruct - 90, nArc, xyArc
        aAngle = FN_CalcAngle(xyArc, xyProfileBrief)
        aArcAngle = 20
        For ii = 1 To 3
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle + (aArcAngle / 2)
        Next ii

        nArc = (xyLegPoint.X - xySeamThigh.X) * (3 / 2)
       'PR_MakeXY xyArc, xyLegPoint.X - nArc, xyLegPoint.Y
        PR_CalcPolar xyLegPoint, aConstruct + 90, nArc, xyArc
        aArcAngle = 15
        aAngle = FN_CalcAngle(xyArc, xyLegPoint) + aArcAngle
        For ii = 1 To 3
            PR_CalcPolar xyArc, aAngle, nArc, xyTmp
            cBriefPoly.n = cBriefPoly.n + 1
            cBriefPoly.X(cBriefPoly.n) = xyTmp.X
            cBriefPoly.y(cBriefPoly.n) = xyTmp.y
            aAngle = aAngle - (aArcAngle / 2)
        Next ii

        PR_DrawFitted cBriefPoly
        
        cBriefPoly.X(1) = xyLegPoint.X
        cBriefPoly.y(1) = xyLegPoint.y
        cBriefPoly.X(2) = xyProfileThighMirror1.X
        cBriefPoly.y(2) = xyProfileThighMirror1.y
        If g_bDiffThigh Then
            cBriefPoly.X(3) = xyLT_ProfileThighMirror.X
            cBriefPoly.y(3) = xyLT_ProfileThighMirror.y
        Else
            cBriefPoly.X(3) = xyProfileThighMirror.X
            cBriefPoly.y(3) = xyProfileThighMirror.y
        End If
        cBriefPoly.n = 3
        PR_DrawPoly cBriefPoly
    Else
       'Snap Crotch
        PR_DrawSnapCrotchArc

        cCrotchPoly.X(1) = xyStrap1.X
        cCrotchPoly.y(1) = xyStrap1.y
        cCrotchPoly.X(2) = xyStrap2.X
        cCrotchPoly.y(2) = xyStrap2.y
        cCrotchPoly.X(3) = xyStrap4.X
        cCrotchPoly.y(3) = xyStrap4.y
        cCrotchPoly.n = 3
        PR_DrawPoly cCrotchPoly
        cBriefPoly.X(1) = xyStrap4.X
        cBriefPoly.y(1) = xyStrap4.y
        cBriefPoly.X(2) = xyLegPoint.X
        cBriefPoly.y(2) = xyLegPoint.y
        cBriefPoly.X(3) = xyProfileThighMirror1.X
        cBriefPoly.y(3) = xyProfileThighMirror1.y
        If g_bDiffThigh Then
            cBriefPoly.X(4) = xyLT_ProfileThighMirror.X
            cBriefPoly.y(4) = xyLT_ProfileThighMirror.y
        Else
            cBriefPoly.X(4) = xyProfileThighMirror.X
            cBriefPoly.y(4) = xyProfileThighMirror.y
        End If
        cBriefPoly.n = 4
        PR_DrawPoly cBriefPoly
    End If
End Sub

Private Sub PR_DrawMesh(MeshProfile As BiArc, MeshSeam As BiArc, sSide As String)
   'Subroutime to draw the given mesh
   'used to save code
   'NB the misnomer. Too complicated to change in the DrawMacro
   'procedure so do it here
    PR_SetLayer "Notes"
    PR_DrawArc MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent
    PR_DrawArc MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd
    PR_SetLayer "Template" & sSide
    PR_DrawArc MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent
    PR_DrawArc MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd

End Sub

Private Sub PR_DrawMeshConstruction(xyRaglanStart As XY, xyMeshStart As XY, sAxillaText As String, nMeshLength#)
   'Draws the custruction lines for a mesh
   'Save code
   'Returns
   '    xyMeshStart point
   '    AxillaText
   '    Length of Mesh
    Dim nOffset#, xyTmp As XY, xyTmp1 As XY
    PR_SetLayer "1"
    If g_nAge < 14 Then
        nOffset = 0.75
        nMeshLength = ONE_and_THREE_QUARTER_GUSSET '2.58
        sAxillaText = "GUSSET 1-3/4\"""
    Else
        nOffset = 0.8125
        nMeshLength = BOYS_GUSSET   '2.8
        sAxillaText = "BOY'S GUSSET"
    End If
    
    PR_MakeXY xyTmp, xyRaglanStart.X - nOffset, xyRaglanStart.y
    PR_DrawLine xyRaglanStart, xyTmp
    PR_MakeXY xyTmp1, xyTmp.X, xyTmp.y + 0.25
    PR_MakeXY xyTmp, xyTmp1.X, xyTmp1.y - 0.5
    PR_DrawLine xyTmp1, xyTmp
    
    PR_MakeXY xyMeshStart, xyRaglanStart.X - nOffset, xyRaglanStart.y

End Sub

Private Sub PR_DrawSnapCrotchArc()
    Dim aAngle As Double
    Dim xyTmp As XY, xyTmp1 As XY
    Dim nRadius As Double
    Dim cSnapCrotchPoly As curve
    Dim sLayer As String
    
   'Check if the arc is too close to the cut-out
   'it should be no closer than 1.25 inches
    nRadius = FN_CalcLength(xyGussetArcCentre, xyProfileBrief)

    If FN_CalcCirCirInt(xyGussetArcCentre, nRadius, xyCutOut5, 1.25, xyTmp, xyTmp1) Then
       'The arc is too close
       'draw as a fitted curve to allow for user adjustment
        MsgBox "The calculated arc for the Snap Crotch was too close to the cut-out. It will be drawn as fitted curve.  Adjust this as required."
        cSnapCrotchPoly.n = 3
        cSnapCrotchPoly.X(1) = xyProfileBrief.X
        cSnapCrotchPoly.y(1) = xyProfileBrief.y
        aAngle = FN_CalcAngle(xyGussetArcCentre, xyCutOut5)
        nRadius = FN_CalcLength(xyGussetArcCentre, xyCutOut5)
        PR_CalcPolar xyGussetArcCentre, aAngle, nRadius + 1.25, xyTmp
        cSnapCrotchPoly.X(2) = xyTmp.X
        cSnapCrotchPoly.y(2) = xyTmp.y
        cSnapCrotchPoly.X(cSnapCrotchPoly.n) = xyStrap3.X
        cSnapCrotchPoly.y(cSnapCrotchPoly.n) = xyStrap3.y
        PR_DrawFitted cSnapCrotchPoly

       'Draw a circle on the layer construct for use by the user
        sLayer = g_sCurrentLayer
        PR_SetLayer "Construct"
        PR_DrawCircle xyCutOut5, 1.25
        PR_SetLayer sLayer

    Else
       'Draw arc as normal
        PR_DrawArc xyGussetArcCentre, xyProfileBrief, xyStrap3
    End If

'PR_DrawMarker xyGussetArcCentre
'PR_DrawText "xyGussetArcCentre", xyGussetArcCentre, .1
'PR_DrawMarker xyProfileBrief
'PR_DrawText "xyProfileBrief", xyProfileBrief, .1
'PR_DrawMarker xyStrap3
'PR_DrawText "xyStrap3", xyStrap3, .1
'PR_DrawMarker xyStrap1
'PR_DrawText "xyStrap1", xyStrap1, .1
'PR_DrawMarker xyStrap2
'PR_DrawText "xyStrap2", xyStrap2, .1
'PR_DrawMarker xyCutOut5
'PR_DrawText "xyCutOut5", xyCutOut5, .1
'PR_DrawMarker xyCutOut6
'PR_DrawText "xyCutOut6", xyCutOut6, .1
'PR_DrawMarker xyCutOut7
'PR_DrawText "xyCutOut7", xyCutOut7, .1
'PR_DrawMarker xyCutOut8
'PR_DrawText "xyCutOut8", xyCutOut7, .1
    
End Sub

Private Static Sub PR_DrawVerticalLineThruPoint(xyPnt As XY, nHeight As Double)
   'Function to reduce code in PR_DrawMacro
    Dim xy1 As XY, xy2  As XY
    PR_MakeXY xy1, xyPnt.X, xyPnt.y - nHeight
    PR_MakeXY xy2, xyPnt.X, xyPnt.y + nHeight
    PR_DrawLine xy1, xy2

End Sub

Private Sub PR_EnableCalculateDiskButton()
   'Procedure to check if there is enough data to
   'enable the caclculate disks command button
   '
    Dim sLeftCup$, sRightCup$

    sLeftCup = cboLeftCup.Text
    sRightCup = cboRightCup.Text

   'Check if enough data to caclulate
   'Exit sub if not
    If Val(txtCir(5).Text) = 0 Then
       'No chest cir.
        cmdCalculateBraDisks.Enabled = False
        Exit Sub
    ElseIf Val(txtCir(10).Text) = 0 Then
       'No under breast cir.
        cmdCalculateBraDisks.Enabled = False
        Exit Sub
    ElseIf Val(txtCir(11).Text) = 0 And (sLeftCup = "" Or sLeftCup = "None") And (sRightCup = "" Or sRightCup = "None") Then
       'No over nipple circumference and no bra cups given for either side
        cmdCalculateBraDisks.Enabled = False
        Exit Sub
    End If
    
    cmdCalculateBraDisks.Enabled = True

End Sub

Private Sub PR_GetComboListFromFile(Combo_Name As Control, sFileName As String)
   'General procedure to create the list section of
   'a combo box reading the data from a file
    
    Dim sLine As String, fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48, "BodySuit Dialogue"
        Exit Sub
    End If

    Open sFileName For Input As fFileNum
    Do While Not EOF(fFileNum)
        Line Input #fFileNum, sLine
        Combo_Name.AddItem sLine
        Loop
    Close #fFileNum

End Sub

Private Sub PR_Select_Text(Text_Box_Name As Control)
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub

Private Sub PR_SetCrotchSizes()

    'Sets two global variables:
    '   g_nFrontCrotchSize (type double)
    '   g_nBackCrotchSize (type double)
    '
    'NOTE: DOES NOT CATER FOR 1 LEG PEOPLE

    If g_sSex = "Male" Then
            If g_nAdult Then
                    If g_nButtocksCirGiven < 45 Then   '45"
                            g_nFrontCrotchSize = 4
                            g_nBackCrotchSize = 3
                        Else
                            g_nFrontCrotchSize = 4.5
                            g_nBackCrotchSize = 3.5
                    End If
                Else
                    g_nFrontCrotchSize = 2.25
                    g_nBackCrotchSize = 1.5
            End If
        Else
            If g_nAdult Then
                    If g_nButtocksCirGiven < 45 Then   '45"
                            g_nFrontCrotchSize = 3
                            g_nBackCrotchSize = 4
                        Else
                            g_nFrontCrotchSize = 3.5
                            g_nBackCrotchSize = 4.5
                    End If
                Else
                    g_nFrontCrotchSize = 1.5
                    g_nBackCrotchSize = 2.25
            End If
    End If

End Sub

Private Sub PR_SetFlySize(sType As String)

    'Sets two global variables:
    '   g_sFlySize (type string)
    '   g_nFlyLength (type double) not at this time
    
    Dim nCutOutDia As Double
    'nCutOutDia is in inches
    nCutOutDia = fnRoundInches(g_nCutOutRadius * 2)
'MsgBox "sType=" & sType & "nCutOutDia=" & Str$(nCutOutDia) & "g_nAge=" & Str$(g_nAge)
    If sType = "Hor" Then
            Select Case g_nAge
                Case 3 To 6
                    If nCutOutDia <= 1.5 Then
                            g_sFlySize = "C-1"
                            'g_nFlyLength = "2"
                        Else
                            g_sFlySize = "C-2"
                            'g_nFlyLength = 2.5
                    End If
                Case 7 To 11
                    If nCutOutDia <= 2 Then
                            g_sFlySize = "C-2"
                            'g_nFlyLength = 2.5
                        Else
                            g_sFlySize = "C-3"
                            'g_nFlyLength = 2.5
                    End If
                Case 12 To 14
                    Select Case nCutOutDia
                        Case Is <= 2.25
                            g_sFlySize = "C-3"
                            'g_nFlyLength = 2.5
                        Case 2.375 To 2.5
                            g_sFlySize = "C-4"
                            'g_nFlyLength = 2.5
                        Case Is > 2.5
                            g_sFlySize = "C-5"
                            'g_nFlyLength = 3.125
                    End Select
                Case Is >= 15
                    Select Case nCutOutDia
                        Case Is <= 3.5
                            g_sFlySize = "1"
                            'g_nFlyLength = 3.5
                        Case 3.625 To 4.25
                            g_sFlySize = "2"
                            'g_nFlyLength = 3.75
                        Case 4.375 To 5
                            g_sFlySize = "3"
                            'g_nFlyLength = 4
                        Case Is > 5.125
                            g_sFlySize = "Oversize"
                            'g_nFlyLength = 4
                    End Select
            End Select
        Else
            Select Case g_nAge
                Case Is < 10
                    g_sFlySize = "Small"
                    'g_nFlyLength = 1.75
                Case 10 To 14
                    g_sFlySize = "Medium"
                    'g_nFlyLength = 2.75
                Case Is >= 15
                    g_sFlySize = "Large"
                    'g_nFlyLength = 3.75
            End Select
    End If

End Sub
    

Private Sub PR_SetGussetSize()

    'Sets two global variables:
    '   g_sGussetSize (type string)
    '   g_nGussetLength (type double) not at this time

    Dim nCutOutDia As Double
    
    'nCutOutDia is in inches
    nCutOutDia = fnRoundInches(g_nCutOutRadius * 2)
    If g_sSex = "Male" Then
            Select Case g_nAge
                Case Is <= 3
                    g_sGussetSize = "1-3/4\"""
                    g_nGussetLength = 2.125
                Case 4 To 14
                    g_sGussetSize = "Boy's"
                    g_nGussetLength = 3.875
                Case Is >= 15
                    g_sGussetSize = "Male"
                    g_nGussetLength = 5.375
            End Select
        Else
            Select Case g_nAge
                Case Is <= 3
                    g_sGussetSize = "1-3/4\"""
                    g_nGussetLength = 2.125
                Case 4 To 14
                    Select Case nCutOutDia
                        Case Is <= 1.75
                            g_sGussetSize = "1\"""
                            g_nGussetLength = 2.125
                        Case 1.875 To 3.125
                            g_sGussetSize = "1-1/4\"""
                            g_nGussetLength = 2.75
                        Case Is >= 3.25
                            g_sGussetSize = "Regular"
                            g_nGussetLength = 3.625
                    End Select
                Case Is >= 15
                    If g_nButtocksCirGiven <= 45 Then
                            g_sGussetSize = "Regular"
                            g_nGussetLength = 3.625
                        Else
                            g_sGussetSize = "Oversize"
                            g_nGussetLength = 4.125
                    End If
            End Select
    End If

End Sub

Private Sub PR_UpdateDB()
   'Procedure called from PR_CreateDrawMacro
    
    Dim sSymbol As String, sMeshData  As String
    
    sSymbol = "suitbody"

   'Use existing symbol
    PR_PutLine "hBody = UID (" & QQ & "find" & QC & Val(txtUidVB.Text) & ");"
    PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

   'Update the BODY Box symbol
    
    PR_PutLine "SetDBData( hBody" & CQ & "AxillaBackNeckRad" & QCQ & g_nAxillaBackNeckRad & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "AxillaFrontNeckRad" & QCQ & g_nAxillaFrontNeckRad & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ABNRadRight" & QCQ & g_nABNRadRight & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "AFNRadRight" & QCQ & g_nAFNRadRight & QQ & ");"
    
   'Mesh data for use by sleeve drawing programme
   'We use commas as the delimiters to show empty data
    If g_sMeshLeft = "" And g_sMeshRight = "" Then
       sMeshData = ",,,"

    ElseIf g_sMeshLeft = "" Then
       sMeshData = ",," & g_sMeshRight

    ElseIf g_sMeshRight = "" Then
       sMeshData = g_sMeshLeft & ",,"

    Else
       sMeshData = g_sMeshLeft & "," & g_sMeshRight
    
    End If
    
    PR_PutLine "SetDBData( hBody" & CQ & "Data" & QCQ & sMeshData & QQ & ");"

End Sub

Private Sub prOptLeftLeg(i0 As Integer, i1 As Integer)
   'Procedure to set the option buttons
   'sets them according to the pattern i0-i2
    
    optLeftLeg(0).Value = i0
    optLeftLeg(1).Value = i1

End Sub

Private Sub prOptRightLeg(i0 As Integer, i1 As Integer)
   'Procedure to set the option buttons
   'sets them according to the pattern i0-i2

    optRightLeg(0).Value = i0
    optRightLeg(1).Value = i1

End Sub

Private Sub PrOutPut()

'    Dim fOutPut As Integer
'
'    'Open file
'    fOutPut = FreeFile
'    Open "C:\JOBST\VARS.TXT" For Output As fOutPut
'
'    Print #fOutPut, "BodySuit variables"
'    Print #fOutPut,
'    Print #fOutPut, "g_sCurrentLayer= ", g_sCurrentLayer
'    Print #fOutPut, "g_nCurrTextHt= ", g_nCurrTextHt
'    Print #fOutPut, "g_nCurrTextAspect= ", g_nCurrTextAspect
'    Print #fOutPut, "g_nCurrTextHorizJust= ", g_nCurrTextHorizJust
'    Print #fOutPut, "g_nCurrTextVertJust= ", g_nCurrTextVertJust
'    Print #fOutPut, "g_nCurrTextFont= ", g_nCurrTextFont
'    Print #fOutPut, "g_nCurrTextAngle= ", g_nCurrTextAngle
'
'    Print #fOutPut,
'    Print #fOutPut, "Patient details"
'    Print #fOutPut,
'    Print #fOutPut, "g_sFileNo= ", g_sFileNo
'    Print #fOutPut, "g_sSide= ", g_sSide
'    Print #fOutPut, "g_sPatient= ", g_sPatient
'    Print #fOutPut, "g_sID= ", g_sID
'    Print #fOutPut, "g_sDiagnosis= ", g_sDiagnosis
'    Print #fOutPut, "g_nAge= ", g_nAge
'    Print #fOutPut, "g_sSex= ", g_sSex
'    Print #fOutPut, "g_sWorkOrder= ", g_sWorkOrder
'    Print #fOutPut, "g_nUnitsFac= ", g_nUnitsFac
'    Print #fOutPut, "g_sBackNeck= ", g_sBackNeck
'    Print #fOutPut, "g_sFrontNeck= ", g_sFrontNeck
'    Print #fOutPut, "g_nUnderBreast= ", g_nUnderBreast
'    Print #fOutPut, "g_nNipple= ", g_nNipple
'    Print #fOutPut, "g_nChest= ", g_nChest
'    Print #fOutPut, "g_sLeftCup= ", g_sLeftCup
'    Print #fOutPut, "g_sRightCup= ", g_sRightCup
'    Print #fOutPut, "g_sLeftDisk= ", g_sLeftDisk
'    Print #fOutPut, "g_sRightDisk= ", g_sRightDisk
'
'    Print #fOutPut,
'    Print #fOutPut, "Response & DiffSides"
'    Print #fOutPut,
 '   Print #fOutPut, "g_bResponse= ", g_bResponse
'    Print #fOutPut, "g_bDiffThigh= ", g_bDiffAxillaType
'
'    Print #fOutPut,
 '   Print #fOutPut, "Values Given"
'    Print #fOutPut,
'    Print #fOutPut, "g_nLeftShldCirGiven= ", g_nLeftShldCirGiven
'    Print #fOutPut, "g_nRightShldCirGiven= ", g_nRightShldCirGiven
'    Print #fOutPut, "g_nLeftAxillaCirGiven= ", g_nLeftAxillaCirGiven
'    Print #fOutPut, "g_nRightAxillaCirGiven= ", g_nRightAxillaCirGiven
'    Print #fOutPut, "g_nChestCirGiven= ", g_nChestCirGiven
'    Print #fOutPut, "g_nUnderBreastCirGiven= ", g_nUnderBreastCirGiven
'    Print #fOutPut, "g_nWaistCirGiven= ", g_nWaistCirGiven
'    Print #fOutPut, "g_nButtocksCirGiven= ", g_nButtocksCirGiven
'    Print #fOutPut, "g_nLeftThighCirGiven= ", g_nLeftThighCirGiven
'    Print #fOutPut, "g_nRightThighCirGiven= ", g_nRightThighCirGiven
'    Print #fOutPut, "g_nNeckCirGiven= ", g_nNeckCirGiven
'    Print #fOutPut, "g_nShldToFoldGiven= ", g_nShldToFoldGiven
'    Print #fOutPut, "g_nShldToWaistGiven= ", g_nShldToWaistGiven
'    Print #fOutPut, "g_nShldToBreastGiven= ", g_nShldToBreastGiven
'    Print #fOutPut, "g_nLeftShldToAxilla= ", g_nLeftShldToAxilla
'    Print #fOutPut, "g_nRightShldToAxilla= ", g_nRightShldToAxilla
'    Print #fOutPut, "g_nNippleCirGiven= ", g_nNippleCirGiven
'    Print #fOutPut, "g_sFrontNeckStyle= ", g_sFrontNeckStyle
'    Print #fOutPut, "g_nFrontNeckSize= ", g_nFrontNeckSize
'    Print #fOutPut, "g_sBackNeckStyle= ", g_sBackNeckStyle
'    Print #fOutPut, "g_nBackNeckSize= ", g_nBackNeckSize
'    Print #fOutPut, "g_sLeftSleeve= ", g_sLeftSleeve
'    Print #fOutPut, "g_sRightSleeve= ", g_sRightSleeve
'    Print #fOutPut, "g_sLeftLeg= ", g_sLeftLeg
'    Print #fOutPut, "g_sRightLeg= ", g_sRightLeg
'    Print #fOutPut, "g_sClosure= ", g_sClosure
'    Print #fOutPut, "g_sFabric= ", g_sFabric
'    Print #fOutPut, "g_sCrotchStyle= ", g_sCrotchStyle
'
'    Print #fOutPut,
'    Print #fOutPut, "Circumferences"
'    Print #fOutPut,
'    Print #fOutPut, "g_nShldWidth= ", g_nShldWidth
'    Print #fOutPut, "g_nLeftThighCir= ", g_nLeftThighCir
'    Print #fOutPut, "g_nRightThighCir= ", g_nRightThighCir
'    Print #fOutPut, "g_nShldToAxilla= ", g_nShldToAxilla
'    Print #fOutPut, "g_nAxillaCir= ", g_nAxillaCir
'    Print #fOutPut, "g_nThighCir= ", g_nThighCir
'    Print #fOutPut, "g_nLeftAxillaCir= ", g_nLeftAxillaCir
'    Print #fOutPut, "g_nRightAxillaCir= ", g_nRightAxillaCir
'    Print #fOutPut, "g_nChestCir= ", g_nChestCir
'    Print #fOutPut, "g_nUnderBreastCir= ", g_nUnderBreastCir
'    Print #fOutPut, "g_nWaistCir= ", g_nWaistCir
'    Print #fOutPut, "g_nButtocksCir= ", g_nButtocksCir
'    Print #fOutPut, "g_nNeckCir= ", g_nNeckCir
'    Print #fOutPut, "g_nShldToFold= ", g_nShldToFold
'    Print #fOutPut, "g_nShldToWaist= ", g_nShldToWaist
'    Print #fOutPut, "g_nShldToBreast= ", g_nShldToBreast
'    Print #fOutPut, "g_nButtocksLength= ", g_nButtocksLength
'    Print #fOutPut, "g_nButtBackSeamRatio= ", g_nButtBackSeamRatio
'    Print #fOutPut, "g_nButtFrontSeamRatio= ", g_nButtFrontSeamRatio
'    Print #fOutPut, "g_nButtRedIncreased= ", g_nButtRedIncreased
'    Print #fOutPut, "g_nThighRedDecreased= ", g_nThighRedDecreased
'    Print #fOutPut, "g_nButtCir= ", g_nButtCir
'    Print #fOutPut, "g_nHalfGroinHeight= ", g_nHalfGroinHeight
'    Print #fOutPut, "g_nButtRadius= ", g_nButtRadius
'    Print #fOutPut, "g_nButtCirCalc= ", g_nButtCirCalc
'    Print #fOutPut, "g_nButtFrontSeam= ", g_nButtFrontSeam
'    Print #fOutPut, "g_nButtBackSeam= ", g_nButtBackSeam
'    Print #fOutPut, "g_nButtCutOut= ", g_nButtCutOut
'    Print #fOutPut, "g_nWaistFrontSeam= ", g_nWaistFrontSeam
'    Print #fOutPut, "g_nWaistBackSeam= ", g_nWaistBackSeam
'    Print #fOutPut, "g_nWaistCutOut= ", g_nWaistCutOut
'    Print #fOutPut, "g_nChestFrontSeam= ", g_nChestFrontSeam
'    Print #fOutPut, "g_nChestBackSeam= ", g_nChestBackSeam
'    Print #fOutPut, "g_nChestCutOut= ", g_nChestCutOut
'
'    Print #fOutPut,
'    Print #fOutPut, "Reductions"
'    Print #fOutPut,
'    Print #fOutPut, "g_nChestCirRed= ", g_nChestCirRed
'    Print #fOutPut, "g_nUnderBreastCirRed= ", g_nUnderBreastCirRed
'    Print #fOutPut, "g_nWaistCirRed= ", g_nWaistCirRed
'    Print #fOutPut, "g_nLargeButtCirRed= ", g_nLargeButtCirRed
'    Print #fOutPut, "g_nThighCirRed= ", g_nThighCirRed
'    Print #fOutPut, "g_nNeckCirRed= ", g_nNeckCirRed
'    Print #fOutPut, "g_nShldToFoldRed= ", g_nShldToFoldRed
'    Print #fOutPut, "g_nShldToWaistRed= ", g_nShldToWaistRed
'    Print #fOutPut, "g_nShldToBreastRed= ", g_nShldToBreastRed
'    Print #fOutPut, "g_nButtocksCirRed= ", g_nButtocksCirRed
'
'    Print #fOutPut,
'    Print #fOutPut, "Co-ordinate points"
'    Print #fOutPut,
'    Print #fOutPut, "xySeamFold= ", xySeamFold.X, xySeamFold.Y
'    Print #fOutPut, "xySeamThigh= ", xySeamThigh.X, xySeamThigh.Y
'    Print #fOutPut, "xySeamButt= ", xySeamButt.X, xySeamButt.Y
'    Print #fOutPut, "xySeamHighShld= ", xySeamHighShld.X, xySeamHighShld.Y
'    Print #fOutPut, "xySeamLowShld= ", xySeamLowShld.X, xySeamLowShld.Y
'    Print #fOutPut, "xySeamChest= ", xySeamChest.X, xySeamChest.Y
'    Print #fOutPut, "xySeamWaist= ", xySeamWaist.X, xySeamWaist.Y
'    Print #fOutPut, "xyProfileThigh= ", xyProfileThigh.X, xyProfileThigh.Y
'    Print #fOutPut, "xyFold= ", xyFold.X, xyFold.Y
'    Print #fOutPut, "xyProfileButt= ", xyProfileButt.X, xyProfileButt.Y
'    Print #fOutPut, "xyCutOut3= ", xyCutout3.X, xyCutout3.Y
'    Print #fOutPut, "xyCutOut4= ", xyCutOut4.X, xyCutOut4.Y
'    Print #fOutPut, "xyCutOut5= ", xyCutOut5.X, xyCutOut5.Y
'    Print #fOutPut, "xyCutOut6= ", xyCutOut6.X, xyCutOut6.Y
'    Print #fOutPut, "xyCutOut2= ", xyCutOut2.X, xyCutOut2.Y
'    Print #fOutPut, "xyCutOut7= ", xyCutOut7.X, xyCutOut7.Y
'    Print #fOutPut, "xyProfileWaist= ", xyProfileWaist.X, xyProfileWaist.Y
'    Print #fOutPut, "xyProfileChest= ", xyProfileChest.X, xyProfileChest.Y
'    Print #fOutPut, "xyProfileNeck= ", xyProfileNeck.X, xyProfileNeck.Y
'    Print #fOutPut, "xyProfileNeckMirror= ", xyProfileNeckMirror.X, xyProfileNeckMirror.Y
'    Print #fOutPut, "xyCutOutBackNeck= ", xyCutOutBackNeck.X, xyCutOutBackNeck.Y
'    Print #fOutPut, "xyCutOut8= ", xyCutOut8.X, xyCutOut8.Y
 '   Print #fOutPut, "xyCutOutFrontNeck= ", xyCutOutFrontNeck.X, xyCutOutFrontNeck.Y
'    Print #fOutPut, "xyRaglan1= ", xyRaglan1.X, xyRaglan1.Y
'    Print #fOutPut, "xyRaglan2= ", xyRaglan2.X, xyRaglan2.Y
'    Print #fOutPut, "xyRaglan3= ", xyRaglan3.X, xyRaglan3.Y
'    Print #fOutPut, "xyRaglan4= ", xyRaglan4.X, xyRaglan4.Y
'    Print #fOutPut, "xyRaglan5= ", xyRaglan5.X, xyRaglan5.Y
'    Print #fOutPut, "xyRaglan6= ", xyRaglan6.X, xyRaglan6.Y
 
'    Close #fOutPut

End Sub

Private Sub Select_Text_In_Box(Text_Box_Name As Control)
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub

Private Sub Timer1_Timer()
    'It is assumed that the link open from Drafix has failed
    'Therefore we "End" here
     End
End Sub

Private Sub txtBackNeck_GotFocus()
    PR_Select_Text txtBackNeck
End Sub

Private Sub txtBackNeck_LostFocus()
    Dim nLen#
    nLen = FN_InchesValue(txtBackNeck)
    If nLen >= 0 Then
        lblBackNeck.Caption = fnInchestoText(nLen)
        g_sBackNeck = txtBackNeck
    Else
        lblBackNeck.Caption = ""
    End If
End Sub

Private Sub txtCir_GotFocus(Index As Integer)
    PR_Select_Text txtCir(Index)
End Sub

Private Sub txtCir_LostFocus(Index As Integer)
    Dim nLen#
    nLen = FN_InchesValue(txtCir(Index))
    If nLen > 0 Then
        lblCir(Index).Caption = fnInchestoText(nLen)
    Else
        lblCir(Index).Caption = ""
    End If
    
    If Index > 10 Or Index = 5 Then PR_EnableCalculateDiskButton
    

End Sub

Private Sub txtFrontNeck_GotFocus()
    PR_Select_Text txtFrontNeck
End Sub

Private Sub txtFrontNeck_LostFocus()
    Dim nLen#
    nLen = FN_InchesValue(txtFrontNeck)
    If nLen >= 0 Then
        lblFrontNeck.Caption = fnInchestoText(nLen)
        g_sFrontNeck = txtFrontNeck 'Keep this for restore
    Else
        lblFrontNeck.Caption = ""
    End If
End Sub

Private Sub txtLeftDisk_GotFocus()
    PR_Select_Text txtLeftDisk
End Sub

Private Sub txtLeftDisk_KeyPress(KeyAscii As Integer)
   'Only Disks sizes 1 to 9 allowed
    Select Case KeyAscii
        Case 32
            KeyAscii = 0
        Case 8, 9, 10, 13, 49 To 57
           'Do nothing
        Case Else
            KeyAscii = 0
            MsgBox "Bra disk sizes 1 to 9 only.", 48, "BodySuit - Dialogue"
    End Select

End Sub

Private Sub txtRightDisk_GotFocus()
    PR_Select_Text txtRightDisk
End Sub

Private Sub txtRightDisk_KeyPress(KeyAscii As Integer)
   'Only Disks sizes 1 to 9 allowed
    Select Case KeyAscii
        Case 32
            KeyAscii = 0
        Case 8, 9, 10, 13, 49 To 57
           'Do nothing
        Case Else
            KeyAscii = 0
            MsgBox "Bra disk sizes 1 to 9 only.", 48, "BodySuit - Dialogue"
    End Select
End Sub

Private Sub Validate_Text(Text_Box_Name As Control)
   'Subroutine that is activated when the focus is lost.
    Dim rTextBoxValue As Double
    Dim ii As Integer, iTest As Integer

   'Checks that input data is valid.
   'If not valid then display a message and returns focus
   'to the text in question

   'Get the text value
    rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))

   'Only if relevant units are Inches
    If rTextBoxValue < 0 Then
        MsgBox "Invalid Format for inches", 48, "Data input Error"
        Text_Box_Name.SetFocus
        Exit Sub
    End If


   'Check that each character is numeric or a decimal point
   'N.B.
   '    Asc("0") = 48
   '    Asc("9") = 57
   '    Asc(".") = 46
   '
    For ii = 1 To Len(Text_Box_Name)
        iTest = Asc(Mid$(Text_Box_Name.Text, ii, 1))
        If (iTest < 48 Or iTest > 57) And iTest <> 46 Then
            rTextBoxValue = -1
        End If
    Next ii
    
    If rTextBoxValue < 0 Then
        MsgBox "Invalid or Negative value given", 48, "Data input Error"
        Text_Box_Name.SetFocus
    End If

    If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
        MsgBox "Zero value given", 48, "Data input Error"
        Text_Box_Name.SetFocus
    End If

    If rTextBoxValue > 999.9 Then
        MsgBox "Given value too Large", 48, "Data input Error"
        Text_Box_Name.SetFocus
    End If

End Sub

Private Sub Validate_Text_In_Box(Text_Box_Name As Control, Label_Name As Control)
   'Subroutine that is activated when the focus is lost.
    Dim rTextBoxValue As Double, rDec As Double
    Dim ii As Integer, iTest As Integer
    Dim iEighths As Integer, iInt As Integer
    Dim sString As String
   
   'Checks that input data is valid.
   'If not valid then display a message and returns focus
   'to the text in question

   'Get the text value
    rTextBoxValue = fnDisplayToInches(Val(Text_Box_Name.Text))

   'Only relevant if units are Inches
    If rTextBoxValue < 0 Then
        MsgBox "Invalid Format for inches", 48, "Data input Error"
        Text_Box_Name.SetFocus
        Exit Sub
    End If

   'Check that each character is numeric or a decimal point
   'N.B.
   '    Asc("0") = 48
   '    Asc("9") = 57
   '    Asc(".") = 46
   '
    For ii = 1 To Len(Text_Box_Name)
        iTest = Asc(Mid$(Text_Box_Name.Text, ii, 1))
        If (iTest < 48 Or iTest > 57) And iTest <> 46 Then
            rTextBoxValue = -1
        End If
    Next ii
    
    If rTextBoxValue < 0 Then
        MsgBox "Invalid or Negative value given", 48, "Data input Error"
        Text_Box_Name.SetFocus
        Exit Sub
    End If

    If rTextBoxValue = 0 And Len(Text_Box_Name.Text) > 0 Then
        MsgBox "Zero value given", 48, "Data input Error"
        Text_Box_Name.SetFocus
        Exit Sub
    End If

    If rTextBoxValue > 999.9 Then
        MsgBox "Given value too Large", 48, "Data input Error"
        Text_Box_Name.SetFocus
        Exit Sub
    End If

    Label_Name.Caption = fnInchestoText(rTextBoxValue)

End Sub


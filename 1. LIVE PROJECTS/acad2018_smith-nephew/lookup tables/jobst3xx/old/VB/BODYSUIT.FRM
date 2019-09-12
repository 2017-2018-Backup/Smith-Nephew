VERSION 5.00
Begin VB.Form bodysuit 
   Caption         =   "Body Brief/Suit - Dialogue"
   ClientHeight    =   8160
   ClientLeft      =   1950
   ClientTop       =   780
   ClientWidth     =   9600
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
   LinkTopic       =   "bodysuit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8160
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdTab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   147
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton btnSave 
      Cancel          =   -1  'True
      Caption         =   "Save"
      Height          =   375
      Left            =   2700
      TabIndex        =   146
      Top             =   7560
      Width           =   1275
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   12
      Left            =   9840
      TabIndex        =   145
      Text            =   "Thigh Red."
      Top             =   6075
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   11
      Left            =   9840
      TabIndex        =   144
      Text            =   "Lg. Butt. Red."
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCrotchStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11040
      TabIndex        =   139
      Top             =   7290
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11040
      TabIndex        =   138
      Top             =   7650
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9840
      Top             =   6960
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   9840
      TabIndex        =   115
      Text            =   "Fabric"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   9840
      TabIndex        =   114
      Text            =   "Closure"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   9840
      TabIndex        =   113
      Text            =   "Back Neck"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   9840
      TabIndex        =   112
      Text            =   "Front Neck"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   9840
      TabIndex        =   111
      Text            =   "Right Axilla"
      Top             =   4260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   9840
      TabIndex        =   110
      Text            =   "Left Axilla"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   9840
      TabIndex        =   109
      Text            =   "Right Bracup"
      Top             =   3660
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   9840
      TabIndex        =   108
      Text            =   "Left Bra Cup"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   9840
      TabIndex        =   107
      Text            =   "Nipple Reduction"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   9840
      TabIndex        =   106
      Text            =   "Chest reduction"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   9840
      TabIndex        =   105
      Text            =   "Waist Reduction"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidVB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9840
      TabIndex        =   102
      Text            =   "txtUidVB"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9840
      TabIndex        =   100
      Text            =   "txtWorkOrder"
      Top             =   930
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10140
      TabIndex        =   99
      Text            =   "txtUidMPD"
      Top             =   675
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Design Choices"
      Height          =   5955
      Left            =   6060
      TabIndex        =   97
      Top             =   2115
      Width           =   3435
      Begin VB.Frame Frame7 
         Caption         =   "Left Leg"
         Height          =   1020
         Left            =   420
         TabIndex        =   142
         Top             =   4215
         Width           =   1035
         Begin VB.OptionButton optLeftLeg 
            Caption         =   "Panty"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton optLeftLeg 
            Caption         =   "Brief"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   143
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.ComboBox cboLegStyle 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3855
         Width           =   1995
      End
      Begin VB.ComboBox cboCrotchStyle 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   5415
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Right Leg"
         Height          =   1035
         Left            =   2040
         TabIndex        =   117
         Top             =   4215
         Width           =   1035
         Begin VB.OptionButton optRightLeg 
            Caption         =   "Panty"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   119
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton optRightLeg 
            Caption         =   "Brief"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Text            =   "cboFabric"
         Top             =   3300
         Width           =   3075
      End
      Begin VB.ComboBox cboClosure 
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Text            =   "cboClosure"
         Top             =   2700
         Width           =   1935
      End
      Begin VB.TextBox txtBackNeck 
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2310
         Width           =   555
      End
      Begin VB.ComboBox cboBackNeck 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Text            =   "cboBackNeck"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtFrontNeck 
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1530
         Width           =   555
      End
      Begin VB.ComboBox cboFrontNeck 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Text            =   "cboFrontNeck"
         Top             =   1140
         Width           =   1935
      End
      Begin VB.ComboBox cboRightAxilla 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Text            =   "cboRightAxilla"
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cboLeftAxilla 
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Text            =   "cboLeftAxilla"
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Leg Style:"
         Height          =   195
         Left            =   180
         TabIndex        =   121
         Top             =   3915
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Crotch Style:"
         Height          =   195
         Left            =   180
         TabIndex        =   120
         Top             =   5475
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fabric:"
         Height          =   285
         Index           =   33
         Left            =   180
         TabIndex        =   72
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Closure:"
         Height          =   285
         Index           =   32
         Left            =   180
         TabIndex        =   73
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label labBackNeck 
         Caption         =   "Back Neck size:"
         Height          =   285
         Left            =   180
         TabIndex        =   74
         Top             =   2340
         Width           =   1995
      End
      Begin VB.Label lblBackNeck 
         Caption         =   "inches"
         Height          =   165
         Left            =   2340
         TabIndex        =   75
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Back Neck:"
         Height          =   285
         Index           =   30
         Left            =   180
         TabIndex        =   76
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label labFrontNeck 
         Caption         =   "Front Neck size:"
         Height          =   285
         Left            =   180
         TabIndex        =   77
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label lblFrontNeck 
         Caption         =   "inches"
         Height          =   165
         Left            =   2340
         TabIndex        =   78
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Front Neck:"
         Height          =   285
         Index           =   28
         Left            =   180
         TabIndex        =   79
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Right Axilla:"
         Height          =   285
         Index           =   25
         Left            =   180
         TabIndex        =   80
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Left Axilla:"
         Height          =   285
         Index           =   27
         Left            =   180
         TabIndex        =   81
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   26
         Left            =   1980
         TabIndex        =   98
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame frmBra 
      Caption         =   "Measurements for Bra"
      Height          =   2895
      Left            =   120
      TabIndex        =   63
      Top             =   4560
      Width           =   5835
      Begin VB.CommandButton cmdCalculateBraDisks 
         Caption         =   "Calculate Disks"
         Height          =   375
         Left            =   2100
         TabIndex        =   116
         Top             =   2430
         Width           =   1455
      End
      Begin VB.TextBox txtRightDisk 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4680
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cboRightCup 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1995
         Width           =   1035
      End
      Begin VB.TextBox txtLeftDisk 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cboLeftCup 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1995
         Width           =   1035
      End
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   37
         Text            =   "cboRed"
         Top             =   690
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   11
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   14
         Top             =   990
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   10
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   13
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtCir 
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
         Caption         =   "Disk:"
         Height          =   285
         Index           =   5
         Left            =   4680
         TabIndex        =   96
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Cup:"
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   95
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Disk:"
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   94
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Cup:"
         Height          =   285
         Index           =   2
         Left            =   660
         TabIndex        =   93
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label LabBra 
         Alignment       =   2  'Center
         Caption         =   "Right Bra Cup"
         Height          =   285
         Index           =   1
         Left            =   3540
         TabIndex        =   92
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Alignment       =   2  'Center
         Caption         =   "Left Bra Cup"
         Height          =   285
         Index           =   0
         Left            =   660
         TabIndex        =   91
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Caption         =   "Circ. over nipple"
         Height          =   285
         Index           =   10
         Left            =   180
         TabIndex        =   90
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Caption         =   "C"
         Height          =   285
         Index           =   11
         Left            =   3300
         TabIndex        =   89
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   11
         Left            =   4260
         TabIndex        =   88
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label LabBra 
         Caption         =   "Circ. under breast"
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   87
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Caption         =   "B"
         Height          =   285
         Index           =   9
         Left            =   3300
         TabIndex        =   86
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   10
         Left            =   4260
         TabIndex        =   85
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LabBra 
         Caption         =   "Shoulder to under breast"
         Height          =   405
         Index           =   6
         Left            =   180
         TabIndex        =   84
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Caption         =   "A"
         Height          =   285
         Index           =   7
         Left            =   3300
         TabIndex        =   83
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   9
         Left            =   4260
         TabIndex        =   82
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Measurments for Body"
      Height          =   4425
      Left            =   120
      TabIndex        =   62
      Top             =   45
      Width           =   5835
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   4
         Left            =   4920
         TabIndex        =   68
         Text            =   "cboRed"
         Top             =   3780
         Width           =   850
      End
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   3
         Left            =   4920
         TabIndex        =   69
         Text            =   "cboRed"
         Top             =   3293
         Width           =   850
      End
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   70
         Text            =   "cboRed"
         Top             =   1973
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   12
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2625
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   13
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   8
         Top             =   2955
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   14
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   9
         Top             =   3285
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   15
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   10
         Top             =   3615
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   16
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   11
         Top             =   3945
         Width           =   555
      End
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   71
         Text            =   "cboRed"
         Top             =   2303
         Width           =   850
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   6
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   6
         Top             =   2295
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   5
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1965
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   4
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1635
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   3
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1305
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   2
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   975
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   1
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   1
         Top             =   645
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   0
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "0"
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Proximal Thigh Circ. - Right"
         Height          =   285
         Index           =   20
         Left            =   180
         TabIndex        =   122
         Top             =   4005
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Proximal Thigh Circ. - Left"
         Height          =   285
         Index           =   38
         Left            =   180
         TabIndex        =   123
         Top             =   3675
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Circ. of Lg. part of Buttocks"
         Height          =   285
         Index           =   24
         Left            =   180
         TabIndex        =   124
         Top             =   3345
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder to Lg. part of Buttocks"
         Height          =   285
         Index           =   21
         Left            =   180
         TabIndex        =   125
         Top             =   3015
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder to Fold of Buttocks"
         Height          =   285
         Index           =   42
         Left            =   180
         TabIndex        =   126
         Top             =   2685
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "8"
         Height          =   285
         Index           =   41
         Left            =   3300
         TabIndex        =   127
         Top             =   2685
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   12
         Left            =   4260
         TabIndex        =   128
         Top             =   2685
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "9"
         Height          =   285
         Index           =   39
         Left            =   3300
         TabIndex        =   129
         Top             =   3015
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   13
         Left            =   4260
         TabIndex        =   130
         Top             =   3015
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "10"
         Height          =   285
         Index           =   37
         Left            =   3300
         TabIndex        =   131
         Top             =   3345
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   14
         Left            =   4260
         TabIndex        =   132
         Top             =   3345
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "11"
         Height          =   285
         Index           =   29
         Left            =   3300
         TabIndex        =   133
         Top             =   3675
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   15
         Left            =   4260
         TabIndex        =   134
         Top             =   3675
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "12"
         Height          =   285
         Index           =   23
         Left            =   3300
         TabIndex        =   135
         Top             =   4005
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   16
         Left            =   4260
         TabIndex        =   136
         Top             =   4005
         Width           =   600
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   6
         Left            =   4260
         TabIndex        =   66
         Top             =   2355
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "7"
         Height          =   285
         Index           =   14
         Left            =   3300
         TabIndex        =   65
         Top             =   2355
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Waist circ."
         Height          =   285
         Index           =   13
         Left            =   180
         TabIndex        =   64
         Top             =   2355
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   5
         Left            =   4260
         TabIndex        =   61
         Top             =   2025
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "6"
         Height          =   285
         Index           =   12
         Left            =   3300
         TabIndex        =   60
         Top             =   2025
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Chest circ."
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   58
         Top             =   2025
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   4
         Left            =   4260
         TabIndex        =   56
         Top             =   1695
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "5"
         Height          =   285
         Index           =   10
         Left            =   3300
         TabIndex        =   54
         Top             =   1695
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder to Waist"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   52
         Top             =   1695
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   3
         Left            =   4260
         TabIndex        =   50
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "4"
         Height          =   285
         Index           =   8
         Left            =   3300
         TabIndex        =   48
         Top             =   1365
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder width"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   47
         Top             =   1365
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   2
         Left            =   4260
         TabIndex        =   45
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "3"
         Height          =   285
         Index           =   5
         Left            =   3300
         TabIndex        =   44
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Neck circ."
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   43
         Top             =   1035
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   1
         Left            =   4260
         TabIndex        =   42
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         Height          =   285
         Index           =   3
         Left            =   3300
         TabIndex        =   41
         Top             =   705
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Right shoulder circ."
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   40
         Top             =   705
         Width           =   1815
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   0
         Left            =   4260
         TabIndex        =   39
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         Height          =   285
         Index           =   1
         Left            =   3300
         TabIndex        =   38
         Top             =   375
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Left shoulder circ."
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   36
         Top             =   375
         Width           =   1635
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   46
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   6060
      TabIndex        =   15
      Top             =   45
      Width           =   3435
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   55
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2580
         TabIndex        =   51
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2400
         TabIndex        =   59
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   57
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   49
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   53
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   1920
         TabIndex        =   21
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1620
         Width           =   375
      End
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Crotch Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9780
      TabIndex        =   141
      Top             =   7350
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9780
      TabIndex        =   140
      Top             =   7710
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shoulder to Fold"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   19
      Left            =   600
      TabIndex        =   137
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
      TabIndex        =   104
      Top             =   2250
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vest Body"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   9960
      TabIndex        =   103
      Top             =   1500
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   9960
      TabIndex        =   101
      Top             =   420
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "bodysuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   BODYSUIT.MAK
'Purpose:   Body Dialogue
'
'
'Version:   3.00
'Date:      3.May.96
'Author:    Paul O'Rawe
'           Gary George
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
   
   
'   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

   'MsgBox constant
'    Const IDCANCEL = 2
'    Const IDYES = 6
'    Const IDNO = 7



Private Sub btnCancel_Click()
    
    Dim Response As Integer, sTask$, sCurrentValues$
    
   'Check if data has been modified
    sCurrentValues = FN_ValuesString()

    If sCurrentValues <> g_sChangeChecker Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "Bodysuit - Glove Dialogue")
        Select Case Response
            Case IDYES
                PR_CreateMacro_Save "c:\jobst\draw.d"
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate sTask
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "Bodysuit - Dialogue"
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

Private Sub btnSave_Click()
    
    Dim Response As Integer, sTask$, sCurrentValues$
    
   'Check if data has been modified
    sCurrentValues = FN_ValuesString()
    If sCurrentValues <> g_sChangeChecker Then
        'Check all values are present
        PR_GetValuesFromDialogue
        If FN_ValidateAndCalculateData(True) Then 'else stay in dialog
            
            PR_CreateMacro_Save "c:\jobst\draw.d"
            sTask = fnGetDrafixWindowTitleText()
            If sTask <> "" Then
                AppActivate sTask
                SendKeys "@c:\jobst\draw.d{enter}", True
                End
            Else
                MsgBox "Can't find a Drafix Drawing to update!", 16, "BodySuit - Dialogue"
            End If
        End If
    Else
        End
    End If

End Sub

Private Sub cboBackNeck_Click()
    Select Case Left$(cboBackNeck.Text, 1)
        Case "R", "S"   'Regular or Scoop
            txtBackNeck = ""
            lblBackNeck.Caption = ""
            txtBackNeck.Enabled = False
            labBackNeck.Enabled = False
        Case Else  'Measured Scoop
            If g_sBackNeck <> "" Then
                txtBackNeck = g_sFrontNeck
                txtBackNeck_LostFocus   'Display inches
            End If
            txtBackNeck.Enabled = True
            labBackNeck.Enabled = True
    End Select
End Sub

Private Sub cboCrotchStyle_Click()
    Dim sString As String, sError As String
    Dim iindex As Integer
    
    sError = ""

   'Check Crotch style against sex
    sString = cboCrotchStyle.List(cboCrotchStyle.ListIndex)

   'Check for male crotch and female patient
    iindex = 0
    iindex = InStr(1, sString, "Fly", 0) + InStr(1, sString, "Male", 0)
    If iindex <> 0 And txtSex = "Female" Then
        sError = sError + "Male Crotch type specified for Female patient" + NL
    End If

   'Check for female crotch and male patient
    iindex = 0
    iindex = InStr(1, sString, "Female", 0)
    If iindex <> 0 And txtSex = "Male" Then
        sError = sError + "Female Crotch type specified for Male patient" + NL
    End If


   'Display Error message (if required) and return
   'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 0, "Crotch Style Error"
    End If

End Sub

Private Sub cboFrontNeck_Click()
    Select Case Left$(cboFrontNeck.Text, 1)
        Case "R", "S", "V"   'Regular or Scoop or V neck
            txtFrontNeck = ""
            lblFrontNeck.Caption = ""
            txtFrontNeck.Enabled = False
            labFrontNeck.Enabled = False
        Case Else   'Measured Scoop, Measured V or Turtle neck
            If g_sFrontNeck <> "" Then
                txtFrontNeck = g_sFrontNeck
                txtFrontNeck_LostFocus   'Display inches
            End If
            txtFrontNeck.Enabled = True
            labFrontNeck.Enabled = True
    End Select
End Sub

Private Sub cboLeftAxilla_Click()
    If cboLeftAxilla.Text = "Sleeveless" Then
        cboRightAxilla.Text = "Sleeveless"
    Else
        If cboRightAxilla.Text = "Sleeveless" Then
             cboRightAxilla.Text = cboLeftAxilla.Text
        End If
    End If
End Sub

Private Sub cboLeftCup_Click()
    If bodysuit.Visible And (cboLeftCup.Text = "" Or cboLeftCup.Text = "None") Then
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
    
    
   'Get the currently selected leg style
    sLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)
    If sLegStyle = g_sPreviousLegStyle Then Exit Sub

    Select Case sLegStyle
     Case "Panty"
        optLeftLeg(0).Enabled = True
        optRightLeg(0).Enabled = True

        optLeftLeg(0).Value = 1     'Left leg Panty
        optLeftLeg(1).Value = 0     'Left leg Brief

        optRightLeg(0).Value = 1    'Right leg Panty
        optRightLeg(1).Value = 0    'Right leg Brief


     Case "Brief", "Brief-French"
        optLeftLeg(0).Value = 0     'Left leg Panty
        optLeftLeg(0).Enabled = False
        optLeftLeg(1).Value = 1     'Left leg Brief

        optRightLeg(0).Value = 0    'Right leg Panty
        optRightLeg(0).Enabled = False
        optRightLeg(1).Value = 1    'Right leg Brief

        optLeftLeg(1).Enabled = True    'We enable just in case as
        optRightLeg(1).Enabled = True   '
    
    End Select

   'Store leg style for use when it changes
    g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)

End Sub

Private Sub cboRightAxilla_Click()
    If cboRightAxilla.Text = "Sleeveless" Then
        cboLeftAxilla.Text = "Sleeveless"
    Else
        If cboLeftAxilla.Text = "Sleeveless" Then
             cboLeftAxilla.Text = cboRightAxilla.Text
        End If
    End If
End Sub

Private Sub cboRightCup_Click()
    If bodysuit.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
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

Private Sub cmdTab_Click()
   'Allows the user to use enter as a tab
    SendKeys "{TAB}", False
End Sub

Private Function FN_CalcLength(xyStart As XY, xyEnd As XY) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

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
        If Char = QQ Then
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
    Print #fNum, "//by Visual Basic, BodySuit"
    Print #fNum, "//type - "; sType; ""


End Function



Private Function Fn_ValidateData() As Integer
'This function is used only to make gross checks
'for missing data.
'It does not perform any sensibility checks on the
'data
'Check CutOut, Back Seams and Shoulder width to see
'if any changes are required - if there is the dialog
'is updated.

    Dim sError As String
    Dim iFatalError As Integer
    Dim nShoulderWidthGiven As Double
    Dim nChestCirGiven As Double
    Dim nWaistCirGiven As Double
    Dim nShldToFoldGiven As Double
    Dim nButtocksCirGiven As Double
    Dim nWaistCirRed As Double
    Dim nChestCirRed As Double
    Dim nButtocksCirRed As Double
    Dim nThighCirRed As Double
    Dim sCrotchStyle As String
    Dim nSmallestThighGiven As Double
    Dim nWaistRem As Double
    Dim nChestRem As Double
    Dim nChestCir As Double
    Dim nWaistCir As Double
    Dim nButtocksCir As Double
    Dim nGroinHeight As Double
    Dim nShldToFold As Double
    Dim nShldToWaist As Double
    Dim nButtocksLength As Double
    Dim nButtBackSeamRatio As Double
    Dim nButtFrontSeamRatio As Double
    Dim nButtRedAtLimit As Integer
    Dim nThighRedAtLimit As Integer
    Dim nButtRedIncreased As Integer
    Dim nThighRedDecreased As Integer
    Dim nHalfGroinHeight As Double
    Dim nButtRadius As Double
    Dim nButtCirCalc As Double
    Dim nButtBackSeam As Double
    Dim nButtFrontSeam As Double
    Dim nButtCutOut As Double
    Dim nWaistFrontSeam As Double
    Dim nChestFrontSeam As Double
    Dim nWaistCutOut As Double
    Dim nWaistCutOutModified As Integer

    'Initialise
    Fn_ValidateData = False
    sError = ""
    iFatalError = False
    
    'Get values from dialog
    nShldWidthGiven = fnDisplayToInches(Val(txtCir(3).Text))
    nShldToWaistGiven = fnDisplayToInches(Val(txtCir(4).Text))
    nChestCirGiven = fnDisplayToInches(Val(txtCir(5).Text))
    nWaistCirGiven = fnDisplayToInches(Val(txtCir(6).Text))
    nShldToFoldGiven = fnDisplayToInches(Val(txtCir(12).Text))
    nButtocksCirGiven = fnDisplayToInches(Val(txtCir(14).Text))
    nLeftThighCirGiven = fnDisplayToInches(Val(txtCir(15).Text))
    nRightThighCirGiven = fnDisplayToInches(Val(txtCir(16).Text))
    sCrotchStyle = cboCrotchStyle.Text
    
    'Reductions
    nShldToFoldRed = 0.95             '95%
    nShldToWaistRed = 0.95            '95%
    If cboRed(1).Text <> "" Then
            nChestCirRed = Val(cboRed(1).Text)
        Else
            nChestCirRed = 0.85 '85% default
    End If
    If cboRed(0).Text <> "" Then
            nWaistCirRed = Val(cboRed(0).Text)
        Else
            nWaistCirRed = 0.85 '85% default
    End If
    If cboRed(3).Text <> "" Then
            nButtocksCirRed = Val(cboRed(3).Text)
        Else
            nButtocksCirRed = 0.85 '85% default
    End If
    If cboRed(4).Text <> "" Then
            nThighCirRed = Val(cboRed(4).Text)
        Else
            nThighCirRed = 0.85   '85% default
    End If
    
    'Test minimum value for Shoulder width
    If nShldWidthGiven < 0.5 Then
        sError = sError & "Shoulder Width is less than 1/2""!" & NL
    End If
    
    ReDim sCircum(0 To 16)
    sCircum(0) = "Left shoulder circ."
    sCircum(1) = "Right shoulder circ."
    sCircum(2) = "Neck circ."
    sCircum(3) = "Shoulder width."
    sCircum(4) = "Shoulder to waist."
    sCircum(5) = "Chest circ."
    sCircum(6) = "Waist circ."
    sCircum(9) = "Shoulder to under breast."
    sCircum(10) = "Circ. under breast."
    sCircum(11) = "Circ. over nipple."
    sCircum(12) = "Shoulder to Fold of Buttocks."
    sCircum(13) = "Shoulder to Large Part of Buttocks."
    sCircum(14) = "Circ. of Large Part of Buttocks."
    sCircum(15) = "Left Thigh Circ."
    sCircum(16) = "Right Thigh Circ."

    'FATAL Errors
    '~~~~~~~~~~~~
    'Body measurements (all must be present)
    For ii = 0 To 6
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing " & sCircum(ii) & NL
        End If
    Next ii
    For ii = 12 To 16
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing " & sCircum(ii) & NL
        End If
    Next ii

    'Bra Cups
    'Note:
    '    The Circumference over nipple is optional unless a cup has been
    '    specified
    '
    If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(9) & NL
    End If
    If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
            sError = sError & "Missing " & sCircum(10) & NL
    End If

    If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
            sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated!" & NL
    End If

    If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
            sError = sError & "Bra disks given! But Missing " & sCircum(9) & NL
    End If

    'If cups or dimensions given then a disk must be present
    'NB
    '    cboXXXXCup.ListIndex = 6 = "None"
    '    cboXXXXCup.ListIndex = 7 = ""

    If cboLeftCup.ListIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Left BRA Cup!" & NL
    End If
    If cboRightCup.ListIndex < 6 And Val(txtRightDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Right BRA Cup!" & NL
    End If
    
   'Sex error
    If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.ListIndex < 0 Or cboRightCup.ListIndex < 0) Then
            sError = sError & "Male patient but Bra Measurements or Bra Cups requested!" & NL
    End If

   'Neck at back and front
    Dim sChar As String * 1
    sChar = Left$(cboBackNeck.Text, 1)
    If sChar = "M" And txtBackNeck.Text = "" Then
            sError = sError & "No dimension for Back Meck Measured Scoop!" & NL
    End If

    sChar = Left$(cboFrontNeck.Text, 1)
    If sChar = "M" And txtFrontNeck.Text = "" Then
            sError = sError & "No dimension for Front Neck Measured Scoop!" & NL
    End If

   '
    If cboLeftAxilla.Text = "" Then
            sError = sError & "Left Axilla not given!" & NL
    End If

    If cboRightAxilla.Text = "" Then
            sError = sError & "Right Axilla not given!" & NL
    End If

    If cboFrontNeck.Text = "" Then
            sError = sError & "Neck not given!" & NL
    End If

    If cboBackNeck.Text = "" Then
            sError = sError & "Back neck not given!" & NL
    End If

    If cboClosure.Text = "" Then
            sError = sError & "Closure not given!" & NL
    End If

    If cboFabric.Text = "" Then
            sError = sError & "Fabric not given!" & NL
    End If
    
    'Extra Values
    If cboLegStyle.Text = "" Then
            sError = sError & "Leg Style not given!" & NL
    End If
    If cboCrotchStyle.Text = "" Then
            sError = sError & "Crotch Style not given!" & NL
    End If
    
    'Display Error message (if required) and return
    'These are fatal errors
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        Fn_ValidateData = False
        Exit Function
    Else
        Fn_ValidateData = True
    End If

    'Possible FATAL Errors and WARNINGS
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    'Figured Variables for tests
    'Note: Shoulder Width & Nipple Circumference not figured
    nChestCir = fnRoundInches((nChestCirGiven * nChestCirRed)) / 2   'half scale
    nWaistCir = fnRoundInches((nWaistCirGiven * nWaistCirRed)) / 2   'half-scale
    nShldToFold = fnRoundInches((nShldToFoldGiven * nShldToFoldRed)) + INCH1_2
    nShldToWaist = fnRoundInches((nShldToWaistGiven * nShldToWaistRed)) + INCH1_2
    nButtocksLength = fnRoundInches((nShldToFold - nShldToWaist) / 3)
    
    'Test Measurements
    
    'NOTE:-
    '    There can be no more that 5% difference between the reductions
    '    at each circumference.

    'Use smallest thigh size.
    If (nLeftThighCirGiven < nRightThighCirGiven) Then
            nSmallestThighGiven = nLeftThighCirGiven
        Else
            nSmallestThighGiven = nRightThighCirGiven
    End If
    
    If InStr(sCrotchStyle, "Fly") <> 0 Then
            nButtBackSeamRatio = 0.6 '60%
            nButtFrontSeamRatio = 0.4 '40%
        Else
            nButtBackSeamRatio = 0.5 '50%
            nButtFrontSeamRatio = 0.5 '50%
    End If

    'CutOut and Back Seam Test
    nButtRedAtLimit = False
    nThighRedAtLimit = False
    nButtRedIncreased = False
    nThighRedDecreased = False
    Do
        If nButtocksCirRed = 0.9 Then 'at Reduction limit of 90%
            nButtRedAtLimit = True
        End If
        'I had to use Val and Str$ functions because VB3 has a problem
        'when returning values from the minus operation,
        'i.e. nThighCirRed was not equal to .8 exactly (but it should be!!!)
        If Val(Str$(nThighCirRed)) = 0.8 Then 'at Reduction limit of 80%
                nThighRedAtLimit = True
        End If
        If nThighRedAtLimit And nButtRedAtLimit Then
            Exit Do
        End If
        nButtocksCir = fnRoundInches((nButtocksCirGiven * nButtocksCirRed)) / 2  'half scale
        nGroinHeight = fnRoundInches((nSmallestThighGiven * nThighCirRed)) / 2    'half scale
        nHalfGroinHeight = (nGroinHeight / 2)
        nButtRadius = Sqr((nButtocksLength ^ 2) + (nHalfGroinHeight ^ 2))
        nButtCirCalc = nButtRadius + nHalfGroinHeight
        nButtBackSeam = (nButtocksCir - nButtCirCalc) * nButtBackSeamRatio
        If nButtBackSeam < 1.5 Then     'less than 3" full scale
                If nButtRedAtLimit Then
                        nThighCirRed = nThighCirRed - 0.05    'decrease the reduction by 5%
                        nThighRedDecreased = True
                    Else
                        nButtocksCirRed = nButtocksCirRed + 0.05 'increase the reduction by 5%
                        nButtRedIncreased = True
                End If
            Else
                Exit Do     'greater than 3" full scale - acceptable
        End If
    Loop
    If nButtRedIncreased Then
        cboRed(3).Text = nButtocksCirRed
        sError = sError + "With given Buttock Reduction, distance from back of Cut-Out" + NL
        sError = sError + "to largest part of buttocks is less than 3 inches!" + NL
        sError = sError + "Buttocks Reduction increased to rectify this." + NL
        'sError = sError + "ref:- " + NL
    End If
    If nThighRedDecreased Then
        cboRed(4).Text = nThighCirRed
        sError = sError + "With given Thigh Reduction, distance from back of Cut-Out" + NL
        sError = sError + "to largest part of buttocks is less than 3 inches!" + NL
        sError = sError + "Thigh Reduction decreased to rectify this." + NL
        'sError = sError + "ref:- " + NL
    End If
    'One last time with Reductions
    nButtocksCir = fnRoundInches((nButtocksCirGiven * nButtocksCirRed)) / 2  'half scale
    nGroinHeight = fnRoundInches((nSmallestThighGiven * nThighCirRed)) / 2    'half scale
    nHalfGroinHeight = (nGroinHeight / 2)
    nButtRadius = Sqr((nButtocksLength ^ 2) + (nHalfGroinHeight ^ 2))
    nButtCirCalc = nButtRadius + nHalfGroinHeight
    nButtBackSeam = (nButtocksCir - nButtCirCalc) * nButtBackSeamRatio
    If nButtBackSeam <= 0 Then
            sError = sError + NL + "Severe - Warning!" + NL
            sError = sError + "Distance from back of Cut-Out to largest part of buttocks" + NL
            sError = sError + "is still negative or zero!" + NL
            sError = sError + "Even with the modified reductions." + NL
            iFatalError = True
        ElseIf nButtBackSeam < 1.5 Then     'less than 3" full scale
            sError = sError + NL + "Severe - Warning!" + NL
            sError = sError + "Distance from back of Cut-Out to largest part of buttocks" + NL
            sError = sError + "is still less than 3 inches!" + NL
            sError = sError + "Even with the modified reductions." + NL
    End If
   
    'All other seams test
    nButtFrontSeam = (nButtocksCir - nButtCirCalc) * nButtFrontSeamRatio
    nButtCutOut = nButtCirCalc - (nButtBackSeam + nButtFrontSeam)
    If (nWaistCirGiven < nButtocksCirGiven) Then
            nWaistFrontSeam = Abs(((nButtocksCirGiven - nWaistCirGiven) / 2) - nButtFrontSeam)
        Else
            nWaistFrontSeam = nButtFrontSeam + (Int(nWaistCirGiven - nButtocksCirGiven) * INCH1_8)
    End If
    If (nChestCirGiven < nButtocksCirGiven) Then
            nChestFrontSeam = Abs(((nButtocksCirGiven - nChestCirGiven) / 2) - nButtFrontSeam)
        Else
            nChestFrontSeam = nButtFrontSeam + (Int(nChestCirGiven - nButtocksCirGiven) * INCH1_8)
    End If
    nWaistCutOut = (nButtCutOut * 0.87)
    nWaistCutOutModified = 0
    Do
        nWaistRem = nWaistCir - (nWaistCutOut + (nWaistFrontSeam * 2))
        nChestRem = nChestCir - (nWaistCutOut + nWaistFrontSeam + nChestFrontSeam)
        If nWaistRem < 1.5 Or nChestRem < 1.5 Then  'less than 3" @ full scale
                If nWaistRem <= 0 Then
                    sError = sError + NL + "Severe - Warning!" + NL
                    sError = sError + "Distance from back of Cut-Out to Waist" + NL
                    sError = sError + "is negative or zero!" + NL
                    If nWaistCutOutModified > 0 Then
                        sError = sError + "Even with the Waist CutOut decreased." + NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                If nChestRem <= 0 Then
                    sError = sError + NL + "Severe - Warning!" + NL
                    sError = sError + "Distance from back of Cut-Out to Chest" + NL
                    sError = sError + "is negative or zero!" + NL
                    If nWaistCutOutModified > 0 Then
                        sError = sError + "Even with the Waist CutOut decreased." + NL
                    End If
                    iFatalError = True
                    Exit Do
                End If
                'Decreae Waist CutOut
                nWaistCutOut = nWaistCutOut - (2 * (1.5 - nWaistRem))
                nWaistCutOutModified = nWaistCutOutModified + 1
                If nWaistCutOut <= 0 Then
                    sError = sError + NL + "Severe - Warning!" + NL
                    sError = sError + "Distance from back of Cut-Out to Waist or Chest" + NL
                    sError = sError + "is less than 3 inches!" + NL
                    Select Case nWaistCutOutModified
                        Case 1
                            sError = sError & "Waist CutOut had to be decreased " & Str$(nWaistCutOutModified) & " time" & NL
                            'sError = sError + "ref:- " + NL
                        Case Is > 1
                            sError = sError & "Waist CutOut had to be decreased " & Str$(nWaistCutOutModified) & " times" & NL
                            'sError = sError + "ref:- " + NL
                    End Select
                    sError = sError + "but is still negative or zero!" + NL
                    iFatalError = True
                    Exit Do
                End If
            Else
                If nWaistCutOutModified > 0 Then
                    sError = sError + "Distance from back of Cut-Out to Waist or Chest" + NL
                    sError = sError + "was less than 3 inches!" + NL
                    Select Case nWaistCutOutModified
                        Case 1
                            sError = sError & "Waist CutOut had to be decreased " & Str$(nWaistCutOutModified) & " time." & NL
                            'sError = sError + "ref:- " + NL
                        Case Is > 1
                            sError = sError & "Waist CutOut had to be decreased " & Str$(nWaistCutOutModified) & " times." & NL
                            'sError = sError + "ref:- " + NL
                    End Select
                End If
                Exit Do     'greater than 3" - acceptable
        End If
    Loop

    'Check that all of the reductions are within 5% of each other
    Dim nMinRed%, nMaxRed%, nCurrentValue%
    nMinRed = 0
    nMaxRed = 0
    
    'turn all reductions into integers for easy comparing
    'Waist
    nMinRed = (nWaistCirRed * 100)
    nMaxRed = nMinRed
    'Chest
    nCurrentValue = (nChestCirRed * 100)
    If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
    End If
    'Buttocks
    nCurrentValue = (nButtocksCirRed * 100)
    If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
    End If
    'Thigh
    nCurrentValue = (nThighCirRed * 100)
    If nMinRed > nCurrentValue Then
            nMinRed = nCurrentValue
        Else
            If nMaxRed < nCurrentValue Then
                nMaxRed = nCurrentValue
            End If
    End If

    If (nMaxRed - nMinRed) > 5 Then
        sError = sError + NL + "Severe - Warning!" + NL
        sError = sError + "All Reductions should be within 5% of each other" + NL
 '       iFatalError = True   'gg
    End If
   
   'Display Error message (if required) and return
   'There could be fatal errors found
    If Len(sError) > 0 Then
            If iFatalError = True Then
                    MsgBox sError, 64, "Warning - Problems with data"
                    Fn_ValidateData = False
                Else
                    sError = sError + NL + "The above problems have been found in the data do you" + NL
                    sError = sError + "wish to continue ?"
                    iError = MsgBox(sError, 52, "Severe Problems with data")
                    If iError = IDYES Then
                        Fn_ValidateData = True
                    Else
                        Fn_ValidateData = False
                    End If
            End If
        Else
            Fn_ValidateData = True
    End If

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

    For ii = 0 To 4
        sString = sString & cboRed(ii).Text
    Next ii

    For ii = 0 To 6
        sString = sString & txtCir(ii)
    Next ii
    'Reason for skip is that 7 & 8 were taken out during development
    For ii = 9 To 16
        sString = sString & txtCir(ii)
    Next ii
    
    
    sString = sString & cboLeftCup.Text & txtLeftDisk.Text
    
    sString = sString & cboRightCup.Text & txtRightDisk.Text

    sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text

    sString = sString & cboFrontNeck.Text & txtFrontNeck.Text

    sString = sString & cboBackNeck.Text & txtBackNeck.Text
    
    sString = sString & cboClosure.Text

    sString = sString & cboFabric.Text

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

Private Function fnDisplayToInches(ByVal nDisplay As Double) As Double
   'This function takes the value given and converts it
   'into a decimal version in inches, rounded to the nearest eighth
   'of an inch.
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
   'Globals:-
   '        g_nUnitsFac = 1       => nDisplay in Inches
   '        g_nUnitsFac = 10/25.5 => nDisplay in CMs
   'Returns:-
   '        Single, Inches rounded to the nearest eighth (0.125)
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
        fnDisplayToInches = fnRoundInches(nDisplay * g_nUnitsFac) * iSign
        Exit Function
    End If

   'Imperial units
    iInt = Int(nDisplay)
    nDec = nDisplay - iInt

   'Check that conversion is possible (return -1 if not)
    If nDec > 0.8 Then
        fnDisplayToInches = -1
    Else
        fnDisplayToInches = fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
    End If
End Function



Private Function fnGetString(ByVal sString As String, iindex As Integer) As String
   'Function to return as a string the iIndexth item in a string
   'that uses blanks (spaces) as delimiters.
   'EG
   '    sString = "Sam Spade Hello"
   '    fnGetNumber( sString, 2) = "Spade"
   '
   'If the iIndexth item is not found then return -1 to indicate an error.
   'This assumes that the string will not be used to store -ve numbers.
   'Indexing starts from 1

    Dim ii As Integer, iPos As Integer
    Dim sItem As String
    
   'Initial error checking
    sString = Trim$(sString) 'Remove leading and trailing blanks

    If Len(sString) = 0 Then
        fnGetString = ""
        Exit Function
    End If

   'Prepare string
    sString = sString & " " 'Trailing blank as stopper for last item

   'Get iIndexth item
    For ii = 1 To iindex
        iPos = InStr(sString, " ")
        If ii = iindex Then
            sString = Left$(sString, iPos - 1)
            fnGetString = sString
            Exit Function
        Else
            sString = LTrim$(Mid$(sString, iPos))
            If Len(sString) = 0 Then
                fnGetString = ""
                Exit Function
            End If
        End If
    Next ii

   'The function should have exited befor this, however just in case
   '(iIndex = 0) we indicate an error,
    fnGetString = ""

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

Private Sub Form_LinkClose()
    Dim ii%

   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

   'Check that a "MainPatientDetails" Symbol has been
   'found
    If txtUidMPD = "" Then
        MsgBox "No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue"
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
    If sChar = "R" Or sChar = "S" Or sChar = "V" Then
       'Disable length box for Regular, Scoop, V neck
        txtBackNeck.Enabled = False
        labBackNeck.Enabled = False
    Else
        txtBackNeck.Enabled = True
        labBackNeck.Enabled = True
    End If

    g_sFrontNeck = txtFrontNeck
    sChar = Left$(cboFrontNeck.Text, 1)
    If sChar = "R" Or sChar = "S" Then
        txtFrontNeck.Enabled = False
        labFrontNeck.Enabled = False
    Else
        labFrontNeck.Enabled = True
        txtFrontNeck.Enabled = True
    End If
   
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
         If cboLegStyle.List(ii) = fnGetString(txtLegStyle.Text, 1) Then
             cboLegStyle.ListIndex = ii
         End If
    Next ii
 
   'Set this now to avoid problems with cboLegStyle_Click
    If cboLegStyle.ListIndex = -1 Then
         g_sPreviousLegStyle = "None"
    Else
         g_sPreviousLegStyle = cboLegStyle.List(cboLegStyle.ListIndex)
    End If

    If fnGetString(txtLegStyle.Text, 2) <> "" Then
        If fnGetString(txtLegStyle.Text, 2) = "Panty" Then
            optLeftLeg(0).Value = True
        Else
            optLeftLeg(1).Value = True 'Brief
        End If
    End If
    If fnGetString(txtLegStyle.Text, 3) <> "" Then
        If fnGetString(txtLegStyle.Text, 3) = "Panty" Then
            optRightLeg(0).Value = True
        Else
            optRightLeg(1).Value = True 'Brief
        End If
    End If
   
   'Save the values in the text to a string
   'this can then be used to check if they have changed
   'on use of the close button
'    g_sChangeChecker = FN_ValuesString()
    Screen.MousePointer = 0    ' Change pointer back to normal.

    'Disable bras for Males
    If txtSex.Text = "Male" Then
        frmBra.Enabled = False
        For ii = 0 To 11
            LabBra(ii).Enabled = False
        Next ii
        cboLeftCup.Enabled = False
        txtLeftDisk.Enabled = False
        cboRightCup.Enabled = False
        txtRightDisk.Enabled = False
        cboRed(2).Enabled = False
        cmdCalculateBraDisks.Enabled = False
    Else
        PR_EnableCalculateDiskButton
    End If
    Show
    g_sChangeChecker = FN_ValuesString()

End Sub

Private Sub Form_Load()
    Dim ii%

    Hide

    Set MainForm = bodysuit

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
        MsgBox "BodySuit Dialogue is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting BodySuit - Dialogue"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - bodysuit.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - bodysuit.Height) / 2  ' Center form vertically.


   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

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
    cboRed(0).AddItem "0.85"     'gg
    cboRed(0).AddItem "0.80"
    cboRed(0).AddItem ""
    
    'Chest Default 90% +/- 5%
    cboRed(1).AddItem "0.95"
    cboRed(1).AddItem "0.90"     'gg
    cboRed(1).AddItem "0.85"
    cboRed(1).AddItem ""

    'Under Breast 90% +/- 5%
    cboRed(2).AddItem "0.95"
    cboRed(2).AddItem "0.90"     'gg
    cboRed(2).AddItem "0.85"
    cboRed(2).AddItem ""

    'Large part of Buttocks 85% +/- 5%
    cboRed(3).AddItem "0.90"
    cboRed(3).AddItem "0.85"     'gg
    cboRed(3).AddItem "0.80"
    cboRed(3).AddItem ""
    
    'Left Thigh 85% +/- 5%
    cboRed(4).AddItem "0.90"
    cboRed(4).AddItem "0.85"     'gg
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
    cboFrontNeck.AddItem "Measured V neck"
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

Private Sub optLeftLeg_Click(Index As Integer)
    
    If cboLegStyle.List(cboLegStyle.ListIndex) = "Panty" Then
        Select Case Index
         Case 0 'Panty
            'Enable Right side
             optRightLeg(1).Enabled = True

         Case 1 'Brief
            'Disable Right side
             optRightLeg(1).Enabled = False
        End Select
    End If
End Sub

Private Sub optRightLeg_Click(Index As Integer)
    If cboLegStyle.List(cboLegStyle.ListIndex) = "Panty" Then
        Select Case Index
         Case 0 'Panty
            'Enable left side
             optLeftLeg(1).Enabled = True

         Case 1 'Brief
            'Disable Left side
             optLeftLeg(1).Enabled = False
        End Select
    End If

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a body then Define the DATA Base
   'fields for the BodySuit and insert the BODYSUIT symbol
    PR_PutLine "HANDLE hMPD, hBody;"
   
    PR_UpdateDB

    Close #fNum
   
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

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable

    Print #fNum, sLine

End Sub

Private Sub PR_PutNumberAssign(sVariableName As String, nAssignedNumber As Variant)
   'Procedure to put a number assignment
   'Adds a semi-colon
   '    fNum is global variable

    Print #fNum, sVariableName; "="; Str$(nAssignedNumber); ";"

End Sub

Private Sub PR_PutStringAssign(sVariableName As String, sAssignedString As Variant)
   'Procedure to put a string assignment
   'Encloses String in quotes and adds a semi-colon
   '    fNum is global variable

    Print #fNum, sVariableName; "="; QQ; sAssignedString; QQ; ";"

End Sub

Private Sub PR_Select_Text(Text_Box_Name As Control)
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
End Sub

Private Sub PR_SetLayer(sNewLayer As String)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the current LAYER.
   'For this to work it assumes that hLayer is defined in DRAFIX as
   'a HANDLE.
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_sCurrentLayer are globals initialised by FN_Open
   '
   'To reduce unessesary writing of DRAFIX code check that the new layer
   'is different from the Current layer, change only if it is different.
   '

    If g_sCurrentLayer = sNewLayer Then Exit Sub
    g_sCurrentLayer = sNewLayer
    
    Print #fNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; sNewLayer; QQ; ");"
    Print #fNum, "if ( hLayer > %zero && hLayer != 32768)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Private Sub PR_UpdateDB()
   'Procedure called from
   '    PR_CreateMacro_Save
   'and
   '    PR_CreateMacro_Data
   '
   'Used to stop duplication on code

    Dim sSymbol As String
    Dim sLeftleg As String, sRightleg As String
    
    sSymbol = "suitbody"

    If txtUidVB.Text = "" Then
       'Define DB Fields
        PR_PutLine "@" & g_sPathJOBST & "\BODY\BFIELDS.D;"

       'Find "mainpatientdetails" and get position
        PR_PutLine "XY     xyMPD_Origin, xyMPD_Scale ;"
        PR_PutLine "STRING sMPD_Name;"
        PR_PutLine "ANGLE  aMPD_Angle;"

        PR_PutLine "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");"
        PR_PutLine "if (hMPD)"
        PR_PutLine "  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");"

       'Insert bodybox
        PR_PutLine "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){"
        PR_PutLine "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));"
        PR_PutLine "  hBody = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);"
        PR_PutLine "  }"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");"
    Else
       'Use existing symbol
        PR_PutLine "hBody = UID (" & QQ & "find" & QC & Val(txtUidVB.Text) & ");"
        PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

    End If

   'Update the BODY Box symbol
    PR_PutLine "SetDBData( hBody" & CQ & "LtSCir" & QCQ & txtCir(0).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "RtSCir" & QCQ & txtCir(1).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "NeckCir" & QCQ & txtCir(2).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "SWidth" & QCQ & txtCir(3).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "S_Waist" & QCQ & txtCir(4).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ChestCir" & QCQ & txtCir(5).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtCir(6).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "S_Breast" & QCQ & txtCir(9).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BreastCir" & QCQ & txtCir(10).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "NippleCir" & QCQ & txtCir(11).Text & QQ & ");"
    
    'Extra Values
    PR_PutLine "SetDBData( hBody" & CQ & "SFButt" & QCQ & txtCir(12).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "SLgButt" & QCQ & txtCir(13).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "LgButtCir" & QCQ & txtCir(14).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "LtThCir" & QCQ & txtCir(15).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "RtThCir" & QCQ & txtCir(16).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "BraLtCup" & QCQ & cboLeftCup.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BraRtCup" & QCQ & cboRightCup.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BraLtDisk" & QCQ & txtLeftDisk.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BraRtDisk" & QCQ & txtRightDisk.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "LtAxillaType" & QCQ & FN_EscapeQuotesInString(cboLeftAxilla) & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "RtAxillaType" & QCQ & FN_EscapeQuotesInString(cboRightAxilla) & QQ & ");"


    PR_PutLine "SetDBData( hBody" & CQ & "NeckType" & QCQ & cboFrontNeck.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "NeckDimension" & QCQ & txtFrontNeck.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BackNeckType" & QCQ & cboBackNeck.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BackNeckDim" & QCQ & txtBackNeck.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "Closure" & QCQ & cboClosure.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "Fabric" & QCQ & cboFabric.Text & QQ & ");"
    
    'Extra Values
    'Index 0   =>  Panty
    'Index 1   =>  Brief
    If optLeftLeg(0).Value = True Then
        sLeftleg = "Panty"
    Else
        sLeftleg = "Brief"
    End If
    If optRightLeg(0).Value = True Then
        sRightleg = "Panty"
    Else
        sRightleg = "Brief"
    End If
    'txtLegStyle = cboLegStyle.Text & " " & sLeftleg & " " & sRightleg & " " & Str$(nLargestCir)
    txtLegStyle = cboLegStyle.Text & " " & sLeftleg & " " & sRightleg

    PR_PutLine "SetDBData( hBody" & CQ & "LegStyle" & QCQ & txtLegStyle.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "CrotchStyle" & QCQ & cboCrotchStyle.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "WaistCirUserFac" & QCQ & cboRed(0).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ChestCirUserFac" & QCQ & cboRed(1).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BreastCirUserFac" & QCQ & cboRed(2).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "LgButtCirUserFac" & QCQ & cboRed(3).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ThCirUserFac" & QCQ & cboRed(4).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"

End Sub

Private Function round(ByVal nNumber As Double) As Integer
   'Fuction to return the rounded value of a decimal number
   'E.G.
   '    round(1.35)  = 1
   '    round(1.55)  = 2
   '    round(2.50)  = 3
   '    round(-2.50) = -3
   '    round(0)     = 0
   '

    Dim iInt As Integer, iSign As Integer
    
   'Avoid extra work. Return 0 if input is 0
    If nNumber = 0 Then
        round = 0
        Exit Function
    End If

   'Split input
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)

   'Effect rounding
    If (nNumber - iInt) >= 0.5 Then
        round = (iInt + 1) * iSign
    Else
        round = iInt * iSign
    End If
    
End Function

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


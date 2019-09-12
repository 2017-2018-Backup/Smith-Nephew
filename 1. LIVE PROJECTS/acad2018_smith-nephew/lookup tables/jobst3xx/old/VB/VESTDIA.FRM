VERSION 5.00
Begin VB.Form vestdia 
   Caption         =   "VEST Body - Dialogue"
   ClientHeight    =   7170
   ClientLeft      =   390
   ClientTop       =   1290
   ClientWidth     =   8385
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
   LinkTopic       =   "vestdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   8385
   Begin VB.CommandButton cmdTab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   9540
      TabIndex        =   117
      Top             =   5880
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   5880
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
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
      Left            =   8880
      TabIndex        =   106
      Text            =   "EOS reduction"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   8880
      TabIndex        =   105
      Text            =   "Waist Reduction"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidVB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   102
      Text            =   "txtUidVB"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   100
      Text            =   "txtWorkOrder"
      Top             =   930
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   99
      Text            =   "txtUidMPD"
      Top             =   630
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Design Choices"
      Height          =   3645
      Left            =   4800
      TabIndex        =   97
      Top             =   2250
      Width           =   3435
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Text            =   "cboFabric"
         Top             =   3180
         Width           =   3075
      End
      Begin VB.ComboBox cboClosure 
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Text            =   "cboClosure"
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtBackNeck 
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2190
         Width           =   555
      End
      Begin VB.ComboBox cboBackNeck 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Text            =   "cboBackNeck"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtFrontNeck 
         Height          =   315
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1410
         Width           =   555
      End
      Begin VB.ComboBox cboFrontNeck 
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Text            =   "cboFrontNeck"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.ComboBox cboRightAxilla 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Text            =   "cboRightAxilla"
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cboLeftAxilla 
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Text            =   "cboLeftAxilla"
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Fabric:"
         Height          =   285
         Index           =   33
         Left            =   180
         TabIndex        =   72
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Closure:"
         Height          =   285
         Index           =   32
         Left            =   180
         TabIndex        =   73
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label labBackNeck 
         Caption         =   "Back Neck size:"
         Height          =   285
         Left            =   180
         TabIndex        =   74
         Top             =   2220
         Width           =   1995
      End
      Begin VB.Label lblBackNeck 
         Caption         =   "inches"
         Height          =   165
         Left            =   2340
         TabIndex        =   75
         Top             =   2220
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Back Neck:"
         Height          =   285
         Index           =   30
         Left            =   180
         TabIndex        =   76
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label labFrontNeck 
         Caption         =   "Front Neck size:"
         Height          =   285
         Left            =   180
         TabIndex        =   77
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblFrontNeck 
         Caption         =   "inches"
         Height          =   165
         Left            =   2340
         TabIndex        =   78
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Front Neck:"
         Height          =   285
         Index           =   28
         Left            =   180
         TabIndex        =   79
         Top             =   1080
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
      Height          =   2985
      Left            =   120
      TabIndex        =   62
      Top             =   4140
      Width           =   4575
      Begin VB.CommandButton cmdCalculateBraDisks 
         Caption         =   "Calculate Disks"
         Height          =   375
         Left            =   1380
         TabIndex        =   116
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox txtRightDisk 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3540
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1980
         Width           =   375
      End
      Begin VB.ComboBox cboRightCup 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox txtLeftDisk 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1950
         Width           =   375
      End
      Begin VB.ComboBox cboLeftCup 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1950
         Width           =   1035
      End
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   2
         Left            =   3600
         TabIndex        =   34
         Text            =   "cboRed"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   11
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   12
         Top             =   990
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   10
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   11
         Top             =   690
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   9
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "9"
         Top             =   330
         Width           =   555
      End
      Begin VB.Label LabBra 
         Caption         =   "Disk:"
         Height          =   285
         Index           =   5
         Left            =   3540
         TabIndex        =   96
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Cup:"
         Height          =   285
         Index           =   4
         Left            =   2340
         TabIndex        =   95
         Top             =   1710
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Disk:"
         Height          =   285
         Index           =   3
         Left            =   1380
         TabIndex        =   94
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label LabBra 
         Caption         =   "Cup:"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   93
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label LabBra 
         Alignment       =   2  'Center
         Caption         =   "Right Bra Cup"
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   92
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label LabBra 
         Alignment       =   2  'Center
         Caption         =   "Left Bra Cup"
         Height          =   285
         Index           =   0
         Left            =   240
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
         Left            =   1980
         TabIndex        =   89
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   11
         Left            =   2940
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
         Left            =   1980
         TabIndex        =   86
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   10
         Left            =   2940
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
         Left            =   1980
         TabIndex        =   83
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   9
         Left            =   2940
         TabIndex        =   82
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vest below waist"
      Height          =   1125
      Left            =   120
      TabIndex        =   61
      Top             =   2970
      Width           =   4575
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   33
         Text            =   "cboRed"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   8
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   9
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   7
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "7"
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Circ. at EOS"
         Height          =   285
         Index           =   18
         Left            =   180
         TabIndex        =   71
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   17
         Left            =   1980
         TabIndex        =   70
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   8
         Left            =   2940
         TabIndex        =   69
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder to EOS"
         Height          =   285
         Index           =   16
         Left            =   180
         TabIndex        =   68
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   15
         Left            =   1980
         TabIndex        =   67
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   7
         Left            =   2940
         TabIndex        =   66
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Measurments for Vest"
      Height          =   2835
      Left            =   120
      TabIndex        =   60
      Top             =   90
      Width           =   4575
      Begin VB.ComboBox cboRed 
         Height          =   315
         Index           =   0
         Left            =   3600
         TabIndex        =   32
         Text            =   "cboRed"
         Top             =   2370
         Width           =   795
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   6
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2340
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   5
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   6
         Top             =   2010
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   4
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   3
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   2
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   1
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   1
         Top             =   690
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   0
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   6
         Left            =   2940
         TabIndex        =   65
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "7"
         Height          =   285
         Index           =   14
         Left            =   1980
         TabIndex        =   64
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Waist circ."
         Height          =   285
         Index           =   13
         Left            =   180
         TabIndex        =   63
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   5
         Left            =   2940
         TabIndex        =   59
         Top             =   2070
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "6"
         Height          =   285
         Index           =   12
         Left            =   1980
         TabIndex        =   58
         Top             =   2070
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Chest circ."
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   56
         Top             =   2070
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   4
         Left            =   2940
         TabIndex        =   54
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "5"
         Height          =   285
         Index           =   10
         Left            =   1980
         TabIndex        =   52
         Top             =   1740
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder to Waist"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   50
         Top             =   1740
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   3
         Left            =   2940
         TabIndex        =   48
         Top             =   1410
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "4"
         Height          =   285
         Index           =   8
         Left            =   1980
         TabIndex        =   46
         Top             =   1410
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Shoulder width"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   45
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   2
         Left            =   2940
         TabIndex        =   43
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "3"
         Height          =   285
         Index           =   5
         Left            =   1980
         TabIndex        =   41
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Neck circ."
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   40
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   1
         Left            =   2940
         TabIndex        =   39
         Top             =   750
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         Height          =   285
         Index           =   3
         Left            =   1980
         TabIndex        =   38
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Right shoulder circ."
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   0
         Left            =   2940
         TabIndex        =   36
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   35
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Left shoulder circ."
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   375
      Left            =   5100
      TabIndex        =   42
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   44
      Top             =   6180
      Width           =   1275
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   4800
      TabIndex        =   4
      Top             =   90
      Width           =   3435
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   53
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2580
         TabIndex        =   49
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2400
         TabIndex        =   57
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   55
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   47
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   51
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   1920
         TabIndex        =   18
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1620
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Combo Box DDE "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   8880
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
      Left            =   9000
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
      Left            =   9000
      TabIndex        =   101
      Top             =   420
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "vestdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   VESTDIA
'Purpose:   Vest Dialogue
'
'
'Version:   3.01
'Date:      4.Oct.95
'Author:    Gary George
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

'   'MsgBox constant
'    Const IDCANCEL = 2
'    Const IDYES = 6
'    Const IDNO = 7


    
    Option Explicit



Private Sub Cancel_Click()
    Dim Response As Integer, sTask$, sCurrentValues$
    
   'Check if data has been modified
    sCurrentValues = FN_ValuesString()

    If sCurrentValues <> g_sChangeChecker Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
        Select Case Response
            Case IDYES
                PR_CreateMacro_Save "c:\jobst\draw.d"
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate sTask
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "VEST Body - Dialogue"
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

Private Sub cboFrontNeck_Click()
    Select Case Left$(cboFrontNeck.Text, 1)
        Case "R", "S"   'Regular or Scoop
            txtFrontNeck = ""
            lblFrontNeck.Caption = ""
            txtFrontNeck.Enabled = False
            labFrontNeck.Enabled = False
        Case Else   'Measured Scoop or Turtle neck
            If g_sFrontNeck <> "" Then
                txtFrontNeck = g_sFrontNeck
                txtFrontNeck_LostFocus   'Display inches
            End If
            txtFrontNeck.Enabled = True
            labFrontNeck.Enabled = True
    End Select
End Sub

Private Sub cboLeftCup_Click()
    If vestdia.Visible And (cboLeftCup.Text = "" Or cboLeftCup.Text = "None") Then
        txtLeftDisk = ""
        g_sLeftCup = ""
        g_sLeftDisk = ""
    Else
        PR_EnableCalculateDiskButton
    End If
End Sub

Private Sub cboRightCup_Click()
    If vestdia.Visible And (cboRightCup.Text = "" Or cboRightCup.Text = "None") Then
        txtRightDisk = ""
        g_sRightCup = ""
        g_sRightDisk = ""
    Else
        PR_EnableCalculateDiskButton
       'PR_DoBraCupsAndDisks
    End If
End Sub

Private Sub cmdCalculateBraDisks_Click()
    PR_DoBraCupsAndDisks
End Sub

Private Sub cmdTab_Click()
   'Allows the user to use enter as a tab
    SendKeys "{TAB}", False
End Sub

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
            MsgBox sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")"
            Exit Function
        End If
        
        nDiff = nNipple - nChest
        If nDiff < 0 Then
            FN_CalculateBra = False
            sError = "Circumference over Nipple line is smaller than Chest circumference"
            MsgBox sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")"
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
        MsgBox sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")"
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
            MsgBox sError, 0, "VEST Body - Bra Cup" & "(" & sSide & ")"
        End If
    Else
        iDisk = nSelctedDisk
    End If
    
    FN_CalculateBra = True


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
    Print #fNum, "//by Visual Basic, VEST Body"
    Print #fNum, "//type - "; sType; ""


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

    ReDim sCircum(0 To 11)
    sCircum(0) = "Left shoulder circ."
    sCircum(1) = "Right shoulder circ."
    sCircum(2) = "Neck circ."
    sCircum(3) = "Shoulder width"
    sCircum(4) = "Shoulder to waist"
    sCircum(5) = "Chest circ."
    sCircum(6) = "Waist circ."
    sCircum(7) = "Shoulder to EOS"
    sCircum(8) = "Circ. at EOS"
    sCircum(9) = "Shoulder to under breast"
    sCircum(10) = "Circ. under breast"
    sCircum(11) = "Circ. over nipple"

   'Vest measurements (all must be present)
    For ii = 0 To 6
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing dimension for " & sCircum(ii) & "!" & NL
        End If
    Next ii

   'EOS Measurements (if one given both must be given)
    If Val(txtCir(7).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(7) & "!" & NL
    End If
    If Val(txtCir(8).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(8) & "!" & NL
    End If

   'Bra Cups
   'Note:
   '    The Circumference over nipple is optional unless a cup has been
   '    specified
   '
    If Val(txtCir(9).Text) = 0 And Val(txtCir(10).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(9) & "!" & NL
    End If
    If Val(txtCir(10).Text) = 0 And Val(txtCir(9).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(10) & "!" & NL
    End If

    If (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0) And ((cboLeftCup.Text <> "None" And txtLeftDisk.Text = "") Or (txtRightDisk.Text = "" And cboRightCup.Text <> "None")) Then
            sError = sError & "Bra Measurements or Bra Cups requested but no disks calculated" & "!" & NL
    End If

    If Val(txtCir(9).Text) = 0 And (txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "") Then
            sError = sError & "Bra disks given! But missing dimension for " & sCircum(9) & "!" & NL
    End If

   'If cups or dimensions given then a disk must be present
   'NB
   '    cboXXXXCup.ListIndex = 6 = "None"
   '    cboXXXXCup.ListIndex = 7 = ""

    If cboLeftCup.ListIndex < 6 And Val(txtLeftDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Left BRA Cup" & "!" & NL
    End If
    If cboRightCup.ListIndex < 6 And Val(txtRightDisk.Text) = 0 Then
            sError = sError & "No disk calculated for Right BRA Cup" & "!" & NL
    End If
    
   'Sex error
    If txtSex.Text = "Male" And (Val(txtCir(9).Text) <> 0 Or Val(txtCir(10).Text) <> 0 Or cboLeftCup.ListIndex < 0 Or cboRightCup.ListIndex < 0) Then
            sError = sError & "Male patient but Bra Measurements or Bra Cups requested " & "!" & NL
    End If

   'Neck at back and front
    Dim sChar As String * 1
    sChar = Left$(cboBackNeck.Text, 1)
    If sChar = "M" And txtBackNeck.Text = "" Then
            sError = sError & "No dimension for Back Meck Measured Scoop! " & NL
    End If

    sChar = Left$(cboFrontNeck.Text, 1)
    If sChar = "M" And txtFrontNeck.Text = "" Then
            sError = sError & "No dimension for Front Neck Measured Scoop! " & NL
    End If

   '
    If cboLeftAxilla.Text = "" Then
            sError = sError & "Left Axilla not given! " & NL
    End If

    If cboRightAxilla.Text = "" Then
            sError = sError & "Right Axilla not given! " & NL
    End If

    If cboFrontNeck.Text = "" Then
            sError = sError & "Neck not given! " & NL
    End If

    If cboBackNeck.Text = "" Then
            sError = sError & "Back neck not given! " & NL
    End If

    If cboClosure.Text = "" Then
            sError = sError & "Closure not given! " & NL
    End If

    If cboFabric.Text = "" Then
            sError = sError & "Fabric not given! " & NL
    End If
    
    If sError <> "" Then
        MsgBox sError, 16, "VEST Body - Dialogue"
        FN_ValidateData = False
    Else
        FN_ValidateData = True
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

    For ii = 0 To 11
        sString = sString & txtCir(ii)
    Next ii

    For ii = 0 To 2
        sString = sString & cboRed(ii).Text
    Next ii
    
    sString = sString & cboLeftCup.Text & txtLeftDisk.Text
    
    sString = sString & cboRightCup.Text & txtRightDisk.Text

    sString = sString & cboLeftAxilla.Text & cboRightAxilla.Text

    sString = sString & cboFrontNeck.Text & txtFrontNeck.Text

    sString = sString & cboBackNeck.Text & txtBackNeck.Text
    
    sString = sString & cboClosure.Text

    sString = sString & cboFabric.Text

    FN_ValuesString = sString



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
    For ii = 0 To 11
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
    If sChar = "R" Or sChar = "S" Then
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

   'Save the values in the text to a string
   'this can then be used to check if they have changed
   'on use of the close button
    g_sChangeChecker = FN_ValuesString()
    Screen.MousePointer = 0    ' Change pointer to hourglass.
    
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

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "VEST Body Dialogue is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting VEST Body - Dialogue"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - vestdia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - vestdia.Height) / 2  ' Center form vertically.

    Set MainForm = vestdia

   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    g_nUnitsFac = 1
    g_PathJOBST = fnPathJOBST()

   'Clear fields
   'Circumferences and lengths
    For ii = 0 To 11
        txtCir(ii) = ""
    Next ii

   'The data from these DDE text boxes is copied
   'to the combo boxes on Link close
   '
    For ii = 0 To 10
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
    cboRed(0).AddItem "0.95"
    cboRed(0).AddItem "0.85"
    cboRed(0).AddItem ""

    cboRed(1).AddItem "0.95"
    cboRed(1).AddItem "0.85"
    cboRed(1).AddItem ""

    cboRed(2).AddItem "0.95"
    cboRed(2).AddItem "0.85"
    cboRed(2).AddItem ""

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
    cboFrontNeck.AddItem "Turtle"
    cboFrontNeck.AddItem "Turtle - Fabric same as Vest"
    cboFrontNeck.AddItem "Turtle Detachable"
    cboFrontNeck.AddItem "Turtle Detach. Fabric"

    cboBackNeck.AddItem "Regular"
    cboBackNeck.AddItem "Scoop"
    cboBackNeck.AddItem "Measured Scoop"

    cboClosure.AddItem "Velcro"
    cboClosure.AddItem "Zip"
    cboClosure.AddItem "Front Velcro"
    cboClosure.AddItem "Front Velcro (Reversed)"
    cboClosure.AddItem "Back Velcro"
    cboClosure.AddItem "Back Velcro (Reversed)"
    cboClosure.AddItem "Front Zip"
    cboClosure.AddItem "Back Zip"

    PR_GetComboListFromFile cboFabric, g_PathJOBST & "\FABRIC.DAT"


End Sub

Private Sub OK_Click()
    Dim sTask$
   'Don't allow multiple clicking
   '
    OK.Enabled = False
    If FN_ValidateData() Then
        Screen.MousePointer = 11
        Hide

        PR_CreateMacro_Data "c:\jobst\draw.d"
        PR_CreateMacro_Axilla "c:\jobst\draw_1.d", "c:\jobst\draw_3.d"
        PR_CreateMacro_Bra "c:\jobst\draw_2.d"
        
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            AppActivate sTask
            SendKeys "@" & g_PathJOBST & "\VEST\BODY.D{enter}", True
            Screen.MousePointer = 0
            End
        Else
            MsgBox "Can't find a Drafix Drawing to update!", 16, "CAD Glove Dialogue"
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub PR_CreateMacro_Axilla(sfile As String, sFile2 As String)
   'Assumes that data has been validated befor this
   'procedure is called.

    g_bMeshAxilla = False

    Dim sLeftAxilla$, sRightAxilla$, sAxilla$
    Dim iLoop%, ii%

   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sfile, "Axilla Drawing MACRO/S", (txtPatientName.Text), (txtFileNo.Text))

    sLeftAxilla = Left$(cboLeftAxilla.Text, 1)
    sRightAxilla = Left$(cboRightAxilla.Text, 1)

    If sLeftAxilla <> sRightAxilla Then iLoop = 2 Else iLoop = 1

    For ii = 1 To iLoop
        If ii = 1 Then
            sAxilla = sLeftAxilla
        Else
            sAxilla = sRightAxilla
        End If
        Select Case sAxilla
            Case "R"
               'Regular Axilla
                PR_PutLine "@" & g_PathJOBST & "\VEST\BODREGUL.D;"
            Case "O", "L"
               'Open or Lining Axilla
                PR_PutLine "@" & g_PathJOBST & "\VEST\BODOTHRS.D;"
            Case "M"
               'Mesh Axilla
                g_bMeshAxilla = True
                PR_PutLine "@" & g_PathJOBST & "\VEST\BODMESH.D;"
            Case "S"
               'Sleeveless Axilla
                PR_PutLine "@" & g_PathJOBST & "\VEST\BODSLESS.D;"
        End Select
    Next ii

    Close #fNum


    fNum = FN_Open(sFile2, "Draw mesh", (txtPatientName.Text), (txtFileNo.Text))
     If g_bMeshAxilla Then Print #fNum, "Execute ("; QQ; "application"; QC; "sPathJOBST + "; QQ; "\\raglan\\meshvest"; QCQ; "normal"; QQ; " );"
    Close #fNum

End Sub

Private Sub PR_CreateMacro_Bra(sfile As String)
   'Assumes that data has been validated before this
   'procedure is called.
    Dim nBraCLOffset#, nDiskXoff#, nDiskYoff#
    Dim ii%, iDisk%, iLeftDisk%, iRightDisk%

   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sfile, "Load Bradisks", (txtPatientName.Text), (txtFileNo.Text))

    If cboLeftCup.ListIndex < 7 Or cboLeftCup.ListIndex < 7 Or txtLeftDisk.Text <> "" Or txtRightDisk.Text <> "" Then
       'A cup or a disk of some sort has been given
        If cboLeftCup.Text = "None" Then iLeftDisk = -1 Else iLeftDisk = Val(txtLeftDisk.Text)
        If cboRightCup.Text = "None" Then iRightDisk = -1 Else iRightDisk = Val(txtRightDisk.Text)

       'Loop through both cups and give positioning information
        For ii = 1 To 2
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
                   nDiskXoff = 2.293
                   nDiskYoff = 2.882
                Case 6
                   nBraCLOffset = 0.625
                   nDiskXoff = 2.625
                   nDiskYoff = 3.025
                Case 7
                   nBraCLOffset = 0.5
                   nDiskXoff = 2.94
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
            
            If ii = 1 Then
               'Left Disks
                PR_PutNumberAssign "nDiskLt", iLeftDisk
                PR_PutNumberAssign "nBraCLOffsetLt", nBraCLOffset
                PR_PutNumberAssign "nDiskXoffLt", nDiskXoff
                PR_PutNumberAssign "nDiskYoffLt", nDiskYoff
            Else
               'Right Disks
                PR_PutNumberAssign "nDiskRt", iRightDisk
                PR_PutNumberAssign "nBraCLOffsetRt", nBraCLOffset
                PR_PutNumberAssign "nDiskXoffRt", nDiskXoff
                PR_PutNumberAssign "nDiskYoffRt", nDiskYoff

            End If
        Next ii

       'Insert disks and text using the Assignments from above
        PR_PutLine "@" & g_PathJOBST & "\VEST\BODYBRA.D;"

    Else
       'No bra cups are required
        PR_PutLine "// --------------------"
        PR_PutLine "// NO BRA CUPS REQUIRED"
        PR_PutLine "// --------------------"
    End If

    Close #fNum

End Sub

Private Sub PR_CreateMacro_Data(sfile As String)
    
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sfile, "Save Data and Set Variables for BODY.D", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
    PR_PutLine "HANDLE hMPD, hBody;"
   
    PR_UpdateDB
   
   'Variable data
   'If display is in inches then convert to decimal inches
   'The conversion of CMs to Inches is done in BODY.D
   'in this way we do not add any extra problems due to rounding errors
   '
   'Patient details
    PR_PutStringAssign "sFileNo", (txtFileNo.Text)
    PR_PutStringAssign "sPatient", (txtPatientName.Text)
    PR_PutStringAssign "sAge", (txtAge.Text)
    PR_PutStringAssign "sSEX", (txtSex.Text)
    PR_PutStringAssign "sDiagnosis", (txtDiagnosis.Text)
    PR_PutNumberAssign "nUnitsFac", g_nUnitsFac
    PR_PutNumberAssign "nAge", Val(txtAge.Text)
   'PR_PutStringAssign "sUnitType", txtUnits.Text
    
   'VEST Body Details
    PR_PutNumberAssign "nLtSCir", FN_Decimalise(Val(txtCir(0)))
    PR_PutNumberAssign "nRtSCir", FN_Decimalise(Val(txtCir(1)))
    PR_PutNumberAssign "nNeckCir", FN_Decimalise(Val(txtCir(2)))
    PR_PutNumberAssign "nSWidth", FN_Decimalise(Val(txtCir(3)))
    PR_PutNumberAssign "nS_Waist", FN_Decimalise(Val(txtCir(4)))
    PR_PutNumberAssign "nChestCir", FN_Decimalise(Val(txtCir(5)))
    PR_PutNumberAssign "nChestCirActual", FN_Decimalise(Val(txtCir(5)))
    PR_PutNumberAssign "nWaistCir", FN_Decimalise(Val(txtCir(6)))
    PR_PutNumberAssign "nS_EOS", FN_Decimalise(Val(txtCir(7)))
    PR_PutNumberAssign "nEOSCir", FN_Decimalise(Val(txtCir(8)))
    PR_PutNumberAssign "nS_Breast", FN_Decimalise(Val(txtCir(9)))
    PR_PutNumberAssign "nBreastCir", FN_Decimalise(Val(txtCir(10)))
    PR_PutNumberAssign "nBreastCirActual", FN_Decimalise(Val(txtCir(10)))
    PR_PutNumberAssign "nNippleCir", FN_Decimalise(Val(txtCir(11)))

    PR_PutStringAssign "sBraLtCup", (cboLeftCup.Text)
    PR_PutStringAssign "sBraRtCup", (cboRightCup.Text)
    PR_PutStringAssign "sBraLtDisk", (txtLeftDisk.Text)
    PR_PutStringAssign "sBraRtDisk", (txtRightDisk.Text)

    PR_PutStringAssign "sLtAxillaType", FN_EscapeQuotesInString(cboLeftAxilla)
    PR_PutStringAssign "sRtAxillaType", FN_EscapeQuotesInString(cboRightAxilla)

    PR_PutStringAssign "sNeckType", (cboFrontNeck.Text)
    PR_PutNumberAssign "nNeckDimension", FN_Decimalise(Val(txtFrontNeck.Text))
    PR_PutStringAssign "sBackNeckType", (cboBackNeck.Text)
    PR_PutNumberAssign "nBackNeckDim", FN_Decimalise(Val(txtBackNeck.Text))

    PR_PutStringAssign "sClosure", (cboClosure.Text)
    PR_PutStringAssign "sFabric", (cboFabric.Text)

    PR_PutNumberAssign "nWaistCirUserFac", Val(cboRed(0).Text)
    PR_PutNumberAssign "nEOSCirUserFac", Val(cboRed(1).Text)
    PR_PutNumberAssign "nBreastCirUserFac", Val(cboRed(0).Text)

    If txtWorkOrder.Text = "" Then
        PR_PutStringAssign "sWorkOrder", "-"
    Else
        PR_PutStringAssign "sWorkOrder", (txtWorkOrder.Text)
    End If

    Close #fNum

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
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
    
    sSymbol = "vestbody"

    If txtUidVB.Text = "" Then
       'Define DB Fields
        PR_PutLine "@" & g_PathJOBST & "\VEST\VFIELDS.D;"

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
    
    PR_PutLine "SetDBData( hBody" & CQ & "S_EOS" & QCQ & txtCir(7).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "EOSCir" & QCQ & txtCir(8).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "S_Breast" & QCQ & txtCir(9).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BreastCir" & QCQ & txtCir(10).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "NippleCir" & QCQ & txtCir(11).Text & QQ & ");"
    
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
    
    PR_PutLine "SetDBData( hBody" & CQ & "WaistCirUserFac" & QCQ & cboRed(0).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "EOSCirUserFac" & QCQ & cboRed(1).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "BreastCirUserFac" & QCQ & cboRed(2).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"

End Sub

Private Sub Timer1_Timer()
    'It is assumed that the link open from Drafix has failed
    'Therefor we "End" here
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
            MsgBox "Bra disk sizes 1 to 9 only.", 48, "VEST Body - Dialogue"
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
            MsgBox "Bra disk sizes 1 to 9 only.", 48, "VEST Body - Dialogue"
    End Select
End Sub


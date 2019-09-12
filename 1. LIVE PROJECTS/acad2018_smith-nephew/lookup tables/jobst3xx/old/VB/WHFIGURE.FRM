VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form whfigure 
   Caption         =   "Waist Height - Figure"
   ClientHeight    =   8460
   ClientLeft      =   450
   ClientTop       =   480
   ClientWidth     =   11235
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
   LinkTopic       =   "whfigure"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   11235
   Begin VB.Timer Timer1 
      Left            =   6840
      Top             =   3570
   End
   Begin VB.TextBox txtLeftFirstTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   249
      Top             =   7230
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightFirstTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   248
      Top             =   7230
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftLastTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   247
      Top             =   7590
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightLastTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   246
      Top             =   7590
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidBody 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   244
      Top             =   4350
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   217
      Top             =   6870
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   218
      Top             =   6870
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightStr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   219
      Top             =   6510
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftStr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   220
      Top             =   6510
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   221
      Top             =   6150
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   222
      Top             =   6150
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Right 
      Caption         =   "Right Leg"
      Height          =   8415
      Left            =   7440
      TabIndex        =   238
      Top             =   0
      Width           =   3735
      Begin MSFlexGridLib.MSFlexGrid grdRightDisplay 
         Height          =   6390
         Left            =   2580
         TabIndex        =   257
         Top             =   1890
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   11271
         _Version        =   65541
         Rows            =   24
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   29
         Left            =   2040
         TabIndex        =   117
         Top             =   8070
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   28
         Left            =   2040
         TabIndex        =   116
         Top             =   7800
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   27
         Left            =   2040
         TabIndex        =   115
         Top             =   7530
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   26
         Left            =   2040
         TabIndex        =   114
         Top             =   7260
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   25
         Left            =   2040
         TabIndex        =   113
         Top             =   6990
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   24
         Left            =   2040
         TabIndex        =   112
         Top             =   6720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   23
         Left            =   2040
         TabIndex        =   111
         Top             =   6450
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   22
         Left            =   2040
         TabIndex        =   109
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   21
         Left            =   2040
         TabIndex        =   108
         Top             =   5910
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   20
         Left            =   2040
         TabIndex        =   107
         Top             =   5640
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   19
         Left            =   2040
         TabIndex        =   106
         Top             =   5370
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   18
         Left            =   2040
         TabIndex        =   105
         Top             =   5100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   17
         Left            =   2040
         TabIndex        =   104
         Top             =   4830
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   16
         Left            =   2040
         TabIndex        =   110
         Top             =   4560
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   15
         Left            =   2040
         TabIndex        =   103
         Top             =   4290
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   14
         Left            =   2040
         TabIndex        =   102
         Top             =   4020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   13
         Left            =   2040
         TabIndex        =   101
         Top             =   3750
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   12
         Left            =   2040
         TabIndex        =   100
         Top             =   3480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   11
         Left            =   2040
         TabIndex        =   99
         Top             =   3210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   10
         Left            =   2040
         TabIndex        =   98
         Top             =   2940
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   9
         Left            =   2040
         TabIndex        =   97
         Top             =   2670
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   8
         Left            =   2040
         TabIndex        =   96
         Top             =   2400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   6
         Left            =   2040
         TabIndex        =   94
         Top             =   1860
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtRightMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   7
         Left            =   2040
         TabIndex        =   95
         Top             =   2130
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Frame Frame1 
         Caption         =   "Template"
         Height          =   915
         Left            =   1980
         TabIndex        =   173
         Top             =   180
         Width           =   1640
         Begin VB.CheckBox chkRightZipper 
            Caption         =   "Ankle Zipper"
            Height          =   225
            Left            =   120
            TabIndex        =   93
            Top             =   630
            Width           =   1515
         End
         Begin VB.ComboBox cboRightTemplate 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   270
            Width           =   1395
         End
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   29
         Left            =   600
         MaxLength       =   7
         TabIndex        =   91
         Top             =   8070
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   28
         Left            =   600
         MaxLength       =   7
         TabIndex        =   90
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   27
         Left            =   600
         MaxLength       =   7
         TabIndex        =   89
         Top             =   7530
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   26
         Left            =   600
         MaxLength       =   7
         TabIndex        =   88
         Top             =   7260
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   25
         Left            =   600
         MaxLength       =   7
         TabIndex        =   87
         Top             =   6990
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   24
         Left            =   600
         MaxLength       =   7
         TabIndex        =   86
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   23
         Left            =   600
         MaxLength       =   7
         TabIndex        =   85
         Top             =   6450
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   22
         Left            =   600
         MaxLength       =   7
         TabIndex        =   84
         Top             =   6180
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   21
         Left            =   600
         MaxLength       =   7
         TabIndex        =   83
         Top             =   5910
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   20
         Left            =   600
         MaxLength       =   7
         TabIndex        =   82
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   19
         Left            =   600
         MaxLength       =   7
         TabIndex        =   81
         Top             =   5370
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   18
         Left            =   600
         MaxLength       =   7
         TabIndex        =   80
         Top             =   5100
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   17
         Left            =   600
         MaxLength       =   7
         TabIndex        =   79
         Top             =   4830
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   16
         Left            =   600
         MaxLength       =   7
         TabIndex        =   78
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   15
         Left            =   600
         MaxLength       =   7
         TabIndex        =   76
         Top             =   4290
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   14
         Left            =   600
         MaxLength       =   7
         TabIndex        =   75
         Top             =   4020
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   13
         Left            =   600
         MaxLength       =   7
         TabIndex        =   74
         Top             =   3750
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   12
         Left            =   600
         MaxLength       =   7
         TabIndex        =   73
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   11
         Left            =   600
         MaxLength       =   7
         TabIndex        =   72
         Top             =   3210
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   10
         Left            =   600
         MaxLength       =   7
         TabIndex        =   71
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   9
         Left            =   600
         MaxLength       =   7
         TabIndex        =   70
         Top             =   2670
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   600
         MaxLength       =   7
         TabIndex        =   69
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   600
         MaxLength       =   7
         TabIndex        =   68
         Top             =   2130
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   600
         MaxLength       =   7
         TabIndex        =   67
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   600
         MaxLength       =   7
         TabIndex        =   66
         Top             =   1590
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   600
         MaxLength       =   7
         TabIndex        =   65
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   600
         MaxLength       =   7
         TabIndex        =   64
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   600
         MaxLength       =   7
         TabIndex        =   63
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   600
         MaxLength       =   7
         TabIndex        =   62
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   600
         MaxLength       =   7
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid grdRightInches 
         Height          =   8100
         Left            =   1250
         TabIndex        =   256
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   14288
         _Version        =   65541
         Rows            =   30
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label labRightDisp 
         Caption         =   "MMHg"
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   55
         Top             =   1620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label labRightDisp 
         Alignment       =   2  'Center
         Caption         =   "Gms"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   56
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labRightDisp 
         Alignment       =   2  'Center
         Caption         =   "Red."
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   77
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labRightDisp 
         Caption         =   "Max Stretch:"
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   92
         Top             =   1230
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label labRightMaxStr 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   3180
         TabIndex        =   118
         Top             =   1230
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   119
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   120
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   121
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   123
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   124
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   125
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   126
         Top             =   1890
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   127
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   128
         Top             =   2430
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   129
         Top             =   2700
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   130
         Top             =   2970
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   131
         Top             =   3240
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "10½"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   132
         Top             =   3510
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "12"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   133
         Top             =   3780
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "13½"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   134
         Top             =   4050
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "15"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   135
         Top             =   4320
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "16½"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   136
         Top             =   4590
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "18"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   137
         Top             =   4860
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "19½"
         Height          =   195
         Index           =   18
         Left            =   180
         TabIndex        =   138
         Top             =   5130
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "21"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   139
         Top             =   5400
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "22½"
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   140
         Top             =   5670
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   195
         Index           =   21
         Left            =   180
         TabIndex        =   141
         Top             =   5940
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "25½"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   142
         Top             =   6210
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "27"
         Height          =   195
         Index           =   23
         Left            =   180
         TabIndex        =   143
         Top             =   6480
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "28½"
         Height          =   195
         Index           =   24
         Left            =   180
         TabIndex        =   144
         Top             =   6750
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   195
         Index           =   25
         Left            =   180
         TabIndex        =   151
         Top             =   7020
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "31½"
         Height          =   195
         Index           =   26
         Left            =   180
         TabIndex        =   154
         Top             =   7290
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "33"
         Height          =   195
         Index           =   27
         Left            =   180
         TabIndex        =   155
         Top             =   7560
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "34½"
         Height          =   195
         Index           =   28
         Left            =   180
         TabIndex        =   164
         Top             =   7830
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "36"
         Height          =   195
         Index           =   29
         Left            =   180
         TabIndex        =   165
         Top             =   8100
         Width           =   315
      End
   End
   Begin VB.Frame Left 
      Caption         =   "Left Leg"
      Height          =   8415
      Left            =   60
      TabIndex        =   215
      Top             =   0
      Width           =   3735
      Begin MSFlexGridLib.MSFlexGrid grdLeftInches 
         Height          =   8100
         Left            =   1320
         TabIndex        =   254
         Top             =   265
         Width           =   700
         _ExtentX        =   1244
         _ExtentY        =   14288
         _Version        =   65541
         Rows            =   30
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   29
         Left            =   2100
         TabIndex        =   54
         Top             =   8100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   28
         Left            =   2100
         TabIndex        =   53
         Top             =   7830
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   27
         Left            =   2100
         TabIndex        =   52
         Top             =   7560
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   26
         Left            =   2100
         TabIndex        =   51
         Top             =   7290
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   25
         Left            =   2100
         TabIndex        =   50
         Top             =   7020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   24
         Left            =   2100
         TabIndex        =   49
         Top             =   6750
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   23
         Left            =   2100
         TabIndex        =   48
         Top             =   6480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   22
         Left            =   2100
         TabIndex        =   47
         Top             =   6210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   21
         Left            =   2100
         TabIndex        =   46
         Top             =   5940
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   20
         Left            =   2100
         TabIndex        =   45
         Top             =   5670
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   19
         Left            =   2100
         TabIndex        =   44
         Top             =   5400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   18
         Left            =   2100
         TabIndex        =   43
         Top             =   5130
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   17
         Left            =   2100
         TabIndex        =   42
         Top             =   4860
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   16
         Left            =   2100
         TabIndex        =   41
         Top             =   4590
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   15
         Left            =   2100
         TabIndex        =   40
         Top             =   4320
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   14
         Left            =   2100
         TabIndex        =   39
         Top             =   4050
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   13
         Left            =   2100
         TabIndex        =   38
         Top             =   3780
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   12
         Left            =   2100
         TabIndex        =   37
         Top             =   3510
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   11
         Left            =   2100
         TabIndex        =   36
         Top             =   3240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   10
         Left            =   2100
         TabIndex        =   35
         Top             =   2970
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   9
         Left            =   2100
         TabIndex        =   34
         Top             =   2700
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   8
         Left            =   2100
         TabIndex        =   33
         Top             =   2430
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   7
         Left            =   2100
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   6
         Left            =   2100
         TabIndex        =   30
         Top             =   1890
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Frame Frame2 
         Caption         =   "Template"
         Height          =   915
         Left            =   2040
         TabIndex        =   207
         Top             =   180
         Width           =   1635
         Begin VB.ComboBox cboLeftTemplate 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   270
            Width           =   1395
         End
         Begin VB.CheckBox chkLeftZipper 
            Caption         =   "Ankle Zipper"
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   630
            Width           =   1515
         End
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   660
         MaxLength       =   7
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   660
         MaxLength       =   7
         TabIndex        =   1
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   660
         MaxLength       =   7
         TabIndex        =   2
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   660
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   660
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   660
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1590
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   660
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   660
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2130
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   660
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   9
         Left            =   660
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2670
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   10
         Left            =   660
         MaxLength       =   7
         TabIndex        =   10
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   11
         Left            =   660
         MaxLength       =   7
         TabIndex        =   11
         Top             =   3210
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   12
         Left            =   660
         MaxLength       =   7
         TabIndex        =   12
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   13
         Left            =   660
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3750
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   14
         Left            =   660
         MaxLength       =   7
         TabIndex        =   14
         Top             =   4020
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   15
         Left            =   660
         MaxLength       =   7
         TabIndex        =   216
         Top             =   4290
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   16
         Left            =   660
         MaxLength       =   7
         TabIndex        =   15
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   17
         Left            =   660
         MaxLength       =   7
         TabIndex        =   16
         Top             =   4830
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   18
         Left            =   660
         MaxLength       =   7
         TabIndex        =   17
         Top             =   5100
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   19
         Left            =   660
         MaxLength       =   7
         TabIndex        =   18
         Top             =   5370
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   20
         Left            =   660
         MaxLength       =   7
         TabIndex        =   19
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   21
         Left            =   660
         MaxLength       =   7
         TabIndex        =   20
         Top             =   5910
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   22
         Left            =   660
         MaxLength       =   7
         TabIndex        =   21
         Top             =   6180
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   23
         Left            =   660
         MaxLength       =   7
         TabIndex        =   22
         Top             =   6450
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   24
         Left            =   660
         MaxLength       =   7
         TabIndex        =   23
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   25
         Left            =   660
         MaxLength       =   7
         TabIndex        =   24
         Top             =   6990
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   26
         Left            =   660
         MaxLength       =   7
         TabIndex        =   25
         Top             =   7260
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   27
         Left            =   660
         MaxLength       =   7
         TabIndex        =   26
         Top             =   7530
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   28
         Left            =   660
         MaxLength       =   7
         TabIndex        =   27
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   29
         Left            =   660
         MaxLength       =   7
         TabIndex        =   28
         Top             =   8070
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid grdLeftDisplay 
         Height          =   6384
         Left            =   2640
         TabIndex        =   255
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   11271
         _Version        =   65541
         Rows            =   24
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label labLeftMaxStr 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   3180
         TabIndex        =   224
         Top             =   1260
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label labLeftDisp 
         Caption         =   "Max Stretch:"
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   225
         Top             =   1260
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Caption         =   "Red."
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   226
         Top             =   1650
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Caption         =   "Gms"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   227
         Top             =   1650
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Caption         =   "MMHg"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   228
         Top             =   1650
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   60
         Left            =   180
         TabIndex        =   174
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   59
         Left            =   180
         TabIndex        =   175
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   58
         Left            =   180
         TabIndex        =   176
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   57
         Left            =   180
         TabIndex        =   177
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   56
         Left            =   180
         TabIndex        =   178
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   195
         Index           =   55
         Left            =   180
         TabIndex        =   179
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   54
         Left            =   180
         TabIndex        =   180
         Top             =   1890
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   53
         Left            =   180
         TabIndex        =   181
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   52
         Left            =   180
         TabIndex        =   182
         Top             =   2430
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   51
         Left            =   180
         TabIndex        =   183
         Top             =   2700
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   50
         Left            =   180
         TabIndex        =   184
         Top             =   2970
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   195
         Index           =   49
         Left            =   180
         TabIndex        =   185
         Top             =   3240
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "10½"
         Height          =   195
         Index           =   48
         Left            =   180
         TabIndex        =   186
         Top             =   3510
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "12"
         Height          =   195
         Index           =   47
         Left            =   180
         TabIndex        =   187
         Top             =   3780
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "13½"
         Height          =   195
         Index           =   46
         Left            =   180
         TabIndex        =   188
         Top             =   4050
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "15"
         Height          =   195
         Index           =   45
         Left            =   180
         TabIndex        =   189
         Top             =   4320
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "16½"
         Height          =   195
         Index           =   44
         Left            =   180
         TabIndex        =   190
         Top             =   4590
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "18"
         Height          =   195
         Index           =   43
         Left            =   180
         TabIndex        =   191
         Top             =   4860
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "19½"
         Height          =   195
         Index           =   42
         Left            =   180
         TabIndex        =   192
         Top             =   5130
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "21"
         Height          =   195
         Index           =   41
         Left            =   180
         TabIndex        =   193
         Top             =   5400
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "22½"
         Height          =   195
         Index           =   40
         Left            =   180
         TabIndex        =   194
         Top             =   5670
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   195
         Index           =   39
         Left            =   180
         TabIndex        =   195
         Top             =   5940
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "25½"
         Height          =   195
         Index           =   38
         Left            =   180
         TabIndex        =   196
         Top             =   6210
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "27"
         Height          =   195
         Index           =   37
         Left            =   180
         TabIndex        =   197
         Top             =   6480
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "28½"
         Height          =   195
         Index           =   36
         Left            =   180
         TabIndex        =   198
         Top             =   6750
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   195
         Index           =   35
         Left            =   180
         TabIndex        =   199
         Top             =   7020
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "31½"
         Height          =   195
         Index           =   34
         Left            =   180
         TabIndex        =   200
         Top             =   7290
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "33"
         Height          =   195
         Index           =   33
         Left            =   180
         TabIndex        =   201
         Top             =   7560
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "34½"
         Height          =   195
         Index           =   32
         Left            =   180
         TabIndex        =   202
         Top             =   7830
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "36"
         Height          =   195
         Index           =   31
         Left            =   180
         TabIndex        =   203
         Top             =   8100
         Width           =   315
      End
   End
   Begin VB.TextBox txtRightTemplate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   214
      Top             =   5790
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidRightLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   212
      Top             =   4710
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   209
      Top             =   4350
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftTemplate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   206
      Top             =   5790
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   1395
      Left            =   3900
      TabIndex        =   204
      Top             =   2160
      Width           =   3435
      Begin VB.OptionButton optFabric 
         Caption         =   "Gradient"
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   60
         Top             =   870
         Width           =   1095
      End
      Begin VB.OptionButton optFabric 
         Caption         =   "Jobstex"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   59
         Top             =   870
         Width           =   1155
      End
      Begin VB.OptionButton optFabric 
         Caption         =   "Powernet"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   870
         Width           =   1215
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   330
         Width           =   2475
      End
      Begin VB.Label Label27 
         Caption         =   "Fabric:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   205
         Top             =   390
         Width           =   1035
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   3900
      TabIndex        =   166
      Top             =   0
      Width           =   3435
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   237
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   233
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   240
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2520
         TabIndex        =   241
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2520
         TabIndex        =   235
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   239
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   149
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "File No:"
         Height          =   195
         Left            =   120
         TabIndex        =   171
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   1980
         TabIndex        =   170
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Units:"
         Height          =   255
         Left            =   1980
         TabIndex        =   169
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   168
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   167
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   161
      Top             =   3990
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   158
      Top             =   5070
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtRightAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   157
      Top             =   5430
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   153
      Top             =   5430
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   150
      Top             =   5070
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidLeftLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      TabIndex        =   146
      Top             =   4710
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   145
      Top             =   3990
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   172
      Top             =   7860
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   223
      Top             =   7860
      Width           =   1215
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LeftFirstTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   253
      Top             =   7290
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RightFirstTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   252
      Top             =   7290
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LeffLastTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   251
      Top             =   7650
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RightLastTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   250
      Top             =   7650
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Body UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   245
      Top             =   4410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label11 
      Caption         =   "Messages"
      Height          =   195
      Left            =   3900
      TabIndex        =   243
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label labMessage 
      BorderStyle     =   1  'Fixed Single
      Height          =   3525
      Left            =   3900
      TabIndex        =   242
      Top             =   4110
      Width           =   3435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   229
      Top             =   6930
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   230
      Top             =   6930
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   231
      Top             =   6570
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   232
      Top             =   6570
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   234
      Top             =   6210
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   236
      Top             =   6210
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Templt"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   211
      Top             =   5850
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   213
      Top             =   4770
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   210
      Top             =   4410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Templt"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   208
      Top             =   5850
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DDE Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   163
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   162
      Top             =   4050
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lable14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   160
      Top             =   5130
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   159
      Top             =   5490
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   156
      Top             =   5490
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   152
      Top             =   5130
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5700
      TabIndex        =   148
      Top             =   4770
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   147
      Top             =   4050
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "whfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   WHFIGURE.MAK
'Form:      WHFIGURE.FRM
'Purpose:   Dialog to figure the ankles of both legs
'           for a Waist Height
'
'Version:   3.02
'Date:      16.Jan.95
'Author:    Gary George
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'27.Apr.95  GG      Bug Fix, For JOBSTEX_FL the last tape
'                   was not being found correctly
'Dec 98     GG      Port to VB5
'
'Notes:-
'
   'MsgBox constant
'    Const IDYES = 6
'    Const IDNO = 7
'    Const IDCANCEL = 2

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
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height - Figure"
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

Private Sub cboFabric_Click()
    PR_FigureLeftAnkle
    PR_FigureRightAnkle
End Sub

Private Sub chkLeftZipper_Click()
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
        If chkLeftZipper.Value = 1 Then
            cboLeftTemplate.ListIndex = 1  '9DS
        Else
            cboLeftTemplate.ListIndex = 0  '13DS
        End If
    End If
    PR_FigureLeftAnkle
End Sub

Private Sub chkRightZipper_Click()
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
        If chkRightZipper.Value = 1 Then
            cboRightTemplate.ListIndex = 1  '9DS
        Else
            cboRightTemplate.ListIndex = 0  '13DS
        End If
    End If
    PR_FigureRightAnkle
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
    Print #fNum, "//by Visual Basic, Waist Height - Figure"
    Print #fNum, "//type - "; sType; ""

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetColor"; QC; "Table("; QQ; "find"; QCQ; "color"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"

   'Display Hour Glass symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; sType; QQ; ");"


End Function

Private Function FN_POWERNET_Pressure(iReduction As Integer, nAnkleCir As Double, iModulus As Integer) As Integer
   'nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
   'This fuction is the reverse of FN_POWERNET_reduction
   '
   'Input
   '    iReduction  a Reduction in the range 10 to 32
   '    iModulus    Modulus of chosen fabric
   '                Eg Fabric = "Pow 210-3B Cream" => Modulus = 210.
   '    nAnkleCir   Ankle circumferance used with the derved grams to
   '                back calculat the pressue at the given reductio
   '
   'Globals
   '    POWERNET    The Fabric conversion chart loaded from file.
   '                Maps modulus to reduction at the given grams.
   '
   'Output
   '    nMMHg       Pressure reverse calclated from the conversion chart.
   '
   '
   'NOTE:-
   '    No range checking is done on the reduction values.
   '    Similarly the modulus is not checked.
   '    So it had better be right.
   '
    Dim sConversion As String
    Dim iVal As Integer, iGrams As Integer, ii As Integer

   'Get conversion string based on modulus
    sConversion = ""
    For ii = 0 To 17
        If iModulus = Val(POWERNET.Modulus(ii)) Then
            sConversion = POWERNET.Conversion(ii)
            Exit For
        End If
    Next ii

    iVal = iReduction - 10
    iGrams = Val(Mid$(sConversion, (iVal * 4) + 1, 4))
    
    FN_POWERNET_Pressure = round(iGrams / nAnkleCir)
    
End Function

Private Function FN_POWERNET_Reduction(iGrams As Integer, iModulus As Integer) As Integer
   'nReduction = FN_POWERNET_Reduction(nMMHg, nGrams, iModulus)
   'Input
   '    iGrams      Grams at Ankle
   '    iModulus    Modulus of chosen fabric
   '                Eg Fabric = "Pow 210-3B Cream" => Modulus = 210.
   '
   'Globals
   '    POWERNET    The Fabric conversion chart loaded from file.
   '                Maps modulus to reduction at the given grams.
   '
   'Output
   '    iReduction  Reduction established from the conversion chart.
   '                In the range 10 to 32
   '
   '
   'NOTE
   '    This fuction is derived from the DRAFIX function
   '    FNCalcReduction()
   '
    Dim sConversion As String
    Dim ii As Integer, iPrevVal As Integer, iVal As Integer

   'Get conversion string based on modulus
    sConversion = ""
    For ii = 0 To 17
        If iModulus = Val(POWERNET.Modulus(ii)) Then
            sConversion = POWERNET.Conversion(ii)
            Exit For
        End If
    Next ii

    If sConversion = "" Then
        FN_POWERNET_Reduction = -1000
        Exit Function
    End If

    iPrevVal = 0
    For ii = 0 To 22
        iVal = Val(Mid$(sConversion, (ii * 4) + 1, 4))
        If iVal >= iGrams Then Exit For
        iPrevVal = iVal
    Next ii
    
   'Default to a 10 reduction
    If ii = 0 Then
        FN_POWERNET_Reduction = 10
        Exit Function
    End If
    
   'Get reduction closest to given grams
    If (iGrams - iPrevVal) < (iVal - iGrams) Then
        FN_POWERNET_Reduction = ii + 9
    Else
        FN_POWERNET_Reduction = ii + 10
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

   'Dimensions

    FN_ValuesString = sString

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

  'Disable Time out mechanism
   Timer1.Enabled = False
   Dim ii As Integer, nn As Integer, nAge  As Integer, iFabricClass As Integer
   Dim iMMHg As Integer, iZipper As Integer, iValue As Integer
   Dim nValue As Double, nStretch As Double

   g_sDiagnosis = txtDiagnosis.Text

   g_iUidBody = txtUidBody
   g_sLastFabric = txtFabric.Text
   
   If fnGetNumber(txtLegStyle.Text, 1) = 2 Then  'Briefs
        Unload whfigure
        Load whfigbrf
        whfigbrf.Show
        Exit Sub
   End If

  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If
  
  'Establish how the leg is to be figured
  'Base this on chosen fabric and availability of mm
  'If it is not given explicitly
  
  'Check for an explititly given fabric class from the Ankle field
  'This Field will be emplty if no previous figuring has been done
  'Note This is a multi data. I the first number is -ve then there is no ankle.
  '     Also the function fnGetNumber returns -1 if the number does not exist.
  '
   iFabricClass = -1
   If fnGetNumber(txtLeftAnkle.Text, 1) <> -1 Then
        iFabricClass = fnGetNumber(txtLeftAnkle.Text, 8)
   End If
   If fnGetNumber(txtRightAnkle.Text, 1) <> -1 Then
        iFabricClass = fnGetNumber(txtRightAnkle.Text, 8)
   End If

  'Get fabric class from other data
   If iFabricClass < 0 Then
       If txtFabric.Text = "" Then
            If txtDiagnosis.Text = "Burns" Then
                g_POWERNET = True
            Else
                If Left$(txtDiagnosis.Text, 5) = "Lymph" Then
                    g_JOBSTEX_FL = True
                Else
                    g_JOBSTEX = True
                End If
           End If
       Else
           If Left$(txtFabric.Text, 3) = "Pow" Then
                    g_POWERNET = True
           Else
                If txtLeftMMs.Text <> "" Or txtRightMMs.Text <> "" Then
                    g_JOBSTEX_FL = True
                Else
                    g_JOBSTEX = True
                End If
           End If
       End If
   Else
        If iFabricClass = 0 Then g_POWERNET = True
        If iFabricClass = 1 Then g_JOBSTEX = True
        If iFabricClass = 2 Then g_JOBSTEX_FL = True
   End If
   
   If g_JOBSTEX = True Then optFabric(1).Value = True
   If g_JOBSTEX_FL = True Then optFabric(2).Value = True

  'Setup fabric dropdown box
   If g_POWERNET = True Then
        optFabric(0).Value = True
        PR_GetComboListFromFile cboFabric, g_sPathJOBST + "\WHFABRIC.DAT"
        PR_LoadFabricFromFile g_sPathJOBST + "\TEMPLTS\POWERNET.DAT"
   End If
   
   If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
       'Jobstex fabric
        cboFabric.AddItem "53 - JOBSTEX"
        cboFabric.AddItem "55 - JOBSTEX"
        cboFabric.AddItem "57 - JOBSTEX"
        cboFabric.AddItem "63 - JOBSTEX"
        cboFabric.AddItem "65 - JOBSTEX"
        cboFabric.AddItem "67 - JOBSTEX"
        cboFabric.AddItem "73 - JOBSTEX"
        cboFabric.AddItem "75 - JOBSTEX"
        cboFabric.AddItem "77 - JOBSTEX"
        cboFabric.AddItem "83 - JOBSTEX"
        cboFabric.AddItem "85 - JOBSTEX"
        cboFabric.AddItem "87 - JOBSTEX"
   End If

  'Set fabric value
  'NB   It might be that the fabric originally used is not one on the
  '     current fabric list.  In this case add the given fabric to
  '     the start of the list
   For ii = 0 To (cboFabric.ListCount - 1)
        If cboFabric.List(ii) = txtFabric.Text Then
            cboFabric.ListIndex = ii
        End If
   Next ii
   If txtFabric.Text <> "" And cboFabric.ListIndex = -1 Then
        cboFabric.AddItem txtFabric.Text, 0
        cboFabric.ListIndex = 0
   End If

  'Update tape boxes
   grdLeftInches.Col = 0
   For ii = 0 To 29
        nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
        If nValue > 0 Then
            txtLeft(ii) = nValue
            nValue = fnDisplayToInches(nValue)
            grdLeftInches.Row = ii
            grdLeftInches.Text = fnInchesToText(nValue)
        End If
   Next ii

   grdRightInches.Col = 0
   For ii = 0 To 29
        nValue = Val(Mid(txtRightLengths.Text, (ii * 4) + 1, 4)) / 10
        If nValue > 0 Then
            txtRight(ii) = nValue
            nValue = fnDisplayToInches(nValue)
            grdRightInches.Row = ii
            grdRightInches.Text = fnInchesToText(nValue)
        End If
   Next ii
   
   PR_EstablishAnkles

   If g_POWERNET = True Then PR_EnablePOWERNET
   If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX
   If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX

  'From the ankle tapes extract the saved ankle figuring
  'We have to be careful in the order in which this is set up as in the case
  'of JOBSTEX_FL we do not want to cause recaculation on the ankle.
  'We also must take care only to add values only if they exist
  '(thus we check that the returned value from fnGetNumber is not -ve)
  '
   Dim iFabric%
   If txtLeftAnkle.Text <> "" And g_iLtAnkle <> 0 Then
        iZipper = fnGetNumber(txtLeftAnkle.Text, 7)
        If iZipper >= 0 Then chkLeftZipper.Value = iZipper  'Order is v.important here
        iMMHg = fnGetNumber(txtLeftAnkle.Text, 2)
        If iMMHg >= 0 Then
            g_iLtMM(g_iLtAnkle) = iMMHg
            txtLeftMM(g_iLtAnkle) = iMMHg
        End If
        If g_JOBSTEX_FL = True Then
            iFabric = Val(Left$(cboFabric.Text, 2))
            nStretch = fnGetNumber(txtLeftAnkle.Text, 3)
            If nStretch >= 0 Then
                    PR_DisplayFiguredAnkle "Left", nStretch, g_nLtLastAnkle, g_nLtLastHeel, iFabric
                    g_iLtStretch(g_iLtAnkle) = nStretch
            End If
        End If
   End If

   If txtRightAnkle.Text <> "" And g_iRtAnkle <> 0 Then
        iZipper = fnGetNumber(txtRightAnkle.Text, 7)
        If iZipper >= 0 Then chkRightZipper.Value = iZipper
        iMMHg = fnGetNumber(txtRightAnkle.Text, 2)
        If iMMHg >= 0 Then
            g_iRtMM(g_iRtAnkle) = iMMHg
            txtRightMM(g_iRtAnkle) = iMMHg
        End If
        If g_JOBSTEX_FL = True Then
            iFabric = Val(Left$(cboFabric.Text, 2))
            nStretch = fnGetNumber(txtRightAnkle.Text, 3)
            If nStretch >= 0 Then
                    PR_DisplayFiguredAnkle "Right", nStretch, g_nRtLastAnkle, g_nRtLastHeel, iFabric
                    g_iRtStretch(g_iRtAnkle) = nStretch
            End If
        End If
   End If

  'Don't figure ankle for fabric class JOBSTEX_FL as the user can modify the pressures
  'at each leg tape manually and these are stored.  Figuring would overwrite
  'any manual changes
   If g_JOBSTEX_FL <> True Then
        PR_FigureLeftAnkle
        PR_FigureRightAnkle
   Else
        If txtRightMMs.Text <> "" And g_iRtAnkle <> 0 Then
            For ii = g_iRtAnkle + 1 To 29
                iValue = Val(Mid(txtRightMMs.Text, (ii * 3) + 1, 3))
                If iValue > 0 Then
                    txtRightMM(ii) = iValue
                    g_iRtMM(ii) = iValue
                    g_iRtStretch(ii) = Val(Mid(txtRightStr.Text, (ii * 3) + 1, 3))
                    g_iRtRed(ii) = Val(Mid(txtRightRed.Text, (ii * 3) + 1, 3))
                    grdRightDisplay.Row = ii - 6 'Display from ankle
                    grdRightDisplay.Col = 0
                    grdRightDisplay = Str$(g_iRtStretch(ii))
                    grdRightDisplay.Col = 1
                    grdRightDisplay = Str$(g_iRtRed(ii))
                End If
            Next ii
        End If
        If txtLeftMMs.Text <> "" And g_iLtAnkle <> 0 Then
            For ii = g_iLtAnkle + 1 To 29
                iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
                If iValue > 0 Then
                    txtLeftMM(ii) = iValue
                    g_iLtMM(ii) = iValue
                    g_iLtStretch(ii) = Val(Mid(txtLeftStr.Text, (ii * 3) + 1, 3))
                    g_iLtRed(ii) = Val(Mid(txtLeftRed.Text, (ii * 3) + 1, 3))
                    grdLeftDisplay.Row = ii - 6 'Display from ankle
                    grdLeftDisplay.Col = 0
                    grdLeftDisplay = Str$(g_iLtStretch(ii))
                    grdLeftDisplay.Col = 1
                    grdLeftDisplay = Str$(g_iLtRed(ii))
                End If
            Next ii
        End If
        
   End If
   
  'Store initial values for use in Cancel_Click
   g_sChangeChecker = FN_ValuesString()
   
   Show
   Screen.MousePointer = 0    ' Change pointer to default.
   whfigure.SetFocus
   
    
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Dim ii As Integer
    Hide
   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "The Figure Module is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting Figure"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - whfigure.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - whfigure.Height) / 2  ' Center form vertically.
    
    Set MainForm = whfigure

    g_nUnitsFac = 1 'Default to inches
    g_JOBSTEX = False
    g_JOBSTEX_FL = False
    g_POWERNET = False

   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    g_sPathJOBST = fnPathJOBST()

   'Setup display inches grid
    grdLeftInches.ColWidth(0) = 700
    grdRightInches.ColWidth(0) = 700
    grdRightInches.ColAlignment(0) = 2
    grdLeftInches.ColAlignment(0) = 2
    
    For ii = 0 To 29
        grdLeftInches.RowHeight(ii) = 266
        grdRightInches.RowHeight(ii) = 266
    Next ii
        grdRightDisplay.ColWidth(0) = 488
        grdRightDisplay.ColWidth(1) = 488

   'Setup display of results grid
    For ii = 0 To 1
        grdLeftDisplay.ColWidth(ii) = 488
        grdRightDisplay.ColWidth(ii) = 488
        grdLeftDisplay.ColAlignment(ii) = 2
        grdRightDisplay.ColAlignment(ii) = 2
    Next ii
    
    For ii = 0 To 23
        grdLeftDisplay.RowHeight(ii) = 266
        grdRightDisplay.RowHeight(ii) = 266
    Next ii

   'Enable timer
   'This is a time out mechanism that will
   'end the VB programme if the DDE link with DRAFIX
   'is not sucessfully completed
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
            MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height - Figure"
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub optFabric_Click(index As Integer)
   'Allows the selection of fabric class
   'Used to override the fabric class selection established at link close

   'Clean Message box
    labMessage.Caption = ""
 
   'Do nothing if same fabric class selected
    If g_POWERNET = True And index = 0 Then Exit Sub
    If g_JOBSTEX = True And index = 1 Then Exit Sub
    If g_JOBSTEX_FL = True And index = 2 Then Exit Sub
   
   'Restablish ankles just in case they have changed
    PR_EstablishAnkles

   'Set the display for the different fabric classes
   'We must first reset the form to the blank form that we use on form load
    PR_ResetFormWHFIGURE
    
    g_POWERNET = False
    g_JOBSTEX = False
    g_JOBSTEX_FL = False

    Select Case index
     Case 0
       'POWERNET
        g_POWERNET = True
        PR_EnablePOWERNET
     Case 1
       'JOBSTEX
        g_JOBSTEX = True
        PR_EnableJOBSTEX
     Case 2
       'JOBSTEX_FL, Pressure calulated at every leg tape
        g_JOBSTEX_FL = True
        PR_EnableJOBSTEX
        PR_EnableFL_JOBSTEX
    End Select

  'Setup fabric dropdown box
   cboFabric.Clear
   If g_POWERNET = True Then
        PR_GetComboListFromFile cboFabric, g_sPathJOBST + "\WHFABRIC.DAT"
        PR_LoadFabricFromFile g_sPathJOBST + "\TEMPLTS\POWERNET.DAT"
   End If
   
   If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
       'Jobstex fabric
        cboFabric.AddItem "53 - JOBSTEX"
        cboFabric.AddItem "55 - JOBSTEX"
        cboFabric.AddItem "57 - JOBSTEX"
        cboFabric.AddItem "63 - JOBSTEX"
        cboFabric.AddItem "65 - JOBSTEX"
        cboFabric.AddItem "67 - JOBSTEX"
        cboFabric.AddItem "73 - JOBSTEX"
        cboFabric.AddItem "75 - JOBSTEX"
        cboFabric.AddItem "77 - JOBSTEX"
        cboFabric.AddItem "83 - JOBSTEX"
        cboFabric.AddItem "85 - JOBSTEX"
        cboFabric.AddItem "87 - JOBSTEX"
   End If

   'If we have values for Ankle Pressure and fabric lets use them
    If g_iLtLastFabric < cboFabric.ListCount Then
        cboFabric.ListIndex = g_iLtLastFabric
        g_iLtLastFabric = -1  'This will force re-figuring of Left
        g_iRtLastFabric = -1  'This will force re-figuring of Right
        If g_iLtMM(g_iLtAnkle) > 0 Then
            txtLeftMM(g_iLtAnkle) = g_iLtMM(g_iLtAnkle)
            PR_FigureLeftAnkle
        End If
        If g_iRtMM(g_iRtAnkle) > 0 Then
            txtRightMM(g_iRtAnkle) = g_iRtMM(g_iRtAnkle)
            PR_FigureRightAnkle
        End If
        g_iLtLastFabric = -1
        g_iRtLastFabric = -1
    End If

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
    PR_PutLine "HANDLE hMPD, hBody, hLeftLeg, hRightLeg;"
   
    Update_DDE_Text_Boxes

    Dim sSymbol As String
    
    sSymbol = "waistbody"

    If txtUidBody.Text <> "" Then
       'Update the Waist BODY symbol with the fabric
        PR_PutLine "hBody = UID (" & QQ & "find" & QC & Val(txtUidBody.Text) & ");"
        PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"
        PR_PutLine "SetDBData( hBody" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");"
    End If

    If txtUidRightLeg <> "" Then
       'Update the Right Leg symbol
        PR_PutLine "hRightLeg = UID (" & QQ & "find" & QC & Val(txtUidRightLeg.Text) & ");"
        PR_PutLine "if (!hRightLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"
        
        PR_PutLine "SetDBData( hRightLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid$(txtRightLengths.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid$(txtRightLengths.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hRightLeg" & CQ & "TapeMMs" & QCQ & Mid$(txtRightMMs.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "TapeMMs2" & QCQ & Mid$(txtRightMMs.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hRightLeg" & CQ & "Reduction" & QCQ & Mid$(txtRightRed.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "Reduction2" & QCQ & Mid$(txtRightRed.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hRightLeg" & CQ & "Grams" & QCQ & Mid$(txtRightStr.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "Grams2" & QCQ & Mid$(txtRightStr.Text, 61) & QQ & ");"
    
        PR_PutLine "SetDBData( hRightLeg" & CQ & "AnkleTape" & QCQ & txtRightAnkle.Text & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "FirstTape" & QCQ & txtRightFirstTape.Text & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "LastTape" & QCQ & txtRightLastTape.Text & QQ & ");"

        PR_PutLine "SetDBData( hRightLeg" & CQ & "Pressure" & QCQ & txtRightTemplate.Text & QQ & ");"
        PR_PutLine "SetDBData( hRightLeg" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");"
    
    End If

    If txtUidLeftLeg <> "" Then
       'Update Left leg symbol
        PR_PutLine "hLeftLeg = UID (" & QQ & "find" & QC & Val(txtUidLeftLeg.Text) & ");"
        PR_PutLine "if (!hLeftLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"
        
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid$(txtLeftLengths.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid$(txtLeftLengths.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "TapeMMs" & QCQ & Mid$(txtLeftMMs.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "TapeMMs2" & QCQ & Mid$(txtLeftMMs.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Reduction" & QCQ & Mid$(txtLeftRed.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Reduction2" & QCQ & Mid$(txtLeftRed.Text, 61) & QQ & ");"
        
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Grams" & QCQ & Mid$(txtLeftStr.Text, 1, 60) & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Grams2" & QCQ & Mid$(txtLeftStr.Text, 61) & QQ & ");"
    
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "AnkleTape" & QCQ & txtLeftAnkle.Text & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "FirstTape" & QCQ & txtLeftFirstTape.Text & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "LastTape" & QCQ & txtLeftLastTape.Text & QQ & ");"

        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Pressure" & QCQ & txtLeftTemplate.Text & QQ & ");"
        PR_PutLine "SetDBData( hLeftLeg" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");"
    
    End If


    Close #fNum
   
End Sub

Private Sub PR_DisplayErrorMessage(iErrorNum, sContext)
   'Procedure to display error messages
   '
   '    iErrorNum   Local error number
   '    sContext    User supplied string to be displayed in the Message Label
   '
   'NOTE
   '    In many cases values returned fron functions are -ve to indicate that
   '    an error has occured.
   '    This meams that an error can be quickly established without having to
   '    know the actual error details, mearly by checking if the returned value
   '    is -ve.
   '    This works well in this particular set of modules as therv are very few
   '    time that a -ve value is required to be returned.
   '
   '    All error numbers are local to this Module.
   '
   '
   '

    Dim iError As Integer
    Dim sError As String, NL As String

    NL = Chr$(13)   'New Line character
   'Within this procedure the error number is always +ve.
    iError = Abs(iErrorNum)

    sError = ""
    Select Case iError
    'Ankle selection
     Case 900
        sError = sError + "As one ankle is at the +1½ tape and the other is at the +3 tape. "
        sError = sError + "Figuring for both Ankles at the +3 tape."
    'Ankle figuring
     Case 1000
        sError = "The modulus for the chosen fabric " + cboFabric.List(cboFabric.ListIndex) + NL
        sError = sError + "Not listed in the POWERNET conversion chart."
     Case 1001
        sError = "The given Pressure resulted in a reduction of less than 14. "
        sError = sError + "Using a 14 reduction and back calculating the Pressure"
     Case 1002
        sError = "The given Pressure resulted in a reduction of over 26. "
        sError = sError + "As the Diagnosis is for Burns, "
        sError = sError + "using a 26 reduction and back calculating the Pressure"
     Case 1003
        sError = "The given Pressure resulted in a reduction of over 32. "
        sError = sError + "As the Diagnosis is other than Burns, "
        sError = sError + "using a 32 reduction and back calculating the Pressure"
     Case 1004
        sError = "At this reduction the 100% stretch is exceeded. "
        sError = sError + "Decrease the reduction."
     Case 2001
        sError = "Maximum Stretch exceeded. "
     Case 2002
        sError = "Minimum Stretch exceeded. "
     Case 2003
        sError = "Donning Stretch of 110% exceeded. "
     Case 2004
        sError = "Reduction exceeds 32. "
     Case Else
        sError = "Unknown Error" + NL
    End Select
    
   'If a message is in the box then add
    If labMessage.Caption = "" Then
        labMessage.Caption = sContext + NL + sError
    Else
        labMessage.Caption = labMessage.Caption + NL + NL + sContext + NL + sError
    End If
    Beep

End Sub

Private Sub PR_DisplayFiguredAnkle(sLeg As String, nStretch As Double, nAnkleCir As Double, nHeelCir As Double, iFabric As Integer)
    Dim nPatternDim As Double, iReduction As Integer, iMaxStretch  As Integer
    Dim sContext As String, iDonningStretch  As Integer, iMinStretch  As Integer

    nPatternDim = (nAnkleCir / (1 + (0.01 * nStretch))) / g_nUnitsFac
    iReduction = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
    iDonningStretch = round((((nHeelCir / g_nUnitsFac) - nPatternDim) / nPatternDim) * 100)

    If sLeg = "Left" Then
        grdLeftDisplay.Row = g_iLtAnkle - 6
        grdLeftDisplay.Col = 0
       'Stretch
        grdLeftDisplay = Str$(nStretch)
       'Reduction
        grdLeftDisplay.Col = 1
        grdLeftDisplay = Str$(iReduction)
        g_iLtRed(g_iLtAnkle) = iReduction
       'Max Stretch
        labLeftMaxStr = iDonningStretch
    Else
        grdRightDisplay.Row = g_iRtAnkle - 6
        grdRightDisplay.Col = 0
       'Stretch
        grdRightDisplay = Str$(nStretch)
       'Reduction
        grdRightDisplay.Col = 1
        grdRightDisplay = Str$(iReduction)
        g_iRtRed(g_iRtAnkle) = iReduction
       'Max Stretch
        labRightMaxStr = iDonningStretch
    End If
       
      'Set Max and Min stretches
       iMaxStretch = 44
       If iFabric <= 55 Then iMinStretch = 16 Else iMinStretch = 18
       
       If nStretch < iMinStretch Then
            PR_DisplayErrorMessage 2002, sLeg & " Ankle Figuring"
       End If

       If nStretch > iMaxStretch Then
            PR_DisplayErrorMessage 2001, sLeg & " Ankle Figuring"
       End If
       
       If iDonningStretch > 110 Then
            PR_DisplayErrorMessage 2003, sLeg & " Ankle Figuring"
       End If
       
       If iReduction > 32 Then
            PR_DisplayErrorMessage 2004, sLeg & " Ankle Figuring"
       End If
       


End Sub

Private Sub PR_EnableFL_JOBSTEX()
   'Setup the fields to allow calculation of
   'stretch at each tape
    Dim ii As Integer, iValue  As Integer
   
   'Right Leg
    If g_iRtAnkle <> 0 Then
        For ii = 8 To 29
            txtRightMM(ii).Visible = True
        Next ii
        txtRightMM(g_iRtAnkle + 1).Visible = True
       
       'Load MM data from string
       'Ankle data will be stored seperatly
        If txtRightMMs.Text <> "" Then
            For ii = g_iRtAnkle + 1 To 29
                iValue = Val(Mid(txtRightMMs.Text, (ii * 3) + 1, 3))
                If iValue > 0 Then
                    txtRightMM(ii) = iValue
                End If
            Next ii
        End If

    End If

   'Left Leg
    If g_iLtAnkle <> 0 Then
       'Make Boxes visible
        For ii = 8 To 29
            txtLeftMM(ii).Visible = True
        Next ii
        
        txtLeftMM(g_iLtAnkle + 1).Visible = True

       'Load MM data from string
       'Ankle data will be stored seperatly
        If txtLeftMMs.Text <> "" Then
            For ii = g_iLtAnkle + 1 To 29
                iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
                If iValue > 0 Then
                    txtLeftMM(ii) = iValue
                End If
            Next ii
        End If
    End If
    
End Sub

Private Sub PR_EnableJOBSTEX()
   'Set up the display for JOBSTEX fabric
   '
    Dim ii As Integer, nn As Integer, iDisable As Integer, iEnable As Integer
    Dim nValue As Double

    cboLeftTemplate.Clear
    cboLeftTemplate.AddItem "13 DS"
    cboLeftTemplate.AddItem "09 DS"
    cboLeftTemplate.Enabled = False

    cboRightTemplate.Clear
    cboRightTemplate.AddItem "13 DS"
    cboRightTemplate.AddItem "09 DS"
    cboRightTemplate.Enabled = False


   'Set value
    For ii = 0 To (cboLeftTemplate.ListCount - 1)
        If cboLeftTemplate.List(ii) = txtLeftTemplate.Text Then
            cboLeftTemplate.ListIndex = ii
        End If
    Next ii

    For ii = 0 To (cboRightTemplate.ListCount - 1)
        If cboRightTemplate.List(ii) = txtRightTemplate.Text Then
            cboRightTemplate.ListIndex = ii
        End If
    Next ii

  'Set the display for each leg depending on the ankle and fabric and options.
  'g_JOBSTEX_FL is a special case where all of the stretch at tapes above the
  'ankle are calculated
  '
  'Panty Legs
  'If a template is not given then
  'then set 13DS for JOBSTEX, this is first item in list
  '

   If g_iRtAnkle = 0 Then
        If cboRightTemplate.ListIndex = -1 Then cboRightTemplate.ListIndex = 0
        cboRightTemplate.Enabled = False
        chkRightZipper.Enabled = False
        chkRightZipper.Value = 0
   End If

   If g_iLtAnkle = 0 Then
        If cboLeftTemplate.ListIndex = -1 Then cboLeftTemplate.ListIndex = 0
        cboLeftTemplate.Enabled = False
        chkLeftZipper.Enabled = False
        chkLeftZipper.Value = 0
   End If


   If g_iRtAnkle <> 0 Then
       'Set display headings
        labRightMaxStr.Visible = True
        labRightDisp(1).Caption = "Str"
        For ii = 0 To 3
            labRightDisp(ii).Visible = True
        Next

       'Set display value depending on Ankle position
        iEnable = g_iRtAnkle
        If iEnable = 6 Then
            iDisable = 7
        Else
            iDisable = 6
        End If

        txtRightMM(iEnable).Visible = True
        txtRightMM(iDisable).Visible = False

   End If


   If g_iLtAnkle <> 0 Then
        labLeftMaxStr.Visible = True
        labLeftDisp(1).Caption = "Str"
        For ii = 0 To 3
            labLeftDisp(ii).Visible = True
        Next
       'Set display value depending on Ankle position
        iEnable = g_iLtAnkle
        If iEnable = 6 Then
            iDisable = 7
        Else
            iDisable = 6
        End If

        txtLeftMM(iEnable).Visible = True
        txtLeftMM(iDisable).Visible = False
   End If


End Sub

Private Sub PR_EnablePOWERNET()

    Dim ii As Integer, nn As Integer, iDisable As Integer, iEnable As Integer
    Dim nValue As Double

    cboLeftTemplate.Clear
    cboLeftTemplate.AddItem "30mm Hg"
    cboLeftTemplate.AddItem "35mm Hg"
    cboLeftTemplate.AddItem "40mm Hg"
    cboLeftTemplate.AddItem "50mm Hg"

    cboRightTemplate.Clear
    cboRightTemplate.AddItem "30mm Hg"
    cboRightTemplate.AddItem "35mm Hg"
    cboRightTemplate.AddItem "40mm Hg"
    cboRightTemplate.AddItem "50mm Hg"
    
    cboLeftTemplate.Enabled = True
    cboRightTemplate.Enabled = True
   
   'Set value
    For ii = 0 To (cboLeftTemplate.ListCount - 1)
        If cboLeftTemplate.List(ii) = txtLeftTemplate.Text Then
            cboLeftTemplate.ListIndex = ii
        End If
    Next ii

    For ii = 0 To (cboRightTemplate.ListCount - 1)
        If cboRightTemplate.List(ii) = txtRightTemplate.Text Then
            cboRightTemplate.ListIndex = ii
        End If
    Next ii
 
  'Set the display for each leg depending on the ankle and fabric and options.
  '     g_XtAnkle = 0   => Panty leg
  '     g_XtAnkle = 6   => Ankle at +1-1/2
  '     g_XtAnkle = 0   => Ankle at +3
  '
  
  'Panty Legs
  'If a template is not given then
  'then set to 35mm Hg for Powernet

   If g_iRtAnkle = 0 Then
        If cboRightTemplate.ListIndex = -1 Then cboRightTemplate.ListIndex = 0
        chkRightZipper.Enabled = False
        chkRightZipper.Value = 0
   End If

   If g_iLtAnkle = 0 Then
        If cboLeftTemplate.ListIndex = -1 Then cboLeftTemplate.ListIndex = 0
        chkLeftZipper.Enabled = False
        chkLeftZipper.Value = 0
   End If

   If g_iRtAnkle <> 0 Then
        labRightDisp(1).Caption = "Gms"
        labRightDisp(3).Visible = False
        labRightMaxStr.Visible = False
        For ii = 0 To 2
            labRightDisp(ii).Visible = True
        Next

       'Set display value depending on Ankle position
        iEnable = g_iRtAnkle
        If iEnable = 6 Then
            iDisable = 7
        Else
            iDisable = 6
        End If
        
        txtRightMM(iEnable).Visible = True
        txtRightMM(iDisable).Visible = False
   End If

   If g_iLtAnkle <> 0 Then
        labLeftDisp(1).Caption = "Gms"
        labLeftDisp(3).Visible = False
        labLeftMaxStr.Visible = False
        For ii = 0 To 2
            labLeftDisp(ii).Visible = True
        Next

       'Set display value depending on Ankle position
        iEnable = g_iLtAnkle
        If iEnable = 6 Then
            iDisable = 7
        Else
            iDisable = 6
        End If
        
        txtLeftMM(iEnable).Visible = True
        txtLeftMM(iDisable).Visible = False
   End If

End Sub

Private Sub PR_EstablishAnkles()
   'For both legs
   'Setup global variables.

    Dim ii As Integer, nn As Integer
    Dim nValue As Double

  'Find ankle tapes
  'Note:
  '     This depends on the heel dimension.
  '     For a heel of less than 9" the ankle is at the +1-1/2 tape
  '     otherwise it is at the +3 tape.
  '     The exception to this rule is that if one leg is at 1-1/2 and one
  '     is at +3, then both legs must be figured at the +3 tape.
  '
  '     Later a check is made to see if there are enough tapes to draw the
  '     foot properley, at this point we are not interested we just want the
  '     ankles.
  '
  '     g_XtAnkle = 0   => Panty leg
  '     g_XtAnkle = 6   => Ankle at +1-1/2
  '     g_XtAnkle = 7   => Ankle at +3
  '
  ' Note also use of LAST global variables

   g_iLtAnkle = 0
   g_nLtLastAnkle = 0
   g_nLtLastHeel = fnDisplayToInches(Val(txtLeft(5).Text))
   If g_nLtLastHeel <> 0 Then
        If g_nLtLastHeel < 9 Then
            g_iLtAnkle = 6    'Small heel Ankle at +1 1/2 tape
        Else
            g_iLtAnkle = 7    'Large heel Ankle at +3 tape
        End If
   End If

   g_iRtAnkle = 0
   g_nRtLastAnkle = 0
   g_nRtLastHeel = fnDisplayToInches(Val(txtRight(5).Text))
   If g_nRtLastHeel <> 0 Then
        If g_nRtLastHeel < 9 Then
            g_iRtAnkle = 6    'Small heel Ankle at +1 1/2 tape
        Else
            g_iRtAnkle = 7    'Large heel Ankle at +3 tape
        End If
   End If
   
  'Check for asymetric ankles
   If g_iRtAnkle <> g_iLtAnkle And g_iRtAnkle <> 0 And g_iLtAnkle <> 0 Then
        g_iLtAnkle = 7    'Large heel Ankle at +3 tape
        g_iRtAnkle = 7    'Large heel Ankle at +3 tape
        PR_DisplayErrorMessage 900, "Ankle selection"
   End If

   g_nLtLastAnkle = fnDisplayToInches(Val(txtLeft(g_iLtAnkle)))
   g_nRtLastAnkle = fnDisplayToInches(Val(txtRight(g_iRtAnkle)))

End Sub

Private Sub PR_FigureLeftAnkle()
   'Figures the Pressure/Stretch at the LEFT ankle
   'Input from form:-
   '   Left(5-7) where     Left(5) = Heel Tape
   '                       Left(6) = Ankle Tape (Heel < 9")
   '                       Left(7) = Ankle Tape (Heel > 9")
   '   cboFabric
   '   chkLeftZipper
   '   txtLeftMM(g_iLtAnkle)
   '   txtLeftStretch(g_iLtAnkle)   JOBSTEX fabric only
   '
   'Input from globals
   '   g_JOBSTEX           Flag to indicate use of JOBSTEX fabric
   '   g_JOBSTEX_FL        Flag to indicate that JOBSTEX fabric in use
   '                       and that the stretch will be calculated at all
   '                       leg tape
   '   g_POWERNET          Flag to indicate use of Powernet fabric
   '
   '   g_iLtAnkle          Indicates the Left ankle tape
   '   g_nLtLastHeel
   '   g_nLtLastAnkle
   '   g_iLtLastMM
   '   g_iLtLastStretch    JOBSTEX fabric only
   '   g_iLtLastZipper
   '   g_iLastFabric
   '
   'JOBSTEX fabric Input and Output
   '   txtLeftMM(g_iLtAnkle)        User Input, if txtLeftStretch has been modified.
   '   txtLeftStretch(g_iLtAnkle)   User Input, if txtLeftMM has been modified.
   '   labLeftRed(g_iLtAnkle)       Displays Reduction
   '   labLeftMaxStr(2)             Displays MaxStretch
   '
   'Note:-
   '   The user inputs a required pressure in MMHg the stretch is then calculated
   '   and the stretch dependant values displayed.
   '   Alternately the stretch can be entered and the pressure in MMHg calculated.
   '   Given input of both (eg at link close) then MMHg takes precidence
   '
   'POWERNET fabric Input and Output
   '   txtLeftMM(g_iLtAnkle)            User Input
   '   txtLeftStretch(g_iLtAnkle)       not used (disabled)
   '   labLeftFiguredGrams(g_iLtAnkle)  Displays Grams
   '   labLeftRed(g_iLtAnkle)           Displays Reduction
   'Note:-
   '   The user inputs a required pressure in MMHg the reduction is then calculated
   '   and the reduction dependant values displayed.
   '
   '
    Dim nMMHg As Double, nStretch As Double
    Dim nAnkleCir As Double, nHeelCir As Double
    Dim nTapeCir As Double
    Dim iZipper As Integer, iFabric As Integer, iPrevMMHg As Integer
    Dim iModulus As Integer, iGrams As Integer, iReduction As Integer
    Dim ii As Integer
    Dim nLastTape As Integer, nLastTapeMMHg As Integer, nLastTapeCir As Double
    Dim nPatternDim As Double

   'Clean Messaage box
    labMessage.Caption = ""

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly

   'Leg tapes at heel and ankles and Fabric
    If Val(txtLeft(5).Text) = 0 Or Val(txtLeft(g_iLtAnkle).Text) = 0 Or cboFabric.ListIndex = -1 Then
        Exit Sub
    End If
   
   'MMHg for POWERNET fabric
    If g_POWERNET = True And Val(txtLeftMM(g_iLtAnkle)) = 0 Then Exit Sub
    
   'MMHg and Stretch for JOBSTEX fabric
    If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle)) = 0 Then Exit Sub

   'Fabric Selection
   'NOTE:-
   '    The POWERNET fabric is based on the file g_sPathJOBST + "\WHFABRIC.DAT"
   '    For clarity this file can contain blank lines.  Therefor we must check
   '    that the fabric selected is not blank.
    If cboFabric.List(cboFabric.ListIndex) = "" Then Exit Sub

   'Establish if any of the inputs have changed compared to the last stored
   'values.
   'If they have changed then recalculate and redisplay
   'NB  Use of GOTO, (structured programmers get stuffed!)

   'Heel Tape
    If fnDisplayToInches(Val(txtLeft(5))) <> g_nLtLastHeel Then GoTo CALC
   
   'Ankle Tape
    If fnDisplayToInches(Val(txtLeft(g_iLtAnkle))) <> g_nLtLastAnkle Then GoTo CALC
   
   'Pressure MMHg
    If Val(txtLeftMM(g_iLtAnkle).Text) <> g_iLtLastMM Then GoTo CALC

   'Zippers
    If chkLeftZipper.Value <> g_iLtLastZipper Then GoTo CALC
    
   'Fabric
    If cboFabric.ListIndex <> g_iLtLastFabric Then GoTo CALC

   'If nothing has changed then Exit this sub
    Exit Sub

   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
CALC:
   'Calculate stretch based on given pressure (MMHg).
   'NB  All dimensions are converted to in inches except that last dimesions
   '    are stored using their display values.
    
    nHeelCir = fnDisplayToInches(Val(txtLeft(5).Text))
    g_nLtLastHeel = nHeelCir
        
    nAnkleCir = fnDisplayToInches(Val(txtLeft(g_iLtAnkle).Text))
    g_nLtLastAnkle = nAnkleCir

    iZipper = chkLeftZipper.Value
    g_iLtLastZipper = iZipper

    g_iLtLastFabric = cboFabric.ListIndex
    nLastTape = -1
    
    nMMHg = Val(txtLeftMM(g_iLtAnkle).Text)
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
      'JOBSTEX - Fabric
       iFabric = Val(Left$(cboFabric.List(cboFabric.ListIndex), 2))
       nStretch = round(FN_JOBSTEX_Stretch(nAnkleCir, nHeelCir, nMMHg, iZipper, iFabric))
      
      'Calculate and Display the other stretch dependent values
       PR_DisplayFiguredAnkle "Left", nStretch, nAnkleCir, nHeelCir, iFabric
       

       g_iLtLastStretch = nStretch
      'g_iLtRed(g_iLtAnkle) this is set in PR_DisplayFiguredAnkle above
       g_iLtStretch(g_iLtAnkle) = nStretch
       g_iLtMM(g_iLtAnkle) = nMMHg


      'Check if template is set
       If cboLeftTemplate.ListIndex = -1 Then cboLeftTemplate.ListIndex = iZipper
       
    Else
      'POWERNET - Fabric
       iGrams = round(nAnkleCir * nMMHg)
       iModulus = Val(Mid$(cboFabric.List(cboFabric.ListIndex), 5, 3))
       iReduction = FN_POWERNET_Reduction(iGrams, iModulus)
      'Exit if error
       If iReduction < 0 Then
            PR_DisplayErrorMessage iReduction, "Error - Left Ankle Figuring"
            Exit Sub
       End If

      'Minimum Reduction for all diagnosis
       If iReduction < 14 Then
            iReduction = 14
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1001, "Left Ankle Figuring"
       End If
       
      'Maximum Reduction for Burns
       If iReduction > 26 And txtDiagnosis.Text = "Burns" Then
            iReduction = 26
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1002, "Left Ankle Figuring"
       End If
      
      'Maximum reduction for all other diagnosis
       If iReduction > 32 Then
            iReduction = 32
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1003, "Left Ankle Figuring"
       End If

      'Display Calculations
       txtLeftMM(g_iLtAnkle).Text = Str$(nMMHg)
       grdLeftDisplay.Row = g_iLtAnkle - 6
       grdLeftDisplay.Col = 0
       grdLeftDisplay = Str$(iGrams)
       grdLeftDisplay.Col = 1
       grdLeftDisplay = Str$(iReduction)

      'Set Template
       Select Case iReduction
        Case 0 To 14, 14 To 18
            cboLeftTemplate.ListIndex = 0   '30 mmhg
        Case 19 To 21
            cboLeftTemplate.ListIndex = 1   '35 mmhg
        Case 21 To 23
            cboLeftTemplate.ListIndex = 2   '40 mmhg
        Case 24 To 32, Is > 32
            cboLeftTemplate.ListIndex = 3   '50 mmhg
       End Select

      'Max Stretch (Only if no zipper given)
       If chkLeftZipper.Value = 0 Then
            nHeelCir = nHeelCir / 4
            nAnkleCir = (nAnkleCir * ((100 - iReduction) / 100)) / 2
            If nAnkleCir < nHeelCir Then PR_DisplayErrorMessage 1004, "Left Ankle Figuring"
       End If
        
       g_iLtStretch(g_iLtAnkle) = iGrams 'Bit of a misnomer but what the Hell!
       g_iLtRed(g_iLtAnkle) = iReduction
       g_iLtMM(g_iLtAnkle) = nMMHg
       
    End If
    g_iLtLastMM = Val(txtLeftMM(g_iLtAnkle).Text)

   'Calculate the pressure at all leg tapes
   '
    If g_JOBSTEX_FL = True Then
       'iZipper = 0     'Don't use zipper

       'Establish Last tape position
       'This allows for the addition of leg tapes to extend up the leg
       'It's not very efficient but what the hell!!!!
        nLastTape = 0
        For ii = g_iLtAnkle + 1 To 29
            If txtLeft(ii).Text = "" Then Exit For
        Next ii
        nLastTape = ii - 1
       
       'Establish pressure at last tape
       'As the last tape is at a 20 reduction we know that the stretch must
       'be 25%
        nLastTapeCir = fnDisplayToInches(Val(txtLeft(nLastTape).Text))
        nLastTapeMMHg = round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
        
        txtLeftMM(nLastTape) = nLastTapeMMHg
        g_iLtStretch(nLastTape) = 25
        g_iLtRed(nLastTape) = 20
        g_iLtMM(nLastTape) = nLastTapeMMHg
        
       'First pass from Ankle to LastTape
        nStretch = g_iLtLastStretch
        iPrevMMHg = g_iLtLastMM
        For ii = g_iLtAnkle + 1 To nLastTape - 1
            nTapeCir = fnDisplayToInches(Val(txtLeft(ii).Text))
            nMMHg = 1000 'Force While
            While nMMHg > iPrevMMHg And nStretch > 1
                nMMHg = round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
                If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
            Wend
            txtLeftMM(ii) = nMMHg
            g_iLtMM(ii) = nMMHg
            iPrevMMHg = nMMHg
            
           'Revise stretch to that based on the Pressure
            nStretch = round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))

            g_iLtStretch(ii) = nStretch
            g_iLtRed(ii) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
        Next ii

       'Second pass from LastTape to Ankle
       'Based on last tape Pressure recalculate the tapes until the pressures
       'become the same
        iPrevMMHg = nLastTapeMMHg

        For ii = nLastTape - 1 To g_iLtAnkle + 1 Step -1
            nTapeCir = fnDisplayToInches(Val(txtLeft(ii).Text))

            nMMHg = Val(txtLeftMM(ii).Text)
            nStretch = g_iLtStretch(ii)
            If nMMHg >= iPrevMMHg Then Exit For
            While nMMHg < iPrevMMHg
                nMMHg = round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
                If nMMHg < iPrevMMHg Then nStretch = nStretch + 1
            Wend
            txtLeftMM(ii) = nMMHg
            g_iLtMM(ii) = nMMHg
            iPrevMMHg = nMMHg

           'Revise stretch to that based on the Pressure
            nStretch = round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))

            g_iLtStretch(ii) = nStretch
            g_iLtRed(ii) = round((1 - 1 / ((0.01 * nStretch) + 1)) * 100)
        Next ii

       'Display Values
       'STRETCH
        grdLeftDisplay.Col = 0
        For ii = g_iLtAnkle + 1 To nLastTape
            grdLeftDisplay.Row = ii - 6
            grdLeftDisplay = Str$(g_iLtStretch(ii))
        Next ii
        
       'REDUCTION
        grdLeftDisplay.Col = 1
        For ii = g_iLtAnkle + 1 To nLastTape
            grdLeftDisplay.Row = ii - 6
            grdLeftDisplay = Str$(g_iLtRed(ii))
        Next ii
        
    End If

End Sub

Private Sub PR_FigureLeftTape(index As Integer)
   'Figures the Pressure/Stretch at a LEFT LegTape
   'This is only available to JOBSTEX_FL
   'See PR_FigureLeftTape for details
    
    Dim nMMHg As Double, nStretch As Double
    Dim nTapeCir As Double
    Dim iZipper As Integer, iFabric As Integer

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly
    If Val(txtLeftMM(index).Text) <= 0 Then Exit Sub

   'Establish if pressure has changed from last time
   'If it hasn't then exit
   'This is used to cure a problem in that the stretch for a tape is based not on the
   'pressure but on the previous or subsequent stretch, depending if it was found
   'on the forward pass or the backward pass.
   'The pressure was then back calculated fron the stretch.
   'We had a problem in that the same pressue can give a different stretch
   'To stop this happening when tabbing through the pressures this check has
   'been introduced
   '
    If Val(txtLeftMM(index).Text) = g_iLtMM(index) Then Exit Sub
    
   'MMHg and Stretch for JOBSTEX fabric
   'Don't Calculate if no values for Ankle tape
    If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle)) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub

   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
    
    nTapeCir = fnDisplayToInches(Val(txtLeft(index).Text))
    iZipper = chkLeftZipper.Value
    iFabric = Val(Left$(cboFabric.List(cboFabric.ListIndex), 2))
    nMMHg = Val(txtLeftMM(index).Text)
    g_iLtMM(index) = nMMHg

   'Calculate stretch
    nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, iZipper, iFabric)
    
   'Store values
    g_iLtStretch(index) = round(nStretch)
    g_iLtRed(index) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)

   'Display values
   'Note, the display grid displays from the ankle only
   '
    grdLeftDisplay.Row = index - 6 'Display from ankle
    grdLeftDisplay.Col = 0
    grdLeftDisplay = Str$(g_iLtStretch(index))
    grdLeftDisplay.Col = 1
    grdLeftDisplay = Str$(g_iLtRed(index))

End Sub

Private Sub PR_FigureRightAnkle()
   'Figures the Pressure/Stretch at the Right ankle
   'Input from form:-
   '   Right(5-7) where     Right(5) = Heel Tape
   '                       Right(6) = Ankle Tape (Heel < 9")
   '                       Right(7) = Ankle Tape (Heel > 9")
   '   cboFabric
   '   chkRightZipper
   '   txtRightMM(g_iRtAnkle)
   '   txtRightStretch(g_iRtAnkle)   JOBSTEX fabric only
   '
   'Input from globals
   '   g_JOBSTEX           Flag to indicate use of JOBSTEX fabric
   '   g_JOBSTEX_FL        Flag to indicate that JOBSTEX fabric in use
   '                       and that the stretch will be calculated at all
   '                       leg tape
   '   g_POWERNET          Flag to indicate use of Powernet fabric
   '
   '   g_iRtAnkle          Indicates the Right ankle tape
   '   g_nRtLastHeel
   '   g_nRtLastAnkle
   '   g_iRtLastMM
   '   g_iRtLastStretch    JOBSTEX fabric only
   '   g_iRtLastZipper
   '   g_iLastFabric
   '
   'JOBSTEX fabric Input and Output
   '   txtRightMM(g_iRtAnkle)        User Input, if txtRightStretch has been modified.
   '   txtRightStretch(g_iRtAnkle)   User Input, if txtRightMM has been modified.
   '   labRightRed(g_iRtAnkle)       Displays Reduction
   '   labRightMaxStr(2)             Displays MaxStretch
   '
   'Note:-
   '   The user inputs a required pressure in MMHg the stretch is then calculated
   '   and the stretch dependant values displayed.
   '   ARternately the stretch can be entered and the pressure in MMHg calculated.
   '   Given input of both (eg at link close) then MMHg takes precidence
   '
   'POWERNET fabric Input and Output
   '   txtRightMM(g_iRtAnkle)            User Input
   '   txtRightStretch(g_iRtAnkle)       not used (disabled)
   '   labRightFiguredGrams(g_iRtAnkle)  Displays Grams
   '   labRightRed(g_iRtAnkle)           Displays Reduction
   'Note:-
   '   The user inputs a required pressure in MMHg the reduction is then calculated
   '   and the reduction dependant values displayed.
   '
   '
    Dim nMMHg As Double, nStretch As Double
    Dim nAnkleCir As Double, nHeelCir As Double
    Dim nTapeCir As Double
    Dim iZipper As Integer, iFabric As Integer, iPrevMMHg As Integer
    Dim iModulus As Integer, iGrams As Integer, iReduction As Integer
    Dim ii As Integer
    Dim nLastTape As Integer, nLastTapeMMHg As Integer, nLastTapeCir As Double

   'Clean Messaage box
    labMessage.Caption = ""

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly

   'Leg tapes at heel and ankles and Fabric
    If Val(txtRight(5).Text) = 0 Or Val(txtRight(g_iRtAnkle).Text) = 0 Or cboFabric.ListIndex = -1 Then
        Exit Sub
    End If
   
   'MMHg for POWERNET fabric
    If g_POWERNET = True And Val(txtRightMM(g_iRtAnkle)) = 0 Then Exit Sub
    
   'MMHg and Stretch for JOBSTEX fabric
    If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtRightMM(g_iRtAnkle)) = 0 Then Exit Sub

   'Fabric Selection
   'NOTE:-
   '    The POWERNET fabric is based on the file g_sPathJOBST + "\WHFABRIC.DAT"
   '    For clarity this file can contain blank lines.  Therefor we must check
   '    that the fabric selected is not blank.
    If cboFabric.List(cboFabric.ListIndex) = "" Then Exit Sub

   'Establish if any of the inputs have changed compared to the last stored
   'values.
   'If they have changed then recalculate and redisplay
   'NB  Use of GOTO, (structured programmers get stuffed!)

   'Heel Tape
    If fnDisplayToInches(Val(txtRight(5))) <> g_nRtLastHeel Then GoTo RT_CALC
   
   'Ankle Tape
    If fnDisplayToInches(Val(txtRight(g_iRtAnkle))) <> g_nRtLastAnkle Then GoTo RT_CALC
   
   'Pressure MMHg
    If Val(txtRightMM(g_iRtAnkle).Text) <> g_iRtLastMM Then GoTo RT_CALC

   'Zippers
    If chkRightZipper.Value <> g_iRtLastZipper Then GoTo RT_CALC
    
   'Fabric
    If cboFabric.ListIndex <> g_iRtLastFabric Then GoTo RT_CALC

   'If nothing has changed then Exit this sub
    Exit Sub

   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
RT_CALC:
   'Calculate stretch based on given pressure (MMHg).
   'NB  All dimensions are converted to in inches except that last dimesions
   '    are stored using their display values.
    
    nHeelCir = fnDisplayToInches(Val(txtRight(5).Text))
    g_nRtLastHeel = nHeelCir
        
    nAnkleCir = fnDisplayToInches(Val(txtRight(g_iRtAnkle).Text))
    g_nRtLastAnkle = nAnkleCir

    iZipper = chkRightZipper.Value
    g_iRtLastZipper = iZipper

    g_iRtLastFabric = cboFabric.ListIndex
    nLastTape = -1
    
    nMMHg = Val(txtRightMM(g_iRtAnkle).Text)
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
      'JOBSTEX - Fabric
       iFabric = Val(Left$(cboFabric.List(cboFabric.ListIndex), 2))
       nStretch = round(FN_JOBSTEX_Stretch(nAnkleCir, nHeelCir, nMMHg, iZipper, iFabric))
      
      'Calculate and Display the other stretch dependent values
      'This procedure also issues warnings a well
       PR_DisplayFiguredAnkle "Right", nStretch, nAnkleCir, nHeelCir, iFabric
    
       g_iRtLastStretch = nStretch
      'g_iRtRed(g_iRtAnkle) this is set in PR_DisplayFiguredAnkle above
       g_iRtStretch(g_iRtAnkle) = nStretch
       g_iRtMM(g_iRtAnkle) = nMMHg


      'Check if template is set
       If cboRightTemplate.ListIndex = -1 Then cboRightTemplate.ListIndex = iZipper
       
    Else
      'POWERNET - Fabric
       iGrams = round(nAnkleCir * nMMHg)
       iModulus = Val(Mid$(cboFabric.List(cboFabric.ListIndex), 5, 3))
       iReduction = FN_POWERNET_Reduction(iGrams, iModulus)
      'Exit if error
       If iReduction < 0 Then
            PR_DisplayErrorMessage iReduction, "Error - Right Ankle Figuring"
            Exit Sub
       End If

      'Minimum Reduction for all diagnosis
       If iReduction < 14 Then
            iReduction = 14
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1001, "Right Ankle Figuring"
       End If
       
      'Maximum Reduction for Burns
       If iReduction > 26 And txtDiagnosis.Text = "Burns" Then
            iReduction = 26
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1002, "Right Ankle Figuring"
       End If
      
      'Maximum reduction for all other diagnosis
       If iReduction > 32 Then
            iReduction = 32
            nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
            iGrams = round(nAnkleCir * nMMHg)
            PR_DisplayErrorMessage 1003, "Right Ankle Figuring"
       End If
      
      'Display Calculations
       txtRightMM(g_iRtAnkle).Text = Str$(nMMHg)
       grdRightDisplay.Row = g_iRtAnkle - 6
       grdRightDisplay.Col = 0
       grdRightDisplay = Str$(iGrams)
       grdRightDisplay.Col = 1
       grdRightDisplay = Str$(iReduction)

      'Set Template
       Select Case iReduction
        Case 0 To 14, 14 To 18
            cboRightTemplate.ListIndex = 0   '30 mmhg
        Case 19 To 21
            cboRightTemplate.ListIndex = 1   '35 mmhg
        Case 21 To 23
            cboRightTemplate.ListIndex = 2   '40 mmhg
        Case 24 To 32, Is > 32
            cboRightTemplate.ListIndex = 3   '50 mmhg
       End Select

      'Max Stretch (Only if no zipper given)
       If chkRightZipper.Value = 0 Then
            nHeelCir = nHeelCir / 4
            nAnkleCir = (nAnkleCir * ((100 - iReduction) / 100)) / 2
            If nAnkleCir < nHeelCir Then PR_DisplayErrorMessage 1004, "Right Ankle Figuring"
       End If
        
       g_iRtStretch(g_iRtAnkle) = iGrams 'Bit of a misnomer but what the Hell!
       g_iRtRed(g_iRtAnkle) = iReduction
       g_iRtMM(g_iRtAnkle) = nMMHg
       
    End If
    g_iRtLastMM = Val(txtRightMM(g_iRtAnkle).Text)

   'Calculate the pressure at all leg tapes
   '
    If g_JOBSTEX_FL = True Then
       'iZipper = 0     'Don't use zipper

       'Establish Last tape position
       'This allows for the addition of leg tapes to extend up the leg
       'It's not very efficient but what the hell!!!!
        nLastTape = 0
        For ii = g_iRtAnkle + 1 To 29
            If txtRight(ii).Text = "" Then Exit For
        Next ii
        nLastTape = ii - 1
       
       'Establish pressure at last tape
       'As the last tape is at a 20 reduction we know that the stretch must
       'be 25%
        nLastTapeCir = fnDisplayToInches(Val(txtRight(nLastTape).Text))
        nLastTapeMMHg = round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
        
        txtRightMM(nLastTape) = nLastTapeMMHg
        g_iRtStretch(nLastTape) = 25
        g_iRtRed(nLastTape) = 20
        g_iRtMM(nLastTape) = nLastTapeMMHg
        
       'First pass from Ankle to LastTape
        nStretch = g_iRtLastStretch
        iPrevMMHg = g_iRtLastMM
        For ii = g_iRtAnkle + 1 To nLastTape - 1
            nTapeCir = fnDisplayToInches(Val(txtRight(ii).Text))
            nMMHg = 1000 'Force While
            While nMMHg > iPrevMMHg And nStretch > 1
                nMMHg = round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
                If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
            Wend
            txtRightMM(ii) = nMMHg
            g_iRtMM(ii) = nMMHg
            iPrevMMHg = nMMHg
           
           'Revise stretch to that based on the Pressure
            nStretch = round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))

            g_iRtStretch(ii) = nStretch
            g_iRtRed(ii) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
        Next ii

       'Second pass from LastTape to Ankle
       'Based on last tape Pressure recalculate the tapes until the pressures
       'become the same
        iPrevMMHg = nLastTapeMMHg

        For ii = nLastTape - 1 To g_iRtAnkle + 1 Step -1
            nTapeCir = fnDisplayToInches(Val(txtRight(ii).Text))

            nMMHg = Val(txtRightMM(ii).Text)
            nStretch = g_iRtStretch(ii)
            If nMMHg >= iPrevMMHg Then Exit For
            While nMMHg < iPrevMMHg
                nMMHg = round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
                If nMMHg < iPrevMMHg Then nStretch = nStretch + 1
            Wend
            txtRightMM(ii) = nMMHg
            g_iRtMM(ii) = nMMHg
            iPrevMMHg = nMMHg

           'Revise stretch to that based on the Pressure
            nStretch = round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))

            g_iRtStretch(ii) = nStretch
            g_iRtRed(ii) = round((1 - 1 / ((0.01 * nStretch) + 1)) * 100)
        Next ii

       'Display Values
       'STRETCH
        grdRightDisplay.Col = 0
        For ii = g_iRtAnkle + 1 To nLastTape
            grdRightDisplay.Row = ii - 6
            grdRightDisplay = Str$(g_iRtStretch(ii))
        Next ii
        
       'REDUCTION
        grdRightDisplay.Col = 1
        For ii = g_iRtAnkle + 1 To nLastTape
            grdRightDisplay.Row = ii - 6
            grdRightDisplay = Str$(g_iRtRed(ii))
        Next ii
        
    End If

End Sub

Private Sub PR_FigureRightTape(index As Integer)
    Dim nMMHg As Double, nStretch As Double
    Dim nTapeCir As Double
    Dim iZipper As Integer, iFabric As Integer

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly
    If Val(txtRightMM(index).Text) <= 0 Then Exit Sub

    If Val(txtRightMM(index).Text) = g_iRtMM(index) Then Exit Sub
    
   'MMHg and Stretch for JOBSTEX fabric
   'Don't Calculate if no values for Ankle tape
    If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtRightMM(g_iRtAnkle)) = 0 Or g_iRtStretch(g_iRtAnkle) = 0 Then Exit Sub
    
   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
    
    nTapeCir = fnDisplayToInches(Val(txtRight(index).Text))
    iZipper = chkRightZipper.Value
    iFabric = Val(Left$(cboFabric.List(cboFabric.ListIndex), 2))
    nMMHg = Val(txtRightMM(index).Text)
    g_iRtMM(index) = nMMHg

   'Calculate stretch
    nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nRtLastHeel, nMMHg, iZipper, iFabric)
    
   'Store values
    g_iRtStretch(index) = round(nStretch)
    g_iRtRed(index) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)

   'Display values
   'Note, the display grid displays from the ankle only
   '
    grdRightDisplay.Row = index - 6 'Display from ankle
    grdRightDisplay.Col = 0
    grdRightDisplay = Str$(g_iRtStretch(index))
    grdRightDisplay.Col = 1
    grdRightDisplay = Str$(g_iRtRed(index))

End Sub

Private Sub PR_HeelChange(sLeg As String)
   'Procedure to modify the display if heel changed

    Dim iLtExistingAnkle  As Integer, iRtExistingAnkle As Integer
   'Clean Message box
    labMessage.Caption = ""

   'Exit if heel has not been changed
    If sLeg = "Left" Then
        If fnDisplayToInches(Val(txtLeft(5).Text)) = g_nLtLastHeel Then Exit Sub
    Else
        If fnDisplayToInches(Val(txtRight(5).Text)) = g_nRtLastHeel Then Exit Sub
    End If

   'Store existing ankle position and get new ones
    iLtExistingAnkle = g_iLtAnkle
    iRtExistingAnkle = g_iRtAnkle

    PR_EstablishAnkles

   'If revised the same as existing then exit
    If iLtExistingAnkle = g_iLtAnkle And iRtExistingAnkle = g_iRtAnkle Then Exit Sub

   'Modify display
    PR_ResetFormWHFIGURE
    If g_POWERNET = True Then PR_EnablePOWERNET
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX
    If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX

End Sub

Private Sub PR_LoadFabricFromFile(sFileName As String)
   'Procedure to load the POWERNET conversion chart from file
   'N.B. File opening Errors etc. are not handled (so tough titty!)
    
    Dim fNum As Integer, ii As Integer
    fNum = FreeFile
    Open sFileName For Input As fNum
    ii = 0
    Do Until EOF(fNum)
        Input #fNum, POWERNET.Modulus(ii), POWERNET.Conversion(ii)
        ii = ii + 1
    Loop
    Close fNum

End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable

    Print #fNum, sLine

End Sub

Private Sub PR_ResetFormWHFIGURE()
   'Reset display of form to a clean state
    Dim ii As Integer

   'Reset Max Stretch
    labLeftMaxStr.Caption = ""
    labLeftMaxStr.Visible = False
    labLeftDisp(3).Visible = False
    labRightMaxStr.Caption = ""
    labRightMaxStr.Visible = False
    labRightDisp(3).Visible = False

   'Clean Display grid
    For ii = 0 To 23
        
        grdLeftDisplay.Row = ii
        grdLeftDisplay.Col = 0
        grdLeftDisplay = ""
        grdLeftDisplay.Col = 1
        grdLeftDisplay = ""

        grdRightDisplay.Row = ii
        grdRightDisplay.Col = 0
        grdRightDisplay = ""
        grdRightDisplay.Col = 1
        grdRightDisplay = ""

    Next ii

   'Switch off MM text boxes
    For ii = 6 To 29
        txtLeftMM(ii).Text = ""
        txtLeftMM(ii).Visible = False
        txtRightMM(ii).Text = ""
        txtRightMM(ii).Visible = False
    Next ii

End Sub

Private Sub Timer1_Timer()
    'If this event happens then the DDE link to
    'DRAFIX has not been terminated sucessfully.
    'IE No link close.
    'We assume a failure of some sort and we End
     
     End

End Sub

Private Sub txtLeft_GotFocus(index As Integer)
    Select_Text_In_Box txtLeft(index)
End Sub

Private Sub txtLeft_LostFocus(index As Integer)
    Validate_And_Display_Text_In_Box txtLeft(index), grdLeftInches, index
    If index = g_iLtAnkle Then PR_FigureLeftAnkle
    If index = 5 Then PR_HeelChange "Left"
End Sub

Private Sub txtLeftMM_LostFocus(index As Integer)
    If index = g_iLtAnkle Then
        PR_FigureLeftAnkle
    Else
        PR_FigureLeftTape index
    End If
End Sub

Private Sub txtRight_GotFocus(index As Integer)
    Select_Text_In_Box txtRight(index)
End Sub

Private Sub txtRight_LostFocus(index As Integer)
    Validate_And_Display_Text_In_Box txtRight(index), grdRightInches, index
    If index = g_iRtAnkle Then PR_FigureRightAnkle
    If index = 5 Then PR_HeelChange "Right"
End Sub

Private Sub txtRightMM_LostFocus(index As Integer)
    If index = g_iRtAnkle Then
        PR_FigureRightAnkle
    Else
        PR_FigureRightTape index
    End If

End Sub

Private Sub Update_DDE_Text_Boxes()
   'Called from OK_Click
   'Update the text boxes used for DDE transfer
    Dim ii As Integer, sString As String, sJustifiedString As String
    Dim sFabricClass As String
    Dim iRed As Integer, iStr  As Integer, iMM As Integer
    Dim sRed As String, sStr As String, sMM As String
    Dim nValue As Single

    If g_POWERNET = True Then sFabricClass = "0"
    If g_JOBSTEX = True Then sFabricClass = "1"
    If g_JOBSTEX_FL = True Then sFabricClass = "2"

    txtFabric = cboFabric.List(cboFabric.ListIndex)

   'LEFT Leg
   'Assume that data has been validated earlier
    
   'Set initial values
    txtLeftLengths = ""
    txtLeftRed = ""
    txtLeftStr = ""
    txtLeftMMs = ""
    txtLeftFirstTape = -1
    txtLeftLastTape = -1
    txtLeftTemplate = cboLeftTemplate.List(cboLeftTemplate.ListIndex)

    For ii = 0 To 29
        nValue = Val(txtLeft(ii).Text)
        If nValue <> 0 Then
            nValue = nValue * 10 'Shift decimal place to right
            sJustifiedString = String(4, " ")
            RSet sJustifiedString = Trim$(Str$(nValue))
        Else
            sJustifiedString = String(4, " ")
        End If

       'Tape values
        txtLeftLengths.Text = txtLeftLengths.Text & sJustifiedString
       
       'Set first and last tape (assumes no holes in data)
        If txtLeftFirstTape < 0 And nValue > 0 Then txtLeftFirstTape = ii + 1
        If txtLeftLastTape < 0 And txtLeftFirstTape > 0 And nValue = 0 Then txtLeftLastTape = ii
    Next ii

   'Where JOBSTEX_FL has been choosen then we update the MMs, Reduction and Stretch
   'These will then be picked up by the Leg Drawing modules
    If g_JOBSTEX_FL = True And g_iLtAnkle <> 0 Then
        For ii = 0 To 29
            iMM = g_iLtMM(ii)
            iRed = g_iLtRed(ii)
            iStr = g_iLtStretch(ii)
            If iMM <> 0 Then
                sMM = String(3, " ")
                RSet sMM = Trim$(Str$(iMM))
                sRed = String(3, " ")
                RSet sRed = Trim$(Str$(iRed))
                sStr = String(3, " ")
                RSet sStr = Trim$(Str$(iStr))
            Else
                sMM = String(3, " ")
                sRed = String(3, " ")
                sStr = String(3, " ")
            End If
            txtLeftMMs.Text = txtLeftMMs.Text & sMM
            txtLeftRed.Text = txtLeftRed.Text & sRed
            txtLeftStr.Text = txtLeftStr.Text & sStr
        Next ii
    End If

   'Transfer to the "legbox" the following data in blank delimited format
   '
   'txtLeftAnkleTape = "AnkleTape#
   '                    Pressure
   '                    [Grams|Stretch]
   '                    Reduction
   '                    AnkleLength
   '                    HeelLength
   '                    Zipper Status
   '                    FabricClass
   '
   'This means that the CutOut drawing Macro (WH_CUT.D) need not load all
   'of the routines to read leg data, as it is only interested in the
   'heel length
   '
   'g_iLtAnkle=0 => footless, so don't worry about the ankle
   'Note:-
   '    The DRAFIX macros index tape lengths from 1 not 0 as VB does.
   '    Therefor g_iLtAnkle is incremented by 1
   '
    sString = ""
    If g_iLtAnkle <> 0 Then
        sString = sString & Str$(g_iLtAnkle + 1) & " "
        sString = sString & Str$(g_iLtMM(g_iLtAnkle)) & " "
        sString = sString & Str$(g_iLtStretch(g_iLtAnkle)) & " "
        sString = sString & Str$(g_iLtRed(g_iLtAnkle)) & " "
        sString = sString & Str$(g_nLtLastAnkle) & " "
        sString = sString & Str$(g_nLtLastHeel) & " "
        sString = sString & Str$(g_iLtLastZipper) & " "
        txtLeftAnkle.Text = sString & sFabricClass
    Else
       'Set the txtAnkleTape to reflect the footless state
        txtLeftAnkle.Text = "-1 0 0 0 0 0 0 " & sFabricClass
    End If


   'RIGHT Leg

    txtRightLengths = ""
    txtRightRed = ""
    txtRightStr = ""
    txtRightMMs = ""
    txtRightFirstTape = -1
    txtRightLastTape = -1
    txtRightTemplate = cboRightTemplate.List(cboRightTemplate.ListIndex)

    For ii = 0 To 29
        nValue = Val(txtRight(ii).Text)
        If nValue <> 0 Then
            nValue = nValue * 10 'Shift decimal place to right
            sJustifiedString = String(4, " ")
            RSet sJustifiedString = Trim$(Str$(nValue))
        Else
            sJustifiedString = String(4, " ")
        End If

       'Tape values
        txtRightLengths.Text = txtRightLengths.Text & sJustifiedString
       
       'Set first and last tape (assumes no holes in data)
        If txtRightFirstTape < 0 And nValue > 0 Then txtRightFirstTape = ii + 1
        If txtRightLastTape < 0 And txtRightFirstTape > 0 And nValue = 0 Then txtRightLastTape = ii
    Next ii

   'Where JOBSTEX_FL has been choosen then we update the MMs, Reduction and Stretch
   'These will then be picked up by the Leg Drawing modules
    If g_JOBSTEX_FL = True And g_iRtAnkle <> 0 Then
        For ii = 0 To 29
            iMM = g_iRtMM(ii)
            iRed = g_iRtRed(ii)
            iStr = g_iRtStretch(ii)
            If iMM <> 0 Then
                sMM = String(3, " ")
                RSet sMM = Trim$(Str$(iMM))
                sRed = String(3, " ")
                RSet sRed = Trim$(Str$(iRed))
                sStr = String(3, " ")
                RSet sStr = Trim$(Str$(iStr))
            Else
                sMM = String(3, " ")
                sRed = String(3, " ")
                sStr = String(3, " ")
            End If
            txtRightMMs.Text = txtRightMMs.Text & sMM
            txtRightRed.Text = txtRightRed.Text & sRed
            txtRightStr.Text = txtRightStr.Text & sStr
        Next ii
    End If

   'Transfer to the "legbox" the following data in blank delimited format
    
    sString = ""
    If g_iRtAnkle <> 0 Then
        sString = sString & Str$(g_iRtAnkle + 1) & " "
        sString = sString & Str$(g_iRtMM(g_iRtAnkle)) & " "
        sString = sString & Str$(g_iRtStretch(g_iRtAnkle)) & " "
        sString = sString & Str$(g_iRtRed(g_iRtAnkle)) & " "
        sString = sString & Str$(g_nRtLastAnkle) & " "
        sString = sString & Str$(g_nRtLastHeel) & " "
        sString = sString & Str$(g_iRtLastZipper) & " "
        txtRightAnkle.Text = sString & sFabricClass
    Else
       'Set the txtAnkleTape to reflect the footless state
        txtRightAnkle.Text = "-1 0 0 0 0 0 0 " & sFabricClass
    End If
  
End Sub

Private Function Validate_Data() As Integer
   'Called from OK_Click
    
    Dim sError As String, NL As String, sTextList As String
    Dim sLeftError As String, sRightError As String
    Dim ii As Integer, nFirstTape As Integer, nLastTape As Integer
    Dim iError As Integer
    Dim nHeelLength As Double, nFootLength As Double, nTapeX As Double
    Dim nValue As Double

    NL = Chr$(10) 'new line
    sTextList = "  7  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"


   'LEFT LEG Checks
   'Check Tape length data text boxes for holes (ie missing values)
   'Establish First and last tape
    
    sError = ""

    For ii = 0 To 29 Step 1
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    nFirstTape = ii

    For ii = 29 To 0 Step -1
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    nLastTape = ii

    For ii = nFirstTape To nLastTape
        If Val(txtLeft(ii).Text) = 0 Then
            sError = sError + "Missing Tape length - " & LTrim$(Mid$(sTextList, (ii * 3) + 1, 3)) + NL
        End If
    Next ii
   
   'Check if -3 tape exists for Large Heel
    If g_iLtAnkle = 7 And Val(txtLeft(3).Text) = 0 Then
        sError = sError + "As the Heel is 9 inches and over." + NL
        sError = sError + "and there is no -3 tape the foot will not draw properly " + NL
    End If

   'Check if -1 1/2 tape exists for Small Heel
    If g_iLtAnkle = 6 And Val(txtLeft(4).Text) = 0 Then
        sError = sError + "As the Heel is smaller than 9 inches." + NL
        sError = sError + "and there is no -1 1/2 tape the foot will not draw properly " + NL
    End If
   
   'Check that figuring has been done at the ankle
    If g_iLtAnkle <> 0 And (g_iLtStretch(g_iLtAnkle) = 0 Or g_iLtRed(g_iLtAnkle) = 0 Or g_iLtMM(g_iLtAnkle) = 0) Then
        sError = sError + "The Ankle has not been figured!." + NL
        sError = sError + "Figure the Ankle or use CANCEL to exit" + NL
    End If

    sLeftError = sError

   'RIGHT LEG Checks
   'Check Tape length data text boxes for holes (ie missing values)
   'Establish First and last tape
    
    sError = ""

    For ii = 0 To 29 Step 1
        If Val(txtRight(ii).Text) > 0 Then Exit For
    Next ii
    nFirstTape = ii

    For ii = 29 To 0 Step -1
        If Val(txtRight(ii).Text) > 0 Then Exit For
    Next ii
    nLastTape = ii

    For ii = nFirstTape To nLastTape
        If Val(txtRight(ii).Text) = 0 Then
            sError = sError + "Missing Tape length - " & LTrim$(Mid$(sTextList, (ii * 3) + 1, 3)) + NL
        End If
    Next ii
   
   'Check if -3 tape exists for Large Heel
    If g_iRtAnkle = 7 And Val(txtRight(3).Text) = 0 Then
        sError = sError + "As the Heel is 9 inches and over." + NL
        sError = sError + "and there is no -3 tape the foot will not draw properly " + NL
    End If

   'Check if -1 1/2 tape exists for Small Heel
    If g_iRtAnkle = 6 And Val(txtRight(4).Text) = 0 Then
        sError = sError + "As the Heel is smaller than 9 inches." + NL
        sError = sError + "and there is no -1 1/2 tape the foot will not draw properly " + NL
    End If
   
   'Check that figuring has been done at the ankle
    If g_iRtAnkle <> 0 And (g_iRtStretch(g_iRtAnkle) = 0 Or g_iRtRed(g_iRtAnkle) = 0 Or g_iRtMM(g_iRtAnkle) = 0) Then
        sError = sError + "The Ankle has not been figured!." + NL
        sError = sError + "Figure the Ankle or use CANCEL to exit" + NL
    End If
    
    sRightError = sError

    If sRightError <> "" Then sError = "----------------------------- RIGHT LEG -----------------------------" + NL + sRightError
    If sLeftError <> "" Then sError = sError + "----------------------------- LEFT  LEG -----------------------------" + NL + sLeftError
   
   'General check
    If cboFabric.List(cboFabric.ListIndex) = "" Then
        sError = sError + "A Fabric has not been chosen!." + NL
    End If

   'Display Error message (if required) and return
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        Validate_Data = False
        Exit Function
    Else
        Validate_Data = True
    End If



   'Yes / No type errors
   'In this case we warn the user that there is a problem!
   'They can continue or they can return to the dialog to make changes

   'Initialize error variables
    sError = ""
    iError = False

   'Display Error message (if required) and return
   'These are non-fatal errors
    If iError = True Then
        iError = MsgBox(sError, 36, "Warning, Problems with data")
        If iError = 6 Then  'IDYES
            Validate_Data = True
        Else
            Validate_Data = False
        End If
    Else
        Validate_Data = True
    End If

End Function


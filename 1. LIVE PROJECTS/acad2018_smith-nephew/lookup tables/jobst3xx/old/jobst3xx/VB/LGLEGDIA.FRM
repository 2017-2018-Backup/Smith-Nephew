VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "comct232.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form lglegdia 
   Caption         =   "LEG - Details"
   ClientHeight    =   8532
   ClientLeft      =   1716
   ClientTop       =   468
   ClientWidth     =   9012
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "lglegdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8532
   ScaleWidth      =   9012
   Begin VB.Timer Timer1 
      Left            =   8940
      Top             =   8100
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   62
      Top             =   7740
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   9600
      TabIndex        =   66
      Top             =   8130
      Width           =   675
   End
   Begin VB.TextBox txtLegTitleFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   186
      Top             =   4140
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtThighBandBK 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   184
      Top             =   7440
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtKneeBand 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   181
      Top             =   7140
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtThighBandAK 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   179
      Top             =   6840
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtKneeLength 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   177
      Top             =   6540
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtThighLength 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   175
      Top             =   6240
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtAnklet 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   173
      Top             =   5940
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton Draw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   5820
      TabIndex        =   172
      Top             =   7860
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Leg Style"
      Height          =   1815
      Left            =   5220
      TabIndex        =   77
      Top             =   870
      Width           =   3675
      Begin ComCtl2.UpDown spnFirstTape 
         Height          =   375
         Left            =   1440
         TabIndex        =   188
         Top             =   1320
         Width           =   240
         _ExtentX        =   339
         _ExtentY        =   656
         _Version        =   327681
      End
      Begin ComCtl2.UpDown spnLastTape 
         Height          =   375
         Left            =   3360
         TabIndex        =   187
         Top             =   1320
         Width           =   240
         _ExtentX        =   339
         _ExtentY        =   656
         _Version        =   327681
      End
      Begin VB.OptionButton optType 
         Caption         =   "Thigh B/K"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   183
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox txtFirstTape 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   900
         TabIndex        =   170
         Top             =   1350
         Width           =   495
      End
      Begin VB.TextBox txtLastTape 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2820
         TabIndex        =   169
         Top             =   1350
         Width           =   495
      End
      Begin VB.OptionButton optType 
         Caption         =   "Thigh A/K"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   86
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Thigh Length"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Top             =   960
         Width           =   1515
      End
      Begin VB.OptionButton optType 
         Caption         =   "Knee Band"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   93
         Top             =   330
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Knee Length"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   94
         Top             =   630
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "Anklet"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   95
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label labLastTape 
         Caption         =   "Last tape"
         Height          =   225
         Left            =   1980
         TabIndex        =   82
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label labFirstTape 
         Caption         =   "First tape"
         Height          =   225
         Left            =   60
         TabIndex        =   83
         Top             =   1380
         Width           =   855
      End
   End
   Begin VB.TextBox txtDiagnosis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9660
      TabIndex        =   168
      Top             =   1170
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtPatientName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10800
      TabIndex        =   166
      Top             =   570
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUnits 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9660
      TabIndex        =   164
      Top             =   870
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtFileNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9660
      TabIndex        =   162
      Top             =   570
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10800
      TabIndex        =   160
      Top             =   1170
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10800
      TabIndex        =   158
      Top             =   870
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   156
      Top             =   5640
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtToeStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   154
      Top             =   5340
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame4 
      Caption         =   "Styles / Other Measurements"
      Height          =   1290
      Left            =   5220
      TabIndex        =   54
      Top             =   3945
      Width           =   3675
      Begin VB.ComboBox cboToeStyle 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   420
         Width           =   1995
      End
      Begin VB.TextBox txtFootLength 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   61
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "FootLength:"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Toe Style:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label labFootLengthInInches 
         Height          =   255
         Left            =   1800
         TabIndex        =   68
         Top             =   870
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Leg"
      Height          =   615
      Left            =   5220
      TabIndex        =   69
      Top             =   120
      Width           =   3675
      Begin VB.OptionButton optLeftLeg 
         Caption         =   "Left"
         Height          =   195
         Left            =   660
         TabIndex        =   70
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optRightLeg 
         Caption         =   "Right"
         Height          =   195
         Left            =   2160
         TabIndex        =   71
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.TextBox txtUidLegTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   72
      Top             =   2640
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   137
      Top             =   5040
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftStr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   138
      Top             =   4740
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   139
      Top             =   4440
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Left 
      Caption         =   "Leg Details"
      Height          =   8415
      Left            =   120
      TabIndex        =   136
      Top             =   60
      Width           =   4995
      Begin MSFlexGridLib.MSFlexGrid grdLeftDisplay 
         Height          =   6450
         Left            =   3780
         TabIndex        =   193
         Top             =   1880
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   11367
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
      Begin MSFlexGridLib.MSFlexGrid grdLeftInches 
         Height          =   8085
         Left            =   2265
         TabIndex        =   192
         Top             =   240
         Width           =   800
         _ExtentX        =   1397
         _ExtentY        =   14266
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
      Begin VB.CommandButton ShiftTapesUp 
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton ShiftTapesDown 
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton ExtendLegTapes 
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtFootPleat1 
         Height          =   315
         Left            =   240
         MaxLength       =   7
         TabIndex        =   145
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtFootPleat2 
         Height          =   315
         Left            =   240
         MaxLength       =   7
         TabIndex        =   147
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtTopLegPleat2 
         Height          =   315
         Left            =   240
         MaxLength       =   7
         TabIndex        =   149
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox txtTopLegPleat1 
         Height          =   315
         Left            =   240
         MaxLength       =   7
         TabIndex        =   151
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   29
         Left            =   3240
         TabIndex        =   56
         Top             =   8070
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   28
         Left            =   3240
         TabIndex        =   55
         Top             =   7800
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   27
         Left            =   3240
         TabIndex        =   53
         Top             =   7530
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   26
         Left            =   3240
         TabIndex        =   52
         Top             =   7260
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   25
         Left            =   3240
         TabIndex        =   51
         Top             =   6990
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   24
         Left            =   3240
         TabIndex        =   50
         Top             =   6720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   23
         Left            =   3240
         TabIndex        =   49
         Top             =   6450
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   22
         Left            =   3240
         TabIndex        =   48
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   21
         Left            =   3240
         TabIndex        =   47
         Top             =   5910
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   20
         Left            =   3240
         TabIndex        =   46
         Top             =   5640
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   19
         Left            =   3240
         TabIndex        =   45
         Top             =   5370
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   18
         Left            =   3240
         TabIndex        =   44
         Top             =   5100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   17
         Left            =   3240
         TabIndex        =   43
         Top             =   4830
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   16
         Left            =   3240
         TabIndex        =   42
         Top             =   4560
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   15
         Left            =   3240
         TabIndex        =   41
         Top             =   4290
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   14
         Left            =   3240
         TabIndex        =   40
         Top             =   4020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   39
         Top             =   3750
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   12
         Left            =   3240
         TabIndex        =   38
         Top             =   3480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   11
         Left            =   3240
         TabIndex        =   37
         Top             =   3210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   10
         Left            =   3240
         TabIndex        =   36
         Top             =   2940
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   9
         Left            =   3240
         TabIndex        =   35
         Top             =   2670
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   8
         Left            =   3240
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   7
         Left            =   3240
         TabIndex        =   33
         Top             =   2130
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   6
         Left            =   3240
         TabIndex        =   31
         Top             =   1860
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Frame Frame2 
         Caption         =   "Template"
         Height          =   915
         Left            =   3180
         TabIndex        =   130
         Top             =   150
         Width           =   1755
         Begin VB.ComboBox cboLeftTemplate 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   270
            Width           =   1395
         End
         Begin VB.CheckBox chkLeftZipper 
            Caption         =   "Ankle Zipper"
            Height          =   225
            Left            =   120
            TabIndex        =   32
            Top             =   630
            Width           =   1515
         End
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   0
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   2
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1290
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1830
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2370
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   9
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   10
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   10
         Top             =   2910
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   11
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   11
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   12
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   12
         Top             =   3450
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   13
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   14
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   14
         Top             =   3990
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   15
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   15
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   16
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   16
         Top             =   4530
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   17
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   17
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   18
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   18
         Top             =   5070
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   19
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   19
         Top             =   5340
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   20
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   20
         Top             =   5610
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   21
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   21
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   22
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   22
         Top             =   6150
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   23
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   23
         Top             =   6420
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   24
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   24
         Top             =   6690
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   25
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   25
         Top             =   6960
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   26
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   26
         Top             =   7230
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   27
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   27
         Top             =   7500
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   28
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   28
         Top             =   7770
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   29
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   29
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "Pleats:"
         Height          =   195
         Left            =   240
         TabIndex        =   128
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label32 
         Caption         =   "Pleats:"
         Height          =   195
         Left            =   240
         TabIndex        =   153
         Top             =   7140
         Width           =   675
      End
      Begin VB.Label Label31 
         Caption         =   "move"
         Height          =   195
         Left            =   360
         TabIndex        =   152
         Top             =   4020
         Width           =   555
      End
      Begin VB.Label Label30 
         Caption         =   "tapes"
         Height          =   195
         Left            =   360
         TabIndex        =   150
         Top             =   4260
         Width           =   555
      End
      Begin VB.Label Label28 
         Caption         =   "Leg Tapes"
         Height          =   195
         Left            =   120
         TabIndex        =   148
         Top             =   6000
         Width           =   915
      End
      Begin VB.Label Label21 
         Caption         =   "Extend"
         Height          =   195
         Left            =   240
         TabIndex        =   146
         Top             =   5760
         Width           =   675
      End
      Begin VB.Label labLeftMaxStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4320
         TabIndex        =   140
         Top             =   1215
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label labLeftDisp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Max Stretch:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   141
         Top             =   1230
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Red."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   142
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gms"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   143
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "MMHg"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3180
         TabIndex        =   144
         Top             =   1620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   97
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   98
         Top             =   510
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   99
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   100
         Top             =   1050
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   4
         Left            =   1080
         TabIndex        =   101
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   102
         Top             =   1590
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   103
         Top             =   1860
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   7
         Left            =   1080
         TabIndex        =   104
         Top             =   2130
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   8
         Left            =   1080
         TabIndex        =   105
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   9
         Left            =   1080
         TabIndex        =   106
         Top             =   2670
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   10
         Left            =   1080
         TabIndex        =   107
         Top             =   2940
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   195
         Index           =   11
         Left            =   1080
         TabIndex        =   108
         Top             =   3210
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "10½"
         Height          =   195
         Index           =   12
         Left            =   1080
         TabIndex        =   109
         Top             =   3480
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "12"
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   110
         Top             =   3750
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "13½"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   111
         Top             =   4020
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "15"
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   112
         Top             =   4290
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "16½"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   113
         Top             =   4560
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "18"
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   114
         Top             =   4830
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "19½"
         Height          =   195
         Index           =   18
         Left            =   1080
         TabIndex        =   115
         Top             =   5100
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "21"
         Height          =   195
         Index           =   19
         Left            =   1080
         TabIndex        =   116
         Top             =   5370
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "22½"
         Height          =   195
         Index           =   20
         Left            =   1080
         TabIndex        =   117
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   195
         Index           =   21
         Left            =   1080
         TabIndex        =   118
         Top             =   5910
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "25½"
         Height          =   195
         Index           =   22
         Left            =   1080
         TabIndex        =   119
         Top             =   6180
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "27"
         Height          =   195
         Index           =   23
         Left            =   1080
         TabIndex        =   120
         Top             =   6450
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "28½"
         Height          =   195
         Index           =   24
         Left            =   1080
         TabIndex        =   121
         Top             =   6720
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   195
         Index           =   25
         Left            =   1080
         TabIndex        =   122
         Top             =   6990
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "31½"
         Height          =   195
         Index           =   26
         Left            =   1080
         TabIndex        =   123
         Top             =   7260
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "33"
         Height          =   195
         Index           =   27
         Left            =   1080
         TabIndex        =   124
         Top             =   7530
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "34½"
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   125
         Top             =   7800
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "36"
         Height          =   195
         Index           =   29
         Left            =   1080
         TabIndex        =   126
         Top             =   8070
         Width           =   315
      End
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   132
      Top             =   2040
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftTemplate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   129
      Top             =   3840
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fabric"
      Height          =   1035
      Left            =   5220
      TabIndex        =   127
      Top             =   2790
      Width           =   3675
      Begin VB.OptionButton optFabric 
         Caption         =   "Jobstex"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   59
         Top             =   660
         Width           =   1035
      End
      Begin VB.OptionButton optFabric 
         Caption         =   "Powernet"
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   58
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   330
         Width           =   3195
      End
   End
   Begin VB.TextBox txtChosenStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   87
      Top             =   2340
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   81
      Top             =   3540
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   78
      Top             =   3240
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidLeftLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   74
      Top             =   2940
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10260
      TabIndex        =   73
      Top             =   1740
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7380
      TabIndex        =   79
      Top             =   7860
      Width           =   1215
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "txtWorkOorder"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8940
      TabIndex        =   63
      Top             =   7770
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   67
      Top             =   4170
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ThighBandBK"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8940
      TabIndex        =   185
      Top             =   7470
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "KneeBand"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8940
      TabIndex        =   182
      Top             =   7170
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ThighBandAK"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8940
      TabIndex        =   180
      Top             =   6870
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "KneeLength"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8940
      TabIndex        =   178
      Top             =   6570
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ThighLength"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8940
      TabIndex        =   176
      Top             =   6270
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Anklet"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8940
      TabIndex        =   174
      Top             =   5970
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label11 
      Caption         =   "Messages / Errors:"
      Height          =   225
      Left            =   5400
      TabIndex        =   171
      Top             =   5280
      Width           =   1635
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Diag:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   96
      Top             =   1230
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10140
      TabIndex        =   167
      Top             =   630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Units:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9120
      TabIndex        =   165
      Top             =   930
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "File No:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   163
      Top             =   630
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sex:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   10320
      TabIndex        =   161
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Age:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   10320
      TabIndex        =   159
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "leg"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8940
      TabIndex        =   157
      Top             =   5670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ToeStyle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   155
      Top             =   5400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LegTitleUID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   85
      Top             =   2670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label labMessage 
      BorderStyle     =   1  'Fixed Single
      Height          =   2100
      Left            =   5220
      TabIndex        =   90
      Top             =   5505
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   91
      Top             =   5070
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   131
      Top             =   4770
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   135
      Top             =   4470
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Templt"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   134
      Top             =   3870
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   133
      Top             =   2070
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DDE Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9660
      TabIndex        =   89
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Chosen Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   88
      Top             =   2370
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   84
      Top             =   3570
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   80
      Top             =   3270
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   76
      Top             =   2970
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8940
      TabIndex        =   75
      Top             =   1770
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "lglegdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module :   LGLEGDIA.MAK
'Purpose:   Input and Figure leg
'
'Version:   3.01
'Date:      1995
'Author:    Gary George
'
'Used in:
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'Oct/Nov 95 GG      Modifications w.r.t Triton / Imageable
'
'19.Dec.95  GG      Bug fix
'                   Close button always assumed a change
'                   had been made.
'                   FN_ConcatData() added
'
'Jan 99     GG      Ported to VB5
'-------------------------------------------------------
'
'Notes:-
'    Much of the code and form has been hacked from
'    WHFIGURE.FRM and WHLEGDIA.FRM
'    As both of these are proven, there has been little
'    to no changes made except to disable JOBSTEX_FL
'
'    Reference to the left leg should be taken to mean
'    the current leg
'

   'MsgBox constants
'    Const IDYES = 6
'    Const IDNO = 7
'    Const IDCANCEL = 2

   'Other constants
    Const NEWLINE = 13

    Option Explicit

Private Sub Cancel_Click()
    Dim Response%
    Dim sTask$

   'Check if data has been modified
    If g_sChangeChecker <> FN_ConcatData() Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "LEG Details Dialogue")
        Select Case Response
            Case IDYES
                Update_DDE_Text_Boxes
                fnum = FN_SaveOpen("c:\jobst\draw.d", txtPatientName, txtFileNo, txtLeg)
                PR_SaveLeg
                Close fnum
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate fnGetDrafixWindowTitleText()
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "LEG Details Dialogue"
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

Private Sub Draw_Click()
    Dim sCAD_App$
   'Check that data is all present and insert into drafix
    If Validate_Data() Then
        Update_DDE_Text_Boxes
        PR_DrawAndSaveLeg
        sCAD_App = fnGetDrafixWindowTitleText()
        If sCAD_App <> "" Then
            AppActivate sCAD_App
            SendKeys "@C:\JOBST\DRAW.D{enter}", True
            End
        Else
            MsgBox "Can't find drafix", 16, "Leg Dialogue"
        End If
    End If
End Sub

Private Sub ExtendLegTapes_Click()
   'Takes the last tape given and add the required value
   'GOP - 01-02/12, 6.2.1
   '
    Dim ii As Integer
    Dim nValue As Double

   'Locate last tape
    For ii = 29 To 0 Step -1
        If Val(txtLeft(ii)) > 0 Then Exit For
    Next ii

   'Check that there are some tapes.
   'Check that there is room to add a new tape
    If ii = 0 Or ii = 29 Then
        Beep
        Exit Sub
    End If

   'Convert given value to inches
    nValue = fnDisplaytoInches(Val(txtLeft(ii)))

   'Extend tape
    If txtSex = "Male" Then
        nValue = nValue + 0.5
    Else
        nValue = nValue + 1
    End If

    If g_iStyleLastTape = ii Then PR_SetLastTape ii + 1

    txtLeft(ii + 1) = fnInchesToDisplay(nValue)
    grdLeftInches.Row = ii + 1
    grdLeftInches = fnInchesToText(nValue)

End Sub

Private Function FN_BuildStyleString() As String

    Dim sString As String, sFabricClass  As String
    Dim sFootLength As String, nSave As Integer

    If g_POWERNET = True Then sFabricClass = "0"
    If g_JOBSTEX = True Then sFabricClass = "1"
    If g_JOBSTEX_FL = True Then sFabricClass = "2"

    nSave = 0

   'Transfer to the "legbox" the following set of data in blank delimited format
   'these will be stored in the fields relevent data fields
   '
   '                            DATA - Field
   '                            ~~~~~~~~~~~~
   '                LegStyle                            (1)
   '                First Tape of style                 (2)
   '                Last Tape of style                  (3)
   '                AnkleTape#                          (4)
   '                Pressure                            (5)
   '                [Grams|Stretch]                     (6)
   '                Reduction                           (7)
   '                AnkleLength                         (8)
   '                HeelLength                          (9)
   '                Zipper Status                       (10)
   '                FabricClass                         (11)
   '                Toe Style %%%                       (12)
   '                Foot Length                         (13)
   '                Template                            (14)
   '                Heel Style  ***                     (15)
   '                Heel Reinforcement ***              (16)
   '
   '          *** = Not yet implemented
   '          %%% = Above Knee / Below Knee for Footless styles
   '
   'Where Leg Style has the following meanings :-
   '
   '    0 = Anklet
   '    1 = Knee High
   '    2 = Thigh High
   '    3 = Knee band
   '    4 = Thigh band
   '
   'The First and Last tape positions are style dependant.
    If txtFootLength.Text = "" Then
        sFootLength = "0"
    Else
        sFootLength = Str$(txtFootLength.Text)
    End If

    sString = Str$(g_iLegStyle) & " "
    If g_iLegStyle < 3 Then
        sString = sString & Str$(g_iFirstTape + nSave) & " "
    Else
        sString = sString & Str$(g_iStyleFirstTape + nSave) & " "
    End If
    sString = sString & Str$(g_iStyleLastTape + nSave) & " "
    
    Select Case g_iLegStyle
        Case 0
            If g_iLtAnkle <> 0 Then
                sString = sString & Str$(g_iLtAnkle + 1) & " "
                sString = sString & "0 "
                sString = sString & "0 "
                sString = sString & "0 "
                sString = sString & Str$(g_nLtLastAnkle) & " "
                sString = sString & Str$(g_nLtLastHeel) & " "
                sString = sString & Str$(g_iLtLastZipper) & " "
                sString = sString & sFabricClass & " "
                sString = sString & Str$(cboToeStyle.ListIndex) & " "
                sString = sString & sFootLength & " "
                sString = sString & Str$(cboLeftTemplate.ListIndex) & " -1 -1"
            Else
               'Set to reflect the footless state
                sString = sString & "-1 0 0 0 0 0 0 "
                sString = sString & sFabricClass
                sString = sString & " 0 0 "
                sString = sString & Str$(cboLeftTemplate.ListIndex) & " -1 -1"
            End If
        Case 1, 2
            If g_iLtAnkle <> 0 Then
                sString = sString & Str$(g_iLtAnkle + 1) & " "
                sString = sString & Str$(g_iLtMM(g_iLtAnkle)) & " "
                sString = sString & Str$(g_iLtStretch(g_iLtAnkle)) & " "
                sString = sString & Str$(g_iLtRed(g_iLtAnkle)) & " "
                sString = sString & Str$(g_nLtLastAnkle) & " "
                sString = sString & Str$(g_nLtLastHeel) & " "
                sString = sString & Str$(g_iLtLastZipper) & " "
                sString = sString & sFabricClass & " "
                sString = sString & Str$(cboToeStyle.ListIndex) & " "
                sString = sString & sFootLength & " "
                sString = sString & Str$(cboLeftTemplate.ListIndex) & " -1 -1"
            Else
               'Set to reflect the footless state
                sString = sString & "-1 0 0 0 0 0 0 "
                sString = sString & sFabricClass
                sString = sString & " 0 0 "
                sString = sString & Str$(cboLeftTemplate.ListIndex) & " -1 -1"
            End If

        Case 3, 4, 5
                sString = sString & "-1 0 0 0 0 0 0 "
                sString = sString & sFabricClass
                sString = sString & " 0 0 "
                sString = sString & Str$(cboLeftTemplate.ListIndex) & " -1 -1"
    End Select

    FN_BuildStyleString = sString

End Function

Private Function FN_ConcatData() As String
   'Concatenates the displayed data
   'this can then be used to check if any modifications have been made
   '
    Dim ii As Integer
    Dim vData As Variant

   'Initialise to blank string
    vData = ""

    For ii = 0 To 29
        vData = vData + txtLeft(ii).Text
        If ii >= 6 Then vData = vData + txtLeftMM(ii).Text
    Next ii
    For ii = 0 To 1
        vData = vData + Str$(optFabric(ii).Value)
    Next ii

    For ii = 0 To 5
        vData = vData + Str$(optType(ii).Value)
    Next ii

    vData = vData + cboFabric.Text + cboToeStyle.Text + txtFootLength.Text + txtFirstTape.Text + txtLastTape.Text + cboLeftTemplate.Text + Str$(chkLeftZipper.Value)
    vData = vData + txtFootPleat1.Text + txtFootPleat2.Text + txtTopLegPleat1.Text + txtTopLegPleat1.Text
    
    FN_ConcatData = vData
    
End Function

Private Function FN_DrawOpen(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fnum = FreeFile
    Open sDrafixFile For Output As fnum
    FN_DrawOpen = fnum

   'Initialise String globals
    CC = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & CC & QQ  'Quote Comma Quote ( "," )
    QC = QQ & CC        'Quote Comma ( ", )
    CQ = CC & QQ        'Comma Quote ( ," )

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName
    
   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextHt = 0.125
    g_nCurrTextAspect = 0.6
    g_nCurrTextHorizJust = 1    'Left
    g_nCurrTextVertJust = 8    'Bottom
    g_nCurrTextFont = 0         'BLOCK

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fnum, "//DRAFIX Leg Drawing Macro created - "; Date$; "  "; Time$
    Print #fnum, "//Patient - "; g_sPatient; CC; " "; g_sFileNo; CC; " SIDE - "; g_sSide
    Print #fnum, "//by Visual Basic"

   'Text data
    Print #fnum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fnum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fnum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fnum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fnum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"
    
    PR_PutLine "STRING sTitleName, sLeg, sTmp, sWorkOrder, sID, sPathJOBST;"

   'Path to JOBST installed directory
    PR_PutStringAssign "sPathJOBST", FN_EscapeSlashesInString(g_sPathJOBST)

End Function

Private Function FN_EscapeSlashesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeSlashesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = "\" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeSlashesInString = sEscapedString

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

Private Function FN_SaveOpen(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
 
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fnum = FreeFile
    Open sDrafixFile For Output As fnum
    FN_SaveOpen = fnum

   'Initialise String globals
    CC = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & CC & QQ  'Quote Comma Quote ( "," )
    QC = QQ & CC        'Quote Comma ( ", )
    CQ = CC & QQ        'Comma Quote ( ," )

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName
    
   'Write header information etc. to the DRAFIX macro file
   '
    Print #fnum, "//DRAFIX Leg Drawing Macro created - "; Date$; "  "; Time$
    Print #fnum, "//Patient - "; g_sPatient; CC; " "; g_sFileNo; CC; " SIDE - "; g_sSide
    Print #fnum, "//by Visual Basic"
    PR_PutLine "STRING sTitleName, sLeg, sTmp, sWorkOrder, sID, sPathJOBST;"

   'Path to JOBST installed directory
    PR_PutStringAssign "sPathJOBST", FN_EscapeSlashesInString(g_sPathJOBST)
    
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
   Dim sMessage As String, iLastStyleFound As Integer
   Dim sSpacer

  'Disable timeout timer
   Timer1.Enabled = False
    
  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

  'Use Title fabric if a fabric is not already given
   If txtLegTitleFabric.Text <> "" And txtFabric.Text = "" Then txtFabric.Text = txtLegTitleFabric.Text

  'Set dropdown combo boxes
  'Toe Style
   cboToeStyle.AddItem ""
   cboToeStyle.AddItem "Curved"
   cboToeStyle.AddItem "Cut-Back"
   cboToeStyle.AddItem "Straight"
   cboToeStyle.AddItem "Soft Enclosed"
   cboToeStyle.AddItem "Soft Enclosed B/M"
   cboToeStyle.AddItem "Self Enclosed"
  
  'Set value
   For ii = 0 To (cboToeStyle.ListCount - 1)
        If cboToeStyle.List(ii) = txtToeStyle.Text Then
            cboToeStyle.ListIndex = ii
        End If
   Next ii

  'Leg Option buttons
   If txtLeg.Text = "Left" Then
        lglegdia.Caption = "LEG Details - Left"
        optLeftLeg.Value = True
        optRightLeg.Enabled = False
   End If
   If txtLeg.Text = "Right" Then
        lglegdia.Caption = "LEG Details - Right"
        optRightLeg.Value = True
        optLeftLeg.Enabled = False
   End If

  'Update tape boxes
  'Get first and last tapes in the text boxes
   g_iFirstTape = -1
   g_iLastTape = -1

   g_iStyleFirstTape = -1
   g_iStyleLastTape = 30

   nValue = 0
   grdLeftInches.Col = 0
   For ii = 0 To 29
        nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
        If nValue > 0 Then
            txtLeft(ii) = nValue
            nValue = fnDisplaytoInches(nValue)
            grdLeftInches.Row = ii
            grdLeftInches.Text = fnInchesToText(nValue)
        End If
        If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii
        If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii - 1
   Next ii
   If nValue > 0 Then g_iLastTape = 29

  'Set leg style options
  'The chosen leg style is given when the user has selected a
  'pattern from the drawing rather than using one of the buttons.

   g_iLegStyle = fnGetNumber(txtChosenStyle.Text, 1)
   
   If g_iLegStyle >= 0 Then
        optType_Click g_iLegStyle
        optType(g_iLegStyle).Value = True
        PR_SetLastTape g_iStyleLastTape
   Else
        PR_LastTapeDisplay "Disabled"
        PR_FirstTapeDisplay "Disabled"
   End If

   If g_iLegStyle >= 3 Then PR_SetFirstTape g_iStyleFirstTape


  'Establish how the leg is to be figured
  'Base this on chosen fabric and availability of mm
  'If it is not given explicitly
  
  'Check for an explititly given fabric class from the Style field
  'This Field will be emplty if no previous figuring has been done
  'Note This is a multi data. I the first number is -ve then there is no ankle.
  '     Also the function fnGetNumber returns -1 if the number does not exist.
  '
   iFabricClass = -1
   If fnGetNumber(g_sStyleString, 4) <> -1 Then
        iFabricClass = fnGetNumber(g_sStyleString, 11)
   End If

  'Get fabric class from other data
   If iFabricClass < 0 Then
       If txtFabric.Text = "" Then
            If txtDiagnosis.Text = "Burns" Then
                g_POWERNET = True
            Else
'                If Left$(txtDiagnosis.Text, 5) = "Lymph" Then
'                    g_JOBSTEX_FL = True
'                Else
'                    g_JOBSTEX = True
'                End If
                g_JOBSTEX = True
           End If
       Else
           If Left$(txtFabric.Text, 3) = "Pow" Then
                g_POWERNET = True
           Else
'                If txtLeftMMs.Text <> Then
'                    g_JOBSTEX_FL = True
'                Else
'                    g_JOBSTEX = True
'                End If
                g_JOBSTEX = True
           End If
       End If
   Else
        If iFabricClass = 0 Then g_POWERNET = True
        If iFabricClass = 1 Then g_JOBSTEX = True
        If iFabricClass = 2 Then g_JOBSTEX = True
'        If iFabricClass = 2 Then g_JOBSTEX_FL = True
   End If
   
   If g_JOBSTEX = True Then optFabric(1).Value = True
'   If g_JOBSTEX_FL = True Then optFabric(2).Value = True

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

  'Set up depending on leg style chosen
  'if only a single style available then display the values for that
  'style only
  'other wise display as blank untill the user selects an option

   PR_EstablishAnkles

   If g_POWERNET = True Then PR_EnablePOWERNET
   If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX
'   If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX


  'From the g_sStyleString set above extract the saved ankle figuring
  'We have to be careful in the order in which this is set up as in the case
  'of JOBSTEX_FL we do not want to cause recaculation on the ankle.
  'We also must take care only to add values only if they exist
  '(thus we check that the returned value from fnGetNumber is not -ve)
  '
   Dim iFabric%
   If g_sStyleString <> "" And g_iLtAnkle <> 0 Then
        iZipper = fnGetNumber(g_sStyleString, 10)
        If iZipper >= 0 Then chkLeftZipper.Value = iZipper  'Order is v.important here
        iMMHg = fnGetNumber(g_sStyleString, 5)
        If iMMHg >= 0 Then
            g_iLtMM(g_iLtAnkle) = iMMHg
            txtLeftMM(g_iLtAnkle) = iMMHg
        End If
        If g_JOBSTEX_FL = True Then
            nStretch = fnGetNumber(g_sStyleString, 6)
            iFabric = Val(Left$(cboFabric.Text, 2))
            If nStretch >= 0 Then
                    PR_DisplayFiguredAnkle "Left", nStretch, g_nLtLastAnkle, g_nLtLastHeel, iFabric
                    g_iLtStretch(g_iLtAnkle) = nStretch
            End If
        End If
   End If


  'Don't figure ankle for fabric class JOBSTEX_FL as the user can modify the pressures
  'at each leg tape manually and these are stored.  Figuring would overwrite
  'any manual changes
   If g_JOBSTEX_FL <> True Then
        PR_FigureLeftAnkle
   Else
        If txtLeftMMs.Text <> "" And g_iLtAnkle <> 0 Then
            For ii = g_iLtAnkle + 1 To 29
                iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
                If iValue > 0 Then
                    txtLeftMM(ii) = iValue
                    PR_FigureLeftTape ii
                End If
            Next ii
        End If
   End If

  'Disable anklet and knee if no ankle tape
   If g_iLtAnkle = 0 Then
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
   End If
   
  'Find existing styles display last found
   sMessage = ""
   sSpacer = ""
   iLastStyleFound = -1

   If txtAnklet.Text <> "" Then
        sMessage = sMessage + sSpacer + "Anklet"
        sSpacer = ", "
        iLastStyleFound = 0
   End If

   If txtKneeLength.Text <> "" Then
        sMessage = sMessage + sSpacer + "Knee Length"
        sSpacer = ", "
        iLastStyleFound = 1
   End If
   If txtThighLength.Text <> "" Then
        sMessage = sMessage + sSpacer + "Thigh Length"
        sSpacer = ", "
        iLastStyleFound = 2
   End If
   If txtKneeBand.Text <> "" Then
        sMessage = sMessage + sSpacer + "Knee Band"
        sSpacer = ", "
        iLastStyleFound = 3
   End If
   If txtThighBandAK.Text <> "" Then
        sMessage = sMessage + sSpacer + "Thigh (A/K)"
        sSpacer = ", "
        iLastStyleFound = 4
   End If
   If txtThighBandBK.Text <> "" Then
        sMessage = sMessage + sSpacer + "Thigh (B/K)"
        sSpacer = ", "
        iLastStyleFound = 5
   End If

   If iLastStyleFound >= 0 Then
        optType_Click (iLastStyleFound)
        optType(iLastStyleFound).Value = True
        labMessage.Caption = "Existing Styles: " + Chr$(13) + sMessage
   End If

   g_sChangeChecker = FN_ConcatData()
   Show
   Screen.MousePointer = 0    'Change pointer to default.
   lglegdia.SetFocus

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
        MsgBox "The Leg input Module is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting Figure"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - lglegdia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - lglegdia.Height) / 2  ' Center form vertically.
    
    Set MainForm = lglegdia

    g_nUnitsFac = 1 'Default to inches
    g_iLegStyle = -1
    g_JOBSTEX = False
    g_JOBSTEX_FL = False
    g_POWERNET = False
    g_sTextList = "-7½ -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

   'Setup display inches grid
    grdLeftInches.ColWidth(0) = 880
    grdLeftInches.ColAlignment(0) = 2

    For ii = 0 To 29
        grdLeftInches.RowHeight(ii) = 266
    Next ii

   'Setup display of results grid
    For ii = 0 To 1
        grdLeftDisplay.ColWidth(ii) = 488
        grdLeftDisplay.ColAlignment(ii) = 2
    Next ii
    
    For ii = 0 To 23
        grdLeftDisplay.RowHeight(ii) = 266
    Next ii


    g_sPathJOBST = fnPathJOBST()

   'Enable time out timer
    Timer1.Interval = 6000
    Timer1.Enabled = True

End Sub

Private Sub optFabric_Click(Index As Integer)
   'Allows the selection of fabric class
   'Used to override the fabric class selection established at link close

   'Clean Message box
    labMessage.Caption = ""
 
   'Do nothing if same fabric class selected
    If g_POWERNET = True And Index = 0 Then Exit Sub
    If g_JOBSTEX = True And Index = 1 Then Exit Sub
    If g_JOBSTEX_FL = True And Index = 2 Then Exit Sub
   
   'Restablish ankles just in case they have changed
    PR_EstablishAnkles

   'Set the display for the different fabric classes
   'We must first reset the form to the blank form that we use on form load
    PR_ResetFormLGLEGDIA
    
    g_POWERNET = False
    g_JOBSTEX = False
    g_JOBSTEX_FL = False

    Select Case Index
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
        'PR_EnableFL_JOBSTEX
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
        g_iLtLastFabric = -1000  'This will force re-figuring
        If g_iLtMM(g_iLtAnkle) > 0 Then
            txtLeftMM(g_iLtAnkle) = g_iLtMM(g_iLtAnkle)
            PR_FigureLeftAnkle
        End If
        g_iLtLastFabric = -1
    End If

End Sub

Private Sub optType_Click(Index As Integer)
    
    Dim ii As Integer, nValue  As Double

    g_iLastLegStyle = g_iLegStyle
    g_iLegStyle = Index
    
   'get the Style String g_sStyleString from the relevent DDE box
   'or rebuild string from existing displayed data
    Select Case g_iLegStyle
        Case 0
            g_sStyleString = txtAnklet.Text
'            g_sBestCatNo = "0105"   'Anklet
'            g_nOptions = 0
        Case 1
            g_sStyleString = txtKneeLength.Text
'            g_sBestCatNo = "0101"   'Knee length
'            g_nOptions = 0
        Case 2
            g_sStyleString = txtThighLength.Text
'            g_sBestCatNo = "0201"   'Thigh length
'            g_nOptions = 0
        Case 3
            g_sStyleString = txtKneeBand.Text
'            g_sBestCatNo = "0015"   'Knee band
'            g_nOptions = 3
'            g_sCatNo(1) = "0015"    'Knee band
'            g_sCatNo(2) = "1131"    'Stump Support Below Knee
'            g_sCatNo(3) = "0101"    'Footless Knee Length
        Case 4
            g_sStyleString = txtThighBandAK.Text
'            g_sBestCatNo = "0019"   'Thigh band
'            g_nOptions = 3
'            g_sCatNo(1) = "0019"    'Thigh band
'            g_sCatNo(2) = "1130"    'Stump Support above Knee
'            g_sCatNo(3) = "0201"    'Footless thigh Length

        Case 5
            g_sStyleString = txtThighBandBK.Text
'            g_sBestCatNo = "0019"    'Thigh band
'            g_nOptions = 3
'            g_sCatNo(1) = "0019"    'Thigh band
'            g_sCatNo(2) = "1131"    'Stump Support Below Knee
'            g_sCatNo(3) = "0201"    'Footless thigh Length

    End Select

   'Find First and last Tapes as displayed
   'Don't worry about holes for now
    g_iFirstTape = -1
    For ii = 0 To 29
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    If ii < 30 Then g_iFirstTape = ii

    g_iLastTape = -1
    For ii = 29 To 0 Step -1
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    If ii >= 0 Then g_iLastTape = ii

    If g_iLastTape = g_iFirstTape Then
        g_iFirstTape = -1
        g_iLastTape = -1
    End If

    PR_LastTapeDisplay "Enabled"

    Select Case Index
        Case 0      'Anklet
           'Disable FirstTape display
            PR_FirstTapeDisplay "Disabled"
           'Disable heel figuring
           'Set last tape to ankle + 1
            If g_iLtAnkle <> 0 Then
                PR_SetLastTape g_iLtAnkle + 1
                txtLeftMM(g_iLtAnkle).Enabled = False
                txtLeftMM(g_iLtAnkle).Text = ""
                grdLeftDisplay.Row = g_iLtAnkle - 6
                grdLeftDisplay.Col = 0
                grdLeftDisplay = ""
                grdLeftDisplay.Col = 1
                grdLeftDisplay = ""
                cboToeStyle.ListIndex = fnGetNumber(g_sStyleString, 12)
            End If

            If g_POWERNET = True Then cboLeftTemplate.ListIndex = 0
        
        Case 1 To 2   'Knee Length, Thigh Length
           'Disable FirstTape display
            PR_FirstTapeDisplay "Disabled"
           'Enable and figure ankle
            If g_iLtAnkle <> 0 Then
                txtLeftMM(g_iLtAnkle).Enabled = True
                cboToeStyle.ListIndex = fnGetNumber(g_sStyleString, 12)
                If g_iLtLastMM <> 0 Then
                    txtLeftMM(g_iLtAnkle) = g_iLtLastMM
                    g_iLtLastFabric = -1000 'Use this to force a refigure
                    PR_FigureLeftAnkle
                End If
            End If
           
           'Set Last style tape from StyleString
            If g_sStyleString <> "" Then
                PR_SetLastTape fnGetNumber(g_sStyleString, 3) - 1
            End If
           
           'Ensure that the style last tape is valid, reset if not
            If g_iStyleLastTape > 0 And g_iStyleLastTape <= g_iLastTape Then
                PR_SetLastTape g_iStyleLastTape
            Else
                PR_SetLastTape g_iLastTape
            End If

        Case 3, 4, 5  'Knee Band, Thigh Band AK and Thigh Band BK
           'Enable FirstTape display
            PR_FirstTapeDisplay "Enabled"
           
           'Disable Heel figuring
            If g_iLtAnkle <> 0 Then
                txtLeftMM(g_iLtAnkle).Text = ""
                txtLeftMM(g_iLtAnkle).Enabled = False
                grdLeftDisplay.Row = g_iLtAnkle - 6
                grdLeftDisplay.Col = 0
                grdLeftDisplay = ""
                grdLeftDisplay.Col = 1
                grdLeftDisplay = ""
            End If

           'Set Style first tape
           'Use existing if available else use first tape as above
           'don't go past ankle
           'Set First and Last style tapes from StyleString
            If g_sStyleString <> "" Then
                PR_SetFirstTape fnGetNumber(g_sStyleString, 2) - 1
                PR_SetLastTape fnGetNumber(g_sStyleString, 3) - 1
            End If
           
           'Ensure that the style last tape is valid, reset if not
            If g_iStyleFirstTape > 0 And g_iStyleFirstTape >= g_iFirstTape Then
                PR_SetFirstTape g_iStyleFirstTape
            Else
                If g_iLtAnkle <> 0 Then
                    PR_SetFirstTape g_iLtAnkle
                Else
                    PR_SetFirstTape g_iFirstTape
                End If
            End If

            If g_iStyleLastTape > 0 And g_iStyleLastTape <= g_iLastTape Then
                PR_SetLastTape g_iStyleLastTape
            Else
                PR_SetLastTape g_iLastTape
            End If
     
            cboLeftTemplate.ListIndex = 0

     End Select

     If g_sStyleString <> "" Then
       'Restore values from style string
        chkLeftZipper.Value = fnGetNumber(g_sStyleString, 10)

        nValue = fnGetNumber(g_sStyleString, 13)
        If nValue <> 0 Then
            txtFootLength = nValue
        Else
            txtFootLength = ""
        End If

        txtFootLength_LostFocus
        
        If fnGetNumber(g_sStyleString, 14) <= cboLeftTemplate.ListCount - 1 Then
            cboLeftTemplate.ListIndex = fnGetNumber(g_sStyleString, 14)
        Else
            cboLeftTemplate.ListIndex = -1
        End If
        If g_iLtAnkle <> 0 And g_iLegStyle > 0 And g_iLegStyle < 3 And fnGetNumber(g_sStyleString, 5) <> 0 Then
            txtLeftMM(g_iLtAnkle) = fnGetNumber(g_sStyleString, 5)
            g_iLtLastFabric = -1000 'Use this to force a refigure
            PR_FigureLeftAnkle
        End If
     Else
       'Build style string from exisisting data
        g_sStyleString = FN_BuildStyleString()
     End If

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
    'Messages with repect to available Styles
     Case 500
        sError = "Anklet"
     Case 501
        sError = "Knee Length"
     Case 502
        sError = "Thigh Length"
     Case 503
        sError = "Knee Band"
     Case 504
        sError = "Thigh Band A/K"
     Case 505
        sError = "Thigh Band B/K"
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
        labMessage.Caption = sError
    Else
        labMessage.Caption = labMessage.Caption + NL + NL + sError
    End If
    Beep

End Sub

Private Sub PR_DisplayFiguredAnkle(sLeg As String, nStretch As Double, nAnkleCir As Double, nHeelCir As Double, iFabric As Integer)
    Dim nPatternDim As Double, iReduction As Integer, iMaxStretch  As Integer
    Dim sContext As String, iDonningStretch  As Integer, iMinStretch  As Integer

    nPatternDim = (nAnkleCir / (1 + (0.01 * nStretch))) / g_nUnitsFac
    iReduction = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
    iDonningStretch = round((((nHeelCir / g_nUnitsFac) - nPatternDim) / nPatternDim) * 100)

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


      'Set Max and Min stretches
       iMaxStretch = 44
       If iFabric <= 55 Then iMinStretch = 16 Else iMinStretch = 18
       
       If nStretch < iMinStretch Then
            PR_DisplayErrorMessage 2002, "Ankle Figuring"
       End If

       If nStretch > iMaxStretch Then
            PR_DisplayErrorMessage 2001, "Ankle Figuring"
       End If
       
       If iDonningStretch > 110 Then
            PR_DisplayErrorMessage 2003, "Ankle Figuring"
       End If
       
       If iReduction > 32 Then
            PR_DisplayErrorMessage 2004, "Ankle Figuring"
       End If
       

  
End Sub

Private Sub PR_DrawAndSaveLeg()
   'Procedure to create a Macro to draw a leg lower options
   'Anklet,  Knee High, Thigh High, Knee Band and Thigh Band
    
    Dim sFile As String, sBody  As String
    Dim sLengths As String, sData As String, sLegStyle As String
    Dim nLegStyle As Integer, nStyleLastRed As Integer
    Dim nFirstTape As Integer, nLastTape As Integer, nFabricClass As Integer
    Dim nStyleFirstTape As Integer, nStyleLastTape As Integer, nAge  As Integer
    Dim ii As Integer, FootLess As Integer, itemplate As Integer

    Dim nThighPltXoff As Double, nThighPltYoff As Double
    Dim nThighPltRad As Double, nTopThigh As Double
    Dim nThighPltDeltaAngle As Double
    Dim nThighPltLen As Double, nHeel As Double, nThighTopExtension  As Double
    Dim nA As Double, nB As Double, nThighPltStartAngle As Double

    Dim nFoldHt As Double, nValue As Double

   'Open Macro file (fNum is declared as Global)
    sFile = "C:\JOBST\DRAW.D"
    fnum = FN_DrawOpen(sFile, txtPatientName, txtFileNo, txtLeg)

    PR_SaveLeg

    PR_PutLine "@" & g_sPathJOBST & "\LEG\LG_LEG.D;"

   
   'Patient Details
    PR_PutStringAssign "sPatient", txtPatientName
    PR_PutStringAssign "sFileNo", txtFileNo

    If txtWorkOrder.Text = "" Then
        PR_PutStringAssign "sWorkOrder", "-"
    Else
        PR_PutStringAssign "sWorkOrder", txtWorkOrder
    End If
    
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

   'Body details
    PR_PutStringAssign "sFabric", txtFabric

   'Leg Details
    sLengths = txtLeftLengths.Text  'Use later to establish 1st & Last tapes
       
    If Len(txtFabric.Text) > 0 Then PR_PutStringAssign "sFabric", txtFabric

    If fnGetNumber(g_sStyleString, 4) < 0 Or fnGetNumber(g_sStyleString, 1) > 2 Or g_iLtAnkle = 0 Then
        PR_PutLine "FootLess = %true;"
        FootLess = True
    Else
        PR_PutLine "FootLess = %false;"
       'Ankle tape values
        PR_PutNumberAssign "nAnkleTape", fnGetNumber(g_sStyleString, 4)
        PR_PutStringAssign "sAnkleTape", Str$(fnGetNumber(g_sStyleString, 4))
        PR_PutStringAssign "sMMAnkle", Str$(fnGetNumber(g_sStyleString, 5))
        PR_PutStringAssign "sGramsAnkle", Str(fnGetNumber(g_sStyleString, 6))
        PR_PutStringAssign "sReductionAnkle", Str$(fnGetNumber(g_sStyleString, 7))
        PR_PutNumberAssign "nReductionAnkle", fnGetNumber(g_sStyleString, 7)
        nHeel = fnGetNumber(g_sStyleString, 9)
    End If
    
    nFabricClass = fnGetNumber(g_sStyleString, 11)
    If nFabricClass = 2 Then
        PR_PutStringAssign "sReduction", txtLeftRed
        PR_PutStringAssign "sTapeMMs", txtLeftMMs
        PR_PutStringAssign "sStretch", txtLeftStr
    End If

    PR_PutStringAssign "sLeg", txtLeg
    
    PR_PutStringAssign "sPressure", txtLeftTemplate
    itemplate = Left(txtLeftTemplate.Text, 2)
    
    PR_PutStringAssign "sTapeLengths", txtLeftLengths
    PR_PutStringAssign "sToeStyle", txtToeStyle
    
    PR_PutNumberAssign "nFootPleat1", fnDisplaytoInches(Val(txtFootPleat1.Text))
    PR_PutNumberAssign "nFootPleat2", fnDisplaytoInches(Val(txtFootPleat2.Text))
    
    PR_PutStringAssign "sFootLength", txtFootLength

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

   'First and last tapes for style
    nStyleFirstTape = fnGetNumber(g_sStyleString, 2)
    PR_PutNumberAssign "nStyleFirstTape", nStyleFirstTape
    
    nStyleLastTape = fnGetNumber(g_sStyleString, 3)
    PR_PutNumberAssign "nStyleLastTape", nStyleLastTape
    
   'Get last tape value for use in calculating thigh length ending
    nValue = fnDisplaytoInches(Val(Mid(sLengths, ((nStyleLastTape - 1) * 4) + 1, 4)) / 10)

   'Style type etc
   'NOTE for a footless thigh length treat as a thigh band (BK)
    nLegStyle = fnGetNumber(g_sStyleString, 1)
    If nLegStyle = 2 And FootLess Then nLegStyle = 5

    PR_PutNumberAssign "nLegStyle", nLegStyle
    PR_PutStringAssign "sType", g_sStyleString
    PR_PutNumberAssign "nTopLegPleat1", fnDisplaytoInches(Val(txtTopLegPleat1.Text))
    PR_PutNumberAssign "nTopLegPleat2", fnDisplaytoInches(Val(txtTopLegPleat2.Text))
    
    Select Case nLegStyle
    Case 1  'Knee High
        sLegStyle = "KLN"

        If nFabricClass = 0 And nValue < 15 Then
            nStyleLastRed = 10
        Else
            nStyleLastRed = 14
        End If

    Case 2 'Thigh Length
        sLegStyle = "TLN"

        If nFabricClass = 0 Then
            nStyleLastRed = 14
        Else
            nStyleLastRed = 16
        End If
       
       'Calculate parameters for Thigh Plate Template
        nTopThigh = ((nValue * ((100 - nStyleLastRed) / 100)) / 2) + 0.1875
        nThighPltRad = 23

        Select Case nTopThigh   'Get closest thigh plate
         Case 0 To 7
            nThighPltXoff = 0.52
            nThighPltLen = 6
         Case 7 To 9
            nThighPltXoff = 0.73
            nThighPltLen = 8
         Case 9 To 11
            nThighPltXoff = 1
            nThighPltLen = 10
         Case 11 To 13
            nThighPltXoff = 1.27
            nThighPltLen = 12
         Case Is > 13
            nThighPltXoff = 1.5
            nThighPltLen = 14
        End Select
       
       'Depending on Heel set top of thigh extension
        If nHeel < 9 Then
            nThighTopExtension = 0.5
            PR_PutNumberAssign "nThighTopExtension", nThighTopExtension
        Else
            nThighTopExtension = 1
            PR_PutNumberAssign "nThighTopExtension", nThighTopExtension
        End If

       'Revise thigh plate to make it fit (if required)
       'Note the thigh plate line must land on the profile, on or
       'in front of the last tape
        If nThighPltXoff < nThighTopExtension Then nThighPltXoff = 1

        PR_PutNumberAssign "nThighPltRad", nThighPltRad
        PR_PutNumberAssign "nThighPltXoff", nThighPltXoff

    Case 3 'Knee Band
        sLegStyle = "KBN"
       
        PR_PutNumberAssign "nStyleFirstRed", 8
        nStyleLastRed = 8

       'Elastic at distal always
        PR_PutNumberAssign "nElastic", 1

    Case 4, 5 'Thigh Band Above Knee and ThighBand Below Knee
        If nFabricClass = 0 Then
            nStyleLastRed = 14
        Else
            nStyleLastRed = 16
        End If

       'Release the distal tape to a 95% reduction,  92% reduction if +3 or +1-1/2 tape
        If nStyleFirstTape > 8 Then
            PR_PutNumberAssign "nStyleFirstRed", 5
        Else
            PR_PutNumberAssign "nStyleFirstRed", 8
        End If
       
       'Calculate parameters for Thigh Plate Template
        nTopThigh = ((nValue * ((100 - nStyleLastRed) / 100)) / 2) + 0.1875
        nThighPltRad = 23

        Select Case nTopThigh   'Get closest thigh plate
         Case 0 To 7
            nThighPltXoff = 0.52
            nThighPltLen = 6
         Case 7 To 9
            nThighPltXoff = 0.73
            nThighPltLen = 8
         Case 9 To 11
            nThighPltXoff = 1
            nThighPltLen = 10
         Case 11 To 13
            nThighPltXoff = 1.27
            nThighPltLen = 12
         Case Is > 13
            nThighPltXoff = 1.5
            nThighPltLen = 14
        End Select
       
       'Depending on age set top of thigh extension
        If nAge < 10 Then
            nThighTopExtension = 0.5
            PR_PutNumberAssign "nThighTopExtension", nThighTopExtension
        Else
            nThighTopExtension = 1
            PR_PutNumberAssign "nThighTopExtension", nThighTopExtension
        End If

       'Revise thigh plate to make it fit (if required)
       'Note the thigh plate line must land on the profile, on or
       'in front of the last tape
        If nThighPltXoff < nThighTopExtension Then nThighPltXoff = 1

        PR_PutNumberAssign "nThighPltRad", nThighPltRad
        PR_PutNumberAssign "nThighPltXoff", nThighPltXoff

       'Elastic
        If nLegStyle = 4 Then
            PR_PutNumberAssign "nElastic", 1
            sLegStyle = "TBA"   'Thigh Band (A/K)
        Else
            PR_PutNumberAssign "nElastic", 0
            sLegStyle = "TBB"   'Thigh Band (B/K)
        End If


    Case Else 'Anklet,
        sLegStyle = "ANK"
       'Anklet reductions done by counting and therefor done in the
       'macro LGLEGDWG.D
        nStyleLastRed = -1
        PR_PutNumberAssign "nReductionAnkle", 14

    End Select

    PR_PutNumberAssign "nStyleLastRed", nStyleLastRed
    PR_PutStringAssign "sLegStyle", sLegStyle

   'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
   'For a footless style the only valid classes are 0 and 1
    If FootLess And nFabricClass = 2 Then nFabricClass = 1
    PR_PutNumberAssign "nFabricClass", nFabricClass

   'Load universal Procedures and Defaults
    PR_PutLine "@" & g_sPathJOBST & "\LEG\LGLEGDEF.D;"
   
   'Get Origin & Draw template
    PR_PutLine "@" & g_sPathJOBST & "\LEG\LGLEGTMP.D;"

   'Calculate Foot Points (If a foot exists)
   'Draw Leg and foot profile
    If Not FootLess Then
        PR_PutLine "@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;"
        PR_PutLine "@" & g_sPathJOBST & "\LEG\LGLEGDWG.D;"
    End If

   'Footless ThighLength and Thigh/Knee Bands
    If FootLess Then
        PR_PutLine "@" & g_sPathJOBST & "\LEG\LGBNDDWG.D;"
    End If

   'Draw closing lines
    Select Case nLegStyle
     Case 0
        PR_PutLine "@" & g_sPathJOBST & "\LEG\LGANKCLS.D;"
     Case 1, 2
        PR_PutLine "@" & g_sPathJOBST & "\LEG\LGLEGCLS.D;"
    End Select
   
   'Close macro file
    Close #fnum

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

   'Set value
    For ii = 0 To (cboLeftTemplate.ListCount - 1)
        If cboLeftTemplate.List(ii) = txtLeftTemplate.Text Then
            cboLeftTemplate.ListIndex = ii
        End If
    Next ii

  'Set the display for leg depending on the ankle and fabric and options.
  'g_JOBSTEX_FL is a special case where all of the stretch at tapes above the
  'ankle are calculated
  '
  'Panty Legs
  'If a template is not given then
  'then set 13DS for JOBSTEX, this is first item in list
  '
   If g_iLtAnkle = 0 Then
        If cboLeftTemplate.ListIndex = -1 Then cboLeftTemplate.ListIndex = 0
        cboLeftTemplate.Enabled = False
        chkLeftZipper.Enabled = False
        chkLeftZipper.Value = 0
   End If
    
   
   If g_iLtAnkle <> 0 Then
        For ii = 0 To 2
            labLeftDisp(ii).Visible = True
        Next ii
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
        chkLeftZipper.Enabled = True
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

    cboLeftTemplate.Enabled = True
   
   'Set value
    For ii = 0 To (cboLeftTemplate.ListCount - 1)
        If cboLeftTemplate.List(ii) = txtLeftTemplate.Text Then
            cboLeftTemplate.ListIndex = ii
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

   If g_iLtAnkle = 0 Then
        If cboLeftTemplate.ListIndex = -1 Then cboLeftTemplate.ListIndex = 1
        chkLeftZipper.Enabled = False
        chkLeftZipper.Value = 0
   End If

   If g_iLtAnkle <> 0 Then
        For ii = 0 To 2
            labLeftDisp(ii).Visible = True
        Next ii
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
        
        chkLeftZipper.Enabled = True

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
   g_nLtLastHeel = fnDisplaytoInches(Val(txtLeft(5).Text))
   If g_nLtLastHeel <> 0 Then
        If g_nLtLastHeel < 9 Then
            g_iLtAnkle = 6    'Small heel Ankle at +1 1/2 tape
        Else
            g_iLtAnkle = 7    'Large heel Ankle at +3 tape
        End If
   End If

   g_nLtLastAnkle = fnDisplaytoInches(Val(txtLeft(g_iLtAnkle)))

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
   '    The POWERNET fabric is based on the file g_sPathJOBST & "\WHFABRIC.DAT
   '    For clarity this file can contain blank lines.  Therefor we must check
   '    that the fabric selected is not blank.
    If cboFabric.List(cboFabric.ListIndex) = "" Then Exit Sub

   'Establish if any of the inputs have changed compared to the last stored
   'values.
   'If they have changed then recalculate and redisplay
   'NB  Use of GOTO, (structured programmers get stuffed!)

   'Heel Tape
    If fnDisplaytoInches(Val(txtLeft(5))) <> g_nLtLastHeel Then GoTo CALC
   
   'Ankle Tape
    If fnDisplaytoInches(Val(txtLeft(g_iLtAnkle))) <> g_nLtLastAnkle Then GoTo CALC
   
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
    
    nHeelCir = fnDisplaytoInches(Val(txtLeft(5).Text))
    g_nLtLastHeel = nHeelCir
        
    nAnkleCir = fnDisplaytoInches(Val(txtLeft(g_iLtAnkle).Text))
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
        If ii = 29 Then
            nLastTape = 29
        Else
            nLastTape = ii - 1
        End If
       
       'Establish pressure at last tape
       'As the last tape is at a 20 reduction we know that the stretch must
       'be 25%
        nLastTapeCir = fnDisplaytoInches(Val(txtLeft(nLastTape).Text))
        nLastTapeMMHg = round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
        
        txtLeftMM(nLastTape) = nLastTapeMMHg
        g_iLtStretch(nLastTape) = 25
        g_iLtRed(nLastTape) = 20
        g_iLtMM(nLastTape) = nLastTapeMMHg
        
       'First pass from Ankle to LastTape
        nStretch = g_iLtLastStretch
        iPrevMMHg = g_iLtLastMM
        For ii = g_iLtAnkle + 1 To nLastTape - 1
            nTapeCir = fnDisplaytoInches(Val(txtLeft(ii).Text))
            nMMHg = 1000 'Force While
            While nMMHg > iPrevMMHg And nStretch > 1
                nMMHg = round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
                If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
            Wend
            txtLeftMM(ii) = nMMHg
            g_iLtMM(ii) = nMMHg
            iPrevMMHg = nMMHg
            g_iLtStretch(ii) = nStretch
            g_iLtRed(ii) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
        Next ii

       'Second pass from LastTape to Ankle
       'Based on last tape Pressure recalculate the tapes until the pressures
       'become the same
        iPrevMMHg = nLastTapeMMHg

        For ii = nLastTape - 1 To g_iLtAnkle + 1 Step -1
            nTapeCir = fnDisplaytoInches(Val(txtLeft(ii).Text))

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
            g_iLtStretch(ii) = nStretch
            g_iLtRed(ii) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
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

Private Sub PR_FigureLeftTape(Index As Integer)
   'Figures the Pressure/Stretch at a LEFT LegTape
   'This is only available to JOBSTEX_FL
   'See PR_FigureLeftTape for details
    
    Dim nMMHg As Double, nStretch As Double
    Dim nTapeCir As Double
    Dim iZipper As Integer, iFabric As Integer

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly
    If Val(txtLeftMM(Index).Text) <= 0 Then Exit Sub

   'Establish if pressure has changed from last time
   'If it hasn't then exit
   'This is used to cure a problem in that the stretch for a tape is not based not on the
   'pressure but on the previous or subsequent stretch, depending if it was found
   'on the forward pass or the backward pass.
   'The pressure was then back calculated fron the stretch.
   'We had a problem in that the same pressue can give a different stretch
   'To stop this happening when tabbing through the pressures this check has
   'been introduced
   '
    If Val(txtLeftMM(Index).Text) = g_iLtMM(Index) Then Exit Sub
    

   'MMHg and Stretch for JOBSTEX fabric
   'Don't Calculate if no values for Ankle tape
    If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle)) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub


   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
    
    nTapeCir = fnDisplaytoInches(Val(txtLeft(Index).Text))
    iZipper = chkLeftZipper.Value
    iFabric = Val(Left$(cboFabric.List(cboFabric.ListIndex), 2))
    nMMHg = Val(txtLeftMM(Index).Text)
    g_iLtMM(Index) = nMMHg

   'Calculate stretch
    nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, iZipper, iFabric)
    
   'Store values
    g_iLtStretch(Index) = round(nStretch)
    g_iLtRed(Index) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)

   'Display values
   'Note, the display grid displays from the ankle only
   '
    grdLeftDisplay.Row = Index - 6 'Display from ankle
    grdLeftDisplay.Col = 0
    grdLeftDisplay = Str$(g_iLtStretch(Index))
    grdLeftDisplay.Col = 1
    grdLeftDisplay = Str$(g_iLtRed(Index))

End Sub

Private Sub PR_FirstTapeDisplay(sType As String)
    
    If sType = "Disabled" Then
        txtFirstTape.Text = ""
        txtFirstTape.Enabled = False
        labFirstTape.Enabled = False
        If g_iStyleFirstTape >= 0 And g_iStyleFirstTape <= 29 Then Label1(g_iStyleFirstTape).ForeColor = RGB(0, 0, 0)
        spnFirstTape.Enabled = False
    Else
        txtFirstTape.Enabled = True
        labFirstTape.Enabled = True
        spnFirstTape.Enabled = True
    End If

End Sub

Private Sub PR_HeelChange(sLeg As String)
   'Procedure to modify the display if heel changed

    Dim iLtExistingAnkle  As Integer, iRtExistingAnkle As Integer
   'Clean Message box
    labMessage.Caption = ""

   'Exit if heel has not been changed
    If fnDisplaytoInches(Val(txtLeft(5).Text)) = g_nLtLastHeel Then Exit Sub

   'Store existing ankle position and get new ones
    iLtExistingAnkle = g_iLtAnkle
    
    PR_EstablishAnkles

   'If revised the same as existing then exit
    If iLtExistingAnkle = g_iLtAnkle Then Exit Sub

   'Modify display
    PR_ResetFormLGLEGDIA
   
   'Disable / Enable Anklet | Knee
    If g_iLtAnkle = 0 Then
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
    Else
        optType(0).Enabled = True
        optType(1).Enabled = True
        optType(2).Enabled = True
    End If

    If g_POWERNET = True Then PR_EnablePOWERNET
    If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX
   ' If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX

End Sub

Private Sub PR_LastTapeDisplay(sType As String)
    If sType = "Disabled" Then
        txtLastTape.Text = ""
        txtLastTape.Enabled = False
        labLastTape.Enabled = False
        If g_iStyleLastTape >= 0 And g_iStyleLastTape <= 29 Then Label1(g_iStyleLastTape).ForeColor = RGB(0, 0, 0)
        spnLastTape.Enabled = False
    Else
        txtLastTape.Enabled = True
        labLastTape.Enabled = True
        spnLastTape.Enabled = True
    End If

End Sub

Private Sub PR_LoadFabricFromFile(sFileName As String)
   'Procedure to load the POWERNET conversion chart from file
   'N.B. File opening Errors etc. are not handled (so tough titty!)
    
    Dim fnum As Integer, ii As Integer
    fnum = FreeFile
    Open sFileName For Input As fnum
    ii = 0
    Do Until EOF(fnum)
        Input #fnum, POWERNET.Modulus(ii), POWERNET.Conversion(ii)
        ii = ii + 1
    Loop
    Close fnum

End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable
   '
    Print #fnum, sLine
End Sub

Private Sub PR_PutNumberAssign(sVariableName As String, nAssignedNumber As Variant)

   'Procedure to put a number assignment
   'Adds a semi-colon

    Print #fnum, sVariableName; "="; Str$(nAssignedNumber); ";"


End Sub

Private Sub PR_PutStringAssign(sVariableName As String, sAssignedString As Variant)
   'Procedure to put a string assignment
   'Encloses String in quotes and adds a semi-colon

    Print #fnum, sVariableName; "="; QQ; sAssignedString; QQ; ";"

End Sub

Private Sub PR_ResetFormLGLEGDIA()
   'Reset display of form to a clean state
    Dim ii As Integer

   'Reset Max Stretch
    labLeftMaxStr.Caption = ""
    labLeftMaxStr.Visible = False
    labLeftDisp(3).Visible = False

   'Clean Display grid
    For ii = 0 To 1     'First two rows only, used in this case
        
        grdLeftDisplay.Row = ii
        grdLeftDisplay.Col = 0
        grdLeftDisplay = ""
        grdLeftDisplay.Col = 1
        grdLeftDisplay = ""

    Next ii

   'Switch off MM text boxes
    For ii = 6 To 7
        txtLeftMM(ii).Text = ""
        txtLeftMM(ii).Visible = False
    Next ii

   '
    For ii = 0 To 2
        labLeftDisp(ii).Visible = False
    Next ii
    
End Sub

Private Sub PR_SaveLeg()
   'Procedure to create a macro to save the leg data
    If txtUidLeftLeg.Text <> "" Then
        PR_PutLine "HANDLE  hLeg;"
        Print #fnum, "hLeg = UID ("; QQ; "find"; QC; Val(txtUidLeftLeg.Text); ");"
    Else
        PR_PutLine "HANDLE  hLeg, hTitle, hLegTitle;"
        PR_PutLine "XY xyTitleOrigin, xyTitleScale;"
        PR_PutLine "ANGLE aTitleAngle;"
        Print #fnum, "hTitle = UID ("; QQ; "find"; QC; Val(txtUidTitle.Text); ");"
        PR_PutStringAssign "sLeg", txtLeg

        Print #fnum, "SetSymbolLibrary( sPathJOBST + "; QQ; "\\JOBST.SLB"; QQ; ");"
        Print #fnum, "if ( !Symbol("; QQ; "find"; QCQ; "legleg"; QQ; "))"
        Print #fnum, "Exit(%cancel, "; QQ; "Can't find >legleg< symbol to insert\nCheck your installation, that JOBST.SLB exists"; QQ; ");"

        Print #fnum, "if ( hTitle )"
        Print #fnum, "GetGeometry( hTitle, &sTitleName, &xyTitleOrigin, &xyTitleScale, &aTitleAngle);"
        Print #fnum, "else"
        Print #fnum, "Exit(%cancel,"; QQ; "Can't find > mainpatientdetails <, Use TITLE to insert Patient Data"; QQ; ");"
        Print #fnum, "Execute("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        
        If txtUidLegTitle = "" Then
            Print #fnum, "if ( !Symbol("; QQ; "find"; QCQ; "legcommon"; QQ; "))"
            Print #fnum, "Exit(%cancel, "; QQ; "Can't find > legcommon < symbol to insert\nCheck your installation, that JOBST.SLB exists"; QQ; ");"
            Print #fnum, "hLegTitle = AddEntity("; QQ; "symbol"; QCQ; "legcommon"; QC; "xyTitleOrigin);"
            Print #fnum, "SetDBData( hLegTitle"; CQ; "Fabric"; QCQ; txtFabric; QQ; ");"
            Print #fnum, "SetDBData( hLegTitle"; CQ; "fileno"; QCQ; txtFileNo; QQ; ");"
           'Run macro to setup DB fields
            Print #fnum, "@"; g_sPathJOBST; "\LEG\LGFIELDS.D;"

        End If
    
        Print #fnum, "if (StringCompare("; QQ; "Left"; QC; "sLeg))"
        Print #fnum, "xyTitleOrigin.x = xyTitleOrigin.x + 1.5;"
        Print #fnum, "else"
        Print #fnum, "xyTitleOrigin.x = xyTitleOrigin.x + 3;"

        Print #fnum, "hLeg = AddEntity("; QQ; "symbol"; QCQ; "legleg"; QC; "xyTitleOrigin);"
        Print #fnum, "SetDBData( hLeg, "; QQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"

    End If

    Print #fnum, "if (!hLeg)Exit(%cancel,"; QQ; "Can't find LEGBOX to Update"; QQ; ");"
    
    Print #fnum, "SetDBData( hLeg, "; QQ; "Leg"; QCQ; txtLeg.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "TapeLengthsPt1"; QCQ; Mid$(txtLeftLengths.Text, 1, 60); QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "TapeLengthsPt2"; QCQ; Mid$(txtLeftLengths.Text, 61, 60); QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "Pressure"; QCQ; txtLeftTemplate.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "FootLength"; QCQ; txtFootLength.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "Fabric"; QCQ; txtFabric.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "FootPleat1"; QCQ; txtFootPleat1.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "TopLegPleat1"; QCQ; txtTopLegPleat1.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "FootPleat2"; QCQ; txtFootPleat2.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "TopLegPleat2"; QCQ; txtTopLegPleat2.Text; QQ; ");"
    
    Print #fnum, "SetDBData( hLeg, "; QQ; "ToeStyle"; QCQ; txtToeStyle.Text; QQ; ");"
    
    Print #fnum, "SetDBData( hLeg, "; QQ; "Anklet"; QCQ; txtAnklet.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "ThighLength"; QCQ; txtThighLength.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "KneeLength"; QCQ; txtKneeLength.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "ThighBand"; QCQ; txtThighBandAK.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "ThighBandBK"; QCQ; txtThighBandBK.Text; QQ; ");"
    Print #fnum, "SetDBData( hLeg, "; QQ; "KneeBand"; QCQ; txtKneeBand.Text; QQ; ");"

End Sub

Private Sub PR_SetFirstTape(iNewTape As Integer)

    If iNewTape >= 0 And iNewTape <= 29 Then
        If g_iStyleFirstTape >= 0 And g_iStyleFirstTape <= 29 Then
            Label1(g_iStyleFirstTape).ForeColor = RGB(0, 0, 0)
        End If
        g_iStyleFirstTape = iNewTape
        Label1(iNewTape).ForeColor = RGB(255, 0, 0)
        txtFirstTape.Text = LTrim$(Mid$(g_sTextList, (iNewTape * 3) + 1, 3))
    End If
    
End Sub

Private Sub PR_SetLastTape(iNewTape As Integer)

    If iNewTape >= 0 And iNewTape <= 29 Then
        If g_iStyleLastTape >= 0 And g_iStyleLastTape <= 29 Then
            Label1(g_iStyleLastTape).ForeColor = RGB(0, 0, 0)
        End If
        g_iStyleLastTape = iNewTape
        Label1(iNewTape).ForeColor = RGB(255, 0, 0)
        txtLastTape.Text = LTrim$(Mid$(g_sTextList, (iNewTape * 3) + 1, 3))
    End If

End Sub

Private Sub PR_SetTextData(nHoriz As Variant, nVert As Variant, nHt As Variant, nAspect As Variant, nFont As Variant)
   'To the DRAFIX macro file (given by the global fNum).
   'Write the syntax to set the TEXT default attributes, these are
   'based on the values in the arguments.  Where the value is -ve then this
   'attribute is not set.
   'where :-
   '    nHoriz      Horizontal justification (1=Left, 2=Cen, 4=Right)
   '    nVert       Verticalal justification (8=Top, 16=Cen, 32=Bottom)
   '    nHt         Text height
   '    nAspect     Text aspect ratio (heigth/width)
   '    nFont       Text font (0 to 18)
   '
   'N.B. No checking is done on the values given
   '
   'Note:-
   '    fNum, CC, QQ, NL, g_nCurrTextHt, g_nCurrTextAspect,
   '    g_nCurrTextHorizJust, g_nCurrTextVertJust, g_nCurrTextFont
   '    are globals initialised by FN_Open
   '

    If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then
    Print #fnum, "SetData("; QQ; "TextHorzJust"; QC; nHoriz; ");"
    g_nCurrTextHorizJust = nHoriz
    End If

    If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
    Print #fnum, "SetData("; QQ; "TextVertJust"; QC; nVert; ");"
    g_nCurrTextVertJust = nVert
    End If

    If nHt >= 0 And g_nCurrTextHt <> nHt Then
    Print #fnum, "SetData("; QQ; "TextHeight"; QC; nHt; ");"
    g_nCurrTextHt = nHt
    End If

    If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
    Print #fnum, "SetData("; QQ; "TextAspect"; QC; nAspect; ");"
    g_nCurrTextAspect = nAspect
    End If

    If nFont >= 0 And g_nCurrTextFont <> nFont Then
    Print #fnum, "SetData("; QQ; "TextFont"; QC; nFont; ");"
    g_nCurrTextFont = nFont
    End If


End Sub

Private Sub PR_StripBadChar(Text_Box_Name As Control)
   'Strip the excess characters from the text boxes.
   'This is due to a DRAFIX bug with DDE Poke.
    Dim sString As String
    Dim iLength As Integer, ii As Integer
    sString = Text_Box_Name.Text
    For ii = 1 To 2
        iLength = Len(sString)
        If iLength = 0 Then Exit For
        Select Case Asc(Right$(sString, 1))
            Case 32 To 126, 160 To 255
                Exit For
            Case Else
                sString = Left$(sString, iLength - 1)
        End Select
    Next ii
    Text_Box_Name.Text = sString
End Sub

Private Sub ShiftTapesDown_Click()
   'The effect of a single click on this button
   'is to shift all of the tapes in the direction of the
   'arrow
    Dim ii As Integer
    Dim nValue As Double

   'Check that last tape is empty so that the previous
   'tape can be shifted into it
    If Len(txtLeft(29).Text) > 0 Then
       'Beep and exit function
        Beep
        Exit Sub
    End If

   'Shift all tapes down 1
    For ii = 29 To 1 Step -1
        txtLeft(ii) = txtLeft(ii - 1)
        nValue = fnDisplaytoInches(Val(txtLeft(ii - 1)))
        grdLeftInches.Row = ii
        grdLeftInches = fnInchesToText(nValue)
    Next ii

   'Initiate Heel Change (if Required)
    If Val(txtLeft(5).Text) > 0 And g_iLtAnkle = 0 Then PR_HeelChange "Left"
    If Val(txtLeft(5).Text) = 0 And g_iLtAnkle <> 0 Then PR_HeelChange "Left"
    
   'For Thigh and Knee Bands move style first and last tapes
    If g_iLegStyle = 3 Or g_iLegStyle = 4 Then
        PR_SetFirstTape g_iStyleFirstTape + 1
        PR_SetLastTape g_iStyleLastTape + 1
    End If

   'Clean first position
    txtLeft(0) = ""
    grdLeftInches.Row = 0
    grdLeftInches = ""

End Sub

Private Sub ShiftTapesUp_Click()
   'The effect of a single click on this button
   'is to shift all of the tapes in the direction of the
   'arrow
    Dim ii As Integer
    Dim nValue As Double

   'Check that first tape is empty so that the following
   'tape can be shifted into it
    If Len(txtLeft(0).Text) > 0 Then
       'Beep and exit Procedure
        Beep
        Exit Sub
    End If

   'Shift all tapes up 1
    For ii = 1 To 29
        txtLeft(ii - 1) = txtLeft(ii)
        nValue = fnDisplaytoInches(Val(txtLeft(ii)))
        grdLeftInches.Row = ii - 1
        grdLeftInches = fnInchesToText(nValue)
    Next ii

   'Initiate Heel Change (if Required)
    If Val(txtLeft(5).Text) > 0 And g_iLtAnkle = 0 Then PR_HeelChange "Left"
    If Val(txtLeft(5).Text) = 0 And g_iLtAnkle <> 0 Then PR_HeelChange "Left"
    
   'For Thigh and Knee Bands move style first and last tapes
    If g_iLegStyle = 3 Or g_iLegStyle = 4 Then
        If g_iLtAnkle = 0 Then
            PR_SetFirstTape g_iStyleFirstTape - 1
        Else
            If g_iStyleFirstTape - 1 >= g_iLtAnkle Then
                PR_SetFirstTape g_iStyleFirstTape - 1
            Else
                PR_SetFirstTape g_iLtAnkle
            End If
        End If
        PR_SetLastTape g_iStyleLastTape - 1
    End If

   'Clean last position
    txtLeft(29) = ""
    grdLeftInches.Row = 29
    grdLeftInches = ""
End Sub

Private Sub spnFirstTape_DownClick()

    Dim ii As Integer
    If g_iStyleFirstTape <= 0 Or g_iStyleFirstTape - 1 < g_iFirstTape Then
        Beep
        Exit Sub
    End If

    PR_SetFirstTape g_iStyleFirstTape - 1

End Sub

Private Sub spnFirstTape_UpClick()
    Dim ii As Integer

    If g_iStyleFirstTape >= 28 Or g_iStyleFirstTape + 1 >= g_iStyleLastTape Then
        Beep
        Exit Sub
    End If

    PR_SetFirstTape g_iStyleFirstTape + 1

End Sub

Private Sub spnLastTape_DownClick()
    Dim ii As Integer

    If g_iStyleLastTape <= 1 Or g_iStyleLastTape - 1 <= g_iStyleFirstTape Then
        Beep
        Exit Sub
    End If

    PR_SetLastTape g_iStyleLastTape - 1

End Sub

Private Sub spnLastTape_UpClick()

    Dim ii As Integer
    
    If g_iStyleLastTape >= 29 Or g_iStyleLastTape + 1 > g_iLastTape Then
        Beep
        Exit Sub
    End If
    
    PR_SetLastTape g_iStyleLastTape + 1

End Sub

Private Sub Tab_Click()
   'Allows the user to use enter as a tab

    SendKeys "{TAB}", False

End Sub

Private Sub Timer1_Timer()
    'Timeout timer
    'If DRAFIX to VB DDE link fails then time out
    'after 6 seconds
     End
End Sub

Private Sub txtFootLength_GotFocus()
    Select_Text_In_Box txtFootLength
End Sub

Private Sub txtFootLength_LostFocus()
    Validate_And_Display_Text_In_Box txtFootLength, labFootLengthInInches, -1
End Sub

Private Sub txtLeft_GotFocus(Index As Integer)
    Select_Text_In_Box txtLeft(Index)
End Sub

Private Sub txtLeft_KeyPress(Index As Integer, KeyAscii As Integer)
   'Use enter to act as a TAB
  '  If KeyAscii = NEWLINE Then txtLeft((Index + 1) Mod 30).SetFocus

End Sub

Private Sub txtLeft_LostFocus(Index As Integer)
    Validate_And_Display_Text_In_Box txtLeft(Index), grdLeftInches, Index
    If Index = g_iLtAnkle Then PR_FigureLeftAnkle
    If Index = 5 Then PR_HeelChange "Left"
End Sub

Private Sub txtLeftMM_LostFocus(Index As Integer)
    If Index = g_iLtAnkle Then
        PR_FigureLeftAnkle
    Else
        PR_FigureLeftTape Index
    End If
End Sub

Private Sub Update_DDE_Text_Boxes()
   'Called from OK_Click and Draw
   'Update the text boxes used for DDE transfer
    Dim ii As Integer, sString As String, sJustifiedString As String
    Dim sFabricClass As String
    Dim iRed As Integer, iStr  As Integer, iMM As Integer
    Dim sRed As String, sStr As String, sMM As String
    Dim nValue As Single

    txtFabric = cboFabric.List(cboFabric.ListIndex)
    txtToeStyle = cboToeStyle.List(cboToeStyle.ListIndex)

   'LEFT Leg
   'Assume that data has been validated earlier
    
   'Set initial values
    txtLeftLengths = ""
    txtLeftRed = ""
    txtLeftStr = ""
    txtLeftMMs = ""
    g_iLastTape = -1
    g_iFirstTape = -1
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
        If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii + 1
        If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii
    Next ii

    If g_iLastTape < 0 Then g_iLastTape = 30

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

    g_iStyleLastTape = g_iStyleLastTape + 1
    g_iStyleFirstTape = g_iStyleFirstTape + 1

    g_sStyleString = FN_BuildStyleString()

    Select Case g_iLegStyle
        Case 0
            txtAnklet.Text = g_sStyleString
        Case 1
            txtKneeLength.Text = g_sStyleString
        Case 2
            txtThighLength.Text = g_sStyleString
        Case 3
            txtKneeBand.Text = g_sStyleString
        Case 4
            txtThighBandAK.Text = g_sStyleString
        Case 5
            txtThighBandBK.Text = g_sStyleString
    End Select

End Sub

Private Function Validate_Data() As Integer
   'Called from OK_Click
    
    Dim sError As String, NL As String, sTextList As String
    Dim sLeftError As String, sRightError As String
    Dim ii As Integer, nFirstTape As Integer, nLastTape As Integer
    Dim iError As Integer, nLargeHeelX As Double
    Dim nHeelLength As Double, nFootLength As Double, nTapeX As Double
    Dim nValue As Double

    NL = Chr$(10) 'new line

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
            sError = sError + "Missing Tape length - " & LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3)) + NL
        End If
    Next ii
   
   'Check that a minimum of 3 tapes are given
    If (nLastTape - nFirstTape) < 2 Then
        sError = sError + "Minimum of 3 Tapes must be given." + NL
    End If

   'Check if -3 tape exists for Large Heel
    If g_iLtAnkle = 7 And Val(txtLeft(3).Text) = 0 And g_iLegStyle < 3 Then
        sError = sError + "As the Heel is 9 inches and over." + NL
        sError = sError + "and there is no -3 tape the foot will not draw properly " + NL
    End If

   'Check if -1 1/2 tape exists for Small Heel
    If g_iLtAnkle = 6 And Val(txtLeft(4).Text) = 0 And g_iLegStyle < 3 Then
        sError = sError + "As the Heel is smaller than 9 inches." + NL
        sError = sError + "and there is no -1 1/2 tape the foot will not draw properly " + NL
    End If
   
   'Check that figuring has been done at the ankle
   'Does not apply to Anklets or Bands
    If g_iLegStyle > 0 And g_iLegStyle < 3 And g_iLtAnkle <> 0 And (g_iLtStretch(g_iLtAnkle) = 0 Or g_iLtRed(g_iLtAnkle) = 0 Or g_iLtMM(g_iLtAnkle) = 0) Then
        sError = sError + "The Ankle has not been figured!." + NL
        sError = sError + "Figure the Ankle or use CANCEL to exit" + NL
    End If


   'Toe style need only be given if a heel tape is given, Heel is at tape 5
    If (g_iFirstTape < 5) And (cboToeStyle.ListIndex <= 0) And g_iLegStyle < 3 Then
        sError = sError + "Missing Toe Style." + NL
    End If

   'Check that a foot length has been given for self enclosed
    If cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" And Val(txtFootLength.Text) = 0 Then
        sError = sError + "A foot length must be given for Self Enclosed toes." + NL
    End If

   'Check that a leg style has been chosen
    If g_iLegStyle < 0 Then
        sError = sError + "A Leg Style has not been chosen!." + NL
        sError = sError + "Choose a Style or use CANCEL to exit" + NL
    End If
    
   'General check
    If cboFabric.List(cboFabric.ListIndex) = "" Then
        sError = sError + "A Fabric has not been chosen!." + NL
    End If

   'General check for atemplate
    If cboLeftTemplate.ListIndex = -1 Then
        sError = sError + "A template has not been chosen!." + NL
    End If


   'Display Error message (if required) and return
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        Validate_Data = False
        Exit Function
    Else
        Validate_Data = True
    End If

   'For the combo boxes check if a toe selection has been made
   'and a heel tape is not given.
   'This is a none fatal error, the data will be ammended and
   'an information message given
   'will be given
    
    sError = ""
    If nFirstTape > 5 Then
        If cboToeStyle.ListIndex > 0 Then
            sError = sError + "Toe style, " + cboToeStyle.List(cboToeStyle.ListIndex) + NL
        End If
        cboToeStyle.ListIndex = 0
    End If

   'Display Warning message (if required)
    If Len(sError) > 0 Then
        sError = "The TOE style given can't apply. As there is no Heel." + NL
        sError = sError + "Removing TOE style to reflect this."
        MsgBox sError, 64, "Warning - Errors in Data"
    End If



   'Yes / No type errors
   'In this case we warn the user that there is a problem!
   'They can continue or they can return to the dialog to make changes

   'Initialize error variables
    sError = ""
    iError = False
   
   'Make a stab at seeing if the foot length will be long enough.
   'For Toe = "Self enclosed" or Toe = "Soft enclosed" and a foot length is given
   'This checks that the given foot length is longer than the distance
   'from the heel to the -3 tape at the seam (-1 1/2 for heels less than 9 ")
   'it is a rough giude only.
    If g_nLtLastHeel > 0 And (cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" Or (cboToeStyle.List(cboToeStyle.ListIndex) = "Soft Enclosed" And Val(txtFootLength.Text) <> 0)) Then
              
       'Set checking parameters based on fabric and template
        If g_POWERNET = True Then
            nLargeHeelX = 2.71875
           'Scale heel to the 0 tape of a 30mmHg template
           'This is the worst case for POWERNET fabric
            nHeelLength = g_nLtLastHeel * 0.402
           'Reduce heel by 3
            nHeelLength = nHeelLength + (3 * 0.05025)
        Else
            Select Case cboLeftTemplate.ListIndex
            Case 0  ' 13 DS
                nLargeHeelX = 2.73
            Case 1  ' 09 DS
                nLargeHeelX = 2.88
            End Select
           'Scale heel to the 0 tape of a JOBSTEX template
            nHeelLength = g_nLtLastHeel * 0.391
           'Reduce heel by 3
            nHeelLength = nHeelLength + (3 * 0.048875)
        End If

        If g_nLtLastHeel >= 9 Then
            nTapeX = nLargeHeelX 'Large heel, distance from heel to -3 tape
        Else
            nTapeX = 1.5         'Small heel, distance from heel to -1 1/2 tape
        End If

        nFootLength = fnDisplaytoInches(Val(txtFootLength.Text))
         
        nValue = Sqr(nHeelLength ^ 2 + nTapeX ^ 2)

        If nValue > nFootLength Then
            iError = True
            sError = sError + "The given foot length may be too small to correctly" + NL
            sError = sError + "position the toe!" + NL + NL
            sError = sError + "Do you wish to continue anyway ?"
        End If
    End If


   'Display Error message (if required) and return
   'These are non-fatal errors
    If iError = True Then
        iError = MsgBox(sError, 36, "Warning, Problems with data")
        If iError = IDYES Then
            Validate_Data = True
        Else
            Validate_Data = False
        End If
    Else
        Validate_Data = True
    End If

End Function


VERSION 5.00
Begin VB.Form bdlegdia 
   Caption         =   "Waist Height Leg Data"
   ClientHeight    =   8376
   ClientLeft      =   5220
   ClientTop       =   1476
   ClientWidth     =   6960
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
   LinkTopic       =   "bdlegdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8376
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdExtendLegTapes 
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdShiftTapesDown 
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdShiftTapesUp 
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   7140
      TabIndex        =   72
      Top             =   5490
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Left            =   7260
      Top             =   4890
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   147
      Top             =   4320
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   3420
      TabIndex        =   101
      Top             =   120
      Width           =   3435
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   107
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   106
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   105
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2520
         TabIndex        =   104
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2520
         TabIndex        =   103
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   102
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   360
         TabIndex        =   78
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "File No:"
         Height          =   195
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   1980
         TabIndex        =   111
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Units:"
         Height          =   255
         Left            =   1980
         TabIndex        =   110
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox txtLegStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   97
      Top             =   3900
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtFirstTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   94
      Top             =   3180
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLastTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   93
      Top             =   3540
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame2 
      Caption         =   "Leg"
      Height          =   615
      Left            =   3420
      TabIndex        =   87
      Top             =   2370
      Width           =   3435
      Begin VB.OptionButton optRightLeg 
         Caption         =   "Right"
         Height          =   195
         Left            =   2160
         TabIndex        =   92
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optLeftLeg 
         Caption         =   "Left"
         Height          =   195
         Left            =   660
         TabIndex        =   89
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txtAnkleTape 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   86
      Top             =   2460
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtToeStyle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   85
      Top             =   2820
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtTapeLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   81
      Top             =   2100
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   79
      Top             =   1740
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   74
      Top             =   1380
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8220
      TabIndex        =   73
      Top             =   1020
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Styles / Other Measurements"
      Height          =   1815
      Left            =   3420
      TabIndex        =   66
      Top             =   3270
      Width           =   3435
      Begin VB.OptionButton optAboveKnee 
         Caption         =   "Above Knee"
         Height          =   255
         Left            =   1740
         TabIndex        =   68
         Top             =   1320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optBelowKnee 
         Caption         =   "Below Knee"
         Height          =   255
         Left            =   180
         TabIndex        =   69
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtFootLength 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   84
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cboToeStyle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   420
         Width           =   1995
      End
      Begin VB.Label labFootLengthInInches 
         Height          =   255
         Left            =   1800
         TabIndex        =   146
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Toe Style:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "FootLength:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   900
         Width           =   1035
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3780
      TabIndex        =   80
      Top             =   7860
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   77
      Top             =   7860
      Width           =   1215
   End
   Begin VB.TextBox txtTopLegPleat1 
      Height          =   315
      Left            =   240
      MaxLength       =   7
      TabIndex        =   62
      Top             =   7860
      Width           =   615
   End
   Begin VB.TextBox txtTopLegPleat2 
      Height          =   315
      Left            =   240
      MaxLength       =   7
      TabIndex        =   64
      Top             =   7500
      Width           =   615
   End
   Begin VB.TextBox txtFootPleat2 
      Height          =   315
      Left            =   240
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtFootPleat1 
      Height          =   315
      Left            =   240
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   29
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   60
      Top             =   8010
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   28
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   58
      Top             =   7740
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   27
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   56
      Top             =   7470
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   26
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   54
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   25
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   52
      Top             =   6930
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   24
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   50
      Top             =   6660
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   23
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   48
      Top             =   6390
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   22
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   46
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   21
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   44
      Top             =   5850
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   20
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   42
      Top             =   5580
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   19
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   40
      Top             =   5310
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   18
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   38
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   17
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   36
      Top             =   4770
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   16
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   34
      Top             =   4500
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   15
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   32
      Top             =   4230
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   14
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   30
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   13
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   28
      Top             =   3690
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   12
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   26
      Top             =   3420
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   11
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   24
      Top             =   3150
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   10
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   22
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   9
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   20
      Top             =   2610
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   8
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2340
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   7
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   16
      Top             =   2070
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   6
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   5
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1530
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   8
      Top             =   990
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   4
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   3
      Top             =   180
      Width           =   615
   End
   Begin VB.Label labData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Misc. Data"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   143
      Top             =   4380
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label20 
      Caption         =   "Extend"
      Height          =   195
      Left            =   240
      TabIndex        =   145
      Top             =   5820
      Width           =   675
   End
   Begin VB.Label Label11 
      Caption         =   "Leg Tapes"
      Height          =   195
      Left            =   120
      TabIndex        =   144
      Top             =   6060
      Width           =   915
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   29
      Left            =   2400
      TabIndex        =   142
      Top             =   8040
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   28
      Left            =   2400
      TabIndex        =   141
      Top             =   7770
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   27
      Left            =   2400
      TabIndex        =   140
      Top             =   7500
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   26
      Left            =   2400
      TabIndex        =   139
      Top             =   7230
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   25
      Left            =   2400
      TabIndex        =   138
      Top             =   6960
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   24
      Left            =   2400
      TabIndex        =   137
      Top             =   6690
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   23
      Left            =   2400
      TabIndex        =   136
      Top             =   6420
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   22
      Left            =   2400
      TabIndex        =   135
      Top             =   6150
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   21
      Left            =   2400
      TabIndex        =   134
      Top             =   5880
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   20
      Left            =   2400
      TabIndex        =   133
      Top             =   5610
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   19
      Left            =   2400
      TabIndex        =   132
      Top             =   5340
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   18
      Left            =   2400
      TabIndex        =   131
      Top             =   5070
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   130
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   129
      Top             =   4530
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   15
      Left            =   2400
      TabIndex        =   128
      Top             =   4260
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   127
      Top             =   3990
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   126
      Top             =   3720
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   125
      Top             =   3450
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   124
      Top             =   3180
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   123
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   122
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   121
      Top             =   2370
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   120
      Top             =   2100
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   119
      Top             =   1830
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   118
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   117
      Top             =   1290
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   116
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   115
      Top             =   750
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   114
      Top             =   480
      Width           =   585
   End
   Begin VB.Label labTapeInInches 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   113
      Top             =   210
      Width           =   585
   End
   Begin VB.Label Label22 
      Caption         =   "tapes"
      Height          =   195
      Left            =   360
      TabIndex        =   90
      Top             =   4320
      Width           =   555
   End
   Begin VB.Label Label21 
      Caption         =   "move"
      Height          =   195
      Left            =   360
      TabIndex        =   100
      Top             =   4080
      Width           =   555
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
      Left            =   7140
      TabIndex        =   99
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Style"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   98
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lable14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FirstTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   96
      Top             =   3240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LastTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   95
      Top             =   3600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AnkleTape"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   91
      Top             =   2520
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ToeStyle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   88
      Top             =   2880
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tape Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   83
      Top             =   2160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   82
      Top             =   1800
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   76
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Title UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   75
      Top             =   1080
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Pleats:"
      Height          =   195
      Left            =   240
      TabIndex        =   65
      Top             =   7200
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Pleats:"
      Height          =   195
      Left            =   240
      TabIndex        =   63
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "36"
      Height          =   195
      Index           =   29
      Left            =   1140
      TabIndex        =   61
      Top             =   8070
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "34½"
      Height          =   195
      Index           =   28
      Left            =   1140
      TabIndex        =   59
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "33"
      Height          =   195
      Index           =   27
      Left            =   1140
      TabIndex        =   57
      Top             =   7530
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "31½"
      Height          =   195
      Index           =   26
      Left            =   1140
      TabIndex        =   55
      Top             =   7260
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "30"
      Height          =   195
      Index           =   25
      Left            =   1140
      TabIndex        =   53
      Top             =   6990
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "28½"
      Height          =   195
      Index           =   24
      Left            =   1140
      TabIndex        =   51
      Top             =   6720
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "27"
      Height          =   195
      Index           =   23
      Left            =   1140
      TabIndex        =   49
      Top             =   6450
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "25½"
      Height          =   195
      Index           =   22
      Left            =   1140
      TabIndex        =   47
      Top             =   6180
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "24"
      Height          =   195
      Index           =   21
      Left            =   1140
      TabIndex        =   45
      Top             =   5910
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "22½"
      Height          =   195
      Index           =   20
      Left            =   1140
      TabIndex        =   43
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "21"
      Height          =   195
      Index           =   19
      Left            =   1140
      TabIndex        =   41
      Top             =   5370
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "19½"
      Height          =   195
      Index           =   18
      Left            =   1140
      TabIndex        =   39
      Top             =   5100
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "18"
      Height          =   195
      Index           =   17
      Left            =   1140
      TabIndex        =   37
      Top             =   4830
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "16½"
      Height          =   195
      Index           =   16
      Left            =   1140
      TabIndex        =   35
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "15"
      Height          =   195
      Index           =   15
      Left            =   1140
      TabIndex        =   33
      Top             =   4290
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "13½"
      Height          =   195
      Index           =   14
      Left            =   1140
      TabIndex        =   31
      Top             =   4020
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "12"
      Height          =   195
      Index           =   13
      Left            =   1140
      TabIndex        =   29
      Top             =   3750
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "10½"
      Height          =   195
      Index           =   12
      Left            =   1140
      TabIndex        =   27
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "9"
      Height          =   195
      Index           =   11
      Left            =   1140
      TabIndex        =   25
      Top             =   3210
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "7½"
      Height          =   195
      Index           =   10
      Left            =   1140
      TabIndex        =   23
      Top             =   2940
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   195
      Index           =   9
      Left            =   1140
      TabIndex        =   21
      Top             =   2670
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4½"
      Height          =   195
      Index           =   8
      Left            =   1140
      TabIndex        =   19
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   195
      Index           =   7
      Left            =   1140
      TabIndex        =   17
      Top             =   2130
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1½"
      Height          =   195
      Index           =   6
      Left            =   1140
      TabIndex        =   15
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   5
      Left            =   1140
      TabIndex        =   13
      Top             =   1590
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1½"
      Height          =   195
      Index           =   4
      Left            =   1140
      TabIndex        =   11
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   195
      Index           =   3
      Left            =   1140
      TabIndex        =   9
      Top             =   1050
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4½"
      Height          =   195
      Index           =   2
      Left            =   1140
      TabIndex        =   7
      Top             =   780
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   195
      Index           =   1
      Left            =   1140
      TabIndex        =   5
      Top             =   510
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "7½"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      Top             =   240
      Width           =   315
   End
End
Attribute VB_Name = "bdlegdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   '* Windows API Functions Declarations
 '   Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
 '   Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
 '   Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
 '   Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
 '   Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
   
   'Constanst used by GetWindow
 '   Const GW_CHILD = 5
 '   Const GW_HWNDFIRST = 0
 '   Const GW_HWNDLAST = 1
 '   Const GW_HWNDNEXT = 2
 '   Const GW_HWNDPREV = 3
 '   Const GW_OWNER = 4

   'Constants
    Const HEEL_TOL = 9 '9" heel

   'MsgBox constants
  '  Const IDYES = 6
  '  Const IDNO = 7
  '  Const IDCANCEL = 2

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
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height Leg Data"
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

Private Sub cmdExtendLegTapes_Click()
   'Takes the last tape given and add the required value
   'GOP - 01-02/12, 6.2.1
   '
    Dim ii As Integer
    Dim nValue As Double

   'Locate last tape
    For ii = 29 To 0 Step -1
        If Val(Text1(ii)) > 0 Then Exit For
    Next ii

   'Check that there are some tapes.
   'Check that there is room to add a new tape
    If ii = 0 Or ii = 29 Then
        Beep
        Exit Sub
    End If

   'Convert given value to inches
    nValue = fnDisplayToInches(Val(Text1(ii)))

   'Extend tape
    If txtSex = "Male" Then
        nValue = nValue + 0.5
    Else
        nValue = nValue + 1
    End If

    Text1(ii + 1) = fnInchesToDisplay(nValue)
    labTapeInInches(ii + 1).Caption = fnInchesToText(nValue)

End Sub

Private Sub cmdShiftTapesDown_Click()
   'The effect of a single click on this button
   'is to shift all of the tapes in the direction of the
   'arrow
    Dim ii As Integer
    Dim nValue As Double

   'Check that last tape is empty so that the previous
   'tape can be shifted into it
    If Len(Text1(29).Text) > 0 Then
       'Beep and exit function
        Beep
        Exit Sub
    End If

   'Shift all tapes down 1
    For ii = 29 To 1 Step -1
        Text1(ii) = Text1(ii - 1)
        nValue = fnDisplayToInches(Val(Text1(ii - 1)))
        labTapeInInches(ii).Caption = fnInchesToText(nValue)
    Next ii

   'Clean first position
    Text1(0) = ""
    labTapeInInches(0).Caption = ""

   'Check for a heel value
    If Val(Text1(5).Text) > 0 Then
        optBelowKnee.Value = False
        optAboveKnee.Value = False
        optBelowKnee.Enabled = False
        optAboveKnee.Enabled = False
        Else
        optBelowKnee.Enabled = True
        optAboveKnee.Enabled = True
    End If
End Sub

Private Sub cmdShiftTapesUp_Click()
   'The effect of a single click on this button
   'is to shift all of the tapes in the direction of the
   'arrow
    Dim ii As Integer
    Dim nValue As Double

   'Check that first tape is empty so that the following
   'tape can be shifted into it
    If Len(Text1(0).Text) > 0 Then
        'Beep and exit function
        Beep
        Exit Sub
    End If

   'Shift all tapes up 1
    For ii = 1 To 29
        Text1(ii - 1) = Text1(ii)
        nValue = fnDisplayToInches(Val(Text1(ii)))
        labTapeInInches(ii - 1).Caption = fnInchesToText(nValue)
    Next ii
   'Clean last position
    Text1(29) = ""
    labTapeInInches(29).Caption = ""

   'Check for a heel value
    If Val(Text1(5).Text) > 0 Then
        optBelowKnee.Value = False
        optAboveKnee.Value = False
        optBelowKnee.Enabled = False
        optAboveKnee.Enabled = False
        Else
        optBelowKnee.Enabled = True
        optAboveKnee.Enabled = True
    End If

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
    Print #fNum, "//by Visual Basic, Waist Height Leg"
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
    For ii = 0 To 29
        sString = sString & Str$(Val(Text1(ii).Text))
    Next ii
    
    sString = sString & txtFootPleat1 & txtFootPleat2
    
    sString = sString & txtTopLegPleat1 & txtTopLegPleat2 & txtFootLength
    
    sString = sString & cboToeStyle.Text & Str$(optAboveKnee.Value) & Str$(optBelowKnee.Value)
    
    FN_ValuesString = sString

End Function


Private Sub Form_LinkClose()

   Dim nValue As Double
   Dim nAge, ii As Integer

  'Disable Time out
   Timer1.Enabled = False
  
  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

  'Foot Length
   nValue = fnDisplayToInches(Val(txtFootLength.Text))
   If nValue > 0 Then labFootLengthInInches.Caption = fnInchesToText(nValue)
  
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
        optLeftLeg.Value = True
        optRightLeg.Enabled = False
        bdlegdia.Caption = "Body Suit Leg Data - Left"
   Else
        optRightLeg.Value = True
        optLeftLeg.Enabled = False
        bdlegdia.Caption = "Body Suit Leg Data - Right"
   End If
  
  'Update tape boxes
   For ii = 0 To 29
        nValue = Val(Mid(txtTapeLengths.Text, (ii * 4) + 1, 4)) / 10
        If nValue > 0 Then
            Text1(ii) = nValue
            nValue = fnDisplayToInches(nValue)
            labTapeInInches(ii).Caption = fnInchesToText(nValue)
        End If
   Next ii

  'Store original Heel and Ankle tape values
  'Used on DDE_Update check if values have been modified
  'A heel length of 0 => footless

   g_nAnkleTape = 0
   g_nAnkleLen = 0
   g_nHeelLen = fnDisplayToInches(Val(Text1(5)))
   
   If (g_nHeelLen <> 0) And (g_nHeelLen < HEEL_TOL) Then
        g_nAnkleTape = 6    'Small heel Ankle at +1 1/2 tape
        g_nAnkleLen = fnDisplayToInches(Val(Text1(g_nAnkleTape)))
   Else
        g_nAnkleTape = 7    'Large heel Ankle at +3 tape
        g_nAnkleLen = fnDisplayToInches(Val(Text1(g_nAnkleTape)))
   End If
   
  'Set below and above knee option buttons
  'Disable both if a heel value is given
   optBelowKnee.Value = False
   optAboveKnee.Value = True       'default for body suit
   optBelowKnee.Enabled = False
   optAboveKnee.Enabled = False

   If g_nHeelLen = 0 Then
        optBelowKnee.Enabled = True
        optAboveKnee.Enabled = True
        Select Case Val(txtData.Text)
          Case -1
            optBelowKnee.Value = True
          Case 1
            optAboveKnee.Value = True
        End Select
   End If

  'Store initial values for use in Cancel_Click
   g_sChangeChecker = FN_ValuesString()

   Show
   Screen.MousePointer = 0    ' Change pointer to default.
   bdlegdia.SetFocus

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Hide
   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - bdlegdia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - bdlegdia.Height) / 2  ' Center form vertically.
   
    Set MainForm = bdlegdia
   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    g_sPathJOBST = fnPathJOBST()

    g_nUnitsFac = 1 'Default to inches

   'Start a timer that will ensure that the
   'the VB programme ends if the DRAFIX DDE
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
            MsgBox "Can't find a Drafix Drawing to update!", 16, "WAIST Height Leg Data"
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
    PR_PutLine "HANDLE hMPD, hLeg;"
   
    Update_DDE_Text_Boxes

    Dim sSymbol As String
    
    sSymbol = "waistleg"

    If txtUidLeg.Text = "" Then
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

       'Insert symbol
        PR_PutLine "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){"
        PR_PutLine "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));"
        If txtLeg.Text = "Left" Then
            PR_PutLine "  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 1.5,xyMPD_Origin.y );"
        Else
            PR_PutLine "  hLeg = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin.x + 3,xyMPD_Origin.y );"
        End If
        PR_PutLine "  }"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");"
    Else
       'Use existing symbol
        PR_PutLine "hLeg = UID (" & QQ & "find" & QC & Val(txtUidLeg.Text) & ");"
        PR_PutLine "if (!hLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

    End If

   'Update the BODY Box symbol
    PR_PutLine "SetDBData( hLeg" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "Leg" & QCQ & txtLeg.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid$(txtTapeLengths.Text, 1, 60) & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid$(txtTapeLengths.Text, 61) & QQ & ");"

    PR_PutLine "SetDBData( hLeg" & CQ & "AnkleTape" & QCQ & txtAnkleTape.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "FirstTape" & QCQ & txtFirstTape.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "LastTape" & QCQ & txtLastTape.Text & QQ & ");"
    
    PR_PutLine "SetDBData( hLeg" & CQ & "ToeStyle" & QCQ & txtToeStyle.Text & QQ & ");"

    
    PR_PutLine "SetDBData( hLeg" & CQ & "FootPleat1" & QCQ & txtFootPleat1.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "FootPleat2" & QCQ & txtFootPleat2.Text & QQ & ");"

    PR_PutLine "SetDBData( hLeg" & CQ & "TopLegPleat1" & QCQ & txtTopLegPleat1.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "TopLegPleat2" & QCQ & txtTopLegPleat2.Text & QQ & ");"

    PR_PutLine "SetDBData( hLeg" & CQ & "FootLength" & QCQ & txtFootLength.Text & QQ & ");"
    PR_PutLine "SetDBData( hLeg" & CQ & "Data" & QCQ & txtData.Text & QQ & ");"

    Close #fNum
   
End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable

    Print #fNum, sLine

End Sub

Private Sub Tab_Click()
   'Allows the user to use enter as a tab

    SendKeys "{TAB}", False

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Select_Text_In_Box Text1(Index)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim nValue As Double
    nValue = Validate_And_Display_Text_In_Box(Text1(Index), labTapeInInches(Index))
   
   'If this is for the heel tape ie Index = 5 then
   'Enable or disable the Above or Below option controls
    If nValue >= 0 And Index = 5 Then
        If nValue = 0 Then
           'No heel given Enable
            optBelowKnee.Enabled = True
            optAboveKnee.Enabled = True
        Else
           'Heel value given Disable controls
            optBelowKnee.Value = False
            optAboveKnee.Value = False
            optBelowKnee.Enabled = False
            optAboveKnee.Enabled = False
        End If
    End If

End Sub

Private Sub Timer1_Timer()
    'If this event is called then the
    'DDE link with DRAFIX has not closed
    'Therefor we time out after 6 seconds
     End

End Sub

Private Sub txtFootLength_GotFocus()
    Select_Text_In_Box txtFootLength
End Sub

Private Sub txtFootLength_LostFocus()
    Dim nValue As Double
    nValue = Validate_And_Display_Text_In_Box(txtFootLength, labFootLengthInInches)
End Sub

Private Sub txtFootPleat1_GotFocus()
    Select_Text_In_Box txtFootPleat1
End Sub

Private Sub txtFootPleat1_LostFocus()
    Validate_Text_In_Box txtFootPleat1
End Sub

Private Sub txtFootPleat2_GotFocus()
    Select_Text_In_Box txtFootPleat2
End Sub

Private Sub txtFootPleat2_LostFocus()
    Validate_Text_In_Box txtFootPleat2
End Sub

Private Sub txtTopLegPleat1_GotFocus()
    Select_Text_In_Box txtTopLegPleat1
End Sub

Private Sub txtTopLegPleat1_LostFocus()
    Validate_Text_In_Box txtTopLegPleat1
End Sub

Private Sub txtTopLegPleat2_GotFocus()
    Select_Text_In_Box txtTopLegPleat2
End Sub

Private Sub txtTopLegPleat2_LostFocus()
    Validate_Text_In_Box txtTopLegPleat2
End Sub

Private Sub Update_DDE_Text_Boxes()
   'Called from OK_Click and Cancel_Click
   'Update the text boxes used for DDE transfer
    Dim ii As Integer, sString As String, sRightJustified As String
    Dim sJustifiedString As String
    Dim nValue As Single
    Dim nAnkleTape As Single, nAnkleLen As Single, nHeelLen As Single

   'Combo Boxes
    txtToeStyle.Text = cboToeStyle.List(cboToeStyle.ListIndex)

   'Option Boxes
    If optLeftLeg.Value = True Then
        txtLeg = "Left"
    Else
        txtLeg = "Right"
    End If

    txtData = "0"
    If optBelowKnee.Value = True Then txtData = "-1"
    If optAboveKnee.Value = True Then txtData = "1"

   'Tape boxes
   'Assume that data has been validated earlier
    
   'Set initial values
    txtTapeLengths = ""
    txtFirstTape = -1
    txtLastTape = 30 'Take care of special case where last tape is text1(29)

    For ii = 0 To 29
        nValue = Val(Text1(ii).Text)
        If nValue <> 0 Then
            nValue = nValue * 10 'Shift decimal place to right
            sJustifiedString = String(4, " ")
            RSet sJustifiedString = Trim$(Str$(nValue))
        Else
            sJustifiedString = String(4, " ")
        End If

       'Tape values
        txtTapeLengths.Text = txtTapeLengths.Text & sJustifiedString
       
       'Set first and last tape (assumes no holes in data)
        If txtFirstTape < 0 And nValue > 0 Then txtFirstTape = ii + 1
        If txtLastTape = 30 And txtFirstTape > 0 And nValue = 0 Then txtLastTape = ii
    Next ii
   

   'Check if heel or ankle changed, If they have make the changes and warn
   'that configuring is required
    nAnkleTape = 0
    nAnkleLen = 0
    nHeelLen = fnDisplayToInches(Val(Text1(5)))
   
    If (nHeelLen <> 0) And (nHeelLen < HEEL_TOL) Then
        nAnkleTape = 6    'Small heel Ankle at +1 1/2 tape
        nAnkleLen = fnDisplayToInches(Val(Text1(nAnkleTape)))
    Else
        nAnkleTape = 7    'Large heel Ankle at +3 tape
        nAnkleLen = fnDisplayToInches(Val(Text1(nAnkleTape)))
    End If

   'Note:-
   '        The DRAFIX macros index tape lengths from 1 not 0 as VB does.
   '        Therefor nAnkleTape is incremented by 1
   '
    If (nHeelLen <> g_nHeelLen) Or (nAnkleLen <> g_nAnkleLen) Then
'        MsgBox "Figure Ankle for this leg.", 48    'Not required  for body
        If nHeelLen = 0 Then txtAnkleTape = "-1 0 0 0 0 0 0 0"
    End If

   'Set the txtAnkleTape to reflect the footless state
   'Only if it has not been set before
    If nHeelLen = 0 And txtAnkleTape = "" Then txtAnkleTape = "-1 0 0 0 0 0 0 0"

  
End Sub

Private Function Validate_Data() As Integer
   'Called from OK_Click
    
    Dim sError As String, NL As String, sTextList As String
    Dim ii As Integer, nFirstTape As Integer, nLastTape As Integer
    Dim iError As Integer
    Dim nHeelLength As Double, nFootLength As Double, nTapeX As Double
    Dim nValue As Double

    NL = Chr$(10) 'new line
    sTextList = "  7  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

   'Check Tape length data text boxes for holes (ie missing values)
   'Establish First and last tape
    For ii = 0 To 29 Step 1
        If Val(Text1(ii).Text) > 0 Then Exit For
    Next ii
    nFirstTape = ii

    For ii = 29 To 0 Step -1
        If Val(Text1(ii).Text) > 0 Then Exit For
    Next ii
    nLastTape = ii

    If nFirstTape = 30 Then
        sError = sError + "No Tape lengths given." + NL
    Else
        For ii = nFirstTape To nLastTape
            If Val(Text1(ii).Text) = 0 Then
            sError = sError + "Missing Tape length " & LTrim$(Mid$(sTextList, (ii * 3) + 1, 3)) + NL
            End If
        Next ii
    End If

   'Check combo boxes, not re-inforcement
   'NB - the first item in each combo box is a blank
   'Toe style need only be given if a heel tape is given, Heel is at tape 5
    If (nFirstTape < 5) And (cboToeStyle.ListIndex <= 0) Then
        sError = sError + "Missing Toe Style." + NL
    End If

   'Check Tape length data text boxes for holes (ie missing values)
   'Establish First and last tape
    For ii = 0 To 29 Step 1
        If Val(Text1(ii).Text) > 0 Then Exit For
    Next ii
    nFirstTape = ii

    For ii = 29 To 0 Step -1
        If Val(Text1(ii).Text) > 0 Then Exit For
    Next ii
    nLastTape = ii

    If nFirstTape = 29 Then
        sError = sError + "No Tape lengths given." + NL
    End If

   'Check that a minimum of 3 tapes are given
    If (nLastTape - nFirstTape) = 0 Then
        sError = sError + "More than 1 tape must be given." + NL
    End If

   'Check that a foot length has been given for self enclosed
'    If cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" And Val(txtFootLength.Text) = 0 Then
'        sError = sError + "A foot length must be given for Self Enclosed toes." + NL
'    End If

   'Get Heel Length if it applies
    If nFirstTape <= 5 Then
        nHeelLength = fnDisplayToInches(Val(Text1(5).Text))
    Else
        nHeelLength = 0
    End If

   'Check if either the "above" or "below" knee options are
   'given for a footless
'    If nHeelLength = 0 And optBelowKnee.Value = False And optAboveKnee.Value = False Then
'        sError = sError + "This is a footless style." + NL
'        sError = sError + "ABOVE or BELOW Knee must be selected." + NL
'    End If
   
   'Check if -3 tape exists for Large Heel
'    If nFirstTape <= 5 And nHeelLength >= 9 And Val(Text1(3)) = 0 Then
'        sError = sError + "As the Heel is 9 inches and over." + NL
'        sError = sError + "and there is no -3 tape the foot will not draw properly " + NL
'    End If

   'Check if -1 1/2 tape exists for Small Heel
'    If nFirstTape <= 5 And nHeelLength < 9 And Val(Text1(4)) = 0 Then
'        sError = sError + "As the Heel is smaller than 9 inches." + NL
'        sError = sError + "and there is no -1 1/2 tape the foot will not draw properly " + NL
'    End If
   

   'Display Error message (if required) and return
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Fatal - Errors in Data"
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
'    If nFirstTape > 5 Then
'        If cboToeStyle.ListIndex > 0 Then
'            sError = sError + "Toe style, " + cboToeStyle.List(cboToeStyle.ListIndex) + NL
'        End If
'        cboToeStyle.ListIndex = 0
'    End If

   'Display Warning message (if required)
 '   If Len(sError) > 0 Then
 '       sError = "The TOE style given can't apply. As there is no Heel." + NL
 '       sError = sError + "Removing TOE style to reflect this."
 '       MsgBox sError, 64, "Warning - Errors in Data"
 '   End If


   'Yes / No type errors
   'In this case we warn the user that there is a problem!
   'They can continue or they can return to the dialog to make changes

   'Initialize error variables
    sError = ""
    iError = False

   'Make a stab at seeing if the foot length will be long enough
   'For Toe = "Self enclosed" or Toe = "Soft enclosed" and a foot length is given
 '   If nHeelLength > 0 And (cboToeStyle.List(cboToeStyle.ListIndex) = "Self Enclosed" Or (cboToeStyle.List(cboToeStyle.ListIndex) = "Soft Enclosed" And Val(txtFootLength.Text) <> 0)) Then
        
       'Revise heel length to worst case
       'Scale heel to the 0 tape of a 30mmHg template
 '       nHeelLength = nHeelLength * .402
       'Reduce heel by 3
 '       nHeelLength = nHeelLength + (3 * .05025)

       '
 '       If nHeelLength >= 9 Then
 '           nTapeX = 2.71875 'Large heel, distance from heel to -3 tape
 '       Else
 '           nTapeX = 1.5     'Small heel, distance from heel to -1 1/2 tape
 '       End If

 '       nFootLength = fnDisplayToInches(Val(txtFootLength.Text))
        
 '       nValue = Sqr(nHeelLength ^ 2 + nTapeX ^ 2)
'
'        If nValue > nFootLength Then
'            iError = True
'            sError = sError + "The given foot length may be too small to correctly" + NL
'            sError = sError + "position the toe!" + NL + NL
'            sError = sError + "Do you wish to continue anyway ?"
'        End If
'    End If
   
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


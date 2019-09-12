VERSION 5.00
Begin VB.Form cadglved 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6705
   ClientLeft      =   3870
   ClientTop       =   1170
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "cadglved"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6705
   ScaleWidth      =   3510
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   6150
   End
   Begin VB.TextBox txtUIDPALM6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   139
      Text            =   "txtUIDPALM6"
      Top             =   780
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   138
      Text            =   "txtAge"
      Top             =   5100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   137
      Text            =   "txtID"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtTapeLengths2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   136
      Text            =   "txtTapeLengths2"
      Top             =   2130
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtSide 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   135
      Text            =   "txtSide"
      Top             =   4830
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   134
      Text            =   "txtData"
      Top             =   3210
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDTempLFS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   133
      Text            =   "txtUIDTempLFS"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtMM 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   128
      Text            =   "txtMM"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox cboLining 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   131
      Top             =   5760
      Width           =   2115
   End
   Begin VB.TextBox txtLining 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   130
      Text            =   "txtLining"
      Top             =   4290
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDTempETS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   63
      Text            =   "txtUIDTempETS"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   540
      TabIndex        =   64
      Top             =   4980
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   540
      TabIndex        =   65
      Top             =   4710
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   540
      TabIndex        =   66
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   540
      TabIndex        =   67
      Top             =   4170
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   540
      TabIndex        =   68
      Top             =   3900
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   540
      TabIndex        =   69
      Top             =   3630
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   540
      TabIndex        =   70
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   540
      TabIndex        =   71
      Top             =   3090
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   540
      TabIndex        =   72
      Top             =   2820
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   540
      TabIndex        =   73
      Top             =   2550
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   540
      TabIndex        =   74
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   540
      TabIndex        =   75
      Top             =   2010
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   540
      TabIndex        =   76
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   540
      TabIndex        =   77
      Top             =   1470
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   540
      TabIndex        =   78
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   79
      Top             =   930
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   540
      TabIndex        =   80
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox EditTapes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   129
      Top             =   390
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   390
      Width           =   495
   End
   Begin VB.TextBox txtTapeMMs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   30
      Text            =   "txtTapeMMs"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtGrams 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   31
      Text            =   "txtGrams"
      Top             =   2670
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtReduction 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   32
      Text            =   "txtReduction"
      Top             =   2940
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   930
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   1470
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   2010
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   2550
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   2820
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   10
      Top             =   3090
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1800
      TabIndex        =   12
      Top             =   3630
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   1800
      TabIndex        =   13
      Top             =   3900
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   1800
      TabIndex        =   14
      Top             =   4170
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   1800
      TabIndex        =   15
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   1800
      TabIndex        =   16
      Top             =   4710
      Width           =   495
   End
   Begin VB.TextBox EditMMs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   1800
      TabIndex        =   17
      Top             =   4980
      Width           =   495
   End
   Begin VB.TextBox txtUIDPALMER 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   29
      Text            =   "txtUIDPALMER"
      Top             =   510
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   28
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtUIDEOS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   27
      Text            =   "txtUIDEOS"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDETS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   26
      Text            =   "txtUIDETS"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   25
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Finished"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   6240
      Width           =   855
   End
   Begin VB.ComboBox cboContracture 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   23
      Top             =   5370
      Width           =   2115
   End
   Begin VB.TextBox txtTapeLengthsPt1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   22
      Text            =   "txtTapeLengthsPt1"
      Top             =   1860
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtContracture 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   21
      Text            =   "txtContracture"
      Top             =   4020
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUnits 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   20
      Text            =   "txtUnits"
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDLFS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   19
      Text            =   "txtUIDLFS"
      Top             =   1590
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   18
      Text            =   "txtFabric"
      Top             =   3750
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblContracture 
      Caption         =   "Linings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   132
      Top             =   5790
      Width           =   1095
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2355
      TabIndex        =   33
      Top             =   420
      Width           =   495
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   34
      Top             =   690
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   35
      Top             =   960
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   36
      Top             =   1230
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   37
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   38
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   39
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   40
      Top             =   2310
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   41
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   42
      Top             =   2850
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   43
      Top             =   3120
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   44
      Top             =   3390
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   81
      Top             =   3660
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   82
      Top             =   3930
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   83
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   2400
      TabIndex        =   84
      Top             =   4470
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   85
      Top             =   4740
      Width           =   435
   End
   Begin VB.Label EditGrams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   109
      Top             =   5010
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2940
      TabIndex        =   110
      Top             =   420
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   111
      Top             =   690
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2940
      TabIndex        =   112
      Top             =   960
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2940
      TabIndex        =   113
      Top             =   1230
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   114
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   115
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2940
      TabIndex        =   116
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2940
      TabIndex        =   117
      Top             =   2310
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2940
      TabIndex        =   127
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2940
      TabIndex        =   126
      Top             =   2850
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2940
      TabIndex        =   125
      Top             =   3120
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2940
      TabIndex        =   124
      Top             =   3390
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2940
      TabIndex        =   123
      Top             =   3660
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2940
      TabIndex        =   122
      Top             =   3930
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2940
      TabIndex        =   121
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   2940
      TabIndex        =   120
      Top             =   4470
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   2940
      TabIndex        =   119
      Top             =   4740
      Width           =   435
   End
   Begin VB.Label EditReductions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   2940
      TabIndex        =   118
      Top             =   5010
      Width           =   435
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   108
      Top             =   420
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   107
      Top             =   690
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   106
      Top             =   960
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   105
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   104
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   103
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   102
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   101
      Top             =   2310
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   100
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   99
      Top             =   2850
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   98
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   97
      Top             =   3390
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   96
      Top             =   3660
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   95
      Top             =   3930
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   94
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   93
      Top             =   4470
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   92
      Top             =   4740
      Width           =   375
   End
   Begin VB.Label num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   91
      Top             =   5010
      Width           =   375
   End
   Begin VB.Label lblLengths 
      Alignment       =   2  'Center
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   90
      Top             =   90
      Width           =   735
   End
   Begin VB.Label lblMMs 
      Alignment       =   2  'Center
      Caption         =   "MMs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   89
      Top             =   90
      Width           =   495
   End
   Begin VB.Label lblGrams 
      Alignment       =   2  'Center
      Caption         =   "Gram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   88
      Top             =   90
      Width           =   615
   End
   Begin VB.Label lblReduction 
      Alignment       =   2  'Center
      Caption         =   "Reds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   87
      Top             =   90
      Width           =   495
   End
   Begin VB.Label lblContractures 
      Caption         =   "Contracture:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   62
      Top             =   420
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   61
      Top             =   690
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   60
      Top             =   960
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   59
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   58
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   57
      Top             =   1770
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   56
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   55
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   54
      Top             =   2580
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   53
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   52
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1080
      TabIndex        =   51
      Top             =   3390
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   1080
      TabIndex        =   50
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1080
      TabIndex        =   49
      Top             =   3930
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   48
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1080
      TabIndex        =   47
      Top             =   4470
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1080
      TabIndex        =   46
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   1080
      TabIndex        =   45
      Top             =   5010
      Width           =   735
   End
End
Attribute VB_Name = "cadglved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   CADGLVED.MAK
'Purpose:   Editor for CAD glove that extends to the
'           elbow or axilla.
'
'Version:   3.01
'Date:      21.Sep.96
'Author:    Gary George
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Jan 98     GG      Ported to VB5
'
'Notes:-
'
' This is a FOC programme, I've said it before and I'll say
' it again. 'cause I'm a stupid bastard.
'
' Known Bugs
'
' 1 Fails if DRAFIX Help is in use.
'
'
'
    Option Explicit

Private Sub cboContracture_Click()
    Dim sType As String

    If g_NoElbowTape Then Exit Sub
    
    If cboContracture.Text <> txtContracture.Text Then
        txtContracture.Text = cboContracture.Text
        sType = Left$(txtContracture.Text, 2)
        Select Case sType
        Case "No"
           'Reset mms at elbow to standard and recalculate
            Select Case txtMM
                Case "15mm"
                    EditMMs(g_iElbowTape) = 12
                Case "20mm"
                    EditMMs(g_iElbowTape) = 16
                Case "25mm"
                    EditMMs(g_iElbowTape) = 20
            End Select
            PR_CalculateFromMMs g_iElbowTape
            cboContracture.Text = ""
            txtContracture.Text = ""
        Case "10"
           'set the reduction and back calculate grams and mms
            PR_CalculateFromReduction g_iElbowTape, 12
        Case "36"
           'set the reduction and back calculate grams and mms
            PR_CalculateFromReduction g_iElbowTape, 15
        Case "71"
           'set the reduction and back calculate grams and mms
            PR_CalculateFromReduction g_iElbowTape, 17
        End Select

    End If
End Sub

Private Sub cboLining_Click()
    If cboLining.Text <> txtLining.Text Then
        If cboLining.Text = "None" Then cboLining.Text = ""
        txtLining.Text = cboLining.Text
    End If
End Sub

Private Sub cmdCancel_Click()
   'This cancel will delete the tempory curve then exit.
   'Checks to see if there is an Available drafix instance to
   'sendkeys to.
    Dim sTask As String
    If g_ReDrawn = True Then
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            PR_CancelProfileEdits
            AppActivate sTask
            SendKeys "@C:\JOBST\DRAW.D{enter}", True
        End If
    End If
    End
End Sub

Private Sub cmdClose_Click()
   'Check that data is all present and commit into drafix
    If FN_Validate_Data() Then
        PR_DrawAndCommitProfileEdits
        AppActivate fnGetDrafixWindowTitleText()
        SendKeys "@C:\JOBST\DRAW.D{enter}", True
        End
    End If
End Sub

Private Sub cmdDraw_Click()
   'Check that data is all present and insert into drafix
    If FN_Validate_Data() Then
        PR_DrawProfileEdits
        AppActivate fnGetDrafixWindowTitleText()
        g_ReDrawn = True
        SendKeys "@C:\JOBST\DRAW.D{enter}", True
    End If
End Sub

Private Sub EditMMs_GotFocus(Index As Integer)
    Select_Text_In_Box EditMMs(Index)
End Sub

Private Sub EditMMs_LostFocus(Index As Integer)
    If IsNumeric(EditMMs(Index)) Then PR_CalculateFromMMs Index
End Sub

Private Sub EditTapes_GotFocus(Index As Integer)
    Select_Text_In_Box EditTapes(Index)
End Sub

Private Sub EditTapes_LostFocus(Index As Integer)
   'Check length given is valid
   'Display value in inches
    If IsNumeric(EditTapes(Index)) Then
        InchText(Index).Caption = fnInchesToText(fnDisplayToInches(Val(EditTapes(Index))))
        PR_CalculateFromMMs Index
    End If
End Sub

Private Static Function FN_CirLinInt(xyStart As XY, xyEnd As XY, xyCen As XY, nRad As Double, xyInt As XY) As Integer
   'Function to calculate the intersection between
   'a line and a circle.
   'Note:-
   '    Returns true if intersection found.
   '    The first intersection (only) is found.
   '    Ported from DRAFIX CAD DLG version.
   '

    Dim nSlope, nA, nB, nC, nK, nM, nCalcTmp, nRoot As Double
    Dim nSign As Integer

    nSlope = FN_CalcAngle(xyStart, xyEnd)

   'Horizontal Line
    If nSlope = 0 Or nSlope = 180 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.X + Sqr(nC) * nSign
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.y = xyStart.y
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Vertical Line
    If nSlope = 90 Or nSlope = 270 Then
        nSlope = -1
        nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.y + Sqr(nC) * nSign
            If nRoot >= min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
                xyInt.y = nRoot
                xyInt.X = xyStart.X
                FN_CirLinInt = True
                Exit Function
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False
        Exit Function
    End If

   'Non-othogonal line
    If nSlope > 0 Then
        nM = (xyEnd.y - xyStart.y) / (xyEnd.X - xyStart.X)  'Slope
        nK = xyStart.y - nM * xyStart.X                     'Y-Axis intercept
        nA = (1 + nM ^ 2)
        nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.y * nM))
        nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
        nCalcTmp = (nB ^ 2) - (4 * nC * nA)

        If (nCalcTmp < 0) Then
            FN_CirLinInt = False         'No Roots
            Exit Function
        End If
        nSign = 1
        While nSign > -2
            nRoot = (-nB + (Sqr(nCalcTmp) / nSign)) / (2 * nA)
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.y = nM * nRoot + nK
                FN_CirLinInt = True
                Exit Function   'Return first root found
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False 'Should never get to here
    End If
    FN_CirLinInt = False

End Function

Private Function FN_DrawOpen(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_DrawOpen = fNum

   'Initialise String globals
    CC = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & CC & QQ  'Quote Comma Quote ( "," )
    QC = QQ & CC        'Quote Comma ( ", )
    CQ = CC & QQ        'Comma Quote ( ," )

   'Globals to reduced drafix code written to file
    g_sCurrentLayer = ""
    g_nCurrTextAspect = 0.6

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Arm Editing Macro created - "; Date$; "  "; Time$
    Print #fNum, "//by Visual Basic"

End Function

Private Function FN_Validate_Data() As Integer
    
    Dim sError As String, NL As String
    Dim ii As Integer, nFirstTape As Integer, nLastTape As Integer
    Dim iError As Integer
    Dim nValue As Double

    NL = Chr$(10) 'new line
   

   'Display Error message (if required) and return
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        FN_Validate_Data = False
        Exit Function
    Else
        FN_Validate_Data = True
    End If


End Function

Private Sub Form_LinkClose()
   Dim ii As Integer, nn As Integer, iValue As Integer, filenum  As Integer
   Dim nValue As Double, iFirstTape, iLastTape
   Dim textline As String, sString As String

   Timer1.Enabled = False
  'Test that this is an editable glove (Warn and exit if Not)
   g_GloveType = Val(Mid(txtData.Text, (4 * 2) + 1, 2))
   g_EOSType = Val(Mid(txtData.Text, (5 * 2) + 1, 2))
   
   If g_GloveType = GLOVE_NORMAL Then
        MsgBox "This is a NORMAL Glove, Can't EDIT the arm of this type of glove.", 16, "Glove Edit Warning"
        End
   End If

    
  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

  'Other globals
   g_sID = txtID.Text
   g_sSide = txtSide.Text


  'Title bar display
   If g_sSide = "Left" Then
        cadglved.Caption = "CAD GLOVE Edit - Left"
        g_Direction = -1
   End If
   If g_sSide = "Right" Then
        cadglved.Caption = "CAD GLOVE Edit - Right"
        g_Direction = 1
   End If

  'Display - Given Length, Length in inches, MMs, Grams and Reduction at each tape
  'Store values into working arrays. Store also the initial values for change checking
   txtTapeLengthsPt1.Text = "   " & txtTapeLengthsPt1.Text
   txtTapeMMs.Text = "   " & txtTapeMMs.Text
   txtGrams.Text = "   " & txtGrams.Text
   txtReduction.Text = "   " & txtReduction.Text

   For ii = 0 To 17
        nValue = Val(Mid(txtTapeLengthsPt1.Text, (ii * 3) + 1, 3)) / 10
        If nValue > 0 Then

            EditTapes(ii).Text = nValue
            g_nLengths(ii) = nValue
            g_nLengthsInit(ii) = nValue
            InchText(ii).Caption = fnInchesToText(fnDisplayToInches(nValue))
            
            iValue = Val(Mid(txtTapeMMs.Text, (ii * 3) + 1, 3))
            EditMMs(ii) = iValue
            g_iMMs(ii) = iValue
            g_iMMsInit(ii) = iValue
            
            iValue = Val(Mid(txtGrams.Text, (ii * 3) + 1, 3))
            EditGrams(ii).Caption = Str(iValue)
            g_iGms(ii) = iValue
            g_iGmsInit(ii) = iValue

            iValue = Val(Mid(txtReduction.Text, (ii * 3) + 1, 3))
            EditReductions(ii).Caption = Str(iValue)
            g_iRed(ii) = iValue
            g_iRedInit(ii) = iValue

        End If
   Next ii

  'Establish First tape
  'Disable unused tapes
   g_iFirstTape = -1
   iFirstTape = Val(Mid(txtData.Text, (2 * 2) + 1, 2))
   If iFirstTape > 0 Then
        g_iFirstTape = iFirstTape + 1
        For ii = 0 To g_iFirstTape - 1
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).Text = ""
            EditGrams(ii).Caption = ""
            EditReductions(ii).Caption = ""
            InchText(ii).Caption = ""
            EditMMs(ii).BackColor = &HFFFFFF
        Next ii
   Else
        nValue = 0
        For ii = 0 To 17
            nValue = Val(EditTapes(ii).Text)
            If nValue > 0 Then Exit For
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).BackColor = &HFFFFFF
        Next ii
        g_iFirstTape = ii
   End If


  'Establish Last tape (DISABLE down to Last Tape)
   g_iLastTape = -1
   iLastTape = Val(Mid(txtData.Text, (3 * 2) + 1, 2))
   If iLastTape > 0 Then
       g_iLastTape = (17 - iLastTape) + 1
       For ii = 17 To g_iLastTape + 1 Step -1
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).Text = ""
            EditGrams(ii).Caption = ""
            EditReductions(ii).Caption = ""
            InchText(ii).Caption = ""
            EditMMs(ii).BackColor = &HFFFFFF
       Next ii
   Else
       nValue = 0
       For ii = 17 To 0 Step -1
            nValue = Val(EditTapes(ii).Text)
            If nValue > 0 Then Exit For
            EditTapes(ii).Enabled = False
            EditMMs(ii).Enabled = False
            EditMMs(ii).BackColor = &HFFFFFF
       Next ii
       g_iLastTape = ii
   End If

  'If the Proximal style is Flap
  'then disable last tape editing
   If Val(Mid(txtData.Text, (5 * 2) + 1, 2)) = 1 Then
        EditTapes(g_iLastTape).Enabled = False
        EditMMs(g_iLastTape).Enabled = False
        EditMMs(g_iLastTape).BackColor = &HFFFFFF
        EditGrams(g_iLastTape).Enabled = False
        EditReductions(g_iLastTape).Enabled = False
        InchText(g_iLastTape).Enabled = False
   End If
  'Disable first tape editing (CAD glove is always a 10 reduction)
   EditTapes(g_iFirstTape).Enabled = False
   EditMMs(g_iFirstTape).Enabled = False
   EditMMs(g_iFirstTape).BackColor = &HFFFFFF
   EditGrams(g_iFirstTape).Enabled = False
   EditReductions(g_iFirstTape).Enabled = False
    InchText(g_iFirstTape).Enabled = False

  'Always drawn on the fold
   g_OffFold = True
   
  'Establish insert size
  'Insert size is stored in Eighths
   sString = Mid$(txtTapeLengths2.Text, 15, 2)
   If Val(sString) > 0 Then g_nInsertSize = Val(sString) * 0.125

  
  'Load the original profile from file
  'Setup profile vertex to tape mappings
   PR_GetProfileFromFile "C:\JOBST\ARMCURVE.DAT"


  'Load fabric conversion charts from file
  'Establish Fabric
   If Mid(txtFabric.Text, 1, 3) = "Pow" Then
        PR_LoadFabricFromFile g_MATERIAL, g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"
   Else
        MsgBox "Unsupported fabric type, Can't EDIT using this type of Fabric.", 16, "Glove Edit Warning"
        End
   End If

  'Establish Fabric modulus
   g_iModulus = Val(Mid(txtFabric.Text, 5, 3))

  'If there is no elbow tape or elbow tape is the last tape
  'then disable linings and contractures
   g_iElbowTape = 10
   g_sOriginalContracture = txtContracture.Text
   g_sOriginalLining = txtLining.Text
   If Val(EditTapes(g_iElbowTape)) = 0 Or g_iElbowTape = g_iFirstTape Or g_iElbowTape = g_iLastTape Then
        cboContracture.Enabled = False
        cboLining.Enabled = False
        g_NoElbowTape = True
   Else
        cboContracture.Text = txtContracture.Text
        cboLining.Text = txtLining.Text
   End If

  'Set redraw flag to False
   g_ReDrawn = False
   g_BlendingCurveChanged = False

   Show
   Screen.MousePointer = 0    'Change pointer to default.
   cmdCancel.SetFocus

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Dim filenum As Integer, textline As String
   'Hide form while loading
    Hide

   'Start Time out timer
    Timer1.Interval = 6000
    Timer1.Enabled = True

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "The Glove Edit Module is already running!" + Chr(13) + "Use ALT-TAB to access it .", 16, "Leg Edit Warning"
        End
    End If

   'Position to "off" center
    Left = (Screen.Width - cadglved.Width) / 10
    Top = (Screen.Height - cadglved.Height) / 2

    Set MainForm = cadglved

   'Set Tape Numbers
    num(0) = "-6"
    num(1) = "-4" & Chr$(189)
    num(2) = "-3"
    num(3) = "-1" & Chr$(189)
    num(4) = "0"
    num(5) = "1" & Chr$(189)
    num(6) = "3"
    num(7) = "4" & Chr$(189)
    num(8) = "6"
    num(9) = "7" & Chr$(189)
    num(10) = "9"
    num(11) = "10" & Chr$(189)
    num(12) = "12"
    num(13) = "13" & Chr$(189)
    num(14) = "15"
    num(15) = "16" & Chr$(189)
    num(16) = "18"
    num(17) = "19" & Chr$(189)

    txtUnits = ""
    txtUIDETS = ""
    txtUIDLFS = ""
    txtUIDEOS = ""
    txtUIDTempETS = ""
    txtUIDTempLFS = ""
    txtUIDPALMER = ""
    txtSide = ""
    txtID = ""
    txtData = ""
    txtTapeLengthsPt1 = ""
    txtTapeLengths2 = ""
    txtTapeMMs = ""
    txtMM = ""
    txtGrams = ""
    txtReduction = ""
    txtContracture = ""
    txtLining = ""
    txtFabric = ""
    
   'Linings Combo
    cboLining.AddItem "Inside Lining"
    cboLining.AddItem "Outside Lining"
    cboLining.AddItem "Lining"
    cboLining.AddItem "Full Lining"
    cboLining.AddItem "Reinforced Elbow"
    cboLining.AddItem "None"

   'Contractures Combo
    cboContracture.AddItem "10-35 Degrees"
    cboContracture.AddItem "36-70 Degrees"
    cboContracture.AddItem "71 Degrees and Over"
    cboContracture.AddItem "None"

    g_nUnitsFac = 1 'Default to inches
    g_NoElbowTape = False


    g_sPathJOBST = fnPathJOBST()

End Sub

Private Function max(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the maximum of two numbers
    If nFirst >= nSecond Then
        max = nFirst
    Else
        max = nSecond
    End If
End Function

Private Function min(nFirst As Variant, nSecond As Variant) As Variant
    ' Returns the minimum of two numbers
    If nFirst <= nSecond Then
        min = nFirst
    Else
        min = nSecond
    End If
End Function

Private Sub PR_CalculateFromMMs(Index As Integer)
   'Calculate the reduction based on the given lengths
   'and the given MMs

   'If no change has been made to either length or MMs then exit
    If Val(EditTapes(Index).Text) = g_nLengths(Index) And Val(EditMMs(Index).Text) = g_iMMs(Index) Then Exit Sub

    Dim sConversion As String
    Dim nLength As Double
    Dim ii As Integer, iPrevVal As Integer, iVal As Integer
    Dim iReduction As Integer, iGrams As Integer, iMMs As Integer

   'Calculate grams
   'NB allows use of decimal MMs to fudge value
   '
    nLength = fnDisplayToInches(EditTapes(Index).Text)
    iGrams = Int(nLength * Val(EditMMs(Index).Text))
    iMMs = Int(Val(EditMMs(Index).Text))

   'Get conversion string based on modulus
    sConversion = ""
    For ii = 0 To 17
        If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
            sConversion = g_MATERIAL.Conversion(ii)
            Exit For
        End If
    Next ii

    If sConversion = "" Then
        MsgBox "Fabric Modulus not found in conversion chart", 16, "ARM Edit"
        cmdCancel_Click
        End
    End If

    iPrevVal = 0
    For ii = 0 To 22
        iVal = Val(Mid$(sConversion, (ii * 4) + 1, 4))
        If iVal >= iGrams Then Exit For
        iPrevVal = iVal
    Next ii
    
    Select Case ii
        Case 0
           'Minimum reduction
            iReduction = 10
        Case 23
           'Maximum reduction
            iReduction = 32
        Case Else
           'Get reduction closest to given grams
            If (iGrams - iPrevVal) < (iVal - iGrams) Then
                iReduction = ii + 9
            Else
                iReduction = ii + 10
            End If
    End Select
    
   'Modify stored values
    g_iMMs(Index) = iMMs
    g_nLengths(Index) = Val(EditTapes(Index).Text)
    g_iRed(Index) = iReduction
    g_iGms(Index) = iGrams

   'Change display
    EditMMs(Index).Text = iMMs  'Reset MMs to integer value
    EditGrams(Index).Caption = Str(iGrams)
    EditReductions(Index).Caption = Str(iReduction)

   'Show that this tape has been modified
    g_iChanged(Index) = 1

End Sub

Private Sub PR_CalculateFromReduction(Index As Integer, iReduction As Integer)

   'Back calculate from the given reduction revised grams and mms
    
    Dim sConversion As String
    Dim nLength As Double
    Dim ii As Integer, iPrevVal As Integer, iVal As Integer
    Dim iGrams As Integer, iMMs As Integer

   'Quiet exit for bad data
    nLength = fnDisplayToInches(EditTapes(Index).Text)
    If nLength <= 0 Then Exit Sub
    
   'Get conversion string based on modulus
    sConversion = ""
    For ii = 0 To 17
        If g_iModulus = Val(g_MATERIAL.Modulus(ii)) Then
            sConversion = g_MATERIAL.Conversion(ii)
            Exit For
        End If
    Next ii

    If sConversion = "" Then
        MsgBox "Fabric Modulus not found in conversion chart", 16, "ARM Edit"
        cmdCancel_Click
        End
    End If

    
   'Get grams from conversion string based on given reduction
   'N.B. Reductions only run from 10 to 32
   '
    Select Case iReduction
        Case Is < 10
           'Use lowest available grams for those less than 10
            iGrams = Val(Mid$(sConversion, 1, 4))
        Case Is > 32
           'Use highest available grams for those greater than 32
            iGrams = Val(Mid$(sConversion, (22 * 4) + 1, 4))
        Case Else
            iGrams = Val(Mid$(sConversion, ((iReduction - 10) * 4) + 1, 4))
    End Select

   'Back calculate MMs from grams and length
    iMMs = iGrams / nLength

   'Modify stored values
    g_iMMs(Index) = iMMs
    g_nLengths(Index) = Val(EditTapes(Index).Text)
    g_iRed(Index) = iReduction
    g_iGms(Index) = iGrams

   'Change display
    EditMMs(Index).Text = iMMs  'Reset MMs to integer value
    EditGrams(Index).Caption = Str(iGrams)
    EditReductions(Index).Caption = Str(iReduction)

   'Show that this tape has been modified
    g_iChanged(Index) = 1

End Sub

Private Sub PR_CancelProfileEdits()
    Dim ii As Integer

    If g_ReDrawn = True Then
      fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Cancel", g_sSide)
       'Delete Curve Copy
        PR_PutLine "HANDLE hTempCurv;"
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempETS.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"
        If g_OffFold Then
            Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempLFS.Text); ");"
            Print #fNum, "DeleteEntity(hTempCurv);"
        End If
        Print #fNum, "Execute ("; QQ; "menu"; QCQ; "ViewRedraw"; QQ; ");"
      Close fNum
    End If
End Sub

Private Sub PR_ChangeEOS(xyPt As XY, xyPt1 As XY)
   'The EOS is given by the UID txtUIDEOS
    Print #fNum, "hEOS = UID ("; QQ; "find"; QC; Val(txtUIDEOS.Text); ");"
    Print #fNum, "if (hEOS) {"
    Print #fNum, "          SetGeometry (hEOS,"; xyPt.X; ","; xyPt.y; ","; xyPt1.X; ","; xyPt1.y; ");"
    Print #fNum, "          }"
    

End Sub

Private Static Sub PR_DrawAndCommitProfileEdits()
   'Create a macro to commit the changes to the Original profile
    
    Dim xyPt As XY, xyPt1 As XY
    Dim ii As Integer, iProfileVertex  As Integer
    Dim ProfileChanged As Integer
    Dim nFiguredValue As Double
    Dim aAngle#, nOffset#, nLength#
    Dim sLen As String, sMM As String, sGms As String, sRed As String
    Dim sPackedLengths As String, sPackedMMs As String, sPackedGms As String, sPackedRed As String
    Dim iLen As Integer, iMM As Integer, iGms As Integer, iRed  As Integer

    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Manual Glove", g_sSide)
    PR_PutLine "HANDLE  hCurvETS, hCurvLFS, hTempCurv, hArm, hChan, hEnt, hEOS, hPALMER, hPALM6;"
       
    PR_SetLayer "Construct"

   'Delete Curve Copy (Only if redraw has been used)
    If g_ReDrawn = True Then
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempETS.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"
        If g_OffFold Then
            Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempLFS.Text); ");"
            Print #fNum, "DeleteEntity(hTempCurv);"
        End If
    End If

   'Modify Original curve  (need not have been redrawn first)
    ProfileChanged = False
   
    If g_OffFold Then
      Print #fNum, "hCurvETS = UID ("; QQ; "find"; QC; Val(txtUIDETS.Text); ");"
      Print #fNum, "hCurvLFS = UID ("; QQ; "find"; QC; Val(txtUIDLFS.Text); ");"
      If g_BlendingCurveChanged Then
            For ii = 1 To 8
                iProfileVertex = g_iVertexLFSBlendMap(ii)
                Print #fNum, "SetVertex(hCurvLFS,"; iProfileVertex; ","; g_LFSBlendProfile.X(ii + 1); ","; g_LFSBlendProfile.y(ii + 1); ");"
            Next ii
            Print #fNum, "hPALMER = UID ("; QQ; "find"; QC; Val(txtUIDPALMER.Text); ");"
            Print #fNum, "SetGeometry(hPALMER,"; QQ; "xmarker"; QC; xyPALMER.X; CC; xyPALMER.y; ", 0.0625,0.0625, 0);"
            PR_PutTapeLabel g_iFirstTape, fnDisplayToInches(g_nLengths(g_iFirstTape)), g_iMMs(g_iFirstTape), g_iGms(g_iFirstTape), g_iRed(g_iFirstTape)
            ProfileChanged = True
      End If

      For ii = 0 To 17
        If g_iChanged(ii) < 0 Or g_iChanged(ii) > 0 Then
            iProfileVertex = g_iVertexETSMap(ii)
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPt.y = xyProfileETS(iProfileVertex - 1).y
            xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
            Print #fNum, "SetVertex(hCurvETS,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
            iProfileVertex = g_iVertexLFSMap(ii)
            xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
            xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
            Print #fNum, "SetVertex(hCurvLFS,"; iProfileVertex; ","; xyPt1.X; ","; xyPt1.y; ");"
            PR_PutTapeLabel ii, fnDisplayToInches(g_nLengths(ii)), g_iMMs(ii), g_iGms(ii), g_iRed(ii)

            ProfileChanged = True
        End If
      Next ii


     'Modify EOS (Keeps it tidy for Zipper module)
      If g_GloveType = GLOVE_ELBOW And g_iLastTape = ELBOW_TAPE And ((g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0) Or (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0)) Then
           'Glove to elbow, if either the last tape or tape befor the last has changed
            iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPt.y = xyProfileETS(iProfileVertex - 1).y
            xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)

            iProfileVertex = g_iVertexETSMap(g_iLastTape)
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape)) * (Abs(g_iRed(g_iLastTape) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPt1.y = xyProfileETS(iProfileVertex - 1).y
            xyPt1.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
            
            aAngle = Abs(90 - FN_CalcAngle(xyPt1, xyPt))
            If Val(txtAge.Text) > 10 Then nOffset = 0.75 Else nOffset = 0.375
            nLength = nOffset / Cos(aAngle * (PI / 180))
            PR_CalcPolar xyPt1, FN_CalcAngle(xyPt1, xyPt), nLength, xyPt
           
           'Reflect point in X= xyTemplate.
            xyPt1.X = xyOtemplate.X
            xyPt1.y = xyPt.y
            xyPt1.X = xyOtemplate.X + (FN_CalcLength(xyPt1, xyPt) * g_Direction)
            
            PR_ChangeEOS xyPt, xyPt1
        
        ElseIf g_GloveType = GLOVE_AXILLA And g_EOSType = ARM_FLAP And (g_iChanged(g_iLastTape - 1) < 0 Or g_iChanged(g_iLastTape - 1) > 0) Then
           'Glove to axilla with a flap, if the tape befor last has change.
            iProfileVertex = g_iVertexETSMap(g_iLastTape - 1)
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iLastTape - 1)) * (Abs(g_iRed(g_iLastTape - 1) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPt.y = xyProfileETS(iProfileVertex - 1).y
            xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
            iProfileVertex = g_iVertexLFSMap(g_iLastTape - 1)
            xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
            xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
            
            PR_ChangeEOS xyPt, xyPt1
        
        ElseIf g_iChanged(g_iLastTape) < 0 Or g_iChanged(g_iLastTape) > 0 Then
           'All others, if last tape has changed.
           'in this case we can reuse xyPt & xyPt1 set above
            PR_ChangeEOS xyPt, xyPt1

      End If
      
    End If


   'Update contracture if it has been changed
    If txtContracture.Text <> g_sOriginalContracture And g_NoElbowTape = False Then
        PR_DeleteByID g_sID & "Contracture"
        PR_DrawContracture
    End If

   'Add revised lining if it has been changed
    If txtLining.Text <> g_sOriginalLining And g_NoElbowTape = False Then
        PR_DrawLining
    End If

   'Update Origin Marker (only if Profile has been changed)
    If ProfileChanged = True Then

        For ii = 1 To 17
            iLen = g_nLengths(ii) * 10 'Shift decimal place
            iMM = g_iMMs(ii)
            iRed = g_iRed(ii)
            iGms = g_iGms(ii)

            If iLen <> 0 Then
                sLen = String(3, " ")
                RSet sLen = Trim$(Str$(iLen))
                sMM = String(3, " ")
                RSet sMM = Trim$(Str$(iMM))
                sRed = String(3, " ")
                RSet sRed = Trim$(Str$(iRed))
                sGms = String(3, " ")
                RSet sGms = Trim$(Str$(iGms))
            Else
                sLen = String(3, " ")
                sMM = String(3, " ")
                sRed = String(3, " ")
                sGms = String(3, " ")
            End If

            sPackedLengths = sPackedLengths & sLen
            sPackedMMs = sPackedMMs & sMM
            sPackedRed = sPackedRed & sRed
            sPackedGms = sPackedGms & sGms

        Next ii

        Print #fNum, "hArm = UID ("; QQ; "find"; QC; Val(txtUIDPALMER.Text); ");"
        Print #fNum, "if (!hArm)Exit(%cancel,"; QQ; "Can't find PALMER marker to Update"; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "TapeLengthsPt1"; QCQ; sPackedLengths; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "TapeMMs"; QCQ; sPackedMMs; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "Grams"; QCQ; sPackedGms; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "Reduction"; QCQ; sPackedRed; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "Contracture"; QCQ; txtContracture.Text; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "Lining"; QCQ; txtLining.Text; QQ; ");"
    
    End If
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "ViewRedraw"; QQ; ");"

    Close fNum

End Sub

Private Sub PR_DrawContracture()
    Dim xyContractLFS As XY, xyContractETS As XY, xyText As XY
    Dim iElbowVertex As Integer
    Dim nContractureWidth As Double, nFiguredValue As Double, nProfileOffset As Double
    Dim sContracture As String
    Dim Contracture As Curve

    If g_NoElbowTape Or txtContracture.Text = "" Then Exit Sub

   'Get elbow position on profile
    iElbowVertex = g_iVertexETSMap(g_iElbowTape)
    xyText.y = xyProfileETS(iElbowVertex - 1).y
    
    If g_OffFold Then
        nFiguredValue = (((fnDisplayToInches(g_nLengths(g_iElbowTape)) * (Abs(g_iRed(g_iElbowTape) - 100) / 100)) - (2 * g_nInsertSize)) / 2)
        xyContractETS.y = xyProfileETS(iElbowVertex - 1).y
        xyContractETS.X = xyOtemplate.X - (nFiguredValue / 2) * g_Direction
        xyContractLFS.y = xyProfileLFS(iElbowVertex - 1).y
        xyContractLFS.X = xyOtemplate.X + (nFiguredValue / 2) * g_Direction
        xyText.X = xyOtemplate.X
    Else
        nFiguredValue = (((fnDisplayToInches(g_nLengths(g_iElbowTape)) * (Abs(g_iRed(g_iElbowTape) - 100) / 100)) - g_nInsertSize) / 2)
        xyContractETS.y = xyProfileETS(iElbowVertex - 1).y
        xyContractETS.X = xyOtemplate.X - nFiguredValue * g_Direction
        xyContractLFS.y = xyProfileLFS(iElbowVertex - 1).y
        xyContractLFS.X = xyOtemplate.X
        xyText.X = xyOtemplate.X - (nFiguredValue / 2)
    End If

   'Contracture width
    sContracture = Left$(txtContracture.Text, 2)
    Select Case sContracture
        Case "No"
           'Just in case
            Exit Sub
        Case "10"
            nContractureWidth = 0.5
        Case "36"
            nContractureWidth = 1
        Case "71"
            nContractureWidth = 1.5
    End Select

   'Create contracture polyline
    Contracture.n = 5
    
   'Start at Thumbside
    Contracture.X(1) = xyContractETS.X
    Contracture.y(1) = xyContractETS.y

   'To the top
    Contracture.y(2) = xyContractETS.y + (nContractureWidth / 2)
    Contracture.X(2) = xyContractETS.X + ((xyContractLFS.X - xyContractETS.X) / 2)

   'To Little finger side
    Contracture.X(3) = xyContractLFS.X
    Contracture.y(3) = xyContractLFS.y
   
   'To the bottom
    Contracture.y(4) = Contracture.y(2) - nContractureWidth
    Contracture.X(4) = Contracture.X(2)

   'Close to the thumbside
    Contracture.X(5) = xyContractETS.X
    Contracture.y(5) = xyContractETS.y

   'Draw contracture
    PR_SetLayer "Notes"
    PR_DrawPoly Contracture
    PR_AddEntityID g_sID, "", "Contracture"
    PR_SetTextData 2, 32, -1, -1, -1
    PR_DrawText "Remove For Contracture", xyText, 0.125

End Sub

Private Sub PR_DrawLining()
    Dim xyElbow As XY
    Dim iElbowVertex As Integer
    Dim nFiguredValue As Double

    If g_NoElbowTape Or txtLining.Text = "" Then Exit Sub
    
   'Get elbow position on profile
    iElbowVertex = g_iVertexETSMap(g_iElbowTape)
    xyElbow.y = xyProfileETS(iElbowVertex - 1).y
    
    If g_OffFold Then
        xyElbow.X = xyOtemplate.X
    Else
        nFiguredValue = (((fnDisplayToInches(g_nLengths(g_iElbowTape)) * (Abs(g_iRed(g_iElbowTape) - 100) / 100)) - g_nInsertSize) / 2)
        xyElbow.X = xyOtemplate.X - (nFiguredValue / 2) * g_Direction
    End If

   'Draw Lining Text
    PR_SetLayer "Notes"
    PR_SetTextData 2, 8, -1, -1, -1
    PR_DrawText txtLining, xyElbow, 0.125

End Sub

Private Sub PR_DrawProfileEdits()
   'Create a macro to copy part of the arm profile and
   'modify this copy
   'Modifications will only be committed to the original profile on "Finish"
   'The exception to this is the ShortArm where due to DRAFIX bugs the edits
   'have to be done directly on the profile.

    Dim xyPt As XY, xyNull As XY, xyPt1 As XY
    Dim ii As Integer, iProfileVertex As Integer
    Dim nFiguredValue As Double


    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Profile Edits", g_sSide)
    PR_PutLine "HANDLE  hCurvETS, hCurvLFS;"

       
    If g_ReDrawn <> True Then
       'Make Copy of curve (First time through only)

        PR_SetLayer "Construct"
        Print #fNum, "hCurvETS = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        
        For ii = 0 To g_iProfileETS
            Print #fNum, CC; xyProfileETS(ii).X; CC; xyProfileETS(ii).y;
        Next ii
        
        Print #fNum, ");"

        If g_OffFold Then
            Print #fNum, "hCurvLFS = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
            
            For ii = 0 To g_iProfileLFS
                Print #fNum, CC; xyProfileLFS(ii).X; CC; xyProfileLFS(ii).y;
            Next ii
            
            Print #fNum, ");"
        End If

    Else
       'Modify Curve Copy
        If txtUIDTempETS.Text <> "" Then
            Print #fNum, "hCurvETS = UID ("; QQ; "find"; QC; Val(txtUIDTempETS.Text); ");"
            If g_OffFold Then Print #fNum, "hCurvLFS = UID ("; QQ; "find"; QC; Val(txtUIDTempLFS.Text); ");"
        End If
    End If

   'Loop through all tape values.  If changed then modify profile copy
    If g_OffFold Then
      If g_iChanged(g_iFirstTape) > 0 Or g_iChanged(g_iFirstTape + 1) > 0 Then
            g_iChanged(g_iFirstTape) = 0

            g_BlendingCurveChanged = True
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape)) * (Abs(g_iRed(g_iFirstTape) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPALMER.X = xyPALM6.X + (nFiguredValue * g_Direction)

            nFiguredValue = ((((fnDisplayToInches(g_nLengths(g_iFirstTape + 1)) * (Abs(g_iRed(g_iFirstTape + 1) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            iProfileVertex = g_iVertexLFSMap(g_iFirstTape + 1)
            xyPt1.y = xyProfileLFS(iProfileVertex - 1).y
            xyPt1.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
            PR_CalcWristBlendLFS xyArcCenter, aArcSweep, xyArcStart, xyPALMER, xyPt1, g_LFSBlendProfile
            For ii = 1 To 8
                iProfileVertex = g_iVertexLFSBlendMap(ii)
                Print #fNum, "SetVertex(hCurvLFS,"; iProfileVertex; ","; g_LFSBlendProfile.X(ii + 1); ","; g_LFSBlendProfile.y(ii + 1); ");"
            Next ii
      
      End If
      For ii = 0 To 17
        If g_iChanged(ii) > 0 Then
            g_iChanged(ii) = -1
            iProfileVertex = g_iVertexETSMap(ii)
            nFiguredValue = ((((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) + 0.375) - (2 * g_nInsertSize)) / 2)
            xyPt.y = xyProfileETS(iProfileVertex - 1).y
            xyPt.X = xyOtemplate.X - ((nFiguredValue / 2) * g_Direction)
            Print #fNum, "SetVertex(hCurvETS,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
            iProfileVertex = g_iVertexLFSMap(ii)
            xyPt.y = xyProfileLFS(iProfileVertex - 1).y
            xyPt.X = xyOtemplate.X + ((nFiguredValue / 2) * g_Direction)
            Print #fNum, "SetVertex(hCurvLFS,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
        End If
      Next ii
    End If
    If g_ReDrawn <> True Then
        PR_PutLine "HANDLE  hDDE;"

       'Poke Curve UID back to the VB program
        Print #fNum, "hDDE = Open ("; QQ; "dde"; QCQ; "cadglved"; QCQ; "cadglved"; QQ; ");"
        Print #fNum, "Poke ( hDDE, "; QQ; "txtUIDTempETS"; QC; "MakeString("; QQ; "long"; QC; "UID("; QQ; "get"; QQ; ",hCurvETS)));"
        If g_OffFold Then Print #fNum, "Poke ( hDDE, "; QQ; "txtUIDTempLFS"; QC; "MakeString("; QQ; "long"; QC; "UID("; QQ; "get"; QQ; ",hCurvLFS)));"

    End If
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "ViewRedraw"; QQ; ");"

    Close fNum

End Sub

Private Sub PR_GetProfileFromFile(sFileName As String)
   'Procedure to read curve data from file
   'Data is a series of x, y values.
   'In the following format
   '
   '    Line    Type
   '-----------------------------------------------
   '    1       PALMER Marker xyOTemplate
   '    2       PALM6 Marker
   '    3       Little finger arc properites
   '    4       Centerline
   '    5       n_LFS = number of vertices Little Finger side
   '    6  }
   '    .  }}
   '    .  }}}  Vertices of profile (Profile)
   '    .  }}
   '    n  }
   '    n_LFS+2 N_ETS = number of vertices Thumb side
   '    n+2}
   '    .  }}
   '    .  }}}  Vertices of profile (Profile)
   '    .  }}
   '    n_ETS  }
   '
   'The editable vertices are setup by use of an array that maps the
   'vertex of the profile to the tapes.

    Dim ii As Integer, iPalm As Integer, iVertex As Integer
    Dim fFileNum As Integer, nTolerance As Double, nLength As Double
    Dim xyCLStart As XY, xyCLEnd As XY
    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48
        Exit Sub
    End If

    Open sFileName For Input As fFileNum

   'Get control points
    Input #fFileNum, xyOtemplate.X, xyOtemplate.y
    xyPALMER = xyOtemplate

    Input #fFileNum, xyPALM6.X, xyPALM6.y
    Input #fFileNum, xyArcCenter.X, xyArcCenter.y, nArcRadius, aArcStart, aArcSweep
    If g_sSide = "Right" Then
        PR_CalcPolar xyArcCenter, aArcStart + aArcSweep, nArcRadius, xyArcStart
        aArcSweep = -aArcSweep
    Else
        PR_CalcPolar xyArcCenter, aArcStart, nArcRadius, xyArcStart
    End If
    Input #fFileNum, xyCLStart.X, xyCLStart.y, xyCLEnd.X, xyCLEnd.y

   'Get profile points
    g_iProfileLFS = 0
    Do While Not EOF(fFileNum)
        Input #fFileNum, g_iProfileLFS
        g_iProfileLFS = g_iProfileLFS - 1
        For ii = 0 To g_iProfileLFS
            Input #fFileNum, xyProfileLFS(ii).X, xyProfileLFS(ii).y
        Next ii
        Input #fFileNum, g_iProfileETS
        g_iProfileETS = g_iProfileETS - 1
        For ii = 0 To g_iProfileETS
            Input #fFileNum, xyProfileETS(ii).X, xyProfileETS(ii).y
        Next ii
    Loop
    Close #fFileNum

   
   'For editing purposes we must find the profile vertex that maps to
   'a tape. We therefor create an array that contains the profile vertex
   'number for each tape
   '

    
   'by starting at the last tape and working forward
   'we can map vertex to editable profile
    iVertex = g_iProfileETS + 1
    For ii = g_iLastTape To g_iFirstTape - 1 Step -1
        g_iVertexETSMap(ii) = iVertex
        iVertex = iVertex - 1
    Next ii

   'The next stage is map the blending curve vertex
   '
    iVertex = iVertex + 2
    For ii = 8 To 1 Step -1
        g_iVertexETSBlendMap(ii) = iVertex
        iVertex = iVertex - 1
    Next ii

    If g_OffFold Then

       'Map editable profile
        iVertex = g_iProfileLFS + 1
        For ii = g_iLastTape To g_iFirstTape - 1 Step -1
            g_iVertexLFSMap(ii) = iVertex
            iVertex = iVertex - 1
        Next ii
    
       'The next stage is map the blending curve vertex
       '
        iVertex = iVertex + 2
        For ii = 8 To 1 Step -1
            g_iVertexLFSBlendMap(ii) = iVertex
            iVertex = iVertex - 1
        Next ii

       'Reset the origin to lie on the midpoint of the wrist line
'        nLength = Abs(xyPALM6.x - xyOtemplate.x)
'        If g_sSide = "Left" Then
'            xyOtemplate.x = xyOtemplate.x + nLength / 2
'        Else
'            xyOtemplate.x = xyOtemplate.x - nLength / 2
'        End If
        xyOtemplate.X = xyCLStart.X
    End If


End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable
   '
    Print #fNum, sLine
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
    Print #fNum, "if ( hLayer != %badtable)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Private Sub Timer1_Timer()
    End
End Sub


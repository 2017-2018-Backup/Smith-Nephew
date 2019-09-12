VERSION 5.00
Begin VB.Form armeddia 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6750
   ClientLeft      =   2835
   ClientTop       =   495
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
   LinkTopic       =   "armeddia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6750
   ScaleWidth      =   3510
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
      TabIndex        =   129
      Text            =   "txtMM"
      Top             =   5670
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
      TabIndex        =   133
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
      TabIndex        =   132
      Text            =   "txtLining"
      Top             =   5250
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtGauntlet 
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
      TabIndex        =   131
      Text            =   "txtGauntlet"
      Top             =   4890
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDTempArm 
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
      TabIndex        =   64
      Text            =   "txtUIDTempArm"
      Top             =   4500
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
      TabIndex        =   65
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
      TabIndex        =   66
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
      TabIndex        =   67
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
      TabIndex        =   68
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
      TabIndex        =   69
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
      TabIndex        =   70
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
      TabIndex        =   71
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
      TabIndex        =   72
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
      TabIndex        =   73
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
      TabIndex        =   74
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
      TabIndex        =   75
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
      TabIndex        =   76
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
      TabIndex        =   77
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
      TabIndex        =   78
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
      TabIndex        =   79
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
      TabIndex        =   80
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
      TabIndex        =   81
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
      TabIndex        =   130
      Top             =   390
      Width           =   495
   End
   Begin VB.TextBox txtStump 
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
      Text            =   "txtStump"
      Top             =   3780
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
      TabIndex        =   31
      Text            =   "txtTapeMMs"
      Top             =   2340
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
      TabIndex        =   32
      Text            =   "txtGrams"
      Top             =   2700
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
      TabIndex        =   33
      Text            =   "txtReduction"
      Top             =   3060
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
   Begin VB.TextBox txtUIDArm 
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
      Text            =   "txtUIDArm"
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
   Begin VB.TextBox txtArm 
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
      Text            =   "txtArm"
      Top             =   1230
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtUIDCurve 
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
      Text            =   "txtUIDCurve"
      Top             =   870
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
      Appearance      =   0  'Flat
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
   Begin VB.TextBox txtTapeLengths 
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
      Text            =   "txtTapeLengths"
      Top             =   1980
      Visible         =   0   'False
      Width           =   1500
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
      Top             =   3420
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
      Top             =   150
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
      TabIndex        =   19
      Text            =   "txtID"
      Top             =   1620
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
      Top             =   4140
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
      TabIndex        =   134
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
      TabIndex        =   34
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
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   37
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
      TabIndex        =   38
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
      TabIndex        =   39
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
      TabIndex        =   40
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
      TabIndex        =   41
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
      TabIndex        =   42
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
      TabIndex        =   43
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
      TabIndex        =   44
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
      TabIndex        =   45
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
      TabIndex        =   82
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
      TabIndex        =   83
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
      TabIndex        =   84
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
      TabIndex        =   85
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
      TabIndex        =   86
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
      TabIndex        =   110
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
      TabIndex        =   111
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
      TabIndex        =   112
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
      TabIndex        =   113
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
      TabIndex        =   114
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
      TabIndex        =   115
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
      TabIndex        =   116
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
      TabIndex        =   117
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
      TabIndex        =   118
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
      TabIndex        =   128
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
      TabIndex        =   127
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
      TabIndex        =   126
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
      TabIndex        =   125
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
      TabIndex        =   124
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
      TabIndex        =   123
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
      TabIndex        =   122
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
      TabIndex        =   121
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
      TabIndex        =   120
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
      TabIndex        =   119
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
      TabIndex        =   109
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
      TabIndex        =   108
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
      TabIndex        =   107
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
      TabIndex        =   106
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
      TabIndex        =   105
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
      TabIndex        =   104
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
      TabIndex        =   103
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
      TabIndex        =   102
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
      TabIndex        =   101
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
      TabIndex        =   100
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
      TabIndex        =   99
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
      TabIndex        =   98
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      TabIndex        =   95
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
      TabIndex        =   94
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
      TabIndex        =   93
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
      TabIndex        =   92
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
      TabIndex        =   91
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
      TabIndex        =   90
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
      TabIndex        =   89
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
      TabIndex        =   88
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
      TabIndex        =   87
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
      TabIndex        =   63
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
      TabIndex        =   62
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
      TabIndex        =   61
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
      TabIndex        =   60
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
      TabIndex        =   59
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   56
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   52
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
      TabIndex        =   51
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
      Top             =   5010
      Width           =   735
   End
End
Attribute VB_Name = "armeddia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   ARMEDDIA
'Purpose:   Editor for Arms and Vest sleeves.
'
'Version:   3.01
'Date:      28.Mar.95
'Author:    Gary George
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Ported to VB 5
'
'Notes:-
'
' Known Bugs
'
' 1 With a gauntlet extra or missing points above the wrist
'   are not checked for.
'
' 2 The vertex mapping that is currently used only with the
'   gaunlets between the palm and wrist should be extended
'   to included the whole arm.
'
' 3 Fails if DRAFIX Help is in use.
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

Private Sub cboLinings_Change()
    If g_NoElbowTape Then Exit Sub
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
    If FN_CheckValue(EditMMs(Index), "MM's") Then
          PR_CalculateFromMMs Index
    End If
End Sub

Private Sub EditTapes_GotFocus(Index As Integer)
    Select_Text_In_Box EditTapes(Index)
End Sub

Private Sub EditTapes_LostFocus(Index As Integer)
   'Check length given is valid
   'Display value in inches
    If FN_CheckValue(EditTapes(Index), "Tape Length") Then
        InchText(Index).Caption = fnInchesToText(fnDisplayToInches(Val(EditTapes(Index))))
        PR_CalculateFromMMs Index
    End If
End Sub

Private Function FN_CheckValue(TextBox As Control, sMessage As String) As Integer
   'Check that a valid numeric value has been Entered
    Dim sChar As String, sText As String
    Dim iLen As Integer, nn As Integer
    sText = TextBox.Text
    iLen = Len(sText)
    FN_CheckValue = True
    For nn = 1 To iLen
        sChar = Mid$(sText, nn, 1)
        If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
            MsgBox "Invalid " & sMessage & " has been entered", 48, "Arm Details"
            TextBox.SetFocus
            FN_CheckValue = False
            Exit For
        End If
    Next nn
End Function

Private Function FN_DrawOpen(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_DrawOpen = fNum

   'Initialise String globals
    cc = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & cc & QQ  'Quote Comma Quote ( "," )
    QC = QQ & cc        'Quote Comma ( ", )
    CQ = cc & QQ        'Comma Quote ( ," )

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
   Dim nValue As Double
   Dim textline As String
   Dim sDistalStyle As String, sProximalStyle As String

  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

   Set MainForm = armeddia
  'Other globals
   g_sID = txtID.Text & txtArm.Text
   g_sFileNo = Mid(txtID.Text, 5)
   g_sStyle = Mid(txtID.Text, 1, 4)
   g_sSide = txtArm.Text

  'Proximal and Distal end tape Styles
   sDistalStyle = Mid(txtID.Text, 1, 2)
   sProximalStyle = Mid(txtID.Text, 3, 2)

  'Get gauntlet Details
   g_Gauntlet = fnGetNumber(txtGauntlet.Text, 1)
   If g_Gauntlet Then
       g_iWristNo = fnGetNumber(txtGauntlet.Text, 5) - 1
       g_iPalmNo = fnGetNumber(txtGauntlet.Text, 6) - 1
       g_nPalmWristDist = fnDisplayToInches(fnGetNumber(txtGauntlet.Text, 9))
   End If

  'Title bar display
   If txtArm.Text = "Left" Then
        armeddia.Caption = "ARM Edit - Left [" & g_sFileNo & "]"
   End If
   If txtArm.Text = "Right" Then
        armeddia.Caption = "ARM Edit - Right [" & g_sFileNo & "]"
   End If

  'Display - Given Length, Length in inches, MMs, Grams and Reduction at each tape
  'Store values into working arrays. Store also the initial values for change checking
   For ii = 0 To 17
        nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3)) / 10
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
   nValue = 0
   For ii = 0 To 17
        nValue = Val(EditTapes(ii).Text)
        If nValue > 0 Then Exit For
        EditTapes(ii).Enabled = False
        EditMMs(ii).Enabled = False
        EditMMs(ii).BackColor = &HFFFFFF
   Next ii
   g_iFirstTape = ii


  'If the first tape is on the elbow then disable
  'first tape editing
   If g_iFirstTape = 10 Then
        EditTapes(g_iFirstTape).Enabled = False
        EditMMs(g_iFirstTape).Enabled = False
        EditMMs(g_iFirstTape).BackColor = &HFFFFFF
   End If

  'Establish Last tape (DISABLE down to Last Tape)
   g_iLastTape = -1
   nValue = 0
   For ii = 17 To 0 Step -1
        nValue = Val(EditTapes(ii).Text)
        If nValue > 0 Then Exit For
        EditTapes(ii).Enabled = False
        EditMMs(ii).Enabled = False
        EditMMs(ii).BackColor = &HFFFFFF
   Next ii
   g_iLastTape = ii

  'If the Proximal style is Flap or is a Vest Raglan
  'then disable last tape editing
   If sProximalStyle = "FP" Or sProximalStyle = "VR" Then
        EditTapes(g_iLastTape).Enabled = False
        EditMMs(g_iLastTape).Enabled = False
        EditMMs(g_iLastTape).BackColor = &HFFFFFF
   End If

  'Set redraw flag to False
   g_ReDrawn = False
   g_ShortArm = False
  
  'Load the original profile from file
  'Setup profile vertex to tape mappings
   PR_GetProfileFromFile "C:\JOBST\ARMCURVE.DAT"
   

  'Load fabric conversion charts from file
  'Establish Fabric
   If Mid(txtFabric.Text, 1, 3) = "Pow" Then
        PR_LoadFabricFromFile g_MATERIAL, g_sPathJOBST & "\TEMPLTS\POWERNET.DAT"
   Else
        PR_LoadFabricFromFile g_MATERIAL, g_sPathJOBST & "\TEMPLTS\BOBINNET.DAT"
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
        cboContracture = txtContracture
        cboLining = txtLining
   End If

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

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "The Arm Edit Module is already running!" + Chr(13) + "Use ALT-TAB to access it .", 16, "Leg Edit Warning"
        End
    End If

   'Position to center of screen
    Left = (Screen.Width - armeddia.Width) / 10
    Top = (Screen.Height - armeddia.Height) / 2

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
    txtUIDArm = ""
    txtUIDTempArm = ""
    txtUIDCurve = ""
    txtArm = ""
    txtID = ""
    txtTapeLengths = ""
    txtTapeMMs = ""
    txtGrams = ""
    txtReduction = ""
    txtContracture = ""
    txtLining = ""
    txtStump = ""
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
    g_ShortArm = False
    g_NoElbowTape = False


    g_sPathJOBST = fnPathJOBST()

End Sub

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
      fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Cancel", txtArm)
      
      If g_ShortArm = True Then
       'Restore original curve (Short Arms only)
        PR_PutLine "HANDLE  hCurv;"
        Print #fNum, "hCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempArm.Text); ");"
        For ii = 0 To g_iProfile
            Print #fNum, "SetVertex(hCurv,"; ii + 1; ","; xyOriginal(ii).X; ","; xyOriginal(ii).y; ");"
        Next ii
      Else
       'Delete Curve Copy
        PR_PutLine "HANDLE hTempCurv;"
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempArm.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"
      End If
      Close fNum
    End If
End Sub

Private Sub PR_DeleteTapeLabel(xyPt As XY, nTape)
    'Deletes the tape label and the text at the given point
     Dim slayer As String
     
     Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double

     x1 = xyPt.X
     x2 = xyPt.X + 1
     y1 = xyPt.y
     y2 = xyPt.y + 2
     slayer = g_sFileNo & Mid$(g_sSide, 1, 1) & nTape

     Print #fNum, "hChan=Open("; QQ; "selection"; QCQ; "(type = 'Text' OR type = 'circle') AND layer = '"; slayer; "' AND TOTally INside "; x1; y1; x2; y2; QQ; ");"
     Print #fNum, "if(hChan)"
     Print #fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}"
     Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"

End Sub

Private Sub PR_DrawAndCommitProfileEdits()

   'Create a macro to commit the changes to the Original profile
    
    Dim xyPt As XY
    Dim ii As Integer, iProfileVertex  As Integer
    Dim ProfileChanged As Integer
    Dim nSeam As Double
    Dim sLen As String, sMM As String, sGms As String, sRed As String
    Dim sPackedLengths As String, sPackedMMs As String, sPackedGms As String, sPackedRed As String
    Dim iLen As Integer, iMM As Integer, iGms As Integer, iRed  As Integer
    
    nSeam = 0.1875

    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtArm)
    PR_PutLine "HANDLE  hCurv, hTempCurv, hArm, hChan, hEnt;"
       
    PR_SetLayer "Construct"

   'Delete Curve Copy (Only if redraw has been used)
    If g_ReDrawn = True And g_ShortArm = False Then
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempArm.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"
    End If

   'Modify Original curve  (need not have been redrawn first)
    Print #fNum, "hCurv = UID ("; QQ; "find"; QC; Val(txtUIDCurve.Text); ");"
    ProfileChanged = False
    For ii = 0 To 17
        If g_iChanged(ii) < 0 Or g_iChanged(ii) > 0 Then
            iProfileVertex = g_iVertexMap(ii)
            xyPt.X = xyProfile(iProfileVertex - 1).X
            xyPt.y = ((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
            Print #fNum, "SetVertex(hCurv,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
            xyPt.y = xyOtemplate.y
            PR_DeleteTapeLabel xyPt, ii + 1
            PR_PutTapeLabel ii + 1, xyPt, fnDisplayToInches(g_nLengths(ii)), g_iMMs(ii), g_iGms(ii), g_iRed(ii)
            ProfileChanged = True
        End If
    Next ii
   'Update contracture if it has been changed
    If txtContracture.Text <> g_sOriginalContracture And g_NoElbowTape = False Then
        PR_DeleteByID g_sID & "Contracture"
        PR_DrawContracture
    End If

   'Add revised lining if it has been changed
    If txtLining.Text <> g_sOriginalLining And g_NoElbowTape = False Then
        PR_DrawLining
    End If

   'Update Arm Box or Origin Marker (only if Profile has been changed)
    If ProfileChanged = True Then
        For ii = 0 To 17
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


        Print #fNum, "hArm = UID ("; QQ; "find"; QC; Val(txtUIDArm.Text); ");"
        Print #fNum, "if (!hArm)Exit(%cancel,"; QQ; "Can't find ARMBOX to Update"; QQ; ");"
        Print #fNum, "SetDBData( hArm, "; QQ; "TapeLengths"; QCQ; sPackedLengths; QQ; ");"
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
    Dim xyContractTop As XY, xyContractBott As XY
    Dim iElbowVertex As Integer
    Dim nContractureWidth As Double, nSeam As Double, nProfileOffset As Double
    Dim sContracture As String
    Dim Contracture As curve

    If g_NoElbowTape Or txtContracture.Text = "" Then Exit Sub
    
   'Get elbow position on profile
    nSeam = 0.1875
    nProfileOffset = 0.5
    iElbowVertex = g_iVertexMap(g_iElbowTape)
    xyContractTop.X = xyProfile(iElbowVertex - 1).X
    xyContractTop.y = ((fnDisplayToInches(g_nLengths(g_iElbowTape)) * (Abs(g_iRed(g_iElbowTape) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
    xyContractTop.y = xyContractTop.y - nProfileOffset

   'Elbow position on the edge of the template at the fold line
    xyContractBott.X = xyContractTop.X
    xyContractBott.y = xyOtemplate.y

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
    
   'Start at bottom
    Contracture.X(1) = xyContractBott.X
    Contracture.y(1) = xyContractBott.y

   'To the left
    Contracture.X(2) = xyContractBott.X - (nContractureWidth / 2)
    Contracture.y(2) = xyContractBott.y + ((xyContractTop.y - xyContractBott.y) / 2)

   'To the top
    Contracture.X(3) = xyContractTop.X
    Contracture.y(3) = xyContractTop.y
   
   'To the right
    Contracture.X(4) = xyContractBott.X + (nContractureWidth / 2)
    Contracture.y(4) = Contracture.y(2)

   'Close to the bottom
    Contracture.X(5) = xyContractBott.X
    Contracture.y(5) = xyContractBott.y

   'Draw contracture
    PR_SetLayer "Notes"
    PR_DrawPoly Contracture
    PR_AddEntityID g_sID, "", "Contracture"
    PR_SetTextData 1, 32, -1, -1, -1
    PR_DrawText "Remove For Contracture", xyContractTop, 0.125

End Sub

Private Sub PR_DrawLining()
    Dim xyElbow As XY
    Dim iElbowVertex As Integer
    Dim nProfileOffset As Double, nSeam As Double

    If g_NoElbowTape Or txtLining.Text = "" Then Exit Sub
    
   'Get elbow position on profile
    nProfileOffset = 1
    nSeam = 0.1875
    iElbowVertex = g_iVertexMap(g_iElbowTape)
    xyElbow.X = xyProfile(iElbowVertex - 1).X
    xyElbow.y = ((fnDisplayToInches(g_nLengths(g_iElbowTape)) * (Abs(g_iRed(g_iElbowTape) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
    xyElbow.y = xyElbow.y - nProfileOffset

   'Draw Lining Text
    PR_SetLayer "Notes"
    PR_SetTextData 2, 32, -1, -1, -1
    PR_DrawText txtLining, xyElbow, 0.125

End Sub

Private Sub PR_DrawProfileEdits()
   'Create a macro to copy part of the arm profile and
   'modify this copy
   'Modifications will only be committed to the original profile on "Finish"
   'The exception to this is the ShortArm where due to DRAFIX bugs the edits
   'have to be done directly on the profile.

    Dim xyPt   As XY
    Dim ii As Integer, iProfileVertex As Integer
    Dim nSeam As Double

    nSeam = 0.1875

    If g_ShortArm = True Then txtUIDTempArm = txtUIDCurve

    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Profile Edits", txtArm)
    PR_PutLine "HANDLE  hCurv;"
       
    If g_ReDrawn <> True And g_ShortArm = False Then
       'Make Copy of curve (First time through only)
       'Not for short arms (3 tapes)

        PR_SetLayer "Construct"
        Print #fNum, "hCurv = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        
        For ii = 0 To g_iProfile
            Print #fNum, cc; xyProfile(ii).X; cc; xyProfile(ii).y;
        Next ii
        
        Print #fNum, ");"

    Else
       'Modify Curve Copy
        If txtUIDTempArm.Text <> "" Then
            Print #fNum, "hCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempArm.Text); ");"
        Else
            
        End If
    End If

   'Loop through all tape values.  If changed then modify profile copy
    For ii = 0 To 17
        If g_iChanged(ii) > 0 Then
            g_iChanged(ii) = -1
            iProfileVertex = g_iVertexMap(ii)
            xyPt.X = xyProfile(iProfileVertex - 1).X
            xyPt.y = ((fnDisplayToInches(g_nLengths(ii)) * (Abs(g_iRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
            Print #fNum, "SetVertex(hCurv,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
        End If
    Next ii

    If g_ReDrawn <> True And g_ShortArm = False Then
        PR_PutLine "HANDLE  hDDE;"

       'Poke Curve UID back to the VB program
        Print #fNum, "hDDE = Open ("; QQ; "dde"; QCQ; "armeddia"; QCQ; "armeddia"; QQ; ");"
        Print #fNum, "Poke ( hDDE, "; QQ; "txtUIDTempArm"; QC; "MakeString("; QQ; "long"; QC; "UID("; QQ; "get"; QQ; ",hCurv)));"

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
   '    1       Template Origin (xyOtemplate)
   '    2  }
   '    .  }}
   '    .  }}}  Vertices of profile (Profile)
   '    .  }}
   '    N  }
   '
   '
   'The editable vetices are setup by use of an array that maps the
   'vertex of the profile to the tapes.

    Dim ii As Integer, iPalm As Integer, iVertex As Integer
    Dim fFileNum As Integer, nTolerance As Double

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48
        Exit Sub
    End If

    Open sFileName For Input As fFileNum

   'Get control points
    Input #fFileNum, xyOtemplate.X, xyOtemplate.y

   'Get profile points
    g_iProfile = 0
    Do While Not EOF(fFileNum)
        Input #fFileNum, xyProfile(g_iProfile).X, xyProfile(g_iProfile).y
       'Store original profile for use in cancel for short arms
        xyOriginal(g_iProfile).X = xyProfile(g_iProfile).X
        xyOriginal(g_iProfile).y = xyProfile(g_iProfile).y
        g_iProfile = g_iProfile + 1
        Loop
    g_iProfile = g_iProfile - 1
    Close #fFileNum

   'If a gauntlet has been given then the user may have added additional
   'Vertices to smooth the flow of the curve into the gaunlet end.  These
   'will be between the WristTape and the PalmTape.
   '
   'For editing purposes we must find the profile vertex that maps to
   'a tape. We therefor create an array that contains the profile vertex
   'number for each tape
   '
    If g_Gauntlet Then
       'One to One mapping up to Palm
        iVertex = 1
        ii = g_iFirstTape
        Do
            g_iVertexMap(ii) = iVertex
            iVertex = iVertex + 1
            ii = ii + 1
        Loop Until ii > g_iPalmNo
        iPalm = iVertex - 1

       'Skip Vertex between palm and wrist
       'NB Use of tolerance
        nTolerance = 0.01
        For ii = iPalm To g_iProfile
            If xyProfile(ii).X > xyProfile(iPalm - 1).X + (g_nPalmWristDist + nTolerance) Then Exit For
        Next ii
        iVertex = ii

       'One to One mapping from wrist onwards
        For ii = g_iWristNo To g_iLastTape
            g_iVertexMap(ii) = iVertex
            iVertex = iVertex + 1
        Next ii

    Else
       'For ordinary arms we assume a 1 to 1 mapping
        iVertex = 1
        For ii = g_iFirstTape To g_iLastTape
            g_iVertexMap(ii) = iVertex
            iVertex = iVertex + 1
        Next ii
       'Check to see if the user has inserted extra points
       'Warn and exit, as this will cause errors
        If (g_iLastTape - g_iFirstTape) > g_iProfile Then
            MsgBox "There are missing points in the Plain Arm profile.  The editor will not work!  Redraw arm or insert missing point", 16, "Arm EDIT"
            End
        End If
        If (g_iLastTape - g_iFirstTape) < g_iProfile Then
            MsgBox "There are extra points in the Plain Arm profile.  The editor will not work!  Remove extra points and try again", 16, "Arm EDIT"
            End
        End If
    End If

    

   'DRAFIX only allows the creation by a macro of an Open Fitted curve if the curve
   'has more than 3 points.
   'Therefor the method of duplicating the curve to display the results is
   'not available.  Edits will be applied directly to the curve
    If g_iProfile = 2 Then g_ShortArm = True


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


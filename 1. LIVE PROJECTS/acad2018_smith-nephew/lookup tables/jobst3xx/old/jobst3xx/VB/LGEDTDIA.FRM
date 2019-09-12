VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form lgedtdia 
   Caption         =   "Loading - EDIT"
   ClientHeight    =   7908
   ClientLeft      =   636
   ClientTop       =   780
   ClientWidth     =   4044
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
   LinkTopic       =   "lgedtdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7908
   ScaleWidth      =   4044
   Begin VB.Timer Timer1 
      Left            =   4140
      Top             =   5400
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   75
      Top             =   1200
      Width           =   435
   End
   Begin VB.TextBox txtUIDTempLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   90
      Top             =   4890
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUIDCurv 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   92
      Top             =   4530
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton Finished 
      Caption         =   "Finished"
      Height          =   375
      Left            =   360
      TabIndex        =   118
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Draw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   1500
      TabIndex        =   120
      Top             =   7320
      Width           =   1035
   End
   Begin VB.TextBox txtUnits 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   73
      Top             =   840
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtFileNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   74
      Top             =   480
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   82
      Top             =   4140
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftRed 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   127
      Top             =   3720
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftStretch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   128
      Top             =   3360
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   129
      Top             =   3000
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Left 
      Height          =   7185
      Left            =   120
      TabIndex        =   126
      Top             =   0
      Width           =   3795
      Begin MSFlexGridLib.MSFlexGrid grdLeftDisplay 
         Height          =   6495
         Left            =   3120
         TabIndex        =   131
         Top             =   570
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   11451
         _Version        =   65541
         Rows            =   24
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         BackColorBkg    =   -2147483633
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grdLeftInches 
         Height          =   8120
         Left            =   1320
         TabIndex        =   130
         Top             =   -1050
         Width           =   700
         _ExtentX        =   1228
         _ExtentY        =   14330
         _Version        =   65541
         Rows            =   30
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483648
         BackColorBkg    =   -2147483633
         GridLines       =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   6
         Left            =   2640
         TabIndex        =   49
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   7
         Left            =   2640
         TabIndex        =   50
         Top             =   810
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   8
         Left            =   2640
         TabIndex        =   51
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   9
         Left            =   2640
         TabIndex        =   52
         Top             =   1350
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   10
         Left            =   2640
         TabIndex        =   53
         Top             =   1620
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   11
         Left            =   2640
         TabIndex        =   54
         Top             =   1890
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   12
         Left            =   2640
         TabIndex        =   55
         Top             =   2160
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   13
         Left            =   2640
         TabIndex        =   56
         Top             =   2430
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   14
         Left            =   2640
         TabIndex        =   57
         Top             =   2700
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   15
         Left            =   2640
         TabIndex        =   58
         Top             =   2970
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   16
         Left            =   2640
         TabIndex        =   59
         Top             =   3240
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   17
         Left            =   2640
         TabIndex        =   60
         Top             =   3510
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   18
         Left            =   2640
         TabIndex        =   61
         Top             =   3780
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   19
         Left            =   2640
         TabIndex        =   62
         Top             =   4050
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   20
         Left            =   2640
         TabIndex        =   63
         Top             =   4320
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   21
         Left            =   2640
         TabIndex        =   64
         Top             =   4590
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   22
         Left            =   2640
         TabIndex        =   65
         Top             =   4860
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   23
         Left            =   2640
         TabIndex        =   66
         Top             =   5130
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   24
         Left            =   2640
         TabIndex        =   67
         Top             =   5400
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   25
         Left            =   2640
         TabIndex        =   68
         Top             =   5670
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   26
         Left            =   2640
         TabIndex        =   69
         Top             =   5940
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   27
         Left            =   2640
         TabIndex        =   70
         Top             =   6230
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   28
         Left            =   2640
         TabIndex        =   71
         Top             =   6480
         Width           =   435
      End
      Begin VB.TextBox txtLeftStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   29
         Left            =   2640
         TabIndex        =   72
         Top             =   6750
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   29
         Left            =   2100
         TabIndex        =   48
         Top             =   6750
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   28
         Left            =   2100
         TabIndex        =   47
         Top             =   6480
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   27
         Left            =   2100
         TabIndex        =   46
         Top             =   6210
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   26
         Left            =   2100
         TabIndex        =   45
         Top             =   5940
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   25
         Left            =   2100
         TabIndex        =   44
         Top             =   5670
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   24
         Left            =   2100
         TabIndex        =   43
         Top             =   5400
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   23
         Left            =   2100
         TabIndex        =   42
         Top             =   5130
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   22
         Left            =   2100
         TabIndex        =   41
         Top             =   4860
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   21
         Left            =   2100
         TabIndex        =   40
         Top             =   4590
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   20
         Left            =   2100
         TabIndex        =   39
         Top             =   4320
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   19
         Left            =   2100
         TabIndex        =   38
         Top             =   4050
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   18
         Left            =   2100
         TabIndex        =   37
         Top             =   3780
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   17
         Left            =   2100
         TabIndex        =   36
         Top             =   3510
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   16
         Left            =   2100
         TabIndex        =   35
         Top             =   3240
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   15
         Left            =   2100
         TabIndex        =   34
         Top             =   2970
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   14
         Left            =   2100
         TabIndex        =   33
         Top             =   2700
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   13
         Left            =   2100
         TabIndex        =   32
         Top             =   2430
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   12
         Left            =   2100
         TabIndex        =   31
         Top             =   2160
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   11
         Left            =   2100
         TabIndex        =   30
         Top             =   1890
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   10
         Left            =   2100
         TabIndex        =   29
         Top             =   1620
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   9
         Left            =   2100
         TabIndex        =   28
         Top             =   1350
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   8
         Left            =   2100
         TabIndex        =   27
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   7
         Left            =   2100
         TabIndex        =   26
         Top             =   810
         Width           =   435
      End
      Begin VB.TextBox txtLeftMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   6
         Left            =   2100
         TabIndex        =   25
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   660
         MaxLength       =   7
         TabIndex        =   0
         Top             =   270
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   660
         MaxLength       =   7
         TabIndex        =   1
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   660
         MaxLength       =   7
         TabIndex        =   2
         Top             =   810
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   660
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   9
         Left            =   660
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1350
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   10
         Left            =   660
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1620
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   11
         Left            =   660
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1890
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   12
         Left            =   660
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   13
         Left            =   660
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2430
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   14
         Left            =   660
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2700
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   15
         Left            =   660
         MaxLength       =   7
         TabIndex        =   10
         Top             =   2970
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   16
         Left            =   660
         MaxLength       =   7
         TabIndex        =   11
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   17
         Left            =   660
         MaxLength       =   7
         TabIndex        =   12
         Top             =   3510
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   18
         Left            =   660
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3780
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   19
         Left            =   660
         MaxLength       =   7
         TabIndex        =   14
         Top             =   4050
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   20
         Left            =   660
         MaxLength       =   7
         TabIndex        =   15
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   21
         Left            =   660
         MaxLength       =   7
         TabIndex        =   16
         Top             =   4590
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   22
         Left            =   660
         MaxLength       =   7
         TabIndex        =   17
         Top             =   4860
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   23
         Left            =   660
         MaxLength       =   7
         TabIndex        =   18
         Top             =   5130
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   24
         Left            =   660
         MaxLength       =   7
         TabIndex        =   19
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   25
         Left            =   660
         MaxLength       =   7
         TabIndex        =   20
         Top             =   5670
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   26
         Left            =   660
         MaxLength       =   7
         TabIndex        =   21
         Top             =   5940
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   27
         Left            =   660
         MaxLength       =   7
         TabIndex        =   22
         Top             =   6210
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   28
         Left            =   660
         MaxLength       =   7
         TabIndex        =   23
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   29
         Left            =   660
         MaxLength       =   7
         TabIndex        =   24
         Top             =   6750
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   1920
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Caption         =   "Red."
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   77
         Top             =   300
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Alignment       =   2  'Center
         Caption         =   "Str"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   84
         Top             =   300
         Width           =   435
      End
      Begin VB.Label labLeftDisp 
         Caption         =   "MMHg"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   85
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   93
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "1½"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   94
         Top             =   570
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   95
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "4½"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   96
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   97
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7½"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   98
         Top             =   1650
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   99
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "10½"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   100
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "12"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   101
         Top             =   2460
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "13½"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   102
         Top             =   2730
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "15"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   103
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "16½"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   104
         Top             =   3270
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "18"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   105
         Top             =   3540
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "19½"
         Height          =   195
         Index           =   18
         Left            =   180
         TabIndex        =   106
         Top             =   3810
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "21"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   107
         Top             =   4080
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "22½"
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   108
         Top             =   4350
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   195
         Index           =   21
         Left            =   180
         TabIndex        =   109
         Top             =   4620
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "25½"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   110
         Top             =   4890
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "27"
         Height          =   195
         Index           =   23
         Left            =   180
         TabIndex        =   111
         Top             =   5160
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "28½"
         Height          =   195
         Index           =   24
         Left            =   180
         TabIndex        =   112
         Top             =   5430
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   195
         Index           =   25
         Left            =   180
         TabIndex        =   113
         Top             =   5700
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "31½"
         Height          =   195
         Index           =   26
         Left            =   180
         TabIndex        =   114
         Top             =   5970
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "33"
         Height          =   195
         Index           =   27
         Left            =   180
         TabIndex        =   115
         Top             =   6240
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "34½"
         Height          =   195
         Index           =   28
         Left            =   180
         TabIndex        =   116
         Top             =   6510
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "36"
         Height          =   195
         Index           =   29
         Left            =   180
         TabIndex        =   117
         Top             =   6780
         Width           =   315
      End
   End
   Begin VB.TextBox txtLeftAnkle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   81
      Top             =   2280
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtLeftLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   79
      Top             =   1920
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtUidLeftLeg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      TabIndex        =   76
      Top             =   1560
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2700
      TabIndex        =   122
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fabric:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   91
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Temp UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   123
      Top             =   4920
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leg Curv UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   124
      Top             =   4560
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "File No:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   87
      Top             =   540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Units:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   119
      Top             =   870
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "leg"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   89
      Top             =   4170
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Red"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   88
      Top             =   3750
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Str"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   121
      Top             =   3390
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left MM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   125
      Top             =   3030
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
      Left            =   4200
      TabIndex        =   86
      Top             =   180
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Ankle"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   83
      Top             =   2310
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Lengths"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   80
      Top             =   1950
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Leg UID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      TabIndex        =   78
      Top             =   1590
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "lgedtdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   LGEDTDIA.D
'Purpose:   Editor for Waist Height with two legs.
'           For legs drawn with JOBSTEX material and
'           gradient pressure
'
'Version:   1.01
'Date:      13.Jan.94
'Author:    Gary George
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Jan 99     GG      Ported to VB5
'
'Note:-
'
'    Much of the code and form has been hacked from
'    WHFIGURE.FRM and WHLEGDIA.FRM
'    As both of these are proven.
'
'    Reference to the left leg should be taken to mean
'    the current leg
'

'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)

    Option Explicit

Private Sub Cancel_Click()
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

Private Sub Draw_Click()
   'Check that data is all present and insert into drafix
    If Validate_Data() Then
        PR_DrawProfileEdits
        AppActivate fnGetDrafixWindowTitleText()
        g_ReDrawn = True
        SendKeys "@C:\JOBST\DRAW.D{enter}", True
    End If
End Sub

Private Sub Finished_Click()
   'Check that data is all present and commit into drafix
    If Validate_Data() Then
        PR_DrawAndCommitProfileEdits
        AppActivate fnGetDrafixWindowTitleText()
        SendKeys "@C:\JOBST\DRAW.D{enter}", True
        End
    End If

End Sub

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

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Leg Editing Macro created - "; Date$; "  "; Time$
    Print #fNum, "//by Visual Basic"

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
   Dim iMMValue As Integer, iStrValue As Integer, iRedValue As Integer
   Dim nValue As Double, nStretch As Double
   Dim sMessage As String, iLastStyleFound As Integer
   Dim sSpacer

  'Disable TIMEOUT
   Timer1.Enabled = False

  'Units
   If txtUnits = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If


  'Leg Option buttons
   If txtLeg.Text = "Left" Then
        lgedtdia.Caption = "LEG Edit - Left [" & txtFileNo & "]"
   End If
   If txtLeg.Text = "Right" Then
        lgedtdia.Caption = "LEG Edit - Right [" & txtFileNo & "]"
   End If

  'From the AnkleString extract the saved ankle figuring
   g_iLtAnkle = fnGetNumber(txtLeftAnkle.Text, 1) - 1
   g_iLtLastMM = fnGetNumber(txtLeftAnkle.Text, 2)
   g_iLtLastStretch = fnGetNumber(txtLeftAnkle.Text, 3)
   g_iLtLastZipper = fnGetNumber(txtLeftAnkle.Text, 7)
   
  'Fabric
   g_iLtLastFabric = Val(Left$(txtFabric.Text, 2))
   
   g_iLtMM(g_iLtAnkle) = g_iLtLastMM
   g_iLtRed(g_iLtAnkle) = fnGetNumber(txtLeftAnkle.Text, 4)
   g_iLtStretch(g_iLtAnkle) = g_iLtLastStretch
   
  'Enable Display of the extracted values
  'Set display value depending on Ankle position
   If g_iLtAnkle = 7 Then
        txtLeftMM(6).Visible = False
        txtLeftStr(6).Visible = False
   End If
    
        
  'Display ankle values
   txtLeftMM(g_iLtAnkle) = g_iLtLastMM
   txtLeftMM(g_iLtAnkle).Enabled = False
   txtLeftMM(g_iLtAnkle).BackColor = &HFFFFFF

   txtLeftStr(g_iLtAnkle).BackColor = &HFFFFFF
   txtLeftStr(g_iLtAnkle) = g_iLtStretch(g_iLtAnkle)
   txtLeftStr(g_iLtAnkle).Enabled = False
   grdLeftDisplay.Col = 0
   grdLeftDisplay.Row = g_iLtAnkle - 6 'Display from ankle
   grdLeftDisplay = Str$(g_iLtRed(g_iLtAnkle))

  'Update tape boxes
  'Get first and last tapes in the text boxes
   g_iFirstTape = -1
   g_iLastTape = -1

   nValue = 0
   grdLeftInches.Col = 0
   For ii = 5 To 29
        nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
        If nValue > 0 Then
            txtLeft(ii) = nValue
            nValue = fnDisplaytoInches(nValue)
            grdLeftInches.Row = ii
            grdLeftInches.Text = fnInchesToText(nValue)
        End If
        g_nLtLengths(ii) = nValue
        If ii <= g_iLtAnkle Then txtLeft(ii).Enabled = False
        If ii = 5 Then g_nLtLastHeel = nValue
        If ii = g_iLtAnkle Then g_nLtLastAnkle = nValue
        If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii
        If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii - 1
   Next ii
   If nValue > 0 Then g_iLastTape = 29

   g_JOBSTEX_FL = True
  
  'Initialise display from stored data
   For ii = g_iLtAnkle + 1 To 29
       iMMValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
       iStrValue = Val(Mid(txtLeftStretch.Text, (ii * 3) + 1, 3))
       iRedValue = Val(Mid(txtLeftRed.Text, (ii * 3) + 1, 3))
       If iMMValue > 0 Then
            txtLeftMM(ii) = iMMValue
            g_iLtMM(ii) = iMMValue
            txtLeftStr(ii) = iStrValue
            g_iLtStretch(ii) = iStrValue
            grdLeftDisplay.Row = ii - 6
            grdLeftDisplay = Str$(iRedValue)
            g_iLtRed(ii) = iRedValue
       End If
   Next ii

  'Store inital values for comparision and for use in cancel command
   For ii = 0 To 29
        g_iLtStretchInit(ii) = g_iLtStretch(ii)
        g_iLtRedInit(ii) = g_iLtRed(ii)
        g_iLtMMInit(ii) = g_iLtMM(ii)
        g_nLtLengthsInit(ii) = g_nLtLengths(ii)
        g_iLtChanged(ii) = 0
   Next ii

   g_ReDrawn = False

   PR_GetProfileFromFile "C:\JOBST\LEGCURVE.DAT"
   
  'Disable Tape above fold Ht
   For ii = g_iLtAnkle + (g_iLastEditTape - g_iFirstEditTape) + 2 To 29
        txtLeft(ii).Enabled = False
        txtLeftMM(ii).Enabled = False
        txtLeftMM(ii).BackColor = &HFFFFFF
        txtLeftStr(ii).Enabled = False
        txtLeftStr(ii).BackColor = &HFFFFFF
   Next ii

   Show
   Screen.MousePointer = 0    'Change pointer to default.
   Cancel.SetFocus

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

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - lgedtdia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - lgedtdia.Height) / 2  ' Center form vertically.

    Set MainForm = lgedtdia

    g_nUnitsFac = 1 'Default to inches
    g_JOBSTEX_FL = True
    g_sTextList = "-7½ -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

   'Setup display inches grid
    grdLeftInches.ColWidth(0) = 700
    grdLeftInches.ColAlignment(0) = 2
    For ii = 0 To 29
        grdLeftInches.RowHeight(ii) = 266
    Next ii

   'Setup display of results grid
    grdLeftDisplay.ColWidth(0) = 488
    grdLeftDisplay.ColAlignment(0) = 2
    For ii = 0 To 23
        grdLeftDisplay.RowHeight(ii) = 266
    Next ii

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "The Edit Module is already running!" + Chr(13) + "Use ALT-TAB to access it .", 16, "Leg Edit Warning"
        End
    End If

    g_sPathJOBST = fnPathJOBST()

   'Set up to \timeout after 6 seconds
   'This is disabled in link close
    Timer1.Interval = 6000
    Timer1.Enabled = True

End Sub

Private Sub PR_CancelProfileEdits()

    If g_ReDrawn = True Then
        fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
        PR_PutLine "HANDLE hTempCurv;"
       
        'Delete Curve Copy
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempLeg.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"

        Close fNum
    End If

End Sub

Private Sub PR_DeleteTapeLabel(xyPt As xy)
    'Deletes the tape label and the text at the given point
    'Uses a marquee to do so
     Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double

     x1 = xyPt.X - 0.5
     x2 = xyPt.X + 0.5
     y1 = xyPt.y - 1.5
     y2 = xyPt.y + 0.5

     Print #fNum, "hChan=Open("; QQ; "selection"; QCQ; " ( type = 'Text' OR type = 'Symbol' ) AND layer = 'Construct' AND TOTally INside "; x1; y1; x2; y2; QQ; ");"
     Print #fNum, "if(hChan)"
     Print #fNum, "{ResetSelection(hChan);while(hEnt=GetNextSelection(hChan))DeleteEntity(hEnt);}"
     Print #fNum, "Close("; QQ; "selection"; QC; "hChan);"

End Sub

Private Sub PR_DrawAndCommitProfileEdits()

   'Create a macro to commit the changes to the Original profile
    
    Dim xyPt As xy
    Dim ii As Integer, iTempLegVertex As Integer, iProfileVertex As Integer
    Dim iStart As Integer, iEnd As Integer
    Dim ProfileChanged As Integer
    Dim nSeam As Double
    Dim sLen As String, sMM As String, sStr As String, sRed As String
    Dim sLeftLengths As String, sLeftMMs As String, sLeftStr As String, sLeftRed As String
    Dim nLen As Double, iMM As Integer, iStr  As Integer, iRed  As Integer
    Dim sLine As String

    nSeam = 0.1875

    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
    PR_PutLine "HANDLE  hCurv, hTempCurv, hLeg, hChan, hEnt;"
       
    sLine = FN_EscapeSlashesInString(g_sPathJOBST) & "\\JOBST.SLB"

    Print #fNum, "SetSymbolLibrary("; QQ; sLine; QQ; ");"
    PR_SetLayer "Construct"

    iStart = g_iFirstEditTape - 3
    If iStart > g_iFirstEditTape Then iStart = g_iFirstEditTape
        
    iEnd = g_iLastEditTape + 3
    If iEnd > g_iLegProfile Then iEnd = g_iLegProfile

   'Delete Curve Copy (Only if redraw has been used)
    If g_ReDrawn = True Then
        Print #fNum, "hTempCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempLeg.Text); ");"
        Print #fNum, "DeleteEntity(hTempCurv);"
    End If

   'Modify Original curve  (need not have been redrawn first)
    Print #fNum, "hCurv = UID ("; QQ; "find"; QC; Val(txtUIDCurv.Text); ");"
    ProfileChanged = False
    For ii = g_iLtAnkle + 1 To g_iLastTape
        If g_iLtChanged(ii) < 0 Or g_iLtChanged(ii) > 0 Then
            iProfileVertex = ((ii - g_iLtAnkle) + g_iFirstEditTape)
            xyPt.X = xyLegProfile(iProfileVertex - 1).X
            xyPt.y = ((g_nLtLengths(ii) * (Abs(g_iLtRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
            Print #fNum, "SetVertex(hCurv,"; iProfileVertex; ","; xyPt.X; ","; xyPt.y; ");"
            PR_DeleteTapeLabel xyLegProfile(iProfileVertex - 1)
            PR_DrawTapeLabel xyPt, ii
            ProfileChanged = True
        End If
    Next ii

   'Update Leg Box (only if Profile has been changed)
    If ProfileChanged = True Then
        sLeftLengths = Left(txtLeftLengths.Text, 20)
        For ii = 5 To 29
            nLen = Val(txtLeft(ii).Text)

            If nLen <> 0 Then
                nLen = nLen * 10 'Shift decimal place to right
                sLen = String(4, " ")
                RSet sLen = Trim$(Str$(nLen))
            Else
                sLen = String(4, " ")
            End If
        sLeftLengths = sLeftLengths & sLen

        Next ii

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

            sLeftMMs = sLeftMMs & sMM
            sLeftRed = sLeftRed & sRed
            sLeftStr = sLeftStr & sStr

        Next ii


        Print #fNum, "hLeg = UID ("; QQ; "find"; QC; Val(txtUidLeftLeg.Text); ");"
        Print #fNum, "if (!hLeg)Exit(%cancel,"; QQ; "Can't find LEGBOX to Update"; QQ; ");"

        Print #fNum, "SetDBData( hLeg, "; QQ; "TapeLengthsPt1"; QCQ; Mid$(sLeftLengths, 1, 60); QQ; ");"
        Print #fNum, "SetDBData( hLeg, "; QQ; "TapeLengthsPt2"; QCQ; Mid$(sLeftLengths, 61, 60); QQ; ");"
    
        Print #fNum, "SetDBData( hLeg, "; QQ; "TapeMMs"; QCQ; Mid$(sLeftMMs, 1, 60); QQ; ");"
        Print #fNum, "SetDBData( hLeg, "; QQ; "TapeMMs2"; QCQ; Mid$(sLeftMMs, 61, 60); QQ; ");"
    
        Print #fNum, "SetDBData( hLeg, "; QQ; "Grams"; QCQ; Mid$(sLeftStr, 1, 60); QQ; ");"
        Print #fNum, "SetDBData( hLeg, "; QQ; "Grams2"; QCQ; Mid$(sLeftStr, 61, 60); QQ; ");"
    
        Print #fNum, "SetDBData( hLeg, "; QQ; "Reduction"; QCQ; Mid$(sLeftRed, 1, 60); QQ; ");"
        Print #fNum, "SetDBData( hLeg, "; QQ; "Reduction2"; QCQ; Mid$(sLeftRed, 61, 60); QQ; ");"
    
    End If
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "ViewRedraw"; QQ; ");"

    Close fNum

End Sub

Private Sub PR_DrawProfileEdits()
   'Create a macro to copy part of the leg profile and
   'modify this copy
   'Modifications will only be committed to the original profile on "Finish"
    
    Dim xyPt   As xy
    Dim ii As Integer, iTempLegVertex As Integer, iProfileVertex As Integer
    Dim iStart As Integer, iEnd As Integer
    Dim nSeam As Double

    nSeam = 0.1875

    fNum = FN_DrawOpen("C:\JOBST\DRAW.D", "EDIT", "Test", txtLeg)
    PR_PutLine "HANDLE  hCurv;"
       
    iStart = g_iFirstEditTape - 3
    If iStart > g_iFirstEditTape Then iStart = g_iFirstEditTape
        
    iEnd = g_iLastEditTape + 3
    If iEnd > g_iLegProfile Then iEnd = g_iLegProfile

    If g_ReDrawn <> True Then
       'Make Copy of curve
        PR_SetLayer "Construct"
        Print #fNum, "hCurv = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        
        For ii = iStart To iEnd
            Print #fNum, CC; xyLegProfile(ii).X; CC; xyLegProfile(ii).y;
        Next ii
        
        Print #fNum, ");"

    Else
       'Modify Curve Copy
        If txtUIDTempLeg.Text <> "" Then
            Print #fNum, "hCurv = UID ("; QQ; "find"; QC; Val(txtUIDTempLeg.Text); ");"
        Else
            
        End If
    End If

    For ii = g_iLtAnkle + 1 To g_iLastTape
        If g_iLtChanged(ii) > 0 Then
            g_iLtChanged(ii) = -1
            iTempLegVertex = (ii - g_iLtAnkle) + (g_iFirstEditTape - iStart)
            iProfileVertex = ((ii - g_iLtAnkle) + g_iFirstEditTape) - 1
            If iTempLegVertex > iEnd Then Exit For
            xyPt.X = xyLegProfile(iProfileVertex).X
            xyPt.y = ((g_nLtLengths(ii) * (Abs(g_iLtRed(ii) - 100) / 100)) / 2) + nSeam + xyOtemplate.y
            Print #fNum, "SetVertex(hCurv,"; iTempLegVertex; ","; xyPt.X; ","; xyPt.y; ");"
        End If
    Next ii

    If g_ReDrawn <> True Then
        PR_PutLine "HANDLE  hDDE;"

       'Poke Curve UID back to the VB program
        Print #fNum, "hDDE = Open ("; QQ; "dde"; QCQ; "lgedtdia"; QCQ; "lgedtdia"; QQ; ");"
        Print #fNum, "Poke ( hDDE, "; QQ; "txtUIDTempLeg"; QC; "MakeString("; QQ; "long"; QC; "UID("; QQ; "get"; QQ; ",hCurv)));"

    End If
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "ViewRedraw"; QQ; ");"

    Close fNum

End Sub

Private Sub PR_DrawTapeLabel(xyPt As xy, Index As Integer)
    Dim sSymbol As String, sText

    sSymbol = Trim$(Str$(Index)) & "tape"
    Print #fNum, "Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ");"
    Print #fNum, "hEnt = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; xyPt.X; CC; xyPt.y; ");"

    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; "2);"
    sText = fnInchesToText(g_nLtLengths(Index))
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; xyPt.X; CC; xyPt.y - 0.5; CC; "0.075"; CC; "0.125"; ",0);"
    
    sText = Trim$(Str$(g_iLtRed(Index)))
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; xyPt.X; CC; xyPt.y - 0.7; CC; "0.075"; CC; "0.125"; ",0);"
    
    sText = Trim$(Str$(g_iLtStretch(Index)))
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; xyPt.X; CC; xyPt.y - 0.9; CC; "0.075"; CC; "0.125"; ",0);"
    
    sText = Trim$(Str$(g_iLtMM(Index)))
    Print #fNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; xyPt.X; CC; xyPt.y - 1.1; CC; "0.075"; CC; "0.125"; ",0);"


End Sub

Private Sub PR_FigureLeftTape(Index As Integer, Force As Integer)
   'Figures the Pressure/Stretch at a LEFT LegTape
    
    Dim nMMHg As Double, nStretch As Double
    Dim nTapeCir As Double
    Dim UseMMHg As Integer

    grdLeftDisplay.Col = 0
    grdLeftDisplay.Row = Index - 6 'Display from ankle

   'Establish if enough data exists to calculate.
   'If not then EXIT quietly
    If Val(txtLeftMM(Index).Text) <= 0 And Val(txtLeftStr(Index).Text) <= 0 Then
        Beep
        grdLeftDisplay = ""
       'Select_Text_in_Box txtLeftMM(Index)
        Exit Sub
    End If

   'Establish if pressure or stretch has changed from last time
   'If it hasn't then exit
   '
    If Val(txtLeftMM(Index).Text) = g_iLtMM(Index) And Val(txtLeftStr(Index).Text) = g_iLtStretch(Index) And Force = False Then Exit Sub
    
   'MMHg and Stretch for JOBSTEX fabric
   'Don't Calculate if no values for Ankle tape
    If Val(txtLeftMM(g_iLtAnkle)) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub
    
   'If we get to here we can assume that there is enough data to process and
   'we can calculate the relevant values.
    
    nTapeCir = fnDisplaytoInches(Val(txtLeft(Index).Text))
    nMMHg = Val(txtLeftMM(Index).Text)
    nStretch = Val(txtLeftStr(Index).Text)
   
   'Establish if caclculation is based on given Stretch or Pressure (MMhg)
    UseMMHg = -1
   
   'Simple case where one of the values is missing
    If nMMHg <= 0 And nStretch > 0 Then UseMMHg = 0
    If nStretch <= 0 And nMMHg > 0 Then UseMMHg = 1

    If nStretch > 0 And nMMHg > 0 Then
        If nMMHg <> g_iLtMM(Index) And nStretch <> g_iLtStretch(Index) Then
           'Changes in Pressure take precedence
            UseMMHg = 1
        Else
            If nMMHg <> g_iLtMM(Index) Then
                UseMMHg = 1
            Else
                UseMMHg = 0
            End If
        End If
    End If
   
   'Exit if UseMMHg flag not set
    If UseMMHg < 0 Then Exit Sub
   
   'Calculate using Pressure solving for stretch
    If UseMMHg = 1 Then
        nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, g_iLtLastZipper, g_iLtLastFabric)
        If nStretch <= 0 Then
            Beep
            grdLeftDisplay = ""
            txtLeftStr(Index) = ""
            Select_Text_in_Box txtLeftMM(Index)
            Exit Sub
        End If
    Else
        nMMHg = FN_JOBSTEX_Pressure(nTapeCir, g_nLtLastHeel, nStretch, g_iLtLastZipper, g_iLtLastFabric)
        If nMMHg <= 0 Then
            Beep
            grdLeftDisplay = ""
            txtLeftMM(Index) = ""
            Select_Text_in_Box txtLeftStr(Index)
            Exit Sub
        End If
    End If

   'Store values
    g_iLtStretch(Index) = round(nStretch)
    g_iLtRed(Index) = round((1 - 1 / (0.01 * nStretch + 1)) * 100)
    g_iLtMM(Index) = nMMHg

   'Display values
    txtLeftMM(Index) = g_iLtMM(Index)
    txtLeftStr(Index) = g_iLtStretch(Index)
    grdLeftDisplay = Str$(g_iLtRed(Index))

   'Flag Change position
    g_iLtChanged(Index) = 1

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
    Print #fNum, "if ( hLayer > %zero && hLayer != 32768)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Private Sub Timer1_Timer()
    'This event means:-
    'That the DDE link to Drafix has not been established
    'therefore we End.
     End
End Sub

Private Sub txtLeft_GotFocus(Index As Integer)
    Select_Text_in_Box txtLeft(Index)
End Sub

Private Sub txtLeft_LostFocus(Index As Integer)
    Validate_And_Display_Text_In_Box txtLeft(Index), grdLeftInches, Index
End Sub

Private Sub txtLeftMM_GotFocus(Index As Integer)
    'Select_Text_in_Box txtLeftMM(Index)
End Sub

Private Sub txtLeftMM_LostFocus(Index As Integer)
    If Index <> g_iLtAnkle Then
        PR_FigureLeftTape Index, 0
    End If
End Sub

Private Sub txtLeftStr_GotFocus(Index As Integer)
    'Select_Text_in_Box txtLeftStr(Index)
End Sub

Private Sub txtLeftStr_LostFocus(Index As Integer)
    If Index <> g_iLtAnkle Then
        PR_FigureLeftTape Index, 0
    End If
End Sub

Private Sub Validate_And_Display_Text_In_Box(Text_Box_Name As Control, Display_Control As Control, Index As Integer)
   'Subroutine that is activated when the focus is lost.
   'Displays value in inches in the grid given
   '
   'If Value is changed then
   '    Modifies Stored Value
   '    Recalculates tape reduction and stretch
   '
    Dim rTextBoxValue As Double
    Dim iFirstTape As Integer, iLastTape As Integer
    Dim ii As Integer, iTest As Integer

   'Checks that input data is valid.
   'If not valid then display a message and returns focus
   'to the text in question

   'Get the text value
    rTextBoxValue = fnDisplaytoInches(Val(Text_Box_Name.Text))

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

    If Index >= 0 And g_nLtLengths(Index) <> rTextBoxValue Then
       'Treat Display as Grid control
        Display_Control.Row = Index
        Display_Control.Text = fnInchesToText(rTextBoxValue)
        g_nLtLengths(Index) = rTextBoxValue
        PR_FigureLeftTape Index, True
    End If


End Sub

Private Function Validate_Data() As Integer
    
    Dim sError As String, NL As String
    Dim ii As Integer, nFirstTape As Integer, nLastTape As Integer
    Dim iError As Integer
    Dim nValue As Double

    NL = Chr$(10) 'new line
   
   'Check Tape length data text boxes for holes (ie missing values)
   'Establish First and last tape
    
    sError = ""

    For ii = 5 To 29 Step 1
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    nFirstTape = ii

    For ii = 29 To 5 Step -1
        If Val(txtLeft(ii).Text) > 0 Then Exit For
    Next ii
    nLastTape = ii

    For ii = nFirstTape To nLastTape
        If Val(txtLeft(ii).Text) = 0 Then
            sError = sError + "Missing Tape length - " & LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3)) + NL
        End If
    Next ii
   
   'Check that a minimum of 2 tapes are given
    If (nLastTape - nFirstTape) = 0 Then
        sError = sError + "More than 1 tape must be given." + NL
    End If

    If (nLastTape - nFirstTape) < 0 Then
        sError = sError + "No Tape values given." + NL
    End If

   'Display Error message (if required) and return
    If Len(sError) > 0 Then
        MsgBox sError, 48, "Errors in Data"
        Validate_Data = False
        Exit Function
    Else
        Validate_Data = True
    End If


End Function


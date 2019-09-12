VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form webspacr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Spacers - Draw"
   ClientHeight    =   9024
   ClientLeft      =   792
   ClientTop       =   252
   ClientWidth     =   8196
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000040&
   LinkMode        =   1  'Source
   LinkTopic       =   "webspacr"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9024
   ScaleWidth      =   8196
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   3615
      _ExtentX        =   6371
      _ExtentY        =   15473
      _Version        =   327681
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Hand Data"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Above Wrist"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   117
         Top             =   120
         Width           =   3345
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   12
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   123
            Top             =   2100
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   23
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   134
            Top             =   5070
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   22
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   133
            Top             =   4800
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   21
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   132
            Top             =   4530
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   20
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   131
            Top             =   4260
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   19
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   130
            Top             =   3990
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   18
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   129
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   17
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   128
            Top             =   3450
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   127
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   15
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   126
            Top             =   2910
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   14
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   125
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   13
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   124
            Top             =   2370
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   11
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   122
            Top             =   1830
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat2 
            Height          =   285
            Left            =   360
            TabIndex        =   138
            Top             =   4350
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat2 
            Height          =   285
            Left            =   360
            TabIndex        =   137
            Top             =   2070
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat1 
            Height          =   285
            Left            =   360
            TabIndex        =   139
            Top             =   5040
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat1 
            Height          =   285
            Left            =   360
            TabIndex        =   136
            Top             =   1350
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   10
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   121
            Top             =   1560
            Width           =   495
         End
         Begin VB.ComboBox cboProximalTape 
            Height          =   315
            Left            =   2340
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   5970
            Width           =   735
         End
         Begin VB.ComboBox cboDistalTape 
            Height          =   315
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   5970
            Width           =   735
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   8
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   119
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   9
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   120
            Top             =   1290
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid grdInches 
            Height          =   4335
            Left            =   2160
            TabIndex        =   118
            Top             =   1000
            Width           =   615
            _ExtentX        =   1080
            _ExtentY        =   7662
            _Version        =   65541
            Rows            =   16
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
         Begin VB.Label lblArm 
            Caption         =   "Web Spacer To Elbow"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   165
            Top             =   450
            Width           =   2175
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   164
            Top             =   810
            Width           =   555
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "19½"
            Height          =   255
            Index           =   23
            Left            =   1020
            TabIndex        =   163
            Top             =   5130
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "18"
            Height          =   255
            Index           =   22
            Left            =   1020
            TabIndex        =   162
            Top             =   4830
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "16½"
            Height          =   255
            Index           =   21
            Left            =   1020
            TabIndex        =   161
            Top             =   4590
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "15"
            Height          =   255
            Index           =   20
            Left            =   1020
            TabIndex        =   160
            Top             =   4320
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "13½"
            Height          =   255
            Index           =   19
            Left            =   1020
            TabIndex        =   159
            Top             =   4050
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "12"
            Height          =   255
            Index           =   18
            Left            =   1020
            TabIndex        =   158
            Top             =   3780
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "10½"
            Height          =   255
            Index           =   17
            Left            =   1020
            TabIndex        =   157
            Top             =   3510
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "9"
            Height          =   255
            Index           =   16
            Left            =   1020
            TabIndex        =   156
            Top             =   3270
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "7½"
            Height          =   255
            Index           =   15
            Left            =   1020
            TabIndex        =   155
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   255
            Index           =   14
            Left            =   1020
            TabIndex        =   154
            Top             =   2730
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "4½"
            Height          =   255
            Index           =   13
            Left            =   1020
            TabIndex        =   153
            Top             =   2430
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   12
            Left            =   1020
            TabIndex        =   152
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "1½"
            Height          =   255
            Index           =   11
            Left            =   1020
            TabIndex        =   151
            Top             =   1890
            Width           =   375
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   150
            Top             =   1110
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Elbow"
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   149
            Top             =   3270
            Width           =   555
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   10
            Left            =   1020
            TabIndex        =   148
            Top             =   1620
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "EOS"
            Height          =   195
            Index           =   10
            Left            =   1920
            TabIndex        =   147
            Top             =   6030
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Wrist"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   146
            Top             =   6030
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-3"
            Height          =   255
            Index           =   8
            Left            =   1020
            TabIndex        =   145
            Top             =   1050
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-1½"
            Height          =   255
            Index           =   9
            Left            =   1020
            TabIndex        =   144
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   143
            Top             =   1830
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   142
            Top             =   4110
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   140
            Top             =   4800
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   8145
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   3345
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   55
            Top             =   4140
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   54
            Top             =   3840
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   64
            Top             =   7620
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   56
            Top             =   5100
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   57
            Top             =   5400
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   58
            Top             =   5700
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   59
            Top             =   6000
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   60
            Top             =   6300
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   61
            Top             =   6720
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   62
            Top             =   7020
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   63
            Top             =   7320
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   53
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   52
            Top             =   3240
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   51
            Top             =   2940
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   50
            Top             =   2640
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   49
            Top             =   2340
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   48
            Top             =   2040
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   47
            Top             =   1740
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   46
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   45
            Top             =   1140
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   44
            Top             =   840
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   43
            Top             =   540
            Width           =   555
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   12
            Left            =   2460
            TabIndex        =   116
            Top             =   4170
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   11
            Left            =   2460
            TabIndex        =   115
            Top             =   3900
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   10
            Left            =   2460
            TabIndex        =   114
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   2
            Left            =   2460
            TabIndex        =   113
            Top             =   5730
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "3"" Beyond Wrist"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   112
            Top             =   4200
            Width           =   1560
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "1½"" Beyond Wrist"
            Height          =   255
            Index           =   11
            Left            =   180
            TabIndex        =   111
            Top             =   3900
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Lengths"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   110
            Top             =   4740
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Circumferences"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   109
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 12"
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   108
            Top             =   5130
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 13"
            Height          =   255
            Index           =   1
            Left            =   1020
            TabIndex        =   107
            Top             =   5415
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 14"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   106
            Top             =   5715
            Width           =   975
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Index 15"
            Height          =   255
            Index           =   3
            Left            =   1020
            TabIndex        =   105
            Top             =   6030
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 16"
            Height          =   255
            Index           =   4
            Left            =   900
            TabIndex        =   104
            Top             =   6330
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 20"
            Height          =   255
            Index           =   31
            Left            =   900
            TabIndex        =   103
            Top             =   7665
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 19"
            Height          =   255
            Index           =   30
            Left            =   840
            TabIndex        =   102
            Top             =   7350
            Width           =   975
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 18"
            Height          =   255
            Index           =   29
            Left            =   1020
            TabIndex        =   101
            Top             =   7050
            Width           =   795
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 17"
            Height          =   255
            Index           =   28
            Left            =   1020
            TabIndex        =   100
            Top             =   6765
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   2  'Center
            Caption         =   "Tip To Web"
            Height          =   645
            Index           =   5
            Left            =   360
            TabIndex        =   99
            Top             =   5460
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Wrist To Web"
            Height          =   645
            Index           =   26
            Left            =   360
            TabIndex        =   98
            Top             =   6990
            Width           =   435
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   97
            Top             =   5130
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   1
            Left            =   2460
            TabIndex        =   96
            Top             =   5430
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   3
            Left            =   2460
            TabIndex        =   95
            Top             =   6030
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   4
            Left            =   2460
            TabIndex        =   94
            Top             =   6330
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   5
            Left            =   2460
            TabIndex        =   93
            Top             =   6780
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   6
            Left            =   2460
            TabIndex        =   92
            Top             =   7050
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   7
            Left            =   2460
            TabIndex        =   91
            Top             =   7350
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   8
            Left            =   2460
            TabIndex        =   90
            Top             =   7650
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Wrist 11"
            Height          =   255
            Index           =   10
            Left            =   780
            TabIndex        =   89
            Top             =   3600
            Width           =   1020
         End
         Begin VB.Label labCir 
            Caption         =   "Index Finger"
            Height          =   405
            Index           =   23
            Left            =   360
            TabIndex        =   88
            Top             =   2490
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Middle Finger"
            Height          =   405
            Index           =   22
            Left            =   360
            TabIndex        =   87
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Ring Finger"
            Height          =   405
            Index           =   21
            Left            =   360
            TabIndex        =   86
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Little Finger"
            Height          =   405
            Index           =   20
            Left            =   360
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   9
            Left            =   2460
            TabIndex        =   84
            Top             =   3315
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Palm 10"
            Height          =   255
            Index           =   9
            Left            =   780
            TabIndex        =   83
            Top             =   3285
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   8
            Left            =   2460
            TabIndex        =   82
            Top             =   2985
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb  9"
            Height          =   255
            Index           =   8
            Left            =   780
            TabIndex        =   81
            Top             =   2985
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   7
            Left            =   2460
            TabIndex        =   80
            Top             =   2685
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  8"
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   79
            Top             =   2700
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   6
            Left            =   2460
            TabIndex        =   78
            Top             =   2385
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  7"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   77
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   5
            Left            =   2460
            TabIndex        =   76
            Top             =   2085
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  6"
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   75
            Top             =   2100
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   4
            Left            =   2460
            TabIndex        =   74
            Top             =   1785
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  5"
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   73
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   3
            Left            =   2460
            TabIndex        =   72
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  4"
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   71
            Top             =   1500
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   2
            Left            =   2460
            TabIndex        =   70
            Top             =   1185
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  3"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   69
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   1
            Left            =   2460
            TabIndex        =   68
            Top             =   885
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  2"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   67
            Top             =   900
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  1"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   66
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   65
            Top             =   585
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   8700
      TabIndex        =   40
      Top             =   6930
      Width           =   675
   End
   Begin VB.TextBox txtShoulderPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8460
      TabIndex        =   33
      Text            =   "txtShoulderPleat"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWristPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8460
      TabIndex        =   39
      Text            =   "txtWristPleat"
      Top             =   4980
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengthPt1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   38
      Text            =   "txtTapeLengthPt1"
      Top             =   3150
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame9 
      Caption         =   "Options"
      Height          =   2325
      Left            =   3960
      TabIndex        =   35
      Top             =   2670
      Width           =   4035
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Text            =   "cboFabric"
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox cboInsert 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   870
         Width           =   1275
      End
      Begin VB.OptionButton optExtendTo 
         Caption         =   "Normal Web Spacer"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   1350
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optExtendTo 
         Caption         =   "Web Spacer to Elbow"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   1710
         Width           =   2355
      End
      Begin VB.Label labInsert 
         Caption         =   "Inches"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   37
         Top             =   900
         Width           =   615
      End
      Begin VB.Label labInsert 
         Caption         =   "Insert Size:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   36
         Top             =   930
         Width           =   975
      End
   End
   Begin VB.TextBox txtDataGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   34
      Text            =   "txtDataGlove"
      Top             =   4110
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   6060
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8460
      TabIndex        =   31
      Text            =   "txtUidMPD"
      Top             =   450
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSide 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   30
      Text            =   "txtSide"
      Top             =   3810
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   5820
      TabIndex        =   29
      Top             =   7200
      Width           =   1395
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   435
      Left            =   4140
      TabIndex        =   28
      Top             =   7200
      Width           =   1395
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8520
      TabIndex        =   27
      Text            =   "txtFabric"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   26
      Text            =   "txtUidGlove"
      Top             =   2550
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   25
      Text            =   "txtTapeLengths2"
      Top             =   3510
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8400
      TabIndex        =   24
      Text            =   "txtTapeLengths"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidGC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8520
      TabIndex        =   22
      Text            =   "txtUidGC"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8460
      TabIndex        =   21
      Text            =   "txtWorkOrder"
      Top             =   750
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Glove"
      Height          =   585
      Left            =   3960
      TabIndex        =   19
      Top             =   120
      Width           =   4035
      Begin VB.OptionButton optHand 
         Caption         =   "Right"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   12
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Left"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   11
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   1785
      Left            =   3960
      TabIndex        =   20
      Top             =   690
      Width           =   4035
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   660
         Width           =   3075
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   13
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   1380
         Width           =   525
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   3060
         TabIndex        =   18
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   16
         Top             =   1020
         Width           =   3075
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   2580
         TabIndex        =   8
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2700
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8580
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Glove Data"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8580
      TabIndex        =   0
      Top             =   2310
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Glove Common"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8640
      TabIndex        =   23
      Top             =   1170
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "webspacr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   WEBSPACR.MAK
'Purpose:   Draw web spacers
'
'
'Version:   3.02
'Date:      20.Sep.96
'Author:    Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Port to VB5
'
'Notes:-
'   This has been quickly cobbled from the MANUAL and
'   CAD GLOVE code so it's a bit quirky.
'   but it works and was cheap to do. (I didn't write this bit !Honest Guv!)
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

   'Open tip constants
    Const AGE_CUTOFF = 10           'Years
    Const ADULT_STD_OPEN_TIP = 0.5  'Inches
    Const CHILD_STD_OPEN_TIP = 0.375 'Inches

    Option Explicit

Private Sub Cancel_Click()
    End
End Sub

Private Function FN_CalculateWebSpacer() As Integer
   'Procedure to setup Module level variables based on the data in the
   'dialogue
   'Returns true if there is enough data to work with
   'else returns false if there isn't

    Dim nWeb#, nHighWeb#, nThumbWeb#, nPalmCir#, nWristCir#, nThumbCir#
    Dim nWebToWeb#, nStrapLengthMiddleAndIndex#, nStrapLengthLittle#, nStrapWidth#
    Dim nStrapConstruct#, nGaunletCrook#, nFigure#, nExtendThumb#
    Dim nSeam#, nInsert#, nSpace#, nLength#
    Dim nBotRadFac#, nWristNotch#
    Dim aAngle#, aThumb#, nTaperNotchRad#, nThumbOffset#, nWristOffset#
    Dim xyTmp As XY, xyTmp1 As XY
    Dim ii%, iAge%, iError%
    Dim sError$, NL$
    ReDim nWebHt(1 To 4) As Double
    Dim nThumbHt1#
   'Initialise
    NL = Chr$(13) 'New Line character
    FN_CalculateWebSpacer = False   'Default to false

   'Patients age
    iAge = Val(txtAge.Text)

   'Circumferences
    nThumbCir = fnDisplayToInches(Val(txtCir(8).Text))  'Thumb
    nPalmCir = fnDisplayToInches(Val(txtCir(9).Text))      'Palm
    nWristCir = fnDisplayToInches(Val(txtCir(10).Text))    'Wrist

   'Webs
   'The thumb web can be lowered when the thumb
   'is being calculated in the Manual glove drawing
   'programme
   'Therefor we use the value that has been calculated and
   'stored

    nHighWeb = 0
    For ii = 1 To 4
        nWeb = fnDisplayToInches(Val(txtLen(ii + 4).Text))
        If ii = 4 Then nWeb = nWeb - g_nThumbWebDrop 'Calculated thumb drop
        nWebHt(ii) = nWeb
        If nWeb > nHighWeb Then nHighWeb = nWeb
    Next ii
   'Last web is thumb web
    nThumbWeb = nWeb


   'Check on available data
    sError = ""
    If nThumbCir = 0 Then
        sError = sError & "Missing THUMB circumference!" & NL
    End If
    If nPalmCir = 0 Then
        sError = sError & "Missing PALM circumference!" & NL
    End If
    If nWristCir = 0 Then
        sError = sError & "Missing WRIST circumference!" & NL
    End If
    If nHighWeb = 0 Then
        sError = sError & "Missing WEB lengths!" & NL
    End If
    If nThumbWeb = 0 Then
        sError = sError & "Missing THUMB WEB length!" & NL
    End If

    If g_iInsertSize = 0 Then
        sError = sError & "Insert not specified!" & NL
    End If

    If g_iInsertSize < 3 Or g_iInsertSize > 8 Then
        sError = sError & "Specified insert is not found on table ""ONE INSERT WITH LARGE SEAMS"", ref GOP 01-02/28, page 4!" & NL
    End If
    
    If sError <> "" Then
        sError = "Unable to draw Web Spacer because of missing or incomplete data." & NL & sError
        MsgBox sError, 16, g_sDialogID
        FN_CalculateWebSpacer = False
        Exit Function
    End If

   'Setup figuring and construction factors
   'NB age related construction factors
    nGaunletCrook = 3 * SIXTEENTH
    nFigure = 0.86
    nSeam = 0.25
    nExtendThumb = 0.5
    nInsert = g_iInsertSize * EIGHTH
    
    If iAge > 10 Then
       'Adults
        nStrapLengthMiddleAndIndex = 1.25
        nStrapLengthLittle = 1
        nStrapWidth = 0.5
        nStrapConstruct = 0.25
    Else
       'Children
        nStrapLengthMiddleAndIndex = 1
        nStrapLengthLittle = 0.875
        nStrapWidth = 0.375
        nStrapConstruct = 0.1875
    End If

   'Figure dimensions
    nPalmCir = ((nPalmCir * nFigure) - nInsert) + QUARTER
    nWristCir = nWristCir * nFigure
    nWebToWeb = (nHighWeb - nThumbWeb) * nFigure
    nHighWeb = nHighWeb * nFigure
    nThumbCir = FN_FigureThumb(nThumbCir)
    If nThumbCir <= 0 Then
        sError = "Unable to figure the thumb with the given insert and Thumb circumference." & NL
        MsgBox sError, 16, g_sDialogID
        FN_CalculateWebSpacer = False
        Exit Function
    End If
   
   'Calculate points
   'Work w.r.t the right hand first
   'when we have caculated the thumb and we know the
   'full extents of the drawing then we can mirror and
   'translate the final image etc.

   'NOTE
   'The construction points xyW(n) are completly
   'arbitary and given im no particular order
   'refer to drawing
    
    PR_MakeXY xyFold, 10, 0

    PR_MakeXY xyW(1), xyFold.X - (nWristCir / 2), 0
    PR_MakeXY xyW(4), xyFold.X + (nWristCir / 2), xyW(1).y

    PR_MakeXY xyW(2), xyFold.X - (nPalmCir / 2), nHighWeb
    PR_MakeXY xyW(3), xyFold.X + (nPalmCir / 2), xyW(2).y

    PR_MakeXY xyW(8), xyW(3).X - 0.5, xyW(3).y
    PR_MakeXY xyW(9), xyFold.X + nSeam, xyW(3).y - nSeam
    PR_MakeXY xyW(5), xyW(2).X + nSeam, xyW(2).y

    PR_MakeXY xyW(6), xyW(3).X, xyW(3).y - nWebToWeb
    
   'Finger strap points
   'Assume for the mean time no missing fingers
    nLength = FN_CalcLength(xyW(5), xyW(9))
    nSpace = (nLength - (3 * nStrapWidth)) / 4

    PR_MakeXY xyW(10), xyW(5).X + nSpace, xyW(2).y

    PR_CalcPolar xyW(10), 90, nStrapLengthMiddleAndIndex, xyW(20)
    PR_CalcPolar xyW(20), 0, nStrapWidth, xyW(21)

    nLength = (nStrapLengthMiddleAndIndex + nSeam) - (nSpace / 2)
    PR_CalcPolar xyW(21), 270, nLength, xyW(11)
    PR_CalcPolar xyW(11), 0, nSpace, xyW(12)
    PR_CalcPolar xyW(12), 90, nLength, xyW(22)
    PR_CalcPolar xyW(22), 0, nStrapWidth, xyW(23)
    PR_CalcPolar xyW(23), 270, nLength, xyW(13)
    PR_CalcPolar xyW(13), 0, nSpace, xyW(14)

    nLength = (nStrapLengthLittle + nSeam) - (nSpace / 2)
    PR_CalcPolar xyW(14), 90, nLength, xyW(24)
    PR_CalcPolar xyW(24), 0, nStrapWidth, xyW(25)
    PR_CalcPolar xyW(25), 270, nLength, xyW(15)

    PR_MakeXY xyW(26), xyW(15).X + (nSpace / 2), xyW(9).y


   'Thumb points, construction circles etc.
'    nNotchRad = .09375
'    nBotRadFac = 1.3
'    aThumb = 53
    nNotchRad = 0.05
    nBotRadFac = 1.4
    aThumb = 70
    nThumbOffset = 0.5
    nWristOffset = 0.5
    nWristNotch = 0.125

    nTopThumbRad = nThumbCir + nNotchRad
    nBotThumbRad = ((xyW(6).y - xyW(4).y) - nThumbCir)
    
    If ((xyW(6).y - xyW(4).y) - nThumbCir) <= 0 Then
        sError = "Query thumb circumference as it is impossible to draw the thumb with the given dimension!." & NL
        MsgBox sError, 16, g_sDialogID
        FN_CalculateWebSpacer = False
        Exit Function
    Else
        nBotThumbRad = nBotThumbRad * nBotRadFac
    End If
    
    xyTopThumbRad.X = xyW(6).X + nNotchRad
    xyTopThumbRad.y = xyW(6).y + nNotchRad

   'Calculate notch points
   'Note use of factor to create a taper of the thumb and the
   'seam into the notch
    nTaperNotchRad = nNotchRad * 0.95
    xyT(1).X = xyTopThumbRad.X - nTaperNotchRad
    xyT(1).y = xyTopThumbRad.y
    PR_CalcPolar xyTopThumbRad, 270 + aThumb, nTopThumbRad + 1, xyTmp

    iError = FN_CirLinInt(xyTopThumbRad, xyTmp, xyTopThumbRad, nTaperNotchRad, xyT(2))
    
   'Other thumb points at top
    PR_CalcPolar xyTopThumbRad, 270 + aThumb, nTopThumbRad, xyT(5)
    PR_CalcPolar xyT(5), aThumb, nThumbOffset, xyT(4)
    PR_CalcPolar xyT(4), aThumb + 90, nThumbCir, xyT(3)

   'Calculate bottom circle
   'Note use of cosine rule
    Dim nA#, nB#, nC#, nCosA#
    nA = nTopThumbRad + nBotThumbRad
    nB = FN_CalcLength(xyTopThumbRad, xyW(4))
    nC = nBotThumbRad
    nCosA = ((nB ^ 2 + nC ^ 2) - nA ^ 2) / (2 * nB * nC)
    aAngle = FN_CalcAngle(xyW(4), xyTopThumbRad) - (Arccos(nCosA) * (180 / PI))
    PR_CalcPolar xyW(4), aAngle, nBotThumbRad, xyBotThumbRad

   'Notch at EOS
    PR_GetDlgAboveWrist
    If g_ExtendTo = GLOVE_NORMAL Then
       'Calculate notch
        xyTmp.X = xyBotThumbRad.X
        xyTmp.y = xyFold.y + nWristOffset
        xyTmp1.X = xyFold.X
        xyTmp1.y = xyTmp.y
    
        If FN_CirLinInt(xyTmp, xyTmp1, xyBotThumbRad, nBotThumbRad, xyN(3)) Then
           'Result in xyN(3)
        Else
           'No intesection on bottom circle
            aAngle = FN_CalcAngle(xyW(4), xyT(4))
            nLength = nWristOffset / Sin(aAngle * (PI / 180))
            PR_CalcPolar xyW(4), aAngle, nLength, xyN(3)
        End If
        xyN(2) = xyN(3)
        xyN(2).X = xyN(2).X - nWristNotch
        xyN(1) = xyW(4)
    Else
       'If there is data for an extension then we calculate it here
        PR_CalculateExtension
    End If

   'Exit with true if we get this far
    FN_CalculateWebSpacer = True

End Function

Private Function FN_FigureThumb(nThumbCir) As Double
   'Figure the thumb
   'Based on the Table
   '    "ONE INSERT WITH LARGE SEAMS"
   'GOP 01-02/28, page 4
   'Date May 1991
   '
   ' Globals
   '    g_iInsertSize   -   Given insert size on EIGHTHs
   '
   ' Returns
   '    Figured dimension in inches
   '    0 if not able to calculate

    Dim iThumb%

    FN_FigureThumb = 0

    If g_iInsertSize < 3 Or g_iInsertSize > 8 Then Exit Function

   'Translate thumb circumference to EIGHTHs
    iThumb = round(nThumbCir / EIGHTH)

   'Check that thumb lies within range
    If iThumb < 8 Then Exit Function

   'Calculate Table entry for given thumb and insert
    iThumb = (3 + (7 - g_iInsertSize)) + (iThumb - 8)

    If iThumb < 0 Then Exit Function

    FN_FigureThumb = iThumb * SIXTEENTH

End Function

Private Static Function FN_LengthWristToEOS() As Double
   'Calculates the distance from the EOS to the
   'wrist based on values given in the dialogue
   '
    Dim ii%, nLen#, nSpace#

   'Get the data from the dialogue
    nLen = 0

    Select Case g_ExtendTo
     Case GLOVE_NORMAL
        nLen = 0
     Case GLOVE_ELBOW, GLOVE_PASTWRIST
        For ii = 1 To g_iNumTapesWristToEOS - 1
            nSpace = 1.375
            If ii = 1 And g_nPleats(1) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(1))
            If ii = 2 And g_nPleats(2) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(2))
            If ii = g_iNumTapesWristToEOS - 2 And ii <> 1 And g_nPleats(3) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(3))
            If ii = g_iNumTapesWristToEOS - 1 And ii <> 2 And g_nPleats(4) <> 0 Then nSpace = fnDisplayToInches(g_nPleats(4))
            nLen = nLen + nSpace
        Next ii
     
    End Select

    FN_LengthWristToEOS = nLen

End Function

Private Function FN_Open(sDrafixFile As String, sName As Variant, sPatientFile As Variant, sLeftorRight As Variant) As Integer
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_Open = fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

   'Initialise patient globals
    g_sFileNo = sPatientFile
    g_sSide = sLeftorRight
    g_sPatient = sName
    
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
    Print #fNum, "//by Visual Basic, GLOVES - Drawing"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hChan, hEnt, hSym, hOrigin, hMPD;"
    Print #fNum, "XY     xyStart, xyOrigin, xyScale, xyO;"
    Print #fNum, "STRING sFileNo, sSide, sID, sName;"
    Print #fNum, "ANGLE  aAngle;"

   'Text data
    Print #fNum, "SetData("; QQ; "TextHorzJust"; QC; g_nCurrTextHorizJust; ");"
    Print #fNum, "SetData("; QQ; "TextVertJust"; QC; g_nCurrTextVertJust; ");"
    Print #fNum, "SetData("; QQ; "TextHeight"; QC; g_nCurrTextHt; ");"
    Print #fNum, "SetData("; QQ; "TextAspect"; QC; g_nCurrTextAspect; ");"
    Print #fNum, "SetData("; QQ; "TextFont"; QC; g_nCurrTextFont; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Data"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ZipperLength"; QCQ; "length"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Zipper"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "curvetype"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Leg"; QCQ; "string"; QQ; ");"

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
   'Note this is a bit dodgey if the drawing contains more than 9999 entities
    PR_Setlayer "Construct"
    Print #fNum, "hOrigin = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart"; CC; "0.125);"
    Print #fNum, "if (hOrigin) {"
    Print #fNum, "  sID=StringMiddle(MakeString("; QQ; "long"; QQ; ",UID("; QQ; "get"; QQ; ",hOrigin)), 1, 4) ; "
    Print #fNum, "  while (StringLength(sID) < 4) sID = sID + "; QQ; " "; QQ; ";"
    Print #fNum, "  sID = sID + sFileNo + sSide ;"
    Print #fNum, "  SetDBData(hOrigin,"; QQ; "ID"; QQ; ",sID);"
    Print #fNum, "  }"
    
   'Display Hour Glass Symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
   'Set values for use futher on by other macros
    Print #fNum, "xyOrigin = xyStart"; ";"

    Print #fNum, "hMPD = UID ("; QQ; "find"; QC; Val(txtUidMPD.Text); ");"
    Print #fNum, "if (hMPD)"
    Print #fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);"
    Print #fNum, "else"
    Print #fNum, "  Exit(%cancel,"; QQ; "Can't find > mainpatientdetails < symbol, Insert Patient Data"; QQ; ");"

   'Start drawing on correct side
    PR_Setlayer "Template" & g_sSide

End Function

Private Sub Form_LinkClose()
   
   Dim ii%, nn%, nValue#, nAge%, jj%
   Dim sTapesBeyondWrist$

  'Disable timeout timer as a link close event has happened
   Timer1.Enabled = False


  'Check if MainPatientDetails exists
  'If it dosn't then exit
   If txtUidMPD = "" Then
        MsgBox "No Patient details found!", 16, g_sDialogID
        End
   End If

  'Scaling factor
   If txtUnits.Text = "cm" Then
       g_nUnitsFac = 10 / 25.4
   Else
       g_nUnitsFac = 1
   End If

   If txtUidGlove = "" Then
       'Disable text boxes that will not recieve data

        For ii = 20 To 23
            labCir_DblClick ii
        Next ii
        
        labCir_DblClick 11
        txtLen(4) = 0.5 / g_nUnitsFac
        txtLen_LostFocus 4

        optExtendTo(1).Enabled = False
   End If

  'Hand Option buttons
   If txtSide.Text = "Left" Then
        optHand(0).Value = True
        optHand(1).Enabled = False
        g_sSide = "Left"
   End If
   If txtSide.Text = "Right" Then
        optHand(1).Value = True
        optHand(0).Enabled = False
        g_sSide = "Right"
   End If

  'Circumferences
   For ii = 0 To 10
        nValue = Val(Mid(txtTapeLengths.Text, (ii * 3) + 1, 3))
        If nValue > 0 Then
            txtCir(ii) = nValue / 10
            txtCir_LostFocus (ii)
        End If
   Next ii

  'Lengths
   For ii = 0 To 8
        nn = ii + 11
        nValue = Val(Mid(txtTapeLengths.Text, (nn * 3) + 1, 3))
        If nValue > 0 Then
            txtLen(ii) = nValue / 10
            txtLen_LostFocus (ii)
        End If
   Next ii

   'Tapes beyond wrist
   'normal glove
    ii = 0
    sTapesBeyondWrist = Mid$(txtTapeLengths2.Text, 8)
    For nn = 11 To 12
        nValue = Val(Mid(sTapesBeyondWrist, (ii * 3) + 1, 3))
        If nValue > 0 Then
            txtCir(nn) = nValue / 10
            txtCir_LostFocus (nn)
        End If
        ii = ii + 1
    Next nn

   cboInsert.AddItem "Calculate"
   cboInsert.AddItem "1/8"
   cboInsert.AddItem "1/4"
   cboInsert.AddItem "3/8"
   cboInsert.AddItem "1/2"
   cboInsert.AddItem "5/8"
   cboInsert.AddItem "3/4"
   cboInsert.AddItem "7/8"
   cboInsert.AddItem "1"
   cboInsert.AddItem "1-1/8"
   cboInsert.AddItem "1-1/4"
   cboInsert.AddItem "1-3/8"
   cboInsert.AddItem "1-1/2"
   cboInsert.AddItem "1-5/8"
   cboInsert.AddItem "1-3/4"
   cboInsert.AddItem "1-7/8"

  'Stored insert value (This is stored in 8ths)
   Dim sInsert As String, nInsert As Integer
   
   sInsert = Mid$(txtTapeLengths2.Text, 15, 2)
   If sInsert <> "" Then
        g_iInsertSize = Val(sInsert)
   Else
        g_iInsertSize = 4
   End If
   If g_iInsertSize <= cboInsert.ListCount Then
        cboInsert.ListIndex = g_iInsertSize
   Else
        cboInsert.ListIndex = 4
   End If
        
  
   cboFabric.Text = txtFabric.Text

  'Default always to normal webspacer
   MainForm!optExtendTo(0).Value = True

   If Val(txtTapeLengthPt1.Text) > 0 Then
        PR_GetExtensionDDE_Data
        If Not g_TapesFound Then
           'disable access to the extension
            MainForm!SSTab1.TabEnabled(1) = False
            MainForm!optExtendTo(1).Enabled = False
        End If
   Else
       'disable access to the extension
        MainForm!SSTab1.TabEnabled(1) = False
        MainForm!optExtendTo(1).Enabled = False
   End If


  'Revised Web hts
   If MainForm!txtDataGlove.Text <> "" Then
        ii = 6
        nValue = Val(Mid(MainForm!txtDataGlove.Text, (ii * 2) + 1, 2))
       'Amount thumbweb was dropped in 1/16ths
        g_nThumbWebDrop = nValue * 0.0625
   End If


  'Show after link is closed
   Show
  'Use ok click to draw the web spacer automatically
  ' OK_Click
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()
    Dim ii%
   'Hide during DDE transfer
    Hide

   'Enable Timeout timer
   'Use 6 seconds (a 1/10th of a minute).
   'Timeout is disabled in Link close
    Timer1.Interval = 10000
    Timer1.Enabled = True

    Set MainForm = webspacr

   'Position to center of screen
    Left = (Screen.Width - MainForm.Width) / 2
    Top = (Screen.Height - MainForm.Height) / 2

   'Get the path to the JOBST system installation directory
    g_sPathJOBST = fnPathJOBST()

   'Clear DDE transfer boxes
    txtUidGlove = ""
    txtTapeLengths = ""
    txtTapeLengths2 = ""
    txtTapeLengthPt1 = ""
    txtWristPleat = ""
    txtShoulderPleat = ""
    txtSide = ""
    txtWorkOrder = ""
    txtUidGC = ""
    txtFabric = ""
    txtUidMPD = ""
    txtDataGlove = ""

   'Default units
   'g_nUnitsFac = 1             'Metric
    g_nUnitsFac = 10 / 25.4     'Inches

   'Set up fabric list
    PR_GetComboListFromFile cboFabric, g_sPathJOBST & "\FABRIC.DAT"

   'Setup display inches grid
    grdInches.ColWidth(0) = 615
    grdInches.ColAlignment(0) = 2
    For ii = 0 To 15
        grdInches.RowHeight(ii) = 270
    Next ii

    cboDistalTape.AddItem "1st"
    For ii = 2 To 17
        cboDistalTape.AddItem LTrim$(Mid$(g_sTapeText, (ii * 3) + 1, 3))
    Next ii
    cboDistalTape.ListIndex = 0

    cboProximalTape.AddItem "Last"
    For ii = 17 To 2 Step -1
        cboProximalTape.AddItem LTrim$(Mid$(g_sTapeText, (ii * 3) + 1, 3))
    Next ii
    cboProximalTape.ListIndex = 0

End Sub

Private Sub labCir_DblClick(Index As Integer)
   'Procedure to disable the circumference fields
   'to enable the use to select only the relevent
   'data to be drawn
   'Saves having to delele and re-enter to get variations
   'on a glove
    Dim iDIP%, iPIP%
    If labCir(Index).ForeColor = &HC0C0C0 Then
        Select Case Index
        Case 0 To 7, 11 To 12
            PR_EnableCir Index
        Case 8
            PR_EnableCir Index
           'Disable thumb length
            labLen(4).ForeColor = &H0&
            txtLen(4).Enabled = True
            lblLen(4).Enabled = True
        Case 20 To 23
            iDIP = (Index - 20) * 2
            iPIP = iDIP + 1
            labCir(Index).ForeColor = &H0&
            PR_EnableCir iDIP
            PR_EnableCir iPIP
           'Enable finger length
            labLen(Index - 20).ForeColor = &H0&
            txtLen(Index - 20).Enabled = True
            lblLen(Index - 20).Enabled = True
         End Select
    Else
        Select Case Index
        Case 0 To 7, 12
            PR_DisableCir Index
        Case 8
            PR_DisableCir Index
           'Disable thumb length
            labLen(4).ForeColor = &HC0C0C0
            txtLen(4).Enabled = False
            lblLen(4).Enabled = False
        Case 11
            PR_DisableCir Index
            PR_DisableCir Index + 1
        Case 20 To 23
            iDIP = (Index - 20) * 2
            iPIP = iDIP + 1
            labCir(Index).ForeColor = &HC0C0C0
            PR_DisableCir iDIP
            PR_DisableCir iPIP
           'Disable finger length
            labLen(Index - 20).ForeColor = &HC0C0C0
            txtLen(Index - 20).Enabled = False
            lblLen(Index - 20).Enabled = False
        End Select
    End If

End Sub

Private Sub labLen_DblClick(Index As Integer)
    If Index = 5 Then Exit Sub
    If labLen(Index).ForeColor = &HC0C0C0 Then
        labLen(Index).ForeColor = &H0&
        txtLen(Index).Enabled = True
        lblLen(Index).Enabled = True
    Else
        labLen(Index).ForeColor = &HC0C0C0
        txtLen(Index).Enabled = False
        lblLen(Index).Enabled = False
    End If
End Sub

Private Sub OK_Click()
    Dim sTask$
   
   'Don't allow multiple clicking
   '
    OK.Enabled = False
    
    If FN_CalculateWebSpacer() Then
       'Display and hourglass
        Screen.MousePointer = 11
        
       'N.B. Use of local JOBST Directory C:\JOBST
        PR_CreateDrawMacro "C:\JOBST\DRAW.D"

       'Find Drafix and start macro
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            AppActivate sTask
            SendKeys "@c:\jobst\draw.d{enter}", True
            Screen.MousePointer = 0
            End
        Else
            MsgBox "Can't find a Drafix Drawing to update!", 16, g_sDialogID
        End If
    End If

    OK.Enabled = True
    Screen.MousePointer = 0
    Show
    Exit Sub
    

End Sub

Private Sub optHand_Click(Index As Integer)
    If Index = 0 Then
        g_sSide = "Left"
    Else
        g_sSide = "Right"
    End If
End Sub

Private Sub PR_CalculateExtension()
   'Procedure to calculate the POINTS
   'used to draw the extension of the glove above the
   'wrist
   'The Procedure also supplies the data that is to be
   'printed at each tape
   'N.B.
     
    Dim nMidWristX As Double, nValue As Double, nSpacing As Double
    Dim ii%, nFiguredValue As Double, iStartTape%, iLastTape%, iError%
    Dim iVertex%, nHt#, nWristOffset#, nWristNotch#
    Dim xyTmp As XY



    If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_PASTWRIST Then
       'As we have recalculated the wrist we need to start at the
       'wrist
        nWristOffset = 0.5
        nWristNotch = 0.125

        nHt = FN_LengthWristToEOS()
        UlnarProfile.n = g_iNumTapesWristToEOS
        UlnarProfile.y(1) = nHt
        
        RadialProfile.n = g_iNumTapesWristToEOS
        RadialProfile.y(1) = nHt
        
        iStartTape = g_iWristPointer
        iLastTape = g_iWristPointer + (g_iNumTapesWristToEOS - 1)

       'Note: As we can allow the wrist to be any tape we need a
       'vertex counter
        iVertex = 1
        
        For ii = iStartTape To iLastTape
           'NB use of 1/2 scale
            nValue = g_nCir(ii) * 0.86
            UlnarProfile.X(iVertex) = xyFold.X - (nValue / 2)
            RadialProfile.X(iVertex) = xyFold.X + (nValue / 2)

           'Setup the notes for the vertex
           'We do this here as we have all the data and it simplifies
           'the drawing side
            TapeNote(iVertex).sTapeText = LTrim$(Mid$(g_sTapeText, ((ii + 1) * 3) + 1, 3))
            TapeNote(iVertex).iTapePos = ii
            TapeNote(iVertex).nCir = g_nCir(ii)
           
           'Standard spacing
            nSpacing = 1.375
           
           'Account all for pleats
           'Wrist (as we have started at the wrist we use, iStartTape + 1
           'and iStartTape + 2 in this case)
            If ii = iStartTape + 1 And g_nPleats(1) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(1))
            If ii = iStartTape + 2 And g_nPleats(2) > 0 Then nSpacing = fnDisplayToInches(g_nPleats(2))
           
           'Axilla
            If ii = iLastTape - 1 And g_nPleats(3) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(3))
            If ii = iLastTape And g_nPleats(4) <> 0 Then nSpacing = fnDisplayToInches(g_nPleats(4))
    
           'We are working from the top (wrist) down
           'iVertex = 1 is set before the For Loop (Values from wrist)
            If iVertex <> 1 Then
                UlnarProfile.y(iVertex) = UlnarProfile.y(iVertex - 1) - nSpacing
                RadialProfile.y(iVertex) = RadialProfile.y(iVertex - 1) - nSpacing
                If nSpacing <> 1.375 Then TapeNote(iVertex).sNote = "PLEAT"
            End If

           'Increment vertex count
            iVertex = iVertex + 1
            
        Next ii
      
      '
      'Calculate Notch
       PR_MakeXY xyN(1), RadialProfile.X(RadialProfile.n), RadialProfile.y(RadialProfile.n)
       PR_MakeXY xyTmp, RadialProfile.X(RadialProfile.n - 1), RadialProfile.y(RadialProfile.n - 1)
       If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
           'Cutback by 3/4 of an inch
            iError = FN_CirLinInt(xyN(1), xyTmp, xyN(1), 0.75, xyN(1))
       End If
      
       iError = FN_CirLinInt(xyN(1), xyTmp, xyN(1), nWristOffset, xyN(3))
       xyN(2) = xyN(3)
       xyN(2).X = xyN(2).X - nWristNotch

      'Reset end of profile
       RadialProfile.X(RadialProfile.n) = xyN(3).X
       RadialProfile.y(RadialProfile.n) = xyN(3).y
       UlnarProfile.X(UlnarProfile.n) = xyFold.X - (xyN(3).X - xyFold.X)
       UlnarProfile.y(UlnarProfile.n) = xyN(3).y
 
    End If

End Sub

Private Sub PR_CreateDrawMacro(sFileName As String)
   '
   ' GLOBALS
   '            g_sSide
   '

    Dim nInsert#, nFingerSpace#
    Dim ii%
    Dim xyPt1 As XY, xyPt2 As XY, xyPt3 As XY, xyTextAlign As XY
    Dim xyWristPt1 As XY, xyWristPt2 As XY
    Dim xyTmp As XY
    Dim aAngle#, aInc#, nTranslate#
    Dim Thumb As Curve
    Dim sText$, sWorkOrder$

   'Initialise
    fNum = FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)

   'Draw on correct layer
    PR_Setlayer "Template" & g_sSide
    
'    PR_DebugPoint "xyFold", xyFold
'    For ii = 1 To 20
'        PR_DebugPoint "W" & Trim$(Str$(ii)), xyW(ii)
'    Next ii
'    For ii = 1 To 8
'        PR_DebugPoint "T" & Trim$(Str$(ii)), xyT(ii)
'    Next ii

'PR_DebugPoint "xyTopThumbRad", xyTopThumbRad
'PR_DebugPoint "xyBotThumbRad", xyBotThumbRad

   'Having established the points we now draw them
   'We must mirror for the right hand side and
   'translate all points relative to our temporary datum xyFold
    
   'Translate points relative to xyFold
    Dim nYTranslate#
    nYTranslate = FN_LengthWristToEOS()
    nTranslate = -(xyFold.X + (xyFold.X - xyT(4).X))
    xyFold.X = xyFold.X + nTranslate
    xyFold.y = xyFold.y + nYTranslate

    For ii = 1 To 5
        xyT(ii).X = xyT(ii).X + nTranslate
        xyT(ii).y = xyT(ii).y + nYTranslate
    Next ii

    For ii = 1 To 3
        xyN(ii).X = xyN(ii).X + nTranslate
    Next ii

    For ii = 1 To 26
        xyW(ii).X = xyW(ii).X + nTranslate
        xyW(ii).y = xyW(ii).y + nYTranslate
        If g_sSide = "Left" Then PR_MirrorPointInYaxis xyFold.X, xyW(ii)
    Next ii
    xyTopThumbRad.X = xyTopThumbRad.X + nTranslate
    xyTopThumbRad.y = xyTopThumbRad.y + nYTranslate
    xyBotThumbRad.X = xyBotThumbRad.X + nTranslate
    xyBotThumbRad.y = xyBotThumbRad.y + nYTranslate

    
   'Drawfingerstraps
    PR_StartPoly
        PR_AddVertex xyW(2), 0
        PR_AddVertex xyW(10), 0
        PR_AddVertex xyW(20), 0
        PR_AddVertex xyW(21), 0
        If g_sSide = "Left" Then PR_AddVertex xyW(11), -1 Else PR_AddVertex xyW(11), 1
        PR_AddVertex xyW(12), 0
        PR_AddVertex xyW(22), 0
        PR_AddVertex xyW(23), 0
        If g_sSide = "Left" Then PR_AddVertex xyW(13), -1 Else PR_AddVertex xyW(13), 1
        PR_AddVertex xyW(14), 0
        PR_AddVertex xyW(24), 0
        PR_AddVertex xyW(25), 0
        If g_sSide = "Left" Then PR_AddVertex xyW(15), -0.414 Else PR_AddVertex xyW(15), 0.414
        PR_AddVertex xyW(26), 0
        PR_AddVertex xyW(9), 0
        PR_AddVertex xyW(8), 0
        PR_AddVertex xyW(3), 0
    PR_EndPoly
   
    Dim nLength#

    If g_ExtendTo = GLOVE_ELBOW Or g_ExtendTo = GLOVE_PASTWRIST Then

        For ii = 1 To RadialProfile.n
            UlnarProfile.X(ii) = UlnarProfile.X(ii) + nTranslate
            RadialProfile.X(ii) = RadialProfile.X(ii) + nTranslate
        Next ii

        PR_DrawPoly UlnarProfile
        PR_DrawPoly RadialProfile

       'Draw notes for each tape
       If g_ExtendTo = GLOVE_ELBOW Then
        nLength = -1
        For ii = 1 To RadialProfile.n
            PR_MakeXY xyPt1, UlnarProfile.X(ii), UlnarProfile.y(ii)
            PR_MakeXY xyPt2, RadialProfile.X(ii), RadialProfile.y(ii)
            If ii = 1 Then
                xyWristPt1 = xyPt1
                xyWristPt2 = xyPt2
            End If
            If ii <> RadialProfile.n Then
                PR_DrawTapeNotes ii, xyPt1, xyPt2, nLength, ""
            Else
                If TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
                    xyTmp = xyN(1)
                    xyPt3 = xyN(1)
                Else
                    xyTmp = xyN(1)
                    xyTmp.y = 0
                    xyPt3 = xyTmp
                End If
                PR_MirrorPointInYaxis xyFold.X, xyPt3
                PR_DrawTapeNotes ii, xyTmp, xyPt3, nLength, ""
            End If
            
            If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
               'Draw elbow line
                PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
                PR_Setlayer "Notes"
                PR_DrawLine xyPt1, xyPt2
                PR_CalcMidPoint xyPt1, xyPt2, xyPt3
                xyPt3.y = xyPt3.y + 0.0625
                PR_DrawText "E", xyPt3, 0.1
            End If
        Next ii

       'Center line
        PR_Setlayer "Construct"
        PR_CalcMidPoint xyWristPt1, xyWristPt2, xyTmp
        PR_CalcMidPoint xyPt1, xyPt2, xyPt3
        PR_DrawLine xyPt3, xyTmp
       End If

       'Wrist line
        PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT

        PR_Setlayer "Notes"
        PR_DrawLine xyW(1), xyW(4)
        PR_CalcMidPoint xyW(1), xyW(4), xyPt3
        xyPt3.y = xyPt3.y + 0.0625
        PR_DrawText "W", xyPt3, 0.1

       'Elastic Text
        xyPt1 = xyN(1)
        xyPt2 = xyN(1)
        
        PR_MirrorPointInYaxis xyFold.X, xyPt2

        sText = Trim$(fnInchestoText(FN_CalcLength(xyPt1, xyPt2) + 1.5))
        PR_CalcMidPoint xyPt1, xyPt2, xyPt3
        xyPt3.y = xyPt3.y + 0.375
        sText = sText & "\"" ELASTIC"
        PR_DrawText sText, xyPt3, 0.1

       'Elastic for children
        If Val(txtAge.Text) <= 10 Then
 '           PR_Setlayer "Notes"
 '           PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.y = xyPt3.y - 0.25
            PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
            PR_DrawText "1/2\"" ELASTIC", xyPt3, 0.1
        End If


        'EOS
        PR_Setlayer "Template" & g_sSide
        
        PR_DrawLine xyPt1, xyPt2

    Else
       
       'Elastic text at EOS
        PR_Setlayer "Notes"
        sText = Trim$(fnInchestoText(FN_CalcLength(xyW(1), xyW(4)) + 1.5))
        PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
        PR_CalcMidPoint xyW(1), xyW(4), xyPt3
        xyPt3.y = xyPt3.y + 0.25
        sText = sText & "\"" ELASTIC"
        PR_DrawText sText, xyPt3, 0.1

       'EOS at wrist
        PR_Setlayer "Template" & g_sSide
        PR_DrawLine xyW(1), xyW(4)

    End If

   'Drawthumb (LEFT)
    If g_sSide = "Left" Then PR_DrawLine xyT(1), xyW(2) Else PR_DrawLine xyT(1), xyW(3)
    
    PR_DrawArc xyTopThumbRad, xyT(1), xyT(2)
    PR_StartPoly
        PR_AddVertex xyT(2), 0
        PR_AddVertex xyT(3), 0
        PR_AddVertex xyT(4), 0
        PR_AddVertex xyT(5), 0
    PR_EndPoly

   'Notch
    PR_StartPoly
        PR_AddVertex xyN(1), 0
        PR_AddVertex xyN(2), 0
        PR_AddVertex xyN(3), 0
    PR_EndPoly

   'Calculate thumb curve
   'For the left then mirror for right
    Thumb.n = 8

    If g_ExtendTo = GLOVE_NORMAL Then
        xyTmp = xyN(3)
    Else
        If g_sSide = "Left" Then
            xyTmp = xyW(1)
        Else
            xyTmp = xyW(4)
        End If
    End If
    Thumb.n = 1
    Thumb.X(Thumb.n) = xyTmp.X
    Thumb.y(Thumb.n) = xyTmp.y
    
    aAngle = FN_CalcAngle(xyBotThumbRad, xyTmp)
    aInc = (aAngle - FN_CalcAngle(xyBotThumbRad, xyTopThumbRad)) / 5
    For ii = 2 To 4
        aAngle = aAngle - aInc
        PR_CalcPolar xyBotThumbRad, aAngle, nBotThumbRad, xyTmp
        If xyTmp.y > xyN(3).y Then
            Thumb.n = Thumb.n + 1
            Thumb.X(Thumb.n) = xyTmp.X
            Thumb.y(Thumb.n) = xyTmp.y
        End If
    Next ii

    aAngle = FN_CalcAngle(xyTopThumbRad, xyBotThumbRad)
    aInc = (FN_CalcAngle(xyTopThumbRad, xyT(5)) - aAngle) / 4
    For ii = 5 To 8
        aAngle = aAngle + aInc
        PR_CalcPolar xyTopThumbRad, aAngle, nTopThumbRad, xyTmp
        If xyTmp.y > xyN(3).y Then
            Thumb.n = Thumb.n + 1
            Thumb.X(Thumb.n) = xyTmp.X
            Thumb.y(Thumb.n) = xyTmp.y
        End If
    Next ii

    PR_DrawFitted Thumb

   'Draw thumb (Right)
    For ii = 1 To 5
        PR_MirrorPointInYaxis xyFold.X, xyT(ii)
    Next ii
    For ii = 1 To 3
        PR_MirrorPointInYaxis xyFold.X, xyN(ii)
    Next ii
    PR_MirrorPointInYaxis xyFold.X, xyTopThumbRad
    PR_MirrorCurveInYaxis xyFold.X, Thumb
    If g_sSide = "Left" Then PR_DrawLine xyT(1), xyW(3) Else PR_DrawLine xyT(1), xyW(2)
    
    PR_DrawArc xyTopThumbRad, xyT(1), xyT(2)
    PR_StartPoly
        PR_AddVertex xyT(2), 0
        PR_AddVertex xyT(3), 0
        PR_AddVertex xyT(4), 0
        PR_AddVertex xyT(5), 0
    PR_EndPoly

   'notch
    PR_StartPoly
        PR_AddVertex xyN(1), 0
        PR_AddVertex xyN(2), 0
        PR_AddVertex xyN(3), 0
    PR_EndPoly
    PR_DrawFitted Thumb

'PR_DrawCircle xyBotThumbRad, nBotThumbRad
'PR_DrawCircle xyTopThumbRad, nTopThumbRad
'PR_DrawCircle xyTopThumbRad, nNotchRad

   'Add notes etc.
    PR_Setlayer "Notes"


   'Insert Size (Text)
   'Position Insert text Size w.r.t Little Finger web
    nInsert = g_iInsertSize * EIGHTH
    If nInsert = 0.375 Then
       'Substitute 1/2" for 3/8" inch for printing only
        sText = Trim$(fnInchestoText(0.5))
    Else
        sText = Trim$(fnInchestoText(nInsert))
    End If
    If Mid$(sText, 1, 1) = "-" Then sText = Mid$(sText, 2) 'Strip leading "-" sign
    sText = sText & "\" & QQ
    sText = "INSERT " & sText

    If g_sSide = "Left" Then
        PR_CalcMidPoint xyW(15), xyW(14), xyTextAlign
    Else
        PR_CalcMidPoint xyW(10), xyW(11), xyTextAlign
    End If
    xyTextAlign.y = xyTextAlign.y - 0.5

    xyPt3 = xyTextAlign
    PR_DrawText sText, xyPt3, 0.1

   'Arrow
    PR_DrawMarkerNamed "closed arrow", xyW(9), 0.125, 0.05, 270

   'Patient Details
    PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
    PR_MakeXY xyPt3, xyPt3.X, xyPt3.y - 0.3
    If txtWorkOrder.Text = "" Then
        sWorkOrder = "-"
    Else
        sWorkOrder = txtWorkOrder.Text
    End If
    sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim$(Mid$(txtFabric.Text, 4))
    PR_DrawText sText, xyPt3, 0.1

   'Other patient details in black on layer construct
    PR_Setlayer "Construct"
    PR_MakeXY xyPt3, xyPt3.X, xyPt3.y - 0.8
    sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
    PR_DrawText sText, xyPt3, 0.1

   'This will only save if there is no existing glove data
    PR_UpdateDDE
    PR_UpdateDBFields

Close #fNum

End Sub

Private Sub PR_DisableCir(Index As Integer)
   'Read as part of labCir_DblClick
    labCir(Index).ForeColor = &HC0C0C0
    txtCir(Index).Enabled = False
    lblCir(Index).Enabled = False
End Sub

Private Sub PR_DrawTapeNotes(iVertex As Integer, xyStart As XY, xyEnd As XY, nOffSet As Double, sZipperID As String)
   'Draws the the notes at ecah tape
    Dim xyDatum As XY, xyInsert As XY
    Dim aAngle As Double, nLength As Double, nDec As Double
    Dim sText As String
    Dim iInt As Integer
      
   'Exit from sub if no Tape position Given as
   'this implies that the TapeNote is Empty
   '(EG with GLOVE_NORMAL)
    If TapeNote(iVertex).iTapePos = 0 Then Exit Sub

   'Draw Construction line
    PR_Setlayer "Construct"

    PR_DrawLine xyStart, xyEnd

   'Get center point on which text positions are based
   'Doing it this way accounts for case when xyStart.x > xyEnd.x
    If nOffSet > 0 Then
        aAngle = FN_CalcAngle(xyStart, xyEnd)
        PR_CalcPolar xyStart, aAngle, nOffSet, xyDatum
    Else
        PR_CalcMidPoint xyStart, xyEnd, xyDatum
    End If

   'Insert Symbol
    PR_InsertSymbol "glvTapeNotes", xyDatum, 1, 0
    
   'Setup for Tape Lable
    PR_AddDBValueToLast "Data", TapeNote(iVertex).sTapeText
   
    
   'Circumference  in Inches and Eighths format
    iInt = Int(TapeNote(iVertex).nCir)
    PR_AddDBValueToLast "TapeLengths", Trim$(Str$(iInt))

    nDec = (TapeNote(iVertex).nCir - iInt)
    If nDec > 0 Then
        iInt = round((nDec / EIGHTH) * 10) / 10
        If iInt <> 0 Then PR_AddDBValueToLast "TapeLengths2", Trim$(Str$(iInt))
    End If

    PR_AddDBValueToLast "Reduction", "14"
    
   'Check for Notes
    If TapeNote(iVertex).sNote <> "" Then
       'set text justification etc.
        PR_SetTextData RIGHT_, VERT_CENTER, EIGHTH, CURRENT, CURRENT
        PR_MakeXY xyInsert, xyDatum.X - 0.1, xyDatum.y + 0.35
        If TapeNote(iVertex).sNote <> "PLEAT" Then PR_Setlayer "Notes"
        PR_DrawText TapeNote(iVertex).sNote, xyInsert, 0.1
    End If

End Sub

Private Sub PR_EnableCir(Index As Integer)
   'Read as part of labCir_DblClick
    labCir(Index).ForeColor = &H0&
    txtCir(Index).Enabled = True
    lblCir(Index).Enabled = True
End Sub

Private Sub PR_GetComboListFromFile(Combo_Name As Control, sFileName As String)
   'General procedure to create the list section of
   'a combo box reading the data from a file
    
    Dim sLine As String, fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48, g_sDialogID
        Exit Sub
    End If

    Open sFileName For Input As fFileNum
    Do While Not EOF(fFileNum)
        Line Input #fFileNum, sLine
        Combo_Name.AddItem sLine
        Loop
    Close #fFileNum

End Sub

Private Sub PR_GetDlgAboveWrist()
   'General procedure to read the data given in the
   'dialogue controls and copy to module level variables.
   'Saves having to do this more than once.
   '
   'Updates:-
   '
   '    g_nCir(1, NOFF_ARMTAPES) As Double
   '    g_iFirstTape            As Integer
   '    g_iLastTape             As Integer
   '    g_iWristPointer         As Integer
   '    g_iEOSPointer           As Integer
   '    g_iNumTotalTapes        As Integer
   '    g_iNumTapesWristToEOS   As Integer
   '    g_EOSType               As Integer
   '    g_OnFold                As Integer
   '    g_ExtendTo              As Integer
   '    g_nPleats(1 To 4)       As Double
   '    g_iPressure             As Integer
   '
   'NOTE:
   '    We ignore the fact that there may be missing tapes.
   '
   '    g_iWristPointer and g_iEOSPointer are used to indicate the
   '    start and finish tapes in the arrays and not in the
   '    txtExtCir() array

    Dim nCir#
    Dim ii%


    g_iFirstTape = -1
    g_iLastTape = -1
    g_iWristPointer = -1
    g_iEOSPointer = -1
    g_iNumTotalTapes = 0
    g_iNumTapesWristToEOS = 0

   'Glove type
    If MainForm!optExtendTo(0).Value = True Then
        g_ExtendTo = GLOVE_NORMAL
    ElseIf MainForm!optExtendTo(1).Value = True Then
        g_ExtendTo = GLOVE_ELBOW
    End If

    
    If g_ExtendTo = GLOVE_NORMAL Then
       'Get values from Hand Data TAB
        For ii = 10 To 12
            nCir = FN_InchesValue(MainForm!txtCir(ii))
            g_nCir(ii - 9) = nCir
        Next ii
        g_iFirstTape = 1

       'Assumes no holes in data
        If g_nCir(2) <= 0 Then
            g_iNumTotalTapes = 1
            g_iNumTapesWristToEOS = 1
            g_iLastTape = 1
        ElseIf g_nCir(3) > 0 Then
            g_iNumTotalTapes = 3
            g_iNumTapesWristToEOS = 3
            g_iLastTape = 3
            g_ExtendTo = GLOVE_PASTWRIST
        Else
            g_iNumTotalTapes = 2
            g_iNumTapesWristToEOS = 2
            g_iLastTape = 2
            g_ExtendTo = GLOVE_PASTWRIST
        End If
'        g_iWristPointer = 1
        g_iWristPointer = 1
    Else
        If MainForm!SSTab1.Tab <> 1 Then MainForm!SSTab1.Tab = 1

       'Get First and last tape (if any)
        For ii = 8 To 23
            nCir = FN_InchesValue(MainForm!txtExtCir(ii))
            If nCir > 0 And g_iFirstTape = -1 Then g_iFirstTape = ii - 7
            g_nCir(ii - 7) = nCir
        Next ii
        
        For ii = 23 To 8 Step -1
            If Val(MainForm!txtExtCir(ii)) > 0 Then
                g_iLastTape = ii - 7
                Exit For
            End If
        Next ii
    
       'Total number of tapes
        If (g_iLastTape = g_iFirstTape) Then
            If g_iLastTape <> -1 Then g_iNumTotalTapes = 1
        Else
            g_iNumTotalTapes = (g_iLastTape - g_iFirstTape) + 1
        End If
       'Get values from "Above Wrist" TAB
        Dim iIndex%
       'Wrist tape
       'cboDistalTape.ListIndex = 0 => use first (Defaults to this)
        iIndex = MainForm!cboDistalTape.ListIndex
        If iIndex = 0 Or iIndex = -1 Then
            g_iWristPointer = g_iFirstTape
        Else
            g_iWristPointer = iIndex
        End If
    
       'EOS tape
       'cboProximalTape.ListIndex = 0 => use last (Defaults to this)
        iIndex = MainForm!cboProximalTape.ListIndex
        If iIndex = 0 Or iIndex = -1 Then
            g_iEOSPointer = g_iLastTape
        Else
           'NB. The order of tapes is reversed in this list
           'starting at 19-1/2 and finishing at 0
            g_iEOSPointer = 17 - iIndex
        End If
    
        
       'Number of tapes between
        If (g_iWristPointer = g_iEOSPointer) Then
           'Just in case there are no tapes
           'And "first" and "last" are used for Wrist and EOS
            If g_iLastTape <> -1 Then g_iNumTapesWristToEOS = 1
        Else
            g_iNumTapesWristToEOS = (g_iEOSPointer - g_iWristPointer) + 1
        End If
    
       'Pleats
        If MainForm!txtWristPleat1.Enabled = True Then g_nPleats(1) = Val(MainForm!txtWristPleat1.Text) Else g_nPleats(1) = 0
        If MainForm!txtWristPleat2.Enabled = True Then g_nPleats(2) = Val(MainForm!txtWristPleat2.Text) Else g_nPleats(2) = 0
        If MainForm!txtShoulderPleat2.Enabled = True Then g_nPleats(3) = Val(MainForm!txtShoulderPleat2.Text) Else g_nPleats(3) = 0
        If MainForm!txtShoulderPleat1.Enabled = True Then g_nPleats(4) = Val(MainForm!txtShoulderPleat1.Text) Else g_nPleats(4) = 0

    
    End If



End Sub

Private Sub PR_GetExtensionDDE_Data()
   'Procedure to use the data given in the Form for the
   'glove extension (if any) along the arm
   '
    Dim nAge As Integer, Flap%
    Dim ii%, jj%, nValue#, nn%
    Dim iMMs%, iGms%, iRed%
    Dim sDiag As String
    
    MainForm!SSTab1.Tab = 1

   'Defaults
    MainForm!optExtendTo(0).Value = True     'Normal Glove
    MainForm!cboDistalTape.ListIndex = 0     '1st Tape
    MainForm!cboProximalTape.ListIndex = 0   'Last Tape
    Flap = False

    
   'Set default pressure w.r.t diagnosis
    nAge = Val(MainForm!txtAge.Text)


  'Glove extensions
  'Code for extended glove
  'Glove Option Buttons
   Dim nLen#

   If MainForm!txtDataGlove.Text <> "" Then
     For ii = 0 To 5
        nValue = Val(Mid(MainForm!txtDataGlove.Text, (ii * 2) + 1, 2))
        If ii = 0 Then
           'Fold options
           'Do nothing Off fold is default
        ElseIf ii = 1 And nValue >= 0 Then
           'Pressure w.r.t Figuring and MMs
            'ignore
        ElseIf ii = 2 Then
           'Wrist Tape
            MainForm!cboDistalTape.ListIndex = nValue
        ElseIf ii = 3 Then
           'Proximal Tape
            MainForm!cboProximalTape.ListIndex = nValue
        ElseIf ii = 4 Then
            'Ignore
        ElseIf ii = 5 Then
            'Ignore flaps
        End If
    Next ii
   End If

  'Glove to elbow and Glove to axilla
  'These can start at -3 (However to allow for possible
  'changes later we save up to -4-1/2 but ignore -4-1/2 for the
  'mean time)
   g_TapesFound = False
   If MainForm!txtTapeLengthPt1.Text <> "" Then
    ii = 1
    For nn = 8 To 23
        nValue = Val(Mid(MainForm!txtTapeLengthPt1.Text, (ii * 3) + 1, 3))
        If nValue > 0 Then
            MainForm!txtExtCir(nn) = nValue / 10
            nLen = FN_InchesValue(MainForm!txtExtCir(nn))
            If nLen <> -1 Then
                PR_GrdInchesDisplay nn - 8, nLen
                g_TapesFound = True
            End If
            g_nCir(ii) = nLen
        Else
            PR_GrdInchesDisplay nn - 8, 0
        End If
        ii = ii + 1
    Next nn
   End If

  'Pleats
    For ii = 0 To 1
        nValue = Val(Mid(MainForm!txtWristPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtWristPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtWristPleat2 = nValue / 10
        End If
        nValue = Val(Mid(MainForm!txtShoulderPleat.Text, (ii * 3) + 1, 3))
        If ii = 0 And nValue > 0 Then
            MainForm!txtShoulderPleat1 = nValue / 10
        ElseIf ii = 1 And nValue > 0 Then
            MainForm!txtShoulderPleat2 = nValue / 10
        End If
    Next ii
    
    MainForm!SSTab1.Tab = 0

End Sub

Private Sub PR_GrdInchesDisplay(iIndex As Integer, nLen As Double)
    MainForm!grdInches.Row = iIndex
    MainForm!grdInches.Col = 0
    MainForm!grdInches.Text = fnInchestoText(nLen)
End Sub

Private Sub PR_UpdateDBFields()

    Dim sSymbol$

   'Glove common
    sSymbol = "glovecommon"
    If txtUidGC.Text = "" Then
       'Insert a new symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ")){"
        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x, xyO.y);"
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find >"; sSymbol; "< symbol to insert\nCheck your installation, that JOBST.SLB exists!"; QQ; ");"
       'Update DB fields
        Print #fNum, "SetDBData( hSym"; CQ; "Fabric"; QCQ; txtFabric.Text; QQ; ");"
        Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"
    End If
    

   'Glove
    sSymbol = "gloveglove"
    If txtUidGlove.Text = "" Then
       'Insert a new symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; sSymbol; QQ; ")){"
        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "Data"; QQ; "));"
        If optHand(0) = True Then
           'Left Hand
            Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x+1.5 , xyO.y);"
        Else
           'Right Hand
            Print #fNum, "  hSym = AddEntity("; QQ; "symbol"; QCQ; sSymbol; QC; "xyO.x+3 , xyO.y);"
        End If
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find >"; sSymbol; "< symbol to insert\nCheck your installation, that JOBST.SLB exists!"; QQ; ");"
       'Update DB fields
        Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths"; QCQ; txtTapeLengths.Text; QQ; ");"
        Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths2"; QCQ; txtTapeLengths2.Text; QQ; ");"
        Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"
        Print #fNum, "SetDBData( hSym"; CQ; "Sleeve"; QCQ; txtSide.Text; QQ; ");"
        Print #fNum, "SetDBData( hSym"; CQ; "Data"; QCQ; txtDataGlove.Text; QQ; ");"
    End If
    

End Sub

Private Sub PR_UpdateDDE()
        
   'Procedure to update the fields used when data is transfered
   'from DRAFIX using DDE.
   'Although the transfer back to DRAFIX is not via DDE we use the same controls
   'simply to illustrate the method by which the data is packed into the
   'fields

   'The decisons on format for each is not very neat!
   'The reason is the iteritive development process and the 63 char length
   'limitation on DRAFIX DB Fields.

   'Backward compatibility with the CAD-Glove is required in this
   'procedure

    MainForm!SSTab1.Tab = 0

    Dim iLen%, ii%
    Dim sLen$, sPacked$

   'Initialise
    sPacked = ""
   
   'Pack Circumferences
    For ii = 0 To 10
        iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Pack Lengths
    For ii = 0 To 8
        iLen = Val(txtLen(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Store to DDE text boxes
    txtTapeLengths.Text = sPacked

   'Tip options
    sPacked = ""
    For ii = 0 To 4
        sPacked = sPacked & "0"
    Next ii

   'CheckBoxes
    sPacked = sPacked & "0"
    sPacked = sPacked & "0"
    
   'Tapes past wrist (Store here due to limitations on DB field length
   'Pack Circumferences
    For ii = 11 To 12
        iLen = Val(txtCir(ii).Text) * 10 'Shift decimal place
    
        If iLen <> 0 Then
            sLen = String(3, " ")
            RSet sLen = Trim$(Str$(iLen))
        Else
            sLen = String(3, " ")
        End If
    
        sPacked = sPacked & sLen

    Next ii

   'Insert size
    sLen = String(3, " ")
    RSet sLen = Trim$(Str$(g_iInsertSize))
    sPacked = sPacked & sLen

   'Reinforced dorsal
    sPacked = sPacked & "0"

   'Thumb Options
    sPacked = sPacked & "0"
   
   'Insert Options
    sPacked = sPacked & "0"

   'Store to DDE text boxes
    txtTapeLengths2.Text = sPacked


   'Fabric
    If cboFabric.Text <> "" Then txtFabric = cboFabric.Text

    
End Sub

Private Sub Tab_Click()
   'Allows the user to use enter as a tab
    
    SendKeys "{TAB}", False

End Sub

Private Sub Timer1_Timer()
   'End as the link close event has not
   'occured.
    End
End Sub

Private Sub txtCir_GotFocus(Index As Integer)
    PR_Select_Text txtCir(Index)
End Sub

Private Sub txtCir_LostFocus(Index As Integer)
        
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtCir(Index))

   'Display value in inches if Valid
    If nLen <> -1 Then lblCir(Index).Caption = fnInchestoText(nLen)

End Sub

Private Sub txtLen_GotFocus(Index As Integer)
    PR_Select_Text txtLen(Index)
End Sub

Private Sub txtLen_LostFocus(Index As Integer)
    Dim nLen#
    Dim nTipCutBack#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtLen(Index))
    If nLen < 0 Then
        lblLen(Index).Caption = ""
        Exit Sub
    End If
    
    
   'Display length text
    If nLen > 0 Then
        lblLen(Index).Caption = fnInchestoText(nLen)
    Else
        lblLen(Index).Caption = ""
    End If

End Sub


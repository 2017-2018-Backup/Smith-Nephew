VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form manglove 
   Caption         =   "MANUAL Glove - Dialogue"
   ClientHeight    =   8892
   ClientLeft      =   -360
   ClientTop       =   912
   ClientWidth     =   9552
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
   LinkTopic       =   "manglove"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8892
   ScaleWidth      =   9552
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   88
      Top             =   120
      Width           =   4815
      _ExtentX        =   8488
      _ExtentY        =   15473
      _Version        =   327681
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Hand Data"
      TabPicture(0)   =   "manglove.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Above Wrist"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   186
         Top             =   120
         Width           =   4485
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   9
            Left            =   2640
            TabIndex        =   226
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   8
            Left            =   2640
            TabIndex        =   225
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   9
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   111
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   8
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   110
            Top             =   330
            Width           =   495
         End
         Begin VB.Frame frmCalculate 
            Caption         =   "Figure Arm"
            Height          =   825
            Left            =   300
            TabIndex        =   221
            Top             =   5220
            Width           =   4035
            Begin VB.ComboBox cboPressure 
               Height          =   315
               Left            =   780
               Style           =   2  'Dropdown List
               TabIndex        =   223
               Top             =   330
               Width           =   795
            End
            Begin VB.CommandButton cmdCalculate 
               Caption         =   "Calculate"
               Height          =   375
               Left            =   2220
               TabIndex        =   222
               Top             =   300
               Width           =   1275
            End
            Begin VB.Label lblArm 
               Caption         =   "Press:"
               Height          =   255
               Index           =   7
               Left            =   180
               TabIndex        =   224
               Top             =   360
               Width           =   555
            End
         End
         Begin VB.ComboBox cboDistalTape 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   220
            Top             =   4830
            Width           =   735
         End
         Begin VB.Frame frmGloveToAxilla 
            Caption         =   "Glove to Axilla, EOS Style"
            Height          =   1965
            Left            =   300
            TabIndex        =   207
            Top             =   6090
            Width           =   4035
            Begin VB.OptionButton optProximalTape 
               Caption         =   "Flap"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   210
               Top             =   360
               Width           =   915
            End
            Begin VB.OptionButton optProximalTape 
               Caption         =   "Plain"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   209
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.TextBox txtCustFlapLength 
               Height          =   285
               Left            =   2340
               TabIndex        =   205
               Top             =   1320
               Width           =   495
            End
            Begin VB.ComboBox cboFlaps 
               Height          =   315
               Left            =   1860
               TabIndex        =   208
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtStrapLength 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2340
               TabIndex        =   203
               Top             =   780
               Width           =   495
            End
            Begin VB.TextBox txtWaistCir 
               Height          =   285
               Left            =   2340
               TabIndex        =   206
               Top             =   1590
               Width           =   495
            End
            Begin VB.TextBox txtFrontStrapLength 
               Height          =   285
               Left            =   2340
               TabIndex        =   204
               Top             =   1050
               Width           =   495
            End
            Begin VB.Label lblFlap 
               Caption         =   "Requested Flap length:"
               Height          =   285
               Index           =   3
               Left            =   240
               TabIndex        =   219
               Top             =   1350
               Width           =   1995
            End
            Begin VB.Label lblFlap 
               Caption         =   "Back Strap:"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   218
               Top             =   810
               Width           =   1035
            End
            Begin VB.Label lblFlap 
               Caption         =   "Style"
               Height          =   285
               Index           =   0
               Left            =   3360
               TabIndex        =   217
               Top             =   390
               Width           =   555
            End
            Begin VB.Label labStrap 
               Height          =   195
               Left            =   2880
               TabIndex        =   216
               Top             =   810
               Width           =   675
            End
            Begin VB.Label labCustFlapLength 
               Height          =   195
               Left            =   2880
               TabIndex        =   215
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label labWaistCir 
               Height          =   195
               Left            =   2880
               TabIndex        =   214
               Top             =   1620
               Width           =   675
            End
            Begin VB.Label lblFlap 
               Caption         =   "Style ""D"" Waist Circum:"
               Height          =   285
               Index           =   4
               Left            =   180
               TabIndex        =   213
               Top             =   1620
               Width           =   2115
            End
            Begin VB.Label labFrontStrapLength 
               Height          =   195
               Left            =   2880
               TabIndex        =   212
               Top             =   1080
               Width           =   675
            End
            Begin VB.Label lblFlap 
               Caption         =   "Front Strap:"
               Height          =   285
               Index           =   2
               Left            =   1200
               TabIndex        =   211
               Top             =   1080
               Width           =   1095
            End
         End
         Begin VB.ComboBox cboProximalTape 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   4830
            Width           =   735
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   10
            Left            =   2640
            TabIndex        =   201
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   10
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   112
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat1 
            Height          =   285
            Left            =   180
            TabIndex        =   150
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat1 
            Height          =   285
            Left            =   180
            TabIndex        =   153
            Top             =   4350
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat2 
            Height          =   285
            Left            =   180
            TabIndex        =   151
            Top             =   1380
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat2 
            Height          =   285
            Left            =   180
            TabIndex        =   152
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   11
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   113
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   13
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   115
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   14
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   116
            Top             =   1950
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   15
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   117
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   118
            Top             =   2490
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   17
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   119
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   18
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   120
            Top             =   3030
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   19
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   121
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   20
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   122
            Top             =   3570
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   21
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   123
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   22
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   124
            Top             =   4110
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   23
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   125
            Top             =   4380
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   11
            Left            =   2640
            TabIndex        =   200
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   12
            Left            =   2640
            TabIndex        =   199
            Top             =   1410
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   13
            Left            =   2640
            TabIndex        =   198
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   14
            Left            =   2640
            TabIndex        =   197
            Top             =   1950
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   15
            Left            =   2640
            TabIndex        =   196
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   2640
            TabIndex        =   195
            Top             =   2490
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   18
            Left            =   2640
            TabIndex        =   194
            Top             =   3030
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   19
            Left            =   2640
            TabIndex        =   193
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   20
            Left            =   2640
            TabIndex        =   192
            Top             =   3570
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   21
            Left            =   2640
            TabIndex        =   191
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   22
            Left            =   2640
            TabIndex        =   190
            Top             =   4110
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   23
            Left            =   2640
            TabIndex        =   189
            Top             =   4380
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   17
            Left            =   2640
            TabIndex        =   188
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   12
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   114
            Top             =   1410
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid grdInches 
            Height          =   4275
            Left            =   1860
            TabIndex        =   187
            Top             =   330
            Width           =   615
            _ExtentX        =   1080
            _ExtentY        =   7535
            _Version        =   65541
            Rows            =   16
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483638
            GridLines       =   0
            ScrollBars      =   0
            BorderStyle     =   0
         End
         Begin MSFlexGridLib.MSFlexGrid grdDisplay 
            Height          =   4275
            Left            =   3240
            TabIndex        =   227
            Top             =   330
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   7535
            _Version        =   65541
            Rows            =   16
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483638
            GridLines       =   0
            ScrollBars      =   0
            BorderStyle     =   0
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   256
            Top             =   4110
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   255
            Top             =   3420
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   254
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-1½"
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   253
            Top             =   630
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-3"
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   252
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Wrist"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   251
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   195
            Index           =   9
            Left            =   1680
            TabIndex        =   250
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "EOS"
            Height          =   195
            Index           =   10
            Left            =   2580
            TabIndex        =   249
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   195
            Index           =   8
            Left            =   3840
            TabIndex        =   248
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   10
            Left            =   840
            TabIndex        =   247
            Top             =   930
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Elbow"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   246
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   245
            Top             =   420
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "1½"
            Height          =   255
            Index           =   11
            Left            =   840
            TabIndex        =   244
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   243
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "4½"
            Height          =   255
            Index           =   13
            Left            =   840
            TabIndex        =   242
            Top             =   1740
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   255
            Index           =   14
            Left            =   840
            TabIndex        =   241
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "7½"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   240
            Top             =   2310
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "9"
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   239
            Top             =   2580
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "10½"
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   238
            Top             =   2820
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "12"
            Height          =   255
            Index           =   18
            Left            =   840
            TabIndex        =   237
            Top             =   3090
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "13½"
            Height          =   255
            Index           =   19
            Left            =   840
            TabIndex        =   236
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "15"
            Height          =   255
            Index           =   20
            Left            =   840
            TabIndex        =   235
            Top             =   3630
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "16½"
            Height          =   255
            Index           =   21
            Left            =   840
            TabIndex        =   234
            Top             =   3900
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "18"
            Height          =   255
            Index           =   22
            Left            =   840
            TabIndex        =   233
            Top             =   4140
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "19½"
            Height          =   255
            Index           =   23
            Left            =   840
            TabIndex        =   232
            Top             =   4440
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   231
            Top             =   120
            Width           =   555
         End
         Begin VB.Label lblArm 
            Caption         =   " MMs"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   230
            Top             =   150
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Gram"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   229
            Top             =   150
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Red"
            Height          =   255
            Index           =   6
            Left            =   3780
            TabIndex        =   228
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   8145
         Left            =   120
         TabIndex        =   89
         Top             =   120
         Width           =   4605
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   20
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   21
            Top             =   780
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   22
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   23
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   24
            Top             =   1680
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   25
            Top             =   1980
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   26
            Top             =   2280
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   27
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   28
            Top             =   2880
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   29
            Top             =   3180
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   30
            Top             =   3480
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   40
            Top             =   7260
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   39
            Top             =   6960
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   38
            Top             =   6690
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   37
            Top             =   6240
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   36
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   35
            Top             =   5640
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   34
            Top             =   5340
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   33
            Top             =   5040
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   41
            Top             =   7560
            Width           =   555
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   106
            Top             =   5040
            Width           =   1095
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   109
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   108
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   107
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   102
            Top             =   5340
            Width           =   1095
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   105
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   104
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   103
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   98
            Top             =   5640
            Width           =   1095
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   101
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   100
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   99
               Top             =   60
               Width           =   195
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   94
            Top             =   5940
            Width           =   1095
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   97
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   96
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   95
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   90
            Top             =   6240
            Width           =   1095
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   93
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   92
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   91
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   31
            Top             =   3780
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   32
            Top             =   4080
            Width           =   555
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   185
            Top             =   525
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  1"
            Height          =   255
            Index           =   0
            Left            =   1140
            TabIndex        =   184
            Top             =   540
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  2"
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   183
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   182
            Top             =   825
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  3"
            Height          =   255
            Index           =   2
            Left            =   1140
            TabIndex        =   181
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   180
            Top             =   1125
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  4"
            Height          =   255
            Index           =   3
            Left            =   1140
            TabIndex        =   179
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   178
            Top             =   1425
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  5"
            Height          =   255
            Index           =   4
            Left            =   1140
            TabIndex        =   177
            Top             =   1740
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   176
            Top             =   1725
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  6"
            Height          =   255
            Index           =   5
            Left            =   1140
            TabIndex        =   175
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   174
            Top             =   2025
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  7"
            Height          =   255
            Index           =   6
            Left            =   1140
            TabIndex        =   173
            Top             =   2340
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   172
            Top             =   2325
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  8"
            Height          =   255
            Index           =   7
            Left            =   1140
            TabIndex        =   171
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   170
            Top             =   2625
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb  9"
            Height          =   255
            Index           =   8
            Left            =   720
            TabIndex        =   169
            Top             =   2925
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   168
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Palm 10"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   167
            Top             =   3225
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   166
            Top             =   3255
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Little Finger"
            Height          =   405
            Index           =   20
            Left            =   300
            TabIndex        =   165
            Top             =   540
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Ring Finger"
            Height          =   405
            Index           =   21
            Left            =   300
            TabIndex        =   164
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Middle Finger"
            Height          =   405
            Index           =   22
            Left            =   300
            TabIndex        =   163
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Index Finger"
            Height          =   405
            Index           =   23
            Left            =   300
            TabIndex        =   162
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Wrist 11"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   161
            Top             =   3540
            Width           =   1020
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   160
            Top             =   7590
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   159
            Top             =   7290
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   158
            Top             =   6990
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   157
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Clsd."
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   156
            Top             =   4770
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Std."
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   155
            Top             =   4770
            Width           =   375
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   154
            Top             =   6270
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   149
            Top             =   5970
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   148
            Top             =   5370
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   147
            Top             =   5070
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Wrist To Web"
            Height          =   645
            Index           =   26
            Left            =   300
            TabIndex        =   146
            Top             =   6930
            Width           =   435
         End
         Begin VB.Label labLen 
            Alignment       =   2  'Center
            Caption         =   "Tip To Web"
            Height          =   645
            Index           =   5
            Left            =   300
            TabIndex        =   145
            Top             =   5400
            Width           =   435
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 17"
            Height          =   255
            Index           =   28
            Left            =   960
            TabIndex        =   144
            Top             =   6705
            Width           =   795
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 18"
            Height          =   255
            Index           =   29
            Left            =   960
            TabIndex        =   143
            Top             =   6990
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 19"
            Height          =   255
            Index           =   30
            Left            =   780
            TabIndex        =   142
            Top             =   7290
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 20"
            Height          =   255
            Index           =   31
            Left            =   840
            TabIndex        =   141
            Top             =   7605
            Width           =   915
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 16"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   140
            Top             =   6270
            Width           =   915
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Index 15"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   139
            Top             =   5970
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 14"
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   138
            Top             =   5655
            Width           =   975
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 13"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   137
            Top             =   5355
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 12"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   136
            Top             =   5070
            Width           =   795
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
            Left            =   180
            TabIndex        =   135
            Top             =   210
            Width           =   1575
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
            Left            =   180
            TabIndex        =   134
            Top             =   4680
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tips"
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
            Index           =   5
            Left            =   3480
            TabIndex        =   133
            Top             =   4500
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Desird"
            Height          =   195
            Index           =   2
            Left            =   3900
            TabIndex        =   132
            Top             =   4770
            Width           =   555
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "1½"" Beyond Wrist"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   131
            Top             =   3840
            Width           =   1620
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "3"" Beyond Wrist"
            Height          =   255
            Index           =   12
            Left            =   180
            TabIndex        =   130
            Top             =   4140
            Width           =   1560
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   129
            Top             =   5670
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   128
            Top             =   3540
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   127
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   12
            Left            =   2400
            TabIndex        =   126
            Top             =   4110
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   5325
      Left            =   5160
      TabIndex        =   19
      Top             =   2460
      Width           =   4275
      Begin VB.CheckBox chkFusedFingers 
         Caption         =   "Fused Fingers"
         Height          =   225
         Left            =   2040
         TabIndex        =   87
         Top             =   720
         Width           =   1755
      End
      Begin VB.Frame Frame7 
         Caption         =   "Fold "
         Height          =   645
         Left            =   120
         TabIndex        =   82
         Top             =   3420
         Width           =   4035
         Begin VB.CheckBox chkPrintFold 
            Caption         =   "Print"
            Height          =   225
            Left            =   2760
            TabIndex        =   86
            Top             =   270
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.OptionButton optFold 
            Caption         =   "OFF Fold"
            Height          =   315
            Index           =   0
            Left            =   1380
            TabIndex        =   84
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optFold 
            Caption         =   "ON Fold"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Insert"
         Height          =   1995
         Left            =   120
         TabIndex        =   74
         Top             =   1380
         Width           =   4035
         Begin VB.CommandButton cmdCalculateInsert 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   2640
            TabIndex        =   85
            Top             =   270
            Width           =   975
         End
         Begin VB.ComboBox cboInsert 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   300
            Width           =   1275
         End
         Begin VB.OptionButton optInsertStyle 
            Caption         =   "One Insert Glove"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   78
            Top             =   1560
            Width           =   1935
         End
         Begin VB.OptionButton optInsertStyle 
            Caption         =   "Normal"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   660
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optInsertStyle 
            Caption         =   "One Insert, Little Finger Only (Off Fold)"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   3855
         End
         Begin VB.OptionButton optInsertStyle 
            Caption         =   "One Insert, Little and Index Finger"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   75
            Top             =   1260
            Width           =   3255
         End
         Begin VB.Label labInsert 
            Caption         =   "Size:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   555
         End
         Begin VB.Label labInsert 
            Caption         =   "Inches"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   80
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Thumb"
         Height          =   1185
         Left            =   2160
         TabIndex        =   70
         Top             =   4050
         Width           =   1995
         Begin VB.OptionButton optThumbStyle 
            Caption         =   "Radial Set In"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   73
            Top             =   810
            Width           =   1575
         End
         Begin VB.OptionButton optThumbStyle 
            Caption         =   "Edged Opening"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   510
            Width           =   1815
         End
         Begin VB.OptionButton optThumbStyle 
            Caption         =   "Normal"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Variations"
         Height          =   1185
         Left            =   120
         TabIndex        =   61
         Top             =   4050
         Width           =   1995
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Normal"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   270
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Glove to Elbow"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   540
            Width           =   1815
         End
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Glove to Axilla"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   60
            Top             =   810
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkDorsal 
         Caption         =   "Reinforced Dorsal"
         Height          =   225
         Left            =   2040
         TabIndex        =   57
         Top             =   1050
         Width           =   1875
      End
      Begin VB.CheckBox chkPalm 
         Caption         =   "Reinforced Palm"
         Height          =   225
         Left            =   120
         TabIndex        =   55
         Top             =   1050
         Width           =   1755
      End
      Begin VB.CheckBox chkSlantedInserts 
         Caption         =   "Slant Inserts"
         Height          =   225
         Left            =   120
         TabIndex        =   54
         Top             =   750
         Width           =   1395
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "cboFabric"
         Top             =   330
         Width           =   3735
      End
   End
   Begin VB.TextBox txtFlap 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   69
      Text            =   "txtFlap"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWristPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   68
      Text            =   "txtWristPleat"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtShoulderPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   67
      Text            =   "txtShoulderPleat"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDataGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   66
      Text            =   "txtDataGlove"
      Top             =   5130
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtReduction 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   65
      Text            =   "txtReduction"
      Top             =   4830
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   64
      Text            =   "txtTapeMMs"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtGrams 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   63
      Text            =   "txtGrams"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengthPt1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   0
      Text            =   "txtTapeLengthsPt1"
      Top             =   6510
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   7080
   End
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      TabIndex        =   56
      Top             =   7080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSide 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   51
      Text            =   "txtSide"
      Top             =   3630
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   50
      Top             =   8040
      Width           =   1395
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   435
      Left            =   5520
      TabIndex        =   49
      Top             =   8040
      Width           =   1395
   End
   Begin VB.TextBox txtUidGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   47
      Text            =   "txtUidGlove"
      Top             =   2730
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   46
      Text            =   "txtTapeLengths2"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   45
      Text            =   "txtTapeLengths"
      Top             =   3030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Glove"
      Height          =   585
      Left            =   5160
      TabIndex        =   17
      Top             =   0
      Width           =   4275
      Begin VB.OptionButton optHand 
         Caption         =   "Right"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Left"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   1785
      Left            =   5160
      TabIndex        =   18
      Top             =   630
      Width           =   4275
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   13
         Top             =   660
         Width           =   3075
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   1380
         Width           =   525
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3060
         TabIndex        =   16
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtUnits 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   1020
         Width           =   3075
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   2580
         TabIndex        =   5
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2700
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   42
      Text            =   "txtWorkOrder"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidGC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   43
      Text            =   "txtUidGC"
      Top             =   300
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   48
      Text            =   "txtFabric"
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   52
      Text            =   "txtUidMPD"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDataGC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   62
      Text            =   "txtDataGC"
      Top             =   870
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   9240
      TabIndex        =   53
      Top             =   1440
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
      Left            =   9300
      TabIndex        =   1
      Top             =   2490
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
      Left            =   9240
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "manglove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   MANGLOVE.MAK
'Purpose:   Manual glove
'
'
'Version:   3.01
'Date:      17.Jan.96
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
   'Windows API Functions Declarations
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

   'Open tip constants
    Const AGE_CUTOFF = 10           'Years
    Const ADULT_STD_OPEN_TIP = 0.5  'Inches
    Const CHILD_STD_OPEN_TIP = 0.375 'Inches

    Option Explicit

Private Sub Cancel_Click()
    Dim Response As Integer, sTask$
    
   'Check if data has been modified
    If FN_GloveDataChanged() Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, g_sDialogID)
        Select Case Response
            Case IDYES
                PR_UpdateDDE
                PR_UpdateDDE_Extension
                PR_CreateSaveMacro "c:\jobst\draw.d"
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate fnGetDrafixWindowTitleText()
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, g_sDialogID
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

Private Sub cboFlaps_Click()
    
    If InStr(1, MainForm!cboFlaps.Text, "D") > 0 Then
        MainForm!lblFlap(4).Enabled = True
        MainForm!txtWaistCir.Enabled = True
        MainForm!labWaistCir.Enabled = True
    ElseIf MainForm!lblFlap(4).Enabled = True Then
        MainForm!txtWaistCir.Text = ""
        MainForm!labWaistCir.Caption = ""
        MainForm!lblFlap(4).Enabled = False
        MainForm!txtWaistCir.Enabled = False
        MainForm!labWaistCir.Enabled = False
    End If

    If MainForm.Visible Then PR_CalculateArmTapeReductions

End Sub

Private Sub cmdCalculate_Click()
    PR_CalculateArmTapeReductions
End Sub

Private Static Sub cmdCalculateInsert_Click()
    Dim iError%, iIndex%
    If FN_ValidateData("Ignore Webs") Then
       'Common Dialogue elements
        If optFold(1).Value = True Then g_OnFold = True Else g_OnFold = False
        PR_GetDlgHandData
        PR_LoadReductionCharts (cboFabric.Text)
        If FN_CalculateInsert(g_iInsertSize, "CALC") Then
           'Translate and display
            Select Case g_iInsertSize
                Case 1 To 5
                    iIndex = g_iInsertSize + 3
                Case 6
                    iIndex = 3
            End Select
    
           'We need to force a calculation if the user changes
           'fold options but if this is done before we have enough
           'data we will confuse the user as they have not explititly
           'calculated.

            g_DataIsCalcuable = True

        Else
           'Set Drop down to calculate on fail
            iIndex = 0
        End If
        
        cboInsert.ListIndex = iIndex

    End If

End Sub

Private Function FN_CalcGlovePoints(iInsert As Integer) As Integer
   'Calculate the points on the glove based on the insert given
   'and the data in the dialogue
   'Note:-
   ' Figure both Left and Right the same.
   ' We then mirror the result in the vertical line
   ' given by X = xyDatum.x
    FN_CalcGlovePoints = True
    
    Dim nFiguredPalm As Double, nFiguredWrist As Double

    Dim nPIPs As Double, nMaxWeb As Double
    Dim nFiguredPIPs As Double, nMissingPIPs As Double
    Dim nn%, ii%, Response%
    Dim nInsert#, nSlantInsert#, nThumbSlantLength#, aPalmThumbEnd#, nMinDist#, nInc#
    Dim nThumbLineToWeb#, nMinRadius#, nLength#
    Dim xyMinRadius As XY, xyMinRadiusC1  As XY, xyThumbOriginal As XY, xyThumbPalm5 As XY
    Dim Success%, ForcedFail%, iMaxLoopCount%, iLoopCount%
    Dim nCheckDistance#

    nInsert = g_iInsertValue(iInsert) * EIGHTH
    nSlantInsert = 0
    nInc = 0.03125 '1/32nd of an inch

   'Calculate Figured DIPs and PIPs for the fingers
    For nn = 1 To 4
        If g_iFINGER_CHART(nn) = THUMB Then
            g_Finger(nn).PIP = FN_FigureFinger(THUMB, g_nFingerPIPCir(nn), iInsert)
            g_Finger(nn).DIP = FN_FigureFinger(THUMB, g_nFingerDIPCir(nn), iInsert)
        Else
            g_Finger(nn).PIP = FN_FigureFinger(PIP, g_nFingerPIPCir(nn), iInsert)
            g_Finger(nn).DIP = FN_FigureFinger(DIP, g_nFingerDIPCir(nn), iInsert)
        End If
       'If not missing then minimun value is 3/8th"
        If g_Finger(nn).PIP > 0 Then
            g_Finger(nn).PIP = g_Finger(nn).PIP + g_nFINGER_FIGURE(nn)
            If g_Finger(nn).PIP < 0.375 And g_Finger(nn).PIP > 0 Then g_Finger(nn).PIP = 0.375
            g_Finger(nn).DIP = g_Finger(nn).DIP + g_nFINGER_FIGURE(nn)
            If g_Finger(nn).DIP < 0.375 And g_Finger(nn).DIP > 0 Then g_Finger(nn).DIP = 0.375
        End If
        nFiguredPIPs = nFiguredPIPs + g_Finger(nn).PIP
    Next nn

   'Figure palm and wrist
    nFiguredPalm = (FN_FigureTape(PALM, g_nPalm) + g_nPALM_FIGURE) / 2
    nFiguredWrist = (FN_FigureTape(WRIST, g_nWrist) + g_nPALM_FIGURE) / 2
  
   'Thumb
    If g_iThumbStyle = 2 Then
       'Setin thumb
        g_Finger(5).PIP = FN_FigureFinger(PIP, g_nFingerPIPCir(5), iInsert)
    ElseIf g_iThumbStyle = 0 Then
       'Normal thumbs
        g_Finger(5).PIP = FN_FigureFinger(THUMB, g_nFingerPIPCir(5), iInsert)
    Else
       'Edged thumbs
       'Do not figure
        g_Finger(5).PIP = g_nFingerPIPCir(5)
    End If
    g_Finger(5).DIP = 0

   'Check for missing fingers
   'If there are missing fingers then we need establish the width of the
   'PIP for the missing fingers, we use the Figured palm as the total
   'width of the PIPs, subtract the actual total PIPs and spread the rest
   'over the missing fingers
   'Note the exception for Fused Fingers
    If g_MissingFingers > 0 And Not g_FusedFingers Then
        nMissingPIPs = Abs(nFiguredPalm - nFiguredPIPs) / g_MissingFingers
        For ii = 1 To 4
            If g_Missing(ii) = True Then g_Finger(ii).PIP = nMissingPIPs
        Next ii
    End If
    
   
   'Establish Datum point
    If g_OnFold Then
        PR_MakeXY xyDatum, 0, FN_LengthWristToEOS()
    Else
        PR_MakeXY xyDatum, 1, FN_LengthWristToEOS()
    End If

   'Missing web heights can represent amputated fingers etc
   'we need to account for these

   'Missing thumb web (use lowest - 0.375)
    If g_nWeb(4) = 0 Then
        g_nWeb(4) = 1000000 'Dummy high value
        For ii = 3 To 1 Step -1
            If g_nWeb(ii) < g_nWeb(4) Then g_nWeb(4) = g_nWeb(ii)
        Next ii
        g_nWeb(4) = g_nWeb(4) - 0.375
    End If

   'Missing middle
    If g_nWeb(3) = 0 Then g_nWeb(3) = g_nWeb(4) + 0.25
    
   'Missing Ring
    If g_nWeb(2) = 0 And g_nWeb(1) <> 0 Then g_nWeb(2) = g_nWeb(1) 'Use Little
    If g_nWeb(2) = 0 And g_nWeb(1) = 0 Then g_nWeb(2) = g_nWeb(3)  'Use Middle if no little

   'Missing Little
    If g_nWeb(1) = 0 Then g_nWeb(1) = g_nWeb(2)  'Use Ring
    
   'Little Finger (at web)
    nPIPs = g_Finger(1).PIP
    PR_MakeXY xyLittle, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(1)
    
   'Little Finger at side of hand
    PR_MakeXY xyLFS, xyDatum.X, xyLittle.Y
    
   'Ring finger
    nPIPs = g_Finger(1).PIP + g_Finger(2).PIP
    PR_MakeXY xyRing, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(2)
    
   'Middle finger
    nPIPs = g_Finger(1).PIP + g_Finger(2).PIP + g_Finger(3).PIP
    PR_MakeXY xyMiddle, xyDatum.X + nPIPs, xyDatum.Y + g_nWeb(3)

   'Index finger
    nPIPs = g_Finger(1).PIP + g_Finger(2).PIP + g_Finger(3).PIP + g_Finger(4).PIP
    PR_MakeXY xyIndex, xyDatum.X + nPIPs, xyMiddle.Y

   'Thumb
    If g_OnFold Then
        PR_MakeXY xyThumb, xyDatum.X + nFiguredPalm, xyDatum.Y + g_nWeb(4)
    Else
        If nPIPs <> nFiguredPalm Then
           'Avoid overflow
            PR_MakeXY xyThumb, xyDatum.X + (((nPIPs - nFiguredPalm) / 2) + nFiguredPalm), xyDatum.Y + g_nWeb(4)
        Else
            PR_MakeXY xyThumb, xyDatum.X + nFiguredPalm, xyDatum.Y + g_nWeb(4)
        End If
    End If
   
   'Set finger points
    g_Finger(1).xyUlnar = xyLFS
    g_Finger(1).xyRadial = xyLittle
    g_Finger(1).WebHt = xyLittle.Y

    g_Finger(2).xyUlnar = xyLittle
    g_Finger(2).xyRadial = xyRing
    g_Finger(2).WebHt = xyRing.Y

    g_Finger(3).xyUlnar = xyRing
    g_Finger(3).xyRadial = xyMiddle
    g_Finger(3).WebHt = xyMiddle.Y

    g_Finger(4).xyUlnar = xyMiddle
    g_Finger(4).xyRadial = xyIndex
    g_Finger(4).WebHt = xyMiddle.Y


   'Palm points
   'Wrist line
    If g_OnFold Then
        PR_MakeXY xyPalm(1), xyLFS.X, xyDatum.Y
    Else
        If nPIPs <> nFiguredWrist Then
           'Avoids overflow
            PR_MakeXY xyPalm(1), xyDatum.X + ((nPIPs - nFiguredWrist) / 2), xyDatum.Y
        Else
            PR_MakeXY xyPalm(1), xyDatum.X, xyDatum.Y
        End If
    End If

   'Offset from wrist line
    If g_OnFold Then
        xyPalm(2).X = xyLFS.X
    Else
        If nPIPs <> nFiguredPalm Then
           'Avoids overflow
            xyPalm(2).X = xyLFS.X + (nPIPs - nFiguredPalm) / 2
        Else
            xyPalm(2).X = xyLFS.X
        End If
    End If

   'Establish maximum Web Height to get offset
    nMaxWeb = 0
    For ii = 1 To 3
        If g_nWeb(ii) > nMaxWeb Then nMaxWeb = g_nWeb(ii)
    Next ii

    If nMaxWeb > 3.5 Then
        xyPalm(2).Y = xyDatum.Y + 1
    Else
        xyPalm(2).Y = xyDatum.Y + 0.5
    End If

   'Offset from wrist at Thumb Side
    xyPalm(5).X = xyPalm(2).X + nFiguredPalm
    xyPalm(5).Y = xyPalm(2).Y

   'Wrist at Thumb Side
    xyPalm(6).X = xyPalm(1).X + nFiguredWrist
    xyPalm(6).Y = xyPalm(1).Y

   
   'Figure Lengths for fingers and thumbs
    Dim nTipCutBack#, iDigitType%

    If Val(txtAge.Text) <= AGE_CUTOFF Then
        nTipCutBack = CHILD_STD_OPEN_TIP
    Else
        nTipCutBack = ADULT_STD_OPEN_TIP
    End If

    iDigitType = FINGER_LEN
    For ii = 1 To 5
        If ii = 5 Then iDigitType = THUMB_LEN 'Swap for thumbs
        Select Case g_Finger(ii).Tip
        Case 0, 10
           'Closed tips (Figured Length)
            g_Finger(ii).Len = FN_FigureLength(iDigitType, g_nFingerLen(ii))
        Case 1, 11
           'Standard Cut back for open tips (GivenLength - StdCutBack)
            g_Finger(ii).Len = g_nFingerLen(ii) - nTipCutBack
        Case 2, 12
           'Open at a desired given length
            g_Finger(ii).Len = g_nFingerLen(ii)
        End Select
    Next ii

   'Thumbpoints on palm
   'Angles
    aPalmThumbEnd = 0
    
    xyThumbOriginal = xyThumb  'Store original to give a message later
    xyPalmThumb(1) = xyThumb

    If g_iThumbStyle <> 1 Then
       'Normal or SetIn thumbs
        xyPalmThumb(1).Y = xyThumb.Y + 0.75

       'Ensure mininum distance of 3/8" from Index web
        If (xyIndex.Y - xyPalmThumb(1).Y) < 0.375 Then
            xyPalmThumb(1).Y = xyIndex.Y - 0.375
        End If
       'Ensure mininum distance of 1/8" from top of thumb curve to thumb web
        If xyPalmThumb(1).Y - xyThumb.Y < 0.125 Then
                xyThumb.Y = xyPalmThumb(1).Y - 0.125
        End If
        nThumbSlantLength = g_Finger(5).PIP / (1 / Sqr(2))

       'This loop counter ensures that the bottom of thumb curve does not go closer
       'than 0.25" to the EOS

'        iMaxLoopCount = Int((((xyThumb.Y - xyPalm(6).Y) - nThumbSlantLength)) / .0625)
        iMaxLoopCount = min(25, Int((((xyThumb.Y - 0.25) - nThumbSlantLength)) / 0.0625))
        iLoopCount = 0
        Do 'at least once
           'Get blending curve and revise xyPalmThumb(1) X value
            PR_CalcFingerBlendingProfile xyIndex, xyThumb, xyMiddle, g_Finger(4), g_FingerThumbBlend, xyPalmThumb(1)
        
            xyPalmThumb(2).Y = xyPalmThumb(1).Y
            xyPalmThumb(2).X = xyPalmThumb(1).X - nInsert
        
            xyPalmThumb(5).Y = xyPalmThumb(2).Y
 '           xyPalmThumb(5).X = xyPalmThumb(2).X - .125
            xyPalmThumb(5).X = xyPalmThumb(2).X - 0.25
        
            xyPalmThumb(3).Y = xyThumb.Y
            xyPalmThumb(3).X = xyPalmThumb(1).X - ((nFiguredPalm / 2) - 0.125)

           'Check that Point 3 is always less than point 5, if it is not
           'then make it so by an eighth
            If xyPalmThumb(5).X - xyPalmThumb(3).X < 0.125 Then xyPalmThumb(3).X = xyPalmThumb(5).X - 0.125

    
           'Get top thumb curve
           'We find the top curve by iteration, accepting the first
           'by changing the X value of xyPalmThumb(5)
            nMinDist = min(Abs(xyPalmThumb(3).Y - xyPalmThumb(5).Y), Abs(xyPalmThumb(3).X - xyPalmThumb(5).X))
            xyThumbPalm5 = xyPalmThumb(5)  'retain for later use
            Success = False
            ForcedFail = False
            nCheckDistance = Abs(xyPalmThumb(3).X - xyPalmThumb(5).X)
           'Try first moving towards the center of the palm away from xyPalmThumb(2)
            While Not Success And nCheckDistance > nMinDist And Not ForcedFail
                ii = FN_BiArcCurve(xyPalmThumb(3), 90, xyPalmThumb(5), aPalmThumbEnd, g_ThumbTopCurve)

                If ii Then
                   'Check that the distance from the thumbline is more than 3/8th from xyMiddle
                   'allow adjustement to continue the other way if it is
                    nLength = FN_CalcLength(g_ThumbTopCurve.xyR1, xyRing) - g_ThumbTopCurve.nR1
                    nLength = round(nLength / 0.0625)
'                    If nLength < .375 Then
                    If nLength < 6 Then
                        ForcedFail = True
                    Else
                        Success = True
                    End If
                Else
                   'Adjust and try again BUT!
                   'Check that xyPalmThumb(5).x - xyPalmThumb(3) is not silly
                   'IE stop any infinite loops.
'                    xyPalmThumb(5).X = xyPalmThumb(5).X + nInc
                    xyPalmThumb(5).X = xyPalmThumb(5).X - nInc
                    If Abs(xyPalmThumb(3).X - xyPalmThumb(5).X) > nCheckDistance Then ForcedFail = True
                    nCheckDistance = Abs(xyPalmThumb(3).X - xyPalmThumb(5).X)
                End If
            Wend
           
           'If the above does not work then try moving away from the center
           'of the palm towards (but not past) xyPalmThumb(2)
            If Not Success Then
             xyPalmThumb(5) = xyThumbPalm5
             For nLength = xyPalmThumb(5).X To xyPalmThumb(2).X Step (nInc / 2)
                ii = FN_BiArcCurve(xyPalmThumb(3), 90, xyPalmThumb(5), aPalmThumbEnd, g_ThumbTopCurve)
                If ii Then
                   'Check that the distance from the thumbline is more than 3/8th from xyRing
                   'allow adjustement to continue until fail
                    If (FN_CalcLength(g_ThumbTopCurve.xyR1, xyRing) - g_ThumbTopCurve.nR1) >= 0.375 Then
                        Success = True
                        Exit For
                    Else
                       'Adjust and try again
                        xyPalmThumb(5).X = nLength
                    End If
                Else
                   'Adjust and try again
                    xyPalmThumb(5).X = nLength
                End If
             Next nLength
            End If

            iLoopCount = iLoopCount + 1

            If Success Or iLoopCount >= iMaxLoopCount Then Exit Do

    
           'If we don't have a sucessfull top thumb line then lower
           'the top line or if this is not possible then drop xyThumb
           'Do this in increments of nDrop
            Dim nDrop#
            nDrop = 0.0625
           'Drop top thumb curve (Ensure a minimum of 1/8th inch)
            If (xyPalmThumb(1).Y - xyThumb.Y) >= 0.25 Then
               'Drop top curve
                xyPalmThumb(1).Y = xyPalmThumb(1).Y - 0.125
            Else
               'Drop thumb
                xyThumb.Y = xyThumb.Y - nDrop
            End If
    
        
        Loop

        If Not Success Then
            MsgBox "Can't form TOP part of thumb curve on palm!. Adjust the palm curve and use the Transfer tool to update the thumb", 48, g_sDialogID
           'Degenerate "ThumbTopCurve" to a straight line
            xyPalmThumb(5) = xyThumbPalm5
            g_ThumbTopCurve.xyStart = xyPalmThumb(3)
            g_ThumbTopCurve.xyEnd = xyThumbPalm5
            g_ThumbTopCurve.nR1 = 0
            g_ThumbTopCurve.nR2 = 0
        Else
           'Check and warn
            If FN_CalcLength(g_ThumbTopCurve.xyR1, xyLittle) - g_ThumbTopCurve.nR1 < 0.375 Then
                Response = MsgBox("The insert used creates a Top thumb curve on the palm that is closer than 3/8ths to the little finger web!" & Chr$(13) & "Use YES to continue or NO to return to the dialogue.", 52, g_sDialogID)
            End If
            If xyThumbOriginal.Y <> xyThumb.Y Then
                nLength = Abs(xyThumbOriginal.Y - xyThumb.Y)
                Response = MsgBox("Warnimg the thumb web position has been lowered by" & Str$(nLength) & " inches, so that the thumb curve is 3/8ths away from closest web!" & Chr$(13) & "Use YES to continue or NO to return to the dialogue.", 52, g_sDialogID)
               'Save the drop for later use with Web spacers
                g_iThumbWebDrop = Int(nLength / 0.0625) 'this will give the value in 1/16 ths
            End If
            Select Case Response
                Case IDYES
                    FN_CalcGlovePoints = True
                Case IDNO, IDCANCEL
                    FN_CalcGlovePoints = False
                    Exit Function
            End Select

        End If

    Else

       'Edged thumb
        nThumbSlantLength = g_Finger(5).PIP / 2
        PR_CalcFingerBlendingProfile xyIndex, xyThumb, xyMiddle, g_Finger(4), g_FingerThumbBlend, xyPalmThumb(1)

    End If

    xyPalmThumb(4).X = xyThumb.X
    xyPalmThumb(4).Y = xyThumb.Y - nThumbSlantLength

   'Palm points (Points dependant on thumb y)
   'LFS at Thumb web height
    If g_OnFold Then
        xyPalm(3).X = xyLFS.X
    Else
        xyPalm(3).X = xyPalm(2).X
    End If

    xyPalm(3).Y = xyThumb.Y
    xyPalm(4) = xyThumb
    
End Function

Private Function FN_CalculateInsert(iInsert As Integer, sType As String)
   'This procedure uses the data for the finger PIP and the
   'Palm circumference to calculate the insert sizes
   'based on the procedure
   'GOP 01/02/26, Effective date 2nd June 1992
   '
   '
   'Notes:-
   '
    Dim nn%, ii%, Response%, nAge
    Dim nFiguredPIPs#, nFiguredPalm#, nPermittedDiff#
    Dim sText$
    nAge = Val(MainForm!txtAge.Text)

'Dim ifile%
   'Initially true
    FN_CalculateInsert = True

   'CALCULATE an insert
   'setup the parameters based on radio buttons
'ifile = FreeFile
'Open "C:\TMP\ERR.DAT" For Output As ifile
'If g_FusedFingers = True Then Print #ifile, "Fused Fingers"

   'Loop through the inserts starting at
   '1/2" (ii=1) to 1" (ii=5) then try as a last resort 3/8" (ii=6)
   'If we find an insert that satisfies our test criteria then break the
   'for next loop
   'NB addition wrt Fused Fingers
   '
    For ii = 1 To 4
        g_nFINGER_FIGURE(ii) = 0
    Next ii

    Dim iStart%, iEnd%
    If sType = "CALC" Then
        iStart = 1
        iEnd = 6
    Else
        iStart = iInsert
        iEnd = iInsert
    End If

    For iInsert = iStart To iEnd
        Select Case g_iInsertStyle
          Case 0
           'Normal
            If g_OnFold Then
               'Use thumb for little finger
                g_iFINGER_CHART(1) = THUMB
                g_nFINGER_FIGURE(1) = -(2 * SIXTEENTH)
                For ii = 2 To 4
                    g_iFINGER_CHART(ii) = PIP
                Next ii
                g_nPALM_FIGURE = EIGHTH - (g_iInsertValue(iInsert) * EIGHTH)
            Else
                For ii = 1 To 4
                    g_iFINGER_CHART(ii) = PIP
                Next ii
                g_nPALM_FIGURE = (3 * EIGHTH) - ((g_iInsertValue(iInsert) * EIGHTH) * 2)
            End If
          
          Case 1
            g_iFINGER_CHART(1) = THUMB
            For ii = 2 To 4
                g_iFINGER_CHART(ii) = PIP
            Next ii
            g_nPALM_FIGURE = (3 * EIGHTH) - (g_iInsertValue(iInsert) * EIGHTH)
          
          Case 2
            g_iFINGER_CHART(1) = THUMB
            g_iFINGER_CHART(2) = PIP
            g_iFINGER_CHART(3) = PIP
            g_iFINGER_CHART(4) = THUMB
            g_nPALM_FIGURE = (2 * EIGHTH)
          
          Case 3
            For ii = 1 To 4
                g_iFINGER_CHART(ii) = THUMB
            Next ii
            g_nPALM_FIGURE = (2 * EIGHTH)

        End Select

       'set permitted differencses
        If g_OnFold Then
            nPermittedDiff = EIGHTH
        Else
            nPermittedDiff = 2 * EIGHTH
        End If

       'Common figuring used latter
        g_iFINGER_CHART(5) = THUMB
        g_nWRIST_FIGURE = g_nPALM_FIGURE

       'Figure fingers
        nFiguredPIPs = 0
        For ii = 1 To 4
            nFiguredPIPs = nFiguredPIPs + FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert) + g_nFINGER_FIGURE(ii)
        Next ii
        
       'Figure palm
        nFiguredPalm = (FN_FigureTape(PALM, g_nPalm) + g_nPALM_FIGURE) / 2
'Dim nValue#
'Print #ifile, "iInsert= "; iInsert
'Print #ifile, "nFiguredPalm = "; nFiguredPalm
'Print #ifile, "nFiguredPIPs = "; nFiguredPIPs
'For ii = 1 To 4
'nValue = FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert) + g_nFINGER_FIGURE(ii)
'Print #ifile, "Finger"; ii
'Print #ifile, "Figured"; nValue; "Cir="; g_nFingerPIPCir(ii); "Figured= "; FN_FigureFinger(g_iFINGER_CHART(ii), g_nFingerPIPCir(ii), iInsert); "Figure all="; g_nFINGER_FIGURE(ii)
'Next ii
       'Exit "FOR" if conditions met
        If Abs(nFiguredPalm - nFiguredPIPs) <= nPermittedDiff Or (g_MissingFingers <> 0 And Not g_FusedFingers) Then Exit For
        If sType <> "CALC" Then Exit For  'Exit if we have been given an Insert size
        
    Next iInsert

'Close #ifile
   'Check to see if all of the above have been tried
   'issue a warning message and exit this sub
    If iInsert >= 7 Then
        If g_OnFold Then
            sText = "Unable to calculate an insert. Try OFF the fold"
        Else
            sText = "Unable to calculate an insert. Try again using one of the other options."
        End If
        MsgBox sText, 16, "MANUAL Glove - Insert Calculation"
        FN_CalculateInsert = False
        Exit Function
    End If

   'Check and warn with respect to age and insert size
    If ((iInsert = 4 Or iInsert = 5) And nAge < 6) Then
        sText = "It is not recommended that an insert of 7/8"" or greater is given for children under 6 years old. Try using a ""one insert glove - on the fold""." & Chr$(13) & "Do you wish to continue?"
        Response = MsgBox(sText, 36, "MANUAL Glove - Insert Calculation")
        Select Case Response
            Case IDYES
                FN_CalculateInsert = True
            Case IDNO, IDCANCEL
                FN_CalculateInsert = False
        End Select
    End If

    If (sType <> "CALC" And Abs(nFiguredPalm - nFiguredPIPs) > nPermittedDiff) And g_MissingFingers = 0 And Not g_FusedFingers Then
        sText = "Incorrect figuring with the insert specified." & Chr$(13) & "Do you wish to continue?"
        Response = MsgBox(sText, 36, "MANUAL Glove - Insert Calculation")
        Select Case Response
            Case IDYES
                FN_CalculateInsert = True
            Case IDNO, IDCANCEL
                FN_CalculateInsert = False
        End Select
    End If


End Function

Private Static Function FN_CompareWithOtherGlove(sFabric$, Onfold%, iInsertSize%) As Integer
  'Compares the current glove with the other glove
  '
  'Globals
  '     g_iInsertOtherGlv
  '     g_OnFoldOtherGlv
  '     txtFabric.Text
  '     g_sSide
  '
    Dim sText As String, Response As Integer, sThisInsert$, sOtherInsert$

    FN_CompareWithOtherGlove = True
    sText = ""
    
    If g_iInsertOtherGlv = -1 Then Exit Function    'Other glove does not exist

    If sFabric <> txtFabric.Text Then
        sText = g_sSide & " glove fabric is " & sFabric & "." & Chr$(13)
        sText = sText & "Other glove fabric is " & txtFabric.Text & "." & Chr$(13)
    End If

    If g_iInsertOtherGlv <> iInsertSize Then
        sText = sText & "Insert sizes are different." & Chr$(13)
        sThisInsert = Trim$(fnInchestoText(g_iInsertValue(iInsertSize) * EIGHTH))
        If Mid$(sThisInsert, 1, 1) = "-" Then sThisInsert = Mid$(sThisInsert, 2)
        sOtherInsert = Trim$(fnInchestoText(g_iInsertValue(g_iInsertOtherGlv) * EIGHTH))
        If Mid$(sOtherInsert, 1, 1) = "-" Then sOtherInsert = Mid$(sOtherInsert, 2)
        sText = sText & "This glove uses a " & sThisInsert & " insert, The other glove uses a " & sOtherInsert & " insert." & Chr$(13)
    End If
   
    If Onfold <> g_OnFoldOtherGlv Then
        If g_sSide = "Left" And g_OnFoldOtherGlv Then sText = sText & "Right glove is Drawn on the fold" & Chr$(13)
        If g_sSide = "Left" And Not g_OnFoldOtherGlv Then sText = sText & "Right glove is Drawn off the fold" & Chr$(13)
        If g_sSide = "Right" And g_OnFoldOtherGlv Then sText = sText & "Left glove is Drawn on the fold" & Chr$(13)
        If g_sSide = "Right" And Not g_OnFoldOtherGlv Then sText = sText & "Left glove is Drawn off the fold" & Chr$(13)
    End If

   'Check if error has occured
    If sText <> "" Then
        If g_sSide = "Left" Then
            sText = "This glove differs from the RIGHT glove!" & Chr$(13) & sText & Chr$(13) & "Do you wish to continue and draw the glove anyway?"
        Else
            sText = "This glove differs from the LEFT glove!" & Chr$(13) & sText & Chr$(13) & "Do you wish to continue and draw the glove anyway?"
        End If
        Response = MsgBox(sText, 36, g_sDialogID)
        Select Case Response
            Case IDYES
                FN_CompareWithOtherGlove = True
            Case IDNO, IDCANCEL
                FN_CompareWithOtherGlove = False
        End Select
    End If

End Function

Private Static Function FN_GloveDataChanged() As Integer
   'Function that will be used to check if data has changed
   'The first Time that it is called the function will always
   'return False
   '
   'NOTE:-
   '    The change checked for is always the state of the data the
   '    first time the function is called
   '
   '    It is intendee to be run After Link Close to save the state
   '    at that point
   '    If the user uses Close then we can check if any changes have been
   '    made


    Dim sCheck As String, sTmp  As String
    Dim ii%
   
   'Get data from the controls
    sTmp = ""
'    MainForm!SSTab1.Tab = 0
    
    For ii = 0 To 12
        sTmp = sTmp & txtCir(ii).Text
    Next ii
    
    For ii = 0 To 8
        sTmp = sTmp & txtLen(ii).Text
    Next ii

'    MainForm!SSTab1.Tab = 1
    
    For ii = 8 To 23
        sTmp = sTmp & txtExtCir(ii).Text & mms(ii).Text
    Next ii
    
    sTmp = sTmp & txtWristPleat1.Text & txtWristPleat2.Text & txtShoulderPleat1.Text & txtShoulderPleat2.Text

    sTmp = sTmp & cboFlaps.Text & txtStrapLength.Text & txtFrontStrapLength.Text & txtCustFlapLength.Text & txtWaistCir.Text
    
    sTmp = sTmp & Str$(optExtendTo(0).Value) & Str$(optExtendTo(1).Value) & Str$(optExtendTo(2).Value)
    
    sTmp = sTmp & Str$(optThumbStyle(0).Value) & Str$(optThumbStyle(1).Value) & Str$(optThumbStyle(2).Value)
    
    sTmp = sTmp & Str$(optProximalTape(0).Value) & Str$(optProximalTape(1).Value)

    sTmp = sTmp & Str$(optFold(0).Value) & Str$(optFold(1).Value)
    
    sTmp = sTmp & cboFabric.Text & cboDistalTape.Text & cboProximalTape.Text & cboPressure.Text & cboInsert.Text

    sTmp = sTmp & Str$(chkSlantedInserts.Value) & Str$(chkPalm.Value) & Str$(chkDorsal.Value)
   
   'Check for changes
    If sCheck = "" Then
        FN_GloveDataChanged = False
        sCheck = sTmp
    Else
        If sTmp <> sCheck Then FN_GloveDataChanged = True
    End If

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
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengths"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengths2"; QCQ; "string"; QQ; ");"
    Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengthPt1"; QCQ; "string"; QQ; ");"
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
   'Note this is a bit dogey if the drawing contains more than 9999 entities
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
Private Function FN_TranslateInsert(iListIndex As Integer, iInsert As Integer) As Integer
    'Function to translate the cboInsert value to one that can be
    'used in a manual glove
    'It will also attempt reconcile the two
    '
    '   cboInsert.AddItem "Calculate"   0
    '   cboInsert.AddItem "1/8"         1
    '   cboInsert.AddItem "1/4"         2
    '   cboInsert.AddItem "3/8"         3      Valid Manual
    '   cboInsert.AddItem "1/2"         4      Valid Manual
    '   cboInsert.AddItem "5/8"         5      Valid Manual
    '   cboInsert.AddItem "3/4"         6      Valid Manual
    '   cboInsert.AddItem "7/8"         7      Valid Manual
    '   cboInsert.AddItem "1"           8      Valid Manual
    '   cboInsert.AddItem "1-1/8"       9
    '   cboInsert.AddItem "1-1/4"       10
    '   cboInsert.AddItem "1-3/8"       11
    '   cboInsert.AddItem "1-1/2"       12
    '   cboInsert.AddItem "1-5/8"       13
    '   cboInsert.AddItem "1-3/4"       14
    '   cboInsert.AddItem "1-7/8"       15
    '
    FN_TranslateInsert = True
    Select Case iListIndex
        Case 0 To 2, 9 To 15
            MsgBox "Not valid insert size for MANUAL-Gloves. Use a value in the range 3/8th to 1 inch.  Or use Calculate to force a calculation ", 16, g_sDialogID
            FN_TranslateInsert = False
            iInsert = 0
        Case 4 To 8
            iInsert = iListIndex - 3
        Case 3
            iInsert = 6
    End Select
            

End Function

Private Function FN_ValidateData(sOption As String) As Integer
   'This function is used only to make gross checks
   'for missing data.
   'It does not perform any sensibility checks on the
   'data
   'The sOption allows for a less rigourous check
   '    "Check All"
   '    "Ignore Webs"
   '
    Dim sError$, NL$, ii%, nn%, sWebError$, nThumbWebHt#
    
   'Initialise
    FN_ValidateData = False
    sError = ""
    NL = Chr$(13)

    ReDim sCircum(0 To 12)
    sCircum(0) = "Little Finger"
    sCircum(1) = "Little Finger"
    sCircum(2) = "Ring Finger"
    sCircum(3) = "Ring Finger"
    sCircum(4) = "Middle Finger"
    sCircum(5) = "Middle Finger"
    sCircum(6) = "Index Finger"
    sCircum(7) = "Index Finger"
    sCircum(8) = "Thumb"
    sCircum(9) = "Palm"
    sCircum(10) = "Wrist"
    sCircum(11) = "Tape 1½"
    sCircum(12) = "Tape 3"

    ReDim nTotalCir(0 To 4)

    ReDim sLengths(0 To 8)
    sLengths(0) = "Little Finger tip  to web"
    sLengths(1) = "Ring Finger tip to web"
    sLengths(2) = "Middle Finger tip to web"
    sLengths(3) = "Index Finger tip to web"
    sLengths(4) = "Thumb to tip web"
    sLengths(5) = "Wrist to web at Little Finger"
    sLengths(6) = "Wrist to web at Middle Finger"
    sLengths(7) = "Wrist to web at Index Finger"
    sLengths(8) = "Wrist to web at Thumb"


   'Check circumferences
    nn = 0
    For ii = 0 To 7 Step 2
        If Val(txtCir(ii).Text) <> 0 And Val(txtCir(ii + 1).Text) = 0 Then
            sError = sError & "Missing PIP for " & sCircum(ii) & "!" & NL
        End If
        nTotalCir(nn) = Val(txtCir(ii).Text) + Val(txtCir(ii + 1).Text)
        nn = nn + 1
    Next ii

   'Settings for thumb (as not set above)
    nTotalCir(4) = Val(txtCir(8).Text)

    For ii = 8 To 10
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing circumference for " & sCircum(ii) & "!" & NL
        End If
    Next ii

    
   'Check on lengths
    For ii = 0 To 3
        If Val(txtLen(ii).Text) = 0 And nTotalCir(ii) <> 0 Then
            sError = sError & "Missing length " & sLengths(ii) & "!" & NL
        End If
        If Val(txtLen(ii).Text) <> 0 And nTotalCir(ii) = 0 Then
            sError = sError & "Length given but missing circumference " & sCircum(ii * 2) & "!" & NL
        End If
    Next ii

   'Length not required for edged thumbs
    If optThumbStyle(1).Value <> True Then
        If Val(txtLen(4).Text) = 0 And nTotalCir(4) <> 0 Then
            sError = sError & "Missing length " & sLengths(4) & "!" & NL
        End If
        If Val(txtLen(4).Text) <> 0 And nTotalCir(4) = 0 Then
            sError = sError & "Length given but missing circumference " & sCircum(4 * 2) & "!" & NL
        End If

    End If
    
    If sOption <> "Ignore Webs" Then
        sWebError = ""
        For ii = 5 To 8
            If Val(txtLen(ii).Text) = 0 Then
                sWebError = sWebError & "Missing length for " & sLengths(ii) & "!" & NL
            End If
        Next ii
        If sWebError <> "" Then
            If MsgBox(sWebError & NL & "If the missing webs represent amputated fingers then, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
                FN_ValidateData = False
                Exit Function
            End If
        End If

       'Check that the thumb web is lower that all the other webs
        sWebError = ""
        nThumbWebHt = Val(txtLen(8).Text)
        For ii = 5 To 7
            If ((Val(txtLen(ii).Text)) <= nThumbWebHt) And nThumbWebHt <> 0 Then
                sWebError = sWebError & "Thumb web is higher than " & sLengths(ii) & "!" & NL
            End If
        Next ii
        If sWebError <> "" Then
            If MsgBox(sWebError & NL & "Use, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
                FN_ValidateData = False
                Exit Function
            End If
        End If
        
    End If

    If cboFabric.Text = "" Then
            sError = sError & "Fabric not given! " & NL
    Else
        If UCase$(Mid$(cboFabric.Text, 1, 3)) <> "POW" Then
            sError = sError & "Fabric chosen is not Powernet"
        End If
        If Val(Mid$(cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid$(cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
            sError = sError & "Modulus of fabric chosen is not available on Gram / Tension reduction chart"
        End If

    End If

   'Validate extension data
    If sOption = "Check All" Then
        sError = sError & FN_ValidateExtensionData()
    End If

    If sError <> "" Then
        MsgBox sError, 16, g_sDialogID
        FN_ValidateData = False
    Else
        FN_ValidateData = True
    End If

End Function


Private Sub Form_LinkClose()
   
   Dim ii%, nn%, nValue#, nAge%, jj%
   Dim sTapesBeyondWrist$, sValue$

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

  'Tip Option Buttons
  '     0 => Closed tip
  '     1 => Standard Open Tip
  '     2 => Disired Open Tip
   If txtTapeLengths2.Text <> "" Then
     For ii = 0 To 6
        nValue = Val(Mid(txtTapeLengths2.Text, ii + 1, 1))
        If ii = 0 Then
            optLittleTip(nValue).Value = True
        ElseIf ii = 1 Then
            optRingTip(nValue).Value = True
        ElseIf ii = 2 Then
            optMiddleTip(nValue).Value = True
        ElseIf ii = 3 Then
            optIndexTip(nValue).Value = True
        ElseIf ii = 4 Then
            optThumbTip(nValue).Value = True
        ElseIf ii = 5 Then
            chkSlantedInserts.Value = nValue
        ElseIf ii = 6 Then
            chkPalm.Value = nValue
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
   End If


  'Setup for specific insert size
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
   
  'This is a flag used to allow for automatic recalculation
  'if the user switches between on and off the fold
  'This stops confusing the user with warning messages about
  'missing data when they have not requested anything to happen
  'but have simply used the Fold option buttons
   g_DataIsCalcuable = False
   
   sInsert = Mid$(txtTapeLengths2.Text, 15, 2)
   If sInsert <> "" Then
        nInsert = Val(sInsert)
        If nInsert <= cboInsert.ListCount Then
            cboInsert.ListIndex = nInsert
        Else
            cboInsert.ListIndex = 0
        End If
   Else
        cboInsert.ListIndex = 0
   End If

   If cboInsert.ListIndex > 0 Then
        g_DataIsCalcuable = True
        g_iInsertSize = nInsert
   End If

  'Reinforced DORSAL
   nValue = Val(Mid(txtTapeLengths2.Text, 17, 1))
   chkDorsal.Value = nValue

  'Thumb Options
   nValue = Val(Mid(txtTapeLengths2.Text, 18, 1))
   optThumbStyle(nValue) = True

  'Inset Options
   nValue = Val(Mid(txtTapeLengths2.Text, 19, 1))
   optInsertStyle(nValue) = True

  'Print Fold Options
   sValue = Mid(txtTapeLengths2.Text, 20, 1)
   chkPrintFold.Value = 1
   If sValue = "0" Then chkPrintFold.Value = 0
   
  'Fused fingers options
   nValue = Val(Mid(txtTapeLengths2.Text, 21, 1))
   chkFusedFingers.Value = nValue

  'Fabric
   If txtFabric.Text <> "" Then cboFabric.Text = txtFabric.Text

  'Get DDE data for extension
  'This procedure is in the module GLVEXTEN.BAS
   g_CalculatedExtension = False
   
   PR_GetExtensionDDE_Data
  
  'Update the values in the glove extension module
  'This takes the values from the dialogue and a sets it up for subsequent
  'procedures in this module
   PR_GetDlgAboveWrist

  'Having set all the fields with raw data we NOW!
  'refine the display by executing the relevent procedure
   For ii = 0 To 2
    If optExtendTo(ii).Value = True Then
        PR_ExtendTo_Click (ii)
        Exit For
    End If
   Next ii

  'Get data for other hand to ensure that the two gloves are
  'similar
   g_iInsertOtherGlv = -1
   g_OnFoldOtherGlv = False
   If g_sSide = "Right" Then
        If Val(Mid$(txtDataGC, 1, 2)) > 0 Then
            g_iInsertOtherGlv = Val(Mid$(txtDataGC, 1, 2))
            If Val(Mid$(txtDataGC, 3, 2)) > 0 Then g_OnFoldOtherGlv = True
        End If
   Else
        If Val(Mid$(txtDataGC, 5, 2)) <> 0 Then
            g_iInsertOtherGlv = Val(Mid$(txtDataGC, 5, 2))
            If Val(Mid$(txtDataGC, 7, 2)) > 0 Then g_OnFoldOtherGlv = True
        End If
   End If

  
  'Setup to test for changes
   ii = FN_GloveDataChanged()


   MainForm!SSTab1.Tab = 0

  'Show after link is closed
   Show
   
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
    Timer1.Interval = 6000
    Timer1.Enabled = True

    Set MainForm = manglove

   'Position to center of screen
    Left = (Screen.Width - MainForm.Width) / 2
    Top = (Screen.Height - MainForm.Height) / 2

   'Get the path to the JOBST system installation directory
    g_sPathJOBST = fnPathJOBST()

   'Set up fabric list
    PR_GetComboListFromFile cboFabric, g_sPathJOBST & "\FABRIC.DAT"
   
   'Clear DDE transfer boxes
    cboFabric.Text = ""
    txtUidGlove = ""
    txtTapeLengths = ""
    txtTapeLengths2 = ""
    txtTapeLengthPt1 = ""
    txtSide = ""
    txtWorkOrder = ""
    txtUidGC = ""
    txtFabric = ""
    txtUidMPD = ""
    txtGrams = ""
    txtReduction = ""
    txtDataGlove = ""
    txtDataGC = ""
    txtFlap = ""
    txtTapeMMs = ""
    txtWristPleat = ""
    txtShoulderPleat = ""

   'Default units
   'g_nUnitsFac = 1             'Metric
    g_nUnitsFac = 10 / 25.4     'Inches

   'Set Default Tips to Closed
    optLittleTip(0).Value = 1
    optRingTip(0).Value = 1
    optMiddleTip(0).Value = 1
    optIndexTip(0).Value = 1
    optThumbTip(0).Value = 1


   'Setup display inches grid
    grdInches.ColWidth(0) = 615
    grdInches.ColAlignment(0) = 2
    For ii = 0 To 15
'        grdInches.RowHeight(ii) = 266
        grdInches.RowHeight(ii) = 270
    Next ii

   'Setup display of results grid
    For ii = 0 To 1
        grdDisplay.ColWidth(ii) = 488
        grdDisplay.ColAlignment(ii) = 2
    Next ii
    
    For ii = 0 To 15
'        grdDisplay.RowHeight(ii) = 266
        grdDisplay.RowHeight(ii) = 270
    Next ii

   'ListBoxes
   'g_sTapeText, From GLVEXTEN.BAS

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

    cboFlaps.AddItem "Style A "
    cboFlaps.AddItem "Style B "
    cboFlaps.AddItem "Style C "
    cboFlaps.AddItem "Style D "
    cboFlaps.AddItem "Style E "
    cboFlaps.AddItem "Style F "
    cboFlaps.AddItem "Raglan A"
    cboFlaps.AddItem "Raglan B"
    cboFlaps.AddItem "Raglan C"
    cboFlaps.AddItem "Raglan D"
    cboFlaps.AddItem "Raglan E"
    cboFlaps.AddItem "Raglan F"

    cboPressure.AddItem "15"
    cboPressure.AddItem "20"
    cboPressure.AddItem "25"
    cboPressure.ListIndex = 0

   'Default side
    g_sSide = "Left"

   'load
    PR_LoadPowernetChart

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

Private Sub lblPleat_DblClick(Index As Integer)
   'Procedure to disable the pleat fields
   'to enable the user to select only the relevent
   'data to be drawn
   'Saves having to delele and re-enter to get variations
   'on a glove
    Dim iDIP%, iPIP%
    If lblPleat(Index).ForeColor = &HC0C0C0 Then
        lblPleat(Index).ForeColor = &H0&
        Select Case Index
        Case 0
            txtWristPleat1.Enabled = True
        Case 1
            txtWristPleat2.Enabled = True
        Case 2
            txtShoulderPleat2.Enabled = True
        Case 3
            txtShoulderPleat1.Enabled = True
        End Select
    Else
        lblPleat(Index).ForeColor = &HC0C0C0
        Select Case Index
        Case 0
            txtWristPleat1.Enabled = False
        Case 1
            txtWristPleat2.Enabled = False
        Case 2
            txtShoulderPleat2.Enabled = False
        Case 3
            txtShoulderPleat1.Enabled = False
        End Select
    End If

End Sub

Private Sub OK_Click()
    Dim sTask$
    Dim iInsert As Integer, iError%
   'Don't allow multiple clicking
   '
    OK.Enabled = False
   
   'Get data from the Dialogue
    PR_GetDlgData
    
    If FN_ValidateData("Check All") Then
       'Display and hourglass
        Screen.MousePointer = 11
        PR_LoadReductionCharts (cboFabric.Text)

       'These functions, FN_CalculateInsert() and FN_TranslateInsert()
       'interact with the user.
       'If they returns false then either an error has occoured
       'or the user has chosen to make changes.
       'In either case the user is returned to the dialogue

        If cboInsert.Text = "Calculate" Then
            If Not FN_CalculateInsert(g_iInsertSize, "CALC") Then GoTo EXIT_OK
        Else
           'Translate cboInsert.Text to correct format
           'for use below
            If Not FN_TranslateInsert((cboInsert.ListIndex), g_iInsertSize) Then
                GoTo EXIT_OK
            Else
               'Use the given insertsize
                If Not FN_CalculateInsert(g_iInsertSize, "USE") Then GoTo EXIT_OK
            End If
        End If

       'Check with the glove common
        If Not FN_CompareWithOtherGlove((cboFabric.Text), g_OnFold, g_iInsertSize) Then
            GoTo EXIT_OK
        End If

       'Calculate glove points based on insert given above and data.
        If Not FN_CalcGlovePoints(g_iInsertSize) Then GoTo EXIT_OK

       'Calculate glove extensions (if any)
        PR_CalculateExtension xyPalm(1), xyPalm(6), g_iInsertValue(g_iInsertSize) * EIGHTH
        
        PR_UpdateDDE

        PR_UpdateDDE_Extension
        
       'N.B. Use of local JOBST Directory C:\JOBST
        PR_CreateDrawMacro "C:\JOBST\DRAW.D", g_iInsertSize
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

EXIT_OK:

    OK.Enabled = True
    Screen.MousePointer = 0
    Show
    Exit Sub
    

End Sub

Private Sub optExtendTo_Click(Index As Integer)
   'Use Procedure in module GLVEXTEN.BAS
   If MainForm.Visible Then PR_ExtendTo_Click Index
End Sub

Private Sub optFold_Click(Index As Integer)
    If g_DataIsCalcuable Then cmdCalculateInsert_Click
End Sub

Private Sub optHand_Click(Index As Integer)
    If Index = 0 Then
        g_sSide = "Left"
    Else
        g_sSide = "Right"
    End If
End Sub

Private Sub optIndexTip_Click(Index As Integer)
    PR_CutBackTip 3, Index
End Sub

Private Sub optInsertStyle_Click(Index As Integer)
    If Index = 1 Then
        optFold(0).Value = True
        optFold(1).Enabled = False
    Else
        optFold(1).Enabled = True
    End If
End Sub

Private Sub optLittleTip_Click(Index As Integer)
    PR_CutBackTip 0, Index
End Sub

Private Sub optMiddleTip_Click(Index As Integer)
    PR_CutBackTip 2, Index
End Sub

Private Sub optProximalTape_Click(Index As Integer)
   'Use Procedure in module GLVEXTEN.BAS
    PR_ProximalTape_Click Index
End Sub

Private Sub optRingTip_Click(Index As Integer)
    PR_CutBackTip 1, Index
End Sub

Private Sub optThumbTip_Click(Index As Integer)
    PR_CutBackTip 4, Index
End Sub

Private Sub PR_CreateDrawMacro(sFileName As String, iInsertSize As Integer)
   '
   ' GLOBALS
   '
   '    g_sSide
   '    g_FingerThumbBlend
   '
    Dim ii%, xyPt1 As XY, xyPt2 As XY, xyPt3 As XY, xyTextAlign As XY, xyTmp As XY
    Dim nAge%, iiStart%, nInsert As Double, TmpX#, TmpY#, nOffSet#
    Dim FingerLFSBlend As Curve, WristLFSBlend As Curve
    Dim WristThumbBlend As Curve
    Dim LFS As Curve, ThumbSide As Curve, Reinforced  As Curve
    Dim aAngle As Double, nLength As Double, nTranslate  As Double
    Dim sWorkOrder As String, sText As String
                            

   'Initialise
    fNum = FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)

   'Draw on correct layer
    PR_Setlayer "Template" & g_sSide
    
   'Get Blending profiles on thumbside
   'N.B. Modifications to xyPalmThumb(1) and xyPalmThumb(4)
    If RadialProfile.n = 1 Then iiStart = 1 Else iiStart = 2
    xyPt1.X = RadialProfile.X(iiStart)
    xyPt1.Y = RadialProfile.Y(iiStart)
   
'    PR_CalcFingerBlendingProfile xyIndex, xyPalm(4), xyMiddle, g_Finger(4), FingerThumbBlend, xyPalmThumb(1)
    PR_CalcWristBlendingProfile xyPalm(4), xyPalm(5), xyPalm(6), xyPt1, WristThumbBlend, xyPalmThumb(4)
   
   'Get Blending profiles on Little Finger Side
    If Not g_OnFold Then
        xyPt1.X = UlnarProfile.X(iiStart)
        xyPt1.Y = UlnarProfile.Y(iiStart)
        PR_CalcWristBlendingProfile xyPalm(3), xyPalm(2), xyPalm(1), xyPt1, WristLFSBlend, xyTmp
    End If
    PR_MakeXY xyPt3, xyPalm(3).X, xyPalm(4).Y
    PR_CalcFingerBlendingProfile xyLFS, xyPt3, xyLittle, g_Finger(1), FingerLFSBlend, xyTmp

   'Calculate other PalmThumb points based on point xyPalmThumb(1)
   'Done here so it can be handed below
   '
    nInsert = g_iInsertValue(iInsertSize) * EIGHTH

   'Calculate a point that will be used to align text
    PR_MakeXY xyTextAlign, xyPalm(1).X + 0.25, xyPalm(1).Y + 1.625

   'Mirror and translate for right hand side
   'To now all the data has been caculated for the LEFT
   'side
    If g_sSide = "Right" Then
        'Find farthest point from datum
         nTranslate = max(g_Finger(4).xyRadial.X, RadialProfile.X(RadialProfile.n))
         nTranslate = nTranslate - xyDatum.X
        'Fingers
         For ii = 1 To 4
            PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_Finger(ii).xyRadial
            PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_Finger(ii).xyUlnar
         Next ii
        'Curves
         PR_MirrorCurveInYaxis xyDatum.X, nTranslate, g_FingerThumbBlend
         PR_MirrorCurveInYaxis xyDatum.X, nTranslate, WristThumbBlend
         If Not g_OnFold Then
            PR_MirrorCurveInYaxis xyDatum.X, nTranslate, WristLFSBlend
         End If
         PR_MirrorCurveInYaxis xyDatum.X, nTranslate, FingerLFSBlend
         PR_MirrorCurveInYaxis xyDatum.X, nTranslate, RadialProfile
         PR_MirrorCurveInYaxis xyDatum.X, nTranslate, UlnarProfile
        'Points
         For ii = 1 To 6
            PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyPalm(ii)
         Next ii
         For ii = 1 To 5
            PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyPalmThumb(ii)
         Next ii
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyLFS
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyLittle
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyRing
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyMiddle
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyIndex
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, xyThumb

        'Top of thumb curve
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_ThumbTopCurve.xyStart
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_ThumbTopCurve.xyTangent
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_ThumbTopCurve.xyEnd
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_ThumbTopCurve.xyR1
         PR_MirrorPointInYaxis xyDatum.X, nTranslate, g_ThumbTopCurve.xyR2
        
        'Text alignment
         nLength = max((Len(txtPatientName.Text) * 0.075) + 0.125, 1.5)
         PR_MakeXY xyTextAlign, xyPalm(1).X - nLength, xyPalm(1).Y + 1.625
    End If

   
   'Draw Fingers (For missing see later)
    For ii = 1 To 4
        PR_DrawFingers g_Finger(ii)
    Next ii
'For ii = 1 To 5
'    PR_DrawMarkerNamed "detail", xyPalmThumb(ii), .1, .1, 0
'    PR_DrawText Trim$(Str$(ii)), xyPalmThumb(ii), .075
'Next ii
   'Draw Thumb
    If g_nFingerLen(5) <> 0 Then
        Select Case g_iThumbStyle
            Case 0
                PR_DrawNormalThumb nInsert
            Case 1
               'do nothing as this is an edged thumb which does not require
               'a length, drawn below
            Case 2
                PR_DrawSetInThumb nInsert
        End Select
    Else
        If g_iThumbStyle = 1 Then PR_DrawEdgedThumb nInsert
    End If

   'Draw arm extension
   'Create profile for LitteFingerSide and ThumbSide
   'Little Finger Side
    If Not g_OnFold Then
       'Join blending profiles to extension
        LFS.n = 0
        For ii = 1 To FingerLFSBlend.n
            LFS.n = LFS.n + 1
            LFS.X(LFS.n) = FingerLFSBlend.X(ii)
            LFS.Y(LFS.n) = FingerLFSBlend.Y(ii)
        Next ii
        For ii = 2 To WristLFSBlend.n - 1
            LFS.n = LFS.n + 1
            LFS.X(LFS.n) = WristLFSBlend.X(ii)
            LFS.Y(LFS.n) = WristLFSBlend.Y(ii)
        Next ii
       'Add wrist point, only if RadialProfile.n = 1
        If RadialProfile.n = 1 Then
            LFS.n = LFS.n + 1
            LFS.X(LFS.n) = WristLFSBlend.X(WristLFSBlend.n)
            LFS.Y(LFS.n) = WristLFSBlend.Y(WristLFSBlend.n)
        End If
        For ii = 2 To UlnarProfile.n
            LFS.n = LFS.n + 1
            LFS.X(LFS.n) = UlnarProfile.X(ii)
            LFS.Y(LFS.n) = UlnarProfile.Y(ii)
        Next ii
    Else
       'To reduce code in the Zipper drawing module we
       'artificially create a polyline of three points
       'We know that the macro language always creates a polyline
       'if there are only 3 points in a fitted curve.
       'This is a bug in Drafix 3.01, which we take care of here
       'however this may be fixed for a subsequent release so beware!
        
        LFS.n = 0
        For ii = 1 To FingerLFSBlend.n
            LFS.n = LFS.n + 1
            LFS.X(LFS.n) = FingerLFSBlend.X(ii)
            LFS.Y(LFS.n) = FingerLFSBlend.Y(ii)
        Next ii
        LFS.X(LFS.n) = xyLFS.X
        LFS.Y(LFS.n) = UlnarProfile.Y(UlnarProfile.n) + ((xyLFS.Y - UlnarProfile.Y(UlnarProfile.n)) / 2)
        
        LFS.n = LFS.n + 1
        LFS.X(LFS.n) = UlnarProfile.X(UlnarProfile.n)
        LFS.Y(LFS.n) = UlnarProfile.Y(UlnarProfile.n)
    End If

   'Draw
    PR_Setlayer "Template" & g_sSide
    PR_DrawFitted LFS
    PR_AddDBValueToLast "Zipper", "LFS"
    PR_AddDBValueToLast "ZipperLength", "%Length"
   'Only add this data base field if this is an editable curve
    If g_ExtendTo <> GLOVE_NORMAL Then PR_AddDBValueToLast "curvetype", "LFS"

   'Thumb side
   'Join blending profiles to extension
    ThumbSide.n = 0
    For ii = 1 To g_FingerThumbBlend.n
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = g_FingerThumbBlend.X(ii)
        ThumbSide.Y(ThumbSide.n) = g_FingerThumbBlend.Y(ii)
    Next ii
    For ii = 2 To WristThumbBlend.n - 1
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(ii)
        ThumbSide.Y(ThumbSide.n) = WristThumbBlend.Y(ii)
    Next ii
   'Add wrist point, only if RadialProfile.n = 1
    If RadialProfile.n = 1 Then
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(WristThumbBlend.n)
        ThumbSide.Y(ThumbSide.n) = WristThumbBlend.Y(WristThumbBlend.n)
    End If
    For ii = 2 To RadialProfile.n
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = RadialProfile.X(ii)
        ThumbSide.Y(ThumbSide.n) = RadialProfile.Y(ii)
    Next ii
   'Draw
    PR_DrawFitted ThumbSide
    PR_AddDBValueToLast "Zipper", "ETS"
    PR_AddDBValueToLast "ZipperLength", "%Length"
   'Only add this data base field if this is an editable curve
    If g_ExtendTo <> GLOVE_NORMAL Then PR_AddDBValueToLast "curvetype", "ETS"
    
   'Add labels and construction lines
    PR_Setlayer "Construct"


   'Draw notes for each tape
   'If on the fold draw notes at a fixed distance
    If g_OnFold Then
        nLength = Abs(RadialProfile.X(1) - UlnarProfile.X(1)) / 2
    Else
        nLength = -1
    End If
    For ii = 1 To RadialProfile.n
        PR_MakeXY xyPt1, UlnarProfile.X(ii), UlnarProfile.Y(ii)
        PR_MakeXY xyPt2, RadialProfile.X(ii), RadialProfile.Y(ii)
        If ii = RadialProfile.n - 1 And g_EOSType = ARM_FLAP Then
           'Provide a dummy EOS for the zipper programme
            PR_DrawTapeNotes ii, xyPt1, xyPt2, nLength, "EOS"
        Else
            PR_DrawTapeNotes ii, xyPt1, xyPt2, nLength, ""
        End If
        If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
           'Draw elbow line
            PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
            PR_Setlayer "Notes"
            PR_DrawLine xyPt1, xyPt2
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.Y = xyPt3.Y + 0.0625
            PR_DrawText "E", xyPt3, 0.1
        End If
    Next ii


    If g_EOSType = ARM_FLAP Then
        PR_DrawShoulderFlaps xyPt1, xyPt2
    Else
       'Draw EOS, account for elastic if last tape is at elbow
       'N.B. our elbow is fixed at tape no 9
        PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
        PR_Setlayer "Notes"
        If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
            PR_MakeXY xyPt3, RadialProfile.X(RadialProfile.n - 1), RadialProfile.Y(RadialProfile.n - 1)
            aAngle = Abs(90 - FN_CalcAngle(xyPt2, xyPt3))
            If Val(txtAge.Text) > 10 Then nOffSet = 0.75 Else nOffSet = 0.375
            nLength = nOffSet / Cos(aAngle * (PI / 180))
            PR_CalcPolar xyPt2, FN_CalcAngle(xyPt2, xyPt3), nLength, xyPt2
            If Not g_OnFold Then
                PR_MakeXY xyPt3, UlnarProfile.X(UlnarProfile.n - 1), UlnarProfile.Y(UlnarProfile.n - 1)
                PR_CalcPolar xyPt1, FN_CalcAngle(xyPt1, xyPt3), nLength, xyPt1
            Else
                xyPt1.Y = xyPt1.Y + nOffSet
            End If
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.Y = xyPt3.Y + 0.0625
            PR_DrawText "E", xyPt3, 0.1
        End If
    
       'Elastic for children
        If Val(txtAge.Text) <= 10 Then
            PR_Setlayer "Notes"
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.Y = xyPt3.Y + 0.25
            PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
            PR_DrawText "1/2\"" Elastic", xyPt3, 0.1
        End If
    
       'Draw EOS line
        PR_Setlayer "Template" & g_sSide
        PR_DrawLine xyPt1, xyPt2
        PR_AddDBValueToLast "Zipper", "EOS"
        PR_AddDBValueToLast "curvetype", "EOS"
    
    End If

   'Draw Centre line
    If Not g_OnFold Then
        PR_Setlayer "Construct"
        aAngle = FN_CalcAngle(xyPt1, xyPt2)
        nLength = FN_CalcLength(xyPt1, xyPt2)
        PR_CalcPolar xyPt1, aAngle, nLength / 2, xyPt1
        PR_MakeXY xyPt1, xyPt1.X, UlnarProfile.Y(1)
        PR_MakeXY xyPt2, xyPt1.X, UlnarProfile.Y(UlnarProfile.n)
        PR_DrawLine xyPt1, xyPt2
    End If


   'Insert Size (Text)
   'Position Insert text Size w.r.t Little Finger web
    PR_Setlayer "Notes"
    If nInsert = 0.375 Then
       'Substitute 1/2" for 3/8" inch for printing only
        sText = Trim$(fnInchestoText(0.5))
    Else
        sText = Trim$(fnInchestoText(nInsert))
    End If
    If Mid$(sText, 1, 1) = "-" Then sText = Mid$(sText, 2) 'Strip leading "-" sign
    sText = sText & "\" & QQ
    sText = "INSERT " & sText
    PR_MakeXY xyPt3, xyLittle.X, xyLittle.Y - (nInsert + 0.25)
    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
    PR_DrawText sText, xyPt3, 0.1

   'Inserts
    Select Case g_iInsertStyle
        Case 1  'One insert, Little Finger Only, Off the Fold
            PR_InsertSymbol "GloveInsertU", xyLittle, 1, 0
            PR_InsertSymbol "GloveInsertU", xyRing, 1, 0
            PR_InsertSymbol "GloveInsertU", xyMiddle, 1, 0
            If g_sSide = "Left" Then
                PR_MakeXY xyPt1, xyPalmThumb(1).X - 0.0625, xyPalmThumb(1).Y + 0.03125
            Else
                PR_MakeXY xyPt1, xyPalmThumb(1).X + 0.0625, xyPalmThumb(1).Y + 0.03125
            End If
            PR_MakeXY xyPt2, xyPt1.X, xyPalmThumb(1).Y + 0.25
            PR_DrawLine xyPt1, xyPt2
            PR_DrawMarkerNamed "Open Short Arrow", xyPt1, 0.1, 0.125, 90
            PR_DrawMarkerNamed "Open Short Arrow", xyPt2, 0.1, 0.125, 270
        Case 2  'One insert, Little and Index Finger, On or Off the Fold
            PR_InsertSymbol "GloveInsertU", xyLittle, 1, 0
            PR_InsertSymbol "GloveInsertU", xyRing, 1, 0
            PR_InsertSymbol "GloveInsertU", xyMiddle, 1, 0
        Case 3  'One insert Glove, On or Off the Fold
            PR_InsertSymbol "GloveInsertU", xyLittle, 1, 0
            PR_InsertSymbol "GloveInsertU", xyMiddle, 1, 0
    End Select

   'Add wrist line on notes
    PR_DrawLine xyPalm(1), xyPalm(6)
        PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
    PR_AddDBValueToLast "curvetype", "WRIST"
    PR_CalcMidPoint xyPalm(1), xyPalm(6), xyPt3
    xyPt3.Y = xyPt3.Y + 0.0625
    PR_DrawText "W", xyPt3, 0.1

   'print fold line if drawn on the fold and print requested
    If g_PrintFold And g_OnFold Then
        xyPt3 = xyPalm(1)
        xyPt3.Y = xyPalm(1).Y - 0.25
        If g_sSide = "Right" Then
            g_nCurrTextAngle = 90
        Else
            g_nCurrTextAngle = -90
        End If
        PR_DrawText "Fold", xyPt3, 0.1
        g_nCurrTextAngle = 0
    End If

   'Slanted inserts
    If chkSlantedInserts.Value = 1 Then
        nInsert = nInsert - 0.125
        For ii = 2 To 4
            If g_Missing(ii) <> True Then
                If ii = 2 Then
                    TmpX = xyLittle.X
                    TmpY = xyLittle.Y
                ElseIf ii = 3 Then
                    TmpX = xyRing.X
                    TmpY = xyRing.Y
                Else
                    TmpX = xyMiddle.X
                    TmpY = xyMiddle.Y
                End If
               'Vertical line
                PR_MakeXY xyPt1, TmpX, TmpY
                PR_MakeXY xyPt2, TmpX, TmpY - nInsert
                PR_DrawLine xyPt1, xyPt2
 
               'Horizontal line
                PR_MakeXY xyPt1, TmpX + EIGHTH, TmpY - nInsert
                PR_MakeXY xyPt2, TmpX - EIGHTH, TmpY - nInsert
                PR_DrawLine xyPt1, xyPt2
 
            End If
        Next ii
    End If
 

   'Reinforced Palms
    If chkPalm.Value = 1 Or chkDorsal.Value = 1 Then
        Reinforced.n = 5
        Reinforced.X(1) = xyLFS.X
        Reinforced.Y(1) = xyLFS.Y - 0.25
        Reinforced.X(2) = xyLittle.X
        Reinforced.Y(2) = xyLittle.Y - 0.25
        Reinforced.X(3) = xyRing.X
        Reinforced.Y(3) = xyRing.Y - 0.25
        Reinforced.X(4) = xyMiddle.X
        Reinforced.Y(4) = xyMiddle.Y - 0.25
        Reinforced.X(5) = xyIndex.X
        Reinforced.Y(5) = xyMiddle.Y - 0.25

       'Draw reinforcing lines using Long Dash line type
        PR_SetLineStyle "Long Dash"
        PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT

       'Establish a text insert point
        PR_CalcMidPoint xyPalm(3), xyPalm(4), xyPt3
        xyPt3.Y = xyPt3.Y - 0.75

        If chkPalm.Value = 1 Then
            PR_DrawFitted Reinforced
           'Add REINFORCED text
            PR_DrawText "REINFORCED\nPALM", xyPt3, 0.1
        End If

        If chkDorsal.Value = 1 Then
            If chkSlantedInserts.Value = 1 Then
                For ii = 1 To 5
                    Reinforced.Y(ii) = Reinforced.Y(ii) - nInsert
                Next ii
            End If
            PR_DrawFitted Reinforced
           'Add REINFORCED text
            If chkPalm.Value = 1 Then
               'Shift text so that it misses PALM Text
               If txtSide.Text = "Right" Then
                    xyPt3.Y = xyPt3.Y - 0.4
               Else
                    xyPt3.Y = xyPt3.Y + 0.4
              End If
            End If
            PR_DrawText "REINFORCED\nDORSAL", xyPt3, 0.1
        End If
        
       'Close at wrist
        PR_DrawLine xyPalm(1), xyPalm(6)
        PR_SetLineStyle "bylayer"
    End If


   'Missing Fingers
    PR_DrawMissingFingers


   'Patient Details
    PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
    xyPt3 = xyTextAlign
    If txtWorkOrder.Text = "" Then
        sWorkOrder = "-"
    Else
        sWorkOrder = txtWorkOrder.Text
    End If
    sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim$(Mid$(cboFabric.Text, 4))
    PR_DrawText sText, xyPt3, 0.1
    
   'Other patient details in black on layer construct
    PR_Setlayer "Construct"
    PR_MakeXY xyPt3, xyPt3.X, xyPt3.Y - 0.8
    sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
    PR_DrawText sText, xyPt3, 0.1


   'Markers for zippers
   'Place markers at lowest Web positions
   'For latter use with Zips, if a slant insert given then
   'supply the length of the slant insert
   'Added here capture mods to nInsert and drawn on layer construct
    
    PR_DrawMarker xyPalm(1)
    PR_AddDBValueToLast "Zipper", "PALMER"
    PR_NamedHandle "hPalmer"    'Save a HANDLE for later use

    PR_DrawMarker xyLittle
    PR_AddDBValueToLast "Zipper", "PALMER-WEB"
    If chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If
    
    PR_DrawMarker xyRing
    PR_AddDBValueToLast "Zipper", "DORSAL-WEB"
    If chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If
    
    PR_DrawMarker xyPalmThumb(3)
    PR_AddDBValueToLast "Zipper", "DORSAL"
    
    PR_DrawMarker xyMiddle
    PR_AddDBValueToLast "Zipper", "OUTSIDE-WEB"
    If chkSlantedInserts.Value = 1 Then
        PR_AddDBValueToLast "ZipperLength", Str$(nInsert)
    Else
        PR_AddDBValueToLast "ZipperLength", "0"
    End If

   'Markers for editors
    For ii = 2 To 6
        PR_DrawMarker xyPalm(ii)
        PR_AddDBValueToLast "Data", "PALM" & Trim$(Str$(ii))
    Next ii

    
   'Update symbol gloveglove, glovecommon and marker palmer with data
    PR_UpdateDBFields "Draw"

    PR_Setlayer "1"


debugclose:

   Close #fNum

End Sub

Private Sub PR_CreateSaveMacro(sDrafixFile As String)
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
   
   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient "; txtPatientName.Text; ", "; txtFileNo.Text; ", Hand - "; g_sSide
    Print #fNum, "//by Visual Basic, GLOVES - Save only"

   'Define DRAFIX variables
    Print #fNum, "HANDLE hLayer, hChan, hSym, hEnt, hOrigin, hMPD;"
    Print #fNum, "XY     xyO, xyScale;"
    Print #fNum, "STRING sFileNo, sSide, sID, sName;"
    Print #fNum, "ANGLE  aAngle;"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    
   'Display Hour Glass symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Saving"; QQ; ");"
   
   'Get Handle and Origin of "mainpatientdetails symbol"
    Print #fNum, "hMPD = UID ("; QQ; "find"; QC; Val(txtUidMPD.Text); ");"
    Print #fNum, "if (hMPD)"
    Print #fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);"
    Print #fNum, "else"
    Print #fNum, "  Exit(%cancel,"; QQ; "Can't find > mainpatientdetails < symbol, Insert Patient Data"; QQ; ");"

   'Save DB fileds
    PR_UpdateDBFields "Save"

    Close #fNum

End Sub

Private Sub PR_CutBackTip(iFinger%, iOption)
   'Function to cut down code w.r.t.
   ' optLittleTip_Click
   ' optMiddleTip_Click
   ' optRingTip_Click
   ' optIndexTip_Click
   ' optThumbTip_Click

    Dim nLen#
    Dim nTipCutBack#

    If txtLen(iFinger).Enabled = False Then Exit Sub

    nLen = FN_InchesValue(txtLen(iFinger))
    If nLen <= 0 Then Exit Sub

    If Val(txtAge.Text) <= AGE_CUTOFF Then
        nTipCutBack = CHILD_STD_OPEN_TIP
    Else
        nTipCutBack = ADULT_STD_OPEN_TIP
    End If

    If iOption = 1 Then nLen = nLen - nTipCutBack
        
    If nLen > 0 Then lblLen(iFinger).Caption = fnInchestoText(nLen)


End Sub

Private Sub PR_DisableCir(Index As Integer)
   'Read as part of labCir_DblClick
    labCir(Index).ForeColor = &HC0C0C0
    txtCir(Index).Enabled = False
    lblCir(Index).Enabled = False
End Sub

Private Static Sub PR_DrawEdgedThumb(nInsert)
   'Procedure to draw an edged thumb
    Dim aAngle As Integer
    Dim nFillet As Single
    Dim xyTmp1 As XY, xyTmp2 As XY, xyTmp3 As XY, xyKeep As XY

   'Set angle depending on side
    If g_sSide = "Right" Then aAngle = 360 Else aAngle = 180
    
    nFillet = 0.125

   'Thumb on layer Notes
    PR_Setlayer "Notes"

    PR_CalcPolar xyPalmThumb(1), aAngle, 0.25 - nFillet, xyTmp1
    PR_DrawLine xyPalmThumb(1), xyTmp1
    PR_CalcPolar xyTmp1, 270, nFillet, xyTmp2
    PR_CalcPolar xyTmp2, aAngle, nFillet, xyTmp3
    PR_DrawArc xyTmp2, xyTmp1, xyTmp3
    xyKeep = xyTmp3

    PR_CalcPolar xyPalmThumb(4), aAngle, 0.25 - nFillet, xyTmp1
    PR_DrawLine xyPalmThumb(4), xyTmp1
    PR_CalcPolar xyTmp1, 90, nFillet, xyTmp2
    PR_CalcPolar xyTmp2, aAngle, nFillet, xyTmp3
    If g_sSide = "Right" Then PR_DrawArc xyTmp2, xyTmp3, xyTmp1 Else PR_DrawArc xyTmp2, xyTmp1, xyTmp3
    PR_DrawLine xyKeep, xyTmp3

    PR_CalcMidPoint xyKeep, xyTmp3, xyTmp1
    PR_CalcPolar xyTmp1, aAngle, 0.05, xyTmp1
    g_nCurrTextAngle = aAngle - 90
    PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
    PR_DrawText "EDGED", xyTmp1, 0.1
    g_nCurrTextAngle = 0


        

End Sub

Private Static Sub PR_DrawFingers(Digit As Finger)
   'Procedure to draw a Finger based on the data given in the structure
   'Finger
   'NB
   '    NOT USED TO DRAW A THUMB
   '
   ' Fingers are drawn vertically
   ' All dimensions assummed to be figured
   '
   '
    Dim xyPt1 As XY, xyPt2 As XY, xyPt3  As XY, xyPt4 As XY
    Dim xyPt5 As XY, xyPt6 As XY, xyPt7  As XY, xyPt8 As XY
    Dim xyPt9 As XY, xyMid As XY
    Dim nX As Double, nY As Double
    Dim aTheta As Double, nA As Double, nRadius As Double
    Dim rTheta As Double
    Dim nHt As Double, nWidth As Double
    Dim nTipWidth As Double, nBulge As Double
    Dim Sign As Integer
    
   'Simple case for missing fingers
   'Don't draw anything
    If Digit.Len <= 0 Then Exit Sub
    
   'Establish if handing is required
    Sign = 1
    If Digit.xyUlnar.X > Digit.xyRadial.X Then Sign = -1

   'Case for missing DIP
    If Digit.DIP <= 0 Then
     PR_MakeXY xyPt1, Digit.xyRadial.X, Digit.WebHt + Digit.Len
     If Sign > 0 Then
        PR_CalcPolar xyPt1, 180, Digit.PIP, xyPt2
     Else
        PR_CalcPolar xyPt1, 0, Digit.PIP, xyPt2
     End If
     Select Case Digit.Tip
        Case 0, 10 'Closed tip (ON and OFF Fold)
            nHt = FN_CalcLength(xyPt1, xyPt2) / 2   'Bug fix 04/02/97
            If xyPt1.Y - nHt <= Digit.WebHt Then    '     "    "
                xyPt1.Y = Digit.WebHt               '     "    "
                xyPt2.Y = xyPt1.Y                   '     "    "
            Else                                    '     "    "
                xyPt1.Y = xyPt1.Y - nHt             '     "    "
                xyPt2.Y = xyPt1.Y                   '     "    "
            End If                                  '     "    "
            PR_StartPoly
              PR_AddVertex Digit.xyRadial, 0
              PR_AddVertex xyPt1, (Sign)
              PR_AddVertex xyPt2, 0
              PR_AddVertex Digit.xyUlnar, 0
            PR_EndPoly

        
        Case 1, 11, 2, 12 'Standard open tip & Desired length open tip (ON Fold)
           'The cutback for standard open tips has been done in
           'FN_CalculateGlove
            PR_StartPoly
              PR_AddVertex Digit.xyRadial, 0
              PR_AddVertex xyPt1, 0
              PR_AddVertex xyPt2, 0
              PR_AddVertex Digit.xyUlnar, 0
            PR_EndPoly

     End Select

    Else


     Select Case Digit.Tip
        Case 0  'Closed tip (OFF Fold)
            nHt = Digit.Len / 3
            nTipWidth = Digit.DIP * 0.8
            If nTipWidth < 0.375 Then nTipWidth = 0.375

           'PIP
            PR_MakeXY xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt

           'DIP
            xyPt2.Y = xyPt1.Y + nHt
            xyPt2.X = xyPt1.X + (Sign * (Abs(Digit.PIP - Digit.DIP) / 2))

           'At Tip
            xyPt3.Y = xyPt2.Y + nHt
            xyPt3.X = xyPt2.X + (Sign * (Abs(nTipWidth - Digit.DIP) / 2))
            
           'Find center of arc
            rTheta = Atn((Abs(nTipWidth - Digit.DIP) / 2) / nHt)

            aTheta = rTheta * (180 / PI)
            rTheta = (rTheta + (PI / 2)) / 2
            nA = (nTipWidth / 2) / Cos(rTheta)
            nRadius = nA * Sin(rTheta)

            If Sign > 0 Then
                PR_CalcPolar xyPt3, 315 - (aTheta / 2), nA, xyPt4
                PR_CalcPolar xyPt4, 180 - aTheta, nRadius, xyPt5
                PR_CalcPolar xyPt4, aTheta, nRadius, xyPt6
            Else
                PR_CalcPolar xyPt3, 225 + (aTheta / 2), nA, xyPt4
                PR_CalcPolar xyPt4, aTheta, nRadius, xyPt5
                PR_CalcPolar xyPt4, 180 - aTheta, nRadius, xyPt6
            End If

            PR_MakeXY xyPt7, xyPt3.X + (Sign * nTipWidth), xyPt3.Y
            PR_MakeXY xyPt8, xyPt2.X + (Sign * Digit.DIP), xyPt2.Y
            PR_MakeXY xyPt9, xyPt1.X + (Sign * Digit.PIP), xyPt1.Y

           'Draw Finger
            PR_StartPoly
              PR_AddVertex Digit.xyUlnar, 0
              PR_AddVertex xyPt1, 0
              PR_AddVertex xyPt2, 0
             'Calculate bulge
              nBulge = ((nRadius - Abs(xyPt5.Y - xyPt4.Y)) / (FN_CalcLength(xyPt5, xyPt6) / 2)) * (Sign * -1)
              PR_AddVertex xyPt5, nBulge
              PR_AddVertex xyPt6, 0
              PR_AddVertex xyPt8, 0
              PR_AddVertex xyPt9, 0
              PR_AddVertex Digit.xyRadial, 0
            PR_EndPoly

        Case 1, 2 'Standard open tip & Desired length open tip (OFF Fold
           'The cutback for standard open tips has been done in
           'FN_CalculateGlove
            nHt = Digit.Len / 2

           'PIP
           ' xyPt1.x = Digit.xyUlnar.x
           ' xyPt1.y = Digit.xyUlnar.y + nHt
            PR_MakeXY xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt

            xyPt4 = xyPt1
            xyPt4.X = xyPt1.X + (Sign * Digit.PIP)

           'DIP
            xyPt2.Y = xyPt1.Y + nHt
            xyPt2.X = xyPt1.X + (Sign * (Abs(Digit.PIP - Digit.DIP) / 2))

            xyPt3 = xyPt2
            xyPt3.X = xyPt2.X + (Sign * Digit.DIP)

           'Draw Finger
            PR_StartPoly
             PR_AddVertex Digit.xyUlnar, 0
             PR_AddVertex xyPt1, 0
             PR_AddVertex xyPt2, 0
             PR_AddVertex xyPt3, 0
             PR_AddVertex xyPt4, 0
             PR_AddVertex Digit.xyRadial, 0
            PR_EndPoly

        Case 10 'Closed tip (ON Fold)
            nHt = Digit.Len / 3
            nTipWidth = Digit.DIP * 0.8
            If nTipWidth < 0.375 Then nTipWidth = 0.375

           'PIP
            PR_MakeXY xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt

           'DIP
            xyPt2.Y = xyPt1.Y + nHt
            xyPt2.X = Digit.xyUlnar.X

           'At Tip
            xyPt3.Y = xyPt2.Y + nHt
            xyPt3.X = Digit.xyUlnar.X
            
           'Center of arc
            nRadius = nTipWidth / 2
            PR_MakeXY xyPt4, xyPt3.X + (Sign * nRadius), xyPt3.Y - nRadius

           'Find Ends of arc
            rTheta = Atn(Abs(nTipWidth - Digit.DIP) / nHt)
            aTheta = rTheta * (180 / PI)

            PR_MakeXY xyPt5, xyPt3.X, xyPt3.Y - nRadius
            
            If Sign > 0 Then
                PR_CalcPolar xyPt4, aTheta, nRadius, xyPt6
            Else
                PR_CalcPolar xyPt4, 180 - aTheta, nRadius, xyPt6
            End If

            PR_MakeXY xyPt7, xyPt3.X + (Sign * nTipWidth), xyPt3.Y
            PR_MakeXY xyPt8, xyPt2.X + (Sign * Digit.DIP), xyPt2.Y
            PR_MakeXY xyPt9, xyPt1.X + (Sign * Digit.PIP), xyPt1.Y

           'Draw finger
            PR_StartPoly
              PR_AddVertex Digit.xyUlnar, 0
              PR_AddVertex xyPt5, 0
             'Calculate bulge
              PR_CalcMidPoint xyPt5, xyPt6, xyMid
              nBulge = ((nRadius - FN_CalcLength(xyMid, xyPt4)) / (FN_CalcLength(xyPt5, xyPt6) / 2)) * (Sign * -1)
              PR_AddVertex xyPt5, nBulge
              PR_AddVertex xyPt6, 0
              PR_AddVertex xyPt8, 0
              PR_AddVertex xyPt9, 0
              PR_AddVertex Digit.xyRadial, 0
            PR_EndPoly

        
        Case 11, 12 'Standard open tip & Desired length open tip (ON Fold)
           'The cutback for standard open tips has been done in
           'FN_CalculateGlove
            nHt = Digit.Len / 2

           'PIP
            PR_MakeXY xyPt1, Digit.xyUlnar.X, Digit.WebHt + nHt

            xyPt4 = xyPt1
            xyPt4.X = xyPt1.X + (Sign * Digit.PIP)

           'DIP
            xyPt2.Y = xyPt1.Y + nHt
            xyPt2.X = xyPt1.X

            xyPt3 = xyPt2
            xyPt3.X = xyPt2.X + (Sign * Digit.DIP)

           'Draw Finger
            PR_StartPoly
             PR_AddVertex Digit.xyUlnar, 0
             PR_AddVertex xyPt2, 0
             PR_AddVertex xyPt3, 0
             PR_AddVertex xyPt4, 0
             PR_AddVertex Digit.xyRadial, 0
            PR_EndPoly

     End Select
    End If

End Sub

Private Sub PR_DrawMissing(iFingers%, iTip%, xy1 As XY, xy2 As XY, xy3 As XY, xy4 As XY)
    
    Dim xyMidPoint As XY
    PR_Setlayer "Template" & g_sSide

    Select Case iFingers
        Case 1
            PR_CalcMidPoint xy1, xy2, xyMidPoint
            PR_StartPoly
                PR_AddVertex xy1, 0
                PR_AddVertex xyMidPoint, 0
                PR_AddVertex xy2, 0
            PR_EndPoly
        Case 2
            PR_StartPoly
                PR_AddVertex xy1, 0
                PR_AddVertex xy2, 0
                PR_AddVertex xy3, 0
            PR_EndPoly
            xyMidPoint = xy2
        Case 3
            PR_StartPoly
                PR_AddVertex xy1, 0
                PR_AddVertex xy2, 0
                PR_AddVertex xy3, 0
                PR_AddVertex xy4, 0
            PR_EndPoly
            PR_CalcMidPoint xy2, xy3, xyMidPoint
        Case Is < 1
           'Add note only
            PR_CalcMidPoint xy1, xy2, xyMidPoint
            xyMidPoint.Y = xyMidPoint.Y - (EIGHTH * 1)
        Case Else
            Exit Sub
    End Select

    PR_Setlayer "Notes"
    xyMidPoint.Y = xyMidPoint.Y - (EIGHTH * 2)
    If iTip = 0 Or iTip = 10 Then
        PR_DrawText "CLOSED", xyMidPoint, 0.1
    Else
        PR_DrawText "OPEN", xyMidPoint, 0.1
    End If

End Sub

Private Sub PR_DrawMissingFingers()
    Dim iMissingFingers%, ClosedLittle%, ClosedIndex%
    Dim ii%
    Dim xyNull As XY

    If g_FusedFingers Then Exit Sub

    PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT

   'Check for missing fingers
    iMissingFingers = 0
    For ii = 0 To 3
        If g_Finger(ii + 1).Len <= 0 Then iMissingFingers = iMissingFingers + (2 ^ ii)
    Next ii
            
    ClosedLittle = False
    ClosedIndex = False
    If g_Finger(1).Tip = 0 Or g_Finger(1).Tip = 10 Then ClosedLittle = True
    If g_Finger(4).Tip = 0 Or g_Finger(4).Tip = 10 Then ClosedIndex = True

    Select Case iMissingFingers
       'NB -1 => text only
        Case 1
            If ClosedLittle Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            End If
        Case 2
            PR_DrawMissing 1, g_Finger(2).Tip, xyLittle, xyRing, xyNull, xyNull
        Case 3
            If ClosedLittle Then
                PR_DrawMissing 1, g_Finger(1).Tip, xyLittle, xyRing, xyNull, xyNull
            Else
                PR_DrawMissing 2, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyNull
            End If
        Case 4
            PR_DrawMissing 1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull
        Case 5
            If ClosedLittle Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            End If
            PR_DrawMissing 1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull
        Case 6
            PR_DrawMissing 2, g_Finger(2).Tip, xyLittle, xyRing, xyMiddle, xyNull
        Case 7
            If ClosedLittle Then
                PR_DrawMissing 2, g_Finger(1).Tip, xyLittle, xyRing, xyMiddle, xyNull
            Else
                PR_DrawMissing 3, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyMiddle
            End If
        Case 8
            If ClosedIndex Then
                PR_DrawMissing -1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            End If
        Case 9
            If ClosedLittle Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            End If
            If ClosedIndex Then
                PR_DrawMissing -1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            End If
        Case 10
            PR_DrawMissing 1, g_Finger(2).Tip, xyLittle, xyRing, xyNull, xyNull
            If ClosedIndex Then
                PR_DrawMissing -1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            End If
        Case 11
            If ClosedLittle Then
                PR_DrawMissing 1, g_Finger(1).Tip, xyLittle, xyRing, xyNull, xyNull
            Else
                PR_DrawMissing 2, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyNull
            End If
            If ClosedIndex Then
                PR_DrawMissing -1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(4).Tip, xyMiddle, xyIndex, xyNull, xyNull
            End If
        Case 12
            If ClosedIndex Then
                PR_DrawMissing 1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull
            Else
                PR_DrawMissing 2, g_Finger(3).Tip, xyRing, xyMiddle, xyIndex, xyNull
            End If
        Case 13
            If ClosedLittle Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            Else
                PR_DrawMissing 1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
            End If
            If ClosedIndex Then
                PR_DrawMissing 1, g_Finger(3).Tip, xyRing, xyMiddle, xyNull, xyNull
            Else
                PR_DrawMissing 2, g_Finger(3).Tip, xyRing, xyMiddle, xyIndex, xyNull
            End If
        Case 14
            If ClosedIndex Then
                PR_DrawMissing 2, g_Finger(2).Tip, xyLittle, xyRing, xyMiddle, xyNull
            Else
                PR_DrawMissing 3, g_Finger(3).Tip, xyLittle, xyRing, xyMiddle, xyIndex
            End If
        Case 15
            If ClosedIndex And ClosedLittle Then
                PR_DrawMissing 1, g_Finger(1).Tip, xyLittle, xyMiddle, xyNull, xyNull
            ElseIf ClosedLittle Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyLFS, xyLittle, xyNull, xyNull
                PR_DrawMissing 3, g_Finger(1).Tip, xyLittle, xyRing, xyMiddle, xyIndex
            ElseIf ClosedIndex Then
                PR_DrawMissing -1, g_Finger(1).Tip, xyMiddle, xyIndex, xyNull, xyNull
                PR_DrawMissing 3, g_Finger(1).Tip, xyLFS, xyLittle, xyRing, xyMiddle
            Else
                PR_DrawMissing 1, g_Finger(1).Tip, xyLFS, xyIndex, xyNull, xyNull
            End If
        Case Else
            '=> No missing fingers so "DO NOTHING"
    End Select

End Sub

Private Static Sub PR_DrawNormalThumb(nInsert As Double)
    ReDim xyT(1 To 12) As XY
    Dim xyTmp As XY, xyTmpEnd As XY, xyTmpStart As XY, xyTranslate As XY
    Dim xyThumbPalm5 As XY
    Dim nLength#, aAngle#, aThumbAngle#
    Dim nSeam#, nJ_Radius#, nInc#
    Dim nSlantInsert#, nThumbSlantLength#
    Dim aPalmThumbStart#, aPalmThumbEnd#
    Dim ThumbTopCurve As BiArc, ThumbBottCurve As BiArc
    Dim ii%, iAngleDirection%, iThumbDirection%, Success%
    Dim nMinDist#, nOffSet#, nFiguredPalm#, nBulge#
    Dim sText$
   
   'Globals
   '     g_ThumbTopCurve, from FN_CalculateGlove
   '

   'Angles
    aPalmThumbStart = FN_CalcAngle(xyPalmThumb(4), xyPalmThumb(3)) + 20
    aPalmThumbEnd = 0
    iAngleDirection = 1
    iThumbDirection = -1
    nOffSet = 3
   
   'Mirror Thumb angles for right side
    If g_sSide = "Right" Then
        'Angles
         aPalmThumbStart = aPalmThumbStart - 40
         aPalmThumbEnd = 180
         iAngleDirection = -1
         iThumbDirection = 1
         nOffSet = 6
    End If

    nInc = 0.03125 '1/32nd of an inch

   'Thumb curve on palm
    PR_Setlayer "Notes"

   'BOTTOM Curve
   'We find by iteration the first curve that fits the data.  We change the
   'start angle as part of the iteration
    Success = False
    While Not Success And (aPalmThumbStart < 180 Or aPalmThumbStart > 0)
        ii = FN_BiArcCurve(xyPalmThumb(4), aPalmThumbStart, xyPalmThumb(3), 90, ThumbBottCurve)
        If ii Then
            Success = True
        Else
           'Adjust and try again
            aPalmThumbStart = aPalmThumbStart + (iAngleDirection * 1)
        End If
    Wend

    If Not Success Then
        MsgBox "Can't form BOTTOM part of thumb curve on palm! Adjust the palm curve and use the Transfer tool to update the thumb", 48, g_sDialogID
       'Degenerate "ThumbBottCurve" to a straight line
        ThumbBottCurve.xyStart = xyPalmThumb(4)
        ThumbBottCurve.xyEnd = xyPalmThumb(3)
        ThumbBottCurve.nR1 = 0
        ThumbBottCurve.nR2 = 0
    End If

   'Draw thumb curve on palm
    PR_DrawThumbCurve ThumbBottCurve, g_ThumbTopCurve
    PR_AddDBValueToLast "curvetype", "ThumbPalmCurve"

'    PR_DrawLine xyPalmThumb(5), xyPalmThumb(1)
    PR_DrawLine g_ThumbTopCurve.xyEnd, xyPalmThumb(1)
         
   'Draw the actual thumb
   'Translate thumb curves
    nSeam = 0.125
    nJ_Radius = 0.125

    ThumbBottCurve.xyR1.X = ThumbBottCurve.xyR1.X + nOffSet
    ThumbBottCurve.xyR2.X = ThumbBottCurve.xyR2.X + nOffSet
    ThumbBottCurve.xyTangent.X = ThumbBottCurve.xyTangent.X + nOffSet
    ThumbBottCurve.xyEnd.X = ThumbBottCurve.xyEnd.X + nOffSet
    ThumbBottCurve.xyStart.X = ThumbBottCurve.xyStart.X + nOffSet
    g_ThumbTopCurve.xyR1.X = g_ThumbTopCurve.xyR1.X + nOffSet
    g_ThumbTopCurve.xyR2.X = g_ThumbTopCurve.xyR2.X + nOffSet
    g_ThumbTopCurve.xyTangent.X = g_ThumbTopCurve.xyTangent.X + nOffSet
    g_ThumbTopCurve.xyEnd.X = g_ThumbTopCurve.xyEnd.X + nOffSet
    g_ThumbTopCurve.xyStart.X = g_ThumbTopCurve.xyStart.X + nOffSet
   
   'Draw thumb curve (Seam line)
    PR_DrawThumbCurve ThumbBottCurve, g_ThumbTopCurve
    PR_AddDBValueToLast "curvetype", "ThumbPieceSeam"

    aAngle = FN_CalcAngle(xyPalmThumb(4), xyPalmThumb(2))
    aThumbAngle = aAngle + (iThumbDirection * 45)
    xyT(1) = ThumbBottCurve.xyStart
    
    nLength = g_Finger(5).PIP / (1 / Sqr(2))
    PR_CalcPolar xyT(1), aAngle, nLength, xyT(4)
    PR_CalcPolar xyT(4), aThumbAngle, g_Finger(5).Len, xyT(3)
    PR_CalcPolar xyT(3), aThumbAngle + (iThumbDirection * 90), g_Finger(5).PIP, xyT(2)
    
    nLength = nJ_Radius / Tan((45 / 2) * (PI / 180))
    PR_CalcPolar xyT(4), aThumbAngle, nLength, xyT(5)
    PR_CalcPolar xyT(4), aAngle, nLength, xyT(6)
    PR_CalcPolar xyT(6), aAngle + (iThumbDirection * 90), nJ_Radius, xyT(7)

    nLength = FN_CalcLength(xyPalmThumb(4), xyPalmThumb(2))
    PR_CalcPolar xyT(1), aAngle, nLength, xyT(10)
    PR_DrawLine g_ThumbTopCurve.xyEnd, xyT(10)
    
    nLength = nSeam / Abs(Sin(aAngle * (PI / 180)))
    nLength = FN_CalcLength(xyPalmThumb(4), xyPalmThumb(2)) + nLength
    PR_CalcPolar xyT(1), aAngle, nLength, xyT(8)

    PR_CalcPolar xyT(1), aThumbAngle + 180, 1, xyTmp
    ii = FN_CirLinInt(xyT(1), xyTmp, ThumbBottCurve.xyR1, ThumbBottCurve.nR1 + nSeam, xyT(9))
    
    PR_DrawLine xyT(6), xyT(1)

   'Offset line
    PR_CalcPolar xyT(6), aAngle + (iAngleDirection * 90), 0.25, xyTmp
    PR_CalcPolar xyTmp, aAngle, 10, xyT(11)
    If Not FN_CirLinInt(xyT(11), xyTmp, g_ThumbTopCurve.xyR2, g_ThumbTopCurve.nR2 + nSeam, xyT(11)) Then
        'make a best guess
        PR_CalcPolar xyT(8), aAngle + (iAngleDirection * 90), 0.25, xyT(11)
    End If
    
    PR_CalcPolar xyTmp, aAngle + 180, 10, xyT(12)
    If Not FN_CirLinInt(xyT(12), xyTmp, ThumbBottCurve.xyR1, ThumbBottCurve.nR1 + nSeam, xyT(12)) Then
        'make a best guess
        PR_CalcPolar xyT(1), aAngle + (iAngleDirection * 90), 0.25, xyT(12)
    End If

    PR_DrawLine xyT(11), xyT(12)
   
   'Draw Text
    PR_CalcMidPoint xyT(1), xyT(2), xyTmp
    
    PR_CalcPolar xyTmp, aThumbAngle + (iAngleDirection * 90), g_Finger(5).PIP * 2 / 3, xyTmp
    PR_SetTextData HORIZ_CENTER, VERT_CENTER, EIGHTH, CURRENT, CURRENT
    If g_sSide = "Left" Then g_nCurrTextAngle = aThumbAngle Else g_nCurrTextAngle = 180 + aThumbAngle
    sText = txtSide.Text & "\n" & txtPatientName.Text & "\n" & txtWorkOrder.Text
    PR_DrawText sText, xyTmp, 0.1
    g_nCurrTextAngle = 0

   'Draw Outside profile
    PR_Setlayer "Template" & g_sSide
    PR_StartPoly
      PR_AddVertex xyT(9), 0

     'Draw tip
      If g_Finger(5).Tip = 0 Then
         'Closed tip
          PR_CalcPolar xyT(2), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(2)
          PR_CalcPolar xyT(3), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(3)
          PR_CalcMidPoint xyT(2), xyT(3), xyTmp
          nBulge = 1 * (iThumbDirection * -1)
      Else
         'Open tip
          nBulge = 0
      End If

      PR_AddVertex xyT(2), nBulge
      PR_AddVertex xyT(3), 0

     'Calculate bulge
      PR_CalcMidPoint xyT(5), xyT(6), xyTmp
      nBulge = ((nJ_Radius - FN_CalcLength(xyTmp, xyT(7))) / (FN_CalcLength(xyT(5), xyT(6)) / 2)) * iThumbDirection
      PR_AddVertex xyT(5), nBulge
      
      PR_AddVertex xyT(6), 0
      PR_AddVertex xyT(8), 0
    
    PR_EndPoly

    If ThumbBottCurve.nR1 > 0 And ThumbBottCurve.nR1 > 0 Then
       'Offset curves
        ThumbBottCurve.nR1 = ThumbBottCurve.nR1 + nSeam
        ThumbBottCurve.nR2 = ThumbBottCurve.nR2 + nSeam
        aAngle = FN_CalcAngle(ThumbBottCurve.xyR1, ThumbBottCurve.xyStart)
        PR_CalcPolar ThumbBottCurve.xyR1, aAngle, ThumbBottCurve.nR1, ThumbBottCurve.xyStart
        aAngle = FN_CalcAngle(ThumbBottCurve.xyR1, ThumbBottCurve.xyTangent)
        PR_CalcPolar ThumbBottCurve.xyR1, aAngle, ThumbBottCurve.nR1, ThumbBottCurve.xyTangent
        aAngle = FN_CalcAngle(ThumbBottCurve.xyR2, ThumbBottCurve.xyEnd)
        PR_CalcPolar ThumbBottCurve.xyR2, aAngle, ThumbBottCurve.nR2, ThumbBottCurve.xyEnd
    Else
        ThumbBottCurve.xyStart = xyT(9)
        ThumbBottCurve.xyEnd.X = ThumbBottCurve.xyEnd.X + (iThumbDirection * nSeam)
    End If

    If g_ThumbTopCurve.nR1 > 0 And g_ThumbTopCurve.nR1 > 0 Then
        g_ThumbTopCurve.nR1 = g_ThumbTopCurve.nR1 + nSeam
        g_ThumbTopCurve.nR2 = g_ThumbTopCurve.nR2 + nSeam
        aAngle = FN_CalcAngle(g_ThumbTopCurve.xyR1, g_ThumbTopCurve.xyStart)
        PR_CalcPolar g_ThumbTopCurve.xyR1, aAngle, g_ThumbTopCurve.nR1, g_ThumbTopCurve.xyStart
        aAngle = FN_CalcAngle(g_ThumbTopCurve.xyR1, g_ThumbTopCurve.xyTangent)
        PR_CalcPolar g_ThumbTopCurve.xyR1, aAngle, g_ThumbTopCurve.nR1, g_ThumbTopCurve.xyTangent
        aAngle = FN_CalcAngle(g_ThumbTopCurve.xyR2, g_ThumbTopCurve.xyEnd)
        PR_CalcPolar g_ThumbTopCurve.xyR2, aAngle, g_ThumbTopCurve.nR2, g_ThumbTopCurve.xyEnd
    Else
        g_ThumbTopCurve.xyEnd = xyT(8)
        g_ThumbTopCurve.xyStart = ThumbBottCurve.xyEnd
    End If

   'Draw thumb curve
    PR_DrawThumbCurve ThumbBottCurve, g_ThumbTopCurve
    PR_AddDBValueToLast "curvetype", "ThumbPieceOutsideEdge"
    
    PR_DrawLine g_ThumbTopCurve.xyEnd, xyT(8)


End Sub

Private Sub PR_DrawSetInThumb(nInsert)

   'Procedure to draw a Radial Set In thumb
    
    ReDim xyT(1 To 12) As XY
    Dim xyTmp As XY, xyTmpEnd As XY, xyTmpStart As XY, xyTranslate As XY
    Dim nLength#, aAngle#, aThumbAngle#
    Dim nSeam#, nJ_Radius#
    Dim nSlantInsert#, nThumbSlantLength#
    Dim aPalmThumbStart#, aPalmThumbEnd#
    Dim ThumbTopCurve As BiArc, ThumbBottCurve As BiArc
    Dim ii%, iAngleDirection%, iThumbDirection%, Success%
    Dim nMinDist#, nOffSet#
    

   'Angles
    iAngleDirection = 1
    iThumbDirection = -1
    aThumbAngle = 45
   
   'Mirror Thumb angles for right side
    If g_sSide = "Right" Then
        'Angles
         iAngleDirection = -1
         iThumbDirection = 1
         aThumbAngle = 135
    End If

   'Calculate thumb points
   'Draw profile
    PR_Setlayer "Template" & g_sSide

    xyT(4) = xyThumb
    xyT(1) = xyThumb
    xyT(1).Y = xyThumb.Y - (Sqr(2) * g_Finger(5).PIP)

    PR_CalcPolar xyT(4), aThumbAngle, g_Finger(5).Len, xyT(3)
    PR_CalcPolar xyT(3), aThumbAngle + (iThumbDirection * 90), g_Finger(5).PIP, xyT(2)
    
    nJ_Radius = 0.125
    nLength = nJ_Radius * Tan((22.5) * (PI / 180))
    PR_CalcPolar xyT(4), aThumbAngle, nJ_Radius, xyT(5)
    PR_CalcPolar xyT(4), 90, nJ_Radius, xyT(6)
    PR_CalcPolar xyT(6), 90 + (iThumbDirection * 90), nLength, xyT(7)

    nJ_Radius = 0.25
    nLength = nJ_Radius * Cos((67.5) * (PI / 180))
    PR_CalcPolar xyT(1), aThumbAngle, nLength, xyT(8)
    PR_CalcPolar xyT(1), 270, nLength, xyT(9)
    PR_CalcPolar xyT(9), 90 + (iThumbDirection * 90), nJ_Radius, xyT(10)

   'Draw tip
    If g_Finger(5).Tip = 0 Then
       'Closed tip
        PR_CalcPolar xyT(2), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(2)
        PR_CalcPolar xyT(3), aThumbAngle + 180, g_Finger(5).PIP / 2, xyT(3)
        PR_CalcMidPoint xyT(2), xyT(3), xyTmp
        PR_DrawArc xyTmp, xyT(2), xyT(3)
    Else
       'Open tip
        PR_DrawLine xyT(2), xyT(3)
    End If

    PR_DrawLine xyT(8), xyT(2)
    PR_DrawLine xyT(3), xyT(5)
    PR_DrawArc xyT(7), xyT(6), xyT(5)
    PR_DrawArc xyT(10), xyT(8), xyT(9)


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
    If sZipperID <> "" Then
        If sZipperID <> "EOS" Then PR_AddDBValueToLast "curvetype", "EOS"
        PR_AddDBValueToLast "Zipper", sZipperID
    End If

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
   
   'MMs
    PR_AddDBValueToLast "MM", Trim$(Str$(TapeNote(iVertex).iMMs)) & "mm"
    
   'Grams
    PR_AddDBValueToLast "Grams", Trim$(Str$(TapeNote(iVertex).iGms)) & "gm"

   'Reduction
    PR_AddDBValueToLast "Reduction", Trim$(Str$(TapeNote(iVertex).iRed))
    
   'Circumference  in Inches and Eighths format
    iInt = Int(TapeNote(iVertex).nCir)
    PR_AddDBValueToLast "TapeLengths", Trim$(Str$(iInt))

    nDec = (TapeNote(iVertex).nCir - iInt)
    If nDec > 0 Then
        iInt = round((nDec / EIGHTH) * 10) / 10
        If iInt <> 0 Then PR_AddDBValueToLast "TapeLengths2", Trim$(Str$(iInt))
    End If
    
   'Check for Notes
    If TapeNote(iVertex).sNote <> "" Then
       'set text justification etc.
        PR_SetTextData RIGHT_, VERT_CENTER, EIGHTH, CURRENT, CURRENT
        PR_MakeXY xyInsert, xyDatum.X - 0.1, xyDatum.Y + 0.35
        If TapeNote(iVertex).sNote <> "PLEAT" Then PR_Setlayer "Notes"
        PR_DrawText TapeNote(iVertex).sNote, xyInsert, 0.1
    End If

End Sub

Private Sub PR_DrawThumbCurve(BottCurve As BiArc, TopCurve As BiArc)
   'Procedure to draw the specific thumb curve as a fitted polyline.
   'Created to reduce code in PR_DrawThumb.
   '
   'N.B. (Very Carefully)
   '
   '    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   '    This is not a general procedure it is specific to the
   '    thumb case defined by the Two bi-arcs.  It is order
   '    and data dependant.
   '    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   '
   '
    Dim Profile As Curve
    Dim ii%, iVertex%
    Dim xyVertex As XY
    Dim aFirst#, aEnd#, aTangent#, aStep#, aDiff#

    iVertex = 1
    Profile.X(iVertex) = BottCurve.xyStart.X
    Profile.Y(iVertex) = BottCurve.xyStart.Y
    
    If BottCurve.nR1 > 0 And BottCurve.nR2 > 0 Then

        aFirst = FN_CalcAngle(BottCurve.xyR1, BottCurve.xyStart)
        aTangent = FN_CalcAngle(BottCurve.xyR1, BottCurve.xyTangent)
        aEnd = FN_CalcAngle(BottCurve.xyR2, BottCurve.xyEnd)
        aStep = (aTangent - aFirst) / 4
        For ii = 1 To 3
            iVertex = iVertex + 1
            PR_CalcPolar BottCurve.xyR1, aFirst + (ii * aStep), BottCurve.nR1, xyVertex
            Profile.X(iVertex) = xyVertex.X
            Profile.Y(iVertex) = xyVertex.Y
        Next ii
        
        iVertex = iVertex + 1
        Profile.X(iVertex) = BottCurve.xyTangent.X
        Profile.Y(iVertex) = BottCurve.xyTangent.Y

        aDiff = aEnd - aTangent
        If Abs(aDiff) >= 180 Then aDiff = 360 - Abs(aDiff)
        aStep = aDiff / 4
        For ii = 1 To 3
            iVertex = iVertex + 1
            PR_CalcPolar BottCurve.xyR2, aTangent + (ii * aStep), BottCurve.nR2, xyVertex
            Profile.X(iVertex) = xyVertex.X
            Profile.Y(iVertex) = xyVertex.Y
        Next ii
    
    End If
    
    iVertex = iVertex + 1
    Profile.X(iVertex) = BottCurve.xyEnd.X
    Profile.Y(iVertex) = BottCurve.xyEnd.Y

    If TopCurve.nR1 > 0 And TopCurve.nR2 > 0 Then
        aFirst = FN_CalcAngle(TopCurve.xyR1, TopCurve.xyStart)
        aTangent = FN_CalcAngle(TopCurve.xyR1, TopCurve.xyTangent)
        aEnd = FN_CalcAngle(TopCurve.xyR2, TopCurve.xyEnd)
        
        If aFirst > 270 Then
            aStep = (aTangent + (360 - aFirst)) / 4
        Else
            aStep = (aTangent - aFirst) / 4
        End If
        For ii = 1 To 3
            iVertex = iVertex + 1
            PR_CalcPolar TopCurve.xyR1, aFirst + (ii * aStep), TopCurve.nR1, xyVertex
            Profile.X(iVertex) = xyVertex.X
            Profile.Y(iVertex) = xyVertex.Y
        Next ii
            
        iVertex = iVertex + 1
        Profile.X(iVertex) = TopCurve.xyTangent.X
        Profile.Y(iVertex) = TopCurve.xyTangent.Y
        
        aStep = (aEnd - aTangent) / 4
        For ii = 1 To 3
            iVertex = iVertex + 1
            PR_CalcPolar TopCurve.xyR2, aTangent + (ii * aStep), TopCurve.nR2, xyVertex
            Profile.X(iVertex) = xyVertex.X
            Profile.Y(iVertex) = xyVertex.Y
        Next ii
    
    End If

    iVertex = iVertex + 1
    Profile.X(iVertex) = TopCurve.xyEnd.X
    Profile.Y(iVertex) = TopCurve.xyEnd.Y

    Profile.n = iVertex
    PR_DrawFitted Profile

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

Private Sub PR_GetDlgData()
   'Subroutine to extract from the Dialogue the data
   'putting it into global variables
   '
   'Common Dialogue elements
    If optFold(1).Value = True Then g_OnFold = True Else g_OnFold = False
    
   'Get data from tabbed
    PR_GetDlgHandData
    PR_GetDlgAboveWrist

End Sub

Private Sub PR_GetDlgHandData()
   'Procedure to setup Module level variables based on the data in the
   'dialogue
   'Save having to do this more than once
    
    Dim ii%

        
    If MainForm!SSTab1.Tab <> 0 Then MainForm!SSTab1.Tab = 0

   'Insert values in Eighths
    g_iInsertValue(1) = 4
    g_iInsertValue(2) = 5
    g_iInsertValue(3) = 6
    g_iInsertValue(4) = 7
    g_iInsertValue(5) = 8
    g_iInsertValue(6) = 3

   'Missing Finger
    For ii = 1 To 5
        g_Missing(ii) = False
    Next ii
    
   'Given Dimensions and missing fingers
    g_MissingFingers = 0

    If txtCir(0).Enabled Then g_nFingerDIPCir(1) = fnDisplayToInches(Val(txtCir(0).Text)) 'Little finger DIP
    If txtCir(1).Enabled Then g_nFingerPIPCir(1) = fnDisplayToInches(Val(txtCir(1).Text)) 'Little finger PIP
    If g_nFingerPIPCir(1) = 0 Then
        g_MissingFingers = g_MissingFingers + 1
        g_Missing(1) = True
    End If
    
    
    If txtCir(2).Enabled Then g_nFingerDIPCir(2) = fnDisplayToInches(Val(txtCir(2).Text)) 'Ring finger DIP
    If txtCir(3).Enabled Then g_nFingerPIPCir(2) = fnDisplayToInches(Val(txtCir(3).Text)) 'Ring finger PIP
    If g_nFingerPIPCir(2) = 0 Then
        g_MissingFingers = g_MissingFingers + 1
        g_Missing(2) = True
    End If
    
    If txtCir(4).Enabled Then g_nFingerDIPCir(3) = fnDisplayToInches(Val(txtCir(4).Text)) 'Middle Finger DIP
    If txtCir(5).Enabled Then g_nFingerPIPCir(3) = fnDisplayToInches(Val(txtCir(5).Text)) 'Middle Finger PIP
    If g_nFingerPIPCir(3) = 0 Then
        g_MissingFingers = g_MissingFingers + 1
        g_Missing(3) = True
    End If
    
    If txtCir(6).Enabled Then g_nFingerDIPCir(4) = fnDisplayToInches(Val(txtCir(6).Text)) 'Index Finger DIP
    If txtCir(7).Enabled Then g_nFingerPIPCir(4) = fnDisplayToInches(Val(txtCir(7).Text)) 'Index Finger PIP
    If g_nFingerPIPCir(4) = 0 Then
        g_MissingFingers = g_MissingFingers + 1
        g_Missing(4) = True
    End If

    If txtCir(8).Enabled Then g_nFingerPIPCir(5) = fnDisplayToInches(Val(txtCir(8).Text)) 'Thumb
    If g_nFingerPIPCir(5) = 0 Then
        g_Missing(5) = True
    End If

    g_nPalm = fnDisplayToInches(Val(txtCir(9).Text))  'Palm
    g_nWrist = fnDisplayToInches(Val(txtCir(10).Text))  'Wrist

   'Finger Lengths
    For ii = 0 To 4
        If txtLen(ii).Enabled Then g_nFingerLen(ii + 1) = fnDisplayToInches(Val(txtLen(ii).Text))
    Next ii

   'Webs
    For ii = 1 To 4
        g_nWeb(ii) = fnDisplayToInches(Val(txtLen(ii + 4).Text))
    Next ii

   'Set finger tips
    For ii = 0 To 2
        If optLittleTip(ii).Value = True Then
            If g_OnFold Then
                g_Finger(1).Tip = ii + 10
            Else
                g_Finger(1).Tip = ii
            End If
        End If
        If optRingTip(ii).Value = True Then g_Finger(2).Tip = ii
        If optMiddleTip(ii).Value = True Then g_Finger(3).Tip = ii
        If optIndexTip(ii).Value = True Then g_Finger(4).Tip = ii
        If optThumbTip(ii).Value = True Then g_Finger(5).Tip = ii
    Next ii

    g_nCir(1) = FN_InchesValue(MainForm!txtCir(10))
    If MainForm!txtCir(11).Enabled Then g_nCir(2) = FN_InchesValue(MainForm!txtCir(11))
    If MainForm!txtCir(12).Enabled Then g_nCir(3) = FN_InchesValue(MainForm!txtCir(12))
    
   'Thumb styles
    For ii = 0 To 2
        If optThumbStyle(ii) = True Then g_iThumbStyle = ii
    Next ii

   'Insert styles
    For ii = 0 To 3
        If optInsertStyle(ii) = True Then g_iInsertStyle = ii
    Next ii
    
   'print fold
    g_PrintFold = False
    If MainForm!chkPrintFold.Value <> 0 Then g_PrintFold = True

   'Fused fingers
    g_FusedFingers = False
    If MainForm!chkFusedFingers.Value <> 0 Then g_FusedFingers = True

End Sub

Private Sub PR_UpdateDBFields(sType As String)
   'sType = "Save" for use with save only
   'sType = "Draw" for use with a drawing macro

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
    Else
       'Use existing symbol
        Print #fNum, "hSym = UID ("; QQ; "find"; QC; Val(txtUidGC.Text); ");"
        Print #fNum, "if (!hSym) Exit(%cancel,"; QQ; "Can't find >"; sSymbol; "< symbol to update!"; QQ; ");"
    End If
    
   'Update DB fields
    Print #fNum, "SetDBData( hSym"; CQ; "Fabric"; QCQ; txtFabric.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Data"; QCQ; txtDataGC.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"

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
    Else
       'Use existing symbol
        Print #fNum, "hSym = UID ("; QQ; "find"; QC; Val(txtUidGlove.Text); ");"
        Print #fNum, "if (!hSym) Exit(%cancel,"; QQ; "Can't find >"; sSymbol; "< symbol to update!"; QQ; ");"
    End If
    
   'Update DB fields
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths"; QCQ; txtTapeLengths.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengths2"; QCQ; txtTapeLengths2.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeLengthPt1"; QCQ; txtTapeLengthPt1.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Sleeve"; QCQ; txtSide.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Grams"; QCQ; txtGrams.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Reduction"; QCQ; txtReduction.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "TapeMMs"; QCQ; txtTapeMMs.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Data"; QCQ; txtDataGlove.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "WristPleat"; QCQ; txtWristPleat.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "ShoulderPleat"; QCQ; txtShoulderPleat.Text; QQ; ");"
    Print #fNum, "SetDBData( hSym"; CQ; "Flap"; QCQ; txtFlap.Text; QQ; ");"

    If sType = "Draw" Then
       'Update the fields associated with the PALMER marker
       'This was given an explicit handle in the procedure PR_CreateDrawMacro
       '
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengths"; QCQ; txtTapeLengths.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengths2"; QCQ; txtTapeLengths2.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeLengthPt1"; QCQ; txtTapeLengthPt1.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeMMs"; QCQ; txtTapeMMs.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "fileno"; QCQ; txtFileNo.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Sleeve"; QCQ; txtSide.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Data"; QCQ; txtDataGlove.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "age"; QCQ; txtAge.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "units"; QCQ; txtUnits.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Grams"; QCQ; txtGrams.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Reduction"; QCQ; txtReduction.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "TapeMMs"; QCQ; txtTapeMMs.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Data"; QCQ; txtDataGlove.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "WristPleat"; QCQ; txtWristPleat.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "ShoulderPleat"; QCQ; txtShoulderPleat.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Flap"; QCQ; txtFlap.Text; QQ; ");"
        Print #fNum, "SetDBData( hPalmer"; CQ; "Fabric"; QCQ; txtFabric.Text; QQ; ");"
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
    ReDim Closed(0 To 4), StdOpen(0 To 4)
    Closed(0) = optLittleTip(0).Value
    Closed(1) = optRingTip(0).Value
    Closed(2) = optMiddleTip(0).Value
    Closed(3) = optIndexTip(0).Value
    Closed(4) = optThumbTip(0).Value
    
    StdOpen(0) = optLittleTip(1).Value
    StdOpen(1) = optRingTip(1).Value
    StdOpen(2) = optMiddleTip(1).Value
    StdOpen(3) = optIndexTip(1).Value
    StdOpen(4) = optThumbTip(1).Value

    sPacked = ""
    For ii = 0 To 4
        If Closed(ii) = True Then
            sPacked = sPacked & "0"
        ElseIf StdOpen(ii) = True Then
            sPacked = sPacked & "1"
        Else
            sPacked = sPacked & "2"
       End If
    Next ii

   'CheckBoxes
    sPacked = sPacked & Trim$(Str$(chkSlantedInserts.Value))
    sPacked = sPacked & Trim$(Str$(chkPalm.Value))
    
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
   'Translate from Manual glove format to Eighths
   'NOTE:-
   '    This is rather stupid but it made life easier earlier
    If g_iInsertSize = 6 Then
        iLen = 3
    ElseIf g_iInsertSize > 0 Then
        iLen = 3 + g_iInsertSize
    End If
    sLen = String(3, " ")
    RSet sLen = Trim$(Str$(iLen))
    sPacked = sPacked & sLen

   'Reinforced dorsal
    sPacked = sPacked & Trim$(Str$(chkDorsal.Value))

   'Thumb Options
    For ii = 0 To 2
        If optThumbStyle(ii) = True Then sPacked = sPacked & Trim$(Str$(ii))
    Next ii
   'Insert Options
    For ii = 0 To 3
        If optInsertStyle(ii) = True Then sPacked = sPacked & Trim$(Str$(ii))
    Next ii

   'Print Fold
    sPacked = sPacked & Trim$(Str$(chkPrintFold.Value))

   'Fused Fingers
    sPacked = sPacked & Trim$(Str$(chkFusedFingers.Value))

   'Store to DDE text boxes
    txtTapeLengths2.Text = sPacked

   'Hand & Glove common
   'Glove common is used to carry the fabric and insert size and
   'on\off fold data, so that both gloves can be harmonized
    sLen = String(2, " ")
    RSet sLen = Trim$(Str$(g_iInsertSize))
    If MainForm!optFold(0).Value = True Then sPacked = sLen & " 0" Else sPacked = sLen & " 1"
    
    If optHand(0).Value = True Then
        txtSide.Text = "Left"
        sPacked = sPacked & Mid$(txtDataGC, 5, 4)
    Else
        txtSide.Text = "Right"
        If Mid$(txtDataGC, 1, 4) <> "" Then
            sPacked = Mid$(txtDataGC, 1, 4) & sPacked
        Else
            sPacked = "    " & sPacked
        End If
    End If
    txtDataGC.Text = sPacked

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
    Dim sString$

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtCir(Index))
   'Display value in inches if Valid
    If nLen <> -1 Then
        sString = Trim$(fnInchestoText(nLen))
        If Left$(sString, 1) = "-" Then sString = Mid$(sString, 2)
        lblCir(Index).Caption = sString
    End If

End Sub

Private Sub txtCustFlapLength_GotFocus()

    PR_Select_Text txtCustFlapLength

End Sub

Private Sub txtCustFlapLength_LostFocus()

    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtCustFlapLength)

   'Display value in inches if Valid
    If nLen <> -1 Then labCustFlapLength.Caption = fnInchestoText(nLen)


End Sub

Private Sub txtExtCir_GotFocus(Index As Integer)
    PR_Select_Text txtExtCir(Index)
End Sub

Private Sub txtExtCir_LostFocus(Index As Integer)
        
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtExtCir(Index))

   'Display value in inches if Valid
    If nLen <> -1 Then
       'Note we use a grid here that starts at 0
        PR_GrdInchesDisplay Index - 8, nLen
    End If

End Sub

Private Sub txtFrontStrapLength_GotFocus()

    PR_Select_Text txtFrontStrapLength

End Sub

Private Sub txtFrontStrapLength_LostFocus()
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtFrontStrapLength)

   'Display value in inches if Valid
    If nLen <> -1 Then labFrontStrapLength.Caption = fnInchestoText(nLen)

End Sub

Private Sub txtLen_GotFocus(Index As Integer)
    PR_Select_Text txtLen(Index)
End Sub

Private Sub txtLen_LostFocus(Index As Integer)
    Dim nLen#
    Dim nTipCutBack#
    Dim sString$

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtLen(Index))
    If nLen < 0 Then
        lblLen(Index).Caption = ""
        Exit Sub
    End If
    
    If Index <= 4 Then
       'Special case for finger lengths w.r.t. tips
        If Val(txtAge.Text) <= AGE_CUTOFF Then
            nTipCutBack = CHILD_STD_OPEN_TIP
        Else
            nTipCutBack = ADULT_STD_OPEN_TIP
        End If

        Select Case Index
            Case 0
               'Little finger
                If optLittleTip(1).Value Then nLen = nLen - nTipCutBack
            Case 1
               'Ring finger
                If optRingTip(1).Value Then nLen = nLen - nTipCutBack
            Case 2
               'Middle finger
                If optMiddleTip(1).Value Then nLen = nLen - nTipCutBack
            Case 3
               'Index finger
                If optIndexTip(1).Value Then nLen = nLen - nTipCutBack
            Case 4
               'Thumb finger
                If optThumbTip(1).Value Then nLen = nLen - nTipCutBack
        End Select
    End If
    
   'Display length text
    If nLen > 0 Then
        sString = Trim$(fnInchestoText(nLen))
        If Left$(sString, 1) = "-" Then sString = Mid$(sString, 2)
        lblLen(Index).Caption = sString
    Else
        lblLen(Index).Caption = ""
    End If

End Sub

Private Sub txtStrapLength_GotFocus()

    PR_Select_Text txtStrapLength

End Sub

Private Sub txtStrapLength_LostFocus()
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtStrapLength)

   'Display value in inches if Valid
    If nLen <> -1 Then labStrap.Caption = fnInchestoText(nLen)

End Sub

Private Sub txtWaistCir_GotFocus()

    PR_Select_Text txtWaistCir

End Sub

Private Sub txtWaistCir_LostFocus()
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtWaistCir)

   'Display value in inches if Valid
    If nLen <> -1 Then labWaistCir.Caption = fnInchestoText(nLen)


End Sub


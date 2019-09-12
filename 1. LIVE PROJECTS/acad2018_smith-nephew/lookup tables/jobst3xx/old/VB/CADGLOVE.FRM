VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form cadglove 
   Caption         =   "CAD Glove - Dialogue"
   ClientHeight    =   8760
   ClientLeft      =   1260
   ClientTop       =   360
   ClientWidth     =   9060
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
   LinkTopic       =   "cadglove"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8760
   ScaleWidth      =   9060
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   4815
      _ExtentX        =   8488
      _ExtentY        =   15261
      _Version        =   327681
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Hand Data"
      TabPicture(0)   =   "cadglove.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Above Wrist"
      TabPicture(1)   =   "cadglove.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   170
         Top             =   120
         Width           =   4485
         Begin MSFlexGridLib.MSFlexGrid grdDisplay 
            Height          =   4335
            Left            =   3165
            TabIndex        =   244
            Top             =   290
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   7641
            _Version        =   65541
            Rows            =   16
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483629
            BackColorBkg    =   -2147483638
            GridLines       =   0
            ScrollBars      =   0
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid grdInches 
            Height          =   4335
            Left            =   1845
            TabIndex        =   243
            Top             =   290
            Width           =   615
            _ExtentX        =   1080
            _ExtentY        =   7641
            _Version        =   65541
            Rows            =   16
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483648
            BackColorBkg    =   -2147483638
            GridLines       =   0
            ScrollBars      =   0
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   9
            Left            =   2640
            TabIndex        =   213
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   8
            Left            =   2640
            TabIndex        =   212
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   9
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   101
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   8
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   100
            Top             =   330
            Width           =   495
         End
         Begin VB.Frame frmCalculate 
            Caption         =   "Figure Arm"
            Height          =   825
            Left            =   300
            TabIndex        =   204
            Top             =   5220
            Width           =   4035
            Begin VB.ComboBox cboPressure 
               Height          =   315
               Left            =   780
               Style           =   2  'Dropdown List
               TabIndex        =   206
               Top             =   330
               Width           =   795
            End
            Begin VB.CommandButton cmdCalculate 
               Cancel          =   -1  'True
               Caption         =   "Calculate"
               Height          =   375
               Left            =   2220
               TabIndex        =   205
               Top             =   300
               Width           =   1275
            End
            Begin VB.Label lblArm 
               Caption         =   "Press:"
               Height          =   255
               Index           =   7
               Left            =   180
               TabIndex        =   207
               Top             =   360
               Width           =   555
            End
         End
         Begin VB.ComboBox cboDistalTape 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   203
            Top             =   4830
            Width           =   735
         End
         Begin VB.Frame frmGloveToAxilla 
            Caption         =   "Glove to Axilla, EOS Style"
            Height          =   1965
            Left            =   300
            TabIndex        =   186
            Top             =   6090
            Width           =   4035
            Begin VB.OptionButton optProximalTape 
               Caption         =   "Flap"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   189
               Top             =   360
               Width           =   795
            End
            Begin VB.OptionButton optProximalTape 
               Caption         =   "Plain"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   188
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.TextBox txtCustFlapLength 
               Height          =   285
               Left            =   2340
               TabIndex        =   200
               Top             =   1320
               Width           =   495
            End
            Begin VB.ComboBox cboFlaps 
               Height          =   315
               Left            =   1860
               TabIndex        =   187
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtStrapLength 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2340
               TabIndex        =   196
               Top             =   780
               Width           =   495
            End
            Begin VB.TextBox txtWaistCir 
               Height          =   285
               Left            =   2340
               TabIndex        =   202
               Top             =   1590
               Width           =   495
            End
            Begin VB.TextBox txtFrontStrapLength 
               Height          =   285
               Left            =   2340
               TabIndex        =   198
               Top             =   1050
               Width           =   495
            End
            Begin VB.Label lblFlap 
               Caption         =   "Requested Flap length:"
               Height          =   285
               Index           =   3
               Left            =   240
               TabIndex        =   201
               Top             =   1350
               Width           =   1995
            End
            Begin VB.Label lblFlap 
               Caption         =   "Back Strap:"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   199
               Top             =   810
               Width           =   1035
            End
            Begin VB.Label lblFlap 
               Caption         =   "Style"
               Height          =   285
               Index           =   0
               Left            =   3360
               TabIndex        =   197
               Top             =   390
               Width           =   555
            End
            Begin VB.Label labStrap 
               Height          =   195
               Left            =   2880
               TabIndex        =   195
               Top             =   810
               Width           =   675
            End
            Begin VB.Label labCustFlapLength 
               Height          =   195
               Left            =   2880
               TabIndex        =   194
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label labWaistCir 
               Height          =   195
               Left            =   2880
               TabIndex        =   193
               Top             =   1620
               Width           =   675
            End
            Begin VB.Label lblFlap 
               Caption         =   "Style ""D"" Waist Circum:"
               Height          =   285
               Index           =   4
               Left            =   180
               TabIndex        =   192
               Top             =   1620
               Width           =   2115
            End
            Begin VB.Label labFrontStrapLength 
               Height          =   195
               Left            =   2880
               TabIndex        =   191
               Top             =   1080
               Width           =   675
            End
            Begin VB.Label lblFlap 
               Caption         =   "Front Strap:"
               Height          =   285
               Index           =   2
               Left            =   1200
               TabIndex        =   190
               Top             =   1080
               Width           =   1095
            End
         End
         Begin VB.ComboBox cboProximalTape 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   185
            Top             =   4830
            Width           =   735
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   10
            Left            =   2640
            TabIndex        =   184
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   10
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   102
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat1 
            Height          =   285
            Left            =   180
            TabIndex        =   208
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat1 
            Height          =   285
            Left            =   180
            TabIndex        =   211
            Top             =   4350
            Width           =   495
         End
         Begin VB.TextBox txtWristPleat2 
            Height          =   285
            Left            =   180
            TabIndex        =   209
            Top             =   1380
            Width           =   495
         End
         Begin VB.TextBox txtShoulderPleat2 
            Height          =   285
            Left            =   180
            TabIndex        =   210
            Top             =   3660
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   11
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   103
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   13
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   105
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   14
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   106
            Top             =   1950
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   15
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   107
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   108
            Top             =   2490
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   17
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   109
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   18
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   110
            Top             =   3030
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   19
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   111
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   20
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   112
            Top             =   3570
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   21
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   113
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   22
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   114
            Top             =   4110
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   23
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   115
            Top             =   4380
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   11
            Left            =   2640
            TabIndex        =   183
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   12
            Left            =   2640
            TabIndex        =   182
            Top             =   1410
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   13
            Left            =   2640
            TabIndex        =   181
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   14
            Left            =   2640
            TabIndex        =   180
            Top             =   1950
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   15
            Left            =   2640
            TabIndex        =   179
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   2640
            TabIndex        =   178
            Top             =   2490
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   18
            Left            =   2640
            TabIndex        =   177
            Top             =   3030
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   19
            Left            =   2640
            TabIndex        =   176
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   20
            Left            =   2640
            TabIndex        =   175
            Top             =   3570
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   21
            Left            =   2640
            TabIndex        =   174
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   22
            Left            =   2640
            TabIndex        =   173
            Top             =   4110
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   23
            Left            =   2640
            TabIndex        =   172
            Top             =   4380
            Width           =   495
         End
         Begin VB.TextBox mms 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   17
            Left            =   2640
            TabIndex        =   171
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtExtCir 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   12
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   104
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   242
            Top             =   4110
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   241
            Top             =   3420
            Width           =   495
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   240
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-1½"
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   239
            Top             =   630
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "-3"
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   238
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Wrist"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   237
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   195
            Index           =   9
            Left            =   1680
            TabIndex        =   236
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "EOS"
            Height          =   195
            Index           =   10
            Left            =   2580
            TabIndex        =   235
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   195
            Index           =   8
            Left            =   3840
            TabIndex        =   234
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   10
            Left            =   840
            TabIndex        =   233
            Top             =   930
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Elbow"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   232
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPleat 
            Caption         =   "Pleat"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   231
            Top             =   420
            Width           =   495
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "1½"
            Height          =   255
            Index           =   11
            Left            =   840
            TabIndex        =   230
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   229
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "4½"
            Height          =   255
            Index           =   13
            Left            =   840
            TabIndex        =   228
            Top             =   1740
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   255
            Index           =   14
            Left            =   840
            TabIndex        =   227
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "7½"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   226
            Top             =   2310
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "9"
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   225
            Top             =   2580
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "10½"
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   224
            Top             =   2820
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "12"
            Height          =   255
            Index           =   18
            Left            =   840
            TabIndex        =   223
            Top             =   3090
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "13½"
            Height          =   255
            Index           =   19
            Left            =   840
            TabIndex        =   222
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "15"
            Height          =   255
            Index           =   20
            Left            =   840
            TabIndex        =   221
            Top             =   3630
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "16½"
            Height          =   255
            Index           =   21
            Left            =   840
            TabIndex        =   220
            Top             =   3900
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "18"
            Height          =   255
            Index           =   22
            Left            =   840
            TabIndex        =   219
            Top             =   4140
            Width           =   375
         End
         Begin VB.Label lblTape 
            Alignment       =   2  'Center
            Caption         =   "19½"
            Height          =   255
            Index           =   23
            Left            =   840
            TabIndex        =   218
            Top             =   4440
            Width           =   375
         End
         Begin VB.Label lblArm 
            Caption         =   "Tape"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   217
            Top             =   120
            Width           =   555
         End
         Begin VB.Label lblArm 
            Caption         =   " MMs"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   216
            Top             =   150
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Gram"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   215
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblArm 
            Caption         =   "Red"
            Height          =   255
            Index           =   6
            Left            =   3780
            TabIndex        =   214
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   8025
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   4605
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   62
            Top             =   4080
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   61
            Top             =   3780
            Width           =   555
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   94
            Top             =   6240
            Width           =   1095
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   97
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   96
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optThumbTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   95
               Top             =   30
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
            TabIndex        =   90
            Top             =   5940
            Width           =   1095
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   93
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   92
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optIndexTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   91
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
            TabIndex        =   86
            Top             =   5640
            Width           =   1095
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   89
               Top             =   60
               Width           =   195
            End
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   88
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optMiddleTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   87
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
            TabIndex        =   82
            Top             =   5340
            Width           =   1095
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   85
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   84
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optRingTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   83
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   3120
            ScaleHeight     =   312
            ScaleWidth      =   1092
            TabIndex        =   78
            Top             =   5040
            Width           =   1095
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   2
               Left            =   780
               TabIndex        =   81
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   1
               Left            =   420
               TabIndex        =   80
               Top             =   30
               Width           =   195
            End
            Begin VB.OptionButton optLittleTip 
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   79
               Top             =   30
               Width           =   195
            End
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   71
            Top             =   7560
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   63
            Top             =   5040
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   64
            Top             =   5340
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   65
            Top             =   5640
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   66
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   67
            Top             =   6240
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   68
            Top             =   6690
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   69
            Top             =   6960
            Width           =   555
         End
         Begin VB.TextBox txtLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   70
            Top             =   7260
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   60
            Top             =   3480
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   59
            Top             =   3180
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   58
            Top             =   2880
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   57
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   56
            Top             =   2280
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   55
            Top             =   1980
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   54
            Top             =   1680
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   53
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   52
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   51
            Top             =   780
            Width           =   555
         End
         Begin VB.TextBox txtCir 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   50
            Top             =   480
            Width           =   555
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   12
            Left            =   2400
            TabIndex        =   169
            Top             =   4110
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   168
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   167
            Top             =   3540
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   166
            Top             =   5670
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "3"" Beyond Wrist"
            Height          =   255
            Index           =   12
            Left            =   180
            TabIndex        =   165
            Top             =   4140
            Width           =   1560
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "1½"" Beyond Wrist"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   164
            Top             =   3840
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Desird"
            Height          =   195
            Index           =   2
            Left            =   3900
            TabIndex        =   163
            Top             =   4770
            Width           =   555
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
            TabIndex        =   162
            Top             =   4500
            Width           =   795
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
            TabIndex        =   161
            Top             =   4680
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
            Left            =   180
            TabIndex        =   160
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 12"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   159
            Top             =   5070
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 13"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   158
            Top             =   5355
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 14"
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   157
            Top             =   5655
            Width           =   975
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Index 15"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   156
            Top             =   5970
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 16"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   155
            Top             =   6270
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb 20"
            Height          =   255
            Index           =   31
            Left            =   840
            TabIndex        =   154
            Top             =   7605
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle 19"
            Height          =   255
            Index           =   30
            Left            =   780
            TabIndex        =   153
            Top             =   7290
            Width           =   975
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ring 18"
            Height          =   255
            Index           =   29
            Left            =   960
            TabIndex        =   152
            Top             =   6990
            Width           =   795
         End
         Begin VB.Label Lable1 
            Alignment       =   1  'Right Justify
            Caption         =   "Little 17"
            Height          =   255
            Index           =   28
            Left            =   960
            TabIndex        =   151
            Top             =   6705
            Width           =   795
         End
         Begin VB.Label labLen 
            Alignment       =   2  'Center
            Caption         =   "Tip To Web"
            Height          =   645
            Index           =   5
            Left            =   300
            TabIndex        =   150
            Top             =   5400
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Wrist To Web"
            Height          =   645
            Index           =   26
            Left            =   300
            TabIndex        =   149
            Top             =   6930
            Width           =   435
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   148
            Top             =   5070
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   147
            Top             =   5370
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   146
            Top             =   5970
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   145
            Top             =   6270
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Std."
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   144
            Top             =   4770
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Clsd."
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   143
            Top             =   4770
            Width           =   495
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   142
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   141
            Top             =   6990
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   140
            Top             =   7290
            Width           =   615
         End
         Begin VB.Label lblLen 
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   139
            Top             =   7590
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Wrist 11"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   138
            Top             =   3540
            Width           =   1020
         End
         Begin VB.Label labCir 
            Caption         =   "Index Finger"
            Height          =   405
            Index           =   23
            Left            =   300
            TabIndex        =   137
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Middle Finger"
            Height          =   405
            Index           =   22
            Left            =   300
            TabIndex        =   136
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Ring Finger"
            Height          =   405
            Index           =   21
            Left            =   300
            TabIndex        =   135
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label labCir 
            Caption         =   "Little Finger"
            Height          =   405
            Index           =   20
            Left            =   300
            TabIndex        =   134
            Top             =   540
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   133
            Top             =   3255
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Palm 10"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   132
            Top             =   3225
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   131
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "Thumb  9"
            Height          =   255
            Index           =   8
            Left            =   720
            TabIndex        =   130
            Top             =   2925
            Width           =   1020
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   129
            Top             =   2625
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  8"
            Height          =   255
            Index           =   7
            Left            =   1140
            TabIndex        =   128
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   127
            Top             =   2325
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  7"
            Height          =   255
            Index           =   6
            Left            =   1140
            TabIndex        =   126
            Top             =   2340
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   125
            Top             =   2025
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  6"
            Height          =   255
            Index           =   5
            Left            =   1140
            TabIndex        =   124
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   123
            Top             =   1725
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  5"
            Height          =   255
            Index           =   4
            Left            =   1140
            TabIndex        =   122
            Top             =   1740
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   121
            Top             =   1425
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  4"
            Height          =   255
            Index           =   3
            Left            =   1140
            TabIndex        =   120
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   119
            Top             =   1125
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  3"
            Height          =   255
            Index           =   2
            Left            =   1140
            TabIndex        =   118
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   117
            Top             =   825
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "PIP  2"
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   116
            Top             =   840
            Width           =   615
         End
         Begin VB.Label labCir 
            Alignment       =   1  'Right Justify
            Caption         =   "DIP  1"
            Height          =   255
            Index           =   0
            Left            =   1140
            TabIndex        =   99
            Top             =   540
            Width           =   615
         End
         Begin VB.Label lblCir 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   98
            Top             =   525
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   3225
      Left            =   4920
      TabIndex        =   1
      Top             =   2460
      Width           =   4035
      Begin VB.Frame Frame9 
         Caption         =   "Variations"
         Height          =   1755
         Left            =   120
         TabIndex        =   2
         Top             =   1380
         Width           =   3795
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Glove to Axilla"
            Height          =   192
            Index           =   2
            Left            =   180
            TabIndex        =   3
            Top             =   1350
            Width           =   1575
         End
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Glove to Elbow"
            Height          =   192
            Index           =   1
            Left            =   180
            TabIndex        =   4
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton optExtendTo 
            Caption         =   "Normal"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   720
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cboInsert 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label labInsert 
            Caption         =   "Insert Size:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.Label labInsert 
            Caption         =   "Inches"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   8
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.CheckBox chkDorsal 
         Caption         =   "Reinforced Dorsal"
         Height          =   225
         Left            =   1980
         TabIndex        =   9
         Top             =   1050
         Width           =   1875
      End
      Begin VB.CheckBox chkPalm 
         Caption         =   "Reinforced Palm"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Width           =   1755
      End
      Begin VB.CheckBox chkSlantedInserts 
         Caption         =   "Slant Inserts"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   750
         Width           =   1395
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "cboFabric"
         Top             =   330
         Width           =   3735
      End
   End
   Begin VB.TextBox txtFlap 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   13
      Text            =   "txtFlap"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWristPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   14
      Text            =   "txtWristPleat"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtShoulderPleat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   15
      Text            =   "txtShoulderPleat"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDataGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   16
      Text            =   "txtDataGlove"
      Top             =   5130
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtReduction 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   17
      Text            =   "txtReduction"
      Top             =   4830
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeMMs 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   18
      Text            =   "txtTapeMMs"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtGrams 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   19
      Text            =   "txtGrams"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDataGC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   20
      Text            =   "txtDataGC"
      Top             =   870
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
      TabIndex        =   21
      Top             =   7080
      Width           =   675
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   22
      Text            =   "txtUidMPD"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSide 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   23
      Text            =   "txtSide"
      Top             =   3630
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   7080
      TabIndex        =   24
      Top             =   7800
      Width           =   1395
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   435
      Left            =   5280
      TabIndex        =   25
      Top             =   7800
      Width           =   1395
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   26
      Text            =   "txtFabric"
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidGlove 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   27
      Text            =   "txtUidGlove"
      Top             =   2730
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   28
      Text            =   "txtTapeLengths2"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTapeLengths 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   29
      Text            =   "txtTapeLengths"
      Top             =   3030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidGC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   30
      Text            =   "txtUidGC"
      Top             =   300
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9120
      TabIndex        =   31
      Text            =   "txtWorkOrder"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Glove"
      Height          =   585
      Left            =   4920
      TabIndex        =   32
      Top             =   0
      Width           =   4035
      Begin VB.OptionButton optHand 
         Caption         =   "Right"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   33
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Left"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   34
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   1785
      Left            =   4920
      TabIndex        =   35
      Top             =   630
      Width           =   4035
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   36
         Top             =   660
         Width           =   3075
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   37
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   38
         Top             =   1380
         Width           =   525
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   3060
         TabIndex        =   39
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   3240
         TabIndex        =   40
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   41
         Top             =   1020
         Width           =   3075
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   180
         TabIndex        =   75
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   2580
         TabIndex        =   73
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2700
         TabIndex        =   72
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fold "
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   5880
      TabIndex        =   43
      Top             =   5070
      Width           =   2415
      Begin VB.OptionButton optFold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ON Fold"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   44
         Top             =   210
         Width           =   1035
      End
      Begin VB.OptionButton optFold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OFF Fold"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   45
         Top             =   210
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   9240
      TabIndex        =   46
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
      TabIndex        =   42
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
      TabIndex        =   47
      Top             =   120
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "cadglove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   CADGLOVE.MAK
'Purpose:   To draw the CAD glove with the addition
'           of glove to elbow / axilla
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
'15.Aug.96  GG      F.O.C. (Because I'm a Stupid bastard!)
'                   Upgrade to the Original CAD Glove
'                   program, to extend above the wrist in
'                   the same way as the Manual Glove
'                   This has cost ME! the proverbial arm
'                   if not a leg as well.
'
'04.Jun.98  GG      Bug fix in respect of insert size
'                   Warning given if insert size below 1/2"
'
'06.Dec.98  GG      Upgrade to VB5.
'                   The code for KEVIN.EXE (the original CAD glove) is now
'                   contained in a Module KEVIN
'
'
'Notes:-
'
'
'
'
   'These are now defined in the common module Public.bas
   'Windows API Functions Declarations
   'Private Declare Function GetActiveWindow Lib "User" () As Integer
   'Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
   'Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
   'Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
   'Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
   'Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
    
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
                    AppActivate sTask
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


Private Function FN_ValidateData(sOption As String) As Integer
   'This function is used only to make gross checks
   'for missing data.
   'It does not perform any sensibility checks on the
   'data
   'The sOption allows for a less rigourous check
   '    "Check All"
   '    "Ignore Webs"
   '
    Dim sError$, nL$, ii%, nn%, sWebError$
    
   'Initialise
    FN_ValidateData = False
    sError = ""
    nL = Chr$(13)

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
            sError = sError & "Missing PIP for " & sCircum(ii) & "!" & nL
        End If
        nTotalCir(nn) = Val(txtCir(ii).Text) + Val(txtCir(ii + 1).Text)
        nn = nn + 1
    Next ii
   'Settings for thumb (as not set above)
    nTotalCir(4) = Val(txtCir(8).Text)


    For ii = 8 To 10
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing circumference for " & sCircum(ii) & "!" & nL
        End If
    Next ii

    
   'Check on lengths
    For ii = 0 To 4
        If Val(txtLen(ii).Text) = 0 And nTotalCir(ii) <> 0 Then
            sError = sError & "Missing length " & sLengths(ii) & "!" & nL
        End If
        If Val(txtLen(ii).Text) <> 0 And nTotalCir(ii) = 0 Then
            sError = sError & "Length given but missing circumference " & sCircum(ii * 2) & "!" & nL
        End If
    Next ii

    If sOption <> "Ignore Webs" Then
        sWebError = ""
        For ii = 5 To 8
            If Val(txtLen(ii).Text) = 0 Then
                sWebError = sWebError & "Missing length for " & sLengths(ii) & "!" & nL
            End If
        Next ii
        If sWebError <> "" Then
            If MsgBox(sWebError & nL & "If the missing webs represent amputated fingers then, ""OK"" to continue or ""CANCEL"" to return to the dialogue to check measurments", 49, g_sDialogID) <> IDOK Then
                FN_ValidateData = False
                Exit Function
            End If
        End If
    End If

    If cboFabric.Text = "" Then
            sError = sError & "Fabric not given! " & nL
    Else
        If UCase$(Mid$(cboFabric.Text, 1, 3)) <> "POW" Then
            sError = sError & "Fabric chosen is not Powernet"
        End If
        If Val(Mid$(cboFabric.Text, 5, 3)) < LOW_MODULUS Or Val(Mid$(cboFabric.Text, 5, 3)) > HIGH_MODULUS Then
            sError = sError & "Modulus of fabric chosen is not available on Gram / Tension reduction chart"
        End If

    End If

   'Validate extension data
    If sOption = "Check All" And g_ExtendTo <> GLOVE_NORMAL Then
        sError = sError & FN_ValidateExtensionData()
    End If

    If sError <> "" Then
        MsgBox sError, 16, g_sDialogID
        FN_ValidateData = False
    Else
        FN_ValidateData = True
    End If

End Function

Private Static Sub Form_LinkClose()

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
   Dim sINSERT As String, nInsert As Integer
   
   g_iInsertSize = 0
   cboInsert.ListIndex = 0
   
   nInsert = Val(Mid$(txtTapeLengths2.Text, 15, 2))
   If nInsert > 0 And nInsert <= cboInsert.ListCount Then
        cboInsert.ListIndex = nInsert
        g_iInsertSize = nInsert
   End If
  

  'Reinforced DORSAL
   nValue = Val(Mid(txtTapeLengths2.Text, 17, 1))
   chkDorsal.Value = nValue

  'Fabric
   If txtFabric.Text <> "" Then cboFabric.Text = txtFabric.Text
  
  'Get DDE data for extension
  'This procedure is in the module GLVEXTEN.BAS
   PR_GetExtensionDDE_Data

  'Update the values in the glove extension module
  'This takes the values from the dialogue and a sets it up for subsequent
  'procedures in this module
   PR_GetDlgAboveWrist


  'Having set all the fields with raw data we NOW!
  'refine the display by executing the relevent procedure
   For ii = 0 To 2
    If MainForm!optExtendTo(ii).Value = True Then
       PR_ExtendTo_Click (ii)
        Exit For
    End If
   Next ii


  'Setup to test for changes
   ii = FN_GloveDataChanged()

   MainForm!SSTab1.Tab = 0

   g_OnFold = False
  
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

    Set MainForm = cadglove

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
        grdInches.RowHeight(ii) = 270
    Next ii

   'Setup display of results grid
    For ii = 0 To 1
        grdDisplay.ColWidth(ii) = 488
        grdDisplay.ColAlignment(ii) = 2
    Next ii
    
    For ii = 0 To 15
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

Private Sub labCir_DblClick(INDEX As Integer)
   'Procedure to disable the circumference fields
   'to enable the use to select only the relevent
   'data to be drawn
   'Saves having to delele and re-enter to get variations
   'on a glove
    Dim iDIP%, iPIP%
    If labCir(INDEX).ForeColor = &HC0C0C0 Then
        Select Case INDEX
        Case 0 To 7, 11 To 12
            PR_EnableCir INDEX
        Case 8
            PR_EnableCir INDEX
           'Disable thumb length
            labLen(4).ForeColor = &H0&
            txtLen(4).Enabled = True
            lblLen(4).Enabled = True
        Case 20 To 23
            iDIP = (INDEX - 20) * 2
            iPIP = iDIP + 1
            labCir(INDEX).ForeColor = &H0&
            PR_EnableCir iDIP
            PR_EnableCir iPIP
           'Enable finger length
            labLen(INDEX - 20).ForeColor = &H0&
            txtLen(INDEX - 20).Enabled = True
            lblLen(INDEX - 20).Enabled = True
         End Select
    Else
        Select Case INDEX
        Case 0 To 7, 12
            PR_DisableCir INDEX
        Case 8
            PR_DisableCir INDEX
           'Disable thumb length
            labLen(4).ForeColor = &HC0C0C0
            txtLen(4).Enabled = False
            lblLen(4).Enabled = False
        Case 11
            PR_DisableCir INDEX
            PR_DisableCir INDEX + 1
        Case 20 To 23
            iDIP = (INDEX - 20) * 2
            iPIP = iDIP + 1
            labCir(INDEX).ForeColor = &HC0C0C0
            PR_DisableCir iDIP
            PR_DisableCir iPIP
           'Disable finger length
            labLen(INDEX - 20).ForeColor = &HC0C0C0
            txtLen(INDEX - 20).Enabled = False
            lblLen(INDEX - 20).Enabled = False
        End Select
    End If

End Sub

Private Sub labLen_DblClick(INDEX As Integer)
    If INDEX = 5 Then Exit Sub
    If labLen(INDEX).ForeColor = &HC0C0C0 Then
        labLen(INDEX).ForeColor = &H0&
        txtLen(INDEX).Enabled = True
        lblLen(INDEX).Enabled = True
    Else
        labLen(INDEX).ForeColor = &HC0C0C0
        txtLen(INDEX).Enabled = False
        lblLen(INDEX).Enabled = False
    End If
End Sub

Private Sub lblPleat_DblClick(INDEX As Integer)
   'Procedure to disable the pleat fields
   'to enable the user to select only the relevent
   'data to be drawn
   'Saves having to delele and re-enter to get variations
   'on a glove
    Dim iDIP%, iPIP%
    If lblPleat(INDEX).ForeColor = &HC0C0C0 Then
        lblPleat(INDEX).ForeColor = &H0&
        Select Case INDEX
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
        lblPleat(INDEX).ForeColor = &HC0C0C0
        Select Case INDEX
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

    If MainForm!SSTab1.Tab <> 0 Then MainForm!SSTab1.Tab = 0
    
    If FN_ValidateData("Check All") Then
'    If True Then
       'Display and hourglass
        Screen.MousePointer = 11
        
       'Create the data for the EXTERNAL CAD Glove
       'programme
        PR_CreateCADGloveData

       'Run the CAD Glove
        PR_RunCADGlove
        
        fNum = FN_Open("C:\JOBST\DRAW.D", txtPatientName, txtFileNo, g_sSide)

       'N.B. Use of local JOBST Dirctory C:\JOBST
       'Set translations and rotations
        XTRANSLATE = -12.937
        YTRANSLATE = FN_LengthWristToEOS()
        ROTATION = -90

        PR_DrawGloveHPGL_File ("C:\JOBST\GLOVDATA.PLT")

        If g_ExtendTo <> GLOVE_NORMAL Then
           'Calculate glove extensions (if any)
            PR_CalculateExtension
            PR_DrawExtension
            PR_UpdateDDE_Extension
        End If
            
        PR_UpdateDDE
        PR_UpdateDBFields "Draw"

        Close #fNum

'
'      'Check that given or returned insert is above a minimum of 3/8"
'       If g_iInsertSize < 3 Then
'            If IDNO = MsgBox("The insert size is less than 3/8""! Do you want to continue.", 36, g_sDialogID) Then GoTo EXIT_OK
'       End If
'
' The above was replaces at the request of G.Dunne 04.Jun.98 so that the minimum
' is now 1/2".

      'Check that given or returned insert is above a minimum of 3/8"
       If g_iInsertSize <= 3 Then
            If IDNO = MsgBox("The insert size is less than 1/2""! Do you want to continue.", 36, g_sDialogID) Then GoTo EXIT_OK
       End If
        
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

Private Sub optExtendTo_Click(INDEX As Integer)
   'Use Procedure in module GLVEXTEN.BAS
   If MainForm.Visible Then PR_ExtendTo_Click INDEX
End Sub

Sub PR_CreateCADGloveData()
    Dim fGloveData As Integer
    Dim X As Integer

    Const GLOVEDATA = "C:\JOBST\GLOVDATA.DAT"

   'Create Glove data file that is input to
   'the CAD glove programme
    
    fGloveData = FreeFile
    Open GLOVEDATA For Output As #fGloveData
       'Path to executables
        Print #fGloveData, g_sPathJOBST; "\GLOVECAD"

       'Standard header
        Print #fGloveData, "VBDATA" 'Name
        Print #fGloveData, "1234"   'File Number
        Print #fGloveData, "1"      'Product types 1 only
        Print #fGloveData, "2"      'Burns
        Print #fGloveData, "N"      'One Glove only

       'Which glove Left or Right
        If MainForm!optHand(0).Value = True Then
            Print #fGloveData, "L"  'Left Hand
        Else
            Print #fGloveData, "R"  'Right Hand
        End If

        Print #fGloveData, "N"      'No Zipper
        Print #fGloveData, "N"      'No Slant Inserts
        Print #fGloveData, "N"      'No Reinforcements

       'Number of Fingers (Excluding thumb)
        Dim ii%, iFingers%
        iFingers = 0
        For ii = 0 To 3
            If Val(MainForm!txtLen(ii).Text) <> 0 Then iFingers = iFingers + 1
        Next ii
        Print #fGloveData, Trim$(Str$(iFingers))  'Number of fingers

        Print #fGloveData, "Y"      'Name is correct
        Print #fGloveData, "Y"      'File# is correct
        Print #fGloveData, "Y"      'Hand is correct
        Print #fGloveData, "1"      '1 glove only
        Print #fGloveData, "N"      'N to Lymphedema

       'Fabric range
        Dim sFabric$, nFabric%
'        sFabric = MainForm!cboFabric.List(MainForm!cboFabric.ListIndex)
        sFabric = MainForm!cboFabric.Text
        nFabric = Val(Mid(sFabric, 5, 3))

        If nFabric <= 190 Then
            Print #fGloveData, "A"  'A Range Fabric
        Else
            Print #fGloveData, "B"  'B Range Fabric
        End If

        Print #fGloveData, "E"      'English measurments

       'Circumferences and Lengths
        For ii = 0 To 10
            Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(ii).Text)))
        Next ii
        For ii = 0 To 8
            Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtLen(ii).Text)))
        Next ii

       'Tapes extending past wrist
        If Val(MainForm!txtCir(11).Text) <> 0 Then
            Print #fGloveData, "Y"      'Does extend beyond wrist
            If Val(MainForm!txtCir(12).Text) = 0 Then
                Print #fGloveData, "1"  'Only one tape past wrist
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(11).Text)))
            Else
                Print #fGloveData, "2"  'Two tapes past wrist
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(11).Text)))
                Print #fGloveData, FN_CADGloveFormat(fnDisplayToInches(Val(MainForm!txtCir(12).Text)))
            End If
        Else
            Print #fGloveData, "N"      'Doesn't extend beyond wrist
        End If
        
        Print #fGloveData, "Y"      'Everything correct
        Print #fGloveData, "VB"     'Designers Name
        
       'Tips
        Dim nClosed%, nStdOpen%
        ReDim Closed(0 To 4) As Integer
        ReDim StdOpen(0 To 4) As Integer

        nClosed = 0
        nStdOpen = 0
        nClosed = Abs(MainForm!optLittleTip(0).Value) + Abs(MainForm!optRingTip(0).Value) + Abs(MainForm!optMiddleTip(0).Value) + Abs(MainForm!optIndexTip(0).Value) + Abs(MainForm!optThumbTip(0).Value)
        nStdOpen = Abs(MainForm!optLittleTip(1).Value) + Abs(MainForm!optRingTip(1).Value) + Abs(MainForm!optMiddleTip(1).Value) + Abs(MainForm!optIndexTip(1).Value) + Abs(MainForm!optThumbTip(1).Value)
       
       'The simple case is either all open or all closed
        If nClosed = 5 Then
           'All tips closed
            Print #fGloveData, "N"  'Open tips Not required
' The following commented out in response to the fact that
' the cutback for standard tips when performed by
' KEVIN.EXE produces strange and varing results.
' Now force all lengths to desired
'
'        ElseIf nStdOpen = 5 Then
'           'All tips open
'            Print #fGloveData, "Y"  'Open tips required
'            Print #fGloveData, "Y"  'On all 4 fingers and thumb
        Else
           'Some open, some closed, some desired
           'If a std tip is required then the finger length is
           'the given - std length
            Print #fGloveData, "Y"  'Open tips required
            Print #fGloveData, "N"  'On all fingers and thumb
            Dim nLen#
            Dim nTipCutBack#
        
           'Setup
            If Val(MainForm!txtAge.Text) <= 10 Then
                nTipCutBack = CHILD_STD_OPEN_TIP
            Else
                nTipCutBack = ADULT_STD_OPEN_TIP
            End If
            
            Closed(0) = MainForm!optLittleTip(0).Value
            Closed(1) = MainForm!optRingTip(0).Value
            Closed(2) = MainForm!optMiddleTip(0).Value
            Closed(3) = MainForm!optIndexTip(0).Value
            Closed(4) = MainForm!optThumbTip(0).Value
            
            StdOpen(0) = MainForm!optLittleTip(1).Value
            StdOpen(1) = MainForm!optRingTip(1).Value
            StdOpen(2) = MainForm!optMiddleTip(1).Value
            StdOpen(3) = MainForm!optIndexTip(1).Value
            StdOpen(4) = MainForm!optThumbTip(1).Value
            
            For ii = 0 To 4
                nLen = fnDisplayToInches(Val(MainForm!txtLen(ii).Text))

                If nLen > 0 And Closed(ii) <> True Then
                    Print #fGloveData, "Y"  'Open tip required
                    If StdOpen(ii) = True Then
                       'Cut back for a standard tip
                        Print #fGloveData, "N" 'No length is wrong
                        Print #fGloveData, FN_CADGloveFormat(nLen - nTipCutBack)
                    Else
                       'Yes this is the desired length
                        Print #fGloveData, "Y"
                    End If
                Else
                    Print #fGloveData, "N"  'Open tip not required
                End If
            Next ii
        End If

        Print #fGloveData, "Y"  'Force plot

    Close #fGloveData

End Sub




Private Sub optHand_Click(INDEX As Integer)
    If INDEX = 0 Then
        g_sSide = "Left"
    Else
        g_sSide = "Right"
    End If
End Sub

Private Sub optIndexTip_Click(INDEX As Integer)
    PR_CutBackTip 3, INDEX
End Sub

Private Sub optLittleTip_Click(INDEX As Integer)
    PR_CutBackTip 0, INDEX
End Sub

Private Sub optMiddleTip_Click(INDEX As Integer)
    PR_CutBackTip 2, INDEX
End Sub

Private Sub optProximalTape_Click(INDEX As Integer)
   'Use Procedure in module GLVEXTEN.BAS
    PR_ProximalTape_Click INDEX
End Sub

Private Sub optRingTip_Click(INDEX As Integer)
    PR_CutBackTip 1, INDEX
End Sub

Private Sub optThumbTip_Click(INDEX As Integer)
    PR_CutBackTip 4, INDEX
End Sub

Private Sub PR_CreateSaveMacro(sDrafixFile As String)
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
   
   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    nL = Chr$(10) 'The new line character
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

Private Sub PR_DisableCir(INDEX As Integer)
   'Read as part of labCir_DblClick
    labCir(INDEX).ForeColor = &HC0C0C0
    txtCir(INDEX).Enabled = False
    lblCir(INDEX).Enabled = False
End Sub

Private Sub PR_DrawExtension()
   '
   ' GLOBALS
   '
   '    g_sSide
   '
   '
    Dim ii%, xyPt1 As XY, xyPt2 As XY, xyPt3 As XY, xyTextAlign As XY, xyTmp As XY
    Dim nAge%, iiStart%, nInsert As Double, TmpX#, TmpY#, nOffSet#
    Dim WristLFSBlend  As Curve, WristThumbBlend As Curve
    Dim LFS As Curve, ThumbSide As Curve, Reinforced  As Curve
    Dim aAngle As Double, nLength As Double, nTranslate  As Double
    Dim sWorkOrder As String, sText As String
    Dim xyNull As XY, xyWrist1 As XY, xyWrist2  As XY

   'Initialise
    nInsert = g_iInsertSize * EIGHTH
    xyDatum.X = 0.80409                 'I don't know why either
    xyDatum.y = YTRANSLATE              'I only wrote the %$$&^&*() thing

   'Draw on correct layer
    PR_SetLayer "Template" & g_sSide
    
   'Get Blending profiles on thumbside
    xyWrist1.X = RadialProfile.X(2)
    xyWrist1.y = RadialProfile.y(2)
    xyWrist2.X = UlnarProfile.X(2)
    xyWrist2.y = UlnarProfile.y(2)
    
    If g_sSide = "Right" Then
        xyTmp = xyWrist2
        xyWrist2 = xyWrist1
        xyWrist1 = xyTmp
        xyTmp = xyPalm(6)
        xyPalm(6) = xyPalm(1)
        xyPalm(1) = xyTmp
    End If

    PR_CalcWristBlendThumbSide xyPalm(4), xyPalm(5), xyPalm(6), xyWrist1, WristThumbBlend
   
   'Get Blending profiles on Little Finger Side
    PR_CalcWristBlendLFS xyLFAcen, aLFAsweep, xyPalm(3), xyPalm(1), xyWrist2, WristLFSBlend

   'Calculate a point that will be used to align text
    PR_MakeXY xyTextAlign, xyPalm(1).X + 0.25, xyPalm(1).y + 1.625

   
   'Mirror and translate for right hand side
   'To now all the data has been caculated for the LEFT
   'side
    If g_sSide = "Right" Then
        'Find farthest point from datum
         nTranslate = 0
         'Curves
         PR_MirrorCurveInYaxis xyPalm(7).X, nTranslate, RadialProfile
         PR_MirrorCurveInYaxis xyPalm(7).X, nTranslate, UlnarProfile
        'Points
         For ii = 1 To 6
            PR_MirrorPointInYaxis xyPalm(7).X, nTranslate, xyPalm(ii)
         Next ii
        'Text alignment
         nLength = max((Len(txtPatientName.Text) * 0.075) + 0.125, 1.5)
         PR_MakeXY xyTextAlign, xyPalm(1).X - nLength, xyPalm(1).y + 1.625
    End If


   'Draw arm extension
   'Create profile for LitteFingerSide and ThumbSide
   'Little Finger Side
   'Join blending profiles to extension
    LFS.n = 0
    For ii = 1 To WristLFSBlend.n
        LFS.n = LFS.n + 1
        LFS.X(LFS.n) = WristLFSBlend.X(ii)
        LFS.y(LFS.n) = WristLFSBlend.y(ii)
    Next ii
    For ii = 2 To UlnarProfile.n
        LFS.n = LFS.n + 1
        LFS.X(LFS.n) = UlnarProfile.X(ii)
        LFS.y(LFS.n) = UlnarProfile.y(ii)
    Next ii

   'Draw
    PR_SetLayer "Template" & g_sSide
    PR_DrawFitted LFS
    PR_AddDBValueToLast "Zipper", "LFS"
    PR_AddDBValueToLast "ZipperLength", "%Length"
    PR_AddDBValueToLast "curvetype", "LFS"

   'Thumb side
   'Join blending profiles to extension
    ThumbSide.n = 0
    For ii = 1 To WristThumbBlend.n
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = WristThumbBlend.X(ii)
        ThumbSide.y(ThumbSide.n) = WristThumbBlend.y(ii)
    Next ii
    For ii = 2 To RadialProfile.n
        ThumbSide.n = ThumbSide.n + 1
        ThumbSide.X(ThumbSide.n) = RadialProfile.X(ii)
        ThumbSide.y(ThumbSide.n) = RadialProfile.y(ii)
    Next ii
   'Draw
    PR_DrawFitted ThumbSide
    PR_AddDBValueToLast "Zipper", "ETS"
    PR_AddDBValueToLast "ZipperLength", "%Length"
    PR_AddDBValueToLast "curvetype", "ETS"
    
   'Add labels and construction lines
    PR_SetLayer "Construct"


   'Draw notes for each tape
    nLength = -1
    For ii = 1 To RadialProfile.n
        PR_MakeXY xyPt1, UlnarProfile.X(ii), UlnarProfile.y(ii)
        PR_MakeXY xyPt2, RadialProfile.X(ii), RadialProfile.y(ii)
        If ii = RadialProfile.n - 1 And g_EOSType = ARM_FLAP Then
           'Provide a dummy EOS for the zipper programme
            PR_DrawTapeNotes ii, xyPt1, xyPt2, nLength, "EOS"
        Else
            PR_DrawTapeNotes ii, xyPt1, xyPt2, nLength, ""
        End If
        If TapeNote(ii).iTapePos = ELBOW_TAPE And TapeNote(RadialProfile.n).iTapePos <> ELBOW_TAPE Then
           'Draw elbow line
            PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
            PR_SetLayer "Notes"
            PR_DrawLine xyPt1, xyPt2
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.y = xyPt3.y + 0.0625
            PR_DrawText "E", xyPt3, 0.1
        End If
    Next ii


    If g_EOSType = ARM_FLAP Then
        PR_DrawShoulderFlaps xyPt1, xyPt2
    Else
       'Draw EOS, account for elastic if last tape is at elbow
       'N.B. our elbow is fixed at tape no 9
        PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
        PR_SetLayer "Notes"
        If TapeNote(RadialProfile.n).iTapePos = ELBOW_TAPE Then
            PR_MakeXY xyPt3, RadialProfile.X(RadialProfile.n - 1), RadialProfile.y(RadialProfile.n - 1)
            aAngle = Abs(90 - FN_CalcAngle(xyPt2, xyPt3))
            If Val(txtAge.Text) > 10 Then nOffSet = 0.75 Else nOffSet = 0.375
            nLength = nOffSet / Cos(aAngle * (PI / 180))
            PR_CalcPolar xyPt2, FN_CalcAngle(xyPt2, xyPt3), nLength, xyPt2
            If Not g_OnFold Then
                PR_MakeXY xyPt3, UlnarProfile.X(UlnarProfile.n - 1), UlnarProfile.y(UlnarProfile.n - 1)
                PR_CalcPolar xyPt1, FN_CalcAngle(xyPt1, xyPt3), nLength, xyPt1
            Else
                xyPt1.y = xyPt1.y + nOffSet
            End If
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.y = xyPt3.y + 0.0625
            PR_DrawText "E", xyPt3, 0.1
        End If
    
       'Elastic for children
        If Val(txtAge.Text) <= 10 Then
            PR_SetLayer "Notes"
            PR_CalcMidPoint xyPt1, xyPt2, xyPt3
            xyPt3.y = xyPt3.y + 0.25
            PR_SetTextData HORIZ_CENTER, VERT_CENTER, CURRENT, CURRENT, CURRENT
            PR_DrawText "1/2\"" Elastic", xyPt3, 0.1
        End If
    
       'Draw EOS line
        PR_SetLayer "Template" & g_sSide
        PR_DrawLine xyPt1, xyPt2
        PR_AddDBValueToLast "Zipper", "EOS"
        PR_AddDBValueToLast "curvetype", "EOS"
    
    End If

   'Draw Centre line
    PR_SetLayer "Construct"
    aAngle = FN_CalcAngle(xyPt1, xyPt2)
    nLength = FN_CalcLength(xyPt1, xyPt2)
    PR_CalcPolar xyPt1, aAngle, nLength / 2, xyPt1
    PR_MakeXY xyPt1, xyPt1.X, UlnarProfile.y(1)
    PR_MakeXY xyPt2, xyPt1.X, UlnarProfile.y(UlnarProfile.n)
    PR_DrawLine xyPt1, xyPt2
    PR_AddDBValueToLast "curvetype", "CL"


   'Add wrist line on notes
    PR_SetLayer "Notes"
    PR_DrawLine xyPalm(1), xyPalm(6)
        PR_SetTextData HORIZ_CENTER, BOTTOM_, CURRENT, CURRENT, CURRENT
    PR_AddDBValueToLast "curvetype", "WRIST"
    PR_CalcMidPoint xyPalm(1), xyPalm(6), xyPt3
    xyPt3.y = xyPt3.y + 0.0625
    PR_DrawText "W", xyPt3, 0.1
    
   'Update symbol gloveglove, glovecommon and marker palmer with data
'    PR_UpdateDBFields "Draw"
    If g_sSide = "Right" Then
        PR_RepositionPalmerMarker xyPalm(6)
    Else
        PR_RepositionPalmerMarker xyPalm(1)
    End If
    PR_SetLayer "1"



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
    PR_SetLayer "Construct"

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
        PR_MakeXY xyInsert, xyDatum.X - 0.1, xyDatum.y + 0.35
        If TapeNote(iVertex).sNote <> "PLEAT" Then PR_SetLayer "Notes"
        PR_DrawText TapeNote(iVertex).sNote, xyInsert, 0.1
    End If

End Sub

Private Sub PR_EnableCir(INDEX As Integer)
   'Read as part of labCir_DblClick
    labCir(INDEX).ForeColor = &H0&
    txtCir(INDEX).Enabled = True
    lblCir(INDEX).Enabled = True
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
    Dim ii
   'Common Dialogue elements
    g_OnFold = False

   'Glove type
    If MainForm!optExtendTo(0).Value = True Then
        g_ExtendTo = GLOVE_NORMAL
    ElseIf MainForm!optExtendTo(1).Value = True Then
        g_ExtendTo = GLOVE_ELBOW
    Else
        g_ExtendTo = GLOVE_AXILLA
    End If

    g_nWrist = fnDisplayToInches(Val(MainForm!txtCir(10).Text))
    If g_ExtendTo <> GLOVE_NORMAL Then
        PR_GetDlgAboveWrist
    Else
        g_iNumTapesWristToEOS = 0
        For ii = 11 To 12
            If Val(MainForm!txtCir(ii).Text) > 0 And MainForm!txtCir(ii).Enabled Then g_iNumTapesWristToEOS = g_iNumTapesWristToEOS + 1
        Next ii
    End If

End Sub

Private Sub PR_RepositionPalmerMarker(xyPalm As XY)
   'As the extenstion has  changed the position of
   'the wrist we need to adjust the palmer marker
    Print #fNum, "SetGeometry(hPalmer,"; QQ; "xmarker"; QQ; CC; "xyStart.x+"; xyPalm.X; CC; "xyStart.y+"; xyPalm.y; ",0.0625);"
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

Private Sub txtCir_GotFocus(INDEX As Integer)
    PR_Select_Text txtCir(INDEX)
End Sub

Private Sub txtCir_LostFocus(INDEX As Integer)
        
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtCir(INDEX))

   'Display value in inches if Valid
    If nLen <> -1 Then lblCir(INDEX).Caption = fnInchestoText(nLen)

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

Private Sub txtExtCir_GotFocus(INDEX As Integer)
    PR_Select_Text txtExtCir(INDEX)
End Sub

Private Sub txtExtCir_LostFocus(INDEX As Integer)
        
    Dim nLen#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtExtCir(INDEX))

   'Display value in inches if Valid
    If nLen <> -1 Then
       'Note we use a grid here that starts at 0
        PR_GrdInchesDisplay INDEX - 8, nLen
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

Private Sub txtLen_GotFocus(INDEX As Integer)
    PR_Select_Text txtLen(INDEX)
End Sub

Private Sub txtLen_LostFocus(INDEX As Integer)
    Dim nLen#
    Dim nTipCutBack#

   'Check that number is valid (-ve value => invalid)
    nLen = FN_InchesValue(txtLen(INDEX))
    If nLen < 0 Then
        lblLen(INDEX).Caption = ""
        Exit Sub
    End If
    
    If INDEX <= 4 Then
       'Special case for finger lengths w.r.t. tips
        If Val(txtAge.Text) <= AGE_CUTOFF Then
            nTipCutBack = CHILD_STD_OPEN_TIP
        Else
            nTipCutBack = ADULT_STD_OPEN_TIP
        End If

        Select Case INDEX
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
        lblLen(INDEX).Caption = fnInchestoText(nLen)
    Else
        lblLen(INDEX).Caption = ""
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


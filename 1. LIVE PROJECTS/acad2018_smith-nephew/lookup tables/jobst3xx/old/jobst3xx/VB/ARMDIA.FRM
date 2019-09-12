VERSION 5.00
Begin VB.Form armdia 
   Caption         =   "ARM"
   ClientHeight    =   8328
   ClientLeft      =   708
   ClientTop       =   1968
   ClientWidth     =   9540
   ClipControls    =   0   'False
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
   LinkTopic       =   "armdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8328
   ScaleWidth      =   9540
   Begin VB.TextBox txtVestID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11580
      TabIndex        =   222
      Text            =   "txtVestID"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMeshData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10620
      TabIndex        =   221
      Text            =   "txtMeshData"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   435
      Left            =   10080
      TabIndex        =   218
      Top             =   7710
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   9660
      Top             =   7710
   End
   Begin VB.TextBox txtUidArm 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   217
      Text            =   "txtUidArm"
      Top             =   7380
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   216
      Text            =   "txtUidMPD"
      Top             =   6300
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtUidAC 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   215
      Text            =   "txtUidAC"
      Top             =   6660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   214
      Text            =   "txtWorkOrder"
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtStrap 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   207
      Text            =   "txtStrap"
      Top             =   3330
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   203
      Text            =   "txtID"
      Top             =   4050
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   150
      Text            =   "txtAge"
      Top             =   1170
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   195
      Text            =   "txtSex"
      Top             =   1530
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtTapeLent 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   194
      Text            =   "txtTapeLent"
      Top             =   450
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtTapeMM 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   193
      Text            =   "txtTapeMM"
      Top             =   810
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   192
      Text            =   "txtWeight"
      Top             =   1170
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtReduction 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   191
      Text            =   "txtReduction"
      Top             =   1530
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFirstTape 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   190
      Text            =   "txtFirstTape"
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLastTape 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   189
      Text            =   "txtLastTape"
      Top             =   2610
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSleeve 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   188
      Text            =   "txtSleeve"
      Top             =   6975
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSecondTape 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   187
      Text            =   "txtSecondTape"
      Top             =   2250
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSecondLastTape 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   186
      Text            =   "txtSecondLastTape"
      Top             =   2970
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtStump 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   185
      Text            =   "txtStump"
      Top             =   6210
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtTCircum 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   184
      Text            =   "txtTCircum"
      Top             =   4050
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtWristNo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   183
      Text            =   "txtWristNo"
      Top             =   3330
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPalmNo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   181
      Text            =   "txtPalmNo"
      Top             =   2970
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtDetach 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   179
      Text            =   "txtDetach"
      Top             =   2610
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtEnclosed 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   177
      Text            =   "txtEnclosed"
      Top             =   3690
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtGauntletFlag 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   175
      Text            =   "txtGauntletFlag"
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtThumLen 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   174
      Text            =   "txtThumLen"
      Top             =   4410
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtThumbDist 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   173
      Text            =   "txtThumbDist"
      Top             =   5490
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPalmWristDist 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   172
      Text            =   "txtPalmWristDist"
      Top             =   5130
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtNoThumb 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   171
      Text            =   "txtNoThumb"
      Top             =   4770
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFlap 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   170
      Text            =   "txtFlap"
      Top             =   2250
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   169
      Text            =   "txtFabric"
      Top             =   7740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtMM 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   9660
      TabIndex        =   168
      Text            =   "txtMM"
      Top             =   7335
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtModulus 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   167
      Text            =   "txtModulus"
      Top             =   3690
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtVestRaglan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10620
      TabIndex        =   166
      Text            =   "txtVestRaglan"
      Top             =   450
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtGauntlet 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10620
      TabIndex        =   165
      Text            =   "txtGauntlet"
      Top             =   4770
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   10620
      TabIndex        =   164
      Text            =   "txtType"
      Top             =   5130
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame6 
      Caption         =   "Distal tape Style"
      Height          =   3285
      Left            =   4560
      TabIndex        =   43
      Top             =   2220
      Width           =   4875
      Begin VB.Frame frmGauntlet 
         Caption         =   "Gauntlet Details"
         Height          =   2415
         Left            =   60
         TabIndex        =   153
         Top             =   780
         Width           =   4755
         Begin VB.TextBox txtGauntletExtension 
            Height          =   285
            Left            =   3300
            MaxLength       =   4
            TabIndex        =   159
            Top             =   2010
            Width           =   495
         End
         Begin VB.TextBox txtThumb 
            Height          =   285
            Left            =   3300
            MaxLength       =   4
            TabIndex        =   158
            Top             =   1680
            Width           =   495
         End
         Begin VB.CheckBox chkDetachable 
            Caption         =   "Detachable"
            Height          =   255
            Left            =   120
            TabIndex        =   160
            Top             =   810
            Width           =   1335
         End
         Begin VB.CheckBox chkEnclosed 
            Caption         =   "Enclosed "
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   1530
            Width           =   1215
         End
         Begin VB.TextBox txtCircum 
            Height          =   285
            Left            =   3300
            TabIndex        =   157
            Top             =   1350
            Width           =   495
         End
         Begin VB.TextBox txtPalmWrist 
            Height          =   285
            Left            =   3300
            TabIndex        =   156
            Top             =   1020
            Width           =   495
         End
         Begin VB.CheckBox chkNoThumb 
            Caption         =   "No Thumb"
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   1170
            Width           =   1335
         End
         Begin VB.ComboBox cboPalm 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3300
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   270
            Width           =   735
         End
         Begin VB.ComboBox cboWrist 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3300
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   630
            Width           =   735
         End
         Begin VB.Label labGauntletExtension 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3900
            TabIndex        =   213
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblGauntletExtension 
            Caption         =   "Distal Extension:"
            Height          =   255
            Left            =   1500
            TabIndex        =   212
            Top             =   2010
            Width           =   1815
         End
         Begin VB.Label lblPalm 
            Caption         =   "Palm at Tape No :"
            Height          =   255
            Left            =   1500
            TabIndex        =   44
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label lblWrist 
            Caption         =   "Wrist at Tape No :"
            Height          =   255
            Left            =   1500
            TabIndex        =   136
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label lblThumb 
            Caption         =   "Thumb Length :"
            Height          =   255
            Left            =   1500
            TabIndex        =   137
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblCircum 
            Caption         =   "Thumb Circum :"
            Height          =   255
            Left            =   1500
            TabIndex        =   138
            Top             =   1350
            Width           =   1575
         End
         Begin VB.Label lblPalmWrist 
            Caption         =   "Palm-Wrist Distance:"
            Height          =   255
            Left            =   1500
            TabIndex        =   139
            Top             =   1020
            Width           =   2055
         End
         Begin VB.Label PWDist 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3900
            TabIndex        =   140
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label TCirc 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3900
            TabIndex        =   141
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label TLen 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3900
            TabIndex        =   142
            Top             =   1710
            Width           =   735
         End
      End
      Begin VB.OptionButton chkGauntlets 
         Caption         =   "Gauntlet"
         Height          =   255
         Left            =   2040
         TabIndex        =   180
         Top             =   330
         Width           =   1155
      End
      Begin VB.OptionButton chkStump 
         Caption         =   "Stump"
         Height          =   255
         Left            =   1080
         TabIndex        =   178
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton chkNone 
         Caption         =   "Plain"
         Height          =   255
         Left            =   180
         TabIndex        =   176
         Top             =   330
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox cboDistalTape 
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tape"
         Height          =   195
         Index           =   1
         Left            =   4320
         TabIndex        =   200
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "At"
         Height          =   195
         Index           =   0
         Left            =   3300
         TabIndex        =   199
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proximal tape Style"
      Height          =   2625
      Left            =   4560
      TabIndex        =   149
      Top             =   5610
      Width           =   4875
      Begin VB.OptionButton optProximalTape 
         Caption         =   "Flap"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   198
         Top             =   390
         Width           =   735
      End
      Begin VB.Frame frmFlaps 
         Caption         =   "Flap Details"
         Height          =   1755
         Left            =   60
         TabIndex        =   162
         Top             =   750
         Width           =   4755
         Begin VB.TextBox txtFrontStrapLength 
            Height          =   285
            Left            =   3420
            TabIndex        =   206
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtWaistCir 
            Height          =   285
            Left            =   3420
            TabIndex        =   211
            Top             =   1380
            Width           =   495
         End
         Begin VB.TextBox txtStrapLength 
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   3420
            TabIndex        =   205
            Top             =   300
            Width           =   495
         End
         Begin VB.ComboBox cboFlaps 
            Height          =   315
            Left            =   660
            TabIndex        =   204
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtCustFlapLength 
            Height          =   285
            Left            =   3420
            TabIndex        =   208
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label lblFrontStrapLength 
            Caption         =   "Front Strap:"
            Height          =   285
            Left            =   2220
            TabIndex        =   147
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label labFrontStrapLength 
            Height          =   195
            Left            =   3960
            TabIndex        =   148
            Top             =   690
            Width           =   675
         End
         Begin VB.Label lblWaistCir 
            Caption         =   "Style ""D"" Waist Circum:"
            Height          =   285
            Left            =   1260
            TabIndex        =   209
            Top             =   1410
            Width           =   2115
         End
         Begin VB.Label labWaistCir 
            Height          =   195
            Left            =   3960
            TabIndex        =   210
            Top             =   1410
            Width           =   675
         End
         Begin VB.Label labCustFlapLength 
            Height          =   195
            Left            =   3960
            TabIndex        =   196
            Top             =   1050
            Width           =   675
         End
         Begin VB.Label labStrap 
            Height          =   195
            Left            =   3960
            TabIndex        =   197
            Top             =   330
            Width           =   675
         End
         Begin VB.Label lblStyle 
            Caption         =   "Style:"
            Height          =   285
            Left            =   120
            TabIndex        =   146
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lblStrapLength 
            Caption         =   "Back Strap:"
            Height          =   255
            Left            =   2220
            TabIndex        =   151
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label lblCustFlapLength 
            Caption         =   "Requested Flap length:"
            Height          =   285
            Left            =   1260
            TabIndex        =   152
            Top             =   1050
            Width           =   1995
         End
      End
      Begin VB.ComboBox cboProximalTape 
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   390
         Width           =   735
      End
      Begin VB.OptionButton optProximalTape 
         Caption         =   "Plain"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   144
         Top             =   390
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Tape"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   202
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "At"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   201
         Top             =   420
         Width           =   195
      End
   End
   Begin VB.TextBox txtFileNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   54
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtWristPleat1 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   1050
      Width           =   495
   End
   Begin VB.TextBox txtShoulderPleat1 
      Height          =   285
      Left            =   180
      TabIndex        =   21
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtPatientName 
      Height          =   285
      Left            =   780
      TabIndex        =   53
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   7920
      TabIndex        =   52
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton Draw 
      Caption         =   "Draw"
      Height          =   435
      Left            =   5220
      TabIndex        =   51
      Top             =   150
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Design for"
      Height          =   1425
      Left            =   4560
      TabIndex        =   42
      Top             =   690
      Width           =   4875
      Begin VB.ComboBox cboPressure 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   660
         Width           =   1575
      End
      Begin VB.ComboBox cboFabric 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1380
         TabIndex        =   47
         Top             =   1020
         Width           =   3135
      End
      Begin VB.TextBox txtDiagnosis 
         Height          =   285
         Left            =   1380
         TabIndex        =   45
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblPressure 
         Caption         =   "Pressure:"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblDiag 
         Caption         =   "Diagnosis:"
         Height          =   255
         Left            =   180
         TabIndex        =   49
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblFabric 
         Caption         =   "Fabric:"
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   1050
         Width           =   1095
      End
   End
   Begin VB.TextBox txtWristPleat2 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   1710
      Width           =   495
   End
   Begin VB.TextBox txtShoulderPleat2 
      Height          =   285
      Left            =   180
      TabIndex        =   22
      Top             =   6570
      Width           =   495
   End
   Begin VB.TextBox txtinchflag 
      Height          =   285
      Left            =   1620
      TabIndex        =   39
      Top             =   7770
      Width           =   855
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1410
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1770
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2130
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   4
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2490
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   5
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   7
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   10
      Top             =   3570
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   8
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   11
      Top             =   3930
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   9
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   12
      Top             =   4290
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   10
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4650
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   11
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   14
      Top             =   5010
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   12
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   15
      Top             =   5370
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   13
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   16
      Top             =   5730
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   14
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   17
      Top             =   6090
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   15
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   18
      Top             =   6450
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   16
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   19
      Top             =   6810
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   17
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   20
      Top             =   7170
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   23
      Top             =   1050
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      Top             =   1410
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   25
      Top             =   1770
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   26
      Top             =   2130
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   27
      Top             =   2490
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   28
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   29
      Top             =   3210
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   30
      Top             =   3570
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   8
      Left            =   2640
      TabIndex        =   31
      Top             =   3930
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   9
      Left            =   2640
      TabIndex        =   32
      Top             =   4290
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   10
      Left            =   2640
      TabIndex        =   33
      Top             =   4650
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   12
      Left            =   2640
      TabIndex        =   35
      Top             =   5370
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   13
      Left            =   2640
      TabIndex        =   36
      Top             =   5730
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   14
      Left            =   2640
      TabIndex        =   37
      Top             =   6090
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   15
      Left            =   2640
      TabIndex        =   38
      Top             =   6450
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   16
      Left            =   2640
      TabIndex        =   40
      Top             =   6810
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   17
      Left            =   2640
      TabIndex        =   41
      Top             =   7170
      Width           =   495
   End
   Begin VB.TextBox mms 
      Height          =   285
      Index           =   11
      Left            =   2640
      TabIndex        =   34
      Top             =   5010
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   6
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3210
      Width           =   495
   End
   Begin VB.TextBox Length 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      Height          =   435
      Left            =   6600
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblPleat 
      Caption         =   "Pleat"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   112
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label lblPleat 
      Caption         =   "Pleat"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   220
      Top             =   6270
      Width           =   495
   End
   Begin VB.Label lblPleat 
      Caption         =   "Pleat"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   219
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "Elbow"
      Height          =   285
      Left            =   180
      TabIndex        =   145
      Top             =   4680
      Width           =   555
   End
   Begin VB.Label lblPleat 
      Caption         =   "Pleat"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   135
      Top             =   810
      Width           =   495
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "-6"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   134
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "-4½"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   133
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "-3"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   132
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "-1½"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   131
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   130
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "1½"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   129
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   128
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "4½"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   127
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   126
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "7½"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   125
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "9"
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   124
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "10½"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   123
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "12"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   122
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "13½"
      Height          =   255
      Index           =   13
      Left            =   840
      TabIndex        =   121
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "15"
      Height          =   255
      Index           =   14
      Left            =   840
      TabIndex        =   120
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "16½"
      Height          =   255
      Index           =   15
      Left            =   840
      TabIndex        =   119
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "18"
      Height          =   255
      Index           =   16
      Left            =   840
      TabIndex        =   118
      Top             =   6810
      Width           =   375
   End
   Begin VB.Label lblTape 
      Alignment       =   2  'Center
      Caption         =   "19½"
      Height          =   255
      Index           =   17
      Left            =   840
      TabIndex        =   117
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label23 
      Caption         =   " Length"
      Height          =   255
      Left            =   1200
      TabIndex        =   116
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   " MMs"
      Height          =   255
      Left            =   2640
      TabIndex        =   115
      Top             =   810
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "Gram"
      Height          =   255
      Left            =   3240
      TabIndex        =   114
      Top             =   810
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "Red"
      Height          =   255
      Left            =   3840
      TabIndex        =   113
      Top             =   810
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   111
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   110
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   109
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   108
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   107
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   106
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   105
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   104
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   103
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   102
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   101
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   100
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   99
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   98
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   97
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   96
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   16
      Left            =   3240
      TabIndex        =   95
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label gms 
      Height          =   255
      Index           =   17
      Left            =   3240
      TabIndex        =   94
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   93
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   92
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   91
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   90
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   89
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   88
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   87
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   86
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   85
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   84
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   83
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   82
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   12
      Left            =   3840
      TabIndex        =   81
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   13
      Left            =   3840
      TabIndex        =   80
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   14
      Left            =   3840
      TabIndex        =   79
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   15
      Left            =   3840
      TabIndex        =   78
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   16
      Left            =   3840
      TabIndex        =   77
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label reds 
      Height          =   255
      Index           =   17
      Left            =   3840
      TabIndex        =   76
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "File No:"
      Height          =   255
      Left            =   3120
      TabIndex        =   75
      Top             =   270
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Patient:"
      Height          =   255
      Left            =   60
      TabIndex        =   74
      Top             =   270
      Width           =   675
   End
   Begin VB.Label lblUnits 
      Caption         =   "Units:"
      Height          =   255
      Left            =   1020
      TabIndex        =   73
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   72
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   71
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   70
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   69
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   68
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   67
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   66
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   65
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   64
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   63
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   62
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   61
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   12
      Left            =   1920
      TabIndex        =   60
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   13
      Left            =   1920
      TabIndex        =   59
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   14
      Left            =   1920
      TabIndex        =   58
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   15
      Left            =   1920
      TabIndex        =   57
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   16
      Left            =   1920
      TabIndex        =   56
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label InchText 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   17
      Left            =   1920
      TabIndex        =   55
      Top             =   7200
      Width           =   615
   End
End
Attribute VB_Name = "armdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   ARMDIA.MAK
'Purpose:   Drawing Arms and Vest/Bodysuit sleeves.
'
'Version:   3.07
'Date:      94/95
'Author:    Ciaran McKavangh    (Armdia.FRM, mostly)
'           Gary George         (Utility.BAS)
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Feb.95     CMcK    Error and bug fixes
'                   G.Dunne Fax, 20.Feb.95
'
'Mar.95     GG      Error and bug fixes
'                   G.Dunne Fax, 20.Mar.95 and 03.Apr.95    z
'
'Apr.95     GG      Allow multiple ARM styles in a single
'                   drawing
'
'Jul.95     GG      Triton to CAD link
'
'Oct.95     GG      Triton to CAD link
'                   and Sleeveles Vests
'
'Dec.95     GG      Mods to decrease code
'
'Sep.97     GG      Added support for bodysuit and bodybrief
'
'Jan.98     GG      Added support for new mesh method.
'                   Added support for side scooped necks
'
'Dec 98     GG      Ported to VB5
'Note:-
'
'
Dim Measurements(1 To 18)  As Variant
Dim ChangeChecker(1 To 18)  As Variant
Dim MMsChecker(1 To 18)  As Variant
Dim tapewidths(1 To 18) As Variant
Dim Pow(1 To 19, 1 To 23) As Variant
Dim Bob(1 To 19, 1 To 23) As Variant
Dim Zero, one, two, three, four, five, six, seven, eight, nine, ten

'* Windows API Functions Declarations
' Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
' Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
' Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
' Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
' Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
   
'Constanst used by GetWindow
' Const GW_CHILD = 5
' Const GW_HWNDFIRST = 0
' Const GW_HWNDLAST = 1
' Const GW_HWNDNEXT = 2
' Const GW_HWNDPREV = 3
' Const GW_OWNER = 4

   'MsgBox constant
'    Const IDCANCEL = 2
'    Const IDYES = 6
'    Const IDNO = 7

 Dim EmptySleeve 'Flag to suppress calculation


Private Sub BackCalcMMs()
Dim Backgram
If Length(w) <> "" Then
' check MMs
    If (w < 9) And (w = (txtFirstTape - 1)) And txtStump = "False" And (Val(reds(w) > 14)) Then
        reds(w) = 14
        If Left$(g_sFabnam, 3) = "Pow" Then
            Backgram = Pow(Modulus, 5)
        Else
            Backgram = Bob(Modulus, 5)
        End If
        mms(w) = round(Backgram / Measurements(X))
        gms(w) = round(Backgram)
    ElseIf (w < 9) And (w = (txtFirstTape - 1)) And (Val(reds(w) < 10)) Then
        If Left$(g_sFabnam, 3) = "Pow" Then
            Backgram = Pow(Modulus, 1)
            If Val(gms(w)) < Backgram Then
                mms(w) = round(Backgram / Measurements(X))
                reds(w) = 10
                gms(w) = round(Backgram)
            End If
        Else
            Backgram = Bob(Modulus, 1)
            If Val(gms(w)) < Backgram Then
                mms(w) = round(Backgram / Measurements(X))
                reds(w) = 10
                gms(w) = round(Backgram)
            End If
        End If
    End If
End If
' end check wrist reduction


End Sub

Private Function BobReduction(Modul, lblGram)
one = 1
two = 2
nine = 9
eight = 8
ten = 10
           j = one
                If (Val(Bob(Modul, 1)) = Val(lblGram)) Then
                       lblred = ten
                End If
                j = two
                Do While j <= 22
                    If (Val(Bob(Modul, j)) = Val(lblGram)) Then
                        lblred = Val(j + nine)
                    ElseIf (Val(lblGram) - Val(Bob(Modul, j - one))) <= (Val(Bob(Modul, j)) - Val(lblGram)) Then
                        lblred = Val(j + eight)
                        Exit Do
                    End If
                    j = j + one
                Loop
                BobReduction = lblred
End Function

Private Sub Calculate_Click()
   'Force a recalculation
   '
    Dim ii As Integer
   '
    If cboPressure.Text = "" Then
        If armdia.Visible = True Then MsgBox "Pressure has not been selected", 48, "Arm Details"
        Exit Sub
    End If
    If cboFabric.Text = "" Then
        If armdia.Visible = True Then MsgBox "Fabric has not been selected", 48, "Arm Details"
        Exit Sub
    End If

    PR_CheckValues
            
    If cboProximalTape.ListIndex = 0 Then cboProximalTape_Click
    If cboDistalTape.ListIndex = 0 Then cboDistalTape_Click
    If g_iStyleLastTape < g_iStyleFirstTape Then
        If armdia.Visible = True Then MsgBox "Chosen distal tape is above Proximal tape!", 48, "Arm Details"
        Exit Sub
    End If
    
    If chkGauntlets.Value = True Then
        If cboPalm.ListIndex < 0 Then
            If armdia.Visible = True Then MsgBox "Palm tape position not given", 48, "Gauntlet Details"
            Exit Sub
        End If
        If cboWrist.ListIndex < 0 Then
            If armdia.Visible = True Then MsgBox "Wrist tape position not given", 48, "Gauntlet Details"
            Exit Sub
        End If
        PR_GetGauntletDetails
        PR_GauntletHoleCheck
    End If

    g_sPressureChange = "True" 'Use this to force a recalculation
    PR_GetStartEndTapes
    PR_GetPressure
    PR_GetFabric
    PR_GetGauntletDetails
    For ii = 0 To 17
        If Val(Length(ii)) > 0 Then
            PR_Pressurecalc ii
        End If
    Next ii
    PR_SetRresults
    PR_SetGlobals
    g_sPressureChange = "False"

   'Clean up displayed values
    For ii = 0 To txtFirstTape - 2
        If mms(ii).Text = "" Or mms(ii).Text = "0" Then
            mms(ii) = ""
            gms(ii) = ""
            reds(ii) = ""
        End If
    Next ii

    For ii = 17 To txtLastTape Step -1
        If mms(ii).Text = "" Or mms(ii).Text = "0" Then
            mms(ii) = ""
            gms(ii) = ""
            reds(ii) = ""
        End If
    Next ii

    If armdia.Visible = True Then g_Modified = True
    EmptySleeve = False

End Sub

Private Sub cboDistalTape_Click()
   'Set style first tape
   'Modify tape labels to show this (lblTape index starts at 0)
   'NOTE:-
   '    cboDistalTape.ListIndex = 0 = 1st
   '    cboDistalTape.ListIndex = 1 = -6
   '    cboDistalTape.ListIndex = 2 = -4½
   '    cboDistalTape.ListIndex = 3 = -3
   '                .
   '                .
   '                .
   '    cboDistalTape.ListIndex = 18 = 19½
   '
    Dim iTape As Integer

   'Reset display of previous to Black
    If g_iStyleFirstTape > 0 Then
        lblTape(g_iStyleFirstTape - 1).ForeColor = RGB(0, 0, 0)
    End If

   'set style first tape value
    iTape = cboDistalTape.ListIndex - 1
    If iTape < 0 Then
       'Use last tape
        If g_iFirstTape > 0 Then
            g_iStyleFirstTape = g_iFirstTape
        Else
            g_iStyleFirstTape = -1
        End If
    Else
       'use tape pointed at by index
        g_iStyleFirstTape = cboDistalTape.ListIndex
    End If

   'Set colour of tape label to green
    iTape = g_iStyleFirstTape - 1 'lblTape Index starts at 0
    If iTape >= 0 Then
        lblTape(iTape).ForeColor = RGB(0, 255, 0)
    End If

    If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub cboFabric_Click()
Dim fabnam, Fabtype
    ' check for fabric modulus
    If cboFabric <> "" Then
        fabnam = Left$(cboFabric.Text, 3)
        Fabtype = Mid$(cboFabric.Text, 5, 3)
        If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
            MsgBox "Fabric Modulus Not Recognised", 48, "Arm Details"
            cboFabric.Text = ""
            Exit Sub
        End If
        If (fabnam <> "Pow") And (fabnam <> "Bob") Then
            MsgBox "Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details"
            cboFabric.Text = ""
            Exit Sub
        End If
        If armdia.Visible = True Then g_Modified = True

    End If
    
End Sub

Private Sub cboFlaps_Click()
Dim w, X
If cboFlaps.Text = "none" Then
    cboFlaps.Text = ""
    txtStrap = ""
    txtFlap = ""
    labStrap = ""
    txtCustFlapLength = ""
    labCustFlapLength = ""
End If
X = Val(txtLastTape)
w = X - 1
If X > 1 Then
    If Length(w) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w) <> "" And reds(w) <> "" Then
        txtFlap = cboFlaps.Text
        PR_LastTapeRed w, X, g_sModulus
        PR_CheckFlaps w
        PR_SetRresults
        PR_SetGlobals
    End If
End If

   'Enable waist circumference for style "D"
    If InStr(1, cboFlaps.Text, "D", 0) > 0 Then
        labWaistCir.Enabled = True      'Inches convesion
        txtWaistCir.Enabled = True      'Edit box
        lblWaistCir.Enabled = True
    ElseIf txtWaistCir.Enabled = True Then
       'Disable only if previously enabled
        txtWaistCir.Text = ""       'Delete display of value
        labWaistCir.Caption = ""    'Delete inches display of value
        labWaistCir.Enabled = False
        txtWaistCir.Enabled = False
        lblWaistCir.Enabled = False
    End If

   'Disable all boxes if Vest Raglan chosen
    If Left$(cboFlaps.Text, 4) = "Vest" Or Left$(cboFlaps.Text, 4) = "Body" Then
        txtStrapLength.Text = ""
        txtStrapLength.Enabled = False
        txtFrontStrapLength.Text = ""
        txtFrontStrapLength.Enabled = False
        lblFrontStrapLength.Enabled = False
        labFrontStrapLength.Caption = ""
        labStrap.Caption = ""
        lblStrapLength.Enabled = False
        txtCustFlapLength.Text = ""
        txtCustFlapLength.Enabled = False
        labCustFlapLength.Caption = ""
        lblCustFlapLength.Enabled = False
    Else
       'Enable if not enbaled only
        If txtStrapLength.Enabled = False Then
            txtStrapLength.Enabled = True
            lblStrapLength.Enabled = True
        End If
        If txtCustFlapLength.Enabled = False Then
            txtCustFlapLength.Enabled = True
            lblCustFlapLength.Enabled = True
        End If
        If txtFrontStrapLength.Enabled = False Then
            txtFrontStrapLength.Enabled = True
            lblFrontStrapLength.Enabled = True
        End If
    End If
    If armdia.Visible = True Then g_Modified = True


End Sub

Private Sub cboPalm_LostFocus()
If cboWrist.ListIndex > -1 Then
    If cboWrist.ListIndex <= cboPalm.ListIndex Then
        MsgBox "Palm No. must be less than Wrist No.", 48, "Gauntlet Details"
        cboWrist.SetFocus
        Exit Sub
    End If
End If

    If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub cboPressure_Click()
    Dim Index As Integer
    
    If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub cboProximalTape_Click()
   'Set style last tape
   'Modify tape labels to show this (lblTape index starts at 0)
   'NOTE:-
   '    cboProximal.ListIndex = 0 = Last
   '    cboProximal.ListIndex = 1 = 19½
   '    cboProximal.ListIndex = 2 = 18
   '    cboProximal.ListIndex = 3 = 16½
   '                .
   '                .
   '                .
   '    cboProximal.ListIndex = 18 = -6
   '
    Dim iTape As Integer

   'Reset display of previous to Black
    If g_iStyleLastTape > 0 Then
        lblTape(g_iStyleLastTape - 1).ForeColor = RGB(0, 0, 0)
    End If

   'set style last tape value
    iTape = cboProximalTape.ListIndex - 1
    If iTape < 0 Then
       'Use last tape
        If g_iLastTape > 0 Then
            g_iStyleLastTape = g_iLastTape
        Else
            g_iStyleLastTape = -1
        End If
    Else
       'use tape pointed at by index
        g_iStyleLastTape = (18 - cboProximalTape.ListIndex) + 1
    End If

   'Set colour of tape label to red
    iTape = g_iStyleLastTape - 1 'lblTape Index starts at 0
    If iTape >= 0 Then
        lblTape(iTape).ForeColor = RGB(255, 0, 0)
    End If


    If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub cboWrist_LostFocus()
If cboPalm.ListIndex > -1 Then
    If cboPalm.ListIndex >= cboWrist.ListIndex Then
        MsgBox "Wrist No. must be greater than Palm No.", 48, "Gauntlet Details"
        cboPalm.SetFocus
        Exit Sub
    End If
End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub chkDetachable_Click()
'Setup for Detachable gauntlet

    If chkDetachable.Value = 1 Then
       'For detachable gauntlets
       'Note:- cboProximalTape gives the tapes in reverse order
        cboProximalTape.ListIndex = 13 '+1-1/2 tape
        
        'If Val(txtAge) > 14 Then
        '    cboProximalTape.ListIndex = 12 '+3 tape for adults
        'Else
        '    cboProximalTape.ListIndex = 13 '+1-1/2 tape for children
        'End If
        cboProximalTape_Click
        optProximalTape(0).Value = 1
        optProximalTape(1).Value = 0
    ElseIf chkDetachable.Value = 0 Then
       'For normal gauntlets us last tape
        cboProximalTape.ListIndex = 0
        cboProximalTape_Click
    End If
End Sub

Private Sub chkEnclosed_Click()
    If chkEnclosed.Value = 1 Then
        chkNoThumb.Enabled = False
        chkNoThumb.Value = False
        txtCircum.Enabled = True
        txtThumb.Enabled = True
        lblCircum.Enabled = True
        lblThumb.Enabled = True
    ElseIf chkEnclosed.Value = 0 Then
        chkNoThumb.Enabled = True
    
    End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub chkGauntlets_Click()
   If chkGauntlets.Value = True Then
        frmGauntlet.Enabled = True
        chkStump.Value = False
        lblPalm.Enabled = True
        lblWrist.Enabled = True
        lblThumb.Enabled = True
        cboPalm.Enabled = True
        cboWrist.Enabled = True
        txtThumb.Enabled = True
        chkEnclosed.Enabled = True
        chkDetachable.Enabled = True
        chkNoThumb.Enabled = True
        lblPalmWrist.Enabled = True
        txtPalmWrist.Enabled = True
        txtCircum.Enabled = True
        lblCircum.Enabled = True
        lblGauntletExtension.Enabled = True
        txtGauntletExtension.Enabled = True

       'Set distal tape position to first
        cboDistalTape.Enabled = True
        cboDistalTape.ListIndex = 0
'        cboDistalTape_Click
        cboDistalTape.Enabled = False

   ElseIf chkGauntlets.Value = False Then
        frmGauntlet.Enabled = False
        chkNoThumb.Enabled = False
        lblPalmWrist.Enabled = False
        txtPalmWrist.Enabled = False
        lblPalm.Enabled = False
        lblWrist.Enabled = False
        lblThumb.Enabled = False
        cboPalm.Enabled = False
        cboWrist.Enabled = False
        txtThumb.Enabled = False
        chkEnclosed.Enabled = False
        chkDetachable.Enabled = False
        txtCircum.Enabled = False
        lblCircum.Enabled = False
        lblGauntletExtension.Enabled = True
        txtGauntletExtension.Enabled = True

   End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub chkNone_Click()
   'Disable gauntlet
    g_sHoleCheck = "False"
    frmGauntlet.Enabled = False
    lblPalm.Enabled = False
    lblWrist.Enabled = False
    lblThumb.Enabled = False
    cboPalm.ListIndex = -1
    cboWrist.ListIndex = -1
    cboPalm.Enabled = False
    cboWrist.Enabled = False
        lblGauntletExtension.Enabled = False
        txtGauntletExtension.Enabled = False

    
    If chkDetachable.Value = 1 Then
       'Reset the proximal tape to last
        chkDetachable.Value = 0
        chkDetachable_Click
    End If

    chkNoThumb.Value = 0
    chkEnclosed.Value = 0
    PWDist = ""
    TCirc = ""
    TLen = ""
    txtThumb = ""
    labGauntletExtension = ""
    txtGauntletExtension = ""
    txtThumb.Enabled = False
    chkEnclosed.Enabled = False
    chkDetachable.Enabled = False
    chkNoThumb.Enabled = False
    lblPalmWrist.Enabled = False
    txtPalmWrist = ""
    txtPalmWrist.Enabled = False
    txtCircum = ""
    txtCircum.Enabled = False
    lblCircum.Enabled = False

    txtGauntletFlag.Text = "False"

   'Enable distal tape position
    cboDistalTape.Enabled = True

   'Disable stump
    chkStump.Value = False
    txtStump.Text = "False"
    
If Val(txtFirstTape) > 0 Then
    Dim w, X
    X = Val(txtFirstTape)
    w = X - 1
    If Length(w) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w) <> "" And reds(w) <> "" Then
        PR_WristRed w, X, g_sModulus
        PR_Reductions w, X
        PR_SetRresults
        PR_SetGlobals
    End If
End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub chkNoThumb_Click()
    If chkNoThumb.Value = 1 Then
        chkEnclosed.Enabled = False
        chkEnclosed.Value = 0
        txtCircum.Enabled = False
        txtThumb.Enabled = False
        lblCircum.Enabled = False
        lblThumb.Enabled = False
        txtCircum = ""
        TCirc = ""
        txtThumb = ""
        TLen = ""
    
    ElseIf chkNoThumb.Value = 0 And chkGauntlets.Value = True Then
        txtCircum.Enabled = True
        txtThumb.Enabled = True
        lblCircum.Enabled = True
        lblThumb.Enabled = True
        chkEnclosed.Enabled = True
     End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub chkStump_Click()
Dim entries, i
If chkStump.Value = True Then
   'Disable gauntlet
    g_sHoleCheck = "False"
    frmGauntlet.Enabled = False
    lblPalm.Enabled = False
    lblWrist.Enabled = False
    lblThumb.Enabled = False
    cboPalm.ListIndex = -1
    cboPalm.Enabled = False
    cboWrist.ListIndex = -1
    cboWrist.Enabled = False
    If chkDetachable.Value = 1 Then
       'Reset the proximal tape to last
        chkDetachable.Value = 0
        chkDetachable_Click
    End If
    chkNoThumb.Value = 0
    chkEnclosed.Value = 0
    txtGauntletExtension.Enabled = False
    lblGauntletExtension.Enabled = False
    labGauntletExtension = ""
    txtGauntletExtension = ""
    PWDist = ""
    TCirc = ""
    TLen = ""
    txtThumb = ""
    txtThumb.Enabled = False
    chkEnclosed.Enabled = False
    chkDetachable.Enabled = False
    chkNoThumb.Enabled = False
    lblPalmWrist.Enabled = False
    txtPalmWrist = ""
    txtPalmWrist.Enabled = False
    txtCircum = ""
    txtCircum.Enabled = False
    lblCircum.Enabled = False
    txtGauntletFlag.Text = "False"


   'Set distal tape position to first
    cboDistalTape.Enabled = True
    cboDistalTape.ListIndex = 0
    cboDistalTape_Click
    'cboDistalTape.Enabled = False

ElseIf chkStump.Value = 0 Then
    chkGauntlets.Enabled = True
End If

If Val(txtFirstTape) > 0 Then
    Dim w, X
    X = Val(txtFirstTape)
    w = X - 1
    If Length(w) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w) <> "" And reds(w) <> "" Then
        If chkStump.Value = True Then
            PR_Reductions w, X
            PR_Stump w, X
            PR_SetRresults
            PR_SetGlobals
        End If
    End If
End If

If armdia.Visible = True Then g_Modified = True

End Sub

Private Sub cmdCancel_Click()
    Dim Response As Integer
    If g_Modified Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "Arm Details")
        Select Case Response
            Case IDYES
                fNum = FN_Open("c:\jobst\draw.d", txtPatientName, txtFileNo, txtSleeve, "Save")
               'Check Flaps
                txtFlap = cboFlaps.Text
                If txtFlap = "" Then g_sFlap = "None" Else g_sFlap = txtFlap
                If txtFlap <> "" And txtStrapLength = "" Then
                    If txtinchflag = "cm" Then
                        txtStrap = "61"
                    Else
                        txtStrap = "24"
                    End If
                End If
                PR_GetStartEndTapes
                PR_GetGauntletDetails
                PR_SetRresults
                PR_SetGlobals
                g_sFabric = cboFabric.Text
               'Update
                DataBaseDataUpDate "Save"
                Close fNum
                AppActivate fnGetDrafixWindowTitleText()
                SendKeys "@c:\jobst\draw.d{enter}", True
                End
            Case IDNO
                End
            Case IDCANCEL
                Exit Sub
        End Select
    Else
        End
    End If

End Sub

Private Function cmtoinches(flag, cm)
Dim inch, remainder As Double, temp As Integer, temprem, rnder
Dim Value As Double

If flag = "cm" Then
    inch = Val(cm) / 2.54
    temp = Int(inch)
    remainder = (inch - temp)
    remainder = remainder / 0.125
    realrem = Int(remainder)
    temprem = remainder - Int(remainder)
    If temprem >= 0.5 Then
        realrem = realrem + 1
    End If
    remainder = realrem * 0.125
    inch = temp + remainder
    cmtoinches = inch
ElseIf flag = "inches" Then
    Value = cm
    If Value >= 1 Then
        temp = Int(Value)
        remainder = Value - temp
        If remainder > 0.75 Then
            MsgBox "Invalid Measurements, Inches and Eights only", 48, "Arm Details"
            cmtoinches = "-1"
            Exit Function
        Else
            remainder = remainder * 10
            remainder = remainder * 0.125
            cmtoinches = temp + remainder
        End If
    Else
        remainder = Value
        If remainder > 0.75 Then
            MsgBox "Invalid Measurements, Inches and Eights only", 48, "Arm Details"
            cmtoinches = "-1"
            Exit Function
        Else
            remainder = remainder * 10
            remainder = remainder * 0.125
            cmtoinches = remainder
        End If
    
    End If
 End If

End Function

Private Sub Draw_Click()

Dim Profile As curve, conures As curve, Gauntend As curve, ii, nn, tps, numtapes, fNum As Integer
Dim xyPt As XY, xyPt1 As XY, xyTHole1 As XY, xyCen As XY, xyMidPoint As XY, xyBegin As XY, xyFin As XY, xyTHole2 As XY, xyThumbhole As XY, xyPt2 As XY, xyPleat As XY, xyElbow1 As XY
Dim actfabric, Pleatsgn, pl, GauntSize, hole, nPalm, nWrist, lenty
Dim nOpp1, nOpp2, nAdj1, nAdj2, nAng1, nAng2, nlen, midylen, nghk, n, m, valcheck
Dim xyThumbStart As XY, xyThumb1 As XY, xyThumb2 As XY, xyThumb3 As XY, xyThumb4 As XY, xyThumb5 As XY, xyThumb6 As XY, xyThumb7 As XY, xyThumb8 As XY
Dim thm1, thm2, thm3, thmAng1, thmang2, nTextHt, numtap, xyP As XY
Dim SeamDist, w, X, v As Integer, xyElbMark1 As XY, xyElbMark2 As XY
Dim iPleats As Integer
SeamDist = 0.1875
valcheck = NoVals()

If cboPressure.Text = "" Then
    MsgBox "Pressure has not been selected", 48, "Arm Details"
    Exit Sub
End If
If cboFabric.Text = "" Then
    MsgBox "Fabric has not been selected", 48, "Arm Details"
    Exit Sub
End If

PR_GetStartEndTapes
'Set style first and last tape
If g_iStyleFirstTape = -1 Then g_iStyleFirstTape = Val(txtFirstTape.Text)
If g_iStyleLastTape = -1 Then g_iStyleLastTape = Val(txtLastTape.Text)


If valcheck = 0 Then
    MsgBox "No tape measurements have been entered", 48, "Arm Details"
    Exit Sub
End If

If g_iStyleLastTape < g_iStyleFirstTape Then
    MsgBox "Chosen distal tape is above Proximal tape!", 48, "Arm Details"
    Exit Sub
End If

numtapes = (Val(txtLastTape) - Val(txtFirstTape)) + 1

If numtapes = 1 Then
    MsgBox "At least two tapes have to be entered before drawing", 48, "Arm Details"
    Exit Sub
End If


'Check if flap is style D and that a waist circumference is given
If InStr(1, cboFlaps.Text, "D", 0) > 0 And Val(txtWaistCir) = 0 Then
    MsgBox "A Waist circumference must be given for a D-Style flap", 48, "Arm Details"
    Exit Sub
End If

' Check if tape has been deleted
Dim offnumtapes, DeleteChecker, actnumtapes, tmpFirst, tmpLast
offnumtapes = (txtLastTape - txtFirstTape) + 1
    gg = 0
    tmpFirst = ""
    Do While gg < 18
         If Length(gg) <> "" Then
            tmpFirst = gg + 1
            Exit Do
         End If
         gg = gg + 1
    Loop

    ' secondlast & last tapes
    ff = 17
    tmpLast = ""
    Do While ff > 0
         If Length(ff) <> "" Then
            tmpLast = ff + 1
            
            Exit Do
         End If
         ff = ff - 1
    Loop
actnumtapes = (tmpLast - tmpFirst) + 1

If chkStump.Value = False Then
    txtStump = False
End If

    PR_GetPressure
    PR_GetFabric
    PR_CheckValues
    PR_GetGauntletDetails
    PR_GauntletHoleCheck
    PR_GetStartEndTapes
'    PR_CalculateMMSGmsReds 1
    PR_SetRresults
    PR_SetGlobals
    
'check Gauntlet details
    If chkGauntlets.Value = True Then
        If cboPalm.ListIndex = -1 Then
            MsgBox "No value for palm tape", 48, "Gauntlet Details"
            Exit Sub
        End If
        If cboWrist.ListIndex = -1 Then
            MsgBox "No value for wrist tape", 48, "Gauntlet Details"
            Exit Sub
        End If
    End If

' set tapewidths
For w = 0 To 17
    X = w + 1
    PR_SetDistances w, X
    PR_CheckFlaps w
Next w

Profile.n = numtapes
    If numtapes > 3 Then
    ii = 1
    For tps = Val(txtFirstTape) To Val(txtLastTape)
        If Measurements(tps) = "" Then
                tps = tps + 1
                Profile.n = Profile.n - 1
            End If
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
            Profile.X(ii) = Val(tapewidths(tps))
            Profile.y(ii) = actfabric + SeamDist
            ii = ii + 1
        Next tps
    End If
    If numtapes = 3 Then
        tps = Val(txtFirstTape)
        actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
        Profile.X(1) = Val(tapewidths(tps))
        Profile.y(1) = actfabric + SeamDist
        tps = tps + 2
        actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
        Profile.X(3) = Val(tapewidths(tps))
        Profile.y(3) = actfabric + SeamDist
        tps = tps - 1
        If Measurements(tps) <> "" Then                      'And cboFlaps.Text = ""
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
            Profile.X(2) = Val(tapewidths(tps))
            Profile.y(2) = actfabric + SeamDist
        End If
    End If
    If numtapes = 3 Then
        tps = Val(txtFirstTape)
        If Measurements(tps + 1) = "" And chkGauntlets.Value = True Then
            Profile.n = 2
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
            Profile.X(1) = Val(tapewidths(tps))
            Profile.y(1) = actfabric + SeamDist
            tps = tps + 2
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
            Profile.X(2) = Profile.X(1) + Val(txtPalmWristDist)
            Profile.y(2) = actfabric + SeamDist
       End If
    End If
    If numtapes = 2 And chkGauntlets.Value = False Then
        ii = 1
        For tps = Val(txtFirstTape) To Val(txtLastTape)
            If Measurements(tps) = "" Then
                tps = tps + 1
                Profile.n = Profile.n - 1
            End If
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1)))) / 100) / 2
            Profile.X(ii) = Val(tapewidths(tps))
            Profile.y(ii) = actfabric + SeamDist
            ii = ii + 1
        Next tps
    End If

'Check that there are more tape spaces than pleats
'I.E. So that pleats will fit
iPleats = 0
If Val(txtWristPleat1) > 0 And txtWristPleat1.Enabled Then iPleats = iPleats + 1
If Val(txtWristPleat2) > 0 And txtWristPleat2.Enabled Then iPleats = iPleats + 1
If Val(txtShoulderPleat1) > 0 And txtShoulderPleat1.Enabled Then iPleats = iPleats + 1
If Val(txtShoulderPleat2) > 0 And txtShoulderPleat2.Enabled Then iPleats = iPleats + 1
'As we can't have a pleat between palm and wrist we fudge it here to cut down on code.
If chkGauntlets.Value = True Then iPleats = iPleats + 1
If Profile.n <= iPleats Then
    MsgBox "There are more pleats given, than spaces available between tapes. Remove or temporarily disable pleats, by double clicking on the ""Pleat"" label!", 48, "Arm Details"
    Exit Sub
End If

   fNum = FN_Open("c:\jobst\draw.d", txtPatientName, txtFileNo, txtSleeve, "Draw")

    PR_SetLayer "Template" & txtSleeve
    PR_DrawFitted Profile
    PR_AddEntityID g_sID, "", "Profile"    'g_sID from FN_Open
    Print #fNum, "SetDBData(hEnt,"; QQ; "curvetype"; QCQ; "sleeveprofile"; QQ; ");"
     

    PR_NamedHandle "hSleeveProfile"


    PR_MakeXY xyPt1, Profile.X(1), 0
    PR_MakeXY xyPt2, Profile.X(Profile.n), 0
    PR_DrawLine xyPt1, xyPt2                        'Fold Line
    PR_AddEntityID g_sID, "", "FoldLine"   'sFileNo, sSide from FN_Open
     
    PR_SetLayer "Notes"
    PR_DrawLineOffset xyPt1, xyPt2, 0.6875      'Inner tram line
    PR_DrawLineOffset xyPt1, xyPt2, 0.1875      'seam line at 3/16ths
   
    If txtGauntletFlag = "False" Then
        PR_SetLayer "Template" & txtSleeve
        PR_MakeXY xyPt1, Profile.X(1), 0
        If Val(txtFirstTape) = 11 Then
            PR_MakeXY xyPt2, Profile.X(1), Profile.y(1) / 2
        Else
            PR_MakeXY xyPt2, Profile.X(1), Profile.y(1)
        End If
        PR_DrawLine xyPt1, xyPt2                        'Wrist Line
        PR_AddEntityID g_sID, "", "WristLine"  'sFileNo, sSide from FN_Open
         

    End If
    
    If cboFlaps = "" Or cboFlaps = " " Then
        PR_SetLayer "Template" & txtSleeve
        If Val(txtLastTape) = 11 Then
            PR_MakeXY xyPt1, Profile.X(Profile.n), Profile.y(Profile.n) / 2
        Else
            PR_MakeXY xyPt1, Profile.X(Profile.n), Profile.y(Profile.n)
        End If
        PR_MakeXY xyPt2, Profile.X(Profile.n), 0
        PR_DrawLine xyPt1, xyPt2                           'Shoulder Line
        PR_AddEntityID g_sID, "", "ShoulderLine"  'sFileNo, sSide from FN_Open
         
    End If

    ' if Old Drawing add DB Fields
    If g_sOld = "True" Then
        PR_OldDrawing
    End If

    nTextHt = 0.125

    ' Draw Circular Stump
    If txtStump = "True" Then
        PR_MakeXY xyPt1, Profile.X(1) - 5, Profile.y(1) / 2
        PR_DrawCircularStump xyPt1, Profile.y(1)
       'Add Stump text
        PR_SetLayer "Notes"
        g_nCurrTextAngle = 90
        PR_MakeXY xyPt1, Profile.X(1) + 0.25, Profile.y(1) / 2
        PR_DrawText "STUMP", xyPt1, nTextHt
        g_nCurrTextAngle = 0
    End If
    ' Insert Elbow mark
    ii = 1
    For tps = Val(txtFirstTape) To Val(txtLastTape)
        'Elbow Mark
        If tps = 11 And g_sHoleCheck <> "True" Then
            PR_SetLayer "Notes"
            PR_MakeXY xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4
            PR_MakeXY xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4
            PR_DrawLine xyElbMark1, xyElbMark2
            PR_DrawAssignDrafixVariable "xyElbow.x", Profile.X(ii)
            PR_DrawAssignDrafixVariable "xyElbow.y", Profile.y(ii)
        End If
        
        If tps = 10 And g_sHoleCheck = "True" Then
            PR_SetLayer "Notes"
            PR_MakeXY xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4
            PR_MakeXY xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4
            PR_DrawLine xyElbMark1, xyElbMark2
            PR_DrawAssignDrafixVariable "xyElbow.x", Profile.X(ii)
            PR_DrawAssignDrafixVariable "xyElbow.y", Profile.y(ii)
        End If
        ii = ii + 1
    Next tps

    
    ' insert Pleats
   If Val(txtWristPleat1) > 0 And txtWristPleat1.Enabled Then
        nTextHt = 0.125
        PR_SetLayer "Construct"
        If chkGauntlets = True Then
            n1 = max(Val(txtWristNo) - Val(txtFirstTape), 2)
            n2 = n1 + 1
        Else
            n1 = 1
            n2 = 2
        End If
        pl = (Profile.X(n2) - Profile.X(n1)) / 2
        PR_MakeXY xyPleat, Profile.X(n1) + pl, 0.3375
        PR_DrawText "Pleat", xyPleat, nTextHt
     End If
   If Val(txtWristPleat2) > 0 And txtWristPleat2.Enabled Then
        nTextHt = 0.125
        PR_SetLayer "Construct"
        If chkGauntlets = True Then
            n1 = max(Val(txtWristNo) - Val(txtFirstTape), 2) + 1
            n2 = n1 + 1
        Else
            n1 = 2
            n2 = 3
        End If
        pl = (Profile.X(n2) - Profile.X(n1)) / 2
        PR_MakeXY xyPleat, Profile.X(n1) + pl, 0.3375
        PR_DrawText "Pleat", xyPleat, nTextHt
     End If
   If Val(txtShoulderPleat1) > 0 And txtShoulderPleat1.Enabled Then
        nTextHt = 0.125
        PR_SetLayer "Construct"
        n = Profile.n
        m = Profile.n - 1
        pl = (Profile.X(n) - Profile.X(m)) / 2
        PR_MakeXY xyPleat, Profile.X(m) + pl, 0.3375
        PR_DrawText "Pleat", xyPleat, nTextHt
     End If
   If Val(txtShoulderPleat2) > 0 And txtShoulderPleat2.Enabled Then
        nTextHt = 0.125
        PR_SetLayer "Construct"
        n = Profile.n - 1
        m = Profile.n - 2
        pl = (Profile.X(n) - Profile.X(m)) / 2
        PR_MakeXY xyPleat, Profile.X(m) + pl, 0.3375
        PR_DrawText "Pleat", xyPleat, nTextHt
     End If
     
    'Draw Gauntlet
    If chkGauntlets.Value = True Then
        DrawGauntlet Profile, numtapes
    End If

    'Patient Details
    TypeDetails Profile, numtapes

    ' Elastics
    PR_SetLayer "Template" & txtSleeve
  
    Elastic Profile

    ' Shoulder Flaps
    If cboFlaps.Text = "" Or Left$(cboFlaps, 4) = "None" Then
        g_sFlap = "None"
    ElseIf (Left$(cboFlaps.Text, 4) = "Vest") Then
        g_sFlap = ""
        g_sFlapStrap = ""
        g_sFlap = cboFlaps.Text
        PR_DrawRaglan "Vest", txtVestRaglan, txtAge, txtVestID

    ElseIf (Left$(cboFlaps.Text, 4) = "Body") Then
        g_sFlap = ""
        g_sFlapStrap = ""
        g_sFlap = cboFlaps.Text
        PR_DrawRaglan "Body", txtVestRaglan, txtAge, txtVestID

    Else
        g_sFlap = cboFlaps.Text
        PR_DrawShoulderFlaps Profile, numtapes
    End If
    
    g_sFabric = cboFabric.Text

   'Vestarm and Profile Origin Marker data base update
   'For editing purposes the tape lengths are stored with
   'the Profile Origin Marker
    Dim sJustifiedString As String
    For ii = 0 To 17
        nValue = Val(Length(ii).Text)
        If nValue <> 0 And ii >= (Val(txtFirstTape) - 1) And ii <= (Val(txtLastTape) - 1) Then
            nValue = nValue * 10 'Shift decimal place to right
            sJustifiedString = String(3, " ")
            RSet sJustifiedString = Trim$(Str$(nValue))
        Else
            sJustifiedString = String(3, " ")
        End If

       'Tape values
        g_sEditLengths = g_sEditLengths & sJustifiedString
    Next ii

   'Update
    DataBaseDataUpDate "Draw"
    
    ii = 1
    For tps = Val(txtFirstTape) To Val(txtLastTape)
        If Measurements(tps) = "" Then
            tps = tps + 1
        End If
        PR_MakeXY xyPt1, Profile.X(ii), 0
        PR_PutTapeLabel tps, xyPt1, Measurements(tps), mms(tps - 1), gms(tps - 1), reds(tps - 1)
        ii = ii + 1
    Next tps

   'For Body suit et.al. start program to draw the mesh
    PR_DrawMesh

    Close fNum
    AppActivate fnGetDrafixWindowTitleText()
    SendKeys "@c:\jobst\draw.d{enter}", True
    End

End Sub

Private Sub DrawGauntlet(Profile As curve, numtapes As Variant)
    Dim Gauntend As curve, GauntDetachEnd As curve
    Dim xyPt1  As XY, xyTHole1 As XY, xyCen As XY, xyMidPoint As XY
    Dim xyThumbConstruct1 As XY, xyThumbConstruct2  As XY
    Dim GauntSize, hole, nPalm, lenty, nTextHt
    Dim nOpp1, nOpp2, nAdj1, nAdj2, nAng1, nlen, midylen, num, Ext
    Dim xyBegin As XY, xyFin As XY, xyTHole2 As XY, xyThumbhole As XY, xyPt2 As XY
    Dim xyGend1   As XY, xyGend2 As XY
    Dim ii, nn, tps, fNum, xyPt As XY, xyPleat  As XY
    Dim xyElbow1  As XY, xyElbow2 As XY, xyElbow3 As XY, xyElbow4 As XY
    Dim actfabric, Pleatsgn, pl, xyP  As XY, nAng2, nghk
    Dim xyThumbStart As XY, xyThumb1 As XY, xyThumb2 As XY, xyThumb3 As XY, xyThumb4 As XY, xyThumb5 As XY, xyThumb6 As XY, xyThumb7 As XY, xyThumb8 As XY
    Dim thm1, thm2, thm3, thmAng1, thmang2
    Dim xyThumbNotes As XY

       'Calaculate Gauntlet extension
        If txtAge > 14 And txtSex = "Male" Then
            GauntSize = 1.125
        ElseIf txtAge > 14 And txtSex = "Female" Then
            GauntSize = 0.625
        ElseIf txtAge < 15 Then
            GauntSize = 0.5
        End If
       'Use given extension if one is given
        If Val(txtGauntletExtension) > 0 Then GauntSize = cmtoinches(txtinchflag, txtGauntletExtension)

        num = (Val(txtLastTape) - Val(txtFirstTape)) + 2

       'Draw gauntlet extension
        PR_SetLayer "Template" & txtSleeve
        Gauntend.n = 4
        Gauntend.X(1) = Profile.X(1)
        Gauntend.y(1) = 0
        Gauntend.X(2) = Profile.X(1) - GauntSize
        Gauntend.y(2) = 0
        Gauntend.X(3) = Profile.X(1) - GauntSize
        Gauntend.y(3) = Profile.y(1)
        Gauntend.X(4) = Profile.X(1)
        Gauntend.y(4) = Profile.y(1)
        PR_DrawPoly Gauntend
        PR_AddEntityID g_sID, "", "GauntEnd"  'sFileNo, sSide from FN_Open
         

' start of added lines
    PR_MakeXY xyPt1, Gauntend.X(2), 0
    
    PR_MakeXY xyPt2, Gauntend.X(1), 0

    PR_DrawLine xyPt1, xyPt2                        'Fold Line
    PR_AddEntityID g_sID, "", "FoldLine"   'sFileNo, sSide from FN_Open
     
   
    PR_SetLayer "Notes"
    PR_DrawLineOffset xyPt1, xyPt2, 0.6875      'Inner tram line
    PR_DrawLineOffset xyPt1, xyPt2, 0.1875      'seam line at 3/16ths

' end of added lines

        
       ' Thumb opening
        nPalm = Val(txtPalmNo) - Val(txtFirstTape) + 1
        Ydiff = Profile.y(nPalm + 1) - Profile.y(nPalm)
        If Abs(Ydiff) > 0 Then 'Avoids overflow
            xyThumbhole.y = Profile.y(nPalm) + (Ydiff / 2)
        Else
            xyThumbhole.y = Profile.y(nPalm)
        End If
       'Draw thumb hole construction line
        PR_SetLayer "Construct"
        xyThumbConstruct1.y = xyThumbhole.y
        xyThumbConstruct1.X = Profile.X(nPalm) + (Val(txtPalmWristDist) / 2)
        PR_MakeXY xyThumbConstruct2, xyThumbConstruct1.X, 0.1875
        PR_DrawLine xyThumbConstruct1, xyThumbConstruct2
        PR_AddEntityID g_sID, "", "ThumbHoleConstruct"  'g_sID from FN_Open
         
        PR_AddMarker xyThumbConstruct1
        PR_AddEntityID g_sID, "", "ThumbHoleConstruct"  'g_sID from FN_Open
         
        
        ' Lower line
        xyThumbhole.y = ((xyThumbhole.y - 0.1875) / 2) + 0.375
        xyThumbhole.X = Profile.X(nPalm)
        
        PR_SetLayer "Template" & txtSleeve
        PR_MakeXY xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875
        PR_MakeXY xyTHole2, xyThumbhole.X + txtThumbDist - 0.1875, xyThumbhole.y - 0.1875
        PR_DrawLine xyTHole1, xyTHole2
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open
         
        ' Upper Line
        PR_SetLayer "Template" & txtSleeve
        PR_MakeXY xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875
        PR_MakeXY xyTHole2, xyThumbhole.X + txtThumbDist - 0.1875, xyThumbhole.y + 0.1875
        PR_DrawLine xyTHole1, xyTHole2
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open
         
        ' Thumb Reminder in layer notes
        PR_SetLayer "Notes"
        PR_MakeXY xyThumbNotes, xyThumbhole.X + 0.375, xyThumbhole.y + 0.5575
        If chkNoThumb.Value = 0 Then
          PR_DrawText "Thumb Attached", xyThumbNotes, 0.125
        Else
          PR_DrawText "No Thumb", xyThumbNotes, 0.125
        End If

        PR_SetLayer "Template" & txtSleeve
        
        ' Right Curve
        PR_MakeXY xyCen, xyThumbhole.X + txtThumbDist - 0.1875, xyThumbhole.y
        PR_MakeXY xyBegin, xyCen.X, xyCen.y + 0.1875
        PR_MakeXY xyFin, xyCen.X, xyCen.y - 0.1875
        PR_RightThumbHole xyCen, xyBegin, xyFin
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open
         
        ' Left top Curve
        PR_MakeXY xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875
        PR_MakeXY xyFin, xyThumbhole.X, xyThumbhole.y
        lenty = FN_CalcLength(xyBegin, xyFin)
        midylen = lenty / 2
        nAng1 = FN_CalcAngle(xyFin, xyBegin)
        nAng1 = nAng1 * (PI / 180)
        nAdj1 = Cos(nAng1) * midylen
        nOpp1 = Sin(nAng1) * midylen
        PR_MakeXY xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1
        nlen = Sqr(0.25 - (midylen * midylen))
        nAdj2 = Cos(nAng1) * nlen
        nOpp2 = Sin(nAng1) * nlen
        PR_MakeXY xyCen, xyMidPoint.X + nOpp2, xyMidPoint.y - nAdj2
        PR_DrawArc xyCen, xyBegin, xyFin
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open
        
        ' Left Bottom Curve
        PR_MakeXY xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875
        PR_MakeXY xyFin, xyThumbhole.X, xyThumbhole.y
        lenty = FN_CalcLength(xyBegin, xyFin)
        midylen = lenty / 2
        nAng1 = FN_CalcAngle(xyFin, xyBegin)
        nAng1 = nAng1 * (PI / 180)
        nAdj1 = Cos(nAng1) * midylen
        nOpp1 = Sin(nAng1) * midylen
        PR_MakeXY xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1
        nlen = Sqr(0.25 - (midylen * midylen))
        nAdj2 = Cos(nAng1) * nlen
        nOpp2 = Sin(nAng1) * nlen
        PR_MakeXY xyCen, xyMidPoint.X - nOpp2, xyMidPoint.y + nAdj2
        PR_DrawArc xyCen, xyBegin, xyFin
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open
        
        'Cutting line to Distal EOS
        PR_MakeXY xyBegin, xyThumbhole.X - GauntSize, xyThumbhole.y
        PR_DrawLine xyBegin, xyThumbhole
        PR_AddEntityID g_sID, "", "ThumbHole"  'g_sID from FN_Open

        If chkNoThumb.Value = 0 Then
            'Draw Thumb
            DrawThumb xyThumbhole
        End If
        If chkDetachable.Value = 1 Then
        End If

End Sub

Private Sub DrawThumb(xyThumbhole As XY)
    Dim xyThumbStart As XY, xyThumb1 As XY, xyThumb2 As XY, xyThumb3 As XY, xyThumb4 As XY, xyThumb5 As XY, xyThumb6 As XY, xyThumb7 As XY, xyThumb8 As XY
    Dim xyThumbBase1 As XY, xyThumbBase2 As XY, xyThumbBase3 As XY, xyThumbBase4 As XY
    Dim thm1, thm2, thm3, thmAng1, thmang2, nTextHt
    Dim xyFin As XY, xyBegin As XY, xyMidPoint  As XY, xyCen As XY, xyPt As XY
    Dim ThumbBase As curve, ThumbBaseBit, ThumbBaseLen
    Dim nThumbAngle As Double, nGauntletPlateAngle As Double
    Dim nSpacing As Double
           'Draw Thumb
            PR_MakeXY xyThumbStart, xyThumbhole.X, xyThumbhole.y + 3.1
            PR_MakeXY xyThumb1, xyThumbStart.X - 0.3125, xyThumbStart.y
            PR_MakeXY xyThumb2, xyThumbStart.X, xyThumb1.y + 0.25

           'Calculate Thumb Angle
            nGauntletPlateAngle = 90 - 37
            PR_CalcPolar xyThumb2, nGauntletPlateAngle, Val(txtTCircum.Text), xyPt
            PR_MakeXY xyThumb5, xyThumbStart.X + txtThumbDist, xyThumbStart.y
            nThumbAngle = FN_CalcAngle(xyThumb5, xyPt)

           'Thumb length adjustments for enclosed tip
            If chkEnclosed.Value = 1 Then
               'Adjust by age
                If Val(txtAge) > 14 Then
                    txtThumLen = txtThumLen - 0.25
                Else
                    txtThumLen = txtThumLen - 0.125
                End If
               'Adjust w.r.t radius
                txtThumLen = txtThumLen - (Val(txtTCircum.Text) / 2)
                If txtThumLen < 0 Then txtThumLen = 0.125
            End If

            PR_CalcPolar xyThumb2, nThumbAngle, Val(txtThumLen.Text), xyThumb3
            PR_CalcPolar xyThumb3, nThumbAngle - 90, Val(txtTCircum.Text), xyThumb4

            thmAng1 = FN_CalcAngle(xyThumb5, xyThumb4)
            thmAng1 = 180 - thmAng1
            thmAng1 = thmAng1 * (PI / 180)
            thm1 = Cos(thmAng1) * 0.1875
            thm2 = Sin(thmAng1) * 0.1875
            PR_MakeXY xyThumb6, xyThumb5.X + thm1, xyThumb5.y - thm2
            PR_MakeXY xyThumb7, xyThumb5.X - 0.1875, xyThumbStart.y - 0.375
            PR_MakeXY xyThumb8, xyThumbStart.X + 0.375, xyThumbStart.y - 0.375
            PR_DrawLine xyThumb2, xyThumb3
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbLength"  'sFileNo, sSide from FN_Open
            
           ' thumb tip
            If chkEnclosed.Value = 0 Then
                PR_DrawLine xyThumb3, xyThumb4
                PR_AddEntityID g_sFileNo, g_sSide, "ThumbCircum"  'sFileNo, sSide from FN_Open
            ElseIf chkEnclosed.Value = 1 Then
                ThumbTip xyThumb3, xyThumb4
                PR_AddEntityID g_sFileNo, g_sSide, "ThumbCircum"  'sFileNo, sSide from FN_Open
            End If
            PR_DrawLine xyThumb4, xyThumb5
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbBack"  'sFileNo, sSide from FN_Open
            PR_DrawLine xyThumb5, xyThumb6
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbBack2"  'sFileNo, sSide from FN_Open
           
           'Draw base of thumb
            PR_DrawLine xyThumb7, xyThumb8
            PR_AddEntityID g_sFileNo, g_sSide, "ThumBase"  'sFileNo, sSide from FN_Open
            
           'Proximal bottom curve of thumb
            lenty = FN_CalcLength(xyThumb7, xyThumb6)
            midylen = lenty / 2
            nAng = FN_CalcAngle(xyThumb7, xyThumb6)
            nAng = (90 - nAng) * (PI / 180)
            nHyp = midylen / Cos(nAng)
            PR_MakeXY xyCen, xyThumb7.X, xyThumb7.y + nHyp
            PR_DrawArc xyCen, xyThumb7, xyThumb6
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbRightCurve"  'sFileNo, sSide from FN_Open
            
           'Distal Bottom thumb Curve
            lenty = FN_CalcLength(xyThumb1, xyThumb8)
            midylen = lenty / 2
            nAng = FN_CalcAngle(xyThumb8, xyThumb1)
            nAng = (nAng - 90) * (PI / 180)
            nHyp = midylen / Cos(nAng)
            PR_MakeXY xyCen, xyThumb8.X, xyThumb8.y + nHyp
            PR_DrawArc xyCen, xyThumb1, xyThumb8
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbLeftCurve"  'sFileNo, sSide from FN_Open
        
           'J Curve
            PR_MakeXY xyBegin, xyThumb2.X, xyThumb2.y
            PR_MakeXY xyFin, xyThumb1.X, xyThumb1.y
            lenty = FN_CalcLength(xyFin, xyBegin)
            midylen = lenty / 2
            nAng1 = FN_CalcAngle(xyFin, xyBegin)
            nAng1 = nAng1 * (PI / 180)
            nAdj1 = Cos(nAng1) * midylen
            nOpp1 = Sin(nAng1) * midylen
            PR_MakeXY xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1
            nlen = Sqr(0.1225 - (midylen * midylen))
            nAng2 = FN_CalcAngle(xyFin, xyBegin)
            nAng2 = 90 - nAng2
            nAng2 = nAng2 * (PI / 180)
            nAdj2 = Cos(nAng2) * nlen
            nOpp2 = Sin(nAng2) * nlen
            PR_MakeXY xyCen, xyMidPoint.X - nAdj2, xyMidPoint.y + nOpp2
            PR_DrawArc xyCen, xyBegin, xyFin
            PR_AddEntityID g_sFileNo, g_sSide, "ThumbJCurve"  'sFileNo, sSide from FN_Open
        
            PR_SetLayer "Notes"
            nTextHt = 0.125
            PR_MakeXY xyPt, xyThumbStart.X, xyThumbStart.y + 0.375
   
   'Main details in Green on layer notes
    Dim sWorkOrder$, sText$

    PR_SetLayer "Notes"
    PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
    
    PR_SetLayer "Notes"
    If txtWorkOrder.Text = "" Then
        sWorkOrder = "-"
    Else
        sWorkOrder = txtWorkOrder.Text
    End If
    sText = txtSleeve.Text & "\n" & txtPatientName.Text & "\n" & sWorkOrder
    PR_DrawText sText, xyPt, 0.125
    

End Sub

Private Sub Elastic(Profile As curve)
Dim elastictop As curve, xyEpt1 As XY, xyEpt2 As XY, xyEpt3 As XY

    g_nCurrTextAngle = 90
    PR_SetTextData 2, 16, -1, -1, -1                'Horiz center, Vertical center

    elastictop.n = 3
    elastictop.X(1) = Profile.X(Profile.n)
    elastictop.y(1) = Profile.y(Profile.n) / 2
    elastictop.y(2) = Profile.y(Profile.n) / 2
    elastictop.y(3) = Profile.y(Profile.n) + 0.375

   'Proximal Elastic and Cutback
    If Val(txtAge) > 10 And Val(txtLastTape) = 11 Then
        elastictop.X(2) = Profile.X(Profile.n) - 0.75
        elastictop.X(3) = Profile.X(Profile.n) - 0.75
        PR_SetLayer "Template" & txtSleeve
        PR_DrawPoly elastictop
    ElseIf Val(txtAge) <= 10 And Val(txtLastTape) = 11 Then
       'Proximal Elastic @ Elbow
        elastictop.X(2) = Profile.X(Profile.n) - 0.375
        elastictop.X(3) = Profile.X(Profile.n) - 0.375
        PR_SetLayer "Template" & txtSleeve
        PR_DrawPoly elastictop
        PR_MakeXY xyEpt1, Profile.X(Profile.n) - 0.75, Profile.y(Profile.n) / 2
        PR_SetLayer "Notes"
        PR_DrawText "1/2'' Elastic", xyEpt1, 0.125
    ElseIf cboFlaps.Text = "" And Val(txtAge) <= 10 And Val(txtLastTape) <> 11 Then
       'Proximal Elastic on plain arms for children under 10
        PR_MakeXY xyEpt1, Profile.X(Profile.n) - 0.25, Profile.y(Profile.n) / 2
        PR_SetLayer "Notes"
        PR_DrawText "1/2'' Elastic", xyEpt1, 0.125
    End If



   'Distal Elastic and Cutback
    elastictop.X(1) = Profile.X(1)
    elastictop.y(1) = Profile.y(1) / 2
    elastictop.y(2) = Profile.y(1) / 2
    elastictop.X(3) = Profile.X(1) + 0.375
    elastictop.y(3) = Profile.y(1) + 0.375
    PR_MakeXY xyEpt1, Profile.X(1) + 0.75, Profile.y(1) / 2

    If Val(txtFirstTape) = 11 And Val(txtAge) <= 10 Then
        elastictop.X(2) = Profile.X(1) + 0.375
        elastictop.X(3) = Profile.X(1) + 0.375
        PR_SetLayer "Template" & txtSleeve
        PR_DrawPoly elastictop
        PR_SetLayer "Notes"
        PR_DrawText "1/2'' Elastic", xyEpt1, 0.125

     ElseIf Val(txtFirstTape) = 11 And Val(txtAge) > 10 Then
        elastictop.X(2) = Profile.X(1) + 0.75
        elastictop.X(3) = Profile.X(1) + 0.75
        PR_SetLayer "Template" & txtSleeve
        PR_DrawPoly elastictop

     ElseIf Val(txtFirstTape) > 11 And Val(txtAge) <= 10 Then
       'For children with sleeves above elbow then use 1/2" elastic
        PR_SetLayer "Notes"
        PR_DrawText "1/2'' Elastic", xyEpt1, 0.125

    End If
    
    g_nCurrTextAngle = 0
    

End Sub

Private Function FN_CheckValue(TextBox As Control, sMessage As String) As Integer
' Check for numeric values
'Dim sChar As String, sText As String
'Dim iLen As Integer, nn As Integer
'sText = TextBox.Text
'iLen = Len(sText)
'FN_CheckValue = True
'For nn = 1 To iLen
'    sChar = Mid$(sText, nn, 1)
'    If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
'        MsgBox "Invalid " & sMessage & " Length has been entered", 48, "Arm Details"
'        TextBox.SetFocus
'        FN_CheckValue = False
'        Exit For
'    End If
'Next nn

    FN_CheckValue = True
    If Not IsNumeric(TextBox.Text) And Len(TextBox.Text) > 0 Then
        MsgBox "Invalid " & sMessage & " Length has been entered", 48, "Arm Details"
        TextBox.SetFocus
        FN_CheckValue = False
    End If

End Function


Private Function FN_PleatDistances(flag, cm)
Dim inch, remainder, temp, temprem, rnder
If flag = "cm" Then
    inch = Val(cm) / 2.54
    FN_PleatDistances = inch
ElseIf flag = "inches" Then
    temp = Int(cm)
    remainder = cm - temp
    If remainder > 0.7 Then
        MsgBox "Invalid Measurements, Inches and Eights only", 48, "Arm Details"
        FN_PleatDistances = "-1"
        Exit Function
    Else
        remainder = remainder * 10
        remainder = remainder * 0.125
        FN_PleatDistances = temp + remainder
    End If
 End If
End Function



Private Function fnInchesToText(nInches As Double) As String
   'Function returns a decimal value in inches as a string
   '
    Dim nPrecision As Double, nDec As Double
    Dim iInt As Integer, iEighths As Integer
    Dim sString As String
    nPrecision = 0.125

   'Split into decimal parts
    iInt = Int(nInches)
    nDec = nInches - iInt
    If nDec <> 0 Then           'Avoid overflow
        iEighths = Int(nDec / nPrecision)
    Else
        iEighths = 0
    End If

   'Format string
    If iInt <> 0 Then
        sString = LTrim$(Str$(iInt))
    Else
        sString = " 0"
    End If
    If iEighths <> 0 Then
        Select Case iEighths
        Case 2, 6
            sString = sString + "-" + LTrim$(Str$(iEighths / 2)) + "/4"
        Case 4
            sString = sString + "-" + "1/2"
        Case Else
            sString = sString + "-" + LTrim$(Str$(iEighths)) + "/8"
        End Select
    Else
        sString = sString + "   "
    End If
       
   'Return formatted string
    fnInchesToText = sString

End Function

Private Sub Form_LinkClose()
Dim i, X, sTmp, rtmp, def1, def2, cmp
Dim nValue As Double, nAge As Integer
Dim sFlap As String, sDiag  As String
   'Disable timer
   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

    If txtUidMPD = "" Then
        MsgBox "Can't find Patient Details", 16, "Arm Details Dialogue"
        End
    End If

    cboFabric.Text = txtFabric.Text
    g_sUnits = txtinchflag

    ' set form caption
    If txtSleeve = "Left" Then
        armdia.Caption = "ARM Details - Left"
    Else
        armdia.Caption = "ARM Details - Right"
    End If
    
   'Pressure
    cboPressure.Text = txtMM
   ' if pressure empty set from diagnosis
    nAge = Val(txtAge.Text)
    sDiag = UCase$(Mid$(txtDiagnosis.Text, 1, 6))
    If txtMM = "" Then
       'Arterial Insufficiency
        If sDiag = "ARTERI" Then
            cboPressure = "15mm"
       'Assist fluid dynamics
        ElseIf sDiag = "ASSIST" Then
            cboPressure = "20mm"
       'Blood Clots
        ElseIf sDiag = "BLOOD " Then
            cboPressure = "20mm"
       'Burns
        ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then
            cboPressure = "15mm"
       'Cancer
        ElseIf sDiag = "CANCER" Then
            cboPressure = "20mm"
       'Cardial Vascular Arrest
        ElseIf sDiag = "CARDIA" Then
            cboPressure = "15mm"
       'Carpal Tunnel Syndrome
        ElseIf sDiag = "CARPAL" Then
            cboPressure = "20mm"
       'Cellulitis
        ElseIf sDiag = "CELLUL" Then
            cboPressure = "20mm"
       'Chronic Venous Insufficiency
        ElseIf sDiag = "CHRONI" Then
            cboPressure = "20mm"
       'Heart Condition
        ElseIf sDiag = "HEART " Then
            cboPressure = "15mm"
       'Hemangioma
        ElseIf sDiag = "HEMANG" Then
            cboPressure = "15mm"
       'Lymphedema, Lymphedema 1+, Lymphedema 2+
        ElseIf sDiag = "LYMPHE" Then
            If InStr(txtDiagnosis.Text, "2") > 0 Then
                If nAge <= 70 Then
                    cboPressure = "25mm"
                Else
                    cboPressure = "20mm"
                End If
            Else
                If nAge <= 70 Then
                    cboPressure = "20mm"
                Else
                    cboPressure = "15mm"
                End If
            End If
       'Night Wear
        ElseIf sDiag = "NIGHT " Then
            cboPressure = "15mm"
       'Post Fracture
        ElseIf sDiag = "POST F" Then
            cboPressure = "20mm"
       'Post Mastectomy
        ElseIf sDiag = "POST M" Then
            cboPressure = "20mm"
       'Postphlebetic Syndrome
        ElseIf sDiag = "POSTPH" Then
            cboPressure = "20mm"
       'Renal Disease (Kidney)
        ElseIf sDiag = "RENAL " Then
            cboPressure = "20mm"
       'Skin Graft
        ElseIf sDiag = "SKIN G" Then
            cboPressure = "15mm"
       'Stroke
        ElseIf sDiag = "STROKE" Then
            cboPressure = "15mm"
       'Tendonitis
        ElseIf sDiag = "TENDON" Then
            cboPressure = "15mm"
       'Thrombophlebitis
        ElseIf sDiag = "THROMB" Then
            cboPressure = "20mm"
       'Trauma
        ElseIf sDiag = "TRAUMA" Then
            cboPressure = "20mm"
       'Varicose Veins
        ElseIf sDiag = "VARICO" Then
            cboPressure = "20mm"
       'Request for 30 m/m, 40 m/m and 50 m/m
        ElseIf sDiag = "REQUES" Then
            If InStr(txtDiagnosis.Text, "30") > 0 Then
                cboPressure = "15mm"
            ElseIf InStr(txtDiagnosis.Text, "40") > 0 Then
                cboPressure = "20mm"
            ElseIf InStr(txtDiagnosis.Text, "50") > 0 Then
                cboPressure = "25mm"
            End If
       'Vein Removal
        ElseIf sDiag = "VEIN R" Then
            cboPressure = "20mm"
        End If
    End If

   'Pleats, break up multiple fields
   'NB the order
    nValue = FN_GetNumber(txtWristPleat1.Text, 2)
    If nValue > 0 Then txtWristPleat2 = nValue Else txtWristPleat2 = ""
    nValue = FN_GetNumber(txtWristPleat1.Text, 1)
    If nValue > 0 Then txtWristPleat1 = nValue Else txtWristPleat1 = ""
   ' Strip_Bad_Char txtShoulderPleat1
   'NB the order
    nValue = FN_GetNumber(txtShoulderPleat1.Text, 2)
    If nValue > 0 Then txtShoulderPleat2 = nValue Else txtShoulderPleat2 = ""
    nValue = FN_GetNumber(txtShoulderPleat1.Text, 1)
    If nValue > 0 Then txtShoulderPleat1 = nValue Else txtShoulderPleat1 = ""
    
   'Stump
   ' Strip_Bad_Char txtStump
    If txtStump.Text = "True" Then
        chkStump.Value = 1
        chkStump_Click
    Else
        chkStump.Value = 0
    End If
    
    LowLength txtTapeLent
    LowWeight txtWeight
    LowRed txtReduction
    LowTapeMM txtTapeMM
    For i = 0 To 17
        k = i + 1
        ChangeChecker(k) = Length(i)
    Next i

    
   'Gauntlet
   'Break up gauntlet multiple field
    If FN_GetNumber(txtGauntlet.Text, 1) = 1 Then txtGauntletFlag = "True"
    If FN_GetNumber(txtGauntlet.Text, 1) = 0 Then txtGauntletFlag = "False"

    If txtGauntletFlag <> "True" Or txtGauntletFlag <> "False" Then
        g_sOld = "True"
    Else
        g_sOld = "False"
    End If
    
    If txtGauntletFlag = "True" Then
        
       'Flags
        If FN_GetNumber(txtGauntlet.Text, 2) Then txtEnclosed = "True"
        If FN_GetNumber(txtGauntlet.Text, 3) Then txtDetach = "True"
        If FN_GetNumber(txtGauntlet.Text, 4) Then txtNoThumb = "True"
       'Data
        txtWristNo = FN_GetNumber(txtGauntlet.Text, 5)
        txtPalmNo = FN_GetNumber(txtGauntlet.Text, 6)
        nValue = FN_GetNumber(txtGauntlet.Text, 7)
        If nValue > 0 Then txtThumLen = nValue
        txtTCircum = FN_GetNumber(txtGauntlet.Text, 8)
        txtPalmWristDist = FN_GetNumber(txtGauntlet.Text, 9)
        nValue = FN_GetNumber(txtGauntlet.Text, 10)
        If nValue > 0 Then txtGauntletExtension = nValue
       'Setup rest as per normal
        cboWrist.ListIndex = Val(txtWristNo) - 1
        cboPalm.ListIndex = Val(txtPalmNo) - 1
        txtCircum.Enabled = True
        lblCircum.Enabled = True
        OldGauntlet txtPalmNo, txtWristNo, txtPalmWristDist, txtTCircum, txtThumLen
        chkGauntlets.Value = True
        lblPalm.Enabled = True
        lblWrist.Enabled = True
        lblThumb.Enabled = True
        cboPalm.Enabled = True
        cboWrist.Enabled = True
        txtThumb.Enabled = True
        chkEnclosed.Enabled = True
        If txtEnclosed = "True" Then
            chkEnclosed.Value = 1
        End If
        chkDetachable.Enabled = True
        If txtDetach = "True" Then
            chkDetachable.Value = 1
            chkDetachable_Click
        End If
        chkNoThumb.Enabled = True
        If txtNoThumb = "True" Then
            chkNoThumb.Value = 1
            txtThumb = ""
            txtCircum = ""
            txtCircum.Enabled = False
            lblCircum.Enabled = False
        End If
        lblPalmWrist.Enabled = True
        txtPalmWrist.Enabled = True
    ElseIf txtStump.Text <> "True" Then
        chkNone.Value = 1
        chkNone_Click
    End If
    
   'Flap, break up flap multiple field
   'In this case the flap is 35 char long and the strap length is from char 36
    If txtType = "SLEEVE" Then cboFlaps.AddItem "Vest"
    If txtType = "SLEEVE-BODY" Then cboFlaps.AddItem "Body"

    If (txtFlap.Text = "") Or Mid(txtFlap.Text, 1, 4) = "None" Or txtType = "SLEEVE" Or txtType = "SLEEVE-BODY" Then
       'For new vests raglan set default to vest
        If txtType = "SLEEVE" Then
            cboFlaps.Text = "Vest"
            optProximalTape(1).Value = True
            optProximalTape_Click (1)
            cboFlaps_Click
        ElseIf txtType = "SLEEVE-BODY" Then
            cboFlaps.Text = "Body"
            optProximalTape(1).Value = True
            optProximalTape_Click (1)
            cboFlaps_Click
        Else
            txtFlap = ""
            optProximalTape(0).Value = True
            optProximalTape_Click (0)
        End If
    Else
        optProximalTape(1).Value = True
        optProximalTape_Click (1)
        sData = Mid(txtFlap, 36) 'NB order is very important
        txtFlap = Left(txtFlap, 35)
        cboFlaps.Text = txtFlap
        nValue = FN_GetNumber(sData, 1) 'Strap length
        If nValue > 0 Then
            txtStrapLength = nValue
            txtStrap = nValue
            labStrap = fnInchesToText(cmtoinches(txtinchflag, txtStrap))
        End If
        nValue = FN_GetNumber(sData, 2) 'Custom flap length
        If nValue > 0 Then
           'Enable custom flap
            txtCustFlapLength.Enabled = True
            labCustFlapLength.Enabled = True
            txtCustFlapLength = nValue
            labCustFlapLength = fnInchesToText(cmtoinches(txtinchflag, txtCustFlapLength))
        End If
        nValue = FN_GetNumber(sData, 3) 'Waist Circumference (Style-D only)
        If nValue > 0 Then
           'Enable custom flap
            txtWaistCir.Enabled = True
            lblWaistCir.Enabled = True
            labWaistCir.Enabled = True
            txtWaistCir = nValue
            g_sWaistCir = nValue
            labWaistCir = fnInchesToText(cmtoinches(txtinchflag, txtWaistCir))
        End If
        nValue = FN_GetNumber(sData, 4) 'Front Strap
        If nValue > 0 Then
           'Enable custom flap
            txtFrontStrapLength.Enabled = True
            lblFrontStrapLength.Enabled = True
            labFrontStrapLength.Enabled = True
            txtFrontStrapLength = nValue
            g_sFrontStrapLength = nValue
            labFrontStrapLength.Caption = fnInchesToText(cmtoinches(txtinchflag, txtFrontStrapLength))
        End If

        cboFlaps_Click
    End If

    EmptySleeve = True
    For w = 0 To 17
        X = w + 1
        If Length(w) <> "" Then
            EmptySleeve = False
            Measurements(X) = cmtoinches(txtinchflag, Length(w))
        End If
    Next w

    For i = 0 To 17
    If Val(Length(i)) > 0 Then
        Dim inchbit As Double
        inchbit = cmtoinches(txtinchflag, Length(i))
        If inchbit = -1 Then
            Length(i).SetFocus
            Exit Sub
        End If
        InchText(i) = fnInchesToText(inchbit)
    End If
    Next
     
     Dim pwinch As Double, circinch, thminch
     If Val(txtPalmWrist) > 0 Then
        pwinch = cmtoinches(txtinchflag, txtPalmWrist)
        PWDist = fnInchesToText(pwinch)
     End If

     If Val(txtCircum) > 0 Then
        pwinch = cmtoinches(txtinchflag, txtCircum)
        TCirc = fnInchesToText(pwinch)
     End If

     If Val(txtThumb) > 0 Then
        pwinch = cmtoinches(txtinchflag, txtThumb)
        TLen = fnInchesToText(pwinch)
     End If

     If Val(txtGauntletExtension) > 0 Then
        labGauntletExtension.Caption = fnInchesToText(cmtoinches(txtinchflag, txtGauntletExtension))
     End If


PR_GetStartEndTapes
For Index = 0 To 17
  If Length(Index) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And mms(Index) <> "" Then
    Dim Gram
    inchbit = cmtoinches(txtinchflag, Length(Index))
    If inchbit = -1 Then
        Length(Index).SetFocus
        Exit Sub
    End If
    InchText(Index) = fnInchesToText(inchbit)
    w = Index
    X = Index + 1
    PR_GetFabric
'    Gram = Fix(Val(mms(Index)) * Val(Measurements(Index + 1)))
'    gms(Index) = round(Gram)
'    reds(Index) = GetReduction(g_sModulus, Gram, g_sFabnam)
    PR_WristRed w, X, g_sModulus
    PR_PalmRed w, X, g_sModulus
    PR_LastTapeRed w, X, g_sModulus
    PR_SetRresults
    PR_SetGlobals
  End If
Next Index

   'Set up first and last tapes depending on ID
    nValue = Val(Mid$(txtID, 1, 2))
    If nValue > 0 Then
        cboDistalTape.ListIndex = nValue
    Else
        cboDistalTape.ListIndex = 0
    End If
    cboDistalTape_Click
    
    nValue = Val(Mid$(txtID, 3, 2))
    If nValue > 0 Then
        cboProximalTape.ListIndex = (18 - nValue) + 1
    ElseIf Mid$(txtID, 3, 2) <> "GT" Then 'IE Not for Detachable Gauntlet
        cboProximalTape.ListIndex = 0
    End If
    cboProximalTape_Click

Show
Screen.MousePointer = 0
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = "Cancel" Then
        Cancel = 0
        End
    End If
End Sub

Private Sub Form_Load()

Hide
' Maintain while loading DDE data
Screen.MousePointer = 11  ' change pointer to hourglass
' Reset in form_linkclose

' Position to center of screen
Left = (Screen.Width - armdia.Width) / 2
Top = (Screen.Height - armdia.Height) / 2

Set MainForm = armdia

Dim i
i = 0
Do While i < 18
    reds(i) = ""
    gms(i) = ""
    mms(i) = ""
    Length(i) = ""
    i = i + 1
Loop
txtVestRaglan = ""
txtVestID = ""
txtPatientName.Text = ""
txtFileNo.Text = ""
txtinchflag.Text = ""
cboFabric.Text = ""
cboPressure.Text = ""
txtDiagnosis.Text = ""
cboFlaps.Text = ""
txtStrapLength.Text = ""
g_sFlapLength = ""
g_sFlapChk = ""
txtTapeLent.Text = ""
g_sPressureChange = "False"
g_sFabricChange = "False"
txtWeight.Text = ""
txtTapeMM.Text = ""
txtReduction.Text = ""
txtThumLen.Text = ""
txtThumbDist.Text = ""
txtTCircum.Text = ""
txtDetach.Text = ""
txtPalmNo.Text = ""
txtWristNo.Text = ""
txtNoThumb.Text = ""
txtPalmWristDist.Text = ""
txtStump.Text = ""
txtGauntletExtension.Text = ""
txtFrontStrapLength.Text = ""
txtFlap.Text = ""
txtCustFlapLength.Text = ""
txtFabric.Text = ""
txtSleeve.Text = ""
txtMM.Text = ""
txtType.Text = ""
txtFirstTape.Text = ""
txtSecondTape.Text = ""
txtSecondLastTape.Text = ""
txtLastTape.Text = ""
txtSex.Text = ""
txtAge.Text = ""
txtModulus.Text = ""
txtWristPleat1.Text = ""
txtWristPleat2.Text = ""
txtShoulderPleat1.Text = ""
txtShoulderPleat2.Text = ""
chkNone.Value = True
chkStump.Value = False
frmGauntlet.Enabled = 0
chkGauntlets.Value = False
chkDetachable.Value = 0
chkEnclosed.Value = 0
chkNoThumb.Value = 0
txtPalmWrist.Text = ""
txtCircum.Text = ""
txtThumb.Text = ""
txtWaistCir = ""
i = 0
txtUidMPD = ""
txtUidAC = ""
txtACwo = ""
txtUidArm = ""

'initalise globals
g_sFileNo = ""
g_sSide = ""
g_sPatient = ""
g_sFabric = ""
g_sFabnam = ""
g_sWpleat1 = ""
g_sWpleat2 = ""
g_sSpleat1 = ""
g_sSpleat2 = ""
g_sFlapStrap = ""
g_sFlap = ""
g_sMM = ""
g_sGaunt = ""
g_sDetGaunt = ""
g_sNoThumb = ""
g_sPalmNo = ""
g_sWristNo = ""
g_sPalmWristDist = ""
g_sThumbCircum = ""
g_sThumbLength = ""
g_sEnclosedThumb = ""
g_sSecondLastTape = ""
g_sSecondTape = ""
g_sFirstTape = ""
g_sLastTape = ""
g_sModulus = ""
g_sStump = ""
g_sTapeLengths = ""
g_sTapeMMs = ""
g_sGrams = ""
g_sFlapLength = ""
g_sReduction = ""
g_sAmm = ""
g_sBmm = ""
g_sCmm = ""
g_sDmm = ""

g_iStyleLastTape = -1
g_iStyleFirstTape = -1
g_sWaistCir = ""

g_sPathJOBST = fnPathJOBST()

Dim textline
    fileNum = FreeFile
    Open g_sPathJOBST & "\Defaults\Pressure.dat" For Input As fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textline
        cboPressure.AddItem textline
    Loop
    Close

    
    fileNum = FreeFile
    Open g_sPathJOBST & "\fabric.dat" For Input As fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textline
        cboFabric.AddItem textline
    Loop
    Close

    fileNum = FreeFile
    Open g_sPathJOBST & "\Defaults\flaps.dat" For Input As fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textline
        cboFlaps.AddItem textline
    Loop
    Close

'
'Load Palm Tapes and Wrist
'
    g_sTextList = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

    For ii = 0 To 9
        cboWrist.AddItem LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3))
        cboPalm.AddItem LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3))
    Next ii

    cboDistalTape.AddItem "1st"
    For ii = 0 To 17
        cboDistalTape.AddItem LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3))
    Next ii

    cboProximalTape.AddItem "Last"
    For ii = 17 To 0 Step -1
        cboProximalTape.AddItem LTrim$(Mid$(g_sTextList, (ii * 3) + 1, 3))
    Next ii

'   ------------------------------------------------
'   Routine to read in the Bobinnet conversion Chart
'   ------------------------------------------------
    Dim j, offset, Modulus  As String, Reduction As String
    'Static bob(1 To 19, 1 To 23) As Variant
    ReDim indi(1 To 19) As Variant
    fileNum = FreeFile
    Open g_sPathJOBST & "\templts\bobinnet.dat" For Input As fileNum
    For i = 1 To 19
        Input #fileNum, Modulus, Reduction
        indi(i) = Modulus
        offset = 1
        For j = 1 To 23
            Bob(i, j) = Mid$(Reduction, offset, 4)
            offset = offset + 4
            Next j
    Next i
    Close
'   ------------------------------------------------
'            End Of Bobinnet Routine
'   ------------------------------------------------

'   ------------------------------------------------
'   Routine to read in the Powernet conversion Chart
'   ------------------------------------------------
    Dim P, q, offset2, modulus2 As String, Reduction2 As String
'    ReDim pow(1 To 19, 1 To 23) As Variant
    ReDim pindi(1 To 19) As Variant
    fileNum = FreeFile
    Open g_sPathJOBST & "\templts\powernet.dat" For Input As fileNum
    For P = 1 To 19
        Input #fileNum, modulus2, Reduction2
        pindi(P) = modulus2
        offset2 = 1
        For q = 1 To 23
            Pow(P, q) = Mid$(Reduction2, offset2, 4)
            offset2 = offset2 + 4
           Next q
    Next P
    Close
'   ------------------------------------------------
'            End Of Powernet Routine
'   ------------------------------------------------
 
   'Start a timer
   'The Timer is disabled in LinkClose
   'If after 10 seconds the timer event will "End" the programme
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 10000  'Approx 10 Seconds
    Timer1.Enabled = True


End Sub

Private Sub Gauntlet(NoThumb, Palm, Wrist, PalmWrist, thumlength, Sex, age, Enclosed, Circum, Detach, txtinchflag)
        
        txtGauntletFlag = "True"
        tempwrist = Wrist
        temppalm = Palm
        thumdist = PalmWrist
        If thumdist = "" Then thumdist = 0
        thumdist = cmtoinches(txtinchflag, thumdist)
       
       'No Thumb
        If chkNoThumb.Value = 1 Then
            txtNoThumb = "True"
        Else
            txtNoThumb = "False"
        End If
        
       ' Check values entered
        If Palm = "" Then
            MsgBox "Palm tape number has not been entered.", 48, "Gauntlet Details"
            Exit Sub
        End If
        If Wrist = "" Then
            MsgBox "Wrist tape number has not been entered.", 48, "Gauntlet Details"
            Exit Sub
        End If
        If PalmWrist = "" Then
            MsgBox "Palm to Wrist Distance has not been entered.", 48, "Gauntlet Details"
            Exit Sub
        End If
        ' Thumbhole Dist
        If thumdist > 2.5 Then
            txtThumbDist = 2.5
            txtPalmWristDist = thumdist
        Else
            txtThumbDist = thumdist
            txtPalmWristDist = thumdist
        End If
        ' End Thumbhole Dist
        
        'Thumb length
        If Val(thumlength) > 0 Then
            thum = thumlength
            thum = cmtoinches(txtinchflag, thum)
            txtThumLen = thum
        Else: txtThumLen = 0
            thum = 0
        End If
        
        If thum = 0 And Val(age) > 14 Then
            thum = 0.75
            txtThumLen = thum
        ElseIf txtThumLen = 0 And Val(age) <= 14 Then
            thum = 0.5
            txtThumLen = thum
        End If
        'End Thumb length

        'Enclosed Thumb
        If Enclosed = 1 And Val(age) > 14 Then
            If thumlength = "" Then
                MsgBox "Thumb Length has not been entered.", 48, "Gauntlet Details"
                Exit Sub
            End If
            txtEnclosed = "True"
            txtThumblen = Val(thumlength) - 0.25
        ElseIf Enclosed = 1 And Val(age) < 15 Then
            If thumlength = 0 Then
                MsgBox "Thumb Length has not been entered.", 48, "Gauntlet Details"
                Exit Sub
            End If
            txtEnclosed = "True"
            txtThumblen = Val(thumlength) - 0.125
        ElseIf Enclosed = 0 Then
            txtEnclosed = "False"
        End If
        'End Enclosed Thumb

        ' Thumb Circumference
        If NoThumb = 0 And Val(Circum) = 0 Then
            MsgBox "Thumb Circumference has not been entered.", 48, "Gauntlet Details"
            Exit Sub
        ElseIf NoThumb = 0 And Val(Circum) > 0 Then
            tempCircum = cmtoinches(txtinchflag, Circum)
            tempCircum = Val(tempCircum) / 2
            txtTCircum = tempCircum
        End If
        ' End Thumb Circumference
 
        ' Detachable
        If Detach = 1 Then
            txtDetach = "True"
        Else
            txtDetach = "False"
        End If
        ' End Detachable
 
        If Val(Palm) > Val(Wrist) Then
            MsgBox "Palm Tape No Should be less than Wrist Tape No. Enter new values and re-calculate. ", 48, "Gauntlet Details"
            Exit Sub
        End If
        ' check palm and wrist tapes
        
End Sub

Private Function GetReduction(Modul, Gram, fabnam)
Dim tempred
        If Left$(fabnam, 3) = "Pow" Then
            tempred = PowReduction(Modul, Gram)
        ElseIf Left$(fabnam, 3) = "Bob" Then
            tempred = BobReduction(Modul, Gram)
        End If
        GetReduction = tempred
End Function

Private Sub lblPleat_DblClick(Index As Integer)
   'Procedure to disable the pleat fields
   'to enable the User to select only the relevent
   'data to be drawn
   'Saves having to delele and re-enter to get variations
   'on pleats
    If lblPleat(Index).ForeColor = &HC0C0C0 Then
       'Enable pleat if previously disabled
        lblPleat(Index).ForeColor = &H0&
        Select Case Index
         Case 0
            txtWristPleat1.Enabled = True
         Case 1
            txtWristPleat2.Enabled = True
         Case 3
            txtShoulderPleat1.Enabled = True
         Case 2
            txtShoulderPleat2.Enabled = True
        End Select
    Else
       'Disable pleat if previously enabled
        lblPleat(Index).ForeColor = &HC0C0C0
        Select Case Index
         Case 0
            txtWristPleat1.Enabled = False
         Case 1
            txtWristPleat2.Enabled = False
         Case 3
            txtShoulderPleat1.Enabled = False
         Case 2
            txtShoulderPleat2.Enabled = False
        End Select
    End If

End Sub

Private Sub Length_GotFocus(Index As Integer)
    select_text_in_Box1 Length(Index)
End Sub

Private Sub Length_LostFocus(Index As Integer)

If (ChangeChecker(Index + 1) <> Length(Index)) Then
 
If armdia.Visible = True Then g_Modified = True

' Delete tape values
If Length(Index) = "" And mms(Index) <> "" And gms(Index) <> "" And reds(Index) <> "" And InchText(Index) <> "" Then
     mms(Index) = ""
     gms(Index) = ""
     reds(Index) = ""
     InchText(Index) = ""
     PR_GetStartEndTapes
    'Reset the first and last tape indicators
     If cboProximalTape.ListIndex = 0 Then cboProximalTape_Click
     If cboDistalTape.ListIndex = 0 Then cboDistalTape_Click

End If


' Check for numeric values
    If Not FN_CheckValue(Length(Index), "Tape") Then Exit Sub

 
' check for too large a number
If (Val(Length(Index)) > 99 And txtinchflag = "cm") Or (Val(Length(Index)) > 50 And txtinchflag = "inches") Then
    MsgBox "Tape length entered is too Large", 48, "Arm Details"
    Length(Index).SetFocus
    Exit Sub
End If

' check for too small a number
' tapelen from above
If Val(Length(Index)) < 1 And Len(Length(Index)) > 0 Then
    MsgBox "Tape length entered is too Small", 48, "Arm Details"
    Length(Index).SetFocus
    Exit Sub
End If


If Val(Length(Index)) > 0 Then
    Dim inchbit As Double
    inchbit = cmtoinches(txtinchflag, Length(Index))
    If inchbit = -1 Then
        Length(Index).SetFocus
        Exit Sub
    End If
    InchText(Index) = fnInchesToText(inchbit)
    Measurements(Index + 1) = cmtoinches(txtinchflag, Length(Index))
   
   'Exit if no fabric or no pressure has been selected
   'We also supress calculations if this is an empty sleeve
   'the flag EmptySleeve is set in Form "LinkClose" & Command "Calculate"
   'I.E. don't do any calculations
    If cboFabric.Text = "" Or cboPressure.Text = "" Or EmptySleeve = True Then Exit Sub

   'Recalculate etc
    PR_GetPressure
    PR_GetFabric
    PR_CheckValues
    PR_GetGauntletDetails
    PR_GetStartEndTapes
    PR_CalculateMMSGmsReds Index
    PR_ReSetMMs Index, g_sModulus
    PR_SetRresults
    PR_SetGlobals


'check for extra measurements in Gauntlet
PR_GauntletHoleCheck


'If a tape has been added then reset the tape indicators
 If cboProximalTape.ListIndex = 0 Then cboProximalTape_Click
 If cboDistalTape.ListIndex = 0 Then cboDistalTape_Click

End If
   'Reset ChangeChecker to revised value GG.17.Mar.95
    ChangeChecker(Index + 1) = Length(Index)
End If
End Sub

Private Sub LowLength(tapes As Variant)
Dim i, nValue
For i = 0 To 17
    nValue = Val(Mid(tapes, (i * 3) + 1, 3)) / 10
    If nValue > 0 Then
        Length(i) = nValue
    End If
Next i
End Sub

Private Sub LowRed(tapes)
Dim i, nValue
For i = 0 To 17
    nValue = Val(Mid(tapes, (i * 3) + 1, 3))
    If nValue > 0 Then
        reds(i) = nValue
    End If
Next i
End Sub

Private Sub LowTapeMM(tapes)
Dim i, nValue
For i = 0 To 17
    nValue = Val(Mid(tapes, (i * 3) + 1, 3))
    If nValue > 0 Then
        mms(i) = nValue
    End If
Next i
End Sub

Private Sub LowWeight(tapes)
Dim i, nValue
For i = 0 To 17
    nValue = Val(Mid(tapes, (i * 3) + 1, 3))
    If nValue > 0 Then
        gms(i) = nValue
    End If
Next i
End Sub

Private Function MissingVal(Before, Present, after)
'   *** Function to Check For Missing Measurements
    If Present = "" And Before <> "" And after <> "" Then
        MissingVal = 1
    Else
        MissingVal = 0
    End If
End Function

Private Sub mms_GotFocus(Index As Integer)
    select_text_in_Box1 mms(Index)
    MMsChecker(Index + 1) = mms(Index)
End Sub

Private Sub mms_LostFocus(Index As Integer)

If mms(Index) <> "" And Length(Index) = "" And gms(Index) = "" Then
    mms(Index) = ""
End If

If MMsChecker(Index + 1) <> mms(Index) Then

If armdia.Visible = True Then g_Modified = True

If Length(Index) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(Index) <> "" Then
    Dim inchbit As Double
    Dim w, X, Gram
    
    inchbit = cmtoinches(txtinchflag, Length(Index))
    If inchbit = -1 Then
        Length(Index).SetFocus
        Exit Sub
    End If
    InchText(Index) = fnInchesToText(inchbit)
    PR_GetFabric
    PR_GetStartEndTapes
        w = Index
        X = Index + 1
            Gram = Fix(Val(mms(Index)) * Val(Measurements(Index + 1)))
            gms(Index) = round(Gram)
            reds(Index) = GetReduction(g_sModulus, Gram, g_sFabnam)
            
            'Check for Grams Being Too Big
            If Val(reds(Index)) = 0 And gms(Index) > 205 Then
                reds(Index) = 32
                If g_sFabnam = "Pow" Then
                    Backgram = Pow(txtModulus, 23)
                ElseIf g_sFabnam = "Bob" Then
                    Backgram = Bob(txtModulus, 23)
                End If
                gms(Index) = Backgram
                mms(Index) = round(Backgram / Val(Measurements(Index + 1)))
            End If
        
            'Check for grams being too little
            If Val(reds(Index)) <= 10 Then
                reds(Index) = 10
                If g_sFabnam = "Pow" Then
                    Backgram = Pow(txtModulus, 1)
                ElseIf g_sFabnam = "Bob" Then
                    Backgram = Bob(txtModulus, 1)
                End If
                gms(Index) = Backgram
                mms(Index) = round(Backgram / Val(Measurements(Index + 1)))
            End If
        
        PR_WristRed w, X, g_sModulus
        PR_PalmRed w, X, g_sModulus
        PR_LastTapeRed w, X, g_sModulus
        PR_Stump w, X
        PR_CheckFlaps w
        PR_ReSetMMs Index, g_sModulus
    PR_SetRresults
    PR_SetGlobals
End If
End If
End Sub

Private Function NoVals()
'   *** Function to check if no values have been entered ***
Dim i, contents, total
    i = 0
    contents = 0
    total = 0
    Do While i < 18
        contents = Val(Length(i))
        total = total + contents
        i = i + 1
    Loop
    NoVals = total
'    *** End Check values have been entered ***
End Function

Private Sub OldGauntlet(txtPalmNo As Variant, txtWristNo As Variant, txtPalmWristDist As Variant, txtTCircum As Variant, txtThumLen As Variant)
Dim differ
    differ = Val(txtWristNo) - Val(txtPalmNo)
    If differ > 1 Then
        g_sHoleCheck = "True"
    ElseIf differ = 1 Then
        g_sHoleCheck = "False"
    End If
    txtPalm = (Val(txtPalmNo) * 1.5) + (-7.5)
    txtWrist = (Val(txtWristNo) * 1.5) + (-7.5)
    txtPalmWrist = txtPalmWristDist
    txtCircum = txtTCircum
    txtThumb = txtThumLen
End Sub

Private Sub optProximalTape_Click(Index As Integer)
'only two choices
'
    If Index = 0 Then
       'Disable flaps
        frmFlaps.Enabled = False
        cboFlaps.Text = "none"
        cboFlaps_Click          'Force change of pressure
        cboFlaps.Enabled = False

        lblStyle.Enabled = False
        txtStrapLength.Text = ""
        txtStrapLength.Enabled = False
        labStrap.Caption = ""
        lblStrapLength.Enabled = False

        txtFrontStrapLength.Text = ""
        txtFrontStrapLength.Enabled = False
        labFrontStrapLength.Caption = ""
        lblFrontStrapLength.Enabled = False

        txtCustFlapLength.Text = ""
        txtCustFlapLength.Enabled = False
        labCustFlapLength.Caption = ""
        lblCustFlapLength.Enabled = False

        txtWaistCir.Text = ""
        txtWaistCir.Enabled = False
        labWaistCir.Caption = ""
        lblWaistCir.Enabled = False

       'Enable proximal tape position
        cboProximalTape.Enabled = True

    Else
       'Reset proximal tape position to last
        cboProximalTape.ListIndex = 0
        cboProximalTape_Click
        cboProximalTape.Enabled = False

       'Enable flaps
        frmFlaps.Enabled = True
        frmFlaps.Enabled = True
        cboFlaps.Enabled = True
        lblStyle.Enabled = True

        txtStrapLength.Enabled = True
        lblStrapLength.Enabled = True

        txtCustFlapLength.Enabled = True
        lblCustFlapLength.Enabled = True

        txtFrontStrapLength.Enabled = True
        lblFrontStrapLength.Enabled = True

       'Can't have a detachable gauntlet with a flap
        If chkDetachable.Value = 1 Then chkDetachable.Value = 0

'        txtWaistCir.Enabled = True
'        lblWaistCir.Enabled = True
    End If

    If armdia.Visible = True Then g_Modified = True

End Sub

Private Function PadGram(Gram)
If Gram <> "" Then
    If Val(Gram) < 10 Then
        PadGram = "00" & Gram
    ElseIf Val(Gram) < 100 And Val(Gram) > 9 Then
        PadGram = "0" & Gram
    ElseIf Val(Gram) > 99 Then
        PadGram = Gram
    End If
ElseIf Gram = "" Then
    PadGram = "000"
End If

End Function

Private Function PadMM(mm)
If mm <> "" Then
    If Len(mm) = 1 Then
        PadMM = "00" & mm
    ElseIf Len(mm) = 2 Then
        PadMM = "0" & mm
    ElseIf Len(mm) = 3 Then
        PadMM = mm
    End If
ElseIf mm = "" Then
    PadMM = "000"
End If
End Function

Private Function PadRed(red)
If red <> "" Then
   If Len(red) = 1 Then
        PadRed = "00" & red
   ElseIf Len(red) = 2 Then
        PadRed = "0" & red
   ElseIf Len(red) = 3 Then
        PadRed = red
   End If
ElseIf red = "" Then
    PadRed = "000"
End If
End Function

Private Function PadTape(measure As Variant)
Dim pos, dig, bit, Result
If Len(measure) = 4 Then
    pos = InStr(measure, ".")
    If pos <> 0 Then
        dig = Int(measure)
        bit = Right$(measure, 1)
        Result = dig & bit
    End If
         PadTape = Result
ElseIf Len(measure) = 3 Then
    pos = InStr(measure, ".")
    If pos <> 0 Then
        dig = Int(measure)
        bit = Right$(measure, 1)
        Result = "0" & dig & bit
    End If
    PadTape = Result
ElseIf Len(measure) = 2 Then
    Result = measure
    PadTape = Result & "0"
ElseIf Len(measure) = 1 Then
    Result = measure
    PadTape = "0" & Result & "0"
ElseIf measure = "" Then
    PadTape = "000"
End If
End Function

Private Function PowReduction(Modul, lblGram)
one = 1
two = 2
nine = 9
eight = 8
ten = 10
j = one
    If (Val(Pow(Modul, 1)) = Val(lblGram)) Then
        lblred = ten
    End If
    j = two
    Do While j <= 22
        If (Val(Pow(Modul, j)) = Val(lblGram)) Then
            lblred = Val(j + nine)
        ElseIf (Val(lblGram) - Val(Pow(Modul, j - one))) <= (Val(Pow(Modul, j)) - Val(lblGram)) Then
            lblred = Val(j + eight)
            Exit Do
        End If
            j = j + one
    Loop
    PowReduction = lblred
End Function

Private Sub PR_AddMarker(xyInsert As XY)
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC
    Print #fNum, "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.y); CC
    Print #fNum, 0.1; CC & 0.1; CC; 0; ");"
End Sub

Private Sub PR_CalculateMMSGmsReds(Index As Integer)
Dim w, X
    w = 0
    X = 1
    Do While X < 19
    Do While w < 18

        PR_InitaliseEmpty w              'Zero any empty fields
        PR_InitaliseArrays w, X
        If (Length(Index) <> "" And mms(Index) = "") Or (Length(Index) <> "" And g_sPressureChange = "True") Then
            PR_SetMMs w, g_sAmm, g_sBmm, g_sCmm, g_sDmm
        End If
        PR_Reductions w, X
        PR_WristRed w, X, g_sModulus
        PR_PalmRed w, X, g_sModulus
        PR_Stump w, X
        PR_LastCircum w, X
        PR_LastTapeRed w, X, g_sModulus
        PR_CheckFlaps w
        PR_Detachable w, X, g_sModulus, g_sFabnam
        w = w + 1
        X = X + 1
    Loop
    Loop

End Sub

Private Sub PR_CheckFlaps(w As Variant)
If Length(w) <> "" Then
' Check Flaps
    txtFlap = cboFlaps.Text
    g_sFlap = txtFlap
    If txtFlap <> "" And txtStrapLength = "" Then
        If txtinchflag = "cm" Then
            txtStrap = "61"
        Else
            txtStrap = "24"
        End If
        'labStrap = fnInchesToText(cmtoinches(txtinchflag, txtStrap))
    End If
End If
' End Check Flaps
End Sub

Private Sub PR_CheckValues()
''   *** Check values have been entered ***
Dim i, contents, total
    i = 0
    contents = 0
    total = 0
    Do While i < 18
        contents = Val(Length(i))
        total = total + contents
        i = i + 1
    Loop
    If total = 0 Then
        MsgBox "No tape measurements have been entered!", 48, "Arm Details"
        Exit Sub
    End If
'   *** End Check values have been entered ***

End Sub

Private Sub PR_Detachable(w As Variant, X As Variant, Modul As Variant, fabnam As Variant)
Dim Backgram
If Length(w) <> "" Then
    'CheckDetachable Gauntlet
    If chkDetachable.Value = 1 Then
        If X = Val(txtWristNo) And X = Val(txtSecondLastTape) Then
            reds(w) = 10
            If Left$(fabnam, 3) = "Pow" Then
                Backgram = Pow(Modul, 1)
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
            Else
                Backgram = Bob(Modul, 1)
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
            End If
        ElseIf X = Val(txtWristNo) + 1 And X = Val(txtLastTape) Then
           'If length of wrist is same as wrist+1 then assume detachable gauntlet
           'originally based on 2 tapes
            If Length(w) = Length(w - 1) Then
                reds(w) = 5
                gms(w) = ""
                mms(w) = ""
            Else
                reds(w) = 10
                If Left$(fabnam, 3) = "Pow" Then
                    Backgram = Pow(Modul, 1)
                    If Val(gms(w)) <> Backgram Then
                        gms(w) = round(Backgram)
                        mms(w) = round(Backgram / Measurements(X))
                    End If
                Else
                    Backgram = Bob(Modul, 1)
                    If Val(gms(w)) <> Backgram Then
                        gms(w) = round(Backgram)
                        mms(w) = round(Backgram / Measurements(X))
                    End If
                End If
            End If
        ElseIf X = Val(txtWristNo) And X = Val(txtLastTape) Then
            If Length(w + 1) = "" Then
                Length(w + 1) = Length(w)
                Measurements(X + 1) = Measurements(X)
                InchText(w + 1) = fnInchesToText(cmtoinches(txtinchflag, Length(w)))
                reds(w) = 10
                If Left$(fabnam, 3) = "Pow" Then
                    Backgram = Pow(Modul, 1)
                    If Val(gms(w)) <> Backgram Then
                        gms(w) = round(Backgram)
                        mms(w) = round(Backgram / Measurements(X))
                    End If
                Else
                    Backgram = Bob(Modul, 1)
                    If Val(gms(w)) <> Backgram Then
                        gms(w) = round(Backgram)
                        mms(w) = round(Backgram / Measurements(X))
                    End If
                End If
            Else
                reds(w) = 5
                gms(w) = ""
                mms(w) = ""
            End If
        End If
    End If
End If


End Sub

Private Sub PR_DrawShoulderFlaps(Profile As curve, numtapes As Variant)
Dim ShoulderFlap As curve, RaglanFlap As curve
Dim LastTapeValue, xyPt1 As XY, xyPt2 As XY
Dim xyTopFlap1 As XY, xyTopFlap2 As XY, xyTopFlap3 As XY, xyTopFlap4 As XY, xyTopFlap5 As XY
Dim xyTopFlap6 As XY, xyTopFlap7 As XY, xyTopFlap8 As XY, xyTopFlap9 As XY, xyTopFlap10 As XY
Dim xyBotFlap1 As XY, xyBotFlap2 As XY, xyBotFlap3 As XY, xyBotFlap4 As XY, xyBotFlap5 As XY
Dim xyFlap1 As XY, xyFlap2 As XY, xyFlap3 As XY, xyFlap4 As XY, xyFlap5 As XY, xyPotA As XY, xyPotB As XY
Dim Beta, Omega, Phi, Delta, Theta, BottomRadius, PI, Alpha, TopArcIncrement
Dim x1, y1, h1, adj, opp, hyp, LittleBit, Change, TopRadius, FlapMarkerAngle, xyFlapText1  As XY
Dim xyMid As XY, xyCentre As XY, FlapLength, FlapMarkerLength, xyFlapMarker As XY, xyFlapText2 As XY, xyStrapText As XY
Dim ArcLength, CircleCircum, TopArcLength, BottomArcLength, TopArcRadius, BottomArcRadius, TopArcAngle, BottomArcAngle
Dim TempMarker, Arrow As String, xfabby, xfablen, xfabdist
Dim xyRaglan1 As XY, xyRaglan2 As XY, xyRaglan3 As XY, xyRaglan4 As XY, xyRaglan5 As XY, xyRaglan6 As XY, xyRaglan7 As XY
Dim Rag1, Rag2, Rag3, Rag4, RagAng1, RagAng2, RagAng3, RaglanBottom, RaglanTip, TopRagLanAngle, InnerRaglanTip
Dim nNotchOffset As Double
BottomArcAngle = 47.5
Arrow = "closed arrow"
PI = 4 * Atn(1)
'If txtFlap <> "" And txtFlap <> "Vest Raglan" Then
If txtFlap <> "" And txtFlap <> "Vest" And txtFlap <> "Body" Then
    ' Check for Custom Flap Length
    If Val(txtCustFlapLength) > 0 Then
        g_sinchflag = txtinchflag
        g_sFlapLength = cmtoinches(txtinchflag, txtCustFlapLength)
        LastTapeValue = g_sFlapLength
    Else
       'Standard flap length
        LastTapeValue = Val(txtLastTape) - 1
        LastTapeValue = cmtoinches(txtinchflag, Length(LastTapeValue))
        LastTapeValue = (LastTapeValue / 3.14) * 0.92
    End If

    
    ' Calculate Bottom part of curve with 5 Points
    If (Profile.y(Profile.n)) >= 5.5 Then
        BottomRadius = LastTapeValue - 0.625
        RaglanBottom = 3.5
        'TopRagLanAngle = 20
        TopRagLanAngle = 17
        RaglanTip = 1.9375
        InnerRaglanTip = 2.173
        nNotchOffset = 0.875   '14/16 ths Raglan Template s287
    ElseIf ((Profile.y(Profile.n)) < 5.5) And ((Profile.y(Profile.n)) > 3) Then
        BottomRadius = 2.8
        RaglanBottom = 2.5
        'TopRagLanAngle = 20
        TopRagLanAngle = 17
        RaglanTip = 1.05
        InnerRaglanTip = 1.22
        nNotchOffset = 0.75     '12/16 ths Raglan Template s289
    ElseIf ((Profile.y(Profile.n)) <= 3) Then
        BottomRadius = 2
        RaglanBottom = 2.0625
        RaglanTip = 0.625
        'TopRagLanAngle = 15
        TopRagLanAngle = 15
        InnerRaglanTip = 0.727
        nNotchOffset = 0.6875   '11/16 ths, Raglan Template s290
    End If
    nTemplateRadius = 3.5
    nTemplateAngle = 40
    nNotchToTangent = 2.125

    'Calc length of bottom Arc
    BottomArcRadius = BottomRadius
    CircleCircum = PI * (2 * (BottomArcRadius))
    BottomArcLength = (BottomArcAngle / 360) * CircleCircum
    
    ShoulderFlap.X(1) = Profile.X(Profile.n) + LastTapeValue
    ShoulderFlap.y(1) = 0

    For ii = 1 To 4
        Theta = (ii * 11.875 * PI) / 180
        adj = Cos(Theta) * BottomRadius
        opp = Sin(Theta) * BottomRadius
        LittleBit = BottomRadius - adj
        ShoulderFlap.X(ii + 1) = Profile.X(Profile.n) + LastTapeValue - LittleBit
        ShoulderFlap.y(ii + 1) = 0 + opp
    Next ii
    PR_MakeXY xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.y(5)
    
    'Calculate 6 angles and points for top part of curve
    PR_MakeXY xyTopFlap1, Profile.X(Profile.n), Profile.y(Profile.n)
    ShoulderFlap.X(10) = xyTopFlap1.X
    ShoulderFlap.y(10) = xyTopFlap1.y

    'midpoint
    h1 = FN_CalcLength(xyTopFlap1, xyBotFlap5)
    h1 = h1 / 2
    Omega = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Omega = 180 - Omega
    Omega = (Omega * PI) / 180
    opp = Abs(h1 * Sin(Omega))
    adj = Abs(h1 * Cos(Omega))
    PR_MakeXY xyMid, xyBotFlap5.X - adj, xyBotFlap5.y + opp
   
    'Centre of Top Arc
    Delta = FN_CalcAngle(xyBotFlap5, xyTopFlap1)
    Delta = Delta - 90
    Delta = (Delta * PI) / 180
    hyp = 4 * h1
    opp = Abs(hyp * Sin(Delta))
    adj = Abs(hyp * Cos(Delta))
    PR_MakeXY xyCentre, xyMid.X + adj, xyMid.y + opp
    TopRadius = FN_CalcLength(xyCentre, xyTopFlap1)
    TopArcRadius = TopRadius
    Phi = FN_CalcAngle(xyCentre, xyTopFlap1)
    Phi = Phi - 180
    Alpha = FN_CalcAngle(xyCentre, xyBotFlap5)
    Alpha = Alpha - 180
    TopArcAngle = Alpha - Phi
    TopArcIncrement = (Alpha - Phi) / 5

   'Calc Flap Length for position of marker
    CircleCircum = PI * (2 * (TopArcRadius))
    TopArcLength = (TopArcAngle / 360) * CircleCircum
    FlapLength = TopArcLength + BottomArcLength
    FlapMarkerLength = FlapLength / 2.7
   'Minimum length from market to seam is 2.5"
    If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
    If FlapMarkerLength < TopArcLength Then
        FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
        FlapMarkerAngle = FlapMarkerAngle + Phi
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(TopRadius * Sin(FlapMarkerAngle))
        adj = Abs(TopRadius * Cos(FlapMarkerAngle))
        PR_MakeXY xyFlapMarker, xyCentre.X - adj, xyCentre.y - opp
    ElseIf FlapMarkerLength > TopArcLength Then
        TempMarker = FlapMarkerLength - TopArcLength
        TempMarker = BottomArcLength - TempMarker
        FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
        FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
        opp = Abs(BottomArcRadius * Sin(FlapMarkerAngle))
        adj = Abs(BottomArcRadius * Cos(FlapMarkerAngle))
        LittleBit = BottomRadius - adj
        PR_MakeXY xyFlapMarker, Profile.X(Profile.n) + LastTapeValue - LittleBit, 0 + opp
    ElseIf FlapMarkerLength = TopArcLength Then
        PR_MakeXY xyFlapMarker, xyBotFlap5.X, xyBotFlap5.y
    End If
    PR_SetLayer "Notes"
    Print #fNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "closed arrow"; QC; "xyStart.x +"; Str$(xyFlapMarker.X); CC; "xyStart.y +"; Str$(xyFlapMarker.y); ",0.25,0.1,225);"
    PR_SetLayer "Template" & txtSleeve

    For ii = 1 To 4
        PhiExtra = Phi + (ii * TopArcIncrement)
        PhiExtra = (PhiExtra * PI) / 180
        opp = Abs(TopRadius * Sin(PhiExtra))
        adj = Abs(TopRadius * Cos(PhiExtra))
        ShoulderFlap.X(10 - ii) = xyCentre.X - adj
        ShoulderFlap.y(10 - ii) = xyCentre.y - opp
    Next ii

    
    If Left$(txtFlap, 6) = "Raglan" Then
       'Copy standard shoulder from profile down
       'only go one vetex past the FlapMarker
        RaglanFlap.n = 0
        For ii = 10 To 1 Step -1
            RaglanFlap.n = RaglanFlap.n + 1
            RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
            RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
            If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
        Next ii
        PR_DrawFitted RaglanFlap
    Else
        ShoulderFlap.n = 10
        PR_DrawFitted ShoulderFlap
    End If
 
    'Draw Raglan
    If Left$(txtFlap, 6) = "Raglan" Then
        Rag1 = Sqr((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
        Rag1 = BottomRadius - Rag1
        PR_MakeXY xyRaglan1, Profile.X(Profile.n) + LastTapeValue - Rag1, nNotchOffset
        PR_MakeXY xyRaglan2, xyRaglan1.X - nNotchToTangent, 0
        RagAng1 = FN_CalcAngle(xyRaglan2, xyRaglan1)
        Rag2 = FN_CalcLength(xyRaglan2, xyRaglan1)
        Rag2 = Rag2 / 2
        RagAng1 = (RagAng1 * PI) / 180
        opp = Abs(Rag2 * Sin(RagAng1))
        adj = Abs(Rag2 * Cos(RagAng1))
        PR_MakeXY xyRaglan3, xyRaglan2.X + adj, xyRaglan2.y + opp
        
        RagAng1 = FN_CalcAngle(xyRaglan3, xyRaglan1)
        RagAng1 = 90 - RagAng1
        RagAng1 = (RagAng1 * PI) / 180
        Rag3 = FN_CalcLength(xyRaglan3, xyRaglan1)
        Rag4 = Sqr((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
        opp = Abs(Rag4 * Sin(RagAng1))
        adj = Abs(Rag4 * Cos(RagAng1))
        PR_MakeXY xyRaglan4, xyRaglan3.X - adj, xyRaglan3.y + opp
        PR_DrawArc xyRaglan4, xyRaglan1, xyRaglan2
        
        RagAng2 = (nTemplateAngle * PI) / 180
        opp = Abs(RaglanBottom * Sin(RagAng2))
        adj = Abs(RaglanBottom * Cos(RagAng2))
        PR_MakeXY xyRaglan5, xyRaglan1.X + adj, xyRaglan1.y + opp
        PR_DrawLine xyRaglan1, xyRaglan5
        RagAng3 = FN_CalcAngle(xyRaglan5, xyFlapMarker)
        
        If RagAng3 > 180 Then
            RagAng3 = RagAng3 - 180
            RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
        ElseIf RagAng3 <= 180 Then
            RagAng3 = RagAng3 - 90
            RagAng4 = 90 - RagAng3
            RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
        End If
        
        RagAng5 = RagAng4 - 13
        RagAng4 = (RagAng4 * PI) / 180
        opp = Abs(RaglanTip * Sin(RagAng4))
        adj = Abs(RaglanTip * Cos(RagAng4))
        PR_MakeXY xyRaglan6, xyRaglan5.X - adj, xyRaglan5.y + opp
        
        'Draw Front line of tip
        RagAng5 = (RagAng5 * PI) / 180
        opp = Abs(InnerRaglanTip * Sin(RagAng5))
        adj = Abs(InnerRaglanTip * Cos(RagAng5))
        PR_MakeXY xyRaglan7, xyRaglan5.X - adj, xyRaglan5.y + opp
     
        PR_DrawLine xyRaglan5, xyRaglan6
        PR_DrawLine xyRaglan6, xyFlapMarker
        PR_SetLayer "Notes"
        PR_DrawLine xyRaglan5, xyRaglan7

'    PR_AddMarker xyRaglan1, "xyRaglan1"
'    PR_AddMarker xyRaglan2, "xyRaglan2"
'    PR_AddMarker xyRaglan3, "xyRaglan3"
'    PR_AddMarker xyRaglan4, "xyRaglan4"
'    PR_AddMarker xyRaglan5, "xyRaglan5"
'    PR_AddMarker xyRaglan6, "xyRaglan6"
'    PR_AddMarker xyRaglan7, "xyRaglan7"
'    PR_AddMarker xyRaglan6, "xyRaglan5"
'    PR_AddMarker xyFlapMarker, "xyFlapMarker"

    End If
    
    PR_SetLayer "Notes"
    PR_MakeXY xyStrapText, xyFlapMarker.X - 1.5, xyFlapMarker.y
    Strap = "STRAP = " & Trim(fnInchesToText(cmtoinches(txtinchflag, txtStrap))) & "\" & QQ
    If Val(txtFrontStrapLength) > 0 Then
      'Display FRONT strap length if given
       Strap = "BACK  STRAP = " & Trim(fnInchesToText(cmtoinches(txtinchflag, txtStrap))) & "\" & QQ
       Strap = Strap & "\nFRONT STRAP = " & Trim(fnInchesToText(cmtoinches(txtinchflag, txtFrontStrapLength))) & "\" & QQ
    End If
    If Val(txtWaistCir) > 0 Then
      'Display Waist Circumference for D-Style flaps only
       Strap = Strap & "\nWAIST = " & Trim(fnInchesToText(cmtoinches(txtinchflag, txtWaistCir))) & "\" & QQ
    End If
    PR_DrawText Strap, xyStrapText, 0.125
    
    xfabby = txtFlap
    xfablen = Len(xfabby)
    xfabdist = (xfablen * 0.075) + 0.1
    PR_MakeXY xyFlapText1, xyFlapMarker.X - xfabdist, xyFlapMarker.y - 0.75
    PR_DrawText txtFlap, xyFlapText1, 0.125

    PR_MakeXY xyPt1, Profile.X(Profile.n), 0
    PR_MakeXY xyPt2, Profile.X(Profile.n) + LastTapeValue, 0

    If Left$(txtFlap, 6) = "Raglan" Then
        PR_MakeXY xyPotA, xyPt2.X - 0.2875, xyPt2.y
        PR_MakeXY xyPotB, xyPt2.X - 1.225, xyPt2.y
        PR_DrawLineOffset xyPt1, xyPotA, 0.6875      'Inner tram line
        PR_DrawLineOffset xyPt1, xyPotB, 0.1875      'seam line at 3/16ths
    Else
        PR_MakeXY xyPotA, xyPt2.X - 0.09375, xyPt2.y
        PR_MakeXY xyPotB, xyPt2.X - 0.03125, xyPt2.y
        PR_DrawLineOffset xyPt1, xyPotA, 0.6875      'Inner tram line
        PR_DrawLineOffset xyPt1, xyPotB, 0.1875      'seam line at 3/16ths
    End If

    PR_SetLayer "Template" & txtSleeve
    
    If Left$(txtFlap, 6) = "Raglan" Then
        PR_MakeXY xyPotA, xyPt2.X - 1.75, xyPt2.y
        PR_DrawLine xyPt1, xyRaglan2
    Else
        PR_DrawLine xyPt1, xyPt2                        'Fold Line
    End If
    PR_AddEntityID g_sFileNo, g_sSide, "FoldLineFlap"   'sFileNo, sSide from FN_Open

End If
End Sub

Private Sub PR_GauntletHoleCheck()

If txtGauntletFlag = "True" Then
    Dim Gauntletpwdiff, count
    Gauntletpwdiff = (Val(txtWristNo) - Val(txtPalmNo))
    Gauntletpwdiff = Gauntletpwdiff - 1
    If Gauntletpwdiff = 1 Then
        g_sHoleCheck = "True"
        Length(txtPalmNo) = ""
        Measurements(txtPalmNo + 1) = ""
        gms(txtPalmNo) = ""
        reds(txtPalmNo) = ""
        mms(txtPalmNo) = ""
        InchText(txtPalmNo) = ""
    Else
    g_sHoleCheck = "False"
    End If
End If

End Sub

Private Sub PR_GetFabric()
'   ------------------------
'   *** Get Fabric array ***
'   ------------------------
Dim fabnumber, Fabtype, fabnam
Dim Modul, fabcount
fabnam = Left$(cboFabric.Text, 3)
Fabtype = Mid$(cboFabric.Text, 5, 3)
If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
    MsgBox "Fabric Modulus Not Recognised", 48, "Arm Details"
    Exit Sub
End If
If (fabnam <> "Pow") And (fabnam <> "Bob") Then
    MsgBox "Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details"
    Exit Sub
End If
g_sFabnam = cboFabric.Text
txtFabric = cboFabric.Text
fabnumber = 160
fabcount = 1
   Do While fabcount < 20
        If Val(Fabtype) = fabnumber Then
            Modul = fabcount
            g_sModulus = Modul
            txtModulus = Modul
        End If
        fabcount = fabcount + 1
        fabnumber = fabnumber + 10
    Loop
'   *** End Get Fabric ***
End Sub

Private Sub PR_GetGauntletDetails()
'       *** Start Gauntlet Details ***

    If chkGauntlets.Value = True Then
        txtWristNo = cboWrist.ListIndex + 1
        If Val(txtWristNo) = 0 Then txtWristNo = ""
        
        txtPalmNo = cboPalm.ListIndex + 1
        If Val(txtPalmNo) = 0 Then txtPalmNo = ""
        
        Call Gauntlet(chkNoThumb, txtPalmNo, txtWristNo, txtPalmWrist, txtThumb, txtSex, txtAge, chkEnclosed, txtCircum, chkDetachable, txtinchflag)
    ElseIf chkGauntlets.Value = False Then
        txtGauntletFlag = "False"
    End If

''       *** End Gauntlet Details ***

End Sub

Private Sub PR_GetPressure()
'   --------------------
'   *** get Pressure ***
'   --------------------
    If cboPressure.Text = "15mm" Then
        g_sAmm = 12
        g_sBmm = 15
        g_sCmm = 10
        g_sDmm = 8
    ElseIf cboPressure.Text = "20mm" Then
        g_sAmm = 16
        g_sBmm = 20
        g_sCmm = 13
        g_sDmm = 10
    ElseIf cboPressure.Text = "25mm" Then
        g_sAmm = 20
        g_sBmm = 25
        g_sCmm = 17
        g_sDmm = 13
    End If
    txtMM.Text = cboPressure.Text
'   --- end get pressure ---

End Sub

Private Sub PR_GetStartEndTapes()
'Now sets start and last tape text values w.r.t style start and end
Dim i, ii
   'first  & second tapes
    i = 0
    txtFirstTape = ""
    Do While i < 18
         If Length(i) <> "" Then
            g_iFirstTape = i + 1
            g_iSecondTape = i + 2
            Exit Do
         End If
         i = i + 1
    Loop

    If g_iStyleFirstTape > g_iFirstTape And g_iStyleFirstTape >= 0 And cboDistalTape.ListIndex <> 0 Then
       'Use style tape value if it is given ( g_iStyleFirstTape = -1 => not given)
        txtFirstTape = g_iStyleFirstTape
        txtSecondTape = g_iStyleFirstTape + 1
    Else
        txtFirstTape = g_iFirstTape
        txtSecondTape = g_iSecondTape
    End If

   'secondlast & last tapes
    txtLastTape = ""
    ii = 17
    Do While ii > 0
        If Length(ii) <> "" Then
            g_iLastTape = ii + 1
            g_iSecondLastTape = ii
            Exit Do
        End If
        ii = ii - 1
    Loop

    If g_iStyleLastTape < g_iLastTape And g_iStyleLastTape >= 0 And cboProximalTape.ListIndex <> 0 Then
       'Use style tape value if it is given ( g_iStyleLastTape = -1 => not given)
        txtLastTape = g_iStyleLastTape
        txtSecondLastTape = g_iStyleLastTape - 1
    Else
        txtLastTape = g_iLastTape
        txtSecondLastTape = g_iSecondLastTape
    End If

End Sub

Private Sub PR_InitaliseArrays(w As Variant, X As Variant)
        If Length(w) = "" Then
           Measurements(X) = ""
           tapewidths(X) = ""
        End If
End Sub

Private Sub PR_InitaliseEmpty(w As Variant)
'initalise empty fields on front end
        If Length(w) = "" Then
           mms(w) = ""
           gms(w) = ""
           reds(w) = ""
           InchText(w) = ""
        End If
End Sub

Private Sub PR_LastCircum(w As Variant, X As Variant)
Dim temptapes1, temptapes2
If w > 9 Then

    If Length(w) <> "" And Length(w - 1) <> "" Then
'check circum of last tape
    temptapes1 = cmtoinches(txtinchflag, Length(w))

    temptapes2 = cmtoinches(txtinchflag, Length(w - 1))

    If (w = (Val(txtLastTape) - 1)) And (Val(txtAge) > 10) And (Val(temptapes1 - temptapes2) > 1.5) And (Val(txtLastTape) > 11) Then

        MsgBox "Last tape is measured over axilla. Exclude measurement", 48, "Arm Details"

       ' Exit Sub
    ElseIf (w = (Val(txtLastTape) - 1)) And (Val(txtAge) < 11) And (Val(temptapes1 - temptapes2) > 0.75) And (Val(txtLastTape) > 11) Then

        MsgBox "Last tape is measured over axilla. Exclude measurement", 48, "Arm Details, child"

       ' Exit Sub

    End If

End If
End If

End Sub

Private Sub PR_LastTapeRed(w As Variant, X As Variant, Modulus As Variant)
Dim Backgram
If Length(w) <> "" Then
'check last tape reduction
    If (w > 1) And (w = Val(txtLastTape) - 1) And (cboFlaps = "") Then
        If Left$(g_sFabnam, 3) = "Pow" Then
            Backgram = Pow(Modulus, 1)
            If Val(gms(w)) <> Backgram Then
                gms(w) = round(Backgram)
                mms(w) = round(Backgram / Measurements(X))
            End If
        Else
            Backgram = Bob(Modulus, 1)
            If Val(gms(w)) <> Backgram Then
                gms(w) = round(Backgram)
                mms(w) = round(Backgram / Measurements(X))
            End If
        End If
        reds(w) = 10
    
    ElseIf (w > 1) And (w = Val(txtLastTape) - 1) And ((txtFlap <> "") Or (Left$(cboFlaps, 4) = "None") Or (Left$(cboFlaps, 5) = "Style")) Then
        If Left$(cboFlaps, 4) <> "Vest" And Left$(cboFlaps, 4) <> "Body" Then
            If Left$(g_sFabnam, 3) = "Pow" Then
                Backgram = Pow(Modulus, 3) '12 reduction for Flaps
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
             Else
                Backgram = Bob(Modulus, 3) '12 reduction for Flaps
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
             End If
             reds(w) = 12
        Else
            If Left$(g_sFabnam, 3) = "Pow" Then
                Backgram = Pow(Modulus, 1) '10 reduction for Vest & Body
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
             Else
                Backgram = Bob(Modulus, 1) '10 reduction for Vest & Body
                If Val(gms(w)) <> Backgram Then
                    gms(w) = round(Backgram)
                    mms(w) = round(Backgram / Measurements(X))
                End If
             End If
             reds(w) = 10
        End If
   End If
End If
' end check last tape reduction

End Sub

Private Sub PR_OldDrawing()
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Flap"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Gauntlet"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Stump"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeLengths"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "TapeMMs"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Reduction"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Grams"; QCQ; "string"; QQ; ");"
Print #fNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ID"; QCQ; "string"; QQ; ");"
End Sub

Private Sub PR_PalmRed(w As Variant, X As Variant, Modulus As Variant)
Dim Backgram
If Length(w) <> "" Then
'check palm reduction
    If (chkGauntlets.Value = True And X = txtPalmNo) Or (chkGauntlets.Value = True And w = (txtFirstTape - 1)) Then

        reds(w) = 15

        If Left$(g_sFabnam, 3) = "Pow" Then
            Backgram = Pow(Modulus, 6)
        Else
            Backgram = Bob(Modulus, 6)
        End If
        
        mms(w) = round(Backgram / Measurements(X))

        gms(w) = round(Backgram)

    End If
End If

End Sub

Private Sub PR_Pressurecalc(Index As Integer)
Dim w, X
    w = Index
    X = Index + 1
        PR_InitaliseEmpty w              'Zero any empty fields
        PR_InitaliseArrays w, X
        If (Length(Index) <> "" And mms(Index) = "") Or (Length(Index) <> "" And g_sPressureChange = "True") Then
            PR_SetMMs w, g_sAmm, g_sBmm, g_sCmm, g_sDmm
        End If
        
        PR_Reductions w, X
        PR_WristRed w, X, g_sModulus
        PR_PalmRed w, X, g_sModulus
        PR_Stump w, X
        PR_LastCircum w, X
        PR_LastTapeRed w, X, g_sModulus
        PR_CheckFlaps w
        PR_Detachable w, X, g_sModulus, g_sFabnam
End Sub

Private Sub PR_Reductions(w As Variant, X As Variant)
Dim Gram
If Length(w) <> "" Then
    Gram = Fix(Val(mms(w)) * Val(Measurements(X)))
    If Gram <> 0 Then
        gms(w) = round(Gram)
        reds(w) = GetReduction(g_sModulus, Gram, g_sFabnam)
    End If
End If

End Sub

Private Sub PR_ReSetMMs(w As Variant, Modulus As Variant)
Dim givenreduction, X
X = w + 1
If Length(w) <> "" And mms(w) <> "" And reds(w) <> "" Then
    givenreduction = Val(reds(w) - 9)
    If givenreduction > 32 Or givenreduction < 1 Then
        MsgBox "Reduction Value Not Recognised", 48, "Arm Details"
        Exit Sub
    End If
    If Left$(g_sFabnam, 3) = "Pow" Then
        Backgram = Pow(Modulus, givenreduction)
        If Val(gms(w)) <> Backgram Then
            gms(w) = round(Backgram)
            newmms = Backgram / Measurements(X)
            tempmms = round(newmms)
            tempmms = (newmms - tempmms) * 10
            If tempmms >= 5 Then
                mms(w) = round(newmms) + 1
            Else
                mms(w) = round(newmms)
            End If
           
        End If
    Else
        Backgram = Bob(Modulus, givenreduction)
        If Val(gms(w)) <> Backgram Then
            gms(w) = round(Backgram)
            newmms = Backgram / Measurements(X)
            tempmms = round(newmms)
            tempmms = (newmms - tempmms) * 10
            If tempmms >= 5 Then
                mms(w) = round(newmms) + 1
            Else
                mms(w) = round(newmms)
            End If
        End If
    End If
End If
End Sub

Private Sub PR_SetDistances(w As Variant, X As Variant)

'Length(w)
'Measurements(X)

'Check Distances between tapes
Dim tdist, tpace
tpace = 1.375
If Length(w) <> "" Then
   'First tape
    If X = Val(txtFirstTape) Then
        tapewidths(X) = "0"
   
   'secondTape
    ElseIf X = Val(txtSecondTape) And chkGauntlets = False Then
        If Val(txtWristPleat1) = 0 Or Not txtWristPleat1.Enabled Then
            tapewidths(X) = tpace
        ElseIf Val(txtWristPleat1) > 0 Then
            tdist = Val(txtWristPleat1)
            tdist = FN_PleatDistances(txtinchflag, tdist)
            tapewidths(X) = tapewidths(X - 1) + tdist
        End If

    ElseIf X = Val(txtSecondTape) And chkGauntlets = True Then
        If Val(Length(w)) = 0 Then
            tapewidths(X) = ""
        ElseIf Val(Length(w)) > 0 And Val(txtWristNo) = X Then
            tapewidths(X) = Val(txtPalmWristDist)
        ElseIf Val(Length(w)) > 0 And Val(txtWristNo) <> X Then
            tapewidths(X) = tpace
        End If
        
   'thirdtape
    ElseIf X = Val(txtSecondTape + 1) And chkGauntlets = False Then
        If Val(txtWristPleat2) = 0 Or Not txtWristPleat2.Enabled Then
            tapewidths(X) = tapewidths(X - 1) + tpace
        ElseIf Val(txtWristPleat2) > 0 Then
            tdist = Val(txtWristPleat2)
            tdist = FN_PleatDistances(txtinchflag, tdist)
            tapewidths(X) = Val(tapewidths(X - 1)) + tdist
        End If
    
    ElseIf X = Val(txtSecondTape + 1) And chkGauntlets = True Then
        If (Length(w - 1) = "") Then
            tapewidths(X) = Val(tapewidths(X - 2)) + Val(txtPalmWristDist)
        ElseIf Length(w - 1) <> "" And Val(txtWristNo) = X Then
            tapewidths(X) = tapewidths(X - 1) + Val(txtPalmWristDist)
        ElseIf Length(w - 1) <> "" And Val(txtWristNo) <> X Then
            tapewidths(X) = tapewidths(X - 1) + tpace
        ElseIf Length(w - 1) <> "" Then
            tapewidths(X) = ""
        End If
         
   'fourthtape
    ElseIf X = Val(txtSecondTape + 2) And chkGauntlets = False Then
        tapewidths(X) = tapewidths(X - 1) + tpace

    ElseIf X = Val(txtSecondTape + 2) And chkGauntlets = True Then
        If Length(w - 1) = "" And Val(txtWristNo) = X Then
            tapewidths(X) = tapewidths(X - 2) + Val(txtPalmWristDist)
        ElseIf Length(w - 1) <> "" And Val(txtWristNo) = X Then
            tapewidths(X) = tapewidths(X - 1) + Val(txtPalmWristDist)
        ElseIf Length(w - 1) <> "" And Val(txtWristNo) <> X Then
            tapewidths(X) = tapewidths(X - 1) + tpace
        End If
        
   'Tapes above fourthtape
    ElseIf X > (Val(txtSecondTape) + 2) And X < Val(txtSecondLastTape) Then
        tapewidths(X) = tapewidths(X - 1) + tpace

   'Lasttape -1
    ElseIf X = Val(txtSecondLastTape) And (X > 1) Then
        If Val(txtShoulderPleat2) = 0 Or Not txtShoulderPleat2.Enabled Then
            tapewidths(X) = Val(tapewidths(X - 1)) + tpace
        ElseIf Val(txtShoulderPleat2) > 0 Then
            tdist = Val(txtShoulderPleat2)
            tdist = FN_PleatDistances(txtinchflag, tdist)
            tapewidths(X) = Val(tapewidths(X - 1)) + tdist
        End If

   'Last tape
    ElseIf X = Val(txtLastTape) And (X > 1) Then
        If Val(txtShoulderPleat1) = 0 Or Not txtShoulderPleat1.Enabled Then
            tapewidths(X) = Val(tapewidths(X - 1)) + tpace
        ElseIf Val(txtShoulderPleat1) > 0 Then
            tdist = Val(txtShoulderPleat1)
            tdist = FN_PleatDistances(txtinchflag, tdist)
            tapewidths(X) = Val(tapewidths(X - 1)) + tdist
        End If
    End If
    
   'The following statements duplicate some of the above.  They are here
   'as bug fixes as I can't understand the above code.  The above does work
   '(mostly) so I have left well enough alone.

   'Account for wrist pleats if a gauntlet given
    If Val(txtWristNo) + 1 = X And (Val(txtWristPleat1) > 0 And txtWristPleat1.Enabled) And chkGauntlets = True Then
        tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtWristPleat1))
    End If
    If Val(txtWristNo) + 2 = X And (Val(txtWristPleat2) > 0 And txtWristPleat2.Enabled) And chkGauntlets = True Then
        tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtWristPleat2))
    End If

   'Account for 1ST shoulder pleat
    If Val(txtLastTape) = X And (Val(txtShoulderPleat1) > 0 And txtShoulderPleat1.Enabled) Then
        tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtShoulderPleat1))
    End If
   'Account for 2nd shoulder pleat
    If Val(txtSecondLastTape) = X And (Val(txtShoulderPleat2) > 0 And txtShoulderPleat2.Enabled) Then
        tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtShoulderPleat2))
    End If
     

End If
    'End Check Distances between tape
End Sub

Private Sub PR_SetGlobals()
'    SetTapeLengths
    g_sTapeLengths = txtTapeLent

'    SetMMs
    g_sTapeMMs = txtTapeMM
    '    SetGrams
    g_sGrams = txtWeight
'    SetReductions
    g_sReduction = txtReduction
    g_sFlapStrap = txtStrap
   'Revised method w.r.t. multiple styles
    If optProximalTape(0).Value = True Then
        g_sFlapChk = "False"
    Else
        g_sFlapChk = "True"
    End If
    If optProximalTape(1).Value = True Then
        g_sFlapChk = "True"
    End If
    g_sPatient = txtPatientName
    g_sGaunt = txtGauntletFlag
    g_sNoThumb = txtNoThumb
    g_sPatient = txtPatientName
    g_sWpleat1 = txtWristPleat1
    If g_sWpleat1 = "" Then g_sWpleat1 = "0"
    g_sWpleat2 = txtWristPleat2
    g_sSpleat1 = txtShoulderPleat1
    If g_sSpleat1 = "" Then g_sSpleat1 = "0"
    g_sSpleat2 = txtShoulderPleat2
    g_sDetGaunt = txtDetach
    g_sEnclosedThumb = txtEnclosed
    g_sPalmNo = txtPalmNo
    g_sWristNo = txtWristNo
    g_sPalmWristDist = txtPalmWrist
    g_sThumbCircum = txtCircum
    g_sThumbLength = txtThumb
    g_sSecondTape = txtSecondTape
    g_sSecondLastTape = txtSecondLastTape
    g_sMM = cboPressure.Text
    g_sFirstTape = txtFirstTape
    g_sLastTape = txtLastTape
    g_sStump = txtStump
    g_sSide = txtSleeve
    g_sWaistCir = txtWaistCir

End Sub

Private Sub PR_SetMMs(w As Variant, Amm As Variant, Bmm As Variant, Cmm As Variant, Dmm As Variant)
Dim GenMMs
       If Length(w) <> "" And w >= txtFirstTape - 1 And w <= txtLastTape - 1 Then
            If (w = txtFirstTape - 1) Or (w = 9) Then
                GenMMs = Amm
            ElseIf (w > 11) And (w < (txtSecondLastTape - 1)) Then
                GenMMs = Amm
            ElseIf ((w > txtFirstTape - 1) And (w < 9)) Then
                GenMMs = Bmm
            ElseIf (w = 11) Or (w = (txtSecondLastTape - 1)) Then
                GenMMs = Cmm
            ElseIf (w = 10) Or (w = (txtLastTape - 1)) Then
                GenMMs = Dmm
            End If
            If (txtWristNo <> "" And w = Val(txtWristNo) - 1) Then
                GenMMs = Amm
            End If
            mms(w) = Val(GenMMs)
       Else
            mms(w) = ""
       End If
End Sub

Private Sub PR_SetRresults()
    SetTapeLengths
    SetMMs
    SetGrams
    SetReductions
End Sub

Private Sub PR_Stump(w As Variant, X As Variant)
 If Length(w) <> "" Then
    If (w = (txtFirstTape - 1)) And chkStump.Value = True Then
        reds(w) = 14
        If Left$(g_sFabnam, 3) = "Pow" Then
            Backgram = Pow(txtModulus, 5)
        Else
            Backgram = Bob(txtModulus, 5)
        End If
        mms(w) = round(Backgram / Measurements(X))
        gms(w) = round(Backgram)
        txtStump = "True"
    End If
 End If
End Sub

Private Sub PR_WristRed(w As Variant, X As Variant, Modulus As Variant)
Dim Backgram, Reduction

If Length(w) <> "" And X = Val(txtFirstTape) And txtStump <> "True" And txtGauntletFlag <> "True" Then

' check wrist reduction
    If (w < 9) And reds(w) = "" Then
            reds(w) = 14
    ElseIf reds(w) > 14 And (w < 9) Then
            reds(w) = 14
   'Set elbow reduction to 5
    ElseIf reds(w) <> 5 And (w = 10) Then
            reds(w) = 5
   'Above elbow is also a 5
    ElseIf reds(w) <> 5 And (w > 10) Then
            reds(w) = 5
    ElseIf (w < 9) And ((reds(w) < 10) And (reds(w) > 0)) Then
            reds(w) = 10
    End If

   'Set up an index into the reduction to grams mapping table
   'minimum is 1
    Reduction = reds(w) - 9
    If Reduction < 1 Then Reduction = 1

   'From either powernet or bobinnet table get the grams and back calculate the
   'mms
    If Left$(g_sFabnam, 3) = "Pow" Then
        Backgram = Pow(Modulus, Reduction)
    Else
        Backgram = Bob(Modulus, Reduction)
    End If

    mms(w) = round(Backgram / Measurements(X))
    gms(w) = round(Backgram)

End If

' end check wrist reduction
End Sub

Private Sub select_text_in_Box1(Text_Box_Name As Control)
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = 4
End Sub

Private Sub SetGrams()
'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
Dim ii%
    txtWeight = ""
    For ii = 0 To 17
       txtWeight = txtWeight & PadGram(gms(ii))
    Next ii
'a = PadGram(gms(0))
'b = PadGram(gms(1))
'c = PadGram(gms(2))
'd = PadGram(gms(3))
'e = PadGram(gms(4))
'f = PadGram(gms(5))
'G = PadGram(gms(6))
'h = PadGram(gms(7))
'i = PadGram(gms(8))
'j = PadGram(gms(9))
'k = PadGram(gms(10))
'l = PadGram(gms(11))
'm = PadGram(gms(12))
'n = PadGram(gms(13))
'o = PadGram(gms(14))
'P = PadGram(gms(15))
'q = PadGram(gms(16))
'r = PadGram(gms(17))
'txtWeight = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

End Sub

Private Sub SetMMs()
'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
Dim ii%
    txtTapeMM = ""
    For ii = 0 To 17
       txtTapeMM = txtTapeMM & PadMM(mms(ii))
    Next ii
'a = PadMM(mms(0))
'b = PadMM(mms(1))
'c = PadMM(mms(2))
'd = PadMM(mms(3))
'e = PadMM(mms(4))
'f = PadMM(mms(5))
'G = PadMM(mms(6))
'h = PadMM(mms(7))
'i = PadMM(mms(8))
'j = PadMM(mms(9))
'k = PadMM(mms(10))
'l = PadMM(mms(11))
'm = PadMM(mms(12))
'n = PadMM(mms(13))
'o = PadMM(mms(14))
'P = PadMM(mms(15))
'q = PadMM(mms(16))
'r = PadMM(mms(17))
'txtTapeMM = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

End Sub

Private Sub SetReductions()
'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
Dim ii%
    txtReduction = ""
    For ii = 0 To 17
       txtReduction = txtReduction & PadRed(reds(ii))
    Next ii
'a = PadRed(reds(0))
'b = PadRed(reds(1))
'c = PadRed(reds(2))
'd = PadRed(reds(3))
'e = PadRed(reds(4))
'f = PadRed(reds(5))
'G = PadRed(reds(6))
'h = PadRed(reds(7))
'i = PadRed(reds(8))
'j = PadRed(reds(9))
'k = PadRed(reds(10))
'l = PadRed(reds(11))
'm = PadRed(reds(12))
'n = PadRed(reds(13))
'o = PadRed(reds(14))
'P = PadRed(reds(15))
'q = PadRed(reds(16))
'r = PadRed(reds(17))
'txtReduction = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

End Sub

Private Function SetStump(Length)
Dim Circum
Circum = Length
Circum = Circum * 0.86
SetStump = Circum
End Function

Private Sub SetTapeLengths()
Dim ii%
    txtTapeLent = ""
    For ii = 0 To 17
       txtTapeLent = txtTapeLent & PadTape(Trim(Length(ii)))
    Next ii

End Sub

Private Sub Tab_Click()
   'Allows the user to use enter as a tab

    SendKeys "{TAB}", False

End Sub

Private Sub ThumbTip(xyArcStart As XY, xyArcEnd As XY)
' This Calculates and Draws the curve at the
' right of the ThumbHole on the template
Dim xyCen As XY
Dim nRad, nStartAng, nEndAng, nDeltaAng
    nRad = FN_CalcLength(xyArcStart, xyArcEnd)
    nRad = nRad / 2
    nStartAng = FN_CalcAngle(xyArcStart, xyArcEnd)
    nStartAng = nStartAng * (PI / 180)
    nAdj = Cos(nStartAng) * nRad
    nOpp = Sin(nStartAng) * nRad
    xyCen.X = xyArcStart.X + nAdj
    xyCen.y = xyArcStart.y + nOpp
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nDeltaAng = -180
    Print #fNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"
End Sub

Private Sub Timer1_Timer()
    'This is enabled on Load
    'It is disabled in Link close
    'If the event happens then it is assumed that
    'for some reason the link between DRAFIX and This dialogue
    'has failed'
    'Therefor we "End" here
     End
     End
End Sub

Private Sub txtCircum_GotFocus()
    select_text_in_Box1 txtCircum
End Sub

Private Sub txtCircum_LostFocus()
    If Not FN_CheckValue(txtCircum, "Thumb Circumference") Then
        Exit Sub
    ElseIf Val(txtCircum) > 0 Then
            TCirc = fnInchesToText(cmtoinches(txtinchflag, txtCircum))
        Else
            TCirc = ""
    End If
End Sub

Private Sub txtCustFlapLength_GotFocus()
    select_text_in_Box1 txtCustFlapLength
End Sub

Private Sub txtCustFlapLength_LostFocus()
    If Not FN_CheckValue(txtCustFlapLength, "Custom Flap") Then
        Exit Sub
    ElseIf Val(txtCustFlapLength) > 0 Then
            labCustFlapLength.Caption = fnInchesToText(cmtoinches(txtinchflag, txtCustFlapLength))
        Else
            labCustFlapLength.Caption = ""
    End If
End Sub

Private Sub txtFrontStrapLength_GotFocus()
   select_text_in_Box1 txtFrontStrapLength
End Sub

Private Sub txtFrontStrapLength_LostFocus()
    If Not FN_CheckValue(txtFrontStrapLength, "Front Strap") Then
        Exit Sub
    ElseIf Val(txtFrontStrapLength) > 0 Then
            labFrontStrapLength.Caption = fnInchesToText(cmtoinches(txtinchflag, txtFrontStrapLength))
        Else
            labFrontStrapLength.Caption = ""
    End If
End Sub

Private Sub txtGauntletExtension_GotFocus()
    select_text_in_Box1 txtGauntletExtension
End Sub

Private Sub txtGauntletExtension_LostFocus()
    If Not FN_CheckValue(txtGauntletExtension, "Gauntlet Extension") Then
        Exit Sub
    ElseIf Val(txtGauntletExtension) > 0 Then
            labGauntletExtension.Caption = fnInchesToText(cmtoinches(txtinchflag, txtGauntletExtension))
        Else
            labGauntletExtension.Caption = ""
    End If
End Sub

Private Sub txtPalmWrist_GotFocus()
     select_text_in_Box1 txtPalmWrist
End Sub

Private Sub txtPalmWrist_LostFocus()
    If Not FN_CheckValue(txtPalmWrist, "Palm to Wrist") Then
        Exit Sub
    ElseIf Val(txtPalmWrist) > 0 Then
            PWDist = fnInchesToText(cmtoinches(txtinchflag, txtPalmWrist))
        Else
            PWDist = ""
    End If
End Sub

Private Sub txtShoulderPleat1_GotFocus()
    select_text_in_Box1 txtShoulderPleat1
End Sub

Private Sub txtShoulderPleat1_LostFocus()
    If Not FN_CheckValue(txtShoulderPleat1, "First Shoulder Pleat") Then
        Exit Sub
    End If
End Sub

Private Sub txtShoulderPleat2_GotFocus()
    select_text_in_Box1 txtShoulderPleat2
End Sub

Private Sub txtShoulderPleat2_LostFocus()
    If Not FN_CheckValue(txtShoulderPleat2, "Second Shoulder Pleat") Then
        Exit Sub
    End If
End Sub

Private Sub txtStrapLength_GotFocus()
    select_text_in_Box1 txtStrapLength
End Sub

Private Sub txtStrapLength_LostFocus()

    If Not FN_CheckValue(txtStrapLength, "Strap") Then
        Exit Sub
    ElseIf Val(txtStrapLength) > 0 Then
            labStrap = fnInchesToText(cmtoinches(txtinchflag, txtStrapLength))
            txtStrap = txtStrapLength
        Else
            labStrap = ""
    End If


Dim w, X
X = Val(txtLastTape)
w = X - 1

If X > 0 Then

   If Length(w) <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w) <> "" And reds(w) <> "" Then
        
        txtFlap = cboFlaps.Text
        
        PR_LastTapeRed w, X, g_sModulus

        PR_CheckFlaps w
        
        PR_SetRresults
        
        PR_SetGlobals
        
     End If
End If

End Sub

Private Sub txtThumb_GotFocus()
    select_text_in_Box1 txtThumb
End Sub

Private Sub txtThumb_LostFocus()
    If Not FN_CheckValue(txtThumb, "Thumb") Then
        Exit Sub
    ElseIf Val(txtThumb) > 0 Then
            TLen = fnInchesToText(cmtoinches(txtinchflag, txtThumb))
        Else
            TLen = ""
    End If
End Sub

Private Sub txtWaistCir_GotFocus()
    select_text_in_Box1 txtWaistCir
End Sub

Private Sub txtWaistCir_LostFocus()
    If Not FN_CheckValue(txtWaistCir, "Style-D waist circumference") Then
        Exit Sub
    ElseIf Val(txtWaistCir) > 0 Then
            labWaistCir = fnInchesToText(cmtoinches(txtinchflag, txtWaistCir))
        Else
            labWaistCir = ""
    End If

End Sub

Private Sub txtWristPleat1_GotFocus()
    select_text_in_Box1 txtWristPleat1
End Sub

Private Sub txtWristPleat1_LostFocus()
    If Not FN_CheckValue(txtWristPleat1, "First Wrist Pleat") Then
        Exit Sub
    End If
End Sub

Private Sub txtWristPleat2_GotFocus()
    select_text_in_Box1 txtWristPleat2
End Sub

Private Sub txtWristPleat2_LostFocus()
    If Not FN_CheckValue(txtWristPleat2, "Second Wrist Pleat") Then
        Exit Sub
    End If
End Sub

Private Sub TypeDetails(Profile As curve, numtapes)
Dim n, xyDetails As XY, nTextHt

Dim sDetails, NL, sSymbol$, nSpacing#
    
    nSpacing = 0.2

    PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
    
    If Profile.X(Profile.n) < 6 Then
        PR_MakeXY xyDetails, Profile.X(Profile.n) - 2, 2.5
    Else
        PR_MakeXY xyDetails, Profile.X(1) + 4, 2.5
    End If
    If Left$(cboFlaps.Text, 4) = "Vest" Or Left$(cboFlaps.Text, 4) = "Body" Then
        PR_MakeXY xyDetails, Profile.X(Profile.n) + 1.5, 1.75
    End If

   'Main details in Green on layer notes
    PR_SetLayer "Notes"
    Dim sText$

    sText = txtSleeve.Text & "\n" & txtPatientName.Text & "\n" & g_sWorkOrder & "\n" & Trim$(Mid$(txtFabric.Text, 4))
    PR_DrawText sText, xyDetails, 0.125


   'Other patient details in black on layer construct
    PR_SetLayer "Construct"

    PR_MakeXY xyDetails, xyDetails.X, xyDetails.y - 0.8
    sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
    PR_DrawText sText, xyDetails, 0.125


End Sub

Private Function WristRed(lblred)
    If Val(lblred) > 14 Then
        WristRed = 14
    ElseIf Val(lblred) < 10 Then
        WristRed = 10
    End If
End Function


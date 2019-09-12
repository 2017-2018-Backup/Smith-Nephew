VERSION 5.00
Begin VB.Form HeadNeck 
   Caption         =   "Head & Neck"
   ClientHeight    =   7140
   ClientLeft      =   2430
   ClientTop       =   1155
   ClientWidth     =   6900
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
   LinkTopic       =   "headneck"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7140
   ScaleWidth      =   6900
   Begin VB.Timer Timer1 
      Left            =   8580
      Top             =   5670
   End
   Begin VB.TextBox txtUidHN 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   123
      Text            =   "txtUidHN"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtTopRightEar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   120
      Text            =   "txtTopRightEar"
      Top             =   3450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtTopLeftEar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   119
      Text            =   "txtTopLeftEar"
      Top             =   1350
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtHeadNeckUID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   118
      Text            =   "txtHeadNeckUID"
      Top             =   1950
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtOptionChoice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   117
      Text            =   "txtOptionChoice"
      Top             =   2250
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   116
      Top             =   4050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   115
      Text            =   "txtWorkOrder"
      Top             =   5250
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      MaxLength       =   19
      TabIndex        =   86
      Text            =   "txtData"
      Top             =   4950
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   390
      Left            =   3900
      TabIndex        =   36
      Top             =   6600
      Width           =   1440
   End
   Begin VB.TextBox txtMeasurements 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      MaxLength       =   43
      TabIndex        =   111
      Text            =   "txtMeasurements"
      Top             =   1650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtMouthRightY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   110
      Text            =   "txtMouthRightY"
      Top             =   3150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtxyNeckTopFrontY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   109
      Text            =   "txtxyNeckTopFrontY"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtxyNeckTopFrontX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   108
      Text            =   "txtxyNeckTopFrontX"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtxyChinTopY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   107
      Text            =   "txtxyChinTopY"
      Top             =   4350
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtxyChinTopX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   106
      Text            =   "txtxyChinTopX"
      Top             =   4050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSEarBotHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   105
      Text            =   "txtCSEarBotHeight"
      Top             =   3750
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSEarTopHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   104
      Text            =   "txtCSEarTopHeight"
      Top             =   3450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSMouthWidth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   103
      Text            =   "txtCSMouthWidth"
      Top             =   3150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtMouthRightX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   67
      Text            =   "txtMouthRightX"
      Top             =   1050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSForeHead 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   99
      Text            =   "txtCSForeHead"
      Top             =   2850
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSChinAngle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   98
      Text            =   "txtCSChinAngle"
      Top             =   2550
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSNeckCircum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   97
      Text            =   "txtCSNeckCircum"
      Top             =   2250
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCSChintoMouth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   96
      Text            =   "txtCSChintoMouth"
      Top             =   1950
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtNoseCoverX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   95
      Text            =   "txtNoseCoverX"
      Top             =   750
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtNoseCoverY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   94
      Text            =   "txtNoseCoverY"
      Top             =   450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtEyeWidth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   93
      Text            =   "txtEyeWidth"
      Top             =   1350
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtDartEndX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   92
      Text            =   "txtDartEndX"
      Top             =   750
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtDartEndY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   91
      Text            =   "txtDartEndY"
      Top             =   1050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtDartStartY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   90
      Text            =   "txtDartStartY"
      Top             =   450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtDartStartX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   89
      Text            =   "txtDartStartX"
      Top             =   150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtChinLeftBotY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   88
      Text            =   "txtChinLeftBotY"
      Top             =   2850
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtChinLeftBotX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   85
      Text            =   "txtChinLeftBotX"
      Top             =   2550
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtNoseBottomY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8550
      TabIndex        =   84
      Text            =   "txtNoseBottomY"
      Top             =   150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCircumferenceTotal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   83
      Text            =   "txtCircumferenceTotal"
      Top             =   3750
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtMouthHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   82
      Text            =   "txtMouthHeight"
      Top             =   1650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtBotOfEyeY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   81
      Text            =   "txtBotOfEyeY"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtBotOfEyeX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   80
      Text            =   "txtBotOfEyeX"
      Top             =   4350
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtLowerEarHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   79
      Text            =   "txtLowerEarHeight"
      Top             =   4050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtfNum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   78
      Text            =   "txtfNum"
      Top             =   3450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrTextFont 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   77
      Text            =   "txtCurrTextFont"
      Top             =   3150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrTextVertJust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   76
      Text            =   "txtCurrTextVertJust"
      Top             =   2850
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrTextHorizJust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   75
      Text            =   "txtCurrTextHorizJust"
      Top             =   2550
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrTextAspect 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   74
      Text            =   "txtCurrTextAspect"
      Top             =   2250
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrTextHt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   73
      Text            =   "txtCurrTextHt"
      Top             =   1950
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCurrentLayer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   72
      Text            =   "txtCurrentLayer"
      Top             =   1650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtMidToEyeTop 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   71
      Text            =   "txtMidToEyeTop"
      Top             =   1350
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtLipStrapWidth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   70
      Text            =   "txtLipStrapWidth"
      Top             =   1050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   69
      Text            =   "txtUidMPD"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtRadiusNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7050
      TabIndex        =   68
      Text            =   "txtRadiusNo"
      Top             =   450
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtDraw 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   102
      Text            =   "txtDraw"
      Top             =   150
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   1350
      TabIndex        =   35
      Top             =   6600
      Width           =   1440
   End
   Begin VB.Frame frmMeasurements 
      Caption         =   "Measurements"
      Height          =   2415
      Left            =   75
      TabIndex        =   51
      Top             =   4050
      Width           =   6765
      Begin VB.TextBox txtChinCollarMin 
         Height          =   285
         Left            =   5475
         TabIndex        =   33
         Top             =   1125
         Width           =   540
      End
      Begin VB.TextBox txtRightEarLength 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5475
         MaxLength       =   4
         TabIndex        =   32
         Top             =   825
         Width           =   540
      End
      Begin VB.TextBox txtLeftEarLength 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5475
         MaxLength       =   4
         TabIndex        =   31
         Top             =   525
         Width           =   540
      End
      Begin VB.TextBox txtHeadBandDepth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5475
         MaxLength       =   4
         TabIndex        =   30
         Top             =   225
         Width           =   540
      End
      Begin VB.TextBox txtLengthOfNose 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   29
         Top             =   2025
         Width           =   540
      End
      Begin VB.TextBox txtTipOfNose 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   28
         Top             =   1725
         Width           =   540
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   3825
         TabIndex        =   34
         Top             =   1875
         Width           =   2745
      End
      Begin VB.TextBox txtThroatToSternal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   27
         Top             =   1425
         Width           =   540
      End
      Begin VB.TextBox txtCircOfNeck 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1125
         Width           =   540
      End
      Begin VB.TextBox txtCircChinAngle 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   25
         Top             =   825
         Width           =   540
      End
      Begin VB.TextBox txtCircEyeBrow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   24
         Top             =   525
         Width           =   540
      End
      Begin VB.TextBox txtChinToMouth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2550
         MaxLength       =   4
         TabIndex        =   23
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblInch11 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   6000
         TabIndex        =   122
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label lblChinCollarMin 
         Caption         =   "Collar Contour"
         Height          =   240
         Left            =   3825
         TabIndex        =   121
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label lblRightEarLen 
         Caption         =   "Right Ear Length"
         Height          =   240
         Left            =   3825
         TabIndex        =   87
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblLeftEarLen 
         Caption         =   "Left Ear Length"
         Height          =   240
         Left            =   3825
         TabIndex        =   114
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label lblInch10 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   6000
         TabIndex        =   113
         Top             =   870
         Width           =   690
      End
      Begin VB.Label lblInch9 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   6000
         TabIndex        =   112
         Top             =   570
         Width           =   690
      End
      Begin VB.Label lblInch8 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   6000
         TabIndex        =   101
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lblHeadBandDepth 
         Caption         =   "Head Band Depth"
         Height          =   240
         Left            =   3825
         TabIndex        =   100
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label lblFabric 
         Caption         =   "Fabric : "
         Height          =   240
         Left            =   3825
         TabIndex        =   66
         Top             =   1650
         Width           =   690
      End
      Begin VB.Label lblInch7 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   65
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label lblInch6 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   64
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label LblLengthOfNose 
         Caption         =   "Length of Nose"
         Height          =   240
         Left            =   150
         TabIndex        =   63
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label lblAcrossTipOfNose 
         Caption         =   "Across Tip of Nose"
         Height          =   240
         Left            =   150
         TabIndex        =   62
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblInch5 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   61
         Top             =   1470
         Width           =   690
      End
      Begin VB.Label lblInch4 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   60
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label lblInch3 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   59
         Top             =   870
         Width           =   690
      End
      Begin VB.Label lblInch2 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   58
         Top             =   570
         Width           =   690
      End
      Begin VB.Label lblInch1 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   3075
         TabIndex        =   57
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lblThroatToSternal 
         Caption         =   "Throat to Sternal Notch"
         Height          =   240
         Left            =   150
         TabIndex        =   56
         Top             =   1500
         Width           =   2190
      End
      Begin VB.Label lblCircOfNeck 
         Caption         =   "CIRC of Neck"
         Height          =   240
         Left            =   150
         TabIndex        =   55
         Top             =   1200
         Width           =   2265
      End
      Begin VB.Label lblCircChinAngle 
         Caption         =   "CIRC of Head at Chin Angle"
         Height          =   240
         Left            =   150
         TabIndex        =   54
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label lblCircEyeBrow 
         Caption         =   "CIRC above Eyebrow"
         Height          =   240
         Left            =   150
         TabIndex        =   53
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblChinToMouth 
         Caption         =   "Chin to Mouth"
         Height          =   240
         Left            =   150
         TabIndex        =   52
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frmModifications 
      Caption         =   "Modifications"
      Height          =   2715
      Left            =   2850
      TabIndex        =   44
      Top             =   1275
      Width           =   3990
      Begin VB.CheckBox chkVelcro 
         Caption         =   "2"" Wide Velcro"
         Height          =   240
         Left            =   1950
         TabIndex        =   124
         Top             =   2400
         Width           =   1950
      End
      Begin VB.CheckBox chkEyes 
         Caption         =   "Include Eyes"
         Height          =   240
         Left            =   1950
         TabIndex        =   20
         Top             =   1500
         Width           =   1590
      End
      Begin VB.CheckBox chkRightEarClosed 
         Caption         =   "Right Ear Closed"
         Height          =   240
         Left            =   1950
         TabIndex        =   13
         Top             =   900
         Width           =   1980
      End
      Begin VB.CheckBox chkLeftEarClosed 
         Caption         =   "Left Ear Closed"
         Height          =   240
         Left            =   150
         TabIndex        =   12
         Top             =   900
         Width           =   1710
      End
      Begin VB.CheckBox chkEarSize 
         Caption         =   "Ear Size"
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CheckBox chkLipStrap 
         Caption         =   "Lip Strap"
         Enabled         =   0   'False
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   1500
         Width           =   1665
      End
      Begin VB.CheckBox chkOpenHeadMask 
         Caption         =   "Open Head Mask"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   2400
         Width           =   1980
      End
      Begin VB.CheckBox chkNeckElastic 
         Caption         =   "Neck Elastic"
         Height          =   240
         Left            =   1950
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkLining 
         Caption         =   "Lining"
         Height          =   240
         Left            =   1950
         TabIndex        =   22
         Top             =   1800
         Width           =   915
      End
      Begin VB.CheckBox chkLeftEyeFlap 
         Caption         =   "Left Eye Flap"
         Height          =   240
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   1590
      End
      Begin VB.CheckBox chkRightEyeFlap 
         Caption         =   "Right Eye Flap"
         Height          =   240
         Left            =   1950
         TabIndex        =   9
         Top             =   300
         Width           =   1665
      End
      Begin VB.CheckBox chkZipper 
         Caption         =   "Zipper"
         Height          =   240
         Left            =   1950
         TabIndex        =   21
         Top             =   2100
         Width           =   915
      End
      Begin VB.CheckBox chkNoseCovering 
         Caption         =   "Nose Covering"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   2100
         Width           =   1740
      End
      Begin VB.CheckBox chkLipCovering 
         Caption         =   "Lip Covering"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   1800
         Width           =   1665
      End
      Begin VB.CheckBox chkRightEarFlap 
         Caption         =   "Right Ear Flap"
         Height          =   240
         Left            =   1950
         TabIndex        =   11
         Top             =   600
         Width           =   1590
      End
      Begin VB.CheckBox chkLeftEarFlap 
         Caption         =   "Left Ear Flap"
         Height          =   240
         Left            =   150
         TabIndex        =   10
         Top             =   600
         Width           =   1890
      End
   End
   Begin VB.Frame frmDesignChoices 
      Caption         =   "Design Choices"
      Height          =   2715
      Left            =   75
      TabIndex        =   43
      Top             =   1275
      Width           =   2640
      Begin VB.OptionButton optContouredChinCollar 
         Caption         =   "Contoured Chin Collar"
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   1800
         Width           =   2190
      End
      Begin VB.OptionButton optChinCollar 
         Caption         =   "Chin Collar"
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   1500
         Width           =   1440
      End
      Begin VB.OptionButton optModifiedChinStrap 
         Caption         =   "Modified Chin Strap"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   1200
         Width           =   2040
      End
      Begin VB.OptionButton optChinStrap 
         Caption         =   "Chin Strap"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   900
         Width           =   1590
      End
      Begin VB.OptionButton optHeadBand 
         Caption         =   "Head Band"
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   2100
         Width           =   1815
      End
      Begin VB.OptionButton optOpenFaceMask 
         Caption         =   "Open Face Mask"
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   2220
      End
      Begin VB.OptionButton optFaceMask 
         Caption         =   "Face Mask"
         Height          =   240
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.Frame frmPatientDetails 
      Caption         =   "Patient Details"
      Height          =   1140
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6765
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   5820
         TabIndex        =   42
         Top             =   675
         Width           =   810
      End
      Begin VB.TextBox txtUnits 
         Height          =   285
         Left            =   4275
         TabIndex        =   41
         Top             =   675
         Width           =   945
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   5925
         TabIndex        =   39
         Top             =   300
         Width           =   690
      End
      Begin VB.TextBox txtFileNo 
         Height          =   285
         Left            =   4275
         TabIndex        =   38
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txtDiagnosis 
         Height          =   285
         Left            =   1050
         TabIndex        =   40
         Top             =   675
         Width           =   2415
      End
      Begin VB.TextBox txtPatientName 
         Height          =   285
         Left            =   825
         TabIndex        =   37
         Top             =   300
         Width           =   2640
      End
      Begin VB.Label lblSex 
         Caption         =   "Sex"
         Height          =   240
         Left            =   5400
         TabIndex        =   50
         Top             =   690
         Width           =   465
      End
      Begin VB.Label lblUnits 
         Caption         =   "Units"
         Height          =   240
         Left            =   3600
         TabIndex        =   49
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lblAge 
         Caption         =   "Age"
         Height          =   240
         Left            =   5475
         TabIndex        =   48
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblFileNo 
         Caption         =   "File No"
         Height          =   240
         Left            =   3600
         TabIndex        =   47
         Top             =   330
         Width           =   690
      End
      Begin VB.Label lblDiagnosis 
         Caption         =   "Diagnosis"
         Height          =   240
         Left            =   150
         TabIndex        =   46
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblPatientName 
         Caption         =   "Patient"
         Height          =   240
         Left            =   150
         TabIndex        =   45
         Top             =   300
         Width           =   690
      End
   End
End
Attribute VB_Name = "HeadNeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   HeadNeck
'File:
'Purpose:   To draw face masks and chin straps
'
'Version:   3.00
'Date:      95
'Author:    Cieran McCavanagh
'           Gary George
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Ported to VB5
'
'Notes:-
    Option Explicit
    
    
    ' Declare arrays
    Dim FaceMaskChartA(1 To 3, 1 To 25) As Double
    Dim FaceMaskChart2(1 To 3, 1 To 16) As Double
    Dim FaceMaskChart3(1 To 3, 1 To 47) As Double

    Dim TopArcAdj(1 To 19) As Double
    Dim TopArcOpp(1 To 19) As Double
    Dim TopArcHyp(1 To 19) As Double

    Dim TopRightArcAdjIn(1 To 19) As Double
    Dim TopRightArcAdjOut(1 To 19) As Double
    Dim TopRightArcOppIn(1 To 19) As Double
    Dim TopRightArcOppOut(1 To 19) As Double
   
    Dim RightArcAdjUp(1 To 20)  As Double
    Dim RightArcAdjDown(1 To 20)  As Double
    Dim RightArcOppUp(1 To 20)  As Double
    Dim RightArcOppDown(1 To 20)  As Double
    Dim RightArcOpenAdjUp(1 To 20)  As Double

    Dim LeftArcAdjIn(1 To 19)  As Double
    Dim LeftArcAdjOut(1 To 19)  As Double
    Dim LeftArcOppIn(1 To 19)  As Double
    Dim LeftArcOppOut(1 To 19)  As Double
    Dim LeftArcHyp(1 To 19)  As Double

    Dim BotFaceAdj(1 To 19) As Double
    Dim BotFaceOpp(1 To 19) As Double
    
    Dim NeckRightAdj(1 To 19) As Double
    Dim NeckRightOpp(1 To 19) As Double
    
    Dim ChinProfileX(1 To 15) As Double
    Dim ChinProfileY(1 To 15) As Double

    Dim LipProfileX(1 To 30) As Double
    Dim LipProfileY(1 To 30) As Double

    Dim OpenFaceProfileX(1 To 12) As Double
    Dim OpenFaceProfileY(1 To 12) As Double

    Dim NoseTipLenAdj(1 To 50) As Double
    Dim NoseTipLenOpp(1 To 50) As Double
    Dim NoseTipWidthAdj(1 To 50) As Double
    Dim NoseTipWidthOpp(1 To 50) As Double

    Dim ChinStrapFrontX(1 To 8) As Double
    Dim ChinStrapFrontY(1 To 8) As Double
    Dim ChinStrapBackX(1 To 19) As Double
    Dim ChinStrapBackY(1 To 19) As Double
    Dim ChinStrapChinX(1 To 13) As Double
    Dim ChinStrapChinY(1 To 13) As Double
    Dim ChinStrapLipX(1 To 25) As Double
    Dim ChinStrapLipY(1 To 25) As Double
    Dim xyChinTopCen As xy
    Dim xyChinBotCen As xy
    Dim g_nChinBotRadius As Double
    Dim g_nChinTopRadius As Double
    Dim xyNoseDiagTop As xy, xyNoseDiagBot As xy  'Used DrawNose & DrawEyeFlap
        
'    '* Windows API Functions Declarations
'    Private Declare Function GetWindow Lib "User" (ByVal hwnd%, ByVal wCmd%) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd%, ByVal lpString$, ByVal aint%) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd%) As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
   
    'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

Private Function Arccos(X As Double) As Double
     Arccos = Atn(-X / Sqr(-X * X + 1)) + 1.5708
End Function

Private Sub chkLeftEarClosed_Click()

    ' if the left ear is closed, then
    ' the Left Ear Flap is disabled
    ' Else it is enabled

    If chkLeftEarClosed.Value = 1 Then
        chkLeftEarFlap.Value = 0
        chkLeftEarFlap.Enabled = False
    ElseIf chkLeftEarClosed.Value = 0 Then
        chkLeftEarFlap.Enabled = True
    End If

End Sub

Private Sub chkLeftEarFlap_Click()

    ' if the Left Ear Flap is included, then
    ' the Left Ear Closed is disabled
    ' Else it is enabled

    If chkLeftEarFlap.Value = 1 Then
        chkLeftEarClosed.Value = 0
        chkLeftEarClosed.Enabled = False
    ElseIf chkLeftEarFlap.Value = 0 Then
        chkLeftEarClosed.Enabled = True
    End If

End Sub

Private Sub chkLining_Click()

    ' if Lining is Chosen, then the
    ' Open Head Mask check box is disabled
    ' else it is ensbled

    If chkLining.Value = 1 Then
        chkOpenHeadMask.Enabled = False
        chkOpenHeadMask.Value = 0
    ElseIf chkLining.Value = 0 Then
        chkOpenHeadMask.Enabled = True
    End If


End Sub

Private Sub chkLipStrap_Click()

    ' The Lip Covering is only available if there is a Lip Strap with the Following:
    ' an Open Face Mask, a Chin Strap, or a Modified Chin Strap
    ' If there is no lipstrap then the Lip Covering is disabled

'GG-19.Sep.96
'    If (optOpenFaceMask.Value = True Or optChinStrap.Value = True Or optModifiedChinStrap.Value = True) And chkLipStrap.Value = 1 Then
'        chkLipCovering.Enabled = True
'    ElseIf (optOpenFaceMask.Value = True Or optChinStrap.Value = True Or optModifiedChinStrap.Value = True) And chkLipStrap.Value = 0 Then
'        chkLipCovering.Enabled = False
'        chkLipCovering.Value = 0
'    End If

End Sub

Private Sub chkOpenHeadMask_Click()

    ' if an Open Head Mask is Chosen, then the
    ' Lining check box is disabled
    ' else it is ensbled

    If chkOpenHeadMask.Value = 1 Then
        chkLining.Enabled = False
        chkLining.Value = 0
    ElseIf chkOpenHeadMask.Value = 0 Then
        chkLining.Enabled = True
    End If

End Sub

Private Sub chkRightEarClosed_Click()

    ' if the Right ear is closed, then
    ' the Right Ear Flap is disabled
    ' Else it is enabled

    If chkRightEarClosed.Value = 1 Then
        chkRightEarFlap.Value = 0
        chkRightEarFlap.Enabled = False
    ElseIf chkRightEarClosed.Value = 0 Then
        chkRightEarFlap.Enabled = True
    End If

End Sub

Private Sub chkRightEarFlap_Click()

    ' if the Right Ear Flap is included, then
    ' the Right Ear Closed is disabled
    ' Else it is enabled

    If chkRightEarFlap.Value = 1 Then
        chkRightEarClosed.Value = 0
        chkRightEarClosed.Enabled = False
    ElseIf chkRightEarFlap.Value = 0 Then
        chkRightEarClosed.Enabled = True
    End If

End Sub

Private Sub chkVelcro_Click()
    
    If chkVelcro.Value = 1 Then
        chkZipper.Enabled = False
        chkZipper.Value = 0
    ElseIf chkVelcro.Value = 0 Then
        chkZipper.Enabled = True
    End If

End Sub

Private Sub chkZipper_Click()
    
    If chkZipper.Value = 1 Then
        chkVelcro.Enabled = False
        chkVelcro.Value = 0
    ElseIf chkZipper.Value = 0 Then
        chkVelcro.Enabled = True
    End If

End Sub

Private Sub cmdClose_Click()

    ' This Procedure Compares the Values on the Form with
    ' the Values last saved.  If changes have been made, the
    ' option is given to save the changes before closing

    Dim Response%
    Dim DataString$, MeasurementString$
    Dim OptionChoice$, Options$

    ' Check For any Changes
    MeasurementString = FN_SetMeasurementString()

    DataString = FN_ModificationString()
    Options = FN_ChooseDesign()
    DataString = Options & DataString

    If (DataString = txtData And MeasurementString = txtMeasurements) Or (txtData = "txtData" And txtMeasurements = "txtMeasurements") Then
        ' no changes
        End
    Else
        ' changes have been made, find out what to do?
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "Head & Neck Details")
        
        ' Determine response of message box
        If Response = 6 Then
            ' Yes button selected
            txtData = DataString
            txtMeasurements = MeasurementString
            txtfNum = FN_OpenSave("c:\jobst\draw.d", txtPatientName, txtFileNo)
            PR_SaveDetails
            Close txtfNum
            AppActivate fnGetDrafixWindowTitleText()
            SendKeys "@c:\jobst\draw.d{enter}", True
            End
        ElseIf Response = 7 Then
            ' No button selected
            End
        ElseIf Response = 2 Then
            ' Cancel button selected
            Exit Sub
        End If

    End If

End Sub

Private Sub cmdDraw_Click()

    ' This Procedure Checks that the required values have been
    ' entered, It produces code for the required drawing, draws the
    ' design, and saves the results to the Head and Neck symbol

    Dim DataString$, OptionChoice$
    Dim HeadBand#, NeckDepth#, CatalogueNo$


    ' Check Correct Values have been entered for choices
    
    ' Check for Fabric type
    If cboFabric.Text = "" Then
        MsgBox "Fabric Type Has Not Been Entered", 48, "Head & Neck Details"
        cboFabric.SetFocus
        Exit Sub
    End If
    
    ' Check for Extension Collar Values
    If optContouredChinCollar.Value = True Or optChinCollar.Value = True Then
        If Val(txtCircOfNeck) <= 0 Then
            MsgBox "Invalid or No Measurement entered for CIRC of Neck", 48, "Head & Neck Details"
            txtCircOfNeck.SetFocus
            Exit Sub
        End If
    End If

    ' Check Neck Depth
    Dim Response%
    If txtThroatToSternal <> "" Then
        NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
        If NeckDepth < 1.125 Then
            Response = MsgBox("Neck Depth is less than the specification minimum", 1, "Head & Neck Details")
            If Response = 2 Then
                txtThroatToSternal.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    ' Check For Face Mask Values
    If optFaceMask.Value = True Or optOpenFaceMask.Value = True Or optChinStrap.Value = True Or optModifiedChinStrap.Value = True Then
        If Val(txtChinToMouth) <= 0 Then
            MsgBox "Invalid or No Measurement entered for Chin To Mouth", 48, "Head & Neck Details"
            txtChinToMouth.SetFocus
            Exit Sub
        End If
        If Val(txtCircEyeBrow) <= 0 Then
            MsgBox "Invalid or No Measurement entered for CIRC above Eyebrow", 48, "Head & Neck Details"
            txtCircEyeBrow.SetFocus
            Exit Sub
        End If
        If Val(txtCircChinAngle) <= 0 Then
            MsgBox "Invalid or No Measurement entered for CIRC of Head at Chin Angle", 48, "Head & Neck Details"
            txtCircChinAngle.SetFocus
            Exit Sub
        End If
        If Val(txtCircOfNeck) <= 0 Then
            MsgBox "Invalid or No Measurement entered for CIRC of Neck", 48, "Head & Neck Details"
            txtCircOfNeck.SetFocus
            Exit Sub
        End If
        If chkNoseCovering.Value = 1 Then
            If Val(txtTipOfNose) <= 0 Then
                MsgBox "Invalid or No Measurement entered for Across Tip of Nose", 48, "Head & Neck Details"
                txtTipOfNose.SetFocus
                Exit Sub
            End If
            If Val(txtLengthOfNose) <= 0 Then
                MsgBox "Invalid or No Measurement entered for Length of Nose", 48, "Head & Neck Details"
                txtLengthOfNose.SetFocus
                Exit Sub
            End If
        End If
    End If
    txtOptionChoice = FN_ChooseDesign()
    
    ' Set Data String, the data string contains the design choice and modifications
    DataString = FN_ModificationString()
    txtData = txtOptionChoice & DataString
    
    ' Set Measurements String, the Measurements string contains the dimensions of the patients head
    txtMeasurements = FN_SetMeasurementString()

    ' Select the required template radius for patient
    If optContouredChinCollar.Value = False And optChinCollar.Value = False Then
        PR_SelectTemplateRadius
    End If

    ' Check For HeadBand Values
    If optHeadBand.Value = True Then
        If txtHeadBandDepth = "" Then
            MsgBox "Invalid or No Measurement entered for Head Band Depth", 48, "Head & Neck Details"
            txtHeadBandDepth.SetFocus
            Exit Sub
        End If
        If Val(txtCircEyeBrow) <= 0 Then
            MsgBox "Invalid or No Measurement entered for CIRC above Eyebrow", 48, "Head & Neck Details"
            txtCircEyeBrow.SetFocus
            Exit Sub
        End If
        HeadBand = FN_CmToInches(txtUnits, txtHeadBandDepth)
        If HeadBand > TopArcAdj(txtRadiusNo) Then
            MsgBox "Head Band depth is too Large.", 48, "Head Band"
            Exit Sub
        End If
    End If


If chkNoseCovering.Value = 1 Then
    ' Check Values have Been Entered
    If (Val(txtTipOfNose) <= 0) Then
        MsgBox "No value has been entered for distance Across Tip of Nose.", 48, "Nose Covering"
        txtTipOfNose.SetFocus
    End If
    If (Val(txtLengthOfNose) <= 0) Then
        MsgBox "No value has been entered for the Length of Nose.", 48, "Nose Covering"
        txtLengthOfNose.SetFocus
    End If
End If

    ' Open file for drawing code
    txtfNum = FN_Open("c:\jobst\draw.d", txtPatientName, txtFileNo)
    PR_SetLayer "TemplateLeft"
    
    ' Draw relevant design option
    PR_FaceMask
    PR_OpenFaceMask
    PR_DrawHeadBand
    PR_DrawChinStrap
    PR_DrawChinCollar
    
    ' Save details to HeadNeck symbol
    PR_SaveDetails
    
    ' Close drawing code file
    Close txtfNum
    
    ' Activate drafix file and draw design
    AppActivate fnGetDrafixWindowTitleText()
    SendKeys "@c:\jobst\draw.d{enter}", True
    End


End Sub

Private Function FN_CalcAngle(xyStart As xy, xyEnd As xy) As Double
   
   'Function to return the angle between two points in degrees
   'in the range 0 - 360
   'Zero is always 0 and is never 360
    
    Dim X, Y, rAngle#
    Dim PI#
    PI = 3.141592654
    
    X = xyEnd.X - xyStart.X
    Y = xyEnd.Y - xyStart.Y

   'Horizomtal
    If X = 0 Then
        If Y > 0 Then
            FN_CalcAngle = 90
        Else
            FN_CalcAngle = 270
        End If
        Exit Function
    End If
    
   'Vertical (avoid divide by zero later)
    If Y = 0 Then
        If X > 0 Then
            FN_CalcAngle = 0
        Else
            FN_CalcAngle = 180
        End If
        Exit Function
    End If

   'All other cases
    rAngle = Atn(Y / X) * (180 / PI)         'Convert to degrees
    
    If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
    
    If Y > 0 Then
        FN_CalcAngle = rAngle
    Else
        FN_CalcAngle = rAngle + 180
    End If

End Function

Private Function FN_CalcCirCirInt(xyCen1 As xy, nRad1 As Double, xyCen2 As xy, nRad2 As Double, xyInt1 As xy, xyInt2 As xy) As Integer
   'Function that will return
   '    TRUE    if two circles intersect
   '    FALSE   if two circles don't intersect
   '
   'The intersection points are returned in the values
   '
   '    xyInt1 & xyInt2
   '
   'with intersection with lowest X value as xyInt1
   '
    Dim nLength As Double, aAngle  As Double, aTheta As Double, nCosTheta  As Double
    Dim xyTmp As xy

    FN_CalcCirCirInt = False

   'Check that the circles can intersect
   'It is a theorem of plane geometry that no three real numbers
   'a, b and c can be the lenghts of the sides of a triangle unless
   'the sum of any two is greater than the third.
   'We use this as our main test of possible intersection, we also check for silly
   'data.

   'Test for silly data
    If xyCen1.X = xyCen2.X And xyCen1.Y = xyCen2.Y Then Exit Function
    If nRad1 <= 0 Or nRad2 <= 0 Then Exit Function

   'Test for intersection
    nLength = FN_CalcLength(xyCen1, xyCen2)

    If (nLength + nRad1 < nRad2) Or (nLength + nRad2 < nRad1) Or (nRad1 + nRad2 < nLength) Then
        Exit Function
    Else
        FN_CalcCirCirInt = True
    End If

   'Calculate intesection points
   '
   'Special case where circles touch (ie Intersect at one point only)
    
    
   'Angle between centers
   'Note: Length between centers from above
    aAngle = FN_CalcAngle(xyCen1, xyCen2)

   'Get angle w.r.t line between centers to the intersection point
   'use cosine rule
   '
    nCosTheta = -((nRad2 ^ 2 - (nLength ^ 2 + nRad1 ^ 2)) / (2 * nLength * nRad1))

    aTheta = Atn(-nCosTheta / Sqr(-(nCosTheta ^ 2) + 1)) + 1.5708

    aTheta = aTheta * (180 / PI) 'convert to degrees

    aAngle = aAngle - aTheta
    PR_CalcPolar xyCen1, aAngle, nRad1, xyInt1

    aAngle = aAngle + aTheta
    PR_CalcPolar xyCen1, aAngle, nRad1, xyInt2

    If xyInt2.X < xyInt1.X Then
        xyTmp = xyInt1
        xyInt1 = xyInt2
        xyInt2 = xyTmp
    End If

    
End Function

Private Function FN_CalcLength(xyStart As xy, xyEnd As xy) As Double
   'Fuction to return the length between two points
   'Greatfull thanks to Pythagorus

    FN_CalcLength = Sqr((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

End Function

Private Function FN_ChooseDesign()

' this function returns the code word for the chosen design
    
    If optFaceMask.Value = True Then
        FN_ChooseDesign = "RFM"
    ElseIf optHeadBand.Value = True Then
        FN_ChooseDesign = "RHB"
    ElseIf optOpenFaceMask.Value = True Then
        FN_ChooseDesign = "OFM"
    ElseIf optChinStrap.Value = True Then
        FN_ChooseDesign = "RCS"
    ElseIf optModifiedChinStrap.Value = True Then
        FN_ChooseDesign = "MCS"
    ElseIf optChinCollar.Value = True Then
        FN_ChooseDesign = "RCC"
    ElseIf optContouredChinCollar.Value = True Then
        FN_ChooseDesign = "CCC"
    End If

End Function

Private Static Function FN_CirLinInt(xyStart As xy, xyEnd As xy, xyCen As xy, nRad As Double, xyInt As xy) As Integer
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
        nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
        If nC < 0 Then
            FN_CirLinInt = False     'no roots
            Exit Function
        End If
        nSign = 1                               'test each root
        While nSign > -2
            nRoot = xyCen.X + Sqr(nC) * nSign
            If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                xyInt.X = nRoot
                xyInt.Y = xyStart.Y
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
            nRoot = xyCen.Y + Sqr(nC) * nSign
            If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                xyInt.Y = nRoot
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
        nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X)  'Slope
        nK = xyStart.Y - nM * xyStart.X                     'Y-Axis intercept
        nA = (1 + nM ^ 2)
        nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
        nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
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
                xyInt.Y = nM * nRoot + nK
                FN_CirLinInt = True
                Exit Function   'Return first root found
            End If
            nSign = nSign - 2
        Wend
        FN_CirLinInt = False 'Should never get to here
    End If
    FN_CirLinInt = False

End Function

Private Function FN_CmToInches(flag, cm)

    ' This function converts cm to inches

    Dim inch#, remainder#, temp#, temprem#, rnder#
    Dim Value#, realrem#

    If flag = "cm" Then                     ' value is in cm
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
        FN_CmToInches = Abs(inch)
    ElseIf flag = "inches" Then
        Value = cm
        If Value >= 1 Then
            temp = Int(Value)
            remainder = Value - temp
            If remainder > 0.75 Then
                MsgBox "Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details"
                FN_CmToInches = "-1"
                Exit Function
            Else
                remainder = remainder * 10
                remainder = remainder * 0.125
                FN_CmToInches = Abs(temp + remainder)
            End If
        Else
            temp = Int(Value)
            remainder = Value - temp
            If remainder > 0.75 Then
                MsgBox "Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details"
                FN_CmToInches = "-1"
                Exit Function
            Else
                remainder = remainder * 10
                remainder = remainder * 0.125
                FN_CmToInches = Abs(remainder)
            End If
        End If
    End If

End Function


Private Function FN_InchesToCm(flag, inch)

    ' This function converts inches to cm

    Dim cm, remainder, temp, temprem, rnder
    If flag = "cm" Then
        cm = Val(inch) * 2.54
        temp = Int(cm)
        remainder = (cm - temp)
        temprem = remainder * 10
        temprem = Int(temprem)
        rnder = (remainder * 10) - temprem
        If rnder > 0.5 Then
            temprem = temprem + 1
        End If
        temprem = temprem / 10
        cm = temp + temprem
        FN_InchesToCm = cm
    ElseIf flag = "inches" Then
        temp = Int(inch)
        remainder = inch - temp
        rnder = remainder / 0.125
        rnder = Int(rnder)
        cm = (rnder / 10) + temp
        FN_InchesToCm = cm
    End If

End Function

Private Function FN_InchesToText(nInches#) As String
   
   'Function returns a decimal value in inches as a string
   '
    Dim nPrecision#, nDec#
    Dim iInt%, iEighths%
    Dim sString$
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
        sString = "  "
    End If
    If iEighths <> 0 Then
        Select Case iEighths
        Case 2, 6
            sString = sString + " " + LTrim$(Str$(iEighths / 2)) + "/4"
        Case 4
            sString = sString + " " + "1/2"
        Case Else
            sString = sString + " " + LTrim$(Str$(iEighths)) + "/8"
        End Select
    Else
        sString = sString
    End If
       
   'Return formatted string
    FN_InchesToText = sString

End Function

Private Function FN_LinLinInt(xyLine1Start As xy, xyLine1End As xy, xyLine2Start As xy, xyLine2End As xy, xyInt As xy) As Integer
   'Function:
   '       BOOLEAN = FN_LinLinInt( xyLine1Start, xyLine1End, xyLine2Start, xyLine2End, xyInt);
   'Parameters:
   '       xyLine1Start = xyLine1Start.X, xyLine1Start.Y
   '       xyLine1End = xyLine1End.X, xyLine1End.Y
   '       xyLine2Start = xyLine2Start.X, xyLine2Start.Y
   '       xyLine2End = xyLine2End.X, xyLine2End.Y
   '
   'Returns:
   '       True if intersection found and lies on the line
   '       False if no intesection
   '       xyInt =  intersection
   '
    Dim nK1#, nK2#, nM1#, nM2, nSlope1#, nSlope2, nCase#, nY#, nX#

   'Initialy false
    FN_LinLinInt = False

   'Calculate slope of lines
    nCase = 0
    nSlope1 = FN_CalcAngle(xyLine1Start, xyLine1End)
    If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
    If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2
    
    nSlope2 = FN_CalcAngle(xyLine2Start, xyLine2End)
    If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
    If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

    Select Case nCase
     
     Case 0
       'Both lines are Non-Orthogonal Lines
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nM2 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        If (nM1 = nM2) Then Exit Function               'Parallel lines
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)   'Y-Axis intercept
        nK2 = xyLine2Start.Y - (nM2 * xyLine2Start.X)   'Y-Axis intercept
        If (nK1 = nK2) Then Exit Function
       'Find X
        nX = (nK2 - nK1) / (nM1 - nM2)
       'Find Y
        nY = (nM1 * nX) + nK1

     Case 1
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X)                           'Y-Axis intercept
        nY = xyLine1Start.Y
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1

     Case 2
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
        nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X)                           'Y-Axis intercept
        nX = xyLine1Start.X
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1

     Case 4
       'Line 1 is Horizontal or Line 2 is horizontal
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)                           'Y-Axis intercept
        nY = xyLine2Start.Y
        
       'Solve for X at the given Y value
        nX = (nY - nK1) / nM1
    
     Case 5
       'Parallel orthogonal lines, no intersection possible
        Exit Function

     Case 6
       'Line1 is Vertical and the Line2 is Horizontal
        nX = xyLine1Start.X
        nY = xyLine2Start.Y

     Case 8
       'Line 1 is Vertical or Line 2 is Vertical
        nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
        nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X)                           'Y-Axis intercept
        nX = xyLine2Start.X
       'Solve for Y at the given X value
        nY = (nM1 * nX) + nK1
     
     Case 9
       'Line1 is Horizontal and the Line2 is Vertical
        nX = xyLine2Start.X
        nY = xyLine1Start.Y

     Case 10
       'Parallel orthogonal lines, no intersection possible
        Exit Function
     
     Case Else
        Exit Function
    
    End Select

   'Ensure that the points X and Y are on the lines
    xyInt.X = nX
    xyInt.Y = nY

   'Line 1
    If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
    If (nY < min(xyLine1Start.Y, xyLine1End.Y) Or nY > max(xyLine1Start.Y, xyLine1End.Y)) Then Exit Function
   
   'Line 2
    If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
    If (nY < min(xyLine2Start.Y, xyLine2End.Y) Or nY > max(xyLine2Start.Y, xyLine2End.Y)) Then Exit Function
   
    FN_LinLinInt = True
   

End Function

Private Function FN_ModificationString()

    ' This Function sets a string containing the chosen modifications

    Dim Modifications 'As String

    ' Set Data String
    Modifications = chkLeftEarFlap.Value
    Modifications = Modifications & chkRightEarFlap.Value
    Modifications = Modifications & chkNeckElastic.Value
    Modifications = Modifications & chkLeftEyeFlap.Value
    Modifications = Modifications & chkRightEyeFlap.Value
    Modifications = Modifications & chkOpenHeadMask.Value
    Modifications = Modifications & chkLipStrap.Value
    Modifications = Modifications & chkLipCovering.Value
    Modifications = Modifications & chkNoseCovering.Value
    Modifications = Modifications & chkZipper.Value
    Modifications = Modifications & chkLining.Value
    Modifications = Modifications & chkEarSize.Value
    Modifications = Modifications & chkRightEarClosed.Value
    Modifications = Modifications & chkLeftEarClosed.Value
    Modifications = Modifications & chkEyes.Value
    Modifications = Modifications & chkVelcro.Value

    FN_ModificationString = Modifications

End Function

Private Function FN_Open(sDrafixFile$, sName As Variant, sPatientFile As Variant) As Integer
   
   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ
   '
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    txtfNum = FreeFile
    Open sDrafixFile For Output As txtfNum
    FN_Open = txtfNum
    
   'Initialise patient globals
    txtFileNo = sPatientFile
    txtPatientName = sName
    
   'Globals to reduced drafix code written to file
    txtCurrentLayer = ""
    txtCurrTextHt = 0.125
    txtCurrTextAspect = 0.6
    txtCurrTextHorizJust = 1    'Left
    txtCurrTextVertJust = 32    'Bottom
    txtCurrTextFont = 0         'BLOCK

   'Write header information etc. to the DRAFIX macro file
   
    Print #txtfNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #txtfNum, "//Patient - "; txtPatientName; CC; " "; txtFileNo; CC; " Head & Neck "
    Print #txtfNum, "//by Visual Basic"

    'Define DRAFIX variables
    Print #txtfNum, "HANDLE hLayer, hFacemask, hTitle, hChan, hEnt;"
    Print #txtfNum, "XY     xyTitleBoxOrigin,xyTitleOrigin, xyStart,TitleOrigin, xyTitleScale, xyOrigin;"
    Print #txtfNum, "ANGLE  aTitleAngle;"
    Print #txtfNum, "STRING sTitleName, sFileNo;"

   'Set up data base fields
    Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ID"; QCQ; "string"; QQ; ");"
    Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "HeadNeck"; QCQ; "string"; QQ; ");"
    Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "units"; QCQ; "string"; QQ; ");"
    Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "Data"; QCQ; "string"; QQ; ");"
    Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "WorkOrder"; QCQ; "string"; QQ; ");"

   'Text data
    Print #txtfNum, "SetData("; QQ; "TextHorzJust"; QC; txtCurrTextHorizJust; ");"
    Print #txtfNum, "SetData("; QQ; "TextVertJust"; QC; txtCurrTextVertJust; ");"
    Print #txtfNum, "SetData("; QQ; "TextHeight"; QC; txtCurrTextHt; ");"
    Print #txtfNum, "SetData("; QQ; "TextAspect"; QC; txtCurrTextAspect; ");"
    Print #txtfNum, "SetData("; QQ; "TextFont"; QC; txtCurrTextFont; ");"
   
   'Get Start point
    Print #txtfNum, "GetUser ("; QQ; "xy"; QCQ; "Indicate Start Point"; QC; "&xyStart);"
   
   'Display Hour Glass symbol
    Print #txtfNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Drawing"; QQ; ");"
   
   'Clear user selections etc
    Print #txtfNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #txtfNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
   
   'Place a marker at the start point later use
    PR_SetLayer "Construct"
    Print #txtfNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "cross"; QC; "xyStart"; CC; "0.125);"
    PR_AddEntityID "OriginMarker"
    ' Update data base fields
    Print #txtfNum, "SetDBData(hEnt,"; QQ; "Fabric"; QCQ; cboFabric.Text; QQ; ");"
    Print #txtfNum, "SetDBData(hEnt,"; QQ; "WorkOrder"; QCQ; txtWorkOrder; QQ; ");"
    Print #txtfNum, "SetDBData(hEnt,"; QQ; "Data"; QCQ; txtData; QQ; ");"
    Print #txtfNum, "SetDBData(hEnt,"; QQ; "units"; QCQ; txtUnits; QQ; ");"

   'Set values for use futher on by other macros
    Print #txtfNum, "xyOrigin = xyStart"; ";"
    Print #txtfNum, "sFileNo = "; QQ; txtFileNo; QQ; ";"

End Function


Private Function FN_OpenSave(sDrafixFile$, sName As Variant, sPatientFile As Variant)
   
   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ
   '
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    txtfNum = FreeFile
    Open sDrafixFile For Output As txtfNum
    FN_OpenSave = txtfNum
    
   'Initialise patient globals
    txtFileNo = sPatientFile
    txtPatientName = sName
    
   'Globals to reduced drafix code written to file
    txtCurrentLayer = ""
    txtCurrTextHt = 0.125
    txtCurrTextAspect = 0.6
    txtCurrTextHorizJust = 1    'Left
    txtCurrTextVertJust = 32    'Bottom
    txtCurrTextFont = 0         'BLOCK

   'Write header information etc. to the DRAFIX macro file
   
    Print #txtfNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #txtfNum, "//Patient - "; txtPatientName; CC; " "; txtFileNo; CC; " Head & Neck "
    Print #txtfNum, "//by Visual Basic"

    'Define DRAFIX variables
    Print #txtfNum, "HANDLE hLayer, hFacemask, hTitle, hChan, hEnt;"
    Print #txtfNum, "XY     xyTitleBoxOrigin,xyTitleOrigin, xyStart,TitleOrigin, xyTitleScale, xyOrigin;"
    Print #txtfNum, "ANGLE  aTitleAngle;"
    Print #txtfNum, "STRING sTitleName, sFileNo;"

   'Set up ID  data base field
   ' Print #txtfNum, "Table("; QQ; "add"; QCQ; "field"; QCQ; "ID"; QCQ; "string"; QQ; ");"

   'Text data
    Print #txtfNum, "SetData("; QQ; "TextHorzJust"; QC; txtCurrTextHorizJust; ");"
    Print #txtfNum, "SetData("; QQ; "TextVertJust"; QC; txtCurrTextVertJust; ");"
    Print #txtfNum, "SetData("; QQ; "TextHeight"; QC; txtCurrTextHt; ");"
    Print #txtfNum, "SetData("; QQ; "TextAspect"; QC; txtCurrTextAspect; ");"
    Print #txtfNum, "SetData("; QQ; "TextFont"; QC; txtCurrTextFont; ");"
   
   
   
   'Clear user selections etc
    Print #txtfNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #txtfNum, "UserSelection ("; QQ; "update"; QQ; ");"
    Print #txtfNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
   

   'Set values for use futher on by other macros
    Print #txtfNum, "sFileNo = "; QQ; txtFileNo; QQ; ";"


End Function


Private Function FN_SetDetailField(Field As Variant)

    ' This function sets the detail field saved
    ' with the drawing
    
    Dim pos, dig, bit, Result

    If Len(Field) = 4 Then
        pos = InStr(Field, ".")
        If pos <> 0 Then
            dig = Int(Field)
            bit = Right$(Field, 1)
            Result = dig & bit
        End If
        FN_SetDetailField = Result
    ElseIf Len(Field) = 3 Then
        pos = InStr(Field, ".")
        If pos <> 0 Then
            dig = Int(Field)
            bit = Right$(Field, 1)
            Result = "0" & dig & bit
        End If
        FN_SetDetailField = Result
    ElseIf Len(Field) = 2 Then
        Dim tempval#
        tempval = Val(Field)
        If tempval > 1 Then
            Result = Field
            FN_SetDetailField = Result & "0"
        Else
            Result = Right$(Field, 1)
            FN_SetDetailField = "00" & Result
        End If
    ElseIf Len(Field) = 1 Then
        Result = Field
        FN_SetDetailField = "0" & Result & "0"
    ElseIf Field = "" Then
        FN_SetDetailField = "000"
    End If

End Function

Private Function FN_SetMeasurementString()

    ' This Function Constructs a String to represent the Measurements

    Dim Rec1, Rec2, Rec3, Rec4, Rec5
    Dim Rec6, Rec7, Rec8, Rec9, Rec10, Rec11

    ' Set Measurements of the Face Mask
    Rec1 = FN_SetDetailField(txtChinToMouth)
    Rec2 = FN_SetDetailField(txtCircEyeBrow)
    Rec3 = FN_SetDetailField(txtCircChinAngle)
    Rec4 = FN_SetDetailField(txtCircOfNeck)
    Rec5 = FN_SetDetailField(txtThroatToSternal)
    Rec6 = FN_SetDetailField(txtTipOfNose)
    Rec7 = FN_SetDetailField(txtLengthOfNose)
    Rec8 = FN_SetDetailField(txtHeadBandDepth)
    Rec9 = FN_SetDetailField(txtLeftEarLength)
    Rec10 = FN_SetDetailField(txtRightEarLength)
    Rec11 = FN_SetDetailField(txtChinCollarMin)
    FN_SetMeasurementString = Rec1 & Rec2 & Rec3 & Rec4 & Rec5 & Rec6 & Rec7 & Rec8 & Rec9 & Rec10 & Rec11

End Function

Private Function fnRoundInches(ByVal nNumber As Double) As Double
   'Function to return the rounded value in decimal inches
   'returns to the nearest eighth (0.125)
   'E.G.
   '    5.67         = 5 inches and 0.67 inches
   '                   0.67 / 0.125 = 5.36 eighths
   '                   5.36 eighths = 5 eighths (rounded to nearest eighth)
   '    5.67         = 5 inches and 5 eighths
   '    5.67         = 5 + ( 5 * 0.125)
   '    5.67         = 6.625 inches
   '

    Dim iInt As Integer, iSign  As Integer
    Dim nPrecision As Double, nDec As Double

   'Return 0 if input is Zero
    If nNumber = 0 Then
       fnRoundInches = 0
       Exit Function
    End If

   'Set precision
    nPrecision = 0.125

   'Break input into components
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)
    nDec = nNumber - iInt

   'Get decimal part in precision units
    If nDec <> 0 Then
        nDec = nDec / nPrecision    'Avoid overflow
    End If
    nDec = round(nDec)

   'Return value
    fnRoundInches = (iInt + (nDec * nPrecision)) * iSign

End Function

Private Sub Form_LinkClose()
   'Disable timer
   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

 Show
    ' Check for previous instance of HeadNeck, if thereis Unload
    If App.PrevInstance Then
        Beep
        MsgBox "The drawing routine for the Head & Neck is already running.", 48, "Head & Neck"
        Unload HeadNeck
        End
    End If

    ' Check for Patient details
    If txtUidMPD = "" Then
        MsgBox "Patient Details have not been entered.", 16, "Head & Neck"
        Unload HeadNeck
        End
    End If

    cboFabric.Text = txtFabric
   
    ' get measurements from txtMeasurements, and place in relevant text boxes
    PR_SeperateMeasurements
  
    ' get Modifications from txtData, and set check boxes
    PR_SeperateModifications
  
    Show

    Screen.MousePointer = 0   ' set mouse pointer to arrow

End Sub

Private Sub Form_Load()

    ' this procedure loads all the required information
    ' to display the form HEADNECK
    
    ' Maintain while loading DDE data
    Hide
    Screen.MousePointer = 11  ' change pointer to hourglass
    ' Reset in form_linkclose

    ' Position to center of screen
    Left = (Screen.Width - HeadNeck.Width) / 2
    Top = (Screen.Height - HeadNeck.Height) / 2
    
    Set MainForm = HeadNeck

    Dim textline$, filenum%

    ' Load Template Arrays, these arrays hold mappings of the
    ' 51508 R1 template, with measurements from the centre
    PR_LoadTopArcArrays
    PR_LoadTopRightArrays
    PR_LoadRightArcArrays
    PR_LoadLeftArcArrays
    PR_LoadBotFaceArrays
    PR_LoadNeckArrays

    ' Load Face Mask Chart Arrays, taken from specifications
    PR_LoadFaceMaskChart1
    PR_LoadFaceMaskChart2
    PR_LoadFaceMaskChart3

    txtUidMPD = ""
    txtUidHN = ""

    g_sPathJOBST = fnPathJOBST()
    ' Load Fabric File
    filenum = FreeFile
    Open g_sPathJOBST & "\fabric.dat" For Input As filenum
    Do While Not EOF(filenum)
        Line Input #filenum, textline
        cboFabric.AddItem textline
    Loop
    Close

   'Start a timer
   'The Timer is disabled in LinkClose
   'If after 10 seconds the timer event will "End" the programme
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 10000  'Approx 10 Seconds
    Timer1.Enabled = True

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

Private Sub optChinCollar_Click()

    ' Initalise available Modifications for Chin Collar

    chkLeftEarFlap.Value = 0
    chkRightEarFlap.Value = 0
    chkNeckElastic.Value = 0
    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkLeftEarClosed.Value = 0
    chkRightEarClosed.Value = 0
    chkOpenHeadMask.Value = 0
    chkLipStrap.Value = 0
    chkLipCovering.Value = 0
    chkNoseCovering.Value = 0
    chkZipper.Value = 0
    chkLining.Value = 0
    chkEarSize.Value = 0
    chkEyes.Value = 0
    chkVelcro.Value = 0
    
    chkEyes.Enabled = False
    chkLeftEarFlap.Enabled = False
    chkRightEarFlap.Enabled = False
    chkNeckElastic.Enabled = False
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarClosed.Enabled = False
    chkRightEarClosed.Enabled = False
    chkOpenHeadMask.Enabled = False
    chkLipStrap.Enabled = False
    chkLipCovering.Enabled = False
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = False
    chkLining.Enabled = False
    chkEarSize.Enabled = False
    chkVelcro.Enabled = False

End Sub

Private Sub optChinStrap_Click()
    
    ' Initalise available Modifications for Chin Strap

    chkLeftEarFlap.Value = 0
    chkRightEarFlap.Value = 0
    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkOpenHeadMask.Value = 0
'    If chkLipStrap.Value = 0 Then
        chkLipCovering.Value = 0
'    End If
    chkNoseCovering.Value = 0
    chkZipper.Value = 0
    chkLining.Value = 0
    chkEarSize.Value = 0
    chkLeftEarClosed.Value = 0
    chkRightEarClosed.Value = 0
    chkEyes.Value = 0
    chkVelcro.Value = 0
    
    chkEyes.Enabled = False
    chkLeftEarFlap.Enabled = False
    chkRightEarFlap.Enabled = False
    chkNeckElastic.Enabled = True
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarClosed.Enabled = False
    chkRightEarClosed.Enabled = False
    chkOpenHeadMask.Enabled = False
    chkLipStrap.Enabled = True
 '   If chkLipStrap.Value = 0 Then
 '       chkLipCovering.Enabled = False
 '   Else
        chkLipCovering.Enabled = True
 '   End If
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = False
    chkLining.Enabled = False
    chkEarSize.Enabled = False
    chkVelcro.Enabled = False

End Sub

Private Sub optContouredChinCollar_Click()

    ' Initalise available Modifications for Contoured Chin Collar

    chkLeftEarFlap.Value = 0
    chkRightEarFlap.Value = 0
    chkNeckElastic.Value = 0
    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkLeftEarClosed.Value = 0
    chkRightEarClosed.Value = 0
    chkOpenHeadMask.Value = 0
    chkLipStrap.Value = 0
    chkLipCovering.Value = 0
    chkNoseCovering.Value = 0
    chkZipper.Value = 0
    chkLining.Value = 0
    chkEarSize.Value = 0
    chkEyes.Value = 0
    chkVelcro.Value = 0
    
    chkEyes.Enabled = False
    chkLeftEarFlap.Enabled = False
    chkRightEarFlap.Enabled = False
    chkNeckElastic.Enabled = False
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarClosed.Enabled = False
    chkRightEarClosed.Enabled = False
    chkOpenHeadMask.Enabled = False
    chkLipStrap.Enabled = False
    chkLipCovering.Enabled = False
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = False
    chkLining.Enabled = False
    chkEarSize.Enabled = False
    chkVelcro.Enabled = False

End Sub

Private Sub optFaceMask_Click()

    ' Initalise available Modifications for Face Mask

    chkLipStrap.Value = 0
    chkEyes.Value = 0
    chkLipCovering.Value = 0
    
    chkEyes.Enabled = False
    chkLeftEarFlap.Enabled = True
    chkRightEarFlap.Enabled = True
    chkNeckElastic.Enabled = True
    chkLeftEyeFlap.Enabled = True
    chkRightEyeFlap.Enabled = True
    chkLeftEarClosed.Enabled = True
    chkRightEarClosed.Enabled = True
    chkOpenHeadMask.Enabled = True
    chkLipStrap.Enabled = False
    chkLipCovering.Enabled = True
    chkNoseCovering.Enabled = True
    chkZipper.Enabled = True
    chkLining.Enabled = True
    chkEarSize.Enabled = True
    chkVelcro.Enabled = True

End Sub

Private Sub optHeadBand_Click()

    ' Initalise available Modifications for Head Band

    chkLeftEarFlap.Value = 0
    chkRightEarFlap.Value = 0
    chkNeckElastic.Value = 0
    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkLeftEarClosed.Value = 0
    chkRightEarClosed.Value = 0
    chkOpenHeadMask.Value = 0
    chkLipStrap.Value = 0
    chkLipCovering.Value = 0
    chkNoseCovering.Value = 0
    chkZipper.Value = 0
    chkLining.Value = 0
    chkEarSize.Value = 0
    chkEyes.Value = 0
    chkVelcro.Value = 0
    
    chkEyes.Enabled = False
    chkLeftEarFlap.Enabled = False
    chkRightEarFlap.Enabled = False
    chkNeckElastic.Enabled = False
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarClosed.Enabled = False
    chkRightEarClosed.Enabled = False
    chkOpenHeadMask.Enabled = False
    chkLipStrap.Enabled = False
    chkLipCovering.Enabled = False
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = False
    chkLining.Enabled = False
    chkEarSize.Enabled = False
    chkVelcro.Enabled = False
    
End Sub

Private Sub optModifiedChinStrap_Click()

    ' Initalise available Modifications for Modified Chin Strap

    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkOpenHeadMask.Value = 0
    chkNoseCovering.Value = 0
    chkZipper.Value = 0
    chkLining.Value = 0
    chkEyes.Value = 0
    chkVelcro.Value = 0

    chkEyes.Enabled = False
    chkNeckElastic.Enabled = True
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarFlap.Enabled = True
    chkRightEarFlap.Enabled = True
    chkLeftEarClosed.Enabled = True
    chkRightEarClosed.Enabled = True
    chkOpenHeadMask.Enabled = False
    chkLipStrap.Enabled = True
'    If chkLipStrap.Value = 0 Then
'        chkLipCovering.Enabled = False
'    Else
        chkLipCovering.Enabled = True
'    End If
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = False
    chkLining.Enabled = False
    chkEarSize.Enabled = True
    chkVelcro.Enabled = False

End Sub

Private Sub optOpenFaceMask_Click()

    ' Initalise available Modifications for Open Face Mask

    chkLeftEyeFlap.Value = 0
    chkRightEyeFlap.Value = 0
    chkLipCovering.Value = 0
    chkNoseCovering.Value = 0
    
    chkEyes.Enabled = True
    chkLeftEarFlap.Enabled = True
    chkRightEarFlap.Enabled = True
    chkNeckElastic.Enabled = True
    chkLeftEyeFlap.Enabled = False
    chkRightEyeFlap.Enabled = False
    chkLeftEarClosed.Enabled = True
    chkRightEarClosed.Enabled = True
    chkOpenHeadMask.Enabled = True
    chkLipStrap.Enabled = True
    chkLipCovering.Enabled = False
    chkNoseCovering.Enabled = False
    chkZipper.Enabled = True
    chkLining.Enabled = True
    chkEarSize.Enabled = True
    chkVelcro.Enabled = True

End Sub

Private Sub PR_AddEntityID(sType$)
   
   'To the DRAFIX macro file (given by the global fNum)
   'write the syntax to add to an ENTITY the database information
   'in the DB variable "ID" that will allow the identity of an entity
   'to be retrieved, by other parts of the system.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined.
   '    HANDLE  hEnt
   '
   'Note:-
   '    fNum, CC, QQ, NL are globals initialised by FN_Open
   '
    Dim sID$, sData$

    sID = txtOptionChoice & txtFileNo & sType
    sData = txtData

    Print #txtfNum, "if (hEnt) SetDBData( hEnt," & QQ & "ID" & QQ & CC & QQ; sID & QQ & ");"
    Print #txtfNum, "if (hEnt) SetDBData( hEnt," & QQ & "Data" & QQ & CC & QQ; sData & QQ & ");"

End Sub

Private Sub PR_CalcPolar(xyStart As xy, ByVal nAngle As Double, nLength As Double, xyReturn As xy)
   'Procedure to return a point at a distance and an angle from a given point
   '
    Dim a As Double, B As Double

    Select Case nAngle
        Case 0
            xyReturn.X = xyStart.X + nLength
            xyReturn.Y = xyStart.Y
        Case 180
            xyReturn.X = xyStart.X - nLength
            xyReturn.Y = xyStart.Y
        Case 90
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y + nLength
        Case 270
            xyReturn.X = xyStart.X
            xyReturn.Y = xyStart.Y - nLength
        Case Else
           'Convert from degees to radians
            nAngle = nAngle * PI / 180
            B = Sin(nAngle) * nLength
            a = Cos(nAngle) * nLength
            xyReturn.X = xyStart.X + a
            xyReturn.Y = xyStart.Y + B
        End Select
End Sub

Private Sub PR_ChinStrapDimensions()

    ' this procedure sets dimensions for the chinstrap

    txtCSChintoMouth = FN_CmToInches(txtUnits, txtChinToMouth)
    txtCSNeckCircum = (FN_CmToInches(txtUnits, txtCircOfNeck)) * 0.46
    txtCSChinAngle = (FN_CmToInches(txtUnits, txtCircEyeBrow) + FN_CmToInches(txtUnits, txtCircChinAngle)) * 0.19
    txtCSForeHead = txtCSChinAngle / 2
    txtCSMouthWidth = txtCSChinAngle / 8
    txtCSEarTopHeight = txtCSForeHead - 1.125
    txtCSEarBotHeight = txtCSChintoMouth * 0.72

End Sub

Private Sub PR_DrawArc(xyCen As xy, xyArcStart As xy, xyArcEnd As xy)

    ' this procedure draws an arc between two points

    Dim nRad, nStartAng, nEndAng, nDeltaAng
    
    nRad = FN_CalcLength(xyCen, xyArcStart)
    nStartAng = FN_CalcAngle(xyCen, xyArcStart)
    nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
    nDeltaAng = nEndAng - nStartAng
    
    Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nStartAng); CC; Str$(nDeltaAng); ");"

End Sub

Private Sub PR_DrawChin()

    ' This Procedure Forms The Mouth, Jobst Procedure No.9

    Dim xyMouthLeft As xy, xyMouthRight As xy
    Dim xy1 As xy, xy2 As xy, xy3 As xy, xy4 As xy, xyTmp As xy
    Dim xyBotCen As xy, xyTopCen As xy, xyChinStart As xy
    Dim n1#, n2#, n3#
    Dim aAngle#, aInc#, rAngle#, nValue#, nBotRadius#, nTopRadius#
    Dim GivenNeckDepth#, DefaultNeckDepth#, nNeckOffset#

    Dim Tempx#, TempY#, ChinDist#, ChinStep2#, ChinStep1#
    Dim MouthWidth#, PI#, Beta#
    Dim TopCurve  As xy, BotCurve As xy, xyChinArc As xy
    Dim XA#, YA#, DiagLen#, DiagStep#
    Dim TopRadius#, BotRadius#, ii%
    Dim DiagStart As xy, DiagEnd As xy, Delta#, v%
    Dim ChinHt#, i%, Theta#, j%

    TopRadius = 1.8 / 2.54
    BotRadius = 0.9 / 2.54
    PI = 3.141592654

    ' Check chart2 for mouth width
    For i = 1 To 16
        If FaceMaskChart2(1, i) = Val(txtCircumferenceTotal) Then
            MouthWidth = FaceMaskChart2(2, i)
        End If
    Next i

    ' Set Cordinates and draw Mouth
    If (optOpenFaceMask.Value = True) And chkLipStrap.Value = 0 Then
        Tempx = -(LeftArcAdjOut(txtRadiusNo)) + 0.375
    Else
        Tempx = -(LeftArcAdjOut(txtRadiusNo))
    End If
    TempY = Val(txtMouthHeight)

    PR_MakeXY xyMouthLeft, Tempx, TempY

    Tempx = -(Abs(LeftArcAdjOut(txtRadiusNo)) - MouthWidth)
    PR_MakeXY xyMouthRight, Tempx, TempY
    txtMouthRightX = xyMouthRight.X
    txtMouthRightY = xyMouthRight.Y
    If (optOpenFaceMask.Value = True) And chkEyes.Value = 1 Then
        PR_MakeXY xyMouthRight, Tempx - 0.125, TempY
    End If

    PR_DrawLine xyMouthLeft, xyMouthRight
    PR_AddEntityID "Mouth"


   'Bug fix code added G George 2.Jul.96
   'Construction points
    nNeckOffset = 1
    If Val(txtAge) > 10 Then
        DefaultNeckDepth = 2
    Else
        DefaultNeckDepth = 1.5
    End If

    If txtThroatToSternal <> "" Then
        GivenNeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
        If GivenNeckDepth > DefaultNeckDepth Then
            nNeckOffset = nNeckOffset + Abs(GivenNeckDepth - DefaultNeckDepth)
        End If
    End If

    PR_MakeXY xyChinStart, Val(txtChinLeftBotX), Val(txtChinLeftBotY)
    PR_MakeXY xy1, xyChinStart.X, xyChinStart.Y + (nNeckOffset - 1)
    PR_MakeXY xy2, xy1.X, xyChinStart.Y + nNeckOffset
    PR_MakeXY xy4, -Abs(LeftArcAdjOut(txtRadiusNo)) + 0.375, Val(txtMouthHeight)
    PR_MakeXY xy3, xy4.X, xy4.Y - 1

   'Fit curve to three lines defined by the 4 points above
    n1 = FN_CalcLength(xy1, xy2)
    n2 = FN_CalcLength(xy2, xy3)
    n3 = FN_CalcLength(xy3, xy4)
    nValue = FN_CalcLength(xy1, xy3)
    nValue = ((n1 ^ 2 + n2 ^ 2) - nValue ^ 2) / (2 * n1 * n2)
    rAngle = Arccos(nValue) / 2
    nBotRadius = Tan(rAngle) * (2 * (n1 / 3))
    nTopRadius = Tan(rAngle) * (n3 / 2)

    ChinProfileX(1) = xy1.X
    ChinProfileY(1) = xy1.Y

    ChinProfileX(2) = xy1.X
    ChinProfileY(2) = xy1.Y + (n1 / 3)

    PR_CalcPolar xy2, FN_CalcAngle(xy2, xy3), (2 * (n1 / 3)), xyTmp
    ChinProfileX(5) = xyTmp.X
    ChinProfileY(5) = xyTmp.Y

    PR_MakeXY xyTmp, ChinProfileX(2), ChinProfileY(2)
    PR_CalcPolar xyTmp, 180, nBotRadius, xyBotCen
    PR_MakeXY xyTmp, ChinProfileX(5), ChinProfileY(5)
    aInc = FN_CalcAngle(xyBotCen, xyTmp) / 3
        
    For ii = 3 To 4
        PR_CalcPolar xyBotCen, aInc, nBotRadius, xyTmp
        aInc = aInc + aInc
        ChinProfileX(ii) = xyTmp.X
        ChinProfileY(ii) = xyTmp.Y
    Next ii

    PR_CalcPolar xy3, FN_CalcAngle(xy3, xy2), (n3 / 2), xyTmp
    ChinProfileX(6) = xyTmp.X
    ChinProfileY(6) = xyTmp.Y

    PR_CalcPolar xyTmp, FN_CalcAngle(xy3, xy2) + 90, nTopRadius, xyTopCen
    aAngle = FN_CalcAngle(xyTopCen, xyTmp)
    aInc = (aAngle - 180) / 3

    For ii = 7 To 9
        aAngle = aAngle - aInc
        PR_CalcPolar xyTopCen, aAngle, nTopRadius, xyTmp
        ChinProfileX(ii) = xyTmp.X
        ChinProfileY(ii) = xyTmp.Y
    Next ii

    ChinProfileX(10) = xy4.X
    ChinProfileY(10) = xy4.Y

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyChinStart.X); CC; "xyStart.y+"; Str$(xyChinStart.Y)
    For ii = 1 To 10
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinProfileX(ii)); CC; "xyStart.y+"; Str$(ChinProfileY(ii))
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "Chin"
   
   'Set center point for subsequent use in lip covering
    xyChinBotCen = xyBotCen
    xyChinTopCen = xyTopCen
    g_nChinBotRadius = nBotRadius
    g_nChinTopRadius = nTopRadius

End Sub


Private Sub PR_DrawChinCollar()

    ' this Procedure calculates and draws the Chin Extension Collar

    Dim xyTextPoint As xy
    Dim xyBotInnerLeft As xy, xyBotInnerRight As xy
    Dim xyTopInnerLeft As xy, xyTopInnerRight As xy
    Dim xyBotOuterLeft As xy, xyBotOuterRight As xy
    Dim xyTopOuterLeft As xy, xyTopOuterRight As xy
    Dim xyTopInnerContour As xy, xyBotInnerContour As xy
    Dim xyTopOuterContour As xy, xyBotOuterContour As xy
    Dim NeckDepth#, NeckCircum#, ContourMin#
    Dim InnerXDist#, OuterXDist#, YDist#
    Dim InnerXContourDist#, OuterXContourDist#, YContourStart#, YContourEnd#
    Dim IntVal%, DecVal#

    If optChinCollar.Value = True Or optContouredChinCollar.Value = True Then

       'Round Up to nearest 1/2"
        If txtCircOfNeck <> "" Then
            NeckCircum = FN_CmToInches(txtUnits, txtCircOfNeck)
            IntVal = Int(NeckCircum)
            DecVal = NeckCircum - IntVal
            If DecVal > 0 And DecVal <= 0.5 Then
                NeckCircum = IntVal + 0.5
            ElseIf DecVal > 0.5 Then
                NeckCircum = IntVal + 1
            End If
        End If

       'Round Up to nearest 1/2"
        If txtThroatToSternal <> "" Then
            NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
            IntVal = Int(NeckDepth)
            DecVal = NeckDepth - IntVal
            If DecVal > 0 And DecVal <= 0.5 Then
                NeckDepth = IntVal + 0.5
            ElseIf DecVal > 0.5 Then
                NeckDepth = IntVal + 1
            End If
        End If

       'Round Up to nearest 1/2"
        If txtChinCollarMin <> "" Then
            ContourMin = FN_CmToInches(txtUnits, txtChinCollarMin)
            IntVal = Int(ContourMin)
            DecVal = ContourMin - IntVal
            If DecVal > 0 And DecVal <= 0.5 Then
                ContourMin = IntVal + 0.5
            ElseIf DecVal > 0.5 Then
                ContourMin = IntVal + 1
            End If
'            If ContourMin < 2.5 Then ContourMin = 2.5
            If ContourMin < 1 Then ContourMin = 1
        Else
            ContourMin = NeckDepth - 0.5
            If ContourMin < 2.5 Then ContourMin = 2.5
        End If

        ' Set Horizontal Distances
        If NeckCircum <= 8 Then
            InnerXDist = 6.1875
            OuterXDist = 7
        Else
            InnerXDist = (((NeckCircum - 8) * 9) / 14) + 6.1875
            OuterXDist = (((NeckCircum - 8) * 3) / 4) + 7
        End If

        ' Set Contour Horizontal Distances
        If NeckCircum <= 8 Then
            InnerXContourDist# = 2.0625
            OuterXContourDist# = 2.3125
        Else
            InnerXContourDist# = ((NeckCircum - 8) / 4.618) + 2.0625
            OuterXContourDist# = ((NeckCircum - 8) / 4) + 2.3125
        End If

        ' Set Vertical Distances
        If NeckDepth <= 0 Then
            If optContouredChinCollar.Value = True Then
                YDist = 3
            Else
                YDist = 3.5
            End If
'        ElseIf NeckDepth > 0 And NeckDepth <= 2.5 Then
'            YDist = 3     'chg
        Else
            YDist = NeckDepth + 0.5
        End If

        If ContourMin > 0 Then
            ContourMin = ContourMin + 0.5
            YContourStart = (YDist - ContourMin) / 2
            YContourEnd = YDist - ((YDist - ContourMin) / 2)
        End If
    
        ' Set Chin Collar Points
        PR_MakeXY xyBotInnerLeft, 0, 0
        PR_MakeXY xyTopInnerLeft, 0, YDist
        PR_MakeXY xyTopOuterLeft, 0, 0
        PR_MakeXY xyBotOuterLeft, 0, -YDist

        If optContouredChinCollar.Value = True Then
            PR_MakeXY xyBotInnerContour, InnerXContourDist, 0
            PR_MakeXY xyBotInnerRight, InnerXDist, YContourStart
            PR_MakeXY xyTopInnerRight, InnerXDist, YContourEnd
            PR_MakeXY xyTopInnerContour, InnerXContourDist, YDist
            PR_MakeXY xyTopOuterContour, OuterXContourDist, 0
            PR_MakeXY xyTopOuterRight, OuterXDist, -YContourStart
            PR_MakeXY xyBotOuterRight, OuterXDist, -YContourEnd
            PR_MakeXY xyBotOuterContour, OuterXContourDist, -YDist
        Else
            PR_MakeXY xyBotInnerRight, InnerXDist, 0
            PR_MakeXY xyTopInnerRight, InnerXDist, YDist
            PR_MakeXY xyTopOuterRight, OuterXDist, 0
            PR_MakeXY xyBotOuterRight, OuterXDist, -YDist
        End If

        ' Draw Inner Collar
        Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotInnerLeft.X); CC; "xyStart.y+"; Str$(xyBotInnerLeft.Y)
        If optContouredChinCollar.Value = True Then
            Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotInnerContour.X); CC; "xyStart.y+"; Str$(xyBotInnerContour.Y)
        End If
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotInnerRight.X); CC; "xyStart.y+"; Str$(xyBotInnerRight.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopInnerRight.X); CC; "xyStart.y+"; Str$(xyTopInnerRight.Y)
        If optContouredChinCollar.Value = True Then
            Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopInnerContour.X); CC; "xyStart.y+"; Str$(xyTopInnerContour.Y)
        End If
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopInnerLeft.X); CC; "xyStart.y+"; Str$(xyTopInnerLeft.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotInnerLeft.X); CC; "xyStart.y+"; Str$(xyBotInnerLeft.Y)
        Print #txtfNum, ");"
        PR_AddEntityID "InnerChinCollar"

        ' Draw Outer Collar
        Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopOuterLeft.X); CC; "xyStart.y+"; Str$(xyTopOuterLeft.Y)
        If optContouredChinCollar.Value = True Then
            Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopOuterContour.X); CC; "xyStart.y+"; Str$(xyTopOuterContour.Y)
        End If
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopOuterRight.X); CC; "xyStart.y+"; Str$(xyTopOuterRight.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotOuterRight.X); CC; "xyStart.y+"; Str$(xyBotOuterRight.Y)
        If optContouredChinCollar.Value = True Then
            Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotOuterContour.X); CC; "xyStart.y+"; Str$(xyBotOuterContour.Y)
        End If
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyBotOuterLeft.X); CC; "xyStart.y+"; Str$(xyBotOuterLeft.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyTopOuterLeft.X); CC; "xyStart.y+"; Str$(xyTopOuterLeft.Y)
        Print #txtfNum, ");"
        PR_AddEntityID "OuterChinCollar"

        PR_MakeXY xyTextPoint, xyTopInnerLeft.X + 1, xyTopInnerLeft.Y - 0.5
        PR_TemplateDetails xyTextPoint, "INNER Chin Collar"
        PR_MakeXY xyTextPoint, xyBotInnerLeft.X + 0.35, xyBotInnerLeft.Y + 1.75
        PR_SetLayer "Notes"
        PR_DrawText "Fold", xyTextPoint, 0.25, 90

        PR_MakeXY xyTextPoint, xyTopOuterLeft.X + 1, xyTopOuterLeft.Y - 0.5
        PR_TemplateDetails xyTextPoint, "OUTER Chin Collar"
        PR_MakeXY xyTextPoint, xyBotOuterLeft.X + 0.35, xyBotOuterLeft.Y + 1.75
        PR_SetLayer "Notes"
        PR_DrawText "Fold", xyTextPoint, 0.25, 90
        PR_SetLayer "TemplateLeft"
    
    End If

End Sub

Private Sub PR_DrawChinStrap()

Dim xyNeckTopFront As xy, xyNeckTopBack As xy
Dim xyNeckBottomFront As xy, xyNeckBottomBack As xy
Dim xyNeckTopVelcro As xy, xyNeckBottomVelcro As xy
Dim xyChinStrapTopFront As xy, xyChinStrapTopBack  As xy
Dim xyFacelineTop As xy, xyFacelineBot As xy
Dim xyBacklineTop As xy, xyBacklineBot As xy
Dim xyMouthFront As xy, xyMouthBack As xy
Dim xyLipStrapTop As xy, xyLipStrapBottom As xy
Dim xyStrapText As xy, xyTextPoint As xy
Dim XDistance#, YDistance#
Dim NeckDepth#, XOffset#
Dim RightArcStart#, RightArcEnd#
Dim RightArcStep#, HeadBack#
Dim TempLen#, StrapLength#
Dim Xlen1#, XLen2#, VelcroLen#, VelcroText

If optChinStrap.Value = True Or optModifiedChinStrap.Value = True Then
    PR_ChinStrapDimensions

    ' Get Neck Points
    XDistance = Val(txtCSNeckCircum) / 2
    RightArcStart = 9 / 2.54
    RightArcEnd = 17.05 / 2.54
    RightArcStep = (RightArcEnd - RightArcStart) / 16
    RightArcStep = RightArcStep * (Val(txtRadiusNo) - 1)
    HeadBack = RightArcStart + RightArcStep
    XOffset = 0

    PR_MakeXY xyNeckTopBack, XDistance, 0
    If xyNeckTopBack.X > HeadBack Then
        XOffset = -(xyNeckTopBack.X - HeadBack)
        PR_MakeXY xyNeckTopBack, xyNeckTopBack.X + XOffset, 0
    End If
    
    PR_MakeXY xyNeckTopFront, -XDistance + XOffset, 0
    txtxyNeckTopFrontX = xyNeckTopFront.X
    txtxyNeckTopFrontY = xyNeckTopFront.Y

    ' Set Min  Neck depth of 2" for adults and 1.5" Under ten
    If txtThroatToSternal = "" Then
        If Val(txtAge) >= 10 Then
            NeckDepth = 2
        Else
            NeckDepth = 1.5
        End If
    ElseIf txtThroatToSternal <> "" Then
        NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
    End If

    If NeckDepth < 1 Then
        MsgBox "Neck Depth should be Greater than 1'', Consult with Supervisor.", 48, "Chin Strap Details"
    End If

    PR_MakeXY xyNeckBottomFront, xyNeckTopFront.X, -NeckDepth
    PR_MakeXY xyNeckBottomBack, xyNeckTopBack.X, -NeckDepth
    PR_MakeXY xyNeckTopVelcro, xyNeckTopBack.X + 1, xyNeckTopBack.Y
    PR_MakeXY xyNeckBottomVelcro, xyNeckBottomBack.X + 1 + XOffset, xyNeckBottomBack.Y

    PR_DrawLine xyNeckBottomVelcro, xyNeckTopVelcro
    PR_AddEntityID "VelcroBack"
    PR_SetLayer "Notes"
    PR_DrawLine xyNeckTopBack, xyNeckBottomBack
    PR_SetLayer "TemplateLeft"
    PR_AddEntityID "NeckBack"

    PR_SetLayer "Notes"
    PR_MakeXY xyStrapText, xyNeckTopBack.X, xyNeckTopBack.Y - 0.25
    
    VelcroLen = NeckDepth - 0.125
    VelcroText = FN_InchesToText(VelcroLen)
    VelcroText = VelcroText & "\"""
    PR_DrawText VelcroText, xyStrapText, 0.1, 0
    
    PR_MakeXY xyStrapText, xyNeckTopBack.X + 0.1, xyNeckTopBack.Y - 0.5
    PR_DrawText " Velcro", xyStrapText, 0.1, 0
    PR_SetLayer "TemplateLeft"

    ' Set Top Front Point
    YDistance = Val(txtCSChinAngle)
    PR_MakeXY xyChinStrapTopFront, 0, YDistance

    ' Set FaceLine and Back Line
    XDistance = Val(txtCSChinAngle) / 2
    PR_MakeXY xyFacelineTop, xyChinStrapTopFront.X - XDistance, xyChinStrapTopFront.Y
    PR_MakeXY xyBacklineTop, xyChinStrapTopFront.X + XDistance, xyChinStrapTopFront.Y

    PR_MakeXY xyFacelineBot, xyFacelineTop.X, 0
    PR_MakeXY xyBacklineBot, xyBacklineTop.X, 0

    ' Set Mouth and lipstrap
    YDistance = Val(txtCSChintoMouth)
    XDistance = Val(txtCSMouthWidth)
    PR_MakeXY xyMouthFront, xyFacelineBot.X, xyFacelineBot.Y + YDistance
    PR_MakeXY xyMouthBack, xyMouthFront.X + XDistance, xyMouthFront.Y

    PR_MakeXY xyLipStrapTop, xyMouthFront.X - 0.125, xyMouthFront.Y + 0.75
    PR_MakeXY xyLipStrapBottom, xyMouthFront.X - 0.125, xyMouthFront.Y
    
    ' Set and draw front curve
    PR_DrawCSFrontCurve xyLipStrapTop, xyLipStrapBottom, xyMouthFront, xyMouthBack, xyFacelineTop, xyChinStrapTopFront

    ' Set and draw back curve
    If optChinStrap.Value = True Then
        PR_DrawCSBackCurve xyBacklineBot, xyFacelineTop, xyChinStrapTopFront, xyNeckTopVelcro
    ElseIf optModifiedChinStrap.Value = True Then
        PR_DrawCSModBackCurve xyNeckTopVelcro, xyChinStrapTopFront
    End If

    ' Draw Top
    If optChinStrap.Value = True Then
        PR_MakeXY xyChinStrapTopBack, xyChinStrapTopFront.X + 2, xyChinStrapTopFront.Y
    ElseIf optModifiedChinStrap.Value = True Then
        PR_MakeXY xyChinStrapTopBack, xyChinStrapTopFront.X + 2.5, xyChinStrapTopFront.Y
    End If
    PR_DrawLine xyChinStrapTopFront, xyChinStrapTopBack
    PR_AddEntityID "ChinStrapTop"

    ' Draw Velcro Top
    PR_DrawLine xyNeckTopBack, xyNeckTopVelcro
    PR_AddEntityID "VelcroTop"
    
    ' Set and draw chin
    PR_DrawCSChin xyMouthFront, xyMouthBack, xyNeckBottomFront, xyNeckBottomVelcro

    ' Calculate Strap Length
    TempLen = fnRoundInches(FN_CalcLength(xyChinStrapTopBack, xyNeckTopBack))
    StrapLength = Int(TempLen)
    TempLen = TempLen - StrapLength
    StrapLength = StrapLength - 1   'Subtract an inch
   'Round down to nearest 1/2"
    If TempLen >= 0.5 Then
        StrapLength = StrapLength + 0.5
    End If

'PR_DrawMarker xyChinStrapTopBack
'PR_DrawText "xyChinStrapTopBack", xyChinStrapTopBack, .05, 0
'PR_DrawMarker xyNeckTopBack
'PR_DrawText "xyNeckTopBack", xyNeckTopBack, .05, 0

    PR_MakeXY xyStrapText, xyChinStrapTopFront.X + 0.1, xyChinStrapTopBack.Y - 0.5
    PR_SetLayer "Notes"
    PR_DrawText "Strap Length: " & FN_InchesToText(StrapLength) & "\""", xyStrapText, 0.1, 0
    PR_SetLayer "TemplateLeft"

    ' Set and draw Lip Covering
    PR_DrawCSLipCovering

    ' Set and Draw Ear Openings
    PR_DrawEars
    PR_DrawEarFlaps


    ' Add Patient Details
    PR_MakeXY xyTextPoint, -1.5, 0.75
    PR_TemplateDetails xyTextPoint, "Patient Details"
End If

End Sub

Private Sub PR_DrawCSBackCurve(xyBacklineBot As xy, xyFacelineTop As xy, xyChinStrapTopFront As xy, xyNeckTopVelcro As xy)

    Dim XDistance#, YDistance#
    Dim XStep#, YStep#, aAngle#, aInc#, nRadius#, nOffSet#
    Dim ii%
    Dim xyTopEar As xy, xyBotEar  As xy, xyArcCen As xy, xyTmp  As xy

    XDistance = Abs(xyBacklineBot.X) - 1
    YDistance = xyFacelineTop.Y - txtCSEarTopHeight

    XStep = XDistance / 20
    YStep = YDistance / 5
    
    PR_MakeXY xyTopEar, 0, (txtCSEarTopHeight)
    PR_MakeXY xyBotEar, 0, (txtCSEarBotHeight)
    nOffSet = (txtCSEarTopHeight - txtCSEarBotHeight) / 8

    ChinStrapBackX(1) = xyChinStrapTopFront.X + 2
    ChinStrapBackY(1) = xyChinStrapTopFront.Y
    ChinStrapBackX(2) = (xyChinStrapTopFront.X + 2) - (1 * XStep)
    ChinStrapBackY(2) = xyChinStrapTopFront.Y - YStep
    ChinStrapBackX(3) = (xyChinStrapTopFront.X + 2) - (2.5 * XStep)
    ChinStrapBackY(3) = xyChinStrapTopFront.Y - (2 * YStep)
    ChinStrapBackX(4) = (xyChinStrapTopFront.X + 2) - (4.5 * XStep)
    ChinStrapBackY(4) = xyChinStrapTopFront.Y - (3 * YStep)
    ChinStrapBackX(5) = (xyChinStrapTopFront.X + 2) - (7.5 * XStep)
    ChinStrapBackY(5) = xyChinStrapTopFront.Y - (4 * YStep)

    PR_MakeXY xyTmp, ChinStrapBackX(5), ChinStrapBackY(5)
    aAngle = FN_CalcAngle(xyTopEar, xyTmp)
    PR_CalcPolar xyTopEar, aAngle, nOffSet, xyTmp
                                   
    aAngle = (aAngle + 90) / 2
    nRadius = nOffSet * Tan(aAngle * (PI / 180))
    PR_MakeXY xyArcCen, xyTopEar.X + nRadius, xyTopEar.Y - nOffSet
    aAngle = FN_CalcAngle(xyArcCen, xyTmp)
    aInc = (180 - aAngle) / 3
    For ii = 6 To 9
        PR_CalcPolar xyArcCen, aAngle, nRadius, xyTmp
        aAngle = aAngle + aInc
        ChinStrapBackX(ii) = xyTmp.X
        ChinStrapBackY(ii) = xyTmp.Y
    Next ii
    
'Original code GG-19.Sep.96
'    ChinStrapBackX(6) = 0
'    ChinStrapBackY(6) = txtCSEarTopHeight
'    ChinStrapBackX(7) = 0
'    ChinStrapBackY(7) = txtCSEarBotHeight + (5 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 6))
'    ChinStrapBackX(8) = 0
'    ChinStrapBackY(8) = txtCSEarBotHeight + (2 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 3))
'    ChinStrapBackX(9) = 0
'    ChinStrapBackY(9) = txtCSEarBotHeight + ((txtCSEarTopHeight - txtCSEarBotHeight) / 2)
'    ChinStrapBackX(10) = 0
'    ChinStrapBackY(10) = txtCSEarBotHeight + ((txtCSEarTopHeight - txtCSEarBotHeight) / 3)
'    ChinStrapBackX(11) = 0
'    ChinStrapBackY(11) = txtCSEarBotHeight + (1 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 6))
'    ChinStrapBackX(12) = 0
'    ChinStrapBackY(12) = txtCSEarBotHeight
'
    
    ChinStrapBackX(13) = 1.5 * XStep
    ChinStrapBackY(13) = txtCSEarBotHeight / 2

    PR_MakeXY xyTmp, ChinStrapBackX(13), ChinStrapBackY(13)
    aAngle = FN_CalcAngle(xyBotEar, xyTmp)
    PR_CalcPolar xyBotEar, aAngle, nOffSet, xyTmp
                                   
    aAngle = ((360 - aAngle) + 90) / 2
    nRadius = nOffSet * Tan(aAngle * (PI / 180))
    PR_MakeXY xyArcCen, xyBotEar.X + nRadius, xyBotEar.Y + nOffSet
    aAngle = FN_CalcAngle(xyArcCen, xyTmp)
    aInc = (aAngle - 180) / 3
    aAngle = 180
    For ii = 10 To 12
        PR_CalcPolar xyArcCen, aAngle, nRadius, xyTmp
        aAngle = aAngle + aInc
        ChinStrapBackX(ii) = xyTmp.X
        ChinStrapBackY(ii) = xyTmp.Y
    Next ii

    ChinStrapBackX(14) = 4 * XStep
    ChinStrapBackY(14) = 0.125
    ChinStrapBackX(15) = 10 * XStep
    ChinStrapBackY(15) = 0
    ChinStrapBackX(16) = 15 * XStep
    ChinStrapBackY(16) = 0
    ChinStrapBackX(17) = xyNeckTopVelcro.X - 1
    ChinStrapBackY(17) = 0
    
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 1 To 17
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapBackX(ii)); CC; "xyStart.y+"; Str$(ChinStrapBackY(ii))
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "ChinStrapBack"

'For ii = 1 To 17
'    PR_MakeXY xyTmp, ChinStrapBackX(ii), ChinStrapBackY(ii)
'    PR_DrawMarker xyTmp
'    PR_DrawText Trim$(Str$(ii)), xyTmp, .05, 0
'Next ii
     
End Sub

Private Sub PR_DrawCSChin(xyMouthFront As xy, xyMouthBack As xy, xyNeckBottomFront As xy, xyNeckBottomVelcro As xy)

Dim xyChinTop  As xy, xyChinBottom As xy
Dim xyBotCurve As xy, xyTopCurve As xy
Dim xyDiagStart As xy, xyDiagEnd As xy, xyTmp  As xy
Dim ChinStep1#, ChinStep2#, ChinDist#
Dim xyChinArc  As xy
Dim PI#, ChinHt#
Dim BotRadius#, TopRadius#
Dim Theta#, Delta#
Dim XA#, YA#, X%
Dim DiagLen#, DiagStep#
Dim j%, i%, ii%, v%
ReDim OffsetChinX(12) As Double, OffsetChinY(12) As Double

TopRadius = 1.8 / 2.54
BotRadius = 0.9 / 2.54
PI = 3.141592654

If optChinStrap.Value = True Or optModifiedChinStrap.Value = True Then

   'Set and Draw Chin
    PR_MakeXY xyChinTop, xyMouthFront.X + 0.25, xyMouthFront.Y
    PR_MakeXY xyChinBottom, xyNeckBottomFront.X + 0.25, xyNeckBottomFront.Y
    txtxyChinTopX = xyChinTop.X
    txtxyChinTopY = xyChinTop.Y
    
xyChinTop.X = xyChinTop.X + 0.125
xyChinBottom.X = xyChinBottom.X + 0.125

    ChinHt = Abs(xyChinTop.Y - xyChinBottom.Y)
    Delta = 30

    If Val(txtThroatToSternal.Text) > 0 Then
        ChinDist = FN_CmToInches(txtUnits, txtThroatToSternal)
    Else
        If Val(txtAge) > 10 Then
            ChinDist = 2
        Else
            ChinDist = 1.5
        End If
    End If


    If ChinDist >= 2 Then
        ChinStep1 = 2
        ChinStep2 = 0.375
    Else
        ChinStep1 = ChinDist
'        ChinStep2 = (3 * ChinDist) / 10
        ChinStep2 = 0.375
    End If

    ChinStrapChinX(1) = (xyChinBottom.X)
    ChinStrapChinY(1) = (xyChinBottom.Y)

    ChinStrapChinX(2) = xyChinBottom.X
    ChinStrapChinY(2) = 0 - ChinStep1

    ChinStrapChinX(3) = xyChinBottom.X
    ChinStrapChinY(3) = 0 - ChinStep2

For i = 1 To 3
    OffsetChinX(i) = ChinStrapChinX(i) - 0.125
    OffsetChinY(i) = ChinStrapChinY(i)
Next i

    PR_MakeXY xyBotCurve, ChinStrapChinX(3) - BotRadius, ChinStrapChinY(3)

    For i = 1 To 3
        Theta = (i * (Delta)) * (PI / 180)
        XA = BotRadius * Cos(Theta)
        YA = BotRadius * Sin(Theta)
        ChinStrapChinX(i + 3) = xyBotCurve.X + XA
        ChinStrapChinY(i + 3) = xyBotCurve.Y + YA
        XA = (BotRadius - 0.125) * Cos(Theta)
        YA = (BotRadius - 0.125) * Sin(Theta)
        OffsetChinX(i + 3) = xyBotCurve.X + XA
        OffsetChinY(i + 3) = xyBotCurve.Y + YA
    Next i

    ChinStrapChinX(12) = xyChinTop.X
    ChinStrapChinY(12) = xyChinTop.Y
    
    ChinStrapChinX(10) = xyChinTop.X
    ChinStrapChinY(10) = ChinStrapChinY(7) + TopRadius

    ChinStrapChinY(10) = xyBotCurve.Y + BotRadius + TopRadius

    ChinStrapChinX(11) = xyChinTop.X
    ChinStrapChinY(11) = xyChinTop.Y - ((ChinStrapChinY(12) - ChinStrapChinY(10)) / 2)

For i = 10 To 12
    OffsetChinX(i) = ChinStrapChinX(i) - 0.125
    OffsetChinY(i) = ChinStrapChinY(i)
Next i
    
    PR_MakeXY xyTopCurve, ChinStrapChinX(10) + TopRadius, ChinStrapChinY(10)
    
    j = 9
    For i = 1 To 3
        Theta = (i * Delta) * (PI / 180)
        XA = TopRadius * Cos(Theta)
        YA = TopRadius * Sin(Theta)
        ChinStrapChinX(j) = xyTopCurve.X - XA
        ChinStrapChinY(j) = xyTopCurve.Y - YA
        XA = (TopRadius + 0.125) * Cos(Theta)
        YA = (TopRadius + 0.125) * Sin(Theta)
        OffsetChinX(j) = xyTopCurve.X - XA
        OffsetChinY(j) = xyTopCurve.Y - YA
        j = j - 1
    Next i
   
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 1 To 12
        Print #txtfNum, CC; "xyStart.x+"; Str$(OffsetChinX(ii)); CC; "xyStart.y+"; Str$(OffsetChinY(ii))
    Next ii
    Print #txtfNum, ");"
    
   'Revise back
    xyChinTop.X = xyChinTop.X - 0.125
    xyChinBottom.X = xyChinBottom.X - 0.125
    
    PR_AddEntityID "Chin"
    PR_DrawLine xyChinTop, xyMouthBack
    PR_AddEntityID "ChinStrapMouth"
    PR_DrawLine xyChinBottom, xyNeckBottomVelcro
    PR_AddEntityID "ChinStrapBottom"

    If chkNeckElastic.Value = 1 Then
        PR_MakeXY xyTmp, xyChinBottom.X + 0.4, xyChinBottom.Y + 0.5
        PR_SetLayer "Notes"
        PR_DrawText "1\"" Elastic Sewn Underneath", xyTmp, 0.1, 0
        PR_SetLayer "TemplateLeft"
    End If

    
    PR_SetLayer "Construct"
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 1 To 12
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapChinX(ii)); CC; "xyStart.y+"; Str$(ChinStrapChinY(ii))
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "LipStrapChinConstructLine"
   
   'Set Global Points for chin curve
    xyChinTopCen = xyTopCurve
    xyChinBotCen = xyBotCurve

End If

End Sub

Private Sub PR_DrawCSFrontCurve(xyLipStrapTop As xy, xyLipStrapBottom As xy, xyMouthFront As xy, xyMouthBack As xy, xyFacelineTop As xy, xyChinStrapTopFront As xy)
    
Dim FrontStartX#, FrontStartY#
Dim XDistance#, YDistance#
Dim XStep#, YStep#
Dim ii%

    ' Set Front Curve
    
    If chkLipStrap.Value = 1 Then
        FrontStartX = xyLipStrapTop.X
        FrontStartY = xyLipStrapTop.Y
        PR_DrawLine xyLipStrapTop, xyLipStrapBottom
        PR_AddEntityID "ChinStrapFront"
        PR_DrawLine xyMouthBack, xyLipStrapBottom
        PR_AddEntityID "Mouth"
        XDistance = Abs(xyLipStrapTop.X)
        YDistance = xyFacelineTop.Y - xyLipStrapTop.Y

    Else
        XDistance = Abs(xyMouthBack.X)
        YDistance = xyFacelineTop.Y - xyMouthFront.Y
        FrontStartX = xyMouthBack.X
        FrontStartY = xyMouthBack.Y
    End If


    XStep = XDistance / 7
    YStep = YDistance / 20
    
    
    ChinStrapFrontX(1) = FrontStartX
    ChinStrapFrontY(1) = FrontStartY
    ChinStrapFrontX(2) = FrontStartX + (XStep)
    ChinStrapFrontY(2) = FrontStartY + (YStep * YStep)
    ChinStrapFrontX(3) = FrontStartX + (XStep * 2)
    ChinStrapFrontY(3) = FrontStartY + ((2 * YStep) * (2 * YStep))
    ChinStrapFrontX(4) = FrontStartX + (XStep * 3)
    ChinStrapFrontY(4) = FrontStartY + ((3 * YStep) * (3 * YStep))
    ChinStrapFrontX(5) = FrontStartX + (XStep * 4)
    ChinStrapFrontY(5) = FrontStartY + ((4 * YStep) * (4 * YStep))
    ChinStrapFrontX(6) = FrontStartX + (XStep * 5)
    ChinStrapFrontY(6) = FrontStartY + ((5 * YStep) * (5 * YStep))
    ChinStrapFrontX(7) = FrontStartX + (XStep * 6)
    ChinStrapFrontY(7) = FrontStartY + ((6 * YStep) * (6 * YStep))
    ChinStrapFrontX(8) = xyChinStrapTopFront.X
    ChinStrapFrontY(8) = xyChinStrapTopFront.Y
    
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 1 To 8
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapFrontX(ii)); CC; "xyStart.y+"; Str$(ChinStrapFrontY(ii))
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "ChinStrapFront"


End Sub

Private Sub PR_DrawCSLipCovering()
    
   'Uses module level variable xyChinTopCen & xyChinBotCen
   '
Dim xyLipCoverTopLeft As xy, xyLipCoverTopRight As xy
Dim xyLipCoverBottomRight As xy, xyLipCoverBottomMid As xy
Dim xyTmp As xy, xy7 As xy, xy6 As xy, xy3 As xy, xy2 As xy, xy1 As xy

Dim nn%, nOffSet#, BotRadius#, aAngle#, aStart#
Dim XPoint#, YPoint#, ChinVal%
Dim i%, ii%, j%, DrawArcPart%
Dim DistStepX#, DistStepY#
Dim ChinCheckY#, XOffset#, xyLipCoveringBotCen As xy
Dim LipDistX#, LipDistY#


'If chkLipCovering.Value = 1 And chkLipStrap.Value = 1 Then
If chkLipCovering.Value = 1 Then

    XOffset = -Abs(txtxyChinTopX)
    BotRadius = 0.9 / 2.54
    
    ' Set Top Left & Right Points
    XPoint = txtxyChinTopX
    YPoint = txtxyChinTopY
'FUDGE Start - GG 18.Sep.95
'    PR_MakeXY xyLipCoverTopLeft, XPoint + .125, YPoint + 1.25
'    PR_MakeXY xyLipCoverTopRight, -.875, YPoint + 1.25
'    PR_MakeXY xyLipCoverBottomRight, -.875, .25
    PR_MakeXY xyLipCoverTopLeft, XPoint + 0.25, YPoint + 1.25
    PR_MakeXY xyLipCoverTopRight, -0.75, YPoint + 1.25
    PR_MakeXY xyLipCoverBottomRight, -0.75, 0.25
'FUDGE End - GG 18.Sep.95
    XPoint = txtxyNeckTopFrontX
    YPoint = txtxyNeckTopFrontY

'PR_DrawMarker xyLipCoverTopLeft
'PR_DrawText "xyLipCoverTopLeft", xyLipCoverTopLeft, .05, 0
'PR_DrawMarker xyLipCoverTopRight
'PR_DrawText "xyLipCoverTopRight", xyLipCoverTopRight, .05, 0
'PR_DrawMarker xyLipCoverBottomRight
'PR_DrawText "xyLipCoverBottomRight", xyLipCoverBottomRight, .05, 0
    
    'Find Bottom Mid Point
    'This is the intesection with the neck line (y=0)
    'at xy6
     PR_MakeXY xy6, ChinStrapChinX(6), ChinStrapChinY(6)
     PR_MakeXY xy7, ChinStrapChinX(7), ChinStrapChinY(7)

       'Take intersection point at end of horizontal line
        DrawArcPart = True
        aStart = FN_CalcAngle(xyChinBotCen, xy6)
        aAngle = aStart - ((0.375 / (2 * PI * BotRadius)) * 360)
        PR_CalcPolar xyChinBotCen, aAngle, BotRadius, xyLipCoverBottomMid

     xyLipCoverBottomMid.X = xyLipCoverBottomMid.X + 0.125
    
    ' Get Points For Curve, From Mid to Bottom Right
    ' Adjust to include Lip
    ChinStrapLipX(1) = xyLipCoverTopLeft.X
    ChinStrapLipY(1) = xyLipCoverTopLeft.Y
    
    ChinStrapLipX(2) = ChinStrapChinX(12) + 0.125
    ChinStrapLipY(2) = ChinStrapChinY(12)
    ChinStrapLipX(3) = ChinStrapChinX(11) + 0.125
    ChinStrapLipY(3) = ChinStrapChinY(11)
    nn = 4
    For ii = 10 To 7 Step -1
        Select Case nn
            Case 4
                nOffSet = 0.08
            Case 5
                nOffSet = 0.125
            Case 6
                nOffSet = 0.125
            Case 7
                nOffSet = 0.06
        End Select
        PR_MakeXY xyTmp, ChinStrapChinX(ii) + 0.125, ChinStrapChinY(ii)
        PR_CalcPolar xyChinTopCen, FN_CalcAngle(xyChinTopCen, xyTmp), FN_CalcLength(xyChinTopCen, xyTmp) - nOffSet, xyTmp
        ChinStrapLipX(nn) = xyTmp.X
        ChinStrapLipY(nn) = xyTmp.Y
        nn = nn + 1
    Next ii

    
    If DrawArcPart Then
        ChinStrapLipX(8) = ChinStrapChinX(6) + 0.125
        ChinStrapLipY(8) = ChinStrapChinY(6)
        
        PR_CalcPolar xyChinBotCen, aAngle + ((aStart - aAngle) / 3), BotRadius, xyTmp
        ChinStrapLipX(10) = xyTmp.X + 0.125
        ChinStrapLipY(10) = xyTmp.Y
        
        PR_CalcPolar xyChinBotCen, aAngle + (2 * ((aStart - aAngle) / 3)), BotRadius, xyTmp
        ChinStrapLipX(9) = xyTmp.X + 0.125
        ChinStrapLipY(9) = xyTmp.Y

        ChinStrapLipX(11) = xyLipCoverBottomMid.X
        ChinStrapLipY(11) = xyLipCoverBottomMid.Y
    Else
        ChinStrapLipX(8) = xyLipCoverBottomMid.X
        ChinStrapLipY(8) = xyLipCoverBottomMid.Y
    End If

   'Add Offset to curve
    For i = 1 To 11
        ChinStrapLipX(i) = ChinStrapLipX(i) + XOffset
    Next i

   'Indicate position on Pattern
   'Somewhere above we have an extra .125" add in the X axis
   'I don't know why etc. etc. etc.
    Dim Point1 As xy, Point2 As xy
    PR_SetLayer "Notes"

    PR_MakeXY Point1, xyLipCoverTopRight.X - 0.125, xyLipCoverTopRight.Y
    PR_MakeXY Point2, Point1.X - 0.125, Point1.Y
    PR_DrawLine Point1, Point2
    PR_MakeXY Point2, Point1.X, Point1.Y - 0.125
    PR_DrawLine Point1, Point2

    PR_MakeXY Point1, xyLipCoverBottomRight.X - 0.125, xyLipCoverBottomRight.Y
    PR_MakeXY Point2, Point1.X - 0.125, Point1.Y
    PR_DrawLine Point1, Point2
    PR_MakeXY Point2, Point1.X, Point1.Y + 0.125
    PR_DrawLine Point1, Point2

    PR_MakeXY Point1, xyLipCoverBottomMid.X - 0.125, xyLipCoverBottomMid.Y
    PR_DrawLine Point1, xyLipCoverBottomMid

    PR_SetLayer "TemplateLeft"

    PR_MakeXY xyLipCoverTopLeft, xyLipCoverTopLeft.X + XOffset, xyLipCoverTopLeft.Y
    PR_MakeXY xyLipCoverTopRight, xyLipCoverTopRight.X + XOffset, xyLipCoverTopRight.Y
    PR_MakeXY xyLipCoverBottomRight, xyLipCoverBottomRight.X + XOffset, xyLipCoverBottomRight.Y
    PR_MakeXY xyLipCoverBottomMid, xyLipCoverBottomMid.X + XOffset, xyLipCoverBottomMid.Y
    
    nOffSet = FN_CalcLength(xyLipCoverBottomMid, xyLipCoverBottomRight) / 2
    aAngle = FN_CalcAngle(xyLipCoverBottomMid, xyLipCoverBottomRight)
    PR_CalcPolar xyLipCoverBottomMid, aAngle, nOffSet, xyTmp
    PR_CalcPolar xyTmp, aAngle + 270, 9, xyLipCoveringBotCen
    PR_DrawArc xyLipCoveringBotCen, xyLipCoverBottomMid, xyLipCoverBottomRight
    PR_AddEntityID "LipCoveringBot"

'PR_DrawMarker xyLipCoverBottomMid
'PR_DrawText "xyLipCoverBottomMid", xyLipCoverBottomMid, .05, 0


    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyLipCoverTopLeft.X); CC; "xyStart.y+"; Str$(xyLipCoverTopLeft.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyLipCoverTopRight.X); CC; "xyStart.y+"; Str$(xyLipCoverTopRight.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyLipCoverBottomRight.X); CC; "xyStart.y+"; Str$(xyLipCoverBottomRight.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "LipCovering"

   'Draw chin area
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    If DrawArcPart Then
        For i = 1 To 11
            Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapLipX(i)); CC; "xyStart.y+"; Str$(ChinStrapLipY(i))
        Next i
    Else
        For i = 1 To 8
            Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapLipX(i)); CC; "xyStart.y+"; Str$(ChinStrapLipY(i))
        Next i
    End If
    Print #txtfNum, ");"
    PR_AddEntityID "ChinStrapLipCoveringCurve"

    Dim length#, VelcroText, TextPoint As xy

    PR_MakeXY TextPoint, (xyLipCoverTopLeft.X + 0.25), (xyLipCoverTopLeft.Y - 0.1)
    PR_TemplateDetails TextPoint, "Chin Strap Lip Covering"
    
    ' Add Velcro
    PR_SetLayer "Notes"
    PR_MakeXY xy1, xyLipCoverTopRight.X - 0.75, xyLipCoverTopRight.Y

    If xy1.X > xyLipCoverBottomMid.X Then
        PR_MakeXY xy2, xy1.X, xyLipCoveringBotCen.Y
        ii = FN_CirLinInt(xy1, xy2, xyLipCoveringBotCen, FN_CalcLength(xyLipCoverBottomRight, xyLipCoveringBotCen), xy2)
    Else
        PR_MakeXY xy2, xy1.X, xyChinBotCen.Y
        PR_MakeXY xyTmp, (xyChinBotCen.X + XOffset) + 0.125, xyChinBotCen.Y
        ii = FN_CirLinInt(xy1, xy2, xyTmp, BotRadius, xy2)
'PR_DrawMarker xyTmp
'PR_DrawText "xyTmp", xyTmp, .05, 0
    End If

    PR_DrawLine xy1, xy2
    length = FN_CalcLength(xy1, xy2) + 0.25
    VelcroText = FN_InchesToText(fnRoundInches(length))
    VelcroText = VelcroText & "\""  Velcro"
    PR_MakeXY TextPoint, xy1.X - 0.25, xy1.Y - 0.25
    PR_DrawText VelcroText, TextPoint, 0.1, -90
    length = FN_CalcLength(xyLipCoverTopRight, xyLipCoverBottomRight) + 0.25
    VelcroText = FN_InchesToText(fnRoundInches(length))
    VelcroText = VelcroText & "\""  Velcro"
    PR_MakeXY TextPoint, xyLipCoverTopRight.X - 0.25, xy1.Y - 0.25
    PR_DrawText VelcroText, TextPoint, 0.1, -90

    
    'Draw an offset chin curve on
    PR_SetLayer "Construct"
    XOffset = XOffset + 0.125
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
     For i = 3 To 12
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapChinX(i) + XOffset); CC; "xyStart.y+"; Str$(ChinStrapChinY(i))
     Next i
    Print #txtfNum, ");"
    PR_AddEntityID "LipStrapChinConstructLine"

    PR_SetLayer "TemplateLeft"

End If

End Sub

Private Sub PR_DrawCSModBackCurve(xyNeckTopVelcro As xy, xyChinStrapTopFront As xy)

    Dim BackStartX#, BackStartY#
    Dim XDistance#, YDistance#
    Dim XStep#, YStep#
    Dim ii%

    ' Set Front Curve
    BackStartX = xyNeckTopVelcro.X - 1
    BackStartY = xyNeckTopVelcro.Y
    XDistance = xyNeckTopVelcro.X - 2.5
    YDistance = xyChinStrapTopFront.Y

    XStep = XDistance / 15
    YStep = YDistance / 20
    
    ChinStrapBackX(1) = BackStartX
    ChinStrapBackY(1) = BackStartY
    ChinStrapBackX(2) = BackStartX - (XStep)
    ChinStrapBackY(2) = BackStartY + (YStep * YStep)
    ChinStrapBackX(3) = BackStartX - (XStep * 2)
    ChinStrapBackY(3) = BackStartY + ((2 * YStep) * (2 * YStep))
    ChinStrapBackX(4) = BackStartX - (XStep * 3)
    ChinStrapBackY(4) = BackStartY + ((3 * YStep) * (3 * YStep))
    ChinStrapBackX(5) = BackStartX - (XStep * 4)
    ChinStrapBackY(5) = BackStartY + ((4 * YStep) * (4 * YStep))
    ChinStrapBackX(6) = BackStartX - (XStep * 5)
    ChinStrapBackY(6) = BackStartY + ((5 * YStep) * (5 * YStep))
    ChinStrapBackX(7) = BackStartX - (XStep * 6)
    ChinStrapBackY(7) = BackStartY + ((6 * YStep) * (6 * YStep))
    ChinStrapBackX(8) = xyChinStrapTopFront.X + 2.5
    ChinStrapBackY(8) = xyChinStrapTopFront.Y
    
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 1 To 8
        Print #txtfNum, CC; "xyStart.x+"; Str$(ChinStrapBackX(ii)); CC; "xyStart.y+"; Str$(ChinStrapBackY(ii))
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "ChinStrapBack"

End Sub

Private Sub PR_DrawEarFlaps()

' This procedure Draws the ear Flap

Dim xyFlapTopLeft As xy, xyFlapTopRight As xy
Dim xyFlapBottomLeft As xy, xyFlapBottomRight As xy
Dim DartAng#, FlapX#, FlapY#
Dim Point1 As xy, Point2 As xy, PI#, Notes As xy
Dim X1#, X2#, Y1#, y2#
Dim LeftXOffSet#, RightXOffSet#, YOffSet#
Dim Left1 As xy, Left2 As xy, Left3 As xy, Left4 As xy
Dim Right1 As xy, Right2 As xy, Right3 As xy, Right4 As xy

PI = 3.141592654
If optModifiedChinStrap.Value = True Or optFaceMask.Value = True Or optOpenFaceMask.Value = True Then

' Calculate Bottom points
If optModifiedChinStrap.Value = False Then
    X1 = Val(txtDartStartX)
    Y1 = Val(txtDartStartY)
    X2 = Val(txtDartEndX)
    y2 = Val(txtDartEndY)
    PR_MakeXY Point1, X1, Y1
    PR_MakeXY Point2, X2, y2
    DartAng = FN_CalcAngle(Point1, Point2)
    DartAng = DartAng * (PI / 180)
    FlapX = Abs(txtDartStartX) + 2
    FlapY = FlapX * Tan(DartAng)
    PR_MakeXY xyFlapBottomRight, Point1.X + FlapX, Point1.Y + FlapY + 0.25
    PR_MakeXY xyFlapBottomLeft, -0.625, xyFlapBottomRight.Y
Else
    PR_MakeXY xyFlapBottomRight, 2.1875, 0.25
    PR_MakeXY xyFlapBottomLeft, -0.375, 0.25
End If
' Set OffSets
LeftXOffSet = -3
RightXOffSet = 1.625
If chkOpenHeadMask.Value = 0 Then
    YOffSet = TopArcOpp(txtRadiusNo) + Abs(xyFlapBottomLeft.Y) + 1.25
Else
    YOffSet = Abs(xyFlapBottomLeft.Y) + 1.75
End If
If optModifiedChinStrap.Value = True Then
    YOffSet = Val(txtCSChinAngle) + Abs(xyFlapBottomLeft.Y) + 0.25
End If

' Draw Left Ear
If chkLeftEarFlap.Value = 1 Then
    ' Set Top Points
    If optModifiedChinStrap.Value = False Then
        PR_MakeXY xyFlapTopLeft, -0.625, txtTopLeftEar + 0.125
        PR_MakeXY xyFlapTopRight, 2, txtTopLeftEar + 0.125
    Else
        PR_MakeXY xyFlapTopLeft, -0.375, txtTopLeftEar + 0.125
        PR_MakeXY xyFlapTopRight, 2.1875, txtTopLeftEar + 0.125
    End If

   'Indicate position on Pattern
    PR_SetLayer "Notes"
    PR_MakeXY Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y
    PR_MakeXY Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125
    PR_DrawLine Point1, xyFlapTopLeft
    PR_DrawLine Point2, xyFlapTopLeft

    PR_MakeXY Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125
    PR_DrawLine Point1, xyFlapTopRight
    PR_DrawLine Point2, xyFlapTopRight

    PR_MakeXY Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomRight
    PR_DrawLine Point2, xyFlapBottomRight

    PR_MakeXY Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y
    PR_MakeXY Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomLeft
    PR_DrawLine Point2, xyFlapBottomLeft

    PR_MakeXY Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet
    PR_MakeXY Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet
    PR_MakeXY Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet
    PR_MakeXY Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet

    PR_SetLayer "TemplateLeft"

    PR_MakeXY Left1, xyFlapTopLeft.X + LeftXOffSet, xyFlapTopLeft.Y + YOffSet
    PR_MakeXY Left2, xyFlapTopRight.X + LeftXOffSet, xyFlapTopRight.Y + YOffSet
    PR_MakeXY Left3, xyFlapBottomRight.X + LeftXOffSet, xyFlapBottomRight.Y + YOffSet
    PR_MakeXY Left4, xyFlapBottomLeft.X + LeftXOffSet, xyFlapBottomLeft.Y + YOffSet

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left1.X); CC; "xyStart.y+"; Str$(Left1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left2.X); CC; "xyStart.y+"; Str$(Left2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left3.X); CC; "xyStart.y+"; Str$(Left3.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left4.X); CC; "xyStart.y+"; Str$(Left4.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left1.X); CC; "xyStart.y+"; Str$(Left1.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "LeftEarFlap"
    
    ' Add notes
     Left1.X = Left1.X + 0.1
     Left1.Y = Left1.Y - 0.1
    PR_TemplateDetails Left1, "Left Ear Flap"
End If

' Draw Right Ear
If chkRightEarFlap.Value = 1 Then
    ' Set Top Points
    If optModifiedChinStrap.Value = False Then
        PR_MakeXY xyFlapTopLeft, -0.625, txtTopRightEar + 0.125
        PR_MakeXY xyFlapTopRight, 2, txtTopRightEar + 0.125
    Else
        PR_MakeXY xyFlapTopLeft, -0.375, txtTopRightEar + 0.125
        PR_MakeXY xyFlapTopRight, 2.1875, txtTopRightEar + 0.125
    End If
   'Indicate position on Pattern
    PR_SetLayer "Notes"
    PR_MakeXY Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y
    PR_MakeXY Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125
    PR_DrawLine Point1, xyFlapTopLeft
    PR_DrawLine Point2, xyFlapTopLeft

    PR_MakeXY Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125
    PR_DrawLine Point1, xyFlapTopRight
    PR_DrawLine Point2, xyFlapTopRight

    PR_MakeXY Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomRight
    PR_DrawLine Point2, xyFlapBottomRight

    PR_MakeXY Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y
    PR_MakeXY Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomLeft
    PR_DrawLine Point2, xyFlapBottomLeft

    PR_MakeXY Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet
    PR_MakeXY Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet
    PR_MakeXY Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet
    PR_MakeXY Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet

    PR_SetLayer "TemplateRight"

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right1.X); CC; "xyStart.y+"; Str$(Right1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right2.X); CC; "xyStart.y+"; Str$(Right2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right3.X); CC; "xyStart.y+"; Str$(Right3.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right4.X); CC; "xyStart.y+"; Str$(Right4.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right1.X); CC; "xyStart.y+"; Str$(Right1.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "RightEarFlap"
    
    ' Add notes
     Right1.X = Right1.X + 0.1
     Right1.Y = Right1.Y - 0.1
    PR_TemplateDetails Right1, "Right Ear Flap"
End If
End If
End Sub

Private Sub PR_DrawEarHole(EarX#, EarTopY#, EarBotY#, Template$)

    Dim xyTopOfEarRight As xy, xyBotOfEarRight As xy
    Dim xyTopOfEarLeft As xy, xyBotOfEarLeft As xy
    Dim xyCen As xy, HalfEarWidth#
    Dim nStartAng#, nDeltaAng#
    Dim nRad#, nEndAng#

    nStartAng = 0
    nDeltaAng = -180

    If optModifiedChinStrap.Value = False Then
        HalfEarWidth = 0.125
        nRad = 0.125
    Else
        HalfEarWidth = 0.09375
        nRad = 0.09375
    End If

    PR_SetLayer Template

    ' Set points and Draw Right of Ear
    PR_MakeXY xyTopOfEarRight, EarX + HalfEarWidth, EarTopY
    PR_MakeXY xyBotOfEarRight, EarX + HalfEarWidth, EarBotY
    PR_DrawLine xyTopOfEarRight, xyBotOfEarRight
    PR_AddEntityID "EarRight"
    
    ' Set points and Draw Left of Ear
    PR_MakeXY xyTopOfEarLeft, EarX - HalfEarWidth, EarTopY
    PR_MakeXY xyBotOfEarLeft, EarX - HalfEarWidth, EarBotY
    PR_DrawLine xyTopOfEarLeft, xyBotOfEarLeft
    PR_AddEntityID "EarLeft"

    ' Draw Top Curve of Ear
    PR_MakeXY xyCen, EarX, EarTopY
    PR_DrawArc xyCen, xyTopOfEarLeft, xyTopOfEarRight
    PR_AddEntityID "EarTop"
    
    'Draw Bottom Curve of Ear
    PR_MakeXY xyCen, EarX, EarBotY
    Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
    PR_AddEntityID "EarBottom"
    PR_SetLayer "TemplateLeft"

End Sub

Private Sub PR_DrawEars()

Dim xyTextPoint  As xy, YEarStep#
Dim EarTopY#, EarBotY#, EarX#

If optModifiedChinStrap.Value = True Or optFaceMask.Value = True Or optOpenFaceMask.Value = True Then
If optModifiedChinStrap.Value = False Then
    YEarStep = 0.125
ElseIf optModifiedChinStrap.Value = True Then
    YEarStep = 0.1875
End If

If chkEarSize.Value = 0 Then
    If optModifiedChinStrap.Value = False Then
        EarTopY = Val(txtLowerEarHeight) - YEarStep
        EarBotY = Val(txtMouthHeight) + YEarStep
        EarX = -0.125
        txtTopLeftEar = EarTopY + YEarStep
        txtTopRightEar = EarTopY + YEarStep
    ElseIf optModifiedChinStrap.Value = True Then
        EarTopY = Val(txtCSEarTopHeight) - YEarStep
        EarBotY = Val(txtCSEarBotHeight) + YEarStep
        EarX = 0.09375
        txtTopLeftEar = EarTopY + YEarStep
        txtTopRightEar = EarTopY + YEarStep
    End If
        
    If chkLeftEarClosed.Value = 0 Then
        PR_DrawEarHole EarX, EarTopY, EarBotY, "TemplateLeft"
    ElseIf chkLeftEarClosed.Value = 1 Then
        PR_SetLayer "Notes"
        PR_MakeXY xyTextPoint, 0.25, EarBotY
        PR_DrawText "Left Ear Closed", xyTextPoint, 0.1, 0
        PR_SetLayer "TemplateLeft"
    End If
        
    If chkRightEarClosed.Value = 0 Then
        PR_DrawEarHole EarX, EarTopY, EarBotY, "TemplateRight"
        PR_SetLayer "TemplateLeft"
    ElseIf chkRightEarClosed.Value = 1 Then
        PR_SetLayer "Notes"
        PR_MakeXY xyTextPoint, 0.25, EarBotY + 0.25
        PR_DrawText "Right Ear Closed", xyTextPoint, 0.1, 0
        PR_SetLayer "TemplateLeft"
    End If

ElseIf chkEarSize.Value = 1 Then
    ' Left Ear
    If Val(txtLeftEarLength) = 0 Then
        If optModifiedChinStrap.Value = False Then
            EarTopY = Val(txtLowerEarHeight) - YEarStep
            EarBotY = Val(txtMouthHeight) + YEarStep
            EarX = -0.125
            txtTopLeftEar = EarTopY + YEarStep
        ElseIf optModifiedChinStrap.Value = True Then
            EarTopY = Val(txtCSEarTopHeight) - YEarStep
            EarBotY = Val(txtCSEarBotHeight) + YEarStep
            EarX = 0.09375
            txtTopLeftEar = EarTopY + YEarStep
        End If
    
    ElseIf Val(txtLeftEarLength) > 0 Then
        If optModifiedChinStrap.Value = False Then
            EarBotY = Val(txtMouthHeight) + YEarStep
            EarX = -0.125
            EarTopY = EarBotY + FN_CmToInches(txtUnits, txtLeftEarLength) - 0.25
            txtTopLeftEar = EarTopY + YEarStep
        ElseIf optModifiedChinStrap.Value = True Then
            EarBotY = Val(txtCSEarBotHeight) + YEarStep
            EarX = 0.09375
            EarTopY = EarBotY + FN_CmToInches(txtUnits, txtLeftEarLength) - 0.25
            txtTopLeftEar = EarTopY + YEarStep
        End If
    End If

    If chkLeftEarClosed.Value = 0 Then
        PR_DrawEarHole EarX, EarTopY, EarBotY, "TemplateLeft"
    ElseIf chkLeftEarClosed.Value = 1 Then
        PR_SetLayer "Notes"
        PR_MakeXY xyTextPoint, 0.25, EarBotY
        PR_DrawText "Left Ear Closed", xyTextPoint, 0.1, 0
        PR_SetLayer "TemplateLeft"
    End If

    ' Right Ear
    If Val(txtRightEarLength) = 0 Then
        If optModifiedChinStrap.Value = False Then
            EarTopY = Val(txtLowerEarHeight) - YEarStep
            EarBotY = Val(txtMouthHeight) + YEarStep
            EarX = -0.125
            txtTopRightEar = EarTopY + 0.125
        ElseIf optModifiedChinStrap.Value = True Then
            EarTopY = Val(txtCSEarTopHeight) - YEarStep
            EarBotY = Val(txtCSEarBotHeight) + YEarStep
            EarX = 0.09375
            txtTopRightEar = EarTopY + YEarStep
        End If
    ElseIf Val(txtRightEarLength) > 0 Then
        If optModifiedChinStrap.Value = False Then
            EarBotY = Val(txtMouthHeight) + YEarStep
            EarX = -0.125
            EarTopY = EarBotY + FN_CmToInches(txtUnits, txtRightEarLength) - 0.25
            txtTopRightEar = EarTopY + YEarStep
        ElseIf optModifiedChinStrap.Value = True Then
            EarBotY = Val(txtCSEarBotHeight) + YEarStep
            EarX = 0.09375
            EarTopY = EarBotY + FN_CmToInches(txtUnits, txtRightEarLength) - 0.25
            txtTopRightEar = EarTopY + YEarStep
        End If
    End If
    If chkRightEarClosed.Value = 0 Then
        PR_DrawEarHole EarX, EarTopY, EarBotY, "TemplateRight"
        PR_SetLayer "TemplateLeft"
    ElseIf chkRightEarClosed.Value = 1 Then
        PR_SetLayer "Notes"
        PR_MakeXY xyTextPoint, 0.25, EarBotY
        PR_DrawText "Right Ear Closed", xyTextPoint, 0.1, 0
        PR_SetLayer "TemplateLeft"
    End If
End If
End If
End Sub

Private Sub PR_DrawEyeFlaps()
' This procedure Draws the eye Flap
' Globals Variables
'
'       xyNoseDiagTop As xy, xyNoseDiagBot As xy
'
Dim xyFlapTopLeft As xy, xyFlapTopRight As xy
Dim xyFlapBottomLeft As xy, xyFlapBottomRight As xy, xySideLineInt  As xy, xyBotLineInt As xy
Dim DartAng#, FlapX#, FlapY#
Dim PI#, Notes  As xy
Dim X1#, X2#, Y1#, y2#, Dist#, Point1    As xy, Point2 As xy
Dim LeftXOffSet#, RightXOffSet#, YOffSet#
Dim Left1 As xy, Left2 As xy, Left3 As xy, Left4 As xy, Left5  As xy, Left6 As xy
Dim Right1 As xy, Right2 As xy, Right3 As xy, Right4 As xy, Right5  As xy, Right6 As xy
Dim LeftYOffSet#, RightYOffSet#, XOffset#

CQ = CC & QQ
PI = 3.141592654
X1 = Val(txtBotOfEyeX)
Y1 = Val(txtBotOfEyeY)
Dist = Val(txtEyeWidth)

' Set Top Points
PR_MakeXY xyFlapTopLeft, X1 + 0.25, Y1 + 1.125
PR_MakeXY xyFlapTopRight, X1 + Dist + 0.25, Y1 + 1.125

' Calculate Bottom points
PR_MakeXY xyFlapBottomRight, xyFlapTopRight.X, Y1 - 0.875
PR_MakeXY xyFlapBottomLeft, xyFlapTopLeft.X, Y1 - 0.875

'Check if Flap intersect nose line
'Step 1 is to extend nose line to ensure intersection

Dim iSide%, iBot%, aAngle#, xyPt1 As xy, xyPt2 As xy, Intersection%

aAngle = FN_CalcAngle(xyNoseDiagBot, xyNoseDiagTop)
PR_CalcPolar xyNoseDiagTop, aAngle, 20, xyPt1
aAngle = FN_CalcAngle(xyNoseDiagTop, xyNoseDiagBot)
PR_CalcPolar xyNoseDiagBot, aAngle, 20, xyPt2

iSide = FN_LinLinInt(xyPt1, xyPt2, xyFlapBottomLeft, xyFlapTopLeft, xySideLineInt)
iBot = FN_LinLinInt(xyPt1, xyPt2, xyFlapBottomLeft, xyFlapBottomRight, xyBotLineInt)

'Cut only if no nose covering
If iBot = True And iSide = True And chkNoseCovering.Value = 0 Then
    Intersection = True
Else
    Intersection = False
End If


' Set OffSets
LeftYOffSet = -0.25
RightYOffSet = 2.5
XOffset = TopArcOpp(txtRadiusNo)
If chkLeftEyeFlap.Value = 1 Or chkRightEyeFlap.Value = 1 Then
   'Indicate position on Pattern
    PR_SetLayer "Notes"
    PR_MakeXY Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y
    PR_MakeXY Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125
    PR_DrawLine Point1, xyFlapTopLeft
    PR_DrawLine Point2, xyFlapTopLeft

    PR_MakeXY Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125
    PR_DrawLine Point1, xyFlapTopRight
    PR_DrawLine Point2, xyFlapTopRight

    PR_MakeXY Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomRight
    PR_DrawLine Point2, xyFlapBottomRight

    If Intersection Then
        PR_MakeXY Point1, xyBotLineInt.X + 0.125, xyBotLineInt.Y
        PR_DrawLine Point1, xyBotLineInt
    Else
        PR_MakeXY Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y
        PR_MakeXY Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125
        PR_DrawLine Point1, xyFlapBottomLeft
        PR_DrawLine Point2, xyFlapBottomLeft
    End If
    PR_MakeXY Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet
    PR_MakeXY Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet
    PR_MakeXY Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet
    PR_MakeXY Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet

End If

' Draw Left Eye
If chkLeftEyeFlap.Value = 1 Then
    PR_SetLayer "TemplateLeft"
    PR_MakeXY Left1, xyFlapTopLeft.X - XOffset, xyFlapTopLeft.Y + LeftYOffSet
    PR_MakeXY Left2, xyFlapTopRight.X - XOffset, xyFlapTopRight.Y + LeftYOffSet
    PR_MakeXY Left3, xyFlapBottomRight.X - XOffset, xyFlapBottomRight.Y + LeftYOffSet
    PR_MakeXY Left4, xyFlapBottomLeft.X - XOffset, xyFlapBottomLeft.Y + LeftYOffSet
    If Intersection Then
        PR_MakeXY Left5, xySideLineInt.X - XOffset, xySideLineInt.Y + LeftYOffSet
        PR_MakeXY Left6, xyBotLineInt.X - XOffset, xyBotLineInt.Y + LeftYOffSet
    End If

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left1.X); CC; "xyStart.y+"; Str$(Left1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left2.X); CC; "xyStart.y+"; Str$(Left2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left3.X); CC; "xyStart.y+"; Str$(Left3.Y)
    If Intersection Then
        Print #txtfNum, CC; "xyStart.x+"; Str$(Left6.X); CC; "xyStart.y+"; Str$(Left6.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(Left5.X); CC; "xyStart.y+"; Str$(Left5.Y)
    Else
        Print #txtfNum, CC; "xyStart.x+"; Str$(Left4.X); CC; "xyStart.y+"; Str$(Left4.Y)
    End If
    Print #txtfNum, CC; "xyStart.x+"; Str$(Left1.X); CC; "xyStart.y+"; Str$(Left1.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "LeftEyeFlap"

    ' Add notes
     Left1.X = Left1.X + 0.1
     Left1.Y = Left1.Y - 0.1
    PR_TemplateDetails Left1, "Left Eye Flap"

End If

' Draw Right Eye
If chkRightEyeFlap.Value = 1 Then
    PR_SetLayer "TemplateRight"
    PR_MakeXY Right1, xyFlapTopLeft.X - XOffset, xyFlapTopLeft.Y + RightYOffSet
    PR_MakeXY Right2, xyFlapTopRight.X - XOffset, xyFlapTopRight.Y + RightYOffSet
    PR_MakeXY Right3, xyFlapBottomRight.X - XOffset, xyFlapBottomRight.Y + RightYOffSet
    PR_MakeXY Right4, xyFlapBottomLeft.X - XOffset, xyFlapBottomLeft.Y + RightYOffSet
    If Intersection Then
        PR_MakeXY Right5, xySideLineInt.X - XOffset, xySideLineInt.Y + RightYOffSet
        PR_MakeXY Right6, xyBotLineInt.X - XOffset, xyBotLineInt.Y + RightYOffSet
    End If

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right1.X); CC; "xyStart.y+"; Str$(Right1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right2.X); CC; "xyStart.y+"; Str$(Right2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right3.X); CC; "xyStart.y+"; Str$(Right3.Y)
    If Intersection Then
        Print #txtfNum, CC; "xyStart.x+"; Str$(Right6.X); CC; "xyStart.y+"; Str$(Right6.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(Right5.X); CC; "xyStart.y+"; Str$(Right5.Y)
    Else
        Print #txtfNum, CC; "xyStart.x+"; Str$(Right4.X); CC; "xyStart.y+"; Str$(Right4.Y)
    End If
    Print #txtfNum, CC; "xyStart.x+"; Str$(Right1.X); CC; "xyStart.y+"; Str$(Right1.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "RightEyeFlap"

    ' Add notes
     Right1.X = Right1.X + 0.1
     Right1.Y = Right1.Y - 0.1
    
    PR_TemplateDetails Right1, "Right Eye Flap"
End If

End Sub

Private Sub PR_DrawEyeOpening()

'Start To form Eye Opening, Jobst Procedure No.5

Dim xyTopOfEye As xy, xyBotOfEye  As xy
Dim xyTopOfEyeStart As xy, xyBotOfEyeStart  As xy
Dim xyTopOfEyeEnd As xy, xyBotOfEyeEnd  As xy
Dim xyCen As xy, xyCen2 As xy, i%
Dim Tempx#, TempY#
Dim MidToEyeBot#, EyeWidthFromFace#
Dim nStartAng, nDeltaAng, nRad#, nEndAng#

Tempx = BotFaceAdj(txtRadiusNo)
PR_MakeXY xyTopOfEye, Tempx, -(Abs(txtMidToEyeTop))
PR_MakeXY xyBotOfEye, Tempx, -(Abs(txtMidToEyeTop)) - 0.25

' Check Chart 2 for Width of eye from faceline
For i = 1 To 16
    If FaceMaskChart2(1, i) = Val(txtCircumferenceTotal) Then
        EyeWidthFromFace = FaceMaskChart2(3, i)
        txtEyeWidth = EyeWidthFromFace
        Exit For
    End If
Next i


' Set points for top of eye
PR_MakeXY xyTopOfEyeStart, xyTopOfEye.X + 0.5, xyTopOfEye.Y
PR_MakeXY xyTopOfEyeEnd, xyTopOfEye.X + EyeWidthFromFace - 0.125, xyTopOfEye.Y

' Set points for bottom of eye
PR_MakeXY xyBotOfEyeStart, xyBotOfEye.X + 0.5, xyBotOfEye.Y
PR_MakeXY xyBotOfEyeEnd, xyBotOfEye.X + EyeWidthFromFace - 0.125, xyBotOfEye.Y

' Set points for Left Curve of eye
PR_MakeXY xyCen, xyTopOfEyeStart.X, xyTopOfEyeStart.Y - 0.125

' Set points for Right Curve of eye
PR_MakeXY xyCen2, xyTopOfEyeEnd.X, xyTopOfEyeEnd.Y - 0.125
nStartAng = 0
nDeltaAng = 90
nRad = 0.125

' Set txtLowerEarHeight with the y-axis value
txtLowerEarHeight = xyBotOfEyeEnd.Y

' Set Coords of Bottom of Eye
txtBotOfEyeX = xyBotOfEye.X
txtBotOfEyeY = xyBotOfEye.Y


If optFaceMask.Value = True Or (optOpenFaceMask.Value = True And chkEyes.Value = 1) Then
    ' Close Face Line in front of eyes
    If chkNoseCovering.Value = 0 Then
        PR_DrawLine xyTopOfEye, xyBotOfEye
        PR_AddEntityID "NoseBridge"
    End If

    PR_DrawLine xyTopOfEyeStart, xyTopOfEyeEnd
    PR_AddEntityID "EyeTop"

    PR_DrawLine xyBotOfEyeStart, xyBotOfEyeEnd
    PR_AddEntityID "EyeBottom"
    
    PR_DrawArc xyCen, xyTopOfEyeStart, xyBotOfEyeStart
    PR_AddEntityID "EyeLeft"

    Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen2.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
    nDeltaAng = -90
    PR_AddEntityID "EyeRightTop"
    Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen2.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
    PR_AddEntityID "EyeRightBottom"

End If

' End Drawing Eye Opening

End Sub

Private Sub PR_DrawForeHead()

Dim xyForeHeadBot As xy, xyForeHeadTop As xy
Dim Tempx#, TempY#

If chkOpenHeadMask.Value = 0 Then
    Tempx = -LeftArcAdjOut(txtRadiusNo)
    TempY = LeftArcOppOut(txtRadiusNo)
    PR_MakeXY xyForeHeadBot, Tempx, 0
    PR_MakeXY xyForeHeadTop, Tempx, TempY

    PR_DrawLine xyForeHeadBot, xyForeHeadTop
    PR_AddEntityID "ForeHead"
End If
End Sub

Private Sub PR_DrawHeadBand()
Dim xyCentre As xy, xyStartLeftArc As xy, xyEndLeftArc As xy
Dim xyStartRightArc As xy, xyEndRightArc As xy, HeadBand#
Dim RightArcStart#, RightArcEnd#, RightArcStep#
Dim Across#, Upwards#, xyTextPoint As xy, xyNotes As xy, xyTmp As xy, xyTmp1  As xy

Dim nRadius As Double
Dim iError As Integer

Dim AcrossOut#, UpwardsOut#

    
If optHeadBand.Value = True Then

    AcrossOut = LeftArcAdjOut(txtRadiusNo)
    UpwardsOut = LeftArcOppOut(txtRadiusNo)
    PR_MakeXY xyStartLeftArc, -AcrossOut, UpwardsOut

    Upwards = FN_CmToInches(txtUnits, txtHeadBandDepth) + 0.5
    RightArcStart = 9 / 2.54
    RightArcEnd = 17.05 / 2.54
    RightArcStep = (RightArcEnd - RightArcStart) / 16
    RightArcStep = RightArcStep * (Val(txtRadiusNo) - 1)

    PR_MakeXY xyCentre, 0, 0

    nRadius = FN_CalcLength(xyStartLeftArc, xyCentre)
    PR_MakeXY xyTmp, 0, Upwards
    PR_MakeXY xyTmp1, xyStartLeftArc.X - 10, Upwards
    iError = FN_CirLinInt(xyTmp, xyTmp1, xyCentre, nRadius, xyEndLeftArc)

    PR_MakeXY xyStartRightArc, RightArcStart + RightArcStep, 0
    Across = Sqr((Abs(xyStartRightArc.X) * Abs(xyStartRightArc.X)) - (Upwards * Upwards))
    PR_MakeXY xyEndRightArc, Across, Upwards
    
    PR_DrawArc xyCentre, xyStartLeftArc, xyEndLeftArc
    PR_AddEntityID "HeadBandLeftArc"
    PR_DrawArc xyCentre, xyStartRightArc, xyEndRightArc
    PR_AddEntityID "HeadBandRightArc"
    PR_DrawLine xyEndLeftArc, xyEndRightArc
    PR_AddEntityID "HeadBandTop"
    
    PR_MakeXY xyTmp, xyStartLeftArc.X, xyStartRightArc.Y
    PR_DrawLine xyTmp, xyStartRightArc
    PR_AddEntityID "HeadBandBottom"
    PR_DrawLine xyTmp, xyStartLeftArc
    PR_AddEntityID "HeadBandLeftLine"
    
   'add closure
    PR_SetLayer "Notes"
    PR_MakeXY xyTextPoint, xyEndRightArc.X - 0.1, Upwards / 2
    PR_DrawText "VELCRO", xyTextPoint, 0.1, FN_CalcAngle(xyEndRightArc, xyStartRightArc)

   ' Add notes
    PR_MakeXY xyTextPoint, xyEndLeftArc.X + 0.125, xyEndLeftArc.Y - 0.13
    PR_TemplateDetails xyTextPoint, "Head Band"

    
End If
End Sub

Private Sub PR_DrawLeftArc()
' This Procedure Draws the Left Arc of the Face Mask
' It uses the value held in txtRadiusNo to determine which
' Arc to Draw.  The Value Held in txtRadius must lie between
' 1 and 19 inclusive.

Dim xyStartArc As xy, xyEndArc As xy, xyCentreArc As xy
Dim AcrossIn#, UpwardsIn#
Dim AcrossOut#, UpwardsOut#

If chkOpenHeadMask.Value = 0 Then
    PR_MakeXY xyCentreArc, 0, 0

    AcrossIn = LeftArcAdjIn(txtRadiusNo)
    UpwardsIn = LeftArcOppIn(txtRadiusNo)

    AcrossOut = LeftArcAdjOut(txtRadiusNo)
    UpwardsOut = LeftArcOppOut(txtRadiusNo)

    PR_MakeXY xyStartArc, -AcrossIn, UpwardsIn
    PR_MakeXY xyEndArc, -AcrossOut, UpwardsOut
    PR_DrawArc xyCentreArc, xyStartArc, xyEndArc
    PR_AddEntityID "LeftArc"
End If
End Sub

Private Sub PR_DrawLeftCutOut()
Dim AcrossLeft#, UpwardsLeft#
Dim AcrossRight#, UpwardsRight#
Dim CentreX#, CentreY#
Dim Centre As xy, LeftTip As xy, RightTip As xy

If chkOpenHeadMask.Value = 0 Then
    CentreX = -3.8 / 2.54
    CentreY = 3.8 / 2.54

    PR_MakeXY Centre, CentreX, CentreY

    AcrossLeft = TopArcAdj(txtRadiusNo)
    UpwardsLeft = TopArcOpp(txtRadiusNo)
    AcrossRight = TopRightArcAdjIn(txtRadiusNo)
    UpwardsRight = TopRightArcOppIn(txtRadiusNo)

    PR_MakeXY LeftTip, -AcrossLeft, UpwardsLeft
    PR_MakeXY RightTip, -AcrossRight, UpwardsRight

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(LeftTip.X); CC; "xyStart.y+"; Str$(LeftTip.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Centre.X); CC; "xyStart.y+"; Str$(Centre.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(RightTip.X); CC; "xyStart.y+"; Str$(RightTip.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "LeftCutOut"

End If
End Sub

Private Sub PR_DrawLine(xyStart As xy, xyFinish As xy)
   'To the DRAFIX macro file (given by the global txtfNum).
   'Write the syntax to draw a LINE between two points.
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
   '    HANDLE  hEnt
   
    Print #txtfNum, "hEnt = AddEntity("
    Print #txtfNum, QQ; "line"; QC
    Print #txtfNum, "xyStart.x+"; Str$(xyStart.X); CC; "xyStart.y+"; Str$(xyStart.Y); CC
    Print #txtfNum, "xyStart.x+"; Str$(xyFinish.X); CC; "xyStart.y+"; Str$(xyFinish.Y)
    Print #txtfNum, ");"
 
End Sub

Private Sub PR_DrawLipCovering()

' This procedure Draws the Lip Covering

    Dim xyFlapTopLeft As xy, xyFlapTopRight As xy, xyFlapBottomMid As xy
    Dim xyFlapBottomLeft As xy, xyFlapBottomRight As xy, xyTextPoint As xy
    Dim Notes As xy, xy1  As xy, xy2 As xy, xyTmp As xy
    Dim xyDartStart As xy, xyDartEnd As xy, xy6 As xy, xy5 As xy
    Dim aStart#, aAngle#, nOffSet#
    Dim DistStepX#, Upwards#, PI#, X1#, X2#, Y1#, y2#, DistX#
    Dim DistStepY#, DistY#, ChinVal#, XOffset#, ChinCheckY#
    Dim nAlong#
    Dim X3#, Y3#, DartAngle#, ii%, i%, nn%
    Dim xyLipCoveringBotCen As xy
    Dim LipDistX#, LipDistY#, xyDartInt   As xy
    Dim length#, VelcroText, TextPoint As xy, nArc#

    PI = 3.141592654
    nAlong = 0.375
'    BotRadius = .9 / 2.54
'    TopRadius = 1.8 / 2.54


    If chkLipCovering.Value = 1 And optChinStrap.Value = False And optModifiedChinStrap.Value = False Then
        
        XOffset = -LeftArcAdjOut(txtRadiusNo) - 0.5

        ' Set Top Left Point
        Y1 = txtNoseBottomY
        PR_MakeXY xyFlapTopLeft, -LeftArcAdjOut(txtRadiusNo) + 0.375, Y1 + 0.5
   

        ' Set Top Right Point
        PR_MakeXY xyFlapTopRight, -0.875, xyFlapTopLeft.Y

        ' Calculate Bottom Right point
        X1 = Val(txtDartStartX)
        Y1 = Val(txtDartStartY)
        X2 = xyFlapTopRight.X
        X3 = Val(txtDartEndX)
        Y3 = Val(txtDartEndY)
        PR_MakeXY xyDartStart, X1, Y1
        PR_MakeXY xyDartEnd, X3, Y3
        
        If X1 >= X2 Then
            PR_MakeXY xyFlapBottomRight, xyFlapTopRight.X, Y1 + 0.25
        Else
            DistX = Abs(Abs(X2) - Abs(X1))
            DartAngle = FN_CalcAngle(xyDartStart, xyDartEnd)
            DartAngle = DartAngle * (PI / 180)
            Upwards = Abs(DistX * Tan(DartAngle))
            PR_MakeXY xyFlapBottomRight, xyFlapTopRight.X, Y1 + Upwards + 0.25
        End If


    ' Find Bottom Mid Point
    PR_MakeXY xy5, ChinProfileX(5), ChinProfileY(5)
    PR_MakeXY xy6, ChinProfileX(6), ChinProfileY(6)
'PR_DrawMarker xy5
'PR_DrawText "xy5", xy5, .05, 0
'PR_DrawMarker xy6
'PR_DrawText "xy6", xy6, .05, 0
'PR_DrawMarker xyDartStart
'PR_DrawText "xyDartStart", xyDartStart, .05, 0

    If xyDartStart.Y > xy6.Y Then
       'Calculate intersection point on toparc
        PR_MakeXY xy1, xyChinTopCen.X, xyDartStart.Y
        PR_MakeXY xy2, xyChinTopCen.X - 10, xyDartStart.Y
'PR_DrawLine xy1, xy2
'Print #txtfNum, "hEnt = AddEntity("; QQ; "circle"; QQ; ",xyStart.x+"; Str$(xyChinTopCen.x); CC; "xyStart.y+"; Str$(xyChinTopCen.y); ","; g_nChinTopRadius; ");"
        i = FN_CirLinInt(xy1, xy2, xyChinTopCen, g_nChinTopRadius, xyDartInt)
'PR_DrawMarker xyDartInt
'PR_DrawText "xyDartInt", xyDartInt, .05, 0
        aAngle = FN_CalcAngle(xyChinTopCen, xy6) - FN_CalcAngle(xyChinTopCen, xyDartInt)
        nArc = ((PI / 180) * aAngle) * g_nChinTopRadius
        nOffSet = FN_CalcLength(xy6, xy5)
        If nArc > nAlong Then
           'This is an impossible case so we don't deal with it
        ElseIf (nOffSet + nArc) > nAlong Then
            PR_CalcPolar xy6, FN_CalcAngle(xy6, xy5), nAlong - nArc, xyFlapBottomMid
'PR_DrawMarker xyFlapBottomMid
'PR_DrawText "xyFlapBottomMid on line", xyFlapBottomMid, .05, 0
        Else
            nOffSet = nAlong - (nOffSet + nArc)
            aStart = FN_CalcAngle(xyChinBotCen, xy5)
            aAngle = aStart - ((nOffSet / (2 * PI * g_nChinBotRadius)) * 360)
            PR_CalcPolar xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid
'PR_DrawMarker xyFlapBottomMid
'PR_DrawText "xyFlapBottomMid on bottom arc", xyFlapBottomMid, .05, 0
        End If
     
     ElseIf xyDartStart.Y > xy5.Y Then
       'Calculate intersection point on line
        PR_MakeXY xyDartInt, xy6.X - (Abs(xy6.Y - xyDartStart.Y) / Tan((FN_CalcAngle(xy5, xy6) * (PI / 180)))), xyDartStart.Y
        nOffSet = FN_CalcLength(xyDartInt, xy5)
        If nOffSet < nAlong Then
            aStart = FN_CalcAngle(xyChinBotCen, xy5)
            aAngle = aStart - (((nAlong - nOffSet) / (2 * PI * g_nChinBotRadius)) * 360)
            PR_CalcPolar xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid
        Else
            PR_CalcPolar xyDartInt, FN_CalcAngle(xy6, xy5), nAlong, xyFlapBottomMid
        End If
'PR_DrawMarker xyDartInt
'PR_DrawText "xyDartInt", xyDartInt, .05, 0
     Else
       'Calculate intersection point on the bottom arc
        PR_MakeXY xy1, xyChinBotCen.X, xyDartEnd.Y
        PR_MakeXY xy2, xyChinBotCen.X + 10, xyDartEnd.Y
'PR_DrawLine xy1, xy2
'Print #txtfNum, "hEnt = AddEntity("; QQ; "circle"; QQ; ",xyStart.x+"; Str$(xyChinBotCen.x); CC; "xyStart.y+"; Str$(xyChinBotCen.y); ","; g_nChinBotRadius; ");"
        i = FN_CirLinInt(xy1, xy2, xyChinBotCen, g_nChinBotRadius, xyDartInt)
'PR_DrawMarker xyDartInt
'PR_DrawText "xyDartInt", xyDartInt, .05, 0
        aStart = FN_CalcAngle(xyChinBotCen, xyDartInt)
        aAngle = aStart - ((nAlong / (2 * PI * g_nChinBotRadius)) * 360)
        PR_CalcPolar xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid

     End If

        ' Adjust to include Lip
        LipProfileX(1) = xyFlapTopLeft.X
        LipProfileY(1) = xyFlapTopLeft.Y
        LipProfileX(2) = ChinProfileX(10)
        LipProfileY(2) = ChinProfileY(10)
        LipProfileX(3) = ChinProfileX(9)
        LipProfileY(3) = ChinProfileY(9)

        nn = 4
        For ii = 8 To 6 Step -1
            Select Case nn
            Case 4
                nOffSet = 0.08
            Case 5
                nOffSet = 0.125
            Case 6
                nOffSet = 0.08
            End Select
            PR_MakeXY xyTmp, ChinProfileX(ii), ChinProfileY(ii)
            PR_CalcPolar xyChinTopCen, FN_CalcAngle(xyChinTopCen, xyTmp), FN_CalcLength(xyChinTopCen, xyTmp) - nOffSet, xyTmp
            LipProfileX(nn) = xyTmp.X
            LipProfileY(nn) = xyTmp.Y
            nn = nn + 1
        Next ii
        
        LipProfileX(7) = xyFlapBottomMid.X
        LipProfileY(7) = xyFlapBottomMid.Y


        ' Add Offset to curve
        For i = 1 To 7
            LipProfileX(i) = LipProfileX(i) + XOffset
        Next i

   'Indicate position on Pattern
    Dim Point1 As xy, Point2 As xy
    PR_SetLayer "Notes"

    PR_MakeXY Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y
    PR_MakeXY Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125
    PR_DrawLine Point1, xyFlapTopLeft
    PR_DrawLine Point2, xyFlapTopLeft

    PR_MakeXY Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125
    PR_DrawLine Point1, xyFlapTopRight
    PR_DrawLine Point2, xyFlapTopRight

    PR_MakeXY Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y
    PR_MakeXY Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomRight
    PR_DrawLine Point2, xyFlapBottomRight

    PR_MakeXY Point1, xyFlapBottomMid.X + 0.125, xyFlapBottomMid.Y
    PR_MakeXY Point2, xyFlapBottomMid.X, xyFlapBottomMid.Y + 0.125
    PR_DrawLine Point1, xyFlapBottomMid

    PR_SetLayer "TemplateLeft"
    
        ' Add OffSet to Points
        PR_MakeXY xyFlapTopLeft, xyFlapTopLeft.X + XOffset, xyFlapTopLeft.Y
        PR_MakeXY xyFlapTopRight, xyFlapTopRight.X + XOffset, xyFlapTopRight.Y
        PR_MakeXY xyFlapBottomRight, xyFlapBottomRight.X + XOffset, xyFlapBottomRight.Y
        PR_MakeXY xyFlapBottomMid, xyFlapBottomMid.X + XOffset, xyFlapBottomMid.Y

        nOffSet = FN_CalcLength(xyFlapBottomMid, xyFlapBottomRight) / 2
        aAngle = FN_CalcAngle(xyFlapBottomMid, xyFlapBottomRight)
        PR_CalcPolar xyFlapBottomMid, aAngle, nOffSet, xyTmp
        PR_CalcPolar xyTmp, aAngle + 270, 9, xyLipCoveringBotCen
        PR_DrawArc xyLipCoveringBotCen, xyFlapBottomMid, xyFlapBottomRight
        PR_AddEntityID "LipCoveringBot"
            
    
        Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyFlapTopLeft.X); CC; "xyStart.y+"; Str$(xyFlapTopLeft.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyFlapTopRight.X); CC; "xyStart.y+"; Str$(xyFlapTopRight.Y)
        Print #txtfNum, CC; "xyStart.x+"; Str$(xyFlapBottomRight.X); CC; "xyStart.y+"; Str$(xyFlapBottomRight.Y)
        Print #txtfNum, ");"
        PR_AddEntityID "LipCovering"

        Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        For ii = 1 To 7
            Print #txtfNum, CC; "xyStart.x+"; Str$(LipProfileX(ii)); CC; "xyStart.y+"; Str$(LipProfileY(ii))
        Next ii
        Print #txtfNum, ");"
        PR_AddEntityID "LipCoveringCurve"

        ' Add Velcro
        PR_SetLayer "Notes"
        PR_MakeXY xy1, xyFlapTopRight.X - 0.75, xyFlapTopRight.Y
        PR_MakeXY xy2, xy1.X, xyLipCoveringBotCen.Y
        ii = FN_CirLinInt(xy1, xy2, xyLipCoveringBotCen, FN_CalcLength(xyLipCoveringBotCen, xyFlapBottomRight), xy2)
        PR_DrawLine xy1, xy2
        length = FN_CalcLength(xy1, xy2) + 0.25
        VelcroText = FN_InchesToText(fnRoundInches(length))
        VelcroText = VelcroText & "\""  Velcro"
        PR_MakeXY TextPoint, xy1.X - 0.25, xy1.Y - 0.25
        PR_DrawText VelcroText, TextPoint, 0.1, -90
        length = FN_CalcLength(xyFlapTopRight, xyFlapBottomRight) + 0.25
        VelcroText = FN_InchesToText(fnRoundInches(length))
        VelcroText = VelcroText & "\""  Velcro"
        PR_MakeXY TextPoint, xyFlapTopRight.X - 0.25, xy1.Y - 0.25
        PR_DrawText VelcroText, TextPoint, 0.1, -90

        ' Add Notes
        PR_MakeXY xyTextPoint, xyFlapTopLeft.X + 0.25, xyFlapTopLeft.Y - 0.1
        PR_TemplateDetails xyTextPoint, "Lip Covering"

        PR_SetLayer "Construct"
        Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
        For ii = 3 To 10
            Print #txtfNum, CC; "xyStart.x+"; Str$(ChinProfileX(ii) + XOffset); CC; "xyStart.y+"; Str$(ChinProfileY(ii))
        Next ii
        Print #txtfNum, ");"

        PR_SetLayer "TemplateLeft"
    End If

End Sub

Private Sub PR_DrawMarker(xyPoint As xy)
   'Draw a Marker at the given point
    Print #txtfNum, "hEnt = AddEntity("; QQ; "marker"; QCQ; "xmarker"; QC; "xyStart.x+"; xyPoint.X; CC; "xyStart.y+"; xyPoint.Y; CC; "0.125);"

End Sub

Private Sub PR_DrawNeckAndDart()

' Procedure to Add Neck Piece, Jobst Procedure No.7

Dim NeckCircum#, FiguredNeck#
Dim NeckRightSideTop As xy, NeckLeftSideTop As xy
Dim NeckRightSideBot As xy, NeckLeftSideBot As xy
Dim DartStart As xy, DartEnd As xy
Dim Tempx#, TempY#
Dim TempX2#, TempY2#
Dim TempDepth#, NeckDepth#
Dim Dart#, DartEndX#, DartEndY#
Dim Circum#, LineLen#
Dim ZipBit#, i%

Circum = Val(txtRadiusNo)
ZipBit = 0
' Set For Zipper
If chkZipper.Value = 1 And Val(txtRadiusNo) > 3 Then
    Circum = Circum - 3
    ZipBit = 0.625
ElseIf chkZipper.Value = 1 And Circum <= 3 Then
    Circum = Circum + 17
    ZipBit = 0.625
End If

' Set For 2" wide Velcro
If chkVelcro.Value = 1 Then
    Circum = Circum + 2
    ZipBit = -0.375
End If


LineLen = 1.35 / 2.54
NeckCircum = FN_CmToInches(txtUnits, txtCircOfNeck)

' Check Chart 3 for Figured Neck
For i = 1 To 46
    If (NeckCircum >= FaceMaskChart3(1, i)) And (NeckCircum <= FaceMaskChart3(2, i)) And (NeckCircum < FaceMaskChart3(1, i + 1)) Then
        FiguredNeck = FaceMaskChart3(3, i)
    ElseIf NeckCircum >= 20 Then
        FiguredNeck = FaceMaskChart3(47, 3)
    End If
Next i

' Top Right point
Tempx = NeckRightAdj(txtRadiusNo)
TempY = NeckRightOpp(txtRadiusNo)
PR_MakeXY NeckRightSideTop, Tempx - ZipBit, TempY

' Top Left Point
TempX2 = Tempx - FiguredNeck
PR_MakeXY NeckLeftSideTop, TempX2, TempY

' Sets Min depth of 2" for adults and 1.5" Under ten
If txtThroatToSternal = "" Then
    If Val(txtAge) >= 10 Then
        NeckDepth = 2
    Else
        NeckDepth = 1.5
    End If
ElseIf txtThroatToSternal <> "" Then
    NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
End If

TempY2 = -(Abs(TempY) + NeckDepth)
PR_MakeXY NeckRightSideBot, Tempx - ZipBit, TempY2

' Bottom Left Point
PR_MakeXY NeckLeftSideBot, TempX2, TempY2

' Set txtChinLeftBotX and txtChinLeftBotY
txtChinLeftBotX = NeckLeftSideBot.X
txtChinLeftBotY = NeckLeftSideBot.Y

' Draw Dart
Dart = Tempx - FiguredNeck + 0.625
PR_MakeXY DartStart, Dart, TempY
txtDartStartX = DartStart.X
txtDartStartY = DartStart.Y
DartEndX = (RightArcAdjDown(Circum)) - LineLen
DartEndY = -RightArcOppDown(Circum)
PR_MakeXY DartEnd, DartEndX, DartEndY
txtDartEndX = DartEnd.X
txtDartEndY = DartEnd.Y

' Draw Neck Piece
Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
Print #txtfNum, CC; "xyStart.x+"; Str$(DartEnd.X); CC; "xyStart.y+"; Str$(DartEnd.Y)
Print #txtfNum, CC; "xyStart.x+"; Str$(DartStart.X); CC; "xyStart.y+"; Str$(DartStart.Y)
Print #txtfNum, CC; "xyStart.x+"; Str$(NeckRightSideTop.X); CC; "xyStart.y+"; Str$(NeckRightSideTop.Y)
Print #txtfNum, CC; "xyStart.x+"; Str$(NeckRightSideBot.X); CC; "xyStart.y+"; Str$(NeckRightSideBot.Y)
Print #txtfNum, CC; "xyStart.x+"; Str$(NeckLeftSideBot.X); CC; "xyStart.y+"; Str$(NeckLeftSideBot.Y)
Print #txtfNum, ");"
PR_AddEntityID "DartAndNeck"

'If 2" velcro then add this
If chkVelcro.Value = 1 Then
    PR_SetLayer "Notes"
    DartEnd.Y = DartStart.Y - 1
    DartEnd.X = DartEnd.X - 1.5
    PR_DrawText "2\"" VELCRO", DartEnd, 0.1, 0
    PR_SetLayer "TemplateLeft"
End If



End Sub

Private Sub PR_DrawNose()

Dim xyNoseTopLeft As xy, xyNoseTopRight As xy
'Dim xyNoseDiagTop As xy, xyNoseDiagBot As xy   'Now Global for use in drawing
                                                'eye flaps
Dim xyNoseBotLeft As xy, xyNoseBotRight As xy, xyNoseBotRight2 As xy
Dim xyTopCen As xy, xyBotCen As xy
Dim nStartAng#, nDeltaAng#, nRad#, nEndAng#
Dim Tempx#, TempY#
Dim TempX2#, TempY2#

    ' Top Part
    Tempx = Val(txtBotOfEyeX)
    TempY = Val(txtBotOfEyeY)
    PR_MakeXY xyNoseTopLeft, Tempx, TempY
    PR_MakeXY xyNoseTopRight, Tempx + 0.125, TempY

    ' Bottom Part
    TempY2 = Val(txtNoseBottomY)
    PR_MakeXY xyNoseBotLeft, xyNoseTopLeft.X, TempY2
    PR_MakeXY xyNoseBotRight, xyNoseTopLeft.X + 0.25, TempY2

    ' Diagonal Line
    PR_MakeXY xyNoseDiagTop, xyNoseTopRight.X + 0.125, xyNoseTopRight.Y - 0.125
    PR_MakeXY xyNoseDiagBot, xyNoseBotRight.X + 0.125, TempY2 + 0.125

    ' Top Curve
    PR_MakeXY xyTopCen, xyNoseTopRight.X, xyNoseDiagTop.Y

    'Bottom Curve
    PR_MakeXY xyBotCen, xyNoseBotRight.X, xyNoseDiagBot.Y
    nStartAng = 0
    nDeltaAng = -90
    nRad = 0.125
    
    ' Draw Nose Opening if No nose covering is requested
    If chkNoseCovering.Value = 0 Then
        PR_DrawLine xyNoseTopLeft, xyNoseTopRight
        PR_AddEntityID "NoseTop"
        PR_DrawLine xyNoseBotLeft, xyNoseBotRight
        PR_AddEntityID "NoseBottom"
        PR_DrawLine xyNoseDiagTop, xyNoseDiagBot
        PR_AddEntityID "NoseSide"
        PR_DrawArc xyTopCen, xyNoseTopRight, xyNoseDiagTop
        PR_AddEntityID "NoseTopCurve"
        Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyBotCen.X); CC; "xyStart.y +"; Str$(xyBotCen.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
        PR_AddEntityID "NoseBottomCurve"
    Else
        PR_MakeXY xyNoseBotRight2, xyNoseBotRight.X + 0.125, xyNoseBotRight.Y
        PR_DrawLine xyNoseBotLeft, xyNoseBotRight2
        PR_AddEntityID "NoseCoveringBottom"
    End If

End Sub

Private Sub PR_DrawNoseCovering()

' This Procedure Calculates and Draws the Nose Covering
' A Nose covering can only be drawn with the "Face Mask" and "Open Head Mask"
    
Dim xyNoseTip As xy, xyTempTip As xy, xyNoseTop As xy, xyNoseTipOut As xy
Dim xyNoseLenCenArc As xy, xyNoseLenStartArc As xy, xyNoseLenEndArc As xy
Dim xyNoseWidthCenArc As xy, xyNoseWidthStartArc As xy, xyNoseWidthEndArc As xy
Dim Adj1#, Opp1#, NoseDiff#, Adj2#, Opp2#, NoseStepAcross#
Dim NoseLen#, Theta#, NoseStepUp#, Delta#, NoseWidth#, X2#, y2#
Dim PI#, Opp3#, X1#, Y1#, TipAng#, TipHyp#, TipX#, TipY#
Dim TopX#, TopY#, TipAng2#, XA1#, YA1#, TipOpp#, TipAdj#, i%
Dim HeadPt1 As xy, HeadPt2 As xy, xyNoseRightTip As xy, xyNoseRightTip2 As xy
Dim TipAng3#
Dim NostralPt1 As xy, NostralPt2 As xy

Dim xyTmp As xy
Dim aAngle#, aTip#, aStartAngle#, aAngleInc#, nLength#, nRad#
Dim ii%


PI = 3.141592654

If chkNoseCovering.Value = 1 Then

   ' Set Values For Nose Length Arc
    Adj1 = Val(txtBotOfEyeX)
    Opp1 = Val(txtBotOfEyeY)
    PR_MakeXY xyNoseLenCenArc, Adj1 + 0.1875, Opp1 + 0.25
    NoseLen = FN_CmToInches(txtUnits, txtLengthOfNose)
   
   ' Set Values for Nose Width Arc
    Opp1 = Val(txtNoseBottomY)
    Adj1 = Val(txtBotOfEyeX) + 0.375
    PR_MakeXY xyNoseWidthCenArc, Adj1, Opp1
    NoseWidth = (FN_CmToInches(txtUnits, txtTipOfNose)) / 2
   'Get intersection point
    If Not FN_CalcCirCirInt(xyNoseLenCenArc, NoseLen, xyNoseWidthCenArc, NoseWidth, xyTempTip, xyTmp) Then
        MsgBox "Can't form nose with the given dimensions for the Nose length and width. The arcs do not intersect.", 48, "Head & Neck"
    End If

'PR_DrawMarker xyTempTip
'PR_DrawText "xyTempTip", xyTempTip, .05, 0
'PR_DrawMarker xyNoseLenCenArc
'PR_DrawText "xyNoseLenCenArc", xyNoseLenCenArc, .05, 0
'PR_DrawMarker xyNoseWidthCenArc
'PR_DrawText "xyNoseWidthCenArc", xyNoseWidthCenArc, .05, 0
'Print #txtfNum, "hEnt = AddEntity("; QQ; "circle"; QQ; ",xyStart.x+"; Str$(xyNoseLenCenArc.x); CC; "xyStart.y+"; Str$(xyNoseLenCenArc.y); ","; NoseLen; ");"
'Print #txtfNum, "hEnt = AddEntity("; QQ; "circle"; QQ; ",xyStart.x+"; Str$(xyNoseWidthCenArc.x); CC; "xyStart.y+"; Str$(xyNoseWidthCenArc.y); ","; NoseWidth; ");"
    
    ' Jobst Procedure 11.5.4 Extending Nose Length
    TipAng = 270 - FN_CalcAngle(xyNoseLenCenArc, xyTempTip)
    TipAng = TipAng * (PI / 180)
    TipX = (NoseLen + 0.25) * Sin(TipAng)
    TipY = (NoseLen + 0.25) * Cos(TipAng)
    PR_MakeXY xyNoseTip, xyNoseLenCenArc.X - TipX, xyNoseLenCenArc.Y - TipY


    TipAng2 = FN_CalcAngle(xyNoseTip, xyTempTip)
    TipAng2 = 90 - TipAng2
    TipAng2 = TipAng2 * (PI / 180)
    TipOpp = 0.1875 * (Sin(TipAng2))
    TipAdj = 0.1875 * (Cos(TipAng2))
    
    PR_MakeXY xyNoseTipOut, xyNoseTip.X - TipAdj, xyNoseTip.Y + TipOpp
    TipAng = FN_CalcAngle(xyTempTip, xyNoseLenCenArc) * (PI / 180)
    nLength = (xyNoseLenCenArc.Y - xyNoseTipOut.Y) - 0.125
    nLength = nLength / Sin(TipAng)
    PR_CalcPolar xyNoseTipOut, FN_CalcAngle(xyTempTip, xyNoseLenCenArc), nLength, xyNoseTop
'PR_DrawMarker xyNoseTop
'PR_DrawText "xyNoseTop-New", xyNoseTop, .05, 0
'PR_DrawMarker xyNoseTipOut
'PR_DrawText "xyNoseTipOut", xyNoseTipOut, .05, 0
'PR_DrawMarker xyNoseTip
'PR_DrawText "xyNoseTip", xyNoseTip, .05, 0
    
    ' Blend Nose into forehead
    TopX = txtNoseCoverX
    TopY = txtNoseCoverY
    PR_MakeXY HeadPt1, TopX, TopY
    PR_MakeXY HeadPt2, TopX - 0.03125, TopY - 0.25
'PR_DrawMarker HeadPt1
'PR_DrawText "HeadPt1", HeadPt1, .05, 0
'PR_DrawMarker HeadPt2
'PR_DrawText "HeadPt2", HeadPt2, .05, 0
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "open"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(HeadPt1.X); CC; "xyStart.y+"; Str$(HeadPt1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(HeadPt2.X); CC; "xyStart.y+"; Str$(HeadPt2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseTop.X); CC; "xyStart.y+"; Str$(xyNoseTop.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "NoseCoveringBridge"
    
    TipX = 0.25 * Cos(TipAng2)
    TipY = 0.25 * Sin(TipAng2)
    PR_MakeXY xyNoseRightTip, xyNoseTip.X + TipX, xyNoseTip.Y - TipY
    
    TipAng3 = (90 - TipAng2 * (180 / PI)) * (PI / 180)
    TipX = 0.25 * Cos(TipAng3)
    TipY = 0.25 * Sin(TipAng3)
    PR_MakeXY xyNoseRightTip2, xyNoseRightTip.X + TipX, xyNoseRightTip.Y + TipY
    
    ' Draw nose
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseTop.X); CC; "xyStart.y+"; Str$(xyNoseTop.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseTipOut.X); CC; "xyStart.y+"; Str$(xyNoseTipOut.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseTip.X); CC; "xyStart.y+"; Str$(xyNoseTip.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseRightTip.X); CC; "xyStart.y+"; Str$(xyNoseRightTip.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(xyNoseRightTip2.X); CC; "xyStart.y+"; Str$(xyNoseRightTip2.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "NoseCovering"
    
   ' Calculate and draw nostril
   ' Mods GG 4.Mar.96
    aTip = FN_CalcAngle(xyNoseLenCenArc, xyNoseTip) + 90
    nLength = FN_CalcLength(xyNoseRightTip2, xyNoseWidthCenArc)
    aAngle = FN_CalcAngle(xyNoseRightTip2, xyNoseWidthCenArc)
    If aAngle > 180 Then
        aAngle = aAngle - aTip
    Else
        aAngle = aAngle + (360 - aTip)
    End If
    nRad = (nLength / 2) / Cos(aAngle * (PI / 180))
    PR_CalcPolar xyNoseRightTip2, aTip, nRad, NostralPt1
    aStartAngle = FN_CalcAngle(NostralPt1, xyNoseRightTip2)
    aAngle = FN_CalcAngle(NostralPt1, xyNoseWidthCenArc)
    If aAngle > 270 Then aAngle = aAngle - 360 'NB Sign in this case will add
    aAngleInc = (aStartAngle - aAngle) / 3

'PR_DrawMarker xyNoseRightTip2
'PR_DrawText "xyNoseRightTip2", xyNoseRightTip2, .05, 0
'PR_DrawMarker xyNoseTip
'PR_DrawText "xyNoseTip", xyNoseTip, .05, 0
'PR_DrawMarker xyNoseWidthCenArc
'PR_DrawText "xyNoseWidthCenArc", xyNoseWidthCenArc, .05, 0
'PR_DrawMarker NostralPt1
'PR_DrawText "NostralPt1", NostralPt1, .05, 0
'PR_DrawMarker xyNoseRightTip2
'PR_DrawText "xyNoseRightTip2", xyNoseRightTip2, .05, 0
'PR_DrawMarker xyNoseLenCenArc
'PR_DrawText "xyNoseLenCenArc", xyNoseLenCenArc, .05, 0
    
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
    For ii = 0 To 3
        PR_CalcPolar NostralPt1, aStartAngle - (aAngleInc * ii), nRad, NostralPt2
        Print #txtfNum, CC; "xyStart.x+"; Str$(NostralPt2.X); CC; "xyStart.y+"; Str$(NostralPt2.Y)
    Next ii
    Print #txtfNum, ");"
    PR_AddEntityID "Nostril"

End If

End Sub



Private Sub PR_DrawOpenFace()

Dim xyFaceStart As xy, xyOpeningTop As xy, xyOpeningBottom As xy
Dim xyLipBottom As xy, xyLipTop As xy
Dim xyPoint1 As xy, xyPoint2 As xy
Dim xDist#, YDist#
Dim X1#, Y1#, ii%
Dim XStep#, YStep#

PR_MakeXY xyFaceStart, -LeftArcAdjOut(txtRadiusNo), 0
PR_MakeXY xyOpeningTop, -LeftArcAdjOut(txtRadiusNo), -0.5
PR_DrawLine xyFaceStart, xyOpeningTop
PR_AddEntityID "OpenFaceEyeBrow"

OpenFaceProfileX(1) = xyOpeningTop.X
OpenFaceProfileY(1) = xyOpeningTop.Y
XStep = (0.375 + txtEyeWidth) / 2
OpenFaceProfileX(2) = xyOpeningTop.X + XStep
OpenFaceProfileY(2) = xyOpeningTop.Y
XStep = (txtEyeWidth)

OpenFaceProfileX(3) = xyOpeningTop.X + XStep
OpenFaceProfileY(3) = xyOpeningTop.Y - 0.0625
XStep = txtEyeWidth + 0.25

YStep = -0.375
OpenFaceProfileX(4) = xyOpeningTop.X + XStep
OpenFaceProfileY(4) = xyOpeningTop.Y + YStep

If chkLipStrap.Value = 1 Then
    X1 = -LeftArcAdjOut(txtRadiusNo)
    Y1 = Val(txtMouthHeight)
    PR_MakeXY xyLipBottom, X1, Y1
    PR_MakeXY xyLipTop, X1, Y1 + Val(txtLipStrapWidth)
    PR_DrawLine xyLipBottom, xyLipTop
    PR_AddEntityID "OpenFaceLipStrap"
    PR_MakeXY xyOpeningBottom, xyLipTop.X, xyLipTop.Y
Else
    X1 = txtMouthRightX
    Y1 = Val(txtMouthHeight)
    PR_MakeXY xyOpeningBottom, X1, Y1
End If

PR_MakeXY xyPoint1, OpenFaceProfileX(4), OpenFaceProfileY(4)
PR_MakeXY xyPoint2, xyOpeningBottom.X, OpenFaceProfileY(4)
xDist = FN_CalcLength(xyPoint1, xyPoint2)
XStep = -xDist / 6

PR_MakeXY xyPoint1, OpenFaceProfileX(4), OpenFaceProfileY(4)
PR_MakeXY xyPoint2, OpenFaceProfileX(4), xyOpeningBottom.Y
YDist = FN_CalcLength(xyPoint1, xyPoint2)
YStep = -YDist / 5

OpenFaceProfileX(5) = OpenFaceProfileX(4) + 0.0625
OpenFaceProfileY(5) = OpenFaceProfileY(4) + YStep

OpenFaceProfileX(6) = OpenFaceProfileX(5) - 0.0625
OpenFaceProfileY(6) = OpenFaceProfileY(5) + YStep

OpenFaceProfileX(7) = OpenFaceProfileX(6) + XStep + 0.0625
OpenFaceProfileY(7) = OpenFaceProfileY(6) + YStep

OpenFaceProfileX(8) = OpenFaceProfileX(7) + XStep - 0.0625
OpenFaceProfileY(8) = OpenFaceProfileY(7) + YStep

OpenFaceProfileX(9) = OpenFaceProfileX(8) + (2 * XStep)
OpenFaceProfileY(9) = xyOpeningBottom.Y + 0.09375

OpenFaceProfileX(10) = xyOpeningBottom.X
OpenFaceProfileY(10) = xyOpeningBottom.Y

Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "fitted"; QQ
For ii = 1 To 10
    Print #txtfNum, CC; "xyStart.x+"; Str$(OpenFaceProfileX(ii)); CC; "xyStart.y+"; Str$(OpenFaceProfileY(ii))
Next ii
Print #txtfNum, ");"
PR_AddEntityID "OpenFace"

End Sub

Private Sub PR_DrawOpenFaceWithEyes()

Dim xyFaceStart As xy, xyOpeningTop As xy, xyOpeningBottom As xy
Dim xyCen As xy, xyMouthRight As xy
Dim Y1#, y2#, X1#, nEndAng#
Dim nDeltaAng#, nStartAng#, nRad#

Y1 = txtBotOfEyeY
y2 = txtMouthHeight
X1 = txtMouthRightX
PR_MakeXY xyFaceStart, -LeftArcAdjOut(txtRadiusNo), 0
PR_MakeXY xyOpeningTop, -LeftArcAdjOut(txtRadiusNo), Y1
PR_MakeXY xyOpeningBottom, X1, y2 + 0.125
PR_MakeXY xyMouthRight, X1 - 0.125, y2
PR_MakeXY xyCen, xyOpeningBottom.X - 0.125, xyOpeningBottom.Y

'Bottom Curve
nStartAng = 0
nDeltaAng = -90
nRad = 0.125
Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCen.X); CC; "xyStart.y +"; Str$(xyCen.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
PR_AddEntityID "OpenFaceBottomCurve"

PR_DrawLine xyFaceStart, xyOpeningTop
PR_AddEntityID "OpenFaceEyeBrow"

PR_DrawLine xyOpeningBottom, xyOpeningTop
PR_AddEntityID "OpenFace"

End Sub

Private Sub PR_DrawOpenHead()

Dim Point1 As xy, Point2 As xy, Point3 As xy
Dim xyClosureText As xy, Circum#, NormCircum#, Total, Cen As xy
Dim Elastic#, xyTextPoint As xy, ArcAngle1#, ArcAngle2#, PI#
Dim Pt1 As xy, Pt2 As xy, Radius#, ArcLen#, TotalArcLen#, ArcAngle#

PI = 3.141592654
PR_MakeXY Cen, 0, 0
Circum = Val(txtRadiusNo)
NormCircum = Val(txtRadiusNo)

' Set For Zipper
If chkZipper.Value = 1 And Val(txtRadiusNo) > 3 Then
    Circum = Circum - 3
ElseIf chkZipper.Value = 1 And Val(txtRadiusNo) <= 3 Then
    Circum = Circum + 17
End If

If chkOpenHeadMask.Value = 1 Then
    PR_MakeXY Point1, -LeftArcAdjOut(NormCircum), 0
    PR_MakeXY Point2, -LeftArcAdjOut(NormCircum), 0.5
    PR_MakeXY Point3, RightArcOpenAdjUp(Circum), 0.5
    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(Point1.X); CC; "xyStart.y+"; Str$(Point1.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Point2.X); CC; "xyStart.y+"; Str$(Point2.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Point3.X); CC; "xyStart.y+"; Str$(Point3.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "OpenHeadTop"
    
    'top arc
    PR_MakeXY Pt1, -TopArcAdj(txtRadiusNo), TopArcOpp(txtRadiusNo)
    PR_MakeXY Pt2, TopArcAdj(txtRadiusNo), TopArcOpp(txtRadiusNo)
    ArcAngle1 = FN_CalcAngle(Cen, Pt1)
    ArcAngle2 = FN_CalcAngle(Cen, Pt2)
    ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
    Radius = TopArcHyp(txtRadiusNo)
    ArcLen = Radius * ArcAngle
    TotalArcLen = ArcLen
    
    ' Left Arc
    PR_MakeXY Pt1, -LeftArcAdjOut(txtRadiusNo), LeftArcOppOut(txtRadiusNo)
    PR_MakeXY Pt2, -LeftArcAdjIn(txtRadiusNo), LeftArcOppIn(txtRadiusNo)
    ArcAngle1 = FN_CalcAngle(Cen, Pt1)
    ArcAngle2 = FN_CalcAngle(Cen, Pt2)
    ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
    Radius = LeftArcHyp(txtRadiusNo)
    ArcLen = Radius * ArcAngle
    TotalArcLen = TotalArcLen + ArcLen

    
    ' Left Line
    ArcLen = Pt1.Y - Point2.Y
    TotalArcLen = TotalArcLen + ArcLen

    'Top Right arc
    PR_MakeXY Pt1, TopRightArcAdjIn(txtRadiusNo), TopRightArcOppIn(txtRadiusNo)
    PR_MakeXY Pt2, TopRightArcAdjOut(txtRadiusNo), TopRightArcOppOut(txtRadiusNo)
    ArcAngle1 = FN_CalcAngle(Cen, Pt1)
    ArcAngle2 = FN_CalcAngle(Cen, Pt2)
    ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
    Radius = TopArcHyp(txtRadiusNo)
    ArcLen = Radius * ArcAngle
    TotalArcLen = TotalArcLen + ArcLen

    
    'Right arc
    PR_MakeXY Pt1, RightArcAdjUp(txtRadiusNo), RightArcOppUp(txtRadiusNo)
    ArcAngle1 = FN_CalcAngle(Cen, Pt1)
    ArcAngle2 = FN_CalcAngle(Cen, Point3)
    ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)

    Radius = Sqr((RightArcOpenAdjUp(Circum) * RightArcOpenAdjUp(Circum)) + 0.25)
    ArcLen = Radius * ArcAngle
    TotalArcLen = TotalArcLen + ArcLen
    Total = FN_InchesToText(fnRoundInches(TotalArcLen))

    PR_MakeXY xyTextPoint, Point1.X + 2, Point1.Y
    PR_SetLayer "Notes"
    PR_DrawText Total & Chr$(39) & Chr$(39) & " Elastic Strap", xyTextPoint, 0.1, 0
    PR_SetLayer "TemplateLeft"

End If

End Sub

Private Sub PR_DrawRightArc()

Dim nRad#, nStartAng#, nEndAng#
Dim nDeltaAng#, Circum#
Dim xyStartArc As xy, xyEndArc As xy, xyCentreArc As xy
Dim xyClosureText As xy, OpenAcross#
Dim AcrossUp#, UpwardsUp#
Dim AcrossDown#, UpwardsDown#

Circum = Val(txtRadiusNo)
PR_MakeXY xyCentreArc, 0, 0

' Set For Zipper
If chkZipper.Value = 1 And Val(txtRadiusNo) > 3 Then
    Circum = Circum - 3
'    PR_MakeXY xyClosureText, RightArcAdjUp(Circum) - .35, .25
'    PR_SetLayer "Notes"
'    PR_DrawText "Zipper", xyClosureText, .25, -90
'    PR_SetLayer "TemplateLeft"
ElseIf chkZipper.Value = 1 And Val(txtRadiusNo) <= 3 Then
    Circum = Circum + 17
'    PR_MakeXY xyClosureText, RightArcAdjUp(Circum) - .35, .25
'    PR_SetLayer "Notes"
'    PR_DrawText "Zipper", xyClosureText, .25, -90
'    PR_SetLayer "TemplateLeft"
End If

' Set For 2" wide Velcro
If chkVelcro.Value = 1 Then
    Circum = Circum + 2
End If

'Draw text for both ZIPPER and VELCRO"
PR_MakeXY xyClosureText, RightArcAdjUp(Circum) - 0.35, 0.25
PR_SetLayer "Notes"
If chkZipper.Value = 1 Then
    PR_DrawText "ZIPPER", xyClosureText, 0.25, -90
Else
    PR_DrawText "VELCRO", xyClosureText, 0.25, -90
End If
PR_SetLayer "TemplateLeft"

AcrossUp = RightArcAdjUp(Circum)
UpwardsUp = RightArcOppUp(Circum)
OpenAcross = RightArcOpenAdjUp(Circum)
AcrossDown = RightArcAdjDown(Circum)
UpwardsDown = RightArcOppDown(Circum)

If chkOpenHeadMask.Value = 1 Then
    PR_MakeXY xyEndArc, OpenAcross, 0.5
Else
    PR_MakeXY xyEndArc, AcrossUp, UpwardsUp
End If

PR_MakeXY xyStartArc, AcrossDown, -UpwardsDown
nRad = FN_CalcLength(xyCentreArc, xyStartArc)
nStartAng = FN_CalcAngle(xyCentreArc, xyStartArc)
nStartAng = 360 - nStartAng
nEndAng = FN_CalcAngle(xyCentreArc, xyEndArc)
nDeltaAng = -(nEndAng + nStartAng)
Print #txtfNum, "hEnt = AddEntity("; QQ; "arc"; QC; "xyStart.x +"; Str$(xyCentreArc.X); CC; "xyStart.y +"; Str$(xyCentreArc.Y); CC; Str$(nRad); CC; Str$(nEndAng); CC; Str$(nDeltaAng); ");"
PR_AddEntityID "RightArc"

End Sub

Private Sub PR_DrawRightCutOut()
Dim AcrossLeft#, UpwardsLeft#
Dim AcrossRight#, UpwardsRight#
Dim CentreX#, CentreY#
Dim Centre As xy, LeftTip As xy, RightTip As xy

If chkOpenHeadMask.Value = 0 Then
    CentreX = 3.8 / 2.54
    CentreY = 3.8 / 2.54
    PR_MakeXY Centre, CentreX, CentreY

    AcrossLeft = TopArcAdj(txtRadiusNo)
    UpwardsLeft = TopArcOpp(txtRadiusNo)
    AcrossRight = TopRightArcAdjIn(txtRadiusNo)
    UpwardsRight = TopRightArcOppIn(txtRadiusNo)

    PR_MakeXY LeftTip, AcrossLeft, UpwardsLeft
    PR_MakeXY RightTip, AcrossRight, UpwardsRight

    Print #txtfNum, "hEnt = AddEntity("; QQ; "poly"; QCQ; "polyline"; QQ
    Print #txtfNum, CC; "xyStart.x+"; Str$(LeftTip.X); CC; "xyStart.y+"; Str$(LeftTip.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(Centre.X); CC; "xyStart.y+"; Str$(Centre.Y)
    Print #txtfNum, CC; "xyStart.x+"; Str$(RightTip.X); CC; "xyStart.y+"; Str$(RightTip.Y)
    Print #txtfNum, ");"
    PR_AddEntityID "RightCutOut"
End If

End Sub

Private Sub PR_DrawRightJoin()

Dim AcrossLeft#, UpwardsLeft#
Dim AcrossRight#, UpwardsRight#
Dim Start  As xy, Finish As xy
Dim Circum#

If chkOpenHeadMask.Value = 0 Then
    Circum = Val(txtRadiusNo)
    ' Set For Zipper if Circumference is Less than 36
    If chkZipper.Value = 1 And Circum > 3 Then
        Circum = Circum - 3
    ElseIf chkZipper.Value = 1 And Circum <= 3 Then
        Circum = Circum + 17
    End If
    ' Set For 2" wide Velcro
    If chkVelcro.Value = 1 Then
        Circum = Circum + 2
    End If
    
    
    AcrossLeft = TopRightArcAdjOut(txtRadiusNo)
    UpwardsLeft = TopRightArcOppOut(txtRadiusNo)
    AcrossRight = RightArcAdjUp(Circum)
    UpwardsRight = RightArcOppUp(Circum)

    PR_MakeXY Start, AcrossLeft, UpwardsLeft
    PR_MakeXY Finish, AcrossRight, UpwardsRight
    PR_DrawLine Start, Finish
    PR_AddEntityID "RightArcJoin"
End If

End Sub

Private Sub PR_DrawRightLine()

Dim xyStart As xy, xyFinish As xy
Dim Across#, Down#
Dim Circum#, LineLen#

Circum = Val(txtRadiusNo)

' Check For Zipper if Circumference is Less than 36
If chkZipper.Value = 1 And Val(txtRadiusNo) > 3 Then
    Circum = Circum - 3
ElseIf chkZipper.Value = 1 And Circum <= 3 Then
    Circum = Circum + 17
End If
    ' Set For 2" wide Velcro
    If chkVelcro.Value = 1 Then
        Circum = Circum + 2
    End If


LineLen = 1.35 / 2.54
Across = RightArcAdjDown(Circum)
Down = -RightArcOppDown(Circum)
 PR_MakeXY xyStart, Across, Down
PR_MakeXY xyFinish, Across - LineLen, Down
PR_DrawLine xyStart, xyFinish
PR_AddEntityID "RightArcNeckJoin"

End Sub

Private Sub PR_DrawText(sText As Variant, xyInsert As xy, nHeight As Variant, nAngle As Variant)
   'To the DRAFIX macro file (given by the global txtfNum).
   'Write the syntax to draw TEXT at the given height.
   '
   'For this to work it assumes that the following DRAFIX variables
   'are defined
   '    XY      xyStart
    
    Dim nWidth
    nWidth = nHeight * txtCurrTextAspect
    Print #txtfNum, "AddEntity("; QQ; "text"; QCQ; sText; QC; "xyStart.x+"; Str$(xyInsert.X); CC; "xyStart.y+"; Str$(xyInsert.Y); CC; nWidth; CC; nHeight; CC; nAngle; ");"
    
End Sub

Private Sub PR_DrawTopArc()
' This Procedure Draws the top Arc of the Face Mask
' It uses the value held in txtRadiusNo to determine which
' Arc to Draw.  The Value Held in txtRadius must lie between
' 1 and 19 inclusive.

Dim xyStartArc As xy, xyEndArc As xy, xyCentreArc As xy
Dim Across#, Upwards#

If chkOpenHeadMask.Value = 0 Then
    PR_MakeXY xyCentreArc, 0, 0
   
    Across = TopArcAdj(txtRadiusNo)
    Upwards = TopArcOpp(txtRadiusNo)
    
    PR_MakeXY xyStartArc, -Across, Upwards
    PR_MakeXY xyEndArc, Across, Upwards
    PR_DrawArc xyCentreArc, xyStartArc, xyEndArc
    PR_AddEntityID "TopArc"
End If
End Sub

Private Sub PR_DrawTopRightArc()

' This Procedure Draws the top Arc of the Face Mask
' It uses the value held in txtRadiusNo to determine which
' Arc to Draw.  The Value Held in txtRadius must lie between
' 1 and 19 inclusive.

Dim xyStartArc As xy, xyEndArc As xy, xyCentreArc As xy
Dim AcrossIn#, UpwardsIn#
Dim AcrossOut#, UpwardsOut#
If chkOpenHeadMask.Value = 0 Then
    PR_MakeXY xyCentreArc, 0, 0
    AcrossIn = TopRightArcAdjIn(txtRadiusNo)
    UpwardsIn = TopRightArcOppIn(txtRadiusNo)

    AcrossOut = TopRightArcAdjOut(txtRadiusNo)
    UpwardsOut = TopRightArcOppOut(txtRadiusNo)

    PR_MakeXY xyStartArc, AcrossIn, UpwardsIn
    PR_MakeXY xyEndArc, AcrossOut, UpwardsOut
    PR_DrawArc xyCentreArc, xyStartArc, xyEndArc
    PR_AddEntityID "TopRightArc"
End If

End Sub

Private Sub PR_FaceMask()
Dim xyTextPoint As xy

If optFaceMask.Value = True Then
    PR_DrawTopArc
    PR_DrawRightCutOut
    PR_DrawTopRightArc
    PR_DrawRightJoin
    PR_DrawRightArc
    PR_DrawRightLine
    PR_DrawLeftArc
    PR_DrawLeftCutOut
    PR_DrawOpenHead
    PR_SetChart1Details
    PR_DrawEyeOpening
    PR_DrawEars
    PR_DrawNeckAndDart
    PR_DrawNose
    PR_DrawForeHead
    PR_DrawChin
    PR_DrawNoseCovering
    PR_DrawEarFlaps
    PR_DrawEyeFlaps
    PR_DrawLipCovering
    PR_StampLining
    PR_NeckElastic
    PR_MakeXY xyTextPoint, 0.3, 0.5
    PR_TemplateDetails xyTextPoint, "Face Mask"

End If
End Sub

Private Sub PR_LoadBotFaceArrays()
' This Procedure Loads the Arrays that hold the positions
' of the end points for the radials of the Face Line
 
Dim BotStart#, BotStep#, i%

BotStart = 10.15 / 2.54
BotStep = (10.1 / 2.54) / 16

For i = 1 To 19
    BotFaceAdj(i) = -(LeftArcAdjOut(i))
    BotFaceOpp(i) = LeftArcOppOut(i) - BotStart
    BotStart = BotStart + BotStep
Next i

End Sub

Private Sub PR_LoadFaceMaskChart1()

' This Procedure Loads Face Mask Chart 1 into an
' array Named: FaceMaskChartA

Dim nMidMouth#, i%

nMidMouth = 2
For i = 1 To 25
    ' Column 1
    FaceMaskChartA(1, i) = nMidMouth
    nMidMouth = nMidMouth + 0.125
    ' Column 2
    If (i = 1) Or (i = 2) Then
        FaceMaskChartA(2, i) = 0.375
    ElseIf (i >= 3) And (i <= 7) Then
        FaceMaskChartA(2, i) = 0.5
    ElseIf (i >= 8) And (i <= 12) Then
        FaceMaskChartA(2, i) = 0.625
    ElseIf (i >= 13) And (i <= 17) Then
        FaceMaskChartA(2, i) = 0.75
    ElseIf (i >= 18) And (i <= 22) Then
        FaceMaskChartA(2, i) = 0.875
    ElseIf (i >= 23) And (i <= 25) Then
        FaceMaskChartA(2, i) = 1
    End If
    ' Column 3
    If (i = 1) Or (i = 2) Then
        FaceMaskChartA(3, i) = 0.5
    ElseIf (i >= 3) And (i <= 6) Then
        FaceMaskChartA(3, i) = 0.625
    ElseIf (i >= 7) And (i <= 10) Then
        FaceMaskChartA(3, i) = 0.75
    ElseIf (i >= 11) And (i <= 14) Then
        FaceMaskChartA(3, i) = 0.825
    ElseIf (i >= 15) And (i <= 18) Then
        FaceMaskChartA(3, i) = 1
    ElseIf (i >= 19) And (i <= 22) Then
        FaceMaskChartA(3, i) = 1.125
    ElseIf (i >= 23) And (i <= 25) Then
        FaceMaskChartA(3, i) = 1.25
    End If
Next i

End Sub

Private Sub PR_LoadFaceMaskChart2()

' This Procedure Loads Face Mask Chart 2 into an
' array Named: FaceMaskChart2

Dim nCircum%, i%

nCircum = 30
For i = 1 To 16
    ' Column 1
    FaceMaskChart2(1, i) = nCircum
    nCircum = nCircum + 2
    ' Column 2
    If (i = 1) Then
        FaceMaskChart2(2, i) = 0.75
    ElseIf (i >= 2) And (i <= 4) Then
        FaceMaskChart2(2, i) = 0.875
    ElseIf (i >= 5) And (i <= 7) Then
        FaceMaskChart2(2, i) = 1
    ElseIf (i >= 8) And (i <= 9) Then
        FaceMaskChart2(2, i) = 1.125
    ElseIf (i >= 10) And (i <= 12) Then
        FaceMaskChart2(2, i) = 1.25
    ElseIf (i >= 13) And (i <= 14) Then
        FaceMaskChart2(2, i) = 1.375
    ElseIf (i >= 15) And (i <= 16) Then
        FaceMaskChart2(2, i) = 1.5
    End If
    ' Column 3
    If (i = 1) Then
        FaceMaskChart2(3, i) = 1.625
    ElseIf (i = 2) Or (i = 3) Then
        FaceMaskChart2(3, i) = 1.75
    ElseIf (i = 4) Or (i = 5) Then
        FaceMaskChart2(3, i) = 1.875
    ElseIf (i = 6) Then
        FaceMaskChart2(3, i) = 2
    ElseIf (i = 7) Or (i = 8) Then
        FaceMaskChart2(3, i) = 2.125
    ElseIf (i = 9) Or (i = 10) Then
        FaceMaskChart2(3, i) = 2.25
    ElseIf (i = 11) Then
        FaceMaskChart2(3, i) = 2.375
    ElseIf (i = 12) Or (i = 13) Then
        FaceMaskChart2(3, i) = 2.5
    ElseIf (i = 14) Then
        FaceMaskChart2(3, i) = 2.625
    ElseIf (i = 15) Or (i = 16) Then
        FaceMaskChart2(3, i) = 2.75
    End If
Next i

End Sub

Private Sub PR_LoadFaceMaskChart3()
' This Procedure Loads Face Mask Chart 3 into an
' array Named: FaceMaskChart3
Dim nNeck As Single, i%

nNeck = 4

' Column 3
For i = 1 To 47
    FaceMaskChart3(3, i) = nNeck
    nNeck = nNeck + 0.125
Next i

' Column 1
FaceMaskChart3(1, 1) = 8
FaceMaskChart3(1, 2) = 8.25
FaceMaskChart3(1, 3) = 8.5
FaceMaskChart3(1, 4) = 8.75
FaceMaskChart3(1, 5) = 9
FaceMaskChart3(1, 6) = 9.25
FaceMaskChart3(1, 7) = 9.625
FaceMaskChart3(1, 8) = 9.875
FaceMaskChart3(1, 9) = 10.125
FaceMaskChart3(1, 10) = 10.375
FaceMaskChart3(1, 11) = 10.625
FaceMaskChart3(1, 12) = 10.875
FaceMaskChart3(1, 13) = 11.125
FaceMaskChart3(1, 14) = 11.375
FaceMaskChart3(1, 15) = 11.625
FaceMaskChart3(1, 16) = 11.875
FaceMaskChart3(1, 17) = 12.125
FaceMaskChart3(1, 18) = 12.375
FaceMaskChart3(1, 19) = 12.75
FaceMaskChart3(1, 20) = 13
FaceMaskChart3(1, 21) = 13.25
FaceMaskChart3(1, 22) = 13.5
FaceMaskChart3(1, 23) = 13.75
FaceMaskChart3(1, 24) = 14
FaceMaskChart3(1, 25) = 14.25
FaceMaskChart3(1, 26) = 14.5
FaceMaskChart3(1, 27) = 14.75
FaceMaskChart3(1, 28) = 15
FaceMaskChart3(1, 29) = 15.25
FaceMaskChart3(1, 30) = 15.5
FaceMaskChart3(1, 31) = 15.875
FaceMaskChart3(1, 32) = 16.125
FaceMaskChart3(1, 33) = 16.375
FaceMaskChart3(1, 34) = 16.625
FaceMaskChart3(1, 35) = 16.875
FaceMaskChart3(1, 36) = 17.125
FaceMaskChart3(1, 37) = 17.375
FaceMaskChart3(1, 38) = 17.625
FaceMaskChart3(1, 39) = 17.875
FaceMaskChart3(1, 40) = 18.125
FaceMaskChart3(1, 41) = 18.375
FaceMaskChart3(1, 42) = 18.625
FaceMaskChart3(1, 43) = 19
FaceMaskChart3(1, 44) = 19.25
FaceMaskChart3(1, 45) = 19.5
FaceMaskChart3(1, 46) = 19.75
FaceMaskChart3(1, 47) = 20

' Column 2
FaceMaskChart3(2, 1) = 8.125
FaceMaskChart3(2, 2) = 8.75
FaceMaskChart3(2, 3) = 8.625
FaceMaskChart3(2, 4) = 8.875
FaceMaskChart3(2, 5) = 9.125
FaceMaskChart3(2, 6) = 9.5
FaceMaskChart3(2, 7) = 9.75
FaceMaskChart3(2, 8) = 10
FaceMaskChart3(2, 9) = 10.25
FaceMaskChart3(2, 10) = 10.5
FaceMaskChart3(2, 11) = 10.75
FaceMaskChart3(2, 12) = 11
FaceMaskChart3(2, 13) = 11.25
FaceMaskChart3(2, 14) = 11.5
FaceMaskChart3(2, 15) = 11.75
FaceMaskChart3(2, 16) = 12
'FaceMaskChart3(2, 17) = 12.125   'GG Bugfix
FaceMaskChart3(2, 17) = 12.25   'GG Bugfix
FaceMaskChart3(2, 18) = 12.625
FaceMaskChart3(2, 19) = 12.875
FaceMaskChart3(2, 20) = 13.125
FaceMaskChart3(2, 21) = 13.375
FaceMaskChart3(2, 22) = 13.625
FaceMaskChart3(2, 23) = 13.875
FaceMaskChart3(2, 24) = 14.125
FaceMaskChart3(2, 25) = 14.375
FaceMaskChart3(2, 26) = 14.625
FaceMaskChart3(2, 27) = 14.875
FaceMaskChart3(2, 28) = 15.125
FaceMaskChart3(2, 29) = 15.375
FaceMaskChart3(2, 30) = 15.75
FaceMaskChart3(2, 31) = 16
FaceMaskChart3(2, 32) = 16.25
FaceMaskChart3(2, 33) = 16.5
FaceMaskChart3(2, 34) = 16.75
FaceMaskChart3(2, 35) = 17
FaceMaskChart3(2, 36) = 17.25
FaceMaskChart3(2, 37) = 17.5
FaceMaskChart3(2, 38) = 17.75
FaceMaskChart3(2, 39) = 18
FaceMaskChart3(2, 40) = 18.25
FaceMaskChart3(2, 41) = 18.5
FaceMaskChart3(2, 42) = 18.875
FaceMaskChart3(2, 43) = 19.125
FaceMaskChart3(2, 44) = 19.375
FaceMaskChart3(2, 45) = 19.625
FaceMaskChart3(2, 46) = 19.875
FaceMaskChart3(2, 47) = 0

End Sub

Private Sub PR_LoadLeftArcArrays()
' This Procedure Loads the Arrays that hold the positions
' of the start and end points for the radials of the Left Arc
 
Dim TopArcUpwards#, LeftArcUpwardsIn#
Dim LeftArcAcrossIn#, LeftUpwardsIn#, i%
Dim LeftArcAcrossOut#, LeftArcUpwardsOut#
Dim LeftAdjIn#, LeftAdjOut#, LeftHyp#

TopArcUpwards = 0.50625 / 2.54
LeftArcAcrossIn = 0.456275 / 2.54
LeftArcAcrossOut = 0.4796875 / 2.54
LeftArcUpwardsIn = 0.2375 / 2.54
LeftArcUpwardsOut = 0.146875 / 2.54

' Left Arc Values
LeftAdjIn = 6.25 / 2.54
LeftAdjOut = 7.65 / 2.54
LeftHyp = 8.03 / 2.54

' Load Arrays
For i = 1 To 19

    LeftArcAdjIn(i) = LeftAdjIn
    LeftArcAdjOut(i) = LeftAdjOut
    LeftArcHyp(i) = LeftHyp
    LeftArcOppIn(i) = Sqr((LeftHyp * LeftHyp) - (LeftAdjIn * LeftAdjIn))
    LeftArcOppOut(i) = Sqr((LeftHyp * LeftHyp) - (LeftAdjOut * LeftAdjOut))
    LeftAdjIn = LeftAdjIn + LeftArcAcrossIn
    LeftAdjOut = LeftAdjOut + LeftArcAcrossOut
    LeftHyp = LeftHyp + TopArcUpwards
   
Next i

End Sub

Private Sub PR_LoadNeckArrays()

' This Procedure Loads the array NeckRight(index)
' with the coordinates of the template for the right
' side of the neck

Dim OppStart#, AdjStart#
Dim OppEnd#, AdjEnd#, i%
Dim OppStep#, AdjStep#

OppStart = -7.65 / 2.54
OppEnd = -15.25 / 2.54
AdjStart = 6.35 / 2.54
AdjEnd = 14 / 2.54
OppStep = (OppEnd - OppStart) / 16
AdjStep = (AdjEnd - AdjStart) / 16

For i = 1 To 19
    NeckRightOpp(i) = OppStart
    NeckRightAdj(i) = AdjStart
    OppStart = OppStart + OppStep
    AdjStart = AdjStart + AdjStep
Next i

End Sub

Private Sub PR_LoadRightArcArrays()

' This Procedure Loads the Arrays that hold the positions
' of the start and end points for the radials of the Right Arc

Dim TopArcUpwards#, TopAdj#, TopHyp#
Dim LeftAdjIn#, LeftAdjOut#, LeftHyp#
Dim RightArcAcrossTop#, RightArcUpwards#
Dim RightArcAcrossBot#, RightArcDownwards#
Dim RightAdjUp#, RightAdjDown#, RightHyp#
Dim RightAdjUp2#, RightAdjDown2#, RightHyp2#
Dim OpenAcross#, i%, j%

RightArcAcrossTop = 0.45 / 2.54
RightArcAcrossBot = 0.4625 / 2.54
RightArcUpwards = 0.23125 / 2.54
RightArcDownwards = 0.23125 / 2.54
TopArcUpwards = 0.50625 / 2.54

' Right Arc Values
RightAdjUp = 7.85 / 2.54
RightAdjDown = 6.4 / 2.54
RightHyp = 9.03 / 2.54
OpenAcross = 8.85 / 2.54

RightAdjUp2 = (7.85 / 2.54) - 3 * (RightArcAcrossTop)
RightAdjDown2 = (6.4 / 2.54) - 3 * (RightArcAcrossBot)
RightHyp2 = (9.03 / 2.54) - 3 * (TopArcUpwards)

' Load Right Arc Arrays
For i = 1 To 19
    RightArcAdjUp(i) = RightAdjUp
    RightArcAdjDown(i) = RightAdjDown
    RightArcOpenAdjUp(i) = Sqr((RightHyp * RightHyp) - (0.25))
    RightArcOppUp(i) = Sqr((RightHyp * RightHyp) - (RightAdjUp * RightAdjUp))
    RightArcOppDown(i) = Sqr((RightHyp * RightHyp) - (RightAdjDown * RightAdjDown))
    RightAdjUp = RightAdjUp + RightArcAcrossTop
    RightAdjDown = RightAdjDown + RightArcAcrossBot
    RightHyp = RightHyp + TopArcUpwards
Next i

For i = 18 To 20
    j = 20
    RightArcAdjUp(i) = RightAdjUp2
    RightArcAdjDown(i) = RightAdjDown2
    RightHyp2 = RightHyp2
    RightArcOpenAdjUp(i) = Sqr((RightHyp2 * RightHyp2) - (0.25))
    RightArcOppUp(i) = Sqr((RightHyp2 * RightHyp2) - (RightAdjUp2 * RightAdjUp2))
    RightArcOppDown(i) = Sqr((RightHyp2 * RightHyp2) - (RightAdjDown2 * RightAdjDown2))
    RightAdjUp2 = RightAdjUp2 + RightArcAcrossTop
    RightAdjDown2 = RightAdjDown2 + RightArcAcrossBot
    RightHyp2 = RightHyp2 + TopArcUpwards
    j = j - 1
Next i

End Sub

Private Sub PR_LoadTopArcArrays()

' This Procedure Loads the Arrays that hold the positions
' of the start and end points for the radials of the Top Arc

Dim TopArcAcross#, TopArcUpwards#
Dim TopAdj#, TopHyp#, i%

TopArcAcross = 0.23125 / 2.54
TopArcUpwards = 0.50625 / 2.54

' Top Arc Values
TopAdj = 5.05 / 2.54
TopHyp = 8 / 2.54

' Load Top Arc Arrays
For i = 1 To 19
    TopArcAdj(i) = TopAdj
    TopArcHyp(i) = TopHyp
    TopArcOpp(i) = Sqr((TopHyp * TopHyp) - (TopAdj * TopAdj))
    TopAdj = TopAdj + TopArcAcross
    TopHyp = TopHyp + TopArcUpwards
Next i

End Sub

Private Sub PR_LoadTopRightArrays()

' This Procedure Loads the Arrays that hold the positions
' of the start and end points for the radials of the Top Right Arc
 
Dim TopRightAdjIn#, TopRightAdjOut#
Dim TopRightHyp#, TopArcUpwards#
Dim TopRightArcAcrossIn#, TopRightArcUpwardsIn#
Dim TopRightArcAcrossOut#, TopRightArcUpwardsOut#
Dim i%

TopRightArcAcrossIn = 0.453125 / 2.54
TopRightArcAcrossOut = 0.4515625 / 2.54
TopRightArcUpwardsIn = 0.23125 / 2.54
TopRightArcUpwardsOut = 0.23125 / 2.54
TopArcUpwards = 0.50625 / 2.54

' Top Right Arc Values
TopRightAdjIn = 6.25 / 2.54
TopRightAdjOut = 6.955 / 2.54
TopRightHyp = 8.03 / 2.54

' Load Arrays
For i = 1 To 19
    ' Top Right Arc Arrays
    TopRightArcAdjIn(i) = TopRightAdjIn
    TopRightArcAdjOut(i) = TopRightAdjOut
    TopRightArcOppIn(i) = Sqr((TopRightHyp * TopRightHyp) - (TopRightAdjIn * TopRightAdjIn))
    TopRightArcOppOut(i) = Sqr((TopRightHyp * TopRightHyp) - (TopRightAdjOut * TopRightAdjOut))
    TopRightAdjIn = TopRightAdjIn + TopRightArcAcrossIn
    TopRightAdjOut = TopRightAdjOut + TopRightArcAcrossOut
    TopRightHyp = TopRightHyp + TopArcUpwards
Next i

End Sub

Private Sub PR_MakeXY(xyReturn As xy, X#, Y#)
   
   'Utility to return a point based on the X and Y values given
   
    xyReturn.X = X
    xyReturn.Y = Y

End Sub

Private Sub PR_NeckElastic()

Dim xyNeckElastic As xy

' Stamp For Lining
If chkNeckElastic.Value = 1 Then
    PR_MakeXY xyNeckElastic, txtChinLeftBotX + 0.4, txtChinLeftBotY + 0.5
    PR_SetLayer "Notes"
    PR_DrawText "1\"" Elastic Sewn Underneath", xyNeckElastic, 0.1, 0
    PR_SetLayer "TemplateLeft"
End If

End Sub

Private Sub PR_OpenFaceMask()

Dim xyTextPoint As xy

If optOpenFaceMask.Value = True Then
    PR_DrawTopArc
    PR_DrawRightCutOut
    PR_DrawTopRightArc
    PR_DrawRightJoin
    PR_DrawRightArc
    PR_DrawRightLine
    PR_DrawLeftArc
    PR_DrawLeftCutOut
    PR_SetChart1Details
    PR_DrawEyeOpening
    PR_DrawEars
    PR_DrawNeckAndDart
    PR_DrawForeHead
    PR_DrawChin
    PR_DrawOpenHead
    PR_DrawEarFlaps
    PR_DrawEyeFlaps
    PR_StampLining
    PR_NeckElastic
    If chkEyes.Value = 0 Then
        PR_DrawOpenFace
    ElseIf chkEyes.Value = 1 Then
        PR_DrawOpenFaceWithEyes
    End If
    PR_DrawLipCovering
    PR_MakeXY xyTextPoint, 0.3, 0.5
    PR_TemplateDetails xyTextPoint, "Open Face Mask"

End If

End Sub

Private Sub PR_PutLine(sLine As String)
   'Puts the contents of sLine to the opened "Macro" file
   'Puts the line with no translation or additions
   '    fNum is global variable
   '
    Print #txtfNum, "//"; sLine

End Sub

Private Sub PR_SaveDetails()

    Dim sSymbol As String
    
    sSymbol = "HEADNECK"

    If txtUidHN.Text = "" Then
       'Find "mainpatientdetails" and get position
        Print #txtfNum, "XY     xyMPD_Origin, xyMPD_Scale ;"
        Print #txtfNum, "STRING sMPD_Name;"
        Print #txtfNum, "ANGLE  aMPD_Angle;"

        Print #txtfNum, "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");"
        Print #txtfNum, "if (hMPD)"
        Print #txtfNum, "  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);"
        Print #txtfNum, "else"
        Print #txtfNum, "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");"

       'Insert headneck
        Print #txtfNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){"
        Print #txtfNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));"
        Print #txtfNum, "  hFacemask = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);"
        Print #txtfNum, "  }"
        Print #txtfNum, "else"
        Print #txtfNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");"
    Else
       'Use existing symbol
        Print #txtfNum, "hFacemask = UID (" & QQ & "find" & QC & Val(txtUidHN.Text) & ");"
        Print #txtfNum, "if (!hFacemask) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

    End If

    
    ' Update data base fields
    Print #txtfNum, "SetDBData(hFacemask,"; QQ; "HeadNeck"; QCQ; txtMeasurements; QQ; ");"
    Print #txtfNum, "SetDBData(hFacemask,"; QQ; "Fabric"; QCQ; cboFabric.Text; QQ; ");"
    Print #txtfNum, "SetDBData(hFacemask,"; QQ; "WorkOrder"; QCQ; txtWorkOrder; QQ; ");"
    Print #txtfNum, "SetDBData(hFacemask,"; QQ; "Data"; QCQ; txtData; QQ; ");"

End Sub

Private Sub PR_SelectTemplateRadius()

' This Procedure determines the radius
' to use on the template

Dim nTotal#, nDecVals#
Dim nEyeBrow#, nChinAngle#
Dim nEyeRem#, nChinRem#, nRemTotal#
Dim nEyeInt%, nChinInt%, nIntTotal%

'Check Values have been entered
If Val(txtCircEyeBrow) <= 0 Then
    MsgBox "Invalid or No value has been entered for CIRC above Eyebrow.", 48, "Head & Neck"
    txtCircEyeBrow.SetFocus
    Exit Sub
End If
If Val(txtCircChinAngle) <= 0 And optHeadBand.Value = False Then
    MsgBox "Invalid or No value has been entered for CIRC around head at Chin Angle.", 48, "Head & Neck"
    txtCircChinAngle.SetFocus
    Exit Sub
End If

' Convert to Inches
nEyeBrow = FN_CmToInches(txtUnits, txtCircEyeBrow)
If optHeadBand.Value = True Then
    nChinAngle = nEyeBrow
Else
    nChinAngle = FN_CmToInches(txtUnits, txtCircChinAngle)
End If

' Add Values and Round up
nEyeInt = Int(nEyeBrow)
nEyeRem = nEyeBrow - nEyeInt
nChinInt = Int(nChinInt)
nChinRem = nChinAngle - nChinInt
nIntTotal = nEyeInt + nChinInt
nRemTotal = nEyeRem + nChinRem
nDecVals = nRemTotal / 0.125
If nDecVals >= 8 Then
    nIntTotal = nIntTotal + 1
    nDecVals = nDecVals - 8
End If
nDecVals = nDecVals * 0.125
nTotal = nIntTotal + nDecVals

' Determine Radius No.
If (nTotal <= 30) Then
        txtRadiusNo = 1
        txtCircumferenceTotal = 30
ElseIf (nTotal > 30) And (nTotal <= 32) Then
        txtRadiusNo = 2
        txtCircumferenceTotal = 32
ElseIf (nTotal > 32) And (nTotal <= 34) Then
        txtRadiusNo = 3
        txtCircumferenceTotal = 34
ElseIf (nTotal > 34) And (nTotal <= 36) Then
        txtRadiusNo = 4
        txtCircumferenceTotal = 36
ElseIf (nTotal > 36) And (nTotal <= 38) Then
        txtRadiusNo = 5
        txtCircumferenceTotal = 38
ElseIf (nTotal > 38) And (nTotal <= 40) Then
        txtRadiusNo = 6
        txtCircumferenceTotal = 40
ElseIf (nTotal > 40) And (nTotal <= 42) Then
        txtRadiusNo = 7
        txtCircumferenceTotal = 42
ElseIf (nTotal > 42) And (nTotal <= 44) Then
        txtRadiusNo = 8
        txtCircumferenceTotal = 44
ElseIf (nTotal > 44) And (nTotal <= 46) Then
        txtRadiusNo = 9
        txtCircumferenceTotal = 46
ElseIf (nTotal > 46) And (nTotal <= 48) Then
        txtRadiusNo = 10
        txtCircumferenceTotal = 48
ElseIf (nTotal > 48) And (nTotal <= 50) Then
        txtRadiusNo = 11
        txtCircumferenceTotal = 50
ElseIf (nTotal > 50) And (nTotal <= 52) Then
        txtRadiusNo = 12
        txtCircumferenceTotal = 52
ElseIf (nTotal > 52) And (nTotal <= 54) Then
        txtRadiusNo = 13
        txtCircumferenceTotal = 54
ElseIf (nTotal > 54) And (nTotal <= 56) Then
        txtRadiusNo = 14
        txtCircumferenceTotal = 56
ElseIf (nTotal > 56) And (nTotal <= 58) Then
        txtRadiusNo = 15
        txtCircumferenceTotal = 58
ElseIf (nTotal > 58) And (nTotal <= 60) Then
        txtRadiusNo = 16
        txtCircumferenceTotal = 60
ElseIf (nTotal > 60) And (nTotal <= 62) Then
        txtRadiusNo = 17
        txtCircumferenceTotal = 62
End If

End Sub

Private Sub PR_SelectTextInBox(Text_Box_Name As Control)
    
    Text_Box_Name.SelStart = 0
    Text_Box_Name.SelLength = 4

End Sub

Private Sub PR_SeperateMeasurements()

Dim tempval As Variant
tempval = Left$(txtMeasurements, 2) & "." & Mid$(txtMeasurements, 3, 1)
If Val(tempval) > 0 Then
    txtChinToMouth = Val(tempval)
    PR_ShowInches tempval, lblInch1
End If
tempval = Mid$(txtMeasurements, 4, 2) & "." & Mid$(txtMeasurements, 6, 1)
If Val(tempval) > 0 Then
    txtCircEyeBrow = Val(tempval)
    PR_ShowInches tempval, lblInch2
End If
tempval = Mid$(txtMeasurements, 7, 2) & "." & Mid$(txtMeasurements, 9, 1)
If Val(tempval) > 0 Then
    txtCircChinAngle = Val(tempval)
    PR_ShowInches tempval, lblInch3
End If
tempval = Mid$(txtMeasurements, 10, 2) & "." & Mid$(txtMeasurements, 12, 1)
If Val(tempval) > 0 Then
    txtCircOfNeck = Val(tempval)
    PR_ShowInches tempval, lblInch4
End If
tempval = Mid$(txtMeasurements, 13, 2) & "." & Mid$(txtMeasurements, 15, 1)
If Val(tempval) > 0 Then
    txtThroatToSternal = Val(tempval)
    PR_ShowInches tempval, lblInch5
End If
tempval = Mid$(txtMeasurements, 16, 2) & "." & Mid$(txtMeasurements, 18, 1)
If Val(tempval) > 0 Then
    txtTipOfNose = Val(tempval)
    PR_ShowInches tempval, lblInch6
End If
tempval = Mid$(txtMeasurements, 19, 2) & "." & Mid$(txtMeasurements, 21, 1)
If Val(tempval) > 0 Then
    txtLengthOfNose = Val(tempval)
    PR_ShowInches tempval, lblInch7
End If
tempval = Mid$(txtMeasurements, 22, 2) & "." & Mid$(txtMeasurements, 24, 1)
If Val(tempval) > 0 Then
    txtHeadBandDepth = Val(tempval)
    PR_ShowInches tempval, lblInch8
End If
tempval = Mid$(txtMeasurements, 25, 2) & "." & Mid$(txtMeasurements, 27, 1)
If Val(tempval) > 0 Then
    txtLeftEarLength = Val(tempval)
    PR_ShowInches tempval, lblInch9
End If
tempval = Mid$(txtMeasurements, 28, 2) & "." & Mid$(txtMeasurements, 30, 1)
If Val(tempval) > 0 Then
    txtRightEarLength = Val(tempval)
    PR_ShowInches tempval, lblInch10
End If
tempval = Mid$(txtMeasurements, 31, 2) & "." & Mid$(txtMeasurements, 33, 1)
If Val(tempval) > 0 Then
    txtChinCollarMin = Val(tempval)
    PR_ShowInches tempval, lblInch11
End If

End Sub

Private Sub PR_SeperateModifications()

Dim tempval As Variant

tempval = Left$(txtData, 3)
If tempval = "RFM" Then
    optFaceMask.Value = True
ElseIf tempval = "RHB" Then
    optHeadBand.Value = True
ElseIf tempval = "OFM" Then
    optOpenFaceMask.Value = True
ElseIf tempval = "RCS" Then
    optChinStrap.Value = True
ElseIf tempval = "MCS" Then
    optModifiedChinStrap.Value = True
ElseIf tempval = "RCC" Then
    optChinCollar.Value = True
ElseIf tempval = "CCC" Then
    optContouredChinCollar.Value = True
End If

tempval = Mid$(txtData, 4, 1)
If Val(tempval) = 1 Then
    chkLeftEarFlap.Value = 1
End If
tempval = Mid$(txtData, 5, 1)
If Val(tempval) = 1 Then
    chkRightEarFlap.Value = 1
End If
tempval = Mid$(txtData, 6, 1)
If Val(tempval) = 1 Then
    chkNeckElastic.Value = 1
End If
tempval = Mid$(txtData, 7, 1)
If Val(tempval) = 1 Then
    chkLeftEyeFlap.Value = 1
End If
tempval = Mid$(txtData, 8, 1)
If Val(tempval) = 1 Then
    chkRightEyeFlap.Value = 1
End If
tempval = Mid$(txtData, 9, 1)
If Val(tempval) = 1 Then
    chkOpenHeadMask.Value = 1
End If
tempval = Mid$(txtData, 10, 1)
If Val(tempval) = 1 Then
    chkLipStrap.Value = 1
End If
tempval = Mid$(txtData, 11, 1)
If Val(tempval) = 1 Then
    chkLipCovering.Value = 1
End If
tempval = Mid$(txtData, 12, 1)
If Val(tempval) = 1 Then
    chkNoseCovering.Value = 1
End If
tempval = Mid$(txtData, 13, 1)
If Val(tempval) = 1 Then
    chkZipper.Value = 1
End If
tempval = Mid$(txtData, 14, 1)
If Val(tempval) = 1 Then
    chkLining.Value = 1
End If
tempval = Mid$(txtData, 15, 1)
If Val(tempval) = 1 Then
    chkEarSize.Value = 1
End If
tempval = Mid$(txtData, 16, 1)
If Val(tempval) = 1 Then
    chkRightEarClosed.Value = 1
End If
tempval = Mid$(txtData, 17, 1)
If Val(tempval) = 1 Then
    chkLeftEarClosed.Value = 1
End If
tempval = Mid$(txtData, 18, 1)
If Val(tempval) = 1 Then
    chkEyes.Value = 1
End If
tempval = Mid$(txtData, 19, 1)
If Val(tempval) = 1 Then
    chkVelcro.Value = 1
End If

End Sub

Private Sub PR_SetChart1Details()

' This Procedure Follows the Jobst Procedures No. 4
' It Gets the Mouth Height from Measurements,
' and checks chart 1 to get width of Lip strap and
' eye Height.
' These are then marked on the drawing

Dim xyMouthStart As xy, xyMouthEnd As xy
Dim xyEyeStart As xy, xyEyeEnd As xy
Dim PointOne#, PointTwo#
Dim Chart1Ref#, MouthHt#
Dim LipWidth#, EyeHt#
Dim nArrayVal#, nChart1Ref#
Dim nArrayValB#, i%

' Get start point for Mouth Height
PointOne = BotFaceAdj(txtRadiusNo)
PointTwo = BotFaceOpp(txtRadiusNo)
MouthHt = FN_CmToInches(txtUnits, txtChinToMouth)
PR_MakeXY xyMouthStart, PointOne, PointTwo + MouthHt

' set txtMouthHeight
txtMouthHeight = xyMouthStart.Y

' Check Chart1
nChart1Ref = Abs(PointTwo) - Abs(MouthHt)
If nChart1Ref = (FaceMaskChartA(1, 25)) Then
    LipWidth = (FaceMaskChartA(2, 25))
    EyeHt = (FaceMaskChartA(3, 25))
Else
    For i = 1 To 24
        nArrayVal = (FaceMaskChartA(1, i))
        nArrayValB = (FaceMaskChartA(1, i + 1))
        If (nChart1Ref >= nArrayVal) And (nChart1Ref < nArrayValB) Then
            LipWidth = (FaceMaskChartA(2, i))
            EyeHt = (FaceMaskChartA(3, i))
        End If
    Next i
End If

' Set Details
txtLipStrapWidth = LipWidth
txtMidToEyeTop = EyeHt

' Mark Mouth Height
PR_MakeXY xyMouthEnd, PointOne, PointTwo + MouthHt + LipWidth

If optFaceMask.Value = True Then
    PR_DrawLine xyMouthStart, xyMouthEnd
    PR_AddEntityID "Lip"
End If
' set txtNoseBottomY
txtNoseBottomY = PointTwo + MouthHt + LipWidth

' Mark Eye Height
PR_MakeXY xyEyeStart, PointOne, 0
If chkNoseCovering.Value = 1 Then
    PR_MakeXY xyEyeEnd, PointOne, -EyeHt + 0.25
    txtNoseCoverX = xyEyeEnd.X
    txtNoseCoverY = xyEyeEnd.Y
Else
    PR_MakeXY xyEyeEnd, PointOne, -EyeHt
End If
If optFaceMask.Value = True Then
    PR_DrawLine xyEyeStart, xyEyeEnd
    PR_AddEntityID "EyeBrow"
End If

End Sub

Private Sub PR_SetLayer(sNewLayer$)
   
   'To the DRAFIX macro file (given by the global txtfNum).
   'Write the syntax to set the current LAYER.
   'For this to work it assumes that hLayer is defined in DRAFIX as
   'a HANDLE.
   '
   'To reduce unessesary writing of DRAFIX code check that the new layer
   'is different from the Current layer, change only if it is different.
   '
    Print #txtfNum, "hLayer = Table("; QQ; "find"; QCQ; "layer"; QCQ; sNewLayer; QQ; ");"
    Print #txtfNum, "if ( hLayer > %zero && hLayer != 32768)"; "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "hLayer);"

End Sub

Private Sub PR_ShowInches(TextBox As Variant, InchLabel As Variant)

' Show Values as Inches
If Val(TextBox) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, TextBox)
    InchLabel = FN_InchesToText(inchbit)
End If

End Sub

Private Sub PR_StampLining()

Dim xyStampLining As xy

' Stamp For Lining
If chkLining.Value = 1 And chkOpenHeadMask.Value <> 1 Then
    PR_MakeXY xyStampLining, -0.5, TopArcOpp(txtRadiusNo) - 0.5
    PR_SetLayer "Notes"
    PR_DrawText "Lining", xyStampLining, 0.25, 0
    PR_SetLayer "TemplateLeft"
End If

End Sub

Private Sub PR_TemplateDetails(xyTextPoint As xy, TitleText$)
    Dim xyNotes As xy
    Dim sText As String

    ' Add notes
    PR_SetLayer "Notes"
    sText = TitleText & "\n" & txtPatientName & "\n" & txtWorkOrder & "\n" & cboFabric.Text
    PR_MakeXY xyNotes, xyTextPoint.X, xyTextPoint.Y - 0.125
    PR_DrawText sText, xyNotes, 0.1, 0
    
    PR_SetLayer "Construct"
    sText = txtFileNo & "\n" & txtDiagnosis & "\n" & txtAge & "\n" & txtSex
    PR_MakeXY xyNotes, xyTextPoint.X, xyTextPoint.Y - 0.9
    PR_DrawText sText, xyNotes, 0.1, 0

    
    PR_SetLayer "TemplateLeft"

End Sub

Private Function round(ByVal nNumber As Double) As Integer
   'Fuction to return the rounded value of a decimal number
   'E.G.
   '    round(1.35)  = 1
   '    round(1.55)  = 2
   '    round(2.50)  = 3
   '    round(-2.50) = -3
   '    round(0)     = 0
   '

    Dim iInt As Integer, iSign As Integer
    
   'Avoid extra work. Return 0 if input is 0
    If nNumber = 0 Then
        round = 0
        Exit Function
    End If

   'Split input
    iSign = Sgn(nNumber)
    nNumber = Abs(nNumber)
    iInt = Int(nNumber)

   'Effect rounding
    If (nNumber - iInt) >= 0.5 Then
        round = (iInt + 1) * iSign
    Else
        round = iInt * iSign
    End If
    
End Function

Private Sub Timer1_Timer()
    'This is enabled on Load
    'It is disabled in Link close
    'If the event happens then it is assumed that
    'for some reason the link between DRAFIX and This dialogue
    'has failed'
    'Therefor we "End" here
     End
End Sub

Private Sub txtChinCollarMin_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtChinCollarMin

End Sub

Private Sub txtChinCollarMin_LostFocus()

    'Check for Numeric Value
    If Not IsNumeric(txtChinCollarMin) And txtChinCollarMin <> "" Then
        MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
        txtChinCollarMin.SetFocus
    End If

    ' Show Values as Inches
    If Val(txtChinCollarMin) > 0 Then
        Dim inchbit#
        inchbit = FN_CmToInches(txtUnits, txtChinCollarMin)
        lblInch11 = FN_InchesToText(inchbit)
    Else
        lblInch11 = ""
    End If

    ' Ensure Value is not Greater then The Throat to Sternal Notch value
    If txtThroatToSternal <> "" Then
        If Val(txtChinCollarMin) >= Val(txtThroatToSternal) Then
            MsgBox "Chin Collar Min Cannot be Greater than the Throat to Sternal Notch Distance", 48, "Head & Neck"
            txtChinCollarMin.SetFocus
        End If
    End If

End Sub

Private Sub txtChinToMouth_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtChinToMouth

End Sub

Private Sub txtChinToMouth_LostFocus()
    
    'Check for Numeric Value
    If Not IsNumeric(txtChinToMouth) And txtChinToMouth <> "" Then
        MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
        txtChinToMouth.SetFocus
    End If

    ' Show Values as Inches
    If Val(txtChinToMouth) > 0 Then
        Dim inchbit#
        inchbit = FN_CmToInches(txtUnits, txtChinToMouth)
        lblInch1 = FN_InchesToText(inchbit)
    Else
        lblInch1 = ""
    End If

End Sub

Private Sub txtCircChinAngle_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtCircChinAngle

End Sub

Private Sub txtCircChinAngle_LostFocus()
    
    'Check for Numeric Value
    If Not IsNumeric(txtCircChinAngle) And txtCircChinAngle <> "" Then
        MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
        txtCircChinAngle.SetFocus
    End If

    ' Show Values as Inches
    If Val(txtCircChinAngle) > 0 Then
        Dim inchbit#
        inchbit = FN_CmToInches(txtUnits, txtCircChinAngle)
        lblInch3 = FN_InchesToText(inchbit)
    Else
        lblInch3 = ""
    End If

End Sub

Private Sub txtCircEyeBrow_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtCircEyeBrow

End Sub

Private Sub txtCircEyeBrow_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtCircEyeBrow) And txtCircEyeBrow <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtCircEyeBrow.SetFocus
End If

' Show Values as Inches
If Val(txtCircEyeBrow) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtCircEyeBrow)
    lblInch2 = FN_InchesToText(inchbit)
Else
    lblInch2 = ""
End If

End Sub

Private Sub txtCircOfNeck_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtCircOfNeck

End Sub

Private Sub txtCircOfNeck_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtCircOfNeck) And txtCircOfNeck <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtCircOfNeck.SetFocus
End If

' Show Values as Inches
If Val(txtCircOfNeck) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtCircOfNeck)
    lblInch4 = FN_InchesToText(inchbit)
Else
    lblInch4 = ""
End If

End Sub

Private Sub txtHeadBandDepth_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtHeadBandDepth

End Sub

Private Sub txtHeadBandDepth_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtHeadBandDepth) And txtHeadBandDepth <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtHeadBandDepth.SetFocus
End If

' Show Values as Inches
If Val(txtHeadBandDepth) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtHeadBandDepth)
    lblInch8 = FN_InchesToText(inchbit)
Else
    lblInch8 = ""
End If

 
End Sub

Private Sub txtLeftEarLength_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtLeftEarLength

End Sub

Private Sub txtLeftEarLength_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtLeftEarLength) And txtLeftEarLength <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtLeftEarLength.SetFocus
End If

' Show Values as Inches
If Val(txtLeftEarLength) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtLeftEarLength)
    lblInch9 = FN_InchesToText(inchbit)
Else
    lblInch9 = ""
End If


End Sub

Private Sub txtLengthOfNose_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtLengthOfNose

End Sub

Private Sub txtLengthOfNose_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtLengthOfNose) And txtLengthOfNose <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtLengthOfNose.SetFocus
End If

' Show Values as Inches
If Val(txtLengthOfNose) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtLengthOfNose)
    lblInch7 = FN_InchesToText(inchbit)
Else
    lblInch7 = ""
End If

End Sub

Private Sub txtRightEarLength_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtRightEarLength

End Sub

Private Sub txtRightEarLength_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtRightEarLength) And txtRightEarLength <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtRightEarLength.SetFocus
End If

' Show Values as Inches
If Val(txtRightEarLength) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtRightEarLength)
    lblInch10 = FN_InchesToText(inchbit)
Else
    lblInch10 = ""
End If


End Sub

Private Sub txtThroatToSternal_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtThroatToSternal

End Sub

Private Sub txtThroatToSternal_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtThroatToSternal) And txtThroatToSternal <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtThroatToSternal.SetFocus
End If
Dim totval
totval = Val(txtThroatToSternal)

' Show Values as Inches
If Val(txtThroatToSternal) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtThroatToSternal)
    lblInch5 = FN_InchesToText(inchbit)
Else
    lblInch5 = ""
End If

' Ensure Chin Collar Min is not Greater then The Throat to Sternal Notch value
If txtThroatToSternal <> "" Then
    If Val(txtChinCollarMin) >= Val(txtThroatToSternal) Then
        MsgBox "Chin Collar Min Cannot be Greater than the Throat to Sternal Notch Distance", 48, "Head & Neck"
        txtChinCollarMin.SetFocus
    End If
End If

End Sub

Private Sub txtTipOfNose_GotFocus()
    
    ' Highlight text in box
    PR_SelectTextInBox txtTipOfNose

End Sub

Private Sub txtTipOfNose_LostFocus()

'Check for Numeric Value
If Not IsNumeric(txtTipOfNose) And txtTipOfNose <> "" Then
    MsgBox "Non-Numeric value  has been entered", 48, "Head & Neck"
    txtTipOfNose.SetFocus
End If

' Show Values as Inches
If Val(txtTipOfNose) > 0 Then
    Dim inchbit#
    inchbit = FN_CmToInches(txtUnits, txtTipOfNose)
    lblInch6 = FN_InchesToText(inchbit)
Else
    lblInch6 = ""
End If

End Sub


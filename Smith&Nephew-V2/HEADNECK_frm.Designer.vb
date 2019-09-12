<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class HeadNeck
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents txtUidHN As System.Windows.Forms.TextBox
	Public WithEvents txtTopRightEar As System.Windows.Forms.TextBox
	Public WithEvents txtTopLeftEar As System.Windows.Forms.TextBox
	Public WithEvents txtHeadNeckUID As System.Windows.Forms.TextBox
	Public WithEvents txtOptionChoice As System.Windows.Forms.TextBox
	Public WithEvents txtFabric As System.Windows.Forms.TextBox
	Public WithEvents txtWorkOrder As System.Windows.Forms.TextBox
	Public WithEvents txtData As System.Windows.Forms.TextBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents txtMeasurements As System.Windows.Forms.TextBox
	Public WithEvents txtMouthRightY As System.Windows.Forms.TextBox
	Public WithEvents txtxyNeckTopFrontY As System.Windows.Forms.TextBox
	Public WithEvents txtxyNeckTopFrontX As System.Windows.Forms.TextBox
	Public WithEvents txtxyChinTopY As System.Windows.Forms.TextBox
	Public WithEvents txtxyChinTopX As System.Windows.Forms.TextBox
	Public WithEvents txtCSEarBotHeight As System.Windows.Forms.TextBox
	Public WithEvents txtCSEarTopHeight As System.Windows.Forms.TextBox
	Public WithEvents txtCSMouthWidth As System.Windows.Forms.TextBox
	Public WithEvents txtMouthRightX As System.Windows.Forms.TextBox
	Public WithEvents txtCSForeHead As System.Windows.Forms.TextBox
	Public WithEvents txtCSChinAngle As System.Windows.Forms.TextBox
	Public WithEvents txtCSNeckCircum As System.Windows.Forms.TextBox
	Public WithEvents txtCSChintoMouth As System.Windows.Forms.TextBox
	Public WithEvents txtNoseCoverX As System.Windows.Forms.TextBox
	Public WithEvents txtNoseCoverY As System.Windows.Forms.TextBox
	Public WithEvents txtEyeWidth As System.Windows.Forms.TextBox
	Public WithEvents txtDartEndX As System.Windows.Forms.TextBox
	Public WithEvents txtDartEndY As System.Windows.Forms.TextBox
	Public WithEvents txtDartStartY As System.Windows.Forms.TextBox
	Public WithEvents txtDartStartX As System.Windows.Forms.TextBox
	Public WithEvents txtChinLeftBotY As System.Windows.Forms.TextBox
	Public WithEvents txtChinLeftBotX As System.Windows.Forms.TextBox
	Public WithEvents txtNoseBottomY As System.Windows.Forms.TextBox
	Public WithEvents txtCircumferenceTotal As System.Windows.Forms.TextBox
	Public WithEvents txtMouthHeight As System.Windows.Forms.TextBox
	Public WithEvents txtBotOfEyeY As System.Windows.Forms.TextBox
	Public WithEvents txtBotOfEyeX As System.Windows.Forms.TextBox
	Public WithEvents txtLowerEarHeight As System.Windows.Forms.TextBox
	Public WithEvents txtfNum As System.Windows.Forms.TextBox
	Public WithEvents txtCurrTextFont As System.Windows.Forms.TextBox
	Public WithEvents txtCurrTextVertJust As System.Windows.Forms.TextBox
	Public WithEvents txtCurrTextHorizJust As System.Windows.Forms.TextBox
	Public WithEvents txtCurrTextAspect As System.Windows.Forms.TextBox
	Public WithEvents txtCurrTextHt As System.Windows.Forms.TextBox
	Public WithEvents txtCurrentLayer As System.Windows.Forms.TextBox
	Public WithEvents txtMidToEyeTop As System.Windows.Forms.TextBox
	Public WithEvents txtLipStrapWidth As System.Windows.Forms.TextBox
	Public WithEvents txtUidMPD As System.Windows.Forms.TextBox
	Public WithEvents txtRadiusNo As System.Windows.Forms.TextBox
	Public WithEvents txtDraw As System.Windows.Forms.TextBox
	Public WithEvents cmdDraw As System.Windows.Forms.Button
	Public WithEvents txtChinCollarMin As System.Windows.Forms.TextBox
	Public WithEvents txtRightEarLength As System.Windows.Forms.TextBox
	Public WithEvents txtLeftEarLength As System.Windows.Forms.TextBox
	Public WithEvents txtHeadBandDepth As System.Windows.Forms.TextBox
	Public WithEvents txtLengthOfNose As System.Windows.Forms.TextBox
	Public WithEvents txtTipOfNose As System.Windows.Forms.TextBox
	Public WithEvents cboFabric As System.Windows.Forms.ComboBox
	Public WithEvents txtThroatToSternal As System.Windows.Forms.TextBox
	Public WithEvents txtCircOfNeck As System.Windows.Forms.TextBox
	Public WithEvents txtCircChinAngle As System.Windows.Forms.TextBox
	Public WithEvents txtCircEyeBrow As System.Windows.Forms.TextBox
	Public WithEvents txtChinToMouth As System.Windows.Forms.TextBox
	Public WithEvents lblInch11 As System.Windows.Forms.Label
	Public WithEvents lblChinCollarMin As System.Windows.Forms.Label
	Public WithEvents lblRightEarLen As System.Windows.Forms.Label
	Public WithEvents lblLeftEarLen As System.Windows.Forms.Label
	Public WithEvents lblInch10 As System.Windows.Forms.Label
	Public WithEvents lblInch9 As System.Windows.Forms.Label
	Public WithEvents lblInch8 As System.Windows.Forms.Label
	Public WithEvents lblHeadBandDepth As System.Windows.Forms.Label
	Public WithEvents lblFabric As System.Windows.Forms.Label
	Public WithEvents lblInch7 As System.Windows.Forms.Label
	Public WithEvents lblInch6 As System.Windows.Forms.Label
	Public WithEvents LblLengthOfNose As System.Windows.Forms.Label
	Public WithEvents lblAcrossTipOfNose As System.Windows.Forms.Label
	Public WithEvents lblInch5 As System.Windows.Forms.Label
	Public WithEvents lblInch4 As System.Windows.Forms.Label
	Public WithEvents lblInch3 As System.Windows.Forms.Label
	Public WithEvents lblInch2 As System.Windows.Forms.Label
	Public WithEvents lblInch1 As System.Windows.Forms.Label
	Public WithEvents lblThroatToSternal As System.Windows.Forms.Label
	Public WithEvents lblCircOfNeck As System.Windows.Forms.Label
	Public WithEvents lblCircChinAngle As System.Windows.Forms.Label
	Public WithEvents lblCircEyeBrow As System.Windows.Forms.Label
	Public WithEvents lblChinToMouth As System.Windows.Forms.Label
	Public WithEvents frmMeasurements As System.Windows.Forms.GroupBox
	Public WithEvents chkVelcro As System.Windows.Forms.CheckBox
	Public WithEvents chkEyes As System.Windows.Forms.CheckBox
	Public WithEvents chkRightEarClosed As System.Windows.Forms.CheckBox
	Public WithEvents chkLeftEarClosed As System.Windows.Forms.CheckBox
	Public WithEvents chkEarSize As System.Windows.Forms.CheckBox
	Public WithEvents chkLipStrap As System.Windows.Forms.CheckBox
	Public WithEvents chkOpenHeadMask As System.Windows.Forms.CheckBox
	Public WithEvents chkNeckElastic As System.Windows.Forms.CheckBox
	Public WithEvents chkLining As System.Windows.Forms.CheckBox
	Public WithEvents chkLeftEyeFlap As System.Windows.Forms.CheckBox
	Public WithEvents chkRightEyeFlap As System.Windows.Forms.CheckBox
	Public WithEvents chkZipper As System.Windows.Forms.CheckBox
	Public WithEvents chkNoseCovering As System.Windows.Forms.CheckBox
	Public WithEvents chkLipCovering As System.Windows.Forms.CheckBox
	Public WithEvents chkRightEarFlap As System.Windows.Forms.CheckBox
	Public WithEvents chkLeftEarFlap As System.Windows.Forms.CheckBox
	Public WithEvents frmModifications As System.Windows.Forms.GroupBox
	Public WithEvents optContouredChinCollar As System.Windows.Forms.RadioButton
	Public WithEvents optChinCollar As System.Windows.Forms.RadioButton
	Public WithEvents optModifiedChinStrap As System.Windows.Forms.RadioButton
	Public WithEvents optChinStrap As System.Windows.Forms.RadioButton
	Public WithEvents optHeadBand As System.Windows.Forms.RadioButton
	Public WithEvents optOpenFaceMask As System.Windows.Forms.RadioButton
	Public WithEvents optFaceMask As System.Windows.Forms.RadioButton
	Public WithEvents frmDesignChoices As System.Windows.Forms.GroupBox
	Public WithEvents txtSex As System.Windows.Forms.TextBox
	Public WithEvents txtUnits As System.Windows.Forms.TextBox
	Public WithEvents txtAge As System.Windows.Forms.TextBox
	Public WithEvents txtFileNo As System.Windows.Forms.TextBox
	Public WithEvents txtDiagnosis As System.Windows.Forms.TextBox
	Public WithEvents txtPatientName As System.Windows.Forms.TextBox
	Public WithEvents lblSex As System.Windows.Forms.Label
	Public WithEvents lblUnits As System.Windows.Forms.Label
	Public WithEvents lblAge As System.Windows.Forms.Label
	Public WithEvents lblFileNo As System.Windows.Forms.Label
	Public WithEvents lblDiagnosis As System.Windows.Forms.Label
	Public WithEvents lblPatientName As System.Windows.Forms.Label
	Public WithEvents frmPatientDetails As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtUidHN = New System.Windows.Forms.TextBox()
        Me.txtTopRightEar = New System.Windows.Forms.TextBox()
        Me.txtTopLeftEar = New System.Windows.Forms.TextBox()
        Me.txtHeadNeckUID = New System.Windows.Forms.TextBox()
        Me.txtOptionChoice = New System.Windows.Forms.TextBox()
        Me.txtFabric = New System.Windows.Forms.TextBox()
        Me.txtWorkOrder = New System.Windows.Forms.TextBox()
        Me.txtData = New System.Windows.Forms.TextBox()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.txtMeasurements = New System.Windows.Forms.TextBox()
        Me.txtMouthRightY = New System.Windows.Forms.TextBox()
        Me.txtxyNeckTopFrontY = New System.Windows.Forms.TextBox()
        Me.txtxyNeckTopFrontX = New System.Windows.Forms.TextBox()
        Me.txtxyChinTopY = New System.Windows.Forms.TextBox()
        Me.txtxyChinTopX = New System.Windows.Forms.TextBox()
        Me.txtCSEarBotHeight = New System.Windows.Forms.TextBox()
        Me.txtCSEarTopHeight = New System.Windows.Forms.TextBox()
        Me.txtCSMouthWidth = New System.Windows.Forms.TextBox()
        Me.txtMouthRightX = New System.Windows.Forms.TextBox()
        Me.txtCSForeHead = New System.Windows.Forms.TextBox()
        Me.txtCSChinAngle = New System.Windows.Forms.TextBox()
        Me.txtCSNeckCircum = New System.Windows.Forms.TextBox()
        Me.txtCSChintoMouth = New System.Windows.Forms.TextBox()
        Me.txtNoseCoverX = New System.Windows.Forms.TextBox()
        Me.txtNoseCoverY = New System.Windows.Forms.TextBox()
        Me.txtEyeWidth = New System.Windows.Forms.TextBox()
        Me.txtDartEndX = New System.Windows.Forms.TextBox()
        Me.txtDartEndY = New System.Windows.Forms.TextBox()
        Me.txtDartStartY = New System.Windows.Forms.TextBox()
        Me.txtDartStartX = New System.Windows.Forms.TextBox()
        Me.txtChinLeftBotY = New System.Windows.Forms.TextBox()
        Me.txtChinLeftBotX = New System.Windows.Forms.TextBox()
        Me.txtNoseBottomY = New System.Windows.Forms.TextBox()
        Me.txtCircumferenceTotal = New System.Windows.Forms.TextBox()
        Me.txtMouthHeight = New System.Windows.Forms.TextBox()
        Me.txtBotOfEyeY = New System.Windows.Forms.TextBox()
        Me.txtBotOfEyeX = New System.Windows.Forms.TextBox()
        Me.txtLowerEarHeight = New System.Windows.Forms.TextBox()
        Me.txtfNum = New System.Windows.Forms.TextBox()
        Me.txtCurrTextFont = New System.Windows.Forms.TextBox()
        Me.txtCurrTextVertJust = New System.Windows.Forms.TextBox()
        Me.txtCurrTextHorizJust = New System.Windows.Forms.TextBox()
        Me.txtCurrTextAspect = New System.Windows.Forms.TextBox()
        Me.txtCurrTextHt = New System.Windows.Forms.TextBox()
        Me.txtCurrentLayer = New System.Windows.Forms.TextBox()
        Me.txtMidToEyeTop = New System.Windows.Forms.TextBox()
        Me.txtLipStrapWidth = New System.Windows.Forms.TextBox()
        Me.txtUidMPD = New System.Windows.Forms.TextBox()
        Me.txtRadiusNo = New System.Windows.Forms.TextBox()
        Me.txtDraw = New System.Windows.Forms.TextBox()
        Me.cmdDraw = New System.Windows.Forms.Button()
        Me.frmMeasurements = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtChinCollarMin = New System.Windows.Forms.TextBox()
        Me.txtRightEarLength = New System.Windows.Forms.TextBox()
        Me.txtLeftEarLength = New System.Windows.Forms.TextBox()
        Me.txtHeadBandDepth = New System.Windows.Forms.TextBox()
        Me.txtLengthOfNose = New System.Windows.Forms.TextBox()
        Me.txtTipOfNose = New System.Windows.Forms.TextBox()
        Me.cboFabric = New System.Windows.Forms.ComboBox()
        Me.txtThroatToSternal = New System.Windows.Forms.TextBox()
        Me.txtCircOfNeck = New System.Windows.Forms.TextBox()
        Me.txtCircChinAngle = New System.Windows.Forms.TextBox()
        Me.txtCircEyeBrow = New System.Windows.Forms.TextBox()
        Me.txtChinToMouth = New System.Windows.Forms.TextBox()
        Me.lblInch11 = New System.Windows.Forms.Label()
        Me.lblChinCollarMin = New System.Windows.Forms.Label()
        Me.lblRightEarLen = New System.Windows.Forms.Label()
        Me.lblLeftEarLen = New System.Windows.Forms.Label()
        Me.lblInch10 = New System.Windows.Forms.Label()
        Me.lblInch9 = New System.Windows.Forms.Label()
        Me.lblInch8 = New System.Windows.Forms.Label()
        Me.lblHeadBandDepth = New System.Windows.Forms.Label()
        Me.lblFabric = New System.Windows.Forms.Label()
        Me.lblInch7 = New System.Windows.Forms.Label()
        Me.lblInch6 = New System.Windows.Forms.Label()
        Me.LblLengthOfNose = New System.Windows.Forms.Label()
        Me.lblAcrossTipOfNose = New System.Windows.Forms.Label()
        Me.lblInch5 = New System.Windows.Forms.Label()
        Me.lblInch4 = New System.Windows.Forms.Label()
        Me.lblInch3 = New System.Windows.Forms.Label()
        Me.lblInch2 = New System.Windows.Forms.Label()
        Me.lblInch1 = New System.Windows.Forms.Label()
        Me.lblThroatToSternal = New System.Windows.Forms.Label()
        Me.lblCircOfNeck = New System.Windows.Forms.Label()
        Me.lblCircChinAngle = New System.Windows.Forms.Label()
        Me.lblCircEyeBrow = New System.Windows.Forms.Label()
        Me.lblChinToMouth = New System.Windows.Forms.Label()
        Me.frmModifications = New System.Windows.Forms.GroupBox()
        Me.chkVelcro = New System.Windows.Forms.CheckBox()
        Me.chkEyes = New System.Windows.Forms.CheckBox()
        Me.chkRightEarClosed = New System.Windows.Forms.CheckBox()
        Me.chkLeftEarClosed = New System.Windows.Forms.CheckBox()
        Me.chkEarSize = New System.Windows.Forms.CheckBox()
        Me.chkLipStrap = New System.Windows.Forms.CheckBox()
        Me.chkOpenHeadMask = New System.Windows.Forms.CheckBox()
        Me.chkNeckElastic = New System.Windows.Forms.CheckBox()
        Me.chkLining = New System.Windows.Forms.CheckBox()
        Me.chkLeftEyeFlap = New System.Windows.Forms.CheckBox()
        Me.chkRightEyeFlap = New System.Windows.Forms.CheckBox()
        Me.chkZipper = New System.Windows.Forms.CheckBox()
        Me.chkNoseCovering = New System.Windows.Forms.CheckBox()
        Me.chkLipCovering = New System.Windows.Forms.CheckBox()
        Me.chkRightEarFlap = New System.Windows.Forms.CheckBox()
        Me.chkLeftEarFlap = New System.Windows.Forms.CheckBox()
        Me.frmDesignChoices = New System.Windows.Forms.GroupBox()
        Me.optContouredChinCollar = New System.Windows.Forms.RadioButton()
        Me.optChinCollar = New System.Windows.Forms.RadioButton()
        Me.optModifiedChinStrap = New System.Windows.Forms.RadioButton()
        Me.optChinStrap = New System.Windows.Forms.RadioButton()
        Me.optHeadBand = New System.Windows.Forms.RadioButton()
        Me.optOpenFaceMask = New System.Windows.Forms.RadioButton()
        Me.optFaceMask = New System.Windows.Forms.RadioButton()
        Me.frmPatientDetails = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.txtSex = New System.Windows.Forms.TextBox()
        Me.txtAge = New System.Windows.Forms.TextBox()
        Me.txtFileNo = New System.Windows.Forms.TextBox()
        Me.txtDiagnosis = New System.Windows.Forms.TextBox()
        Me.txtPatientName = New System.Windows.Forms.TextBox()
        Me.lblSex = New System.Windows.Forms.Label()
        Me.lblAge = New System.Windows.Forms.Label()
        Me.lblFileNo = New System.Windows.Forms.Label()
        Me.lblDiagnosis = New System.Windows.Forms.Label()
        Me.lblPatientName = New System.Windows.Forms.Label()
        Me.txtUnits = New System.Windows.Forms.TextBox()
        Me.lblUnits = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtTempDate = New System.Windows.Forms.TextBox()
        Me.txtDesigner = New System.Windows.Forms.TextBox()
        Me.txtWorkOrder1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.frmMeasurements.SuspendLayout()
        Me.frmModifications.SuspendLayout()
        Me.frmDesignChoices.SuspendLayout()
        Me.frmPatientDetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'txtUidHN
        '
        Me.txtUidHN.AcceptsReturn = True
        Me.txtUidHN.BackColor = System.Drawing.SystemColors.Window
        Me.txtUidHN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUidHN.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUidHN.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUidHN.Location = New System.Drawing.Point(902, 433)
        Me.txtUidHN.MaxLength = 0
        Me.txtUidHN.Name = "txtUidHN"
        Me.txtUidHN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUidHN.Size = New System.Drawing.Size(86, 20)
        Me.txtUidHN.TabIndex = 123
        Me.txtUidHN.Text = "txtUidHN"
        Me.txtUidHN.Visible = False
        '
        'txtTopRightEar
        '
        Me.txtTopRightEar.AcceptsReturn = True
        Me.txtTopRightEar.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopRightEar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTopRightEar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopRightEar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopRightEar.Location = New System.Drawing.Point(1000, 267)
        Me.txtTopRightEar.MaxLength = 0
        Me.txtTopRightEar.Name = "txtTopRightEar"
        Me.txtTopRightEar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopRightEar.Size = New System.Drawing.Size(46, 20)
        Me.txtTopRightEar.TabIndex = 120
        Me.txtTopRightEar.Text = "txtTopRightEar"
        Me.txtTopRightEar.Visible = False
        '
        'txtTopLeftEar
        '
        Me.txtTopLeftEar.AcceptsReturn = True
        Me.txtTopLeftEar.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopLeftEar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTopLeftEar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopLeftEar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopLeftEar.Location = New System.Drawing.Point(1000, 127)
        Me.txtTopLeftEar.MaxLength = 0
        Me.txtTopLeftEar.Name = "txtTopLeftEar"
        Me.txtTopLeftEar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopLeftEar.Size = New System.Drawing.Size(46, 20)
        Me.txtTopLeftEar.TabIndex = 119
        Me.txtTopLeftEar.Text = "txtTopLeftEar"
        Me.txtTopLeftEar.Visible = False
        '
        'txtHeadNeckUID
        '
        Me.txtHeadNeckUID.AcceptsReturn = True
        Me.txtHeadNeckUID.BackColor = System.Drawing.SystemColors.Window
        Me.txtHeadNeckUID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHeadNeckUID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHeadNeckUID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtHeadNeckUID.Location = New System.Drawing.Point(1000, 167)
        Me.txtHeadNeckUID.MaxLength = 0
        Me.txtHeadNeckUID.Name = "txtHeadNeckUID"
        Me.txtHeadNeckUID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHeadNeckUID.Size = New System.Drawing.Size(46, 20)
        Me.txtHeadNeckUID.TabIndex = 118
        Me.txtHeadNeckUID.Text = "txtHeadNeckUID"
        Me.txtHeadNeckUID.Visible = False
        '
        'txtOptionChoice
        '
        Me.txtOptionChoice.AcceptsReturn = True
        Me.txtOptionChoice.BackColor = System.Drawing.SystemColors.Window
        Me.txtOptionChoice.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOptionChoice.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOptionChoice.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOptionChoice.Location = New System.Drawing.Point(1000, 187)
        Me.txtOptionChoice.MaxLength = 0
        Me.txtOptionChoice.Name = "txtOptionChoice"
        Me.txtOptionChoice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOptionChoice.Size = New System.Drawing.Size(46, 20)
        Me.txtOptionChoice.TabIndex = 117
        Me.txtOptionChoice.Text = "txtOptionChoice"
        Me.txtOptionChoice.Visible = False
        '
        'txtFabric
        '
        Me.txtFabric.AcceptsReturn = True
        Me.txtFabric.BackColor = System.Drawing.SystemColors.Window
        Me.txtFabric.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFabric.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFabric.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFabric.Location = New System.Drawing.Point(1000, 307)
        Me.txtFabric.MaxLength = 0
        Me.txtFabric.Name = "txtFabric"
        Me.txtFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFabric.Size = New System.Drawing.Size(46, 20)
        Me.txtFabric.TabIndex = 116
        Me.txtFabric.Visible = False
        '
        'txtWorkOrder
        '
        Me.txtWorkOrder.AcceptsReturn = True
        Me.txtWorkOrder.BackColor = System.Drawing.SystemColors.Window
        Me.txtWorkOrder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtWorkOrder.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWorkOrder.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWorkOrder.Location = New System.Drawing.Point(900, 387)
        Me.txtWorkOrder.MaxLength = 0
        Me.txtWorkOrder.Name = "txtWorkOrder"
        Me.txtWorkOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWorkOrder.Size = New System.Drawing.Size(146, 20)
        Me.txtWorkOrder.TabIndex = 115
        Me.txtWorkOrder.Text = "txtWorkOrder"
        Me.txtWorkOrder.Visible = False
        '
        'txtData
        '
        Me.txtData.AcceptsReturn = True
        Me.txtData.BackColor = System.Drawing.SystemColors.Window
        Me.txtData.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtData.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtData.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.Location = New System.Drawing.Point(900, 367)
        Me.txtData.MaxLength = 19
        Me.txtData.Name = "txtData"
        Me.txtData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtData.Size = New System.Drawing.Size(146, 20)
        Me.txtData.TabIndex = 86
        Me.txtData.Text = "txtData"
        Me.txtData.Visible = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(664, 532)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(96, 26)
        Me.cmdClose.TabIndex = 36
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'txtMeasurements
        '
        Me.txtMeasurements.AcceptsReturn = True
        Me.txtMeasurements.BackColor = System.Drawing.SystemColors.Window
        Me.txtMeasurements.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMeasurements.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMeasurements.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMeasurements.Location = New System.Drawing.Point(1000, 147)
        Me.txtMeasurements.MaxLength = 43
        Me.txtMeasurements.Name = "txtMeasurements"
        Me.txtMeasurements.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMeasurements.Size = New System.Drawing.Size(46, 20)
        Me.txtMeasurements.TabIndex = 111
        Me.txtMeasurements.Text = "txtMeasurements"
        Me.txtMeasurements.Visible = False
        '
        'txtMouthRightY
        '
        Me.txtMouthRightY.AcceptsReturn = True
        Me.txtMouthRightY.BackColor = System.Drawing.SystemColors.Window
        Me.txtMouthRightY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMouthRightY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMouthRightY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMouthRightY.Location = New System.Drawing.Point(1000, 247)
        Me.txtMouthRightY.MaxLength = 0
        Me.txtMouthRightY.Name = "txtMouthRightY"
        Me.txtMouthRightY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMouthRightY.Size = New System.Drawing.Size(46, 20)
        Me.txtMouthRightY.TabIndex = 110
        Me.txtMouthRightY.Text = "txtMouthRightY"
        Me.txtMouthRightY.Visible = False
        '
        'txtxyNeckTopFrontY
        '
        Me.txtxyNeckTopFrontY.AcceptsReturn = True
        Me.txtxyNeckTopFrontY.BackColor = System.Drawing.SystemColors.Window
        Me.txtxyNeckTopFrontY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtxyNeckTopFrontY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtxyNeckTopFrontY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtxyNeckTopFrontY.Location = New System.Drawing.Point(1000, 347)
        Me.txtxyNeckTopFrontY.MaxLength = 0
        Me.txtxyNeckTopFrontY.Name = "txtxyNeckTopFrontY"
        Me.txtxyNeckTopFrontY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtxyNeckTopFrontY.Size = New System.Drawing.Size(46, 20)
        Me.txtxyNeckTopFrontY.TabIndex = 109
        Me.txtxyNeckTopFrontY.Text = "txtxyNeckTopFrontY"
        Me.txtxyNeckTopFrontY.Visible = False
        '
        'txtxyNeckTopFrontX
        '
        Me.txtxyNeckTopFrontX.AcceptsReturn = True
        Me.txtxyNeckTopFrontX.BackColor = System.Drawing.SystemColors.Window
        Me.txtxyNeckTopFrontX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtxyNeckTopFrontX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtxyNeckTopFrontX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtxyNeckTopFrontX.Location = New System.Drawing.Point(950, 347)
        Me.txtxyNeckTopFrontX.MaxLength = 0
        Me.txtxyNeckTopFrontX.Name = "txtxyNeckTopFrontX"
        Me.txtxyNeckTopFrontX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtxyNeckTopFrontX.Size = New System.Drawing.Size(46, 20)
        Me.txtxyNeckTopFrontX.TabIndex = 108
        Me.txtxyNeckTopFrontX.Text = "txtxyNeckTopFrontX"
        Me.txtxyNeckTopFrontX.Visible = False
        '
        'txtxyChinTopY
        '
        Me.txtxyChinTopY.AcceptsReturn = True
        Me.txtxyChinTopY.BackColor = System.Drawing.SystemColors.Window
        Me.txtxyChinTopY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtxyChinTopY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtxyChinTopY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtxyChinTopY.Location = New System.Drawing.Point(950, 327)
        Me.txtxyChinTopY.MaxLength = 0
        Me.txtxyChinTopY.Name = "txtxyChinTopY"
        Me.txtxyChinTopY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtxyChinTopY.Size = New System.Drawing.Size(46, 20)
        Me.txtxyChinTopY.TabIndex = 107
        Me.txtxyChinTopY.Text = "txtxyChinTopY"
        Me.txtxyChinTopY.Visible = False
        '
        'txtxyChinTopX
        '
        Me.txtxyChinTopX.AcceptsReturn = True
        Me.txtxyChinTopX.BackColor = System.Drawing.SystemColors.Window
        Me.txtxyChinTopX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtxyChinTopX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtxyChinTopX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtxyChinTopX.Location = New System.Drawing.Point(950, 307)
        Me.txtxyChinTopX.MaxLength = 0
        Me.txtxyChinTopX.Name = "txtxyChinTopX"
        Me.txtxyChinTopX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtxyChinTopX.Size = New System.Drawing.Size(46, 20)
        Me.txtxyChinTopX.TabIndex = 106
        Me.txtxyChinTopX.Text = "txtxyChinTopX"
        Me.txtxyChinTopX.Visible = False
        '
        'txtCSEarBotHeight
        '
        Me.txtCSEarBotHeight.AcceptsReturn = True
        Me.txtCSEarBotHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSEarBotHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSEarBotHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSEarBotHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSEarBotHeight.Location = New System.Drawing.Point(950, 287)
        Me.txtCSEarBotHeight.MaxLength = 0
        Me.txtCSEarBotHeight.Name = "txtCSEarBotHeight"
        Me.txtCSEarBotHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSEarBotHeight.Size = New System.Drawing.Size(46, 20)
        Me.txtCSEarBotHeight.TabIndex = 105
        Me.txtCSEarBotHeight.Text = "txtCSEarBotHeight"
        Me.txtCSEarBotHeight.Visible = False
        '
        'txtCSEarTopHeight
        '
        Me.txtCSEarTopHeight.AcceptsReturn = True
        Me.txtCSEarTopHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSEarTopHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSEarTopHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSEarTopHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSEarTopHeight.Location = New System.Drawing.Point(950, 267)
        Me.txtCSEarTopHeight.MaxLength = 0
        Me.txtCSEarTopHeight.Name = "txtCSEarTopHeight"
        Me.txtCSEarTopHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSEarTopHeight.Size = New System.Drawing.Size(46, 20)
        Me.txtCSEarTopHeight.TabIndex = 104
        Me.txtCSEarTopHeight.Text = "txtCSEarTopHeight"
        Me.txtCSEarTopHeight.Visible = False
        '
        'txtCSMouthWidth
        '
        Me.txtCSMouthWidth.AcceptsReturn = True
        Me.txtCSMouthWidth.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSMouthWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSMouthWidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSMouthWidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSMouthWidth.Location = New System.Drawing.Point(950, 247)
        Me.txtCSMouthWidth.MaxLength = 0
        Me.txtCSMouthWidth.Name = "txtCSMouthWidth"
        Me.txtCSMouthWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSMouthWidth.Size = New System.Drawing.Size(46, 20)
        Me.txtCSMouthWidth.TabIndex = 103
        Me.txtCSMouthWidth.Text = "txtCSMouthWidth"
        Me.txtCSMouthWidth.Visible = False
        '
        'txtMouthRightX
        '
        Me.txtMouthRightX.AcceptsReturn = True
        Me.txtMouthRightX.BackColor = System.Drawing.SystemColors.Window
        Me.txtMouthRightX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMouthRightX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMouthRightX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMouthRightX.Location = New System.Drawing.Point(1000, 107)
        Me.txtMouthRightX.MaxLength = 0
        Me.txtMouthRightX.Name = "txtMouthRightX"
        Me.txtMouthRightX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMouthRightX.Size = New System.Drawing.Size(46, 20)
        Me.txtMouthRightX.TabIndex = 67
        Me.txtMouthRightX.Text = "txtMouthRightX"
        Me.txtMouthRightX.Visible = False
        '
        'txtCSForeHead
        '
        Me.txtCSForeHead.AcceptsReturn = True
        Me.txtCSForeHead.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSForeHead.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSForeHead.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSForeHead.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSForeHead.Location = New System.Drawing.Point(950, 227)
        Me.txtCSForeHead.MaxLength = 0
        Me.txtCSForeHead.Name = "txtCSForeHead"
        Me.txtCSForeHead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSForeHead.Size = New System.Drawing.Size(46, 20)
        Me.txtCSForeHead.TabIndex = 99
        Me.txtCSForeHead.Text = "txtCSForeHead"
        Me.txtCSForeHead.Visible = False
        '
        'txtCSChinAngle
        '
        Me.txtCSChinAngle.AcceptsReturn = True
        Me.txtCSChinAngle.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSChinAngle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSChinAngle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSChinAngle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSChinAngle.Location = New System.Drawing.Point(950, 207)
        Me.txtCSChinAngle.MaxLength = 0
        Me.txtCSChinAngle.Name = "txtCSChinAngle"
        Me.txtCSChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSChinAngle.Size = New System.Drawing.Size(46, 20)
        Me.txtCSChinAngle.TabIndex = 98
        Me.txtCSChinAngle.Text = "txtCSChinAngle"
        Me.txtCSChinAngle.Visible = False
        '
        'txtCSNeckCircum
        '
        Me.txtCSNeckCircum.AcceptsReturn = True
        Me.txtCSNeckCircum.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSNeckCircum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSNeckCircum.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSNeckCircum.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSNeckCircum.Location = New System.Drawing.Point(950, 187)
        Me.txtCSNeckCircum.MaxLength = 0
        Me.txtCSNeckCircum.Name = "txtCSNeckCircum"
        Me.txtCSNeckCircum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSNeckCircum.Size = New System.Drawing.Size(46, 20)
        Me.txtCSNeckCircum.TabIndex = 97
        Me.txtCSNeckCircum.Text = "txtCSNeckCircum"
        Me.txtCSNeckCircum.Visible = False
        '
        'txtCSChintoMouth
        '
        Me.txtCSChintoMouth.AcceptsReturn = True
        Me.txtCSChintoMouth.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSChintoMouth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSChintoMouth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSChintoMouth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSChintoMouth.Location = New System.Drawing.Point(950, 167)
        Me.txtCSChintoMouth.MaxLength = 0
        Me.txtCSChintoMouth.Name = "txtCSChintoMouth"
        Me.txtCSChintoMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSChintoMouth.Size = New System.Drawing.Size(46, 20)
        Me.txtCSChintoMouth.TabIndex = 96
        Me.txtCSChintoMouth.Text = "txtCSChintoMouth"
        Me.txtCSChintoMouth.Visible = False
        '
        'txtNoseCoverX
        '
        Me.txtNoseCoverX.AcceptsReturn = True
        Me.txtNoseCoverX.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoseCoverX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNoseCoverX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNoseCoverX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNoseCoverX.Location = New System.Drawing.Point(1000, 87)
        Me.txtNoseCoverX.MaxLength = 0
        Me.txtNoseCoverX.Name = "txtNoseCoverX"
        Me.txtNoseCoverX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNoseCoverX.Size = New System.Drawing.Size(46, 20)
        Me.txtNoseCoverX.TabIndex = 95
        Me.txtNoseCoverX.Text = "txtNoseCoverX"
        Me.txtNoseCoverX.Visible = False
        '
        'txtNoseCoverY
        '
        Me.txtNoseCoverY.AcceptsReturn = True
        Me.txtNoseCoverY.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoseCoverY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNoseCoverY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNoseCoverY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNoseCoverY.Location = New System.Drawing.Point(1000, 67)
        Me.txtNoseCoverY.MaxLength = 0
        Me.txtNoseCoverY.Name = "txtNoseCoverY"
        Me.txtNoseCoverY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNoseCoverY.Size = New System.Drawing.Size(46, 20)
        Me.txtNoseCoverY.TabIndex = 94
        Me.txtNoseCoverY.Text = "txtNoseCoverY"
        Me.txtNoseCoverY.Visible = False
        '
        'txtEyeWidth
        '
        Me.txtEyeWidth.AcceptsReturn = True
        Me.txtEyeWidth.BackColor = System.Drawing.SystemColors.Window
        Me.txtEyeWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEyeWidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEyeWidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEyeWidth.Location = New System.Drawing.Point(950, 127)
        Me.txtEyeWidth.MaxLength = 0
        Me.txtEyeWidth.Name = "txtEyeWidth"
        Me.txtEyeWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEyeWidth.Size = New System.Drawing.Size(46, 20)
        Me.txtEyeWidth.TabIndex = 93
        Me.txtEyeWidth.Text = "txtEyeWidth"
        Me.txtEyeWidth.Visible = False
        '
        'txtDartEndX
        '
        Me.txtDartEndX.AcceptsReturn = True
        Me.txtDartEndX.BackColor = System.Drawing.SystemColors.Window
        Me.txtDartEndX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDartEndX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDartEndX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDartEndX.Location = New System.Drawing.Point(950, 87)
        Me.txtDartEndX.MaxLength = 0
        Me.txtDartEndX.Name = "txtDartEndX"
        Me.txtDartEndX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDartEndX.Size = New System.Drawing.Size(46, 20)
        Me.txtDartEndX.TabIndex = 92
        Me.txtDartEndX.Text = "txtDartEndX"
        Me.txtDartEndX.Visible = False
        '
        'txtDartEndY
        '
        Me.txtDartEndY.AcceptsReturn = True
        Me.txtDartEndY.BackColor = System.Drawing.SystemColors.Window
        Me.txtDartEndY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDartEndY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDartEndY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDartEndY.Location = New System.Drawing.Point(950, 107)
        Me.txtDartEndY.MaxLength = 0
        Me.txtDartEndY.Name = "txtDartEndY"
        Me.txtDartEndY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDartEndY.Size = New System.Drawing.Size(46, 20)
        Me.txtDartEndY.TabIndex = 91
        Me.txtDartEndY.Text = "txtDartEndY"
        Me.txtDartEndY.Visible = False
        '
        'txtDartStartY
        '
        Me.txtDartStartY.AcceptsReturn = True
        Me.txtDartStartY.BackColor = System.Drawing.SystemColors.Window
        Me.txtDartStartY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDartStartY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDartStartY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDartStartY.Location = New System.Drawing.Point(950, 67)
        Me.txtDartStartY.MaxLength = 0
        Me.txtDartStartY.Name = "txtDartStartY"
        Me.txtDartStartY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDartStartY.Size = New System.Drawing.Size(46, 20)
        Me.txtDartStartY.TabIndex = 90
        Me.txtDartStartY.Text = "txtDartStartY"
        Me.txtDartStartY.Visible = False
        '
        'txtDartStartX
        '
        Me.txtDartStartX.AcceptsReturn = True
        Me.txtDartStartX.BackColor = System.Drawing.SystemColors.Window
        Me.txtDartStartX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDartStartX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDartStartX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDartStartX.Location = New System.Drawing.Point(950, 47)
        Me.txtDartStartX.MaxLength = 0
        Me.txtDartStartX.Name = "txtDartStartX"
        Me.txtDartStartX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDartStartX.Size = New System.Drawing.Size(46, 20)
        Me.txtDartStartX.TabIndex = 89
        Me.txtDartStartX.Text = "txtDartStartX"
        Me.txtDartStartX.Visible = False
        '
        'txtChinLeftBotY
        '
        Me.txtChinLeftBotY.AcceptsReturn = True
        Me.txtChinLeftBotY.BackColor = System.Drawing.SystemColors.Window
        Me.txtChinLeftBotY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtChinLeftBotY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtChinLeftBotY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtChinLeftBotY.Location = New System.Drawing.Point(1000, 227)
        Me.txtChinLeftBotY.MaxLength = 0
        Me.txtChinLeftBotY.Name = "txtChinLeftBotY"
        Me.txtChinLeftBotY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtChinLeftBotY.Size = New System.Drawing.Size(46, 20)
        Me.txtChinLeftBotY.TabIndex = 88
        Me.txtChinLeftBotY.Text = "txtChinLeftBotY"
        Me.txtChinLeftBotY.Visible = False
        '
        'txtChinLeftBotX
        '
        Me.txtChinLeftBotX.AcceptsReturn = True
        Me.txtChinLeftBotX.BackColor = System.Drawing.SystemColors.Window
        Me.txtChinLeftBotX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtChinLeftBotX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtChinLeftBotX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtChinLeftBotX.Location = New System.Drawing.Point(1000, 207)
        Me.txtChinLeftBotX.MaxLength = 0
        Me.txtChinLeftBotX.Name = "txtChinLeftBotX"
        Me.txtChinLeftBotX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtChinLeftBotX.Size = New System.Drawing.Size(46, 20)
        Me.txtChinLeftBotX.TabIndex = 85
        Me.txtChinLeftBotX.Text = "txtChinLeftBotX"
        Me.txtChinLeftBotX.Visible = False
        '
        'txtNoseBottomY
        '
        Me.txtNoseBottomY.AcceptsReturn = True
        Me.txtNoseBottomY.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoseBottomY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNoseBottomY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNoseBottomY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNoseBottomY.Location = New System.Drawing.Point(1000, 47)
        Me.txtNoseBottomY.MaxLength = 0
        Me.txtNoseBottomY.Name = "txtNoseBottomY"
        Me.txtNoseBottomY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNoseBottomY.Size = New System.Drawing.Size(46, 20)
        Me.txtNoseBottomY.TabIndex = 84
        Me.txtNoseBottomY.Text = "txtNoseBottomY"
        Me.txtNoseBottomY.Visible = False
        '
        'txtCircumferenceTotal
        '
        Me.txtCircumferenceTotal.AcceptsReturn = True
        Me.txtCircumferenceTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCircumferenceTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCircumferenceTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCircumferenceTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCircumferenceTotal.Location = New System.Drawing.Point(900, 287)
        Me.txtCircumferenceTotal.MaxLength = 0
        Me.txtCircumferenceTotal.Name = "txtCircumferenceTotal"
        Me.txtCircumferenceTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCircumferenceTotal.Size = New System.Drawing.Size(46, 20)
        Me.txtCircumferenceTotal.TabIndex = 83
        Me.txtCircumferenceTotal.Text = "txtCircumferenceTotal"
        Me.txtCircumferenceTotal.Visible = False
        '
        'txtMouthHeight
        '
        Me.txtMouthHeight.AcceptsReturn = True
        Me.txtMouthHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtMouthHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMouthHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMouthHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMouthHeight.Location = New System.Drawing.Point(950, 147)
        Me.txtMouthHeight.MaxLength = 0
        Me.txtMouthHeight.Name = "txtMouthHeight"
        Me.txtMouthHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMouthHeight.Size = New System.Drawing.Size(46, 20)
        Me.txtMouthHeight.TabIndex = 82
        Me.txtMouthHeight.Text = "txtMouthHeight"
        Me.txtMouthHeight.Visible = False
        '
        'txtBotOfEyeY
        '
        Me.txtBotOfEyeY.AcceptsReturn = True
        Me.txtBotOfEyeY.BackColor = System.Drawing.SystemColors.Window
        Me.txtBotOfEyeY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBotOfEyeY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBotOfEyeY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBotOfEyeY.Location = New System.Drawing.Point(900, 347)
        Me.txtBotOfEyeY.MaxLength = 0
        Me.txtBotOfEyeY.Name = "txtBotOfEyeY"
        Me.txtBotOfEyeY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBotOfEyeY.Size = New System.Drawing.Size(46, 20)
        Me.txtBotOfEyeY.TabIndex = 81
        Me.txtBotOfEyeY.Text = "txtBotOfEyeY"
        Me.txtBotOfEyeY.Visible = False
        '
        'txtBotOfEyeX
        '
        Me.txtBotOfEyeX.AcceptsReturn = True
        Me.txtBotOfEyeX.BackColor = System.Drawing.SystemColors.Window
        Me.txtBotOfEyeX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBotOfEyeX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBotOfEyeX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBotOfEyeX.Location = New System.Drawing.Point(900, 327)
        Me.txtBotOfEyeX.MaxLength = 0
        Me.txtBotOfEyeX.Name = "txtBotOfEyeX"
        Me.txtBotOfEyeX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBotOfEyeX.Size = New System.Drawing.Size(46, 20)
        Me.txtBotOfEyeX.TabIndex = 80
        Me.txtBotOfEyeX.Text = "txtBotOfEyeX"
        Me.txtBotOfEyeX.Visible = False
        '
        'txtLowerEarHeight
        '
        Me.txtLowerEarHeight.AcceptsReturn = True
        Me.txtLowerEarHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtLowerEarHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLowerEarHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLowerEarHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLowerEarHeight.Location = New System.Drawing.Point(900, 307)
        Me.txtLowerEarHeight.MaxLength = 0
        Me.txtLowerEarHeight.Name = "txtLowerEarHeight"
        Me.txtLowerEarHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLowerEarHeight.Size = New System.Drawing.Size(46, 20)
        Me.txtLowerEarHeight.TabIndex = 79
        Me.txtLowerEarHeight.Text = "txtLowerEarHeight"
        Me.txtLowerEarHeight.Visible = False
        '
        'txtfNum
        '
        Me.txtfNum.AcceptsReturn = True
        Me.txtfNum.BackColor = System.Drawing.SystemColors.Window
        Me.txtfNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfNum.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtfNum.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtfNum.Location = New System.Drawing.Point(900, 267)
        Me.txtfNum.MaxLength = 0
        Me.txtfNum.Name = "txtfNum"
        Me.txtfNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtfNum.Size = New System.Drawing.Size(46, 20)
        Me.txtfNum.TabIndex = 78
        Me.txtfNum.Text = "txtfNum"
        Me.txtfNum.Visible = False
        '
        'txtCurrTextFont
        '
        Me.txtCurrTextFont.AcceptsReturn = True
        Me.txtCurrTextFont.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrTextFont.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrTextFont.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrTextFont.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrTextFont.Location = New System.Drawing.Point(900, 247)
        Me.txtCurrTextFont.MaxLength = 0
        Me.txtCurrTextFont.Name = "txtCurrTextFont"
        Me.txtCurrTextFont.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrTextFont.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrTextFont.TabIndex = 77
        Me.txtCurrTextFont.Text = "txtCurrTextFont"
        Me.txtCurrTextFont.Visible = False
        '
        'txtCurrTextVertJust
        '
        Me.txtCurrTextVertJust.AcceptsReturn = True
        Me.txtCurrTextVertJust.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrTextVertJust.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrTextVertJust.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrTextVertJust.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrTextVertJust.Location = New System.Drawing.Point(900, 227)
        Me.txtCurrTextVertJust.MaxLength = 0
        Me.txtCurrTextVertJust.Name = "txtCurrTextVertJust"
        Me.txtCurrTextVertJust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrTextVertJust.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrTextVertJust.TabIndex = 76
        Me.txtCurrTextVertJust.Text = "txtCurrTextVertJust"
        Me.txtCurrTextVertJust.Visible = False
        '
        'txtCurrTextHorizJust
        '
        Me.txtCurrTextHorizJust.AcceptsReturn = True
        Me.txtCurrTextHorizJust.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrTextHorizJust.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrTextHorizJust.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrTextHorizJust.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrTextHorizJust.Location = New System.Drawing.Point(900, 207)
        Me.txtCurrTextHorizJust.MaxLength = 0
        Me.txtCurrTextHorizJust.Name = "txtCurrTextHorizJust"
        Me.txtCurrTextHorizJust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrTextHorizJust.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrTextHorizJust.TabIndex = 75
        Me.txtCurrTextHorizJust.Text = "txtCurrTextHorizJust"
        Me.txtCurrTextHorizJust.Visible = False
        '
        'txtCurrTextAspect
        '
        Me.txtCurrTextAspect.AcceptsReturn = True
        Me.txtCurrTextAspect.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrTextAspect.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrTextAspect.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrTextAspect.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrTextAspect.Location = New System.Drawing.Point(900, 187)
        Me.txtCurrTextAspect.MaxLength = 0
        Me.txtCurrTextAspect.Name = "txtCurrTextAspect"
        Me.txtCurrTextAspect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrTextAspect.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrTextAspect.TabIndex = 74
        Me.txtCurrTextAspect.Text = "txtCurrTextAspect"
        Me.txtCurrTextAspect.Visible = False
        '
        'txtCurrTextHt
        '
        Me.txtCurrTextHt.AcceptsReturn = True
        Me.txtCurrTextHt.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrTextHt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrTextHt.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrTextHt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrTextHt.Location = New System.Drawing.Point(900, 167)
        Me.txtCurrTextHt.MaxLength = 0
        Me.txtCurrTextHt.Name = "txtCurrTextHt"
        Me.txtCurrTextHt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrTextHt.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrTextHt.TabIndex = 73
        Me.txtCurrTextHt.Text = "txtCurrTextHt"
        Me.txtCurrTextHt.Visible = False
        '
        'txtCurrentLayer
        '
        Me.txtCurrentLayer.AcceptsReturn = True
        Me.txtCurrentLayer.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrentLayer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrentLayer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrentLayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrentLayer.Location = New System.Drawing.Point(900, 147)
        Me.txtCurrentLayer.MaxLength = 0
        Me.txtCurrentLayer.Name = "txtCurrentLayer"
        Me.txtCurrentLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrentLayer.Size = New System.Drawing.Size(46, 20)
        Me.txtCurrentLayer.TabIndex = 72
        Me.txtCurrentLayer.Text = "txtCurrentLayer"
        Me.txtCurrentLayer.Visible = False
        '
        'txtMidToEyeTop
        '
        Me.txtMidToEyeTop.AcceptsReturn = True
        Me.txtMidToEyeTop.BackColor = System.Drawing.SystemColors.Window
        Me.txtMidToEyeTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMidToEyeTop.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMidToEyeTop.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMidToEyeTop.Location = New System.Drawing.Point(900, 127)
        Me.txtMidToEyeTop.MaxLength = 0
        Me.txtMidToEyeTop.Name = "txtMidToEyeTop"
        Me.txtMidToEyeTop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMidToEyeTop.Size = New System.Drawing.Size(46, 20)
        Me.txtMidToEyeTop.TabIndex = 71
        Me.txtMidToEyeTop.Text = "txtMidToEyeTop"
        Me.txtMidToEyeTop.Visible = False
        '
        'txtLipStrapWidth
        '
        Me.txtLipStrapWidth.AcceptsReturn = True
        Me.txtLipStrapWidth.BackColor = System.Drawing.SystemColors.Window
        Me.txtLipStrapWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLipStrapWidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLipStrapWidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLipStrapWidth.Location = New System.Drawing.Point(900, 107)
        Me.txtLipStrapWidth.MaxLength = 0
        Me.txtLipStrapWidth.Name = "txtLipStrapWidth"
        Me.txtLipStrapWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLipStrapWidth.Size = New System.Drawing.Size(46, 20)
        Me.txtLipStrapWidth.TabIndex = 70
        Me.txtLipStrapWidth.Text = "txtLipStrapWidth"
        Me.txtLipStrapWidth.Visible = False
        '
        'txtUidMPD
        '
        Me.txtUidMPD.AcceptsReturn = True
        Me.txtUidMPD.BackColor = System.Drawing.SystemColors.Window
        Me.txtUidMPD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUidMPD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUidMPD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUidMPD.Location = New System.Drawing.Point(902, 413)
        Me.txtUidMPD.MaxLength = 0
        Me.txtUidMPD.Name = "txtUidMPD"
        Me.txtUidMPD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUidMPD.Size = New System.Drawing.Size(86, 20)
        Me.txtUidMPD.TabIndex = 69
        Me.txtUidMPD.Text = "txtUidMPD"
        Me.txtUidMPD.Visible = False
        '
        'txtRadiusNo
        '
        Me.txtRadiusNo.AcceptsReturn = True
        Me.txtRadiusNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtRadiusNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRadiusNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRadiusNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRadiusNo.Location = New System.Drawing.Point(900, 67)
        Me.txtRadiusNo.MaxLength = 0
        Me.txtRadiusNo.Name = "txtRadiusNo"
        Me.txtRadiusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRadiusNo.Size = New System.Drawing.Size(46, 20)
        Me.txtRadiusNo.TabIndex = 68
        Me.txtRadiusNo.Text = "txtRadiusNo"
        Me.txtRadiusNo.Visible = False
        '
        'txtDraw
        '
        Me.txtDraw.AcceptsReturn = True
        Me.txtDraw.BackColor = System.Drawing.SystemColors.Window
        Me.txtDraw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDraw.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDraw.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDraw.Location = New System.Drawing.Point(902, 47)
        Me.txtDraw.MaxLength = 0
        Me.txtDraw.Name = "txtDraw"
        Me.txtDraw.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDraw.Size = New System.Drawing.Size(46, 20)
        Me.txtDraw.TabIndex = 102
        Me.txtDraw.Text = "txtDraw"
        Me.txtDraw.Visible = False
        '
        'cmdDraw
        '
        Me.cmdDraw.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDraw.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDraw.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDraw.Location = New System.Drawing.Point(557, 532)
        Me.cmdDraw.Name = "cmdDraw"
        Me.cmdDraw.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDraw.Size = New System.Drawing.Size(96, 26)
        Me.cmdDraw.TabIndex = 35
        Me.cmdDraw.Text = "&Draw"
        Me.cmdDraw.UseVisualStyleBackColor = False
        '
        'frmMeasurements
        '
        Me.frmMeasurements.BackColor = System.Drawing.SystemColors.Control
        Me.frmMeasurements.Controls.Add(Me.Label11)
        Me.frmMeasurements.Controls.Add(Me.Label10)
        Me.frmMeasurements.Controls.Add(Me.Label9)
        Me.frmMeasurements.Controls.Add(Me.Label8)
        Me.frmMeasurements.Controls.Add(Me.Label7)
        Me.frmMeasurements.Controls.Add(Me.Label3)
        Me.frmMeasurements.Controls.Add(Me.Label2)
        Me.frmMeasurements.Controls.Add(Me.txtChinCollarMin)
        Me.frmMeasurements.Controls.Add(Me.txtRightEarLength)
        Me.frmMeasurements.Controls.Add(Me.txtLeftEarLength)
        Me.frmMeasurements.Controls.Add(Me.txtHeadBandDepth)
        Me.frmMeasurements.Controls.Add(Me.txtLengthOfNose)
        Me.frmMeasurements.Controls.Add(Me.txtTipOfNose)
        Me.frmMeasurements.Controls.Add(Me.cboFabric)
        Me.frmMeasurements.Controls.Add(Me.txtThroatToSternal)
        Me.frmMeasurements.Controls.Add(Me.txtCircOfNeck)
        Me.frmMeasurements.Controls.Add(Me.txtCircChinAngle)
        Me.frmMeasurements.Controls.Add(Me.txtCircEyeBrow)
        Me.frmMeasurements.Controls.Add(Me.txtChinToMouth)
        Me.frmMeasurements.Controls.Add(Me.lblInch11)
        Me.frmMeasurements.Controls.Add(Me.lblChinCollarMin)
        Me.frmMeasurements.Controls.Add(Me.lblRightEarLen)
        Me.frmMeasurements.Controls.Add(Me.lblLeftEarLen)
        Me.frmMeasurements.Controls.Add(Me.lblInch10)
        Me.frmMeasurements.Controls.Add(Me.lblInch9)
        Me.frmMeasurements.Controls.Add(Me.lblInch8)
        Me.frmMeasurements.Controls.Add(Me.lblHeadBandDepth)
        Me.frmMeasurements.Controls.Add(Me.lblFabric)
        Me.frmMeasurements.Controls.Add(Me.lblInch7)
        Me.frmMeasurements.Controls.Add(Me.lblInch6)
        Me.frmMeasurements.Controls.Add(Me.LblLengthOfNose)
        Me.frmMeasurements.Controls.Add(Me.lblAcrossTipOfNose)
        Me.frmMeasurements.Controls.Add(Me.lblInch5)
        Me.frmMeasurements.Controls.Add(Me.lblInch4)
        Me.frmMeasurements.Controls.Add(Me.lblInch3)
        Me.frmMeasurements.Controls.Add(Me.lblInch2)
        Me.frmMeasurements.Controls.Add(Me.lblInch1)
        Me.frmMeasurements.Controls.Add(Me.lblThroatToSternal)
        Me.frmMeasurements.Controls.Add(Me.lblCircOfNeck)
        Me.frmMeasurements.Controls.Add(Me.lblCircChinAngle)
        Me.frmMeasurements.Controls.Add(Me.lblCircEyeBrow)
        Me.frmMeasurements.Controls.Add(Me.lblChinToMouth)
        Me.frmMeasurements.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmMeasurements.Location = New System.Drawing.Point(261, 377)
        Me.frmMeasurements.Name = "frmMeasurements"
        Me.frmMeasurements.Padding = New System.Windows.Forms.Padding(0)
        Me.frmMeasurements.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmMeasurements.Size = New System.Drawing.Size(499, 151)
        Me.frmMeasurements.TabIndex = 51
        Me.frmMeasurements.TabStop = False
        Me.frmMeasurements.Text = "Measurements"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(374, 60)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(13, 13)
        Me.Label11.TabIndex = 129
        Me.Label11.Text = "1"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(374, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(13, 13)
        Me.Label10.TabIndex = 128
        Me.Label10.Text = "1"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(178, 100)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(13, 13)
        Me.Label9.TabIndex = 127
        Me.Label9.Text = "8"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(178, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(13, 13)
        Me.Label8.TabIndex = 126
        Me.Label8.Text = "7"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(178, 60)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(13, 13)
        Me.Label7.TabIndex = 125
        Me.Label7.Text = "6"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(178, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(13, 13)
        Me.Label3.TabIndex = 124
        Me.Label3.Text = "5"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(178, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(13, 13)
        Me.Label2.TabIndex = 123
        Me.Label2.Text = "4"
        '
        'txtChinCollarMin
        '
        Me.txtChinCollarMin.AcceptsReturn = True
        Me.txtChinCollarMin.BackColor = System.Drawing.SystemColors.Window
        Me.txtChinCollarMin.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtChinCollarMin.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtChinCollarMin.Location = New System.Drawing.Point(391, 75)
        Me.txtChinCollarMin.MaxLength = 0
        Me.txtChinCollarMin.Name = "txtChinCollarMin"
        Me.txtChinCollarMin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtChinCollarMin.Size = New System.Drawing.Size(36, 20)
        Me.txtChinCollarMin.TabIndex = 33
        '
        'txtRightEarLength
        '
        Me.txtRightEarLength.AcceptsReturn = True
        Me.txtRightEarLength.BackColor = System.Drawing.SystemColors.Window
        Me.txtRightEarLength.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRightEarLength.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRightEarLength.Location = New System.Drawing.Point(391, 55)
        Me.txtRightEarLength.MaxLength = 4
        Me.txtRightEarLength.Name = "txtRightEarLength"
        Me.txtRightEarLength.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRightEarLength.Size = New System.Drawing.Size(36, 20)
        Me.txtRightEarLength.TabIndex = 32
        Me.txtRightEarLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLeftEarLength
        '
        Me.txtLeftEarLength.AcceptsReturn = True
        Me.txtLeftEarLength.BackColor = System.Drawing.SystemColors.Window
        Me.txtLeftEarLength.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLeftEarLength.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLeftEarLength.Location = New System.Drawing.Point(391, 35)
        Me.txtLeftEarLength.MaxLength = 4
        Me.txtLeftEarLength.Name = "txtLeftEarLength"
        Me.txtLeftEarLength.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLeftEarLength.Size = New System.Drawing.Size(36, 20)
        Me.txtLeftEarLength.TabIndex = 31
        Me.txtLeftEarLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtHeadBandDepth
        '
        Me.txtHeadBandDepth.AcceptsReturn = True
        Me.txtHeadBandDepth.BackColor = System.Drawing.SystemColors.Window
        Me.txtHeadBandDepth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHeadBandDepth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtHeadBandDepth.Location = New System.Drawing.Point(391, 15)
        Me.txtHeadBandDepth.MaxLength = 4
        Me.txtHeadBandDepth.Name = "txtHeadBandDepth"
        Me.txtHeadBandDepth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHeadBandDepth.Size = New System.Drawing.Size(36, 20)
        Me.txtHeadBandDepth.TabIndex = 30
        Me.txtHeadBandDepth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLengthOfNose
        '
        Me.txtLengthOfNose.AcceptsReturn = True
        Me.txtLengthOfNose.BackColor = System.Drawing.SystemColors.Window
        Me.txtLengthOfNose.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLengthOfNose.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLengthOfNose.Location = New System.Drawing.Point(185, 181)
        Me.txtLengthOfNose.MaxLength = 4
        Me.txtLengthOfNose.Name = "txtLengthOfNose"
        Me.txtLengthOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLengthOfNose.Size = New System.Drawing.Size(36, 20)
        Me.txtLengthOfNose.TabIndex = 29
        Me.txtLengthOfNose.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtLengthOfNose.Visible = False
        '
        'txtTipOfNose
        '
        Me.txtTipOfNose.AcceptsReturn = True
        Me.txtTipOfNose.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipOfNose.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipOfNose.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipOfNose.Location = New System.Drawing.Point(185, 161)
        Me.txtTipOfNose.MaxLength = 4
        Me.txtTipOfNose.Name = "txtTipOfNose"
        Me.txtTipOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipOfNose.Size = New System.Drawing.Size(36, 20)
        Me.txtTipOfNose.TabIndex = 28
        Me.txtTipOfNose.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtTipOfNose.Visible = False
        '
        'cboFabric
        '
        Me.cboFabric.BackColor = System.Drawing.SystemColors.Window
        Me.cboFabric.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboFabric.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFabric.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboFabric.Location = New System.Drawing.Point(280, 125)
        Me.cboFabric.Name = "cboFabric"
        Me.cboFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboFabric.Size = New System.Drawing.Size(146, 21)
        Me.cboFabric.TabIndex = 34
        '
        'txtThroatToSternal
        '
        Me.txtThroatToSternal.AcceptsReturn = True
        Me.txtThroatToSternal.BackColor = System.Drawing.SystemColors.Window
        Me.txtThroatToSternal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtThroatToSternal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtThroatToSternal.Location = New System.Drawing.Point(196, 95)
        Me.txtThroatToSternal.MaxLength = 4
        Me.txtThroatToSternal.Name = "txtThroatToSternal"
        Me.txtThroatToSternal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtThroatToSternal.Size = New System.Drawing.Size(36, 20)
        Me.txtThroatToSternal.TabIndex = 27
        Me.txtThroatToSternal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCircOfNeck
        '
        Me.txtCircOfNeck.AcceptsReturn = True
        Me.txtCircOfNeck.BackColor = System.Drawing.SystemColors.Window
        Me.txtCircOfNeck.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCircOfNeck.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCircOfNeck.Location = New System.Drawing.Point(196, 75)
        Me.txtCircOfNeck.MaxLength = 4
        Me.txtCircOfNeck.Name = "txtCircOfNeck"
        Me.txtCircOfNeck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCircOfNeck.Size = New System.Drawing.Size(36, 20)
        Me.txtCircOfNeck.TabIndex = 26
        Me.txtCircOfNeck.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCircChinAngle
        '
        Me.txtCircChinAngle.AcceptsReturn = True
        Me.txtCircChinAngle.BackColor = System.Drawing.SystemColors.Window
        Me.txtCircChinAngle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCircChinAngle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCircChinAngle.Location = New System.Drawing.Point(196, 55)
        Me.txtCircChinAngle.MaxLength = 4
        Me.txtCircChinAngle.Name = "txtCircChinAngle"
        Me.txtCircChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCircChinAngle.Size = New System.Drawing.Size(36, 20)
        Me.txtCircChinAngle.TabIndex = 25
        Me.txtCircChinAngle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCircEyeBrow
        '
        Me.txtCircEyeBrow.AcceptsReturn = True
        Me.txtCircEyeBrow.BackColor = System.Drawing.SystemColors.Window
        Me.txtCircEyeBrow.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCircEyeBrow.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCircEyeBrow.Location = New System.Drawing.Point(196, 35)
        Me.txtCircEyeBrow.MaxLength = 4
        Me.txtCircEyeBrow.Name = "txtCircEyeBrow"
        Me.txtCircEyeBrow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCircEyeBrow.Size = New System.Drawing.Size(36, 20)
        Me.txtCircEyeBrow.TabIndex = 24
        Me.txtCircEyeBrow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtChinToMouth
        '
        Me.txtChinToMouth.AcceptsReturn = True
        Me.txtChinToMouth.BackColor = System.Drawing.SystemColors.Window
        Me.txtChinToMouth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtChinToMouth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtChinToMouth.Location = New System.Drawing.Point(196, 15)
        Me.txtChinToMouth.MaxLength = 4
        Me.txtChinToMouth.Name = "txtChinToMouth"
        Me.txtChinToMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtChinToMouth.Size = New System.Drawing.Size(36, 20)
        Me.txtChinToMouth.TabIndex = 23
        Me.txtChinToMouth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblInch11
        '
        Me.lblInch11.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch11.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch11.Location = New System.Drawing.Point(425, 78)
        Me.lblInch11.Name = "lblInch11"
        Me.lblInch11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch11.Size = New System.Drawing.Size(46, 16)
        Me.lblInch11.TabIndex = 122
        Me.lblInch11.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblChinCollarMin
        '
        Me.lblChinCollarMin.BackColor = System.Drawing.SystemColors.Control
        Me.lblChinCollarMin.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblChinCollarMin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblChinCollarMin.Location = New System.Drawing.Point(280, 80)
        Me.lblChinCollarMin.Name = "lblChinCollarMin"
        Me.lblChinCollarMin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblChinCollarMin.Size = New System.Drawing.Size(101, 16)
        Me.lblChinCollarMin.TabIndex = 121
        Me.lblChinCollarMin.Text = "Collar Contour"
        '
        'lblRightEarLen
        '
        Me.lblRightEarLen.BackColor = System.Drawing.SystemColors.Control
        Me.lblRightEarLen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRightEarLen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRightEarLen.Location = New System.Drawing.Point(280, 60)
        Me.lblRightEarLen.Name = "lblRightEarLen"
        Me.lblRightEarLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRightEarLen.Size = New System.Drawing.Size(92, 16)
        Me.lblRightEarLen.TabIndex = 87
        Me.lblRightEarLen.Text = "Right Ear Length"
        '
        'lblLeftEarLen
        '
        Me.lblLeftEarLen.BackColor = System.Drawing.SystemColors.Control
        Me.lblLeftEarLen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLeftEarLen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLeftEarLen.Location = New System.Drawing.Point(280, 40)
        Me.lblLeftEarLen.Name = "lblLeftEarLen"
        Me.lblLeftEarLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLeftEarLen.Size = New System.Drawing.Size(82, 16)
        Me.lblLeftEarLen.TabIndex = 114
        Me.lblLeftEarLen.Text = "Left Ear Length"
        '
        'lblInch10
        '
        Me.lblInch10.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch10.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch10.Location = New System.Drawing.Point(425, 58)
        Me.lblInch10.Name = "lblInch10"
        Me.lblInch10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch10.Size = New System.Drawing.Size(46, 16)
        Me.lblInch10.TabIndex = 113
        Me.lblInch10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch9
        '
        Me.lblInch9.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch9.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch9.Location = New System.Drawing.Point(425, 38)
        Me.lblInch9.Name = "lblInch9"
        Me.lblInch9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch9.Size = New System.Drawing.Size(46, 16)
        Me.lblInch9.TabIndex = 112
        Me.lblInch9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch8
        '
        Me.lblInch8.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch8.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch8.Location = New System.Drawing.Point(425, 18)
        Me.lblInch8.Name = "lblInch8"
        Me.lblInch8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch8.Size = New System.Drawing.Size(46, 16)
        Me.lblInch8.TabIndex = 101
        Me.lblInch8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblHeadBandDepth
        '
        Me.lblHeadBandDepth.BackColor = System.Drawing.SystemColors.Control
        Me.lblHeadBandDepth.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHeadBandDepth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblHeadBandDepth.Location = New System.Drawing.Point(280, 20)
        Me.lblHeadBandDepth.Name = "lblHeadBandDepth"
        Me.lblHeadBandDepth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHeadBandDepth.Size = New System.Drawing.Size(111, 16)
        Me.lblHeadBandDepth.TabIndex = 100
        Me.lblHeadBandDepth.Text = "Head Band Depth"
        '
        'lblFabric
        '
        Me.lblFabric.BackColor = System.Drawing.SystemColors.Control
        Me.lblFabric.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFabric.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFabric.Location = New System.Drawing.Point(280, 110)
        Me.lblFabric.Name = "lblFabric"
        Me.lblFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFabric.Size = New System.Drawing.Size(46, 16)
        Me.lblFabric.TabIndex = 66
        Me.lblFabric.Text = "Fabric : "
        '
        'lblInch7
        '
        Me.lblInch7.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch7.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch7.Location = New System.Drawing.Point(230, 125)
        Me.lblInch7.Name = "lblInch7"
        Me.lblInch7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch7.Size = New System.Drawing.Size(46, 16)
        Me.lblInch7.TabIndex = 65
        Me.lblInch7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblInch7.Visible = False
        '
        'lblInch6
        '
        Me.lblInch6.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch6.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch6.Location = New System.Drawing.Point(230, 118)
        Me.lblInch6.Name = "lblInch6"
        Me.lblInch6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch6.Size = New System.Drawing.Size(46, 16)
        Me.lblInch6.TabIndex = 64
        Me.lblInch6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblInch6.Visible = False
        '
        'LblLengthOfNose
        '
        Me.LblLengthOfNose.BackColor = System.Drawing.SystemColors.Control
        Me.LblLengthOfNose.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblLengthOfNose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblLengthOfNose.Location = New System.Drawing.Point(10, 186)
        Me.LblLengthOfNose.Name = "LblLengthOfNose"
        Me.LblLengthOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblLengthOfNose.Size = New System.Drawing.Size(101, 16)
        Me.LblLengthOfNose.TabIndex = 63
        Me.LblLengthOfNose.Text = "Length of Nose"
        Me.LblLengthOfNose.Visible = False
        '
        'lblAcrossTipOfNose
        '
        Me.lblAcrossTipOfNose.BackColor = System.Drawing.SystemColors.Control
        Me.lblAcrossTipOfNose.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAcrossTipOfNose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAcrossTipOfNose.Location = New System.Drawing.Point(10, 166)
        Me.lblAcrossTipOfNose.Name = "lblAcrossTipOfNose"
        Me.lblAcrossTipOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAcrossTipOfNose.Size = New System.Drawing.Size(116, 16)
        Me.lblAcrossTipOfNose.TabIndex = 62
        Me.lblAcrossTipOfNose.Text = "Across Tip of Nose"
        Me.lblAcrossTipOfNose.Visible = False
        '
        'lblInch5
        '
        Me.lblInch5.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch5.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch5.Location = New System.Drawing.Point(230, 98)
        Me.lblInch5.Name = "lblInch5"
        Me.lblInch5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch5.Size = New System.Drawing.Size(46, 16)
        Me.lblInch5.TabIndex = 61
        Me.lblInch5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch4
        '
        Me.lblInch4.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch4.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch4.Location = New System.Drawing.Point(230, 78)
        Me.lblInch4.Name = "lblInch4"
        Me.lblInch4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch4.Size = New System.Drawing.Size(46, 16)
        Me.lblInch4.TabIndex = 60
        Me.lblInch4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch3
        '
        Me.lblInch3.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch3.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch3.Location = New System.Drawing.Point(230, 58)
        Me.lblInch3.Name = "lblInch3"
        Me.lblInch3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch3.Size = New System.Drawing.Size(46, 16)
        Me.lblInch3.TabIndex = 59
        Me.lblInch3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch2
        '
        Me.lblInch2.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch2.Location = New System.Drawing.Point(230, 38)
        Me.lblInch2.Name = "lblInch2"
        Me.lblInch2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch2.Size = New System.Drawing.Size(46, 16)
        Me.lblInch2.TabIndex = 58
        Me.lblInch2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInch1
        '
        Me.lblInch1.BackColor = System.Drawing.SystemColors.Control
        Me.lblInch1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInch1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInch1.Location = New System.Drawing.Point(230, 18)
        Me.lblInch1.Name = "lblInch1"
        Me.lblInch1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInch1.Size = New System.Drawing.Size(46, 16)
        Me.lblInch1.TabIndex = 57
        Me.lblInch1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblThroatToSternal
        '
        Me.lblThroatToSternal.BackColor = System.Drawing.SystemColors.Control
        Me.lblThroatToSternal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblThroatToSternal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblThroatToSternal.Location = New System.Drawing.Point(10, 100)
        Me.lblThroatToSternal.Name = "lblThroatToSternal"
        Me.lblThroatToSternal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblThroatToSternal.Size = New System.Drawing.Size(146, 16)
        Me.lblThroatToSternal.TabIndex = 56
        Me.lblThroatToSternal.Text = "Throat desired length"
        '
        'lblCircOfNeck
        '
        Me.lblCircOfNeck.BackColor = System.Drawing.SystemColors.Control
        Me.lblCircOfNeck.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCircOfNeck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCircOfNeck.Location = New System.Drawing.Point(10, 80)
        Me.lblCircOfNeck.Name = "lblCircOfNeck"
        Me.lblCircOfNeck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCircOfNeck.Size = New System.Drawing.Size(151, 16)
        Me.lblCircOfNeck.TabIndex = 55
        Me.lblCircOfNeck.Text = "Neck Circ"
        '
        'lblCircChinAngle
        '
        Me.lblCircChinAngle.BackColor = System.Drawing.SystemColors.Control
        Me.lblCircChinAngle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCircChinAngle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCircChinAngle.Location = New System.Drawing.Point(10, 60)
        Me.lblCircChinAngle.Name = "lblCircChinAngle"
        Me.lblCircChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCircChinAngle.Size = New System.Drawing.Size(146, 14)
        Me.lblCircChinAngle.TabIndex = 54
        Me.lblCircChinAngle.Text = "Circ of Head at Chin Angle"
        '
        'lblCircEyeBrow
        '
        Me.lblCircEyeBrow.BackColor = System.Drawing.SystemColors.Control
        Me.lblCircEyeBrow.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCircEyeBrow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCircEyeBrow.Location = New System.Drawing.Point(10, 40)
        Me.lblCircEyeBrow.Name = "lblCircEyeBrow"
        Me.lblCircEyeBrow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCircEyeBrow.Size = New System.Drawing.Size(121, 16)
        Me.lblCircEyeBrow.TabIndex = 53
        Me.lblCircEyeBrow.Text = "Circ above Eyebrow"
        '
        'lblChinToMouth
        '
        Me.lblChinToMouth.BackColor = System.Drawing.SystemColors.Control
        Me.lblChinToMouth.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblChinToMouth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblChinToMouth.Location = New System.Drawing.Point(10, 20)
        Me.lblChinToMouth.Name = "lblChinToMouth"
        Me.lblChinToMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblChinToMouth.Size = New System.Drawing.Size(81, 16)
        Me.lblChinToMouth.TabIndex = 52
        Me.lblChinToMouth.Text = "Chin to Mouth"
        '
        'frmModifications
        '
        Me.frmModifications.BackColor = System.Drawing.SystemColors.Control
        Me.frmModifications.Controls.Add(Me.chkVelcro)
        Me.frmModifications.Controls.Add(Me.chkEyes)
        Me.frmModifications.Controls.Add(Me.chkRightEarClosed)
        Me.frmModifications.Controls.Add(Me.chkLeftEarClosed)
        Me.frmModifications.Controls.Add(Me.chkEarSize)
        Me.frmModifications.Controls.Add(Me.chkLipStrap)
        Me.frmModifications.Controls.Add(Me.chkOpenHeadMask)
        Me.frmModifications.Controls.Add(Me.chkNeckElastic)
        Me.frmModifications.Controls.Add(Me.chkLining)
        Me.frmModifications.Controls.Add(Me.chkLeftEyeFlap)
        Me.frmModifications.Controls.Add(Me.chkRightEyeFlap)
        Me.frmModifications.Controls.Add(Me.chkZipper)
        Me.frmModifications.Controls.Add(Me.chkNoseCovering)
        Me.frmModifications.Controls.Add(Me.chkLipCovering)
        Me.frmModifications.Controls.Add(Me.chkRightEarFlap)
        Me.frmModifications.Controls.Add(Me.chkLeftEarFlap)
        Me.frmModifications.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmModifications.Location = New System.Drawing.Point(446, 188)
        Me.frmModifications.Name = "frmModifications"
        Me.frmModifications.Padding = New System.Windows.Forms.Padding(0)
        Me.frmModifications.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmModifications.Size = New System.Drawing.Size(314, 188)
        Me.frmModifications.TabIndex = 44
        Me.frmModifications.TabStop = False
        Me.frmModifications.Text = "Modifications"
        '
        'chkVelcro
        '
        Me.chkVelcro.BackColor = System.Drawing.SystemColors.Control
        Me.chkVelcro.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVelcro.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVelcro.Location = New System.Drawing.Point(130, 160)
        Me.chkVelcro.Name = "chkVelcro"
        Me.chkVelcro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVelcro.Size = New System.Drawing.Size(130, 16)
        Me.chkVelcro.TabIndex = 124
        Me.chkVelcro.Text = "2"" Wide Velcro"
        Me.chkVelcro.UseVisualStyleBackColor = False
        '
        'chkEyes
        '
        Me.chkEyes.BackColor = System.Drawing.SystemColors.Control
        Me.chkEyes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEyes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEyes.Location = New System.Drawing.Point(130, 99)
        Me.chkEyes.Name = "chkEyes"
        Me.chkEyes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEyes.Size = New System.Drawing.Size(106, 16)
        Me.chkEyes.TabIndex = 20
        Me.chkEyes.Text = "Include Eyes"
        Me.chkEyes.UseVisualStyleBackColor = False
        '
        'chkRightEarClosed
        '
        Me.chkRightEarClosed.BackColor = System.Drawing.SystemColors.Control
        Me.chkRightEarClosed.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRightEarClosed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRightEarClosed.Location = New System.Drawing.Point(130, 58)
        Me.chkRightEarClosed.Name = "chkRightEarClosed"
        Me.chkRightEarClosed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRightEarClosed.Size = New System.Drawing.Size(132, 19)
        Me.chkRightEarClosed.TabIndex = 13
        Me.chkRightEarClosed.Text = "Right Ear Closed"
        Me.chkRightEarClosed.UseVisualStyleBackColor = False
        '
        'chkLeftEarClosed
        '
        Me.chkLeftEarClosed.BackColor = System.Drawing.SystemColors.Control
        Me.chkLeftEarClosed.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLeftEarClosed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLeftEarClosed.Location = New System.Drawing.Point(10, 58)
        Me.chkLeftEarClosed.Name = "chkLeftEarClosed"
        Me.chkLeftEarClosed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLeftEarClosed.Size = New System.Drawing.Size(114, 16)
        Me.chkLeftEarClosed.TabIndex = 12
        Me.chkLeftEarClosed.Text = "Left Ear Closed"
        Me.chkLeftEarClosed.UseVisualStyleBackColor = False
        '
        'chkEarSize
        '
        Me.chkEarSize.BackColor = System.Drawing.SystemColors.Control
        Me.chkEarSize.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEarSize.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEarSize.Location = New System.Drawing.Point(10, 79)
        Me.chkEarSize.Name = "chkEarSize"
        Me.chkEarSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEarSize.Size = New System.Drawing.Size(71, 16)
        Me.chkEarSize.TabIndex = 14
        Me.chkEarSize.Text = "Ear Size"
        Me.chkEarSize.UseVisualStyleBackColor = False
        '
        'chkLipStrap
        '
        Me.chkLipStrap.BackColor = System.Drawing.SystemColors.Control
        Me.chkLipStrap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLipStrap.Enabled = False
        Me.chkLipStrap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLipStrap.Location = New System.Drawing.Point(10, 99)
        Me.chkLipStrap.Name = "chkLipStrap"
        Me.chkLipStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLipStrap.Size = New System.Drawing.Size(111, 16)
        Me.chkLipStrap.TabIndex = 15
        Me.chkLipStrap.Text = "Lip Strap"
        Me.chkLipStrap.UseVisualStyleBackColor = False
        '
        'chkOpenHeadMask
        '
        Me.chkOpenHeadMask.BackColor = System.Drawing.SystemColors.Control
        Me.chkOpenHeadMask.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOpenHeadMask.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOpenHeadMask.Location = New System.Drawing.Point(10, 140)
        Me.chkOpenHeadMask.Name = "chkOpenHeadMask"
        Me.chkOpenHeadMask.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOpenHeadMask.Size = New System.Drawing.Size(114, 18)
        Me.chkOpenHeadMask.TabIndex = 18
        Me.chkOpenHeadMask.Text = "Open Head Mask"
        Me.chkOpenHeadMask.UseVisualStyleBackColor = False
        '
        'chkNeckElastic
        '
        Me.chkNeckElastic.BackColor = System.Drawing.SystemColors.Control
        Me.chkNeckElastic.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkNeckElastic.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkNeckElastic.Location = New System.Drawing.Point(130, 79)
        Me.chkNeckElastic.Name = "chkNeckElastic"
        Me.chkNeckElastic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkNeckElastic.Size = New System.Drawing.Size(121, 16)
        Me.chkNeckElastic.TabIndex = 19
        Me.chkNeckElastic.Text = "Neck Elastic"
        Me.chkNeckElastic.UseVisualStyleBackColor = False
        '
        'chkLining
        '
        Me.chkLining.BackColor = System.Drawing.SystemColors.Control
        Me.chkLining.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLining.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLining.Location = New System.Drawing.Point(130, 119)
        Me.chkLining.Name = "chkLining"
        Me.chkLining.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLining.Size = New System.Drawing.Size(61, 19)
        Me.chkLining.TabIndex = 22
        Me.chkLining.Text = "Lining"
        Me.chkLining.UseVisualStyleBackColor = False
        '
        'chkLeftEyeFlap
        '
        Me.chkLeftEyeFlap.BackColor = System.Drawing.SystemColors.Control
        Me.chkLeftEyeFlap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLeftEyeFlap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLeftEyeFlap.Location = New System.Drawing.Point(10, 18)
        Me.chkLeftEyeFlap.Name = "chkLeftEyeFlap"
        Me.chkLeftEyeFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLeftEyeFlap.Size = New System.Drawing.Size(106, 16)
        Me.chkLeftEyeFlap.TabIndex = 8
        Me.chkLeftEyeFlap.Text = "Left Eye Flap"
        Me.chkLeftEyeFlap.UseVisualStyleBackColor = False
        '
        'chkRightEyeFlap
        '
        Me.chkRightEyeFlap.BackColor = System.Drawing.SystemColors.Control
        Me.chkRightEyeFlap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRightEyeFlap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRightEyeFlap.Location = New System.Drawing.Point(130, 18)
        Me.chkRightEyeFlap.Name = "chkRightEyeFlap"
        Me.chkRightEyeFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRightEyeFlap.Size = New System.Drawing.Size(111, 19)
        Me.chkRightEyeFlap.TabIndex = 9
        Me.chkRightEyeFlap.Text = "Right Eye Flap"
        Me.chkRightEyeFlap.UseVisualStyleBackColor = False
        '
        'chkZipper
        '
        Me.chkZipper.BackColor = System.Drawing.SystemColors.Control
        Me.chkZipper.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkZipper.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkZipper.Location = New System.Drawing.Point(130, 140)
        Me.chkZipper.Name = "chkZipper"
        Me.chkZipper.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkZipper.Size = New System.Drawing.Size(61, 16)
        Me.chkZipper.TabIndex = 21
        Me.chkZipper.Text = "Zipper"
        Me.chkZipper.UseVisualStyleBackColor = False
        '
        'chkNoseCovering
        '
        Me.chkNoseCovering.BackColor = System.Drawing.SystemColors.Control
        Me.chkNoseCovering.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkNoseCovering.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkNoseCovering.Location = New System.Drawing.Point(10, 160)
        Me.chkNoseCovering.Name = "chkNoseCovering"
        Me.chkNoseCovering.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkNoseCovering.Size = New System.Drawing.Size(116, 23)
        Me.chkNoseCovering.TabIndex = 17
        Me.chkNoseCovering.Text = "Nose Covering"
        Me.chkNoseCovering.UseVisualStyleBackColor = False
        Me.chkNoseCovering.Visible = False
        '
        'chkLipCovering
        '
        Me.chkLipCovering.BackColor = System.Drawing.SystemColors.Control
        Me.chkLipCovering.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLipCovering.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLipCovering.Location = New System.Drawing.Point(10, 119)
        Me.chkLipCovering.Name = "chkLipCovering"
        Me.chkLipCovering.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLipCovering.Size = New System.Drawing.Size(111, 19)
        Me.chkLipCovering.TabIndex = 16
        Me.chkLipCovering.Text = "Lip Covering"
        Me.chkLipCovering.UseVisualStyleBackColor = False
        '
        'chkRightEarFlap
        '
        Me.chkRightEarFlap.BackColor = System.Drawing.SystemColors.Control
        Me.chkRightEarFlap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRightEarFlap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRightEarFlap.Location = New System.Drawing.Point(130, 38)
        Me.chkRightEarFlap.Name = "chkRightEarFlap"
        Me.chkRightEarFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRightEarFlap.Size = New System.Drawing.Size(106, 18)
        Me.chkRightEarFlap.TabIndex = 11
        Me.chkRightEarFlap.Text = "Right Ear Flap"
        Me.chkRightEarFlap.UseVisualStyleBackColor = False
        '
        'chkLeftEarFlap
        '
        Me.chkLeftEarFlap.BackColor = System.Drawing.SystemColors.Control
        Me.chkLeftEarFlap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLeftEarFlap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLeftEarFlap.Location = New System.Drawing.Point(10, 38)
        Me.chkLeftEarFlap.Name = "chkLeftEarFlap"
        Me.chkLeftEarFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLeftEarFlap.Size = New System.Drawing.Size(126, 16)
        Me.chkLeftEarFlap.TabIndex = 10
        Me.chkLeftEarFlap.Text = "Left Ear Flap"
        Me.chkLeftEarFlap.UseVisualStyleBackColor = False
        '
        'frmDesignChoices
        '
        Me.frmDesignChoices.BackColor = System.Drawing.SystemColors.Control
        Me.frmDesignChoices.Controls.Add(Me.optContouredChinCollar)
        Me.frmDesignChoices.Controls.Add(Me.optChinCollar)
        Me.frmDesignChoices.Controls.Add(Me.optModifiedChinStrap)
        Me.frmDesignChoices.Controls.Add(Me.optChinStrap)
        Me.frmDesignChoices.Controls.Add(Me.optHeadBand)
        Me.frmDesignChoices.Controls.Add(Me.optOpenFaceMask)
        Me.frmDesignChoices.Controls.Add(Me.optFaceMask)
        Me.frmDesignChoices.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmDesignChoices.Location = New System.Drawing.Point(261, 188)
        Me.frmDesignChoices.Name = "frmDesignChoices"
        Me.frmDesignChoices.Padding = New System.Windows.Forms.Padding(0)
        Me.frmDesignChoices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmDesignChoices.Size = New System.Drawing.Size(179, 188)
        Me.frmDesignChoices.TabIndex = 43
        Me.frmDesignChoices.TabStop = False
        Me.frmDesignChoices.Text = "Design Choices"
        '
        'optContouredChinCollar
        '
        Me.optContouredChinCollar.BackColor = System.Drawing.SystemColors.Control
        Me.optContouredChinCollar.Cursor = System.Windows.Forms.Cursors.Default
        Me.optContouredChinCollar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optContouredChinCollar.Location = New System.Drawing.Point(10, 120)
        Me.optContouredChinCollar.Name = "optContouredChinCollar"
        Me.optContouredChinCollar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optContouredChinCollar.Size = New System.Drawing.Size(151, 16)
        Me.optContouredChinCollar.TabIndex = 6
        Me.optContouredChinCollar.TabStop = True
        Me.optContouredChinCollar.Text = "Contoured Chin Collar"
        Me.optContouredChinCollar.UseVisualStyleBackColor = False
        '
        'optChinCollar
        '
        Me.optChinCollar.BackColor = System.Drawing.SystemColors.Control
        Me.optChinCollar.Cursor = System.Windows.Forms.Cursors.Default
        Me.optChinCollar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optChinCollar.Location = New System.Drawing.Point(10, 100)
        Me.optChinCollar.Name = "optChinCollar"
        Me.optChinCollar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optChinCollar.Size = New System.Drawing.Size(151, 16)
        Me.optChinCollar.TabIndex = 5
        Me.optChinCollar.TabStop = True
        Me.optChinCollar.Text = "Chin Collar"
        Me.optChinCollar.UseVisualStyleBackColor = False
        '
        'optModifiedChinStrap
        '
        Me.optModifiedChinStrap.BackColor = System.Drawing.SystemColors.Control
        Me.optModifiedChinStrap.Cursor = System.Windows.Forms.Cursors.Default
        Me.optModifiedChinStrap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optModifiedChinStrap.Location = New System.Drawing.Point(10, 80)
        Me.optModifiedChinStrap.Name = "optModifiedChinStrap"
        Me.optModifiedChinStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optModifiedChinStrap.Size = New System.Drawing.Size(151, 16)
        Me.optModifiedChinStrap.TabIndex = 4
        Me.optModifiedChinStrap.TabStop = True
        Me.optModifiedChinStrap.Text = "Modified Chin Strap"
        Me.optModifiedChinStrap.UseVisualStyleBackColor = False
        '
        'optChinStrap
        '
        Me.optChinStrap.BackColor = System.Drawing.SystemColors.Control
        Me.optChinStrap.Cursor = System.Windows.Forms.Cursors.Default
        Me.optChinStrap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optChinStrap.Location = New System.Drawing.Point(10, 60)
        Me.optChinStrap.Name = "optChinStrap"
        Me.optChinStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optChinStrap.Size = New System.Drawing.Size(151, 16)
        Me.optChinStrap.TabIndex = 3
        Me.optChinStrap.TabStop = True
        Me.optChinStrap.Text = "Chin Strap"
        Me.optChinStrap.UseVisualStyleBackColor = False
        '
        'optHeadBand
        '
        Me.optHeadBand.BackColor = System.Drawing.SystemColors.Control
        Me.optHeadBand.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHeadBand.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHeadBand.Location = New System.Drawing.Point(10, 140)
        Me.optHeadBand.Name = "optHeadBand"
        Me.optHeadBand.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHeadBand.Size = New System.Drawing.Size(151, 16)
        Me.optHeadBand.TabIndex = 7
        Me.optHeadBand.TabStop = True
        Me.optHeadBand.Text = "Head Band"
        Me.optHeadBand.UseVisualStyleBackColor = False
        '
        'optOpenFaceMask
        '
        Me.optOpenFaceMask.BackColor = System.Drawing.SystemColors.Control
        Me.optOpenFaceMask.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOpenFaceMask.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOpenFaceMask.Location = New System.Drawing.Point(10, 40)
        Me.optOpenFaceMask.Name = "optOpenFaceMask"
        Me.optOpenFaceMask.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOpenFaceMask.Size = New System.Drawing.Size(151, 16)
        Me.optOpenFaceMask.TabIndex = 2
        Me.optOpenFaceMask.TabStop = True
        Me.optOpenFaceMask.Text = "Open Face Mask"
        Me.optOpenFaceMask.UseVisualStyleBackColor = False
        '
        'optFaceMask
        '
        Me.optFaceMask.BackColor = System.Drawing.SystemColors.Control
        Me.optFaceMask.Checked = True
        Me.optFaceMask.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFaceMask.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFaceMask.Location = New System.Drawing.Point(10, 20)
        Me.optFaceMask.Name = "optFaceMask"
        Me.optFaceMask.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFaceMask.Size = New System.Drawing.Size(151, 16)
        Me.optFaceMask.TabIndex = 1
        Me.optFaceMask.TabStop = True
        Me.optFaceMask.Text = "Face Mask"
        Me.optFaceMask.UseVisualStyleBackColor = False
        '
        'frmPatientDetails
        '
        Me.frmPatientDetails.BackColor = System.Drawing.SystemColors.Control
        Me.frmPatientDetails.Controls.Add(Me.Label12)
        Me.frmPatientDetails.Controls.Add(Me.DateTimePicker1)
        Me.frmPatientDetails.Controls.Add(Me.txtSex)
        Me.frmPatientDetails.Controls.Add(Me.txtAge)
        Me.frmPatientDetails.Controls.Add(Me.txtFileNo)
        Me.frmPatientDetails.Controls.Add(Me.txtDiagnosis)
        Me.frmPatientDetails.Controls.Add(Me.txtPatientName)
        Me.frmPatientDetails.Controls.Add(Me.lblSex)
        Me.frmPatientDetails.Controls.Add(Me.lblAge)
        Me.frmPatientDetails.Controls.Add(Me.lblFileNo)
        Me.frmPatientDetails.Controls.Add(Me.lblDiagnosis)
        Me.frmPatientDetails.Controls.Add(Me.lblPatientName)
        Me.frmPatientDetails.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmPatientDetails.Location = New System.Drawing.Point(12, 64)
        Me.frmPatientDetails.Name = "frmPatientDetails"
        Me.frmPatientDetails.Padding = New System.Windows.Forms.Padding(0)
        Me.frmPatientDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmPatientDetails.Size = New System.Drawing.Size(380, 118)
        Me.frmPatientDetails.TabIndex = 0
        Me.frmPatientDetails.TabStop = False
        Me.frmPatientDetails.Text = "Patient Details"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(251, 27)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(33, 15)
        Me.Label12.TabIndex = 129
        Me.Label12.Text = "DOB:"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd-MM-yyyy"
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(288, 24)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(87, 20)
        Me.DateTimePicker1.TabIndex = 128
        '
        'txtSex
        '
        Me.txtSex.AcceptsReturn = True
        Me.txtSex.BackColor = System.Drawing.SystemColors.Window
        Me.txtSex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSex.Location = New System.Drawing.Point(288, 82)
        Me.txtSex.MaxLength = 0
        Me.txtSex.Name = "txtSex"
        Me.txtSex.ReadOnly = True
        Me.txtSex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSex.Size = New System.Drawing.Size(55, 20)
        Me.txtSex.TabIndex = 42
        '
        'txtAge
        '
        Me.txtAge.AcceptsReturn = True
        Me.txtAge.BackColor = System.Drawing.SystemColors.Window
        Me.txtAge.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAge.Enabled = False
        Me.txtAge.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAge.Location = New System.Drawing.Point(288, 53)
        Me.txtAge.MaxLength = 0
        Me.txtAge.Name = "txtAge"
        Me.txtAge.ReadOnly = True
        Me.txtAge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAge.Size = New System.Drawing.Size(55, 20)
        Me.txtAge.TabIndex = 39
        '
        'txtFileNo
        '
        Me.txtFileNo.AcceptsReturn = True
        Me.txtFileNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtFileNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFileNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFileNo.Location = New System.Drawing.Point(70, 22)
        Me.txtFileNo.MaxLength = 0
        Me.txtFileNo.Name = "txtFileNo"
        Me.txtFileNo.ReadOnly = True
        Me.txtFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFileNo.Size = New System.Drawing.Size(78, 20)
        Me.txtFileNo.TabIndex = 38
        '
        'txtDiagnosis
        '
        Me.txtDiagnosis.AcceptsReturn = True
        Me.txtDiagnosis.BackColor = System.Drawing.SystemColors.Window
        Me.txtDiagnosis.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDiagnosis.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDiagnosis.Location = New System.Drawing.Point(70, 82)
        Me.txtDiagnosis.MaxLength = 0
        Me.txtDiagnosis.Name = "txtDiagnosis"
        Me.txtDiagnosis.ReadOnly = True
        Me.txtDiagnosis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDiagnosis.Size = New System.Drawing.Size(176, 20)
        Me.txtDiagnosis.TabIndex = 40
        '
        'txtPatientName
        '
        Me.txtPatientName.AcceptsReturn = True
        Me.txtPatientName.BackColor = System.Drawing.SystemColors.Window
        Me.txtPatientName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPatientName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPatientName.Location = New System.Drawing.Point(70, 51)
        Me.txtPatientName.MaxLength = 0
        Me.txtPatientName.Name = "txtPatientName"
        Me.txtPatientName.ReadOnly = True
        Me.txtPatientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPatientName.Size = New System.Drawing.Size(176, 20)
        Me.txtPatientName.TabIndex = 37
        '
        'lblSex
        '
        Me.lblSex.BackColor = System.Drawing.SystemColors.Control
        Me.lblSex.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSex.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSex.Location = New System.Drawing.Point(254, 84)
        Me.lblSex.Name = "lblSex"
        Me.lblSex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSex.Size = New System.Drawing.Size(35, 16)
        Me.lblSex.TabIndex = 50
        Me.lblSex.Text = "Sex:"
        '
        'lblAge
        '
        Me.lblAge.BackColor = System.Drawing.SystemColors.Control
        Me.lblAge.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAge.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAge.Location = New System.Drawing.Point(254, 53)
        Me.lblAge.Name = "lblAge"
        Me.lblAge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAge.Size = New System.Drawing.Size(35, 16)
        Me.lblAge.TabIndex = 48
        Me.lblAge.Text = "Age:"
        '
        'lblFileNo
        '
        Me.lblFileNo.BackColor = System.Drawing.SystemColors.Control
        Me.lblFileNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFileNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFileNo.Location = New System.Drawing.Point(18, 24)
        Me.lblFileNo.Name = "lblFileNo"
        Me.lblFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFileNo.Size = New System.Drawing.Size(46, 16)
        Me.lblFileNo.TabIndex = 47
        Me.lblFileNo.Text = "File No:"
        Me.lblFileNo.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDiagnosis
        '
        Me.lblDiagnosis.BackColor = System.Drawing.SystemColors.Control
        Me.lblDiagnosis.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDiagnosis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDiagnosis.Location = New System.Drawing.Point(1, 82)
        Me.lblDiagnosis.Name = "lblDiagnosis"
        Me.lblDiagnosis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDiagnosis.Size = New System.Drawing.Size(64, 18)
        Me.lblDiagnosis.TabIndex = 46
        Me.lblDiagnosis.Text = "Diagnosis:"
        Me.lblDiagnosis.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPatientName
        '
        Me.lblPatientName.BackColor = System.Drawing.SystemColors.Control
        Me.lblPatientName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPatientName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPatientName.Location = New System.Drawing.Point(20, 51)
        Me.lblPatientName.Name = "lblPatientName"
        Me.lblPatientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPatientName.Size = New System.Drawing.Size(43, 18)
        Me.lblPatientName.TabIndex = 45
        Me.lblPatientName.Text = "Name:"
        Me.lblPatientName.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtUnits
        '
        Me.txtUnits.AcceptsReturn = True
        Me.txtUnits.BackColor = System.Drawing.SystemColors.Window
        Me.txtUnits.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUnits.Location = New System.Drawing.Point(86, 82)
        Me.txtUnits.MaxLength = 0
        Me.txtUnits.Name = "txtUnits"
        Me.txtUnits.ReadOnly = True
        Me.txtUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUnits.Size = New System.Drawing.Size(78, 20)
        Me.txtUnits.TabIndex = 41
        Me.txtUnits.Visible = False
        '
        'lblUnits
        '
        Me.lblUnits.BackColor = System.Drawing.SystemColors.Control
        Me.lblUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUnits.Location = New System.Drawing.Point(48, 84)
        Me.lblUnits.Name = "lblUnits"
        Me.lblUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUnits.Size = New System.Drawing.Size(46, 16)
        Me.lblUnits.TabIndex = 49
        Me.lblUnits.Text = "Units:"
        Me.lblUnits.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.txtUnits)
        Me.GroupBox1.Controls.Add(Me.txtTempDate)
        Me.GroupBox1.Controls.Add(Me.txtDesigner)
        Me.GroupBox1.Controls.Add(Me.txtWorkOrder1)
        Me.GroupBox1.Controls.Add(Me.lblUnits)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(398, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(362, 118)
        Me.GroupBox1.TabIndex = 124
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Work Order Details"
        '
        'txtTempDate
        '
        Me.txtTempDate.AcceptsReturn = True
        Me.txtTempDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtTempDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTempDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTempDate.Location = New System.Drawing.Point(277, 20)
        Me.txtTempDate.MaxLength = 0
        Me.txtTempDate.Name = "txtTempDate"
        Me.txtTempDate.ReadOnly = True
        Me.txtTempDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTempDate.Size = New System.Drawing.Size(78, 20)
        Me.txtTempDate.TabIndex = 38
        '
        'txtDesigner
        '
        Me.txtDesigner.AcceptsReturn = True
        Me.txtDesigner.BackColor = System.Drawing.SystemColors.Window
        Me.txtDesigner.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesigner.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesigner.Location = New System.Drawing.Point(86, 51)
        Me.txtDesigner.MaxLength = 0
        Me.txtDesigner.Name = "txtDesigner"
        Me.txtDesigner.ReadOnly = True
        Me.txtDesigner.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesigner.Size = New System.Drawing.Size(164, 20)
        Me.txtDesigner.TabIndex = 40
        '
        'txtWorkOrder1
        '
        Me.txtWorkOrder1.AcceptsReturn = True
        Me.txtWorkOrder1.BackColor = System.Drawing.SystemColors.Window
        Me.txtWorkOrder1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWorkOrder1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWorkOrder1.Location = New System.Drawing.Point(86, 20)
        Me.txtWorkOrder1.MaxLength = 0
        Me.txtWorkOrder1.Name = "txtWorkOrder1"
        Me.txtWorkOrder1.ReadOnly = True
        Me.txtWorkOrder1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWorkOrder1.Size = New System.Drawing.Size(78, 20)
        Me.txtWorkOrder1.TabIndex = 37
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(203, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(71, 16)
        Me.Label4.TabIndex = 47
        Me.Label4.Text = "Temp Date:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(30, 53)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Designer:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(20, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(73, 16)
        Me.Label6.TabIndex = 45
        Me.Label6.Text = "WorkOrder:"
        '
        'PictureBox1
        '
        Me.PictureBox1.ImageLocation = "C:\ACAD2018_SMITH-NEPHEW\Interface\SN-LOGO.bmp"
        Me.PictureBox1.Location = New System.Drawing.Point(12, 7)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(362, 50)
        Me.PictureBox1.TabIndex = 125
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Rounded MT Bold", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(532, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(243, 37)
        Me.Label1.TabIndex = 127
        Me.Label1.Text = "head and neck"
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.White
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox2.ImageLocation = "C:\ACAD2018_SMITH-NEPHEW\Interface\HeadNeck.bmp"
        Me.PictureBox2.Location = New System.Drawing.Point(12, 194)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(229, 335)
        Me.PictureBox2.TabIndex = 126
        Me.PictureBox2.TabStop = False
        '
        'HeadNeck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(770, 563)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtUidHN)
        Me.Controls.Add(Me.txtTopRightEar)
        Me.Controls.Add(Me.txtTopLeftEar)
        Me.Controls.Add(Me.txtHeadNeckUID)
        Me.Controls.Add(Me.txtOptionChoice)
        Me.Controls.Add(Me.txtFabric)
        Me.Controls.Add(Me.txtWorkOrder)
        Me.Controls.Add(Me.txtData)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.txtMeasurements)
        Me.Controls.Add(Me.txtMouthRightY)
        Me.Controls.Add(Me.txtxyNeckTopFrontY)
        Me.Controls.Add(Me.txtxyNeckTopFrontX)
        Me.Controls.Add(Me.txtxyChinTopY)
        Me.Controls.Add(Me.txtxyChinTopX)
        Me.Controls.Add(Me.txtCSEarBotHeight)
        Me.Controls.Add(Me.txtCSEarTopHeight)
        Me.Controls.Add(Me.txtCSMouthWidth)
        Me.Controls.Add(Me.txtMouthRightX)
        Me.Controls.Add(Me.txtCSForeHead)
        Me.Controls.Add(Me.txtCSChinAngle)
        Me.Controls.Add(Me.txtCSNeckCircum)
        Me.Controls.Add(Me.txtCSChintoMouth)
        Me.Controls.Add(Me.txtNoseCoverX)
        Me.Controls.Add(Me.txtNoseCoverY)
        Me.Controls.Add(Me.txtEyeWidth)
        Me.Controls.Add(Me.txtDartEndX)
        Me.Controls.Add(Me.txtDartEndY)
        Me.Controls.Add(Me.txtDartStartY)
        Me.Controls.Add(Me.txtDartStartX)
        Me.Controls.Add(Me.txtChinLeftBotY)
        Me.Controls.Add(Me.txtChinLeftBotX)
        Me.Controls.Add(Me.txtNoseBottomY)
        Me.Controls.Add(Me.txtCircumferenceTotal)
        Me.Controls.Add(Me.txtMouthHeight)
        Me.Controls.Add(Me.txtBotOfEyeY)
        Me.Controls.Add(Me.txtBotOfEyeX)
        Me.Controls.Add(Me.txtLowerEarHeight)
        Me.Controls.Add(Me.txtfNum)
        Me.Controls.Add(Me.txtCurrTextFont)
        Me.Controls.Add(Me.txtCurrTextVertJust)
        Me.Controls.Add(Me.txtCurrTextHorizJust)
        Me.Controls.Add(Me.txtCurrTextAspect)
        Me.Controls.Add(Me.txtCurrTextHt)
        Me.Controls.Add(Me.txtCurrentLayer)
        Me.Controls.Add(Me.txtMidToEyeTop)
        Me.Controls.Add(Me.txtLipStrapWidth)
        Me.Controls.Add(Me.txtUidMPD)
        Me.Controls.Add(Me.txtRadiusNo)
        Me.Controls.Add(Me.txtDraw)
        Me.Controls.Add(Me.cmdDraw)
        Me.Controls.Add(Me.frmMeasurements)
        Me.Controls.Add(Me.frmModifications)
        Me.Controls.Add(Me.frmDesignChoices)
        Me.Controls.Add(Me.frmPatientDetails)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(162, 77)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "HeadNeck"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Head & Neck"
        Me.frmMeasurements.ResumeLayout(False)
        Me.frmMeasurements.PerformLayout()
        Me.frmModifications.ResumeLayout(False)
        Me.frmDesignChoices.ResumeLayout(False)
        Me.frmPatientDetails.ResumeLayout(False)
        Me.frmPatientDetails.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public WithEvents GroupBox1 As GroupBox
    Public WithEvents txtTempDate As TextBox
    Public WithEvents txtDesigner As TextBox
    Public WithEvents txtWorkOrder1 As TextBox
    Public WithEvents Label4 As Label
    Public WithEvents Label5 As Label
    Public WithEvents Label6 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
#End Region
End Class
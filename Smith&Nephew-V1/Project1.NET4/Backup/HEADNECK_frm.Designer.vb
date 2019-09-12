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
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(HeadNeck))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.txtUidHN = New System.Windows.Forms.TextBox
		Me.txtTopRightEar = New System.Windows.Forms.TextBox
		Me.txtTopLeftEar = New System.Windows.Forms.TextBox
		Me.txtHeadNeckUID = New System.Windows.Forms.TextBox
		Me.txtOptionChoice = New System.Windows.Forms.TextBox
		Me.txtFabric = New System.Windows.Forms.TextBox
		Me.txtWorkOrder = New System.Windows.Forms.TextBox
		Me.txtData = New System.Windows.Forms.TextBox
		Me.cmdClose = New System.Windows.Forms.Button
		Me.txtMeasurements = New System.Windows.Forms.TextBox
		Me.txtMouthRightY = New System.Windows.Forms.TextBox
		Me.txtxyNeckTopFrontY = New System.Windows.Forms.TextBox
		Me.txtxyNeckTopFrontX = New System.Windows.Forms.TextBox
		Me.txtxyChinTopY = New System.Windows.Forms.TextBox
		Me.txtxyChinTopX = New System.Windows.Forms.TextBox
		Me.txtCSEarBotHeight = New System.Windows.Forms.TextBox
		Me.txtCSEarTopHeight = New System.Windows.Forms.TextBox
		Me.txtCSMouthWidth = New System.Windows.Forms.TextBox
		Me.txtMouthRightX = New System.Windows.Forms.TextBox
		Me.txtCSForeHead = New System.Windows.Forms.TextBox
		Me.txtCSChinAngle = New System.Windows.Forms.TextBox
		Me.txtCSNeckCircum = New System.Windows.Forms.TextBox
		Me.txtCSChintoMouth = New System.Windows.Forms.TextBox
		Me.txtNoseCoverX = New System.Windows.Forms.TextBox
		Me.txtNoseCoverY = New System.Windows.Forms.TextBox
		Me.txtEyeWidth = New System.Windows.Forms.TextBox
		Me.txtDartEndX = New System.Windows.Forms.TextBox
		Me.txtDartEndY = New System.Windows.Forms.TextBox
		Me.txtDartStartY = New System.Windows.Forms.TextBox
		Me.txtDartStartX = New System.Windows.Forms.TextBox
		Me.txtChinLeftBotY = New System.Windows.Forms.TextBox
		Me.txtChinLeftBotX = New System.Windows.Forms.TextBox
		Me.txtNoseBottomY = New System.Windows.Forms.TextBox
		Me.txtCircumferenceTotal = New System.Windows.Forms.TextBox
		Me.txtMouthHeight = New System.Windows.Forms.TextBox
		Me.txtBotOfEyeY = New System.Windows.Forms.TextBox
		Me.txtBotOfEyeX = New System.Windows.Forms.TextBox
		Me.txtLowerEarHeight = New System.Windows.Forms.TextBox
		Me.txtfNum = New System.Windows.Forms.TextBox
		Me.txtCurrTextFont = New System.Windows.Forms.TextBox
		Me.txtCurrTextVertJust = New System.Windows.Forms.TextBox
		Me.txtCurrTextHorizJust = New System.Windows.Forms.TextBox
		Me.txtCurrTextAspect = New System.Windows.Forms.TextBox
		Me.txtCurrTextHt = New System.Windows.Forms.TextBox
		Me.txtCurrentLayer = New System.Windows.Forms.TextBox
		Me.txtMidToEyeTop = New System.Windows.Forms.TextBox
		Me.txtLipStrapWidth = New System.Windows.Forms.TextBox
		Me.txtUidMPD = New System.Windows.Forms.TextBox
		Me.txtRadiusNo = New System.Windows.Forms.TextBox
		Me.txtDraw = New System.Windows.Forms.TextBox
		Me.cmdDraw = New System.Windows.Forms.Button
		Me.frmMeasurements = New System.Windows.Forms.GroupBox
		Me.txtChinCollarMin = New System.Windows.Forms.TextBox
		Me.txtRightEarLength = New System.Windows.Forms.TextBox
		Me.txtLeftEarLength = New System.Windows.Forms.TextBox
		Me.txtHeadBandDepth = New System.Windows.Forms.TextBox
		Me.txtLengthOfNose = New System.Windows.Forms.TextBox
		Me.txtTipOfNose = New System.Windows.Forms.TextBox
		Me.cboFabric = New System.Windows.Forms.ComboBox
		Me.txtThroatToSternal = New System.Windows.Forms.TextBox
		Me.txtCircOfNeck = New System.Windows.Forms.TextBox
		Me.txtCircChinAngle = New System.Windows.Forms.TextBox
		Me.txtCircEyeBrow = New System.Windows.Forms.TextBox
		Me.txtChinToMouth = New System.Windows.Forms.TextBox
		Me.lblInch11 = New System.Windows.Forms.Label
		Me.lblChinCollarMin = New System.Windows.Forms.Label
		Me.lblRightEarLen = New System.Windows.Forms.Label
		Me.lblLeftEarLen = New System.Windows.Forms.Label
		Me.lblInch10 = New System.Windows.Forms.Label
		Me.lblInch9 = New System.Windows.Forms.Label
		Me.lblInch8 = New System.Windows.Forms.Label
		Me.lblHeadBandDepth = New System.Windows.Forms.Label
		Me.lblFabric = New System.Windows.Forms.Label
		Me.lblInch7 = New System.Windows.Forms.Label
		Me.lblInch6 = New System.Windows.Forms.Label
		Me.LblLengthOfNose = New System.Windows.Forms.Label
		Me.lblAcrossTipOfNose = New System.Windows.Forms.Label
		Me.lblInch5 = New System.Windows.Forms.Label
		Me.lblInch4 = New System.Windows.Forms.Label
		Me.lblInch3 = New System.Windows.Forms.Label
		Me.lblInch2 = New System.Windows.Forms.Label
		Me.lblInch1 = New System.Windows.Forms.Label
		Me.lblThroatToSternal = New System.Windows.Forms.Label
		Me.lblCircOfNeck = New System.Windows.Forms.Label
		Me.lblCircChinAngle = New System.Windows.Forms.Label
		Me.lblCircEyeBrow = New System.Windows.Forms.Label
		Me.lblChinToMouth = New System.Windows.Forms.Label
		Me.frmModifications = New System.Windows.Forms.GroupBox
		Me.chkVelcro = New System.Windows.Forms.CheckBox
		Me.chkEyes = New System.Windows.Forms.CheckBox
		Me.chkRightEarClosed = New System.Windows.Forms.CheckBox
		Me.chkLeftEarClosed = New System.Windows.Forms.CheckBox
		Me.chkEarSize = New System.Windows.Forms.CheckBox
		Me.chkLipStrap = New System.Windows.Forms.CheckBox
		Me.chkOpenHeadMask = New System.Windows.Forms.CheckBox
		Me.chkNeckElastic = New System.Windows.Forms.CheckBox
		Me.chkLining = New System.Windows.Forms.CheckBox
		Me.chkLeftEyeFlap = New System.Windows.Forms.CheckBox
		Me.chkRightEyeFlap = New System.Windows.Forms.CheckBox
		Me.chkZipper = New System.Windows.Forms.CheckBox
		Me.chkNoseCovering = New System.Windows.Forms.CheckBox
		Me.chkLipCovering = New System.Windows.Forms.CheckBox
		Me.chkRightEarFlap = New System.Windows.Forms.CheckBox
		Me.chkLeftEarFlap = New System.Windows.Forms.CheckBox
		Me.frmDesignChoices = New System.Windows.Forms.GroupBox
		Me.optContouredChinCollar = New System.Windows.Forms.RadioButton
		Me.optChinCollar = New System.Windows.Forms.RadioButton
		Me.optModifiedChinStrap = New System.Windows.Forms.RadioButton
		Me.optChinStrap = New System.Windows.Forms.RadioButton
		Me.optHeadBand = New System.Windows.Forms.RadioButton
		Me.optOpenFaceMask = New System.Windows.Forms.RadioButton
		Me.optFaceMask = New System.Windows.Forms.RadioButton
		Me.frmPatientDetails = New System.Windows.Forms.GroupBox
		Me.txtSex = New System.Windows.Forms.TextBox
		Me.txtUnits = New System.Windows.Forms.TextBox
		Me.txtAge = New System.Windows.Forms.TextBox
		Me.txtFileNo = New System.Windows.Forms.TextBox
		Me.txtDiagnosis = New System.Windows.Forms.TextBox
		Me.txtPatientName = New System.Windows.Forms.TextBox
		Me.lblSex = New System.Windows.Forms.Label
		Me.lblUnits = New System.Windows.Forms.Label
		Me.lblAge = New System.Windows.Forms.Label
		Me.lblFileNo = New System.Windows.Forms.Label
		Me.lblDiagnosis = New System.Windows.Forms.Label
		Me.lblPatientName = New System.Windows.Forms.Label
		Me.frmMeasurements.SuspendLayout()
		Me.frmModifications.SuspendLayout()
		Me.frmDesignChoices.SuspendLayout()
		Me.frmPatientDetails.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Head & Neck"
		Me.ClientSize = New System.Drawing.Size(460, 476)
		Me.Location = New System.Drawing.Point(162, 77)
		Me.ForeColor = System.Drawing.SystemColors.WindowText
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "HeadNeck"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.txtUidHN.AutoSize = False
		Me.txtUidHN.Size = New System.Drawing.Size(86, 19)
		Me.txtUidHN.Location = New System.Drawing.Point(472, 396)
		Me.txtUidHN.TabIndex = 123
		Me.txtUidHN.Text = "txtUidHN"
		Me.txtUidHN.Visible = False
		Me.txtUidHN.AcceptsReturn = True
		Me.txtUidHN.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUidHN.BackColor = System.Drawing.SystemColors.Window
		Me.txtUidHN.CausesValidation = True
		Me.txtUidHN.Enabled = True
		Me.txtUidHN.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUidHN.HideSelection = True
		Me.txtUidHN.ReadOnly = False
		Me.txtUidHN.Maxlength = 0
		Me.txtUidHN.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUidHN.MultiLine = False
		Me.txtUidHN.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUidHN.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUidHN.TabStop = True
		Me.txtUidHN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtUidHN.Name = "txtUidHN"
		Me.txtTopRightEar.AutoSize = False
		Me.txtTopRightEar.Size = New System.Drawing.Size(46, 19)
		Me.txtTopRightEar.Location = New System.Drawing.Point(570, 230)
		Me.txtTopRightEar.TabIndex = 120
		Me.txtTopRightEar.Text = "txtTopRightEar"
		Me.txtTopRightEar.Visible = False
		Me.txtTopRightEar.AcceptsReturn = True
		Me.txtTopRightEar.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTopRightEar.BackColor = System.Drawing.SystemColors.Window
		Me.txtTopRightEar.CausesValidation = True
		Me.txtTopRightEar.Enabled = True
		Me.txtTopRightEar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTopRightEar.HideSelection = True
		Me.txtTopRightEar.ReadOnly = False
		Me.txtTopRightEar.Maxlength = 0
		Me.txtTopRightEar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTopRightEar.MultiLine = False
		Me.txtTopRightEar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTopRightEar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTopRightEar.TabStop = True
		Me.txtTopRightEar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtTopRightEar.Name = "txtTopRightEar"
		Me.txtTopLeftEar.AutoSize = False
		Me.txtTopLeftEar.Size = New System.Drawing.Size(46, 19)
		Me.txtTopLeftEar.Location = New System.Drawing.Point(570, 90)
		Me.txtTopLeftEar.TabIndex = 119
		Me.txtTopLeftEar.Text = "txtTopLeftEar"
		Me.txtTopLeftEar.Visible = False
		Me.txtTopLeftEar.AcceptsReturn = True
		Me.txtTopLeftEar.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTopLeftEar.BackColor = System.Drawing.SystemColors.Window
		Me.txtTopLeftEar.CausesValidation = True
		Me.txtTopLeftEar.Enabled = True
		Me.txtTopLeftEar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTopLeftEar.HideSelection = True
		Me.txtTopLeftEar.ReadOnly = False
		Me.txtTopLeftEar.Maxlength = 0
		Me.txtTopLeftEar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTopLeftEar.MultiLine = False
		Me.txtTopLeftEar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTopLeftEar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTopLeftEar.TabStop = True
		Me.txtTopLeftEar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtTopLeftEar.Name = "txtTopLeftEar"
		Me.txtHeadNeckUID.AutoSize = False
		Me.txtHeadNeckUID.Size = New System.Drawing.Size(46, 19)
		Me.txtHeadNeckUID.Location = New System.Drawing.Point(570, 130)
		Me.txtHeadNeckUID.TabIndex = 118
		Me.txtHeadNeckUID.Text = "txtHeadNeckUID"
		Me.txtHeadNeckUID.Visible = False
		Me.txtHeadNeckUID.AcceptsReturn = True
		Me.txtHeadNeckUID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHeadNeckUID.BackColor = System.Drawing.SystemColors.Window
		Me.txtHeadNeckUID.CausesValidation = True
		Me.txtHeadNeckUID.Enabled = True
		Me.txtHeadNeckUID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHeadNeckUID.HideSelection = True
		Me.txtHeadNeckUID.ReadOnly = False
		Me.txtHeadNeckUID.Maxlength = 0
		Me.txtHeadNeckUID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHeadNeckUID.MultiLine = False
		Me.txtHeadNeckUID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHeadNeckUID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHeadNeckUID.TabStop = True
		Me.txtHeadNeckUID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtHeadNeckUID.Name = "txtHeadNeckUID"
		Me.txtOptionChoice.AutoSize = False
		Me.txtOptionChoice.Size = New System.Drawing.Size(46, 19)
		Me.txtOptionChoice.Location = New System.Drawing.Point(570, 150)
		Me.txtOptionChoice.TabIndex = 117
		Me.txtOptionChoice.Text = "txtOptionChoice"
		Me.txtOptionChoice.Visible = False
		Me.txtOptionChoice.AcceptsReturn = True
		Me.txtOptionChoice.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOptionChoice.BackColor = System.Drawing.SystemColors.Window
		Me.txtOptionChoice.CausesValidation = True
		Me.txtOptionChoice.Enabled = True
		Me.txtOptionChoice.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOptionChoice.HideSelection = True
		Me.txtOptionChoice.ReadOnly = False
		Me.txtOptionChoice.Maxlength = 0
		Me.txtOptionChoice.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOptionChoice.MultiLine = False
		Me.txtOptionChoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOptionChoice.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOptionChoice.TabStop = True
		Me.txtOptionChoice.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtOptionChoice.Name = "txtOptionChoice"
		Me.txtFabric.AutoSize = False
		Me.txtFabric.Size = New System.Drawing.Size(46, 19)
		Me.txtFabric.Location = New System.Drawing.Point(570, 270)
		Me.txtFabric.TabIndex = 116
		Me.txtFabric.Visible = False
		Me.txtFabric.AcceptsReturn = True
		Me.txtFabric.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFabric.BackColor = System.Drawing.SystemColors.Window
		Me.txtFabric.CausesValidation = True
		Me.txtFabric.Enabled = True
		Me.txtFabric.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFabric.HideSelection = True
		Me.txtFabric.ReadOnly = False
		Me.txtFabric.Maxlength = 0
		Me.txtFabric.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFabric.MultiLine = False
		Me.txtFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFabric.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFabric.TabStop = True
		Me.txtFabric.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtFabric.Name = "txtFabric"
		Me.txtWorkOrder.AutoSize = False
		Me.txtWorkOrder.Size = New System.Drawing.Size(146, 19)
		Me.txtWorkOrder.Location = New System.Drawing.Point(470, 350)
		Me.txtWorkOrder.TabIndex = 115
		Me.txtWorkOrder.Text = "txtWorkOrder"
		Me.txtWorkOrder.Visible = False
		Me.txtWorkOrder.AcceptsReturn = True
		Me.txtWorkOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWorkOrder.BackColor = System.Drawing.SystemColors.Window
		Me.txtWorkOrder.CausesValidation = True
		Me.txtWorkOrder.Enabled = True
		Me.txtWorkOrder.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWorkOrder.HideSelection = True
		Me.txtWorkOrder.ReadOnly = False
		Me.txtWorkOrder.Maxlength = 0
		Me.txtWorkOrder.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWorkOrder.MultiLine = False
		Me.txtWorkOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWorkOrder.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWorkOrder.TabStop = True
		Me.txtWorkOrder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtWorkOrder.Name = "txtWorkOrder"
		Me.txtData.AutoSize = False
		Me.txtData.Size = New System.Drawing.Size(146, 19)
		Me.txtData.Location = New System.Drawing.Point(470, 330)
		Me.txtData.Maxlength = 19
		Me.txtData.TabIndex = 86
		Me.txtData.Text = "txtData"
		Me.txtData.Visible = False
		Me.txtData.AcceptsReturn = True
		Me.txtData.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtData.BackColor = System.Drawing.SystemColors.Window
		Me.txtData.CausesValidation = True
		Me.txtData.Enabled = True
		Me.txtData.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtData.HideSelection = True
		Me.txtData.ReadOnly = False
		Me.txtData.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtData.MultiLine = False
		Me.txtData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtData.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtData.TabStop = True
		Me.txtData.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtData.Name = "txtData"
		Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClose.Text = "Close"
		Me.cmdClose.Size = New System.Drawing.Size(96, 26)
		Me.cmdClose.Location = New System.Drawing.Point(260, 440)
		Me.cmdClose.TabIndex = 36
		Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClose.CausesValidation = True
		Me.cmdClose.Enabled = True
		Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClose.TabStop = True
		Me.cmdClose.Name = "cmdClose"
		Me.txtMeasurements.AutoSize = False
		Me.txtMeasurements.Size = New System.Drawing.Size(46, 19)
		Me.txtMeasurements.Location = New System.Drawing.Point(570, 110)
		Me.txtMeasurements.Maxlength = 43
		Me.txtMeasurements.TabIndex = 111
		Me.txtMeasurements.Text = "txtMeasurements"
		Me.txtMeasurements.Visible = False
		Me.txtMeasurements.AcceptsReturn = True
		Me.txtMeasurements.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMeasurements.BackColor = System.Drawing.SystemColors.Window
		Me.txtMeasurements.CausesValidation = True
		Me.txtMeasurements.Enabled = True
		Me.txtMeasurements.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMeasurements.HideSelection = True
		Me.txtMeasurements.ReadOnly = False
		Me.txtMeasurements.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMeasurements.MultiLine = False
		Me.txtMeasurements.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMeasurements.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMeasurements.TabStop = True
		Me.txtMeasurements.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMeasurements.Name = "txtMeasurements"
		Me.txtMouthRightY.AutoSize = False
		Me.txtMouthRightY.Size = New System.Drawing.Size(46, 19)
		Me.txtMouthRightY.Location = New System.Drawing.Point(570, 210)
		Me.txtMouthRightY.TabIndex = 110
		Me.txtMouthRightY.Text = "txtMouthRightY"
		Me.txtMouthRightY.Visible = False
		Me.txtMouthRightY.AcceptsReturn = True
		Me.txtMouthRightY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMouthRightY.BackColor = System.Drawing.SystemColors.Window
		Me.txtMouthRightY.CausesValidation = True
		Me.txtMouthRightY.Enabled = True
		Me.txtMouthRightY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMouthRightY.HideSelection = True
		Me.txtMouthRightY.ReadOnly = False
		Me.txtMouthRightY.Maxlength = 0
		Me.txtMouthRightY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMouthRightY.MultiLine = False
		Me.txtMouthRightY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMouthRightY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMouthRightY.TabStop = True
		Me.txtMouthRightY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMouthRightY.Name = "txtMouthRightY"
		Me.txtxyNeckTopFrontY.AutoSize = False
		Me.txtxyNeckTopFrontY.Size = New System.Drawing.Size(46, 19)
		Me.txtxyNeckTopFrontY.Location = New System.Drawing.Point(570, 310)
		Me.txtxyNeckTopFrontY.TabIndex = 109
		Me.txtxyNeckTopFrontY.Text = "txtxyNeckTopFrontY"
		Me.txtxyNeckTopFrontY.Visible = False
		Me.txtxyNeckTopFrontY.AcceptsReturn = True
		Me.txtxyNeckTopFrontY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtxyNeckTopFrontY.BackColor = System.Drawing.SystemColors.Window
		Me.txtxyNeckTopFrontY.CausesValidation = True
		Me.txtxyNeckTopFrontY.Enabled = True
		Me.txtxyNeckTopFrontY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtxyNeckTopFrontY.HideSelection = True
		Me.txtxyNeckTopFrontY.ReadOnly = False
		Me.txtxyNeckTopFrontY.Maxlength = 0
		Me.txtxyNeckTopFrontY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtxyNeckTopFrontY.MultiLine = False
		Me.txtxyNeckTopFrontY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtxyNeckTopFrontY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtxyNeckTopFrontY.TabStop = True
		Me.txtxyNeckTopFrontY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtxyNeckTopFrontY.Name = "txtxyNeckTopFrontY"
		Me.txtxyNeckTopFrontX.AutoSize = False
		Me.txtxyNeckTopFrontX.Size = New System.Drawing.Size(46, 19)
		Me.txtxyNeckTopFrontX.Location = New System.Drawing.Point(520, 310)
		Me.txtxyNeckTopFrontX.TabIndex = 108
		Me.txtxyNeckTopFrontX.Text = "txtxyNeckTopFrontX"
		Me.txtxyNeckTopFrontX.Visible = False
		Me.txtxyNeckTopFrontX.AcceptsReturn = True
		Me.txtxyNeckTopFrontX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtxyNeckTopFrontX.BackColor = System.Drawing.SystemColors.Window
		Me.txtxyNeckTopFrontX.CausesValidation = True
		Me.txtxyNeckTopFrontX.Enabled = True
		Me.txtxyNeckTopFrontX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtxyNeckTopFrontX.HideSelection = True
		Me.txtxyNeckTopFrontX.ReadOnly = False
		Me.txtxyNeckTopFrontX.Maxlength = 0
		Me.txtxyNeckTopFrontX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtxyNeckTopFrontX.MultiLine = False
		Me.txtxyNeckTopFrontX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtxyNeckTopFrontX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtxyNeckTopFrontX.TabStop = True
		Me.txtxyNeckTopFrontX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtxyNeckTopFrontX.Name = "txtxyNeckTopFrontX"
		Me.txtxyChinTopY.AutoSize = False
		Me.txtxyChinTopY.Size = New System.Drawing.Size(46, 19)
		Me.txtxyChinTopY.Location = New System.Drawing.Point(520, 290)
		Me.txtxyChinTopY.TabIndex = 107
		Me.txtxyChinTopY.Text = "txtxyChinTopY"
		Me.txtxyChinTopY.Visible = False
		Me.txtxyChinTopY.AcceptsReturn = True
		Me.txtxyChinTopY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtxyChinTopY.BackColor = System.Drawing.SystemColors.Window
		Me.txtxyChinTopY.CausesValidation = True
		Me.txtxyChinTopY.Enabled = True
		Me.txtxyChinTopY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtxyChinTopY.HideSelection = True
		Me.txtxyChinTopY.ReadOnly = False
		Me.txtxyChinTopY.Maxlength = 0
		Me.txtxyChinTopY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtxyChinTopY.MultiLine = False
		Me.txtxyChinTopY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtxyChinTopY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtxyChinTopY.TabStop = True
		Me.txtxyChinTopY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtxyChinTopY.Name = "txtxyChinTopY"
		Me.txtxyChinTopX.AutoSize = False
		Me.txtxyChinTopX.Size = New System.Drawing.Size(46, 19)
		Me.txtxyChinTopX.Location = New System.Drawing.Point(520, 270)
		Me.txtxyChinTopX.TabIndex = 106
		Me.txtxyChinTopX.Text = "txtxyChinTopX"
		Me.txtxyChinTopX.Visible = False
		Me.txtxyChinTopX.AcceptsReturn = True
		Me.txtxyChinTopX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtxyChinTopX.BackColor = System.Drawing.SystemColors.Window
		Me.txtxyChinTopX.CausesValidation = True
		Me.txtxyChinTopX.Enabled = True
		Me.txtxyChinTopX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtxyChinTopX.HideSelection = True
		Me.txtxyChinTopX.ReadOnly = False
		Me.txtxyChinTopX.Maxlength = 0
		Me.txtxyChinTopX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtxyChinTopX.MultiLine = False
		Me.txtxyChinTopX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtxyChinTopX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtxyChinTopX.TabStop = True
		Me.txtxyChinTopX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtxyChinTopX.Name = "txtxyChinTopX"
		Me.txtCSEarBotHeight.AutoSize = False
		Me.txtCSEarBotHeight.Size = New System.Drawing.Size(46, 19)
		Me.txtCSEarBotHeight.Location = New System.Drawing.Point(520, 250)
		Me.txtCSEarBotHeight.TabIndex = 105
		Me.txtCSEarBotHeight.Text = "txtCSEarBotHeight"
		Me.txtCSEarBotHeight.Visible = False
		Me.txtCSEarBotHeight.AcceptsReturn = True
		Me.txtCSEarBotHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSEarBotHeight.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSEarBotHeight.CausesValidation = True
		Me.txtCSEarBotHeight.Enabled = True
		Me.txtCSEarBotHeight.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSEarBotHeight.HideSelection = True
		Me.txtCSEarBotHeight.ReadOnly = False
		Me.txtCSEarBotHeight.Maxlength = 0
		Me.txtCSEarBotHeight.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSEarBotHeight.MultiLine = False
		Me.txtCSEarBotHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSEarBotHeight.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSEarBotHeight.TabStop = True
		Me.txtCSEarBotHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSEarBotHeight.Name = "txtCSEarBotHeight"
		Me.txtCSEarTopHeight.AutoSize = False
		Me.txtCSEarTopHeight.Size = New System.Drawing.Size(46, 19)
		Me.txtCSEarTopHeight.Location = New System.Drawing.Point(520, 230)
		Me.txtCSEarTopHeight.TabIndex = 104
		Me.txtCSEarTopHeight.Text = "txtCSEarTopHeight"
		Me.txtCSEarTopHeight.Visible = False
		Me.txtCSEarTopHeight.AcceptsReturn = True
		Me.txtCSEarTopHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSEarTopHeight.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSEarTopHeight.CausesValidation = True
		Me.txtCSEarTopHeight.Enabled = True
		Me.txtCSEarTopHeight.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSEarTopHeight.HideSelection = True
		Me.txtCSEarTopHeight.ReadOnly = False
		Me.txtCSEarTopHeight.Maxlength = 0
		Me.txtCSEarTopHeight.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSEarTopHeight.MultiLine = False
		Me.txtCSEarTopHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSEarTopHeight.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSEarTopHeight.TabStop = True
		Me.txtCSEarTopHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSEarTopHeight.Name = "txtCSEarTopHeight"
		Me.txtCSMouthWidth.AutoSize = False
		Me.txtCSMouthWidth.Size = New System.Drawing.Size(46, 19)
		Me.txtCSMouthWidth.Location = New System.Drawing.Point(520, 210)
		Me.txtCSMouthWidth.TabIndex = 103
		Me.txtCSMouthWidth.Text = "txtCSMouthWidth"
		Me.txtCSMouthWidth.Visible = False
		Me.txtCSMouthWidth.AcceptsReturn = True
		Me.txtCSMouthWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSMouthWidth.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSMouthWidth.CausesValidation = True
		Me.txtCSMouthWidth.Enabled = True
		Me.txtCSMouthWidth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSMouthWidth.HideSelection = True
		Me.txtCSMouthWidth.ReadOnly = False
		Me.txtCSMouthWidth.Maxlength = 0
		Me.txtCSMouthWidth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSMouthWidth.MultiLine = False
		Me.txtCSMouthWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSMouthWidth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSMouthWidth.TabStop = True
		Me.txtCSMouthWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSMouthWidth.Name = "txtCSMouthWidth"
		Me.txtMouthRightX.AutoSize = False
		Me.txtMouthRightX.Size = New System.Drawing.Size(46, 19)
		Me.txtMouthRightX.Location = New System.Drawing.Point(570, 70)
		Me.txtMouthRightX.TabIndex = 67
		Me.txtMouthRightX.Text = "txtMouthRightX"
		Me.txtMouthRightX.Visible = False
		Me.txtMouthRightX.AcceptsReturn = True
		Me.txtMouthRightX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMouthRightX.BackColor = System.Drawing.SystemColors.Window
		Me.txtMouthRightX.CausesValidation = True
		Me.txtMouthRightX.Enabled = True
		Me.txtMouthRightX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMouthRightX.HideSelection = True
		Me.txtMouthRightX.ReadOnly = False
		Me.txtMouthRightX.Maxlength = 0
		Me.txtMouthRightX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMouthRightX.MultiLine = False
		Me.txtMouthRightX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMouthRightX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMouthRightX.TabStop = True
		Me.txtMouthRightX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMouthRightX.Name = "txtMouthRightX"
		Me.txtCSForeHead.AutoSize = False
		Me.txtCSForeHead.Size = New System.Drawing.Size(46, 19)
		Me.txtCSForeHead.Location = New System.Drawing.Point(520, 190)
		Me.txtCSForeHead.TabIndex = 99
		Me.txtCSForeHead.Text = "txtCSForeHead"
		Me.txtCSForeHead.Visible = False
		Me.txtCSForeHead.AcceptsReturn = True
		Me.txtCSForeHead.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSForeHead.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSForeHead.CausesValidation = True
		Me.txtCSForeHead.Enabled = True
		Me.txtCSForeHead.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSForeHead.HideSelection = True
		Me.txtCSForeHead.ReadOnly = False
		Me.txtCSForeHead.Maxlength = 0
		Me.txtCSForeHead.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSForeHead.MultiLine = False
		Me.txtCSForeHead.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSForeHead.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSForeHead.TabStop = True
		Me.txtCSForeHead.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSForeHead.Name = "txtCSForeHead"
		Me.txtCSChinAngle.AutoSize = False
		Me.txtCSChinAngle.Size = New System.Drawing.Size(46, 19)
		Me.txtCSChinAngle.Location = New System.Drawing.Point(520, 170)
		Me.txtCSChinAngle.TabIndex = 98
		Me.txtCSChinAngle.Text = "txtCSChinAngle"
		Me.txtCSChinAngle.Visible = False
		Me.txtCSChinAngle.AcceptsReturn = True
		Me.txtCSChinAngle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSChinAngle.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSChinAngle.CausesValidation = True
		Me.txtCSChinAngle.Enabled = True
		Me.txtCSChinAngle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSChinAngle.HideSelection = True
		Me.txtCSChinAngle.ReadOnly = False
		Me.txtCSChinAngle.Maxlength = 0
		Me.txtCSChinAngle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSChinAngle.MultiLine = False
		Me.txtCSChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSChinAngle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSChinAngle.TabStop = True
		Me.txtCSChinAngle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSChinAngle.Name = "txtCSChinAngle"
		Me.txtCSNeckCircum.AutoSize = False
		Me.txtCSNeckCircum.Size = New System.Drawing.Size(46, 19)
		Me.txtCSNeckCircum.Location = New System.Drawing.Point(520, 150)
		Me.txtCSNeckCircum.TabIndex = 97
		Me.txtCSNeckCircum.Text = "txtCSNeckCircum"
		Me.txtCSNeckCircum.Visible = False
		Me.txtCSNeckCircum.AcceptsReturn = True
		Me.txtCSNeckCircum.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSNeckCircum.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSNeckCircum.CausesValidation = True
		Me.txtCSNeckCircum.Enabled = True
		Me.txtCSNeckCircum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSNeckCircum.HideSelection = True
		Me.txtCSNeckCircum.ReadOnly = False
		Me.txtCSNeckCircum.Maxlength = 0
		Me.txtCSNeckCircum.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSNeckCircum.MultiLine = False
		Me.txtCSNeckCircum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSNeckCircum.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSNeckCircum.TabStop = True
		Me.txtCSNeckCircum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSNeckCircum.Name = "txtCSNeckCircum"
		Me.txtCSChintoMouth.AutoSize = False
		Me.txtCSChintoMouth.Size = New System.Drawing.Size(46, 19)
		Me.txtCSChintoMouth.Location = New System.Drawing.Point(520, 130)
		Me.txtCSChintoMouth.TabIndex = 96
		Me.txtCSChintoMouth.Text = "txtCSChintoMouth"
		Me.txtCSChintoMouth.Visible = False
		Me.txtCSChintoMouth.AcceptsReturn = True
		Me.txtCSChintoMouth.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCSChintoMouth.BackColor = System.Drawing.SystemColors.Window
		Me.txtCSChintoMouth.CausesValidation = True
		Me.txtCSChintoMouth.Enabled = True
		Me.txtCSChintoMouth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCSChintoMouth.HideSelection = True
		Me.txtCSChintoMouth.ReadOnly = False
		Me.txtCSChintoMouth.Maxlength = 0
		Me.txtCSChintoMouth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCSChintoMouth.MultiLine = False
		Me.txtCSChintoMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCSChintoMouth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCSChintoMouth.TabStop = True
		Me.txtCSChintoMouth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCSChintoMouth.Name = "txtCSChintoMouth"
		Me.txtNoseCoverX.AutoSize = False
		Me.txtNoseCoverX.Size = New System.Drawing.Size(46, 19)
		Me.txtNoseCoverX.Location = New System.Drawing.Point(570, 50)
		Me.txtNoseCoverX.TabIndex = 95
		Me.txtNoseCoverX.Text = "txtNoseCoverX"
		Me.txtNoseCoverX.Visible = False
		Me.txtNoseCoverX.AcceptsReturn = True
		Me.txtNoseCoverX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNoseCoverX.BackColor = System.Drawing.SystemColors.Window
		Me.txtNoseCoverX.CausesValidation = True
		Me.txtNoseCoverX.Enabled = True
		Me.txtNoseCoverX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNoseCoverX.HideSelection = True
		Me.txtNoseCoverX.ReadOnly = False
		Me.txtNoseCoverX.Maxlength = 0
		Me.txtNoseCoverX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNoseCoverX.MultiLine = False
		Me.txtNoseCoverX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNoseCoverX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNoseCoverX.TabStop = True
		Me.txtNoseCoverX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtNoseCoverX.Name = "txtNoseCoverX"
		Me.txtNoseCoverY.AutoSize = False
		Me.txtNoseCoverY.Size = New System.Drawing.Size(46, 19)
		Me.txtNoseCoverY.Location = New System.Drawing.Point(570, 30)
		Me.txtNoseCoverY.TabIndex = 94
		Me.txtNoseCoverY.Text = "txtNoseCoverY"
		Me.txtNoseCoverY.Visible = False
		Me.txtNoseCoverY.AcceptsReturn = True
		Me.txtNoseCoverY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNoseCoverY.BackColor = System.Drawing.SystemColors.Window
		Me.txtNoseCoverY.CausesValidation = True
		Me.txtNoseCoverY.Enabled = True
		Me.txtNoseCoverY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNoseCoverY.HideSelection = True
		Me.txtNoseCoverY.ReadOnly = False
		Me.txtNoseCoverY.Maxlength = 0
		Me.txtNoseCoverY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNoseCoverY.MultiLine = False
		Me.txtNoseCoverY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNoseCoverY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNoseCoverY.TabStop = True
		Me.txtNoseCoverY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtNoseCoverY.Name = "txtNoseCoverY"
		Me.txtEyeWidth.AutoSize = False
		Me.txtEyeWidth.Size = New System.Drawing.Size(46, 19)
		Me.txtEyeWidth.Location = New System.Drawing.Point(520, 90)
		Me.txtEyeWidth.TabIndex = 93
		Me.txtEyeWidth.Text = "txtEyeWidth"
		Me.txtEyeWidth.Visible = False
		Me.txtEyeWidth.AcceptsReturn = True
		Me.txtEyeWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEyeWidth.BackColor = System.Drawing.SystemColors.Window
		Me.txtEyeWidth.CausesValidation = True
		Me.txtEyeWidth.Enabled = True
		Me.txtEyeWidth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEyeWidth.HideSelection = True
		Me.txtEyeWidth.ReadOnly = False
		Me.txtEyeWidth.Maxlength = 0
		Me.txtEyeWidth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEyeWidth.MultiLine = False
		Me.txtEyeWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEyeWidth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEyeWidth.TabStop = True
		Me.txtEyeWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtEyeWidth.Name = "txtEyeWidth"
		Me.txtDartEndX.AutoSize = False
		Me.txtDartEndX.Size = New System.Drawing.Size(46, 19)
		Me.txtDartEndX.Location = New System.Drawing.Point(520, 50)
		Me.txtDartEndX.TabIndex = 92
		Me.txtDartEndX.Text = "txtDartEndX"
		Me.txtDartEndX.Visible = False
		Me.txtDartEndX.AcceptsReturn = True
		Me.txtDartEndX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDartEndX.BackColor = System.Drawing.SystemColors.Window
		Me.txtDartEndX.CausesValidation = True
		Me.txtDartEndX.Enabled = True
		Me.txtDartEndX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDartEndX.HideSelection = True
		Me.txtDartEndX.ReadOnly = False
		Me.txtDartEndX.Maxlength = 0
		Me.txtDartEndX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDartEndX.MultiLine = False
		Me.txtDartEndX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDartEndX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDartEndX.TabStop = True
		Me.txtDartEndX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDartEndX.Name = "txtDartEndX"
		Me.txtDartEndY.AutoSize = False
		Me.txtDartEndY.Size = New System.Drawing.Size(46, 19)
		Me.txtDartEndY.Location = New System.Drawing.Point(520, 70)
		Me.txtDartEndY.TabIndex = 91
		Me.txtDartEndY.Text = "txtDartEndY"
		Me.txtDartEndY.Visible = False
		Me.txtDartEndY.AcceptsReturn = True
		Me.txtDartEndY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDartEndY.BackColor = System.Drawing.SystemColors.Window
		Me.txtDartEndY.CausesValidation = True
		Me.txtDartEndY.Enabled = True
		Me.txtDartEndY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDartEndY.HideSelection = True
		Me.txtDartEndY.ReadOnly = False
		Me.txtDartEndY.Maxlength = 0
		Me.txtDartEndY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDartEndY.MultiLine = False
		Me.txtDartEndY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDartEndY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDartEndY.TabStop = True
		Me.txtDartEndY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDartEndY.Name = "txtDartEndY"
		Me.txtDartStartY.AutoSize = False
		Me.txtDartStartY.Size = New System.Drawing.Size(46, 19)
		Me.txtDartStartY.Location = New System.Drawing.Point(520, 30)
		Me.txtDartStartY.TabIndex = 90
		Me.txtDartStartY.Text = "txtDartStartY"
		Me.txtDartStartY.Visible = False
		Me.txtDartStartY.AcceptsReturn = True
		Me.txtDartStartY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDartStartY.BackColor = System.Drawing.SystemColors.Window
		Me.txtDartStartY.CausesValidation = True
		Me.txtDartStartY.Enabled = True
		Me.txtDartStartY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDartStartY.HideSelection = True
		Me.txtDartStartY.ReadOnly = False
		Me.txtDartStartY.Maxlength = 0
		Me.txtDartStartY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDartStartY.MultiLine = False
		Me.txtDartStartY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDartStartY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDartStartY.TabStop = True
		Me.txtDartStartY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDartStartY.Name = "txtDartStartY"
		Me.txtDartStartX.AutoSize = False
		Me.txtDartStartX.Size = New System.Drawing.Size(46, 19)
		Me.txtDartStartX.Location = New System.Drawing.Point(520, 10)
		Me.txtDartStartX.TabIndex = 89
		Me.txtDartStartX.Text = "txtDartStartX"
		Me.txtDartStartX.Visible = False
		Me.txtDartStartX.AcceptsReturn = True
		Me.txtDartStartX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDartStartX.BackColor = System.Drawing.SystemColors.Window
		Me.txtDartStartX.CausesValidation = True
		Me.txtDartStartX.Enabled = True
		Me.txtDartStartX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDartStartX.HideSelection = True
		Me.txtDartStartX.ReadOnly = False
		Me.txtDartStartX.Maxlength = 0
		Me.txtDartStartX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDartStartX.MultiLine = False
		Me.txtDartStartX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDartStartX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDartStartX.TabStop = True
		Me.txtDartStartX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDartStartX.Name = "txtDartStartX"
		Me.txtChinLeftBotY.AutoSize = False
		Me.txtChinLeftBotY.Size = New System.Drawing.Size(46, 19)
		Me.txtChinLeftBotY.Location = New System.Drawing.Point(570, 190)
		Me.txtChinLeftBotY.TabIndex = 88
		Me.txtChinLeftBotY.Text = "txtChinLeftBotY"
		Me.txtChinLeftBotY.Visible = False
		Me.txtChinLeftBotY.AcceptsReturn = True
		Me.txtChinLeftBotY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtChinLeftBotY.BackColor = System.Drawing.SystemColors.Window
		Me.txtChinLeftBotY.CausesValidation = True
		Me.txtChinLeftBotY.Enabled = True
		Me.txtChinLeftBotY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChinLeftBotY.HideSelection = True
		Me.txtChinLeftBotY.ReadOnly = False
		Me.txtChinLeftBotY.Maxlength = 0
		Me.txtChinLeftBotY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChinLeftBotY.MultiLine = False
		Me.txtChinLeftBotY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChinLeftBotY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChinLeftBotY.TabStop = True
		Me.txtChinLeftBotY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtChinLeftBotY.Name = "txtChinLeftBotY"
		Me.txtChinLeftBotX.AutoSize = False
		Me.txtChinLeftBotX.Size = New System.Drawing.Size(46, 19)
		Me.txtChinLeftBotX.Location = New System.Drawing.Point(570, 170)
		Me.txtChinLeftBotX.TabIndex = 85
		Me.txtChinLeftBotX.Text = "txtChinLeftBotX"
		Me.txtChinLeftBotX.Visible = False
		Me.txtChinLeftBotX.AcceptsReturn = True
		Me.txtChinLeftBotX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtChinLeftBotX.BackColor = System.Drawing.SystemColors.Window
		Me.txtChinLeftBotX.CausesValidation = True
		Me.txtChinLeftBotX.Enabled = True
		Me.txtChinLeftBotX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChinLeftBotX.HideSelection = True
		Me.txtChinLeftBotX.ReadOnly = False
		Me.txtChinLeftBotX.Maxlength = 0
		Me.txtChinLeftBotX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChinLeftBotX.MultiLine = False
		Me.txtChinLeftBotX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChinLeftBotX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChinLeftBotX.TabStop = True
		Me.txtChinLeftBotX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtChinLeftBotX.Name = "txtChinLeftBotX"
		Me.txtNoseBottomY.AutoSize = False
		Me.txtNoseBottomY.Size = New System.Drawing.Size(46, 19)
		Me.txtNoseBottomY.Location = New System.Drawing.Point(570, 10)
		Me.txtNoseBottomY.TabIndex = 84
		Me.txtNoseBottomY.Text = "txtNoseBottomY"
		Me.txtNoseBottomY.Visible = False
		Me.txtNoseBottomY.AcceptsReturn = True
		Me.txtNoseBottomY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNoseBottomY.BackColor = System.Drawing.SystemColors.Window
		Me.txtNoseBottomY.CausesValidation = True
		Me.txtNoseBottomY.Enabled = True
		Me.txtNoseBottomY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNoseBottomY.HideSelection = True
		Me.txtNoseBottomY.ReadOnly = False
		Me.txtNoseBottomY.Maxlength = 0
		Me.txtNoseBottomY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNoseBottomY.MultiLine = False
		Me.txtNoseBottomY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNoseBottomY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNoseBottomY.TabStop = True
		Me.txtNoseBottomY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtNoseBottomY.Name = "txtNoseBottomY"
		Me.txtCircumferenceTotal.AutoSize = False
		Me.txtCircumferenceTotal.Size = New System.Drawing.Size(46, 19)
		Me.txtCircumferenceTotal.Location = New System.Drawing.Point(470, 250)
		Me.txtCircumferenceTotal.TabIndex = 83
		Me.txtCircumferenceTotal.Text = "txtCircumferenceTotal"
		Me.txtCircumferenceTotal.Visible = False
		Me.txtCircumferenceTotal.AcceptsReturn = True
		Me.txtCircumferenceTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCircumferenceTotal.BackColor = System.Drawing.SystemColors.Window
		Me.txtCircumferenceTotal.CausesValidation = True
		Me.txtCircumferenceTotal.Enabled = True
		Me.txtCircumferenceTotal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCircumferenceTotal.HideSelection = True
		Me.txtCircumferenceTotal.ReadOnly = False
		Me.txtCircumferenceTotal.Maxlength = 0
		Me.txtCircumferenceTotal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCircumferenceTotal.MultiLine = False
		Me.txtCircumferenceTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCircumferenceTotal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCircumferenceTotal.TabStop = True
		Me.txtCircumferenceTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCircumferenceTotal.Name = "txtCircumferenceTotal"
		Me.txtMouthHeight.AutoSize = False
		Me.txtMouthHeight.Size = New System.Drawing.Size(46, 19)
		Me.txtMouthHeight.Location = New System.Drawing.Point(520, 110)
		Me.txtMouthHeight.TabIndex = 82
		Me.txtMouthHeight.Text = "txtMouthHeight"
		Me.txtMouthHeight.Visible = False
		Me.txtMouthHeight.AcceptsReturn = True
		Me.txtMouthHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMouthHeight.BackColor = System.Drawing.SystemColors.Window
		Me.txtMouthHeight.CausesValidation = True
		Me.txtMouthHeight.Enabled = True
		Me.txtMouthHeight.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMouthHeight.HideSelection = True
		Me.txtMouthHeight.ReadOnly = False
		Me.txtMouthHeight.Maxlength = 0
		Me.txtMouthHeight.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMouthHeight.MultiLine = False
		Me.txtMouthHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMouthHeight.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMouthHeight.TabStop = True
		Me.txtMouthHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMouthHeight.Name = "txtMouthHeight"
		Me.txtBotOfEyeY.AutoSize = False
		Me.txtBotOfEyeY.Size = New System.Drawing.Size(46, 19)
		Me.txtBotOfEyeY.Location = New System.Drawing.Point(470, 310)
		Me.txtBotOfEyeY.TabIndex = 81
		Me.txtBotOfEyeY.Text = "txtBotOfEyeY"
		Me.txtBotOfEyeY.Visible = False
		Me.txtBotOfEyeY.AcceptsReturn = True
		Me.txtBotOfEyeY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBotOfEyeY.BackColor = System.Drawing.SystemColors.Window
		Me.txtBotOfEyeY.CausesValidation = True
		Me.txtBotOfEyeY.Enabled = True
		Me.txtBotOfEyeY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBotOfEyeY.HideSelection = True
		Me.txtBotOfEyeY.ReadOnly = False
		Me.txtBotOfEyeY.Maxlength = 0
		Me.txtBotOfEyeY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBotOfEyeY.MultiLine = False
		Me.txtBotOfEyeY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBotOfEyeY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBotOfEyeY.TabStop = True
		Me.txtBotOfEyeY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtBotOfEyeY.Name = "txtBotOfEyeY"
		Me.txtBotOfEyeX.AutoSize = False
		Me.txtBotOfEyeX.Size = New System.Drawing.Size(46, 19)
		Me.txtBotOfEyeX.Location = New System.Drawing.Point(470, 290)
		Me.txtBotOfEyeX.TabIndex = 80
		Me.txtBotOfEyeX.Text = "txtBotOfEyeX"
		Me.txtBotOfEyeX.Visible = False
		Me.txtBotOfEyeX.AcceptsReturn = True
		Me.txtBotOfEyeX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBotOfEyeX.BackColor = System.Drawing.SystemColors.Window
		Me.txtBotOfEyeX.CausesValidation = True
		Me.txtBotOfEyeX.Enabled = True
		Me.txtBotOfEyeX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBotOfEyeX.HideSelection = True
		Me.txtBotOfEyeX.ReadOnly = False
		Me.txtBotOfEyeX.Maxlength = 0
		Me.txtBotOfEyeX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBotOfEyeX.MultiLine = False
		Me.txtBotOfEyeX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBotOfEyeX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBotOfEyeX.TabStop = True
		Me.txtBotOfEyeX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtBotOfEyeX.Name = "txtBotOfEyeX"
		Me.txtLowerEarHeight.AutoSize = False
		Me.txtLowerEarHeight.Size = New System.Drawing.Size(46, 19)
		Me.txtLowerEarHeight.Location = New System.Drawing.Point(470, 270)
		Me.txtLowerEarHeight.TabIndex = 79
		Me.txtLowerEarHeight.Text = "txtLowerEarHeight"
		Me.txtLowerEarHeight.Visible = False
		Me.txtLowerEarHeight.AcceptsReturn = True
		Me.txtLowerEarHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLowerEarHeight.BackColor = System.Drawing.SystemColors.Window
		Me.txtLowerEarHeight.CausesValidation = True
		Me.txtLowerEarHeight.Enabled = True
		Me.txtLowerEarHeight.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLowerEarHeight.HideSelection = True
		Me.txtLowerEarHeight.ReadOnly = False
		Me.txtLowerEarHeight.Maxlength = 0
		Me.txtLowerEarHeight.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLowerEarHeight.MultiLine = False
		Me.txtLowerEarHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLowerEarHeight.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLowerEarHeight.TabStop = True
		Me.txtLowerEarHeight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtLowerEarHeight.Name = "txtLowerEarHeight"
		Me.txtfNum.AutoSize = False
		Me.txtfNum.Size = New System.Drawing.Size(46, 19)
		Me.txtfNum.Location = New System.Drawing.Point(470, 230)
		Me.txtfNum.TabIndex = 78
		Me.txtfNum.Text = "txtfNum"
		Me.txtfNum.Visible = False
		Me.txtfNum.AcceptsReturn = True
		Me.txtfNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtfNum.BackColor = System.Drawing.SystemColors.Window
		Me.txtfNum.CausesValidation = True
		Me.txtfNum.Enabled = True
		Me.txtfNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtfNum.HideSelection = True
		Me.txtfNum.ReadOnly = False
		Me.txtfNum.Maxlength = 0
		Me.txtfNum.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtfNum.MultiLine = False
		Me.txtfNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtfNum.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtfNum.TabStop = True
		Me.txtfNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtfNum.Name = "txtfNum"
		Me.txtCurrTextFont.AutoSize = False
		Me.txtCurrTextFont.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrTextFont.Location = New System.Drawing.Point(470, 210)
		Me.txtCurrTextFont.TabIndex = 77
		Me.txtCurrTextFont.Text = "txtCurrTextFont"
		Me.txtCurrTextFont.Visible = False
		Me.txtCurrTextFont.AcceptsReturn = True
		Me.txtCurrTextFont.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrTextFont.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrTextFont.CausesValidation = True
		Me.txtCurrTextFont.Enabled = True
		Me.txtCurrTextFont.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrTextFont.HideSelection = True
		Me.txtCurrTextFont.ReadOnly = False
		Me.txtCurrTextFont.Maxlength = 0
		Me.txtCurrTextFont.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrTextFont.MultiLine = False
		Me.txtCurrTextFont.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrTextFont.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrTextFont.TabStop = True
		Me.txtCurrTextFont.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrTextFont.Name = "txtCurrTextFont"
		Me.txtCurrTextVertJust.AutoSize = False
		Me.txtCurrTextVertJust.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrTextVertJust.Location = New System.Drawing.Point(470, 190)
		Me.txtCurrTextVertJust.TabIndex = 76
		Me.txtCurrTextVertJust.Text = "txtCurrTextVertJust"
		Me.txtCurrTextVertJust.Visible = False
		Me.txtCurrTextVertJust.AcceptsReturn = True
		Me.txtCurrTextVertJust.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrTextVertJust.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrTextVertJust.CausesValidation = True
		Me.txtCurrTextVertJust.Enabled = True
		Me.txtCurrTextVertJust.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrTextVertJust.HideSelection = True
		Me.txtCurrTextVertJust.ReadOnly = False
		Me.txtCurrTextVertJust.Maxlength = 0
		Me.txtCurrTextVertJust.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrTextVertJust.MultiLine = False
		Me.txtCurrTextVertJust.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrTextVertJust.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrTextVertJust.TabStop = True
		Me.txtCurrTextVertJust.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrTextVertJust.Name = "txtCurrTextVertJust"
		Me.txtCurrTextHorizJust.AutoSize = False
		Me.txtCurrTextHorizJust.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrTextHorizJust.Location = New System.Drawing.Point(470, 170)
		Me.txtCurrTextHorizJust.TabIndex = 75
		Me.txtCurrTextHorizJust.Text = "txtCurrTextHorizJust"
		Me.txtCurrTextHorizJust.Visible = False
		Me.txtCurrTextHorizJust.AcceptsReturn = True
		Me.txtCurrTextHorizJust.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrTextHorizJust.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrTextHorizJust.CausesValidation = True
		Me.txtCurrTextHorizJust.Enabled = True
		Me.txtCurrTextHorizJust.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrTextHorizJust.HideSelection = True
		Me.txtCurrTextHorizJust.ReadOnly = False
		Me.txtCurrTextHorizJust.Maxlength = 0
		Me.txtCurrTextHorizJust.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrTextHorizJust.MultiLine = False
		Me.txtCurrTextHorizJust.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrTextHorizJust.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrTextHorizJust.TabStop = True
		Me.txtCurrTextHorizJust.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrTextHorizJust.Name = "txtCurrTextHorizJust"
		Me.txtCurrTextAspect.AutoSize = False
		Me.txtCurrTextAspect.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrTextAspect.Location = New System.Drawing.Point(470, 150)
		Me.txtCurrTextAspect.TabIndex = 74
		Me.txtCurrTextAspect.Text = "txtCurrTextAspect"
		Me.txtCurrTextAspect.Visible = False
		Me.txtCurrTextAspect.AcceptsReturn = True
		Me.txtCurrTextAspect.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrTextAspect.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrTextAspect.CausesValidation = True
		Me.txtCurrTextAspect.Enabled = True
		Me.txtCurrTextAspect.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrTextAspect.HideSelection = True
		Me.txtCurrTextAspect.ReadOnly = False
		Me.txtCurrTextAspect.Maxlength = 0
		Me.txtCurrTextAspect.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrTextAspect.MultiLine = False
		Me.txtCurrTextAspect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrTextAspect.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrTextAspect.TabStop = True
		Me.txtCurrTextAspect.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrTextAspect.Name = "txtCurrTextAspect"
		Me.txtCurrTextHt.AutoSize = False
		Me.txtCurrTextHt.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrTextHt.Location = New System.Drawing.Point(470, 130)
		Me.txtCurrTextHt.TabIndex = 73
		Me.txtCurrTextHt.Text = "txtCurrTextHt"
		Me.txtCurrTextHt.Visible = False
		Me.txtCurrTextHt.AcceptsReturn = True
		Me.txtCurrTextHt.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrTextHt.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrTextHt.CausesValidation = True
		Me.txtCurrTextHt.Enabled = True
		Me.txtCurrTextHt.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrTextHt.HideSelection = True
		Me.txtCurrTextHt.ReadOnly = False
		Me.txtCurrTextHt.Maxlength = 0
		Me.txtCurrTextHt.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrTextHt.MultiLine = False
		Me.txtCurrTextHt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrTextHt.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrTextHt.TabStop = True
		Me.txtCurrTextHt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrTextHt.Name = "txtCurrTextHt"
		Me.txtCurrentLayer.AutoSize = False
		Me.txtCurrentLayer.Size = New System.Drawing.Size(46, 19)
		Me.txtCurrentLayer.Location = New System.Drawing.Point(470, 110)
		Me.txtCurrentLayer.TabIndex = 72
		Me.txtCurrentLayer.Text = "txtCurrentLayer"
		Me.txtCurrentLayer.Visible = False
		Me.txtCurrentLayer.AcceptsReturn = True
		Me.txtCurrentLayer.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCurrentLayer.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrentLayer.CausesValidation = True
		Me.txtCurrentLayer.Enabled = True
		Me.txtCurrentLayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrentLayer.HideSelection = True
		Me.txtCurrentLayer.ReadOnly = False
		Me.txtCurrentLayer.Maxlength = 0
		Me.txtCurrentLayer.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrentLayer.MultiLine = False
		Me.txtCurrentLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrentLayer.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrentLayer.TabStop = True
		Me.txtCurrentLayer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrentLayer.Name = "txtCurrentLayer"
		Me.txtMidToEyeTop.AutoSize = False
		Me.txtMidToEyeTop.Size = New System.Drawing.Size(46, 19)
		Me.txtMidToEyeTop.Location = New System.Drawing.Point(470, 90)
		Me.txtMidToEyeTop.TabIndex = 71
		Me.txtMidToEyeTop.Text = "txtMidToEyeTop"
		Me.txtMidToEyeTop.Visible = False
		Me.txtMidToEyeTop.AcceptsReturn = True
		Me.txtMidToEyeTop.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMidToEyeTop.BackColor = System.Drawing.SystemColors.Window
		Me.txtMidToEyeTop.CausesValidation = True
		Me.txtMidToEyeTop.Enabled = True
		Me.txtMidToEyeTop.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMidToEyeTop.HideSelection = True
		Me.txtMidToEyeTop.ReadOnly = False
		Me.txtMidToEyeTop.Maxlength = 0
		Me.txtMidToEyeTop.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMidToEyeTop.MultiLine = False
		Me.txtMidToEyeTop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMidToEyeTop.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMidToEyeTop.TabStop = True
		Me.txtMidToEyeTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMidToEyeTop.Name = "txtMidToEyeTop"
		Me.txtLipStrapWidth.AutoSize = False
		Me.txtLipStrapWidth.Size = New System.Drawing.Size(46, 19)
		Me.txtLipStrapWidth.Location = New System.Drawing.Point(470, 70)
		Me.txtLipStrapWidth.TabIndex = 70
		Me.txtLipStrapWidth.Text = "txtLipStrapWidth"
		Me.txtLipStrapWidth.Visible = False
		Me.txtLipStrapWidth.AcceptsReturn = True
		Me.txtLipStrapWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLipStrapWidth.BackColor = System.Drawing.SystemColors.Window
		Me.txtLipStrapWidth.CausesValidation = True
		Me.txtLipStrapWidth.Enabled = True
		Me.txtLipStrapWidth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLipStrapWidth.HideSelection = True
		Me.txtLipStrapWidth.ReadOnly = False
		Me.txtLipStrapWidth.Maxlength = 0
		Me.txtLipStrapWidth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLipStrapWidth.MultiLine = False
		Me.txtLipStrapWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLipStrapWidth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLipStrapWidth.TabStop = True
		Me.txtLipStrapWidth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtLipStrapWidth.Name = "txtLipStrapWidth"
		Me.txtUidMPD.AutoSize = False
		Me.txtUidMPD.Size = New System.Drawing.Size(86, 19)
		Me.txtUidMPD.Location = New System.Drawing.Point(472, 376)
		Me.txtUidMPD.TabIndex = 69
		Me.txtUidMPD.Text = "txtUidMPD"
		Me.txtUidMPD.Visible = False
		Me.txtUidMPD.AcceptsReturn = True
		Me.txtUidMPD.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUidMPD.BackColor = System.Drawing.SystemColors.Window
		Me.txtUidMPD.CausesValidation = True
		Me.txtUidMPD.Enabled = True
		Me.txtUidMPD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUidMPD.HideSelection = True
		Me.txtUidMPD.ReadOnly = False
		Me.txtUidMPD.Maxlength = 0
		Me.txtUidMPD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUidMPD.MultiLine = False
		Me.txtUidMPD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUidMPD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUidMPD.TabStop = True
		Me.txtUidMPD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtUidMPD.Name = "txtUidMPD"
		Me.txtRadiusNo.AutoSize = False
		Me.txtRadiusNo.Size = New System.Drawing.Size(46, 19)
		Me.txtRadiusNo.Location = New System.Drawing.Point(470, 30)
		Me.txtRadiusNo.TabIndex = 68
		Me.txtRadiusNo.Text = "txtRadiusNo"
		Me.txtRadiusNo.Visible = False
		Me.txtRadiusNo.AcceptsReturn = True
		Me.txtRadiusNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRadiusNo.BackColor = System.Drawing.SystemColors.Window
		Me.txtRadiusNo.CausesValidation = True
		Me.txtRadiusNo.Enabled = True
		Me.txtRadiusNo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRadiusNo.HideSelection = True
		Me.txtRadiusNo.ReadOnly = False
		Me.txtRadiusNo.Maxlength = 0
		Me.txtRadiusNo.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRadiusNo.MultiLine = False
		Me.txtRadiusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRadiusNo.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRadiusNo.TabStop = True
		Me.txtRadiusNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtRadiusNo.Name = "txtRadiusNo"
		Me.txtDraw.AutoSize = False
		Me.txtDraw.Size = New System.Drawing.Size(46, 19)
		Me.txtDraw.Location = New System.Drawing.Point(472, 10)
		Me.txtDraw.TabIndex = 102
		Me.txtDraw.Text = "txtDraw"
		Me.txtDraw.Visible = False
		Me.txtDraw.AcceptsReturn = True
		Me.txtDraw.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDraw.BackColor = System.Drawing.SystemColors.Window
		Me.txtDraw.CausesValidation = True
		Me.txtDraw.Enabled = True
		Me.txtDraw.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDraw.HideSelection = True
		Me.txtDraw.ReadOnly = False
		Me.txtDraw.Maxlength = 0
		Me.txtDraw.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDraw.MultiLine = False
		Me.txtDraw.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDraw.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDraw.TabStop = True
		Me.txtDraw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDraw.Name = "txtDraw"
		Me.cmdDraw.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDraw.Text = "Draw"
		Me.cmdDraw.Size = New System.Drawing.Size(96, 25)
		Me.cmdDraw.Location = New System.Drawing.Point(90, 440)
		Me.cmdDraw.TabIndex = 35
		Me.cmdDraw.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDraw.CausesValidation = True
		Me.cmdDraw.Enabled = True
		Me.cmdDraw.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDraw.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDraw.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDraw.TabStop = True
		Me.cmdDraw.Name = "cmdDraw"
		Me.frmMeasurements.Text = "Measurements"
		Me.frmMeasurements.Size = New System.Drawing.Size(451, 161)
		Me.frmMeasurements.Location = New System.Drawing.Point(5, 270)
		Me.frmMeasurements.TabIndex = 51
		Me.frmMeasurements.BackColor = System.Drawing.SystemColors.Control
		Me.frmMeasurements.Enabled = True
		Me.frmMeasurements.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmMeasurements.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmMeasurements.Visible = True
		Me.frmMeasurements.Padding = New System.Windows.Forms.Padding(0)
		Me.frmMeasurements.Name = "frmMeasurements"
		Me.txtChinCollarMin.AutoSize = False
		Me.txtChinCollarMin.Size = New System.Drawing.Size(36, 19)
		Me.txtChinCollarMin.Location = New System.Drawing.Point(365, 75)
		Me.txtChinCollarMin.TabIndex = 33
		Me.txtChinCollarMin.AcceptsReturn = True
		Me.txtChinCollarMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtChinCollarMin.BackColor = System.Drawing.SystemColors.Window
		Me.txtChinCollarMin.CausesValidation = True
		Me.txtChinCollarMin.Enabled = True
		Me.txtChinCollarMin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChinCollarMin.HideSelection = True
		Me.txtChinCollarMin.ReadOnly = False
		Me.txtChinCollarMin.Maxlength = 0
		Me.txtChinCollarMin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChinCollarMin.MultiLine = False
		Me.txtChinCollarMin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChinCollarMin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChinCollarMin.TabStop = True
		Me.txtChinCollarMin.Visible = True
		Me.txtChinCollarMin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChinCollarMin.Name = "txtChinCollarMin"
		Me.txtRightEarLength.AutoSize = False
		Me.txtRightEarLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtRightEarLength.Size = New System.Drawing.Size(36, 19)
		Me.txtRightEarLength.Location = New System.Drawing.Point(365, 55)
		Me.txtRightEarLength.Maxlength = 4
		Me.txtRightEarLength.TabIndex = 32
		Me.txtRightEarLength.AcceptsReturn = True
		Me.txtRightEarLength.BackColor = System.Drawing.SystemColors.Window
		Me.txtRightEarLength.CausesValidation = True
		Me.txtRightEarLength.Enabled = True
		Me.txtRightEarLength.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRightEarLength.HideSelection = True
		Me.txtRightEarLength.ReadOnly = False
		Me.txtRightEarLength.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRightEarLength.MultiLine = False
		Me.txtRightEarLength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRightEarLength.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRightEarLength.TabStop = True
		Me.txtRightEarLength.Visible = True
		Me.txtRightEarLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRightEarLength.Name = "txtRightEarLength"
		Me.txtLeftEarLength.AutoSize = False
		Me.txtLeftEarLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtLeftEarLength.Size = New System.Drawing.Size(36, 19)
		Me.txtLeftEarLength.Location = New System.Drawing.Point(365, 35)
		Me.txtLeftEarLength.Maxlength = 4
		Me.txtLeftEarLength.TabIndex = 31
		Me.txtLeftEarLength.AcceptsReturn = True
		Me.txtLeftEarLength.BackColor = System.Drawing.SystemColors.Window
		Me.txtLeftEarLength.CausesValidation = True
		Me.txtLeftEarLength.Enabled = True
		Me.txtLeftEarLength.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLeftEarLength.HideSelection = True
		Me.txtLeftEarLength.ReadOnly = False
		Me.txtLeftEarLength.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLeftEarLength.MultiLine = False
		Me.txtLeftEarLength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLeftEarLength.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLeftEarLength.TabStop = True
		Me.txtLeftEarLength.Visible = True
		Me.txtLeftEarLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLeftEarLength.Name = "txtLeftEarLength"
		Me.txtHeadBandDepth.AutoSize = False
		Me.txtHeadBandDepth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtHeadBandDepth.Size = New System.Drawing.Size(36, 19)
		Me.txtHeadBandDepth.Location = New System.Drawing.Point(365, 15)
		Me.txtHeadBandDepth.Maxlength = 4
		Me.txtHeadBandDepth.TabIndex = 30
		Me.txtHeadBandDepth.AcceptsReturn = True
		Me.txtHeadBandDepth.BackColor = System.Drawing.SystemColors.Window
		Me.txtHeadBandDepth.CausesValidation = True
		Me.txtHeadBandDepth.Enabled = True
		Me.txtHeadBandDepth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHeadBandDepth.HideSelection = True
		Me.txtHeadBandDepth.ReadOnly = False
		Me.txtHeadBandDepth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHeadBandDepth.MultiLine = False
		Me.txtHeadBandDepth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHeadBandDepth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHeadBandDepth.TabStop = True
		Me.txtHeadBandDepth.Visible = True
		Me.txtHeadBandDepth.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHeadBandDepth.Name = "txtHeadBandDepth"
		Me.txtLengthOfNose.AutoSize = False
		Me.txtLengthOfNose.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtLengthOfNose.Size = New System.Drawing.Size(36, 19)
		Me.txtLengthOfNose.Location = New System.Drawing.Point(170, 135)
		Me.txtLengthOfNose.Maxlength = 4
		Me.txtLengthOfNose.TabIndex = 29
		Me.txtLengthOfNose.AcceptsReturn = True
		Me.txtLengthOfNose.BackColor = System.Drawing.SystemColors.Window
		Me.txtLengthOfNose.CausesValidation = True
		Me.txtLengthOfNose.Enabled = True
		Me.txtLengthOfNose.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLengthOfNose.HideSelection = True
		Me.txtLengthOfNose.ReadOnly = False
		Me.txtLengthOfNose.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLengthOfNose.MultiLine = False
		Me.txtLengthOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLengthOfNose.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLengthOfNose.TabStop = True
		Me.txtLengthOfNose.Visible = True
		Me.txtLengthOfNose.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLengthOfNose.Name = "txtLengthOfNose"
		Me.txtTipOfNose.AutoSize = False
		Me.txtTipOfNose.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtTipOfNose.Size = New System.Drawing.Size(36, 19)
		Me.txtTipOfNose.Location = New System.Drawing.Point(170, 115)
		Me.txtTipOfNose.Maxlength = 4
		Me.txtTipOfNose.TabIndex = 28
		Me.txtTipOfNose.AcceptsReturn = True
		Me.txtTipOfNose.BackColor = System.Drawing.SystemColors.Window
		Me.txtTipOfNose.CausesValidation = True
		Me.txtTipOfNose.Enabled = True
		Me.txtTipOfNose.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTipOfNose.HideSelection = True
		Me.txtTipOfNose.ReadOnly = False
		Me.txtTipOfNose.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTipOfNose.MultiLine = False
		Me.txtTipOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTipOfNose.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTipOfNose.TabStop = True
		Me.txtTipOfNose.Visible = True
		Me.txtTipOfNose.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTipOfNose.Name = "txtTipOfNose"
		Me.cboFabric.Size = New System.Drawing.Size(183, 21)
		Me.cboFabric.Location = New System.Drawing.Point(255, 125)
		Me.cboFabric.TabIndex = 34
		Me.cboFabric.BackColor = System.Drawing.SystemColors.Window
		Me.cboFabric.CausesValidation = True
		Me.cboFabric.Enabled = True
		Me.cboFabric.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboFabric.IntegralHeight = True
		Me.cboFabric.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboFabric.Sorted = False
		Me.cboFabric.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboFabric.TabStop = True
		Me.cboFabric.Visible = True
		Me.cboFabric.Name = "cboFabric"
		Me.txtThroatToSternal.AutoSize = False
		Me.txtThroatToSternal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtThroatToSternal.Size = New System.Drawing.Size(36, 19)
		Me.txtThroatToSternal.Location = New System.Drawing.Point(170, 95)
		Me.txtThroatToSternal.Maxlength = 4
		Me.txtThroatToSternal.TabIndex = 27
		Me.txtThroatToSternal.AcceptsReturn = True
		Me.txtThroatToSternal.BackColor = System.Drawing.SystemColors.Window
		Me.txtThroatToSternal.CausesValidation = True
		Me.txtThroatToSternal.Enabled = True
		Me.txtThroatToSternal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtThroatToSternal.HideSelection = True
		Me.txtThroatToSternal.ReadOnly = False
		Me.txtThroatToSternal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtThroatToSternal.MultiLine = False
		Me.txtThroatToSternal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtThroatToSternal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtThroatToSternal.TabStop = True
		Me.txtThroatToSternal.Visible = True
		Me.txtThroatToSternal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtThroatToSternal.Name = "txtThroatToSternal"
		Me.txtCircOfNeck.AutoSize = False
		Me.txtCircOfNeck.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtCircOfNeck.Size = New System.Drawing.Size(36, 19)
		Me.txtCircOfNeck.Location = New System.Drawing.Point(170, 75)
		Me.txtCircOfNeck.Maxlength = 4
		Me.txtCircOfNeck.TabIndex = 26
		Me.txtCircOfNeck.AcceptsReturn = True
		Me.txtCircOfNeck.BackColor = System.Drawing.SystemColors.Window
		Me.txtCircOfNeck.CausesValidation = True
		Me.txtCircOfNeck.Enabled = True
		Me.txtCircOfNeck.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCircOfNeck.HideSelection = True
		Me.txtCircOfNeck.ReadOnly = False
		Me.txtCircOfNeck.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCircOfNeck.MultiLine = False
		Me.txtCircOfNeck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCircOfNeck.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCircOfNeck.TabStop = True
		Me.txtCircOfNeck.Visible = True
		Me.txtCircOfNeck.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCircOfNeck.Name = "txtCircOfNeck"
		Me.txtCircChinAngle.AutoSize = False
		Me.txtCircChinAngle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtCircChinAngle.Size = New System.Drawing.Size(36, 19)
		Me.txtCircChinAngle.Location = New System.Drawing.Point(170, 55)
		Me.txtCircChinAngle.Maxlength = 4
		Me.txtCircChinAngle.TabIndex = 25
		Me.txtCircChinAngle.AcceptsReturn = True
		Me.txtCircChinAngle.BackColor = System.Drawing.SystemColors.Window
		Me.txtCircChinAngle.CausesValidation = True
		Me.txtCircChinAngle.Enabled = True
		Me.txtCircChinAngle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCircChinAngle.HideSelection = True
		Me.txtCircChinAngle.ReadOnly = False
		Me.txtCircChinAngle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCircChinAngle.MultiLine = False
		Me.txtCircChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCircChinAngle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCircChinAngle.TabStop = True
		Me.txtCircChinAngle.Visible = True
		Me.txtCircChinAngle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCircChinAngle.Name = "txtCircChinAngle"
		Me.txtCircEyeBrow.AutoSize = False
		Me.txtCircEyeBrow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtCircEyeBrow.Size = New System.Drawing.Size(36, 19)
		Me.txtCircEyeBrow.Location = New System.Drawing.Point(170, 35)
		Me.txtCircEyeBrow.Maxlength = 4
		Me.txtCircEyeBrow.TabIndex = 24
		Me.txtCircEyeBrow.AcceptsReturn = True
		Me.txtCircEyeBrow.BackColor = System.Drawing.SystemColors.Window
		Me.txtCircEyeBrow.CausesValidation = True
		Me.txtCircEyeBrow.Enabled = True
		Me.txtCircEyeBrow.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCircEyeBrow.HideSelection = True
		Me.txtCircEyeBrow.ReadOnly = False
		Me.txtCircEyeBrow.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCircEyeBrow.MultiLine = False
		Me.txtCircEyeBrow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCircEyeBrow.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCircEyeBrow.TabStop = True
		Me.txtCircEyeBrow.Visible = True
		Me.txtCircEyeBrow.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCircEyeBrow.Name = "txtCircEyeBrow"
		Me.txtChinToMouth.AutoSize = False
		Me.txtChinToMouth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChinToMouth.Size = New System.Drawing.Size(36, 19)
		Me.txtChinToMouth.Location = New System.Drawing.Point(170, 15)
		Me.txtChinToMouth.Maxlength = 4
		Me.txtChinToMouth.TabIndex = 23
		Me.txtChinToMouth.AcceptsReturn = True
		Me.txtChinToMouth.BackColor = System.Drawing.SystemColors.Window
		Me.txtChinToMouth.CausesValidation = True
		Me.txtChinToMouth.Enabled = True
		Me.txtChinToMouth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChinToMouth.HideSelection = True
		Me.txtChinToMouth.ReadOnly = False
		Me.txtChinToMouth.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChinToMouth.MultiLine = False
		Me.txtChinToMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChinToMouth.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChinToMouth.TabStop = True
		Me.txtChinToMouth.Visible = True
		Me.txtChinToMouth.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChinToMouth.Name = "txtChinToMouth"
		Me.lblInch11.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch11.Size = New System.Drawing.Size(46, 16)
		Me.lblInch11.Location = New System.Drawing.Point(400, 78)
		Me.lblInch11.TabIndex = 122
		Me.lblInch11.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch11.Enabled = True
		Me.lblInch11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch11.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch11.UseMnemonic = True
		Me.lblInch11.Visible = True
		Me.lblInch11.AutoSize = False
		Me.lblInch11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch11.Name = "lblInch11"
		Me.lblChinCollarMin.Text = "Collar Contour"
		Me.lblChinCollarMin.Size = New System.Drawing.Size(101, 16)
		Me.lblChinCollarMin.Location = New System.Drawing.Point(255, 80)
		Me.lblChinCollarMin.TabIndex = 121
		Me.lblChinCollarMin.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblChinCollarMin.BackColor = System.Drawing.SystemColors.Control
		Me.lblChinCollarMin.Enabled = True
		Me.lblChinCollarMin.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblChinCollarMin.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblChinCollarMin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblChinCollarMin.UseMnemonic = True
		Me.lblChinCollarMin.Visible = True
		Me.lblChinCollarMin.AutoSize = False
		Me.lblChinCollarMin.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblChinCollarMin.Name = "lblChinCollarMin"
		Me.lblRightEarLen.Text = "Right Ear Length"
		Me.lblRightEarLen.Size = New System.Drawing.Size(111, 16)
		Me.lblRightEarLen.Location = New System.Drawing.Point(255, 60)
		Me.lblRightEarLen.TabIndex = 87
		Me.lblRightEarLen.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRightEarLen.BackColor = System.Drawing.SystemColors.Control
		Me.lblRightEarLen.Enabled = True
		Me.lblRightEarLen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblRightEarLen.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRightEarLen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRightEarLen.UseMnemonic = True
		Me.lblRightEarLen.Visible = True
		Me.lblRightEarLen.AutoSize = False
		Me.lblRightEarLen.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRightEarLen.Name = "lblRightEarLen"
		Me.lblLeftEarLen.Text = "Left Ear Length"
		Me.lblLeftEarLen.Size = New System.Drawing.Size(111, 16)
		Me.lblLeftEarLen.Location = New System.Drawing.Point(255, 40)
		Me.lblLeftEarLen.TabIndex = 114
		Me.lblLeftEarLen.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLeftEarLen.BackColor = System.Drawing.SystemColors.Control
		Me.lblLeftEarLen.Enabled = True
		Me.lblLeftEarLen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLeftEarLen.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLeftEarLen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLeftEarLen.UseMnemonic = True
		Me.lblLeftEarLen.Visible = True
		Me.lblLeftEarLen.AutoSize = False
		Me.lblLeftEarLen.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLeftEarLen.Name = "lblLeftEarLen"
		Me.lblInch10.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch10.Size = New System.Drawing.Size(46, 16)
		Me.lblInch10.Location = New System.Drawing.Point(400, 58)
		Me.lblInch10.TabIndex = 113
		Me.lblInch10.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch10.Enabled = True
		Me.lblInch10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch10.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch10.UseMnemonic = True
		Me.lblInch10.Visible = True
		Me.lblInch10.AutoSize = False
		Me.lblInch10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch10.Name = "lblInch10"
		Me.lblInch9.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch9.Size = New System.Drawing.Size(46, 16)
		Me.lblInch9.Location = New System.Drawing.Point(400, 38)
		Me.lblInch9.TabIndex = 112
		Me.lblInch9.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch9.Enabled = True
		Me.lblInch9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch9.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch9.UseMnemonic = True
		Me.lblInch9.Visible = True
		Me.lblInch9.AutoSize = False
		Me.lblInch9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch9.Name = "lblInch9"
		Me.lblInch8.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch8.Size = New System.Drawing.Size(46, 16)
		Me.lblInch8.Location = New System.Drawing.Point(400, 18)
		Me.lblInch8.TabIndex = 101
		Me.lblInch8.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch8.Enabled = True
		Me.lblInch8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch8.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch8.UseMnemonic = True
		Me.lblInch8.Visible = True
		Me.lblInch8.AutoSize = False
		Me.lblInch8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch8.Name = "lblInch8"
		Me.lblHeadBandDepth.Text = "Head Band Depth"
		Me.lblHeadBandDepth.Size = New System.Drawing.Size(111, 16)
		Me.lblHeadBandDepth.Location = New System.Drawing.Point(255, 20)
		Me.lblHeadBandDepth.TabIndex = 100
		Me.lblHeadBandDepth.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblHeadBandDepth.BackColor = System.Drawing.SystemColors.Control
		Me.lblHeadBandDepth.Enabled = True
		Me.lblHeadBandDepth.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblHeadBandDepth.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblHeadBandDepth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblHeadBandDepth.UseMnemonic = True
		Me.lblHeadBandDepth.Visible = True
		Me.lblHeadBandDepth.AutoSize = False
		Me.lblHeadBandDepth.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblHeadBandDepth.Name = "lblHeadBandDepth"
		Me.lblFabric.Text = "Fabric : "
		Me.lblFabric.Size = New System.Drawing.Size(46, 16)
		Me.lblFabric.Location = New System.Drawing.Point(255, 110)
		Me.lblFabric.TabIndex = 66
		Me.lblFabric.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFabric.BackColor = System.Drawing.SystemColors.Control
		Me.lblFabric.Enabled = True
		Me.lblFabric.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFabric.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFabric.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFabric.UseMnemonic = True
		Me.lblFabric.Visible = True
		Me.lblFabric.AutoSize = False
		Me.lblFabric.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFabric.Name = "lblFabric"
		Me.lblInch7.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch7.Size = New System.Drawing.Size(46, 16)
		Me.lblInch7.Location = New System.Drawing.Point(205, 138)
		Me.lblInch7.TabIndex = 65
		Me.lblInch7.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch7.Enabled = True
		Me.lblInch7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch7.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch7.UseMnemonic = True
		Me.lblInch7.Visible = True
		Me.lblInch7.AutoSize = False
		Me.lblInch7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch7.Name = "lblInch7"
		Me.lblInch6.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch6.Size = New System.Drawing.Size(46, 16)
		Me.lblInch6.Location = New System.Drawing.Point(205, 118)
		Me.lblInch6.TabIndex = 64
		Me.lblInch6.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch6.Enabled = True
		Me.lblInch6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch6.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch6.UseMnemonic = True
		Me.lblInch6.Visible = True
		Me.lblInch6.AutoSize = False
		Me.lblInch6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch6.Name = "lblInch6"
		Me.LblLengthOfNose.Text = "Length of Nose"
		Me.LblLengthOfNose.Size = New System.Drawing.Size(101, 16)
		Me.LblLengthOfNose.Location = New System.Drawing.Point(10, 140)
		Me.LblLengthOfNose.TabIndex = 63
		Me.LblLengthOfNose.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LblLengthOfNose.BackColor = System.Drawing.SystemColors.Control
		Me.LblLengthOfNose.Enabled = True
		Me.LblLengthOfNose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblLengthOfNose.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblLengthOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblLengthOfNose.UseMnemonic = True
		Me.LblLengthOfNose.Visible = True
		Me.LblLengthOfNose.AutoSize = False
		Me.LblLengthOfNose.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblLengthOfNose.Name = "LblLengthOfNose"
		Me.lblAcrossTipOfNose.Text = "Across Tip of Nose"
		Me.lblAcrossTipOfNose.Size = New System.Drawing.Size(116, 16)
		Me.lblAcrossTipOfNose.Location = New System.Drawing.Point(10, 120)
		Me.lblAcrossTipOfNose.TabIndex = 62
		Me.lblAcrossTipOfNose.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAcrossTipOfNose.BackColor = System.Drawing.SystemColors.Control
		Me.lblAcrossTipOfNose.Enabled = True
		Me.lblAcrossTipOfNose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAcrossTipOfNose.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAcrossTipOfNose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAcrossTipOfNose.UseMnemonic = True
		Me.lblAcrossTipOfNose.Visible = True
		Me.lblAcrossTipOfNose.AutoSize = False
		Me.lblAcrossTipOfNose.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAcrossTipOfNose.Name = "lblAcrossTipOfNose"
		Me.lblInch5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch5.Size = New System.Drawing.Size(46, 16)
		Me.lblInch5.Location = New System.Drawing.Point(205, 98)
		Me.lblInch5.TabIndex = 61
		Me.lblInch5.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch5.Enabled = True
		Me.lblInch5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch5.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch5.UseMnemonic = True
		Me.lblInch5.Visible = True
		Me.lblInch5.AutoSize = False
		Me.lblInch5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch5.Name = "lblInch5"
		Me.lblInch4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch4.Size = New System.Drawing.Size(46, 16)
		Me.lblInch4.Location = New System.Drawing.Point(205, 78)
		Me.lblInch4.TabIndex = 60
		Me.lblInch4.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch4.Enabled = True
		Me.lblInch4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch4.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch4.UseMnemonic = True
		Me.lblInch4.Visible = True
		Me.lblInch4.AutoSize = False
		Me.lblInch4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch4.Name = "lblInch4"
		Me.lblInch3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch3.Size = New System.Drawing.Size(46, 16)
		Me.lblInch3.Location = New System.Drawing.Point(205, 58)
		Me.lblInch3.TabIndex = 59
		Me.lblInch3.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch3.Enabled = True
		Me.lblInch3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch3.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch3.UseMnemonic = True
		Me.lblInch3.Visible = True
		Me.lblInch3.AutoSize = False
		Me.lblInch3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch3.Name = "lblInch3"
		Me.lblInch2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch2.Size = New System.Drawing.Size(46, 16)
		Me.lblInch2.Location = New System.Drawing.Point(205, 38)
		Me.lblInch2.TabIndex = 58
		Me.lblInch2.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch2.Enabled = True
		Me.lblInch2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch2.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch2.UseMnemonic = True
		Me.lblInch2.Visible = True
		Me.lblInch2.AutoSize = False
		Me.lblInch2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch2.Name = "lblInch2"
		Me.lblInch1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblInch1.Size = New System.Drawing.Size(46, 16)
		Me.lblInch1.Location = New System.Drawing.Point(205, 18)
		Me.lblInch1.TabIndex = 57
		Me.lblInch1.BackColor = System.Drawing.SystemColors.Control
		Me.lblInch1.Enabled = True
		Me.lblInch1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInch1.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInch1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInch1.UseMnemonic = True
		Me.lblInch1.Visible = True
		Me.lblInch1.AutoSize = False
		Me.lblInch1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInch1.Name = "lblInch1"
		Me.lblThroatToSternal.Text = "Throat to Sternal Notch"
		Me.lblThroatToSternal.Size = New System.Drawing.Size(146, 16)
		Me.lblThroatToSternal.Location = New System.Drawing.Point(10, 100)
		Me.lblThroatToSternal.TabIndex = 56
		Me.lblThroatToSternal.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblThroatToSternal.BackColor = System.Drawing.SystemColors.Control
		Me.lblThroatToSternal.Enabled = True
		Me.lblThroatToSternal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblThroatToSternal.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblThroatToSternal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblThroatToSternal.UseMnemonic = True
		Me.lblThroatToSternal.Visible = True
		Me.lblThroatToSternal.AutoSize = False
		Me.lblThroatToSternal.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblThroatToSternal.Name = "lblThroatToSternal"
		Me.lblCircOfNeck.Text = "CIRC of Neck"
		Me.lblCircOfNeck.Size = New System.Drawing.Size(151, 16)
		Me.lblCircOfNeck.Location = New System.Drawing.Point(10, 80)
		Me.lblCircOfNeck.TabIndex = 55
		Me.lblCircOfNeck.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCircOfNeck.BackColor = System.Drawing.SystemColors.Control
		Me.lblCircOfNeck.Enabled = True
		Me.lblCircOfNeck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCircOfNeck.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCircOfNeck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCircOfNeck.UseMnemonic = True
		Me.lblCircOfNeck.Visible = True
		Me.lblCircOfNeck.AutoSize = False
		Me.lblCircOfNeck.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCircOfNeck.Name = "lblCircOfNeck"
		Me.lblCircChinAngle.Text = "CIRC of Head at Chin Angle"
		Me.lblCircChinAngle.Size = New System.Drawing.Size(161, 16)
		Me.lblCircChinAngle.Location = New System.Drawing.Point(10, 60)
		Me.lblCircChinAngle.TabIndex = 54
		Me.lblCircChinAngle.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCircChinAngle.BackColor = System.Drawing.SystemColors.Control
		Me.lblCircChinAngle.Enabled = True
		Me.lblCircChinAngle.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCircChinAngle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCircChinAngle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCircChinAngle.UseMnemonic = True
		Me.lblCircChinAngle.Visible = True
		Me.lblCircChinAngle.AutoSize = False
		Me.lblCircChinAngle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCircChinAngle.Name = "lblCircChinAngle"
		Me.lblCircEyeBrow.Text = "CIRC above Eyebrow"
		Me.lblCircEyeBrow.Size = New System.Drawing.Size(121, 16)
		Me.lblCircEyeBrow.Location = New System.Drawing.Point(10, 40)
		Me.lblCircEyeBrow.TabIndex = 53
		Me.lblCircEyeBrow.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCircEyeBrow.BackColor = System.Drawing.SystemColors.Control
		Me.lblCircEyeBrow.Enabled = True
		Me.lblCircEyeBrow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCircEyeBrow.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCircEyeBrow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCircEyeBrow.UseMnemonic = True
		Me.lblCircEyeBrow.Visible = True
		Me.lblCircEyeBrow.AutoSize = False
		Me.lblCircEyeBrow.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCircEyeBrow.Name = "lblCircEyeBrow"
		Me.lblChinToMouth.Text = "Chin to Mouth"
		Me.lblChinToMouth.Size = New System.Drawing.Size(81, 16)
		Me.lblChinToMouth.Location = New System.Drawing.Point(10, 20)
		Me.lblChinToMouth.TabIndex = 52
		Me.lblChinToMouth.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblChinToMouth.BackColor = System.Drawing.SystemColors.Control
		Me.lblChinToMouth.Enabled = True
		Me.lblChinToMouth.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblChinToMouth.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblChinToMouth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblChinToMouth.UseMnemonic = True
		Me.lblChinToMouth.Visible = True
		Me.lblChinToMouth.AutoSize = False
		Me.lblChinToMouth.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblChinToMouth.Name = "lblChinToMouth"
		Me.frmModifications.Text = "Modifications"
		Me.frmModifications.Size = New System.Drawing.Size(266, 181)
		Me.frmModifications.Location = New System.Drawing.Point(190, 85)
		Me.frmModifications.TabIndex = 44
		Me.frmModifications.BackColor = System.Drawing.SystemColors.Control
		Me.frmModifications.Enabled = True
		Me.frmModifications.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmModifications.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmModifications.Visible = True
		Me.frmModifications.Padding = New System.Windows.Forms.Padding(0)
		Me.frmModifications.Name = "frmModifications"
		Me.chkVelcro.Text = "2"" Wide Velcro"
		Me.chkVelcro.Size = New System.Drawing.Size(130, 16)
		Me.chkVelcro.Location = New System.Drawing.Point(130, 160)
		Me.chkVelcro.TabIndex = 124
		Me.chkVelcro.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkVelcro.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkVelcro.BackColor = System.Drawing.SystemColors.Control
		Me.chkVelcro.CausesValidation = True
		Me.chkVelcro.Enabled = True
		Me.chkVelcro.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkVelcro.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkVelcro.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkVelcro.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkVelcro.TabStop = True
		Me.chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkVelcro.Visible = True
		Me.chkVelcro.Name = "chkVelcro"
		Me.chkEyes.Text = "Include Eyes"
		Me.chkEyes.Size = New System.Drawing.Size(106, 16)
		Me.chkEyes.Location = New System.Drawing.Point(130, 100)
		Me.chkEyes.TabIndex = 20
		Me.chkEyes.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkEyes.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkEyes.BackColor = System.Drawing.SystemColors.Control
		Me.chkEyes.CausesValidation = True
		Me.chkEyes.Enabled = True
		Me.chkEyes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkEyes.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkEyes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkEyes.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkEyes.TabStop = True
		Me.chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkEyes.Visible = True
		Me.chkEyes.Name = "chkEyes"
		Me.chkRightEarClosed.Text = "Right Ear Closed"
		Me.chkRightEarClosed.Size = New System.Drawing.Size(132, 16)
		Me.chkRightEarClosed.Location = New System.Drawing.Point(130, 60)
		Me.chkRightEarClosed.TabIndex = 13
		Me.chkRightEarClosed.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRightEarClosed.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRightEarClosed.BackColor = System.Drawing.SystemColors.Control
		Me.chkRightEarClosed.CausesValidation = True
		Me.chkRightEarClosed.Enabled = True
		Me.chkRightEarClosed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRightEarClosed.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRightEarClosed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRightEarClosed.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRightEarClosed.TabStop = True
		Me.chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRightEarClosed.Visible = True
		Me.chkRightEarClosed.Name = "chkRightEarClosed"
		Me.chkLeftEarClosed.Text = "Left Ear Closed"
		Me.chkLeftEarClosed.Size = New System.Drawing.Size(114, 16)
		Me.chkLeftEarClosed.Location = New System.Drawing.Point(10, 60)
		Me.chkLeftEarClosed.TabIndex = 12
		Me.chkLeftEarClosed.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLeftEarClosed.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLeftEarClosed.BackColor = System.Drawing.SystemColors.Control
		Me.chkLeftEarClosed.CausesValidation = True
		Me.chkLeftEarClosed.Enabled = True
		Me.chkLeftEarClosed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLeftEarClosed.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLeftEarClosed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLeftEarClosed.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLeftEarClosed.TabStop = True
		Me.chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLeftEarClosed.Visible = True
		Me.chkLeftEarClosed.Name = "chkLeftEarClosed"
		Me.chkEarSize.Text = "Ear Size"
		Me.chkEarSize.Size = New System.Drawing.Size(71, 16)
		Me.chkEarSize.Location = New System.Drawing.Point(10, 80)
		Me.chkEarSize.TabIndex = 14
		Me.chkEarSize.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkEarSize.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkEarSize.BackColor = System.Drawing.SystemColors.Control
		Me.chkEarSize.CausesValidation = True
		Me.chkEarSize.Enabled = True
		Me.chkEarSize.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkEarSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkEarSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkEarSize.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkEarSize.TabStop = True
		Me.chkEarSize.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkEarSize.Visible = True
		Me.chkEarSize.Name = "chkEarSize"
		Me.chkLipStrap.Text = "Lip Strap"
		Me.chkLipStrap.Enabled = False
		Me.chkLipStrap.Size = New System.Drawing.Size(111, 16)
		Me.chkLipStrap.Location = New System.Drawing.Point(10, 100)
		Me.chkLipStrap.TabIndex = 15
		Me.chkLipStrap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLipStrap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLipStrap.BackColor = System.Drawing.SystemColors.Control
		Me.chkLipStrap.CausesValidation = True
		Me.chkLipStrap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLipStrap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLipStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLipStrap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLipStrap.TabStop = True
		Me.chkLipStrap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLipStrap.Visible = True
		Me.chkLipStrap.Name = "chkLipStrap"
		Me.chkOpenHeadMask.Text = "Open Head Mask"
		Me.chkOpenHeadMask.Size = New System.Drawing.Size(132, 13)
		Me.chkOpenHeadMask.Location = New System.Drawing.Point(10, 160)
		Me.chkOpenHeadMask.TabIndex = 18
		Me.chkOpenHeadMask.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkOpenHeadMask.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkOpenHeadMask.BackColor = System.Drawing.SystemColors.Control
		Me.chkOpenHeadMask.CausesValidation = True
		Me.chkOpenHeadMask.Enabled = True
		Me.chkOpenHeadMask.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkOpenHeadMask.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkOpenHeadMask.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkOpenHeadMask.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkOpenHeadMask.TabStop = True
		Me.chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkOpenHeadMask.Visible = True
		Me.chkOpenHeadMask.Name = "chkOpenHeadMask"
		Me.chkNeckElastic.Text = "Neck Elastic"
		Me.chkNeckElastic.Size = New System.Drawing.Size(121, 16)
		Me.chkNeckElastic.Location = New System.Drawing.Point(130, 80)
		Me.chkNeckElastic.TabIndex = 19
		Me.chkNeckElastic.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkNeckElastic.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNeckElastic.BackColor = System.Drawing.SystemColors.Control
		Me.chkNeckElastic.CausesValidation = True
		Me.chkNeckElastic.Enabled = True
		Me.chkNeckElastic.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkNeckElastic.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNeckElastic.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNeckElastic.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNeckElastic.TabStop = True
		Me.chkNeckElastic.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNeckElastic.Visible = True
		Me.chkNeckElastic.Name = "chkNeckElastic"
		Me.chkLining.Text = "Lining"
		Me.chkLining.Size = New System.Drawing.Size(61, 16)
		Me.chkLining.Location = New System.Drawing.Point(130, 120)
		Me.chkLining.TabIndex = 22
		Me.chkLining.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLining.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLining.BackColor = System.Drawing.SystemColors.Control
		Me.chkLining.CausesValidation = True
		Me.chkLining.Enabled = True
		Me.chkLining.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLining.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLining.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLining.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLining.TabStop = True
		Me.chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLining.Visible = True
		Me.chkLining.Name = "chkLining"
		Me.chkLeftEyeFlap.Text = "Left Eye Flap"
		Me.chkLeftEyeFlap.Size = New System.Drawing.Size(106, 16)
		Me.chkLeftEyeFlap.Location = New System.Drawing.Point(10, 20)
		Me.chkLeftEyeFlap.TabIndex = 8
		Me.chkLeftEyeFlap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLeftEyeFlap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLeftEyeFlap.BackColor = System.Drawing.SystemColors.Control
		Me.chkLeftEyeFlap.CausesValidation = True
		Me.chkLeftEyeFlap.Enabled = True
		Me.chkLeftEyeFlap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLeftEyeFlap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLeftEyeFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLeftEyeFlap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLeftEyeFlap.TabStop = True
		Me.chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLeftEyeFlap.Visible = True
		Me.chkLeftEyeFlap.Name = "chkLeftEyeFlap"
		Me.chkRightEyeFlap.Text = "Right Eye Flap"
		Me.chkRightEyeFlap.Size = New System.Drawing.Size(111, 16)
		Me.chkRightEyeFlap.Location = New System.Drawing.Point(130, 20)
		Me.chkRightEyeFlap.TabIndex = 9
		Me.chkRightEyeFlap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRightEyeFlap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRightEyeFlap.BackColor = System.Drawing.SystemColors.Control
		Me.chkRightEyeFlap.CausesValidation = True
		Me.chkRightEyeFlap.Enabled = True
		Me.chkRightEyeFlap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRightEyeFlap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRightEyeFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRightEyeFlap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRightEyeFlap.TabStop = True
		Me.chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRightEyeFlap.Visible = True
		Me.chkRightEyeFlap.Name = "chkRightEyeFlap"
		Me.chkZipper.Text = "Zipper"
		Me.chkZipper.Size = New System.Drawing.Size(61, 16)
		Me.chkZipper.Location = New System.Drawing.Point(130, 140)
		Me.chkZipper.TabIndex = 21
		Me.chkZipper.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkZipper.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkZipper.BackColor = System.Drawing.SystemColors.Control
		Me.chkZipper.CausesValidation = True
		Me.chkZipper.Enabled = True
		Me.chkZipper.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkZipper.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkZipper.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkZipper.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkZipper.TabStop = True
		Me.chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkZipper.Visible = True
		Me.chkZipper.Name = "chkZipper"
		Me.chkNoseCovering.Text = "Nose Covering"
		Me.chkNoseCovering.Size = New System.Drawing.Size(116, 16)
		Me.chkNoseCovering.Location = New System.Drawing.Point(10, 140)
		Me.chkNoseCovering.TabIndex = 17
		Me.chkNoseCovering.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkNoseCovering.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNoseCovering.BackColor = System.Drawing.SystemColors.Control
		Me.chkNoseCovering.CausesValidation = True
		Me.chkNoseCovering.Enabled = True
		Me.chkNoseCovering.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkNoseCovering.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoseCovering.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoseCovering.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoseCovering.TabStop = True
		Me.chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoseCovering.Visible = True
		Me.chkNoseCovering.Name = "chkNoseCovering"
		Me.chkLipCovering.Text = "Lip Covering"
		Me.chkLipCovering.Size = New System.Drawing.Size(111, 16)
		Me.chkLipCovering.Location = New System.Drawing.Point(10, 120)
		Me.chkLipCovering.TabIndex = 16
		Me.chkLipCovering.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLipCovering.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLipCovering.BackColor = System.Drawing.SystemColors.Control
		Me.chkLipCovering.CausesValidation = True
		Me.chkLipCovering.Enabled = True
		Me.chkLipCovering.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLipCovering.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLipCovering.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLipCovering.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLipCovering.TabStop = True
		Me.chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLipCovering.Visible = True
		Me.chkLipCovering.Name = "chkLipCovering"
		Me.chkRightEarFlap.Text = "Right Ear Flap"
		Me.chkRightEarFlap.Size = New System.Drawing.Size(106, 16)
		Me.chkRightEarFlap.Location = New System.Drawing.Point(130, 40)
		Me.chkRightEarFlap.TabIndex = 11
		Me.chkRightEarFlap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRightEarFlap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRightEarFlap.BackColor = System.Drawing.SystemColors.Control
		Me.chkRightEarFlap.CausesValidation = True
		Me.chkRightEarFlap.Enabled = True
		Me.chkRightEarFlap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRightEarFlap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRightEarFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRightEarFlap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRightEarFlap.TabStop = True
		Me.chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRightEarFlap.Visible = True
		Me.chkRightEarFlap.Name = "chkRightEarFlap"
		Me.chkLeftEarFlap.Text = "Left Ear Flap"
		Me.chkLeftEarFlap.Size = New System.Drawing.Size(126, 16)
		Me.chkLeftEarFlap.Location = New System.Drawing.Point(10, 40)
		Me.chkLeftEarFlap.TabIndex = 10
		Me.chkLeftEarFlap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLeftEarFlap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLeftEarFlap.BackColor = System.Drawing.SystemColors.Control
		Me.chkLeftEarFlap.CausesValidation = True
		Me.chkLeftEarFlap.Enabled = True
		Me.chkLeftEarFlap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLeftEarFlap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLeftEarFlap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLeftEarFlap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLeftEarFlap.TabStop = True
		Me.chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLeftEarFlap.Visible = True
		Me.chkLeftEarFlap.Name = "chkLeftEarFlap"
		Me.frmDesignChoices.Text = "Design Choices"
		Me.frmDesignChoices.Size = New System.Drawing.Size(176, 181)
		Me.frmDesignChoices.Location = New System.Drawing.Point(5, 85)
		Me.frmDesignChoices.TabIndex = 43
		Me.frmDesignChoices.BackColor = System.Drawing.SystemColors.Control
		Me.frmDesignChoices.Enabled = True
		Me.frmDesignChoices.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmDesignChoices.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmDesignChoices.Visible = True
		Me.frmDesignChoices.Padding = New System.Windows.Forms.Padding(0)
		Me.frmDesignChoices.Name = "frmDesignChoices"
		Me.optContouredChinCollar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optContouredChinCollar.Text = "Contoured Chin Collar"
		Me.optContouredChinCollar.Size = New System.Drawing.Size(146, 16)
		Me.optContouredChinCollar.Location = New System.Drawing.Point(10, 120)
		Me.optContouredChinCollar.TabIndex = 6
		Me.optContouredChinCollar.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optContouredChinCollar.BackColor = System.Drawing.SystemColors.Control
		Me.optContouredChinCollar.CausesValidation = True
		Me.optContouredChinCollar.Enabled = True
		Me.optContouredChinCollar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optContouredChinCollar.Cursor = System.Windows.Forms.Cursors.Default
		Me.optContouredChinCollar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optContouredChinCollar.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optContouredChinCollar.TabStop = True
		Me.optContouredChinCollar.Checked = False
		Me.optContouredChinCollar.Visible = True
		Me.optContouredChinCollar.Name = "optContouredChinCollar"
		Me.optChinCollar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optChinCollar.Text = "Chin Collar"
		Me.optChinCollar.Size = New System.Drawing.Size(96, 16)
		Me.optChinCollar.Location = New System.Drawing.Point(10, 100)
		Me.optChinCollar.TabIndex = 5
		Me.optChinCollar.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optChinCollar.BackColor = System.Drawing.SystemColors.Control
		Me.optChinCollar.CausesValidation = True
		Me.optChinCollar.Enabled = True
		Me.optChinCollar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optChinCollar.Cursor = System.Windows.Forms.Cursors.Default
		Me.optChinCollar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optChinCollar.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optChinCollar.TabStop = True
		Me.optChinCollar.Checked = False
		Me.optChinCollar.Visible = True
		Me.optChinCollar.Name = "optChinCollar"
		Me.optModifiedChinStrap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optModifiedChinStrap.Text = "Modified Chin Strap"
		Me.optModifiedChinStrap.Size = New System.Drawing.Size(136, 16)
		Me.optModifiedChinStrap.Location = New System.Drawing.Point(10, 80)
		Me.optModifiedChinStrap.TabIndex = 4
		Me.optModifiedChinStrap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optModifiedChinStrap.BackColor = System.Drawing.SystemColors.Control
		Me.optModifiedChinStrap.CausesValidation = True
		Me.optModifiedChinStrap.Enabled = True
		Me.optModifiedChinStrap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optModifiedChinStrap.Cursor = System.Windows.Forms.Cursors.Default
		Me.optModifiedChinStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optModifiedChinStrap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optModifiedChinStrap.TabStop = True
		Me.optModifiedChinStrap.Checked = False
		Me.optModifiedChinStrap.Visible = True
		Me.optModifiedChinStrap.Name = "optModifiedChinStrap"
		Me.optChinStrap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optChinStrap.Text = "Chin Strap"
		Me.optChinStrap.Size = New System.Drawing.Size(106, 16)
		Me.optChinStrap.Location = New System.Drawing.Point(10, 60)
		Me.optChinStrap.TabIndex = 3
		Me.optChinStrap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optChinStrap.BackColor = System.Drawing.SystemColors.Control
		Me.optChinStrap.CausesValidation = True
		Me.optChinStrap.Enabled = True
		Me.optChinStrap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optChinStrap.Cursor = System.Windows.Forms.Cursors.Default
		Me.optChinStrap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optChinStrap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optChinStrap.TabStop = True
		Me.optChinStrap.Checked = False
		Me.optChinStrap.Visible = True
		Me.optChinStrap.Name = "optChinStrap"
		Me.optHeadBand.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHeadBand.Text = "Head Band"
		Me.optHeadBand.Size = New System.Drawing.Size(121, 16)
		Me.optHeadBand.Location = New System.Drawing.Point(10, 140)
		Me.optHeadBand.TabIndex = 7
		Me.optHeadBand.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHeadBand.BackColor = System.Drawing.SystemColors.Control
		Me.optHeadBand.CausesValidation = True
		Me.optHeadBand.Enabled = True
		Me.optHeadBand.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optHeadBand.Cursor = System.Windows.Forms.Cursors.Default
		Me.optHeadBand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optHeadBand.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optHeadBand.TabStop = True
		Me.optHeadBand.Checked = False
		Me.optHeadBand.Visible = True
		Me.optHeadBand.Name = "optHeadBand"
		Me.optOpenFaceMask.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optOpenFaceMask.Text = "Open Face Mask"
		Me.optOpenFaceMask.Size = New System.Drawing.Size(148, 16)
		Me.optOpenFaceMask.Location = New System.Drawing.Point(10, 40)
		Me.optOpenFaceMask.TabIndex = 2
		Me.optOpenFaceMask.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optOpenFaceMask.BackColor = System.Drawing.SystemColors.Control
		Me.optOpenFaceMask.CausesValidation = True
		Me.optOpenFaceMask.Enabled = True
		Me.optOpenFaceMask.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optOpenFaceMask.Cursor = System.Windows.Forms.Cursors.Default
		Me.optOpenFaceMask.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optOpenFaceMask.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optOpenFaceMask.TabStop = True
		Me.optOpenFaceMask.Checked = False
		Me.optOpenFaceMask.Visible = True
		Me.optOpenFaceMask.Name = "optOpenFaceMask"
		Me.optFaceMask.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFaceMask.Text = "Face Mask"
		Me.optFaceMask.Size = New System.Drawing.Size(106, 16)
		Me.optFaceMask.Location = New System.Drawing.Point(10, 20)
		Me.optFaceMask.TabIndex = 1
		Me.optFaceMask.Checked = True
		Me.optFaceMask.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFaceMask.BackColor = System.Drawing.SystemColors.Control
		Me.optFaceMask.CausesValidation = True
		Me.optFaceMask.Enabled = True
		Me.optFaceMask.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFaceMask.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFaceMask.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFaceMask.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFaceMask.TabStop = True
		Me.optFaceMask.Visible = True
		Me.optFaceMask.Name = "optFaceMask"
		Me.frmPatientDetails.Text = "Patient Details"
		Me.frmPatientDetails.Size = New System.Drawing.Size(451, 76)
		Me.frmPatientDetails.Location = New System.Drawing.Point(5, 5)
		Me.frmPatientDetails.TabIndex = 0
		Me.frmPatientDetails.BackColor = System.Drawing.SystemColors.Control
		Me.frmPatientDetails.Enabled = True
		Me.frmPatientDetails.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmPatientDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmPatientDetails.Visible = True
		Me.frmPatientDetails.Padding = New System.Windows.Forms.Padding(0)
		Me.frmPatientDetails.Name = "frmPatientDetails"
		Me.txtSex.AutoSize = False
		Me.txtSex.Size = New System.Drawing.Size(54, 19)
		Me.txtSex.Location = New System.Drawing.Point(388, 45)
		Me.txtSex.TabIndex = 42
		Me.txtSex.AcceptsReturn = True
		Me.txtSex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSex.BackColor = System.Drawing.SystemColors.Window
		Me.txtSex.CausesValidation = True
		Me.txtSex.Enabled = True
		Me.txtSex.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSex.HideSelection = True
		Me.txtSex.ReadOnly = False
		Me.txtSex.Maxlength = 0
		Me.txtSex.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSex.MultiLine = False
		Me.txtSex.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSex.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSex.TabStop = True
		Me.txtSex.Visible = True
		Me.txtSex.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSex.Name = "txtSex"
		Me.txtUnits.AutoSize = False
		Me.txtUnits.Size = New System.Drawing.Size(63, 19)
		Me.txtUnits.Location = New System.Drawing.Point(285, 45)
		Me.txtUnits.TabIndex = 41
		Me.txtUnits.AcceptsReturn = True
		Me.txtUnits.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUnits.BackColor = System.Drawing.SystemColors.Window
		Me.txtUnits.CausesValidation = True
		Me.txtUnits.Enabled = True
		Me.txtUnits.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUnits.HideSelection = True
		Me.txtUnits.ReadOnly = False
		Me.txtUnits.Maxlength = 0
		Me.txtUnits.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUnits.MultiLine = False
		Me.txtUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUnits.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUnits.TabStop = True
		Me.txtUnits.Visible = True
		Me.txtUnits.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUnits.Name = "txtUnits"
		Me.txtAge.AutoSize = False
		Me.txtAge.Size = New System.Drawing.Size(46, 19)
		Me.txtAge.Location = New System.Drawing.Point(395, 20)
		Me.txtAge.TabIndex = 39
		Me.txtAge.AcceptsReturn = True
		Me.txtAge.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAge.BackColor = System.Drawing.SystemColors.Window
		Me.txtAge.CausesValidation = True
		Me.txtAge.Enabled = True
		Me.txtAge.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAge.HideSelection = True
		Me.txtAge.ReadOnly = False
		Me.txtAge.Maxlength = 0
		Me.txtAge.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAge.MultiLine = False
		Me.txtAge.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAge.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAge.TabStop = True
		Me.txtAge.Visible = True
		Me.txtAge.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAge.Name = "txtAge"
		Me.txtFileNo.AutoSize = False
		Me.txtFileNo.Size = New System.Drawing.Size(71, 19)
		Me.txtFileNo.Location = New System.Drawing.Point(285, 20)
		Me.txtFileNo.TabIndex = 38
		Me.txtFileNo.AcceptsReturn = True
		Me.txtFileNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFileNo.BackColor = System.Drawing.SystemColors.Window
		Me.txtFileNo.CausesValidation = True
		Me.txtFileNo.Enabled = True
		Me.txtFileNo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFileNo.HideSelection = True
		Me.txtFileNo.ReadOnly = False
		Me.txtFileNo.Maxlength = 0
		Me.txtFileNo.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFileNo.MultiLine = False
		Me.txtFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFileNo.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFileNo.TabStop = True
		Me.txtFileNo.Visible = True
		Me.txtFileNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFileNo.Name = "txtFileNo"
		Me.txtDiagnosis.AutoSize = False
		Me.txtDiagnosis.Size = New System.Drawing.Size(161, 19)
		Me.txtDiagnosis.Location = New System.Drawing.Point(70, 45)
		Me.txtDiagnosis.TabIndex = 40
		Me.txtDiagnosis.AcceptsReturn = True
		Me.txtDiagnosis.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDiagnosis.BackColor = System.Drawing.SystemColors.Window
		Me.txtDiagnosis.CausesValidation = True
		Me.txtDiagnosis.Enabled = True
		Me.txtDiagnosis.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDiagnosis.HideSelection = True
		Me.txtDiagnosis.ReadOnly = False
		Me.txtDiagnosis.Maxlength = 0
		Me.txtDiagnosis.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDiagnosis.MultiLine = False
		Me.txtDiagnosis.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDiagnosis.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDiagnosis.TabStop = True
		Me.txtDiagnosis.Visible = True
		Me.txtDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDiagnosis.Name = "txtDiagnosis"
		Me.txtPatientName.AutoSize = False
		Me.txtPatientName.Size = New System.Drawing.Size(176, 19)
		Me.txtPatientName.Location = New System.Drawing.Point(55, 20)
		Me.txtPatientName.TabIndex = 37
		Me.txtPatientName.AcceptsReturn = True
		Me.txtPatientName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPatientName.BackColor = System.Drawing.SystemColors.Window
		Me.txtPatientName.CausesValidation = True
		Me.txtPatientName.Enabled = True
		Me.txtPatientName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPatientName.HideSelection = True
		Me.txtPatientName.ReadOnly = False
		Me.txtPatientName.Maxlength = 0
		Me.txtPatientName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPatientName.MultiLine = False
		Me.txtPatientName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPatientName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPatientName.TabStop = True
		Me.txtPatientName.Visible = True
		Me.txtPatientName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPatientName.Name = "txtPatientName"
		Me.lblSex.Text = "Sex"
		Me.lblSex.Size = New System.Drawing.Size(31, 16)
		Me.lblSex.Location = New System.Drawing.Point(360, 46)
		Me.lblSex.TabIndex = 50
		Me.lblSex.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSex.BackColor = System.Drawing.SystemColors.Control
		Me.lblSex.Enabled = True
		Me.lblSex.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSex.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSex.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSex.UseMnemonic = True
		Me.lblSex.Visible = True
		Me.lblSex.AutoSize = False
		Me.lblSex.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSex.Name = "lblSex"
		Me.lblUnits.Text = "Units"
		Me.lblUnits.Size = New System.Drawing.Size(36, 16)
		Me.lblUnits.Location = New System.Drawing.Point(240, 46)
		Me.lblUnits.TabIndex = 49
		Me.lblUnits.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUnits.BackColor = System.Drawing.SystemColors.Control
		Me.lblUnits.Enabled = True
		Me.lblUnits.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblUnits.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUnits.UseMnemonic = True
		Me.lblUnits.Visible = True
		Me.lblUnits.AutoSize = False
		Me.lblUnits.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUnits.Name = "lblUnits"
		Me.lblAge.Text = "Age"
		Me.lblAge.Size = New System.Drawing.Size(36, 16)
		Me.lblAge.Location = New System.Drawing.Point(365, 20)
		Me.lblAge.TabIndex = 48
		Me.lblAge.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAge.BackColor = System.Drawing.SystemColors.Control
		Me.lblAge.Enabled = True
		Me.lblAge.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAge.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAge.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAge.UseMnemonic = True
		Me.lblAge.Visible = True
		Me.lblAge.AutoSize = False
		Me.lblAge.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAge.Name = "lblAge"
		Me.lblFileNo.Text = "File No"
		Me.lblFileNo.Size = New System.Drawing.Size(46, 16)
		Me.lblFileNo.Location = New System.Drawing.Point(240, 22)
		Me.lblFileNo.TabIndex = 47
		Me.lblFileNo.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFileNo.BackColor = System.Drawing.SystemColors.Control
		Me.lblFileNo.Enabled = True
		Me.lblFileNo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFileNo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFileNo.UseMnemonic = True
		Me.lblFileNo.Visible = True
		Me.lblFileNo.AutoSize = False
		Me.lblFileNo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFileNo.Name = "lblFileNo"
		Me.lblDiagnosis.Text = "Diagnosis"
		Me.lblDiagnosis.Size = New System.Drawing.Size(56, 16)
		Me.lblDiagnosis.Location = New System.Drawing.Point(10, 45)
		Me.lblDiagnosis.TabIndex = 46
		Me.lblDiagnosis.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDiagnosis.BackColor = System.Drawing.SystemColors.Control
		Me.lblDiagnosis.Enabled = True
		Me.lblDiagnosis.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDiagnosis.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDiagnosis.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDiagnosis.UseMnemonic = True
		Me.lblDiagnosis.Visible = True
		Me.lblDiagnosis.AutoSize = False
		Me.lblDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDiagnosis.Name = "lblDiagnosis"
		Me.lblPatientName.Text = "Patient"
		Me.lblPatientName.Size = New System.Drawing.Size(46, 16)
		Me.lblPatientName.Location = New System.Drawing.Point(10, 20)
		Me.lblPatientName.TabIndex = 45
		Me.lblPatientName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPatientName.BackColor = System.Drawing.SystemColors.Control
		Me.lblPatientName.Enabled = True
		Me.lblPatientName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPatientName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPatientName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPatientName.UseMnemonic = True
		Me.lblPatientName.Visible = True
		Me.lblPatientName.AutoSize = False
		Me.lblPatientName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPatientName.Name = "lblPatientName"
		Me.Controls.Add(txtUidHN)
		Me.Controls.Add(txtTopRightEar)
		Me.Controls.Add(txtTopLeftEar)
		Me.Controls.Add(txtHeadNeckUID)
		Me.Controls.Add(txtOptionChoice)
		Me.Controls.Add(txtFabric)
		Me.Controls.Add(txtWorkOrder)
		Me.Controls.Add(txtData)
		Me.Controls.Add(cmdClose)
		Me.Controls.Add(txtMeasurements)
		Me.Controls.Add(txtMouthRightY)
		Me.Controls.Add(txtxyNeckTopFrontY)
		Me.Controls.Add(txtxyNeckTopFrontX)
		Me.Controls.Add(txtxyChinTopY)
		Me.Controls.Add(txtxyChinTopX)
		Me.Controls.Add(txtCSEarBotHeight)
		Me.Controls.Add(txtCSEarTopHeight)
		Me.Controls.Add(txtCSMouthWidth)
		Me.Controls.Add(txtMouthRightX)
		Me.Controls.Add(txtCSForeHead)
		Me.Controls.Add(txtCSChinAngle)
		Me.Controls.Add(txtCSNeckCircum)
		Me.Controls.Add(txtCSChintoMouth)
		Me.Controls.Add(txtNoseCoverX)
		Me.Controls.Add(txtNoseCoverY)
		Me.Controls.Add(txtEyeWidth)
		Me.Controls.Add(txtDartEndX)
		Me.Controls.Add(txtDartEndY)
		Me.Controls.Add(txtDartStartY)
		Me.Controls.Add(txtDartStartX)
		Me.Controls.Add(txtChinLeftBotY)
		Me.Controls.Add(txtChinLeftBotX)
		Me.Controls.Add(txtNoseBottomY)
		Me.Controls.Add(txtCircumferenceTotal)
		Me.Controls.Add(txtMouthHeight)
		Me.Controls.Add(txtBotOfEyeY)
		Me.Controls.Add(txtBotOfEyeX)
		Me.Controls.Add(txtLowerEarHeight)
		Me.Controls.Add(txtfNum)
		Me.Controls.Add(txtCurrTextFont)
		Me.Controls.Add(txtCurrTextVertJust)
		Me.Controls.Add(txtCurrTextHorizJust)
		Me.Controls.Add(txtCurrTextAspect)
		Me.Controls.Add(txtCurrTextHt)
		Me.Controls.Add(txtCurrentLayer)
		Me.Controls.Add(txtMidToEyeTop)
		Me.Controls.Add(txtLipStrapWidth)
		Me.Controls.Add(txtUidMPD)
		Me.Controls.Add(txtRadiusNo)
		Me.Controls.Add(txtDraw)
		Me.Controls.Add(cmdDraw)
		Me.Controls.Add(frmMeasurements)
		Me.Controls.Add(frmModifications)
		Me.Controls.Add(frmDesignChoices)
		Me.Controls.Add(frmPatientDetails)
		Me.frmMeasurements.Controls.Add(txtChinCollarMin)
		Me.frmMeasurements.Controls.Add(txtRightEarLength)
		Me.frmMeasurements.Controls.Add(txtLeftEarLength)
		Me.frmMeasurements.Controls.Add(txtHeadBandDepth)
		Me.frmMeasurements.Controls.Add(txtLengthOfNose)
		Me.frmMeasurements.Controls.Add(txtTipOfNose)
		Me.frmMeasurements.Controls.Add(cboFabric)
		Me.frmMeasurements.Controls.Add(txtThroatToSternal)
		Me.frmMeasurements.Controls.Add(txtCircOfNeck)
		Me.frmMeasurements.Controls.Add(txtCircChinAngle)
		Me.frmMeasurements.Controls.Add(txtCircEyeBrow)
		Me.frmMeasurements.Controls.Add(txtChinToMouth)
		Me.frmMeasurements.Controls.Add(lblInch11)
		Me.frmMeasurements.Controls.Add(lblChinCollarMin)
		Me.frmMeasurements.Controls.Add(lblRightEarLen)
		Me.frmMeasurements.Controls.Add(lblLeftEarLen)
		Me.frmMeasurements.Controls.Add(lblInch10)
		Me.frmMeasurements.Controls.Add(lblInch9)
		Me.frmMeasurements.Controls.Add(lblInch8)
		Me.frmMeasurements.Controls.Add(lblHeadBandDepth)
		Me.frmMeasurements.Controls.Add(lblFabric)
		Me.frmMeasurements.Controls.Add(lblInch7)
		Me.frmMeasurements.Controls.Add(lblInch6)
		Me.frmMeasurements.Controls.Add(LblLengthOfNose)
		Me.frmMeasurements.Controls.Add(lblAcrossTipOfNose)
		Me.frmMeasurements.Controls.Add(lblInch5)
		Me.frmMeasurements.Controls.Add(lblInch4)
		Me.frmMeasurements.Controls.Add(lblInch3)
		Me.frmMeasurements.Controls.Add(lblInch2)
		Me.frmMeasurements.Controls.Add(lblInch1)
		Me.frmMeasurements.Controls.Add(lblThroatToSternal)
		Me.frmMeasurements.Controls.Add(lblCircOfNeck)
		Me.frmMeasurements.Controls.Add(lblCircChinAngle)
		Me.frmMeasurements.Controls.Add(lblCircEyeBrow)
		Me.frmMeasurements.Controls.Add(lblChinToMouth)
		Me.frmModifications.Controls.Add(chkVelcro)
		Me.frmModifications.Controls.Add(chkEyes)
		Me.frmModifications.Controls.Add(chkRightEarClosed)
		Me.frmModifications.Controls.Add(chkLeftEarClosed)
		Me.frmModifications.Controls.Add(chkEarSize)
		Me.frmModifications.Controls.Add(chkLipStrap)
		Me.frmModifications.Controls.Add(chkOpenHeadMask)
		Me.frmModifications.Controls.Add(chkNeckElastic)
		Me.frmModifications.Controls.Add(chkLining)
		Me.frmModifications.Controls.Add(chkLeftEyeFlap)
		Me.frmModifications.Controls.Add(chkRightEyeFlap)
		Me.frmModifications.Controls.Add(chkZipper)
		Me.frmModifications.Controls.Add(chkNoseCovering)
		Me.frmModifications.Controls.Add(chkLipCovering)
		Me.frmModifications.Controls.Add(chkRightEarFlap)
		Me.frmModifications.Controls.Add(chkLeftEarFlap)
		Me.frmDesignChoices.Controls.Add(optContouredChinCollar)
		Me.frmDesignChoices.Controls.Add(optChinCollar)
		Me.frmDesignChoices.Controls.Add(optModifiedChinStrap)
		Me.frmDesignChoices.Controls.Add(optChinStrap)
		Me.frmDesignChoices.Controls.Add(optHeadBand)
		Me.frmDesignChoices.Controls.Add(optOpenFaceMask)
		Me.frmDesignChoices.Controls.Add(optFaceMask)
		Me.frmPatientDetails.Controls.Add(txtSex)
		Me.frmPatientDetails.Controls.Add(txtUnits)
		Me.frmPatientDetails.Controls.Add(txtAge)
		Me.frmPatientDetails.Controls.Add(txtFileNo)
		Me.frmPatientDetails.Controls.Add(txtDiagnosis)
		Me.frmPatientDetails.Controls.Add(txtPatientName)
		Me.frmPatientDetails.Controls.Add(lblSex)
		Me.frmPatientDetails.Controls.Add(lblUnits)
		Me.frmPatientDetails.Controls.Add(lblAge)
		Me.frmPatientDetails.Controls.Add(lblFileNo)
		Me.frmPatientDetails.Controls.Add(lblDiagnosis)
		Me.frmPatientDetails.Controls.Add(lblPatientName)
		Me.frmMeasurements.ResumeLayout(False)
		Me.frmModifications.ResumeLayout(False)
		Me.frmDesignChoices.ResumeLayout(False)
		Me.frmPatientDetails.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class
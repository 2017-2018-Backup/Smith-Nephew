<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class cltxfdia
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
	Public WithEvents txtFileNoPD As System.Windows.Forms.TextBox
	Public WithEvents txtAgePD As System.Windows.Forms.TextBox
	Public WithEvents cboSexPD As System.Windows.Forms.ComboBox
	Public WithEvents txtNamePD As System.Windows.Forms.TextBox
	Public WithEvents cboDiagnosisPD As System.Windows.Forms.ComboBox
	Public WithEvents _labPatientDetails_0 As System.Windows.Forms.Label
	Public WithEvents _labPatientDetails_1 As System.Windows.Forms.Label
	Public WithEvents _labPatientDetails_2 As System.Windows.Forms.Label
	Public WithEvents _labPatientDetails_3 As System.Windows.Forms.Label
	Public WithEvents _labPatientDetails_4 As System.Windows.Forms.Label
	Public WithEvents frmPatientDetails As System.Windows.Forms.GroupBox
	Public CMDialog1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents txtInitialCADFile As System.Windows.Forms.TextBox
	Public WithEvents txtCurrentCADFile As System.Windows.Forms.TextBox
	Public WithEvents CAD_File As System.Windows.Forms.Button
	Public WithEvents lblCADFile As System.Windows.Forms.Label
	Public WithEvents frmCADFile As System.Windows.Forms.GroupBox
	Public WithEvents txtInvokedFrom As System.Windows.Forms.TextBox
	Public WithEvents Cancel As System.Windows.Forms.Button
	Public WithEvents OK As System.Windows.Forms.Button
	Public WithEvents cboDesignerWOD As System.Windows.Forms.ComboBox
	Public WithEvents txtOrderDateWOD As System.Windows.Forms.TextBox
	Public WithEvents cboUnitsWOD As System.Windows.Forms.ComboBox
	Public WithEvents txtWorkOrderWOD As System.Windows.Forms.TextBox
	Public WithEvents _labWODetails_3 As System.Windows.Forms.Label
	Public WithEvents _labWODetails_2 As System.Windows.Forms.Label
	Public WithEvents _labWODetails_1 As System.Windows.Forms.Label
	Public WithEvents _labWODetails_0 As System.Windows.Forms.Label
	Public WithEvents frmWODetails As System.Windows.Forms.GroupBox
	Public WithEvents txtMessage As System.Windows.Forms.TextBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents txtWO_TXF As System.Windows.Forms.TextBox
	Public WithEvents txtWorkOrder As System.Windows.Forms.TextBox
	Public WithEvents txtUidMPD As System.Windows.Forms.TextBox
	Public WithEvents txtTemplateEngineer As System.Windows.Forms.TextBox
	Public WithEvents txtOrderDate As System.Windows.Forms.TextBox
	Public WithEvents txtMPDwo As System.Windows.Forms.TextBox
	Public WithEvents txtAge As System.Windows.Forms.TextBox
	Public WithEvents txtSex As System.Windows.Forms.TextBox
	Public WithEvents txtDiagnosis As System.Windows.Forms.TextBox
	Public WithEvents txtPatientName As System.Windows.Forms.TextBox
	Public WithEvents txtUnits As System.Windows.Forms.TextBox
	Public WithEvents txtFileNo As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents labPatientDetails As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents labWODetails As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(cltxfdia))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.frmPatientDetails = New System.Windows.Forms.GroupBox
		Me.txtFileNoPD = New System.Windows.Forms.TextBox
		Me.txtAgePD = New System.Windows.Forms.TextBox
		Me.cboSexPD = New System.Windows.Forms.ComboBox
		Me.txtNamePD = New System.Windows.Forms.TextBox
		Me.cboDiagnosisPD = New System.Windows.Forms.ComboBox
		Me._labPatientDetails_0 = New System.Windows.Forms.Label
		Me._labPatientDetails_1 = New System.Windows.Forms.Label
		Me._labPatientDetails_2 = New System.Windows.Forms.Label
		Me._labPatientDetails_3 = New System.Windows.Forms.Label
		Me._labPatientDetails_4 = New System.Windows.Forms.Label
		Me.CMDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.txtInitialCADFile = New System.Windows.Forms.TextBox
		Me.txtCurrentCADFile = New System.Windows.Forms.TextBox
		Me.frmCADFile = New System.Windows.Forms.GroupBox
		Me.CAD_File = New System.Windows.Forms.Button
		Me.lblCADFile = New System.Windows.Forms.Label
		Me.txtInvokedFrom = New System.Windows.Forms.TextBox
		Me.Cancel = New System.Windows.Forms.Button
		Me.OK = New System.Windows.Forms.Button
		Me.frmWODetails = New System.Windows.Forms.GroupBox
		Me.cboDesignerWOD = New System.Windows.Forms.ComboBox
		Me.txtOrderDateWOD = New System.Windows.Forms.TextBox
		Me.cboUnitsWOD = New System.Windows.Forms.ComboBox
		Me.txtWorkOrderWOD = New System.Windows.Forms.TextBox
		Me._labWODetails_3 = New System.Windows.Forms.Label
		Me._labWODetails_2 = New System.Windows.Forms.Label
		Me._labWODetails_1 = New System.Windows.Forms.Label
		Me._labWODetails_0 = New System.Windows.Forms.Label
		Me.txtMessage = New System.Windows.Forms.TextBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.txtWO_TXF = New System.Windows.Forms.TextBox
		Me.txtWorkOrder = New System.Windows.Forms.TextBox
		Me.txtUidMPD = New System.Windows.Forms.TextBox
		Me.txtTemplateEngineer = New System.Windows.Forms.TextBox
		Me.txtOrderDate = New System.Windows.Forms.TextBox
		Me.txtMPDwo = New System.Windows.Forms.TextBox
		Me.txtAge = New System.Windows.Forms.TextBox
		Me.txtSex = New System.Windows.Forms.TextBox
		Me.txtDiagnosis = New System.Windows.Forms.TextBox
		Me.txtPatientName = New System.Windows.Forms.TextBox
		Me.txtUnits = New System.Windows.Forms.TextBox
		Me.txtFileNo = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.labPatientDetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.labWODetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.frmPatientDetails.SuspendLayout()
		Me.frmCADFile.SuspendLayout()
		Me.frmWODetails.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Patient and Work Order Details"
		Me.ClientSize = New System.Drawing.Size(429, 459)
		Me.Location = New System.Drawing.Point(47, 133)
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
		Me.Name = "cltxfdia"
		Me.frmPatientDetails.Text = "Patient Details"
		Me.frmPatientDetails.Size = New System.Drawing.Size(414, 105)
		Me.frmPatientDetails.Location = New System.Drawing.Point(8, 8)
		Me.frmPatientDetails.TabIndex = 33
		Me.frmPatientDetails.BackColor = System.Drawing.SystemColors.Control
		Me.frmPatientDetails.Enabled = True
		Me.frmPatientDetails.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmPatientDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmPatientDetails.Visible = True
		Me.frmPatientDetails.Padding = New System.Windows.Forms.Padding(0)
		Me.frmPatientDetails.Name = "frmPatientDetails"
		Me.txtFileNoPD.AutoSize = False
		Me.txtFileNoPD.Size = New System.Drawing.Size(81, 19)
		Me.txtFileNoPD.Location = New System.Drawing.Point(76, 16)
		Me.txtFileNoPD.TabIndex = 38
		Me.txtFileNoPD.AcceptsReturn = True
		Me.txtFileNoPD.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFileNoPD.BackColor = System.Drawing.SystemColors.Window
		Me.txtFileNoPD.CausesValidation = True
		Me.txtFileNoPD.Enabled = True
		Me.txtFileNoPD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFileNoPD.HideSelection = True
		Me.txtFileNoPD.ReadOnly = False
		Me.txtFileNoPD.Maxlength = 0
		Me.txtFileNoPD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFileNoPD.MultiLine = False
		Me.txtFileNoPD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFileNoPD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFileNoPD.TabStop = True
		Me.txtFileNoPD.Visible = True
		Me.txtFileNoPD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFileNoPD.Name = "txtFileNoPD"
		Me.txtAgePD.AutoSize = False
		Me.txtAgePD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtAgePD.Size = New System.Drawing.Size(29, 19)
		Me.txtAgePD.Location = New System.Drawing.Point(332, 43)
		Me.txtAgePD.TabIndex = 37
		Me.txtAgePD.AcceptsReturn = True
		Me.txtAgePD.BackColor = System.Drawing.SystemColors.Window
		Me.txtAgePD.CausesValidation = True
		Me.txtAgePD.Enabled = True
		Me.txtAgePD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAgePD.HideSelection = True
		Me.txtAgePD.ReadOnly = False
		Me.txtAgePD.Maxlength = 0
		Me.txtAgePD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAgePD.MultiLine = False
		Me.txtAgePD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAgePD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAgePD.TabStop = True
		Me.txtAgePD.Visible = True
		Me.txtAgePD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAgePD.Name = "txtAgePD"
		Me.cboSexPD.Size = New System.Drawing.Size(65, 21)
		Me.cboSexPD.Location = New System.Drawing.Point(332, 70)
		Me.cboSexPD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboSexPD.TabIndex = 36
		Me.cboSexPD.BackColor = System.Drawing.SystemColors.Window
		Me.cboSexPD.CausesValidation = True
		Me.cboSexPD.Enabled = True
		Me.cboSexPD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSexPD.IntegralHeight = True
		Me.cboSexPD.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSexPD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSexPD.Sorted = False
		Me.cboSexPD.TabStop = True
		Me.cboSexPD.Visible = True
		Me.cboSexPD.Name = "cboSexPD"
		Me.txtNamePD.AutoSize = False
		Me.txtNamePD.Size = New System.Drawing.Size(213, 19)
		Me.txtNamePD.Location = New System.Drawing.Point(76, 43)
		Me.txtNamePD.TabIndex = 35
		Me.txtNamePD.AcceptsReturn = True
		Me.txtNamePD.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNamePD.BackColor = System.Drawing.SystemColors.Window
		Me.txtNamePD.CausesValidation = True
		Me.txtNamePD.Enabled = True
		Me.txtNamePD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNamePD.HideSelection = True
		Me.txtNamePD.ReadOnly = False
		Me.txtNamePD.Maxlength = 0
		Me.txtNamePD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNamePD.MultiLine = False
		Me.txtNamePD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNamePD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNamePD.TabStop = True
		Me.txtNamePD.Visible = True
		Me.txtNamePD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNamePD.Name = "txtNamePD"
		Me.cboDiagnosisPD.Size = New System.Drawing.Size(213, 21)
		Me.cboDiagnosisPD.Location = New System.Drawing.Point(76, 70)
		Me.cboDiagnosisPD.TabIndex = 34
		Me.cboDiagnosisPD.Text = "cboDiagnosisPD"
		Me.cboDiagnosisPD.BackColor = System.Drawing.SystemColors.Window
		Me.cboDiagnosisPD.CausesValidation = True
		Me.cboDiagnosisPD.Enabled = True
		Me.cboDiagnosisPD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDiagnosisPD.IntegralHeight = True
		Me.cboDiagnosisPD.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDiagnosisPD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDiagnosisPD.Sorted = False
		Me.cboDiagnosisPD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDiagnosisPD.TabStop = True
		Me.cboDiagnosisPD.Visible = True
		Me.cboDiagnosisPD.Name = "cboDiagnosisPD"
		Me._labPatientDetails_0.Text = "File No:"
		Me._labPatientDetails_0.Size = New System.Drawing.Size(45, 13)
		Me._labPatientDetails_0.Location = New System.Drawing.Point(24, 19)
		Me._labPatientDetails_0.TabIndex = 43
		Me._labPatientDetails_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labPatientDetails_0.BackColor = System.Drawing.SystemColors.Control
		Me._labPatientDetails_0.Enabled = True
		Me._labPatientDetails_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labPatientDetails_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._labPatientDetails_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labPatientDetails_0.UseMnemonic = True
		Me._labPatientDetails_0.Visible = True
		Me._labPatientDetails_0.AutoSize = False
		Me._labPatientDetails_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labPatientDetails_0.Name = "_labPatientDetails_0"
		Me._labPatientDetails_1.Text = "Name:"
		Me._labPatientDetails_1.Size = New System.Drawing.Size(41, 13)
		Me._labPatientDetails_1.Location = New System.Drawing.Point(32, 46)
		Me._labPatientDetails_1.TabIndex = 42
		Me._labPatientDetails_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labPatientDetails_1.BackColor = System.Drawing.SystemColors.Control
		Me._labPatientDetails_1.Enabled = True
		Me._labPatientDetails_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labPatientDetails_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._labPatientDetails_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labPatientDetails_1.UseMnemonic = True
		Me._labPatientDetails_1.Visible = True
		Me._labPatientDetails_1.AutoSize = False
		Me._labPatientDetails_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labPatientDetails_1.Name = "_labPatientDetails_1"
		Me._labPatientDetails_2.Text = "Diagnosis:"
		Me._labPatientDetails_2.Size = New System.Drawing.Size(61, 13)
		Me._labPatientDetails_2.Location = New System.Drawing.Point(8, 73)
		Me._labPatientDetails_2.TabIndex = 41
		Me._labPatientDetails_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labPatientDetails_2.BackColor = System.Drawing.SystemColors.Control
		Me._labPatientDetails_2.Enabled = True
		Me._labPatientDetails_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labPatientDetails_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._labPatientDetails_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labPatientDetails_2.UseMnemonic = True
		Me._labPatientDetails_2.Visible = True
		Me._labPatientDetails_2.AutoSize = False
		Me._labPatientDetails_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labPatientDetails_2.Name = "_labPatientDetails_2"
		Me._labPatientDetails_3.Text = "Age:"
		Me._labPatientDetails_3.Size = New System.Drawing.Size(33, 13)
		Me._labPatientDetails_3.Location = New System.Drawing.Point(300, 46)
		Me._labPatientDetails_3.TabIndex = 40
		Me._labPatientDetails_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labPatientDetails_3.BackColor = System.Drawing.SystemColors.Control
		Me._labPatientDetails_3.Enabled = True
		Me._labPatientDetails_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labPatientDetails_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._labPatientDetails_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labPatientDetails_3.UseMnemonic = True
		Me._labPatientDetails_3.Visible = True
		Me._labPatientDetails_3.AutoSize = False
		Me._labPatientDetails_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labPatientDetails_3.Name = "_labPatientDetails_3"
		Me._labPatientDetails_4.Text = "Sex:"
		Me._labPatientDetails_4.Size = New System.Drawing.Size(33, 13)
		Me._labPatientDetails_4.Location = New System.Drawing.Point(300, 73)
		Me._labPatientDetails_4.TabIndex = 39
		Me._labPatientDetails_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labPatientDetails_4.BackColor = System.Drawing.SystemColors.Control
		Me._labPatientDetails_4.Enabled = True
		Me._labPatientDetails_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labPatientDetails_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._labPatientDetails_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labPatientDetails_4.UseMnemonic = True
		Me._labPatientDetails_4.Visible = True
		Me._labPatientDetails_4.AutoSize = False
		Me._labPatientDetails_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labPatientDetails_4.Name = "_labPatientDetails_4"
		Me.CMDialog1Open.Title = "Open CAD File"
		Me.txtInitialCADFile.AutoSize = False
		Me.txtInitialCADFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtInitialCADFile.Size = New System.Drawing.Size(97, 21)
		Me.txtInitialCADFile.Location = New System.Drawing.Point(444, 64)
		Me.txtInitialCADFile.TabIndex = 32
		Me.txtInitialCADFile.Text = "txtInitialCADFile"
		Me.txtInitialCADFile.Visible = False
		Me.txtInitialCADFile.AcceptsReturn = True
		Me.txtInitialCADFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtInitialCADFile.CausesValidation = True
		Me.txtInitialCADFile.Enabled = True
		Me.txtInitialCADFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInitialCADFile.HideSelection = True
		Me.txtInitialCADFile.ReadOnly = False
		Me.txtInitialCADFile.Maxlength = 0
		Me.txtInitialCADFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInitialCADFile.MultiLine = False
		Me.txtInitialCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInitialCADFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInitialCADFile.TabStop = True
		Me.txtInitialCADFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtInitialCADFile.Name = "txtInitialCADFile"
		Me.txtCurrentCADFile.AutoSize = False
		Me.txtCurrentCADFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtCurrentCADFile.Size = New System.Drawing.Size(97, 21)
		Me.txtCurrentCADFile.Location = New System.Drawing.Point(444, 40)
		Me.txtCurrentCADFile.TabIndex = 31
		Me.txtCurrentCADFile.Text = "txtCurrentCADFile"
		Me.txtCurrentCADFile.Visible = False
		Me.txtCurrentCADFile.AcceptsReturn = True
		Me.txtCurrentCADFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtCurrentCADFile.CausesValidation = True
		Me.txtCurrentCADFile.Enabled = True
		Me.txtCurrentCADFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCurrentCADFile.HideSelection = True
		Me.txtCurrentCADFile.ReadOnly = False
		Me.txtCurrentCADFile.Maxlength = 0
		Me.txtCurrentCADFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCurrentCADFile.MultiLine = False
		Me.txtCurrentCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCurrentCADFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCurrentCADFile.TabStop = True
		Me.txtCurrentCADFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrentCADFile.Name = "txtCurrentCADFile"
		Me.frmCADFile.Text = "Drafix CAD File"
		Me.frmCADFile.Size = New System.Drawing.Size(413, 55)
		Me.frmCADFile.Location = New System.Drawing.Point(8, 214)
		Me.frmCADFile.TabIndex = 28
		Me.frmCADFile.BackColor = System.Drawing.SystemColors.Control
		Me.frmCADFile.Enabled = True
		Me.frmCADFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmCADFile.Visible = True
		Me.frmCADFile.Padding = New System.Windows.Forms.Padding(0)
		Me.frmCADFile.Name = "frmCADFile"
		Me.CAD_File.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.CAD_File
		Me.CAD_File.Text = "CAD File ..."
		Me.CAD_File.Size = New System.Drawing.Size(85, 25)
		Me.CAD_File.Location = New System.Drawing.Point(316, 22)
		Me.CAD_File.TabIndex = 30
		Me.CAD_File.BackColor = System.Drawing.SystemColors.Control
		Me.CAD_File.CausesValidation = True
		Me.CAD_File.Enabled = True
		Me.CAD_File.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CAD_File.Cursor = System.Windows.Forms.Cursors.Default
		Me.CAD_File.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CAD_File.TabStop = True
		Me.CAD_File.Name = "CAD_File"
		Me.lblCADFile.Size = New System.Drawing.Size(297, 13)
		Me.lblCADFile.Location = New System.Drawing.Point(12, 24)
		Me.lblCADFile.TabIndex = 29
		Me.lblCADFile.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCADFile.BackColor = System.Drawing.SystemColors.Control
		Me.lblCADFile.Enabled = True
		Me.lblCADFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCADFile.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCADFile.UseMnemonic = True
		Me.lblCADFile.Visible = True
		Me.lblCADFile.AutoSize = False
		Me.lblCADFile.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCADFile.Name = "lblCADFile"
		Me.txtInvokedFrom.AutoSize = False
		Me.txtInvokedFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtInvokedFrom.Size = New System.Drawing.Size(97, 21)
		Me.txtInvokedFrom.Location = New System.Drawing.Point(444, 16)
		Me.txtInvokedFrom.TabIndex = 0
		Me.txtInvokedFrom.Text = "txtInvokedFrom"
		Me.txtInvokedFrom.Visible = False
		Me.txtInvokedFrom.AcceptsReturn = True
		Me.txtInvokedFrom.BackColor = System.Drawing.SystemColors.Window
		Me.txtInvokedFrom.CausesValidation = True
		Me.txtInvokedFrom.Enabled = True
		Me.txtInvokedFrom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInvokedFrom.HideSelection = True
		Me.txtInvokedFrom.ReadOnly = False
		Me.txtInvokedFrom.Maxlength = 0
		Me.txtInvokedFrom.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInvokedFrom.MultiLine = False
		Me.txtInvokedFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInvokedFrom.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInvokedFrom.TabStop = True
		Me.txtInvokedFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtInvokedFrom.Name = "txtInvokedFrom"
		Me.Cancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Cancel.Text = "Cancel"
		Me.Cancel.Size = New System.Drawing.Size(85, 25)
		Me.Cancel.Location = New System.Drawing.Point(216, 416)
		Me.Cancel.TabIndex = 2
		Me.Cancel.BackColor = System.Drawing.SystemColors.Control
		Me.Cancel.CausesValidation = True
		Me.Cancel.Enabled = True
		Me.Cancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Cancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Cancel.TabStop = True
		Me.Cancel.Name = "Cancel"
		Me.OK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OK.Text = "OK"
		Me.OK.Size = New System.Drawing.Size(85, 25)
		Me.OK.Location = New System.Drawing.Point(112, 416)
		Me.OK.TabIndex = 19
		Me.OK.BackColor = System.Drawing.SystemColors.Control
		Me.OK.CausesValidation = True
		Me.OK.Enabled = True
		Me.OK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OK.Cursor = System.Windows.Forms.Cursors.Default
		Me.OK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OK.TabStop = True
		Me.OK.Name = "OK"
		Me.frmWODetails.Text = "Work Order Details"
		Me.frmWODetails.Size = New System.Drawing.Size(414, 88)
		Me.frmWODetails.Location = New System.Drawing.Point(8, 120)
		Me.frmWODetails.TabIndex = 10
		Me.frmWODetails.BackColor = System.Drawing.SystemColors.Control
		Me.frmWODetails.Enabled = True
		Me.frmWODetails.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmWODetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmWODetails.Visible = True
		Me.frmWODetails.Padding = New System.Windows.Forms.Padding(0)
		Me.frmWODetails.Name = "frmWODetails"
		Me.cboDesignerWOD.Size = New System.Drawing.Size(133, 21)
		Me.cboDesignerWOD.Location = New System.Drawing.Point(76, 48)
		Me.cboDesignerWOD.TabIndex = 14
		Me.cboDesignerWOD.BackColor = System.Drawing.SystemColors.Window
		Me.cboDesignerWOD.CausesValidation = True
		Me.cboDesignerWOD.Enabled = True
		Me.cboDesignerWOD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDesignerWOD.IntegralHeight = True
		Me.cboDesignerWOD.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDesignerWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDesignerWOD.Sorted = False
		Me.cboDesignerWOD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDesignerWOD.TabStop = True
		Me.cboDesignerWOD.Visible = True
		Me.cboDesignerWOD.Name = "cboDesignerWOD"
		Me.txtOrderDateWOD.AutoSize = False
		Me.txtOrderDateWOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtOrderDateWOD.Size = New System.Drawing.Size(80, 19)
		Me.txtOrderDateWOD.Location = New System.Drawing.Point(207, 21)
		Me.txtOrderDateWOD.Maxlength = 10
		Me.txtOrderDateWOD.TabIndex = 12
		Me.txtOrderDateWOD.AcceptsReturn = True
		Me.txtOrderDateWOD.BackColor = System.Drawing.SystemColors.Window
		Me.txtOrderDateWOD.CausesValidation = True
		Me.txtOrderDateWOD.Enabled = True
		Me.txtOrderDateWOD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrderDateWOD.HideSelection = True
		Me.txtOrderDateWOD.ReadOnly = False
		Me.txtOrderDateWOD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrderDateWOD.MultiLine = False
		Me.txtOrderDateWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrderDateWOD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrderDateWOD.TabStop = True
		Me.txtOrderDateWOD.Visible = True
		Me.txtOrderDateWOD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrderDateWOD.Name = "txtOrderDateWOD"
		Me.cboUnitsWOD.Size = New System.Drawing.Size(65, 21)
		Me.cboUnitsWOD.Location = New System.Drawing.Point(332, 21)
		Me.cboUnitsWOD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboUnitsWOD.TabIndex = 13
		Me.cboUnitsWOD.BackColor = System.Drawing.SystemColors.Window
		Me.cboUnitsWOD.CausesValidation = True
		Me.cboUnitsWOD.Enabled = True
		Me.cboUnitsWOD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboUnitsWOD.IntegralHeight = True
		Me.cboUnitsWOD.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboUnitsWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboUnitsWOD.Sorted = False
		Me.cboUnitsWOD.TabStop = True
		Me.cboUnitsWOD.Visible = True
		Me.cboUnitsWOD.Name = "cboUnitsWOD"
		Me.txtWorkOrderWOD.AutoSize = False
		Me.txtWorkOrderWOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtWorkOrderWOD.Size = New System.Drawing.Size(57, 19)
		Me.txtWorkOrderWOD.Location = New System.Drawing.Point(76, 21)
		Me.txtWorkOrderWOD.TabIndex = 11
		Me.txtWorkOrderWOD.AcceptsReturn = True
		Me.txtWorkOrderWOD.BackColor = System.Drawing.SystemColors.Window
		Me.txtWorkOrderWOD.CausesValidation = True
		Me.txtWorkOrderWOD.Enabled = True
		Me.txtWorkOrderWOD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWorkOrderWOD.HideSelection = True
		Me.txtWorkOrderWOD.ReadOnly = False
		Me.txtWorkOrderWOD.Maxlength = 0
		Me.txtWorkOrderWOD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWorkOrderWOD.MultiLine = False
		Me.txtWorkOrderWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWorkOrderWOD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWorkOrderWOD.TabStop = True
		Me.txtWorkOrderWOD.Visible = True
		Me.txtWorkOrderWOD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWorkOrderWOD.Name = "txtWorkOrderWOD"
		Me._labWODetails_3.Text = "Units:"
		Me._labWODetails_3.Size = New System.Drawing.Size(41, 16)
		Me._labWODetails_3.Location = New System.Drawing.Point(292, 24)
		Me._labWODetails_3.TabIndex = 20
		Me._labWODetails_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labWODetails_3.BackColor = System.Drawing.SystemColors.Control
		Me._labWODetails_3.Enabled = True
		Me._labWODetails_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labWODetails_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._labWODetails_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labWODetails_3.UseMnemonic = True
		Me._labWODetails_3.Visible = True
		Me._labWODetails_3.AutoSize = False
		Me._labWODetails_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labWODetails_3.Name = "_labWODetails_3"
		Me._labWODetails_2.Text = "Temp Date:"
		Me._labWODetails_2.Size = New System.Drawing.Size(73, 16)
		Me._labWODetails_2.Location = New System.Drawing.Point(140, 24)
		Me._labWODetails_2.TabIndex = 27
		Me._labWODetails_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labWODetails_2.BackColor = System.Drawing.SystemColors.Control
		Me._labWODetails_2.Enabled = True
		Me._labWODetails_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labWODetails_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._labWODetails_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labWODetails_2.UseMnemonic = True
		Me._labWODetails_2.Visible = True
		Me._labWODetails_2.AutoSize = False
		Me._labWODetails_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labWODetails_2.Name = "_labWODetails_2"
		Me._labWODetails_1.Text = "Designer:"
		Me._labWODetails_1.Size = New System.Drawing.Size(57, 16)
		Me._labWODetails_1.Location = New System.Drawing.Point(16, 51)
		Me._labWODetails_1.TabIndex = 26
		Me._labWODetails_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labWODetails_1.BackColor = System.Drawing.SystemColors.Control
		Me._labWODetails_1.Enabled = True
		Me._labWODetails_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labWODetails_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._labWODetails_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labWODetails_1.UseMnemonic = True
		Me._labWODetails_1.Visible = True
		Me._labWODetails_1.AutoSize = False
		Me._labWODetails_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labWODetails_1.Name = "_labWODetails_1"
		Me._labWODetails_0.Text = "WorkOrder:"
		Me._labWODetails_0.Size = New System.Drawing.Size(73, 16)
		Me._labWODetails_0.Location = New System.Drawing.Point(4, 24)
		Me._labWODetails_0.TabIndex = 25
		Me._labWODetails_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._labWODetails_0.BackColor = System.Drawing.SystemColors.Control
		Me._labWODetails_0.Enabled = True
		Me._labWODetails_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._labWODetails_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._labWODetails_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._labWODetails_0.UseMnemonic = True
		Me._labWODetails_0.Visible = True
		Me._labWODetails_0.AutoSize = False
		Me._labWODetails_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._labWODetails_0.Name = "_labWODetails_0"
		Me.txtMessage.AutoSize = False
		Me.txtMessage.Size = New System.Drawing.Size(413, 112)
		Me.txtMessage.Location = New System.Drawing.Point(8, 292)
		Me.txtMessage.MultiLine = True
		Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtMessage.TabIndex = 21
		Me.txtMessage.AcceptsReturn = True
		Me.txtMessage.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMessage.BackColor = System.Drawing.SystemColors.Window
		Me.txtMessage.CausesValidation = True
		Me.txtMessage.Enabled = True
		Me.txtMessage.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMessage.HideSelection = True
		Me.txtMessage.ReadOnly = False
		Me.txtMessage.Maxlength = 0
		Me.txtMessage.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMessage.TabStop = True
		Me.txtMessage.Visible = True
		Me.txtMessage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMessage.Name = "txtMessage"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.txtWO_TXF.AutoSize = False
		Me.txtWO_TXF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtWO_TXF.Size = New System.Drawing.Size(97, 21)
		Me.txtWO_TXF.Location = New System.Drawing.Point(444, 137)
		Me.txtWO_TXF.TabIndex = 23
		Me.txtWO_TXF.Text = "txtWO_TXF"
		Me.txtWO_TXF.Visible = False
		Me.txtWO_TXF.AcceptsReturn = True
		Me.txtWO_TXF.BackColor = System.Drawing.SystemColors.Window
		Me.txtWO_TXF.CausesValidation = True
		Me.txtWO_TXF.Enabled = True
		Me.txtWO_TXF.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWO_TXF.HideSelection = True
		Me.txtWO_TXF.ReadOnly = False
		Me.txtWO_TXF.Maxlength = 0
		Me.txtWO_TXF.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWO_TXF.MultiLine = False
		Me.txtWO_TXF.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWO_TXF.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWO_TXF.TabStop = True
		Me.txtWO_TXF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtWO_TXF.Name = "txtWO_TXF"
		Me.txtWorkOrder.AutoSize = False
		Me.txtWorkOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtWorkOrder.Size = New System.Drawing.Size(97, 21)
		Me.txtWorkOrder.Location = New System.Drawing.Point(444, 113)
		Me.txtWorkOrder.TabIndex = 6
		Me.txtWorkOrder.Text = "txtWorkOrder"
		Me.txtWorkOrder.Visible = False
		Me.txtWorkOrder.AcceptsReturn = True
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
		Me.txtUidMPD.AutoSize = False
		Me.txtUidMPD.Size = New System.Drawing.Size(97, 21)
		Me.txtUidMPD.Location = New System.Drawing.Point(444, 401)
		Me.txtUidMPD.TabIndex = 17
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
		Me.txtTemplateEngineer.AutoSize = False
		Me.txtTemplateEngineer.Size = New System.Drawing.Size(97, 21)
		Me.txtTemplateEngineer.Location = New System.Drawing.Point(444, 377)
		Me.txtTemplateEngineer.TabIndex = 16
		Me.txtTemplateEngineer.Text = "txtTemplateEngineer"
		Me.txtTemplateEngineer.Visible = False
		Me.txtTemplateEngineer.AcceptsReturn = True
		Me.txtTemplateEngineer.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTemplateEngineer.BackColor = System.Drawing.SystemColors.Window
		Me.txtTemplateEngineer.CausesValidation = True
		Me.txtTemplateEngineer.Enabled = True
		Me.txtTemplateEngineer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTemplateEngineer.HideSelection = True
		Me.txtTemplateEngineer.ReadOnly = False
		Me.txtTemplateEngineer.Maxlength = 0
		Me.txtTemplateEngineer.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTemplateEngineer.MultiLine = False
		Me.txtTemplateEngineer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTemplateEngineer.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTemplateEngineer.TabStop = True
		Me.txtTemplateEngineer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtTemplateEngineer.Name = "txtTemplateEngineer"
		Me.txtOrderDate.AutoSize = False
		Me.txtOrderDate.Size = New System.Drawing.Size(97, 21)
		Me.txtOrderDate.Location = New System.Drawing.Point(444, 353)
		Me.txtOrderDate.TabIndex = 15
		Me.txtOrderDate.Text = "txtOrderDate"
		Me.txtOrderDate.Visible = False
		Me.txtOrderDate.AcceptsReturn = True
		Me.txtOrderDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOrderDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtOrderDate.CausesValidation = True
		Me.txtOrderDate.Enabled = True
		Me.txtOrderDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrderDate.HideSelection = True
		Me.txtOrderDate.ReadOnly = False
		Me.txtOrderDate.Maxlength = 0
		Me.txtOrderDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrderDate.MultiLine = False
		Me.txtOrderDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrderDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrderDate.TabStop = True
		Me.txtOrderDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtOrderDate.Name = "txtOrderDate"
		Me.txtMPDwo.AutoSize = False
		Me.txtMPDwo.Size = New System.Drawing.Size(97, 21)
		Me.txtMPDwo.Location = New System.Drawing.Point(444, 329)
		Me.txtMPDwo.TabIndex = 9
		Me.txtMPDwo.Text = "txtMPDwo"
		Me.txtMPDwo.Visible = False
		Me.txtMPDwo.AcceptsReturn = True
		Me.txtMPDwo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMPDwo.BackColor = System.Drawing.SystemColors.Window
		Me.txtMPDwo.CausesValidation = True
		Me.txtMPDwo.Enabled = True
		Me.txtMPDwo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMPDwo.HideSelection = True
		Me.txtMPDwo.ReadOnly = False
		Me.txtMPDwo.Maxlength = 0
		Me.txtMPDwo.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMPDwo.MultiLine = False
		Me.txtMPDwo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMPDwo.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMPDwo.TabStop = True
		Me.txtMPDwo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtMPDwo.Name = "txtMPDwo"
		Me.txtAge.AutoSize = False
		Me.txtAge.Size = New System.Drawing.Size(97, 21)
		Me.txtAge.Location = New System.Drawing.Point(444, 233)
		Me.txtAge.TabIndex = 8
		Me.txtAge.Text = "txtAge"
		Me.txtAge.Visible = False
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
		Me.txtAge.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtAge.Name = "txtAge"
		Me.txtSex.AutoSize = False
		Me.txtSex.Size = New System.Drawing.Size(97, 21)
		Me.txtSex.Location = New System.Drawing.Point(444, 281)
		Me.txtSex.TabIndex = 7
		Me.txtSex.Text = "txtSex"
		Me.txtSex.Visible = False
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
		Me.txtSex.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtSex.Name = "txtSex"
		Me.txtDiagnosis.AutoSize = False
		Me.txtDiagnosis.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtDiagnosis.Size = New System.Drawing.Size(97, 21)
		Me.txtDiagnosis.Location = New System.Drawing.Point(444, 305)
		Me.txtDiagnosis.TabIndex = 5
		Me.txtDiagnosis.Text = "txtDiagnosis"
		Me.txtDiagnosis.Visible = False
		Me.txtDiagnosis.AcceptsReturn = True
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
		Me.txtDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDiagnosis.Name = "txtDiagnosis"
		Me.txtPatientName.AutoSize = False
		Me.txtPatientName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtPatientName.Size = New System.Drawing.Size(97, 21)
		Me.txtPatientName.Location = New System.Drawing.Point(444, 209)
		Me.txtPatientName.TabIndex = 4
		Me.txtPatientName.Text = "txtPatientName"
		Me.txtPatientName.Visible = False
		Me.txtPatientName.AcceptsReturn = True
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
		Me.txtPatientName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtPatientName.Name = "txtPatientName"
		Me.txtUnits.AutoSize = False
		Me.txtUnits.Size = New System.Drawing.Size(97, 21)
		Me.txtUnits.Location = New System.Drawing.Point(444, 257)
		Me.txtUnits.TabIndex = 3
		Me.txtUnits.Text = "txtUnits"
		Me.txtUnits.Visible = False
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
		Me.txtUnits.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtUnits.Name = "txtUnits"
		Me.txtFileNo.AutoSize = False
		Me.txtFileNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtFileNo.Size = New System.Drawing.Size(97, 21)
		Me.txtFileNo.Location = New System.Drawing.Point(444, 185)
		Me.txtFileNo.TabIndex = 1
		Me.txtFileNo.Text = "txtFileNo"
		Me.txtFileNo.Visible = False
		Me.txtFileNo.AcceptsReturn = True
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
		Me.txtFileNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtFileNo.Name = "txtFileNo"
		Me.Label2.Text = "Messages and Errors"
		Me.Label2.Size = New System.Drawing.Size(121, 16)
		Me.Label2.Location = New System.Drawing.Point(8, 272)
		Me.Label2.TabIndex = 22
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Window
		Me._Label1_0.Text = "imageABLE"
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Label1_0.Size = New System.Drawing.Size(85, 13)
		Me._Label1_0.Location = New System.Drawing.Point(456, 95)
		Me._Label1_0.TabIndex = 24
		Me._Label1_0.Visible = False
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.Enabled = True
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Window
		Me._Label1_1.Text = "PatientDetails"
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Label1_1.Size = New System.Drawing.Size(81, 13)
		Me._Label1_1.Location = New System.Drawing.Point(456, 167)
		Me._Label1_1.TabIndex = 18
		Me._Label1_1.Visible = False
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me.Controls.Add(frmPatientDetails)
		Me.Controls.Add(txtInitialCADFile)
		Me.Controls.Add(txtCurrentCADFile)
		Me.Controls.Add(frmCADFile)
		Me.Controls.Add(txtInvokedFrom)
		Me.Controls.Add(Cancel)
		Me.Controls.Add(OK)
		Me.Controls.Add(frmWODetails)
		Me.Controls.Add(txtMessage)
		Me.Controls.Add(txtWO_TXF)
		Me.Controls.Add(txtWorkOrder)
		Me.Controls.Add(txtUidMPD)
		Me.Controls.Add(txtTemplateEngineer)
		Me.Controls.Add(txtOrderDate)
		Me.Controls.Add(txtMPDwo)
		Me.Controls.Add(txtAge)
		Me.Controls.Add(txtSex)
		Me.Controls.Add(txtDiagnosis)
		Me.Controls.Add(txtPatientName)
		Me.Controls.Add(txtUnits)
		Me.Controls.Add(txtFileNo)
		Me.Controls.Add(Label2)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Label1_1)
		Me.frmPatientDetails.Controls.Add(txtFileNoPD)
		Me.frmPatientDetails.Controls.Add(txtAgePD)
		Me.frmPatientDetails.Controls.Add(cboSexPD)
		Me.frmPatientDetails.Controls.Add(txtNamePD)
		Me.frmPatientDetails.Controls.Add(cboDiagnosisPD)
		Me.frmPatientDetails.Controls.Add(_labPatientDetails_0)
		Me.frmPatientDetails.Controls.Add(_labPatientDetails_1)
		Me.frmPatientDetails.Controls.Add(_labPatientDetails_2)
		Me.frmPatientDetails.Controls.Add(_labPatientDetails_3)
		Me.frmPatientDetails.Controls.Add(_labPatientDetails_4)
		Me.frmCADFile.Controls.Add(CAD_File)
		Me.frmCADFile.Controls.Add(lblCADFile)
		Me.frmWODetails.Controls.Add(cboDesignerWOD)
		Me.frmWODetails.Controls.Add(txtOrderDateWOD)
		Me.frmWODetails.Controls.Add(cboUnitsWOD)
		Me.frmWODetails.Controls.Add(txtWorkOrderWOD)
		Me.frmWODetails.Controls.Add(_labWODetails_3)
		Me.frmWODetails.Controls.Add(_labWODetails_2)
		Me.frmWODetails.Controls.Add(_labWODetails_1)
		Me.frmWODetails.Controls.Add(_labWODetails_0)
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_0, CType(0, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_1, CType(1, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_2, CType(2, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_3, CType(3, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_4, CType(4, Short))
		Me.labWODetails.SetIndex(_labWODetails_3, CType(3, Short))
		Me.labWODetails.SetIndex(_labWODetails_2, CType(2, Short))
		Me.labWODetails.SetIndex(_labWODetails_1, CType(1, Short))
		Me.labWODetails.SetIndex(_labWODetails_0, CType(0, Short))
		CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmPatientDetails.ResumeLayout(False)
		Me.frmCADFile.ResumeLayout(False)
		Me.frmWODetails.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class
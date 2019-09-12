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
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.frmPatientDetails = New System.Windows.Forms.GroupBox()
        Me.txtFileNoPD = New System.Windows.Forms.TextBox()
        Me.txtAgePD = New System.Windows.Forms.TextBox()
        Me.cboSexPD = New System.Windows.Forms.ComboBox()
        Me.txtNamePD = New System.Windows.Forms.TextBox()
        Me.cboDiagnosisPD = New System.Windows.Forms.ComboBox()
        Me._labPatientDetails_0 = New System.Windows.Forms.Label()
        Me._labPatientDetails_1 = New System.Windows.Forms.Label()
        Me._labPatientDetails_2 = New System.Windows.Forms.Label()
        Me._labPatientDetails_3 = New System.Windows.Forms.Label()
        Me._labPatientDetails_4 = New System.Windows.Forms.Label()
        Me.CMDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.txtInitialCADFile = New System.Windows.Forms.TextBox()
        Me.txtCurrentCADFile = New System.Windows.Forms.TextBox()
        Me.frmCADFile = New System.Windows.Forms.GroupBox()
        Me.CAD_File = New System.Windows.Forms.Button()
        Me.lblCADFile = New System.Windows.Forms.Label()
        Me.txtInvokedFrom = New System.Windows.Forms.TextBox()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.OK = New System.Windows.Forms.Button()
        Me.frmWODetails = New System.Windows.Forms.GroupBox()
        Me.cboDesignerWOD = New System.Windows.Forms.ComboBox()
        Me.txtOrderDateWOD = New System.Windows.Forms.TextBox()
        Me.cboUnitsWOD = New System.Windows.Forms.ComboBox()
        Me.txtWorkOrderWOD = New System.Windows.Forms.TextBox()
        Me._labWODetails_3 = New System.Windows.Forms.Label()
        Me._labWODetails_2 = New System.Windows.Forms.Label()
        Me._labWODetails_1 = New System.Windows.Forms.Label()
        Me._labWODetails_0 = New System.Windows.Forms.Label()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtWO_TXF = New System.Windows.Forms.TextBox()
        Me.txtWorkOrder = New System.Windows.Forms.TextBox()
        Me.txtUidMPD = New System.Windows.Forms.TextBox()
        Me.txtTemplateEngineer = New System.Windows.Forms.TextBox()
        Me.txtOrderDate = New System.Windows.Forms.TextBox()
        Me.txtMPDwo = New System.Windows.Forms.TextBox()
        Me.txtAge = New System.Windows.Forms.TextBox()
        Me.txtSex = New System.Windows.Forms.TextBox()
        Me.txtDiagnosis = New System.Windows.Forms.TextBox()
        Me.txtPatientName = New System.Windows.Forms.TextBox()
        Me.txtUnits = New System.Windows.Forms.TextBox()
        Me.txtFileNo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.labPatientDetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.labWODetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.frmPatientDetails.SuspendLayout()
        Me.frmCADFile.SuspendLayout()
        Me.frmWODetails.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'frmPatientDetails
        '
        Me.frmPatientDetails.BackColor = System.Drawing.SystemColors.Control
        Me.frmPatientDetails.Controls.Add(Me.txtFileNoPD)
        Me.frmPatientDetails.Controls.Add(Me.txtAgePD)
        Me.frmPatientDetails.Controls.Add(Me.cboSexPD)
        Me.frmPatientDetails.Controls.Add(Me.txtNamePD)
        Me.frmPatientDetails.Controls.Add(Me.cboDiagnosisPD)
        Me.frmPatientDetails.Controls.Add(Me._labPatientDetails_0)
        Me.frmPatientDetails.Controls.Add(Me._labPatientDetails_1)
        Me.frmPatientDetails.Controls.Add(Me._labPatientDetails_2)
        Me.frmPatientDetails.Controls.Add(Me._labPatientDetails_3)
        Me.frmPatientDetails.Controls.Add(Me._labPatientDetails_4)
        Me.frmPatientDetails.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmPatientDetails.Location = New System.Drawing.Point(8, 8)
        Me.frmPatientDetails.Name = "frmPatientDetails"
        Me.frmPatientDetails.Padding = New System.Windows.Forms.Padding(0)
        Me.frmPatientDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmPatientDetails.Size = New System.Drawing.Size(414, 105)
        Me.frmPatientDetails.TabIndex = 33
        Me.frmPatientDetails.TabStop = False
        Me.frmPatientDetails.Text = "Patient Details"
        '
        'txtFileNoPD
        '
        Me.txtFileNoPD.AcceptsReturn = True
        Me.txtFileNoPD.BackColor = System.Drawing.SystemColors.Window
        Me.txtFileNoPD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFileNoPD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFileNoPD.Location = New System.Drawing.Point(76, 16)
        Me.txtFileNoPD.MaxLength = 0
        Me.txtFileNoPD.Name = "txtFileNoPD"
        Me.txtFileNoPD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFileNoPD.Size = New System.Drawing.Size(81, 20)
        Me.txtFileNoPD.TabIndex = 38
        '
        'txtAgePD
        '
        Me.txtAgePD.AcceptsReturn = True
        Me.txtAgePD.BackColor = System.Drawing.SystemColors.Window
        Me.txtAgePD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAgePD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAgePD.Location = New System.Drawing.Point(332, 43)
        Me.txtAgePD.MaxLength = 0
        Me.txtAgePD.Name = "txtAgePD"
        Me.txtAgePD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAgePD.Size = New System.Drawing.Size(29, 20)
        Me.txtAgePD.TabIndex = 37
        Me.txtAgePD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cboSexPD
        '
        Me.cboSexPD.BackColor = System.Drawing.SystemColors.Window
        Me.cboSexPD.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSexPD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSexPD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSexPD.Location = New System.Drawing.Point(332, 70)
        Me.cboSexPD.Name = "cboSexPD"
        Me.cboSexPD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSexPD.Size = New System.Drawing.Size(65, 21)
        Me.cboSexPD.TabIndex = 36
        '
        'txtNamePD
        '
        Me.txtNamePD.AcceptsReturn = True
        Me.txtNamePD.BackColor = System.Drawing.SystemColors.Window
        Me.txtNamePD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNamePD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNamePD.Location = New System.Drawing.Point(76, 43)
        Me.txtNamePD.MaxLength = 0
        Me.txtNamePD.Name = "txtNamePD"
        Me.txtNamePD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNamePD.Size = New System.Drawing.Size(213, 20)
        Me.txtNamePD.TabIndex = 35
        '
        'cboDiagnosisPD
        '
        Me.cboDiagnosisPD.BackColor = System.Drawing.SystemColors.Window
        Me.cboDiagnosisPD.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDiagnosisPD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDiagnosisPD.Location = New System.Drawing.Point(76, 70)
        Me.cboDiagnosisPD.Name = "cboDiagnosisPD"
        Me.cboDiagnosisPD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDiagnosisPD.Size = New System.Drawing.Size(213, 21)
        Me.cboDiagnosisPD.TabIndex = 34
        '
        '_labPatientDetails_0
        '
        Me._labPatientDetails_0.BackColor = System.Drawing.SystemColors.Control
        Me._labPatientDetails_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._labPatientDetails_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labPatientDetails.SetIndex(Me._labPatientDetails_0, CType(0, Short))
        Me._labPatientDetails_0.Location = New System.Drawing.Point(24, 19)
        Me._labPatientDetails_0.Name = "_labPatientDetails_0"
        Me._labPatientDetails_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labPatientDetails_0.Size = New System.Drawing.Size(45, 13)
        Me._labPatientDetails_0.TabIndex = 43
        Me._labPatientDetails_0.Text = "File No:"
        '
        '_labPatientDetails_1
        '
        Me._labPatientDetails_1.BackColor = System.Drawing.SystemColors.Control
        Me._labPatientDetails_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._labPatientDetails_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labPatientDetails.SetIndex(Me._labPatientDetails_1, CType(1, Short))
        Me._labPatientDetails_1.Location = New System.Drawing.Point(32, 46)
        Me._labPatientDetails_1.Name = "_labPatientDetails_1"
        Me._labPatientDetails_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labPatientDetails_1.Size = New System.Drawing.Size(41, 13)
        Me._labPatientDetails_1.TabIndex = 42
        Me._labPatientDetails_1.Text = "Name:"
        '
        '_labPatientDetails_2
        '
        Me._labPatientDetails_2.BackColor = System.Drawing.SystemColors.Control
        Me._labPatientDetails_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._labPatientDetails_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labPatientDetails.SetIndex(Me._labPatientDetails_2, CType(2, Short))
        Me._labPatientDetails_2.Location = New System.Drawing.Point(8, 73)
        Me._labPatientDetails_2.Name = "_labPatientDetails_2"
        Me._labPatientDetails_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labPatientDetails_2.Size = New System.Drawing.Size(61, 13)
        Me._labPatientDetails_2.TabIndex = 41
        Me._labPatientDetails_2.Text = "Diagnosis:"
        '
        '_labPatientDetails_3
        '
        Me._labPatientDetails_3.BackColor = System.Drawing.SystemColors.Control
        Me._labPatientDetails_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._labPatientDetails_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labPatientDetails.SetIndex(Me._labPatientDetails_3, CType(3, Short))
        Me._labPatientDetails_3.Location = New System.Drawing.Point(300, 46)
        Me._labPatientDetails_3.Name = "_labPatientDetails_3"
        Me._labPatientDetails_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labPatientDetails_3.Size = New System.Drawing.Size(33, 13)
        Me._labPatientDetails_3.TabIndex = 40
        Me._labPatientDetails_3.Text = "Age:"
        '
        '_labPatientDetails_4
        '
        Me._labPatientDetails_4.BackColor = System.Drawing.SystemColors.Control
        Me._labPatientDetails_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._labPatientDetails_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labPatientDetails.SetIndex(Me._labPatientDetails_4, CType(4, Short))
        Me._labPatientDetails_4.Location = New System.Drawing.Point(300, 73)
        Me._labPatientDetails_4.Name = "_labPatientDetails_4"
        Me._labPatientDetails_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labPatientDetails_4.Size = New System.Drawing.Size(33, 13)
        Me._labPatientDetails_4.TabIndex = 39
        Me._labPatientDetails_4.Text = "Sex:"
        '
        'CMDialog1Open
        '
        Me.CMDialog1Open.Title = "Open CAD File"
        '
        'txtInitialCADFile
        '
        Me.txtInitialCADFile.AcceptsReturn = True
        Me.txtInitialCADFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtInitialCADFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInitialCADFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInitialCADFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitialCADFile.Location = New System.Drawing.Point(444, 64)
        Me.txtInitialCADFile.MaxLength = 0
        Me.txtInitialCADFile.Name = "txtInitialCADFile"
        Me.txtInitialCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInitialCADFile.Size = New System.Drawing.Size(97, 20)
        Me.txtInitialCADFile.TabIndex = 32
        Me.txtInitialCADFile.Text = "txtInitialCADFile"
        Me.txtInitialCADFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtInitialCADFile.Visible = False
        '
        'txtCurrentCADFile
        '
        Me.txtCurrentCADFile.AcceptsReturn = True
        Me.txtCurrentCADFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrentCADFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrentCADFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrentCADFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrentCADFile.Location = New System.Drawing.Point(444, 40)
        Me.txtCurrentCADFile.MaxLength = 0
        Me.txtCurrentCADFile.Name = "txtCurrentCADFile"
        Me.txtCurrentCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrentCADFile.Size = New System.Drawing.Size(97, 20)
        Me.txtCurrentCADFile.TabIndex = 31
        Me.txtCurrentCADFile.Text = "txtCurrentCADFile"
        Me.txtCurrentCADFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtCurrentCADFile.Visible = False
        '
        'frmCADFile
        '
        Me.frmCADFile.BackColor = System.Drawing.SystemColors.Control
        Me.frmCADFile.Controls.Add(Me.CAD_File)
        Me.frmCADFile.Controls.Add(Me.lblCADFile)
        Me.frmCADFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmCADFile.Location = New System.Drawing.Point(8, 214)
        Me.frmCADFile.Name = "frmCADFile"
        Me.frmCADFile.Padding = New System.Windows.Forms.Padding(0)
        Me.frmCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmCADFile.Size = New System.Drawing.Size(413, 55)
        Me.frmCADFile.TabIndex = 28
        Me.frmCADFile.TabStop = False
        Me.frmCADFile.Text = "Drafix CAD File"
        '
        'CAD_File
        '
        Me.CAD_File.BackColor = System.Drawing.SystemColors.Control
        Me.CAD_File.Cursor = System.Windows.Forms.Cursors.Default
        Me.CAD_File.Enabled = False
        Me.CAD_File.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CAD_File.Location = New System.Drawing.Point(316, 22)
        Me.CAD_File.Name = "CAD_File"
        Me.CAD_File.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CAD_File.Size = New System.Drawing.Size(85, 25)
        Me.CAD_File.TabIndex = 30
        Me.CAD_File.Text = "CAD File ..."
        Me.CAD_File.UseVisualStyleBackColor = False
        '
        'lblCADFile
        '
        Me.lblCADFile.BackColor = System.Drawing.SystemColors.Control
        Me.lblCADFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCADFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCADFile.Location = New System.Drawing.Point(12, 24)
        Me.lblCADFile.Name = "lblCADFile"
        Me.lblCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCADFile.Size = New System.Drawing.Size(297, 13)
        Me.lblCADFile.TabIndex = 29
        '
        'txtInvokedFrom
        '
        Me.txtInvokedFrom.AcceptsReturn = True
        Me.txtInvokedFrom.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvokedFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInvokedFrom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvokedFrom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvokedFrom.Location = New System.Drawing.Point(444, 16)
        Me.txtInvokedFrom.MaxLength = 0
        Me.txtInvokedFrom.Name = "txtInvokedFrom"
        Me.txtInvokedFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvokedFrom.Size = New System.Drawing.Size(97, 20)
        Me.txtInvokedFrom.TabIndex = 0
        Me.txtInvokedFrom.Text = "txtInvokedFrom"
        Me.txtInvokedFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtInvokedFrom.Visible = False
        '
        'Cancel
        '
        Me.Cancel.BackColor = System.Drawing.SystemColors.Control
        Me.Cancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Cancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Cancel.Location = New System.Drawing.Point(216, 416)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Cancel.Size = New System.Drawing.Size(85, 25)
        Me.Cancel.TabIndex = 2
        Me.Cancel.Text = "Cancel"
        Me.Cancel.UseVisualStyleBackColor = False
        '
        'OK
        '
        Me.OK.BackColor = System.Drawing.SystemColors.Control
        Me.OK.Cursor = System.Windows.Forms.Cursors.Default
        Me.OK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OK.Location = New System.Drawing.Point(112, 416)
        Me.OK.Name = "OK"
        Me.OK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OK.Size = New System.Drawing.Size(85, 25)
        Me.OK.TabIndex = 19
        Me.OK.Text = "OK"
        Me.OK.UseVisualStyleBackColor = False
        '
        'frmWODetails
        '
        Me.frmWODetails.BackColor = System.Drawing.SystemColors.Control
        Me.frmWODetails.Controls.Add(Me.cboDesignerWOD)
        Me.frmWODetails.Controls.Add(Me.txtOrderDateWOD)
        Me.frmWODetails.Controls.Add(Me.cboUnitsWOD)
        Me.frmWODetails.Controls.Add(Me.txtWorkOrderWOD)
        Me.frmWODetails.Controls.Add(Me._labWODetails_3)
        Me.frmWODetails.Controls.Add(Me._labWODetails_2)
        Me.frmWODetails.Controls.Add(Me._labWODetails_1)
        Me.frmWODetails.Controls.Add(Me._labWODetails_0)
        Me.frmWODetails.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmWODetails.Location = New System.Drawing.Point(8, 120)
        Me.frmWODetails.Name = "frmWODetails"
        Me.frmWODetails.Padding = New System.Windows.Forms.Padding(0)
        Me.frmWODetails.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmWODetails.Size = New System.Drawing.Size(414, 88)
        Me.frmWODetails.TabIndex = 10
        Me.frmWODetails.TabStop = False
        Me.frmWODetails.Text = "Work Order Details"
        '
        'cboDesignerWOD
        '
        Me.cboDesignerWOD.BackColor = System.Drawing.SystemColors.Window
        Me.cboDesignerWOD.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDesignerWOD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDesignerWOD.Location = New System.Drawing.Point(76, 48)
        Me.cboDesignerWOD.Name = "cboDesignerWOD"
        Me.cboDesignerWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDesignerWOD.Size = New System.Drawing.Size(133, 21)
        Me.cboDesignerWOD.TabIndex = 14
        '
        'txtOrderDateWOD
        '
        Me.txtOrderDateWOD.AcceptsReturn = True
        Me.txtOrderDateWOD.BackColor = System.Drawing.SystemColors.Window
        Me.txtOrderDateWOD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOrderDateWOD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOrderDateWOD.Location = New System.Drawing.Point(207, 21)
        Me.txtOrderDateWOD.MaxLength = 10
        Me.txtOrderDateWOD.Name = "txtOrderDateWOD"
        Me.txtOrderDateWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOrderDateWOD.Size = New System.Drawing.Size(80, 20)
        Me.txtOrderDateWOD.TabIndex = 12
        Me.txtOrderDateWOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cboUnitsWOD
        '
        Me.cboUnitsWOD.BackColor = System.Drawing.SystemColors.Window
        Me.cboUnitsWOD.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboUnitsWOD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUnitsWOD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboUnitsWOD.Location = New System.Drawing.Point(332, 21)
        Me.cboUnitsWOD.Name = "cboUnitsWOD"
        Me.cboUnitsWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboUnitsWOD.Size = New System.Drawing.Size(65, 21)
        Me.cboUnitsWOD.TabIndex = 13
        '
        'txtWorkOrderWOD
        '
        Me.txtWorkOrderWOD.AcceptsReturn = True
        Me.txtWorkOrderWOD.BackColor = System.Drawing.SystemColors.Window
        Me.txtWorkOrderWOD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWorkOrderWOD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWorkOrderWOD.Location = New System.Drawing.Point(76, 21)
        Me.txtWorkOrderWOD.MaxLength = 0
        Me.txtWorkOrderWOD.Name = "txtWorkOrderWOD"
        Me.txtWorkOrderWOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWorkOrderWOD.Size = New System.Drawing.Size(57, 20)
        Me.txtWorkOrderWOD.TabIndex = 11
        Me.txtWorkOrderWOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_labWODetails_3
        '
        Me._labWODetails_3.BackColor = System.Drawing.SystemColors.Control
        Me._labWODetails_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._labWODetails_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labWODetails.SetIndex(Me._labWODetails_3, CType(3, Short))
        Me._labWODetails_3.Location = New System.Drawing.Point(292, 24)
        Me._labWODetails_3.Name = "_labWODetails_3"
        Me._labWODetails_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labWODetails_3.Size = New System.Drawing.Size(41, 16)
        Me._labWODetails_3.TabIndex = 20
        Me._labWODetails_3.Text = "Units:"
        '
        '_labWODetails_2
        '
        Me._labWODetails_2.BackColor = System.Drawing.SystemColors.Control
        Me._labWODetails_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._labWODetails_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labWODetails.SetIndex(Me._labWODetails_2, CType(2, Short))
        Me._labWODetails_2.Location = New System.Drawing.Point(140, 24)
        Me._labWODetails_2.Name = "_labWODetails_2"
        Me._labWODetails_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labWODetails_2.Size = New System.Drawing.Size(73, 16)
        Me._labWODetails_2.TabIndex = 27
        Me._labWODetails_2.Text = "Temp Date:"
        '
        '_labWODetails_1
        '
        Me._labWODetails_1.BackColor = System.Drawing.SystemColors.Control
        Me._labWODetails_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._labWODetails_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labWODetails.SetIndex(Me._labWODetails_1, CType(1, Short))
        Me._labWODetails_1.Location = New System.Drawing.Point(16, 51)
        Me._labWODetails_1.Name = "_labWODetails_1"
        Me._labWODetails_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labWODetails_1.Size = New System.Drawing.Size(57, 16)
        Me._labWODetails_1.TabIndex = 26
        Me._labWODetails_1.Text = "Designer:"
        '
        '_labWODetails_0
        '
        Me._labWODetails_0.BackColor = System.Drawing.SystemColors.Control
        Me._labWODetails_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._labWODetails_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labWODetails.SetIndex(Me._labWODetails_0, CType(0, Short))
        Me._labWODetails_0.Location = New System.Drawing.Point(4, 24)
        Me._labWODetails_0.Name = "_labWODetails_0"
        Me._labWODetails_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labWODetails_0.Size = New System.Drawing.Size(73, 16)
        Me._labWODetails_0.TabIndex = 25
        Me._labWODetails_0.Text = "WorkOrder:"
        '
        'txtMessage
        '
        Me.txtMessage.AcceptsReturn = True
        Me.txtMessage.BackColor = System.Drawing.SystemColors.Window
        Me.txtMessage.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMessage.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMessage.Location = New System.Drawing.Point(8, 292)
        Me.txtMessage.MaxLength = 0
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMessage.Size = New System.Drawing.Size(413, 112)
        Me.txtMessage.TabIndex = 21
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'txtWO_TXF
        '
        Me.txtWO_TXF.AcceptsReturn = True
        Me.txtWO_TXF.BackColor = System.Drawing.SystemColors.Window
        Me.txtWO_TXF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtWO_TXF.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWO_TXF.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWO_TXF.Location = New System.Drawing.Point(444, 137)
        Me.txtWO_TXF.MaxLength = 0
        Me.txtWO_TXF.Name = "txtWO_TXF"
        Me.txtWO_TXF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWO_TXF.Size = New System.Drawing.Size(97, 20)
        Me.txtWO_TXF.TabIndex = 23
        Me.txtWO_TXF.Text = "txtWO_TXF"
        Me.txtWO_TXF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtWO_TXF.Visible = False
        '
        'txtWorkOrder
        '
        Me.txtWorkOrder.AcceptsReturn = True
        Me.txtWorkOrder.BackColor = System.Drawing.SystemColors.Window
        Me.txtWorkOrder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtWorkOrder.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWorkOrder.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWorkOrder.Location = New System.Drawing.Point(444, 113)
        Me.txtWorkOrder.MaxLength = 0
        Me.txtWorkOrder.Name = "txtWorkOrder"
        Me.txtWorkOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWorkOrder.Size = New System.Drawing.Size(97, 20)
        Me.txtWorkOrder.TabIndex = 6
        Me.txtWorkOrder.Text = "txtWorkOrder"
        Me.txtWorkOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtWorkOrder.Visible = False
        '
        'txtUidMPD
        '
        Me.txtUidMPD.AcceptsReturn = True
        Me.txtUidMPD.BackColor = System.Drawing.SystemColors.Window
        Me.txtUidMPD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUidMPD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUidMPD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUidMPD.Location = New System.Drawing.Point(444, 401)
        Me.txtUidMPD.MaxLength = 0
        Me.txtUidMPD.Name = "txtUidMPD"
        Me.txtUidMPD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUidMPD.Size = New System.Drawing.Size(97, 20)
        Me.txtUidMPD.TabIndex = 17
        Me.txtUidMPD.Text = "txtUidMPD"
        Me.txtUidMPD.Visible = False
        '
        'txtTemplateEngineer
        '
        Me.txtTemplateEngineer.AcceptsReturn = True
        Me.txtTemplateEngineer.BackColor = System.Drawing.SystemColors.Window
        Me.txtTemplateEngineer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTemplateEngineer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTemplateEngineer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTemplateEngineer.Location = New System.Drawing.Point(444, 377)
        Me.txtTemplateEngineer.MaxLength = 0
        Me.txtTemplateEngineer.Name = "txtTemplateEngineer"
        Me.txtTemplateEngineer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTemplateEngineer.Size = New System.Drawing.Size(97, 20)
        Me.txtTemplateEngineer.TabIndex = 16
        Me.txtTemplateEngineer.Text = "txtTemplateEngineer"
        Me.txtTemplateEngineer.Visible = False
        '
        'txtOrderDate
        '
        Me.txtOrderDate.AcceptsReturn = True
        Me.txtOrderDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtOrderDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOrderDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOrderDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOrderDate.Location = New System.Drawing.Point(444, 353)
        Me.txtOrderDate.MaxLength = 0
        Me.txtOrderDate.Name = "txtOrderDate"
        Me.txtOrderDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOrderDate.Size = New System.Drawing.Size(97, 20)
        Me.txtOrderDate.TabIndex = 15
        Me.txtOrderDate.Text = "txtOrderDate"
        Me.txtOrderDate.Visible = False
        '
        'txtMPDwo
        '
        Me.txtMPDwo.AcceptsReturn = True
        Me.txtMPDwo.BackColor = System.Drawing.SystemColors.Window
        Me.txtMPDwo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMPDwo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMPDwo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMPDwo.Location = New System.Drawing.Point(444, 329)
        Me.txtMPDwo.MaxLength = 0
        Me.txtMPDwo.Name = "txtMPDwo"
        Me.txtMPDwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMPDwo.Size = New System.Drawing.Size(97, 20)
        Me.txtMPDwo.TabIndex = 9
        Me.txtMPDwo.Text = "txtMPDwo"
        Me.txtMPDwo.Visible = False
        '
        'txtAge
        '
        Me.txtAge.AcceptsReturn = True
        Me.txtAge.BackColor = System.Drawing.SystemColors.Window
        Me.txtAge.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAge.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAge.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAge.Location = New System.Drawing.Point(444, 233)
        Me.txtAge.MaxLength = 0
        Me.txtAge.Name = "txtAge"
        Me.txtAge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAge.Size = New System.Drawing.Size(97, 20)
        Me.txtAge.TabIndex = 8
        Me.txtAge.Text = "txtAge"
        Me.txtAge.Visible = False
        '
        'txtSex
        '
        Me.txtSex.AcceptsReturn = True
        Me.txtSex.BackColor = System.Drawing.SystemColors.Window
        Me.txtSex.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSex.Location = New System.Drawing.Point(444, 281)
        Me.txtSex.MaxLength = 0
        Me.txtSex.Name = "txtSex"
        Me.txtSex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSex.Size = New System.Drawing.Size(97, 20)
        Me.txtSex.TabIndex = 7
        Me.txtSex.Text = "txtSex"
        Me.txtSex.Visible = False
        '
        'txtDiagnosis
        '
        Me.txtDiagnosis.AcceptsReturn = True
        Me.txtDiagnosis.BackColor = System.Drawing.SystemColors.Window
        Me.txtDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDiagnosis.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDiagnosis.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDiagnosis.Location = New System.Drawing.Point(444, 305)
        Me.txtDiagnosis.MaxLength = 0
        Me.txtDiagnosis.Name = "txtDiagnosis"
        Me.txtDiagnosis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDiagnosis.Size = New System.Drawing.Size(97, 20)
        Me.txtDiagnosis.TabIndex = 5
        Me.txtDiagnosis.Text = "txtDiagnosis"
        Me.txtDiagnosis.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDiagnosis.Visible = False
        '
        'txtPatientName
        '
        Me.txtPatientName.AcceptsReturn = True
        Me.txtPatientName.BackColor = System.Drawing.SystemColors.Window
        Me.txtPatientName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPatientName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPatientName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPatientName.Location = New System.Drawing.Point(444, 209)
        Me.txtPatientName.MaxLength = 0
        Me.txtPatientName.Name = "txtPatientName"
        Me.txtPatientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPatientName.Size = New System.Drawing.Size(97, 20)
        Me.txtPatientName.TabIndex = 4
        Me.txtPatientName.Text = "txtPatientName"
        Me.txtPatientName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPatientName.Visible = False
        '
        'txtUnits
        '
        Me.txtUnits.AcceptsReturn = True
        Me.txtUnits.BackColor = System.Drawing.SystemColors.Window
        Me.txtUnits.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUnits.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUnits.Location = New System.Drawing.Point(444, 257)
        Me.txtUnits.MaxLength = 0
        Me.txtUnits.Name = "txtUnits"
        Me.txtUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUnits.Size = New System.Drawing.Size(97, 20)
        Me.txtUnits.TabIndex = 3
        Me.txtUnits.Text = "txtUnits"
        Me.txtUnits.Visible = False
        '
        'txtFileNo
        '
        Me.txtFileNo.AcceptsReturn = True
        Me.txtFileNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtFileNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFileNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFileNo.Location = New System.Drawing.Point(444, 185)
        Me.txtFileNo.MaxLength = 0
        Me.txtFileNo.Name = "txtFileNo"
        Me.txtFileNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFileNo.Size = New System.Drawing.Size(97, 20)
        Me.txtFileNo.TabIndex = 1
        Me.txtFileNo.Text = "txtFileNo"
        Me.txtFileNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFileNo.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 272)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(121, 16)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Messages and Errors"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Window
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(456, 95)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(85, 13)
        Me._Label1_0.TabIndex = 24
        Me._Label1_0.Text = "imageABLE"
        Me._Label1_0.Visible = False
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Window
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(456, 167)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(81, 13)
        Me._Label1_1.TabIndex = 18
        Me._Label1_1.Text = "PatientDetails"
        Me._Label1_1.Visible = False
        '
        'cltxfdia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(429, 459)
        Me.Controls.Add(Me.frmPatientDetails)
        Me.Controls.Add(Me.txtInitialCADFile)
        Me.Controls.Add(Me.txtCurrentCADFile)
        Me.Controls.Add(Me.frmCADFile)
        Me.Controls.Add(Me.txtInvokedFrom)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.frmWODetails)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.txtWO_TXF)
        Me.Controls.Add(Me.txtWorkOrder)
        Me.Controls.Add(Me.txtUidMPD)
        Me.Controls.Add(Me.txtTemplateEngineer)
        Me.Controls.Add(Me.txtOrderDate)
        Me.Controls.Add(Me.txtMPDwo)
        Me.Controls.Add(Me.txtAge)
        Me.Controls.Add(Me.txtSex)
        Me.Controls.Add(Me.txtDiagnosis)
        Me.Controls.Add(Me.txtPatientName)
        Me.Controls.Add(Me.txtUnits)
        Me.Controls.Add(Me.txtFileNo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(47, 133)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "cltxfdia"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Patient and Work Order Details"
        Me.frmPatientDetails.ResumeLayout(False)
        Me.frmPatientDetails.PerformLayout()
        Me.frmCADFile.ResumeLayout(False)
        Me.frmWODetails.ResumeLayout(False)
        Me.frmWODetails.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class
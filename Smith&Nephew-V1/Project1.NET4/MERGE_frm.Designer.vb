<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class merge
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
	Public WithEvents CAD_File As System.Windows.Forms.Button
	Public WithEvents txtCurrentCADFile As System.Windows.Forms.TextBox
	Public CMDialog1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents frmCADFile As System.Windows.Forms.GroupBox
	Public WithEvents Cancel As System.Windows.Forms.Button
	Public WithEvents OK As System.Windows.Forms.Button
	Public WithEvents txtOrderDate As System.Windows.Forms.TextBox
	Public WithEvents txtWorkOrder As System.Windows.Forms.TextBox
	Public WithEvents txtPatientName As System.Windows.Forms.TextBox
	Public WithEvents txtFileNo As System.Windows.Forms.TextBox
	Public WithEvents _labPatientDetails_0 As System.Windows.Forms.Label
	Public WithEvents _labPatientDetails_1 As System.Windows.Forms.Label
	Public WithEvents _labWODetails_2 As System.Windows.Forms.Label
	Public WithEvents _labWODetails_0 As System.Windows.Forms.Label
	Public WithEvents frmWODetails As System.Windows.Forms.GroupBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents labPatientDetails As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents labWODetails As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(merge))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.CAD_File = New System.Windows.Forms.Button
		Me.txtCurrentCADFile = New System.Windows.Forms.TextBox
		Me.CMDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.frmCADFile = New System.Windows.Forms.GroupBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.Cancel = New System.Windows.Forms.Button
		Me.OK = New System.Windows.Forms.Button
		Me.frmWODetails = New System.Windows.Forms.GroupBox
		Me.txtOrderDate = New System.Windows.Forms.TextBox
		Me.txtWorkOrder = New System.Windows.Forms.TextBox
		Me.txtPatientName = New System.Windows.Forms.TextBox
		Me.txtFileNo = New System.Windows.Forms.TextBox
		Me._labPatientDetails_0 = New System.Windows.Forms.Label
		Me._labPatientDetails_1 = New System.Windows.Forms.Label
		Me._labWODetails_2 = New System.Windows.Forms.Label
		Me._labWODetails_0 = New System.Windows.Forms.Label
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.labPatientDetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.labWODetails = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.frmCADFile.SuspendLayout()
		Me.frmWODetails.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Merge Drawings"
		Me.ClientSize = New System.Drawing.Size(354, 248)
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
		Me.Name = "merge"
		Me.CAD_File.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.CAD_File
		Me.CAD_File.Text = "Find ..."
		Me.CAD_File.Size = New System.Drawing.Size(85, 25)
		Me.CAD_File.Location = New System.Drawing.Point(16, 212)
		Me.CAD_File.TabIndex = 13
		Me.CAD_File.BackColor = System.Drawing.SystemColors.Control
		Me.CAD_File.CausesValidation = True
		Me.CAD_File.Enabled = True
		Me.CAD_File.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CAD_File.Cursor = System.Windows.Forms.Cursors.Default
		Me.CAD_File.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CAD_File.TabStop = True
		Me.CAD_File.Name = "CAD_File"
		Me.txtCurrentCADFile.AutoSize = False
		Me.txtCurrentCADFile.Size = New System.Drawing.Size(113, 19)
		Me.txtCurrentCADFile.Location = New System.Drawing.Point(420, 168)
		Me.txtCurrentCADFile.TabIndex = 12
		Me.txtCurrentCADFile.Text = "txtCurrentCADFile"
		Me.txtCurrentCADFile.AcceptsReturn = True
		Me.txtCurrentCADFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
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
		Me.txtCurrentCADFile.Visible = True
		Me.txtCurrentCADFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtCurrentCADFile.Name = "txtCurrentCADFile"
		Me.CMDialog1Open.Title = "Open CAD File"
		Me.frmCADFile.Text = "Drawing to Merge"
		Me.frmCADFile.Size = New System.Drawing.Size(329, 55)
		Me.frmCADFile.Location = New System.Drawing.Point(12, 144)
		Me.frmCADFile.TabIndex = 2
		Me.frmCADFile.BackColor = System.Drawing.SystemColors.Control
		Me.frmCADFile.Enabled = True
		Me.frmCADFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmCADFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmCADFile.Visible = True
		Me.frmCADFile.Padding = New System.Windows.Forms.Padding(0)
		Me.frmCADFile.Name = "frmCADFile"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Use Find to select drawing to merge!"
		Me.Label1.Size = New System.Drawing.Size(293, 13)
		Me.Label1.Location = New System.Drawing.Point(16, 24)
		Me.Label1.TabIndex = 14
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Cancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Cancel.Text = "Cancel"
		Me.Cancel.Size = New System.Drawing.Size(85, 25)
		Me.Cancel.Location = New System.Drawing.Point(252, 212)
		Me.Cancel.TabIndex = 0
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
		Me.OK.Location = New System.Drawing.Point(160, 212)
		Me.OK.TabIndex = 3
		Me.OK.BackColor = System.Drawing.SystemColors.Control
		Me.OK.CausesValidation = True
		Me.OK.Enabled = True
		Me.OK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OK.Cursor = System.Windows.Forms.Cursors.Default
		Me.OK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OK.TabStop = True
		Me.OK.Name = "OK"
		Me.frmWODetails.Text = "Current Drawing"
		Me.frmWODetails.Size = New System.Drawing.Size(329, 121)
		Me.frmWODetails.Location = New System.Drawing.Point(12, 12)
		Me.frmWODetails.TabIndex = 1
		Me.frmWODetails.BackColor = System.Drawing.SystemColors.Control
		Me.frmWODetails.Enabled = True
		Me.frmWODetails.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmWODetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmWODetails.Visible = True
		Me.frmWODetails.Padding = New System.Windows.Forms.Padding(0)
		Me.frmWODetails.Name = "frmWODetails"
		Me.txtOrderDate.AutoSize = False
		Me.txtOrderDate.Size = New System.Drawing.Size(89, 21)
		Me.txtOrderDate.Location = New System.Drawing.Point(224, 84)
		Me.txtOrderDate.TabIndex = 4
		Me.txtOrderDate.Text = "txtOrderDate"
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
		Me.txtOrderDate.Visible = True
		Me.txtOrderDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrderDate.Name = "txtOrderDate"
		Me.txtWorkOrder.AutoSize = False
		Me.txtWorkOrder.Size = New System.Drawing.Size(69, 21)
		Me.txtWorkOrder.Location = New System.Drawing.Point(72, 84)
		Me.txtWorkOrder.TabIndex = 5
		Me.txtWorkOrder.Text = "txtWorkOrder"
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
		Me.txtWorkOrder.Visible = True
		Me.txtWorkOrder.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWorkOrder.Name = "txtWorkOrder"
		Me.txtPatientName.AutoSize = False
		Me.txtPatientName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtPatientName.Size = New System.Drawing.Size(241, 21)
		Me.txtPatientName.Location = New System.Drawing.Point(72, 52)
		Me.txtPatientName.TabIndex = 6
		Me.txtPatientName.Text = "txtPatientName"
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
		Me.txtPatientName.Visible = True
		Me.txtPatientName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPatientName.Name = "txtPatientName"
		Me.txtFileNo.AutoSize = False
		Me.txtFileNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtFileNo.Size = New System.Drawing.Size(97, 21)
		Me.txtFileNo.Location = New System.Drawing.Point(72, 24)
		Me.txtFileNo.TabIndex = 7
		Me.txtFileNo.Text = "txtFileNo"
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
		Me.txtFileNo.Visible = True
		Me.txtFileNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFileNo.Name = "txtFileNo"
		Me._labPatientDetails_0.Text = "File No:"
		Me._labPatientDetails_0.Size = New System.Drawing.Size(45, 13)
		Me._labPatientDetails_0.Location = New System.Drawing.Point(20, 27)
		Me._labPatientDetails_0.TabIndex = 8
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
		Me._labPatientDetails_1.Location = New System.Drawing.Point(28, 54)
		Me._labPatientDetails_1.TabIndex = 9
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
		Me._labWODetails_2.Text = "W/O Date:"
		Me._labWODetails_2.Size = New System.Drawing.Size(73, 16)
		Me._labWODetails_2.Location = New System.Drawing.Point(152, 88)
		Me._labWODetails_2.TabIndex = 10
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
		Me._labWODetails_0.Text = "W/Order:"
		Me._labWODetails_0.Size = New System.Drawing.Size(73, 16)
		Me._labWODetails_0.Location = New System.Drawing.Point(12, 88)
		Me._labWODetails_0.TabIndex = 11
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
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.Controls.Add(CAD_File)
		Me.Controls.Add(txtCurrentCADFile)
		Me.Controls.Add(frmCADFile)
		Me.Controls.Add(Cancel)
		Me.Controls.Add(OK)
		Me.Controls.Add(frmWODetails)
		Me.frmCADFile.Controls.Add(Label1)
		Me.frmWODetails.Controls.Add(txtOrderDate)
		Me.frmWODetails.Controls.Add(txtWorkOrder)
		Me.frmWODetails.Controls.Add(txtPatientName)
		Me.frmWODetails.Controls.Add(txtFileNo)
		Me.frmWODetails.Controls.Add(_labPatientDetails_0)
		Me.frmWODetails.Controls.Add(_labPatientDetails_1)
		Me.frmWODetails.Controls.Add(_labWODetails_2)
		Me.frmWODetails.Controls.Add(_labWODetails_0)
		Me.labPatientDetails.SetIndex(_labPatientDetails_0, CType(0, Short))
		Me.labPatientDetails.SetIndex(_labPatientDetails_1, CType(1, Short))
		Me.labWODetails.SetIndex(_labWODetails_2, CType(2, Short))
		Me.labWODetails.SetIndex(_labWODetails_0, CType(0, Short))
		CType(Me.labWODetails, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.labPatientDetails, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmCADFile.ResumeLayout(False)
		Me.frmWODetails.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class EditUSER
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
	Public WithEvents Change As System.Windows.Forms.Button
	Public WithEvents Del As System.Windows.Forms.Button
	Public WithEvents Add_Renamed As System.Windows.Forms.Button
	Public WithEvents StampList As System.Windows.Forms.ListBox
	Public WithEvents OK As System.Windows.Forms.Button
	Public WithEvents txtStamp As System.Windows.Forms.TextBox
	Public WithEvents Cancel As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(EditUSER))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Change = New System.Windows.Forms.Button
		Me.Del = New System.Windows.Forms.Button
		Me.Add_Renamed = New System.Windows.Forms.Button
		Me.StampList = New System.Windows.Forms.ListBox
		Me.OK = New System.Windows.Forms.Button
		Me.txtStamp = New System.Windows.Forms.TextBox
		Me.Cancel = New System.Windows.Forms.Button
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Edit USER Stamps"
		Me.ClientSize = New System.Drawing.Size(344, 349)
		Me.Location = New System.Drawing.Point(73, 99)
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
		Me.Name = "EditUSER"
		Me.Change.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Change.Text = "Change"
		Me.Change.Size = New System.Drawing.Size(65, 21)
		Me.Change.Location = New System.Drawing.Point(224, 252)
		Me.Change.TabIndex = 6
		Me.Change.BackColor = System.Drawing.SystemColors.Control
		Me.Change.CausesValidation = True
		Me.Change.Enabled = True
		Me.Change.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Change.Cursor = System.Windows.Forms.Cursors.Default
		Me.Change.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Change.TabStop = True
		Me.Change.Name = "Change"
		Me.Del.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Del.Text = "Delete"
		Me.Del.Size = New System.Drawing.Size(65, 21)
		Me.Del.Location = New System.Drawing.Point(132, 252)
		Me.Del.TabIndex = 5
		Me.Del.BackColor = System.Drawing.SystemColors.Control
		Me.Del.CausesValidation = True
		Me.Del.Enabled = True
		Me.Del.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Del.Cursor = System.Windows.Forms.Cursors.Default
		Me.Del.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Del.TabStop = True
		Me.Del.Name = "Del"
		Me.Add_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Add_Renamed.Text = "Add"
		Me.Add_Renamed.Size = New System.Drawing.Size(65, 21)
		Me.Add_Renamed.Location = New System.Drawing.Point(40, 252)
		Me.Add_Renamed.TabIndex = 4
		Me.Add_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.Add_Renamed.CausesValidation = True
		Me.Add_Renamed.Enabled = True
		Me.Add_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Add_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.Add_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Add_Renamed.TabStop = True
		Me.Add_Renamed.Name = "Add_Renamed"
		Me.StampList.Size = New System.Drawing.Size(317, 176)
		Me.StampList.Location = New System.Drawing.Point(12, 48)
		Me.StampList.TabIndex = 3
		Me.StampList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.StampList.BackColor = System.Drawing.SystemColors.Window
		Me.StampList.CausesValidation = True
		Me.StampList.Enabled = True
		Me.StampList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.StampList.IntegralHeight = True
		Me.StampList.Cursor = System.Windows.Forms.Cursors.Default
		Me.StampList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.StampList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StampList.Sorted = False
		Me.StampList.TabStop = True
		Me.StampList.Visible = True
		Me.StampList.MultiColumn = False
		Me.StampList.Name = "StampList"
		Me.OK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OK.Text = "OK"
		Me.OK.Size = New System.Drawing.Size(73, 21)
		Me.OK.Location = New System.Drawing.Point(96, 300)
		Me.OK.TabIndex = 2
		Me.OK.BackColor = System.Drawing.SystemColors.Control
		Me.OK.CausesValidation = True
		Me.OK.Enabled = True
		Me.OK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OK.Cursor = System.Windows.Forms.Cursors.Default
		Me.OK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OK.TabStop = True
		Me.OK.Name = "OK"
		Me.txtStamp.AutoSize = False
		Me.txtStamp.Size = New System.Drawing.Size(317, 21)
		Me.txtStamp.Location = New System.Drawing.Point(12, 16)
		Me.txtStamp.TabIndex = 1
		Me.txtStamp.AcceptsReturn = True
		Me.txtStamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStamp.BackColor = System.Drawing.SystemColors.Window
		Me.txtStamp.CausesValidation = True
		Me.txtStamp.Enabled = True
		Me.txtStamp.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStamp.HideSelection = True
		Me.txtStamp.ReadOnly = False
		Me.txtStamp.Maxlength = 0
		Me.txtStamp.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStamp.MultiLine = False
		Me.txtStamp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStamp.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStamp.TabStop = True
		Me.txtStamp.Visible = True
		Me.txtStamp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStamp.Name = "txtStamp"
		Me.Cancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.Cancel
		Me.Cancel.Text = "Cancel"
		Me.Cancel.Size = New System.Drawing.Size(73, 21)
		Me.Cancel.Location = New System.Drawing.Point(184, 300)
		Me.Cancel.TabIndex = 0
		Me.Cancel.BackColor = System.Drawing.SystemColors.Control
		Me.Cancel.CausesValidation = True
		Me.Cancel.Enabled = True
		Me.Cancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Cancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Cancel.TabStop = True
		Me.Cancel.Name = "Cancel"
		Me.Controls.Add(Change)
		Me.Controls.Add(Del)
		Me.Controls.Add(Add_Renamed)
		Me.Controls.Add(StampList)
		Me.Controls.Add(OK)
		Me.Controls.Add(txtStamp)
		Me.Controls.Add(Cancel)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class
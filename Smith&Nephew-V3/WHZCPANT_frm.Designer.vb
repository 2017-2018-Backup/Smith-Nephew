<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHZCPANT_frm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.chkMedial = New System.Windows.Forms.CheckBox()
        Me.cboProximalElasticList = New System.Windows.Forms.ComboBox()
        Me.lblProximal = New System.Windows.Forms.Label()
        Me.cboDistalElasticList = New System.Windows.Forms.ComboBox()
        Me.lblDistal = New System.Windows.Forms.Label()
        Me.cboLengthList = New System.Windows.Forms.ComboBox()
        Me.lblDistalEOS = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(218, 112)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 17
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(126, 112)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 16
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkMedial
        '
        Me.chkMedial.Location = New System.Drawing.Point(155, 79)
        Me.chkMedial.Name = "chkMedial"
        Me.chkMedial.Size = New System.Drawing.Size(103, 24)
        Me.chkMedial.TabIndex = 15
        Me.chkMedial.Text = "Medial Zipper"
        Me.chkMedial.UseVisualStyleBackColor = True
        '
        'cboProximalElasticList
        '
        Me.cboProximalElasticList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProximalElasticList.FormattingEnabled = True
        Me.cboProximalElasticList.Location = New System.Drawing.Point(278, 43)
        Me.cboProximalElasticList.Name = "cboProximalElasticList"
        Me.cboProximalElasticList.Size = New System.Drawing.Size(121, 21)
        Me.cboProximalElasticList.TabIndex = 14
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(219, 47)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(53, 23)
        Me.lblProximal.TabIndex = 13
        Me.lblProximal.Text = "Proximal:"
        '
        'cboDistalElasticList
        '
        Me.cboDistalElasticList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDistalElasticList.FormattingEnabled = True
        Me.cboDistalElasticList.Location = New System.Drawing.Point(278, 9)
        Me.cboDistalElasticList.Name = "cboDistalElasticList"
        Me.cboDistalElasticList.Size = New System.Drawing.Size(121, 21)
        Me.cboDistalElasticList.TabIndex = 12
        '
        'lblDistal
        '
        Me.lblDistal.Location = New System.Drawing.Point(219, 12)
        Me.lblDistal.Name = "lblDistal"
        Me.lblDistal.Size = New System.Drawing.Size(42, 23)
        Me.lblDistal.TabIndex = 11
        Me.lblDistal.Text = "Distal :"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(91, 9)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(121, 21)
        Me.cboLengthList.TabIndex = 10
        '
        'lblDistalEOS
        '
        Me.lblDistalEOS.Location = New System.Drawing.Point(6, 12)
        Me.lblDistalEOS.Name = "lblDistalEOS"
        Me.lblDistalEOS.Size = New System.Drawing.Size(78, 23)
        Me.lblDistalEOS.TabIndex = 9
        Me.lblDistalEOS.Text = "Distal EOS to"
        '
        'WHZCPANT_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 145)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.chkMedial)
        Me.Controls.Add(Me.cboProximalElasticList)
        Me.Controls.Add(Me.lblProximal)
        Me.Controls.Add(Me.cboDistalElasticList)
        Me.Controls.Add(Me.lblDistal)
        Me.Controls.Add(Me.cboLengthList)
        Me.Controls.Add(Me.lblDistalEOS)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WHZCPANT_frm"
        Me.ShowIcon = False
        Me.Text = "WHZCPANT_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents chkMedial As CheckBox
    Friend WithEvents cboProximalElasticList As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents cboDistalElasticList As ComboBox
    Friend WithEvents lblDistal As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents lblDistalEOS As Label
End Class

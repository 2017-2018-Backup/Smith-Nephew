<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHZIPANK_frm
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
        Me.cboElasticList = New System.Windows.Forms.ComboBox()
        Me.lblProximal = New System.Windows.Forms.Label()
        Me.cboLengthList = New System.Windows.Forms.ComboBox()
        Me.lblAnkle = New System.Windows.Forms.Label()
        Me.lblTemplate = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(203, 78)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 13
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(104, 78)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkMedial
        '
        Me.chkMedial.Location = New System.Drawing.Point(205, 45)
        Me.chkMedial.Name = "chkMedial"
        Me.chkMedial.Size = New System.Drawing.Size(104, 24)
        Me.chkMedial.TabIndex = 11
        Me.chkMedial.Text = "Medial Zipper"
        Me.chkMedial.UseVisualStyleBackColor = True
        '
        'cboElasticList
        '
        Me.cboElasticList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboElasticList.FormattingEnabled = True
        Me.cboElasticList.Location = New System.Drawing.Point(249, 10)
        Me.cboElasticList.Name = "cboElasticList"
        Me.cboElasticList.Size = New System.Drawing.Size(121, 21)
        Me.cboElasticList.TabIndex = 10
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(193, 12)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(52, 20)
        Me.lblProximal.TabIndex = 9
        Me.lblProximal.Text = "Proximal:"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(65, 10)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(121, 21)
        Me.cboLengthList.TabIndex = 8
        '
        'lblAnkle
        '
        Me.lblAnkle.Location = New System.Drawing.Point(7, 12)
        Me.lblAnkle.Name = "lblAnkle"
        Me.lblAnkle.Size = New System.Drawing.Size(52, 20)
        Me.lblAnkle.TabIndex = 7
        Me.lblAnkle.Text = "Ankle to"
        '
        'lblTemplate
        '
        Me.lblTemplate.Location = New System.Drawing.Point(7, 45)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.Size = New System.Drawing.Size(100, 23)
        Me.lblTemplate.TabIndex = 14
        Me.lblTemplate.Text = "Template:"
        '
        'WHZIPANK_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(377, 111)
        Me.Controls.Add(Me.lblTemplate)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.chkMedial)
        Me.Controls.Add(Me.cboElasticList)
        Me.Controls.Add(Me.lblProximal)
        Me.Controls.Add(Me.cboLengthList)
        Me.Controls.Add(Me.lblAnkle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WHZIPANK_frm"
        Me.ShowIcon = False
        Me.Text = "WHZIPANK_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents chkMedial As CheckBox
    Friend WithEvents cboElasticList As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents lblAnkle As Label
    Friend WithEvents lblTemplate As Label
End Class

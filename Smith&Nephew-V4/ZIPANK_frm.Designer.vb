<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZIPANK_frm
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
        Me.lblAnkle = New System.Windows.Forms.Label()
        Me.cboLengthList = New System.Windows.Forms.ComboBox()
        Me.lblProximal = New System.Windows.Forms.Label()
        Me.cboElasticList = New System.Windows.Forms.ComboBox()
        Me.chkMedial = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblAnkle
        '
        Me.lblAnkle.Location = New System.Drawing.Point(5, 15)
        Me.lblAnkle.Name = "lblAnkle"
        Me.lblAnkle.Size = New System.Drawing.Size(52, 20)
        Me.lblAnkle.TabIndex = 0
        Me.lblAnkle.Text = "Ankle to"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(63, 13)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(121, 21)
        Me.cboLengthList.TabIndex = 1
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(191, 15)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(52, 20)
        Me.lblProximal.TabIndex = 2
        Me.lblProximal.Text = "Proximal:"
        '
        'cboElasticList
        '
        Me.cboElasticList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboElasticList.FormattingEnabled = True
        Me.cboElasticList.Location = New System.Drawing.Point(247, 13)
        Me.cboElasticList.Name = "cboElasticList"
        Me.cboElasticList.Size = New System.Drawing.Size(121, 21)
        Me.cboElasticList.TabIndex = 3
        '
        'chkMedial
        '
        Me.chkMedial.Location = New System.Drawing.Point(141, 48)
        Me.chkMedial.Name = "chkMedial"
        Me.chkMedial.Size = New System.Drawing.Size(104, 24)
        Me.chkMedial.TabIndex = 4
        Me.chkMedial.Text = "Medial Zipper"
        Me.chkMedial.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(102, 81)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(201, 81)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ZIPANK_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(377, 111)
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
        Me.Name = "ZIPANK_frm"
        Me.ShowIcon = False
        Me.Text = "ZIPANK_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblAnkle As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents cboElasticList As ComboBox
    Friend WithEvents chkMedial As CheckBox
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnOK As Button
End Class

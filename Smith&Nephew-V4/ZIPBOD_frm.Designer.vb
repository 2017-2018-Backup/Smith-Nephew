<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZIPBOD_frm
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
        Me.lblEOS = New System.Windows.Forms.Label()
        Me.cboLengthList = New System.Windows.Forms.ComboBox()
        Me.lblProximal = New System.Windows.Forms.Label()
        Me.cboElasticList = New System.Windows.Forms.ComboBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblEOS
        '
        Me.lblEOS.Location = New System.Drawing.Point(13, 12)
        Me.lblEOS.Name = "lblEOS"
        Me.lblEOS.Size = New System.Drawing.Size(50, 23)
        Me.lblEOS.TabIndex = 0
        Me.lblEOS.Text = "EOS to"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(70, 10)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(108, 21)
        Me.cboLengthList.TabIndex = 1
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(193, 12)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(54, 23)
        Me.lblProximal.TabIndex = 2
        Me.lblProximal.Text = "Proximal:"
        '
        'cboElasticList
        '
        Me.cboElasticList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboElasticList.FormattingEnabled = True
        Me.cboElasticList.Location = New System.Drawing.Point(254, 10)
        Me.cboElasticList.Name = "cboElasticList"
        Me.cboElasticList.Size = New System.Drawing.Size(108, 21)
        Me.cboElasticList.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(103, 55)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(189, 55)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ZIPBOD_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(373, 94)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cboElasticList)
        Me.Controls.Add(Me.lblProximal)
        Me.Controls.Add(Me.cboLengthList)
        Me.Controls.Add(Me.lblEOS)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ZIPBOD_frm"
        Me.ShowIcon = False
        Me.Text = "ZIPBOD_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblEOS As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents cboElasticList As ComboBox
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnOK As Button
End Class

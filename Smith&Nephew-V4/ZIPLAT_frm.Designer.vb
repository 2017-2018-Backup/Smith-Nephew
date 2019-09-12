<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZIPLAT_frm
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
        Me.cboProximal = New System.Windows.Forms.ComboBox()
        Me.lblDistal = New System.Windows.Forms.Label()
        Me.cboDistal = New System.Windows.Forms.ComboBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblEOS
        '
        Me.lblEOS.Location = New System.Drawing.Point(13, 13)
        Me.lblEOS.Name = "lblEOS"
        Me.lblEOS.Size = New System.Drawing.Size(44, 23)
        Me.lblEOS.TabIndex = 0
        Me.lblEOS.Text = "EOS to"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(64, 13)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(121, 21)
        Me.cboLengthList.TabIndex = 1
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(192, 13)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(55, 21)
        Me.lblProximal.TabIndex = 2
        Me.lblProximal.Text = "Proximal:"
        '
        'cboProximal
        '
        Me.cboProximal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProximal.FormattingEnabled = True
        Me.cboProximal.Location = New System.Drawing.Point(254, 13)
        Me.cboProximal.Name = "cboProximal"
        Me.cboProximal.Size = New System.Drawing.Size(121, 21)
        Me.cboProximal.TabIndex = 3
        '
        'lblDistal
        '
        Me.lblDistal.Location = New System.Drawing.Point(195, 55)
        Me.lblDistal.Name = "lblDistal"
        Me.lblDistal.Size = New System.Drawing.Size(55, 21)
        Me.lblDistal.TabIndex = 4
        Me.lblDistal.Text = "Distal:"
        '
        'cboDistal
        '
        Me.cboDistal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDistal.FormattingEnabled = True
        Me.cboDistal.Location = New System.Drawing.Point(254, 55)
        Me.cboDistal.Name = "cboDistal"
        Me.cboDistal.Size = New System.Drawing.Size(121, 21)
        Me.cboDistal.TabIndex = 5
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(111, 104)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(201, 103)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ZIPLAT_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(385, 136)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cboDistal)
        Me.Controls.Add(Me.lblDistal)
        Me.Controls.Add(Me.cboProximal)
        Me.Controls.Add(Me.lblProximal)
        Me.Controls.Add(Me.cboLengthList)
        Me.Controls.Add(Me.lblEOS)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ZIPLAT_frm"
        Me.ShowIcon = False
        Me.Text = "ZIPLAT_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblEOS As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents cboProximal As ComboBox
    Friend WithEvents lblDistal As Label
    Friend WithEvents cboDistal As ComboBox
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnOK As Button
End Class

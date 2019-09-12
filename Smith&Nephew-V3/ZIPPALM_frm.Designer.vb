<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZIPPALM_frm
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
        Me.cboProximal = New System.Windows.Forms.ComboBox()
        Me.lblProximal = New System.Windows.Forms.Label()
        Me.lblBelowWeb = New System.Windows.Forms.Label()
        Me.cboWebOffsetList = New System.Windows.Forms.ComboBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblEOS
        '
        Me.lblEOS.Location = New System.Drawing.Point(9, 13)
        Me.lblEOS.Name = "lblEOS"
        Me.lblEOS.Size = New System.Drawing.Size(45, 16)
        Me.lblEOS.TabIndex = 0
        Me.lblEOS.Text = "EOS to"
        '
        'cboLengthList
        '
        Me.cboLengthList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLengthList.FormattingEnabled = True
        Me.cboLengthList.Location = New System.Drawing.Point(61, 10)
        Me.cboLengthList.Name = "cboLengthList"
        Me.cboLengthList.Size = New System.Drawing.Size(121, 21)
        Me.cboLengthList.TabIndex = 1
        '
        'cboProximal
        '
        Me.cboProximal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProximal.FormattingEnabled = True
        Me.cboProximal.Location = New System.Drawing.Point(253, 10)
        Me.cboProximal.Name = "cboProximal"
        Me.cboProximal.Size = New System.Drawing.Size(121, 21)
        Me.cboProximal.TabIndex = 5
        '
        'lblProximal
        '
        Me.lblProximal.Location = New System.Drawing.Point(191, 13)
        Me.lblProximal.Name = "lblProximal"
        Me.lblProximal.Size = New System.Drawing.Size(55, 21)
        Me.lblProximal.TabIndex = 4
        Me.lblProximal.Text = "Proximal:"
        '
        'lblBelowWeb
        '
        Me.lblBelowWeb.Location = New System.Drawing.Point(175, 55)
        Me.lblBelowWeb.Name = "lblBelowWeb"
        Me.lblBelowWeb.Size = New System.Drawing.Size(71, 21)
        Me.lblBelowWeb.TabIndex = 6
        Me.lblBelowWeb.Text = "Below Web:"
        '
        'cboWebOffsetList
        '
        Me.cboWebOffsetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWebOffsetList.FormattingEnabled = True
        Me.cboWebOffsetList.Location = New System.Drawing.Point(253, 51)
        Me.cboWebOffsetList.Name = "cboWebOffsetList"
        Me.cboWebOffsetList.Size = New System.Drawing.Size(121, 21)
        Me.cboWebOffsetList.TabIndex = 7
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(199, 100)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(109, 101)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ZIPPALM_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(385, 136)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cboWebOffsetList)
        Me.Controls.Add(Me.lblBelowWeb)
        Me.Controls.Add(Me.cboProximal)
        Me.Controls.Add(Me.lblProximal)
        Me.Controls.Add(Me.cboLengthList)
        Me.Controls.Add(Me.lblEOS)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ZIPPALM_frm"
        Me.ShowIcon = False
        Me.Text = "ZIPPALM_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblEOS As Label
    Friend WithEvents cboLengthList As ComboBox
    Friend WithEvents cboProximal As ComboBox
    Friend WithEvents lblProximal As Label
    Friend WithEvents lblBelowWeb As Label
    Friend WithEvents cboWebOffsetList As ComboBox
    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
End Class

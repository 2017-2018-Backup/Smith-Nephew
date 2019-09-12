<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class KNEECTR_frm
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
        Me.cboContracture = New System.Windows.Forms.ComboBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnDraw = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cboContracture
        '
        Me.cboContracture.FormattingEnabled = True
        Me.cboContracture.Location = New System.Drawing.Point(8, 12)
        Me.cboContracture.Name = "cboContracture"
        Me.cboContracture.Size = New System.Drawing.Size(156, 21)
        Me.cboContracture.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(8, 51)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnDraw
        '
        Me.btnDraw.Location = New System.Drawing.Point(90, 51)
        Me.btnDraw.Name = "btnDraw"
        Me.btnDraw.Size = New System.Drawing.Size(75, 23)
        Me.btnDraw.TabIndex = 2
        Me.btnDraw.Text = "Draw"
        Me.btnDraw.UseVisualStyleBackColor = True
        '
        'KNEECTR_frm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(176, 84)
        Me.Controls.Add(Me.btnDraw)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cboContracture)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "KNEECTR_frm"
        Me.ShowIcon = False
        Me.Text = "KNEECTR_frm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cboContracture As ComboBox
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnDraw As Button
End Class

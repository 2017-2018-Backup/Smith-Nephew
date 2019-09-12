Public Module BodyLeg
    Public BDLeg As New BODYLEG_frm
End Module
Public Class BODYLEG_frm
    Dim frmBodyLeftLeg As New bdlegdia
    Dim frmBodyRightLeg As New BDLEGRight_frm

    Private Sub BODYLEG_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        frmBodyLeftLeg.TopLevel = False
        frmBodyLeftLeg.FormBorderStyle = FormBorderStyle.None
        frmBodyLeftLeg.Dock = DockStyle.Fill
        frmBodyLeftLeg.Visible = True

        BodyLegTabControl.TabPages(0).Controls.Add(frmBodyLeftLeg)
        BodyLegTabControl.TabPages(0).Text = "Left"
        frmBodyLeftLeg.Show()

        frmBodyRightLeg.TopLevel = False
        frmBodyRightLeg.FormBorderStyle = FormBorderStyle.None
        frmBodyRightLeg.Dock = DockStyle.Fill
        frmBodyRightLeg.Visible = True

        BodyLegTabControl.TabPages(1).Controls.Add(frmBodyRightLeg)
        BodyLegTabControl.TabPages(1).Text = "Right"
        frmBodyRightLeg.Show()
    End Sub
End Class
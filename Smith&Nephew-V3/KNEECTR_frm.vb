Public Class KNEECTR_frm
    Public g_sCaption As String
    Public g_bIsCancel As Boolean
    Public g_sContracture As String

    Private Sub KNEECTR_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Hide()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
        Me.Text = g_sCaption
        'Contractures Combo
        cboContracture.Items.Add(" ")
        cboContracture.Items.Add("10 to 35 Degrees")
        cboContracture.Items.Add("36 to 70 Degrees")
        cboContracture.Items.Add("71 Degrees & Over")
        g_bIsCancel = True
        g_sContracture = ""
        Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
        Me.Activate()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        g_bIsCancel = True
        Me.Close()
    End Sub

    Private Sub btnDraw_Click(sender As Object, e As EventArgs) Handles btnDraw.Click
        g_bIsCancel = False
        g_sContracture = cboContracture.Text
        Me.Close()
    End Sub
End Class
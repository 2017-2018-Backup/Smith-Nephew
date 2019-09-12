Public Class ZIPEDIT_frm
    Public g_sCaption As String
    Public g_bIsCancel As Boolean
    Public g_sSymbolText As String

    Private Sub ZIPEDIT_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Hide()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
            Me.Text = g_sCaption
            g_bIsCancel = True
            txtSymbol.Text = g_sSymbolText
            Show()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
            Me.Activate()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        g_bIsCancel = False
        g_sSymbolText = txtSymbol.Text
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        g_bIsCancel = True
        Me.Close()
    End Sub
End Class
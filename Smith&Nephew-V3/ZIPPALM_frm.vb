Public Class ZIPPALM_frm
    Public g_sCaption As String
    Public g_bIsCancel As Boolean
    Public g_sZipLength As String
    Public g_sElasticProximal As String
    Public g_sWebOffset As String
    Public g_nAge As Double

    Private Sub ZIPPALM_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Hide()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
        Me.Text = g_sCaption
        'Lenght List Combo
        cboLengthList.Items.Add("Standard")
        cboLengthList.Items.Add("Give a length")
        cboLengthList.Items.Add("Selected Point")
        cboLengthList.SelectedIndex = 0

        ''Proximal Elastic List Combo
        If (g_nAge < 10) Then
            cboProximal.Items.Add("3/8" & Chr(34) & " Elastic")
        Else
            cboProximal.Items.Add("3/4" & Chr(34) & "Elastic")
        End If
        cboProximal.Items.Add("3/8" & Chr(34) & " Elastic")
        cboProximal.Items.Add("3/4" & Chr(34) & "Elastic")
        cboProximal.Items.Add("1½" & Chr(34) & "Elastic")
        cboProximal.Items.Add("No Elastic")
        cboProximal.SelectedIndex = 0

        ''Weg Offset List Combo
        cboWebOffsetList.Items.Add("1-1/8" & Chr(34))
        cboWebOffsetList.Items.Add("1-1/8" & Chr(34))
        cboWebOffsetList.Items.Add("3/4" & Chr(34))
        cboWebOffsetList.SelectedIndex = 0

        g_bIsCancel = True
        g_sZipLength = ""
        g_sElasticProximal = ""
        g_sWebOffset = ""
        Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
        Me.Activate()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        g_bIsCancel = False
        g_sZipLength = cboLengthList.Text
        g_sElasticProximal = cboProximal.Text
        g_sWebOffset = cboWebOffsetList.Text
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        g_bIsCancel = True
        Me.Close()
    End Sub
End Class
Public Class ZIPANK_frm
    Public g_sCaption As String
    Public g_bIsCancel As Boolean
    Public g_sZipLength As String
    Public g_sElastic As String
    Public g_bIsMedialZipper As Boolean
    Public g_nElasticType As Integer

    Private Sub ZIPANK_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Hide()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
        Me.Text = g_sCaption
        'Lenght List Combo
        cboLengthList.Items.Add("EOS")
        cboLengthList.Items.Add("Selected Point")
        cboLengthList.SelectedIndex = 0
        ''Elastic List Combo
        If g_nElasticType = 2 Then
            cboElasticList.Items.Add("No Elastic")
        Else
            cboElasticList.Items.Add("3/4" & Chr(34) & "Elastic")
        End If
        cboElasticList.Items.Add("3/8" & Chr(34) & " Elastic")
        cboElasticList.Items.Add("3/4" & Chr(34) & "Elastic")
        cboElasticList.Items.Add("1½" & Chr(34) & "Elastic")
        cboElasticList.Items.Add("No Elastic")
        cboElasticList.SelectedIndex = 0

        g_bIsCancel = True
        g_sZipLength = ""
        g_sElastic = ""
        g_bIsMedialZipper = False
        Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
        Me.Activate()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        g_bIsCancel = False
        g_sZipLength = cboLengthList.Text
        g_sElastic = cboElasticList.Text
        g_bIsMedialZipper = chkMedial.Checked
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        g_bIsCancel = True
        Me.Close()
    End Sub
End Class
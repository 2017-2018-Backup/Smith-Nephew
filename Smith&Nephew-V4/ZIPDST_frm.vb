Public Class ZIPDST_frm
    Public g_sCaption As String
    Public g_bIsCancel As Boolean
    Public g_sZipLength As String
    Public g_sElasticProximal As String
    Public g_sElasticDistal As String
    Public g_bIsMedialZipper As Boolean

    Private Sub ZIPDST_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Hide()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
            Me.Text = g_sCaption
            'Lenght List Combo
            cboLengthList.Items.Add("Give a length")
            cboLengthList.Items.Add("Selected Point")
            cboLengthList.SelectedIndex = 0

            ''Proximal Elastic List Combo
            cboProximalElasticList.Items.Add("No Elastic")
            cboProximalElasticList.Items.Add("3/8" & Chr(34) & " Elastic")
            cboProximalElasticList.Items.Add("3/4" & Chr(34) & "Elastic")
            cboProximalElasticList.Items.Add("1½" & Chr(34) & "Elastic")
            cboProximalElasticList.Items.Add("No Elastic")
            cboProximalElasticList.SelectedIndex = 0
            ''Distal Elastic List Combo
            cboDistalElasticList.Items.Add("No Elastic")
            cboDistalElasticList.Items.Add("3/8" & Chr(34) & " Elastic")
            cboDistalElasticList.Items.Add("3/4" & Chr(34) & "Elastic")
            cboDistalElasticList.Items.Add("1½" & Chr(34) & "Elastic")
            cboDistalElasticList.Items.Add("No Elastic")
            cboDistalElasticList.SelectedIndex = 0

            g_bIsCancel = True
            g_sZipLength = ""
            g_sElasticProximal = ""
            g_sElasticDistal = ""
            g_bIsMedialZipper = False
            Show()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
            Me.Activate()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        g_bIsCancel = False
        g_sZipLength = cboLengthList.Text
        g_sElasticProximal = cboProximalElasticList.Text
        g_sElasticDistal = cboDistalElasticList.Text
        g_bIsMedialZipper = chkMedial.Checked
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        g_bIsCancel = True
        Me.Close()
    End Sub
End Class
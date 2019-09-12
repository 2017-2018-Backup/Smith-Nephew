Public Module BodyLeg
    Public BDLeg As New BODYLEG_frm
End Module
Public Class BODYLEG_frm
    Dim frmBodyLeftLeg As New bdlegdia
    Dim frmBodyRightLeg As New BDLEGRight_frm
    Dim g_bIsCloseDlg As Boolean
    Private Sub BODYLEG_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            g_bIsCloseDlg = False
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
        Catch ex As Exception
            g_bIsCloseDlg = True
            Me.Close()
            BodyMain.BodyMainDlg.Close()
        End Try
    End Sub

    Private Sub BODYLEG_frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsCloseDlg = True Then
            g_bIsCloseDlg = False
            Exit Sub
        End If
        If frmBodyLeftLeg.IsDisposed() = False Then
            If frmBodyLeftLeg.PR_CloseBodyLeftLeg() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If frmBodyRightLeg.IsDisposed() = False Then
            If frmBodyRightLeg.PR_CloseBodyRightLeg() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub
    Public Function PR_CloseBodyLeg() As Boolean
        If frmBodyLeftLeg.IsDisposed() = False Then
            If frmBodyLeftLeg.PR_CloseBodyLeftLeg() = False Then
                Return False
            End If
        End If
        If frmBodyRightLeg.IsDisposed() = False Then
            If frmBodyRightLeg.PR_CloseBodyRightLeg() = False Then
                Return False
            End If
        End If
        Return True
    End Function
End Class
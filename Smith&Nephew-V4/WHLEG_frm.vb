Public Module WaistLeg
    Public WHLeg As New whleg
End Module
Public Class whleg
    Dim frmLeftLeg As New whlegdia
    Dim frmRightLeg As New WHRightLeg
    Dim g_bIsCloseDlg As Boolean
    Private Sub whleg_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            g_bIsCloseDlg = False
            frmLeftLeg.TopLevel = False
            frmLeftLeg.FormBorderStyle = FormBorderStyle.None
            frmLeftLeg.Dock = DockStyle.Fill
            frmLeftLeg.Visible = True

            WHLegTabControl.TabPages(0).Controls.Add(frmLeftLeg)
            WHLegTabControl.TabPages(0).Text = "Left"
            frmLeftLeg.Show()

            frmRightLeg.TopLevel = False
            frmRightLeg.FormBorderStyle = FormBorderStyle.None
            frmRightLeg.Dock = DockStyle.Fill
            frmRightLeg.Visible = True

            WHLegTabControl.TabPages(1).Controls.Add(frmRightLeg)
            WHLegTabControl.TabPages(1).Text = "Right"
            frmRightLeg.Show()
        Catch ex As Exception
            g_bIsCloseDlg = True
            Me.Close()
            WaistMain.WaistMainDlg.Close()
        End Try
    End Sub

    Private Sub whleg_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsCloseDlg = True Then
            g_bIsCloseDlg = False
            Exit Sub
        End If
        If frmLeftLeg.IsDisposed() = False Then
            If frmLeftLeg.PR_CloseWaistLeftLeg() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If frmRightLeg.IsDisposed() = False Then
            If frmRightLeg.PR_CloseWaistRightLeg() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Public Function PR_CloseWaistLegDialog() As Boolean
        If frmLeftLeg.IsDisposed() = False Then
            If frmLeftLeg.PR_CloseWaistLeftLeg() = False Then
                Return False
            End If
        End If
        If frmRightLeg.IsDisposed() = False Then
            If frmRightLeg.PR_CloseWaistRightLeg() = False Then
                Return False
            End If
        End If
        Return True
    End Function

    'Private Sub WHLegTabControl_TabIndexChanged(sender As Object, e As EventArgs) Handles WHLegTabControl.TabIndexChanged
    '    Dim TabIndex As Integer = WHLegTabControl.SelectedIndex
    '    If TabIndex = 0 Then
    '        'WHLegTabControl.TabPages(0).Height = frmLeftLeg.Height
    '    ElseIf TabIndex = 1 Then
    '        'WHLegTabControl.TabPages(1).Height = frmRightLeg.Height
    '    End If
    'End Sub
End Class
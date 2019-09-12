Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Public Module VestMain
    Public VestMainDlg As New VESTMAIN_frm
    Public g_bIsSetArmFabric As Boolean
    Public g_bIsSetTorsoFabric As Boolean
    Public g_sVestFabric As String
End Module
Public Class VESTMAIN_frm
    Dim frmVestBody As New vestdia
    Dim frmArm As New VESTARM_frm
    Dim torso As New torsodia
    Dim g_bIsCloseDlg As Boolean
    Private Sub VESTMAIN_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            g_bIsSetArmFabric = False
            g_bIsSetTorsoFabric = False
            g_bIsCloseDlg = False
            g_sVestFabric = ""
            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim obj As New BlockCreation.BlockCreation
            Dim blkId As ObjectId = New ObjectId()
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Can't find Patient Details", 48, "Vest Details Dialog")
                g_bIsCloseDlg = True
                Me.Close()
                Exit Sub
            End If

            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Vest Details Dialog")
                g_bIsCloseDlg = True
                Me.Close()
                Exit Sub
            End If
            txtDiagnosis.Text = diagnosis
            txtFileNo.Text = fileNo
            txtPatientName.Text = patient
            txtinchflag.Text = units

            txtAge1.Text = age
            txtSex1.Text = sex
            txtWorkOrder1.Text = workOrder
            txtDesigner.Text = tempEng
            txtTempDate.Text = tempDate

            Dim resbuf As New ResultBuffer
            resbuf = GetXrecord("NewPatient", "NEWPATIENTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim pattern As String = "dd-MM-yyyy"
                Dim parsedDate As DateTime
                DateTime.TryParseExact(arr(8).Value, pattern, System.Globalization.CultureInfo.CurrentCulture,
                                      System.Globalization.DateTimeStyles.None, parsedDate)
                DateTimePicker1.Value = parsedDate
                Dim startTime As DateTime = Convert.ToDateTime(DateTimePicker1.Value)
                Dim endTime As DateTime = DateTime.Today
                Dim span As TimeSpan = endTime.Subtract(startTime)
                Dim totalDays As Double = span.TotalDays
                Dim totalYears As Double = Math.Truncate(totalDays / 365)
                txtAge1.Text = totalYears.ToString()
            End If

            Dim Rect As Rectangle
            Rect = VestTabControl.ClientRectangle

            frmVestBody.TopLevel = False
            frmVestBody.FormBorderStyle = FormBorderStyle.None
            frmVestBody.Dock = DockStyle.Fill
            frmVestBody.Visible = True

            VestTabControl.TabPages(0).Controls.Add(frmVestBody)
            VestTabControl.TabPages(0).Text = "Vest"
            frmVestBody.Show()

            frmArm.TopLevel = False
            frmArm.FormBorderStyle = FormBorderStyle.None
            frmArm.Dock = DockStyle.Fill
            frmArm.Visible = True

            frmArm.g_sbIsBody = False
            VestTabControl.TabPages(1).Controls.Add(frmArm)
            VestTabControl.TabPages(1).Text = "Arm"
            frmArm.Show()


            torso.TopLevel = False
            torso.FormBorderStyle = FormBorderStyle.None
            torso.Dock = DockStyle.Fill
            torso.Visible = True

            VestTabControl.TabPages(2).Controls.Add(torso)
            VestTabControl.TabPages(2).Text = "Torso"
            VestTabControl.TabPages(2).Height = 500
            VestTabControl.TabPages(2).Width = 100
            torso.Show()
        Catch ex As Exception
            g_bIsCloseDlg = True
            Me.Close()
        End Try
    End Sub

    Private Sub VestTabControl_TabIndexChanged(sender As Object, e As EventArgs) Handles VestTabControl.TabIndexChanged
        Dim TabIndex As Integer = VestTabControl.SelectedIndex
        If TabIndex = 0 Then
            VestTabControl.TabPages(0).Height = frmVestBody.Height
        ElseIf TabIndex = 1 Then
            VestTabControl.TabPages(1).Height = frmArm.Height
        ElseIf TabIndex = 2 Then
            VestTabControl.TabPages(1).Height = torso.Height
        End If

        Select Case VestTabControl.SelectedIndex
            Case 0
                Me.Size = New Size(300, 200)
            Case 1
                Me.Size = New Size(400, 400)
            Case 1
                Me.Size = New Size(500, 200)
        End Select
    End Sub
    Public Function GetXrecord(Optional ByVal dictName As String = "SNInfo", Optional ByVal key As String = "SNDIC") As ResultBuffer
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim NOD As DBDictionary = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
            If Not NOD.Contains(dictName) Then Return Nothing
            Dim dict As DBDictionary = TryCast(tr.GetObject(NOD.GetAt(dictName), OpenMode.ForRead), DBDictionary)
            If dict Is Nothing OrElse Not dict.Contains(key) Then Return Nothing
            Dim xRec As Xrecord = TryCast(tr.GetObject(dict.GetAt(key), OpenMode.ForRead), Xrecord)
            If xRec Is Nothing Then Return Nothing
            Return xRec.Data
        End Using
    End Function

    Private Sub VestTabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles VestTabControl.SelectedIndexChanged
        Dim TabIndex As Integer = VestTabControl.SelectedIndex
        If TabIndex = 1 Then
            frmArm.PR_SetArmFabricValue(g_sVestFabric)
        End If
        If TabIndex = 2 Then
            torso.PR_SetTorsoFabricValue(g_sVestFabric)
        End If
        'If TabIndex = 1 And g_bIsSetArmFabric = True Then
        '    g_bIsSetArmFabric = False
        '    frmArm.PR_SetArmFabricValue()
        'End If
        'If TabIndex = 2 And g_bIsSetTorsoFabric = True Then
        '    g_bIsSetTorsoFabric = False
        '    torso.PR_SetTorsoFabricValue()
        'End If
    End Sub

    Private Sub VESTMAIN_frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsCloseDlg = True Then
            g_bIsCloseDlg = False
            Exit Sub
        End If
        If frmVestBody.IsDisposed() = False Then
            If frmVestBody.PR_CloseVestBodyDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If frmArm.IsDisposed() = False Then
            If frmArm.PR_CloseVestArmDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If torso.IsDisposed() = False Then
            If torso.PR_CloseTorsoDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub
End Class
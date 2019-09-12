Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Public Module LegMain
    Public LegMainDlg As New LEGMAIN_frm
End Module
Public Class LEGMAIN_frm
    'Dim frmLef As New 
    Dim frmLegLeft As New lglegdia
    Dim frmLegRight As New LEGRight_frm
    Dim g_bIsCloseDlg As Boolean

    Private Sub LEGMAIN_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            g_bIsCloseDlg = False
            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim blkId As ObjectId = New ObjectId()
            Dim obj As New BlockCreation.BlockCreation
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Patient Details have not been entered", 48, "Leg Details Dialog")
                g_bIsCloseDlg = True
                Me.Close()
                Exit Sub
            End If
            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Leg Details Dialog")
                g_bIsCloseDlg = True
                Me.Close()
                Exit Sub
            End If
            txtPatientName.Text = patient
            txtFileNo.Text = fileNo
            txtDiagnosis.Text = diagnosis
            txtSex.Text = sex
            txtAge.Text = age
            txtUnits.Text = units

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
                ''CurrentYear - BirthDate  
                Dim startTime As DateTime = Convert.ToDateTime(DateTimePicker1.Value)
                Dim endTime As DateTime = DateTime.Today
                Dim span As TimeSpan = endTime.Subtract(startTime)
                Dim totalDays As Double = span.TotalDays
                Dim totalYears As Double = Math.Truncate(totalDays / 365)
                txtAge.Text = totalYears.ToString()
            End If

            Dim Rect, r1 As Rectangle
            Rect = LegTabControl.ClientRectangle

            frmLegLeft.TopLevel = False
            frmLegLeft.FormBorderStyle = FormBorderStyle.None
            frmLegLeft.Dock = DockStyle.Fill
            frmLegLeft.Visible = True

            LegTabControl.TabPages(0).Controls.Add(frmLegLeft)
            LegTabControl.TabPages(0).Text = "Left"
            frmLegLeft.Show()

            frmLegRight.TopLevel = False
            frmLegRight.FormBorderStyle = FormBorderStyle.None
            frmLegRight.Dock = DockStyle.Fill
            frmLegRight.Visible = True

            LegTabControl.TabPages(1).Controls.Add(frmLegRight)
            LegTabControl.TabPages(1).Text = "Right"
            frmLegRight.Show()
        Catch ex As Exception
            g_bIsCloseDlg = True
            Me.Close()
        End Try
    End Sub

    Private Sub LegTabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LegTabControl.SelectedIndexChanged
        lblLeg.Text = "left leg"
        If LegTabControl.SelectedIndex = 1 Then
            lblLeg.Text = "right leg"
        End If
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

    Private Sub LEGMAIN_frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsCloseDlg = True Then
            g_bIsCloseDlg = False
            Exit Sub
        End If
        If frmLegLeft.IsDisposed() = False Then
            If frmLegLeft.PR_CloseLeftLegDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If frmLegRight.IsDisposed() = False Then
            If frmLegRight.PR_CloseRightLegDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub
End Class
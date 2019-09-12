Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Public Module BodyMain
    Public BodyMainDlg As New BODYMAIN_frm
    Public g_bIsBodyLoad As Boolean
    Public g_sBodyFabric As String
End Module
Public Class BODYMAIN_frm
    Dim frmBody As New bodysuit
    'Dim frmLeg As New whlegdia
    Dim frmLeg As New BODYLEG_frm
    Dim frmArm As New VESTARM_frm

    Private Sub BODYMAIN_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        g_bIsBodyLoad = False
        g_sBodyFabric = ""
        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("Patient Details have not been entered", 48, "Bodysuit Details Dialog")
            Me.Close()
            Exit Sub
        End If
        obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
        If fileNo = "" Then
            MsgBox("Please enter Patient Details", 48, "Bodysuit Details Dialog")
            Me.Close()
            'BodyMain.BodyMainDlg.Hide()
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

        Dim Rect As Rectangle
        Rect = WaistTabControl.ClientRectangle

        frmBody.TopLevel = False
        frmBody.FormBorderStyle = FormBorderStyle.None
        frmBody.Dock = DockStyle.Fill
        frmBody.Visible = True

        WaistTabControl.TabPages(0).Controls.Add(frmBody)
        WaistTabControl.TabPages(0).Text = "Body"
        frmBody.Show()

        frmLeg.TopLevel = False
        frmLeg.FormBorderStyle = FormBorderStyle.None
        frmLeg.Dock = DockStyle.Fill
        frmLeg.Visible = True

        WaistTabControl.TabPages(1).Controls.Add(frmLeg)
        WaistTabControl.TabPages(1).Text = "Leg"
        frmLeg.Show()

        frmArm.TopLevel = False
        frmArm.FormBorderStyle = FormBorderStyle.None
        frmArm.Dock = DockStyle.Fill
        frmArm.g_sbIsBody = True
        frmArm.Visible = True

        WaistTabControl.TabPages(2).Controls.Add(frmArm)
        WaistTabControl.TabPages(2).Text = "Arm"
        WaistTabControl.TabPages(2).Height = 500
        WaistTabControl.TabPages(2).Width = 100
        frmArm.Show()
        'BodyMain.BodyMainDlg.Hide()
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

    Private Sub WaistTabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WaistTabControl.SelectedIndexChanged
        Dim TabIndex As Integer = WaistTabControl.SelectedIndex
        If TabIndex = 0 And g_bIsBodyLoad = True Then
            g_bIsBodyLoad = False
            frmBody.PR_EnableDrawBody()
        End If
        If TabIndex = 2 Then
            frmArm.PR_SetArmFabricValue(g_sBodyFabric)
        End If
    End Sub
End Class
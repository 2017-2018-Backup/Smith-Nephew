Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Public Module ManGlvMain
    Public ManGlvMainDlg As New MANGLOVEMAIN_frm
End Module

Public Class MANGLOVEMAIN_frm
    Dim ManGlvLeft As New manglove
    Dim ManGlvRight As New MANGLOVERight_frm
    Dim WebSpcLeft As New webspacr
    Dim WebSpcRight As New WEBSPACRRight_frm

    Private Sub MANGLOVEMAIN_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Position to center of screen
        Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
        Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

        Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
        Dim blkId As ObjectId = New ObjectId()
        Dim obj As New BlockCreation.BlockCreation
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("Patient Details have not been entered", 48, "ManGlove Details Dialog")
            Me.Close()
            Exit Sub
        End If
        obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
        If fileNo = "" Then
            MsgBox("Please enter Patient Details", 48, "ManGlove Details Dialog")
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
        Rect = ManGlvTabControl.ClientRectangle

        ManGlvLeft.TopLevel = False
        ManGlvLeft.FormBorderStyle = FormBorderStyle.None
        ManGlvLeft.Dock = DockStyle.Fill
        ManGlvLeft.Visible = True

        ManGlvTabControl.TabPages(0).Controls.Add(ManGlvLeft)
        ManGlvTabControl.TabPages(0).Text = "Left Glove"
        ManGlvLeft.Show()

        WebSpcLeft.TopLevel = False
        WebSpcLeft.FormBorderStyle = FormBorderStyle.None
        WebSpcLeft.Dock = DockStyle.Fill
        WebSpcLeft.Visible = True

        ManGlvTabControl.TabPages(1).Controls.Add(WebSpcLeft)
        ManGlvTabControl.TabPages(1).Text = "Left Spacers"
        WebSpcLeft.Show()

        ManGlvRight.TopLevel = False
        ManGlvRight.FormBorderStyle = FormBorderStyle.None
        ManGlvRight.Dock = DockStyle.Fill
        ManGlvRight.Visible = True

        ManGlvTabControl.TabPages(2).Controls.Add(ManGlvRight)
        ManGlvTabControl.TabPages(2).Text = "Right Glove"
        ManGlvRight.Show()

        WebSpcRight.TopLevel = False
        WebSpcRight.FormBorderStyle = FormBorderStyle.None
        WebSpcRight.Dock = DockStyle.Fill
        WebSpcRight.Visible = True

        ManGlvTabControl.TabPages(3).Controls.Add(WebSpcRight)
        ManGlvTabControl.TabPages(3).Text = "Right Spacers"
        WebSpcRight.Show()
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

End Class
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Public Module WaistMain
    Public WaistMainDlg As New WHMAIN_frm
    Structure XY
        Dim X As Double
        Dim Y As Double
    End Structure

    Public Const PI As Double = 3.141592654

    Public xyWHCutInsert As XY
    Public g_sFoldHt As String
    Public g_sLegStyle As String
    Public sCrotchStyle As String
    Public g_sBody As String
    Public g_sAnkleTape As String
    Public LeftLeg As Boolean
    Public nLastTape As Double

    Public nTOSCir As Double
    Public nWaistCir As Double
    Public nMidPointCir As Double
    Public nLargestCir As Double
    Public nWaistHt As Double
    Public nLargestGivenRed As Double
    Public nMidPointGivenRed As Double
    Public nLeftThighCir As Double
    Public nRightThighCir As Double
    Public nThighGivenRed As Double
    Public nWaistGivenRed As Double
    Public nTOSGivenRed As Double
    Public nTOSHt As Double
    Public g_bIsNeedLoad As Boolean
End Module
Public Class WHMAIN_frm
    Dim frmWaistBody As New whboddia
    'Dim frmLeg As New whlegdia
    Dim frmLeg As New whleg
    Dim Figure As New whfigure
    Dim g_bIsCloseDlg As Boolean

    Private Sub WHMAIN_frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.

            g_bIsNeedLoad = False
            g_bIsCloseDlg = False
            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim blkId As ObjectId = New ObjectId()
            Dim obj As New BlockCreation.BlockCreation
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Patient Details have not been entered", 48, "Waist Details Dialog")
                g_bIsCloseDlg = True
                Me.Close()
                Exit Sub
            End If
            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Waist Details Dialog")
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

            Dim Rect As Rectangle
            Rect = WaistTabControl.ClientRectangle

            frmWaistBody.TopLevel = False
            frmWaistBody.FormBorderStyle = FormBorderStyle.None
            frmWaistBody.Dock = DockStyle.Fill
            frmWaistBody.Visible = True

            WaistTabControl.TabPages(0).Controls.Add(frmWaistBody)
            WaistTabControl.TabPages(0).Text = "Body"
            frmWaistBody.Show()

            frmLeg.TopLevel = False
            frmLeg.FormBorderStyle = FormBorderStyle.None
            frmLeg.Dock = DockStyle.Fill
            frmLeg.Visible = True

            WaistTabControl.TabPages(1).Controls.Add(frmLeg)
            WaistTabControl.TabPages(1).Text = "Leg"
            frmLeg.Show()

            Figure.TopLevel = False
            Figure.FormBorderStyle = FormBorderStyle.None
            Figure.Dock = DockStyle.Fill
            Figure.Visible = True

            WaistTabControl.TabPages(2).Controls.Add(Figure)
            WaistTabControl.TabPages(2).Text = "Figure"
            WaistTabControl.TabPages(2).Height = 500
            WaistTabControl.TabPages(2).Width = 100
            Figure.Show()
        Catch ex As Exception
            g_bIsCloseDlg = True
            Me.Close()
        End Try
    End Sub

    Private Sub WaistTabControl_TabIndexChanged(sender As Object, e As EventArgs) Handles WaistTabControl.TabIndexChanged
        Dim TabIndex As Integer = WaistTabControl.SelectedIndex
        If TabIndex = 0 Then
            WaistTabControl.TabPages(0).Height = frmWaistBody.Height
        ElseIf TabIndex = 1 Then
            WaistTabControl.TabPages(1).Height = frmLeg.Height
        ElseIf TabIndex = 2 Then
            WaistTabControl.TabPages(1).Height = Figure.Height
            Figure.Show()
        End If

        Select Case WaistTabControl.SelectedIndex
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

    'To Get Insertion point from user
    Private Sub PR_GetInsertionPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Indicate Start Point "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(WaistMain.xyWHCutInsert, ptStart.X, ptStart.Y)
    End Sub
    Sub PR_MakeXY(ByRef xyReturn As WaistMain.XY, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.Y = y
    End Sub
    Private Function PR_GetBlkAttributeValues() As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If Not acBlkTbl.Has("WAISTBODY") Then
                MsgBox("Can't find WAISTBODY symbol to update!", 16, "Waist - Dialogue")
                Return False
            End If
            Dim blkRecIdBody As ObjectId = ObjectId.Null
            blkRecIdBody = acBlkTbl("WAISTBODY")
            If blkRecIdBody <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecIdBody, OpenMode.ForRead)
                Dim blkRef As BlockReference = New BlockReference(New Point3d(0, 0, 0), blkRecIdBody)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If acAttDef.Tag.ToUpper().Equals("FoldHt", StringComparison.InvariantCultureIgnoreCase) Then
                            g_sFoldHt = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                            g_sStyle = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("TOSCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nTOSCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("WaistCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nWaistCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("MidPointCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nMidPointCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("LargestCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nLargestCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("WaistHt", StringComparison.InvariantCultureIgnoreCase) Then
                            nWaistHt = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("LargestGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                            nLargestGivenRed = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("MidPointGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                            nMidPointGivenRed = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("LeftThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nLeftThighCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("RightThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                            nRightThighCir = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("ThighGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                            nThighGivenRed = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("WaistGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                            nWaistGivenRed = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("TOSGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                            nTOSGivenRed = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("CrotchStyle", StringComparison.InvariantCultureIgnoreCase) Then
                            sCrotchStyle = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("TOSHt", StringComparison.InvariantCultureIgnoreCase) Then
                            nTOSHt = Val(acAttDef.TextString)
                        ElseIf acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                            WaistMain.g_sLegStyle = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("Body", StringComparison.InvariantCultureIgnoreCase) Then
                            g_sBody = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("AnkleTape", StringComparison.InvariantCultureIgnoreCase) Then
                            g_sAnkleTape = acAttDef.TextString
                        End If
                    End If
                Next

            End If

            Dim strLegBlkName As String = ""
            If acBlkTbl.Has("WAISTLEFTLEG") Then
                LeftLeg = True
                strLegBlkName = "WAISTLEFTLEG"
            ElseIf acBlkTbl.Has("WAISTRIGHTLEG") Then
                LeftLeg = False
                strLegBlkName = "WAISTRIGHTLEG"
            Else
                MsgBox("Can't find WAIST LEG symbol to update!", 16, "Waist - Dialogue")
                Return False
            End If
            Dim blkRecIdLeft As ObjectId = acBlkTbl(strLegBlkName)
            If blkRecIdLeft <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecIdLeft, OpenMode.ForRead)
                Dim blkRef As BlockReference = New BlockReference(New Point3d(0, 0, 0), blkRecIdLeft)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If acAttDef.Tag.ToUpper().Equals("LastTape", StringComparison.InvariantCultureIgnoreCase) Then
                            nLastTape = Val(acAttDef.TextString)
                        End If
                    End If
                Next
            End If
        End Using
        Return True
    End Function
    Function FNRound(ByVal nNumber As Single) As Short
        'E.G.
        '    round(1.35) = 1
        '    round(1.55) = 2
        '    round(2.50) = 3
        '    round(-2.50) = -3
        '
        Dim nInt, nSign As Short
        nSign = System.Math.Sign(nNumber)
        nNumber = System.Math.Abs(nNumber)
        nInt = Int(nNumber)
        If (nNumber - nInt) >= 0.5 Then
            FNRound = (nInt + 1) * nSign
        Else
            FNRound = nInt * nSign
        End If
    End Function
    Private Function min(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the minimum of two numbers
        If nFirst <= nSecond Then
            min = nFirst
        Else
            min = nSecond
        End If
    End Function
    Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the maximum of two numbers
        If nFirst >= nSecond Then
            max = nFirst
        Else
            max = nSecond
        End If
    End Function
    Private Function FN_CalcLength(ByRef xyStart As WaistMain.XY, ByRef xyEnd As WaistMain.XY) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

    End Function
    Function FN_CalcAngle(ByRef xyStart As WaistMain.XY, ByRef xyEnd As WaistMain.XY) As Double
        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360
        Dim X, y As Object
        Dim rAngle As Double

        X = xyEnd.X - xyStart.X
        y = xyEnd.Y - xyStart.Y

        'Horizomtal
        If X = 0 Then
            If y > 0 Then
                FN_CalcAngle = 90
            Else
                FN_CalcAngle = 270
            End If
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        If y = 0 Then
            If X > 0 Then
                FN_CalcAngle = 0
            Else
                FN_CalcAngle = 180
            End If
            Exit Function
        End If

        'All other cases
        rAngle = System.Math.Atan(y / X) * (180 / WaistMain.PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        If y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If

    End Function
    Sub PR_CalcPolar(ByRef xyStart As WaistMain.XY, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As WaistMain.XY)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim A, B As Double

        'Convert from degees to radians
        nAngle = nAngle * WaistMain.PI / 180

        B = System.Math.Sin(nAngle) * nLength
        A = System.Math.Cos(nAngle) * nLength

        xyReturn.X = xyStart.X + A
        xyReturn.Y = xyStart.Y + B

    End Sub
    Function FN_CirLinInt(ByRef xyStart As WaistMain.XY, ByRef xyEnd As WaistMain.XY, ByRef xyCen As WaistMain.XY, ByRef nRad As Double, ByRef xyInt As WaistMain.XY) As Short
        'Function to calculate the intersection between
        'a line and a circle.
        'Note:-
        '    Returns true if intersection found.
        '    The first intersection (only) is found.
        '    Ported from DRAFIX CAD DLG version.
        '

        Static nM, nC, nA, nSlope, nB, nK, nCalcTmp As Object
        Static nRoot As Double
        Static nSign As Short

        nSlope = FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        If nSlope = 0 Or nSlope = 180 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = xyStart.Y
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Vertical Line
        If nSlope = 90 Or nSlope = 270 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.Y + System.Math.Sqrt(nC) * nSign
                If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                    xyInt.Y = nRoot
                    xyInt.X = xyStart.X
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Non-othogonal line
        If nSlope > 0 Then
            nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X) 'Slope
            nK = xyStart.Y - nM * xyStart.X 'Y-Axis intercept
            nA = (1 + nM ^ 2)
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False
    End Function
    Sub PR_DrawXMarker(ByRef xyInsertion As WaistMain.XY)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As WaistMain.XY
        PR_CalcPolar(xyBase, 135, 0.0625, xyStart)
        PR_CalcPolar(xyBase, -45, 0.0625, xyEnd)
        PR_CalcPolar(xyBase, 45, 0.0625, xySecSt)
        PR_CalcPolar(xyBase, -135, 0.0625, xySecEnd)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("X Marker") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "X Marker"
                Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.Y, 0), New Point3d(xyEnd.X, xyEnd.Y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xySecSt.X, xySecSt.Y, 0), New Point3d(xySecEnd.X, xySecEnd.Y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl("X Marker")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyInsertion.X, xyInsertion.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawArc(ByRef xyCen As WaistMain.XY, ByRef nRad As Double, ByRef nStartAng As Double, ByRef nEndAng As Double)
        ' this procedure draws an arc between two points

        Dim nDeltaAng As Object
        nDeltaAng = nEndAng - nStartAng

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create an arc that is at 6.25,9.125 with a radius of 6, and
            '' starts at 64 degrees and ends at 204 degrees
            Using acArc As Arc = New Arc(New Point3d(xyCen.X, xyCen.Y, 0),
                                         nRad, nStartAng, nEndAng)

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyCen.X, xyCen.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawLine(ByRef xyStart As WaistMain.XY, ByRef xyFinish As WaistMain.XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                      OpenMode.ForWrite)

            '' Create a line that starts at 5,5 and ends at 12,3
            Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.Y, 0),
                                    New Point3d(xyFinish.X, xyFinish.Y, 0))

            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyStart.X, xyStart.Y, 0)))
            End If
            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''// File Name:	WHCUTDBD.D
        PR_GetBlkAttributeValues()
        Dim nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta, nBottArcToLargestOffset As Double
        ''// Get Origin
        Dim xyO, xyStart, xyEnd As WaistMain.XY
        PR_GetInsertionPoint()
        xyO = WaistMain.xyWHCutInsert
        Dim nDatum, nBodyLegTapePos As Integer
        Dim nFoldHt As Double = Val(g_sFoldHt)
        nDatum = Int(nFoldHt / 1.5) * 1.5
        nBodyLegTapePos = Int(nFoldHt / 1.5) + 6

        ''// NB Explict leg style given
        ''//
        Dim PantyLeg As Boolean = False
        Dim nLegStyle As Double = WHBODDIA1.fnGetNumber(WaistMain.g_sLegStyle, 1)
        Dim sLeg As String = "Right"
        If LeftLeg = True Then
            sLeg = "Left"
        End If
        If (nLegStyle = 1) Then
            PantyLeg = True
        End If

        ''// Calculate Cut Out And TOS
        ''//

        ''// For all subsequent calculations the measurements are worked
        ''// from the right edge of the body template #8654
        Dim nXscale, nYscale, nGivenTOSCir, nGivenWaistCir, nGivenMidPointCir, nGivenLargestCir As Double
        nXscale = 1.25 / 1.5
        nYscale = 0.5

        ''// Scale given values to Decimal inches
        ''// Retain the Given value for later use as the other value will be FIGURED
        ''//
        ''// Circumferences
        nGivenTOSCir = nTOSCir
        nGivenWaistCir = nWaistCir
        nGivenMidPointCir = nMidPointCir
        nGivenLargestCir = nLargestCir

        '' // Start Calculating XY control points
        ''// First Figure Heights And Circumferences
        ''// Largest part of buttocks
        Dim nLength, nLargestHt, nMidPointHt, nThighCirOriginal, nThighCir, nPantyThighRed2, nSeam As Double
        nLength = (nFoldHt + (nWaistHt - nFoldHt) / 3)
        nLargestHt = FNRound((nLength - nDatum) * nXscale)
        nLargestCir = FNRound(nLargestCir * nLargestGivenRed)

        ''// Mid Point (Cleft)
        nMidPointHt = (nLength + (nWaistHt - nFoldHt) / 4)   ''// N.B. Carry through Of nLength from above
        nMidPointHt = FNRound((nMidPointHt - nDatum) * nXscale)
        Dim xyMidPoint, xyFoldPanty, xyButtockArcCen, xyLargest, xyFold As WaistMain.XY
        xyMidPoint.X = xyO.X + nMidPointHt
        nMidPointCir = FNRound(nMidPointCir * nMidPointGivenRed)

        ''// Fold of Buttocks (11.1)
        nPantyThighRed2 = 0.92
        nSeam = 0.1875
        nFoldHt = FNRound((nFoldHt - nDatum) * nXscale)
        nThighCirOriginal = min(nLeftThighCir, nRightThighCir)
        If (PantyLeg) Then
            '' // 92% mark at thigh
            xyFoldPanty.X = xyO.X + nFoldHt
            xyFoldPanty.Y = xyO.Y + FNRound(nThighCirOriginal * nPantyThighRed2) * nYscale + nSeam
        End If
        nThighCir = FNRound(nThighCirOriginal * nThighGivenRed)

        xyFold.X = xyO.X + nFoldHt
        xyFold.Y = xyO.Y + nThighCir * nYscale + nSeam

        ''//  Largest part of Buttocks (12.1)
        xyButtockArcCen.Y = xyO.Y + (nThighCir / 2) * nYscale + nSeam
        xyButtockArcCen.X = xyO.X + nLargestHt
        nLength = FN_CalcLength(xyButtockArcCen, xyFold)

        xyLargest.Y = xyButtockArcCen.Y + nLength
        xyLargest.X = xyButtockArcCen.X

        ''// WaistCir
        nWaistCir = FNRound(nWaistCir * nWaistGivenRed)

        ''// TOSCir
        nTOSCir = FNRound(nTOSCir * nTOSGivenRed)
        ''// CutOut  (13.1)
        Dim OpenCrotch, ClosedCrotch As Boolean
        OpenCrotch = False
        ClosedCrotch = False
        If (sCrotchStyle.Equals("Open Crotch")) Then
            OpenCrotch = True
        Else
            ClosedCrotch = True
        End If

        ''// Get distance to largest part of buttocks from seam line (13)
        ''// nCrotchFrontFactor from WHBODDIA - VB programme
        Dim nCrotchFrontFactor, nOpenFront, nOpenBack, nOpenOff As Double
        nCrotchFrontFactor = WHBODDIA1.fnGetNumber(WaistMain.g_sBody, 1)
        nOpenFront = WHBODDIA1.fnGetNumber(WaistMain.g_sBody, 2)
        nOpenBack = WHBODDIA1.fnGetNumber(WaistMain.g_sBody, 3)

        Dim nAge As Integer = Val(txtAge.Text)
        Dim nHeelLength As Double = WHBODDIA1.fnGetNumber(WaistMain.g_sAnkleTape, 1)
        Dim Footless As Boolean = False
        If WHBODDIA1.fnGetNumber(WaistMain.g_sAnkleTape, 2) = 1 Then
            Footless = True
        End If
        If sCrotchStyle = "Open Crotch" Then
            If txtSex.Text = "Male" Then
                nOpenOff = 0.75
                'Check for footless style or a footless w40c
                If nLegStyle = 1 Or nLegStyle = 2 Or (nLegStyle = 3 And Footless = True) Or (nLegStyle = 4 And Footless = True) Then
                    If nAge <= 10 Then
                        nOpenOff = 0.375
                    End If
                Else
                    If nHeelLength > 0 And nHeelLength < 9 Then
                        nOpenOff = 0.375
                    End If
                End If
            Else
                nOpenOff = 0.375
            End If
        End If


        nLength = xyLargest.Y - xyO.Y - nSeam
        Dim xyCO_LargestBott, xyCO_LargestTop As WaistMain.XY
        xyCO_LargestBott.X = xyLargest.X
        xyCO_LargestBott.Y = xyO.Y + ((nLargestCir * nYscale - nLength) * nCrotchFrontFactor) + nSeam

        xyCO_LargestTop.X = xyLargest.X
        xyCO_LargestTop.Y = xyLargest.Y - ((nLargestCir * nYscale - nLength) * (1 - nCrotchFrontFactor))

        Dim nCO_HalfWidth As Double = FN_CalcLength(xyCO_LargestBott, xyCO_LargestTop) / 2
        ''// Check on Cut Out Diameter
        Dim nCutOutDiaMaxTol As Double = 6.0
        If (FNRound(nCO_HalfWidth * 2) > nCutOutDiaMaxTol) Then
            MsgBox("Warning, Cut Out Diameter is larger than 6 inches" +
                   "\nCalculated diameter is " + Format("length", nCO_HalfWidth * 2), 16, "Waist - Dialog")
        End If

        ''// Cut Out 
        ''// Top Of Support (if given) (Cut Out Front)
        Dim nDiff As Integer
        Dim xyCO_TOSBott, xyCO_WaistBott, xyCO_MidPointBott As WaistMain.XY
        Dim nBodyFrontReduceOff As Double = 0.125
        Dim nBodyFrontIncreaseOff As Double = 0.25
        If (nTOSCir > 0) Then
            nDiff = Int(nGivenTOSCir - nGivenLargestCir)
            If (nDiff < 0) Then nLength = nDiff * nBodyFrontReduceOff
        Else
            nLength = nDiff * nBodyFrontIncreaseOff
            xyCO_TOSBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
            xyCO_TOSBott.X = xyO.X + FNRound((nTOSHt - nDatum) * nXscale)
        End If
        ''// Waist  (Cut Out Front)
        nDiff = Int(nGivenWaistCir - nGivenLargestCir)
        If (nDiff < 0) Then
            nLength = nDiff * nBodyFrontReduceOff
        Else
            nLength = nDiff * nBodyFrontIncreaseOff
            xyCO_WaistBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
            xyCO_WaistBott.X = xyO.X + FNRound((nWaistHt - nDatum) * nXscale)
        End If

        ''// MidPoint  (Cut Out Front)
        ''// Remember that mid point height was established earlier
        nDiff = Int(nGivenMidPointCir - nGivenLargestCir)
        If (nDiff < 0) Then
            nLength = nDiff * nBodyFrontReduceOff
        Else
            nLength = nDiff * nBodyFrontIncreaseOff
            xyCO_MidPointBott.Y = xyCO_LargestBott.Y + nLength * nYscale '' // NB Sign carrys through
            xyCO_MidPointBott.X = xyO.X + nMidPointHt
        End If

        ''// Cut Out (Back)
        ''// Revise Cut Out Diameter to complete cut out
        Dim nCO_Width, nOriginalCO_WaistTopY, nOriginalCO_TOSTopY As Double
        Dim nCutOutConstructFac_1 As Double = 0.87
        Dim nEndBackBodyOff As Double = 1.25
        Dim xyCO_WaistTop, xyCO_TOSTop, xyCO_MidPointTop As WaistMain.XY
        nCO_Width = nCO_HalfWidth * 2 * nCutOutConstructFac_1

        ''// Waist  (Cut Out Back)
        xyCO_WaistTop.Y = xyCO_WaistBott.Y + nCO_Width
        xyCO_WaistTop.X = xyCO_WaistBott.X

        ''// Top Of Support   (Cut Out Back)
        xyCO_TOSTop.Y = xyCO_WaistTop.Y
        If (nTOSCir > 0) Then
            xyCO_TOSTop.X = xyCO_TOSBott.X
        Else
            xyCO_TOSTop.X = xyCO_WaistTop.X + nEndBackBodyOff
        End If

        ''// Back of body
        ''// Get Back body offsets
        ''// Check against Minimun value
        Dim IgnoreMidPoint_CO As Boolean = False
        nOriginalCO_WaistTopY = xyCO_WaistTop.Y
        nOriginalCO_TOSTopY = xyCO_TOSTop.Y

        Dim bIsLoop As Boolean = True
        Dim nWaistBackOff, nTOSBackOff, aAngle, nMidPointBackOff, nMinBackOff As Double
        Dim nBodyBackCutOutMinTol As Double = 3.0
        While (bIsLoop)
            '' // Waist offset
            nWaistBackOff = FNRound(nWaistCir * nYscale)
            nWaistBackOff = (nWaistBackOff - ((xyCO_WaistTop.Y - xyO.Y - nSeam) + (xyCO_WaistBott.Y - xyO.Y - nSeam))) / 2
            ''// TOS Offset
            If (nTOSCir > 0) Then
                nTOSBackOff = FNRound(nTOSCir * nYscale)
                nTOSBackOff = (nTOSBackOff - ((xyCO_TOSTop.Y - xyO.Y - nSeam) + (xyCO_TOSBott.Y - xyO.Y - nSeam))) / 2
            Else
                nTOSBackOff = nWaistBackOff
            End If

            ''// MidPoint Offset
            nLength = xyCO_MidPointBott.X - xyCO_LargestBott.X
            aAngle = FN_CalcAngle(xyCO_LargestTop, xyCO_WaistTop)
            xyCO_MidPointTop.X = xyCO_MidPointBott.X
            xyCO_MidPointTop.Y = xyCO_LargestTop.Y + (System.Math.Tan(aAngle) * nLength)
            nMidPointBackOff = FNRound(nMidPointCir * nYscale)
            nMidPointBackOff = (nMidPointBackOff - ((xyCO_MidPointTop.Y - xyO.Y - nSeam) + (xyCO_MidPointBott.Y - xyO.Y - nSeam))) / 2
            ''// Check that 3" distance is meet
            ''// Note use of 1/2 scale (nYscale) w.r.t.  nBodyBackCutOutMinTol
            bIsLoop = False
            If (IgnoreMidPoint_CO) Then
                nMinBackOff = FNRound(min(nWaistBackOff, nTOSBackOff))
            Else
                nMinBackOff = FNRound(min(min(nWaistBackOff, nTOSBackOff), nMidPointBackOff))
            End If
            If nMinBackOff < (nBodyBackCutOutMinTol * nYscale) Then
                bIsLoop = True
                nDiff = FNRound(((nBodyBackCutOutMinTol * nYscale) - nMinBackOff) * 2)
                If nDiff = 0 Then
                    Exit While
                End If
                xyCO_WaistTop.Y = xyCO_WaistTop.Y - nDiff
                xyCO_TOSTop.Y = xyCO_TOSTop.Y - nDiff
                If xyCO_WaistTop.Y < xyCO_WaistBott.Y Then
                    IgnoreMidPoint_CO = True
                    xyCO_WaistTop.Y = nOriginalCO_WaistTopY
                    xyCO_TOSTop.Y = nOriginalCO_TOSTopY
                    MsgBox("Warning, Mid-Point ignored in calculating back Cut-Out!" +
                           "\nDistance between back of Cut-Out at Mid-Point and Profile may be less than 3", 16, "Waist - Dialog")
                End If
            End If
        End While

        ''// Calculate back body XY points
        ''// Waist
        Dim xyWaist, xyTOS, xyCO_CenterArrow, xyCO_ArcCen As WaistMain.XY
        xyWaist.X = xyCO_WaistTop.X
        xyWaist.Y = xyCO_WaistTop.Y + nWaistBackOff

        ''// Top Of Support
        xyTOS.X = xyCO_TOSTop.X
        xyTOS.Y = xyCO_TOSTop.Y + nTOSBackOff
        ''// MidPoint
        xyMidPoint.X = xyCO_MidPointTop.X
        xyMidPoint.Y = xyCO_MidPointTop.Y + nMidPointBackOff
        ''// Center point of cutout (also used in crotch labeling)
        Dim nCutOutConstructOff_1 As Double = 0.75
        xyCO_CenterArrow.X = xyFold.X + nCutOutConstructOff_1
        xyCO_CenterArrow.Y = xyCO_LargestBott.Y + nCO_HalfWidth

        ''// Establish center, And angles of cutout arc for drawing purposes	
        xyCO_ArcCen.Y = xyCO_CenterArrow.Y
        Dim nCO_ArcRadius As Double = nCO_HalfWidth
        Dim aCO_ArcStart, aCO_ArcDelta, nArcCenToLargestOffset, nTopArcSegment, aCO_ArcStartOriginal As Double
        xyCO_ArcCen.X = xyCO_CenterArrow.X + nCO_ArcRadius
        If (xyCO_ArcCen.X > xyCO_LargestTop.X) Then
            nLength = FN_CalcLength(xyCO_CenterArrow, xyCO_LargestTop) / 2
            aAngle = FN_CalcAngle(xyCO_CenterArrow, xyCO_LargestTop)
            nCO_ArcRadius = nLength / System.Math.Cos(aAngle)
            xyCO_ArcCen.X = xyCO_CenterArrow.X + nCO_ArcRadius
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyCO_LargestTop)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyCO_LargestBott) - aCO_ArcStart
        Else
            aCO_ArcStart = 90
            aCO_ArcDelta = 180
        End If

        ''// Additions for OPEN Crotch
        Dim xyCO_OpenArcTop, xyCO_OpenArcBott, xyPt1, xyPt2, xyInt As WaistMain.XY
        Dim nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta, nTopArcToLargestOffset As Double
        If OpenCrotch Then
            ''//Get Cut out Arc Center to Largest part of buttocks offset
            nArcCenToLargestOffset = xyCO_LargestTop.X - xyCO_ArcCen.X
            If nArcCenToLargestOffset < 0 Then
                nArcCenToLargestOffset = 0 '' //Make sure Not -ve
            End If
            ''// For back / top of crotch establish length of top arc
            nTopArcSegment = (180 - aCO_ArcStart) * WaistMain.PI / 180 * nCO_ArcRadius
            aCO_ArcStartOriginal = aCO_ArcStart '' // Store For possible use With bottom arc
            If nTopArcSegment > nOpenBack Then
                ''	// End Of Top O.C. lands On arc  
                xyCO_OpenArcTop = xyCO_ArcCen
                nCO_OpenArcTopRadius = nCO_ArcRadius
                aCO_OpenArcTopStart = aCO_ArcStart
                aAngle = 180 - (nOpenBack * 180) / (WaistMain.PI * nCO_ArcRadius)
                aCO_OpenArcTopDelta = aAngle - aCO_OpenArcTopStart
                ''// Modify start of Cut-Out arc
                aCO_ArcStart = aCO_OpenArcTopStart + aCO_OpenArcTopDelta
            Else
                ''//Get distance from end of arc to lagest part of buttocks
                nTopArcToLargestOffset = nArcCenToLargestOffset - (nOpenBack - nTopArcSegment)
            End If
            ''// Do Not go past Largest part of buttocks
            If (nTopArcToLargestOffset < 0) Then
                nTopArcToLargestOffset = 0
            End If

            ''// For front / bottom of crotch establish length of arc
            Dim nBottArcSegment As Double
            nBottArcSegment = ((aCO_ArcStartOriginal + aCO_ArcDelta) - 180) * WaistMain.PI / 180 * nCO_ArcRadius
            If (nBottArcSegment > nOpenFront) Then ''	// End Of Bottom O.C. lands On arc  
                xyCO_OpenArcBott = xyCO_ArcCen
                nCO_OpenArcBottRadius = nCO_ArcRadius
                aCO_OpenArcBottStart = 180 + (nOpenFront * 180) / (WaistMain.PI * nCO_ArcRadius)
                aCO_OpenArcBottDelta = aCO_ArcStartOriginal + aCO_ArcDelta - aCO_OpenArcBottStart
                ''// Modify end of Cut-Out arc
                aCO_ArcDelta = aCO_OpenArcBottStart - aCO_ArcStart
            Else
                ''//Get distance from end of arc to lagest part of buttocks
                nBottArcToLargestOffset = nArcCenToLargestOffset - (nOpenFront - nBottArcSegment)
                ''// Change aCO_ArcDelta if Start has moved (New bit)
                If (aCO_ArcStartOriginal <> aCO_ArcStart) Then
                    aCO_ArcDelta = aCO_ArcDelta - (aCO_ArcStart - aCO_ArcStartOriginal)
                End If
            End If
            ''// Do Not go past Largest part of buttocks
            If (nBottArcToLargestOffset < 0) Then
                nBottArcToLargestOffset = 0
            End If

            ''// Establish revised arc for cut out including offset for open crotch
            ''// Ensure that End Points of arc do Not pass largest part of buttocks

            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, xyPt1)
            If (xyPt1.X > xyCO_LargestTop.X) Then
                xyPt1 = xyCO_LargestTop
                nTopArcToLargestOffset = 0
            End If
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta, xyPt2)
            If (xyPt2.X > xyCO_LargestBott.X) Then
                xyPt2 = xyCO_LargestBott
                nBottArcToLargestOffset = 0
            End If
            PR_MakeXY(xyEnd, xyPt1.X, xyPt1.Y + 2)
            Dim nError As Short = FN_CirLinInt(xyPt1, xyEnd, xyCO_ArcCen, nCO_ArcRadius + nOpenOff, xyInt)
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyInt)
            PR_MakeXY(xyEnd, xyPt2.X, xyPt2.Y - 2)
            nError = FN_CirLinInt(xyPt2, xyEnd, xyCO_ArcCen, nCO_ArcRadius + nOpenOff, xyInt)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyInt) - aCO_ArcStart

            nCO_ArcRadius = nCO_ArcRadius + nOpenOff
        End If

        ''//
        ''// DRAW Cutout And TOS
        ''//
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("marker","xmarker",xyO , 0.1, 0.1) 
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "O")
        PR_DrawXMarker(xyO)
        ''hEnt = AddEntity("marker", "xmarker", xyTOS, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "TOS")
        PR_DrawXMarker(xyTOS)
        ''hEnt = AddEntity("marker", "xmarker", xyWaist, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Waist")
        PR_DrawXMarker(xyWaist)
        ''hEnt = AddEntity("marker", "xmarker", xyMidPoint, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "MidPoint")
        PR_DrawXMarker(xyMidPoint)
        ''hEnt = AddEntity("marker", "xmarker", xyLargest, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Largest")
        PR_DrawXMarker(xyLargest)
        ''hEnt = AddEntity("marker", "xmarker", xyButtockArcCen, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "ButtockArcCen")
        PR_DrawXMarker(xyButtockArcCen)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_TOSBott, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_TOSBott")
        PR_DrawXMarker(xyCO_TOSBott)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_TOSTop, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_TOSTop")
        PR_DrawXMarker(xyCO_TOSTop)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_WaistBott, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_WaistBott")
        PR_DrawXMarker(xyCO_WaistBott)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_WaistTop, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_WaistTop")
        PR_DrawXMarker(xyCO_WaistTop)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_MidPointBott, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_MidPointBott")
        PR_DrawXMarker(xyCO_MidPointBott)
        ''hEnt = AddEntity("marker", "xmarker", xyCO_MidPointTop, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_MidPointTop")
        PR_DrawXMarker(xyCO_MidPointTop)
        '' hEnt = AddEntity("marker", "xmarker", xyFold, 0.1, 0.1)
        '' SetDBData(hEnt, "ID", sFileNo + sLeg + "Fold")
        PR_DrawXMarker(xyFold)
        If (PantyLeg) Then
            ''hEnt = AddEntity("marker", "xmarker", xyFoldPanty, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "FoldPanty")
            PR_DrawXMarker(xyFoldPanty)
        End If
        '' hEnt = AddEntity("marker", "xmarker", xyCO_LargestTop, 0.1, 0.1) 	
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_LargestTop")
        ''SetDBData(hEnt, "Data", MakeString("scalar", nTopArcToLargestOffset))
        PR_DrawXMarker(xyCO_LargestTop)

        ''hEnt = AddEntity("marker", "xmarker", xyCO_LargestBott, 0.1, 0.1)
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_LargestBott") ;
        ''SetDBData(hEnt, "Data", MakeString("scalar", nBottArcToLargestOffset));
        PR_DrawXMarker(xyCO_LargestBott)

        ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1) ;		
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_ArcCen") ;
        'SetDBData(hEnt, "Data", MakeString("scalar", nCO_ArcRadius) +
        '          " " + MakeString("scalar", aCO_ArcStart) +
        '          " " + MakeString("scalar", aCO_ArcDelta)) ;
        PR_DrawXMarker(xyCO_ArcCen)

        If xyCO_OpenArcBott.X <> 0 And xyCO_OpenArcBott.Y <> 0 Then
            ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1)
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_OpenArcBott")
            'SetDBData(hEnt, "Data", MakeString("scalar", nCO_OpenArcBottRadius) +
            '          " " + MakeString("scalar", aCO_OpenArcBottStart) +
            '          " " + MakeString("scalar", aCO_OpenArcBottDelta))
            PR_DrawXMarker(xyCO_ArcCen)
        End If

        If xyCO_OpenArcTop.X <> 0 And xyCO_OpenArcTop.Y <> 0 Then
            ''hEnt = AddEntity("marker", "xmarker", xyCO_ArcCen, 0.1, 0.1) ;		
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_OpenArcTop") ;
            'SetDBData(hEnt, "Data", MakeString("scalar", nCO_OpenArcTopRadius) +
            '          " " + MakeString("scalar", aCO_OpenArcTopStart) +
            '          " " + MakeString("scalar", aCO_OpenArcTopDelta)) ;
            PR_DrawXMarker(xyCO_ArcCen)
        End If

        ''hEnt = AddEntity("marker", "xmarker", xyCO_CenterArrow, 0.1, 0.1) ;	
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "CO_CenterArrow");
        ''SetDBData(hEnt, "Data", MakeString("scalar", nOpenOff)) ; 
        PR_DrawXMarker(xyCO_CenterArrow)
        ''// Check for special case where the user has allowed a difference of more than 5%
        ''// Between reductions add a marker 5% away from existing marker at fold.
        Dim xyTmp As WaistMain.XY
        If System.Math.Abs(nLargestGivenRed - nThighGivenRed) > 0.05 Then
            xyTmp.X = xyFold.X
            xyTmp.Y = xyO.Y + FNRound(nThighCirOriginal * (nThighGivenRed + 0.05)) * nYscale + nSeam
            ''hEnt = AddEntity("marker", "xmarker", xyTmp, 0.1, 0.1) ;	
            ''SetDBData(hEnt, "ID", sFileNo + sLeg + "Fold+5%");
            PR_DrawXMarker(xyTmp)
        End If

        If (LeftLeg) Then
            ARMDIA1.PR_SetLayer("TemplateLeft")
        Else
            ARMDIA1.PR_SetLayer("TemplateRight")
        End If
        ''//
        ''// DRAW Cutout And TOS
        ''//
        Dim xyFilletTop As WaistMain.XY
        Dim nFilletRadius As Double = 0.1875
        If (OpenCrotch) Then
            ''// Get Start point of arc 
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, xyPt1)
            ''// TOP Fillet radius
            PR_MakeXY(xyStart, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y)
            PR_MakeXY(xyEnd, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y + nCO_ArcRadius)
            Dim nError As Short = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
            xyFilletTop = xyInt
            ''// TOP
            If xyCO_OpenArcTop.X <> 0 And xyCO_OpenArcTop.Y <> 0 Then
                ''AddEntity("arc", xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta)
                PR_DrawArc(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta)
                PR_CalcPolar(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart + aCO_OpenArcTopDelta, xyPt2)
                ''----AddEntity("line", xyPt1.X, xyFilletTop.Y, xyPt2)
                PR_MakeXY(xyStart, xyPt1.X, xyFilletTop.Y)
                PR_DrawLine(xyStart, xyPt2)
                PR_CalcPolar(xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, xyTmp)
                If (xyTmp.X < xyCO_LargestTop.X) Then
                    ''AddEntity("line", xyTmp, xyCO_LargestTop)
                    PR_DrawLine(xyTmp, xyCO_LargestTop)
                End If
                ''// Fillet
                ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
            Else
                nLength = xyCO_LargestTop.X - xyPt1.X - nTopArcToLargestOffset
                If (nLength > 0) Then
                    PR_CalcPolar(xyPt1, nLength, 0, xyPt2)
                    If (nLength > nFilletRadius) Then
                        xyFilletTop = xyPt1 ''	// gets aCO_ArcStart correct
                        ''AddEntity("arc", xyPt2.X - nFilletRadius, xyPt2.Y - nFilletRadius, nFilletRadius, 0, 90)
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyPt2.Y - nFilletRadius)
                        PR_DrawArc(xyStart, nFilletRadius, 0, 90)
                        ''AddEntity("line", xyPt1, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_DrawLine(xyPt1, xyEnd)
                        ''AddEntity("line", xyPt2.X, xyPt2.Y - nFilletRadius, xyPt2.X, xyCO_LargestTop.Y)
                        PR_MakeXY(xyStart, xyPt2.X, xyPt2.Y - nFilletRadius)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestTop.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    Else
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y + nCO_ArcRadius)
                        nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
                        xyFilletTop = xyInt
                        ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                        PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
                        ''AddEntity("line", xyPt2.X, xyFilletTop.Y, xyPt2.X, xyCO_LargestTop.Y)
                        PR_MakeXY(xyStart, xyPt2.X, xyFilletTop.Y)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestTop.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    End If

                Else
                    ''AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop))
                    PR_DrawArc(xyFilletTop, nFilletRadius, 0, FN_CalcAngle(xyCO_ArcCen, xyFilletTop))
                    ''AddEntity("line", xyPt1.X, xyFilletTop.Y, xyCO_LargestTop)
                    PR_MakeXY(xyStart, xyPt1.X, xyFilletTop.Y)
                    PR_DrawLine(xyStart, xyCO_LargestTop)
                End If

                If (nTopArcToLargestOffset > 0) Then
                    'AddEntity("line",
                    '              CalcXY("relpolar", xyCO_LargestTop, nTopArcToLargestOffset, 180),
                    '                          xyCO_LargestTop)
                    PR_CalcPolar(xyCO_LargestTop, nTopArcToLargestOffset, 180, xyStart)
                    PR_DrawLine(xyStart, xyCO_LargestTop)
                End If
            End If
            ''AddEntity("line", xyCO_LargestTop, xyCO_WaistTop) ;
            PR_DrawLine(xyCO_LargestTop, xyCO_WaistTop)
            '' AddEntity("line", xyCO_WaistTop, xyCO_TOSTop) ;
            PR_DrawLine(xyCO_WaistTop, xyCO_TOSTop)

            ''// BOTTOM
            ''// End points of arc 
            PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta, xyPt1)

            ''// BOTTOM Fillet radius initial value
            Dim xyFilletBott As WaistMain.XY
            PR_MakeXY(xyStart, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y)
            PR_MakeXY(xyEnd, xyPt1.X - nFilletRadius, xyCO_ArcCen.Y - nCO_ArcRadius)
            nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
            xyFilletBott = xyInt

            If xyCO_OpenArcBott.X <> 0 And xyCO_OpenArcBott.Y <> 0 Then
                ''AddEntity("arc", xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta) ;
                PR_DrawArc(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta)
                PR_CalcPolar(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, xyPt2)
                ''AddEntity("line", xyPt1.X, xyFilletBott.Y, xyPt2) ;
                PR_MakeXY(xyStart, xyPt1.X, xyFilletBott.Y)
                PR_DrawLine(xyStart, xyPt2)
                PR_CalcPolar(xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart + aCO_OpenArcBottDelta, xyTmp)
                If (xyTmp.X < xyCO_LargestBott.X) Then
                    ''AddEntity ("line", xyTmp, xyCO_LargestBott)
                    PR_DrawLine(xyTmp, xyCO_LargestBott)
                End If
                ''// Fillet
                aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
            Else
                nLength = xyCO_LargestBott.X - xyPt1.X - nBottArcToLargestOffset
                If (nLength > 0) Then
                    PR_CalcPolar(xyPt1, nLength, 0, xyPt2)
                    If (nLength > nFilletRadius) Then
                        xyFilletBott = xyPt1    ''// gets aCO_ArcStart correct
                        ''AddEntity("arc", xyPt2.X - nFilletRadius, xyPt2.Y + nFilletRadius, nFilletRadius, 270, 90);
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyPt2.Y + nFilletRadius)
                        PR_DrawArc(xyStart, nFilletRadius, 270, 90)
                        ''AddEntity("line", xyPt1, xyPt2.X - nFilletRadius, xyPt2.Y) ;
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyPt2.Y)
                        PR_DrawLine(xyPt1, xyEnd)
                        ''AddEntity("line", xyPt2.X, xyPt2.Y + nFilletRadius, xyPt2.X, xyCO_LargestBott.Y);
                        PR_MakeXY(xyStart, xyPt2.X, xyPt2.Y + nFilletRadius)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestBott.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    Else
                        PR_MakeXY(xyStart, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y)
                        PR_MakeXY(xyEnd, xyPt2.X - nFilletRadius, xyCO_ArcCen.Y - nCO_ArcRadius)
                        nError = FN_CirLinInt(xyStart, xyEnd, xyCO_ArcCen, nCO_ArcRadius - nFilletRadius, xyInt)
                        xyFilletBott = xyInt
                        aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                        ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                        PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
                        ''AddEntity("line", xyPt2.x, xyFilletBott.y, xyPt2.x, xyCO_LargestBott.y);
                        PR_MakeXY(xyStart, xyPt2.X, xyFilletBott.Y)
                        PR_MakeXY(xyEnd, xyPt2.X, xyCO_LargestBott.Y)
                        PR_DrawLine(xyStart, xyEnd)
                    End If
                Else
                    aAngle = FN_CalcAngle(xyCO_ArcCen, xyFilletBott)
                    ''AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle);
                    PR_DrawArc(xyFilletBott, nFilletRadius, aAngle, 360 - aAngle)
                    ''AddEntity("line", xyPt1.x, xyFilletBott.y, xyCO_LargestBott);
                    PR_MakeXY(xyStart, xyPt1.X, xyFilletBott.Y)
                    PR_DrawLine(xyStart, xyCO_LargestBott)
                End If

                If (nBottArcToLargestOffset > 0) Then
                    'AddEntity("line",
                    '      CalcXY("relpolar", xyCO_LargestBott, nBottArcToLargestOffset, 180),
                    '                  xyCO_LargestBott);
                    PR_CalcPolar(xyCO_LargestBott, nBottArcToLargestOffset, 180, xyStart)
                    PR_DrawLine(xyStart, xyCO_LargestBott)
                End If
            End If
            ''StartPoly("openfitted")
            '' AddVertex(xyCO_LargestBott)
            ''AddVertex(xyCO_MidPointBott)
            ''AddVertex(xyCO_WaistBott)
            Dim ptColl As Point3dCollection = New Point3dCollection()
            ptColl.Add(New Point3d(xyCO_LargestBott.X, xyCO_LargestBott.Y, 0))
            ptColl.Add(New Point3d(xyCO_MidPointBott.X, xyCO_MidPointBott.Y, 0))
            ptColl.Add(New Point3d(xyCO_WaistBott.X, xyCO_WaistBott.Y, 0))
            If (nTOSCir > 0) Then
                ''AddVertex(xyCO_TOSBott)
                ptColl.Add(New Point3d(xyCO_TOSBott.X, xyCO_TOSBott.Y, 0))
            End If
            ''EndPoly()
            PR_DrawPoly(ptColl)

            ''// Draw arc taking account of fillets
            aCO_ArcStart = FN_CalcAngle(xyCO_ArcCen, xyFilletTop)
            aCO_ArcDelta = FN_CalcAngle(xyCO_ArcCen, xyFilletBott) - aCO_ArcStart
            ''AddEntity("arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
            PR_DrawArc(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta)
        Else
            ''// Closed Crotch
            ''AddEntity("arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
            PR_DrawArc(xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta)
            ''// TOP
            ''AddEntity("line", xyCO_LargestTop, xyCO_WaistTop) ;
            PR_DrawLine(xyCO_LargestTop, xyCO_WaistTop)
            '' AddEntity("line", xyCO_WaistTop, xyCO_TOSTop) ;
            PR_DrawLine(xyCO_WaistTop, xyCO_TOSTop)
            ''// FIDDLY BITS, If center of cut out Is below largest part of buttocks then put in joining lines 
            If (xyCO_ArcCen.X < xyCO_LargestTop.X) Then
                ''AddEntity("line", CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, 90), xyCO_LargestTop);
                PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, 90, xyStart)
                PR_DrawLine(xyStart, xyCO_LargestTop)
                ''AddEntity("line", CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, 270), xyCO_LargestBott);
                PR_CalcPolar(xyCO_ArcCen, nCO_ArcRadius, 270, xyStart)
                PR_DrawLine(xyStart, xyCO_LargestBott)
            End If
            ''// BOTTOM
            ''StartPoly("openfitted")
            ''AddVertex(xyCO_LargestBott)
            ''AddVertex(xyCO_MidPointBott)
            ''AddVertex(xyCO_WaistBott)
            Dim ptColl1 As Point3dCollection = New Point3dCollection()
            ptColl1.Add(New Point3d(xyCO_LargestBott.X, xyCO_LargestBott.Y, 0))
            ptColl1.Add(New Point3d(xyCO_MidPointBott.X, xyCO_MidPointBott.Y, 0))
            ptColl1.Add(New Point3d(xyCO_WaistBott.X, xyCO_WaistBott.Y, 0))
            If (nTOSCir > 0) Then
                ''AddVertex(xyCO_TOSBott)
                ptColl1.Add(New Point3d(xyCO_TOSBott.X, xyCO_TOSBott.Y, 0))
            End If
            ''EndPoly()
            PR_DrawPoly(ptColl1)
        End If
        ''// Closing Line at waist Or TOS
        ''AddEntity("line", xyCO_TOSTop, xyTOS) ;
        PR_DrawLine(xyCO_TOSTop, xyTOS)
        If (nTOSCir > 0) Then
            ''AddEntity("line", xyCO_TOSBott, xyTOS.X, xyO.Y)
            PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
            PR_DrawLine(xyCO_TOSBott, xyEnd)
        Else
            ''AddEntity("line", xyCO_WaistBott, xyCO_WaistBott.X, xyO.Y)
            PR_MakeXY(xyEnd, xyCO_WaistBott.X, xyO.Y)
            PR_DrawLine(xyCO_WaistBott, xyEnd)
        End If

        ''// Draw construction lines
        ''// 
        ARMDIA1.PR_SetLayer("Construct")
        ''hEnt = AddEntity("line", xyFold.X, xyLargest.Y, xyFold.X, xyO.Y) ;
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "FoldLine") ; 
        PR_MakeXY(xyStart, xyFold.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyFold.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir * nYscale + nSeam) + " " + MakeString("scalar", nGivenWaistCir));     	

        ''AddEntity("line", xyLargest.X, xyLargest.Y,xyLargest.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyLargest.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyLargest.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''AddEntity("line", xyMidPoint.X, xyLargest.Y,xyMidPoint.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyMidPoint.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyMidPoint.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''hEnt = AddEntity("line", xyWaist.X, xyLargest.Y,xyWaist.X, xyO.Y) ;
        PR_MakeXY(xyStart, xyWaist.X, xyLargest.Y)
        PR_MakeXY(xyEnd, xyWaist.X, xyO.Y)
        PR_DrawLine(xyStart, xyEnd)

        If (nTOSCir! = 0) Then
            ''hEnt = AddEntity("line", xyTOS.X, xyLargest.Y, xyTOS.X, xyO.Y) ;
            PR_MakeXY(xyStart, xyTOS.X, xyLargest.Y)
            PR_MakeXY(xyEnd, xyTOS.X, xyO.Y)
        End If
        ''SetDBData(hEnt, "ID", sFileNo + sLeg + "EOSLine") ; 
        ''// Store figured thigh cir. And actual waist cir
        ''SetDBData(hEnt, "Data", MakeString("scalar", nThighCir * nYscale + nSeam) + " " + MakeString("scalar", nGivenWaistCir));     	

        ''AddEntity("arc", xyButtockArcCen, Calc("length", xyButtockArcCen, xyFold), 60, 70) ;
        PR_DrawArc(xyButtockArcCen, FN_CalcLength(xyButtockArcCen, xyFold), 60, 70)

        ''// Add Body template tapes
        ''// From body position given 
        Dim nn As Double
        If (nBodyLegTapePos > 0) Then
            nn = nBodyLegTapePos - 1
        Else
            nn = nLastTape - 1
        End If
        xyPt1.X = xyO.X
        xyPt1.Y = xyO.Y + nSeam + 0.5
        While (xyPt1.X < xyTOS.X)
            'sSymbol = MakeString("long", nn) + "tape"
            'If (!Symbol("find", sSymbol)) Then
            '    MsgBox("Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists\n", 16, "Waist - Dialog")
            '    Exit While
            'End If
            'AddEntity("symbol", sSymbol, xyPt1)
            xyPt1.X = xyPt1.X + 1.5 * nXscale
            nn = nn + 1
        End While
        ''// Seam TRAM Lines 
        ARMDIA1.PR_SetLayer("Notes")
        ''AddEntity("line", xyWHCutInsert.x, xyWHCutInsert.y + nSeam + 0.5, xyPt1.X - (1.5 * nXscale), xyPt1.Y) ; //*
        PR_MakeXY(xyStart, WaistMain.xyWHCutInsert.X, WaistMain.xyWHCutInsert.Y + nSeam + 0.5)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y)
        PR_DrawLine(xyStart, xyEnd)
        ''AddEntity("line", xyWHCutInsert.x, xyWHCutInsert.y + nSeam, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5) ; //*
        PR_MakeXY(xyStart, WaistMain.xyWHCutInsert.X, WaistMain.xyWHCutInsert.Y + nSeam)
        PR_MakeXY(xyEnd, xyPt1.X - (1.5 * nXscale), xyPt1.Y - 0.5)
        PR_DrawLine(xyStart, xyEnd)

        ''// Diagonal Fly Marks
        If (sCrotchStyle.Equals("Diagonal Fly") And sLeg.Equals("Left")) Then
            ''AddEntity("marker", "closed arrow", xyCO_LargestTop, 0.5, 0.125, 90)
            ''AddEntity("marker", "closed arrow", xyCO_LargestBott, 0.5, 0.125, 270)
        End If

        If (LeftLeg) Then
            ARMDIA1.PR_SetLayer("TemplateLeft")
        Else
            ARMDIA1.PR_SetLayer("TemplateRight")
        End If
        If (nTOSCir > 0) Then
            ''AddEntity("line", xyWHCutInsert, xyTOS.X, xyWHCutInsert.y)
            PR_MakeXY(xyEnd, xyTOS.X, WaistMain.xyWHCutInsert.Y)
            PR_DrawLine(WaistMain.xyWHCutInsert, xyEnd)
        Else
            ''AddEntity("line", xyWHCutInsert, xyWaist.X, xyWHCutInsert.y)
            PR_MakeXY(xyEnd, xyWaist.X, WaistMain.xyWHCutInsert.Y)
            PR_DrawLine(WaistMain.xyWHCutInsert, xyEnd)
        End If
    End Sub
    Sub PR_DrawPoly(ByRef PointCollection As Point3dCollection)
        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to draw a POLYLINE through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '
        '
        Dim ii As Short
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                            OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Using acPoly As Polyline = New Polyline()
                For ii = 0 To PointCollection.Count - 1
                    acPoly.AddVertexAt(ii, New Point2d(PointCollection(ii).X, PointCollection(ii).Y), 0, 0, 0)
                Next ii

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(WaistMain.xyWHCutInsert.X, WaistMain.xyWHCutInsert.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub

    Private Sub WaistTabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WaistTabControl.SelectedIndexChanged
        Dim TabIndex As Integer = WaistTabControl.SelectedIndex
        If TabIndex = 2 And g_bIsNeedLoad = True Then
            g_bIsNeedLoad = False
            WaistTabControl.TabPages(1).Height = Figure.Height
            ''Figure.ShowDialog()
            ''Figure.Refresh()
            Figure.PR_LoadFigureFromMain()
        End If
    End Sub

    Private Sub WHMAIN_frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsCloseDlg = True Then
            g_bIsCloseDlg = False
            Exit Sub
        End If
        If frmWaistBody.IsDisposed() = False Then
            If frmWaistBody.PR_CloseWaistBodyDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If frmLeg.IsDisposed() = False Then
            If frmLeg.PR_CloseWaistLegDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
        If Figure.IsDisposed() = False Then
            If Figure.PR_CloseFigureDialog() = False Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub
End Class